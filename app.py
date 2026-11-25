import os
import re
import subprocess
import tempfile

from docx import Document
from pdf2docx import Converter
import pdfplumber
from PIL import Image, ImageOps
import pytesseract
import streamlit as st
import pandas as pd

# إذا كان Tesseract مثبتاً على ويندوز، أزل التعليق وضع المسار الصحيح:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PDF_MIME = "application/pdf"

st.set_page_config(page_title="محول الملفات المنظم", page_icon="🔄", layout="wide")

st.markdown(
    """
<style>
    .stApp { direction: rtl; text-align: right; }
    .stDataFrame { direction: ltr; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("🔄 محول الملفات (PDF - Word - Excel - صور)")
st.caption("معالجة النصوص العربية، الحفاظ على الأرقام والنسب المئوية، تنظيم الأعمدة، ومعاينة النتيجة بعد التحويل")

ARABIC_CHAR_RE = re.compile(r"[\u0600-\u06FF]")
PURE_NUMBER_RE = re.compile(r"^[+-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?%?$|^[+-]?\d+(?:\.\d+)?%?$")
HEADER_KEYWORDS = {
    "الصنف",
    "الاصنف",
    "الكمية",
    "الكميه",
    "السعر",
    "الاجمالي",
    "الإجمالي",
    "المجموع",
    "نسبة",
    "المبيعات",
    "البيان",
    "الوصف",
}

# قاموس تصحيح العبارات والمقاطع المشوهة الناتجة عن ترميز/عكس الخطوط العربية
GARBLED_CORRECTIONS = {
    r"\bم\s+ان\b": "ميزان",
    r"\bطراب\s+ة\b": "ترابيزة",
    r"\bد\s+جتال\b": "ديجتال",
    r"\bصاعق\b": "صاعق",
    r"\bتكي\b": "تكييف",
}


def save_uploaded_file(uploaded_file):
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path, temp_dir


def clean_arabic_spacing(text):
    s = re.sub(r"\s+", " ", text).strip()
    s = re.sub(r"\s+([ىي])\s*$", r"\1", s)
    s = re.sub(r"^([ىي])\s+", r"\1", s)
    s = re.sub(r"\bال\s+", "ال", s)
    return re.sub(r"\s+", " ", s).strip()


def reverse_embedded_numbers(text):
    def _rev_num(match):
        num = match.group(0)
        if len(num) >= 2 and not num.endswith("%"):
            return num[::-1]
        return num

    return re.sub(r"\b\d+(?:\.\d+)?%?\b", _rev_num, text)


def apply_garbled_dictionary(text):
    s = text
    for pattern, replacement in GARBLED_CORRECTIONS.items():
        s = re.sub(pattern, replacement, s)
    return clean_arabic_spacing(s)


def sanitize_garbled_arabic(text):
    s = re.sub(r"\bdic\b|\(\)|:\d+|:\s*dic", " ", text, flags=re.IGNORECASE)
    s = re.sub(r"[()\[\]{}:;]+", " ", s)
    s = re.sub(r"([\u0600-\u06FF])(\d{2,})", r"\1 \2", s)
    s = re.sub(r"(\d{2,})([\u0600-\u06FF])", r"\1 \2", s)
    s = reverse_embedded_numbers(s)

    # إزالة سلاسل الأرقام العشوائية (4 أرقام أو أكثر) الملاصقة لنصوص وصفية مشوهة
    s = re.sub(r"(?:\s+|^)\d{4,}(?=\s|$)", " ", s)
    # إزالة الحروف العربية المفردة المبعثرة وسط الأرقام
    s = re.sub(r"(?<=\d)\s+[\u0600-\u06FF]\s+(?=\d)", " ", s)
    # إزالة مقاطع قصيرة جداً في النهاية فقط إذا كانت منفصلة بمسافة بعد كلمة كاملة (مع الحفاظ على الكلمات المستقلة القصيرة)
    s = re.sub(r"(?<=[\u0600-\u06FF]{3})\s+[\u0600-\u06FF]{1,2}$", "", s)

    s = clean_arabic_spacing(s)
    return apply_garbled_dictionary(s)


def fix_arabic_text(val):
    if val is None:
        return ""
    text = str(val).replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return ""

    if PURE_NUMBER_RE.match(text):
        return text

    if ARABIC_CHAR_RE.search(text):
        reversed_text = text[::-1]
        return sanitize_garbled_arabic(reversed_text)

    return sanitize_garbled_arabic(text)


def is_scattered_letters_column(col_values):
    non_empty = [str(v).strip() for v in col_values if str(v).strip()]
    if not non_empty:
        return True

    scattered_count = 0
    for val in non_empty:
        tokens = val.split()
        if len(tokens) >= 4 and sum(1 for t in tokens if len(t) == 1) / len(tokens) >= 0.65:
            scattered_count += 1

    return scattered_count / len(non_empty) >= 0.6


def parse_numeric_if_possible(val):
    if not isinstance(val, str):
        return val
    s = val.strip()
    if not s or s.endswith("%"):
        return s
    cleaned_num = s.replace(",", "")
    if re.fullmatch(r"[+-]?\d+", cleaned_num):
        try:
            return int(cleaned_num)
        except ValueError:
            return s
    if re.fullmatch(r"[+-]?\d+\.\d+", cleaned_num):
        try:
            return float(cleaned_num)
        except ValueError:
            return s
    return s


def find_header_row_index(rows):
    best_idx = 0
    best_score = -1
    limit = min(len(rows), 5)
    for idx in range(limit):
        row = rows[idx]
        score = 0
        for cell in row:
            c_clean = re.sub(r"[^\u0600-\u06FF\s]", "", str(cell)).strip()
            for word in c_clean.split():
                if word in HEADER_KEYWORDS:
                    score += 2
            if c_clean in HEADER_KEYWORDS:
                score += 2
        if score > best_score:
            best_score = score
            best_idx = idx
    return best_idx if best_score >= 2 else 0


def clean_and_organize_df(raw_rows):
    cleaned_rows = []
    for row in raw_rows:
        fixed_row = [fix_arabic_text(c) for c in row]
        if any(cell != "" for cell in fixed_row):
            cleaned_rows.append(fixed_row)

    if len(cleaned_rows) < 2:
        return None

    max_cols = max(len(r) for r in cleaned_rows)
    padded = [r + [""] * (max_cols - len(r)) for r in cleaned_rows]
    padded = [r[::-1] for r in padded]

    header_idx = find_header_row_index(padded)
    raw_headers = padded[header_idx]
    body_rows = padded[header_idx + 1 :]

    if not body_rows:
        return None

    keep_col_indices = []
    for c_idx in range(max_cols):
        col_vals = [raw_headers[c_idx]] + [r[c_idx] for r in body_rows]
        if not is_scattered_letters_column(col_vals):
            keep_col_indices.append(c_idx)

    if not keep_col_indices:
        return None

    filtered_headers = [raw_headers[i] for i in keep_col_indices]
    filtered_body = [[r[i] for i in keep_col_indices] for r in body_rows]

    filtered_body = [r for r in filtered_body if any(str(cell).strip() != "" for cell in r)]
    if not filtered_body:
        return None

    headers = [h if h else f"عمود_{i + 1}" for i, h in enumerate(filtered_headers)]

    seen = {}
    unique_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)

    data_rows = [[parse_numeric_if_possible(c) for c in r] for r in filtered_body]
    return pd.DataFrame(data_rows, columns=unique_headers)


def pdf_tables_to_dataframes(pdf_path):
    dfs = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables(
                {
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                }
            )
            if not tables:
                tables = page.extract_tables(
                    {
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                    }
                )

            for t_idx, table in enumerate(tables, start=1):
                df = clean_and_organize_df(table)
                if df is not None and not df.empty:
                    sheet_name = f"P{page_num}_T{t_idx}"[:31]
                    dfs.append((sheet_name, df))
    return dfs


def ocr_image_to_lines(img):
    gray = ImageOps.grayscale(img)
    gray = ImageOps.autocontrast(gray)
    raw_text = pytesseract.image_to_string(gray, lang="ara+eng", config="--psm 6")
    lines = []
    for line in raw_text.splitlines():
        cleaned = re.sub(r"\s+", " ", line).strip()
        if cleaned:
            lines.append(cleaned)
    return lines


def convert_office_to_pdf_libreoffice(input_path, temp_dir):
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", input_path, "--outdir", temp_dir],
        check=True,
    )
    base = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(temp_dir, f"{base}.pdf")


def preview_converted_file(file_path, dfs=None, text_lines=None):
    st.divider()
    st.subheader("👁️ معاينة الملف بعد التحويل")

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".xlsx":
        if dfs:
            sheet_names = [name for name, _ in dfs]
            selected_sheet = st.selectbox("اختر الشيت للمعاينة:", sheet_names, key="preview_sheet_select")
            sheet_dict = dict(dfs)
            preview_df = sheet_dict[selected_sheet]
            st.caption(f"عدد الصفوف: {len(preview_df)} | عدد الأعمدة: {len(preview_df.columns)}")
            st.dataframe(preview_df, use_container_width=True)
        else:
            excel_data = pd.ExcelFile(file_path)
            selected_sheet = st.selectbox("اختر الشيت للمعاينة:", excel_data.sheet_names, key="preview_xlsx_select")
            preview_df = excel_data.parse(selected_sheet)
            st.dataframe(preview_df, use_container_width=True)

    elif ext == ".docx":
        doc = Document(file_path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        tables_count = len(doc.tables)
        st.caption(f"عدد الفقرات: {len(paragraphs)} | عدد الجداول: {tables_count}")
        if paragraphs:
            st.text_area("معاينة النصوص داخل ملف Word:", "\n".join(paragraphs[:80]), height=260)
        if tables_count > 0:
            first_tbl = [[c.text.strip() for c in row.cells] for row in doc.tables[0].rows]
            if len(first_tbl) >= 2:
                st.caption("معاينة أول جدول في ملف Word:")
                st.dataframe(pd.DataFrame(first_tbl[1:], columns=first_tbl[0]), use_container_width=True)

    elif ext == ".pdf":
        with pdfplumber.open(file_path) as pdf:
            total_pages = len(pdf.pages)
            st.caption(f"إجمالي الصفحات في ملف PDF الناتج: {total_pages}")
            page_num = st.number_input(
                "رقم الصفحة للمعاينة:",
                min_value=1,
                max_value=max(1, total_pages),
                value=1,
                step=1,
            )
            page = pdf.pages[page_num - 1]
            page_img = page.to_image(resolution=150).original
            st.image(page_img, caption=f"صفحة {page_num} من {total_pages}", use_container_width=True)

    if text_lines:
        with st.expander("عرض الأسطر المستخرجة (OCR)"):
            st.write("\n".join(text_lines))


uploaded_file = st.file_uploader(
    "اختر الملف",
    type=["pdf", "docx", "xlsx", "xls", "png", "jpg", "jpeg"],
)

if uploaded_file is not None:
    file_ext = uploaded_file.name.split(".")[-1].lower()
    input_path, temp_dir = save_uploaded_file(uploaded_file)
    base_name = os.path.splitext(uploaded_file.name)[0]

    if file_ext == "pdf":
        options = ["Excel (.xlsx) - استخراج وتنظيم الجداول", "Word (.docx)"]
    elif file_ext in ["docx", "xlsx", "xls"]:
        options = ["PDF (.pdf)"]
    else:
        options = [
            "Excel (.xlsx) - جدول منظم",
            "Word (.docx) - نص منظم",
            "PDF (.pdf)",
        ]

    target_format = st.selectbox("اختر الصيغة المطلوبة:", options)

    if st.button("تحويل الملف 🚀"):
        with st.spinner("جاري التحويل وتنظيم البيانات..."):
            output_path = ""
            mime_type = ""
            extracted_dfs = None
            extracted_lines = None

            try:
                if file_ext == "pdf":
                    if "Excel" in target_format:
                        dfs = pdf_tables_to_dataframes(input_path)
                        if not dfs:
                            st.warning("لم يتم التعرف على جداول واضحة داخل ملف الـ PDF.")
                            st.stop()
                        output_path = os.path.join(temp_dir, f"{base_name}_organized.xlsx")
                        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                            for sheet_name, df in dfs:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                        extracted_dfs = dfs
                        mime_type = XLSX_MIME

                    elif "Word" in target_format:
                        output_path = os.path.join(temp_dir, f"{base_name}.docx")
                        cv = Converter(input_path)
                        cv.convert(output_path)
                        cv.close()
                        mime_type = DOCX_MIME

                elif file_ext in ["docx", "xlsx", "xls"]:
                    output_path = convert_office_to_pdf_libreoffice(input_path, temp_dir)
                    mime_type = PDF_MIME

                elif file_ext in ["png", "jpg", "jpeg"]:
                    img = Image.open(input_path)
                    if "PDF" in target_format:
                        output_path = os.path.join(temp_dir, f"{base_name}.pdf")
                        img.convert("RGB").save(output_path)
                        mime_type = PDF_MIME

                    elif "Word" in target_format:
                        lines = ocr_image_to_lines(img)
                        fixed_lines = [fix_arabic_text(line) for line in lines]
                        extracted_lines = fixed_lines
                        output_path = os.path.join(temp_dir, f"{base_name}_ocr.docx")
                        doc = Document()
                        for line in fixed_lines:
                            p = doc.add_paragraph(line)
                            p.paragraph_format.bidi_level = 1
                        doc.save(output_path)
                        mime_type = DOCX_MIME

                    elif "Excel" in target_format:
                        lines = ocr_image_to_lines(img)
                        extracted_lines = [fix_arabic_text(line) for line in lines]
                        output_path = os.path.join(temp_dir, f"{base_name}_ocr.xlsx")
                        rows = []
                        for line in lines:
                            cols = [c.strip() for c in re.split(r"\s{2,}|\t", line) if c.strip()]
                            rows.append(cols if cols else [line])
                        max_c = max((len(r) for r in rows), default=1)
                        padded = [r + [""] * (max_c - len(r)) for r in rows]
                        headers = [f"عمود_{i + 1}" for i in range(max_c)]
                        rows_cleaned = [
                            [parse_numeric_if_possible(fix_arabic_text(c)) for c in r]
                            for r in padded
                        ]
                        df = pd.DataFrame(rows_cleaned, columns=headers)
                        df.to_excel(output_path, index=False)
                        extracted_dfs = [("OCR_Sheet", df)]
                        mime_type = XLSX_MIME

                if output_path and os.path.exists(output_path):
                    st.success("تم التحويل وتنظيم البيانات بنجاح! ✅")
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 تحميل الملف بعد التنظيم",
                            data=f,
                            file_name=os.path.basename(output_path),
                            mime=mime_type,
                        )

                    preview_converted_file(
                        output_path,
                        dfs=extracted_dfs,
                        text_lines=extracted_lines,
                    )

            except FileNotFoundError:
                st.error("تأكد من تثبيت Tesseract-OCR أو LibreOffice على الجهاز حسب نوع التحويل.")
            except Exception as e:
                st.error(f"حدث خطأ: {e}")
