from fastapi import FastAPI, UploadFile, File, Form, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
import openpyxl
from openpyxl.styles import Font
from copy import copy
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re, io, zipfile, traceback, os, base64, gc, tempfile, json
from datetime import datetime
import pypdf

app = FastAPI(title="QT1 Doc Control API - Dual Extraction Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Processed-Items"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def clean_safe_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/*?:"<>|]', '-', name).strip()
    return cleaned[:130] if len(cleaned) > 130 else cleaned

def robust_extract_date(full_text: str) -> str:
    """Hàm bóc tách ngày chuyên dụng xử lý toàn bộ các định dạng tiếng Anh / tiếng Việt"""
    month_map = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'september': 9, 'sept': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12
    }
    month_pattern = r'(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)'
    
    # 1. Tìm trực tiếp chữ Date kèm ngày tháng (ví dụ Date: October 19, 2026)
    m_block1 = re.search(r'\bDate\b[^\w\n\r]*[:\s]*[^\w\n\r]*(' + month_pattern + r'\s+\d{1,2}(?:st|nd|rd|th)?,?\s+202\d)', full_text, re.IGNORECASE)
    if m_block1:
        sub_m = re.search(r'(' + month_pattern + r')\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(202\d)', m_block1.group(1), re.IGNORECASE)
        if sub_m:
            m_str = sub_m.group(1).lower()
            d_val = int(sub_m.group(2))
            y_val = sub_m.group(3)
            if m_str in month_map:
                return f"{d_val:02d}/{month_map[m_str]:02d}/{y_val}"

    # 1b. Dạng Date: 19 October 2026
    m_block2 = re.search(r'\bDate\b[^\w\n\r]*[:\s]*[^\w\n\r]*(\d{1,2}(?:st|nd|rd|th)?\s+' + month_pattern + r',?\s+202\d)', full_text, re.IGNORECASE)
    if m_block2:
        sub_m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+(' + month_pattern + r'),?\s+(202\d)', m_block2.group(1), re.IGNORECASE)
        if sub_m:
            d_val = int(sub_m.group(1))
            m_str = sub_m.group(2).lower()
            y_val = sub_m.group(3)
            if m_str in month_map:
                return f"{d_val:02d}/{month_map[m_str]:02d}/{y_val}"

    # 1c. Dạng Date: 19/10/2026
    m_block3 = re.search(r'\bDate\b[^\w\n\r]*[:\s]*[^\w\n\r]*(\d{1,2})[\/.-](\d{1,2})[\/.-](202\d)', full_text, re.IGNORECASE)
    if m_block3:
        d_val = int(m_block3.group(1))
        m_val = int(m_block3.group(2))
        y_val = m_block3.group(3)
        if 1 <= m_val <= 12 and 1 <= d_val <= 31:
            return f"{d_val:02d}/{m_val:02d}/{y_val}"

    # 2. Tìm tất cả các ngày trong 3000 ký tự đầu tiên, ưu tiên ngày gần mã QT1 nhất
    ref_pos = -1
    m_ref = re.search(r'QT1-[A-Za-z0-9_-]+', full_text)
    if m_ref:
        ref_pos = m_ref.start()

    candidates = []
    scope_text = full_text[:3000]

    for m in re.finditer(r'\b(' + month_pattern + r')\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(202\d)\b', scope_text, re.IGNORECASE):
        m_str = m.group(1).lower()
        if m_str in month_map:
            d_val = int(m.group(2))
            y_val = m.group(3)
            dist = abs(m.start() - ref_pos) if ref_pos >= 0 else m.start()
            candidates.append((dist, f"{d_val:02d}/{month_map[m_str]:02d}/{y_val}"))

    for m in re.finditer(r'\b(\d{1,2})(?:st|nd|rd|th)?\s+(' + month_pattern + r'),?\s+(202\d)\b', scope_text, re.IGNORECASE):
        m_str = m.group(2).lower()
        if m_str in month_map:
            d_val = int(m.group(1))
            y_val = m.group(3)
            dist = abs(m.start() - ref_pos) if ref_pos >= 0 else m.start()
            candidates.append((dist, f"{d_val:02d}/{month_map[m_str]:02d}/{y_val}"))

    for m in re.finditer(r'\b(\d{1,2})[\/.-](\d{1,2})[\/.-](202\d)\b', scope_text):
        d_val = int(m.group(1))
        m_val = int(m.group(2))
        y_val = m.group(3)
        if 1 <= m_val <= 12 and 1 <= d_val <= 31:
            dist = abs(m.start() - ref_pos) if ref_pos >= 0 else m.start()
            candidates.append((dist, f"{d_val:02d}/{m_val:02d}/{y_val}"))

    if candidates:
        candidates.sort(key=lambda x: x[0])
        return candidates[0][1]

    return ""

def extract_info_from_filename(filename: str):
    clean_name = os.path.splitext(filename or "")[0].strip()
    m = re.search(r'(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)', clean_name)
    if not m:
        m = re.search(r'(QT1-[A-Za-z0-9_-]+)', clean_name)
    
    doc_no = m.group(1).strip() if m else ""
    subject = ""
    if doc_no:
        parts = clean_name.split(doc_no, 1)
        if len(parts) > 1 and parts[1].strip('_- '):
            subject = parts[1].strip('_- ')
            subject = re.sub(r'(_|-)?Cover$', '', subject, flags=re.IGNORECASE).strip()
    return doc_no, subject

def extract_clean_doc_no(text: str) -> str:
    m = re.search(r'(QT1-[0-9A-Z]+-[A-Z0-9]+-[A-Z0-9]+-\d{2}-\d{5})', text)
    if m: return m.group(1)
    m2 = re.search(r'(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)', text)
    if m2:
        base = m2.group(1)
        base = re.sub(r'_[A-Za-z0-9]+$', '', base)
        base = re.sub(r'_?Rev\.?.*$', '', base, flags=re.IGNORECASE)
        return base
    m3 = re.search(r'(QT1-[A-Za-z0-9_-]+)', text)
    if m3:
        base = m3.group(1)
        base = re.sub(r'_[A-Za-z0-9]+$', '', base)
        base = re.sub(r'_?Rev\.?.*$', '', base, flags=re.IGNORECASE)
        return base
    return ""

def get_latest_status_from_ws(ws, max_r):
    target_col = None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(row=11, column=c).value or "").upper()
        if f"ROUND {max_r} COMMENTS" in val or f"ROUND {max_r} COMMENT" in val:
            target_col = c
            break
            
    if target_col:
        status_val = str(ws.cell(row=13, column=target_col).value or "").upper()
        if "APP" in status_val and "AWC" not in status_val: return "APP"
        elif "RFC" in status_val: return "RFC"
        elif "AWC" in status_val: return "AWC"
        elif "RJ" in status_val: return "RJ"
    return "APP"

def parse_cover_pdf_bytes(pdf_bytes, filename=""):
    fn_doc_no, fn_subject = extract_info_from_filename(filename)
    
    metadata = {
        'tr_no': fn_doc_no if fn_doc_no else 'QT1-Letter',
        'date_str': datetime.now().strftime("%d/%m/%Y"), 
        'subject': fn_subject if fn_subject else 'Official Transmittal',
        'docs': {}
    }
    
    p1_txt = ""
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        if reader.pages:
            p1_txt = reader.pages[0].extract_text() or ""
    except Exception as e:
        print(f"Lỗi đọc PDF pypdf: {e}")

    if p1_txt:
        # 1. Tìm số công văn / số đệ trình
        ref_patterns = [
            r'(?:Ref\.?\s*(?:No\.?)?|Letter\s*(?:No\.?)?|TR\.?\s*(?:No\.?)?|Transmittal\s*(?:No\.?)?|No\.?)\s*[:\s]*([A-Za-z0-9_-]*QT1[A-Za-z0-9_-]+)',
            r'(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-\d+)',
            r'(QT1-[A-Za-z0-9_-]+)'
        ]
        for pat in ref_patterns:
            tr_m = re.search(pat, p1_txt, re.IGNORECASE)
            if tr_m:
                found_ref = tr_m.group(1).strip()
                if len(found_ref) >= 8:
                    metadata['tr_no'] = found_ref
                    break
        
        # 2. Tìm ngày tháng bằng hàm quét chuyên dụng
        extracted_date = robust_extract_date(p1_txt)
        if extracted_date:
            metadata['date_str'] = extracted_date

        # 3. Tìm chủ đề (Subject)
        subj_m = re.search(r'(?:SUBJECT|Subject|Sub|V/v)\s*[:\s]+([^\n\r]+)', p1_txt)
        if subj_m:
            raw_subj = subj_m.group(1).strip()
            raw_subj = re.split(r'(?:Document/Drawing|Attachment|Attn:|Dear)', raw_subj, flags=re.IGNORECASE)[0].strip()
            if len(raw_subj) >= 5:
                metadata['subject'] = raw_subj

    # Ưu tiên dữ liệu từ tên file nếu đọc trong PDF bị thiếu
    if (not metadata['tr_no'] or metadata['tr_no'] == 'QT1-Letter') and fn_doc_no:
        metadata['tr_no'] = fn_doc_no
    if (not metadata['subject'] or metadata['subject'] == 'Official Transmittal') and fn_subject:
        metadata['subject'] = fn_subject

    return metadata

def parse_cover_excel_bytes(excel_bytes, filename=""):
    fn_doc_no, fn_subject = extract_info_from_filename(filename)
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    metadata = {
        'tr_no': fn_doc_no if fn_doc_no else 'QT1-Document',
        'date_str': datetime.now().strftime("%d/%m/%Y"),
        'subject': fn_subject if fn_subject else 'Official Transmittal',
        'docs': {}
    }
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            m = re.match(r'^(QT1-[0-9L]-[A-Z0-9_-]+)$', val)
            if m and 'QT1-HDH' not in val:
                doc_no = extract_clean_doc_no(m.group(1))
                if doc_no:
                    rev_val = str(ws.cell(row=r, column=c+1).value or "1").strip()
                    title_val = str(ws.cell(row=r, column=c+2).value or "").strip()
                    metadata['docs'][doc_no] = {
                        'doc_no': doc_no,
                        'rev': rev_val,
                        'title': title_val
                    }
    return metadata

def find_template(prefix):
    search_dirs = [BASE_DIR, os.getcwd(), r"G:\Shared drives\Test auto"]
    for d in search_dirs:
        if not os.path.exists(d): continue
        for f in os.listdir(d):
            if f.lower().endswith(".docx"):
                clean_f = f.lower().replace("đ", "d").replace("ð", "d")
                clean_p = prefix.lower().replace("đ", "d").replace("ð", "d")
                if clean_p in clean_f and not f.startswith("~$"):
                    return os.path.join(d, f)
    return None

def replace_value_after_keyword(paragraph, keyword, new_text):
    if keyword in paragraph.text:
        if not paragraph.runs: return False
        kw_font = paragraph.runs[0].font
        kw_name, kw_size, kw_bold, kw_italic = kw_font.name, kw_font.size, kw_font.bold, kw_font.italic
        val_font = paragraph.runs[-1].font
        val_name, val_size, val_bold, val_italic = val_font.name, val_font.size, val_font.bold, val_font.italic
        
        paragraph.clear()
        r1 = paragraph.add_run(keyword)
        r1.font.name, r1.font.size, r1.font.bold, r1.font.italic = kw_name, kw_size, kw_bold, kw_italic
        r2 = paragraph.add_run(" " + new_text)
        r2.font.name, r2.font.size, r2.font.bold, r2.font.italic = val_name, val_size, val_bold, val_italic
        return True
    return False

def update_date_paragraph(paragraph, new_date_str):
    spaces = "                                "
    val_font = paragraph.runs[-1].font
    val_name, val_size, val_bold, val_italic = val_font.name, val_font.size, val_font.bold, val_font.italic
    
    paragraph.clear()
    r1 = paragraph.add_run(spaces)
    r2 = paragraph.add_run(new_date_str)
    r2.font.name, r2.font.size, r2.font.bold, r2.font.italic = val_name, val_size, val_bold, val_italic

def sync_footer_content(doc, qti_no, subject):
    new_text = f"{qti_no} {subject}"
    for section in doc.sections:
        footer = section.footer
        if footer:
            for p in footer.paragraphs:
                if p.text.strip():
                    if len(p.runs) > 0:
                        first_font = p.runs[0].font
                        font_name = first_font.name
                        font_size = first_font.size
                        font_color = first_font.color.rgb if first_font.color else None
                        
                        p.clear()
                        r = p.add_run(new_text)
                        r.font.name = font_name
                        r.font.size = font_size
                        if font_color:
                            r.font.color.rgb = font_color
                        else:
                            r.font.color.rgb = RGBColor(128, 128, 128)
                            r.font.size = Pt(10)
                    else:
                        p.text = new_text

def remove_temp_file(path: str):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Lỗi dọn dẹp file tạm {path}: {e}")

@app.post("/api/export-zip-full")
async def export_zip_full(
    files: list[UploadFile] = File(default=[]),
    cover_file: UploadFile = File(None),
    metadata: str = Form("{}"),
    mode: str = Form("auto_bump"),
    qti_no: str = Form("QTI-PECC1-PMB2-26-10-111"),
    start_car_no: str = Form("9999")
):
    temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    temp_zip_path = temp_zip.name
    temp_zip.close()

    try:
        car_metadata = json.loads(metadata) if metadata else {}
        now = datetime.now()
        current_date_str = now.strftime("%d %b %Y")
        current_date_en = now.strftime("%B %d, %Y")
        
        cover_data = {
            'tr_no': 'QT1-Document',
            'date_str': now.strftime("%d/%m/%Y"),
            'subject': 'Official Transmittal',
            'docs': {}
        }
        
        if cover_file is not None:
            cover_bytes = await cover_file.read()
            if cover_file.filename.lower().endswith('.pdf'):
                cover_data = parse_cover_pdf_bytes(cover_bytes, cover_file.filename)
            elif cover_file.filename.lower().endswith(('.xlsx', '.xls')):
                cover_data = parse_cover_excel_bytes(cover_bytes, cover_file.filename)
            del cover_bytes

        vi_template_path = find_template("TV")
        if not vi_template_path: raise FileNotFoundError(f"Vui lòng copy file Word mẫu 'TVÐ1-NÐNLM.docx' vào thư mục: {BASE_DIR}")
        en_template_path = find_template("QTI-PECC1")
        if not en_template_path: raise FileNotFoundError(f"Vui lòng copy file Word mẫu 'QTI-PECC1...docx' vào thư mục: {BASE_DIR}")

        processed_car_list = []

        try:
            start_num = int(re.sub(r'\D', '', str(start_car_no)) or "9999")
        except Exception:
            start_num = 9999

        with zipfile.ZipFile(temp_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            
            # XỬ LÝ CÁC CHẾ ĐỘ CÓ FILE CAR
            if mode != "no_car_letter" and files:
                for file_idx, file in enumerate(files):
                    try:
                        content = await file.read()

                        # 1. AUTO BUMP
                        if mode == "auto_bump":
                            wb = openpyxl.load_workbook(io.BytesIO(content))
                            max_r = 1
                            target_sheet_name = wb.sheetnames[0]
                            target_idx = 0
                            for idx, s_name in enumerate(wb.sheetnames):
                                m = re.search(r'round\s*(\d+)', s_name, re.IGNORECASE)
                                if m and int(m.group(1)) >= max_r:
                                    max_r = int(m.group(1))
                                    target_sheet_name = s_name
                                    target_idx = idx
                                    
                            source_ws = wb[target_sheet_name]
                            car_p = re.search(r'(\d{4,}\s*CaRs)', file.filename, re.IGNORECASE)
                            car_base = car_p.group(1) if car_p else "CaRs"

                            selected_status = car_metadata.get(file.filename, "APP")
                            new_r_num = max_r + 1
                            new_ws = wb.copy_worksheet(source_ws)
                            new_ws.title = f"ROUND {new_r_num} CaRs"
                            wb._sheets.remove(new_ws)
                            wb._sheets.insert(target_idx + 1, new_ws)
                            
                            detected_doc_code = extract_clean_doc_no(file.filename)
                            detected_title = ""
                            for c in range(1, source_ws.max_column + 1):
                                r4_val = str(source_ws.cell(row=4, column=c).value or "")
                                extracted = extract_clean_doc_no(r4_val)
                                if extracted: detected_doc_code = extracted
                                
                                r5_val = str(source_ws.cell(row=5, column=c).value or "").strip()
                                if r5_val and not detected_title:
                                    detected_title = re.sub(r'_(IFA|APD|IFI|RI|HDEC|Reply).*$', '', r5_val, flags=re.IGNORECASE).strip()

                            cover_info = cover_data['docs'].get(detected_doc_code, {})
                            final_doc_no = cover_info.get('doc_no', detected_doc_code if detected_doc_code else "QT1-Document")
                            final_rev = cover_info.get('rev', '1')
                            final_title = cover_info.get('title', detected_title if detected_title else "P&I Diagram")

                            for c in range(1, 10):
                                v = str(new_ws.cell(row=1, column=c).value or "")
                                if re.search(r'-\s*R\d+', v, re.IGNORECASE) or (c == 5 and v):
                                    new_ws.cell(row=1, column=c).value = re.sub(r'R\d+', f'R{new_r_num}', v, flags=re.IGNORECASE)
                                    break

                            prev_com_col, prev_rep_col = None, None
                            for c in range(1, new_ws.max_column + 1):
                                val_h = str(new_ws.cell(row=11, column=c).value or "").upper()
                                if f"ROUND {max_r} COMMENTS" in val_h: prev_com_col = c
                                elif f"ROUND {max_r} REPLIES" in val_h: prev_rep_col = c

                            new_com_col = (prev_rep_col + 1) if prev_rep_col else (new_ws.max_column + 1)
                            new_rep_col = new_com_col + 2

                            new_ws.cell(row=2, column=new_com_col).value = "EVNPMB2"
                            new_ws.cell(row=3, column=new_com_col).value = cover_data['tr_no']
                            new_ws.cell(row=4, column=new_com_col).value = f"{final_doc_no} rev.{final_rev}"
                            new_ws.cell(row=5, column=new_com_col).value = final_title
                            new_ws.cell(row=10, column=new_rep_col).value = "CONTRACTOR's"
                            new_ws.cell(row=11, column=new_com_col).value = f"ROUND {new_r_num} COMMENTS"
                            new_ws.cell(row=11, column=new_rep_col).value = f"ROUND {new_r_num} REPLIES"
                            
                            for r in range(1, 15):
                                if prev_com_col:
                                    src_c = new_ws.cell(row=r, column=prev_com_col)
                                    tgt_c = new_ws.cell(row=r, column=new_com_col)
                                    if src_c.has_style:
                                        tgt_c.font, tgt_c.border, tgt_c.fill, tgt_c.alignment = copy(src_c.font), copy(src_c.border), copy(src_c.fill), copy(src_c.alignment)
                                if prev_rep_col:
                                    src_r = new_ws.cell(row=r, column=prev_rep_col)
                                    tgt_r = new_ws.cell(row=r, column=new_rep_col)
                                    if src_r.has_style:
                                        tgt_r.font, tgt_r.border, tgt_r.fill, tgt_r.alignment = copy(src_r.font), copy(src_r.border), copy(src_r.fill), copy(src_r.alignment)

                            status_full_text = "APPROVED (APP)" if selected_status == "APP" else "RETURN FOR CORRECTIONS (RFC)"
                            if selected_status == "AWC": status_full_text = "APPROVED WITH COMMENTS (AWC)"
                            elif selected_status == "RJ": status_full_text = "REJECTED (RJ)"
                            
                            new_ws.cell(row=12, column=new_com_col).value = f"STATUS AFTER ROUND {new_r_num} COMMENTS:"
                            new_ws.cell(row=13, column=new_com_col).value = status_full_text

                            for r in range(1, 15):
                                for c in range(1, new_ws.max_column + 1):
                                    cell = new_ws.cell(row=r, column=c)
                                    if cell.value and cell.font:
                                        f = cell.font
                                        color = f.color if f.color else None
                                        cell.font = Font(name=f.name, size=f.sz, bold=True, italic=f.i, color=color)

                            red_bold = Font(name="Arial", size=10, bold=True, color="FFFF0000")
                            new_ws.cell(row=12, column=new_com_col).font = red_bold
                            new_ws.cell(row=13, column=new_com_col).font = red_bold

                            for r in range(14, new_ws.max_row + 1):
                                v_prev_c = str(new_ws.cell(row=r, column=prev_com_col).value or "").strip() if prev_com_col else ""
                                v_prev_r = str(new_ws.cell(row=r, column=prev_rep_col).value or "").strip() if prev_rep_col else ""
                                
                                if prev_com_col:
                                    src_c = new_ws.cell(row=r, column=prev_com_col)
                                    tgt_c = new_ws.cell(row=r, column=new_com_col)
                                    if src_c.has_style:
                                        tgt_c.font, tgt_c.border, tgt_c.alignment = copy(src_c.font), copy(src_c.border), copy(src_c.fill), copy(src_c.alignment)

                                is_group = "general comment" in v_prev_c.lower() or "detailed comment" in v_prev_c.lower()
                                if is_group:
                                    new_ws.cell(row=r, column=new_com_col).value = None
                                elif v_prev_c or v_prev_r:
                                    if selected_status == "APP":
                                        new_ws.cell(row=r, column=new_com_col).value = "Noted."
                                else:
                                    new_ws.cell(row=r, column=new_com_col).value = None

                            new_ws.column_dimensions[openpyxl.utils.get_column_letter(new_com_col)].width = 38
                            new_ws.column_dimensions[openpyxl.utils.get_column_letter(new_rep_col)].width = 38

                            safe_title_for_name = clean_safe_filename(final_title)
                            new_excel_name = f"{car_base} R{new_r_num} {current_date_str} {final_doc_no} {safe_title_for_name} rev.{final_rev}.xlsx"

                            excel_out = io.BytesIO()
                            wb.save(excel_out)
                            zip_file.writestr(new_excel_name, excel_out.getvalue())
                            del wb, excel_out

                            processed_car_list.append({
                                'doc_no': final_doc_no,
                                'rev': final_rev,
                                'title': final_title,
                                'status': status_full_text,
                                'car_label': f"{car_base} R{new_r_num}",
                                'assigned_car_name': new_excel_name
                            })

                        # 2. NUMBER FILLED CARS
                        elif mode == "number_filled_cars":
                            cur_num = start_num + file_idx
                            car_base = f"{cur_num:04d} CaRs" if cur_num < 10000 else f"{cur_num} CaRs"
                            
                            wb = openpyxl.load_workbook(io.BytesIO(content))
                            max_r = 1
                            target_ws = wb.active
                            for s_name in wb.sheetnames:
                                m = re.search(r'round\s*(\d+)', s_name, re.IGNORECASE)
                                if m and int(m.group(1)) >= max_r:
                                    max_r = int(m.group(1))
                                    target_ws = wb[s_name]

                            car_label = f"{car_base} R{max_r}"

                            detected_doc_code = extract_clean_doc_no(file.filename)
                            detected_title = ""
                            for c in range(1, target_ws.max_column + 1):
                                r4_val = str(target_ws.cell(row=4, column=c).value or "")
                                extracted = extract_clean_doc_no(r4_val)
                                if extracted: detected_doc_code = extracted
                                r5_val = str(target_ws.cell(row=5, column=c).value or "").strip()
                                if r5_val and not detected_title:
                                    detected_title = re.sub(r'_(IFA|APD|IFI|RI|HDEC|Reply).*$', '', r5_val, flags=re.IGNORECASE).strip()

                            doc_keys = list(cover_data['docs'].keys())
                            if (not detected_doc_code or detected_doc_code not in cover_data['docs']) and file_idx < len(doc_keys):
                                cover_info = cover_data['docs'][doc_keys[file_idx]]
                                final_doc_no = cover_info['doc_no']
                                final_rev = cover_info['rev']
                                final_title = cover_info['title']
                            else:
                                cover_info = cover_data['docs'].get(detected_doc_code, {})
                                final_doc_no = cover_info.get('doc_no', detected_doc_code if detected_doc_code else f"QT1-Doc-{cur_num}")
                                final_rev = cover_info.get('rev', '1')
                                final_title = cover_info.get('title', detected_title if detected_title else "Document")

                            for c in range(1, 10):
                                v = str(target_ws.cell(row=1, column=c).value or "")
                                if re.search(r'CaRs?', v, re.IGNORECASE) or (c == 5 and v):
                                    target_ws.cell(row=1, column=c).value = f"{car_base} - R{max_r}"
                                    break

                            com_col = None
                            for c in range(1, target_ws.max_column + 1):
                                val_h = str(target_ws.cell(row=11, column=c).value or "").upper()
                                if f"ROUND {max_r} COMMENTS" in val_h or f"ROUND {max_r} COMMENT" in val_h:
                                    com_col = c
                                    break

                            selected_status = car_metadata.get(file.filename, None)
                            if not selected_status:
                                selected_status = get_latest_status_from_ws(target_ws, max_r)

                            status_full_text = "APPROVED (APP)" if selected_status == "APP" else "RETURN FOR CORRECTIONS (RFC)"
                            if selected_status == "AWC": status_full_text = "APPROVED WITH COMMENTS (AWC)"
                            elif selected_status == "RJ": status_full_text = "REJECTED (RJ)"

                            if com_col:
                                target_ws.cell(row=2, column=com_col).value = "EVNPMB2"
                                target_ws.cell(row=3, column=com_col).value = cover_data['tr_no']
                                target_ws.cell(row=4, column=com_col).value = f"{final_doc_no} rev.{final_rev}"
                                target_ws.cell(row=5, column=com_col).value = final_title
                                target_ws.cell(row=12, column=com_col).value = f"STATUS AFTER ROUND {max_r} COMMENTS:"
                                target_ws.cell(row=13, column=com_col).value = status_full_text
                                
                                red_bold = Font(name="Arial", size=10, bold=True, color="FFFF0000")
                                target_ws.cell(row=12, column=com_col).font = red_bold
                                target_ws.cell(row=13, column=com_col).font = red_bold

                            safe_title_for_name = clean_safe_filename(final_title)
                            new_excel_name = f"{car_base} R{max_r} {current_date_str} {final_doc_no} {safe_title_for_name} rev.{final_rev}.xlsx"

                            excel_out = io.BytesIO()
                            wb.save(excel_out)
                            zip_file.writestr(new_excel_name, excel_out.getvalue())
                            del wb, excel_out

                            processed_car_list.append({
                                'doc_no': final_doc_no,
                                'rev': final_rev,
                                'title': final_title,
                                'status': status_full_text,
                                'car_label': car_label,
                                'assigned_car_name': new_excel_name
                            })

                        # 3. NEW CARS
                        elif mode == "new_cars":
                            cur_num = start_num + file_idx
                            car_base = f"{cur_num:04d} CaRs" if cur_num < 10000 else f"{cur_num} CaRs"
                            car_label = f"{car_base} R1"
                            selected_status = car_metadata.get(file.filename, "APP")

                            wb = openpyxl.load_workbook(io.BytesIO(content))
                            target_ws = wb.active
                            for s_name in wb.sheetnames:
                                if re.search(r'round\s*1', s_name, re.IGNORECASE):
                                    target_ws = wb[s_name]
                                    break
                            target_ws.title = "ROUND 1 CaRs"

                            detected_doc_code = extract_clean_doc_no(file.filename)
                            detected_title = ""
                            for c in range(1, target_ws.max_column + 1):
                                r4_val = str(target_ws.cell(row=4, column=c).value or "")
                                extracted = extract_clean_doc_no(r4_val)
                                if extracted: detected_doc_code = extracted
                                r5_val = str(target_ws.cell(row=5, column=c).value or "").strip()
                                if r5_val and not detected_title:
                                    detected_title = re.sub(r'_(IFA|APD|IFI|RI|HDEC|Reply).*$', '', r5_val, flags=re.IGNORECASE).strip()

                            doc_keys = list(cover_data['docs'].keys())
                            if (not detected_doc_code or detected_doc_code not in cover_data['docs']) and file_idx < len(doc_keys):
                                cover_info = cover_data['docs'][doc_keys[file_idx]]
                                final_doc_no = cover_info['doc_no']
                                final_rev = cover_info['rev']
                                final_title = cover_info['title']
                            else:
                                cover_info = cover_data['docs'].get(detected_doc_code, {})
                                final_doc_no = cover_info.get('doc_no', detected_doc_code if detected_doc_code else f"QT1-Doc-{cur_num}")
                                final_rev = cover_info.get('rev', '1')
                                final_title = cover_info.get('title', detected_title if detected_title else "Document")

                            row1_done = False
                            for c in range(1, 10):
                                v = str(target_ws.cell(row=1, column=c).value or "")
                                if re.search(r'CaRs?', v, re.IGNORECASE) or (c == 5 and v):
                                    target_ws.cell(row=1, column=c).value = f"{car_base} - R1"
                                    row1_done = True
                                    break
                            if not row1_done:
                                target_ws.cell(row=1, column=5).value = f"{car_base} - R1"

                            com_col, rep_col = None, None
                            for c in range(1, target_ws.max_column + 1):
                                val_h = str(target_ws.cell(row=11, column=c).value or "").upper()
                                if "ROUND 1 COMMENTS" in val_h or "ROUND 1 COMMENT" in val_h:
                                    com_col = c
                                elif "ROUND 1 REPLIES" in val_h or "ROUND 1 REPLY" in val_h:
                                    rep_col = c

                            if not com_col: com_col = 8
                            if not rep_col: rep_col = com_col + 2

                            target_ws.cell(row=2, column=com_col).value = "EVNPMB2"
                            target_ws.cell(row=3, column=com_col).value = cover_data['tr_no']
                            target_ws.cell(row=4, column=com_col).value = f"{final_doc_no} rev.{final_rev}"
                            target_ws.cell(row=5, column=com_col).value = final_title
                            target_ws.cell(row=10, column=rep_col).value = "CONTRACTOR's"
                            target_ws.cell(row=11, column=com_col).value = "ROUND 1 COMMENTS"
                            target_ws.cell(row=11, column=rep_col).value = "ROUND 1 REPLIES"

                            status_full_text = "APPROVED (APP)" if selected_status == "APP" else "RETURN FOR CORRECTIONS (RFC)"
                            if selected_status == "AWC": status_full_text = "APPROVED WITH COMMENTS (AWC)"
                            elif selected_status == "RJ": status_full_text = "REJECTED (RJ)"

                            target_ws.cell(row=12, column=com_col).value = "STATUS AFTER ROUND 1 COMMENTS:"
                            target_ws.cell(row=13, column=com_col).value = status_full_text

                            for r in range(1, 15):
                                for c in range(1, target_ws.max_column + 1):
                                    cell = target_ws.cell(row=r, column=c)
                                    if cell.value and cell.font:
                                        f = cell.font
                                        color = f.color if f.color else None
                                        cell.font = Font(name=f.name, size=f.sz, bold=True, italic=f.i, color=color)

                            red_bold = Font(name="Arial", size=10, bold=True, color="FFFF0000")
                            target_ws.cell(row=12, column=com_col).font = red_bold
                            target_ws.cell(row=13, column=com_col).font = red_bold

                            target_ws.column_dimensions[openpyxl.utils.get_column_letter(com_col)].width = 38
                            target_ws.column_dimensions[openpyxl.utils.get_column_letter(rep_col)].width = 38

                            safe_title_for_name = clean_safe_filename(final_title)
                            new_excel_name = f"{car_base} R1 {current_date_str} {final_doc_no} {safe_title_for_name} rev.{final_rev}.xlsx"

                            excel_out = io.BytesIO()
                            wb.save(excel_out)
                            zip_file.writestr(new_excel_name, excel_out.getvalue())
                            del wb, excel_out

                            processed_car_list.append({
                                'doc_no': final_doc_no,
                                'rev': final_rev,
                                'title': final_title,
                                'status': status_full_text,
                                'car_label': car_label,
                                'assigned_car_name': new_excel_name
                            })

                        # 4. PRE PREPARED
                        else:
                            wb_ro = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
                            max_r = 1
                            target_sheet_name = wb_ro.sheetnames[0]
                            for s_name in wb_ro.sheetnames:
                                m = re.search(r'round\s*(\d+)', s_name, re.IGNORECASE)
                                if m and int(m.group(1)) >= max_r:
                                    max_r = int(m.group(1))
                                    target_sheet_name = s_name

                            source_ws = wb_ro[target_sheet_name]
                            car_p = re.search(r'(\d{4,}\s*CaRs)', file.filename, re.IGNORECASE)
                            car_base = car_p.group(1) if car_p else "CaRs"

                            detected_doc_code = extract_clean_doc_no(file.filename)
                            detected_title = ""
                            for row in source_ws.iter_rows(min_row=4, max_row=5, values_only=True):
                                for cell_val in row:
                                    val_str = str(cell_val or "").strip()
                                    if not detected_doc_code:
                                        extracted = extract_clean_doc_no(val_str)
                                        if extracted: detected_doc_code = extracted
                                    if val_str and not detected_title:
                                        detected_title = re.sub(r'_(IFA|APD|IFI|RI|HDEC|Reply).*$', '', val_str, flags=re.IGNORECASE).strip()

                            cover_info = cover_data['docs'].get(detected_doc_code, {})
                            final_doc_no = cover_info.get('doc_no', detected_doc_code if detected_doc_code else "QT1-Document")
                            final_rev = cover_info.get('rev', '1')
                            final_title = cover_info.get('title', detected_title if detected_title else "Document")
                            
                            actual_status_raw = get_latest_status_from_ws(source_ws, max_r)
                            status_full_text = "APPROVED (APP)" if actual_status_raw == "APP" else "RETURN FOR CORRECTIONS (RFC)"
                            if actual_status_raw == "AWC": status_full_text = "APPROVED WITH COMMENTS (AWC)"
                            elif actual_status_raw == "RJ": status_full_text = "REJECTED (RJ)"

                            del wb_ro
                            zip_file.writestr(clean_safe_filename(file.filename), content)
                            
                            processed_car_list.append({
                                'doc_no': final_doc_no,
                                'rev': final_rev,
                                'title': final_title,
                                'status': status_full_text,
                                'car_label': f"{car_base} R{max_r}",
                                'assigned_car_name': file.filename
                            })

                        del content
                        gc.collect()

                    except Exception as fe:
                        print(f"Lỗi xử lý file {file.filename}: {fe}")
                        traceback.print_exc()

            # 5. CHẾ ĐỘ TRẢ LỜI THƯ, KHÔNG CÓ CAR
            elif mode == "no_car_letter":
                if cover_data['docs']:
                    for doc_k, d_info in cover_data['docs'].items():
                        processed_car_list.append({
                            'doc_no': d_info['doc_no'],
                            'rev': d_info['rev'],
                            'title': d_info['title'],
                            'status': 'REVIEWED (NO CAR)',
                            'car_label': 'N/A',
                            'assigned_car_name': 'N/A (Chỉ thư Cover)'
                        })
                else:
                    processed_car_list.append({
                        'doc_no': cover_data['tr_no'],
                        'rev': '0',
                        'title': cover_data['subject'],
                        'status': 'REVIEWED (NO CAR)',
                        'car_label': 'N/A',
                        'assigned_car_name': 'N/A (Chỉ thư Cover)'
                    })

            safe_subj_name = clean_safe_filename(cover_data['subject'])

            # =========================================================================
            # XUẤT TỜ TRÌNH TIẾNG VIỆT
            # =========================================================================
            doc_vi = Document(vi_template_path)
            for tbl in doc_vi.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "ngày" in p.text and "tháng" in p.text:
                                p.runs[0].text = f"Hà Nội, ngày {now.day} tháng {now.month} năm {now.year}"
                                for r in p.runs[1:]: r.text = ""
                            if "V/v" in p.text:
                                replace_value_after_keyword(p, "V/v", cover_data['subject'])

            for p in doc_vi.paragraphs:
                if "Căn cứ" in p.text or "Công ty Cổ phần" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if "Căn cứ công văn số" in p.text:
                    if not replace_value_after_keyword(p, "Căn cứ công văn số", f"{cover_data['tr_no']} ngày {cover_data['date_str']} của Nhà thầu EPC gói 15 về {cover_data['subject']}."):
                        p.text = f"Căn cứ công văn số {cover_data['tr_no']} ngày {cover_data['date_str']} của Nhà thầu EPC gói 15 về {cover_data['subject']}."

            vi_io = io.BytesIO()
            doc_vi.save(vi_io)
            zip_file.writestr(f"TVÐ1-NÐNLM {safe_subj_name}.docx", vi_io.getvalue())
            del vi_io, doc_vi

            # =========================================================================
            # XUẤT THƯ TIẾNG ANH
            # =========================================================================
            doc_en = Document(en_template_path)
            doc_type_en = "Letter" if "-L-" in cover_data['tr_no'].upper() else "Transmittal"

            for tbl in doc_en.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "Ref. No.:" in p.text:
                                replace_value_after_keyword(p, "Ref. No.:", cover_data['tr_no'])
                            elif "October" in p.text or "202" in p.text:
                                update_date_paragraph(p, current_date_en)
                            elif "Identification Number." in p.text:
                                replace_value_after_keyword(p, "Identification Number.", qti_no.strip())
                            elif "Subject:" in p.text:
                                replace_value_after_keyword(p, "Subject:", cover_data['subject'])

            sync_footer_content(doc_en, qti_no.strip(), cover_data['subject'])

            ref_idx = None
            yours_idx = None
            for idx, p in enumerate(doc_en.paragraphs):
                if "This has reference to the" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    font_name = p.runs[0].font.name if (p.runs and p.runs[0].font.name) else "Times New Roman"
                    font_size = p.runs[0].font.size if (p.runs and p.runs[0].font.size) else Pt(13)
                    p.clear()
                    r = p.add_run(f"This has reference to the {doc_type_en} {cover_data['tr_no']}, we would like to comment as follows:")
                    r.font.name = font_name
                    r.font.size = font_size
                    ref_idx = idx
                elif "Yours sincerely" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    yours_idx = idx
                    break

            if ref_idx is not None and yours_idx is not None:
                paras_to_remove = [doc_en.paragraphs[i] for i in range(ref_idx + 1, yours_idx)]
                for p in paras_to_remove:
                    p._element.getparent().remove(p._element)

            target_p = None
            for p in doc_en.paragraphs:
                if "Yours sincerely" in p.text:
                    target_p = p
                    break

            if target_p is not None:
                if mode == "no_car_letter":
                    if cover_data['docs']:
                        for doc_k, d_info in reversed(list(cover_data['docs'].items())):
                            new_p = target_p.insert_paragraph_before("")
                            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            new_p.paragraph_format.left_indent = Inches(0.35)
                            new_p.paragraph_format.first_line_indent = Inches(-0.35)
                            new_p.paragraph_format.space_before = 0
                            new_p.paragraph_format.space_after = Pt(4)
                            
                            run = new_p.add_run(f"-\tThe document {d_info['doc_no']} {d_info['title']} Rev.{d_info['rev']} has been reviewed and found acceptable with no further comments.")
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(13)
                    else:
                        new_p = target_p.insert_paragraph_before("")
                        new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        new_p.paragraph_format.left_indent = Inches(0.35)
                        new_p.paragraph_format.first_line_indent = Inches(-0.35)
                        new_p.paragraph_format.space_before = 0
                        new_p.paragraph_format.space_after = Pt(4)
                        doc_ref_word = "letter/proposal" if "-L-" in cover_data['tr_no'].upper() else "transmittal/documents"
                        run = new_p.add_run(f"-\tWe have reviewed the submitted {doc_ref_word} and have no objection.")
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(13)
                else:
                    for item in reversed(processed_car_list):
                        new_p = target_p.insert_paragraph_before("")
                        new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        new_p.paragraph_format.left_indent = Inches(0.35)
                        new_p.paragraph_format.first_line_indent = Inches(-0.35)
                        new_p.paragraph_format.space_before = 0
                        new_p.paragraph_format.space_after = Pt(4)
                        
                        run = new_p.add_run(f"-\tThe document {item['doc_no']} {item['title']} Rev.{item['rev']} is {item['status']} indicated as per {item['car_label']}.")
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(13)

            en_io = io.BytesIO()
            doc_en.save(en_io)
            safe_qti_name = clean_safe_filename(qti_no)
            zip_file.writestr(f"{safe_qti_name} {safe_subj_name}.docx", en_io.getvalue())
            del en_io, doc_en

        safe_qti_name = clean_safe_filename(qti_no)
        headers = {
            "Content-Disposition": f'attachment; filename="{safe_qti_name}.zip"'
        }
        
        try:
            items_json = json.dumps([
                {
                    "doc_no": item['doc_no'],
                    "assigned_car": item['assigned_car_name'],
                    "status": item['status']
                } for item in processed_car_list
            ], ensure_ascii=False)
            items_b64 = base64.b64encode(items_json.encode('utf-8')).decode('ascii')
            if len(items_b64) < 6000:
                headers["X-Processed-Items"] = items_b64
        except Exception:
            pass

        return FileResponse(
            path=temp_zip_path,
            media_type="application/zip",
            filename=f"{safe_qti_name}.zip",
            headers=headers,
            background=BackgroundTask(remove_temp_file, temp_zip_path)
        )

    except Exception as ge:
        remove_temp_file(temp_zip_path)
        print("Lỗi hệ thống:", ge)
        traceback.print_exc()
        return Response(content=f"Error: {str(ge)}", status_code=400)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
