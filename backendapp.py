from fastapi import FastAPI, UploadFile, File, Form, Response
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.styles import Font
from copy import copy
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re, io, zipfile, traceback, os, base64
from datetime import datetime
import pypdf

app = FastAPI(title="QT1 Doc Control API - Full Header Footer Sync")

# --- CẤU HÌNH CORS ĐÃ ĐƯỢC CHUẨN HÓA MỞ CỔNG BẢO MẬT ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def clean_safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', '-', name).strip()

def extract_clean_doc_no(text: str) -> str:
    m = re.search(r'(QT1-[0-9L]-[A-Z]+-[A-Z0-9]+-\d{2}-\d{5})', text)
    if m: return m.group(1)
    m2 = re.search(r'(QT1-[0-9L]-[A-Z0-9-]+)', text)
    if m2:
        base = m2.group(1)
        base = re.sub(r'_[A-Za-z0-9]+$', '', base)
        base = re.sub(r'_?Rev\.?.*$', '', base, flags=re.IGNORECASE)
        return base
    return ""

def get_latest_status_from_ws(ws, max_r):
    target_col = None
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(row=11, column=c).value or "").upper()
        if f"ROUND {max_r} COMMENTS" in val:
            target_col = c
            break
            
    if target_col:
        status_val = str(ws.cell(row=13, column=target_col).value or "").upper()
        if "APP" in status_val and "AWC" not in status_val: return "APP"
        elif "RFC" in status_val: return "RFC"
        elif "AWC" in status_val: return "AWC"
        elif "RJ" in status_val: return "RJ"
    return "APP"

def parse_cover_pdf_bytes(pdf_bytes):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    metadata = {
        'tr_no': 'QT1-HDH-PMB2-T-6015',
        'date_str': datetime.now().strftime("%d/%m/%Y"), 
        'subject': 'Entire set of Mechnical P&IDs',
        'docs': {}
    }
    p1_txt = reader.pages[0].extract_text() or ""
    
    tr_m = re.search(r'TR\.\s*No\.?\s*[:\s]*([A-Z0-9_-]+)', p1_txt)
    if tr_m: metadata['tr_no'] = tr_m.group(1).strip()
    
    date_m = re.search(r'Date\s*[:\s]+([\d\s/]+)', p1_txt, re.IGNORECASE)
    if date_m:
        clean_date = re.sub(r'[^\d]', '', date_m.group(1))
        if len(clean_date) >= 8:
            clean_date = clean_date[:8]
            if clean_date[4:8].startswith("202"):
                metadata['date_str'] = f"{clean_date[:2]}/{clean_date[2:4]}/{clean_date[4:8]}"
            
    subj_m = re.search(r'SUBJECT\s*[:\s]*([^\n]+)', p1_txt)
    if subj_m: metadata['subject'] = subj_m.group(1).split('Document/Drawing')[0].strip()

    for page in reader.pages:
        parts = []
        def visitor(text, cm, tm, fontDict, fontSize):
            if text.strip(): parts.append((tm[4], tm[5], text.strip()))
        page.extract_text(visitor_text=visitor)
        
        rows = {}
        for px, py, ptxt in parts:
            matched_y = None
            for y in rows.keys():
                if abs(y - py) <= 3:
                    matched_y = y
                    break
            if matched_y is None:
                matched_y = py
                rows[matched_y] = []
            rows[matched_y].append((px, ptxt))
            
        for y, items in rows.items():
            items.sort(key=lambda x: x[0])
            doc_no = None
            for i, (x, txt) in enumerate(items):
                m = re.search(r'(QT1-[0-9L]-[A-Z]+-[A-Z0-9]+-\d{2}-\d{5})', txt)
                if m:
                    doc_no = m.group(1)
                    rev = "1"
                    title_parts = []
                    if i + 1 < len(items):
                        rev_candidate = items[i+1][1].strip()
                        if len(rev_candidate) <= 3 and rev_candidate.isalnum():
                            rev = rev_candidate
                            title_start_idx = i + 2
                        else:
                            title_start_idx = i + 1
                            
                        for j in range(title_start_idx, len(items)):
                            t = items[j][1]
                            if t not in ['APD', 'IFA', 'IFI', 'RI', '- END -']:
                                title_parts.append(t)
                    
                    title = " ".join(title_parts).strip()
                    title = re.sub(r'\s+(APD|IFA|IFI|RI|- END -).*$', '', title).strip()
                    
                    if doc_no:
                        metadata['docs'][doc_no] = {
                            'doc_no': doc_no,
                            'rev': rev,
                            'title': title
                        }
                    break
    return metadata

def parse_cover_excel_bytes(excel_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    metadata = {
        'tr_no': 'QT1-HDH-PMB2-T-6015',
        'date_str': datetime.now().strftime("%d/%m/%Y"),
        'subject': 'Entire set of Mechnical P&IDs',
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
        val_name, val_size, val_bold, val_italic = val_font.name, val_font.size, val_bold, val_font.italic
        
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

# THUẬT TOÁN ĐỒNG BỘ CHÂN TRANG (FOOTER)
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
                            # Cấu hình màu xám chuẩn cho Footer nếu file mất định dạng
                            r.font.color.rgb = RGBColor(128, 128, 128)
                            r.font.size = Pt(10)
                    else:
                        p.text = new_text

@app.post("/api/export-zip-full")
async def export_zip_full(
    files: list[UploadFile] = File(...),
    cover_file: UploadFile = File(None),
    metadata: str = Form(...),
    mode: str = Form("auto_bump"),
    qti_no: str = Form("QTI-PECC1-PMB2-26-10-111")
):
    import json
    try:
        car_metadata = json.loads(metadata)
        now = datetime.now()
        current_date_str = now.strftime("%d %b %Y")
        current_date_en = now.strftime("%B %d, %Y")
        
        cover_data = {
            'tr_no': 'QT1-HDH-PMB2-T-6015',
            'date_str': now.strftime("%d/%m/%Y"),
            'subject': 'Entire set of Mechnical P&IDs',
            'docs': {}
        }
        
        if cover_file is not None:
            cover_bytes = await cover_file.read()
            if cover_file.filename.lower().endswith('.pdf'):
                cover_data = parse_cover_pdf_bytes(cover_bytes)
            elif cover_file.filename.lower().endswith(('.xlsx', '.xls')):
                cover_data = parse_cover_excel_bytes(cover_bytes)

        vi_template_path = find_template("TV")
        if not vi_template_path: raise FileNotFoundError(f"Vui lòng copy file Word mẫu 'TVÐ1-NÐNLM.docx' vào thư mục: {BASE_DIR}")
        en_template_path = find_template("QTI-PECC1")
        if not en_template_path: raise FileNotFoundError(f"Vui lòng copy file Word mẫu 'QTI-PECC1...docx' vào thư mục: {BASE_DIR}")

        zip_io = io.BytesIO()
        processed_car_list = []

        with zipfile.ZipFile(zip_io, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            
            for file in files:
                try:
                    content = await file.read()
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
                    car_p = re.search(r'(\d{4}\s*CaRs)', file.filename, re.IGNORECASE)
                    car_base = car_p.group(1) if car_p else "CaRs"

                    if mode == "auto_bump":
                        selected_status = car_metadata.get(file.filename, "APP")
                        old_status_raw = get_latest_status_from_ws(source_ws, max_r)
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
                                    tgt_c.font, tgt_c.border, tgt_c.alignment = copy(src_c.font), copy(src_c.border), copy(src_c.alignment)

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
                        
                        processed_car_list.append({
                            'doc_no': final_doc_no,
                            'rev': final_rev,
                            'title': final_title,
                            'status': status_full_text,
                            'car_label': f"{car_base} R{new_r_num}",
                            'assigned_car_name': new_excel_name
                        })

                    else:
                        detected_doc_code = extract_clean_doc_no(file.filename)
                        detected_title = ""
                        if not detected_doc_code:
                            for c in range(1, source_ws.max_column + 1):
                                r4_val = str(source_ws.cell(row=4, column=c).value or "")
                                extracted = extract_clean_doc_no(r4_val)
                                if extracted:
                                    detected_doc_code = extracted
                                    break
                        
                        for c in range(1, source_ws.max_column + 1):
                            r5_val = str(source_ws.cell(row=5, column=c).value or "").strip()
                            if r5_val and not detected_title:
                                detected_title = re.sub(r'_(IFA|APD|IFI|RI|HDEC|Reply).*$', '', r5_val, flags=re.IGNORECASE).strip()

                        cover_info = cover_data['docs'].get(detected_doc_code, {})
                        final_doc_no = cover_info.get('doc_no', detected_doc_code if detected_doc_code else "QT1-Document")
                        final_rev = cover_info.get('rev', '1')
                        final_title = cover_info.get('title', detected_title if detected_title else "Document")
                        
                        actual_status_raw = get_latest_status_from_ws(source_ws, max_r)
                        status_full_text = "APPROVED (APP)" if actual_status_raw == "APP" else "RETURN FOR CORRECTIONS (RFC)"
                        if actual_status_raw == "AWC": status_full_text = "APPROVED WITH COMMENTS (AWC)"
                        elif actual_status_raw == "RJ": status_full_text = "REJECTED (RJ)"

                        zip_file.writestr(clean_safe_filename(file.filename), content)
                        
                        processed_car_list.append({
                            'doc_no': final_doc_no,
                            'rev': final_rev,
                            'title': final_title,
                            'status': status_full_text,
                            'car_label': f"{car_base} R{max_r}",
                            'assigned_car_name': file.filename
                        })

                except Exception as fe:
                    print(f"Lỗi xử lý file {file.filename}: {fe}")
                    traceback.print_exc()

            safe_subj_name = clean_safe_filename(cover_data['subject'])

            # =========================================================================
            # 2. XỬ LÝ TỜ TRÌNH TIẾNG VIỆT
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

            # =========================================================================
            # 3. XỬ LÝ THƯ TIẾNG ANH VỚI TÊN FILE VÀ CHÂN TRANG ĐỒNG BỘ
            # =========================================================================
            doc_en = Document(en_template_path)
            
            # Đồng bộ phần Tiêu đề
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

            # ĐỒNG BỘ CHÂN TRANG (FOOTER)
            sync_footer_content(doc_en, qti_no.strip(), cover_data['subject'])

            # Chèn nội dung thân thư
            ref_idx = None
            yours_idx = None
            for idx, p in enumerate(doc_en.paragraphs):
                if "This has reference to the Transmittal" in p.text:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    replace_value_after_keyword(p, "This has reference to the Transmittal", f"{cover_data['tr_no']}, we would like to comment as follows:")
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

        zip_io.seek(0)
        zip_base64 = base64.b64encode(zip_io.getvalue()).decode('utf-8')
        
        return {
            "status": "success",
            "zip_base64": zip_base64,
            "processed_items": [
                {
                    "doc_no": item['doc_no'],
                    "assigned_car": item['assigned_car_name'],
                    "status": item['status']
                } for item in processed_car_list
            ]
        }
    except Exception as ge:
        print("Lỗi hệ thống:", ge)
        traceback.print_exc()
        return Response(content=f"Error: {str(ge)}", status_code=400)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
