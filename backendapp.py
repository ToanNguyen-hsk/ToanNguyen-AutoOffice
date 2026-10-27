import base64
from copy import copy
from datetime import datetime
import gc
import io
import json
import os
import re
import tempfile
import traceback
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from fastapi import FastAPI, File, Form, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import openpyxl
from openpyxl.styles import Font
import pypdf
from starlette.background import BackgroundTask

app = FastAPI(title="QT1 Doc Control API - Master Precision Engine")

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
  cleaned = re.sub(r'[\\/*?:"<>|]', "-", name).strip()
  return cleaned[:130] if len(cleaned) > 130 else cleaned


def extract_cover_date_bulletproof(text: str) -> str:
  month_map = {
      "jan": 1,
      "january": 1,
      "feb": 2,
      "february": 2,
      "mar": 3,
      "march": 3,
      "apr": 4,
      "april": 4,
      "may": 5,
      "jun": 6,
      "june": 6,
      "jul": 7,
      "july": 7,
      "aug": 8,
      "august": 8,
      "sep": 9,
      "september": 9,
      "sept": 9,
      "oct": 10,
      "october": 10,
      "nov": 11,
      "november": 11,
      "dec": 12,
      "december": 12,
  }
  month_pat = r"(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)"

  m_date = re.search(r"\bDate\b", text, re.IGNORECASE)
  if m_date:
    snippet = text[m_date.end() : m_date.end() + 120]
    snippet_clean = re.sub(
        r"Day\s*Mon\s*Year", "", snippet, flags=re.IGNORECASE
    )
    snippet_clean = re.sub(
        r"Sheet\s*:\s*\d+\s*of\s*\d+", "", snippet_clean, flags=re.IGNORECASE
    )

    # 1. Dạng 8 chữ số cách nhau trong Transmittal: "1 0 1 0 2 0 2 6" hoặc "10 10 2026"
    m_8digits = re.search(
        r"[:\s-]*\s*(\d)\s*(\d)\s*(\d)\s*(\d)\s*(2)\s*(0)\s*(2)\s*(\d)",
        snippet_clean,
    )
    if m_8digits:
      g = m_8digits.groups()
      d_val = int(g[0] + g[1])
      m_val = int(g[2] + g[3])
      y_val = g[4] + g[5] + g[6] + g[7]
      if 1 <= m_val <= 12 and 1 <= d_val <= 31:
        return f"{d_val:02d}/{m_val:02d}/{y_val}"

    # 2. Dạng DD/MM/YYYY
    m_slash = re.search(
        r"[:\s-]*\s*(\d{1,2})[\/.-](\d{1,2})[\/.-](202\d)", snippet_clean
    )
    if m_slash:
      d_val = int(m_slash.group(1))
      m_val = int(m_slash.group(2))
      y_val = m_slash.group(3)
      if 1 <= m_val <= 12 and 1 <= d_val <= 31:
        return f"{d_val:02d}/{m_val:02d}/{y_val}"

    # 3. Dạng chữ tiếng Anh: "October 19, 2026"
    collapsed = re.sub(r"\s+", " ", snippet_clean).strip()
    m_en = re.search(
        r"\b(" + month_pat + r")\s*(\d{1,2})\s*,?\s*(202\s*\d)",
        collapsed,
        re.IGNORECASE,
    )
    if m_en:
      mon_str = m_en.group(1).lower()
      d_val = int(m_en.group(2))
      y_val = re.sub(r"\s+", "", m_en.group(3))
      if mon_str in month_map and 1 <= d_val <= 31:
        return f"{d_val:02d}/{month_map[mon_str]:02d}/{y_val}"

  return ""


def extract_cover_metadata(pdf_bytes: bytes, filename: str = "") -> dict:
  all_texts = []
  try:
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    for p in reader.pages:
      t = p.extract_text()
      if t:
        all_texts.append(t)
  except Exception as e:
    print(f"pypdf error: {e}")

  full_text = "\n".join(all_texts)

  tr_no = ""
  m_fn = re.search(
      r"(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)", filename
  )
  if m_fn:
    tr_no = m_fn.group(1).strip()
  if not tr_no:
    m_ref = re.search(
        r"(?:TR\.?\s*No\.?|Ref\.?)\s*[:\s]*([A-Za-z0-9_-]*QT1[A-Za-z0-9_-]+)",
        full_text,
        re.IGNORECASE,
    )
    if m_ref and "EVNPMB2" not in m_ref.group(1):
      tr_no = m_ref.group(1).strip()
  if not tr_no:
    tr_no = "QT1-Transmittal"

  date_str = extract_cover_date_bulletproof(full_text)

  subject = ""
  m_sub = re.search(
      r"(?:SUBJECT|Subject)\s*[:\s]+([^\n\r]+)", full_text, re.IGNORECASE
  )
  if m_sub:
    raw_sub = m_sub.group(1).strip()
    raw_sub = re.split(
        r"(?:Reference|Dear|Attn|Attachment|Document/Drawing|Sheet:)",
        raw_sub,
        flags=re.IGNORECASE,
    )[0].strip()
    subject = raw_sub

  if not subject and tr_no and filename:
    parts = os.path.splitext(filename)[0].split(tr_no, 1)
    if len(parts) > 1 and parts[1].strip("_- "):
      subject = parts[1].strip("_- ")
      subject = re.sub(r"(_|-)?Cover$", "", subject, flags=re.IGNORECASE).strip()

  if not subject:
    subject = "Official Transmittal"

  docs = {}
  purpose_markers = {
      "IFA",
      "IFI",
      "APD",
      "AB",
      "RI",
      "END",
      "- END -",
      "-END-",
      "-",
  }
  for line in full_text.splitlines():
    m_doc = re.search(r"(QT1-[0-9L]-[A-Za-z0-9_-]+)", line)
    if (
        m_doc
        and "EVNPMB2" not in m_doc.group(1)
        and "QT1-HDH" not in m_doc.group(1)
    ):
      doc_code = m_doc.group(1).strip()
      after = line[m_doc.end() :].strip()

      words = after.split()
      while words and words[-1].upper() in purpose_markers:
        words.pop()
      while words and words[-1] in ["-", "END"]:
        words.pop()

      rev = "1"
      if (
          len(words) >= 2
          and words[0].isdigit()
          and (len(words[1]) <= 3 and words[1].isalnum() and not words[1].isdigit())
      ):
        words.pop(0)
        rev = words.pop(0)
      elif len(words) >= 1 and len(words[0]) <= 3 and words[0].isalnum():
        rev = words.pop(0)
      elif len(words) >= 1 and re.match(
          r"^Rev\.?([A-Za-z0-9]+)$", words[0], re.IGNORECASE
      ):
        m_r = re.match(r"^Rev\.?([A-Za-z0-9]+)$", words[0], re.IGNORECASE)
        rev = m_r.group(1)
        words.pop(0)

      title = " ".join(words).strip()
      if not title:
        title = subject

      docs[doc_code] = {"doc_no": doc_code, "rev": rev, "title": title}

  return {
      "tr_no": tr_no,
      "date_str": date_str if date_str else datetime.now().strftime("%d/%m/%Y"),
      "subject": subject,
      "docs": docs,
  }


def parse_cover_excel_bytes(excel_bytes, filename=""):
  fn_doc_no = ""
  m_fn = re.search(
      r"(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)", filename
  )
  if m_fn:
    fn_doc_no = m_fn.group(1)

  wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
  ws = wb.active
  metadata = {
      "tr_no": fn_doc_no if fn_doc_no else "QT1-Document",
      "date_str": datetime.now().strftime("%d/%m/%Y"),
      "subject": "Official Transmittal",
      "docs": {},
  }
  for r in range(1, ws.max_row + 1):
    for c in range(1, ws.max_column + 1):
      val = str(ws.cell(row=r, column=c).value or "").strip()
      m = re.match(r"^(QT1-[0-9L]-[A-Z0-9_-]+)$", val)
      if m and "QT1-HDH" not in val:
        doc_no = m.group(1)
        rev_val = str(ws.cell(row=r, column=c + 1).value or "1").strip()
        title_val = str(ws.cell(row=r, column=c + 2).value or "").strip()
        metadata["docs"][doc_no] = {
            "doc_no": doc_no,
            "rev": rev_val,
            "title": title_val,
        }
  return metadata


def find_template(prefix):
  search_dirs = [BASE_DIR, os.getcwd(), r"G:\Shared drives\Test auto"]
  for d in search_dirs:
    if not os.path.exists(d):
      continue
    for f in os.listdir(d):
      if f.lower().endswith(".docx"):
        clean_f = f.lower().replace("đ", "d").replace("ð", "d")
        clean_p = prefix.lower().replace("đ", "d").replace("ð", "d")
        if clean_p in clean_f and not f.startswith("~$"):
          return os.path.join(d, f)
  return None


def replace_value_after_keyword(paragraph, keyword, new_text):
  if keyword in paragraph.text:
    if not paragraph.runs:
      paragraph.text = f"{keyword} {new_text}"
      return True
    kw_font = paragraph.runs[0].font
    kw_name = kw_font.name
    kw_size = kw_font.size
    kw_bold = kw_font.bold
    kw_italic = kw_font.italic

    val_font = paragraph.runs[-1].font
    val_name = val_font.name
    val_size = val_font.size
    val_bold = val_font.bold
    val_italic = val_font.italic

    paragraph.clear()
    r1 = paragraph.add_run(keyword)
    r1.font.name = kw_name
    r1.font.size = kw_size
    r1.font.bold = kw_bold
    r1.font.italic = kw_italic

    r2 = paragraph.add_run(" " + new_text)
    r2.font.name = val_name
    r2.font.size = val_size
    r2.font.bold = val_bold
    r2.font.italic = val_italic
    return True
  return False


def update_date_paragraph(paragraph, new_date_str):
  spaces = "                                "
  val_font = paragraph.runs[-1].font
  val_name = val_font.name
  val_size = val_font.size
  val_bold = val_font.bold
  val_italic = val_font.italic

  paragraph.clear()
  r1 = paragraph.add_run(spaces)
  r2 = paragraph.add_run(new_date_str)
  r2.font.name = val_name
  r2.font.size = val_size
  r2.font.bold = val_bold
  r2.font.italic = val_italic


def sync_footer_content(doc, qti_no, subject):
  new_text = f"{qti_no} {subject}"
  for section in doc.sections:
    footer = section.footer
    if footer:
      for p in footer.paragraphs:
        if p.text.strip():
          if len(p.runs) > 0:
            first_font = p.runs[0].font
            p.clear()
            r = p.add_run(new_text)
            r.font.name = first_font.name
            r.font.size = first_font.size
            r.font.color.rgb = (
                first_font.color.rgb
                if first_font.color
                else RGBColor(128, 128, 128)
            )
          else:
            p.text = new_text


def remove_temp_file(path: str):
  try:
    if os.path.exists(path):
      os.remove(path)
  except Exception as e:
    print(f"Cleanup error {path}: {e}")


@app.post("/api/export-zip-full")
async def export_zip_full(
    files: list[UploadFile] = File(default=[]),
    cover_file: UploadFile = File(None),
    metadata: str = Form("{}"),
    mode: str = Form("number_filled_cars"),
    qti_no: str = Form("QTI-PECC1-PMB2-26-10-147"),
    start_car_no: str = Form("9905"),
    custom_date: str = Form(""),
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
        "tr_no": "QT1-Document",
        "date_str": now.strftime("%d/%m/%Y"),
        "subject": "Official Transmittal",
        "docs": {},
    }

    if cover_file is not None:
      cover_bytes = await cover_file.read()
      if cover_file.filename.lower().endswith(".pdf"):
        cover_data = extract_cover_metadata(cover_bytes, cover_file.filename)
      elif cover_file.filename.lower().endswith((".xlsx", ".xls")):
        cover_data = parse_cover_excel_bytes(cover_bytes, cover_file.filename)
      del cover_bytes

    if custom_date and custom_date.strip():
      cover_data["date_str"] = custom_date.strip()

    vi_template_path = find_template("TV")
    if not vi_template_path:
      raise FileNotFoundError(
          f"Thiếu file mẫu TVÐ1-NÐNLM.docx trong thư mục {BASE_DIR}"
      )
    en_template_path = find_template("QTI-PECC1")
    if not en_template_path:
      raise FileNotFoundError(
          f"Thiếu file mẫu QTI-PECC1.docx trong thư mục {BASE_DIR}"
      )

    processed_car_list = []

    try:
      start_num = int(re.sub(r"\D", "", str(start_car_no)) or "9905")
    except Exception:
      start_num = 9905

    with zipfile.ZipFile(temp_zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
      if mode != "no_car_letter" and files:
        for file_idx, file in enumerate(files):
          try:
            content = await file.read()

            # =========================================================================
            # CHẾ ĐỘ: NUMBER FILLED CARS (KỸ SƯ ĐÃ ĐIỀN ĐẦY ĐỦ - CHỈ CẤP SỐ CAR VÀO E1)
            # =========================================================================
            if mode == "number_filled_cars":
              cur_num = start_num + file_idx
              car_num_str = f"{cur_num:04d}" if cur_num < 10000 else f"{cur_num}"
              car_base = f"{car_num_str} CaRs"

              wb = openpyxl.load_workbook(io.BytesIO(content))
              max_r = 1
              target_ws = wb.active
              for s_name in wb.sheetnames:
                m = re.search(r"round\s*(\d+)", s_name, re.IGNORECASE)
                if m and int(m.group(1)) >= max_r:
                  max_r = int(m.group(1))
                  target_ws = wb[s_name]

              car_label = f"{car_base} R{max_r}"

              # Xác định cột nhận xét
              com_col = None
              for c in range(1, target_ws.max_column + 1):
                val_h = str(target_ws.cell(row=11, column=c).value or "").upper()
                if (
                    f"ROUND {max_r} COMMENTS" in val_h
                    or f"ROUND {max_r} COMMENT" in val_h
                ):
                  com_col = c
                  break
              if not com_col:
                com_col = 7

              # 1. ĐỌC TỰ ĐỘNG STATUS TỪ FILE CARS CỦA KỸ SƯ
              detected_status_raw = "APPROVED (APP)"
              found_rfc, found_awc, found_rj = False, False, False
              for r in range(11, 16):
                v_st = str(
                    target_ws.cell(row=r, column=com_col).value or ""
                ).strip().upper()
                if (
                    "RETURN FOR CORRECTIONS" in v_st
                    or "(RFC)" in v_st
                    or v_st == "RFC"
                ):
                  found_rfc = True
                elif (
                    "APPROVED WITH COMMENTS" in v_st
                    or "(AWC)" in v_st
                    or v_st == "AWC"
                ):
                  found_awc = True
                elif "REJECTED" in v_st or "(RJ)" in v_st or v_st == "RJ":
                  found_rj = True

              if found_rfc:
                detected_status_raw = "RETURN FOR CORRECTIONS (RFC)"
              elif found_awc:
                detected_status_raw = "APPROVED WITH COMMENTS (AWC)"
              elif found_rj:
                detected_status_raw = "REJECTED (RJ)"
              else:
                detected_status_raw = "APPROVED (APP)"

              # 2. BÓC TÁCH MÃ DOC VÀ TITLE CỦA KỸ SƯ
              detected_doc_code = ""
              m_fdoc = re.search(r"(QT1-[A-Za-z0-9_-]+)", file.filename)
              if m_fdoc:
                detected_doc_code = m_fdoc.group(1)

              r4_val = str(target_ws.cell(row=4, column=com_col).value or "")
              m_c = re.search(r"(QT1-[A-Za-z0-9_-]+)", r4_val)
              if m_c:
                detected_doc_code = m_c.group(1)

              # Khớp với Transmittal Cover
              doc_keys = list(cover_data["docs"].keys())
              if detected_doc_code in cover_data["docs"]:
                cover_info = cover_data["docs"][detected_doc_code]
              elif doc_keys and file_idx < len(doc_keys):
                cover_info = cover_data["docs"][doc_keys[file_idx]]
              elif doc_keys:
                cover_info = cover_data["docs"][doc_keys[0]]
              else:
                cover_info = {}

              final_doc_no = cover_info.get(
                  "doc_no",
                  detected_doc_code if detected_doc_code else "QT1-Document",
              )
              final_rev = cover_info.get("rev", "1")

              # Lấy title chính xác (loại bỏ chữ 'iv' hoặc '(Rev. A)')
              r5_val = str(
                  target_ws.cell(row=5, column=com_col).value or ""
              ).strip()
              if r5_val.lower() in ["iv", "v", "title of document", ""]:
                r5_val = ""
              r5_val = re.sub(
                  r"\s*\(Rev\.\s*[A-Za-z0-9]+\)", "", r5_val, flags=re.IGNORECASE
              ).strip()

              if cover_info.get("title"):
                final_title = cover_info["title"]
              elif r5_val:
                final_title = r5_val
              else:
                clean_t = re.sub(
                    r"^(?:00xx|Xxxx|\d+)\s*CaRs?\s*R\d+\s*",
                    "",
                    file.filename,
                    flags=re.IGNORECASE,
                )
                clean_t = os.path.splitext(clean_t)[0].strip()
                final_title = clean_t if clean_t else cover_data["subject"]

              # 3. CẤP SỐ CAR CHUẨN VÀO Ô E1 (THAY THẾ 00xx / Xxxx / 00xx-R1 / Xxxx-R1)
              red_title_font = Font(
                  name="Times New Roman", size=16, bold=True, color="FFFF0000"
              )
              e1_text = f"{car_num_str}-R{max_r}"

              # Tìm ô chứa mã tạm ở dòng 1
              found_cell_r1 = False
              for c in range(2, 10):
                val_r1 = str(target_ws.cell(row=1, column=c).value or "").strip()
                if re.search(r"(00xx|xxxx|Xxxx)", val_r1, re.IGNORECASE):
                  target_ws.cell(row=1, column=c).value = e1_text
                  target_ws.cell(row=1, column=c).font = red_title_font
                  found_cell_r1 = True
                  break

              if not found_cell_r1:
                target_ws.cell(row=1, column=5).value = e1_text
                target_ws.cell(row=1, column=5).font = red_title_font

              # ĐỒNG BỘ MÃ DOC & TITLE VÀO DÒNG 4, 5 (BẢO LƯU NGUYÊN VẸN SENDER G2 VÀ Ý KIẾN KỸ SƯ)
              target_ws.cell(row=4, column=com_col).value = (
                  f"{final_doc_no} rev.{final_rev}"
              )
              target_ws.cell(row=5, column=com_col).value = final_title

              safe_title_for_name = clean_safe_filename(final_title)
              new_excel_name = f"{car_base} R{max_r} {current_date_str} {final_doc_no} {safe_title_for_name} rev.{final_rev}.xlsx"

              excel_out = io.BytesIO()
              wb.save(excel_out)
              zip_file.writestr(new_excel_name, excel_out.getvalue())
              del wb, excel_out

              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": detected_status_raw,
                  "car_label": car_label,
                  "assigned_car_name": new_excel_name,
              })

            # =========================================================================
            # CHẾ ĐỘ: AUTO BUMP (TỰ ĐỘNG TẠO ROUND MỚI R+1)
            # =========================================================================
            elif mode == "auto_bump":
              wb = openpyxl.load_workbook(io.BytesIO(content))
              max_r = 1
              target_sheet_name = wb.sheetnames[0]
              target_idx = 0
              for idx, s_name in enumerate(wb.sheetnames):
                m = re.search(r"round\s*(\d+)", s_name, re.IGNORECASE)
                if m and int(m.group(1)) >= max_r:
                  max_r = int(m.group(1))
                  target_sheet_name = s_name
                  target_idx = idx

              source_ws = wb[target_sheet_name]
              car_p = re.search(r"(\d{4,}\s*CaRs)", file.filename, re.IGNORECASE)
              car_base = car_p.group(1) if car_p else "CaRs"

              selected_status = car_metadata.get(file.filename, "APP")
              new_r_num = max_r + 1
              new_ws = wb.copy_worksheet(source_ws)
              new_ws.title = f"ROUND {new_r_num} CaRs"
              wb._sheets.remove(new_ws)
              wb._sheets.insert(target_idx + 1, new_ws)

              prev_com_col, prev_rep_col = None, None
              for c in range(1, new_ws.max_column + 1):
                val_h = str(new_ws.cell(row=11, column=c).value or "").upper()
                if f"ROUND {max_r} COMMENTS" in val_h:
                  prev_com_col = c
                elif f"ROUND {max_r} REPLIES" in val_h:
                  prev_rep_col = c

              detected_doc_code = ""
              m_fdoc = re.search(r"(QT1-[A-Za-z0-9_-]+)", file.filename)
              if m_fdoc:
                detected_doc_code = m_fdoc.group(1)

              if prev_com_col:
                r4_val = str(source_ws.cell(row=4, column=prev_com_col).value or "")
                m_c = re.search(r"(QT1-[A-Za-z0-9_-]+)", r4_val)
                if m_c:
                  detected_doc_code = m_c.group(1)

              cover_info = cover_data["docs"].get(detected_doc_code, {})
              final_doc_no = cover_info.get(
                  "doc_no",
                  detected_doc_code if detected_doc_code else "QT1-Document",
              )
              final_rev = cover_info.get("rev", "1")
              final_title = cover_info.get("title", cover_data["subject"])

              for c in range(2, 10):
                v = str(new_ws.cell(row=1, column=c).value or "")
                if re.search(r"-\s*R\d+", v, re.IGNORECASE) or (c == 5 and v):
                  new_ws.cell(row=1, column=c).value = re.sub(
                      r"R\d+", f"R{new_r_num}", v, flags=re.IGNORECASE
                  )
                  break

              new_com_col = (
                  (prev_rep_col + 1) if prev_rep_col else (new_ws.max_column + 1)
              )
              new_rep_col = new_com_col + 2

              new_ws.cell(row=2, column=new_com_col).value = "EVNPMB2"
              new_ws.cell(row=3, column=new_com_col).value = cover_data["tr_no"]
              new_ws.cell(row=4, column=new_com_col).value = (
                  f"{final_doc_no} rev.{final_rev}"
              )
              new_ws.cell(row=5, column=new_com_col).value = final_title
              new_ws.cell(row=10, column=new_rep_col).value = "CONTRACTOR's"
              new_ws.cell(row=11, column=new_com_col).value = (
                  f"ROUND {new_r_num} COMMENTS"
              )
              new_ws.cell(row=11, column=new_rep_col).value = (
                  f"ROUND {new_r_num} REPLIES"
              )

              for r in range(1, 15):
                if prev_com_col:
                  src_c = new_ws.cell(row=r, column=prev_com_col)
                  tgt_c = new_ws.cell(row=r, column=new_com_col)
                  if src_c.has_style:
                    tgt_c.font, tgt_c.border, tgt_c.fill, tgt_c.alignment = (
                        copy(src_c.font),
                        copy(src_c.border),
                        copy(src_c.fill),
                        copy(src_c.alignment),
                    )
                if prev_rep_col:
                  src_r = new_ws.cell(row=r, column=prev_rep_col)
                  tgt_r = new_ws.cell(row=r, column=new_rep_col)
                  if src_r.has_style:
                    tgt_r.font, tgt_r.border, tgt_r.fill, tgt_r.alignment = (
                        copy(src_r.font),
                        copy(src_r.border),
                        copy(src_r.fill),
                        copy(src_r.alignment),
                    )

              status_full_text = (
                  "APPROVED (APP)"
                  if selected_status == "APP"
                  else "RETURN FOR CORRECTIONS (RFC)"
              )
              if selected_status == "AWC":
                status_full_text = "APPROVED WITH COMMENTS (AWC)"
              elif selected_status == "RJ":
                status_full_text = "REJECTED (RJ)"

              new_ws.cell(row=12, column=new_com_col).value = (
                  f"STATUS AFTER ROUND {new_r_num} COMMENTS:"
              )
              new_ws.cell(row=13, column=new_com_col).value = status_full_text

              for r in range(1, 15):
                for c in range(1, new_ws.max_column + 1):
                  cell = new_ws.cell(row=r, column=c)
                  if cell.value and cell.font:
                    f = cell.font
                    color = f.color if f.color else None
                    cell.font = Font(
                        name=f.name,
                        size=f.sz,
                        bold=True,
                        italic=f.i,
                        color=color,
                    )

              red_bold = Font(
                  name="Arial", size=10, bold=True, color="FFFF0000"
              )
              new_ws.cell(row=12, column=new_com_col).font = red_bold
              new_ws.cell(row=13, column=new_com_col).font = red_bold

              for r in range(14, new_ws.max_row + 1):
                v_prev_c = (
                    str(new_ws.cell(row=r, column=prev_com_col).value or "")
                    .strip()
                    if prev_com_col
                    else ""
                )
                v_prev_r = (
                    str(new_ws.cell(row=r, column=prev_rep_col).value or "")
                    .strip()
                    if prev_rep_col
                    else ""
                )

                if prev_com_col:
                  src_c = new_ws.cell(row=r, column=prev_com_col)
                  tgt_c = new_ws.cell(row=r, column=new_com_col)
                  if src_c.has_style:
                    tgt_c.font, tgt_c.border, tgt_c.alignment = (
                        copy(src_c.font),
                        copy(src_c.border),
                        copy(src_c.alignment),
                    )

                is_group = (
                    "general comment" in v_prev_c.lower()
                    or "detailed comment" in v_prev_c.lower()
                )
                if is_group:
                  new_ws.cell(row=r, column=new_com_col).value = None
                elif v_prev_c or v_prev_r:
                  if selected_status == "APP":
                    new_ws.cell(row=r, column=new_com_col).value = "Noted."
                else:
                  new_ws.cell(row=r, column=new_com_col).value = None

              new_ws.column_dimensions[
                  openpyxl.utils.get_column_letter(new_com_col)
              ].width = 38
              new_ws.column_dimensions[
                  openpyxl.utils.get_column_letter(new_rep_col)
              ].width = 38

              safe_title_for_name = clean_safe_filename(final_title)
              new_excel_name = f"{car_base} R{new_r_num} {current_date_str} {final_doc_no} {safe_title_for_name} rev.{final_rev}.xlsx"

              excel_out = io.BytesIO()
              wb.save(excel_out)
              zip_file.writestr(new_excel_name, excel_out.getvalue())
              del wb, excel_out

              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": status_full_text,
                  "car_label": f"{car_base} R{new_r_num}",
                  "assigned_car_name": new_excel_name,
              })

            # =========================================================================
            # CHẾ ĐỘ: PRE PREPARED (ĐÃ CÓ SẴN CAR)
            # =========================================================================
            else:
              wb_ro = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
              max_r = 1
              target_sheet_name = wb_ro.sheetnames[0]
              for s_name in wb_ro.sheetnames:
                m = re.search(r"round\s*(\d+)", s_name, re.IGNORECASE)
                if m and int(m.group(1)) >= max_r:
                  max_r = int(m.group(1))
                  target_sheet_name = s_name

              source_ws = wb_ro[target_sheet_name]
              car_p = re.search(
                  r"(\d{4,}\s*CaRs)", file.filename, re.IGNORECASE
              )
              car_base = car_p.group(1) if car_p else "CaRs"

              detected_doc_code = ""
              m_fdoc = re.search(r"(QT1-[A-Za-z0-9_-]+)", file.filename)
              if m_fdoc:
                detected_doc_code = m_fdoc.group(1)

              cover_info = cover_data["docs"].get(detected_doc_code, {})
              final_doc_no = cover_info.get(
                  "doc_no",
                  detected_doc_code if detected_doc_code else "QT1-Document",
              )
              final_rev = cover_info.get("rev", "1")
              final_title = cover_info.get("title", cover_data["subject"])

              status_full_text = "APPROVED (APP)"
              for c in range(1, source_ws.max_column + 1):
                val = str(source_ws.cell(row=11, column=c).value or "").upper()
                if f"ROUND {max_r} COMMENTS" in val:
                  st_val = str(
                      source_ws.cell(row=13, column=c).value or ""
                  ).upper()
                  if "RFC" in st_val:
                    status_full_text = "RETURN FOR CORRECTIONS (RFC)"
                  elif "AWC" in st_val:
                    status_full_text = "APPROVED WITH COMMENTS (AWC)"
                  elif "RJ" in st_val:
                    status_full_text = "REJECTED (RJ)"
                  break

              del wb_ro
              zip_file.writestr(clean_safe_filename(file.filename), content)

              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": status_full_text,
                  "car_label": f"{car_base} R{max_r}",
                  "assigned_car_name": file.filename,
              })

            del content
            gc.collect()

          except Exception as fe:
            print(f"Lỗi xử lý file {file.filename}: {fe}")
            traceback.print_exc()

      # =========================================================================
      # CHẾ ĐỘ: TRẢ LỜI THƯ, KHÔNG CÓ CAR
      # =========================================================================
      elif mode == "no_car_letter":
        if cover_data["docs"]:
          for doc_k, d_info in cover_data["docs"].items():
            processed_car_list.append({
                "doc_no": d_info["doc_no"],
                "rev": d_info["rev"],
                "title": d_info["title"],
                "status": "REVIEWED (NO CAR)",
                "car_label": "N/A",
                "assigned_car_name": "N/A (Chỉ thư Cover)",
            })
        else:
          processed_car_list.append({
              "doc_no": cover_data["tr_no"],
              "rev": "0",
              "title": cover_data["subject"],
              "status": "REVIEWED (NO CAR)",
              "car_label": "N/A",
              "assigned_car_name": "N/A (Chỉ thư Cover)",
          })

      safe_subj_name = clean_safe_filename(cover_data["subject"])

      # =========================================================================
      # XUẤT TỜ TRÌNH TIẾNG VIỆT
      # =========================================================================
      doc_vi = Document(vi_template_path)
      for tbl in doc_vi.tables:
        for row in tbl.rows:
          for cell in row.cells:
            for p in cell.paragraphs:
              if "ngày" in p.text and "tháng" in p.text:
                p.runs[0].text = (
                    f"Hà Nội, ngày {now.day} tháng {now.month} năm {now.year}"
                )
                for r in p.runs[1:]:
                  r.text = ""
              if "V/v" in p.text:
                replace_value_after_keyword(p, "V/v", cover_data["subject"])

      for p in doc_vi.paragraphs:
        if "Căn cứ" in p.text or "Công ty Cổ phần" in p.text:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if "Căn cứ công văn số" in p.text:
          target_sentence = (
              f"{cover_data['tr_no']} ngày {cover_data['date_str']} của Nhà"
              f" thầu EPC gói 15 về {cover_data['subject']}."
          )
          if not replace_value_after_keyword(
              p, "Căn cứ công văn số", target_sentence
          ):
            p.text = f"Căn cứ công văn số {target_sentence}"

      vi_io = io.BytesIO()
      doc_vi.save(vi_io)
      zip_file.writestr(f"TVÐ1-NÐNLM {safe_subj_name}.docx", vi_io.getvalue())
      del vi_io, doc_vi

      # =========================================================================
      # XUẤT THƯ TIẾNG ANH
      # =========================================================================
      doc_en = Document(en_template_path)
      doc_type_en = (
          "Letter" if "-L-" in cover_data["tr_no"].upper() else "Transmittal"
      )

      for tbl in doc_en.tables:
        for row in tbl.rows:
          for cell in row.cells:
            for p in cell.paragraphs:
              if "Ref. No.:" in p.text:
                replace_value_after_keyword(p, "Ref. No.:", cover_data["tr_no"])
              elif "October" in p.text or "202" in p.text:
                update_date_paragraph(p, current_date_en)
              elif "Identification Number." in p.text:
                replace_value_after_keyword(
                    p, "Identification Number.", qti_no.strip()
                )
              elif "Subject:" in p.text:
                replace_value_after_keyword(
                    p, "Subject:", cover_data["subject"]
                )

      sync_footer_content(doc_en, qti_no.strip(), cover_data["subject"])

      ref_idx = None
      yours_idx = None
      for idx, p in enumerate(doc_en.paragraphs):
        if "This has reference to the" in p.text:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
          font_name = (
              p.runs[0].font.name
              if (p.runs and p.runs[0].font.name)
              else "Times New Roman"
          )
          font_size = (
              p.runs[0].font.size if (p.runs and p.runs[0].font.size) else Pt(13)
          )
          p.clear()
          r = p.add_run(
              f"This has reference to the {doc_type_en} {cover_data['tr_no']},"
              " we would like to comment as follows:"
          )
          r.font.name = font_name
          r.font.size = font_size
          ref_idx = idx
        elif "Yours sincerely" in p.text:
          p.alignment = WD_ALIGN_PARAGRAPH.LEFT
          yours_idx = idx
          break

      if ref_idx is not None and yours_idx is not None:
        paras_to_remove = [
            doc_en.paragraphs[i] for i in range(ref_idx + 1, yours_idx)
        ]
        for p in paras_to_remove:
          p._element.getparent().remove(p._element)

      target_p = None
      for p in doc_en.paragraphs:
        if "Yours sincerely" in p.text:
          target_p = p
          break

      if target_p is not None:
        if mode == "no_car_letter":
          if cover_data["docs"]:
            for doc_k, d_info in reversed(list(cover_data["docs"].items())):
              new_p = target_p.insert_paragraph_before("")
              new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
              new_p.paragraph_format.left_indent = Inches(0.35)
              new_p.paragraph_format.first_line_indent = Inches(-0.35)
              new_p.paragraph_format.space_before = 0
              new_p.paragraph_format.space_after = Pt(4)

              run = new_p.add_run(
                  f"-\tThe document {d_info['doc_no']} {d_info['title']}"
                  f" Rev.{d_info['rev']} has been reviewed and found acceptable"
                  " with no further comments."
              )
              run.font.name = "Times New Roman"
              run.font.size = Pt(13)
          else:
            new_p = target_p.insert_paragraph_before("")
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.left_indent = Inches(0.35)
            new_p.paragraph_format.first_line_indent = Inches(-0.35)
            new_p.paragraph_format.space_before = 0
            new_p.paragraph_format.space_after = Pt(4)
            doc_ref_word = (
                "letter/proposal"
                if "-L-" in cover_data["tr_no"].upper()
                else "transmittal/documents"
            )
            run = new_p.add_run(
                f"-\tWe have reviewed the submitted {doc_ref_word} and have no"
                " objection."
            )
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
        else:
          for item in reversed(processed_car_list):
            new_p = target_p.insert_paragraph_before("")
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.left_indent = Inches(0.35)
            new_p.paragraph_format.first_line_indent = Inches(-0.35)
            new_p.paragraph_format.space_before = 0
            new_p.paragraph_format.space_after = Pt(4)

            run = new_p.add_run(
                f"-\tThe document {item['doc_no']} {item['title']}"
                f" Rev.{item['rev']} is {item['status']} indicated as per"
                f" {item['car_label']}."
            )
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)

      en_io = io.BytesIO()
      doc_en.save(en_io)
      safe_qti_name = clean_safe_filename(qti_no)
      zip_file.writestr(
          f"{safe_qti_name} {safe_subj_name}.docx", en_io.getvalue()
      )
      del en_io, doc_en

    safe_qti_name = clean_safe_filename(qti_no)
    headers = {
        "Content-Disposition": f'attachment; filename="{safe_qti_name}.zip"'
    }

    try:
      items_json = json.dumps(
          [{
              "doc_no": item["doc_no"],
              "assigned_car": item["assigned_car_name"],
              "status": item["status"],
          }
           for item in processed_car_list],
          ensure_ascii=False,
      )
      items_b64 = base64.b64encode(items_json.encode("utf-8")).decode("ascii")
      if len(items_b64) < 6000:
        headers["X-Processed-Items"] = items_b64
    except Exception:
      pass

    return FileResponse(
        path=temp_zip_path,
        media_type="application/zip",
        filename=f"{safe_qti_name}.zip",
        headers=headers,
        background=BackgroundTask(remove_temp_file, temp_zip_path),
    )

  except Exception as ge:
    remove_temp_file(temp_zip_path)
    print("Lỗi hệ thống:", ge)
    traceback.print_exc()
    return Response(content=f"Error: {str(ge)}", status_code=400)


if __name__ == "__main__":
  import uvicorn

  uvicorn.run(app, host="0.0.0.0", port=8000)
