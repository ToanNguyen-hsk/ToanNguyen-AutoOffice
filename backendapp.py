import base64
from copy import copy
from datetime import datetime, timedelta, timezone
import gc
import io
import json
import os
import re
import tempfile
import time
import traceback
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from fastapi import FastAPI, File, Form, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pypdf
from starlette.background import BackgroundTask

app = FastAPI(title="QT1 Doc Control API - Ultra Safe Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Processed-Items", "X-Cover-Info"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Máy chủ (Render) chạy giờ UTC -> ép về giờ Việt Nam để ngày phát hành không bị lệch
VN_TZ = timezone(timedelta(hours=7))


def clean_safe_filename(name: str) -> str:
  cleaned = re.sub(r'[\\/*?:"<>|]', "-", name).strip()
  return cleaned[:130] if len(cleaned) > 130 else cleaned


def normalize_doc_code(raw: str) -> str:
  if not raw:
    return ""
  m = re.search(r"(QT1-[0-9L]-[A-Za-z0-9_-]+)", raw)
  if not m:
    return ""
  code = m.group(1).strip()
  code = re.sub(
      r"_(?:[A-Za-z0-9]{1,3}|Rev\.?[A-Za-z0-9]+)$", "", code, flags=re.IGNORECASE
  )
  code = re.sub(r"[\s_]+rev\.?.*$", "", code, flags=re.IGNORECASE)
  return code.rstrip("-_")


def clean_document_title(raw_title: str, strip_index: bool = True) -> str:
  """strip_index=False khi dữ liệu đã tách đúng cột (PDF/Excel) -> không cắt nhầm
  số/chữ đầu tiêu đề như "33 kV ..." hay "V ..."."""
  if not raw_title:
    return ""
  t = str(raw_title).strip()
  if strip_index:
    t = re.sub(
        r"^(?:iv|v|vi|vii|viii|ix|x|\d{1,2})\s+", "", t, flags=re.IGNORECASE
    ).strip()
  t = re.sub(
      r"\s*\([A-Za-z0-9\s,.~-]*(?:APD|IFA|IFI|AB|RJ|Rev\.?|R\d)[A-Za-z0-9\s,.~-]*\)$",
      "",
      t,
      flags=re.IGNORECASE,
  ).strip()
  t = re.sub(
      r"VoltageTransformer", "Voltage Transformer", t, flags=re.IGNORECASE
  )
  if re.match(r"^(?:iv|v|vi|vii|viii|ix|x|\d{1,2})$", t, re.IGNORECASE):
    return ""
  return t


def extract_title_from_filename(filename: str) -> str:
  s = os.path.splitext(filename)[0].strip()
  s = re.sub(
      r"^(?:\d+|00xx|xxxx|Xxxx)\s*CaRs?\s*R\d+\s*(?:\d{1,2}\s*[A-Za-z]{3}\s*\d{4}\s*)?",
      "",
      s,
      flags=re.IGNORECASE,
  ).strip()
  s = re.sub(r"^(?:QT1-[A-Za-z0-9_-]+)\s*", "", s, flags=re.IGNORECASE).strip()
  s = re.sub(
      r"[-_](?:reply|comments?|rev\.?[A-Za-z0-9]+)$",
      "",
      s,
      flags=re.IGNORECASE,
  ).strip()
  return clean_document_title(s)


def extract_cover_date_bulletproof(text: str) -> str:
  """Đọc ngày từ một đoạn text bất kỳ. Thử MỌI vị trí chữ 'Date', không chỉ cái đầu tiên."""
  if not text:
    return ""
  for m_date in re.finditer(r"\bDate\b", text, re.IGNORECASE):
    d = _parse_date_snippet(text[m_date.end() : m_date.end() + 120])
    if d:
      return d
  return ""


# =============================================================================
# ĐỌC PDF COVER (TRANSMITTAL) THEO TOẠ ĐỘ: DÒNG -> CỘT -> Ô BẢNG
# =============================================================================
CODE_RE = re.compile(r"(QT1-[0-9L]-[A-Za-z0-9_-]+)")
# Rev hợp lệ: 1 chữ (+1 số) hoặc 1-2 số (+1 chữ), có thể kèm tiền tố "Rev."
# -> không còn "ăn nhầm" từ đầu của tiêu đề như CT, No, kV...
REV_RE = re.compile(r"^(?:Rev\.?)?([A-Za-z]\d?|\d{1,2}[A-Za-z]?)$", re.IGNORECASE)
PURPOSE_MARKERS = {"IFA", "IFI", "APD", "AB", "RI", "END", "-END-", "-"}
STOP_LINE_RE = re.compile(
    r"(Sheet\s*:|Date\s*:|Page\s+\d|END\s*-|-\s*END|Purpose\s*:|Prepared by|Remarks?\s*:)",
    re.IGNORECASE,
)
SUBJECT_END_RE = re.compile(
    r"^\W*(Reference|Ref\b|Dear|Attn|Attachment|Document|To\b|From\b|Cc\b|Date\b)",
    re.IGNORECASE,
)
MAX_PDF_PAGES = 30
MONTH3 = {
    m: i
    for i, m in enumerate(
        ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"],
        1,
    )
}
MONTH_PAT = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"


def _ok_dmy(d, m, y):
  return 1 <= d <= 31 and 1 <= m <= 12 and 2000 <= y <= 2099


def _parse_date_snippet(snippet: str) -> str:
  """Nhận một đoạn text ngay sau nhãn 'Date' và trả về dd/mm/yyyy (hoặc '')."""
  s = re.sub(r"Day\s*Mon(?:th)?\s*Year", " ", snippet or "", flags=re.IGNORECASE)
  s = re.sub(r"Sheet\s*:?\s*\d+\s*of\s*\d+", " ", s, flags=re.IGNORECASE)

  # 05/10/2026, 5-10-2026, 05.10.2026
  m = re.search(r"(?<!\d)(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(20\d\d)(?!\d)", s)
  if m:
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if _ok_dmy(d, mo, y):
      return f"{d:02d}/{mo:02d}/{y}"

  # Ô số rời: 0 5 1 0 2 0 2 6  hoặc  05 10 2026
  m = re.search(r"(?<!\d)(\d)\s*(\d)\s*(\d)\s*(\d)\s*(2)\s*(0)\s*(\d)\s*(\d)(?!\d)", s)
  if m:
    g = m.groups()
    d, mo, y = int(g[0] + g[1]), int(g[2] + g[3]), int("".join(g[4:]))
    if _ok_dmy(d, mo, y):
      return f"{d:02d}/{mo:02d}/{y}"

  # October 05, 2026
  m = re.search(
      r"\b(" + MONTH_PAT + r")\s*(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(20\s*\d\s*\d)",
      s,
      re.IGNORECASE,
  )
  if m:
    mo = MONTH3.get(m.group(1)[:3].lower(), 0)
    d, y = int(m.group(2)), int(re.sub(r"\s+", "", m.group(3)))
    if _ok_dmy(d, mo, y):
      return f"{d:02d}/{mo:02d}/{y}"

  # 05 Oct 2026
  m = re.search(
      r"(?<!\d)(\d{1,2})(?:st|nd|rd|th)?[\s\-/.]*(" + MONTH_PAT + r")[\s\-/.,]*(20\s*\d\s*\d)",
      s,
      re.IGNORECASE,
  )
  if m:
    mo = MONTH3.get(m.group(2)[:3].lower(), 0)
    d, y = int(m.group(1)), int(re.sub(r"\s+", "", m.group(3)))
    if _ok_dmy(d, mo, y):
      return f"{d:02d}/{mo:02d}/{y}"
  return ""


# ----------------------------- nạp trang PDF ---------------------------------
def _ocr_pages(pdf_bytes: bytes, max_pages: int = 4) -> list:
  """PDF scan (không có lớp chữ): OCR vài trang đầu nếu máy chủ có pytesseract."""
  try:
    import pdfplumber
    import pytesseract
  except Exception:
    return []
  out = []
  try:
    dpi = 200
    k = 72.0 / dpi
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
      for page in pdf.pages[:max_pages]:
        img = page.to_image(resolution=dpi).original
        d = pytesseract.image_to_data(
            img, lang="eng", config="--psm 6", output_type=pytesseract.Output.DICT
        )
        words = []
        for i, t in enumerate(d["text"]):
          t = (t or "").strip()
          if t and float(d["conf"][i]) >= 0:
            words.append({
                "t": t,
                "x0": d["left"][i] * k,
                "x1": (d["left"][i] + d["width"][i]) * k,
                "top": d["top"][i] * k,
                "bottom": (d["top"][i] + d["height"][i]) * k,
            })
        out.append({"words": words, "tables": [], "ocr": True})
  except Exception as e:
    print(f"OCR bỏ qua: {e}")
    return []
  return out


def _load_pages(pdf_bytes: bytes) -> list:
  """Mỗi trang: {'words': [{t,x0,x1,top,bottom}], 'tables': [[row,...]], 'ocr': bool}."""
  pages = []
  try:
    import pdfplumber

    t0 = time.monotonic()
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
      for page in pdf.pages[:MAX_PDF_PAGES]:
        words = [
            {"t": w["text"], "x0": w["x0"], "x1": w["x1"], "top": w["top"], "bottom": w["bottom"]}
            for w in page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
        ]
        tables = []
        if time.monotonic() - t0 < 25:  # chống treo với PDF quá nặng
          try:
            for tb in page.find_tables():
              rows = tb.extract()
              if rows:
                tables.append(rows)
          except Exception as te:
            print(f"find_tables bỏ qua: {te}")
        pages.append({"words": words, "tables": tables, "ocr": False})
  except Exception as e:
    print(f"pdfplumber không đọc được (dùng pypdf): {e}")
    pages = []
  if not any(p["words"] for p in pages):
    ocr = _ocr_pages(pdf_bytes)
    if ocr:
      return ocr
  return pages


def _pypdf_text(pdf_bytes: bytes) -> str:
  try:
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes), strict=False)
    return "\n".join((p.extract_text() or "") for p in reader.pages[:MAX_PDF_PAGES])
  except Exception as pe:
    print(f"pypdf reader caught warning/error (safely bypassed): {pe}")
    return ""


# ------------------------------ gom dòng -------------------------------------
def _finish_line(ln):
  ln["words"].sort(key=lambda w: w["x0"])
  ln["top"] = min(w["top"] for w in ln["words"])
  ln["bottom"] = max(w["bottom"] for w in ln["words"])
  ln["text"] = " ".join(w["t"] for w in ln["words"])
  return ln


def _group_lines(words: list) -> list:
  """Gom từ thành dòng theo toạ độ Y (tâm dòng), rồi sắp theo X."""
  items = sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"]))
  lines = []
  for w in items:
    cy = (w["top"] + w["bottom"]) / 2
    h = max(w["bottom"] - w["top"], 1)
    if lines and abs(cy - lines[-1]["cy"]) <= max(2.5, 0.45 * h):
      ln = lines[-1]
      ln["words"].append(w)
      ln["cy"] = sum((x["top"] + x["bottom"]) / 2 for x in ln["words"]) / len(ln["words"])
    else:
      lines.append({"words": [w], "cy": cy})
  return [_finish_line(ln) for ln in lines]


def _text_to_pseudo_lines(text: str) -> list:
  """Dùng khi chỉ có text phẳng (pypdf): tạo toạ độ giả để dùng chung thuật toán."""
  lines = []
  for idx, raw in enumerate(text.splitlines()):
    s = raw.strip()
    if not s:
      continue
    words, pos = [], 0
    for tok in s.split():
      words.append({
          "t": tok,
          "x0": pos * 5.0,
          "x1": (pos + len(tok)) * 5.0,
          "top": idx * 12.0,
          "bottom": idx * 12.0 + 10,
      })
      pos += len(tok) + 1
    lines.append(_finish_line({"words": words, "cy": idx * 12.0 + 5}))
  return lines


# ------------------------------ Ngày / Subject -------------------------------
def _find_date_in_lines(lines: list) -> str:
  """Tìm nhãn 'Date' rồi chỉ đọc vùng ngay bên phải và 2 dòng bên dưới nhãn đó."""
  for i, ln in enumerate(lines):
    for w in ln["words"]:
      if not re.fullmatch(r"date\s*:?", w["t"], re.IGNORECASE):
        continue
      parts = []
      for j in range(i, min(i + 3, len(lines))):
        L = lines[j]
        if j > i and L["top"] - lines[i]["bottom"] > 45:
          break
        if j == i:
          ws = [x for x in L["words"] if x["x0"] >= w["x1"] - 1]
        else:
          ws = [x for x in L["words"] if x["x1"] >= w["x0"] - 6]
        kept = []
        for x in ws:
          if re.match(r"sheet", x["t"], re.IGNORECASE):
            break  # không đọc sang ô 'Sheet: 1 of 2' bên cạnh
          kept.append(x["t"])
        parts.append(" ".join(kept))
        d = _parse_date_snippet(" ".join(parts))
        if d:
          return d
  return ""


def _find_subject_in_lines(lines: list) -> str:
  """Subject có thể xuống nhiều dòng: đọc tiếp các dòng cùng cột cho tới khi gặp nhãn khác."""
  for i, ln in enumerate(lines):
    for w in ln["words"]:
      if not re.fullmatch(r"subject\s*:?", w["t"], re.IGNORECASE):
        continue
      first = []
      for x in ln["words"]:
        if x["x0"] < w["x1"] - 1:
          continue
        if re.match(r"sheet", x["t"], re.IGNORECASE):
          break
        first.append(x)
      left = first[0]["x0"] if first else w["x0"]
      parts = [" ".join(x["t"] for x in first)]
      prev_bottom, lh = ln["bottom"], max(ln["bottom"] - ln["top"], 8)
      for j in range(i + 1, min(i + 6, len(lines))):
        L = lines[j]
        if L["top"] - prev_bottom > 1.8 * lh or SUBJECT_END_RE.match(L["text"]):
          break
        kept = []
        for x in L["words"]:
          if x["x0"] < left - 8:
            continue
          if re.match(r"sheet", x["t"], re.IGNORECASE):
            break
          kept.append(x["t"])
        if not kept:
          break
        parts.append(" ".join(kept))
        prev_bottom = L["bottom"]
      subj = re.split(
          r"(?:Reference|Dear|Attn|Attachment|Document/Drawing|Sheet:)",
          " ".join(parts),
          flags=re.IGNORECASE,
      )[0]
      subj = re.sub(r"\s+", " ", subj).strip(" :-")
      if subj:
        return subj
  return ""


# --------------------------- tách Rev / Tiêu đề ------------------------------
def _split_rev_title(tokens: list):
  """tokens = các từ đứng sau mã tài liệu. Trả về (rev, tokens_tiêu_đề)."""
  if not tokens:
    return "1", []
  if (
      len(tokens) >= 2
      and tokens[0].isdigit()
      and not tokens[1].isdigit()
      and REV_RE.match(tokens[1])
  ):
    return REV_RE.match(tokens[1]).group(1).upper(), tokens[2:]
  m = REV_RE.match(tokens[0])
  if m:
    return m.group(1).upper(), tokens[1:]
  return "1", tokens


def _rev_from_cell(txt: str) -> str:
  tokens = (txt or "").replace("\n", " ").split()
  if not tokens:
    return "1"
  if len(tokens) >= 2 and tokens[0].isdigit() and not tokens[-1].isdigit() and len(tokens[-1]) <= 3:
    tok = tokens[-1]
  else:
    tok = tokens[0]
  m = REV_RE.match(tok)
  if m:
    return m.group(1).upper()
  return re.sub(r"^Rev\.?", "", tok, flags=re.IGNORECASE)[:6] or "1"


# -------------------- Cách A: bảng có đường kẻ (chính xác nhất) --------------
def _flat_cell(s) -> str:
  s = (s or "").replace("\r", "")
  s = re.sub(r"(?<=[A-Za-z0-9])-\n(?=[A-Za-z0-9])", "-", s)
  s = re.sub(r"\s*\n\s*", " ", s)
  return re.sub(r"\s+", " ", s).strip()


def _code_from_cell(cell):
  s = re.sub(r"(?<=[-_])\s*\n\s*", "", cell or "")  # mã bị xuống dòng sau dấu - hoặc _
  s = _flat_cell(s)
  for m in CODE_RE.finditer(s):
    code = m.group(1)
    if "EVNPMB2" in code or "QT1-HDH" in code:
      continue
    return code, s
  return "", s


def _header_cols(row):
  cells = [re.sub(r"\s+", " ", (c or "")).strip().lower() for c in row]
  idx = {}
  for i, c in enumerate(cells):
    if not c:
      continue
    if "purpose" in c and "purpose" not in idx:
      idx["purpose"] = i
    elif re.match(r"^rev", c) and "rev" not in idx:
      idx["rev"] = i
    elif ("title" in c or "description" in c) and "title" not in idx:
      idx["title"] = i
    elif ("document" in c or "drawing" in c or c.startswith("doc")) and "doc" not in idx:
      idx["doc"] = i
  if "title" in idx and ("doc" in idx or "rev" in idx):
    idx.setdefault("doc", idx["title"])
    return idx
  return None


def _infer_cols(tbl):
  """Bảng không có dòng tiêu đề (trang nối tiếp): suy ra cột từ vị trí chứa mã tài liệu."""
  width = max(len(r) for r in tbl)
  counts = [0] * width
  for row in tbl:
    for i, c in enumerate(row):
      if _code_from_cell(c or "")[0]:
        counts[i] += 1
  if not any(counts):
    return None
  d = counts.index(max(counts))
  cols = {"doc": d, "title": d}
  if d + 1 < width:
    cols["rev"] = d + 1
  if d + 2 < width:
    cols["title"] = d + 2
  return cols


def _docs_from_tables(pages: list) -> dict:
  raw_docs = {}
  cols, ncols, last = None, 0, None
  for pg in pages:
    for tbl in pg.get("tables", []):
      width = max(len(r) for r in tbl)
      hdr_row = -1
      for ri, row in enumerate(tbl[:5]):
        found = _header_cols(row)
        if found:
          cols, ncols, hdr_row = found, len(row), ri
          break
      if hdr_row < 0 and (cols is None or width != ncols):
        cols, ncols = _infer_cols(tbl), width
      if cols is None:
        continue

      def cell(row, key):
        i = cols.get(key)
        return row[i] if i is not None and i < len(row) and row[i] else ""

      for row in tbl[hdr_row + 1 :]:
        code, doc_flat = _code_from_cell(cell(row, "doc"))
        if not code:
          # hàng nối tiếp: ô mã trống nhưng ô tiêu đề còn chữ -> nối vào tài liệu trước
          extra = _flat_cell(cell(row, "title"))
          if last and extra and not re.search(r"[A-Za-z]", doc_flat) and not _flat_cell(cell(row, "rev")):
            raw_docs[last]["raw"] += " " + extra
          continue
        norm = normalize_doc_code(code)
        if not norm:
          continue
        if cols["title"] == cols["doc"]:  # mã và tiêu đề chung 1 ô
          title = re.sub(r"^\s*\d{1,3}\s+", "", doc_flat.replace(code, " "))
        else:
          title = _flat_cell(cell(row, "title"))
        words = title.split()
        if "purpose" not in cols:
          while words and words[-1].upper() in PURPOSE_MARKERS:
            words.pop()
        entry = {
            "doc_no": norm,
            "rev": _rev_from_cell(_flat_cell(cell(row, "rev"))) if "rev" in cols else "1",
            "raw": " ".join(words),
        }
        if norm not in raw_docs or not raw_docs[norm]["raw"]:
          raw_docs[norm] = entry
        last = norm
  return _finalize_docs(raw_docs)


def _finalize_docs(raw_docs: dict) -> dict:
  docs = {}
  for k, v in raw_docs.items():
    docs[k] = {
        "doc_no": v["doc_no"],
        "rev": v["rev"],
        "title": clean_document_title(v["raw"], strip_index=False),
    }
  return docs


# -------------- Cách B: bảng không viền, đọc theo toạ độ dòng/cột --------------
def _is_table_header(ln) -> bool:
  toks = [w["t"].lower().strip(".:/,") for w in ln["words"]]
  has_rev = any(t in ("rev", "revision") for t in toks)
  has_title = any(("title" in t or "description" in t) for t in toks)
  return has_rev and has_title


def _is_end_line(ln) -> bool:
  return bool(re.search(r"END\s*-|-\s*END", ln["text"])) and not CODE_RE.search(ln["text"])


def _find_anchor(ln, purpose_cut):
  """Tìm từ chứa mã tài liệu trong dòng và tách Rev / phần đầu tiêu đề đứng cùng dòng."""
  for idx_w, w in enumerate(ln["words"]):
    m = CODE_RE.search(w["t"])
    if not m or "EVNPMB2" in m.group(1) or "QT1-HDH" in m.group(1):
      continue
    before = " ".join(x["t"].lower() for x in ln["words"][:idx_w])
    if re.search(r"\b(reference|ref\.?|attn|attachment)\b", before):
      continue  # mã nằm trong câu tham chiếu, không phải hàng của bảng
    raw_code = m.group(1)
    tokens = ln["words"][idx_w + 1 :]
    if raw_code.endswith(("-", "_")) and tokens and re.match(r"^[A-Za-z0-9]+$", tokens[0]["t"]):
      raw_code += tokens[0]["t"]  # mã bị tách làm 2 từ
      tokens = tokens[1:]
    if purpose_cut is not None:
      tokens = [t for t in tokens if t["x0"] < purpose_cut]
    while tokens and tokens[-1]["t"].upper() in PURPOSE_MARKERS:
      tokens = tokens[:-1]
    texts = [t["t"] for t in tokens]
    rev, rest = _split_rev_title(texts)
    consumed = len(texts) - len(rest)
    return {
        "code": raw_code,
        "rev": rev,
        "rest": rest,
        "rest_x0": tokens[consumed]["x0"] if rest else None,
        "rev_x1": tokens[consumed - 1]["x1"] if consumed else w["x1"],
        "line": ln,
    }
  return None


def _header_line_idx(lines):
  """Chỉ số dòng cuối của phần tiêu đề bảng (No | Document | Rev | Title | Purpose)."""
  for i, ln in enumerate(lines):
    if _is_table_header(ln):
      return i
  for i, ln in enumerate(lines):  # text phẳng của pypdf: mỗi ô nằm 1 dòng riêng
    if ln["text"].strip().lower().strip(".:") in ("rev", "revision"):
      window = [l["text"].strip().lower().strip(".:/") for l in lines[i : i + 5]]
      if any(("title" in w or "description" in w) for w in window):
        last = i
        for k, w in enumerate(window):
          if "title" in w or "description" in w or "purpose" in w:
            last = i + k
        return last
  return None


def _docs_from_layout(lines_by_page: list, column_aware: bool = True) -> dict:
  """Bảng không viền. Mỗi hàng = 1 dòng chứa mã tài liệu + các dòng tiêu đề nối tiếp.
  Hỗ trợ cả bảng căn trên (tiêu đề ở dưới mã) lẫn căn giữa (tiêu đề nằm trên và dưới mã)."""
  header_exists = any(_header_line_idx(lines) is not None for lines in lines_by_page)
  started = not header_exists
  purpose_cut = None

  # ---- Bước 1: tìm dòng-mã (anchor) và các dòng "mồ côi" (chưa biết thuộc hàng nào)
  pages = []
  for lines in lines_by_page:
    a_from, a_to = 0, len(lines)
    if header_exists:
      h = _header_line_idx(lines)
      if h is not None:
        started = True
        for hl in lines[max(0, h - 2) : h + 3]:
          for w in hl["words"]:
            if column_aware and w["t"].lower().startswith("purpose"):
              purpose_cut = w["x0"] - 15
        a_from = h + 1
      elif not started:
        continue
    for i in range(a_from, len(lines)):
      if _is_end_line(lines[i]):
        a_to = i
        if header_exists:
          started = False
        break
    anchors, orphans = [], []
    for i in range(a_from, a_to):
      ln = lines[i]
      a = _find_anchor(ln, purpose_cut)
      if a:
        a["i"] = i
        anchors.append(a)
      elif STOP_LINE_RE.search(ln["text"]) or _is_table_header(ln):
        continue
      elif not column_aware and (
          ln["text"].strip().upper() in PURPOSE_MARKERS
          # số đứng riêng 1 dòng ngay trước dòng có mã = STT của hàng kế (còn Rev "0"/"1" thì giữ)
          or (
              re.fullmatch(r"\d{1,3}", ln["text"].strip())
              and i + 1 < a_to
              and _find_anchor(lines[i + 1], purpose_cut)
          )
      ):
        continue
      else:
        orphans.append(i)
    pages.append({"lines": lines, "anchors": anchors, "orphans": orphans})

  all_anchors = [a for p in pages for a in p["anchors"]]
  if not all_anchors:
    return {}

  # ---- Bước 2: xác định lề trái của cột Tiêu đề từ dữ liệu thật
  title_left = None
  if column_aware:
    xs = sorted(a["rest_x0"] for a in all_anchors if a["rest_x0"] is not None)
    if xs:
      title_left = xs[len(xs) // 2]
    else:
      rx = sorted(a["rev_x1"] for a in all_anchors)
      title_left = rx[len(rx) // 2] + 1

  def eligible(ln):
    if not column_aware:
      return list(ln["words"])
    return [
        w
        for w in ln["words"]
        if w["x0"] >= title_left - 6 and not (purpose_cut is not None and w["x0"] >= purpose_cut)
    ]

  def limit_of(ln):
    return max(1.6 * max(ln["bottom"] - ln["top"], 8), 14) if column_aware else float("inf")

  # ---- Bước 3: bảng căn giữa? (mã nằm giữa, tiêu đề tràn lên dòng trên/dưới mã)
  center_mode = False
  if column_aware:
    for p in pages:
      for a in p["anchors"]:
        if a["rest"]:
          continue
        ln = a["line"]
        lim = limit_of(ln)
        for oi in p["orphans"]:
          o = p["lines"][oi]
          if not eligible(o):
            continue
          # so khoảng cách TÂM dòng (hai dòng căn giữa có thể chồng nhau theo chiều dọc)
          if abs(o["cy"] - ln["cy"]) <= lim + max(ln["bottom"] - ln["top"], 8):
            center_mode = True

  # ---- Bước 4: gán dòng mồ côi cho hàng gần nhất rồi kiểm tra tính liền mạch
  raw_docs = {}
  for p in pages:
    lines, anchors = p["lines"], p["anchors"]
    owned = {a["i"]: [] for a in anchors}
    for oi in p["orphans"]:
      o = lines[oi]
      if not eligible(o):
        continue
      prev_a = next((a for a in reversed(anchors) if a["i"] < oi), None)
      next_a = next((a for a in anchors if a["i"] > oi), None)
      if center_mode:
        cands = [a for a in (prev_a, next_a) if a]
        owner = min(cands, key=lambda a: abs(a["line"]["cy"] - o["cy"])) if cands else None
      else:
        owner = prev_a
      if owner is not None:
        owned[owner["i"]].append(oi)

    for a in anchors:
      ln = a["line"]
      lim = limit_of(ln)
      below, above = [], []
      prev_bottom = ln["bottom"]
      for oi in sorted(x for x in owned[a["i"]] if x > a["i"]):
        o = lines[oi]
        if o["top"] - prev_bottom > lim:
          break
        below.append(oi)
        prev_bottom = o["bottom"]
      prev_top = ln["top"]
      for oi in sorted((x for x in owned[a["i"]] if x < a["i"]), reverse=True):
        o = lines[oi]
        if prev_top - o["bottom"] > lim:
          break
        above.insert(0, oi)
        prev_top = o["top"]

      def line_text(oi):
        t2 = [w["t"] for w in eligible(lines[oi])]
        if (
            not column_aware
            and len(t2) >= 2
            and t2[-1].isdigit()
            and len(t2[-1]) <= 3
            and t2[-2].upper() in PURPOSE_MARKERS
        ):
          t2 = t2[:-1]  # pypdf dính số thứ tự của hàng kế tiếp vào cuối dòng
        while len(t2) > 1 and t2[-1].upper() in PURPOSE_MARKERS:
          t2.pop()
        return " ".join(t2)

      rev = a["rev"]
      below_texts = [line_text(oi) for oi in below]
      if not column_aware and not a["rest"] and below_texts:
        rev, rest2 = _split_rev_title(below_texts[0].split())  # pypdf đẩy Rev xuống dòng sau mã
        below_texts[0] = " ".join(rest2)
      parts = [line_text(oi) for oi in above] + [" ".join(a["rest"])] + below_texts
      norm = normalize_doc_code(a["code"])
      if norm and (norm not in raw_docs or not raw_docs[norm]["raw"]):
        raw_docs[norm] = {"doc_no": norm, "rev": rev, "raw": " ".join(x for x in parts if x)}
  return _finalize_docs(raw_docs)


# ----------------------------- Hàm chính -------------------------------------
def extract_cover_metadata_fast(pdf_bytes: bytes, filename: str = "") -> dict:
  """Đọc Cover PDF theo toạ độ (dòng, cột, ô bảng) thay vì giải nén thô từng stream."""
  pages = _load_pages(pdf_bytes)
  has_words = any(p["words"] for p in pages)
  lines_by_page = [_group_lines(p["words"]) for p in pages if p["words"]] if has_words else []
  is_ocr = bool(pages and pages[0].get("ocr"))

  if has_words:
    full_text = "\n".join(l["text"] for ls in lines_by_page for l in ls)
  else:
    full_text = _pypdf_text(pdf_bytes)

  # 1. Tr_no
  tr_no = ""
  m_fn = re.search(r"(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)", filename)
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

  # 2. Date
  head_lines = [l for ls in lines_by_page[:2] for l in ls]
  date_str = _find_date_in_lines(head_lines) or extract_cover_date_bulletproof(full_text)

  # 3. Subject
  subject = _find_subject_in_lines(head_lines)
  if not subject:
    m_sub = re.search(r"(?:SUBJECT|Subject)\s*[:\s]+([^\n\r]+)", full_text, re.IGNORECASE)
    if m_sub:
      subject = re.split(
          r"(?:Reference|Dear|Attn|Attachment|Document/Drawing|Sheet:)",
          m_sub.group(1).strip(),
          flags=re.IGNORECASE,
      )[0].strip()
  if not subject:
    subject = "Official Transmittal"

  # 4. Danh mục tài liệu: bảng có viền -> bảng không viền -> text phẳng
  docs, method = {}, "none"
  if any(p.get("tables") for p in pages):
    docs = _docs_from_tables(pages)
    method = "table"
  if not docs and lines_by_page:
    docs = _docs_from_layout(lines_by_page, column_aware=True)
    method = "layout"
  if not docs:
    docs = _docs_from_layout([_text_to_pseudo_lines(full_text)], column_aware=False)
    method = "text"
  if not docs:
    method = "none"

  return {
      "tr_no": tr_no,
      "date_str": date_str if date_str else datetime.now(VN_TZ).strftime("%d/%m/%Y"),
      "date_found": bool(date_str),
      "subject": subject,
      "docs": docs,
      "parse_info": {"method": method, "ocr": is_ocr, "pages": len(pages)},
  }


def _cover_from_excel(cover_bytes: bytes, filename: str, cover_data: dict) -> dict:
  """Đọc Cover dạng .xlsx: mã tài liệu, Rev, tiêu đề, Subject, Date, số Transmittal."""
  wb_t = openpyxl.load_workbook(io.BytesIO(cover_bytes), data_only=True)
  ws_t = wb_t.active
  grid = [list(r) for r in ws_t.iter_rows(values_only=True)]

  def sv(v):
    if v is None:
      return ""
    if isinstance(v, float) and v.is_integer():
      v = int(v)  # Rev 0 / 1 không bị thành "0.0" hay bị `or "1"` nuốt mất
    return str(v).strip()

  m_fn = re.search(r"(QT1-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+)", filename)
  if m_fn:
    cover_data["tr_no"] = m_fn.group(1)

  def value_of_label(r, c, val):
    if ":" in val and val.split(":", 1)[1].strip():
      return val.split(":", 1)[1].strip()
    for x in grid[r][c + 1 :]:
      if sv(x):
        return x
    if r + 1 < len(grid) and c < len(grid[r + 1]) and sv(grid[r + 1][c]):
      return grid[r + 1][c]
    return ""

  for r, row in enumerate(grid):
    for c, raw in enumerate(row):
      val = sv(raw)
      if not val:
        continue
      low = val.lower()
      if re.match(r"^subject\s*:?", low) and cover_data["subject"] == "Official Transmittal":
        v = sv(value_of_label(r, c, val))
        if v:
          cover_data["subject"] = v
      elif re.match(r"^date\s*:?$", low) or low.startswith("date:"):
        v = value_of_label(r, c, val)
        d = v.strftime("%d/%m/%Y") if isinstance(v, datetime) else _parse_date_snippet(sv(v))
        if d and not cover_data.get("date_found"):
          cover_data["date_str"], cover_data["date_found"] = d, True
      m = re.match(r"^(QT1-[0-9L]-[A-Za-z0-9_-]+)$", val)
      if m and "QT1-HDH" not in val and "EVNPMB2" not in val:
        doc_no = normalize_doc_code(m.group(1))
        vals = [sv(x) for x in row[c + 1 :] if sv(x)]
        rev, rest = _split_rev_title(vals)
        title = clean_document_title(rest[0] if rest else "", strip_index=False)
        if doc_no and (doc_no not in cover_data["docs"] or not cover_data["docs"][doc_no]["title"]):
          cover_data["docs"][doc_no] = {"doc_no": doc_no, "rev": rev, "title": title}
  return cover_data


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


def find_car_template():
  search_dirs = [BASE_DIR, os.getcwd(), r"G:\Shared drives\Test auto"]
  for d in search_dirs:
    if not os.path.exists(d):
      continue
    for f in os.listdir(d):
      if f.lower().endswith(".xlsx") and not f.startswith("~$"):
        if any(k in f.lower() for k in ["template", "car", "00xx", "xxxx"]):
          return os.path.join(d, f)
  return None


def create_fresh_car_wb():
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "ROUND 1 CaRs"
  ws.views.sheetView[0].showGridLines = True
  col_widths = {
      "A": 4.5,
      "B": 1.5,
      "C": 16.0,
      "D": 1.5,
      "E": 20.0,
      "F": 1.5,
      "G": 48.0,
      "H": 1.5,
      "I": 58.0,
      "J": 1.5,
      "K": 4.5,
      "L": 14.0,
  }
  for col, width in col_widths.items():
    ws.column_dimensions[col].width = width
  dot = Side(border_style="dotted", color="808080")
  dotted_box = Border(top=dot, left=dot, right=dot, bottom=dot)
  f_title = Font(name="Times New Roman", size=16, bold=True)
  f_car_no = Font(name="Times New Roman", size=16, bold=True, color="FFFF0000")
  f_label, f_header = Font(name="Arial", size=10, bold=False), Font(
      name="Arial", size=10, bold=True
  )
  f_status = Font(name="Arial", size=10, bold=True, color="FFFF0000")
  al_center, al_left = Alignment(
      horizontal="center", vertical="center"
  ), Alignment(horizontal="left", vertical="center")

  ws.merge_cells("A1:D1")
  ws["A1"], ws["A1"].font, ws["A1"].alignment = "CaR Nr.", f_title, al_left
  ws["E1"].font, ws["E1"].alignment = f_car_no, al_left

  labels = [
      (2, "i", "Sender of Document"),
      (3, "ii", "Covering Letter No & Date"),
      (4, "iii", "Doc Ident No & Rev"),
      (5, "iv", "Title of Document"),
      (6, "v", "Date Doc Rec'd by EVN/Consultant"),
      (7, "vi", "Reviewed by & date completed"),
      (8, "vii", "Date & No CaRs is issued to PMB2"),
      (9, "viii", "Date & No CaRs is issued to Contractor"),
  ]
  for r, num_rom, text in labels:
    ws.cell(row=r, column=1, value=num_rom).font, ws.cell(
        row=r, column=1
    ).alignment = f_label, al_center
    ws.merge_cells(f"C{r}:E{r}")
    ws.cell(row=r, column=3, value=text).font, ws.cell(
        row=r, column=3
    ).alignment = f_label, al_left
    ws.cell(row=r, column=7).font, ws.cell(row=r, column=7).alignment = (
        f_header,
        al_left,
    )

  ws.cell(row=10, column=3, value="Contract /").font = f_header
  ws.cell(row=10, column=5, value="Doc Cls/Page").font = f_header
  ws.cell(row=10, column=9, value="CONTRACTOR's").font = f_header
  ws.cell(row=10, column=9).alignment = al_left

  headers = [
      (1, "No"),
      (3, "Spec Ref"),
      (5, "Drg Ref etc"),
      (7, "ROUND 1 COMMENTS"),
      (9, "ROUND 1 REPLIES"),
  ]
  for c, h_text in headers:
    ws.cell(row=11, column=c, value=h_text).font = f_header
    ws.cell(row=11, column=c).alignment = al_center if c == 1 else al_left

  ws.cell(row=13, column=7, value="STATUS AFTER ROUND 1 COMMENTS:").font = (
      f_status
  )
  ws.cell(row=14, column=7, value="APPROVED (APP)").font = f_status
  ws.cell(row=15, column=7, value="General Comments").font = f_header
  ws.cell(row=16, column=1, value=1).font, ws.cell(
      row=16, column=1
  ).alignment = f_label, al_center
  ws.cell(row=16, column=7, value="No comment.").font = f_label

  for r in range(17, 27):
    ws.cell(row=r, column=1, value=r - 15).font, ws.cell(
        row=r, column=1
    ).alignment = f_label, al_center
  ws.cell(row=27, column=7, value="Detailed Comments").font = f_header
  for r in range(28, 36):
    ws.cell(row=r, column=1, value=r - 27).font, ws.cell(
        row=r, column=1
    ).alignment = f_label, al_center
  for r in range(1, 36):
    for c in range(1, 13):
      if r >= 11:
        ws.cell(row=r, column=c).border = dotted_box
  return wb


def replace_value_after_keyword(paragraph, keyword, new_text):
  if keyword in paragraph.text:
    if not paragraph.runs:
      paragraph.text = f"{keyword} {new_text}"
      return True
    kw_font = paragraph.runs[0].font
    kw_name, kw_size, kw_bold, kw_italic = (
        kw_font.name,
        kw_font.size,
        kw_font.bold,
        kw_font.italic,
    )
    val_font = paragraph.runs[-1].font
    val_name, val_size, val_bold, val_italic = (
        val_font.name,
        val_font.size,
        val_font.bold,
        val_font.italic,
    )

    paragraph.clear()
    r1 = paragraph.add_run(keyword)
    r1.font.name, r1.font.size, r1.font.bold, r1.font.italic = (
        kw_name,
        kw_size,
        kw_bold,
        kw_italic,
    )
    r2 = paragraph.add_run(" " + new_text)
    r2.font.name, r2.font.size, r2.font.bold, r2.font.italic = (
        val_name,
        val_size,
        val_bold,
        val_italic,
    )
    return True
  return False


def update_date_paragraph(paragraph, new_date_str):
  if not paragraph.runs:
    paragraph.add_run("                                " + new_date_str)
    return
  val_font = paragraph.runs[-1].font
  val_name, val_size, val_bold, val_italic = (
      val_font.name,
      val_font.size,
      val_font.bold,
      val_font.italic,
  )
  paragraph.clear()
  paragraph.add_run("                                ")
  r2 = paragraph.add_run(new_date_str)
  r2.font.name, r2.font.size, r2.font.bold, r2.font.italic = (
      val_name,
      val_size,
      val_bold,
      val_italic,
  )


def sync_footer_content(doc, qti_no, subject):
  new_text = f"{qti_no} {subject}"
  for section in doc.sections:
    if section.footer:
      for p in section.footer.paragraphs:
        if p.text.strip():
          if len(p.runs) > 0:
            f_font = p.runs[0].font
            p.clear()
            r = p.add_run(new_text)
            r.font.name, r.font.size = f_font.name, f_font.size
            r.font.color.rgb = (
                f_font.color.rgb if f_font.color else RGBColor(128, 128, 128)
            )
          else:
            p.text = new_text


def remove_temp_file(path: str):
  try:
    if os.path.exists(path):
      os.remove(path)
  except:
    pass


@app.post("/api/export-zip-full")
async def export_zip_full(
    files: list[UploadFile] = File(default=[]),
    cover_file: UploadFile = File(None),
    metadata: str = Form("{}"),
    mode: str = Form("auto_bump"),
    qti_no: str = Form("QTI-PECC1-PMB2-26-10-155"),
    start_car_no: str = Form("9905"),
    custom_date: str = Form(""),
    default_status: str = Form("APP"),
):
  temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
  temp_zip_path = temp_zip.name
  temp_zip.close()

  try:
    car_metadata = json.loads(metadata) if metadata else {}
    now = datetime.now(VN_TZ)
    warnings = []
    current_date_str = now.strftime("%d %b %Y")
    current_date_en = now.strftime("%B %d, %Y")

    cover_data = {
        "tr_no": "QT1-Document",
        "date_str": now.strftime("%d/%m/%Y"),
        "date_found": False,
        "subject": "Official Transmittal",
        "docs": {},
        "parse_info": {"method": "none", "ocr": False, "pages": 0},
    }

    if cover_file is not None:
      cover_bytes = await cover_file.read()
      if cover_file.filename.lower().endswith(".pdf"):
        cover_data = extract_cover_metadata_fast(
            cover_bytes, cover_file.filename
        )
      elif cover_file.filename.lower().endswith((".xlsx", ".xlsm")):
        cover_data = _cover_from_excel(
            cover_bytes, cover_file.filename, cover_data
        )
        cover_data["parse_info"] = {"method": "excel", "ocr": False, "pages": 1}
      else:
        raise ValueError(
            "Thư Cover chỉ hỗ trợ .pdf hoặc .xlsx (tệp .xls cũ cần lưu lại thành .xlsx)."
        )
      del cover_bytes

    if custom_date and custom_date.strip():
      cover_data["date_str"] = custom_date.strip()
    elif cover_file is not None and not cover_data.get("date_found"):
      warnings.append(
          "Không đọc được ngày trên Cover, đang dùng ngày hôm nay - hãy nhập tay"
          " ở ô 'Ngày phát hành công văn'."
      )
    if cover_file is not None and not cover_data["docs"]:
      warnings.append(
          "Không đọc được danh mục tài liệu từ Cover (Cover scan ảnh hoặc bố cục"
          " lạ) - hãy kiểm tra lại các thư Word xuất ra."
      )

    vi_template_path = find_template("TV")
    en_template_path = find_template("QTI-PECC1")
    if not vi_template_path or not en_template_path:
      raise FileNotFoundError(
          "Thiếu file Word mẫu (TVÐ1-NÐNLM.docx / QTI-PECC1.docx) trong máy"
          " chủ."
      )

    processed_car_list = []
    start_num = int(re.sub(r"\D", "", str(start_car_no)) or "9905")

    with zipfile.ZipFile(temp_zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:

      # 1. TỰ ĐỘNG TẠO CAR ROUND 1
      if mode == "auto_generate_cars":
        docs_to_generate = (
            list(cover_data["docs"].values())
            if cover_data["docs"]
            else [{
                "doc_no": cover_data["tr_no"],
                "rev": "A",
                "title": cover_data["subject"],
            }]
        )
        status_full = (
            "APPROVED (APP)"
            if default_status == "APP"
            else "RETURN FOR CORRECTIONS (RFC)"
        )
        if default_status == "AWC":
          status_full = "APPROVED WITH COMMENTS (AWC)"
        elif default_status == "RJ":
          status_full = "REJECTED (RJ)"

        car_template_path = find_car_template()

        for doc_idx, doc_item in enumerate(docs_to_generate):
          cur_num = start_num + doc_idx
          car_num_str = f"{cur_num:04d}" if cur_num < 10000 else f"{cur_num}"
          car_base = f"{car_num_str} CaRs"

          wb = (
              openpyxl.load_workbook(car_template_path)
              if car_template_path
              else create_fresh_car_wb()
          )
          for ws_item in wb.worksheets:
            if (
                hasattr(ws_item, "sheet_properties")
                and ws_item.sheet_properties
                and hasattr(ws_item.sheet_properties, "pageSetUpPr")
            ):
              ws_item.sheet_properties.pageSetUpPr.fitToPage = None

          ws = (
              wb["ROUND 1 CaRs"]
              if "ROUND 1 CaRs" in wb.sheetnames
              else wb.active
          )
          ws["E1"].value = f"{car_num_str}-R1"
          ws["E1"].font = Font(
              name="Times New Roman", size=16, bold=True, color="FFFF0000"
          )
          ws["E1"].fill = PatternFill(fill_type=None)

          ws["G2"].value = "HYUNDAI - MITSUBISHI - CC1"
          ws["G3"].value = cover_data["tr_no"]
          ws["G4"].value = f"{doc_item['doc_no']} rev.{doc_item['rev']}"
          ws["G5"].value = doc_item['title']
          ws["G6"].value = cover_data["date_str"]

          ws["G13"].value = "STATUS AFTER ROUND 1 COMMENTS:"
          ws["G14"].value = status_full
          if default_status == "APP":
            ws["G16"].value = "No comment."

          new_excel_name = (
              f"{car_base} R1 {current_date_str} {doc_item['doc_no']}"
              f" {clean_safe_filename(doc_item['title'])}"
              f" rev.{doc_item['rev']}.xlsx"
          )
          excel_out = io.BytesIO()
          wb.save(excel_out)
          zip_file.writestr(new_excel_name, excel_out.getvalue())
          del wb, excel_out

          processed_car_list.append({
              "doc_no": doc_item["doc_no"],
              "rev": doc_item["rev"],
              "title": doc_item["title"],
              "status": status_full,
              "car_label": f"{car_base} R1",
              "assigned_car_name": new_excel_name,
          })

      # 2. XỬ LÝ NHANH HÀNG LOẠT TỆP CAR ĐƯỢC TẢI LÊN (BẢO VỆ TIÊU ĐỀ 3 CẤP)
      elif mode != "no_car_letter" and files:
        for file_idx, file in enumerate(files):
          try:
            content = await file.read()
            wb = openpyxl.load_workbook(io.BytesIO(content))

            for ws_item in wb.worksheets:
              if (
                  hasattr(ws_item, "sheet_properties")
                  and ws_item.sheet_properties
                  and hasattr(ws_item.sheet_properties, "pageSetUpPr")
              ):
                ws_item.sheet_properties.pageSetUpPr.fitToPage = None

            max_r = 1
            target_idx = 0
            for idx, s_name in enumerate(wb.sheetnames):
              m = re.search(r"round\s*(\d+)", s_name, re.IGNORECASE)
              if m and int(m.group(1)) >= max_r:
                max_r = int(m.group(1))
                target_idx = idx

            source_ws = wb.worksheets[target_idx]

            prev_com_col, prev_rep_col = None, None
            for c in range(1, source_ws.max_column + 1):
              val_h = str(source_ws.cell(row=11, column=c).value or "").upper()
              if f"ROUND {max_r} COMMENTS" in val_h:
                prev_com_col = c
              elif f"ROUND {max_r} REPLIES" in val_h:
                prev_rep_col = c
            if not prev_com_col:
              prev_com_col = 7

            # Xác định mã Doc No
            detected_doc_code = ""
            for c in range(3, source_ws.max_column + 1):
              v_c = str(source_ws.cell(row=4, column=c).value or "")
              if "QT1-" in v_c:
                detected_doc_code = normalize_doc_code(v_c)
                break
            if not detected_doc_code:
              detected_doc_code = normalize_doc_code(file.filename)

            cover_info = cover_data["docs"].get(detected_doc_code, {})
            final_doc_no = cover_info.get(
                "doc_no",
                detected_doc_code if detected_doc_code else "QT1-Document",
            )
            final_rev = cover_info.get("rev", "1")

            # BẬC THANG BẢO VỆ TIÊU ĐỀ 3 CẤP
            # Cấp 1: Bảng Cover PDF
            if (
                cover_info.get("title")
                and cover_info["title"].lower() != cover_data["subject"].lower()
            ):
              final_title = cover_info["title"]
            else:
              # Cấp 2: Tên file CaR tải lên
              title_from_fn = extract_title_from_filename(file.filename)
              if title_from_fn:
                final_title = title_from_fn
              else:
                # Cấp 3: Ô Hàng 5 Round trước đã làm sạch rác (APD, Rev.0)
                r5_val = clean_document_title(
                    str(
                        source_ws.cell(row=5, column=prev_com_col).value
                        or source_ws.cell(row=5, column=7).value
                        or ""
                    )
                )
                final_title = (
                    r5_val if r5_val else "C&I Cable Route Plan for CPP Area"
                )

            cur_num = start_num + file_idx
            car_num_str = f"{cur_num:04d}" if cur_num < 10000 else f"{cur_num}"
            car_base = f"{car_num_str} CaRs"

            # -------------------------------------------------------------
            # CHẾ ĐỘ 2: CẤP SỐ CHO CAR KỸ SƯ
            # -------------------------------------------------------------
            if mode == "number_filled_cars":
              detected_status_raw = "APPROVED (APP)"
              found_rfc, found_awc, found_rj = False, False, False
              for r in range(11, 16):
                v_st = str(
                    source_ws.cell(row=r, column=prev_com_col).value or ""
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

              found_cell_r1 = False
              for c in range(2, 10):
                val_r1 = str(source_ws.cell(row=1, column=c).value or "").strip()
                if re.search(r"(00xx|xxxx|Xxxx)", val_r1, re.IGNORECASE):
                  source_ws.cell(row=1, column=c).value = (
                      f"{car_base} - R{max_r}"
                  )
                  source_ws.cell(row=1, column=c).font = Font(
                      name="Times New Roman",
                      size=16,
                      bold=True,
                      color="FFFF0000",
                  )
                  source_ws.cell(row=1, column=c).fill = PatternFill(
                      fill_type=None
                  )
                  found_cell_r1 = True
                  break
              if not found_cell_r1:
                source_ws.cell(row=1, column=5).value = f"{car_base} - R{max_r}"
                source_ws.cell(row=1, column=5).font = Font(
                    name="Times New Roman", size=16, bold=True, color="FFFF0000"
                )
                source_ws.cell(row=1, column=5).fill = PatternFill(
                    fill_type=None
                )

              source_ws.cell(row=4, column=prev_com_col).value = (
                  f"{final_doc_no} rev.{final_rev}"
              )
              source_ws.cell(row=5, column=prev_com_col).value = final_title

              new_excel_name = (
                  f"{car_base} R{max_r} {current_date_str} {final_doc_no}"
                  f" {clean_safe_filename(final_title)}"
                  f" rev.{final_rev}.xlsx"
              )
              excel_out = io.BytesIO()
              wb.save(excel_out)
              zip_file.writestr(new_excel_name, excel_out.getvalue())
              del wb, excel_out

              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": detected_status_raw,
                  "car_label": f"{car_base} R{max_r}",
                  "assigned_car_name": new_excel_name,
              })

            # -------------------------------------------------------------
            # CHẾ ĐỘ 3: NÂNG ROUND TIẾP THEO (R+1)
            # -------------------------------------------------------------
            elif mode == "auto_bump":
              selected_status = car_metadata.get(file.filename, "APP")
              new_r_num = max_r + 1
              new_ws = wb.copy_worksheet(source_ws)
              new_ws.title = f"ROUND {new_r_num} CaRs"
              wb._sheets.remove(new_ws)
              wb._sheets.insert(target_idx + 1, new_ws)

              new_com_col = (
                  (prev_rep_col + 1)
                  if prev_rep_col
                  else (new_ws.max_column + 1)
              )
              new_rep_col = new_com_col + 2

              for c in range(2, 10):
                v = str(new_ws.cell(row=1, column=c).value or "")
                if re.search(r"-\s*R\d+", v, re.IGNORECASE) or (c == 5 and v):
                  new_ws.cell(row=1, column=c).value = re.sub(
                      r"R\d+", f"R{new_r_num}", v, flags=re.IGNORECASE
                  )
                  break

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

              status_full = (
                  "APPROVED (APP)"
                  if selected_status == "APP"
                  else "RETURN FOR CORRECTIONS (RFC)"
              )
              if selected_status == "AWC":
                status_full = "APPROVED WITH COMMENTS (AWC)"
              elif selected_status == "RJ":
                status_full = "REJECTED (RJ)"

              new_ws.cell(row=12, column=new_com_col).value = (
                  f"STATUS AFTER ROUND {new_r_num} COMMENTS:"
              )
              new_ws.cell(row=13, column=new_com_col).value = status_full
              red_bold = Font(
                  name="Arial", size=10, bold=True, color="FFFF0000"
              )
              new_ws.cell(row=12, column=new_com_col).font, (
                  new_ws.cell(row=13, column=new_com_col).font
              ) = (red_bold, red_bold)
              new_ws.column_dimensions[
                  openpyxl.utils.get_column_letter(new_com_col)
              ].width = 38
              new_ws.column_dimensions[
                  openpyxl.utils.get_column_letter(new_rep_col)
              ].width = 38

              limit_row = min(new_ws.max_row, 100)
              for r in range(14, limit_row + 1):
                v_prev_c = str(
                    new_ws.cell(row=r, column=prev_com_col).value or ""
                ).strip()
                if v_prev_c and not (
                    "general comment" in v_prev_c.lower()
                    or "detailed comment" in v_prev_c.lower()
                ):
                  if selected_status == "APP":
                    new_ws.cell(row=r, column=new_com_col).value = "Noted."

              new_excel_name = (
                  f"{car_base} R{new_r_num} {current_date_str} {final_doc_no}"
                  f" {clean_safe_filename(final_title)}"
                  f" rev.{final_rev}.xlsx"
              )
              excel_out = io.BytesIO()
              wb.save(excel_out)
              zip_file.writestr(new_excel_name, excel_out.getvalue())
              del wb, excel_out

              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": status_full,
                  "car_label": f"{car_base} R{new_r_num}",
                  "assigned_car_name": new_excel_name,
              })

            # -------------------------------------------------------------
            # CHẾ ĐỘ 4: TRÍCH XUẤT CAR CÓ SẴN
            # -------------------------------------------------------------
            elif mode == "pre_prepared":
              status_full = "APPROVED (APP)"
              for r in [13, 14, 12]:
                st_val = str(
                    source_ws.cell(row=r, column=prev_com_col).value or ""
                ).upper()
                if "RFC" in st_val:
                  status_full = "RETURN FOR CORRECTIONS (RFC)"
                elif "AWC" in st_val:
                  status_full = "APPROVED WITH COMMENTS (AWC)"
                elif "RJ" in st_val:
                  status_full = "REJECTED (RJ)"

              car_base_match = re.search(
                  r"(\d{4,}\s*CaRs?)", file.filename, re.IGNORECASE
              )
              car_lbl = car_base_match.group(1) if car_base_match else "CaRs"

              zip_file.writestr(clean_safe_filename(file.filename), content)
              processed_car_list.append({
                  "doc_no": final_doc_no,
                  "rev": final_rev,
                  "title": final_title,
                  "status": status_full,
                  "car_label": f"{car_lbl} R{max_r}",
                  "assigned_car_name": file.filename,
              })

            del content
            gc.collect()

          except Exception as fe:
            print(f"Lỗi tệp {file.filename}: {fe}")
            warnings.append(f"Không xử lý được tệp {file.filename}: {fe}")

      # 3. TRẢ LỜI THƯ THUẦN TÚY (KHÔNG KÈM CAR)
      elif mode == "no_car_letter":
        if cover_data["docs"]:
          for doc_k, d_info in reversed(list(cover_data["docs"].items())):
            processed_car_list.append({
                "doc_no": d_info["doc_no"],
                "rev": d_info["rev"],
                "title": d_info["title"],
                "status": "REVIEWED (NO CAR)",
                "car_label": "N/A",
                "assigned_car_name": "N/A",
            })
        else:
          processed_car_list.append({
              "doc_no": cover_data["tr_no"],
              "rev": "0",
              "title": cover_data["subject"],
              "status": "REVIEWED (NO CAR)",
              "car_label": "N/A",
              "assigned_car_name": "N/A",
          })

      # =========================================================================
      # ĐỒNG BỘ 2 THƯ WORD COVER
      # =========================================================================
      safe_subj_name = clean_safe_filename(cover_data["subject"])

      doc_vi = Document(vi_template_path)
      for tbl in doc_vi.tables:
        for row in tbl.rows:
          for cell in row.cells:
            for p in cell.paragraphs:
              if "ngày" in p.text and "tháng" in p.text and p.runs:
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
          tgt = (
              f"{cover_data['tr_no']} ngày {cover_data['date_str']} của Nhà"
              f" thầu EPC gói 15 về {cover_data['subject']}."
          )
          if not replace_value_after_keyword(p, "Căn cứ công văn số", tgt):
            p.text = f"Căn cứ công văn số {tgt}"

      vi_io = io.BytesIO()
      doc_vi.save(vi_io)
      zip_file.writestr(f"TVÐ1-NÐNLM {safe_subj_name}.docx", vi_io.getvalue())

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
              elif "Identification Number." in p.text:
                replace_value_after_keyword(
                    p, "Identification Number.", qti_no.strip()
                )
              elif "Subject:" in p.text:
                replace_value_after_keyword(
                    p, "Subject:", cover_data["subject"]
                )
              # Đặt SAU CÙNG: trước đây đoạn Subject/số hiệu có chữ "202x" bị ghi đè thành ngày
              elif "October" in p.text or "202" in p.text:
                update_date_paragraph(p, current_date_en)

      sync_footer_content(doc_en, qti_no.strip(), cover_data["subject"])

      ref_idx, yours_idx = None, None
      for idx, p in enumerate(doc_en.paragraphs):
        if "This has reference to the" in p.text:
          p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
          f_name = (
              p.runs[0].font.name
              if (p.runs and p.runs[0].font.name)
              else "Times New Roman"
          )
          f_size = (
              p.runs[0].font.size if (p.runs and p.runs[0].font.size) else Pt(13)
          )
          p.clear()
          r = p.add_run(
              f"This has reference to the {doc_type_en} {cover_data['tr_no']},"
              " we would like to comment as follows:"
          )
          r.font.name, r.font.size = f_name, f_size
          ref_idx = idx
        elif "Yours sincerely" in p.text:
          p.alignment, yours_idx = WD_ALIGN_PARAGRAPH.LEFT, idx
          break

      if ref_idx is not None and yours_idx is not None:
        for p in [doc_en.paragraphs[i] for i in range(ref_idx + 1, yours_idx)]:
          p._element.getparent().remove(p._element)

      target_p = None
      for p in doc_en.paragraphs:
        if "Yours sincerely" in p.text:
          target_p = p
          break

      if target_p is not None:
        if mode == "no_car_letter":
          items_iter = (
              reversed(list(cover_data["docs"].values()))
              if cover_data["docs"]
              else [{
                  "doc_no": cover_data["tr_no"],
                  "title": cover_data["subject"],
              }]
          )
          for it in items_iter:
            new_p = target_p.insert_paragraph_before("")
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.left_indent, (
                new_p.paragraph_format.first_line_indent
            ) = (Inches(0.35), Inches(-0.35))
            new_p.paragraph_format.space_before, (
                new_p.paragraph_format.space_after
            ) = (0, Pt(4))
            if cover_data["docs"]:
              txt = (
                  f"-\tThe document {it['doc_no']} {it['title']} Rev.{it['rev']}"
                  " has been reviewed and found acceptable with no further"
                  " comments."
              )
            else:
              txt = (
                  "-\tWe have reviewed the submitted transmittal/documents and"
                  " have no objection."
              )
            run = new_p.add_run(txt)
            run.font.name, run.font.size = "Times New Roman", Pt(13)
        else:
          for item in reversed(processed_car_list):
            new_p = target_p.insert_paragraph_before("")
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.left_indent, (
                new_p.paragraph_format.first_line_indent
            ) = (Inches(0.35), Inches(-0.35))
            new_p.paragraph_format.space_before, (
                new_p.paragraph_format.space_after
            ) = (0, Pt(4))
            run = new_p.add_run(
                f"-\tThe document {item['doc_no']} {item['title']}"
                f" Rev.{item['rev']} is {item['status']} indicated as per"
                f" {item['car_label']}."
            )
            run.font.name, run.font.size = "Times New Roman", Pt(13)

      en_io = io.BytesIO()
      doc_en.save(en_io)
      zip_file.writestr(
          f"{clean_safe_filename(qti_no)} {safe_subj_name}.docx",
          en_io.getvalue(),
      )

    headers = {
        "Content-Disposition": (
            f'attachment; filename="{clean_safe_filename(qti_no)}.zip"'
        )
    }
    try:
      items = [
          {
              "doc_no": i["doc_no"],
              "assigned_car": i["assigned_car_name"],
              "status": i["status"],
          }
          for i in processed_car_list
      ]
      total_items = len(items)
      # Header HTTP có giới hạn dung lượng: bớt dần phần tử cuối thay vì bỏ cả danh sách
      while items:
        b64_str = base64.b64encode(
            json.dumps(items, ensure_ascii=False).encode("utf-8")
        ).decode("ascii")
        if len(b64_str) < 6000:
          headers["X-Processed-Items"] = b64_str
          break
        items.pop()
      info = {
          "tr_no": cover_data["tr_no"],
          "date": cover_data["date_str"],
          "subject": cover_data["subject"],
          "doc_count": len(cover_data["docs"]),
          "method": cover_data.get("parse_info", {}).get("method", ""),
          "ocr": cover_data.get("parse_info", {}).get("ocr", False),
          "processed_total": total_items,
          "processed_shown": len(items),
          "warnings": warnings[:10],
      }
      info_b64 = base64.b64encode(
          json.dumps(info, ensure_ascii=False).encode("utf-8")
      ).decode("ascii")
      if len(info_b64) < 6000:
        headers["X-Cover-Info"] = info_b64
    except Exception as he:
      print("Không tạo được header kết quả:", he)

    return FileResponse(
        path=temp_zip_path,
        media_type="application/zip",
        filename=f"{clean_safe_filename(qti_no)}.zip",
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
