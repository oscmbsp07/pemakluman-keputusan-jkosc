import io
import re
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


# ============================================================
# STREAMLIT SETUP
# ============================================================
st.set_page_config(page_title="Pemakluman Keputusan JK OSC", layout="centered")


# ============================================================
# DATA MODEL
# ============================================================
@dataclass
class MeetingInfo:
    bil_str: str          # "01"
    year_str: str         # "2026"
    bil_no_int: int       # 1
    tarikh_surat: str     # "12 Januari 2026"
    hari: str             # "Isnin" or ""


@dataclass
class CaseInfo:
    kind: str             # "PKM" or "BGN"
    paper_no: str         # "01" or "001"
    year: str             # "2026"
    pemohon: str
    perunding: str
    id_permohonan: str
    nama_permohonan: str
    jenis_permohonan: str # "Kebenaran Merancang" or "Bangunan"


# ============================================================
# AGENDA PARSER
# ============================================================
MONTHS = {
    "JANUARI": "Januari",
    "FEBRUARI": "Februari",
    "MAC": "Mac",
    "APRIL": "April",
    "MEI": "Mei",
    "JUN": "Jun",
    "JULAI": "Julai",
    "OGOS": "Ogos",
    "SEPTEMBER": "September",
    "OKTOBER": "Oktober",
    "NOVEMBER": "November",
    "DISEMBER": "Disember",
}

DAYS = {
    "ISNIN": "Isnin",
    "SELASA": "Selasa",
    "RABU": "Rabu",
    "KHAMIS": "Khamis",
    "JUMAAT": "Jumaat",
    "SABTU": "Sabtu",
    "AHAD": "Ahad",
}

# Header kes yang dibenarkan sahaja
HEADER_RE = re.compile(
    r"(?i)^\s*KERTAS\s+MESYUARAT\s+BIL\.\s*OSC\s*/\s*(PKM|BGN)\s*/\s*([0-9]{1,3})\s*/\s*([0-9]{4}).*$"
)
# Ada agenda tulis rapat dan ada status hujung (contoh: (U))
HEADER_RE_TIGHT = re.compile(
    r"(?i)^\s*KERTAS\s+MESYUARAT\s+BIL\.\s*OSC\s*/\s*(PKM|BGN)\s*/\s*([0-9]{1,3})\s*/\s*([0-9]{4})\s*\(?.*$"
)

BIL_RE = re.compile(r"(?i)\bBIL\.\s*(\d{1,2})\s*/\s*(\d{4})\b")
DATE_RE = re.compile(
    r"(?i)\b(\d{1,2})\s+("
    + "|".join(MONTHS.keys())
    + r")\s+(\d{4})\s*(?:\(([^)]+)\))?"
)

# Label-label yang membezakan "Nama Permohonan" vs field
LABEL_RE = re.compile(
    r"(?i)^\s*(Pemohon|Perunding|Lokasi|Koordinat|No\.?\s*Ruj(?:ukan)?(?:\s*Pelan)?|No\.?\s*Ruj(?:ukan)?\s*OSC)\b"
)
PEMOHON_RE = re.compile(r"(?i)^\s*Pemohon\s*:?\s*(.*)$")
PERUNDING_RE = re.compile(r"(?i)^\s*Perunding\s*:?\s*(.*)$")
NO_RUJ_OSC_RE = re.compile(r"(?i)^\s*No\.?\s*Rujukan\s*OSC\s*:?\s*(.*)$")


def _docx_lines(doc: Document) -> List[str]:
    """Ambil teks baris demi baris daripada docx (paragraph + table)."""
    lines: List[str] = []
    for p in doc.paragraphs:
        t = (p.text or "").replace("\r", "")
        lines.append(t.rstrip() if t.strip() else "")
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                t = (cell.text or "").replace("\r", "")
                if t.strip():
                    for ln in t.split("\n"):
                        lines.append(ln.rstrip())
    return lines


def parse_meeting_info(lines: List[str]) -> MeetingInfo:
    """WAJIB: extract Bil. XX/YYYY dan tarikh mesyuarat daripada agenda."""
    blob = "\n".join(lines[:250])

    m_bil = BIL_RE.search(blob)
    if not m_bil:
        raise ValueError("Bil mesyuarat / tarikh mesyuarat tak jumpa dalam agenda (Bil mesyuarat tak dijumpai).")

    bil_raw = m_bil.group(1)
    bil = bil_raw.zfill(2)
    year = m_bil.group(2)
    bil_no_int = int(bil_raw)

    m_date = DATE_RE.search(blob)
    if not m_date:
        raise ValueError("Bil mesyuarat / tarikh mesyuarat tak jumpa dalam agenda (Tarikh mesyuarat tak dijumpai).")

    day_num = int(m_date.group(1))
    mon = MONTHS.get(m_date.group(2).upper().strip(), m_date.group(2).title())
    y = m_date.group(3)

    hari_raw = (m_date.group(4) or "").strip().upper()
    hari = DAYS.get(hari_raw, hari_raw.title() if hari_raw else "")

    tarikh_surat = f"{day_num} {mon} {y}"
    return MeetingInfo(bil_str=bil, year_str=year, bil_no_int=bil_no_int, tarikh_surat=tarikh_surat, hari=hari)


def _is_header(line: str) -> Optional[Tuple[str, str, str]]:
    """Return (kind, paper_no, year) if line is PKM/BGN header; else None."""
    if not line:
        return None
    m = HEADER_RE.match(line) or HEADER_RE_TIGHT.match(line)
    if not m:
        return None

    kind = m.group(1).upper()
    paper_no_raw = m.group(2)
    year = m.group(3)

    # PKM biasa 2 digit, BGN biasa 3 digit (ikut amalan)
    paper_no = paper_no_raw.zfill(2 if kind == "PKM" else 3)
    return kind, paper_no, year


def _clean_join_continuation(rhs: str, cont: List[str]) -> str:
    """Join value + continuation lines until next label."""
    parts: List[str] = []
    if rhs.strip():
        parts.append(rhs.strip())

    for ln in cont:
        s = (ln or "").strip()
        if not s:
            continue
        if LABEL_RE.match(s):
            break
        parts.append(s)

    out = " ".join(parts).strip()
    out = re.sub(r"\s*\+\s*", " + ", out)  # kemaskan simbol +
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out


def _extract_multiline_label(block: List[str], rex: re.Pattern) -> str:
    """Extract label value that may continue to next lines."""
    for i, ln in enumerate(block):
        m = rex.match(ln)
        if not m:
            continue
        rhs = (m.group(1) or "").strip()
        cont = []
        for j in range(i + 1, len(block)):
            nxt = block[j]
            if LABEL_RE.match(nxt):
                break
            cont.append(nxt)
        return _clean_join_continuation(rhs, cont)
    return ""


def _is_list_line(s: str) -> bool:
    t = s.strip()
    if not t:
        return False
    if re.match(r"^\(?\d+\)?\s*[\).]\s+.+", t):
        return True
    if re.match(r"^[A-Za-z]\s*[\).]\s+.+", t):
        return True
    if re.match(r"^[-•]\s+.+", t):
        return True
    return False


def _format_nama_permohonan(raw_lines: List[str]) -> str:
    """
    Kemaskan 'Nama Permohonan':
    - Gabung baris ayat biasa dengan space (supaya kemas, bukan 1 perkataan 1 baris).
    - Kekalkan line-break untuk senarai (1), A), -, •.
    """
    lines = [ln.rstrip() for ln in raw_lines]
    while lines and lines[0].strip() == "":
        lines.pop(0)
    while lines and lines[-1].strip() == "":
        lines.pop()
    if not lines:
        return ""

    out_lines: List[str] = []
    buf = ""

    def flush():
        nonlocal buf
        if buf.strip():
            out_lines.append(re.sub(r"\s{2,}", " ", buf.strip()))
        buf = ""

    for ln in lines:
        t = ln.strip()
        if t == "":
            flush()
            continue
        if _is_list_line(t):
            flush()
            out_lines.append(t)
            continue
        if buf == "":
            buf = t
        else:
            buf = f"{buf} {t}"

    flush()
    return "\n".join(out_lines).strip()


def parse_cases(lines: List[str]) -> List[CaseInfo]:
    """Extract semua kes PKM/BGN ikut turutan dalam agenda."""
    headers: List[Tuple[int, str, str, str]] = []
    for i, ln in enumerate(lines):
        h = _is_header(ln)
        if h:
            kind, paper_no, year = h
            headers.append((i, kind, paper_no, year))

    if not headers:
        return []

    cases: List[CaseInfo] = []
    for idx, (start_i, kind, paper_no, year) in enumerate(headers):
        end_i = headers[idx + 1][0] if idx + 1 < len(headers) else len(lines)
        block = lines[start_i:end_i]

        # Find first label line to split "Nama Permohonan"
        first_label_idx = None
        for j, ln in enumerate(block):
            if LABEL_RE.match(ln):
                first_label_idx = j
                break

        nama_lines = block[1:(first_label_idx if first_label_idx is not None else len(block))]
        nama_permohonan = _format_nama_permohonan(nama_lines)

        pemohon = _extract_multiline_label(block, PEMOHON_RE)
        perunding = _extract_multiline_label(block, PERUNDING_RE)
        id_perm = _extract_multiline_label(block, NO_RUJ_OSC_RE)

        jenis = "Kebenaran Merancang" if kind == "PKM" else "Bangunan"

        cases.append(CaseInfo(
            kind=kind,
            paper_no=paper_no,
            year=year,
            pemohon=pemohon,
            perunding=perunding,
            id_permohonan=id_perm,
            nama_permohonan=nama_permohonan,
            jenis_permohonan=jenis,
        ))

    return cases


# ============================================================
# WORD BUILDER (ikut format contoh manual)
# ============================================================
def _set_doc_layout(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Inches(8.26875)     # A4
    sec.page_height = Inches(11.69375)   # A4
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.top_margin = Inches(0.4402777778)
    sec.bottom_margin = Inches(0.8798611111)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)


def _remove_table_borders(tbl):
    tbl_pr = tbl._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "nil")


def _set_cell(cell, text: str, bold: bool = False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(11)
    run.bold = bold


def _add_blank(doc: Document, n: int = 1):
    for _ in range(n):
        p = doc.add_paragraph("")
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)


def build_doc(meeting: MeetingInfo, case: CaseInfo) -> bytes:
    doc = Document()
    _set_doc_layout(doc)

    # Header reference lines
    doc.add_paragraph("Rujukan Tuan :")

    ruj_kami = f"({meeting.bil_no_int})MBSP/15/1551/(   ){meeting.year_str}"
    p = doc.add_paragraph()
    p.add_run("Rujukan Kami \t:  " + ruj_kami)

    p = doc.add_paragraph()
    p.add_run("Tarikh \t:  " + meeting.tarikh_surat)

    _add_blank(doc, 2)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("PEMAKLUMAN KEPUTUSAN MESYUARAT JAWATANKUASA PUSAT SETEMPAT (OSC)")
    run.bold = True

    _add_blank(doc, 1)

    # Paragraph 1
    p = doc.add_paragraph("Dengan hormatnya saya diarah merujuk perkara di atas.      ")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    _add_blank(doc, 1)

    # Paragraph 2 (wajib ada bil & tarikh)
    hari_part = f" ({meeting.hari})" if meeting.hari else ""
    p2 = (
        f"2.\tAdalah dimaklumkan bahawa Mesyuarat Jawatankuasa Pusat Setempat (OSC) "
        f"Bil.{meeting.bil_str}/{meeting.year_str} yang bersidang pada {meeting.tarikh_surat}{hari_part} "
        f"telah membincangkan permohonan yang telah dikemukakan oleh pihak tuan/ puan seperti mana berikut:"
    )
    p = doc.add_paragraph(p2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    _add_blank(doc, 1)

    # Main info table (5 rows x 3 cols; middle col = ":")
    info_tbl = doc.add_table(rows=5, cols=3)
    _remove_table_borders(info_tbl)

    info_tbl.columns[0].width = Inches(1.1875)
    info_tbl.columns[1].width = Inches(0.2590277778)
    info_tbl.columns[2].width = Inches(5.0652777778)

    rows = [
        ("Kepada (PSP)", ":", case.perunding or ""),
        ("Pemilik Projek", ":", case.pemohon or ""),
        ("Jenis Permohonan", ":", case.jenis_permohonan),
        ("Nama Permohonan", ":", case.nama_permohonan or ""),
        ("ID Permohonan", ":", case.id_permohonan or ""),
    ]
    for i, (a, b, c) in enumerate(rows):
        _set_cell(info_tbl.cell(i, 0), a, bold=False)
        _set_cell(info_tbl.cell(i, 1), b, bold=False)
        _set_cell(info_tbl.cell(i, 2), c, bold=False)

    _add_blank(doc, 1)

    # Paragraph 3
    p3 = (
        "3.\tWalau bagaimanapun, keputusan muktamad bagi permohonan yang berkenaan akan dimaklumkan melalui "
        "Surat Kelulusan atau Surat Penolakan yang akan dikeluarkan oleh Jabatan Induk yang memproses."
    )
    p = doc.add_paragraph(p3)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    _add_blank(doc, 1)

    # Decision table (2 rows x 4 cols) — kosong untuk tanda manual
    dec_tbl = doc.add_table(rows=2, cols=4)
    _remove_table_borders(dec_tbl)

    dec_tbl.columns[0].width = Inches(0.9375)
    dec_tbl.columns[1].width = Inches(2.75)
    dec_tbl.columns[2].width = Inches(0.9375)
    dec_tbl.columns[3].width = Inches(2.28125)

    _set_cell(dec_tbl.cell(0, 0), "", bold=False)
    _set_cell(dec_tbl.cell(0, 1), "LULUS", bold=True)
    _set_cell(dec_tbl.cell(0, 2), "", bold=False)
    _set_cell(dec_tbl.cell(0, 3), "     TOLAK", bold=True)

    _set_cell(dec_tbl.cell(1, 0), "", bold=False)
    _set_cell(dec_tbl.cell(1, 1), "LULUS DENGAN PINDAAN PELAN / LULUS BERSYARAT", bold=True)
    _set_cell(dec_tbl.cell(1, 2), "", bold=False)
    _set_cell(dec_tbl.cell(1, 3), "     TANGGUH", bold=True)

    _add_blank(doc, 1)

    # Optional: untuk kes panjang (contoh SkyWorld), ulang header reference sebelum penutup
    if len(case.nama_permohonan or "") >= 380:
        doc.add_paragraph("Rujukan Tuan :")
        rr = doc.add_paragraph()
        rr.add_run("Rujukan Kami \t:  " + ruj_kami)
        tt = doc.add_paragraph()
        tt.add_run("Tarikh \t:  " + meeting.tarikh_surat)
        _add_blank(doc, 1)

    # Closing block
    doc.add_paragraph("Sekian, terima kasih.")
    _add_blank(doc, 1)
    doc.add_paragraph('"MALAYSIA MADANI"')
    doc.add_paragraph('“BERKHIDMAT UNTUK NEGARA”')
    doc.add_paragraph('“CEKAP, AKAUNTABILITI, TELUS”')
    _add_blank(doc, 1)
    doc.add_paragraph("Saya yang menjalankan amanah,")
    _add_blank(doc, 1)
    doc.add_paragraph("_____________________________")
    doc.add_paragraph("(TPr. ANY NUHAIRAH BINTI ABDUL RAZAK )")
    doc.add_paragraph("Ketua Unit")
    doc.add_paragraph("Unit Pusat Setempat (OSC)")
    doc.add_paragraph("Majlis Bandaraya Seberang Perai")
    doc.add_paragraph(": any.nuhairah@mbsp.gov.my ")
    doc.add_paragraph(": 04-5497419")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============================================================
# UI
# ============================================================
st.title("Pemakluman Keputusan JK OSC")
st.caption("Khas untuk KM (OSC/PKM) & Bangunan (OSC/BGN) sahaja. Keputusan adalah manual (tanda dalam dokumen).")

agenda = st.file_uploader("Upload Agenda Mesyuarat JK OSC (.docx)", type=["docx"])
gen = st.button("JANA PEMAKLUMAN (ZIP)")

if gen:
    if not agenda:
        st.error("Sila muat naik Agenda .docx.")
        st.stop()

    try:
        agenda_doc = Document(io.BytesIO(agenda.getvalue()))
        lines = _docx_lines(agenda_doc)

        meeting = parse_meeting_info(lines)
        cases = parse_cases(lines)

        if not cases:
            st.error("Tiada kes KM (PKM) / BGN dalam agenda.")
            st.stop()

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for c in cases:
                docx_bytes = build_doc(meeting, c)
                zf.writestr(f"{c.kind}_{c.paper_no}_{c.year}.docx", docx_bytes)

        zip_buf.seek(0)

        st.success(f"Siap. Jumlah dokumen dijana: {len(cases)}")
        st.download_button(
            "Muat turun ZIP",
            data=zip_buf.getvalue(),
            file_name=f"Pemakluman_Keputusan_Bil_{meeting.bil_str}_{meeting.year_str}.zip",
            mime="application/zip",
        )

        with st.expander("Ringkasan (semakan cepat)"):
            st.write({
                "Bil Mesyuarat": f"{meeting.bil_str}/{meeting.year_str}",
                "Tarikh Surat": meeting.tarikh_surat,
                "Jumlah dokumen": len(cases),
                "Rujukan Kami": f"({meeting.bil_no_int})MBSP/15/1551/(   ){meeting.year_str}",
            })

    except Exception as e:
        st.error("Gagal proses agenda.")
        st.exception(e)
