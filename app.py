import io
import re
import zipfile
import datetime
import unicodedata

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

APP_TITLE = "Pemakluman Keputusan JKOSC"

# =========================
# Helpers: Word formatting
# =========================

def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    """Margins are in dxa (twips). 1 cm ≈ 567 dxa."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    """Repeat this table row at the top of each page when the table spans pages."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)
    tblHeader.set(qn("w:val"), "true")


def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "nil")


def set_run_font(run, name="Arial", size=11, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)

# =========================
# Helpers: text parsing
# =========================

MALAY_MONTHS = {
    "JANUARI": ("Januari", 1),
    "FEBRUARI": ("Februari", 2),
    "MAC": ("Mac", 3),
    "APRIL": ("April", 4),
    "MEI": ("Mei", 5),
    "JUN": ("Jun", 6),
    "JULAI": ("Julai", 7),
    "OGOS": ("Ogos", 8),
    "SEPTEMBER": ("September", 9),
    "OKTOBER": ("Oktober", 10),
    "NOVEMBER": ("November", 11),
    "DISEMBER": ("Disember", 12),
}

MALAY_DOW = ["Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu", "Ahad"]  # Monday=0


def normalize_space(s: str) -> str:
    s = unicodedata.normalize("NFKC", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_meeting_bil_and_date(full_text: str):
    # Bil. 01/2026
    m = re.search(r"\bBIL\.?\s*(\d{1,3})/(\d{4})\b", full_text, flags=re.IGNORECASE)
    bil_num, year = (int(m.group(1)), int(m.group(2))) if m else (None, None)

    # Date: 12 JANUARI 2026 / 12 Januari 2026
    m2 = re.search(
        r"\b(\d{1,2})\s*(JANUARI|FEBRUARI|MAC|APRIL|MEI|JUN|JULAI|OGOS|SEPTEMBER|OKTOBER|NOVEMBER|DISEMBER)\s*(\d{4})\b",
        full_text,
        flags=re.IGNORECASE,
    )
    dt, date_str, dow = None, None, None
    if m2:
        day = int(m2.group(1))
        mon_key = m2.group(2).upper()
        yr = int(m2.group(3))
        mon_name, mon_num = MALAY_MONTHS.get(mon_key, (m2.group(2).title(), None))
        date_str = f"{day} {mon_name} {yr}"
        if mon_num:
            try:
                dt = datetime.date(yr, mon_num, day)
                dow = MALAY_DOW[dt.weekday()]
            except Exception:
                dt = None
                dow = None

    return bil_num, year, date_str, dt, dow


def parse_agenda_docx(uploaded_file_bytes: bytes):
    doc = Document(io.BytesIO(uploaded_file_bytes))
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    full_text = "\n".join(paras)

    bil_num, meeting_year, date_str, dt, dow = extract_meeting_bil_and_date(full_text)

    # Find all "KERTAS MESYUARAT ..." blocks for OSC/PKM + OSC/BGN
    block_starts = []
    for i, p in enumerate(paras):
        up = p.upper()
        if up.startswith("KERTAS MESYUARAT") and "OSC/" in up:
            block_starts.append(i)

    cases = []
    for idx, start_i in enumerate(block_starts):
        end_i = block_starts[idx + 1] if idx + 1 < len(block_starts) else len(paras)
        block = paras[start_i:end_i]
        header = block[0]

        m = re.search(r"OSC/(PKM|BGN)/(\d{1,3})/(\d{4})", header, flags=re.IGNORECASE)
        if not m:
            continue

        jenis_code = m.group(1).upper()
        kertas_no = int(m.group(2))
        kertas_year = int(m.group(3))

        lines = [normalize_space(x) for x in block[1:] if normalize_space(x)]

        def find_value(prefixes):
            for ln in lines:
                for pref in prefixes:
                    if re.match(rf"^{re.escape(pref)}\s*:?", ln, flags=re.IGNORECASE):
                        parts = re.split(r"\s*:\s*", ln, maxsplit=1)
                        return parts[1].strip() if len(parts) > 1 else ""
            return ""

        pemohon = find_value(["Pemohon"])
        perunding = find_value(["Perunding"])
        no_ruj_osc = find_value(["No. Rujukan OSC"])

        # Nama permohonan: from the first lines after header until the first field line
        field_pat = re.compile(
            r"^(Pemohon|Perunding|Lokasi|Koordinat|No\.\s*Rujukan|Pelan\s*Susunatur|No\.\s*Rujukan\s*OSC)\b",
            re.IGNORECASE,
        )

        nama_lines = []
        for ln in lines:
            if field_pat.search(ln):
                break
            nama_lines.append(ln)

        nama_permohonan = "\n".join(nama_lines).strip()
        # Safety: stop if "Pemohon" is embedded
        nama_permohonan = re.split(r"\bPemohon\b", nama_permohonan, flags=re.IGNORECASE)[0].strip()

        jenis_display = "Kebenaran Merancang" if jenis_code == "PKM" else "Bangunan"

        cases.append(
            {
                "start_i": start_i,
                "jenis_code": jenis_code,
                "jenis_display": jenis_display,
                "kertas_no": kertas_no,
                "kertas_year": kertas_year,
                "nama_permohonan": nama_permohonan,
                "pemohon": pemohon,
                "perunding": perunding,
                "no_ruj_osc": no_ruj_osc,
            }
        )

    cases.sort(key=lambda x: x["start_i"])

    # Assign global sequential number 1..N across PKM+BGN (PKM then BGN, ikut urutan agenda)
    for i, c in enumerate(cases, start=1):
        c["global_idx"] = i

    return {
        "bil_num": bil_num,
        "meeting_year": meeting_year,
        "date_str": date_str,
        "dt": dt,
        "dow": dow,
        "cases": cases,
    }


# =========================
# DOCX builder (ikut format colleague)
# =========================

def _add_rujukan_block(cell, rujukan_kami, tarikh_str):
    # Inner table: 3 rows, 2 cols, right aligned, Arial 11 (BLACK)
    t = cell.add_table(rows=3, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.RIGHT
    remove_table_borders(t)
    t.columns[0].width = Cm(3.2)
    t.columns[1].width = Cm(7.5)

    labels = ["Rujukan Tuan", "Rujukan Kami", "Tarikh"]
    values = ["", rujukan_kami, tarikh_str]

    for r in range(3):
        c0, c1 = t.cell(r, 0), t.cell(r, 1)
        set_cell_margins(c0, top=0, bottom=0, start=0, end=80)
        set_cell_margins(c1, top=0, bottom=0, start=0, end=0)

        p0 = c0.paragraphs[0]
        p1 = c1.paragraphs[0]
        p0.clear()
        p1.clear()
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        r0 = p0.add_run(labels[r] + "  :")
        set_run_font(r0, "Arial", 11, bold=False)

        r1 = p1.add_run(values[r])
        set_run_font(r1, "Arial", 11, bold=False)

        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)

    return t


def _add_info_table(container_cell, case):
    t = container_cell.add_table(rows=5, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    remove_table_borders(t)
    t.columns[0].width = Cm(4.0)
    t.columns[1].width = Cm(0.6)
    t.columns[2].width = Cm(12.4)

    labels = ["Kepada (PSP)", "Pemilik Projek", "Jenis\nPermohonan", "Nama\nPermohonan", "ID\nPermohonan"]
    values = [case["perunding"], case["pemohon"], case["jenis_display"], case["nama_permohonan"], case["no_ruj_osc"]]

    for i in range(5):
        c0, c1, c2 = t.cell(i, 0), t.cell(i, 1), t.cell(i, 2)
        for c in (c0, c1, c2):
            set_cell_margins(c, top=0, bottom=0, start=0, end=0)
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # Label
        p0 = c0.paragraphs[0]
        p0.clear()
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run(labels[i])
        set_run_font(r0, "Arial", 11, bold=True)

        # Colon
        p1 = c1.paragraphs[0]
        p1.clear()
        p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p1.add_run(":")
        set_run_font(r1, "Arial", 11, bold=False)

        # Value
        if i == 3:
            # Nama Permohonan: left aligned (no center), keep line breaks
            first = True
            lines = values[i].split("\n") if values[i] else [""]
            for ln in lines:
                p = c2.paragraphs[0] if first else c2.add_paragraph()
                if first:
                    p.clear()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0

                rr = p.add_run(ln)
                set_run_font(rr, "Arial", 11, bold=False)

                # Indentation for list-like lines (BGN cases)
                if re.match(r"^(\d+\)|[A-Z]\)|[-•])", ln.strip()):
                    p.paragraph_format.left_indent = Cm(0.6)
                    p.paragraph_format.first_line_indent = Cm(-0.2)
                first = False
        else:
            p2 = c2.paragraphs[0]
            p2.clear()
            p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
            rr = p2.add_run(values[i] if values[i] else "-")
            set_run_font(rr, "Arial", 11, bold=False)

        for pp in (c0.paragraphs + c1.paragraphs + c2.paragraphs):
            pp.paragraph_format.space_before = Pt(0)
            pp.paragraph_format.space_after = Pt(0)
            pp.paragraph_format.line_spacing = 1.0

    return t


def _add_checkbox_table(container_cell):
    # 2 rows x 4 cols with fixed widths so boxes align nicely
    t = container_cell.add_table(rows=2, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    remove_table_borders(t)

    widths = [Cm(1.2), Cm(6.0), Cm(1.2), Cm(4.5)]
    for i, w in enumerate(widths):
        t.columns[i].width = w

    items = [
        ("□", "LULUS", "□", "TOLAK"),
        ("□", "LULUS DENGAN\nPINDAAN PELAN /\nLULUS\nBERSYARAT", "□", "TANGGUH"),
    ]

    for r in range(2):
        for c in range(4):
            cell = t.cell(r, c)
            set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

            p = cell.paragraphs[0]
            p.clear()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0

            if c in (0, 2):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                rr = p.add_run(items[r][c])
                set_run_font(rr, "Arial", 11, bold=False)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                rr = p.add_run(items[r][c])
                set_run_font(rr, "Arial", 11, bold=True)

    return t


def build_pemakluman_doc(case, meeting):
    bil_num = meeting["bil_num"] or 0
    year = meeting["meeting_year"] or case["kertas_year"]
    tarikh_str = meeting["date_str"] or ""
    dow = meeting["dow"]
    tarikh_dow = f"{tarikh_str} ({dow})" if tarikh_str and dow else tarikh_str

    idx = case["global_idx"]
    # Special rule: selepas 1551/ jadi ( ) kosong untuk staf isi manual
    rujukan_kami = f"({idx})MBSP/15/1551/( ){year}"

    doc = Document()

    # Page setup (margin ikut gaya sample)
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    # Outer table trick:
    # Row 0 = rujukan block, repeat on each page (so page 2 ada block sama, and it is BLACK, not Word-header grey)
    outer = doc.add_table(rows=2, cols=1)
    outer.alignment = WD_TABLE_ALIGNMENT.LEFT
    outer.autofit = False
    remove_table_borders(outer)
    outer.columns[0].width = Cm(16.0)

    header_cell = outer.cell(0, 0)
    body_cell = outer.cell(1, 0)
    set_cell_margins(header_cell, top=0, bottom=0, start=0, end=0)
    set_cell_margins(body_cell, top=0, bottom=0, start=0, end=0)

    set_repeat_table_header(outer.rows[0])

    _add_rujukan_block(header_cell, rujukan_kami, tarikh_str)

    # Info table
    _add_info_table(body_cell, case)

    # Line
    p_line = body_cell.add_paragraph()
    p_line.paragraph_format.space_before = Pt(6)
    p_line.paragraph_format.space_after = Pt(6)
    add_horizontal_line(p_line)

    # Title
    p_title = body_cell.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rr = p_title.add_run("PEMAKLUMAN   KEPUTUSAN   MESYUARAT   JAWATANKUASA   PUSAT   SETEMPAT\n(OSC)")
    set_run_font(rr, "Arial", 11, bold=True)
    p_title.paragraph_format.space_after = Pt(6)

    # Intro
    p_open = body_cell.add_paragraph()
    rr = p_open.add_run("Dengan hormatnya saya diarahkan merujuk perkara di atas.")
    set_run_font(rr, "Arial", 11, bold=False)
    p_open.paragraph_format.space_after = Pt(6)

    # Para 2
    p2 = body_cell.add_paragraph()
    txt2 = (
        "2.\tAdalah dimaklumkan bahawa Mesyuarat Jawatankuasa Pusat Setempat (OSC)\n"
        f"Bil.{bil_num:02d}/{year} yang bersidang pada {tarikh_dow} telah memaklumkan\n"
        "keputusan ke atas permohonan yang telah dikemukakan oleh pihak tuan/puan seperti mana\n"
        "berikut:"
    )
    rr = p2.add_run(txt2)
    set_run_font(rr, "Arial", 11, bold=False)
    p2.paragraph_format.space_after = Pt(6)

    # Checkbox area (align cantik)
    _add_checkbox_table(body_cell)

    # Para 3
    p3 = body_cell.add_paragraph()
    p3.paragraph_format.space_before = Pt(6)
    txt3 = (
        "3.\tWalau bagaimanapun, keputusan muktamad bagi permohonan yang berkenaan\n"
        "adalah tertakluk kepada surat kelulusan / penolakan yang akan dikeluarkan oleh Jabatan\n"
        "Induk yang memproses."
    )
    rr = p3.add_run(txt3)
    set_run_font(rr, "Arial", 11, bold=False)
    p3.paragraph_format.space_after = Pt(6)

    # Closing
    p_end = body_cell.add_paragraph()
    rr = p_end.add_run("Sekian, terima kasih.")
    set_run_font(rr, "Arial", 11, bold=False)
    p_end.paragraph_format.space_after = Pt(12)

    # Motto
    for line in ['"MALAYSIA MADANI"', '“BERKHIDMAT UNTUK NEGARA”', '“CEKAP, AKAUNTABILITI, TELUS”']:
        p = body_cell.add_paragraph()
        rr = p.add_run(line)
        set_run_font(rr, "Arial", 11, bold=True)
        p.paragraph_format.space_after = Pt(0)

    # Signature
    p_sig_intro = body_cell.add_paragraph()
    rr = p_sig_intro.add_run("\nSaya yang menjalankan amanah,")
    set_run_font(rr, "Arial", 11, bold=False)
    p_sig_intro.paragraph_format.space_after = Pt(24)

    p_sigline = body_cell.add_paragraph()
    rr = p_sigline.add_run("______________________________")
    set_run_font(rr, "Arial", 11, bold=False)
    p_sigline.paragraph_format.space_after = Pt(6)

    sig_lines = [
        "(TPr. ANY NUHAIRAH BINTI ABDUL RAZAK )",
        "Ketua Unit",
        "Unit Pusat Setempat (OSC)",
        "Majlis Bandaraya Seberang Perai",
        "✉ : any.nuhairah@mbsp.gov.my",
        "☎ : 04-5497419",
        "",
        "“Seberang Perai Aspirasi Bandar Masa Hadapan”",
        "(Seberang Perai Aspiring City of Tomorrow)",
    ]
    for i, ln in enumerate(sig_lines):
        p = body_cell.add_paragraph()
        if ln == "":
            p.paragraph_format.space_after = Pt(0)
            continue
        rr = p.add_run(ln)
        set_run_font(rr, "Arial", 11, bold=(i == 0))
        if i >= 7:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rr.font.italic = (i == 8)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)

    return doc


def safe_filename(s: str) -> str:
    s = re.sub(r'[\\/*?:"<>|]', "", (s or "").strip())
    s = re.sub(r"\s+", "_", s)
    return s[:60] if s else "Dokumen"


def make_zip(meeting):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for case in meeting["cases"]:
            doc = build_pemakluman_doc(case, meeting)

            idx = case["global_idx"]
            code = case["jenis_code"]
            if code == "PKM":
                no_str = f"{case['kertas_no']:02d}"
            else:
                no_str = f"{case['kertas_no']:03d}"

            year = meeting["meeting_year"] or case["kertas_year"]
            pemohon = safe_filename(case["pemohon"].replace("Tetuan ", "").strip())
            fname = f"{idx:02d}_OSC_{code}_{no_str}_{year}_{pemohon}.docx"

            f = io.BytesIO()
            doc.save(f)
            z.writestr(fname, f.getvalue())

    buf.seek(0)
    return buf


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Upload Agenda (DOCX) → auto extract kes Kebenaran Merancang (PKM) & Bangunan (BGN) → jana dokumen Word (download ZIP).")

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

if uploaded is None:
    st.stop()

try:
    meeting = parse_agenda_docx(uploaded.read())
except Exception as e:
    st.error(f"Gagal baca DOCX: {e}")
    st.stop()

bil_num = meeting["bil_num"]
date_str = meeting["date_str"]
cnt = len(meeting["cases"])
cnt_pkm = sum(1 for c in meeting["cases"] if c["jenis_code"] == "PKM")
cnt_bgn = sum(1 for c in meeting["cases"] if c["jenis_code"] == "BGN")

st.success(f"Jumpa Bil.: {bil_num:02d}/{meeting['meeting_year']} | Tarikh mesyuarat: {date_str} | Kes PKM: {cnt_pkm} | Kes BGN: {cnt_bgn} | Jumlah: {cnt}")

with st.spinner("Menjana dokumen Word (ZIP)..."):
    zip_buf = make_zip(meeting)

st.download_button(
    "Download semua dokumen (ZIP)",
    data=zip_buf,
    file_name=f"pemakluman_keputusan_Bil_{bil_num:02d}_{meeting['meeting_year']}.zip",
    mime="application/zip",
)
