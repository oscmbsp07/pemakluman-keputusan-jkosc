import io
import re
import zipfile
from datetime import date

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================
# DOCX formatting helpers
# =========================
def set_document_defaults(doc: Document):
    # Default font: Times New Roman 12
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    # Ensure East Asia font set too
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)


def add_horizontal_rule(paragraph):
    """Add a bottom border line to a paragraph (looks like the template line)."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")       # thickness
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_logo_if_exists(doc: Document):
    # Optional: user can add assets/logo.png
    import os
    logo_path = "assets/logo.png"
    if os.path.exists(logo_path):
        p = doc.add_paragraph()
        run = p.add_run()
        try:
            run.add_picture(logo_path, width=Cm(3.0))
        except Exception:
            pass  # ignore if picture load fails


def bold_run(p, text: str):
    r = p.add_run(text)
    r.bold = True
    return r


# =========================
# Malay date parsing
# =========================
MALAY_MONTHS = {
    "JANUARI": 1,
    "FEBRUARI": 2,
    "MAC": 3,
    "APRIL": 4,
    "MEI": 5,
    "JUN": 6,
    "JULAI": 7,
    "OGOS": 8,
    "SEPTEMBER": 9,
    "OKTOBER": 10,
    "NOVEMBER": 11,
    "DISEMBER": 12,
}

MALAY_WEEKDAYS = {
    0: "Isnin",
    1: "Selasa",
    2: "Rabu",
    3: "Khamis",
    4: "Jumaat",
    5: "Sabtu",
    6: "Ahad",
}

def parse_malay_date_anywhere(text: str) -> date | None:
    # e.g. "12 Januari 2026" / "12 JANUARI 2026"
    m = re.search(
        r"(\d{1,2})\s+(JANUARI|FEBRUARI|MAC|APRIL|MEI|JUN|JULAI|OGOS|SEPTEMBER|OKTOBER|NOVEMBER|DISEMBER)\s+(\d{4})",
        text.upper(),
    )
    if not m:
        return None
    d = int(m.group(1))
    mon = MALAY_MONTHS[m.group(2)]
    y = int(m.group(3))
    return date(y, mon, d)

def format_malay_date(d: date) -> str:
    inv = {v: k.title() for k, v in MALAY_MONTHS.items()}
    return f"{d.day} {inv[d.month]} {d.year}"

def format_malay_weekday(d: date) -> str:
    return MALAY_WEEKDAYS[d.weekday()]


# =========================
# Agenda text extraction
# =========================
def extract_text_lines(doc: Document) -> list[str]:
    lines = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            lines.append(t)

    # Sometimes agenda is in tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = (p.text or "").strip()
                    if t:
                        lines.append(t)
    return lines


# =========================
# Meeting info parsing
# =========================
def parse_meeting_info(lines: list[str]) -> dict:
    big = "\n".join(lines)

    # Bil. 01/2026
    m_bil = re.search(r"\bBIL\.?\s*(\d{1,2})\s*/\s*(\d{4})\b", big, flags=re.IGNORECASE)
    if not m_bil:
        raise ValueError("Tak jumpa 'Bil. 01/2026' dalam agenda.")
    bil_no = int(m_bil.group(1))
    year = int(m_bil.group(2))

    # Prefer date from a line containing "Tarikh"
    meeting_date = None
    for ln in lines[:80]:  # early part of agenda usually
        if re.search(r"\bTarikh\b", ln, flags=re.IGNORECASE):
            d = parse_malay_date_anywhere(ln)
            if d:
                meeting_date = d
                break

    # Fallback: first date anywhere
    if not meeting_date:
        meeting_date = parse_malay_date_anywhere(big)

    if not meeting_date:
        raise ValueError("Tak jumpa tarikh mesyuarat (contoh: 12 Januari 2026) dalam agenda.")

    return {
        "bil_no": bil_no,
        "year": year,
        "meeting_date": meeting_date,
        "meeting_date_str": format_malay_date(meeting_date),
        "meeting_weekday": format_malay_weekday(meeting_date),
    }


# =========================
# Case parsing (PKM/BGN only)
# =========================
def parse_cases(lines: list[str]) -> list[dict]:
    cases = []
    i = 0
    n = len(lines)

    # Detect start like: "KERTAS MESYUARAT BIL. OSC/PKM/001/2026" or "OSC/BGN/036/2026"
    pat = re.compile(r"KERTAS\s+MESYUARAT\s+BIL\.?\s*(OSC/(PKM|BGN)/([^/\s]+)/(\d{4}))", re.IGNORECASE)

    while i < n:
        ln = lines[i]
        m = pat.search(ln)
        if not m:
            i += 1
            continue

        full_code = m.group(1).upper()
        case_type = m.group(2).upper()      # PKM / BGN
        raw_item = m.group(3)               # 001 / 36 / 052(U) etc
        year = int(m.group(4))

        # Agenda item number (for rujukan kami bracket)
        digits = re.sub(r"\D", "", raw_item)
        item_no_int = int(digits) if digits else 0

        # Gather block after this line until next "KERTAS MESYUARAT BIL"
        j = i + 1
        block = []
        while j < n and not pat.search(lines[j]):
            block.append(lines[j])
            j += 1

        # Extract key fields inside block
        # Usually:
        #   <nama permohonan lines...>
        #   Perunding : ...
        #   Pemohon / Pemilik Projek : ...
        #   No. Rujukan OSC : ...
        perunding = ""
        pemohon = ""
        id_permohonan = ""

        # Find kv lines, but also build "nama permohonan" from top portion until kv begins
        nama_lines = []
        kv_started = False

        for b in block:
            m_kv = re.match(
                r"^(Perunding|Pemohon|Pemilik\s+Projek|No\.?\s*Rujukan\s*OSC)\s*:\s*(.+)$",
                b,
                flags=re.IGNORECASE,
            )
            if m_kv:
                kv_started = True
                key = m_kv.group(1).strip().lower()
                val = m_kv.group(2).strip()
                if key.startswith("perunding"):
                    perunding = val
                elif key.startswith("pemohon") or key.startswith("pemilik"):
                    pemohon = val
                elif key.startswith("no"):
                    id_permohonan = val
            else:
                if not kv_started:
                    # Still part of nama permohonan chunk
                    # Ignore empty / headings
                    if b.strip():
                        nama_lines.append(b.strip())

        nama_permohonan_raw = " ".join(nama_lines).strip()
        # Normalize spacing
        nama_permohonan_raw = re.sub(r"\s+", " ", nama_permohonan_raw)

        # Build “Jenis” + “Nama” to match colleague style
        if case_type == "PKM":
            jenis_permohonan = "Kebenaran Merancang"
            # Ensure Nama starts with "Permohonan Kebenaran Merancang ..."
            if not nama_permohonan_raw.lower().startswith("permohonan"):
                nama_permohonan = f"Permohonan {jenis_permohonan} {nama_permohonan_raw}".strip()
            else:
                # If starts with permohonan but missing jenis, still prepend
                if "kebenaran merancang" not in nama_permohonan_raw.lower():
                    nama_permohonan = f"Permohonan {jenis_permohonan} {nama_permohonan_raw}".strip()
                else:
                    nama_permohonan = nama_permohonan_raw
        else:
            jenis_permohonan = "Bangunan"
            nama_permohonan = nama_permohonan_raw
            if nama_permohonan and not nama_permohonan.lower().startswith("permohonan"):
                nama_permohonan = f"Permohonan {nama_permohonan}".strip()

        cases.append(
            {
                "code": full_code,
                "case_type": case_type,
                "item_no_int": item_no_int,
                "year": year,
                "perunding": perunding or "-",
                "pemohon": pemohon or "-",
                "jenis_permohonan": jenis_permohonan,
                "nama_permohonan": nama_permohonan or "-",
                "id_permohonan": id_permohonan or "-",
            }
        )

        i = j

    return cases


# =========================
# DOC builder (NO template)
# =========================
def build_letter_doc(meeting: dict, case: dict) -> Document:
    doc = Document()
    set_document_defaults(doc)
    add_logo_if_exists(doc)  # optional assets/logo.png

    bil_no = meeting["bil_no"]
    year = meeting["year"]
    meeting_date = meeting["meeting_date"]
    meeting_date_str = meeting["meeting_date_str"]
    weekday = meeting["meeting_weekday"]

    # Rujukan Kami format: (BilMesyuarat)MBSP/15/1551/(AgendaNo)YYYY
    rujukan_kami = f"({bil_no})MBSP/15/1551/({case['item_no_int']}){year}"

    # --- Header block (top right) using a 1x2 table ---
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    t.columns[0].width = Cm(10)
    t.columns[1].width = Cm(6.5)

    left = t.cell(0, 0).paragraphs[0]
    left.text = ""  # left side empty (logo already above if any)

    right = t.cell(0, 1).paragraphs[0]
    right.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p1 = right
    p1.add_run("Rujukan Tuan :\n")
    p1.add_run(f"Rujukan Kami : {rujukan_kami}\n")
    p1.add_run(f"Tarikh        : {meeting_date_str}")

    # spacing
    doc.add_paragraph("")

    # --- Info table (labels : value) ---
    info = doc.add_table(rows=5, cols=3)
    info.autofit = False
    info.columns[0].width = Cm(4.2)
    info.columns[1].width = Cm(0.6)
    info.columns[2].width = Cm(11.7)

    labels = ["Kepada (PSP)", "Pemilik Projek", "Jenis Permohonan", "Nama Permohonan", "ID Permohonan"]
    values = [
        case["perunding"],
        case["pemohon"],
        case["jenis_permohonan"],
        case["nama_permohonan"],
        case["id_permohonan"],
    ]

    for r in range(5):
        c0 = info.cell(r, 0).paragraphs[0]
        c0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = c0.add_run(labels[r])
        run.bold = True

        info.cell(r, 1).paragraphs[0].text = ":"
        info.cell(r, 2).paragraphs[0].text = values[r]

    # spacing + horizontal line
    p_line = doc.add_paragraph("")
    add_horizontal_rule(p_line)

    # --- Title (bold) ---
    p_title = doc.add_paragraph()
    bold_run(p_title, "PEMAKLUMAN  KEPUTUSAN  MESYUARAT  JAWATANKUASA  PUSAT  SETEMPAT\n(OSC)")
    p_title.paragraph_format.space_after = Pt(6)

    # --- Paragraph 1 ---
    p = doc.add_paragraph("Dengan hormatnya saya diarahkan merujuk perkara di atas.")
    p.paragraph_format.space_after = Pt(6)

    # --- Paragraph 2 (with bold parts like colleague) ---
    bil_str = f"Bil.{bil_no:02d}/{year}"
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(6)
    p2.add_run("2.      Adalah dimaklumkan bahawa ")
    bold_run(p2, "Mesyuarat Jawatankuasa Pusat Setempat (OSC)")
    p2.add_run(f" {bil_str} yang bersidang pada ")
    bold_run(p2, f"{meeting_date_str} ({weekday})")
    p2.add_run(
        " bersetuju untuk memberikan keputusan ke atas permohonan yang telah dikemukakan oleh pihak tuan/puan seperti mana berikut:"
    )

    # --- Checkbox block (2 columns) ---
    cb = doc.add_table(rows=2, cols=2)
    cb.autofit = False
    cb.columns[0].width = Cm(9.0)
    cb.columns[1].width = Cm(7.5)

    # left options
    cb.cell(0, 0).paragraphs[0].text = "☐      LULUS"
    cb.cell(1, 0).paragraphs[0].text = "☐      LULUS DENGAN\n         PINDAAN PELAN /\n         LULUS\n         BERSYARAT"

    # right options
    cb.cell(0, 1).paragraphs[0].text = "☐      TOLAK"
    cb.cell(1, 1).paragraphs[0].text = "☐      TANGGUH"

    doc.add_paragraph("")

    # --- Paragraph 3 ---
    p3 = doc.add_paragraph(
        "3.      Walau bagaimanapun, keputusan muktamad bagi permohonan yang berkenaan adalah tertakluk kepada surat kelulusan / penolakan yang akan dikeluarkan oleh Jabatan Induk yang memproses."
    )
    p3.paragraph_format.space_after = Pt(10)

    doc.add_paragraph("Sekian, terima kasih.")
    doc.add_paragraph("")

    # Slogan + signature placeholder (boleh isi manual)
    s = doc.add_paragraph()
    bold_run(s, "\"MALAYSIA MADANI\"")
    doc.add_paragraph("“BERKHIDMAT UNTUK NEGARA”")
    doc.add_paragraph("“CEKAP, AKAUNTABILITI, TELUS”")
    doc.add_paragraph("")
    doc.add_paragraph("Saya yang menjalankan amanah,")
    doc.add_paragraph("")
    doc.add_paragraph("")
    doc.add_paragraph("(                               )")

    return doc


def doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# =========================
# Streamlit App
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (DOCX) → auto extract kes PKM & BGN sahaja → auto guna Tarikh & Bil dari agenda → jana Word + download ZIP.")

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

if uploaded:
    try:
        agenda_doc = Document(io.BytesIO(uploaded.read()))
        lines = extract_text_lines(agenda_doc)

        meeting = parse_meeting_info(lines)
        cases = parse_cases(lines)

        st.success(
            f"Bil.{meeting['bil_no']:02d}/{meeting['year']} | Tarikh: {meeting['meeting_date_str']} ({meeting['meeting_weekday']}) | "
            f"Kes PKM/BGN dijumpai: {len(cases)}"
        )

        if not cases:
            st.error("Tiada kes PKM (Kebenaran Merancang) atau BGN (Bangunan) dijumpai dalam agenda.")
        else:
            with st.expander("Preview kes (read-only)"):
                for c in cases:
                    st.markdown(f"**{c['code']}**")
                    st.write(f"Perunding: {c['perunding']}")
                    st.write(f"Pemilik Projek/Pemohon: {c['pemohon']}")
                    st.write(f"Jenis: {c['jenis_permohonan']}")
                    st.write(f"Nama Permohonan: {c['nama_permohonan']}")
                    st.write(f"ID Permohonan (No. Rujukan OSC): {c['id_permohonan']}")
                    st.divider()

            if st.button("Jana & Download (ZIP)", type="primary"):
                with st.spinner("Sedang jana dokumen Word..."):
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
                        for c in cases:
                            out_doc = build_letter_doc(meeting, c)
                            out_bytes = doc_to_bytes(out_doc)

                            # Filename ikut pattern awak explain: (BilMesyuarat)MBSP-15-1551-(AgendaNo)YYYY
                            fname = f"({meeting['bil_no']})MBSP-15-1551-({c['item_no_int']}){meeting['year']}.docx"
                            z.writestr(fname, out_bytes)

                    zip_buf.seek(0)
                    st.download_button(
                        "Download ZIP",
                        data=zip_buf.getvalue(),
                        file_name=f"pemakluman_keputusan_Bil{meeting['bil_no']:02d}_{meeting['year']}.zip",
                        mime="application/zip",
                    )

    except Exception as e:
        st.error(f"Ralat: {e}")
