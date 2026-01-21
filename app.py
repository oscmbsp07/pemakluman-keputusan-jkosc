# app.py
import re
import zipfile
from io import BytesIO

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# -----------------------
# Helpers: styling
# -----------------------
def set_doc_defaults(doc: Document, font_name="Times New Roman", font_size_pt=12):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    style.font.size = Pt(font_size_pt)

    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.0)


def add_horizontal_line(doc: Document):
    p = doc.add_paragraph("")
    p_format = p.paragraph_format

    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")      # thickness
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)
    p_format.space_before = Pt(6)
    p_format.space_after = Pt(6)


def clean_spaces(s: str) -> str:
    s = s.replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


# -----------------------
# Parse meeting info (from Agenda)
# -----------------------
def extract_meeting_info(agenda_doc: Document):
    paras = [p.text.strip() for p in agenda_doc.paragraphs if p.text.strip()]
    head = " ".join(paras[:12]).upper()

    m_bil = re.search(r"BIL\.\s*(\d{2})/(\d{4})", head)
    if not m_bil:
        raise ValueError("Tak jumpa 'BIL. 01/2026' dalam agenda.")
    bil_no = int(m_bil.group(1))
    bil_year = int(m_bil.group(2))

    # Example expected: 12 JANUARI 2026 (ISNIN)
    m_date = re.search(r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})\s*\(([^)]+)\)", head)
    if not m_date:
        # fallback without day
        m_date2 = re.search(r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})", head)
        if not m_date2:
            raise ValueError("Tak jumpa tarikh mesyuarat (contoh '12 JANUARI 2026 (ISNIN)').")
        day_num, month_word, year = m_date2.group(1), m_date2.group(2), m_date2.group(3)
        day_name = ""
    else:
        day_num, month_word, year, day_name = m_date.group(1), m_date.group(2), m_date.group(3), m_date.group(4)

    month_map = {
        "JANUARI": "Januari", "FEBRUARI": "Februari", "MAC": "Mac", "APRIL": "April",
        "MEI": "Mei", "JUN": "Jun", "JULAI": "Julai", "OGOS": "Ogos",
        "SEPTEMBER": "September", "OKTOBER": "Oktober", "NOVEMBER": "November", "DISEMBER": "Disember",
    }
    month_nice = month_map.get(month_word.upper(), month_word.title())
    date_str = f"{int(day_num)} {month_nice} {year}"
    day_str = day_name.title() if day_name else ""

    return {
        "bil_no": bil_no,
        "bil_year": bil_year,
        "meeting_bil": f"Bil.{bil_no:02d}/{bil_year}",
        "date_str": date_str,
        "day_str": day_str,
    }


# -----------------------
# Parse cases from Agenda
# Only PKM + BGN
# -----------------------
def parse_cases_from_agenda(agenda_doc: Document):
    paras = [p.text.strip() for p in agenda_doc.paragraphs if p.text.strip()]

    # Start blocks at "KERTAS MESYUARAT BIL. OSC/..."
    header_re = re.compile(r"^KERTAS\s+MESYUARAT\s+BIL\.\s+(OSC/[A-Z]+/\d{2}/\d{4})", re.IGNORECASE)

    blocks = []
    cur = None
    for ln in paras:
        m = header_re.match(ln)
        if m:
            if cur:
                blocks.append(cur)
            cur = {"code": m.group(1).upper(), "lines": []}
        else:
            if cur:
                cur["lines"].append(ln)
    if cur:
        blocks.append(cur)

    cases = []
    for b in blocks:
        code = b["code"]  # OSC/PKM/36/2026
        if "/PKM/" not in code and "/BGN/" not in code:
            continue

        m2 = re.search(r"OSC/[A-Z]+/(\d{2})/(\d{4})", code)
        if not m2:
            continue
        case_no = int(m2.group(1))
        case_year = int(m2.group(2))

        perunding = ""
        pemohon = ""
        id_permohonan = ""

        desc_lines = []
        for ln in b["lines"]:
            # Key lines
            if re.search(r"\bPerunding\b", ln, re.IGNORECASE) and ":" in ln:
                perunding = clean_spaces(ln.split(":", 1)[1])
                continue
            if re.search(r"\bPemohon\b", ln, re.IGNORECASE) and ":" in ln:
                pemohon = clean_spaces(ln.split(":", 1)[1])
                continue
            if re.search(r"\bNo\.?\s*(Fail|Rujukan)\s*OSC\b", ln, re.IGNORECASE) and ":" in ln:
                id_permohonan = clean_spaces(ln.split(":", 1)[1])
                continue

            # Everything else treat as description/title part
            desc_lines.append(ln)

        # Build Nama Permohonan as multiline (keep the “kemas” look)
        # Clean but preserve line breaks
        nama_raw = "\n".join([clean_spaces(x) for x in desc_lines if x.strip()]).strip()

        # Jenis permohonan field (short like colleague)
        if "/BGN/" in code or "PELAN BANGUNAN" in nama_raw.upper():
            jenis = "Bangunan"
        else:
            jenis = "Kebenaran Merancang"

        cases.append({
            "code": code,
            "case_no": case_no,
            "case_year": case_year,
            "perunding": perunding,
            "pemohon": pemohon,
            "jenis_permohonan": jenis,
            "nama_permohonan": nama_raw,
            "id_permohonan": id_permohonan,
        })

    return cases


# -----------------------
# DOCX builder (NO TEMPLATE)
# -----------------------
def add_top_right_ref(doc: Document, rujukan_kami: str, tarikh: str):
    # 1-row 2-col table: left blank, right contains ref lines
    t = doc.add_table(rows=1, cols=2)
    t.autofit = True
    left = t.cell(0, 0)
    right = t.cell(0, 1)

    # make left empty spacer
    left.text = ""

    # right cell lines
    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("Rujukan Tuan : ")
    run.bold = False
    p.add_run("\nRujukan Kami : ").bold = False
    p.add_run(rujukan_kami)
    p.add_run("\nTarikh        : ")
    p.add_run(tarikh)

    # shrink spacing
    for c in (left, right):
        for pp in c.paragraphs:
            pp.paragraph_format.space_before = Pt(0)
            pp.paragraph_format.space_after = Pt(0)


def add_info_table(doc: Document, perunding, pemohon, jenis, nama, id_permohonan):
    t = doc.add_table(rows=5, cols=3)
    t.autofit = True

    labels = [
        "Kepada (PSP)",
        "Pemilik Projek",
        "Jenis Permohonan",
        "Nama Permohonan",
        "ID Permohonan",
    ]
    values = [perunding, pemohon, jenis, nama, id_permohonan]

    for i in range(5):
        t.cell(i, 0).text = labels[i]
        t.cell(i, 1).text = ":"
        # write multiline nicely
        cell = t.cell(i, 2)
        cell.text = ""
        lines = values[i].splitlines() if values[i] else [""]
        for j, line in enumerate(lines):
            if j == 0:
                cell.paragraphs[0].add_run(line)
            else:
                cell.add_paragraph(line)

    # Bold left labels
    for i in range(5):
        for r in t.cell(i, 0).paragraphs[0].runs:
            r.bold = True

    # Reduce spacing
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)


def add_decision_boxes(doc: Document):
    # 2x2 layout using table
    t = doc.add_table(rows=2, cols=4)
    t.autofit = True

    # row1: [box][LULUS] [box][TOLAK]
    t.cell(0, 0).text = "☐"
    t.cell(0, 1).text = "LULUS"
    t.cell(0, 2).text = "☐"
    t.cell(0, 3).text = "TOLAK"

    # row2: [box][LULUS DENGAN PINDAAN...] [box][TANGGUH]
    t.cell(1, 0).text = "☐"
    t.cell(1, 1).text = "LULUS DENGAN PINDAAN PELAN /\nLULUS BERSYARAT"
    t.cell(1, 2).text = "☐"
    t.cell(1, 3).text = "TANGGUH"

    # Bold option text (not the box)
    for (r, c) in [(0, 1), (0, 3), (1, 1), (1, 3)]:
        for run in t.cell(r, c).paragraphs[0].runs:
            run.bold = True

    # Center boxes
    for (r, c) in [(0, 0), (0, 2), (1, 0), (1, 2)]:
        t.cell(r, c).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # tighten spacing
    for row in t.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)


def build_doc(meeting_info, case):
    doc = Document()
    set_doc_defaults(doc, font_name="Times New Roman", font_size_pt=12)

    # Rujukan Kami format: (1)MBSP/15/1551/(36)2026
    rujukan_kami = f"({meeting_info['bil_no']})MBSP/15/1551/({case['case_no']}){case['case_year']}"
    tarikh = meeting_info["date_str"]

    add_top_right_ref(doc, rujukan_kami=rujukan_kami, tarikh=tarikh)
    doc.add_paragraph("")  # small gap

    add_info_table(
        doc,
        perunding=case["perunding"],
        pemohon=case["pemohon"],
        jenis=case["jenis_permohonan"],
        nama=case["nama_permohonan"],
        id_permohonan=case["id_permohonan"],
    )

    add_horizontal_line(doc)

    # Title
    p = doc.add_paragraph("PEMAKLUMAN KEPUTUSAN MESYUARAT JAWATANKUASA PUSAT SETEMPAT\n(OSC)")
    p.runs[0].bold = True

    doc.add_paragraph("Dengan hormatnya saya diarahkan merujuk perkara di atas.")

    # Paragraph 2 (AUTO from agenda)
    day_part = f" ({meeting_info['day_str']})" if meeting_info["day_str"] else ""
    p2 = (
        f"2.     Adalah dimaklumkan bahawa Mesyuarat Jawatankuasa Pusat Setempat (OSC)\n"
        f"{meeting_info['meeting_bil']} yang bersidang pada {meeting_info['date_str']}{day_part} "
        f"bersetuju untuk memberikan keputusan ke atas permohonan yang telah dikemukakan oleh pihak tuan/puan seperti mana berikut:"
    )
    doc.add_paragraph(p2)

    doc.add_paragraph("")  # gap
    add_decision_boxes(doc)
    doc.add_paragraph("")

    p3 = (
        "3.     Walau bagaimanapun, keputusan muktamad bagi permohonan yang berkenaan adalah tertakluk kepada surat kelulusan / "
        "penolakan yang akan dikeluarkan oleh Jabatan Induk yang memproses."
    )
    doc.add_paragraph(p3)

    doc.add_paragraph("")
    doc.add_paragraph("Sekian, terima kasih.")

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue(), rujukan_kami


# -----------------------
# Streamlit UI (simple)
# -----------------------
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.write("Upload Agenda (DOCX) → auto extract **KM (PKM) & Bangunan (BGN)** → generate Word untuk download (ZIP).")

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

if uploaded:
    agenda_doc = Document(BytesIO(uploaded.getvalue()))

    try:
        meeting_info = extract_meeting_info(agenda_doc)
    except Exception as e:
        st.error(f"Ralat baca Bil/Tarikh dari agenda: {e}")
        st.stop()

    cases = parse_cases_from_agenda(agenda_doc)

    st.success(
        f"Bil Mesyuarat: {meeting_info['meeting_bil']} | Tarikh: {meeting_info['date_str']}"
        + (f" ({meeting_info['day_str']})" if meeting_info["day_str"] else "")
        + f" | Kes PKM+BGN: {len(cases)}"
    )

    if len(cases) == 0:
        st.error("Sistem tak jumpa kes PKM/BGN dalam agenda. Pastikan ada blok 'KERTAS MESYUARAT BIL. OSC/PKM/..' atau 'OSC/BGN/..'.")
        st.stop()

    if st.button("Generate & Download (ZIP)"):
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for c in cases:
                doc_bytes, ruj = build_doc(meeting_info, c)
                # filename style: (1)MBSP-15-1551-(36)2026.docx
                fname = ruj.replace("/", "-") + ".docx"
                zf.writestr(fname, doc_bytes)

        zip_buf.seek(0)
        st.download_button(
            "Download ZIP",
            data=zip_buf.getvalue(),
            file_name=f"pemakluman_{meeting_info['bil_no']:02d}_{meeting_info['bil_year']}.zip",
            mime="application/zip",
        )
