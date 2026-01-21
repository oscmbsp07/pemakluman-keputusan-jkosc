# app.py
import re
import zipfile
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document


# =========================
# Utils: DOCX traversal
# =========================
def iter_all_paragraphs(doc: Document):
    """Yield all paragraphs including those inside tables."""
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def clean_spaces(s: str) -> str:
    s = s.replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


# =========================
# Parse Meeting Info (Agenda)
# =========================
def extract_meeting_info(agenda_doc: Document):
    # Grab first few non-empty paragraphs (agenda header is usually here)
    paras = [p.text.strip() for p in agenda_doc.paragraphs if p.text.strip()]
    head = " ".join(paras[:5]).upper()

    # Bil. 01/2026
    m_bil = re.search(r"BIL\.\s*(\d{2})/(\d{4})", head)
    if not m_bil:
        raise ValueError("Tak jumpa format 'BIL. 01/2026' dalam header agenda.")
    bil_no = int(m_bil.group(1))
    bil_year = int(m_bil.group(2))

    # Date + day e.g. 12 JANUARI 2026 (ISNIN)
    # We look specifically at early lines for the meeting date
    head2 = " ".join(paras[:10]).upper()
    m_date = re.search(
        r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})\s*\(([^)]+)\)",
        head2
    )
    if not m_date:
        # fallback: date without day
        m_date2 = re.search(r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})", head2)
        if not m_date2:
            raise ValueError("Tak jumpa tarikh mesyuarat (contoh '12 JANUARI 2026 (ISNIN)') dalam header agenda.")
        day_num, month_word, year = m_date2.group(1), m_date2.group(2), m_date2.group(3)
        day_name = ""
    else:
        day_num, month_word, year, day_name = m_date.group(1), m_date.group(2), m_date.group(3), m_date.group(4)

    # Title line (optional display)
    meeting_title = paras[0] if paras else f"MESYUARAT OSC BIL. {bil_no:02d}/{bil_year}"

    # Keep Malay month capitalization nice: "Januari"
    month_map = {
        "JANUARI": "Januari", "FEBRUARI": "Februari", "MAC": "Mac", "APRIL": "April",
        "MEI": "Mei", "JUN": "Jun", "JULAI": "Julai", "OGOS": "Ogos",
        "SEPTEMBER": "September", "OKTOBER": "Oktober", "NOVEMBER": "November", "DISEMBER": "Disember",
    }
    month_nice = month_map.get(month_word.upper(), month_word.title())
    date_nice = f"{int(day_num)} {month_nice} {year}"

    day_nice = day_name.title() if day_name else ""
    return {
        "meeting_title": meeting_title,
        "bil_no": bil_no,
        "bil_year": bil_year,
        "date_str": date_nice,
        "day_str": day_nice,
    }


# =========================
# Parse Cases (Agenda blocks)
# =========================
def parse_cases_from_agenda(agenda_doc: Document):
    """
    Parse blocks starting with 'KERTAS MESYUARAT BIL. OSC/...'
    Keep only PKM (KM) and BGN (Bangunan).
    """
    paras = [p.text.strip() for p in agenda_doc.paragraphs if p.text.strip()]

    blocks = []
    current = None

    header_re = re.compile(r"^KERTAS MESYUARAT BIL\.\s+(OSC/[A-Z]+/\d{2}/\d{4})", re.IGNORECASE)

    for line in paras:
        m = header_re.match(line)
        if m:
            # flush previous
            if current:
                blocks.append(current)
            current = {
                "header": line,
                "code": m.group(1).upper(),
                "lines": []
            }
        else:
            if current:
                current["lines"].append(line)

    if current:
        blocks.append(current)

    cases = []
    for b in blocks:
        code = b["code"]  # e.g. OSC/PKM/01/2026
        # filter only KM/Bangunan
        if "/PKM/" not in code and "/BGN/" not in code:
            continue

        # extract case no + year from code
        m2 = re.search(r"OSC/[A-Z]+/(\d{2})/(\d{4})", code)
        if not m2:
            continue
        case_no = int(m2.group(1))
        case_year = int(m2.group(2))

        # Split title vs key-value lines
        title_lines = []
        kv = {}

        for ln in b["lines"]:
            # detect key-value style lines (Pemohon : ..., Perunding : ..., No. Rujukan OSC : ...)
            if re.search(r"\bPemohon\b", ln, re.IGNORECASE) and ":" in ln:
                kv["pemohon"] = clean_spaces(ln.split(":", 1)[1])
                continue
            if re.search(r"\bPerunding\b", ln, re.IGNORECASE) and ":" in ln:
                kv["perunding"] = clean_spaces(ln.split(":", 1)[1])
                continue
            if re.search(r"\bNo\.\s*Rujukan\s*OSC\b", ln, re.IGNORECASE) and ":" in ln:
                kv["no_rujukan_osc"] = clean_spaces(ln.split(":", 1)[1])
                continue

            # if not kv line, treat as title/description line (usually the long "Permohonan ...")
            # stop title when we reached many kvs? (we still keep non-kv lines as title)
            title_lines.append(ln)

        title = "\n".join([clean_spaces(x) for x in title_lines if x.strip()])

        # Determine jenis permohonan text
        title_upper = title.upper()
        if "PELAN BANGUNAN" in title_upper or "/BGN/" in code:
            jenis = "PERMOHONAN PELAN BANGUNAN"
            # colleague style: sometimes they remove leading "Permohonan Pelan Bangunan"
            # We'll gently strip ONLY if it starts with that phrase
            title = re.sub(r"^Permohonan\s+Pelan\s+Bangunan\s*", "", title, flags=re.IGNORECASE).strip()
        else:
            jenis = "PERMOHONAN KEBENARAN MERANCANG"
            # For KM, keep full title (matches KB example)

        # Minimal required fields
        cases.append({
            "code": code,
            "case_no": case_no,
            "case_year": case_year,
            "perunding": kv.get("perunding", ""),
            "pemohon": kv.get("pemohon", ""),
            "jenis_permohonan": jenis,
            "nama_permohonan": title,
            "id_permohonan": kv.get("no_rujukan_osc", ""),
        })

    return cases


# =========================
# Fill Template (colleague format)
# =========================
def replace_line_in_cell(cell, startswith_label: str, new_value: str):
    """
    In a cell that has multiple lines (e.g. 'Rujukan Tuan : ...\\nRujukan Kami : ...'),
    replace the line that starts with startswith_label (case-insensitive).
    """
    txt = cell.text
    lines = txt.splitlines() if txt else []
    out = []
    found = False
    for ln in lines:
        if ln.strip().upper().startswith(startswith_label.upper()):
            # keep original label formatting until colon if possible
            if ":" in ln:
                left = ln.split(":", 1)[0]
                out.append(f"{left} : {new_value}")
            else:
                out.append(f"{startswith_label} : {new_value}")
            found = True
        else:
            out.append(ln)
    if not found:
        # append if not exists
        out.append(f"{startswith_label} : {new_value}")
    cell.text = "\n".join(out)


def fill_top_info_table(doc: Document, data: dict, rujukan_kami: str, tarikh_str: str):
    """
    Find the first table that contains 'Kepada (PSP)' and fill third column values.
    Also set Rujukan Kami + Tarikh in the top-right cell (the small header table).
    """
    # Update header rujukan/tarikh (usually in table[0], last col)
    # We do it robustly: search any cell containing "Rujukan Kami"
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if "Rujukan Kami" in cell.text:
                    replace_line_in_cell(cell, "Rujukan Kami", rujukan_kami)
                if cell.text.strip().startswith("Tarikh") or "Tarikh" in cell.text:
                    # only replace lines that start with Tarikh:
                    replace_line_in_cell(cell, "Tarikh", tarikh_str)

    # Now fill the main info table (Kepada/Pemilik/Jenis/Nama/ID)
    target_table = None
    for t in doc.tables:
        if any("Kepada" in c.text for c in t.row_cells(0)):
            # KB/Sky template has Kepada (PSP) in first row first col
            if "Kepada" in t.cell(0, 0).text:
                target_table = t
                break

    if not target_table:
        raise ValueError("Template tak ada jadual 'Kepada (PSP) / Pemilik Projek / Jenis Permohonan / ...'.")

    # Map rows by label in col0
    for row in target_table.rows:
        label = row.cells[0].text.strip()
        if "Kepada" in label:
            row.cells[2].text = data["perunding"]
        elif "Pemilik Projek" in label:
            row.cells[2].text = data["pemohon"]
        elif "Jenis Permohonan" in label:
            row.cells[2].text = data["jenis_permohonan"]
        elif "Nama Permohonan" in label:
            row.cells[2].text = data["nama_permohonan"]
        elif "ID Permohonan" in label or "No Fail OSC" in label:
            row.cells[2].text = data["id_permohonan"]


def update_perenggan_mesyuarat(doc: Document, meeting_bil: str, meeting_date: str, meeting_day: str):
    """
    Update paragraph that contains 'Mesyuarat Jawatankuasa Pusat Setempat (OSC)'.
    Replace Bil.xx/yyyy and date+day inside it.
    """
    bil_re = re.compile(r"Bil\.\s*\d{2}/\d{4}", re.IGNORECASE)
    date_re = re.compile(r"\d{1,2}\s+[A-Za-z]+\s+\d{4}", re.IGNORECASE)
    day_re = re.compile(r"\([A-Za-z]+\)")

    for p in iter_all_paragraphs(doc):
        if "Mesyuarat Jawatankuasa" in p.text and "Bil." in p.text:
            # Replace within whole paragraph text (simple & reliable)
            txt = p.text
            txt = bil_re.sub(meeting_bil, txt)
            txt = date_re.sub(meeting_date, txt)
            if meeting_day:
                txt = day_re.sub(f"({meeting_day})", txt)
            p.text = txt
            return


def generate_single_doc(template_bytes: bytes, meeting_info: dict, case_data: dict, rujukan_prefix: str = "MBSP/15/1551"):
    """
    Clone template and fill fields.
    """
    doc = Document(BytesIO(template_bytes))

    meeting_bil = f"Bil.{meeting_info['bil_no']:02d}/{meeting_info['bil_year']}"
    meeting_date = meeting_info["date_str"]
    meeting_day = meeting_info["day_str"]

    # Rujukan Kami: (1)MBSP/15/1551/(36)2026
    rujukan_kami = f"({meeting_info['bil_no']}){rujukan_prefix}/({case_data['case_no']}){case_data['case_year']}"
    tarikh_str = meeting_date

    fill_top_info_table(doc, case_data, rujukan_kami=rujukan_kami, tarikh_str=tarikh_str)
    update_perenggan_mesyuarat(doc, meeting_bil=meeting_bil, meeting_date=meeting_date, meeting_day=meeting_day)

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue(), rujukan_kami


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")

st.write("Upload Agenda (Word) → auto extract kes **KM & Bangunan** → jana dokumen Word (ikut format template) untuk download.")

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

# Load template from repo
template_path = Path("templates") / "base_template.docx"
if not template_path.exists():
    st.error("Template tak jumpa: `templates/base_template.docx`")
    st.info("Letak 1 fail template format colleague awak dalam folder `templates/` dan rename jadi `base_template.docx`.")
    st.stop()

template_bytes = template_path.read_bytes()

if uploaded:
    agenda_doc = Document(BytesIO(uploaded.getvalue()))

    try:
        meeting_info = extract_meeting_info(agenda_doc)
    except Exception as e:
        st.error(f"Ralat baca header agenda: {e}")
        st.stop()

    cases = parse_cases_from_agenda(agenda_doc)

    st.success(
        f"Jumpa {meeting_info['meeting_title']} | Tarikh: {meeting_info['date_str']}"
        + (f" ({meeting_info['day_str']})" if meeting_info["day_str"] else "")
        + f" | Kes KM/Bangunan: {len(cases)}"
    )

    if len(cases) == 0:
        st.error("Tiada kes KM/Bangunan dijumpai. Semak agenda—pastikan ada blok 'OSC/PKM/..' atau 'OSC/BGN/..'.")
        st.stop()

    if st.button("Generate & Download (ZIP)"):
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for c in cases:
                doc_bytes, ruj = generate_single_doc(
                    template_bytes=template_bytes,
                    meeting_info=meeting_info,
                    case_data=c,
                    rujukan_prefix="MBSP/15/1551",
                )
                # File name: (1)MBSP-15-1551-(036)2026.docx style (safe filename)
                safe_name = ruj.replace("/", "-").replace("(", "(").replace(")", ")")
                filename = f"{safe_name}.docx"
                zf.writestr(filename, doc_bytes)

        zip_buf.seek(0)
        st.download_button(
            "Download ZIP",
            data=zip_buf.getvalue(),
            file_name=f"pemakluman_{meeting_info['bil_no']:02d}_{meeting_info['bil_year']}.zip",
            mime="application/zip"
        )
