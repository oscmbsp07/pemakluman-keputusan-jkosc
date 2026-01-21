import re
import zipfile
import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document


# =========================
# Config: Malay date helpers
# =========================
MONTHS_MS = {
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
MONTHS_MS_REV = {v: k.title() for k, v in MONTHS_MS.items()}

WEEKDAY_MS = ["Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu", "Ahad"]


# =========================
# Low-level docx utilities
# =========================
def wipe_paragraph(p):
    """Remove all XML children in the paragraph (stronger than clearing runs)."""
    for child in list(p._p):
        p._p.remove(child)


def set_paragraph_runs(p, runs):
    """
    runs: list[tuple[text, bold_bool]]
    """
    wipe_paragraph(p)
    for text, bold in runs:
        r = p.add_run(text)
        r.bold = bool(bold)


def set_cell_text(cell, text):
    """
    Replace first paragraph content while keeping cell/paragraph style as much as possible.
    """
    if not cell.paragraphs:
        cell.text = text
        return
    p = cell.paragraphs[0]
    wipe_paragraph(p)
    p.add_run(text)


# =========================
# Parse agenda (DOCX) -> meeting info + cases
# =========================
def parse_meeting_info(full_text: str):
    """
    Extract:
      - bil_no (int) from "BIL. 01/2026"
      - year (int)
      - meeting_date (datetime.date) from "12 Januari 2026"
      - weekday (Malay) computed
    """
    m = re.search(r"\bBIL\.\s*(\d{1,2})\s*/\s*(\d{4})", full_text, re.IGNORECASE)
    if not m:
        raise ValueError("Tak jumpa format 'BIL. 01/2026' dalam agenda.")
    bil_no = int(m.group(1))
    year = int(m.group(2))

    # Find first Malay date (prefer matching year)
    date_matches = []
    for m2 in re.finditer(r"\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b", full_text):
        day = int(m2.group(1))
        mon = m2.group(2).upper()
        yr = int(m2.group(3))
        if mon in MONTHS_MS:
            date_matches.append((m2.start(), day, MONTHS_MS[mon], yr))
    date_matches.sort(key=lambda x: x[0])

    meeting_date = None
    for _, d, mo, yr in date_matches:
        if yr == year:
            meeting_date = datetime.date(yr, mo, d)
            break
    if meeting_date is None and date_matches:
        _, d, mo, yr = date_matches[0]
        meeting_date = datetime.date(yr, mo, d)

    if meeting_date is None:
        raise ValueError("Tak jumpa tarikh mesyuarat (contoh: 12 Januari 2026) dalam agenda.")

    weekday = WEEKDAY_MS[meeting_date.weekday()]
    return bil_no, year, meeting_date, weekday


def parse_cases(paragraph_texts: list[str]):
    """
    Split agenda into case blocks based on:
      'KERTAS MESYUARAT BIL. ...'
    Assign agenda_no = 1..N (ikut turutan dalam agenda).
    """
    cases = []
    current = None

    for raw in paragraph_texts:
        t = (raw or "").strip()
        if not t:
            continue

        m = re.match(r"^KERTAS\s+MESYUARAT\s+BIL\.\s*(.+?)\s*$", t, re.IGNORECASE)
        if m:
            if current:
                cases.append(current)
            code = m.group(1).strip()
            current = {"code": code, "lines": [t]}
        else:
            if current:
                current["lines"].append(t)

    if current:
        cases.append(current)

    for i, c in enumerate(cases, start=1):
        c["agenda_no"] = i

    return cases


def extract_fields(case: dict):
    """
    Extract fields inside one case block.
    Agenda format usually has:
      KERTAS MESYUARAT BIL. OSC/PKM/.. atau OSC/BGN/..
      PERMOHONAN ...
      ALAMAT: ...
      Perunding: ...
      Pemohon: ...
      No. Rujukan OSC: ...

    Output fields:
      perunding, pemohon, nama_permohonan (permohonan + alamat), rujukan_osc, jenis
    """
    lines = case["lines"]
    code = case["code"]

    # Determine jenis
    if "/PKM/" in code:
        jenis = "PERMOHONAN KEBENARAN MERANCANG"
    elif "/BGN/" in code:
        jenis = "PERMOHONAN PELAN BANGUNAN"
    else:
        jenis = ""

    # Helper: get value after "Label:"
    def find_after(prefix_list):
        for l in lines:
            for pref in prefix_list:
                if l.lower().startswith(pref.lower()):
                    parts = l.split(":", 1)
                    return parts[1].strip() if len(parts) == 2 else ""
        return ""

    # Collect permohonan lines until we hit ALAMAT / Perunding / Pemohon / No. Rujukan OSC
    stop_starts = (
        "ALAMAT",
        "Perunding",
        "Pemohon",
        "No. Rujukan OSC",
        "No Rujukan OSC",
        "No. Rujukan  OSC",
        "Lokasi",
        "Koordinat",
    )
    perm_parts = []
    for l in lines[1:]:
        if any(l.strip().startswith(s) for s in stop_starts):
            break
        perm_parts.append(l.strip())
    permohonan = " ".join(perm_parts).strip()

    # Address can span multiple lines: start at "ALAMAT:" and continue until next known label
    alamat = ""
    collecting = False
    alamat_parts = []
    for l in lines[1:]:
        if l.strip().startswith("ALAMAT"):
            collecting = True
            parts = l.split(":", 1)
            if len(parts) == 2:
                alamat_parts.append(parts[1].strip())
            continue
        if collecting:
            if any(l.strip().startswith(s) for s in ("Perunding", "Pemohon", "No. Rujukan OSC", "No Rujukan OSC", "Lokasi", "Koordinat")):
                break
            alamat_parts.append(l.strip())
    alamat = " ".join(alamat_parts).strip()

    nama_permohonan = (permohonan + (" " + alamat if alamat else "")).strip()

    perunding = find_after(["Perunding"])
    pemohon = find_after(["Pemohon"])
    rujukan_osc = find_after(["No. Rujukan OSC", "No Rujukan OSC", "No. Rujukan  OSC"])

    return {
        "agenda_no": case["agenda_no"],
        "kertas": code,
        "jenis": jenis,
        "nama_permohonan": nama_permohonan,
        "perunding": perunding,
        "pemohon": pemohon,
        "rujukan_osc": rujukan_osc,
    }


# =========================
# Fill template docx
# =========================
def update_header(doc: Document, rujukan_kami: str, tarikh_str: str):
    """
    Update paragraphs containing:
      - "Rujukan Kami"
      - "Tarikh"
    """
    for p in doc.paragraphs:
        if "Rujukan Kami" in p.text:
            set_paragraph_runs(p, [("Rujukan Kami\t: ", False), (rujukan_kami, False)])
        if p.text.strip().startswith("Tarikh"):
            set_paragraph_runs(p, [("Tarikh\t\t: ", False), (tarikh_str, False)])


def update_top_fields(doc: Document, perunding: str, pemohon: str, jenis_permohonan: str, nama_permohonan: str, rujukan_osc: str):
    """
    Support 2 formats:
    1) Template has a top table with labels like 'Kepada (PSP)', 'Pemilik Projek', 'Jenis Permohonan', ...
    2) Template uses paragraphs with labels.
    """
    def fill_table_label(label: str, value: str) -> bool:
        for t in doc.tables:
            for r in t.rows:
                for i, cell in enumerate(r.cells):
                    if label.lower() in (cell.text or "").strip().lower():
                        if i + 1 < len(r.cells):
                            set_cell_text(r.cells[i + 1], value)
                            return True
        return False

    # Try tables first
    filled_any = False
    filled_any |= fill_table_label("Kepada (PSP)", perunding)
    filled_any |= fill_table_label("Pemilik Projek", pemohon)
    filled_any |= fill_table_label("Jenis Permohonan", jenis_permohonan)
    filled_any |= fill_table_label("Nama Permohonan", nama_permohonan)
    filled_any |= fill_table_label("ID Permohonan", rujukan_osc)

    if filled_any:
        return

    # Fallback: paragraphs
    for p in doc.paragraphs:
        txt = (p.text or "").strip()
        if txt.startswith("Perunding"):
            set_paragraph_runs(p, [("Perunding\t\t: ", False), (perunding, False)])
        elif txt.startswith("Pemilik Projek"):
            set_paragraph_runs(p, [("Pemilik Projek\t: ", False), (pemohon, False)])
        elif "Jenis Permohonan" in txt:
            set_paragraph_runs(p, [("Jenis Permohonan\t: ", False), (jenis_permohonan, False)])
        elif "Nama Permohonan" in txt:
            set_paragraph_runs(p, [("Nama Permohonan\t: ", False), (nama_permohonan, False)])
        elif "No Fail OSC" in txt or "No. Rujukan OSC" in txt:
            set_paragraph_runs(p, [("No Fail OSC\t\t: ", False), (rujukan_osc, False)])


def update_meeting_paragraph(doc: Document, bil_str: str, tarikh_str: str, weekday_ms: str):
    """
    Replace the paragraph that contains 'Adalah dimaklumkan...' & 'bersidang pada ...'
    with the correct Bil & Tarikh from agenda.
    """
    for p in doc.paragraphs:
        if "Adalah dimaklumkan" in p.text and "bersidang pada" in p.text:
            runs = [
                ("Adalah dimaklumkan bahawa ", False),
                ("Mesyuarat Jawatankuasa Pusat Setempat (OSC) ", True),
                (f"Bil.{bil_str} ", True),
                ("yang bersidang pada ", False),
                (f"{tarikh_str} ({weekday_ms}) ", True),
                ("bersetuju untuk memberikan keputusan ke atas permohonan yang telah dikemukakan oleh pihak tuan/ puan seperti mana berikut:", False),
            ]
            set_paragraph_runs(p, runs)
            return


def build_rujukan_kami(meeting_bil_no: int, year: int, agenda_no: int) -> str:
    """
    Ikut rule yang awak explain:
      (BilNo)MBSP/15/1551/(AgendaNo)Year
    """
    return f"({meeting_bil_no})MBSP/15/1551/({agenda_no}){year}"


def date_ms_str(d: datetime.date) -> str:
    return f"{d.day} {MONTHS_MS_REV[d.month]} {d.year}"


def generate_one_doc(template_bytes: bytes, meeting_bil_no: int, year: int, meeting_date: datetime.date, weekday_ms: str, case_fields: dict) -> bytes:
    """
    Generate 1 docx based on template bytes.
    """
    doc = Document(BytesIO(template_bytes))

    tarikh_str = date_ms_str(meeting_date)
    bil_str = f"{meeting_bil_no:02d}/{year}"  # keep Bil.01/2026 style in body

    rujukan_kami = build_rujukan_kami(meeting_bil_no, year, case_fields["agenda_no"])

    update_header(doc, rujukan_kami, tarikh_str)
    update_top_fields(
        doc,
        perunding=case_fields["perunding"],
        pemohon=case_fields["pemohon"],
        jenis_permohonan=case_fields["jenis"],
        nama_permohonan=case_fields["nama_permohonan"],
        rujukan_osc=case_fields["rujukan_osc"],
    )
    update_meeting_paragraph(doc, bil_str, tarikh_str, weekday_ms)

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → auto extract kes KM (PKM) & Bangunan (BGN) → jana dokumen Word ikut template → download ZIP.")

# Locate template in repo
TEMPLATE_DIR = Path(__file__).parent / "templates"
template_candidates = []
if TEMPLATE_DIR.exists():
    template_candidates = sorted([p for p in TEMPLATE_DIR.glob("*.docx") if p.is_file()])

if not template_candidates:
    st.error("Template .docx tak jumpa dalam folder /templates. Sila upload template Word ke dalam folder 'templates' dalam repo (contoh: templates/template.docx) dan redeploy.")
    st.stop()

# Pick first template docx found (simple, no dropdown)
template_path = template_candidates[0]
st.info(f"Template digunakan: {template_path.name}")

agenda_file = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

if agenda_file:
    agenda_bytes = agenda_file.getvalue()
    agenda_doc = Document(BytesIO(agenda_bytes))
    paras = [p.text for p in agenda_doc.paragraphs]
    full_text = "\n".join(paras)

    try:
        meeting_bil_no, year, meeting_date, weekday_ms = parse_meeting_info(full_text)
    except Exception as e:
        st.error(f"Tak boleh baca info mesyuarat dari agenda. Error: {e}")
        st.stop()

    cases = parse_cases(paras)
    extracted = [extract_fields(c) for c in cases]

    # Filter KM (PKM) & Bangunan (BGN) only
    km_bgn = [x for x in extracted if ("/PKM/" in x["kertas"] or "/BGN/" in x["kertas"])]

    st.success(f"Jumpa BIL. {meeting_bil_no:02d}/{year} | Tarikh mesyuarat: {date_ms_str(meeting_date)} ({weekday_ms}) | Kes KM/Bangunan: {len(km_bgn)}")

    if len(km_bgn) == 0:
        st.error("Tiada kes PKM/BGN dijumpai dalam agenda. Semak format agenda (mesti ada 'KERTAS MESYUARAT BIL. OSC/PKM/..' atau 'OSC/BGN/..').")
        st.stop()

    if st.button("Jana & Muat Turun (ZIP)", type="primary"):
        with st.spinner("Sedang jana dokumen Word..."):
            template_bytes = template_path.read_bytes()

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for case in km_bgn:
                    # Generate docx
                    doc_bytes = generate_one_doc(
                        template_bytes=template_bytes,
                        meeting_bil_no=meeting_bil_no,
                        year=year,
                        meeting_date=meeting_date,
                        weekday_ms=weekday_ms,
                        case_fields=case,
                    )

                    # Filename (Word) - safe and consistent
                    # Example: (1)MBSP-15-1551-(36)2026.docx
                    rujukan_kami = build_rujukan_kami(meeting_bil_no, year, case["agenda_no"])
                    safe_name = rujukan_kami.replace("/", "-") + ".docx"
                    zf.writestr(safe_name, doc_bytes)

            zip_buffer.seek(0)

        st.download_button(
            label="Download ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"Pemakluman_Keputusan_Bil_{meeting_bil_no:02d}_{year}.zip",
            mime="application/zip",
        )
