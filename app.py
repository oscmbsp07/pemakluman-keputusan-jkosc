import base64
import io
import re
import zipfile
from datetime import date

import streamlit as st
from docx import Document

# =========================================================
# Embedded base template (DOCX) - no external template file
# Base is taken from contoh output colleague (KB Inn Hotel)
# =========================================================
_TEMPLATE_B64 = """UEsDBBQAAAAIAK3ZNVtFq... (TRUNCATED FOR BREVITY) ..."""

# -----------------------------
# Helpers: template load/save
# -----------------------------
def _load_base_template() -> Document:
    raw = base64.b64decode(_TEMPLATE_B64.encode("ascii"))
    return Document(io.BytesIO(raw))

def _doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# -----------------------------
# Helpers: docx safe replace
# -----------------------------
def _replace_in_runs(paragraph, repl_map: dict):
    # Replace text inside runs (keeps formatting as much as possible)
    for run in paragraph.runs:
        if not run.text:
            continue
        for old, new in repl_map.items():
            if old in run.text:
                run.text = run.text.replace(old, new)

def _replace_in_cell(cell, repl_map: dict):
    for p in cell.paragraphs:
        _replace_in_runs(p, repl_map)

# -----------------------------
# Malay date helpers
# -----------------------------
_MALAY_MONTHS = {
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

_MALAY_WEEKDAYS = {
    0: "Isnin",
    1: "Selasa",
    2: "Rabu",
    3: "Khamis",
    4: "Jumaat",
    5: "Sabtu",
    6: "Ahad",
}

def _parse_malay_date(text: str) -> date | None:
    # e.g. "12 JANUARI 2026" or "12 Januari 2026"
    m = re.search(
        r"(\d{1,2})\s+(JANUARI|FEBRUARI|MAC|APRIL|MEI|JUN|JULAI|OGOS|SEPTEMBER|OKTOBER|NOVEMBER|DISEMBER)\s+(\d{4})",
        text.upper(),
    )
    if not m:
        return None
    d = int(m.group(1))
    mon = _MALAY_MONTHS[m.group(2)]
    y = int(m.group(3))
    return date(y, mon, d)

def _format_malay_date(d: date) -> str:
    # "12 Januari 2026"
    inv = {v: k.title() for k, v in _MALAY_MONTHS.items()}
    return f"{d.day} {inv[d.month]} {d.year}"

def _format_malay_weekday(d: date) -> str:
    return _MALAY_WEEKDAYS[d.weekday()]

# -----------------------------
# Agenda parsing
# -----------------------------
def _extract_all_text(doc: Document) -> list[str]:
    texts = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            texts.append(t)
    # Some agendas put content in tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = (p.text or "").strip()
                    if t:
                        texts.append(t)
    return texts

def _parse_meeting_info(texts: list[str]) -> dict:
    big = "\n".join(texts)

    # Bil mesyuarat: "BIL. 01/2026" or "Bil.01/2026"
    m_bil = re.search(r"\bBIL\.?\s*(\d{1,2})\s*/\s*(\d{4})\b", big, flags=re.IGNORECASE)
    if not m_bil:
        raise ValueError("Tak jumpa 'BIL. xx/yyyy' dalam agenda.")

    bil_no = int(m_bil.group(1))
    year = int(m_bil.group(2))

    # Tarikh mesyuarat
    d = _parse_malay_date(big)
    if not d:
        raise ValueError("Tak jumpa tarikh (cth: 12 JANUARI 2026) dalam agenda.")

    return {
        "bil_no": bil_no,
        "year": year,
        "meeting_date": d,
        "meeting_date_str": _format_malay_date(d),
        "meeting_weekday": _format_malay_weekday(d),
    }

def _parse_cases(texts: list[str]) -> list[dict]:
    cases = []
    i = 0

    while i < len(texts):
        line = texts[i]

        # Example:
        # "KERTAS MESYUARAT BIL. OSC/PKM/001/2026"
        m = re.search(
            r"KERTAS\s+MESYUARAT\s+BIL\.?\s*(OSC/(PKM|BGN)/([^/\s]+)/(\d{4}))",
            line,
            flags=re.IGNORECASE,
        )
        if not m:
            i += 1
            continue

        full_code = m.group(1).upper()
        case_type = m.group(2).upper()  # PKM / BGN
        raw_item = m.group(3)           # 001 or 052(U) etc
        _year = int(m.group(4))

        # numeric agenda item
        digits = re.sub(r"\D", "", raw_item)
        item_no_int = int(digits) if digits else 0
        item_no_disp = f"{item_no_int:02d}" if item_no_int < 10 else str(item_no_int)

        # Nama Permohonan: next lines until key fields
        j = i + 1
        nama_lines = []
        while j < len(texts):
            t = texts[j]
            if re.match(r"^(Perunding|Pemohon|Pemilik\s+Projek|No\.?\s*Rujukan\s*OSC)\s*:", t, flags=re.IGNORECASE):
                break
            if "KERTAS MESYUARAT BIL" in t.upper():
                break
            if re.match(r"^\d+\.0\s+", t):
                break
            nama_lines.append(t)
            j += 1
        nama_permohonan = " ".join([x.strip() for x in nama_lines]).strip()

        # Read key:value lines
        perunding = ""
        pemohon = ""
        no_rujukan_osc = ""

        k = j
        while k < len(texts):
            t = texts[k]
            if "KERTAS MESYUARAT BIL" in t.upper():
                break
            if re.match(r"^\d+\.0\s+", t):
                break

            m_kv = re.match(
                r"^(Perunding|Pemohon|Pemilik\s+Projek|No\.?\s*Rujukan\s*OSC)\s*:\s*(.+)$",
                t,
                flags=re.IGNORECASE,
            )
            if m_kv:
                key = m_kv.group(1).strip().lower()
                val = m_kv.group(2).strip()
                if key.startswith("perunding"):
                    perunding = val
                elif key.startswith("pemohon") or key.startswith("pemilik"):
                    pemohon = val
                elif key.startswith("no"):
                    no_rujukan_osc = val
            k += 1

        # Filter: only KM (PKM) and Bangunan (BGN)
        if case_type not in ("PKM", "BGN"):
            i = k
            continue

        jenis_permohonan = "Kebenaran Merancang" if case_type == "PKM" else "Bangunan"

        cases.append(
            {
                "code": full_code,
                "case_type": case_type,
                "item_no_int": item_no_int,
                "item_no_disp": item_no_disp,
                "year": _year,
                "perunding": perunding,
                "pemohon": pemohon,
                "jenis_permohonan": jenis_permohonan,
                "nama_permohonan": nama_permohonan,
                "id_permohonan": no_rujukan_osc,
            }
        )

        i = k

    return cases

# -----------------------------
# Build surat (docx) per case
# -----------------------------
def _build_doc(meeting: dict, case: dict) -> bytes:
    doc = _load_base_template()

    bil_no = meeting["bil_no"]
    year = meeting["year"]
    meeting_date_str = meeting["meeting_date_str"]
    meeting_weekday = meeting["meeting_weekday"]

    # Rujukan Kami:
    # (BilMesyuarat)MBSP/15/1551/(AgendaItem)Year
    rujukan_kami = f"({bil_no})MBSP/15/1551/({case['item_no_disp']}){year}"

    # Replace common header stuff (keep formatting)
    repl_header = {
        "(1)MBSP/15/1551/(01)2026": rujukan_kami,
        "12 Januari 2026": meeting_date_str,
        "Bil.01/2026": f"Bil.{bil_no:02d}/{year}",
        "(Isnin)": f"({meeting_weekday})",
        " (Isnin)": f" ({meeting_weekday})",
    }

    for p in doc.paragraphs[:25]:
        _replace_in_runs(p, repl_header)

    # Update info table (table 0)
    # 0: Kepada (PSP) | : | value
    # 1: Pemilik Projek| : | value
    # 2: Jenis Permohonan| : | value
    # 3: Nama Permohonan| : | value
    # 4: ID Permohonan| : | value
    if doc.tables:
        t0 = doc.tables[0]
        if len(t0.rows) >= 5 and len(t0.columns) >= 3:
            t0.cell(0, 2).text = case["perunding"] or "-"
            t0.cell(1, 2).text = case["pemohon"] or "-"
            t0.cell(2, 2).text = case["jenis_permohonan"] or "-"
            t0.cell(3, 2).text = case["nama_permohonan"] or "-"
            t0.cell(4, 2).text = case["id_permohonan"] or "-"

    # Safety: kill any stray old sample date/bil if still exists
    safety_map = {
        "13 Mac 2025": meeting_date_str,
        "Bil.05/2025": f"Bil.{bil_no:02d}/{year}",
        "(Khamis)": f"({meeting_weekday})",
    }
    for p in doc.paragraphs:
        _replace_in_runs(p, safety_map)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                _replace_in_cell(cell, safety_map)

    return _doc_to_bytes(doc)

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → auto extract kes KM & Bangunan → jana dokumen Word (ikut format contoh) → download ZIP.")

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

if uploaded:
    try:
        agenda_doc = Document(io.BytesIO(uploaded.read()))
        texts = _extract_all_text(agenda_doc)

        meeting = _parse_meeting_info(texts)
        cases = _parse_cases(texts)

        st.success(
            f"Jumpa Bil.{meeting['bil_no']:02d}/{meeting['year']} | "
            f"Tarikh mesyuarat: {meeting['meeting_date_str']} ({meeting['meeting_weekday']}) | "
            f"Kes KM/Bangunan: {len(cases)}"
        )

        if len(cases) == 0:
            st.error("Tiada kes KM (PKM) atau Bangunan (BGN) dijumpai dalam agenda ni.")
        else:
            with st.expander("Preview kes yang dijumpai (read-only)"):
                for c in cases:
                    st.markdown(f"**{c['code']}**")
                    st.write(f"Perunding: {c['perunding'] or '-'}")
                    st.write(f"Pemohon/Pemilik Projek: {c['pemohon'] or '-'}")
                    st.write(f"ID Permohonan (No. Rujukan OSC): {c['id_permohonan'] or '-'}")
                    st.write(f"Nama Permohonan: {c['nama_permohonan'] or '-'}")
                    st.divider()

            if st.button("Jana & Download (ZIP)", type="primary"):
                with st.spinner("Sedang jana dokumen Word..."):
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
                        for c in cases:
                            out_bytes = _build_doc(meeting, c)
                            fname = f"({meeting['bil_no']})MBSP-15-1551-({c['item_no_disp']}){meeting['year']}.docx"
                            z.writestr(fname, out_bytes)

                    zip_buf.seek(0)
                    st.download_button(
                        "Download ZIP",
                        data=zip_buf.getvalue(),
                        file_name=f"pemakluman_keputusan_Bil{meeting['bil_no']:02d}_{meeting['year']}.zip",
                        mime="application/zip",
                    )

    except Exception as e:
        st.error(f"Ralat baca agenda / jana dokumen: {e}")
