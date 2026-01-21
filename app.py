import re
import zipfile
from io import BytesIO
from dataclasses import dataclass
from typing import List, Optional, Tuple

import streamlit as st
from docx import Document


# =========================
# Config
# =========================
TEMPLATE_PATH = "templates/template dokumen panggilan.docx"

MONTHS_MS = {
    "januari": 1,
    "februari": 2,
    "mac": 3,
    "april": 4,
    "mei": 5,
    "jun": 6,
    "julai": 7,
    "ogos": 8,
    "september": 9,
    "oktober": 10,
    "november": 11,
    "disember": 12,
}
MONTHS_MS_INV = {v: k.capitalize() for k, v in MONTHS_MS.items()}

DAYS_MS = ["Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu", "Ahad"]


# =========================
# Data model
# =========================
@dataclass
class MeetingInfo:
    bil_no: int
    year: int
    date_str: str           # "12 Januari 2026"
    day_str: str            # "Isnin"
    bil_str: str            # "Bil.01/2026"


@dataclass
class CaseInfo:
    seq: int                # running number (only PKM/BGN)
    paper_no: str           # "OSC/PKM/01/2026"
    nama_permohonan: str
    pemilik_projek: str     # Pemohon
    perunding: str
    no_fail_osc: str
    jenis_permohonan: str   # "PERMOHONAN KEBENARAN MERANCANG" / "PERMOHONAN PELAN BANGUNAN"


# =========================
# Helpers: date parsing
# =========================
def parse_malay_date(text: str):
    import datetime
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text, re.IGNORECASE)
    if not m:
        return None
    d = int(m.group(1))
    mon = m.group(2).lower()
    y = int(m.group(3))
    if mon not in MONTHS_MS:
        return None
    return datetime.date(y, MONTHS_MS[mon], d)


def format_malay_date(dt):
    return f"{dt.day} {MONTHS_MS_INV[dt.month]} {dt.year}"


# =========================
# Template run-safe editing
# (avoid p.text=... (it kills formatting))
# =========================
def clone_run_format(src, dst):
    dst.bold = src.bold
    dst.italic = src.italic
    dst.underline = src.underline
    dst.font.name = src.font.name
    dst.font.size = src.font.size
    if src.font.color and src.font.color.rgb:
        dst.font.color.rgb = src.font.color.rgb


def set_value_after_colon(paragraph, new_value: str):
    """
    Find first ':' in runs, keep label part, replace ONLY the value part after colon.
    Preserves formatting by reusing formatting from first value-run (if any).
    """
    runs = paragraph.runs
    colon_idx = None
    for i, r in enumerate(runs):
        if ":" in r.text:
            colon_idx = i
            break

    if colon_idx is None:
        # fallback
        paragraph.add_run(" " + new_value)
        return

    fmt_src = runs[colon_idx + 1] if colon_idx + 1 < len(runs) else runs[colon_idx]

    # remove all runs after colon
    for r in runs[colon_idx + 1:]:
        r._element.getparent().remove(r._element)

    # ensure colon run has a space at end
    if not paragraph.runs[colon_idx].text.endswith(" "):
        paragraph.runs[colon_idx].text = paragraph.runs[colon_idx].text + " "

    new_run = paragraph.add_run(new_value)
    clone_run_format(fmt_src, new_run)


def replace_regex_in_runs(paragraph, pattern: str, repl: str):
    for r in paragraph.runs:
        r.text = re.sub(pattern, repl, r.text)


# =========================
# Parse Agenda
# =========================
LABELS = [
    "PEMOHON",
    "PERUNDING",
    "LOKASI",
    "KOORDINAT",
    "NO. RUJUKAN",
    "NO. RUJUKAN OSC",
    "NO. FAIL OSC",
    "NO FAIL OSC",
]


def is_label_line(line: str) -> bool:
    return any(re.match(rf"^{re.escape(lab)}\s*:", line, re.IGNORECASE) for lab in LABELS)


def get_labeled_value(block: List[str], label_variants: List[str]) -> str:
    """
    Find 'LABEL: value' then append following lines until next LABEL: ...
    """
    for i, line in enumerate(block):
        for lab in label_variants:
            if re.match(rf"^{re.escape(lab)}\s*:", line, re.IGNORECASE):
                val = line.split(":", 1)[1].strip()
                cont = []
                for nxt in block[i + 1:]:
                    if is_label_line(nxt):
                        break
                    if nxt.strip():
                        cont.append(nxt.strip())
                if cont:
                    val = (val + " " + " ".join(cont)).strip()
                return val
    return ""


def parse_meeting_info(paras: List[str]) -> MeetingInfo:
    """
    Extract Bil.xx/yyyy + Tarikh mesyuarat from top part of agenda.
    """
    head = " ".join(paras[:250])

    m = re.search(r"Bil\.\s*0*(\d+)\s*/\s*(\d{4})", head, re.IGNORECASE)
    if not m:
        raise ValueError("Tak jumpa 'Bil. xx/yyyy' dalam agenda.")
    bil_no = int(m.group(1))
    year = int(m.group(2))
    bil_str = f"Bil.{str(bil_no).zfill(2)}/{year}"

    # find first Malay date with matching year
    dt = None
    for t in paras[:400]:
        dtx = parse_malay_date(t)
        if dtx and dtx.year == year:
            dt = dtx
            break
    if not dt:
        raise ValueError("Tak jumpa Tarikh mesyuarat (cth '12 Januari 2026') dalam agenda.")

    date_str = format_malay_date(dt)
    day_str = DAYS_MS[dt.weekday()]
    return MeetingInfo(bil_no=bil_no, year=year, date_str=date_str, day_str=day_str, bil_str=bil_str)


def parse_cases_from_agenda(paras: List[str]) -> List[CaseInfo]:
    """
    Extract ONLY OSC/PKM & OSC/BGN cases.
    """
    case_idxs = [i for i, t in enumerate(paras) if re.match(r"^KERTAS MESYUARAT BIL\.\s*OSC/(PKM|BGN)/", t, re.IGNORECASE)]
    cases: List[CaseInfo] = []

    for seq, idx in enumerate(case_idxs, start=1):
        header = paras[idx].strip()
        m = re.match(r"^KERTAS MESYUARAT BIL\.\s*(OSC/(PKM|BGN)/[^\s]+)", header, re.IGNORECASE)
        paper_no = m.group(1).strip() if m else header

        # collect block until next case
        end = case_idxs[seq] if seq < len(case_idxs) else len(paras)
        block_all = [x.strip() for x in paras[idx + 1:end] if x is not None]

        # build nama_permohonan from multiple lines until first label line (Pemohon/Perunding/etc.)
        title_parts = []
        rest_block_start = 0
        for j, line in enumerate(block_all):
            if not line.strip():
                continue
            if is_label_line(line):
                rest_block_start = j
                break
            title_parts.append(line.strip())
            rest_block_start = j + 1

        nama_permohonan = " ".join(title_parts).strip()

        rest_block = block_all[rest_block_start:]

        pemohon = get_labeled_value(rest_block, ["PEMOHON"])
        perunding = get_labeled_value(rest_block, ["PERUNDING"])
        no_fail_osc = get_labeled_value(rest_block, ["NO. RUJUKAN OSC", "NO. FAIL OSC", "NO FAIL OSC"])

        # jenis permohonan
        jenis = ""
        if re.search(r"Kebenaran\s+Merancang", nama_permohonan, re.IGNORECASE):
            jenis = "PERMOHONAN KEBENARAN MERANCANG"
        elif re.search(r"Pelan\s+Bangunan", nama_permohonan, re.IGNORECASE):
            jenis = "PERMOHONAN PELAN BANGUNAN"
        else:
            # still PKM/BGN - but fallback if wording pelik
            if "/PKM/" in paper_no.upper():
                jenis = "PERMOHONAN KEBENARAN MERANCANG"
            elif "/BGN/" in paper_no.upper():
                jenis = "PERMOHONAN PELAN BANGUNAN"

        cases.append(
            CaseInfo(
                seq=seq,
                paper_no=paper_no,
                nama_permohonan=nama_permohonan,
                pemilik_projek=pemohon,
                perunding=perunding,
                no_fail_osc=no_fail_osc,
                jenis_permohonan=jenis,
            )
        )

    return cases


# =========================
# Generate DOCX per case
# =========================
def build_rujukan_kami(meeting: MeetingInfo, case: CaseInfo) -> str:
    # Inside document: (1)MBSP/15/1551/(01)2026
    return f"({meeting.bil_no})MBSP/15/1551/({str(case.seq).zfill(2)}){meeting.year}"


def build_filename(meeting: MeetingInfo, case: CaseInfo) -> str:
    # Safe filename (no slash): (1)MBSP-15-1551-(001)2026.docx
    return f"({meeting.bil_no})MBSP-15-1551-({str(case.seq).zfill(3)}){meeting.year}.docx"


def generate_doc_bytes(meeting: MeetingInfo, case: CaseInfo) -> bytes:
    doc = Document(TEMPLATE_PATH)

    # 1) Header fields (tab-based lines)
    rujukan_kami = build_rujukan_kami(meeting, case)

    for p in doc.paragraphs:
        tx = (p.text or "").strip()

        # top header
        if tx.startswith("Rujukan Kami"):
            set_value_after_colon(p, rujukan_kami)
        elif tx.startswith("Tarikh"):
            set_value_after_colon(p, meeting.date_str)

        # case info block
        elif tx.startswith("Perunding"):
            set_value_after_colon(p, case.perunding)
        elif tx.startswith("Pemilik Projek"):
            set_value_after_colon(p, case.pemilik_projek)
        elif tx.startswith("Jenis Permohonan"):
            set_value_after_colon(p, case.jenis_permohonan)
        elif tx.startswith("Nama Permohonan"):
            set_value_after_colon(p, case.nama_permohonan)
        elif tx.startswith("No Fail OSC"):
            set_value_after_colon(p, case.no_fail_osc)

        # 2) Body paragraph that contains meeting bil/date/day
        # Example in template: "... Mesyuarat ... Bil.05/2025 ... 13 Mac 2025 (Khamis) ..."
        if "Mesyuarat" in tx and "Bil." in tx and "bersidang" in tx:
            replace_regex_in_runs(p, r"Bil\.\s*\d+\s*/\s*\d{4}", meeting.bil_str)
            replace_regex_in_runs(p, r"\d{1,2}\s+[A-Za-z]+\s+\d{4}", meeting.date_str)
            replace_regex_in_runs(p, r"\((Isnin|Selasa|Rabu|Khamis|Jumaat|Sabtu|Ahad)\)", f"({meeting.day_str})")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()


def generate_zip(meeting: MeetingInfo, cases: List[CaseInfo]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for c in cases:
            doc_bytes = generate_doc_bytes(meeting, c)
            z.writestr(build_filename(meeting, c), doc_bytes)
    return buf.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → sistem extract OSC/PKM & OSC/BGN → generate dokumen panggilan (Word) untuk download (ZIP).")

if not st.secrets and True:
    # no-op; just to avoid lint complaints
    pass

uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

st.markdown("---")

if uploaded is None:
    st.info("Sila upload Agenda (DOCX).")
    st.stop()

# Read agenda bytes
agenda_bytes = uploaded.read()

# Parse
try:
    agenda_doc = Document(BytesIO(agenda_bytes))
    paras = [(p.text or "").strip() for p in agenda_doc.paragraphs]
    meeting = parse_meeting_info(paras)
    cases = parse_cases_from_agenda(paras)
except Exception as e:
    st.error(f"❌ Gagal baca agenda: {e}")
    st.stop()

# Show summary
st.success(f"✅ Jumpa {meeting.bil_str} | Tarikh: {meeting.date_str} ({meeting.day_str}) | Kes PKM/BGN: {len(cases)}")

with st.expander("Semakan (auto dari agenda)", expanded=True):
    st.write(f"**Bil Mesyuarat:** {meeting.bil_str}")
    st.write(f"**Tarikh Mesyuarat:** {meeting.date_str} ({meeting.day_str})")
    st.write(f"**Jumlah kes PKM/BGN:** {len(cases)}")

    if len(cases) > 0:
        st.write("**Contoh 5 kes terawal (ringkas):**")
        for c in cases[:5]:
            st.write(f"- ({c.seq}) {c.paper_no} | {c.jenis_permohonan} | {c.no_fail_osc}")

# Generate button
col1, col2 = st.columns([1, 3])
with col1:
    gen = st.button("Jana & Download (ZIP)", type="primary", use_container_width=True)

if gen:
    # template existence check
    try:
        _ = Document(TEMPLATE_PATH)
    except Exception:
        st.error(f"❌ Template tak jumpa / tak boleh dibuka: '{TEMPLATE_PATH}'. Pastikan file tu ada dalam folder 'templates/'.")
        st.stop()

    if len(cases) == 0:
        st.error("❌ Tiada kes OSC/PKM atau OSC/BGN dijumpai dalam agenda.")
        st.stop()

    with st.spinner("Sedang jana dokumen..."):
        zip_bytes = generate_zip(meeting, cases)

    st.success("✅ Siap jana dokumen.")
    st.download_button(
        label="Download ZIP",
        data=zip_bytes,
        file_name=f"Pemakluman_{meeting.bil_str.replace('/','-')}_{meeting.year}.zip",
        mime="application/zip",
        use_container_width=True,
    )
