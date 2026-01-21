import re
import io
import zipfile
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

import streamlit as st
from docx import Document
from docx.text.paragraph import Paragraph


# =========================
# UI CONFIG
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → sistem extract KM & Bangunan → generate dokumen (Word) ikut template → download ZIP.")


# =========================
# HELPERS
# =========================
MONTH_MAP = {
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

def normalize_spaces(s: str) -> str:
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def format_tarikh_malay(raw: str) -> str:
    raw = normalize_spaces(raw).upper()
    m = re.search(r"(\d{1,2})\s+([A-Z]+)\s+(\d{4})", raw)
    if not m:
        return raw.title()
    day = int(m.group(1))
    mon = m.group(2)
    year = m.group(3)
    mon2 = MONTH_MAP.get(mon, mon.title())
    return f"{day} {mon2} {year}"

def extract_bil_year(text: str) -> Optional[Tuple[int, int]]:
    t = text.upper()
    m = re.search(r"\bBIL\.?\s*0*(\d{1,2})\s*/\s*(\d{4})\b", t)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))

def extract_tarikh_mesyuarat(lines: List[str]) -> Optional[str]:
    for ln in lines:
        u = ln.upper()
        if "TARIKH" in u:
            m = re.search(r"TARIKH\s*[:\-]\s*(\d{1,2}\s+[A-Z]+\s+\d{4})", u)
            if m:
                return format_tarikh_malay(m.group(1))
    for ln in lines[:80]:
        u = ln.upper()
        m = re.search(r"\b(\d{1,2}\s+[A-Z]+\s+\d{4})\b", u)
        if m:
            return format_tarikh_malay(m.group(1))
    return None


# =========================
# PARSE AGENDA (KM & BANGUNAN)
# =========================
@dataclass
class CaseItem:
    case_no: str
    kertas_bil: str
    jenis: str              # KM / BANGUNAN
    jenis_full: str
    perunding: str
    pemohon: str
    nama_permohonan: str
    id_permohonan: str

def read_docx_lines(file_bytes: bytes) -> List[str]:
    doc = Document(io.BytesIO(file_bytes))
    lines = []
    for p in doc.paragraphs:
        t = normalize_spaces(p.text) if p.text else ""
        if t:
            lines.append(t)
    return lines

def parse_case_block(block_lines: List[str]) -> Optional[CaseItem]:
    text_all = "\n".join(block_lines)

    m_kertas = re.search(r"KERTAS\s+MESYUARAT\s+BIL\.?\s*[:\-]?\s*([A-Z0-9/]+)", text_all.upper())
    if not m_kertas:
        return None
    kertas = m_kertas.group(1).strip()

    jenis = None
    jenis_full = None

    # KM
    if "PKM" in kertas or re.search(r"\bKEBENARAN\s+MERANCANG\b", text_all.upper()):
        jenis = "KM"
        jenis_full = "PERMOHONAN KEBENARAN MERANCANG"

    # Bangunan
    if jenis is None and ("PB" in kertas or ("BANGUNAN" in text_all.upper() and "PELAN" in text_all.upper())):
        jenis = "BANGUNAN"
        jenis_full = "PERMOHONAN PELAN BANGUNAN"

    if jenis is None:
        return None  # bukan KM/Bangunan

    m_no = re.search(r"/(\d{1,3})/(\d{4})$", kertas)
    case_no = m_no.group(1).zfill(2) if m_no else "00"

    def pick(prefix: str) -> str:
        for ln in block_lines:
            if ln.upper().startswith(prefix.upper()):
                parts = re.split(r"\s*[:\-]\s*", ln, maxsplit=1)
                if len(parts) == 2:
                    return parts[1].strip()
        return ""

    perunding = pick("Perunding") or "-"
    pemohon = pick("Pemohon") or "-"
    id_permohonan = pick("No Fail OSC") or pick("No. Fail OSC") or pick("No Fail") or pick("No. Fail") or "-"

    nama_permohonan = ""
    for ln in block_lines:
        if ln.upper().startswith("PERMOHONAN "):
            nama_permohonan = ln.strip()
            break
    if not nama_permohonan:
        nama_permohonan = "-"

    # fallback ID kalau tak ada label (kadang ada MBSP/... dalam perenggan)
    if id_permohonan == "-" or not id_permohonan:
        mm = re.search(r"\bMBSP/[A-Z0-9/().\- ]+\b", text_all)
        if mm:
            id_permohonan = normalize_spaces(mm.group(0))

    return CaseItem(
        case_no=case_no,
        kertas_bil=kertas,
        jenis=jenis,
        jenis_full=jenis_full,
        perunding=perunding,
        pemohon=pemohon,
        nama_permohonan=nama_permohonan,
        id_permohonan=id_permohonan,
    )

def parse_agenda(file_bytes: bytes) -> Tuple[Dict, List[CaseItem]]:
    lines = read_docx_lines(file_bytes)

    bil_year = None
    for ln in lines[:120]:
        by = extract_bil_year(ln)
        if by:
            bil_year = by
            break
    if not bil_year:
        for ln in lines:
            by = extract_bil_year(ln)
            if by:
                bil_year = by
                break

    tarikh_mesyuarat = extract_tarikh_mesyuarat(lines)

    info = {
        "bil": bil_year[0] if bil_year else None,
        "year": bil_year[1] if bil_year else None,
        "bil_str": f"{bil_year[0]:02d}/{bil_year[1]}" if bil_year else "-",
        "tarikh_mesyuarat": tarikh_mesyuarat or "-",
    }

    # blocks by "KERTAS MESYUARAT BIL"
    blocks = []
    cur = []
    for ln in lines:
        if ln.upper().startswith("KERTAS MESYUARAT BIL"):
            if cur:
                blocks.append(cur)
            cur = [ln]
        else:
            if cur:
                cur.append(ln)
    if cur:
        blocks.append(cur)

    cases: List[CaseItem] = []
    for b in blocks:
        item = parse_case_block(b)
        if item:
            cases.append(item)

    def keyfn(x: CaseItem):
        try:
            return int(x.case_no)
        except:
            return 9999
    cases.sort(key=keyfn)

    return info, cases


# =========================
# WORD TEMPLATE FILL (TRANSFER ONLY)
# =========================
def replace_paragraph_text_keep_style(paragraph: Paragraph, new_text: str) -> None:
    for r in paragraph.runs[::-1]:
        r._element.getparent().remove(r._element)
    paragraph.add_run(new_text)

def set_singleline_prefix(doc: Document, prefix: str, value: str) -> bool:
    pref_u = prefix.upper()
    for p in doc.paragraphs:
        t = normalize_spaces(p.text)
        if t.upper().startswith(pref_u):
            replace_paragraph_text_keep_style(p, f"{prefix} {value}".rstrip())
            return True
    return False

def set_field_in_template(doc: Document, label: str, value: str) -> bool:
    # cari paragraph yang bermula label (contoh: "Perunding")
    target = label.strip().upper()
    for p in doc.paragraphs:
        t = normalize_spaces(p.text)
        if t and t.upper().startswith(target):
            replace_paragraph_text_keep_style(p, f"\t{label}\t:\t{value}")
            return True
    return False

def build_rujukan_kami(bil: int, year: int, case_no: str) -> str:
    return f"({bil})MBSP/15/1551/({case_no}){year}"

def generate_doc_for_case(template_bytes: bytes, info: Dict, case: CaseItem) -> bytes:
    doc = Document(io.BytesIO(template_bytes))

    bil = info.get("bil") or 0
    year = info.get("year") or 0
    tarikh = info.get("tarikh_mesyuarat", "-")

    # Rujukan Kami + Tarikh (auto ikut tarikh mesyuarat)
    if bil and year:
        set_singleline_prefix(doc, "Rujukan Kami:", build_rujukan_kami(bil, year, case.case_no))
    set_singleline_prefix(doc, "Tarikh:", tarikh)

    # TRANSFER DATA SAHAJA
    set_field_in_template(doc, "Perunding", case.perunding)
    set_field_in_template(doc, "Pemohon", case.pemohon)
    set_field_in_template(doc, "Jenis Permohonan", case.jenis_full)
    set_field_in_template(doc, "Nama Permohonan", case.nama_permohonan)
    set_field_in_template(doc, "ID Permohonan", case.id_permohonan)

    # Keputusan: BIAR KOSONG / BIAR TEMPLATE ASAL (staf isi manual)
    # -> Kita tak sentuh langsung mana-mana ayat keputusan.

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def make_zip(docs: List[Tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for filename, b in docs:
            z.writestr(filename, b)
    return buf.getvalue()


# =========================
# APP FLOW (NO PILIH-PILIH)
# =========================
uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

TEMPLATE_PATH = "templates/template dokumen panggilan.docx"

if uploaded is None:
    st.info("Upload Agenda dulu. Sistem akan auto jana semua dokumen KM & Bangunan (transfer data sahaja).")
    st.stop()

agenda_bytes = uploaded.read()

try:
    with open(TEMPLATE_PATH, "rb") as f:
        template_bytes = f.read()
except Exception:
    st.error("Template tak jumpa. Pastikan ada: `templates/template dokumen panggilan.docx`")
    st.stop()

info, cases = parse_agenda(agenda_bytes)

st.success(f"Bil: {info.get('bil_str','-')} | Tarikh Mesyuarat: {info.get('tarikh_mesyuarat','-')} | Kes KM/Bangunan: {len(cases)}")

if len(cases) == 0:
    st.error("Sistem tak jumpa kes KM/Bangunan dalam agenda (format agenda tak match parser).")
    st.stop()

st.subheader("Preview (auto detect) — untuk semak je")
for c in cases[:10]:
    st.write(f"- {c.kertas_bil} | {c.perunding} | {c.pemohon}")

st.divider()

if st.button("Generate & Download (ZIP)", type="primary", use_container_width=True):
    with st.spinner("Sedang jana dokumen (transfer data sahaja)..."):
        bil = info.get("bil") or 0
        year = info.get("year") or 0

        docs = []
        for c in cases:
            doc_bytes = generate_doc_for_case(template_bytes, info, c)
            ruj = build_rujukan_kami(bil, year, c.case_no) if (bil and year) else f"CASE_{c.case_no}"
            filename = f"{ruj.replace('/','-')}.docx"
            docs.append((filename, doc_bytes))

        zip_bytes = make_zip(docs)

    st.download_button(
        "Download ZIP",
        data=zip_bytes,
        file_name=f"pemakluman_keputusan_{info.get('bil_str','-').replace('/','-')}.zip",
        mime="application/zip",
        use_container_width=True,
    )
