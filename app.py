import re
import io
import zipfile
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

import streamlit as st
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph


# =========================
# UI CONFIG
# =========================
st.set_page_config(
    page_title="Pemakluman Keputusan JKOSC",
    layout="wide",
)

st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → sistem extract KM/Bangunan → generate dokumen panggilan (Word) untuk download.")


# =========================
# HELPERS: DATE / TEXT
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
    s = s.replace("\u00a0", " ")  # non-breaking space
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def format_tarikh_malay(raw: str) -> str:
    """
    Input expected like: '12 JANUARI 2026' (case-insensitive).
    Output: '12 Januari 2026'
    """
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
    """
    Cari: Bil. 01/2026 atau BIL. 1/2026
    """
    t = text.upper()
    m = re.search(r"\bBIL\.?\s*0*(\d{1,2})\s*/\s*(\d{4})\b", t)
    if not m:
        return None
    bil = int(m.group(1))
    year = int(m.group(2))
    return bil, year

def extract_tarikh_mesyuarat(lines: List[str]) -> Optional[str]:
    """
    Cari baris yang ada TARIKH : 12 JANUARI 2026 (atau variasi)
    """
    for ln in lines:
        u = ln.upper()
        if "TARIKH" in u:
            m = re.search(r"TARIKH\s*[:\-]\s*(\d{1,2}\s+[A-Z]+\s+\d{4})", u)
            if m:
                return format_tarikh_malay(m.group(1))
    # fallback: cari date pattern dekat awal dokumen
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
    case_no: str                 # contoh "01" atau "36"
    kertas_bil: str              # contoh "OSC/PKM/01/2026"
    jenis: str                   # "KM" / "BANGUNAN"
    jenis_full: str              # "PERMOHONAN KEBENARAN MERANCANG" / ...
    perunding: str
    pemohon: str
    nama_permohonan: str
    id_permohonan: str

def read_docx_lines(file_bytes: bytes) -> List[str]:
    doc = Document(io.BytesIO(file_bytes))
    lines = []
    for p in doc.paragraphs:
        t = p.text
        if t:
            t2 = normalize_spaces(t)
            if t2:
                lines.append(t2)
    return lines

def parse_case_block(block_lines: List[str], bil_year: Tuple[int, int]) -> Optional[CaseItem]:
    """
    block_lines bermula dengan 'KERTAS MESYUARAT BIL....'
    """
    text_all = "\n".join(block_lines)

    # kertas bil
    m_kertas = re.search(r"KERTAS\s+MESYUARAT\s+BIL\.?\s*[:\-]?\s*([A-Z0-9/]+)", text_all.upper())
    if not m_kertas:
        return None
    kertas = m_kertas.group(1).strip()

    # decide jenis
    jenis = None
    jenis_full = None

    # KM: PKM or 'KEBENARAN MERANCANG'
    if "PKM" in kertas or re.search(r"\bKEBENARAN\s+MERANCANG\b", text_all.upper()):
        jenis = "KM"
        jenis_full = "PERMOHONAN KEBENARAN MERANCANG"

    # Bangunan: PB or 'BANGUNAN' + 'PELAN'
    if jenis is None and ("PB" in kertas or ("BANGUNAN" in text_all.upper() and "PELAN" in text_all.upper())):
        jenis = "BANGUNAN"
        jenis_full = "PERMOHONAN PELAN BANGUNAN"

    # Kalau bukan KM/Bangunan, skip
    if jenis is None:
        return None

    # case_no dari kertas: OSC/PKM/01/2026 -> ambik "01"
    m_no = re.search(r"/(\d{1,3})/(\d{4})$", kertas)
    case_no = m_no.group(1).zfill(2) if m_no else "00"

    # fields
    def pick(prefix: str) -> str:
        # cari "Prefix : value"
        for ln in block_lines:
            u = ln.upper()
            if u.startswith(prefix.upper()):
                # split colon
                parts = re.split(r"\s*[:\-]\s*", ln, maxsplit=1)
                if len(parts) == 2:
                    return parts[1].strip()
        return ""

    perunding = pick("Perunding")
    pemohon = pick("Pemohon")
    id_permohonan = pick("No Fail OSC") or pick("No. Fail OSC") or pick("No Fail") or pick("No. Fail")

    # nama permohonan: ambik baris bermula "Permohonan ..."
    nama_permohonan = ""
    for ln in block_lines:
        if ln.upper().startswith("PERMOHONAN "):
            nama_permohonan = ln.strip()
            break

    # fallback kalau tak jumpa
    if not nama_permohonan:
        nama_permohonan = f"PERMOHONAN ({jenis_full})"

    # minimal sanity: mesti ada perunding/pemohon/id
    # (kalau agenda ada variasi, jangan terus fail — tapi at least id permohonan kena ada)
    if not id_permohonan:
        # cuba cari pattern MBSP/... dalam block
        mm = re.search(r"\bMBSP/[A-Z0-9/().\- ]+\b", text_all)
        if mm:
            id_permohonan = normalize_spaces(mm.group(0))

    return CaseItem(
        case_no=case_no,
        kertas_bil=kertas,
        jenis=jenis,
        jenis_full=jenis_full,
        perunding=perunding or "-",
        pemohon=pemohon or "-",
        nama_permohonan=nama_permohonan,
        id_permohonan=id_permohonan or "-",
    )

def parse_agenda(file_bytes: bytes) -> Tuple[Dict, List[CaseItem]]:
    lines = read_docx_lines(file_bytes)

    bil_year = None
    for ln in lines[:60]:
        by = extract_bil_year(ln)
        if by:
            bil_year = by
            break
    if not bil_year:
        # fallback: try full scan
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
        "tajuk": None,
        "masa": None,
        "tempat": None,
    }

    # try pull Tajuk/Masa/Tempat basic
    for ln in lines[:120]:
        u = ln.upper()
        if u.startswith("TAJUK"):
            info["tajuk"] = ln
        elif u.startswith("MASA"):
            info["masa"] = ln
        elif u.startswith("TEMPAT"):
            info["tempat"] = ln

    # split blocks by "KERTAS MESYUARAT BIL"
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
    if bil_year:
        for b in blocks:
            item = parse_case_block(b, bil_year)
            if item:
                cases.append(item)

    # sort ikut nombor kes
    def keyfn(x: CaseItem):
        try:
            return int(x.case_no)
        except:
            return 9999
    cases.sort(key=keyfn)

    return info, cases


# =========================
# WORD GENERATION
# =========================
def replace_paragraph_text_keep_style(paragraph: Paragraph, new_text: str) -> None:
    """
    Replace full paragraph text while keeping paragraph style (not run styles).
    We'll clear runs then add one run with new_text.
    """
    # clear existing runs
    for r in paragraph.runs[::-1]:
        r._element.getparent().remove(r._element)
    paragraph.add_run(new_text)

def set_field_in_template(doc: Document, label_startswith: str, value: str) -> bool:
    """
    Find paragraph that starts with label_startswith (case-insensitive), then set it to:
    \\tLabel\\t:\\tVALUE  (maintain tab layout)
    """
    target = label_startswith.strip().upper()
    for p in doc.paragraphs:
        t = normalize_spaces(p.text)
        if not t:
            continue
        if t.upper().startswith(target):
            # standardize to tab format like template uses
            new_text = f"\t{label_startswith}\t:\t{value}"
            replace_paragraph_text_keep_style(p, new_text)
            return True
    return False

def set_singleline_prefix(doc: Document, prefix: str, value: str) -> bool:
    """
    Example: "Rujukan Kami:" -> "Rujukan Kami: (1)MBSP/..."
    """
    pref_u = prefix.upper()
    for p in doc.paragraphs:
        t = normalize_spaces(p.text)
        if not t:
            continue
        if t.upper().startswith(pref_u):
            new_text = f"{prefix} {value}".rstrip()
            replace_paragraph_text_keep_style(p, new_text)
            return True
    return False

def build_rujukan_kami(bil: int, year: int, case_no: str) -> str:
    return f"({bil})MBSP/15/1551/({case_no}){year}"

def generate_doc_for_case(
    template_bytes: bytes,
    agenda_info: Dict,
    case: CaseItem,
    keputusan: str,
) -> bytes:
    """
    Return .docx bytes
    """
    doc = Document(io.BytesIO(template_bytes))

    bil = agenda_info.get("bil")
    year = agenda_info.get("year")
    tarikh = agenda_info.get("tarikh_mesyuarat", "-")

    # --- Header fields
    if bil and year:
        rujukan_kami = build_rujukan_kami(bil, year, case.case_no)
        set_singleline_prefix(doc, "Rujukan Kami:", rujukan_kami)

    # Tarikh letter = tarikh mesyuarat
    set_singleline_prefix(doc, "Tarikh:", tarikh)

    # --- Main info fields (tab style in template)
    set_field_in_template(doc, "Perunding", case.perunding)
    set_field_in_template(doc, "Pemohon", case.pemohon)
    set_field_in_template(doc, "Jenis Permohonan", case.jenis_full)
    set_field_in_template(doc, "Nama Permohonan", case.nama_permohonan)
    set_field_in_template(doc, "ID Permohonan", case.id_permohonan)

    # --- Optional: inject keputusan word into a paragraph if exists
    # We'll look for a paragraph containing 'keputusan' and 'permohonan tersebut'
    # then replace LULUS/TOLAK/TANGGUH if found.
    keputusan_u = keputusan.upper().strip()
    for p in doc.paragraphs:
        t = normalize_spaces(p.text)
        if not t:
            continue
        tu = t.upper()
        if "BERSUTUJU" in tu and "KEPUTUSAN" in tu:
            # rewrite to a standard line (keep paragraph style)
            bil_str = agenda_info.get("bil_str", "-")
            new_t = (
                f"3.\tJawatankuasa OSC telah bersetuju memberi keputusan {keputusan_u} "
                f"bagi permohonan tersebut di atas dan butiran permohonan adalah seperti berikut:"
            )
            replace_paragraph_text_keep_style(p, new_t)
            break

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
# STREAMLIT APP
# =========================
uploaded = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

template_path = "templates/template dokumen panggilan.docx"

colA, colB = st.columns([1, 1])

if uploaded is None:
    st.info("Upload Agenda dulu. Sistem akan auto baca Bil & Tarikh mesyuarat, then senaraikan kes KM/Bangunan.")
    st.stop()

agenda_bytes = uploaded.read()

# load template from repo
try:
    with open(template_path, "rb") as f:
        template_bytes = f.read()
except Exception:
    st.error(
        "Template tak jumpa dalam repo.\n\n"
        "Pastikan fail ada di: `templates/template dokumen panggilan.docx` (nama tepat)."
    )
    st.stop()

# parse
info, cases = parse_agenda(agenda_bytes)

bil_str = info.get("bil_str", "-")
tarikh_mesyuarat = info.get("tarikh_mesyuarat", "-")

st.success(f"Jumpa {bil_str} | Tarikh Mesyuarat: {tarikh_mesyuarat} | KM/Bangunan cases: {len(cases)}")

with colA:
    st.subheader("Info Mesyuarat (auto dari agenda)")
    st.write(f"**Bil:** {bil_str}")
    st.write(f"**Tarikh Mesyuarat (auto jadi Tarikh Surat):** {tarikh_mesyuarat}")
    if info.get("tajuk"):
        st.write("**Tajuk:**")
        st.write(info["tajuk"])
    if info.get("masa"):
        st.write("**Masa:**")
        st.write(info["masa"])
    if info.get("tempat"):
        st.write("**Tempat:**")
        st.write(info["tempat"])

with colB:
    st.subheader("Pilih kes untuk generate")

    if len(cases) == 0:
        st.error(
            "Sistem tak jumpa kes KM/Bangunan dalam agenda.\n\n"
            "Kalau agenda tu memang ada, biasanya format baris 'KERTAS MESYUARAT BIL.' atau label 'No Fail OSC/Perunding/Pemohon' berbeza.\n"
            "Bagitau—aku adjust parser ikut format sebenar agenda kau."
        )
        st.stop()

    # selection UI
    keputusan_default = "LULUS"
    keputusan_options = ["LULUS", "TOLAK", "TANGGUH"]

    selected_idx = []
    keputusan_by_idx = {}

    st.caption("Tick kes yang nak dijana. Keputusan default = LULUS (boleh tukar).")

    for i, c in enumerate(cases):
        left, mid, right = st.columns([0.08, 0.62, 0.30])
        with left:
            picked = st.checkbox(f"#{c.case_no}", key=f"pick_{i}")
        with mid:
            st.markdown(
                f"**{c.kertas_bil}**  \n"
                f"{c.nama_permohonan}  \n"
                f"Perunding: {c.perunding} | Pemohon: {c.pemohon}"
            )
        with right:
            keputusan = st.selectbox(
                "Keputusan",
                keputusan_options,
                index=keputusan_options.index(keputusan_default),
                key=f"keputusan_{i}",
                label_visibility="visible",
            )

        if picked:
            selected_idx.append(i)
            keputusan_by_idx[i] = keputusan

    st.divider()

    if st.button("Generate & Download (ZIP)", type="primary", use_container_width=True):
        if not selected_idx:
            st.error("Takde kes dipilih.")
            st.stop()

        with st.spinner("Sedang jana dokumen..."):
            docs = []
            bil = info.get("bil") or 0
            year = info.get("year") or 0

            for i in selected_idx:
                c = cases[i]
                kep = keputusan_by_idx.get(i, "LULUS")
                doc_bytes = generate_doc_for_case(template_bytes, info, c, kep)

                # filename standard
                ruj = build_rujukan_kami(bil, year, c.case_no) if (bil and year) else f"CASE_{c.case_no}"
                safe_ruj = ruj.replace("/", "-")
                filename = f"{safe_ruj}.docx"
                docs.append((filename, doc_bytes))

            zip_bytes = make_zip(docs)

        st.download_button(
            "Download ZIP",
            data=zip_bytes,
            file_name=f"pemakluman_keputusan_{bil_str.replace('/','-')}.zip",
            mime="application/zip",
            use_container_width=True,
        )
