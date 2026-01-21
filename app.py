import io
import re
import zipfile
from dataclasses import dataclass
from typing import List, Optional, Tuple

import streamlit as st
from docx import Document
from docx.shared import Cm


TEMPLATE_PATH = "templates/template dokumen panggilan.docx"


# =========================
# Helpers: DOCX insert table
# =========================
def _insert_table_after_paragraph(paragraph, rows: int, cols: int):
    """
    Create a table and insert it right after a paragraph.
    """
    doc = paragraph.part.document
    table = doc.add_table(rows=rows, cols=cols)
    # move table XML right after paragraph XML
    paragraph._p.addnext(table._tbl)
    return table


def _safe_get_template_bytes() -> bytes:
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()


# =========================
# Parsing Agenda
# =========================
@dataclass
class AgendaMeta:
    bil_no: int               # e.g. 1 from BIL. 01/2026
    bil_year: int             # e.g. 2026
    mesyuarat_title: str      # e.g. "Mesyuarat Jawatankuasa OSC MBSP"
    tarikh_mesyuarat: str     # keep as-is from agenda
    masa_mesyuarat: str       # keep as-is from agenda
    tempat_mesyuarat: str     # keep as-is from agenda


@dataclass
class CaseItem:
    item_no: int              # e.g. 36 from JKOSC/.../36/2026
    item_year: int            # e.g. 2026
    jenis_permohonan: str     # "Kebenaran Merancang" or "Bangunan"
    pemohon: str
    perunding: str
    nama_permohonan: str
    id_permohonan: str
    kertas_mesyuarat: str


def _norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def parse_agenda(doc_bytes: bytes) -> Tuple[AgendaMeta, List[CaseItem]]:
    doc = Document(io.BytesIO(doc_bytes))
    paras = [_norm_spaces(p.text) for p in doc.paragraphs if _norm_spaces(p.text)]

    # Find BIL. 01/2026
    bil_no, bil_year = None, None
    bil_re = re.compile(r"\bBIL\.\s*(\d{1,2})\s*/\s*(\d{4})\b", re.IGNORECASE)
    for t in paras:
        m = bil_re.search(t)
        if m:
            bil_no = int(m.group(1))
            bil_year = int(m.group(2))
            break
    if bil_no is None or bil_year is None:
        raise ValueError("Tak jumpa format 'BIL. 01/2026' dalam agenda.")

    # Meeting title line containing "Mesyuarat" + "Bil."
    mesyuarat_title = ""
    for t in paras:
        if "Mesyuarat" in t and bil_re.search(t):
            # remove trailing "Bil. xx/yyyy"
            mesyuarat_title = bil_re.sub("", t).strip()
            mesyuarat_title = mesyuarat_title.replace("  ", " ").strip()
            break
    if not mesyuarat_title:
        # fallback: use a generic title
        mesyuarat_title = "Mesyuarat Jawatankuasa OSC MBSP"

    # Tarikh, Masa, Tempat
    tarikh_m = ""
    masa_m = ""
    tempat_m = ""

    for i, t in enumerate(paras):
        if t.lower().startswith("tarikh"):
            tarikh_m = t.split(":", 1)[-1].strip() if ":" in t else t
        if t.lower().startswith("masa"):
            masa_m = t.split(":", 1)[-1].strip() if ":" in t else t
        if t.lower().startswith("tempat"):
            # ambil selepas "Tempat:" dan sambung beberapa baris seterusnya
            first = t.split(":", 1)[-1].strip() if ":" in t else ""
            lines = []
            if first:
                lines.append(first)
            # ambil next lines sampai jumpa sesuatu yang nampak macam section/numbering
            for j in range(i + 1, min(i + 8, len(paras))):
                nxt = paras[j]
                if re.match(r"^\d+\.", nxt) or re.match(r"^[A-Z]\.", nxt) or nxt.upper().startswith("PENGERUSI"):
                    break
                # stop kalau ini header besar agenda
                if "PERMOHONAN" in nxt.upper() and nxt.upper().startswith(("A.", "B.", "C.", "D.", "E.")):
                    break
                lines.append(nxt)
            tempat_m = _norm_spaces(" ".join(lines))
            break

    if not tarikh_m:
        tarikh_m = f"{bil_year}"  # fallback minimal
    if not masa_m:
        masa_m = "-"
    if not tempat_m:
        tempat_m = "-"

    meta = AgendaMeta(
        bil_no=bil_no,
        bil_year=bil_year,
        mesyuarat_title=mesyuarat_title,
        tarikh_mesyuarat=tarikh_m,
        masa_mesyuarat=masa_m,
        tempat_mesyuarat=tempat_m,
    )

    # Extract cases using "No. Rujukan OSC"
    cases: List[CaseItem] = []
    for idx, t in enumerate(paras):
        if "No. Rujukan OSC" not in t:
            continue

        # ID permohonan
        id_perm = t.split(":", 1)[-1].strip() if ":" in t else ""

        # Find pemohon/perunding by scanning backwards
        pemohon = ""
        perunding = ""
        for back in range(idx, max(idx - 20, -1), -1):
            if paras[back].startswith("Pemohon"):
                # usually next line is value
                if back + 1 < len(paras):
                    pemohon = paras[back + 1]
            if paras[back].startswith("Perunding"):
                if back + 1 < len(paras):
                    perunding = paras[back + 1]
            if pemohon and perunding:
                break

        # Find KERTAS MESYUARAT forward
        kertas = ""
        kertas_idx = None
        for fwd in range(idx, min(idx + 15, len(paras))):
            if "KERTAS MESYUARAT" in paras[fwd]:
                kertas = paras[fwd]
                kertas_idx = fwd
                break

        if not kertas:
            continue

        # item_no / year from kertas ".../36/2026"
        m = re.search(r"/(\d{1,3})/(\d{4})\b", kertas)
        if not m:
            continue
        item_no = int(m.group(1))
        item_year = int(m.group(2))

        # jenis_permohonan from code TM/PB in kertas
        jenis = ""
        if "/TM/" in kertas:
            jenis = "Kebenaran Merancang"
        elif "/PB/" in kertas:
            jenis = "Bangunan"

        if jenis not in ("Kebenaran Merancang", "Bangunan"):
            continue  # user wants only these 2

        # nama_permohonan: start after kertas line, collect until next case block markers
        nama_lines = []
        if kertas_idx is not None:
            start = kertas_idx + 1
            for j in range(start, min(start + 30, len(paras))):
                cur = paras[j]
                stop = (
                    cur.startswith("Pemohon")
                    or cur.startswith("Perunding")
                    or cur.startswith("Lokasi")
                    or cur.startswith("Koordinat")
                    or cur.startswith("No. Rujukan")
                    or "KERTAS MESYUARAT" in cur
                    or "No. Rujukan OSC" in cur
                    or re.match(r"^[A-Z]\.\s", cur)  # section headings A./B./...
                )
                if stop:
                    break
                # usually begins with "Permohonan ..."
                if cur:
                    nama_lines.append(cur)

        nama_perm = "\n".join(nama_lines).strip()
        if not nama_perm:
            nama_perm = "-"  # fallback

        cases.append(
            CaseItem(
                item_no=item_no,
                item_year=item_year,
                jenis_permohonan=jenis,
                pemohon=pemohon or "-",
                perunding=perunding or "-",
                nama_permohonan=nama_perm,
                id_permohonan=id_perm or "-",
                kertas_mesyuarat=kertas,
            )
        )

    # de-duplicate by item_no (just in case)
    uniq = {}
    for c in cases:
        uniq[c.item_no] = c
    cases = [uniq[k] for k in sorted(uniq.keys())]

    return meta, cases


# =========================
# Generate Letter (DOCX)
# =========================
def build_rujukan_kami(meta: AgendaMeta, case: CaseItem) -> str:
    # format: (1)MBSP/15/1551/(36)2026
    return f"({meta.bil_no})MBSP/15/1551/({case.item_no}){meta.bil_year}"


def replace_paragraph_text(p, new_text: str):
    # wipe runs then set one run
    for r in p.runs:
        r.text = ""
    if p.runs:
        p.runs[0].text = new_text
    else:
        p.add_run(new_text)


def generate_docx(template_bytes: bytes, meta: AgendaMeta, case: CaseItem, tarikh_surat: str) -> bytes:
    doc = Document(io.BytesIO(template_bytes))

    # 1) Replace Rujukan Kami line
    rujukan = build_rujukan_kami(meta, case)
    for p in doc.paragraphs:
        if "Rujukan Kami" in p.text:
            replace_paragraph_text(p, f"Rujukan Kami : {rujukan}")
            break

    # 2) Replace Tarikh line
    for p in doc.paragraphs:
        if p.text.strip().startswith("Tarikh"):
            replace_paragraph_text(p, f"Tarikh\t: {tarikh_surat}")
            break

    # 3) Replace the main “Dimaklumkan bahawa … berikut:” paragraph
    for p in doc.paragraphs:
        if "Dimaklumkan bahawa" in p.text and "berikut" in p.text:
            sentence = (
                f"2.\tPerkara di atas adalah dirujuk. Dimaklumkan bahawa {meta.mesyuarat_title} "
                f"Bil. {meta.bil_no:02d}/{meta.bil_year} yang diadakan pada {meta.tarikh_mesyuarat} "
                f"jam {meta.masa_mesyuarat} di {meta.tempat_mesyuarat}, telah dikemukakan oleh pihak "
                f"tuan/ puan seperti mana berikut:"
            )
            replace_paragraph_text(p, sentence)
            break

    # 4) Insert details table right after the paragraph that contains "berikut:"
    target_p = None
    for p in doc.paragraphs:
        if "seperti mana berikut" in p.text or p.text.strip().endswith("berikut:"):
            target_p = p
            break

    if target_p is None:
        # fallback: insert near end before first table
        target_p = doc.paragraphs[-1]

    table = _insert_table_after_paragraph(target_p, rows=5, cols=3)
    table.style = "Table Grid"

    # set widths (approx; Word will still adjust)
    widths = [Cm(4.2), Cm(0.6), Cm(12.0)]
    labels = ["Kepada (PSP)", "Pemilik Projek", "Jenis Permohonan", "Nama Permohonan", "ID Permohonan"]
    values = [case.perunding, case.pemohon, case.jenis_permohonan, case.nama_permohonan, case.id_permohonan]

    for r in range(5):
        row = table.rows[r]
        for c in range(3):
            row.cells[c].width = widths[c]

        row.cells[0].text = labels[r]
        row.cells[1].text = ":"
        row.cells[2].text = values[r]

    # Done
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Pemakluman Keputusan JKOSC", layout="wide")
st.title("Pemakluman Keputusan JKOSC")
st.caption("Upload Agenda (Word) → sistem extract KM & Bangunan → generate dokumen panggilan (Word) untuk download.")

agenda_file = st.file_uploader("Upload fail Agenda (DOCX)", type=["docx"])

tarikh_surat = st.text_input(
    "Tarikh Surat (akan masuk dalam template)",
    value="",
    help="Contoh: 23 Januari 2026 (atau ikut format yang Unit guna).",
)

if agenda_file:
    try:
        meta, cases = parse_agenda(agenda_file.read())

        st.success(
            f"Jumpa BIL. {meta.bil_no:02d}/{meta.bil_year} | "
            f"KM/Bangunan cases: {len(cases)}"
        )

        st.write("**Info Mesyuarat (auto dari agenda):**")
        st.write(f"- Tajuk: {meta.mesyuarat_title}")
        st.write(f"- Tarikh: {meta.tarikh_mesyuarat}")
        st.write(f"- Masa: {meta.masa_mesyuarat}")
        st.write(f"- Tempat: {meta.tempat_mesyuarat}")

        # selection
        st.divider()
        st.write("**Pilih kes untuk generate**")
        selected = []
        for c in cases:
            label = f"[{c.item_no}] {c.jenis_permohonan} | {c.id_permohonan}"
            if st.checkbox(label, value=True, key=f"case_{c.item_no}"):
                selected.append(c)

        if st.button("Generate & Download (ZIP)", type="primary"):
            if not tarikh_surat.strip():
                st.error("Tarikh Surat kosong. Isi dulu.")
            elif not selected:
                st.error("Takde kes dipilih.")
            else:
                template_bytes = _safe_get_template_bytes()
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                    for c in selected:
                        docx_bytes = generate_docx(template_bytes, meta, c, tarikh_surat.strip())
                        # filename: (36) <ringkas>.docx
                        safe_id = re.sub(r"[^\w\-\(\)\.\s]", "", c.id_permohonan).strip()
                        fname = f"({c.item_no}) {safe_id}.docx"
                        z.writestr(fname, docx_bytes)

                st.download_button(
                    "Download ZIP",
                    data=zip_buf.getvalue(),
                    file_name=f"pemakluman_keputusan_Bil_{meta.bil_no:02d}_{meta.bil_year}.zip",
                    mime="application/zip",
                )

    except Exception as e:
        st.error(f"Error baca agenda: {e}")

else:
    st.info("Upload Agenda DOCX dulu.")
