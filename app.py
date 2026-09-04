import streamlit as st
from google import genai
from google.genai import types
from datetime import datetime
import re
import io
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# =========================================================
# PERANGKAT GURU - Streamlit
# =========================================================

st.set_page_config(
    page_title="PERANGKAT GURU | Portal Administrasi Kurikulum",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CSS
# =========================================================

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stMarkdown, p, span, label,
h1, h2, h3, h4, h5, h6 {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
    color: #0F172A !important;
}

.stApp { background-color: #F8FAFC !important; }

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 3.5rem !important;
    max-width: 1140px;
}

input, textarea, select {
    color: #0F172A !important;
    background-color: #FFFFFF !important;
    -webkit-text-fill-color: #0F172A !important;
}

.stTextInput label, .stSelectbox label,
.stMultiSelect label, .stRadio label {
    color: #1E293B !important;
    font-weight: 600 !important;
}

.brand-icon-box {
    display: flex;
    justify-content: center;
    align-items: center;
    margin-bottom: 12px;
}

.app-header {
    background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 55%, #2563EB 100%) !important;
    border-radius: 16px;
    padding: 28px 32px;
    margin-bottom: 24px;
    box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.12);
}

.app-header h1 {
    font-size: 26px;
    font-weight: 800;
    margin: 0 0 6px 0;
    color: #FFFFFF !important;
}

.app-header p {
    font-size: 14px;
    color: #CBD5E1 !important;
    margin: 0;
    line-height: 1.5;
}

div[data-testid="stVerticalBlock"] > div[style*="border"] {
    background: #FFFFFF !important;
    border-radius: 16px !important;
    border: 1px solid #E2E8F0 !important;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
    padding: 28px 26px !important;
    margin-bottom: 12px;
}

.stButton > button {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%) !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 12px 24px !important;
    box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
    width: 100% !important;
}

.stButton > button, .stButton > button * {
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 15px !important;
}

.stDownloadButton > button {
    font-weight: 700 !important;
    border-radius: 10px !important;
    padding: 12px 20px !important;
    width: 100% !important;
}

.paper-a4 {
    background: #FFFFFF !important;
    border: 1px solid #CBD5E1;
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08);
    padding: 44px 50px;
    margin: 16px auto;
    border-radius: 4px;
    max-width: 900px;
    color: #0F172A !important;
}

.paper-a4 table {
    width: 100% !important;
    border-collapse: collapse !important;
    margin: 14px 0 !important;
    font-size: 13px !important;
}

.paper-a4 th {
    background: #F1F5F9 !important;
    border: 1px solid #475569 !important;
    padding: 8px 10px !important;
    text-align: left;
    font-weight: 700;
}

.paper-a4 td {
    border: 1px solid #64748B !important;
    padding: 7px 10px !important;
}

.footer-box {
    text-align: center;
    padding: 24px 10px 10px;
    color: #64748B !important;
    font-size: 12px;
    border-top: 1px solid #E2E8F0;
    margin-top: 40px;
}

section[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid #E2E8F0;
}

@media (max-width: 768px) {
    .paper-a4 { padding: 20px 16px; }
    .app-header { padding: 22px 18px; }
}
</style>
""",
    unsafe_allow_html=True,
)

LOGO_URL = "https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f393.png"


def render_logo(width=90):
    st.markdown(
        f"""
        <div class="brand-icon-box">
            <img src="{LOGO_URL}" width="{width}" height="{width}"
                 alt="Logo Perangkat Guru"
                 style="filter:drop-shadow(0 4px 6px rgba(0,0,0,.1));">
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# SESSION STATE
# =========================================================

defaults = {
    "authenticated": False,
    "user_name": "",
    "user_api_key": "",
    "hasil_teks": "",
    "nama_file_base": "Perangkat_Kurikulum",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# EXPORT WORD
# =========================================================

def buat_file_docx(markdown_text: str) -> io.BytesIO:
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    lines = markdown_text.splitlines()
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("```"):
            i += 1
            continue

        if line.startswith("|") and line.endswith("|"):
            table_lines = []

            while (
                i < len(lines)
                and lines[i].strip().startswith("|")
                and lines[i].strip().endswith("|")
            ):
                raw = lines[i].strip()

                if not re.match(r"^\|[\s\-:|]+\|$", raw):
                    cells = [
                        c.strip()
                        for c in raw[1:-1].split("|")
                    ]
                    table_lines.append(cells)

                i += 1

            if table_lines:
                rows = len(table_lines)
                cols = max(len(row) for row in table_lines)

                table = doc.add_table(rows=rows, cols=cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = True

                for r_idx, row_data in enumerate(table_lines):
                    for c_idx in range(cols):
                        value = row_data[c_idx] if c_idx < len(row_data) else ""
                        value = (
                            value.replace("<br>", "\n")
                            .replace("<br/>", "\n")
                            .replace("**", "")
                        )

                        cell = table.cell(r_idx, c_idx)
                        cell.text = value

                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.name = "Calibri"
                                run.font.size = Pt(9.5)
                                if r_idx == 0:
                                    run.font.bold = True

                        if r_idx == 0:
                            shading = parse_xml(
                                r'<w:shd {} w:fill="E2E8F0"/>'.format(
                                    nsdecls("w")
                                )
                            )
                            cell._tc.get_or_add_tcPr().append(shading)

                borders = parse_xml(
                    f"""
                    <w:tblBorders {nsdecls("w")}>
                        <w:top w:val="single" w:sz="4" w:color="94A3B8"/>
                        <w:bottom w:val="single" w:sz="4" w:color="94A3B8"/>
                        <w:insideH w:val="single" w:sz="4" w:color="CBD5E1"/>
                        <w:insideV w:val="single" w:sz="4" w:color="CBD5E1"/>
                    </w:tblBorders>
                    """
                )
                table._tbl.tblPr.append(borders)
                doc.add_paragraph()

            continue

        if line.startswith("# "):
            p = doc.add_heading(line[2:], level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)

        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)

        elif line.startswith("---"):
            p = doc.add_paragraph()
            border = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                '<w:bottom w:val="single" w:sz="12" '
                'w:space="1" w:color="000000"/>'
                "</w:pBdr>"
            )
            p._p.get_or_add_pPr().append(border)

        elif line:
            clean = re.sub(r"<.*?>", "", line)
            clean = re.sub(r"\*\*(.*?)\*\*", r"\1", clean)

            p = doc.add_paragraph(clean)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)

            for run in p.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(11)

        i += 1

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


# =========================================================
# MARKDOWN -> HTML PREVIEW
# =========================================================

def markdown_to_html(text: str) -> str:
    lines = text.splitlines()
    html = []
    in_table = False
    table_rows = []

    def flush_table():
        nonlocal table_rows, in_table
        if not table_rows:
            return

        html.append("<table>")
        for r_idx, row in enumerate(table_rows):
            tag = "th" if r_idx == 0 else "td"
            html.append("<tr>")
            for cell in row:
                cell_html = (
                    cell.replace("<br>", "<br>")
                    .replace("<br/>", "<br>")
                )
                html.append(f"<{tag}>{cell_html}</{tag}>")
            html.append("</tr>")
        html.append("</table>")

        table_rows = []
        in_table = False

    for raw in lines:
        line = raw.strip()

        if line.startswith("```"):
            continue

        if line.startswith("|") and line.endswith("|"):
            if re.match(r"^\|[\s\-:|]+\|$", line):
                continue

            if not in_table:
                in_table = True

            cells = [c.strip() for c in line[1:-1].split("|")]
            table_rows.append(cells)
            continue

        if in_table:
            flush_table()

        if not line:
            html.append("<p>&nbsp;</p>")
            continue

        escaped = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        escaped = re.sub(r"\*\*(.*?)\*\*", r"<strong>\1</strong>", escaped)

        if escaped.startswith("### "):
            html.append(f"<h3>{escaped[4:]}</h3>")
        elif escaped.startswith("## "):
            html.append(f"<h2>{escaped[3:]}</h2>")
        elif escaped.startswith("# "):
            html.append(f"<h1>{escaped[2:]}</h1>")
        elif escaped.startswith("---"):
            html.append("<hr>")
        else:
            html.append(f"<p>{escaped}</p>")

    if in_table:
        flush_table()

    return "\n".join(html)


# =========================================================
# PROMPT
# =========================================================

def buat_instruksi_prompt(data: dict) -> str:
    p3 = (
        ", ".join(data["profil_pancasila"])
        if data["profil_pancasila"]
        else "Sesuai karakteristik materi"
    )

    return f"""
Bertindaklah sebagai ahli penyusunan perangkat pembelajaran Indonesia yang
memahami Kurikulum Merdeka dan regulasi pendidikan yang berlaku.

Tugas: susun dokumen "{data['jenis_perangkat']}" yang lengkap, detail,
operasional, sistematis, dan siap disalin ke dokumen resmi sekolah.

IDENTITAS SATUAN PENDIDIKAN
- Dinas Pendidikan: {data['dinas']}
- Satuan Pendidikan: {data['sekolah']}
- Alamat dan Kontak: {data['alamat']}
- Kota/Kabupaten: {data['kota_sekolah']}

IDENTITAS PENDIDIK DAN PEJABAT
- {data['jabatan_guru']}: {data['guru_nama']} (NIP: {data['guru_nip']})
- Kepala Sekolah: {data['ks_nama']} (NIP: {data['ks_nip']})
- Pengawas Pembina: {data['pengawas_nama']} (NIP: {data['pengawas_nip']})

PARAMETER PEMBELAJARAN
- Mata Pelajaran: {data['mapel']}
- Fase/Kelas: {data['fase_kelas']}
- Semester: {data['semester']}
- Alokasi Waktu: {data['alokasi_waktu']}
- Materi Pokok: {data['materi_pokok']}
- Dimensi Profil Pelajar Pancasila: {p3}
- Tanggal: {data['tanggal_hari_ini']}

ATURAN OUTPUT
1. Langsung mulai dengan KOP SURAT resmi lembaga pendidikan.
2. Gunakan judul dokumen yang jelas menggunakan heading Markdown (`#` atau `##`).
3. Selalu gunakan tabel Markdown yang rapi untuk identitas dokumen, rincian materi, atau kisi-kisi. Pastikan jumlah kolom pembuka dan penutup tabel konsisten (`| kolom 1 | kolom 2 |`).
4. **DILARANG KERAS** menggunakan simbol "...", "[lanjutkan]", "[sesuaikan]", "dst.", atau "dan seterusnya". Semua isi paragraf dan poin harus ditulis secara **lengkap, penuh, dan tuntas sampai selesai**.
5. Jangan mengarang nomor regulasi atau undang-undang spesifik jika tidak yakin.
6. Gunakan bahasa Indonesia formal, baku, edukatif, dan mudah diedit oleh guru.
7. **PENTING:** Jangan pernah menyertakan blok kode Markdown berpagar (seperti ```html atau
