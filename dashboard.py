import streamlit as st
from google import genai
from datetime import datetime
import re
import io
import time

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


# =========================================================
# SIAP AJAR 22 - DASHBOARD
# =========================================================

# Login.py adalah entrypoint utama.
# Jadi dashboard tidak menggunakan st.set_page_config() lagi.

if not st.session_state.get("authenticated", False):
    st.stop()


# =========================================================
# SESSION STATE
# =========================================================

defaults = {
    "user_name": "",
    "user_api_key": "",
    "hasil_teks": "",
    "nama_file_base": "Perangkat_Kurikulum",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# LOGO
# =========================================================

LOGO_URL = (
    "https://raw.githubusercontent.com/"
    "gmandar060-cell/Perangkat-Guru/main/logo.png"
)


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>

    @import url(
        'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:'
        'wght@400;500;600;700;800&display=swap'
    );

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }

    .stApp {
        background: #F8FAFC !important;
    }

    .block-container {
        max-width: 1180px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    input,
    textarea {
        color: #0F172A !important;
        background: #FFFFFF !important;
        -webkit-text-fill-color: #0F172A !important;
    }

    input::placeholder,
    textarea::placeholder {
        color: #94A3B8 !important;
        opacity: 1 !important;
    }

    label {
        color: #1E293B !important;
        font-weight: 600 !important;
    }

    .brand-logo {
        text-align: center;
        margin-bottom: 5px;
    }

    .brand-logo img {
        width: 70px;
        height: 70px;
        object-fit: contain;
    }

    .dashboard-header {
        background:
            linear-gradient(
                135deg,
                #0F172A 0%,
                #1E3A8A 55%,
                #2563EB 100%
            );

        border-radius: 18px;
        padding: 28px 32px;
        margin-bottom: 22px;

        box-shadow:
            0 10px 25px rgba(15, 23, 42, 0.12);
    }

    .dashboard-header h1 {
        color: white !important;
        margin: 0 0 6px 0;
        font-size: 27px;
        font-weight: 800;
    }

    .dashboard-header p {
        color: #CBD5E1 !important;
        margin: 0;
        font-size: 14px;
        line-height: 1.6;
    }

    .info-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 15px;
    }

    .info-number {
        font-size: 25px;
        font-weight: 800;
        color: #1E3A8A !important;
    }

    .info-label {
        color: #64748B !important;
        font-size: 12px;
    }

    .section-title {
        font-size: 18px;
        font-weight: 800;
        color: #0F172A !important;
        margin-bottom: 8px;
    }

    .paper-a4 {
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1;
        border-radius: 5px;
        padding: 45px 50px;
        margin: 18px auto;
        max-width: 900px;
        box-shadow: 0 12px 30px rgba(0,0,0,.08);
        color: #0F172A !important;
    }

    .paper-a4 h1,
    .paper-a4 h2,
    .paper-a4 h3,
    .paper-a4 p,
    .paper-a4 td,
    .paper-a4 th {
        color: #0F172A !important;
    }

    .paper-a4 table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
        font-size: 13px;
    }

    .paper-a4 th {
        background: #E2E8F0 !important;
        border: 1px solid #64748B;
        padding: 8px;
    }

    .paper-a4 td {
        border: 1px solid #94A3B8;
        padding: 8px;
    }

    .footer-box {
        text-align: center;
        color: #64748B !important;
        font-size: 12px;
        border-top: 1px solid #E2E8F0;
        padding-top: 20px;
        margin-top: 35px;
    }

    .stButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        min-height: 44px;
    }

    .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        min-height: 44px;
    }

    section[data-testid="stSidebar"] {
        background: #FFFFFF !important;
        border-right: 1px solid #E2E8F0;
    }

    @media (max-width: 768px) {

        .dashboard-header {
            padding: 22px 18px;
        }

        .dashboard-header h1 {
            font-size: 22px;
        }

        .paper-a4 {
            padding: 22px 16px;
        }

    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNGSI MEMBUAT DOCX
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

        # -------------------------
        # TABEL
        # -------------------------

        if line.startswith("|") and line.endswith("|"):

            table_lines = []

            while (
                i < len(lines)
                and lines[i].strip().startswith("|")
                and lines[i].strip().endswith("|")
            ):

                raw = lines[i].strip()

                if not re.match(
                    r"^\|[\s\-:|]+\|$",
                    raw
                ):

                    cells = [
                        c.strip()
                        for c in raw[1:-1].split("|")
                    ]

                    table_lines.append(cells)

                i += 1

            if table_lines:

                rows = len(table_lines)
                cols = max(
                    len(row)
                    for row in table_lines
                )

                table = doc.add_table(
                    rows=rows,
                    cols=cols
                )

                table.alignment = (
                    WD_TABLE_ALIGNMENT.CENTER
                )

                for r_idx, row_data in enumerate(
                    table_lines
                ):

                    for c_idx in range(cols):

                        value = (
                            row_data[c_idx]
                            if c_idx < len(row_data)
                            else ""
                        )

                        value = (
                            value
                            .replace("<br>", "\n")
                            .replace("<br/>", "\n")
                            .replace("**", "")
                        )

                        cell = table.cell(
                            r_idx,
                            c_idx
                        )

                        cell.text = value

                        for paragraph in cell.paragraphs:

                            for run in paragraph.runs:

                                run.font.name = "Calibri"
                                run.font.size = Pt(9.5)

                                if r_idx == 0:
                                    run.font.bold = True

                doc.add_paragraph()

            continue

        # -------------------------
        # HEADING
        # -------------------------

        if line.startswith("# "):

            p = doc.add_heading(
                line[2:],
                level=1
            )

            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

        elif line.startswith("## "):

            doc.add_heading(
                line[3:],
                level=2
            )

        elif line.startswith("### "):

            doc.add_heading(
                line[4:],
                level=3
            )

        elif line.startswith("---"):

            doc.add_paragraph()

        elif line:

            clean = re.sub(
                r"<.*?>",
                "",
                line
            )

            clean = re.sub(
                r"\*\*(.*?)\*\*",
                r"\1",
                clean
            )

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
# MARKDOWN → HTML
# =========================================================

def markdown_to_html(text: str) -> str:

    lines = text.splitlines()

    html = []

    in_table = False
    table_rows = []

    def flush_table():

        nonlocal table_rows
        nonlocal in_table

        if not table_rows:
            return

        html.append("<table>")

        for r_idx, row in enumerate(table_rows):

            tag = "th" if r_idx == 0 else "td"

            html.append("<tr>")

            for cell in row:

                cell_html = (
                    cell
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )

                cell_html = re.sub(
                    r"\*\*(.*?)\*\*",
                    r"<strong>\1</strong>",
                    cell_html
                )

                html.append(
                    f"<{tag}>{cell_html}</{tag}>"
                )

            html.append("</tr>")

        html.append("</table>")

        table_rows = []
        in_table = False

    for raw in lines:

        line = raw.strip()

        if line.startswith("```"):
            continue

        if line.startswith("|") and line.endswith("|"):

            if re.match(
                r"^\|[\s\-:|]+\|$",
                line
            ):
                continue

            in_table = True

            cells = [
                c.strip()
                for c in line[1:-1].split("|")
            ]

            table_rows.append(cells)

            continue

        if in_table:
            flush_table()

        if not line:

            html.append("<p>&nbsp;</p>")

            continue

        escaped = (
            line
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

        escaped = re.sub(
            r"\*\*(.*?)\*\*",
            r"<strong>\1</strong>",
            escaped
        )

        if escaped.startswith("### "):

            html.append(
                f"<h3>{escaped[4:]}</h3>"
            )

        elif escaped.startswith("## "):

            html.append(
                f"<h2>{escaped[3:]}</h2>"
            )

        elif escaped.startswith("# "):

            html.append(
                f"<h1>{escaped[2:]}</h1>"
            )

        elif escaped.startswith("---"):

            html.append("<hr>")

        else:

            html.append(
                f"<p>{escaped}</p>"
            )

    if in_table:
        flush_table()

    return "\n".join(html)


# =========================================================
# PROMPT GEMINI
# =========================================================

def buat_prompt(data: dict) -> str:

    profil = (
        ", ".join(data["profil_pancasila"])
        if data["profil_pancasila"]
        else "Sesuai karakteristik materi"
    )

    return f"""
Bertindaklah sebagai ahli penyusunan perangkat pembelajaran
Indonesia yang memahami Kurikulum Merdeka.

Susun dokumen berikut:

{data["jenis_perangkat"]}

IDENTITAS SEKOLAH

Dinas Pendidikan:
{data["dinas"]}

Satuan Pendidikan:
{data["sekolah"]}

Alamat:
{data["alamat"]}

Kota/Kabupaten:
{data["kota"]}

IDENTITAS GURU

Jabatan:
{data["jabatan_guru"]}

Nama Guru:
{data["guru_nama"]}

NIP:
{data["guru_nip"]}

Kepala Sekolah:
{data["ks_nama"]}

NIP Kepala Sekolah:
{data["ks_nip"]}

Pengawas Pembina:
{data["pengawas_nama"]}

NIP Pengawas:
{data["pengawas_nip"]}

PARAMETER PEMBELAJARAN

Mata Pelajaran:
{data["mapel"]}

Fase/Kelas:
{data["fase_kelas"]}

Semester:
{data["semester"]}

Alokasi Waktu:
{data["alokasi_waktu"]}

Materi Pokok:
{data["materi_pokok"]}

Profil Pelajar Pancasila:
{profil}

Tanggal:
{data["tanggal"]}

ATURAN DOKUMEN

1. Gunakan bahasa Indonesia formal dan baku.
2. Dokumen harus lengkap dan operasional.
3. Jangan menggunakan placeholder seperti "[isi]",
   "[sesuaikan]", "...", "dst.", atau "dan seterusnya".
4. Gunakan heading Markdown.
5. Gunakan tabel Markdown jika diperlukan.
6. Jangan mengarang nomor regulasi apabila tidak yakin.
7. Jangan memberikan penjelasan di luar dokumen.
8. Jangan menggunakan blok kode Markdown.
9. Buat dokumen yang siap diedit dan digunakan guru.
10. Akhiri dengan bagian pengesahan.

Mulai langsung dari dokumen.
"""


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        f"""
        <div class="brand-logo">
            <img
                src="{LOGO_URL}"
                alt="Logo SIAP AJAR 22"
            >
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<h3 style='text-align:center;'>SIAP AJAR 22</h3>",
        unsafe_allow_html=True,
    )

    st.caption(
        "Portal Administrasi Pembelajaran"
    )

    st.divider()

    st.markdown("### 👤 Pendidik Aktif")

    st.success(
        f"**{st.session_state.user_name}**\n\n"
        "🟢 Sesi AI aktif"
    )

    if st.button(
        "🚪 Keluar / Ganti Akun",
        use_container_width=True,
    ):

        st.session_state.authenticated = False
        st.session_state.user_name = ""
        st.session_state.user_api_key = ""
        st.session_state.hasil_teks = ""

        st.rerun()

    st.divider()

    st.markdown("### ⚙️ Status Sistem")

    st.markdown(
        """
        **Engine:** Google Gemini

        **Output:** DOCX / TXT

        **Perangkat:** 22 Dokumen
        """
    )

# =========================================================
# HEADER DASHBOARD
# =========================================================

st.title(
    f"Selamat Berkarya, {st.session_state.user_name}"
)

st.caption(
    "Susun perangkat pembelajaran secara otomatis, "
    "lengkap, sistematis, dan siap digunakan."
)

# =========================================================
# STATISTIK
# =========================================================

c1, c2, c3 = st.columns(3)

with c1:

    st.markdown(
        """
        <div class="info-card">
            <div class="info-number">22</div>
            <div class="info-label">
                Perangkat Pembelajaran
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c2:

    st.markdown(
        """
        <div class="info-card">
            <div class="info-number">AI</div>
            <div class="info-label">
                Google Gemini
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c3:

    st.markdown(
        """
        <div class="info-card">
            <div class="info-number">DOCX</div>
            <div class="info-label">
                Format Siap Unduh
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# KONFIGURASI PEMBELAJARAN
# =========================================================

with st.container(border=True):

    st.markdown(
        '<div class="section-title">'
        '📋 1. Konfigurasi Pembelajaran'
        '</div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:

        fase_kelas = st.selectbox(
            "Tingkatan Kelas / Fase",
            [
                "Fase A - Kelas 1 SD/MI",
                "Fase A - Kelas 2 SD/MI",
                "Fase B - Kelas 3 SD/MI",
                "Fase B - Kelas 4 SD/MI",
                "Fase C - Kelas 5 SD/MI",
                "Fase C - Kelas 6 SD/MI",
                "Fase D - Kelas 7 SMP/MTs",
                "Fase D - Kelas 8 SMP/MTs",
                "Fase D - Kelas 9 SMP/MTs",
                "Fase E - Kelas 10 SMA/MA/SMK",
                "Fase F - Kelas 11 SMA/MA/SMK",
                "Fase F - Kelas 12 SMA/MA/SMK",
            ],
            index=3,
        )

        is_sd = "SD/MI" in fase_kelas

        match = re.search(
            r"Kelas (\d+)",
            fase_kelas
        )

        angka_kelas = (
            match.group(1)
            if match
            else ""
        )

        if is_sd:

            jabatan_guru = (
                f"Guru Kelas {angka_kelas}"
                if angka_kelas
                else "Guru Kelas"
            )

            label_mapel = (
                "Mata Pelajaran / Muatan Pelajaran"
            )

            default_mapel = "IPAS"

        else:

            jabatan_guru = "Guru Mata Pelajaran"

            label_mapel = "Mata Pelajaran"

            default_mapel = "Fisika"

        mapel = st.text_input(
            label_mapel,
            value=default_mapel,
        )

        materi_pokok = st.text_input(
            "Fokus Topik / Materi Pokok",
            placeholder="Contoh: Usaha dan Energi",
        )

    with col2:

        alokasi_waktu = st.text_input(
            "Alokasi Waktu / Target JP",
            value=(
                "4 JP / Minggu"
                if is_sd
                else "3 JP / Minggu"
            ),
        )

        profil_pancasila = st.multiselect(
            "Dimensi Profil Pelajar Pancasila",
            [
                "Beriman, Bertakwa kepada Tuhan YME "
                "& Berakhlak Mulia",
                "Berkebinekaan Global",
                "Gotong Royong",
                "Mandiri",
                "Bernalar Kritis",
                "Kreatif",
            ],
            default=[
                "Bernalar Kritis",
                "Gotong Royong",
                "Mandiri",
            ],
        )


# =========================================================
# 22 PERANGKAT
# =========================================================

daftar_22_perangkat = [

    "01. Analisis Capaian Pembelajaran (CP) & Pemetaan Elemen",

    "02. Alur Tujuan Pembelajaran (ATP) Lengkap",

    "03. Program Tahunan (PROTA)",

    "04. Program Semester (PROMES)",

    "05. Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)",

    "06. Modul Ajar Lengkap (Format Baku Kurikulum Merdeka)",

    "07. Lembar Kerja Peserta Didik (LKPD) Berdiferensiasi",

    "08. Modul / Panduan Projek Profil Pancasila (P5)",

    "09. Jurnal Mengajar Harian & Agenda Guru Terstruktur",

    "10. Format & Kisi-kisi Asesmen Diagnostik",

    "11. Kisi-kisi & Instrumen Asesmen Formatif",

    "12. Kisi-kisi, Naskah Soal, & Kunci Jawaban Asesmen Sumatif",

    "13. Rubrik Penilaian Kinerja, Portofolio, serta Proyek",

    "14. Rekapitulasi Daftar Nilai Kurikulum Merdeka",

    "15. Buku Presensi / Lembar Absensi Siswa",

    "16. Format Program Pembelajaran Remedial & Pengayaan",

    "17. Distribusi Sumber Belajar & Buku Teks Pembelajaran",

    "18. Analisis Alokasi Waktu Efektif",

    "19. Format Analisis Kuantitatif Butir Soal Evaluasi",

    "20. Jurnal Sikap & Catatan Dimensi Profil Pelajar Pancasila",

    "21. Panduan Layanan Bimbingan & Konsultasi Akademik",

    "22. Format Laporan Evaluasi Diri Guru & Rencana Tindak Lanjut",
]

jenis_perangkat = st.selectbox(
    "📄 Pilih Dokumen yang Akan Dibuat",
    daftar_22_perangkat,
)


# =========================================================
# SEMESTER
# =========================================================

semester = st.radio(
    "Semester Berjalan",
    [
        "Semester Ganjil",
        "Semester Genap",
        "Semester Ganjil & Genap (1 Tahun Penuh)",
    ],
    horizontal=True,
)


# =========================================================
# IDENTITAS SEKOLAH
# =========================================================

tab1, tab2 = st.tabs(
    [
        "🏫 Identitas Satuan Pendidikan",
        "✍️ Pejabat & Penandatangan",
    ]
)


with tab1:

    c1, c2 = st.columns(2)

    with c1:

        dinas = st.text_input(
            "Dinas Pendidikan Pembina",
            value=(
                "DINAS PENDIDIKAN KOTA PONTIANAK"
                if is_sd
                else
                "DINAS PENDIDIKAN PROVINSI "
                "KALIMANTAN BARAT"
            ),
        )

        sekolah = st.text_input(
            "Nama Satuan Pendidikan",
            value=(
                "SDN 01 PONTIANAK"
                if is_sd
                else
                "SMAS NUSA HARAPAN"
            ),
        )

    with c2:

        alamat = st.text_input(
            "Alamat & Kontak Sekolah",
            value=(
                "Jl. Pancasila No. 10, "
                "Telp. (0561) 734567"
            ),
        )

        kota = st.text_input(
            "Kota / Kabupaten Domisili",
            value="Pontianak",
        )


# =========================================================
# PENANDATANGAN
# =========================================================

with tab2:

    c1, c2, c3 = st.columns(3)

    with c1:

        st.markdown(
            f"**{jabatan_guru}**"
        )

        guru_nama = st.text_input(
            "Nama Guru",
            value=st.session_state.user_name,
            key="guru_nama",
        )

        guru_nip = st.text_input(
            "NIP Guru",
            value="-",
            key="guru_nip",
        )

    with c2:

        st.markdown(
            "**Kepala Sekolah**"
        )

        ks_nama = st.text_input(
            "Nama Kepala Sekolah",
            placeholder="Contoh: Zulkifli, S.Pd.",
            key="ks_nama",
        )

        ks_nip = st.text_input(
            "NIP Kepala Sekolah",
            placeholder="Contoh: 197508122005011004",
            key="ks_nip",
        )

    with c3:

        st.markdown(
            "**Pengawas Pembina**"
        )

        pengawas_nama = st.text_input(
            "Nama Pengawas Pembina",
            placeholder="Contoh: Andar",
            key="pengawas_nama",
        )

        pengawas_nip = st.text_input(
            "NIP Pengawas",
            placeholder="Contoh: 196811231993032003",
            key="pengawas_nip",
        )


# =========================================================
# TOMBOL TERBITKAN
# =========================================================

st.divider()

if st.button(
    "✨ TERBITKAN DOKUMEN ADMINISTRASI RESMI",
    use_container_width=True,
    type="primary",
):

    if not mapel.strip():

        st.warning(
            "⚠️ Mata pelajaran wajib diisi."
        )

        st.stop()

    if not sekolah.strip():

        st.warning(
            "⚠️ Nama sekolah wajib diisi."
        )

        st.stop()

    if not materi_pokok.strip():

        st.warning(
            "⚠️ Fokus topik/materi pokok wajib diisi."
        )

        st.stop()

    data = {
        "dinas": dinas,
        "sekolah": sekolah,
        "alamat": alamat,
        "kota": kota,
        "jabatan_guru": jabatan_guru,
        "guru_nama": guru_nama,
        "guru_nip": guru_nip,
        "ks_nama": ks_nama or "-",
        "ks_nip": ks_nip or "-",
        "pengawas_nama": pengawas_nama or "-",
        "pengawas_nip": pengawas_nip or "-",
        "mapel": mapel,
        "fase_kelas": fase_kelas,
        "semester": semester,
        "alokasi_waktu": alokasi_waktu,
        "materi_pokok": materi_pokok,
        "profil_pancasila": profil_pancasila,
        "jenis_perangkat": jenis_perangkat,
        "tanggal": datetime.now().strftime(
            "%d-%m-%Y"
        ),
    }

    prompt = buat_prompt(data)

    progress = st.progress(0)

    status = st.empty()

    try:

        status.info(
            "⚡ Menghubungkan ke Google Gemini..."
        )

        client = genai.Client(
            api_key=st.session_state.user_api_key
        )

        progress.progress(15)

          progress.progress(15)

        model_list = [
            "gemini-3.6-flash",
        ]

        response = None
        model_berhasil = ""

        errors = []      response = None
        model_berhasil = ""

        errors = []

        for model_name in model_list:

            for percobaan in range(3):

                try:

                    status.info(
                        f"📝 Membuat dokumen dengan "
                        f"{model_name} "
                        f"({percobaan + 1}/3)"
                    )

                    response = (
                        client.models.generate_content(
                            model=model_name,
                            contents=prompt,
                        )
                    )

                    if (
                        response
                        and getattr(
                            response,
                            "text",
                            None
                        )
                    ):

                        model_berhasil = model_name

                        break

                except Exception as exc:

                    error_text = str(exc)

                    errors.append(
                        f"{model_name}: "
                        f"{error_text}"
                    )

                    if (
                        "503" in error_text
                        or "UNAVAILABLE"
                        in error_text
                    ):

                        tunggu = 2 ** percobaan

                        status.warning(
                            f"⏳ Gemini sedang sibuk. "
                            f"Mencoba lagi dalam "
                            f"{tunggu} detik..."
                        )

                        time.sleep(tunggu)

                    else:

                        break

            if (
                response
                and getattr(
                    response,
                    "text",
                    None
                )
            ):

                break

        progress.progress(90)

        if (
            not response
            or not getattr(
                response,
                "text",
                None
            )
        ):

            raise RuntimeError(
                "Semua model Gemini gagal.\n\n"
                + "\n".join(errors[-6:])
            )

        # -------------------------
        # SIMPAN HASIL
        # -------------------------

        st.session_state.hasil_teks = response.text

        nama_bersih = re.sub(
            r"^\d+\.\s*",
            "",
            jenis_perangkat,
        )

        st.session_state.nama_file_base = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            nama_bersih,
        ).strip("_")

        progress.progress(100)

        status.success(
            f"✅ Dokumen berhasil dibuat "
            f"dengan {model_berhasil}."
        )

        time.sleep(0.8)

        status.empty()
        progress.empty()

    except Exception as exc:

        progress.empty()
        status.empty()

        st.error(
            "❌ Gagal menerbitkan dokumen."
        )

        st.code(
            str(exc),
            language="text",
        )


# =========================================================
# PREVIEW
# =========================================================

if st.session_state.hasil_teks:

    st.divider()

    st.markdown(
        "### 📄 Preview Lembar Kerja A4"
    )

    preview = markdown_to_html(
        st.session_state.hasil_teks
    )

    st.markdown(
        f"""
        <div class="paper-a4">
            {preview}
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -------------------------
    # DOWNLOAD
    # -------------------------

    c1, c2 = st.columns(2)

    with c1:

        docx_file = buat_file_docx(
            st.session_state.hasil_teks
        )

        st.download_button(
            "📄 Unduh Berkas Word (.DOCX)",
            data=docx_file,
            file_name=(
                f"{st.session_state.nama_file_base}.docx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with c2:

        st.download_button(
            "📝 Unduh Teks Mentah (.TXT)",
            data=st.session_state.hasil_teks,
            file_name=(
                f"{st.session_state.nama_file_base}.txt"
            ),
            mime="text/plain",
            use_container_width=True,
        )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer-box">
        SIAP AJAR 22 • Creator: Andar<br>
        © 2026 SIAP AJAR 22 • Engine AI Pembelajaran
    </div>
    """,
    unsafe_allow_html=True,
)
