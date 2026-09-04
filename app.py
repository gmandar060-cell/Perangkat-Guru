import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import re
import time


# =========================================================
# KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="SIAP AJAR 22",
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

    @import url(
        'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:'
        'wght@400;500;600;700;800&display=swap'
    );

    * {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background: #f8fafc;
    }

    .block-container {
        max-width: 1120px;
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    /* =====================================================
       LOGIN
       ===================================================== */

    .login-hero {
        max-width: 820px;
        margin: 0 auto;
        text-align: center;
        padding: 25px 10px 10px 10px;
    }

    .login-logo {
        width: 92px;
        height: 92px;
        margin: 0 auto 15px auto;
        border-radius: 24px;
        background: white;
        border: 1px solid #e2e8f0;
        box-shadow: 0 15px 35px rgba(15,23,42,.10);
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .login-logo img {
        width: 75px;
        height: 75px;
        object-fit: contain;
    }

    .login-badge {
        display: inline-block;
        padding: 8px 16px;
        border-radius: 999px;
        background: #eff6ff;
        color: #2563eb;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: .6px;
    }

    .login-title {
        margin: 18px 0 7px 0;
        color: #0f172a !important;
        font-size: 46px !important;
        line-height: 1.1 !important;
        font-weight: 800 !important;
    }

    .login-title span {
        color: #2563eb;
    }

    .login-subtitle {
        color: #64748b;
        font-size: 15px;
        line-height: 1.7;
    }

    .login-subtitle strong {
        color: #334155;
    }

    .stats-row {
        display: flex;
        gap: 14px;
        margin: 20px auto 25px auto;
        max-width: 760px;
    }

    .stat-card {
        flex: 1;
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 17px;
        padding: 17px 12px;
        text-align: center;
        box-shadow: 0 8px 25px rgba(15,23,42,.05);
    }

    .stat-number {
        font-size: 22px;
        font-weight: 800;
        color: #2563eb;
    }

    .stat-label {
        margin-top: 4px;
        font-size: 11px;
        color: #64748b;
    }

    .login-form-title {
        text-align: center;
        color: #0f172a;
        font-size: 24px;
        font-weight: 800;
        margin-bottom: 5px;
    }

    .login-form-desc {
        text-align: center;
        color: #64748b;
        font-size: 13px;
        margin-bottom: 20px;
    }

    .api-help {
        text-align: right;
        margin-top: 5px;
        margin-bottom: 12px;
        font-size: 12px;
    }

    .api-help a {
        color: #2563eb;
        text-decoration: none;
        font-weight: 700;
    }

    .creator-box {
        text-align: center;
        color: #94a3b8;
        font-size: 12px;
        margin-top: 20px;
        line-height: 1.7;
    }

    /* =====================================================
       DASHBOARD
       ===================================================== */

    .app-header {
        background: linear-gradient(
            135deg,
            #1d4ed8,
            #2563eb,
            #0ea5e9
        );
        color: white;
        padding: 30px;
        border-radius: 22px;
        margin-bottom: 25px;
        box-shadow: 0 15px 35px rgba(37,99,235,.20);
    }

    .app-header h1 {
        color: white !important;
        margin: 0 0 8px 0;
        font-size: 30px;
        font-weight: 800;
    }

    .app-header p {
        margin: 0;
        color: rgba(255,255,255,.90);
        line-height: 1.7;
    }

    .section-title {
        color: #0f172a;
        font-size: 21px;
        font-weight: 800;
        margin: 14px 0 5px 0;
    }

    .section-desc {
        color: #64748b;
        font-size: 13px;
        margin-bottom: 18px;
    }

    .paper-a4 {
        background: white;
        min-height: 700px;
        padding: 50px;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        box-shadow: 0 8px 30px rgba(15,23,42,.07);
        line-height: 1.7;
        color: #111827;
    }

    .paper-a4 h1,
    .paper-a4 h2,
    .paper-a4 h3 {
        color: #111827;
    }

    .paper-a4 table {
        width: 100%;
        border-collapse: collapse;
        margin: 15px 0;
    }

    .paper-a4 th,
    .paper-a4 td {
        border: 1px solid #cbd5e1;
        padding: 8px;
        vertical-align: top;
    }

    .paper-a4 th {
        background: #f1f5f9;
        font-weight: 700;
    }

    .footer-box {
        text-align: center;
        color: #94a3b8;
        font-size: 11px;
        padding: 30px 0 10px 0;
    }

    section[data-testid="stSidebar"] {
        background: white;
        border-right: 1px solid #e2e8f0;
    }

    @media (max-width: 700px) {

        .login-title {
            font-size: 36px !important;
        }

        .stats-row {
            flex-direction: column;
        }

        .paper-a4 {
            padding: 25px;
        }
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# LOGO
# =========================================================

LOGO_URL = (
    "https://raw.githubusercontent.com/"
    "gmandar060-cell/Perangkat-Guru/"
    "main/logo.png"
)


# =========================================================
# SESSION STATE
# =========================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "user_name" not in st.session_state:
    st.session_state.user_name = ""

if "user_api_key" not in st.session_state:
    st.session_state.user_api_key = ""

if "hasil_teks" not in st.session_state:
    st.session_state.hasil_teks = ""

if "nama_file_base" not in st.session_state:
    st.session_state.nama_file_base = "SIAP_AJAR_22"


# =========================================================
# FUNGSI DOCX
# =========================================================

def buat_file_docx(text):

    document = Document()

    section = document.sections[0]

    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    for line in text.splitlines():

        line = line.strip()

        if not line:
            document.add_paragraph()
            continue

        if line.startswith("### "):

            p = document.add_paragraph()

            run = p.add_run(line[4:])
            run.bold = True
            run.font.size = Pt(13)

            continue

        if line.startswith("## "):

            p = document.add_paragraph()

            run = p.add_run(line[3:])
            run.bold = True
            run.font.size = Pt(15)

            continue

        if line.startswith("# "):

            p = document.add_paragraph()

            run = p.add_run(line[2:])
            run.bold = True
            run.font.size = Pt(17)

            continue

        if line in ["---", "***", "___"]:

            document.add_paragraph("_" * 80)

            continue

        if line.startswith("- "):

            document.add_paragraph(
                line[2:],
                style="List Bullet"
            )

            continue

        if re.match(r"^\d+\.\s", line):

            clean = re.sub(
                r"^\d+\.\s",
                "",
                line
            )

            document.add_paragraph(
                clean,
                style="List Number"
            )

            continue

        if "|" in line:

            parts = [
                x.strip()
                for x in line.strip("|").split("|")
            ]

            if all(
                re.fullmatch(r":?-+:?", x)
                for x in parts
            ):
                continue

            table = document.add_table(
                rows=1,
                cols=len(parts)
            )

            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            for i, value in enumerate(parts):

                table.rows[0].cells[i].text = value

            continue

        p = document.add_paragraph(line)

        for run in p.runs:
            run.font.size = Pt(11)

    output = io.BytesIO()

    document.save(output)

    output.seek(0)

    return output


# =========================================================
# MARKDOWN KE HTML
# =========================================================

def markdown_to_html(text):

    html = []

    in_table = False

    for line in text.splitlines():

        line = line.strip()

        if not line:

            if in_table:

                html.append("</table>")

                in_table = False

            continue

        if line.startswith("### "):

            html.append(
                f"<h3>{line[4:]}</h3>"
            )

            continue

        if line.startswith("## "):

            html.append(
                f"<h2>{line[3:]}</h2>"
            )

            continue

        if line.startswith("# "):

            html.append(
                f"<h1>{line[2:]}</h1>"
            )

            continue

        if line.startswith("- "):

            html.append(
                f"<li>{line[2:]}</li>"
            )

            continue

        if "|" in line:

            parts = [
                x.strip()
                for x in line.strip("|").split("|")
            ]

            if all(
                re.fullmatch(r":?-+:?", x)
                for x in parts
            ):
                continue

            if not in_table:

                html.append("<table>")

                in_table = True

            html.append("<tr>")

            for value in parts:

                html.append(
                    f"<td>{value}</td>"
                )

            html.append("</tr>")

            continue

        html.append(
            f"<p>{line}</p>"
        )

    if in_table:

        html.append("</table>")

    return "\n".join(html)


# =========================================================
# PROMPT AI
# =========================================================

def buat_prompt(data):

    perangkat = "\n".join(
        [
            f"{i + 1}. {item}"
            for i, item in enumerate(data["perangkat"])
        ]
    )

    return f"""
Anda adalah ahli penyusunan perangkat pembelajaran
Kurikulum Merdeka di Indonesia.

Buat perangkat pembelajaran yang lengkap, sistematis,
praktis digunakan guru, dan siap dipindahkan ke Microsoft Word.

IDENTITAS

Nama Guru:
{data["guru_nama"]}

NIP Guru:
{data["guru_nip"]}

Jabatan:
{data["jabatan"]}

Kepala Sekolah:
{data["kepala_sekolah"]}

NIP Kepala Sekolah:
{data["kepala_sekolah_nip"]}

Pengawas Pembina:
{data["pengawas"]}

NIP Pengawas:
{data["pengawas_nip"]}

Dinas Pendidikan:
{data["dinas"]}

Satuan Pendidikan:
{data["sekolah"]}

Alamat:
{data["alamat"]}

Kota/Kabupaten:
{data["kota"]}

Fase/Kelas:
{data["fase_kelas"]}

Mata Pelajaran:
{data["mapel"]}

Materi Pokok:
{data["materi"]}

Alokasi Waktu:
{data["alokasi"]}

Semester:
{data["semester"]}

Profil Pelajar Pancasila:
{", ".join(data["profil_pancasila"])}

PERANGKAT YANG DIMINTA:

{perangkat}

KETENTUAN:

1. Gunakan bahasa Indonesia formal dan jelas.
2. Sesuaikan isi dengan Kurikulum Merdeka.
3. Sesuaikan dengan fase, kelas, mata pelajaran,
   materi, dan alokasi waktu.
4. Gunakan tabel jika lebih mudah digunakan guru.
5. Buat perangkat secara praktis dan siap pakai.
6. Setiap perangkat harus memiliki judul yang jelas.
7. Hubungkan tujuan, kegiatan, materi, dan asesmen.
8. Hindari teori yang tidak diperlukan.
9. Jangan menjelaskan cara membuat perangkat.
10. Langsung hasilkan isi dokumen.

Pada bagian akhir sertakan:

TANDA TANGAN

Guru,
{data["guru_nama"]}

Kepala Sekolah,
{data["kepala_sekolah"]}

Pengawas Pembina,
{data["pengawas"]}
"""


# =========================================================
# LOGIN
# =========================================================

if not st.session_state.authenticated:

    # -----------------------------------------------------
    # HERO
    # -----------------------------------------------------

    st.markdown(
        f"""
        <div class="login-hero">

            <div class="login-logo">
                <img
                    src="{LOGO_URL}"
                    alt="Logo Tut Wuri Handayani"
                >
            </div>

            <div class="login-badge">
                ✦ KURIKULUM MERDEKA
            </div>

            <div class="login-title">
                SIAP <span>AJAR 22</span>
            </div>

            <div class="login-subtitle">
                Satu Portal, Solusi Lengkap
                <strong>22 Perangkat Pembelajaran</strong>
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # STATISTIK
    # -----------------------------------------------------

    st.markdown(
        """
        <div class="stats-row">

            <div class="stat-card">

                <div class="stat-number">
                    4
                </div>

                <div class="stat-label">
                    Kategori Terorganisir
                </div>

            </div>

            <div class="stat-card">

                <div class="stat-number">
                    22
                </div>

                <div class="stat-label">
                    Dokumen AI Generated
                </div>

            </div>

            <div class="stat-card">

                <div class="stat-number">
                    Gemini AI
                </div>

                <div class="stat-label">
                    Flash Model
                </div>

            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # FORM LOGIN
    # -----------------------------------------------------

    st.markdown(
        """
        <div class="login-form-title">
            Masuk ke Portal
        </div>

        <div class="login-form-desc">
            Isi data berikut untuk memulai
        </div>
        """,
        unsafe_allow_html=True
    )


    with st.form(
        "login_form",
        clear_on_submit=False
    ):

        nama = st.text_input(
            "Nama Lengkap & Gelar",
            placeholder="Contoh: Andar, S.Pd.",
            key="login_nama"
        )

        api_key = st.text_input(
            "Gemini API Key",
            type="password",
            placeholder="AIza...",
            key="login_api_key"
        )

        st.markdown(
            """
            <div class="api-help">
                <a
                    href="https://aistudio.google.com/apikey"
                    target="_blank"
                >
                    ⓘ Cara mendapatkan API Key gratis
                </a>
            </div>
            """,
            unsafe_allow_html=True
        )

        masuk = st.form_submit_button(
            "Masuk ke Portal  →",
            type="primary",
            use_container_width=True
        )


    # -----------------------------------------------------
    # PROSES LOGIN
    # -----------------------------------------------------

    if masuk:

        if not nama.strip():

            st.warning(
                "⚠️ Silakan isi Nama Lengkap & Gelar."
            )

        elif not api_key.strip():

            st.warning(
                "⚠️ Silakan masukkan Gemini API Key."
            )

        else:

            st.session_state.user_name = nama.strip()

            st.session_state.user_api_key = api_key.strip()

            st.session_state.authenticated = True

            st.rerun()


    # -----------------------------------------------------
    # CREATOR
    # -----------------------------------------------------

    st.markdown(
        """
        <div class="creator-box">

            Creator:
            <strong>Andar</strong>
            <br>

            SIAP AJAR 22 •
            Portal Administrasi Pembelajaran

        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        f"""
        <div style="
            text-align:center;
            margin-bottom:12px;
        ">

            <img
                src="{LOGO_URL}"
                width="72"
                height="72"
                style="object-fit:contain;"
            >

        </div>

        <div style="
            text-align:center;
            font-size:18px;
            font-weight:800;
            color:#0f172a;
        ">
            SIAP AJAR 22
        </div>

        <div style="
            text-align:center;
            font-size:11px;
            color:#64748b;
            margin-top:4px;
        ">
            Portal Administrasi Pembelajaran
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    st.markdown("#### 👤 Pendidik Aktif")

    st.success(
        f"**{st.session_state.user_name}**\n\n"
        "🟢 Sesi AI aktif"
    )

    if st.button(
        "🔄 Keluar / Ganti Akun",
        use_container_width=True
    ):

        st.session_state.authenticated = False
        st.session_state.user_name = ""
        st.session_state.user_api_key = ""
        st.session_state.hasil_teks = ""

        st.rerun()

    st.divider()

    st.markdown(
        """
        **⚙️ Status Sistem**

        - 🟢 Google Gemini
        - 🟢 Output DOCX / TXT
        - 🟢 22 Perangkat
        """
    )


# =========================================================
# HEADER
# =========================================================

st.markdown(
    f"""
    <div class="app-header">

        <h1>
            Selamat Berkarya,
            {st.session_state.user_name}
        </h1>

        <p>
            Susun perangkat pembelajaran secara terstruktur
            dan ekspor langsung ke Microsoft Word melalui
            SIAP AJAR 22.
        </p>

    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# KONFIGURASI PEMBELAJARAN
# =========================================================

st.markdown(
    '<div class="section-title">'
    '📚 Konfigurasi Pembelajaran'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-desc">'
    'Tentukan identitas pembelajaran sebelum membuat perangkat.'
    '</div>',
    unsafe_allow_html=True
)


col1, col2 = st.columns(2)

with col1:

    fase_kelas = st.selectbox(
        "Fase / Kelas",
        [
            "Fase A — Kelas 1 SD/MI",
            "Fase A — Kelas 2 SD/MI",
            "Fase B — Kelas 3 SD/MI",
            "Fase B — Kelas 4 SD/MI",
            "Fase C — Kelas 5 SD/MI",
            "Fase C — Kelas 6 SD/MI",
            "Fase D — Kelas 7 SMP/MTs",
            "Fase D — Kelas 8 SMP/MTs",
            "Fase D — Kelas 9 SMP/MTs",
            "Fase E — Kelas 10 SMA/MA/SMK",
            "Fase F — Kelas 11 SMA/MA/SMK",
            "Fase F — Kelas 12 SMA/MA/SMK",
        ]
    )

with col2:

    semester = st.radio(
        "Semester",
        ["Ganjil", "Genap"],
        horizontal=True
    )


is_sd = "SD/MI" in fase_kelas

hasil_kelas = re.search(
    r"Kelas\s+(\d+)",
    fase_kelas
)

angka_kelas = (
    hasil_kelas.group(1)
    if hasil_kelas
    else "1"
)


if is_sd:

    jabatan = f"Guru Kelas {angka_kelas}"

    label_mapel = "Mata Pelajaran / Muatan Pelajaran"

    default_mapel = "IPAS"

else:

    jabatan = "Guru Mata Pelajaran"

    label_mapel = "Mata Pelajaran"

    default_mapel = "Fisika"


col1, col2 = st.columns(2)

with col1:

    mapel = st.text_input(
        label_mapel,
        value=default_mapel
    )

with col2:

    materi = st.text_input(
        "Materi Pokok",
        placeholder="Contoh: Pembangunan Ekonomi"
    )


col1, col2 = st.columns(2)

with col1:

    alokasi = st.text_input(
        "Alokasi Waktu",
        value="2 x 45 menit"
    )

with col2:

    profil = st.multiselect(
        "Profil Pelajar Pancasila",
        [
            "Beriman, bertakwa kepada Tuhan YME, dan berakhlak mulia",
            "Berkebinekaan global",
            "Gotong royong",
            "Mandiri",
            "Bernalar kritis",
            "Kreatif",
        ],
        default=[
            "Bernalar kritis",
            "Mandiri",
        ]
    )


# =========================================================
# 22 PERANGKAT
# =========================================================

daftar_perangkat = [

    "Analisis CP & Pemetaan Elemen",

    "ATP Lengkap",

    "PROTA",

    "PROMES",

    "KKTP",

    "Modul Ajar Lengkap",

    "LKPD Berdiferensiasi",

    "Modul/Panduan P5",

    "Jurnal Mengajar Harian & Agenda Guru",

    "Asesmen Diagnostik",

    "Asesmen Formatif",

    "Asesmen Sumatif",

    "Rubrik Penilaian Kinerja/Portofolio/Proyek",

    "Rekap Daftar Nilai",

    "Presensi",

    "Remedial & Pengayaan",

    "Sumber Belajar & Buku Teks",

    "Analisis Alokasi Waktu Efektif",

    "Analisis Kuantitatif Butir Soal",

    "Jurnal Sikap & Catatan P5",

    "Bimbingan & Konsultasi Akademik",

    "Evaluasi Diri Guru & RTL",
]


st.markdown(
    '<div class="section-title">'
    '🗂️ Pilih Perangkat Pembelajaran'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-desc">'
    'Pilih satu atau beberapa perangkat yang ingin dibuat.'
    '</div>',
    unsafe_allow_html=True
)


perangkat = st.multiselect(
    "22 Perangkat",
    daftar_perangkat,
    default=["Modul Ajar Lengkap"]
)


# =========================================================
# IDENTITAS SEKOLAH
# =========================================================

st.markdown(
    '<div class="section-title">'
    '🏫 Identitas Satuan Pendidikan'
    '</div>',
    unsafe_allow_html=True
)


tab1, tab2, tab3, tab4 = st.tabs(
    [
        "Dinas Pendidikan",
        "Satuan Pendidikan",
        "Alamat & Kontak",
        "Kota/Kabupaten",
    ]
)


if is_sd:

    default_dinas = "DINAS PENDIDIKAN KOTA PONTIANAK"

    default_sekolah = "SDN 01 PONTIANAK"

else:

    default_dinas = (
        "DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT"
    )

    default_sekolah = "SMAS NUSA HARAPAN"


with tab1:

    dinas = st.text_input(
        "Dinas Pendidikan",
        value=default_dinas
    )


with tab2:

    sekolah = st.text_input(
        "Nama Satuan Pendidikan",
        value=default_sekolah
    )


with tab3:

    alamat = st.text_input(
        "Alamat & Kontak",
        value="Jl. Pancasila No. 10, Telp. (0561) 734567"
    )


with tab4:

    kota = st.text_input(
        "Kota/Kabupaten",
        value="Pontianak"
    )


# =========================================================
# DATA PENANDATANGAN
# =========================================================

st.markdown(
    '<div class="section-title">'
    '✍️ Data Penandatangan'
    '</div>',
    unsafe_allow_html=True
)


col1, col2 = st.columns(2)

with col1:

    guru_nama = st.text_input(
        "Nama Guru",
        value=st.session_state.user_name
    )

    guru_nip = st.text_input(
        "NIP Guru",
        placeholder="Masukkan NIP jika ada"
    )


with col2:

    kepala_sekolah = st.text_input(
        "Nama Kepala Sekolah",
        placeholder="Nama Kepala Sekolah"
    )

    kepala_sekolah_nip = st.text_input(
        "NIP Kepala Sekolah",
        placeholder="Masukkan NIP jika ada"
    )


col1, col2 = st.columns(2)

with col1:

    pengawas = st.text_input(
        "Nama Pengawas Pembina",
        placeholder="Nama Pengawas"
    )


with col2:

    pengawas_nip = st.text_input(
        "NIP Pengawas",
        placeholder="Masukkan NIP jika ada"
    )


# =========================================================
# GENERATE
# =========================================================

st.divider()


generate = st.button(
    "🚀 GENERATE PERANGKAT PEMBELAJARAN",
    type="primary",
    use_container_width=True
)


if generate:

    if not mapel.strip():

        st.error(
            "❌ Mata pelajaran belum diisi."
        )

        st.stop()


    if not materi.strip():

        st.error(
            "❌ Materi pokok belum diisi."
        )

        st.stop()


    if not sekolah.strip():

        st.error(
            "❌ Nama satuan pendidikan belum diisi."
        )

        st.stop()


    if not perangkat:

        st.error(
            "❌ Pilih minimal satu perangkat."
        )

        st.stop()


    if not st.session_state.user_api_key.strip():

        st.error(
            "❌ Gemini API Key belum tersedia."
        )

        st.stop()


    data = {

        "guru_nama": guru_nama,

        "guru_nip": guru_nip,

        "jabatan": jabatan,

        "kepala_sekolah": kepala_sekolah,

        "kepala_sekolah_nip": kepala_sekolah_nip,

        "pengawas": pengawas,

        "pengawas_nip": pengawas_nip,

        "dinas": dinas,

        "sekolah": sekolah,

        "alamat": alamat,

        "kota": kota,

        "fase_kelas": fase_kelas,

        "mapel": mapel,

        "materi": materi,

        "alokasi": alokasi,

        "semester": semester,

        "profil_pancasila": profil,

        "perangkat": perangkat,

    }


    prompt = buat_prompt(data)


    progress = st.progress(0)

    status = st.empty()


    status.info(
        "⏳ Menghubungkan ke Google Gemini..."
    )


    try:

        client = genai.Client(
            api_key=st.session_state.user_api_key
        )


        # =================================================
        # MODEL GEMINI
        # =================================================

        model_list = [

            "gemini-2.5-flash",

            "gemini-2.0-flash",

        ]


        response = None

        errors = []


        for model in model_list:

            if response is not None:
                break


            status.info(
                f"🤖 Menggunakan {model}..."
            )


            for attempt in range(3):

                try:

                    progress.progress(
                        min(
                            90,
                            10 + attempt * 20
                        )
                    )


                    response = (
                        client.models.generate_content(
                            model=model,
                            contents=prompt
                        )
                    )


                    if response and response.text:

                        break


                except Exception as e:

                    error_text = str(e)

                    errors.append(
                        f"{model} | Percobaan "
                        f"{attempt + 1} | {error_text}"
                    )


                    upper_error = error_text.upper()


                    if (
                        "503" in error_text
                        or "UNAVAILABLE" in upper_error
                        or "429" in error_text
                    ):

                        time.sleep(
                            2 ** attempt
                        )

                    else:

                        break


        progress.progress(100)


        # =================================================
        # GAGAL
        # =================================================

        if response is None or not response.text:

            status.empty()

            st.error(
                "❌ Gagal menghasilkan perangkat."
            )


            with st.expander(
                "Detail error"
            ):

                for error in errors:

                    st.write(error)


        # =================================================
        # BERHASIL
        # =================================================

        else:

            status.success(
                "✅ Perangkat berhasil dibuat."
            )


            st.session_state.hasil_teks = (
                response.text
            )


            safe_name = re.sub(
                r"[^A-Za-z0-9_-]+",
                "_",
                materi.strip()
            )


            st.session_state.nama_file_base = (
                f"SIAP_AJAR_22_{safe_name}"
            )


            st.success(
                "🎉 Dokumen siap dipreview dan diunduh."
            )


    except Exception as e:

        progress.progress(100)

        status.empty()


        st.error(
            "❌ Terjadi kesalahan saat menjalankan AI."
        )


        with st.expander(
            "Detail error"
        ):

            st.code(str(e))


# =========================================================
# PREVIEW
# =========================================================

if st.session_state.hasil_teks:

    st.divider()


    st.markdown(
        '<div class="section-title">'
        '📄 Preview Dokumen'
        '</div>',
        unsafe_allow_html=True
    )


    html_preview = markdown_to_html(
        st.session_state.hasil_teks
    )


    st.markdown(
        f"""
        <div class="paper-a4">

            {html_preview}

        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        "### 📥 Download Dokumen"
    )


    col1, col2 = st.columns(2)


    with col1:

        docx_file = buat_file_docx(
            st.session_state.hasil_teks
        )


        st.download_button(
            "📘 Download DOCX",

            data=docx_file,

            file_name=(
                st.session_state.nama_file_base
                + ".docx"
            ),

            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            ),

            use_container_width=True
        )


    with col2:

        txt_file = (
            st.session_state.hasil_teks
            .encode("utf-8")
        )


        st.download_button(
            "📝 Download TXT",

            data=txt_file,

            file_name=(
                st.session_state.nama_file_base
                + ".txt"
            ),

            mime="text/plain",

            use_container_width=True
        )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer-box">

        SIAP AJAR 22 • Creator: Andar
        <br>

        © 2026 SIAP AJAR 22 • Engine AI Pembelajaran

    </div>
    """,
    unsafe_allow_html=True
)
