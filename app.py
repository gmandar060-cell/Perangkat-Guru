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
# KONFIGURASI HALAMAN
# =========================================================

st.set_page_config(
    page_title="SIAP AJAR 22 | Portal Administrasi Kurikulum",
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

    * {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background: #f8fafc;
    }

    .block-container {
        max-width: 1120px;
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    /* ================= LOGIN ================= */

    .login-page {
        position: fixed;
        inset: 0;
        background:
            radial-gradient(circle at 20% 20%, rgba(37,99,235,.12), transparent 30%),
            radial-gradient(circle at 80% 80%, rgba(14,165,233,.10), transparent 30%),
            #f8fafc;
        z-index: -1;
    }

    .login-hero {
        text-align: center;
        padding: 30px 20px 18px 20px;
    }

    .login-icon {
        width: 82px;
        height: 82px;
        margin: 0 auto;
        border-radius: 22px;
        background: white;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 12px 30px rgba(15,23,42,.12);
        border: 1px solid rgba(15,23,42,.06);
    }

    .login-icon img {
        width: 65px;
        height: 65px;
        object-fit: contain;
    }

    .login-badge {
        display: inline-block;
        padding: 8px 15px;
        border-radius: 999px;
        background: #eff6ff;
        color: #2563eb;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: .7px;
    }

    .login-title {
        margin: 20px 0 8px 0;
        font-size: 46px !important;
        line-height: 1.1;
        font-weight: 800 !important;
        color: #0f172a;
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
        justify-content: center;
        gap: 14px;
        margin: 15px auto 25px auto;
        max-width: 760px;
    }

    .stat-card {
        flex: 1;
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 16px 12px;
        box-shadow: 0 8px 25px rgba(15,23,42,.05);
    }

    .stat-number {
        font-size: 22px;
        font-weight: 800;
        color: #2563eb;
    }

    .stat-label {
        font-size: 11px;
        color: #64748b;
        margin-top: 4px;
    }

    .login-form-title {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        text-align: center;
        margin-bottom: 4px;
    }

    .login-form-desc {
        color: #64748b;
        text-align: center;
        font-size: 13px;
        margin-bottom: 20px;
    }

    .login-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 22px;
        padding: 28px;
        box-shadow: 0 15px 45px rgba(15,23,42,.08);
        max-width: 560px;
        margin: 0 auto;
    }

    .api-help {
        margin-top: -5px;
        margin-bottom: 12px;
        font-size: 12px;
        text-align: right;
    }

    .api-help a {
        color: #2563eb;
        text-decoration: none;
        font-weight: 600;
    }

    .creator-box {
        text-align: center;
        margin-top: 20px;
        color: #94a3b8;
        font-size: 12px;
    }

    /* ================= DASHBOARD ================= */

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
        margin-bottom: 24px;
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
        color: rgba(255,255,255,.88);
        line-height: 1.7;
    }

    .section-title {
        font-size: 21px;
        font-weight: 800;
        color: #0f172a;
        margin: 10px 0 5px 0;
    }

    .section-desc {
        color: #64748b;
        font-size: 13px;
        margin-bottom: 18px;
    }

    .device-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 15px;
        padding: 14px;
        margin-bottom: 10px;
    }

    .paper-a4 {
        background: white;
        width: 100%;
        min-height: 800px;
        padding: 55px;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        box-shadow: 0 8px 30px rgba(15,23,42,.08);
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

    /* ================= SIDEBAR ================= */

    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    /* ================= MOBILE ================= */

    @media (max-width: 700px) {

        .login-title {
            font-size: 36px !important;
        }

        .stats-row {
            flex-direction: column;
        }

        .login-card {
            padding: 20px;
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


def render_logo(width=90):

    st.markdown(
        f"""
        <div style="
            display:flex;
            justify-content:center;
            align-items:center;
            margin-bottom:12px;
        ">

            <img
                src="{LOGO_URL}"
                width="{width}"
                height="{width}"
                alt="Logo SIAP AJAR 22"
                style="
                    object-fit:contain;
                    filter:drop-shadow(
                        0 4px 6px rgba(0,0,0,.1)
                    );
                "
            >

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
# FUNGSI DOCX
# =========================================================

def buat_file_docx(markdown_text: str) -> io.BytesIO:

    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    for line in markdown_text.splitlines():

        line = line.strip()

        if not line:
            document.add_paragraph()
            continue

        # Heading
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

        # Horizontal rule
        if line in ["---", "***", "___"]:

            document.add_paragraph("_" * 90)
            continue

        # Bullet
        if line.startswith("- "):

            document.add_paragraph(
                line[2:],
                style="List Bullet"
            )

            continue

        # Numbered list
        if re.match(r"^\d+\.\s", line):

            text = re.sub(
                r"^\d+\.\s",
                "",
                line
            )

            document.add_paragraph(
                text,
                style="List Number"
            )

            continue

        # Table
        if "|" in line:

            parts = [
                p.strip()
                for p in line.strip("|").split("|")
            ]

            if all(
                re.fullmatch(r":?-+:?", p)
                for p in parts
            ):
                continue

            table = document.add_table(
                rows=1,
                cols=len(parts)
            )

            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            for i, part in enumerate(parts):

                table.rows[0].cells[i].text = part

            continue

        # Normal paragraph
        p = document.add_paragraph(line)

        for run in p.runs:
            run.font.size = Pt(11)

    output = io.BytesIO()

    document.save(output)

    output.seek(0)

    return output


# =========================================================
# MARKDOWN → HTML
# =========================================================

def markdown_to_html(text: str) -> str:

    lines = text.splitlines()

    html = []

    in_table = False

    for line in lines:

        stripped = line.strip()

        if not stripped:

            if in_table:

                html.append("</table>")
                in_table = False

            continue

        if stripped.startswith("### "):

            html.append(
                f"<h3>{stripped[4:]}</h3>"
            )

            continue

        if stripped.startswith("## "):

            html.append(
                f"<h2>{stripped[3:]}</h2>"
            )

            continue

        if stripped.startswith("# "):

            html.append(
                f"<h1>{stripped[2:]}</h1>"
            )

            continue

        if stripped.startswith("- "):

            html.append(
                f"<li>{stripped[2:]}</li>"
            )

            continue

        if "|" in stripped:

            parts = [
                p.strip()
                for p in stripped.strip("|").split("|")
            ]

            if all(
                re.fullmatch(r":?-+:?", p)
                for p in parts
            ):
                continue

            if not in_table:

                html.append("<table>")
                in_table = True

            html.append("<tr>")

            for part in parts:

                html.append(
                    f"<td>{part}</td>"
                )

            html.append("</tr>")

            continue

        html.append(
            f"<p>{stripped}</p>"
        )

    if in_table:
        html.append("</table>")

    return "\n".join(html)


# =========================================================
# PROMPT AI
# =========================================================

def buat_instruksi_prompt(data: dict) -> str:

    perangkat = "\n".join(
        [
            f"{i + 1}. {item}"
            for i, item in enumerate(
                data["perangkat"]
            )
        ]
    )

    prompt = f"""
Anda adalah ahli penyusunan perangkat pembelajaran
Kurikulum Merdeka di Indonesia.

Buat perangkat pembelajaran yang lengkap, rapi,
praktis digunakan guru, dan siap disalin ke Microsoft Word.

IDENTITAS PEMBELAJARAN

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

PERANGKAT YANG DIMINTA

{perangkat}

KETENTUAN:

1. Gunakan bahasa Indonesia yang formal,
   jelas, dan mudah digunakan guru.

2. Sesuaikan seluruh isi dengan Kurikulum Merdeka.

3. Jangan membuat data yang tidak relevan
   dengan fase, kelas, mata pelajaran, dan materi.

4. Gunakan tabel apabila format perangkat
   lebih mudah dibaca dalam bentuk tabel.

5. Buat isi yang praktis dan siap digunakan.

6. Setiap perangkat harus memiliki judul yang jelas.

7. Sertakan identitas sekolah dan guru
   jika relevan.

8. Pastikan tujuan pembelajaran,
   kegiatan pembelajaran, asesmen,
   dan materi saling berkaitan.

9. Hindari penjelasan yang terlalu teoritis.

10. Hasil akhir harus langsung berupa
    dokumen perangkat pembelajaran,
    bukan penjelasan tentang cara membuatnya.

Di bagian akhir, buat tabel tanda tangan:

Guru,
{data["guru_nama"]}

Kepala Sekolah,
{data["kepala_sekolah"]}

Pengawas Pembina,
{data["pengawas"]}
"""

    return prompt

```python
# =========================================================
# HALAMAN LOGIN
# =========================================================

if not st.session_state.authenticated:

    st.markdown(
        '<div class="login-page"></div>',
        unsafe_allow_html=True
    )

    # ================= HERO =================

    st.markdown(
        """
        <div class="login-hero">

            <div style="
                display:flex;
                justify-content:center;
                align-items:center;
                flex-direction:column;
                gap:12px;
            ">

                <div class="login-icon">

                    <img
                        src="https://raw.githubusercontent.com/gmandar060-cell/Perangkat-Guru/main/logo.png"
                        alt="Tut Wuri Handayani"
                    >

                </div>

                <div class="login-badge">
                    ✦ &nbsp; KURIKULUM MERDEKA
                </div>

            </div>

            <h1 class="login-title">
                SIAP AJAR <span>22</span>
            </h1>

            <div class="login-subtitle">
                Satu Portal, Solusi Lengkap
                <strong>
                    22 Perangkat Pembelajaran
                </strong>
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # ================= STATISTIK =================

    st.markdown(
        """
        <div class="stats-row">

            <div class="stat-card">
                <div class="stat-number">4</div>
                <div class="stat-label">
                    Kategori Terorganisir
                </div>
            </div>

            <div class="stat-card">
                <div class="stat-number">22</div>
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


    # ================= FORM LOGIN =================

    st.markdown(
        '<div class="login-card">',
        unsafe_allow_html=True
    )

    with st.form(
        "login_form",
        clear_on_submit=False,
        border=False
    ):

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

        nama_guru_input = st.text_input(
            "Nama Lengkap & Gelar",
            placeholder="Contoh: Andar Prasetyo, S.Pd.",
            key="login_nama"
        )

        api_key_masuk = st.text_input(
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
                    ⓘ &nbsp;
                    Bagaimana cara mendapatkan API key gratis?
                </a>

            </div>
            """,
            unsafe_allow_html=True
        )

        masuk = st.form_submit_button(
            "Masuk ke Portal  →",
            use_container_width=True
        )

    st.markdown(
        '</div>',
        unsafe_allow_html=True
    )


    # ================= VALIDASI LOGIN =================

    if masuk:

        if not nama_guru_input.strip():

            st.warning(
                "⚠️ Silakan isi Nama Lengkap & Gelar."
            )

        elif not api_key_masuk.strip():

            st.warning(
                "⚠️ Silakan masukkan Gemini API Key."
            )

        else:

            st.session_state.user_name = (
                nama_guru_input.strip()
            )

            st.session_state.user_api_key = (
                api_key_masuk.strip()
            )

            st.session_state.authenticated = True

            st.rerun()


    # ================= CREATOR =================

    st.markdown(
        """
        <div class="creator-box">

            Creator: <strong>Andar</strong><br>

            SIAP AJAR 22 • Portal Administrasi Pembelajaran

        </div>
        """,
        unsafe_allow_html=True
    )

    st.stop()
```

# =========================================================
# SIDEBAR DASHBOARD
# =========================================================

with st.sidebar:

    render_logo(70)

    st.markdown(
        """
        <div style="
            text-align:center;
            font-weight:800;
            font-size:18px;
            color:#0f172a;
        ">
            SIAP AJAR 22
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(
        "Portal Administrasi Pembelajaran"
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

        - **Engine:** Google Gemini
        - **Output:** DOCX / TXT
        - **Perangkat:** 22 dokumen
        """
    )


# =========================================================
# HEADER DASHBOARD
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
            platform SIAP AJAR 22.
        </p>

    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# KONFIGURASI PEMBELAJARAN
# =========================================================

st.markdown(
    '<div class="section-title">📚 Konfigurasi Pembelajaran</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-desc">Tentukan identitas pembelajaran sebelum membuat perangkat.</div>',
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
        [
            "Ganjil",
            "Genap",
        ],
        horizontal=True
    )


is_sd = "SD/MI" in fase_kelas

match_kelas = re.search(
    r"Kelas\s+(\d+)",
    fase_kelas
)

angka_kelas = (
    match_kelas.group(1)
    if match_kelas
    else "1"
)

if is_sd:

    jabatan_guru_otomatis = (
        f"Guru Kelas {angka_kelas}"
    )

    label_mapel = (
        "Mata Pelajaran / Muatan Pelajaran"
    )

    default_mapel = "IPAS"

else:

    jabatan_guru_otomatis = (
        "Guru Mata Pelajaran"
    )

    label_mapel = "Mata Pelajaran"

    default_mapel = "Fisika"


col1, col2 = st.columns(2)

with col1:

    mapel = st.text_input(
        label_mapel,
        value=default_mapel
    )

with col2:

    materi_pokok = st.text_input(
        "Materi Pokok",
        placeholder="Contoh: Pembangunan Ekonomi"
    )


col1, col2 = st.columns(2)

with col1:

    alokasi_waktu = st.text_input(
        "Alokasi Waktu",
        value="2 x 45 menit"
    )

with col2:

    profil_pancasila = st.multiselect(
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

daftar_22_perangkat = [

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
    '<div class="section-title">🗂️ Pilih Perangkat Pembelajaran</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-desc">Pilih satu atau beberapa perangkat yang ingin dibuat.</div>',
    unsafe_allow_html=True
)

perangkat_dipilih = st.multiselect(
    "22 Perangkat",
    daftar_22_perangkat,
    default=[
        "Modul Ajar Lengkap"
    ]
)


# =========================================================
# IDENTITAS SEKOLAH
# =========================================================

st.markdown(
    '<div class="section-title">🏫 Identitas Satuan Pendidikan</div>',
    unsafe_allow_html=True
)

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "Dinas Pendidikan Pembina",
        "Nama Satuan Pendidikan",
        "Alamat & Kontak Sekolah",
        "Kota/Kabupaten",
    ]
)

if is_sd:

    default_dinas = (
        "DINAS PENDIDIKAN KOTA PONTIANAK"
    )

    default_sekolah = (
        "SDN 01 PONTIANAK"
    )

else:

    default_dinas = (
        "DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT"
    )

    default_sekolah = (
        "SMAS NUSA HARAPAN"
    )


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
# DATA PEJABAT
# =========================================================

st.markdown(
    '<div class="section-title">✍️ Data Penandatangan</div>',
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
# TOMBOL GENERATE
# =========================================================

st.divider()

generate = st.button(
    "🚀 GENERATE PERANGKAT PEMBELAJARAN",
    type="primary",
    use_container_width=True
)


# =========================================================
# PROSES GENERATE
# =========================================================

if generate:

    if not mapel.strip():

        st.error(
            "❌ Mata pelajaran belum diisi."
        )

        st.stop()

    if not materi_pokok.strip():

        st.error(
            "❌ Materi pokok belum diisi."
        )

        st.stop()

    if not sekolah.strip():

        st.error(
            "❌ Nama satuan pendidikan belum diisi."
        )

        st.stop()

    if not perangkat_dipilih:

        st.error(
            "❌ Pilih minimal satu perangkat."
        )

        st.stop()

    if not st.session_state.user_api_key.strip():

        st.error(
            "❌ Gemini API Key belum tersedia."
        )

        st.stop()


    data_input = {

        "guru_nama": guru_nama,
        "guru_nip": guru_nip,

        "jabatan": jabatan_guru_otomatis,

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
        "materi": materi_pokok,

        "alokasi": alokasi_waktu,
        "semester": semester,

        "profil_pancasila": profil_pancasila,

        "perangkat": perangkat_dipilih,
    }


    prompt_final = buat_instruksi_prompt(
        data_input
    )


    progress = st.progress(0)

    status = st.empty()

    status.info(
        "⏳ Menghubungkan ke Google Gemini..."
    )


    try:

        client = genai.Client(
            api_key=st.session_state.user_api_key
        )

        model_list = [

            "gemini-3.6-flash",

            "gemini-3.8-flash",
        ]

        response = None

        error_messages = []


        for model_name in model_list:

            if response is not None:
                break

            status.info(
                f"🤖 Menggunakan model {model_name}..."
            )

            for attempt in range(3):

                try:

                    progress.progress(
                        min(
                            90,
                            10
                            + (
                                attempt * 15
                            )
                        )
                    )

                    response = client.models.generate_content(

                        model=model_name,

                        contents=prompt_final,
                    )

                    if response and response.text:

                        break


                except Exception as e:

                    error_text = str(e)

                    error_messages.append(
                        f"{model_name} percobaan "
                        f"{attempt + 1}: {error_text}"
                    )

                    if (
                        "503" in error_text
                        or "UNAVAILABLE"
                        in error_text.upper()
                    ):

                        time.sleep(
                            2 ** attempt
                        )

                    else:

                        break


        if response is None or not response.text:

            progress.progress(100)

            status.empty()

            st.error(
                "❌ Gagal menghasilkan perangkat."
            )

            if error_messages:

                with st.expander(
                    "Detail error"
                ):

                    for error in error_messages:

                        st.write(error)

        else:

            progress.progress(100)

            status.success(
                "✅ Perangkat berhasil dibuat."
            )

            st.session_state.hasil_teks = (
                response.text
            )

            safe_name = re.sub(
                r"[^A-Za-z0-9_-]+",
                "_",
                materi_pokok.strip()
            )

            st.session_state.nama_file_base = (
                f"SIAP_AJAR_22_"
                f"{safe_name}"
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
# PREVIEW HASIL
# =========================================================

if st.session_state.hasil_teks:

    st.divider()

    st.markdown(
        '<div class="section-title">📄 Preview Dokumen</div>',
        unsafe_allow_html=True
    )

    preview_html = markdown_to_html(
        st.session_state.hasil_teks
    )

    st.markdown(
        f"""
        <div class="paper-a4">
            {preview_html}
        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown("### 📥 Download Dokumen")

    col1, col2 = st.columns(2)


    # DOCX

    with col1:

        docx_file = buat_file_docx(
            st.session_state.hasil_teks
        )

        st.download_button(

            label="📘 Download DOCX",

            data=docx_file,

            file_name=(
                st.session_state.nama_file_base
                + ".docx"
            ),

            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            ),

            use_container_width=True,
        )


    # TXT

    with col2:

        txt_file = (
            st.session_state.hasil_teks
            .encode("utf-8")
        )

        st.download_button(

            label="📝 Download TXT",

            data=txt_file,

            file_name=(
                st.session_state.nama_file_base
                + ".txt"
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
