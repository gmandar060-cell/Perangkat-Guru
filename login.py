import streamlit as st

# =========================================================
# SIAP AJAR 22 - HALAMAN LOGIN
# =========================================================

st.set_page_config(
    page_title="SIAP AJAR 22 | Login",
    page_icon="🎓",
    layout="centered",
)

# =========================================================
# LOGO
# =========================================================

LOGO_URL = "https://raw.githubusercontent.com/gmandar060-cell/Perangkat-Guru/main/logo.png"

# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, #DBEAFE 0%, transparent 35%),
            radial-gradient(circle at bottom right, #E0E7FF 0%, transparent 35%),
            #F8FAFC !important;
    }

    .block-container {
        max-width: 900px !important;
        padding-top: 45px !important;
        padding-bottom: 30px !important;
    }

    .login-header {
        text-align: center;
        margin-bottom: 25px;
    }

    .login-logo img {
        width: 90px;
        height: 90px;
        object-fit: contain;
        margin-bottom: 10px;
    }

    .kurikulum {
        color: #2563EB !important;
        font-size: 13px;
        font-weight: 800;
        letter-spacing: 1.5px;
        margin-bottom: 8px;
    }

    .login-title {
        color: #0F172A !important;
        font-size: 34px;
        font-weight: 800;
        margin: 0;
    }

    .login-subtitle {
        color: #64748B !important;
        font-size: 14px;
        margin-top: 8px;
    }

    .stats {
        display: flex;
        gap: 12px;
        margin: 25px 0;
    }

    .stat-card {
        flex: 1;
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 17px 10px;
        text-align: center;
    }

    .stat-number {
        color: #1E3A8A !important;
        font-size: 22px;
        font-weight: 800;
    }

    .stat-label {
        color: #64748B !important;
        font-size: 11px;
        margin-top: 3px;
    }

    .login-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 18px;
        padding: 28px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.07);
    }

    .login-card-title {
        color: #0F172A !important;
        font-size: 21px;
        font-weight: 800;
        margin-bottom: 5px;
    }

    .login-card-text {
        color: #64748B !important;
        font-size: 13px;
        margin-bottom: 20px;
    }

    .api-help {
        background: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 10px;
        padding: 12px 14px;
        margin: 8px 0 18px 0;
        font-size: 12px;
    }

    .api-help a {
        color: #1D4ED8 !important;
        text-decoration: none;
        font-weight: 700;
    }

    .creator {
        text-align: center;
        color: #64748B !important;
        font-size: 12px;
        margin-top: 25px;
    }

    .footer {
        text-align: center;
        color: #94A3B8 !important;
        font-size: 11px;
        margin-top: 25px;
        padding-top: 15px;
        border-top: 1px solid #E2E8F0;
    }

    .stButton > button {
        background: linear-gradient(
            135deg,
            #1E3A8A,
            #2563EB
        ) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        min-height: 48px !important;
        font-weight: 800 !important;
        font-size: 14px !important;
    }

    @media (max-width: 600px) {
        .block-container {
            padding: 25px 15px !important;
        }

        .login-title {
            font-size: 28px;
        }

        .stats {
            gap: 6px;
        }

        .stat-card {
            padding: 12px 5px;
        }

        .stat-number {
            font-size: 18px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
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

# =========================================================
# JIKA SUDAH LOGIN
# =========================================================

if st.session_state.authenticated:
    st.switch_page("dashboard.py")

# =========================================================
# HEADER
# =========================================================

st.markdown(
    f"""
    <div class="login-header">

        <div class="login-logo">
            <img
                src="{LOGO_URL}"
                alt="Logo Tut Wuri Handayani"
            >
        </div>

        <div class="kurikulum">
            ✦ KURIKULUM MERDEKA
        </div>

        <div class="login-title">
            SIAP AJAR 22
        </div>

        <div class="login-subtitle">
            Satu Portal, Solusi Lengkap 22 Perangkat Pembelajaran
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# STATISTIK
# =========================================================

st.markdown(
    """
    <div class="stats">

        <div class="stat-card">
            <div class="stat-number">22</div>
            <div class="stat-label">Perangkat Pembelajaran</div>
        </div>

        <div class="stat-card">
            <div class="stat-number">AI</div>
            <div class="stat-label">Google Gemini</div>
        </div>

        <div class="stat-card">
            <div class="stat-number">DOCX</div>
            <div class="stat-label">Siap Diunduh</div>
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# LOGIN CARD
# =========================================================

st.markdown(
    """
    <div class="login-card">

        <div class="login-card-title">
            🔐 Masuk ke Portal
        </div>

        <div class="login-card-text">
            Masukkan nama Anda dan Gemini API Key untuk menggunakan
            SIAP AJAR 22.
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)

nama_guru = st.text_input(
    "Nama Lengkap & Gelar",
    placeholder="Contoh: Andar, S.Pd.",
)

api_key = st.text_input(
    "Gemini API Key",
    type="password",
    placeholder="AIza...",
)

st.markdown(
    """
    <div class="api-help">
        🔑
        <a
            href="https://aistudio.google.com/apikey"
            target="_blank"
        >
            Cara mendapatkan Gemini API Key gratis →
        </a>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.button(
    "🚀 MASUK KE PORTAL",
    use_container_width=True,
):

    if not nama_guru.strip():
        st.warning("⚠️ Nama lengkap wajib diisi.")

    elif not api_key.strip():
        st.warning("⚠️ Gemini API Key wajib diisi.")

    else:

        # Simpan data login ke session
        st.session_state.user_name = nama_guru.strip()
        st.session_state.user_api_key = api_key.strip()
        st.session_state.authenticated = True

        # Masuk ke dashboard
        st.switch_page("dashboard.py")

# =========================================================
# CREATOR
# =========================================================

st.markdown(
    """
    <div class="creator">
        SIAP AJAR 22 • Creator: <b>Andar</b>
    </div>

    <div class="footer">
        © 2026 SIAP AJAR 22 • Engine AI Pembelajaran
    </div>
    """,
    unsafe_allow_html=True,
)
