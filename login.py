import streamlit as st

# =========================================================
# KONFIGURASI
# =========================================================

st.set_page_config(
    page_title="SIAP AJAR 22 | Login",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(
            circle at top left,
            rgba(37, 99, 235, 0.10),
            transparent 35%
        ),
        radial-gradient(
            circle at bottom right,
            rgba(14, 165, 233, 0.08),
            transparent 35%
        ),
        #f8fafc;
}

/* Hilangkan menu Streamlit */
#MainMenu {
    visibility: hidden;
}

header {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

/* Container */
.block-container {
    max-width: 900px;
    padding-top: 45px;
    padding-bottom: 30px;
}

/* =========================================================
   HEADER
   ========================================================= */

.kurikulum {
    text-align: center;
    color: #2563eb;
    font-size: 14px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 10px;
}

.login-title {
    text-align: center;
    color: #0f172a;
    font-size: 42px;
    font-weight: 800;
    line-height: 1.1;
    margin-bottom: 8px;
}

.login-subtitle {
    text-align: center;
    color: #64748b;
    font-size: 15px;
    margin-bottom: 30px;
}

/* =========================================================
   STATISTIK
   ========================================================= */

.stats-container {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
    margin-bottom: 25px;
}

.stat-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 20px 12px;
    text-align: center;
    box-shadow: 0 8px 25px rgba(15, 23, 42, 0.05);
}

.stat-number {
    color: #2563eb;
    font-size: 25px;
    font-weight: 800;
    margin-bottom: 5px;
}

.stat-label {
    color: #64748b;
    font-size: 12px;
    font-weight: 600;
}

/* =========================================================
   LOGIN CARD
   ========================================================= */

.login-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 22px;
    padding: 30px;
    box-shadow: 0 15px 40px rgba(15, 23, 42, 0.08);
    margin-bottom: 20px;
}

.login-card-title {
    color: #0f172a;
    font-size: 23px;
    font-weight: 800;
    margin-bottom: 8px;
}

.login-card-text {
    color: #64748b;
    font-size: 13px;
    line-height: 1.6;
}

/* =========================================================
   INPUT
   ========================================================= */

.stTextInput {
    margin-bottom: 5px;
}

.stTextInput label {
    color: #334155 !important;
    font-weight: 600 !important;
}

.stTextInput input {
    border-radius: 12px !important;
    border: 1px solid #cbd5e1 !important;
    padding: 12px !important;
}

.stTextInput input:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 1px #2563eb !important;
}

/* =========================================================
   BUTTON
   ========================================================= */

.stButton > button {
    width: 100%;
    min-height: 48px;
    border-radius: 12px;
    background: #2563eb;
    color: white;
    border: none;
    font-size: 15px;
    font-weight: 700;
}

.stButton > button:hover {
    background: #1d4ed8;
    color: white;
}

/* =========================================================
   API INFO
   ========================================================= */

.api-info {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 12px;
    padding: 13px 15px;
    margin-top: 10px;
    color: #1e40af;
    font-size: 12px;
    line-height: 1.5;
}

.api-info a {
    color: #1d4ed8;
    font-weight: 700;
    text-decoration: none;
}

/* =========================================================
   FOOTER
   ========================================================= */

.creator {
    text-align: center;
    color: #94a3b8;
    font-size: 11px;
    margin-top: 22px;
    line-height: 1.6;
}

/* =========================================================
   MOBILE
   ========================================================= */

@media (max-width: 650px) {

    .block-container {
        padding-top: 25px;
        padding-left: 18px;
        padding-right: 18px;
    }

    .login-title {
        font-size: 32px;
    }

    .login-subtitle {
        font-size: 13px;
    }

    .stats-container {
        grid-template-columns: 1fr;
        gap: 10px;
    }

    .stat-card {
        padding: 16px;
    }

    .login-card {
        padding: 22px;
    }
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="kurikulum">
    ✦ KURIKULUM MERDEKA
</div>

<div class="login-title">
    SIAP AJAR 22
</div>

<div class="login-subtitle">
    Satu Portal, Solusi Lengkap 22 Perangkat Pembelajaran
</div>
""", unsafe_allow_html=True)


# =========================================================
# STATISTIK
# SATU BLOK HTML UTUH
# =========================================================

st.markdown("""
<div class="stats-container">

    <div class="stat-card">
        <div class="stat-number">22</div>
        <div class="stat-label">
            Perangkat Pembelajaran
        </div>
    </div>

    <div class="stat-card">
        <div class="stat-number">AI</div>
        <div class="stat-label">
            Google Gemini
        </div>
    </div>

    <div class="stat-card">
        <div class="stat-number">DOCX</div>
        <div class="stat-label">
            Siap Diunduh
        </div>
    </div>

</div>
""", unsafe_allow_html=True)


# =========================================================
# LOGIN CARD
# SATU BLOK HTML UTUH
# =========================================================

st.markdown("""
<div class="login-card">

    <div class="login-card-title">
        🔐 Masuk ke Portal
    </div>

    <div class="login-card-text">
        Masukkan nama Anda dan Gemini API Key untuk menggunakan
        SIAP AJAR 22.
    </div>

</div>
""", unsafe_allow_html=True)


# =========================================================
# FORM LOGIN
# =========================================================

nama = st.text_input(
    "Nama Lengkap & Gelar",
    placeholder="Contoh: Andar, S.Pd."
)

api_key = st.text_input(
    "Gemini API Key",
    type="password",
    placeholder="Masukkan Gemini API Key Anda"
)


# =========================================================
# INFORMASI API KEY
# =========================================================

st.markdown("""
<div class="api-info">
    🔑 Belum memiliki Gemini API Key?
    <a
        href="https://aistudio.google.com/app/apikey"
        target="_blank"
    >
        Buat API Key di Google AI Studio
    </a>
</div>
""", unsafe_allow_html=True)

st.write("")


# =========================================================
# TOMBOL LOGIN
# =========================================================

if st.button(
    "🚀 MASUK KE PORTAL",
    use_container_width=True
):

    if not nama.strip():
        st.error("❌ Silakan masukkan Nama Lengkap & Gelar.")

    elif not api_key.strip():
        st.error("❌ Silakan masukkan Gemini API Key.")

    else:

        # Simpan informasi login
        st.session_state.authenticated = True
        st.session_state.user_name = nama.strip()
        st.session_state.user_api_key = api_key.strip()

        # Masuk ke dashboard
        st.switch_page("dashboard.py")


# =========================================================
# FOOTER
# =========================================================

st.markdown("""
<div class="creator">
    SIAP AJAR 22<br>
    Portal Administrasi Pembelajaran Guru<br>
    <b>© 2026 Pak Andar</b>
</div>
""", unsafe_allow_html=True)
