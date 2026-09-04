import streamlit as st

# =========================================================
# KONFIGURASI HALAMAN
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

* {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(37, 99, 235, 0.10), transparent 35%),
        radial-gradient(circle at bottom right, rgba(14, 165, 233, 0.08), transparent 35%),
        #f8fafc;
}

/* Hilangkan menu bawaan */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    visibility: hidden;
}

/* Container */
.block-container {
    max-width: 900px;
    padding-top: 45px;
    padding-bottom: 30px;
}

/* Header */
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
    font-size: 42px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1.1;
    margin-bottom: 8px;
}

.login-subtitle {
    text-align: center;
    color: #64748b;
    font-size: 15px;
    margin-bottom: 32px;
}

/* Statistik */
.stats-container {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 15px;
    margin-bottom: 25px;
}

.stat-card {
    background: rgba(255,255,255,0.90);
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 20px 12px;
    text-align: center;
    box-shadow: 0 8px 25px rgba(15,23,42,0.05);
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

/* Login Card */
.login-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 22px;
    padding: 30px;
    box-shadow: 0 15px 40px rgba(15,23,42,0.08);
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
    margin-bottom: 20px;
}

/* Input */
.stTextInput label {
    font-weight: 600 !important;
    color: #334155 !important;
}

.stTextInput input {
    border-radius: 12px !important;
    border: 1px solid #cbd5e1 !important;
    padding: 12px !important;
}

.stTextInput input:focus {
    border-color: #2563eb !important;
}

/* Button */
.stButton > button {
    width: 100%;
    border-radius: 12px;
    background: #2563eb;
    color: white;
    border: none;
    padding: 12px 20px;
    font-size: 15px;
    font-weight: 700;
    transition: all 0.2s ease;
}

.stButton > button:hover {
    background: #1d4ed8;
    transform: translateY(-1px);
}

/* API info */
.api-info {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 12px;
    padding: 12px 15px;
    margin-top: 15px;
    color: #1e40af;
    font-size: 12px;
    line-height: 1.5;
}

.api-info a {
    color: #1d4ed8;
    font-weight: 700;
    text-decoration: none;
}

/* Footer */
.creator {
    text-align: center;
    color: #94a3b8;
    font-size: 11px;
    margin-top: 20px;
    line-height: 1.6;
}

/* Mobile */
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
        padding: 15px;
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
# =========================================================

st.markdown("""
<div class="stats-container">

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
""", unsafe_allow_html=True)

# =========================================================
# LOGIN CARD
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
# INFO API KEY
# =========================================================

st.markdown("""
<div class="api-info">
    🔑 Belum memiliki Gemini API Key?
    <a href="https://aistudio.google.com/app/apikey" target="_blank">
        Buat API Key di Google AI Studio
    </a>
</div>
""", unsafe_allow_html=True)

st.write("")

# =========================================================
# LOGIN
# =========================================================

if st.button("🚀 MASUK KE PORTAL", use_container_width=True):

    if not nama.strip():
        st.error("❌ Silakan masukkan Nama Lengkap & Gelar.")
        st.stop()

    if not api_key.strip():
        st.error("❌ Silakan masukkan Gemini API Key.")
        st.stop()

    # Simpan sesi login
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
