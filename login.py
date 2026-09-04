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
# SESSION STATE
# =========================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "user_name" not in st.session_state:
    st.session_state.user_name = ""

if "user_api_key" not in st.session_state:
    st.session_state.user_api_key = ""

# =========================================================
# JIKA SUDAH LOGIN → DASHBOARD
# =========================================================

if st.session_state.authenticated:

    halaman = st.navigation(
        [
            st.Page(
                "dashboard.py",
                title="Dashboard",
                icon="🎓"
            )
        ],
        position="hidden"
    )

    halaman.run()
    st.stop()

# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

#MainMenu {
    visibility: hidden;
}

header {
    visibility: hidden;
}

footer {
    visibility: hidden;
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

.block-container {
    max-width: 850px;
    padding-top: 50px;
    padding-bottom: 30px;
}

.kurikulum {
    text-align: center;
    color: #2563eb;
    font-size: 14px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-bottom: 10px;
}

.judul {
    text-align: center;
    color: #0f172a;
    font-size: 42px;
    font-weight: 800;
    line-height: 1.1;
    margin-bottom: 8px;
}

.subjudul {
    text-align: center;
    color: #64748b;
    font-size: 15px;
    margin-bottom: 30px;
}

.stat-box {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 20px 8px;
    text-align: center;
    box-shadow: 0 5px 18px rgba(15, 23, 42, 0.05);
}

.stat-number {
    color: #2563eb;
    font-size: 25px;
    font-weight: 800;
}

.stat-text {
    color: #64748b;
    font-size: 12px;
    font-weight: 600;
}

.login-box {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 20px;
    padding: 28px;
    margin-top: 25px;
    margin-bottom: 20px;
    box-shadow: 0 8px 25px rgba(15, 23, 42, 0.06);
}

.login-title {
    color: #0f172a;
    font-size: 22px;
    font-weight: 800;
    margin-bottom: 8px;
}

.login-text {
    color: #64748b;
    font-size: 13px;
    line-height: 1.6;
}

.api-box {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 12px;
    padding: 12px;
    color: #1e40af;
    font-size: 12px;
    margin-top: 8px;
}

.api-box a {
    color: #1d4ed8;
    font-weight: 700;
    text-decoration: none;
}

.stButton > button {
    width: 100%;
    min-height: 48px;
    border-radius: 12px;
    background: #2563eb;
    color: white;
    border: none;
    font-weight: 700;
    font-size: 15px;
}

.stButton > button:hover {
    background: #1d4ed8;
    color: white;
}

.creator {
    text-align: center;
    color: #94a3b8;
    font-size: 11px;
    line-height: 1.6;
    margin-top: 25px;
}

@media (max-width: 650px) {

    .block-container {
        padding-top: 25px;
        padding-left: 18px;
        padding-right: 18px;
    }

    .judul {
        font-size: 32px;
    }

    .subjudul {
        font-size: 13px;
    }

    .stat-box {
        padding: 15px 5px;
    }

    .login-box {
        padding: 22px;
    }

}

</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="kurikulum">✦ KURIKULUM MERDEKA</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="judul">SIAP AJAR 22</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subjudul">'
    'Satu Portal, Solusi Lengkap 22 Perangkat Pembelajaran'
    '</div>',
    unsafe_allow_html=True
)

# =========================================================
# STATISTIK
# =========================================================

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(
        """
        <div class="stat-box">
            <div class="stat-number">22</div>
            <div class="stat-text">
                Perangkat Pembelajaran
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col2:
    st.markdown(
        """
        <div class="stat-box">
            <div class="stat-number">AI</div>
            <div class="stat-text">
                Google Gemini
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with col3:
    st.markdown(
        """
        <div class="stat-box">
            <div class="stat-number">DOCX</div>
            <div class="stat-text">
                Siap Diunduh
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================================================
# LOGIN CARD
# =========================================================

st.markdown(
    """
    <div class="login-box">
        <div class="login-title">
            🔐 Masuk ke Portal
        </div>

        <div class="login-text">
            Masukkan nama Anda dan Gemini API Key untuk
            menggunakan SIAP AJAR 22.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# FORM
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
# LINK API
# =========================================================

st.markdown(
    """
    <div class="api-box">
        🔑 Belum memiliki Gemini API Key?
        <a
            href="https://aistudio.google.com/app/apikey"
            target="_blank"
        >
            Buat API Key di Google AI Studio
        </a>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")

# =========================================================
# LOGIN
# =========================================================

if st.button(
    "🚀 MASUK KE PORTAL",
    use_container_width=True
):

    if not nama.strip():
        st.error(
            "❌ Silakan masukkan Nama Lengkap & Gelar."
        )

    elif not api_key.strip():
        st.error(
            "❌ Silakan masukkan Gemini API Key."
        )

    else:
        st.session_state.authenticated = True
        st.session_state.user_name = nama.strip()
        st.session_state.user_api_key = api_key.strip()

        st.rerun()

# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="creator">
        SIAP AJAR 22<br>
        Portal Administrasi Pembelajaran Guru<br>
        <b>© 2026 Pak Andar</b>
    </div>
    """,
    unsafe_allow_html=True
)
