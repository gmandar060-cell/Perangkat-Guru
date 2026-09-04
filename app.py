import streamlit as st

# =========================================================
# SIAP AJAR 22 - Halaman Login Utama (app.py)
# =========================================================

st.set_page_config(
    page_title="SIAP AJAR 22",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# CSS LOGIN (Dibersihkan dari bug teks expander)
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
    padding-top: 3rem !important;
    padding-bottom: 3.5rem !important;
    max-width: 1100px;
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

/* Memperbaiki total bug teks arrow berlebih pada expander Streamlit */
details[data-testid="stExpander"] {
    border: 1px solid #E2E8F0 !important;
    border-radius: 10px !important;
    background-color: #FFFFFF !important;
    margin-bottom: 12px;
}

details[data-testid="stExpander"] summary {
    font-weight: 600 !important;
}

/* Menyembunyikan artefak teks ikon bawaan versi streamlit tertentu */
details[data-testid="stExpander"] summary span div p {
    display: inline-block !important;
}

/* Memastikan tinggi kartu login simetris */
div[data-testid="stHorizontalBlock"] > div {
    display: flex;
    flex-direction: column;
}

div[data-testid="stHorizontalBlock"] > div > div[data-testid="stVerticalBlock"] {
    height: 100%;
}

div[data-testid="stVerticalBlock"] > div[style*="border"] {
    background: #FFFFFF !important;
    border-radius: 16px !important;
    border: 1px solid #E2E8F0 !important;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
    padding: 32px 28px !important;
    margin-bottom: 12px;
    height: 100%;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
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

.footer-box {
    text-align: center;
    padding: 24px 10px 10px;
    color: #64748B !important;
    font-size: 12px;
    border-top: 1px solid #E2E8F0;
    margin-top: 40px;
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
                 alt="Logo SIAP AJAR 22"
                 style="filter:drop-shadow(0 4px 6px rgba(0,0,0,.1));">
        </div>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# SESSION STATE INISIALISASI
# =========================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "user_api_key" not in st.session_state:
    st.session_state.user_api_key = ""

# =========================================================
# ROUTING LOGIC (HALAMAN LOGIN VS DASHBOARD)
# =========================================================

if not st.session_state.authenticated:
    col_left, col_right = st.columns(2, gap="large")

    with col_left:
        with st.container(border=True):
            render_logo(85)
            st.markdown(
                "<p style='text-align:center;font-size:11px;font-weight:700;"
                "color:#2563EB;letter-spacing:.5px;margin-bottom:6px;'>"
                " SISTEM INFORMASI ASISTEN PERANGKAT AJAR</p>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<h2 style='text-align:center;font-weight:800;margin:0 0 4px 0;font-size:24px;'>"
                "SIAP AJAR 22</h2>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<p style='text-align:center;font-size:13px;font-style:italic;"
                "color:#64748B;margin-bottom:16px;'>“Satu Portal, Solusi Lengkap 22 Perangkat Pembelajaran”</p>",
                unsafe_allow_html=True,
            )
            st.divider()
            st.markdown("##### 📌 Keunggulan Platform")
            st.markdown(
                """
- 📋 **22 jenis perangkat pembelajaran lengkap**
- 🤖 **Penyusunan cerdas dengan Google Gemini AI**
- 📄 **Ekspor instan ke Microsoft Word (.docx)**
- 📝 **Cadangan teks mentah (.txt) praktis**
- 👁️ **Live preview dokumen gaya A4**
- 🔐 **Keamanan terjamin (API Key berbasis sesi browser)**
                """
            )
            st.caption("SIAP AJAR 22 • Creator: Andar")

    with col_right:
        with st.container(border=True):
            st.markdown("### 🔐 Akses Masuk Pendidik")
            st.write("")

            nama_guru_input = st.text_input(
                "Nama Lengkap & Gelar",
                placeholder="Contoh: Masukan Nama Lengkap Anda, S.Pd.",
            )

            api_key_masuk = st.text_input(
                "Gemini API Key Pribadi",
                type="password",
                placeholder="Masukan Gemini Api Key Anda",
            )

            with st.expander("📖 Cara mendapatkan Gemini API Key"):
                st.markdown(
                    """
1. Buka **Google AI Studio** di browser Anda.
2. Login menggunakan akun Google pribadi.
3. Klik tombol **Get API key**.
4. Buat atau pilih API key yang tersedia.
5. Salin dan tempelkan kunci tersebut ke kolom di atas.
                    """
                )

            st.write("")

            if st.button("MASUK KE PORTAL", use_container_width=True):
                if not nama_guru_input.strip():
                    st.warning("⚠️ Mohon isi Nama Lengkap & Gelar terlebih dahulu.")
                elif not api_key_masuk.strip():
                    st.warning("⚠️ Mohon isi Gemini API Key terlebih dahulu.")
                else:
                    st.session_state.user_name = nama_guru_input.strip()
                    st.session_state.user_api_key = api_key_masuk.strip()
                    st.session_state.authenticated = True
                    st.rerun()

    st.markdown(
        """
<div class="footer-box">
SIAP AJAR 22 • Creator: Andar<br>
© 2026 SIAP AJAR 22 • Engine AI Pembelajaran
</div>
        """,
        unsafe_allow_html=True,
    )
else:
    # Jika sudah berhasil masuk, panggil file dashboard.py
    import dashboard
    dashboard.run()
