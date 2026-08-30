import streamlit as st
import google.generativeai as genai

# 1. KONFIGURASI HALAMAN UTAMA & TEMA PREMIUM
st.set_page_config(
    page_title="Generator 22 Perangkat Kurikulum Merdeka",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# KUSTOMISASI CSS MODERN UNTUK ANTARMUKA PREMIUM
st.markdown("""
    <style>
    .stApp {
        background-color: #F8FAFC;
    }
    div[data-testid="stHeader"] {
        background-color: rgba(0,0,0,0);
    }
    h1 {
        color: #1E3A8A;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    h2, h3, h4 {
        color: #0F172A;
        font-weight: 700;
    }
    div.stButton > button:first-child {
        background-color: #1E3A8A;
        color: white;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.6rem 2rem;
        border: none;
        box-shadow: 0 4px 6px -1px rgba(30, 58, 138, 0.2);
        transition: all 0.2s;
    }
    div.stButton > button:first-child:hover {
        background-color: #1D4ED8;
        transform: translateY(-1px);
    }
    </style>
""", unsafe_allow_html=True)


# 2. STRUKTUR SIDEBAR (BILAH SAMPING)
with st.sidebar:
    logo_base64 = (
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAMAAABwKC9UAAAAeFBMVEX"
        "///8AAADFxcW9vb339/fOzs7V1dXm5ubg4ODGxsaqqqo5OTmUlJSIiIiwsLCYmJiAgIDX19fExM"
        "S0tLRERETu7u7e3t7W1tbOzs7AwMDAwMC4uLitra20tLR4eHhISEhYWFgQEBBQUFAtLS0pKSkhI"
        "SEwMDAEBAQHBwfpf9gBAAAECElEQVR4nO2b23KqMBSGBy6KAgpYvSAt93/FrS0g0mZ7gV6wZf6L"
        "mclMvshOAtv+9SclgW3/+pOSwLZ//UlJYNu//qQksKX0Yid0w2YfX0tB8mIvFOP8YitILhZ6oRJ"
        "7YVscBMmJhUKX0P9wDOR/9EInVEIhzrY4GBIInZALvVAIDw4C2v71JyWBbf/6k5LAtv++6S1F7I"
        "RcmIidshOyEXJhInbKTshGyIWJ2Ck7IRshFyZip+yEbITc6E1v90K7L8Z5N9sWuz0xzs0w3UuUf"
        "mB87Yh+YLzvbInmO7ofmO3OlmC+p4+g3O5fW+z2+U4v0TqXKD6UfeyE/CH037bFXv6Q9gPp97Yt"
        "9vInpA9kX+vN9sh+uVvs7U9mX+u97U9Osh/oK/t77pYob+WvO37bZfIqX9vxO93ZpXxtx+98Z5P"
        "ytR0/vS9fVvnafvqyytf205dbmby9X8vvpGvS9r6I3G87v2M3vC+e8j938v/v3bM75n3xd/vjS9"
        "wTf7cv9uNLLHNP/N2XpCeeuC/p6csW9wXuiydf4p54cl/S0xNvv9N98Xtfcl/S0/uS98XvfZV9S"
        "XtfpXpfeV+S9yVvX7Lel6z3Jet9Fd6XrPdVfF+F9yV9X8X39RdfmO5r9Isc8V7uF7mX+8Ww8Z7f"
        "Z89vG/GfF8PG27uU79+lfP/G7+P7N6T7+f49fH/F91d8X8f39fy+mbyX+8fEeym/v0reS/m/Zf"
        "JeNvy+pT0vxvdlfZ/F91V6X8X39Rdf6X0V+pXpfRfeN9w3/L6+v/N9S/qV+X1X9ivzfcNfP/6+S"
        "uqV7X2V9Cv79GvI9Cv9NfT+7Uv6W/bpt6Dpt9C3oOmv0bexfP8R37/j+3f6W6T6S6S6n+9eKup+"
        "enXfv7fvX+nffv9K/fb7Fervv6m/+0re6Sva6Sva6Sva6Sva6cs8P6CevvSnP/0ZTP/P/bN/pM"
        "K/vCHl/uUJKfcPT0i5f3dCyP0bAAn3jwAkfNf8gC6f+8fP6fKlf/w0Xbr0j59wS7p86R+/pcuT/"
        "vH77/Kkv8W+pEuXfvz69+nSjx9PSpf6/0t9SZf6S7p0qb+kS5f6S7p0qb+kS5f6S7p06Z8eZ+nS"
        "Px1m6dI/fXenS5f+6fE7Xbr0T3fXpEv/9Lid79b76ZseX9ulb3q8pEuhD+mD77EPrv6XfXD1v+"
        "yDK/ulD67slz6Ysl/4YMp688HEenOByXpzYf6CudD3XqLwGZTrxQXUeonCg7wX9H0vof8uOfj+Q"
        "UDbv/6kJLDtX39SEtj2rz8pCWz715+UBLb9609KAtv+9SclgW3/+pOSwLZ//UlJYNu//qQksO1"
        "ff1IS2PavPykJbPvXn5QEtv3rT0oC2/71JyWBbf/6k5LAtn/9SUlg27/+pCT8B6P12D/O1H2qA"
        "AAAAElFTkSuQmCC"
    )
    st.markdown(f"<div style='text-align: center;'><img src='{logo_base64}' width='100'></div>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #1E3A8A; margin-top: 5px; margin-bottom: 20px;'>Sistem Administrasi</h3>", unsafe_allow_html=True)
    st.write("---")
    
    api_key_input = st.text_input(
        "🔑 Masukkan Gemini API Key:",
        type="password",
        help="Gunakan API Key mandiri Anda untuk menghidupkan kecerdasan buatan secara gratis."
    )
    
    with st.expander("💡 Cara Dapatkan API Key Gratis"):
        st.markdown("""
        1. Kunjungi [Google AI Studio](https://google.com)
        2. Masuk dengan akun Gmail Anda.
        3. Tekan tombol **"Get API Key"**.
        4. Salin kode rahasianya dan tempel di atas.
        """)
        
    st.write("---")
    st.caption("🤖 **Versi 2.0.1 (Edisi Perbaikan)**")
    st.caption("Platform Otomasi Administrasi Kurikulum Merdeka © 2026")


# 3. AREA UTAMA HALAMAN APLIKASI
st.title("🤖 Generator Perangkat Pembelajaran Otomatis")
st.markdown("<p style='font-size: 1.15rem; color: #475569;'>Hasilkan 22 jenis berkas administrasi Kurikulum Merdeka dengan format dokumen resmi baku, lengkap dengan KOP Surat Instansi dan kolom tanda tangan struktural.</p>", unsafe_allow_html=True)

# TAHAP 1: FORM PENGISIAN DATA KOP SURAT
with st.expander("🏢 TAHAP 1: Kelola Data KOP Surat Instansi Sekolah", expanded=True):
    col_kop1, col_kop2 = st.columns(2)
    with col_kop1:
        nama_dinas = st.text_input("Nama Dinas Pendidikan Pembuat:", placeholder="Contoh: DINAS PENDIDIKAN PROVINSI JAWA BARAT")
        nama_sekolah = st.text_input("Nama Satuan Pendidikan / Sekolah:", placeholder="Contoh: SMA NEGERI 1 BANDUNG")
    with col_kop2:
        alamat_sekolah = st.text_input("Alamat Lengkap Sekolah & Kontak:", placeholder="Contoh: Jl. Belitung No.22, Telp: (022) 4232648, Email: info@sman1bdg.sch.id")

# TAHAP 2: FORM PENGISIAN DATA STAF PENDIDIK & PEJABAT
with st.expander("✍️ TAHAP 2: Kelola Data Pendidik & Penandatangan Dokumen", expanded=False):
    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        st.markdown("**Data Guru Pengampu**")
        nama_guru = st.text_input("Nama Lengkap Guru:", placeholder="Contoh: Ahmad Subarjo, S.Pd.")
        nip_guru = st.text_input("NIP/NUPTK Guru:", placeholder="Contoh: 198503122010011002")
    with col_g2:
        st.markdown("**Data Kepala Sekolah**")
        nama_kepsek = st.text_input("Nama Kepala Sekolah:", placeholder="Contoh: Dr. H. Supriatna, M.Pd.")
        nip_kepsek = st.text_input("NIP Kepala Sekolah:", placeholder="Contoh: 197305141998021001")
    with col_g3:
        st.markdown("**Data Pengawas Sekolah**")
        nama_pengawas = st.text_input("Nama Pengawas Pembina:", placeholder="Contoh: Dra. Hj. Endah Lestari, M.Si.")
        nip_pengawas = st.text_input("NIP Pengawas Sekolah:", placeholder="Contoh: 196811231993032003")

# TAHAP 3: PARAMETER MATA PELAJARAN & PILIHAN PERANGKAT
with st.container(border=True):
    st.markdown("<h4 style='color: #1E3A8A; margin-top: 0;'>📘 TAHAP 3: Pilih Parameter Mata Pelajaran & Perangkat</h4>", unsafe_allow_html=True)
    col_m1, col_m2 = st.columns(2)
    
    with col_m1:
        nama_mapel = st.text_input("Nama Mata Pelajaran:", placeholder="Contoh: Fisika, Matematika, Kimia")
        fase_kelas = st.selectbox(
            "Pilih Tingkatan Kelas / Fase:",
            [
                "Fase A - Kelas 1", "Fase A - Kelas 2", "Fase B - Kelas 3", "Fase B - Kelas 4",
                "Fase C - Kelas 5", "Fase C - Kelas 6", "Fase D - Kelas 7", "Fase D - Kelas 8",
                "Fase D - Kelas 9", "Fase E - Kelas 10", "Fase F - Kelas 11", "Fase F - Kelas 12"
            ]
        )
        semester = st.radio("Pilih Cakupan Semester:", ["Ganjil", "Genap", "Tahunan (Ganjil & Genap)"], horizontal=True)

    with col_m2:
        alokasi_waktu = st.text_input("Alokasi Waktu / Target JP:", placeholder="Contoh: 3 JP per Minggu atau 108 JP per Tahun")
        perangkat_pilihan = st.selectbox(
            "Jenis Perangkat Kurikulum Merdeka yang Ingin Diterbitkan:",
            [
                "1. Cover Perangkat Pembelajaran", "2. Kalender Pendidikan Sekolah (Format Analisis)",
                "3. Analisis Alokasi Waktu & Minggu Efektif", "4. Pemetaan Capaian Pembelajaran (CP) berdasarkan Elemen",
                "5. Perumusan Tujuan Pembelajaran (TP)", "6. Penyusunan Alur Tujuan Pembelajaran (ATP)",
                "7. Program Tahunan (Prota)", "8. Program Semester (Promes/Prosem)",
                "9. Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)", "10. Modul Ajar (MA) Utama / RPP Plus",
                "11. Bahan Ajar & Ringkasan Materi", "12. LKPD (Lembar Kerja Peserta Didik)",
                "13. Kisi-Kisi Asesmen (Formatif & Sumatif)", "14. Instrumen Asesmen Formatif (Rubrik & Catatan Anekdot)",
                "15. Soal Asesmen Sumatif (Pilihan Ganda & Esai + Kunci Jawaban)", "16. Modul Projek Penguatan Profil Pelajar Pancasila (P5)",
                "17. Jurnal Mengajar Harian Guru", "18. Daftar Absensi & Catatan Perkembangan Karakter Siswa",
                "19. Daftar Nilai Rapor Kurikulum Merdeka", "20. Program Remedial dan Pengayaan",
                "21. Analisis Hasil Asesmen / Evaluasi Pembelajaran", "22. Laporan Refleksi Guru dan Rencana Tindak Lanjut (RTL)"
            ]
        )

st.write("")
tombol_generate = st.button("🚀 Susun & Terbitkan Dokumen Sesuai Format Baku", use_container_width=True)


# 4. LOGIKA MESIN INTELLIGENCE & FORMULASI PROMPT BAKU
if tombol_generate:
    if not api_key_input:
        st.warning("⚠️ **Akses Ditolak:** Harap masukkan **Gemini API Key** Anda pada kolom bilah samping kiri.")
    elif not nama_mapel or not nama_sekolah:
        st.error("❌ **Gagal Memproses:** Kolom 'Nama Mata Pelajaran' dan 'Nama Sekolah' wajib diisi!")
    else:
        with st.spinner(f"⏳ AI sedang menyelaraskan struktur data dan menyusun dokumen '{perangkat_pilihan}'..."):
            try:
                # Amankan isian kosong agar tidak merusak formatting Python string
                dinas_val = nama_dinas if nama_dinas else "DINAS PENDIDIKAN KABUPATEN/KOTA/PROVINSI"
                alamat_val = alamat_sekolah if alamat_sekolah else "Jalan ... Telp ... Email ..."
                waktu_val = alokasi_waktu if alokasi_waktu else "Disesuaikan Ketentuan Standar Kurikulum"
                
                guru_val = nama_guru if nama_guru else "..........................."
                nip_g_val = nip_guru if nip_guru else "..........................."
                kepsek_val = nama_kepsek if nama_kepsek else "..........................."
                nip_k_val = nip_kepsek if nip_kepsek else "..........................."
                pengawas_val = nama_pengawas if nama_pengawas else "..........................."
