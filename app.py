import streamlit as st
import google.generativeai as genai

# 1. KONFIGURASI HALAMAN UTAMA (TAMPILAN MODERN)
st.set_page_config(
    page_title="Generator 22 Perangkat Kurikulum Merdeka",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# KUSTOMISASI MODEREN LEWAT CSS
st.markdown("""
    <style>
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1 {
        color: #1E3A8A;
        font-weight: 700;
    }
    h3 {
        color: #4B5563;
    }
    </style>
""", unsafe_allow_html=True)


# 2. STRUKTUR SIDEBAR (BILAH SAMPING YANG RAPI & BERSIH)
with st.sidebar:
    # PERBAIKAN LOGO: Menggunakan URL cadangan dari repositori aset publik yang stabil
    st.markdown("<div style='text-align: center;'><img src='https://githubusercontent.com' width='120'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin-top: 10px;'>Administrasi Guru</h2>", unsafe_allow_html=True)
    st.write("---")
    
    # Kolom API Key Keamanan Pengguna
    api_key_input = st.text_input(
        "🔑 Masukkan Gemini API Key Anda:",
        type="password",
        help="Dapatkan kunci akses gratis dari Google AI Studio untuk menghidupkan kecerdasan buatan."
    )
    
    # Memasukkan petunjuk ke dalam komponen lipat (Expander)
    with st.expander("💡 Cara Dapat API Key Gratis"):
        st.markdown("""
        1. Buka [Google AI Studio](https://google.com)
        2. Masuk dengan akun Gmail Anda.
        3. Klik tombol **"Get API Key"**.
        4. Klik **"Create API Key"**.
        5. Salin kodenya dan tempel di kolom atas.
        """)
        
    st.write("---")
    st.caption("🤖 **Versi 1.3.0 (Edisi Nasional)**")
    st.caption("Platform Otomasi Administrasi Kurikulum Merdeka © 2026")


# 3. AREA UTAMA (FORMULIR GURU DALAM KOTAK/CONTAINER)
st.title("🤖 Generator 22 Perangkat Pembelajaran")
st.subheader("Kurikulum Merdeka — Otomatis & Standar Kemendikbudristek")
st.write("Silakan isi komponen identitas sekolah dan pilih jenis administrasi yang ingin Anda buat di bawah ini secara teliti.")

# Mengelompokkan input ke dalam Wadah Komponen (Container)
with st.container(border=True):
    st.markdown("<h4 style='color: #1E3A8A; margin-top: 0;'>📋 Form Pengisian Identitas Pembelajaran</h4>", unsafe_allow_html=True)
    
    # Grid 2 Kolom untuk menyusun pertanyaan formulir
    col1, col2 = st.columns(2)
    
    with col1:
        nama_mapel = st.text_input("1. Nama Mata Pelajaran:", placeholder="Contoh: Matematika, Bahasa Indonesia, IPAS")
        fase_kelas = st.selectbox(
            "2. Pilih Kelas & Fase:",
            [
                "Fase A - Kelas 1", "Fase A - Kelas 2",
                "Fase B - Kelas 3", "Fase B - Kelas 4",
                "Fase C - Kelas 5", "Fase C - Kelas 6",
                "Fase D - Kelas 7", "Fase D - Kelas 8", "Fase D - Kelas 9",
                "Fase E - Kelas 10",
                "Fase F - Kelas 11", "Fase F - Kelas 12"
            ]
        )
        semester = st.radio("3. Pilih Semester:", ["Ganjil", "Genap", "Tahunan (Ganjil & Genap)"], horizontal=True)

    with col2:
        alokasi_waktu = st.text_input("4. Total Alokasi Waktu / JP Pembelajaran:", placeholder="Contoh: 144 JP per tahun atau 4 JP per minggu")
        perangkat_pilihan = st.selectbox(
            "5. Pilih Jenis Perangkat Kurikulum Merdeka:",
            [
                "1. Cover Perangkat Pembelajaran",
                "2. Kalender Pendidikan Sekolah (Format Analisis)",
                "3. Analisis Alokasi Waktu & Minggu Efektif",
                "4. Pemetaan Capaian Pembelajaran (CP) berdasarkan Elemen",
                "5. Perumusan Tujuan Pembelajaran (TP)",
                "6. Penyusunan Alur Tujuan Pembelajaran (ATP)",
                "7. Program Tahunan (Prota)",
                "8. Program Semester (Promes/Prosem)",
                "9. Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)",
                "10. Modul Ajar (MA) Utama / RPP Plus",
                "11. Bahan Ajar & Ringkasan Materi",
                "12. LKPD (Lembar Kerja Peserta Didik)",
                "13. Kisi-Kisi Asesmen (Formatif & Sumatif)",
                "14. Instrumen Asesmen Formatif (Rubrik & Catatan Anekdot)",
                "15. Soal Asesmen Sumatif (Pilihan Ganda & Esai + Kunci Jawaban)",
                "16. Modul Projek Penguatan Profil Pelajar Pancasila (P5)",
                "17. Jurnal Mengajar Harian Guru",
                "18. Daftar Absensi & Catatan Perkembangan Karakter Siswa",
                "19. Daftar Nilai Rapor Kurikulum Merdeka",
                "20. Program Remedial dan Pengayaan",
                "21. Analisis Hasil Asesmen / Evaluasi Pembelajaran",
                "22. Laporan Refleksi Guru dan Rencana Tindak Lanjut (RTL)"
            ]
        )

st.write("") # Jarak spasi

# Tombol Eksekusi
tombol_generate = st.button("🚀 Susun & Terbitkan Perangkat Pembelajaran Sekarang", use_container_width=True)


# 4. PROSES LOGIKA MESIN AI & FITUR UNDUH
if tombol_generate:
    if not api_key_input:
        st.warning("⚠️ **Akses Ditolak:** Silakan masukkan **Gemini API Key** Anda pada kolom di bilah samping (*sidebar*) kiri untuk mengaktifkan fitur kecerdasan buatan.")
    elif not nama_mapel:
        st.error("❌ **Kesalahan Input:** Kolom Nama Mata Pelajaran tidak boleh kosong!")
    else:
        with st.spinner(f"⏳ AI sedang memproses dokumen '{perangkat_pilihan}'... Mohon tunggu 10-30 detik."):
            try:
                # Konfigurasi token API
                genai.configure(api_key=api_key_input)
                model = genai.GenerativeModel('gemini-3.5-flash')
                
                master_prompt = f"""
                Bertindaklah sebagai Ahli Kurikulum Merdeka, Pengawas Sekolah Senior, dan AI Administrasi Guru Kemendikbudristek RI.
                
                Tugas Anda adalah membuat dokumen resmi untuk komponen perangkat pembelajaran berikut:
                - Jenis Perangkat: {perangkat_pilihan}
                - Mata Pelajaran: {nama_mapel}
                - Kelas & Fase: {fase_kelas}
                - Semester: {semester}
                - Estimasi Waktu/JP: {alokasi_waktu if alokasi_waktu else "Disesuaikan standar regulasi nasional"}
                
                Ketentuan Penulisan Dokumen:
                1. Buatlah isi dokumen ini secara LENGKAP, UTUH, dan DETAIL. Jangan gunakan singkatan seperti "...dan seterusnya" atau "dst".
                2. Gunakan format tabel Markdown yang rapi jika dokumen tersebut melibatkan data terstruktur.
                3. Gunakan bahasa Indonesia baku, profesional, serta menggunakan kata kerja operasional berdasarkan Taksonomi Bloom yang direvisi.
                4. Pastikan struktur dokumen mengikuti regulasi standar Kurikulum Merdeka terbaru.
                
                Langsung berikan isi dokumen dari judul teratas tanpa ada kalimat pembuka basa-basi.
                """
                
                response = model.generate_content(master_prompt)
                hasil_teks = response.text
                
                # Menampilkan lembar hasil dokumen di layar aplikasi
                st.success(f"✨ Dokumen '{perangkat_pilihan}' Berhasil Disusun!")
                
                with st.container(border=True):
                    st.markdown(hasil_teks)
                
                st.write("---")
                
                # PERBAIKAN TOMBOL DOWNLOAD: Menyediakan file teks mentah (.txt) yang otomatis bisa dibuka rapi di Microsoft Word
                nama_file_unduh = f"{perangkat_pilihan.replace(' ', '_')}_{nama_mapel}_{fase_kelas.replace(' ', '')}.txt"
                
                st.download_button(
                    label="📥 Unduh Dokumen ke Laptop (Format Word/Text)",
                    data=hasil_teks,
                    file_name=nama_file_unduh,
                    mime="text/plain",
                    use_container_width=True,
                    help="Klik untuk langsung mengunduh hasil dokumen ini. File .txt ini dapat langsung dibuka atau di-copy dengan rapi ke Microsoft Word."
                )
                
            except Exception as e:
                st.error(f"❌ Terjadi gangguan komunikasi dengan server Google AI. Pastikan API Key Anda aktif dan benar. Detail Eror: {str(e)}")
