import streamlit as st
import google.generativeai as genai

# 1. KONFIGURASI HALAMAN UTAMA
st.set_page_config(
    page_title="Generator 22 Perangkat Kurikulum Merdeka",
    page_icon="🤖",
    layout="wide"
)

# 2. STRUKTUR SIDEBAR (PENGATURAN & KEAMANAN API)
with st.sidebar:
    # Logo Tut Wuri Handayani untuk identitas pendidikan Indonesia
    st.image("https://wikimedia.org", width=100)
    st.title("Pengaturan Generator")
    st.write("---")
    
    # Kolom API Key Mandiri agar pemilik web (Anda) tidak bangkrut membayar token ribuan guru
    api_key_input = st.text_input(
        "Masukkan Gemini API Key Anda:",
        type="password",
        help="Dapatkan API Key gratis di Google AI Studio (://google.com) menggunakan akun Google Anda."
    )
    
    st.markdown("""
    **Cara Mendapatkan API Key Gratis:**
    1. Buka [Google AI Studio](https://://google.com/)
    2. Login dengan akun Google/Gmail Anda.
    3. Klik tombol **"Get API Key"**.
    4. Salin kodenya dan tempel di kolom atas.
    """)
    st.write("---")
    st.caption("Versi 1.1.0 (Skala Nasional) © 2026")

# 3. AREA FORMULIR UTAMA GURU
st.title("🤖 Generator 22 Perangkat Pembelajaran")
st.subheader("Kurikulum Merdeka — Otomatis & Standar Kemendikbudristek")
st.write("Sistem ini dirancang untuk membantu guru di seluruh Indonesia menyusun administrasi mengajar secara instan dan akurat.")

# Pembuatan Grid 2 Kolom untuk Form Input
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
    alokasi_waktu = st.text_input("4. Total Alokasi Waktu / JP:", placeholder="Contoh: 144 JP per tahun atau 4 JP per minggu")
    
    # 22 Daftar Perangkat Pembelajaran Sesuai Standar Nasional
    perangkat_pilihan = st.selectbox(
        "5. Pilih Jenis Perangkat yang Ingin Dibuat:",
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

st.write("---")

# Tombol Proses Dokumen
tombol_generate = st.button("🚀 Susun Perangkat Pembelajaran Sekarang", use_container_width=True)

# 4. PROSES EKSEKUSI OLEH MESIN AI
if tombol_generate:
    # Validasi Input Penting
    if not api_key_input:
        st.warning("⚠️ **Akses Ditolak:** Mohon masukkan **Gemini API Key** Anda di bilah samping (sidebar) terlebih dahulu untuk menghidupkan kecerdasan buatan.")
    elif not nama_mapel:
        st.error("❌ **Kesalahan Input:** Kolom 'Nama Mata Pelajaran' wajib diisi!")
    else:
        # Menampilkan animasi loading yang interaktif saat AI sedang berpikir
        with st.spinner(f"⏳ AI sedang menyusun dokumen **{perangkat_pilihan}**... Proses ini memakan waktu 10-30 detik. Mohon tidak menutup halaman ini."):
            try:
                # Inisialisasi API Key dari Input Guru
                genai.configure(api_key=api_key_input)
                
                # Menggunakan model Gemini 1.5 Flash yang sangat cepat dan mendukung teks panjang
                model = genai.GenerativeModel('gemini-3.5-flash')
                
                # Formula Rekayasa Perintah (Master Prompt Engineering) Otomatis
                master_prompt = f"""
                Bertindaklah sebagai Ahli Kurikulum Merdeka, Pengawas Sekolah Senior, dan AI Administrasi Guru Kemendikbudristek RI.
                
                Tugas Anda adalah membuat dokumen resmi untuk komponen perangkat pembelajaran berikut:
                - Jenis Perangkat: {perangkat_pilihan}
                - Mata Pelajaran: {nama_mapel}
                - Kelas & Fase: {fase_kelas}
                - Semester: {semester}
                - Estimasi Waktu/JP: {alokasi_waktu if alokasi_waktu else "Disesuaikan standar regulasi nasional"}
                
                Ketentuan Penulisan Dokumen:
                1. Buatlah isi dokumen ini secara LENGKAP, UTUH, dan DETAIL. Jangan gunakan singkatan seperti "...dan seterusnya" atau "dst". Guru membutuhkan dokumen yang siap pakai.
                2. Gunakan format tabel Markdown yang rapi jika dokumen tersebut melibatkan data terstruktur (seperti ATP, Prota, Promes, Kisi-kisi, Rubrik, Nilai, atau Pemetaan).
                3. Gunakan bahasa Indonesia baku, profesional, serta menggunakan kata kerja operasional berdasarkan Taksonomi Bloom yang direvisi.
                4. Pastikan struktur dokumen mengikuti regulasi standar Kurikulum Merdeka terbaru (mengandung elemen profil pelajar pancasila, asesmen formatif/sumatif, dan langkah pembelajaran yang berpusat pada murid jika relevan).
                
                Langsung berikan isi dokumen dari judul teratas tanpa ada kalimat pembuka basa-basi seperti "Berikut adalah dokumen yang Anda minta...".
                """
                
                # Mengirim perintah ke server Google AI
                response = model.generate_content(master_prompt)
                
                # Menampilkan Hasil Pembentukan Dokumen ke Layar
                st.success(f"✨ Dokumen **{perangkat_pilihan}** Berhasil Disusun!")
                st.write("---")
                
                # Menampilkan hasil teks dalam format Markdown yang rapi dan bisa diblok/disalin
                st.markdown(response.text)
                
                st.write("---")
                st.info("💡 **Tips untuk Guru:** Anda bisa memblok teks di atas, menyalinnya (*Copy*), lalu menempelkannya (*Paste*) langsung ke Microsoft Word atau Google Docs untuk dicetak.")
                
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan pada sistem AI. Pastikan API Key Anda valid. Detail Eror: {str(e)}")
