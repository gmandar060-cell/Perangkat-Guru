import streamlit as st
from google import genai
from google.genai import types
from datetime import datetime
import re

# ==========================================
# 1. KONFIGURASI HALAMAN & INJEKSI CSS
# ==========================================
st.set_page_config(
    page_title="Generator 22 Perangkat Kurikulum Merdeka",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling Modern UI
st.markdown("""
<style>
    .stApp {
        background-color: #F8FAFC;
        font-family: 'Inter', sans-serif;
    }
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        padding: 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        color: white;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .main-header p {
        color: #E2E8F0;
        margin: 5px 0 0 0;
        font-size: 14px;
    }
    .stButton>button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 10px 24px;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
        color: white;
        box-shadow: 0 4px 8px rgba(37, 99, 235, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# Logo Tut Wuri Handayani (SVG Base64 Embedded Vector Data)
LOGO_TUT_WURI_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="80" height="80" style="display: block; margin: 0 auto 15px auto;">
  <polygon points="50,5 95,38 78,92 22,92 5,38" fill="#0284C7" stroke="#0369A1" stroke-width="2"/>
  <path d="M50,15 L70,80 L50,65 L30,80 Z" fill="#FACC15"/>
  <circle cx="50" cy="40" r="10" fill="#DC2626"/>
  <path d="M25,50 Q50,70 75,50 Q50,90 25,50" fill="#FFFFFF" opacity="0.9"/>
</svg>
"""

# Inisialisasi Session State
if "hasil_teks" not in st.session_state:
    st.session_state.hasil_teks = ""
if "nama_file" not in st.session_state:
    st.session_state.nama_file = "perangkat_kurikulum_merdeka.txt"


# ==========================================
# 2. FUNGSI MASTER PROMPT ENGINEERING
# ==========================================
def buat_instruksi_prompt(data: dict) -> str:
    """
    Merakit prompt terisolasi untuk dikirimkan ke model Gemini.
    Menghindari syntax error dan menjaga kebakuan format dokumen dinas.
    """
    prompt = f"""
Bertindaklah sebagai Ahli Kurikulum Merdeka Nasional, Lembaga Penjaminan Mutu Pendidikan (LPMP), Pengawas Sekolah Senior, dan AI Administrasi Sekolah Kemendikbudristek RI.

Tugas Anda adalah menerbitkan dokumen resmi perangkat pembelajaran yang LENGKAP, DETAIL, OPERASIONAL, dan SISTEMATIS sesuai regulasi Kurikulum Merdeka terbaru.

=== IDENTITAS SATUAN PENDIDIKAN ===
- Dinas Pendidikan : {data['dinas']}
- Satuan Pendidikan: {data['sekolah']}
- Alamat & Kontak  : {data['alamat']}

=== IDENTITAS STAF & STRUKTURAL ===
- Guru Mata Pelajaran : {data['guru_nama']} (NIP: {data['guru_nip']})
- Kepala Sekolah      : {data['ks_nama']} (NIP: {data['ks_nip']})
- Pengawas Sekolah    : {data['pengawas_nama']} (NIP: {data['pengawas_nip']})

=== PARAMETER PEMBELAJARAN ===
- Jenis Dokumen    : {data['jenis_perangkat']}
- Mata Pelajaran   : {data['mapel']}
- Tingkatan / Fase : {data['fase_kelas']}
- Semester         : {data['semester']}
- Alokasi Waktu    : {data['alokasi_waktu']}
- Tanggal Terbit   : {data['tanggal_hari_ini']}
- Kota Instansi    : {data['kota_sekolah']}

=== KETENTUAN FORMAT DOKUMEN (FORMAT BAKU RESMI INDONESIA) ===
1. BAGIAN PALING ATAS: Tulis KOP SURAT RESMI SEKOLAH huruf kapital terpusat rapi, mencakup nama Dinas Pendidikan, Satuan Pendidikan, dan Alamat lengkap, diikuti garis pembatas horizontal '---'.
2. BAGIAN KEDUA: Judul Dokumen resmi ({data['jenis_perangkat'].upper()}) diikuti TABEL IDENTITAS (Mata Pelajaran, Fase/Kelas, Semester, Tahun Pelajaran, Alokasi Waktu).
3. BAGIAN KETIGA (ISI DOKUMEN UTAMA):
   - Tulis secara UTUH, SANGAT DETAIL, dan OPERASIONAL (Gunakan kata kerja operasional Taksonomi Bloom revisi).
   - DILARANG KERAS memotong materi atau menggunakan kata-kata singkatan seperti '...dst', '[lanjutkan]', '[sesuaikan]', atau 'dan seterusnya'.
   - Jika dokumen memerlukan tabel (misal ATP, Prota, Promes, KKTP, Kisi-kisi, Modul Ajar, Rubrik, Agenda Guru), WAJIB disajikan menggunakan format TABEL MARKDOWN yang rapi, padat, dan jelas.
4. BAGIAN PALING BAWAH (LEMBAR PENGESAHAN):
   Buat lembar tanda tangan 3 kolom horizontal berdampingan dengan format markdown table yang rapi:
   | Mengetahui,<br>Pengawas Pembina | Mengetahui,<br>Kepala Sekolah | {data['kota_sekolah']}, {data['tanggal_hari_ini']}<br>Guru Mata Pelajaran |
   | :---: | :---: | :---: |
   | <br><br><br><br> | <br><br><br><br> | <br><br><br><br> |
   | **{data['pengawas_nama']}**<br>NIP. {data['pengawas_nip']} | **{data['ks_nama']}**<br>NIP. {data['ks_nip']} | **{data['guru_nama']}**<br>NIP. {data['guru_nip']} |

5. ATURAN PENULISAN OUTPUT:
   - Langsung mulai dari baris KOP SURAT pertama tanpa ada teks pembuka seperti 'Tentu, ini dokumen Anda...', 'Berikut adalah...', atau salam robot AI lainnya.
   - Akhiri langsung setelah tabel lembar pengesahan tanpa kalimat penutup.
"""
    return prompt.strip()


# ==========================================
# 3. SIDEBAR (API KEY & PANDUAN)
# ==========================================
with st.sidebar:
    st.markdown(LOGO_TUT_WURI_SVG, unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; margin-bottom: 2px;'>ADMINISTRASI GURU</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 12px; color: #64748B; margin-top: 0;'>Standar Kurikulum Merdeka Kemendikbudristek</p>", unsafe_allow_html=True)
    st.divider()

    st.subheader("🔑 Konfigurasi Akses AI")
    api_key_input = st.text_input(
        "Masukkan Gemini API Key:",
        type="password",
        placeholder="AIzaSy...",
        help="API Key Anda dijamin aman dan hanya tersimpan pada sesi peramban saat ini."
    )

    with st.expander("📖 Panduan Mendapatkan API Key"):
        st.markdown("""
        1. Buka portal [Google AI Studio](https://aistudio.google.com/).
        2. Masuk menggunakan akun Google Anda.
        3. Klik tombol **Get API Key** lalu pilih **Create API Key**.
        4. Salin kode API Key yang didapatkan dan tempelkan pada kolom di atas.
        *Layanan API gratis tersedia dengan kuota harian yang cukup untuk pembuatan perangkat ajar.*
        """)

    st.divider()
    st.info("💡 **Tips:** Pastikan seluruh isian formulir pada Tahap 1, 2, dan 3 telah terisi lengkap sebelum menekan tombol Generate.")


# ==========================================
# 4. HALAMAN UTAMA (3 TAHAPAN FORM)
# ==========================================
st.markdown("""
<div class="main-header">
    <h1>🏛️ Generator 22 Perangkat Pembelajaran Kurikulum Merdeka</h1>
    <p>Penerbitan Dokumen Administrasi Guru & Perangkat Pembelajaran Resmi Berstandar Nasional Berbasis AI</p>
</div>
""", unsafe_allow_html=True)

# TAHAP 1: KOP SURAT INSTANSI (Expander Default Terbuka)
with st.expander("📌 TAHAP 1: Data KOP Surat Instansi Pendidikan", expanded=True):
    col_kop1, col_kop2 = st.columns(2)
    with col_kop1:
        dinas_pendidikan = st.text_input(
            "Nama Dinas Pendidikan Pembina:",
            value="DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT",
            placeholder="Contoh: DINAS PENDIDIKAN DAN KEBUDAYAAN KABUPATEN..."
        )
        nama_sekolah = st.text_input(
            "Nama Satuan Pendidikan (Sekolah):",
            value="SMAS NUSA HARAPAN",
            placeholder="Contoh: SMA NEGERI 1 PONTIANAK"
        )
    with col_kop2:
        alamat_sekolah = st.text_input(
            "Alamat Lengkap & Kontak Satuan Pendidikan:",
            value="Jl. Pancasila No. 10, Telp. (0561) 734567, Pontianak",
            placeholder="Contoh: Jl. Merdeka No. 45..."
        )
        kota_sekolah = st.text_input(
            "Kota / Kabupaten Domisili Instansi:",
            value="Pontianak",
            placeholder="Contoh: Pontianak"
        )

# TAHAP 2: DATA STAF & STRUKTURAL (Expander Default Tertutup)
with st.expander("👥 TAHAP 2: Data Pendidik & Pejabat Penandatangan", expanded=False):
    col_staf1, col_staf2, col_staf3 = st.columns(3)
    with col_staf1:
        st.markdown("**Guru Mata Pelajaran**")
        guru_nama = st.text_input("Nama Guru & Gelar:", value="MUHAMMAD NURZULIANDAR, S.Pd.")
        guru_nip = st.text_input("NIP Guru (Isi '-' jika Non-PNS):", value="-")
    
    with col_staf2:
        st.markdown("**Kepala Sekolah**")
        ks_nama = st.text_input("Nama Kepala Sekolah & Gelar:", value="ZULKIFLI, S.Pd.")
        ks_nip = st.text_input("NIP Kepala Sekolah:", value="197508122005011004")
        
    with col_staf3:
        st.markdown("**Pengawas Sekolah**")
        pengawas_nama = st.text_input("Nama Pengawas Pembina:", value="NURHASANAH, M.Si.")
        pengawas_nip = st.text_input("NIP Pengawas Sekolah:", value="196811231993032003")

# TAHAP 3: PARAMETER PEMBELAJARAN (Container Berbingkai)
with st.container(border=True):
    st.markdown("### ⚙️ TAHAP 3: Parameter Kurikulum & Pilihan Dokumen")
    
    col_param1, col_param2 = st.columns(2)
    with col_param1:
        mapel = st.text_input("Nama Mata Pelajaran:", value="Fisika", placeholder="Contoh: Fisika, Bahasa Indonesia, dsb.")
        fase_kelas = st.selectbox(
            "Tingkatan Kelas / Fase Kurikulum Merdeka:",
            options=[
                "Fase A - Kelas 1 SD/MI",
                "Fase A - Kelas 2 SD/MI",
                "Fase B - Kelas 3 SD/MI",
                "Fase B - Kelas 4 SD/MI",
                "Fase C - Kelas 5 SD/MI",
                "Fase C - Kelas 6 SD/MI",
                "Fase D - Kelas 7 SMP/MTs",
                "Fase D - Kelas 8 SMP/MTs",
                "Fase D - Kelas 9 SMP/MTs",
                "Fase E - Kelas 10 SMA/MA/SMK",
                "Fase F - Kelas 11 SMA/MA/SMK",
                "Fase F - Kelas 12 SMA/MA/SMK"
            ],
            index=9
        )
        semester = st.radio(
            "Cakupan Semester:",
            options=["Semester Ganjil", "Semester Genap", "Semester Ganjil & Genap (1 Tahun Penuh)"],
            horizontal=True
        )

    with col_param2:
        alokasi_waktu = st.text_input("Alokasi Waktu Pembelajaran:", value="3 JP / Minggu (Total 54 JP per Semester)")
        
        daftar_22_perangkat = [
            "01. Analisis Capaian Pembelajaran (CP) & Pemetaan Elemen",
            "02. Alur Tujuan Pembelajaran (ATP) Lengkap",
            "03. Program Tahunan (PROTA)",
            "04. Program Semester (PROMES)",
            "05. Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)",
            "06. Modul Ajar Lengkap (Format Baku Kurikulum Merdeka)",
            "07. Lembar Kerja Peserta Didik (LKPD) Berdiferensiasi",
            "08. Modul / Panduan Projek Penguatan Profil Pelajar Pancasila (P5)",
            "09. Jurnal Mengajar Harian & Agenda Guru Terstruktur",
            "10. Format & Kisi-kisi Asesmen Diagnostik (Kognitif & Non-Kognitif)",
            "11. Kisi-kisi & Instrumen Asesmen Formatif",
            "12. Kisi-kisi, Naskah Soal, & Kunci Jawaban Asesmen Sumatif",
            "13. Rubrik & Format Penilaian Kinerja, Portofolio, serta Proyek",
            "14. Rekapitulasi Buku / Daftar Nilai Siswa Kurikulum Merdeka",
            "15. Buku Presensi / Lembar Absensi Peserta Didik",
            "16. Format Program Pembelajaran Remedial & Pengayaan",
            "17. Distribusi Sumber Belajar & Buku Teks Utama/Pendamping",
            "18. Analisis Alokasi Waktu Efektif (Rincian Minggu Efektif)",
            "19. Format Analisis Kuantitatif Butir Soal Evaluasi",
            "20. Jurnal Sikap & Catatan Karakter Dimensi Profil Pancasila",
            "21. Panduan Layanan Bimbingan Belajar & Konsultasi Akademik Siswa",
            "22. Format Laporan Evaluasi Diri Guru & Rencana Tindak Lanjut (RTL)"
        ]
        
        jenis_perangkat = st.selectbox(
            "Pilih 1 dari 22 Dokumen yang Ingin Diterbitkan:",
            options=daftar_22_perangkat,
            index=1
        )

    st.markdown("<br>", unsafe_allow_html=True)
    tombol_proses = st.button("🚀 Terbitkan Dokumen Administrasi Resmi", use_container_width=True)


# ==========================================
# 5. LOGIKA GENERASI & PEMANGGILAN API
# ==========================================
if tombol_proses:
    if not api_key_input.strip():
        st.error("⚠️ Silakan masukkan Google Gemini API Key Anda terlebih dahulu pada bilah samping (Sidebar)!")
    elif not mapel.strip() or not dinas_pendidikan.strip() or not nama_sekolah.strip():
        st.warning("⚠️ Mohon lengkapi seluruh isian form instansi dan nama mata pelajaran.")
    else:
        tanggal_sekarang = datetime.now().strftime("%d %B %Y")
        data_input = {
            "dinas": dinas_pendidikan,
            "sekolah": nama_sekolah,
            "alamat": alamat_sekolah,
            "kota_sekolah": kota_sekolah,
            "guru_nama": guru_nama,
            "guru_nip": guru_nip,
            "ks_nama": ks_nama,
            "ks_nip": ks_nip,
            "pengawas_nama": pengawas_nama,
            "pengawas_nip": pengawas_nip,
            "mapel": mapel,
            "fase_kelas": fase_kelas,
            "semester": semester,
            "alokasi_waktu": alokasi_waktu,
            "jenis_perangkat": jenis_perangkat,
            "tanggal_hari_ini": tanggal_sekarang
        }

        prompt_final = buat_instruksi_prompt(data_input)

        with st.spinner("🤖 Mengonstruksi dokumen baku Kurikulum Merdeka... Mohon tunggu sejenak."):
            try:
                # Inisialisasi klien SDK GenAI resmi dengan model gemini-2.5-flash
                client = genai.Client(api_key=api_key_input.strip())
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt_final,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                    )
                )

                # Simpan ke session state
                st.session_state.hasil_teks = response.text
                
                # Format nama file unduhan yang bersih
                nama_file_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', f"{jenis_perangkat}_{mapel}_{fase_kelas}")
                st.session_state.nama_file = f"{nama_file_clean}.txt"
                
                st.success("✅ Berkas administrasi berhasil diterbitkan dengan format baku resmi!")

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan saat memproses dokumen: {str(e)}")


# ==========================================
# 6. PENAMPIL HASIL & TOMBOL EKSPOR DOKUMEN
# ==========================================
if st.session_state.hasil_teks:
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📄 Lembar Hasil Dokumen Administrasi")
    
    with st.container(border=True):
        st.markdown(st.session_state.hasil_teks)

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
    with col_dl2:
        st.download_button(
            label="💾 Unduh Naskah Dokumen (.TXT / Siap Salin ke Word)",
            data=st.session_state.hasil_teks,
            file_name=st.session_state.nama_file,
            mime="text/plain; charset=utf-8",
            use_container_width=True
        )
