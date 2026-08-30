import streamlit as st
from google import genai
from google.genai import types
from datetime import datetime
import re

# ==========================================
# 1. KONFIGURASI HALAMAN & INJEKSI CSS MODERN
# ==========================================
st.set_page_config(
    page_title="Kurikulum Merdeka Studio | AI Educator Suite",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background: linear-gradient(180deg, #F1F5F9 0%, #F8FAFC 100%);
    }

    /* Hero Banner Header */
    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 20px;
        padding: 36px 40px;
        color: white;
        margin-bottom: 28px;
        box-shadow: 0 20px 25px -5px rgba(15, 23, 42, 0.15), 0 8px 10px -6px rgba(15, 23, 42, 0.1);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .hero-container::after {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 300px;
        height: 300px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.3) 0%, rgba(0,0,0,0) 70%);
        border-radius: 50%;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(8px);
        padding: 6px 14px;
        border-radius: 30px;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 12px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    .hero-title {
        font-size: 32px;
        font-weight: 800;
        margin: 0 0 8px 0;
        letter-spacing: -0.5px;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 15px;
        color: #94A3B8;
        max-width: 680px;
        margin: 0;
        line-height: 1.6;
    }

    /* Modern Card & Containers */
    [data-testid="stExpander"] {
        background: #FFFFFF !important;
        border-radius: 16px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        margin-bottom: 16px;
    }
    
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background: #FFFFFF !important;
        border-radius: 16px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        padding: 24px !important;
    }

    /* Button Primary */
    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 28px !important;
        box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 15px 20px -3px rgba(37, 99, 235, 0.4) !important;
    }

    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 12px 24px !important;
        box-shadow: 0 10px 15px -3px rgba(5, 150, 105, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 15px 20px -3px rgba(5, 150, 105, 0.4) !important;
    }

    /* Sidebar Customization */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

LOGO_TUT_WURI_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="76" height="76" style="display: block; margin: 0 auto 12px auto;">
  <polygon points="50,5 95,38 78,92 22,92 5,38" fill="#0284C7" stroke="#0369A1" stroke-width="2"/>
  <path d="M50,15 L70,80 L50,65 L30,80 Z" fill="#FACC15"/>
  <circle cx="50" cy="40" r="10" fill="#DC2626"/>
  <path d="M25,50 Q50,70 75,50 Q50,90 25,50" fill="#FFFFFF" opacity="0.9"/>
</svg>
"""

if "hasil_teks" not in st.session_state:
    st.session_state.hasil_teks = ""
if "nama_file" not in st.session_state:
    st.session_state.nama_file = "perangkat_kurikulum_merdeka.txt"


# ==========================================
# 2. FUNGSI LOGIKA PROMPT
# ==========================================
def buat_instruksi_prompt(data: dict) -> str:
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
# 3. SIDEBAR (KEY & STATUS)
# ==========================================
with st.sidebar:
    st.markdown(LOGO_TUT_WURI_SVG, unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; margin: 0; font-size: 18px; font-weight: 700; color: #0F172A;'>STUDIO ADMINISTRASI</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 12px; color: #64748B; margin-top: 4px;'>Kurikulum Merdeka Kemendikbudristek</p>", unsafe_allow_html=True)
    st.divider()

    st.markdown("#### 🔐 Autentikasi AI")
    api_key_input = st.text_input(
        "Gemini API Key:",
        type="password",
        placeholder="AIzaSy...",
        help="API Key tersimpan secara lokal dan privat pada sesi peramban Anda."
    )

    with st.expander("❓ Cara Dapatkan API Key"):
        st.markdown("""
        1. Kunjungi [Google AI Studio](https://aistudio.google.com/).
        2. Masuk menggunakan akun Google.
        3. Klik **Get API Key** ➔ **Create API Key**.
        4. Salin dan tempelkan kuncinya ke kolom di atas.
        """)

    st.divider()
    st.markdown("""
    <div style="background-color: #F8FAFC; border-radius: 10px; padding: 12px; border: 1px solid #E2E8F0; font-size: 12px; color: #475569;">
        <strong>🚀 Engine Version:</strong><br>
        Google Gemini 2.5 Flash<br>
        Format Baku Standar LPMP/BSKAP
    </div>
    """, unsafe_allow_html=True)


# ==========================================
# 4. HALAMAN UTAMA (DASHBOARD MODERN)
# ==========================================
st.markdown("""
<div class="hero-container">
    <div class="hero-badge">⚡ Professional AI Suite for Educators</div>
    <div class="hero-title">Generator 22 Perangkat Pembelajaran</div>
    <div class="hero-subtitle">Susun perangkat pembelajaran baku, operasional, dan siap cetak dengan standar akreditasi dan supervisi pendidikan nasional.</div>
</div>
""", unsafe_allow_html=True)

# Tahap 1 & 2 dalam 2 Tab Modern
tab1, tab2 = st.tabs(["🏛️ 1. Identitas Satuan Pendidikan", "✍️ 2. Pejabat & Penandatangan"])

with tab1:
    col_kop1, col_kop2 = st.columns(2)
    with col_kop1:
        dinas_pendidikan = st.text_input(
            "Dinas Pendidikan Pembina:",
            value="DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT"
        )
        nama_sekolah = st.text_input(
            "Nama Satuan Pendidikan:",
            value="SMAS NUSA HARAPAN"
        )
    with col_kop2:
        alamat_sekolah = st.text_input(
            "Alamat & Kontak Sekolah:",
            value="Jl. Pancasila No. 10, Telp. (0561) 734567, Pontianak"
        )
        kota_sekolah = st.text_input(
            "Kota / Kabupaten Instansi:",
            value="Pontianak"
        )

with tab2:
    col_staf1, col_staf2, col_staf3 = st.columns(3)
    with col_staf1:
        st.markdown("**Guru Mata Pelajaran**")
        guru_nama = st.text_input("Nama Lengkap & Gelar:", value="MUHAMMAD NURZULIANDAR, S.Pd.", key="g_nama")
        guru_nip = st.text_input("NIP (Isi '-' jika Non-PNS):", value="-", key="g_nip")
    
    with col_staf2:
        st.markdown("**Kepala Sekolah**")
        ks_nama = st.text_input("Nama Kepala Sekolah:", value="ZULKIFLI, S.Pd.", key="ks_nama")
        ks_nip = st.text_input("NIP Kepala Sekolah:", value="197508122005011004", key="ks_nip")
        
    with col_staf3:
        st.markdown("**Pengawas Pembina**")
        pengawas_nama = st.text_input("Nama Pengawas Sekolah:", value="NURHASANAH, M.Si.", key="p_nama")
        pengawas_nip = st.text_input("NIP Pengawas:", value="196811231993032003", key="p_nip")

st.markdown("<br>", unsafe_allow_html=True)

# Tahap 3: Parameter Kurikulum
with st.container(border=True):
    st.markdown("### 📋 3. Konfigurasi Dokumen & Kurikulum")
    
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mapel = st.text_input("Mata Pelajaran:", value="Fisika", placeholder="Contoh: Fisika, Biologi, Ekonomi")
        fase_kelas = st.selectbox(
            "Tingkatan Kelas / Fase:",
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
            "Semester Berjalan:",
            options=["Semester Ganjil", "Semester Genap", "Semester Ganjil & Genap (1 Tahun)"],
            horizontal=True
        )

    with col_p2:
        alokasi_waktu = st.text_input("Alokasi Waktu / Target JP:", value="3 JP / Minggu (Total 54 JP per Semester)")
        
        daftar_22_perangkat = [
            "01. Analisis Capaian Pembelajaran (CP) & Pemetaan Elemen",
            "02. Alur Tujuan Pembelajaran (ATP) Lengkap",
            "03. Program Tahunan (PROTA)",
            "04. Program Semester (PROMES)",
            "05. Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)",
            "06. Modul Ajar Lengkap (Format Baku Kurikulum Merdeka)",
            "07. Lembar Kerja Peserta Didik (LKPD) Berdiferensiasi",
            "08. Modul / Panduan Projek Profil Pancasila (P5)",
            "09. Jurnal Mengajar Harian & Agenda Guru Terstruktur",
            "10. Format & Kisi-kisi Asesmen Diagnostik (Kognitif/Non-Kognitif)",
            "11. Kisi-kisi & Instrumen Asesmen Formatif",
            "12. Kisi-kisi, Naskah Soal, & Kunci Jawaban Asesmen Sumatif",
            "13. Rubrik Penilaian Kinerja, Portofolio, serta Proyek",
            "14. Rekapitulasi Daftar Nilai Kurikulum Merdeka",
            "15. Buku Presensi / Lembar Absensi Siswa",
            "16. Format Program Pembelajaran Remedial & Pengayaan",
            "17. Distribusi Sumber Belajar & Buku Teks Pembelajaran",
            "18. Analisis Alokasi Waktu Efektif (Rincian Pekan Efektif)",
            "19. Format Analisis Kuantitatif Butir Soal Evaluasi",
            "20. Jurnal Sikap & Catatan Dimensi Profil Pelajar Pancasila",
            "21. Panduan Layanan Bimbingan & Konsultasi Akademik",
            "22. Format Laporan Evaluasi Diri Guru & Rencana Tindak Lanjut"
        ]
        
        jenis_perangkat = st.selectbox(
            "Pilih Dokumen Perangkat yang Diterbitkan:",
            options=daftar_22_perangkat,
            index=1
        )

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    tombol_proses = st.button("✨ Terbitkan Dokumen Resmi Sekarang", use_container_width=True)


# ==========================================
# 5. PEMROSESAN GENERATIF
# ==========================================
if tombol_proses:
    if not api_key_input.strip():
        st.error("🔑 Harap masukkan Gemini API Key pada bilah samping (Sidebar) terlebih dahulu.")
    elif not mapel.strip() or not dinas_pendidikan.strip() or not nama_sekolah.strip():
        st.warning("⚠️ Mohon pastikan data instansi dan mata pelajaran telah diisi dengan lengkap.")
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

        with st.spinner("⚡ AI sedang merumuskan berkas baku Kurikulum Merdeka..."):
            try:
                client = genai.Client(api_key=api_key_input.strip())
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt_final,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                    )
                )

                st.session_state.hasil_teks = response.text
                
                nama_file_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', f"{jenis_perangkat[:2]}_{mapel}_{fase_kelas[:6]}")
                st.session_state.nama_file = f"{nama_file_clean}.txt"
                
                st.toast("Dokumen berhasil dibuat!", icon="✅")

            except Exception as e:
                st.error(f"❌ Terjadi kendala saat menerbitkan berkas: {str(e)}")


# ==========================================
# 6. PENAMPIL HASIL RESMI & UNDUH
# ==========================================
if st.session_state.hasil_teks:
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    st.markdown("### 📄 Lembar Preview Administrasi Resmi")
    
    with st.container(border=True):
        st.markdown(st.session_state.hasil_teks)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Unduh Naskah Dokumen (.TXT / Siap Salin ke Word)",
        data=st.session_state.hasil_teks,
        file_name=st.session_state.nama_file,
        mime="text/plain; charset=utf-8",
        use_container_width=True
    )
