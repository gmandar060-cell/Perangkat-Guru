import streamlit as st
from google import genai
from google.genai import types
from datetime import datetime
import re
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==========================================
# 1. KONFIGURASI HALAMAN & INJEKSI CSS MODERN
# ==========================================
st.set_page_config(
    page_title="Studio Administrasi Kurikulum Merdeka",
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

    /* Hero Banner */
    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 20px;
        padding: 32px 36px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 15px 25px -5px rgba(15, 23, 42, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(8px);
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 10px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    .hero-title {
        font-size: 30px;
        font-weight: 800;
        margin: 0 0 6px 0;
        letter-spacing: -0.5px;
        color: #FFFFFF;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #94A3B8;
        max-width: 720px;
        margin: 0;
        line-height: 1.5;
    }

    /* Container Card */
    [data-testid="stExpander"] {
        background: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        margin-bottom: 14px;
    }
    
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        padding: 22px !important;
    }

    /* Action Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        box-shadow: 0 8px 12px -2px rgba(37, 99, 235, 0.25) !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 18px -2px rgba(37, 99, 235, 0.35) !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

LOGO_TUT_WURI_SVG = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="70" height="70" style="display: block; margin: 0 auto 10px auto;">
  <polygon points="50,5 95,38 78,92 22,92 5,38" fill="#0284C7" stroke="#0369A1" stroke-width="2"/>
  <path d="M50,15 L70,80 L50,65 L30,80 Z" fill="#FACC15"/>
  <circle cx="50" cy="40" r="10" fill="#DC2626"/>
  <path d="M25,50 Q50,70 75,50 Q50,90 25,50" fill="#FFFFFF" opacity="0.9"/>
</svg>
"""

# Inisialisasi Session State
if "hasil_teks" not in st.session_state:
    st.session_state.hasil_teks = ""
if "nama_file_base" not in st.session_state:
    st.session_state.nama_file_base = "Perangkat_Kurikulum_Merdeka"


# ==========================================
# 2. HELPER GENERATOR DOKUMEN WORD (.DOCX)
# ==========================================
def buat_file_docx(markdown_text: str) -> io.BytesIO:
    """Mengonversi Markdown output AI ke file DOCX berformat standar baku."""
    doc = Document()

    # Set Margin Standar A4 (Normal Margin: 2.54 cm)
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # Handle Tabel Markdown
        if line.startswith("|") and line.endswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                row_raw = lines[i].strip()
                # Skip baris pembatas tabel markdown (| :--- | :---: |)
                if not re.match(r"^\|[\s\-:]+(\|[\s\-:]+)+\|$", row_raw):
                    cells = [c.strip() for c in row_raw[1:-1].split("|")]
                    table_lines.append(cells)
                i += 1

            if table_lines:
                num_rows = len(table_lines)
                num_cols = max(len(r) for r in table_lines)
                table = doc.add_table(rows=num_rows, cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False

                for r_idx, row_data in enumerate(table_lines):
                    for c_idx, cell_value in enumerate(row_data):
                        if c_idx < num_cols:
                            cell = table.cell(r_idx, c_idx)
                            # Bersihkan tag HTML sederhana seperti <br>
                            clean_cell_value = cell_value.replace("<br>", "\n").replace("<br/>", "\n").replace("**", "")
                            cell.text = clean_cell_value
                            
                            # Styling Header Baris Pertama
                            if r_idx == 0:
                                shading_elm = parse_xml(r'<w:shd {} w:fill="E2E8F0"/>'.format(nsdecls('w')))
                                cell._tc.get_or_add_tcPr().append(shading_elm)
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.bold = True
                                        run.font.size = Pt(9.5)
                            else:
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.size = Pt(9.5)

                # Set Border Tipis pada Tabel
                tblPr = table._tbl.tblPr
                borders = parse_xml(r'''
                    <w:tblBorders {} >
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>
                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
                        <w:left w:val="none"/>
                        <w:right w:val="none"/>
                    </w:tblBorders>
                '''.format(nsdecls('w')))
                tblPr.append(borders)

                doc.add_paragraph()  # Jarak setelah tabel
            continue

        # Handle Heading
        if line.startswith("# "):
            p = doc.add_heading(line[2:], level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("---"):
            p = doc.add_paragraph()
            p_border = parse_xml(r'<w:pBdr {}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>'.format(nsdecls('w')))
            p._p.get_or_add_pPr().append(p_border)
        elif line:
            # Format Paragraf Standar
            clean_text = line.replace("**", "")
            p = doc.add_paragraph(clean_text)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(11)
                run.font.name = 'Calibri'
        i += 1

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# ==========================================
# 3. FUNGSI LOGIKA MASTER PROMPT
# ==========================================
def buat_instruksi_prompt(data: dict) -> str:
    p3_str = ", ".join(data['profil_pancasila']) if data['profil_pancasila'] else "Sesuai Karakteristik Materi"

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
- Jenis Dokumen       : {data['jenis_perangkat']}
- Mata Pelajaran      : {data['mapel']}
- Tingkatan / Fase    : {data['fase_kelas']}
- Semester            : {data['semester']}
- Alokasi Waktu       : {data['alokasi_waktu']}
- Fokus Materi Pokok  : {data['materi_pokok']}
- Dimensi Profil P3   : {p3_str}
- Tanggal Terbit      : {data['tanggal_hari_ini']}
- Kota Instansi       : {data['kota_sekolah']}

=== KETENTUAN FORMAT DOKUMEN (FORMAT BAKU RESMI INDONESIA) ===
1. BAGIAN PALING ATAS: Tulis KOP SURAT RESMI SEKOLAH huruf kapital terpusat rapi, mencakup nama Dinas Pendidikan, Satuan Pendidikan, dan Alamat lengkap, diikuti garis pembatas horizontal '---'.
2. BAGIAN KEDUA: Judul Dokumen resmi ({data['jenis_perangkat'].upper()}) diikuti TABEL IDENTITAS (Mata Pelajaran, Fase/Kelas, Semester, Materi Pokok, Dimensi P3, Alokasi Waktu).
3. BAGIAN KETIGA (ISI DOKUMEN UTAMA):
   - Tulis secara UTUH, SANGAT DETAIL, dan OPERASIONAL (Gunakan kata kerja operasional Taksonomi Bloom revisi).
   - Fokuskan konten secara mendalam pada materi pokok: "{data['materi_pokok']}".
   - DILARANG KERAS memotong materi atau menggunakan kata-kata singkatan seperti '...dst', '[lanjutkan]', '[sesuaikan]', atau 'dan seterusnya'.
   - Sajikan komponen data terstruktur menggunakan format TABEL MARKDOWN yang rapi, padat, dan jelas.
4. BAGIAN PALING BAWAH (LEMBAR PENGESAHAN):
   Buat lembar tanda tangan 3 kolom horizontal berdampingan dengan format markdown table yang rapi:
   | Mengetahui,<br>Pengawas Pembina | Mengetahui,<br>Kepala Sekolah | {data['kota_sekolah']}, {data['tanggal_hari_ini']}<br>Guru Mata Pelajaran |
   | :---: | :---: | :---: |
   | <br><br><br><br> | <br><br><br><br> | <br><br><br><br> |
   | **{data['pengawas_nama']}**<br>NIP. {data['pengawas_nip']} | **{data['ks_nama']}**<br>NIP. {data['ks_nip']} | **{data['guru_nama']}**<br>NIP. {data['guru_nip']} |

5. ATURAN PENULISAN OUTPUT:
   - Langsung mulai dari baris KOP SURAT pertama tanpa salam pembuka robot AI.
   - Akhiri langsung setelah tabel lembar pengesahan tanpa kalimat penutup.
"""
    return prompt.strip()


# ==========================================
# 4. SIDEBAR (AUTENTIKASI & STATUS)
# ==========================================
with st.sidebar:
    st.markdown(LOGO_TUT_WURI_SVG, unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; margin: 0; font-size: 17px; font-weight: 700; color: #0F172A;'>STUDIO ADMINISTRASI</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 12px; color: #64748B; margin-top: 2px;'>Kurikulum Merdeka BSKAP Kemendikbud</p>", unsafe_allow_html=True)
    st.divider()

    st.markdown("#### 🔐 Kunci Akses AI")
    api_key_input = st.text_input(
        "Gemini API Key:",
        type="password",
        placeholder="AIzaSy...",
        help="API Key Anda aman dan hanya aktif selama sesi peramban berjalan."
    )

    with st.expander("❓ Cara Mendapatkan API Key"):
        st.markdown("""
        1. Buka [Google AI Studio](https://aistudio.google.com/).
        2. Masuk menggunakan akun Google.
        3. Klik **Get API Key** lalu **Create API Key**.
        4. Salin dan tempel kuncinya ke kolom di atas.
        """)

    st.divider()
    st.markdown("""
    <div style="background-color: #F8FAFC; border-radius: 10px; padding: 12px; border: 1px solid #E2E8F0; font-size: 12px; color: #475569;">
        <strong>⚙️ Spesifikasi Model:</strong><br>
        • Google Gemini 2.5 Flash<br>
        • Format Output Baku Docx / Text<br>
        • Terkalibrasi Standar Akreditasi
    </div>
    """, unsafe_allow_html=True)


# ==========================================
# 5. HALAMAN UTAMA (DASHBOARD)
# ==========================================
st.markdown("""
<div class="hero-container">
    <div class="hero-badge">⚡ Professional AI Suite for Teachers</div>
    <div class="hero-title">Generator 22 Perangkat Pembelajaran</div>
    <div class="hero-subtitle">Terbitkan berkas administrasi dan dokumen pembelajaran terstruktur, operasional, dan siap ekspor langsung ke format Microsoft Word.</div>
</div>
""", unsafe_allow_html=True)

# Tahap 1 & 2 dalam Tab
tab1, tab2 = st.tabs(["🏛️ 1. Identitas Sekolah & Dinas", "✍️ 2. Pejabat & Penandatangan"])

with tab1:
    col_kop1, col_kop2 = st.columns(2)
    with col_kop1:
        dinas_pendidikan = st.text_input("Dinas Pendidikan Pembina:", value="DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT")
        nama_sekolah = st.text_input("Nama Satuan Pendidikan:", value="SMAS NUSA HARAPAN")
    with col_kop2:
        alamat_sekolah = st.text_input("Alamat & Kontak Sekolah:", value="Jl. Pancasila No. 10, Telp. (0561) 734567, Pontianak")
        kota_sekolah = st.text_input("Kota / Kabupaten:", value="Pontianak")

with tab2:
    col_staf1, col_staf2, col_staf3 = st.columns(3)
    with col_staf1:
        st.markdown("**Guru Mata Pelajaran**")
        guru_nama = st.text_input("Nama Guru & Gelar:", value="MUHAMMAD NURZULIANDAR, S.Pd.", key="g_nama")
        guru_nip = st.text_input("NIP (Isi '-' jika Non-PNS):", value="-", key="g_nip")
    with col_staf2:
        st.markdown("**Kepala Sekolah**")
        ks_nama = st.text_input("Nama Kepala Sekolah:", value="ZULKIFLI, S.Pd.", key="ks_nama")
        ks_nip = st.text_input("NIP Kepala Sekolah:", value="197508122005011004", key="ks_nip")
    with col_staf3:
        st.markdown("**Pengawas Pembina**")
        pengawas_nama = st.text_input("Nama Pengawas:", value="NURHASANAH, M.Si.", key="p_nama")
        pengawas_nip = st.text_input("NIP Pengawas:", value="196811231993032003", key="p_nip")

st.markdown("<br>", unsafe_allow_html=True)

# Tahap 3: Parameter Kurikulum & Input Baru
with st.container(border=True):
    st.markdown("### 📋 3. Konfigurasi Kurikulum & Fokus Materi")
    
    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        mapel = st.text_input("Mata Pelajaran:", value="Fisika", placeholder="Contoh: Fisika, Biologi, Matematika")
        materi_pokok = st.text_input(
            "Fokus Topik / Materi Pokok Pembelajaran:",
            value="Pengukuran, Besaran, Satuan, dan Penggunaan Alat Ukur Presisi",
            placeholder="Contoh: Kinematika Gerak Lurus, Ikatan Kimia, Teks Naratif"
        )
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

    with col_p2:
        alokasi_waktu = st.text_input("Alokasi Waktu / Target JP:", value="3 JP / Minggu (Total 54 JP per Semester)")
        
        profil_pancasila = st.multiselect(
            "Dimensi Profil Pelajar Pancasila (P3):",
            options=[
                "Beriman, Bertakwa kepada Tuhan YME, & Berakhlak Mulia",
                "Berkebinekaan Global",
                "Gotong Royong",
                "Mandiri",
                "Bernalar Kritis",
                "Kreatif"
            ],
            default=["Bernalar Kritis", "Gotong Royong", "Mandiri"]
        )

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

    semester = st.radio(
        "Semester Berjalan:",
        options=["Semester Ganjil", "Semester Genap", "Semester Ganjil & Genap (1 Tahun Penuh)"],
        horizontal=True
    )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    tombol_proses = st.button("✨ Terbitkan Dokumen Administrasi Resmi", use_container_width=True)


# ==========================================
# 6. LOGIKA GENERASI
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
            "materi_pokok": materi_pokok,
            "profil_pancasila": profil_pancasila,
            "fase_kelas": fase_kelas,
            "semester": semester,
            "alokasi_waktu": alokasi_waktu,
            "jenis_perangkat": jenis_perangkat,
            "tanggal_hari_ini": tanggal_sekarang
        }

        prompt_final = buat_instruksi_prompt(data_input)

        with st.spinner("⚡ AI sedang menyusun berkas baku sesuai standar BSKAP..."):
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
                st.session_state.nama_file_base = nama_file_clean
                
                st.toast("Dokumen resmi berhasil diterbitkan!", icon="✅")

            except Exception as e:
                st.error(f"❌ Terjadi kendala saat menerbitkan berkas: {str(e)}")


# ==========================================
# 7. PREVIEW & OPSI UNDUH (WORD & TXT)
# ==========================================
if st.session_state.hasil_teks:
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    st.markdown("### 📄 Lembar Preview Administrasi Resmi")
    
    with st.container(border=True):
        st.markdown(st.session_state.hasil_teks)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        # Ekspor Dokumen Word (.docx)
        docx_file = buat_file_docx(st.session_state.hasil_teks)
        st.download_button(
            label="📄 Unduh Berkas Microsoft Word (.DOCX)",
            data=docx_file,
            file_name=f"{st.session_state.nama_file_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    with col_dl2:
        # Ekspor Teks Mentah (.txt)
        st.download_button(
            label="📝 Unduh Naskah Mentah (.TXT)",
            data=st.session_state.hasil_teks,
            file_name=f"{st.session_state.nama_file_base}.txt",
            mime="text/plain; charset=utf-8",
            use_container_width=True
        )
