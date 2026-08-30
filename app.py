import streamlit as st
from google import genai
from google.genai import types
from datetime import datetime
import re
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ==========================================
# 1. KONFIGURASI HALAMAN & CSS RESPONSIF
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
        background-color: #F8FAFC;
    }

    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        max-width: 1200px;
    }

    /* Hero Banner Header */
    .hero-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 60%, #2563EB 100%);
        border-radius: 16px;
        padding: 26px 30px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 20px rgba(15, 23, 42, 0.12);
    }
    .hero-card h1 {
        font-size: 26px;
        font-weight: 800;
        margin: 0 0 6px 0;
        color: #FFFFFF;
    }
    .hero-card p {
        font-size: 13.5px;
        color: #CBD5E1;
        margin: 0;
        line-height: 1.5;
    }

    /* Card & Container */
    [data-testid="stExpander"] {
        background: #FFFFFF !important;
        border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important;
        margin-bottom: 12px;
    }
    
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background: #FFFFFF !important;
        border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important;
        padding: 20px !important;
    }

    /* Tombol Utama */
    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.25) !important;
    }

    .footer-box {
        text-align: center;
        padding: 20px 10px 10px 10px;
        color: #64748B;
        font-size: 12px;
        border-top: 1px solid #E2E8F0;
        margin-top: 36px;
    }

    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

# Logo Tut Wuri Handayani Vektor Murni
LOGO_TUT_WURI_HTML = """
<div style="text-align: center; margin-bottom: 8px;">
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120" width="70" height="70" style="display: inline-block;">
        <polygon points="60,6 114,45 93,109 27,109 6,45" fill="#0284C7" stroke="#0369A1" stroke-width="2"/>
        <polygon points="60,14 105,47 88,101 32,101 15,47" fill="#0EA5E9"/>
        <path d="M60,24 L78,88 L60,74 L42,88 Z" fill="#FACC15"/>
        <circle cx="60" cy="50" r="12" fill="#DC2626"/>
        <circle cx="60" cy="50" r="8" fill="#FFFFFF"/>
        <path d="M28,62 Q60,84 92,62 Q60,104 28,62" fill="#FFFFFF" opacity="0.95"/>
    </svg>
</div>
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
    doc = Document()

    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("|") and line.endswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                row_raw = lines[i].strip()
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
                            clean_val = cell_value.replace("<br>", "\n").replace("<br/>", "\n").replace("**", "")
                            cell.text = clean_val
                            
                            if r_idx == 0:
                                shading_elm = parse_xml(r'<w:shd {} w:fill="E2E8F0"/>'.format(nsdecls('w')))
                                cell._tc.get_or_add_tcPr().append(shading_elm)
                                for p in cell.paragraphs:
                                    for r in p.runs:
                                        r.font.bold = True
                                        r.font.size = Pt(9.5)
                            else:
                                for p in cell.paragraphs:
                                    for r in p.runs:
                                        r.font.size = Pt(9.5)

                tblPr = table._tbl.tblPr
                border_xml = (
                    f'<w:tblBorders {nsdecls("w")}>'
                    f'<w:top w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
                    f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="94A3B8"/>'
                    f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
                    f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
                    f'<w:left w:val="none"/>'
                    f'<w:right w:val="none"/>'
                    f'</w:tblBorders>'
                )
                borders = parse_xml(border_xml)
                tblPr.append(borders)

                doc.add_paragraph()
            continue

        if line.startswith("# "):
            p = doc.add_heading(line[2:], level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("---"):
            p = doc.add_paragraph()
            p_border = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
            p._p.get_or_add_pPr().append(p_border)
        elif line:
            clean_text = line.replace("**", "")
            p = doc.add_paragraph(clean_text)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                r.font.size = Pt(11)
                r.font.name = 'Calibri'
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
# 4. SIDEBAR (NAVIGATOR PANDUAN BERSIH)
# ==========================================
with st.sidebar:
    st.markdown(LOGO_TUT_WURI_HTML, unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; margin: 0; font-size: 16px; font-weight: 700; color: #0F172A;'>STUDIO ADMINISTRASI</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 11px; color: #64748B; margin-top: 2px;'>Standar Kurikulum Merdeka Kemendikbud</p>", unsafe_allow_html=True)
    
    st.divider()

    st.markdown("#### 🔐 Kunci Akses AI (API Key)")
    api_key_input = st.text_input(
        "Tempelkan Gemini API Key Anda:",
        type="password",
        placeholder="AIzaSy...",
        help="API Key adalah kunci digital gratis dari Google."
    )

    with st.expander("📖 Panduan Dapatkan API Key (Gratis)", expanded=False):
        st.markdown("""
        **4 Langkah Mudah Mendapatkan API Key:**
        
        1. **Buka Google AI Studio**  
           Kunjungi: [aistudio.google.com](https://aistudio.google.com/)
        
        2. **Login Akun Google**  
           Gunakan akun Gmail pribadi atau akun belajar.id Anda.
        
        3. **Buat Kunci API**  
           Klik tombol **Get API key** (kiri atas) lalu pilih **Create API key**.
        
        4. **Salin Kode**  
           Salin kode yang berawalan `AIzaSy...` lalu tempelkan pada kolom di atas.

        ---
        *💡 Layanan resmi gratis dari Google & kunci aman tersimpan di peramban Anda.*
        """)

    st.divider()
    st.markdown("""
    **⚙️ Info Sistem:**
    * **Model AI:** Gemini 2.5 Flash
    * **Format Ekspor:** Word (.docx) & Teks (.txt)
    * **Standar:** Regulasi BSKAP Kemendikbud
    """)


# ==========================================
# 5. HALAMAN UTAMA
# ==========================================
st.markdown("""
<div class="hero-card">
    <h1>Generator 22 Perangkat Pembelajaran</h1>
    <p>Terbitkan berkas administrasi dan perangkat pembelajaran Kurikulum Merdeka baku, terstruktur, dan siap ekspor langsung ke Microsoft Word.</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🏛️ 1. Data Sekolah & Dinas", "✍️ 2. Data Penandatangan"])

with tab1:
    col_kop1, col_kop2 = st.columns(2)
    with col_kop1:
        dinas_pendidikan = st.text_input("Dinas Pendidikan Pembina:", value="DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT")
        nama_sekolah = st.text_input("Nama Satuan Pendidikan:", value="SMAS NUSA HARAPAN")
    with col_kop2:
        alamat_sekolah = st.text_input("Alamat Lengkap & Kontak:", value="Jl. Pancasila No. 10, Telp. (0561) 734567")
        kota_sekolah = st.text_input("Kota / Kabupaten Domisili:", value="Pontianak")

with tab2:
    col_staf1, col_staf2, col_staf3 = st.columns(3)
    with col_staf1:
        st.markdown("**Guru Mata Pelajaran**")
        guru_nama = st.text_input("Nama Guru & Gelar:", value="Guru Mata Pelajaran, S.Pd.", key="g_nama")
        guru_nip = st.text_input("NIP (Isi '-' jika Non-PNS):", value="-", key="g_nip")
    with col_staf2:
        st.markdown("**Kepala Sekolah**")
        ks_nama = st.text_input("Nama Kepala Sekolah:", value="ZULKIFLI, S.Pd.", key="ks_nama")
        ks_nip = st.text_input("NIP Kepala Sekolah:", value="197508122005011004", key="ks_nip")
    with col_staf3:
        st.markdown("**Pengawas Pembina**")
        pengawas_nama = st.text_input("Nama Pengawas Pembina:", value="NURHASANAH, M.Si.", key="p_nama")
        pengawas_nip = st.text_input("NIP Pengawas:", value="196811231993032003", key="p_nip")

st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

with st.container(border=True):
    st.markdown("### 📋 3. Konfigurasi Kurikulum & Dokumen")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        mapel = st.text_input("Mata Pelajaran:", value="Fisika", placeholder="Contoh: Fisika, Matematika, Biologi")
        materi_pokok = st.text_input(
            "Fokus Topik / Materi Pokok:",
            value="Pengukuran, Besaran, Satuan, dan Penggunaan Alat Ukur Presisi",
            placeholder="Contoh: Kinematika Gerak Lurus, Usaha dan Energi"
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
            "Pilih Dokumen yang Diterbitkan:",
            options=daftar_22_perangkat,
            index=0
        )

    semester = st.radio(
        "Semester Berjalan:",
        options=["Semester Ganjil", "Semester Genap", "Semester Ganjil & Genap (1 Tahun Penuh)"],
        horizontal=True
    )

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
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

        with st.spinner("⚡ AI sedang menyusun berkas baku sesuai standar Kurikulum Merdeka..."):
            try:
                # Client GenAI Resmi
                client = genai.Client(api_key=api_key_input.strip())
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt_final,
                    config=types.GenerateContentConfig(
                        temperature=0.2,
                    )
                )

                if response and response.text:
                    st.session_state.hasil_teks = response.text
                    nama_file_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', f"{jenis_perangkat[:2]}_{mapel}_{fase_kelas[:6]}")
                    st.session_state.nama_file_base = nama_file_clean
                    st.toast("Dokumen resmi berhasil diterbitkan!", icon="✅")
                else:
                    st.error("❌ Model tidak mengembalikan teks. Silakan coba kembali.")

            except Exception as e:
                st.error(f"❌ Terjadi kendala saat menerbitkan berkas: {str(e)}")


# ==========================================
# 7. PREVIEW & OPSI UNDUH (WORD & TXT)
# ==========================================
if st.session_state.hasil_teks:
    st.markdown("<div style='height: 18px;'></div>", unsafe_allow_html=True)
    st.markdown("### 📄 Lembar Preview Administrasi Resmi")
    
    with st.container(border=True):
        st.markdown(st.session_state.hasil_teks)

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        docx_file = buat_file_docx(st.session_state.hasil_teks)
        st.download_button(
            label="📄 Unduh Dokumen Word (.DOCX)",
            data=docx_file,
            file_name=f"{st.session_state.nama_file_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    with col_dl2:
        st.download_button(
            label="📝 Unduh Teks Mentah (.TXT)",
            data=st.session_state.hasil_teks,
            file_name=f"{st.session_state.nama_file_base}.txt",
            mime="text/plain; charset=utf-8",
            use_container_width=True
        )

# ==========================================
# 8. FOOTER UMUM
# ==========================================
st.markdown("""
<div class="footer-box">
    Studio Administrasi Kurikulum Merdeka Kemendikbudristek RI<br>
    © 2026 Engine AI Perangkat Pembelajaran
</div>
""", unsafe_allow_html=True)
