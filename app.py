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

# Injeksi CSS Responsif
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .stApp {
        background: linear-gradient(180deg, #F1F5F9 0%, #F8FAFC 100%);
    }

    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 1200px;
    }

    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 50%, #2563EB 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 15px 25px -5px rgba(15, 23, 42, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.12);
        position: relative;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.16);
        backdrop-filter: blur(8px);
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 10px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    .hero-title {
        font-size: 28px;
        font-weight: 800;
        margin: 0 0 6px 0;
        letter-spacing: -0.5px;
        color: #FFFFFF;
        line-height: 1.25;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #94A3B8;
        max-width: 720px;
        margin: 0;
        line-height: 1.5;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-top: 1.2rem !important;
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
        }
        .hero-container {
            padding: 20px 18px !important;
            border-radius: 14px !important;
            margin-bottom: 16px !important;
        }
        .hero-title {
            font-size: 20px !important;
        }
        .hero-subtitle {
            font-size: 12.5px !important;
        }
        .stButton > button {
            font-size: 14px !important;
            padding: 10px 16px !important;
        }
    }

    [data-testid="stExpander"] {
        background: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        margin-bottom: 12px;
    }
    
    div[data-testid="stVerticalBlock"] > div[style*="border"] {
        background: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
        padding: 18px !important;
    }

    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        box-shadow: 0 8px 14px -2px rgba(37, 99, 235, 0.25) !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 18px -2px rgba(37, 99, 235, 0.35) !important;
    }

    .stDownloadButton > button {
        font-weight: 700 !important;
        border-radius: 10px !important;
        padding: 12px 18px !important;
    }

    .footer-box {
        text-align: center;
        padding: 24px 10px 12px 10px;
        color: #64748B;
        font-size: 12.5px;
        border-top: 1px solid #E2E8F0;
        margin-top: 40px;
    }

    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
</style>
""", unsafe_allow_html=True)

# Logo Tut Wuri Handayani Vektor Murni (100% Muncul Tanpa Blokir CORS)
LOGO_TUT_WURI_HTML = """
<div style="text-align: center; margin-bottom: 12px;">
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120" width="76" height="76" style="display: inline-block;">
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
# 4. SIDEBAR (AUTENTIKASI & PANDUAN RAMAH GURU)
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
        help="API Key adalah kunci digital gratis dari Google agar aplikasi dapat menulis naskah untuk Anda."
    )

    with st.expander("📖 Panduan Dapatkan API Key (Gratis)", expanded=False):
        st.markdown("""
        <div style="font-size: 12.5px; line-height: 1.6; color: #334155;">
            <p style="margin-bottom: 8px;">Ikuti <strong>4 langkah mudah</strong> berikut untuk mendapatkan kunci resmi gratis dari Google:</p>
            
            <div style="background: #F1F5F9; border-left: 3px solid #2563EB; padding: 6px 10px; border-radius: 4px; margin-bottom: 6px;">
                <strong>Langkah 1:</strong> Buka situs Google AI Studio:<br>
                👉 <a href="https://aistudio.google.com/" target="_blank" style="font-weight: bold; color: #2563EB;">Buka Google AI Studio</a>
            </div>

            <div style="background: #F1F5F9; border-left: 3px solid #2563EB; padding: 6px 10px; border-radius: 4px; margin-bottom: 6px;">
                <strong>Langkah 2:</strong> Masuk (Login) dengan akun Google/Gmail Anda.
            </div>

            <div style="background: #F1F5F9; border-left: 3px solid #2563EB; padding: 6px 10px; border-radius: 4px; margin-bottom: 6px;">
                <strong>Langkah 3:</strong> Klik tombol <strong>"Get API key"</strong> di kiri atas, lalu pilih <strong>"Create API key"</strong>.
            </div>

            <div style="background: #F1F5F9; border-left: 3px solid #2563EB; padding: 6px 10px; border-radius: 4px; margin-bottom: 6px;">
                <strong>Langkah 4:</strong> Klik tombol <strong>Copy</strong> pada kode berawalan <code>AIzaSy...</code>, lalu tempelkan pada kolom di atas.
            </div>

            <div style="background: #FEF3C7; border: 1px solid #FDE68A; padding: 6px 8px; border-radius: 6px; font-size: 11px; color: #92400E; margin-top: 8px;">
                💡 <strong>Catatan:</strong> Layanan ini 100% Gratis dari Google dan kunci Anda aman tersimpan di peramban Anda.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.markdown("""
    <div style="background-color: #F8FAFC; border-radius: 10px; padding: 10px; border: 1px solid #E2E8F0; font-size: 11.5px; color: #475569;">
        <strong>⚙️ Status Mesin AI:</strong><br>
        • Model: Gemini 2.5 Flash / 1.5 Flash<br>
        • Kuota Harian: Tersedia (Gratis)<br>
        • Format Ekspor: .DOCX & .TXT
    </div>
    """, unsafe_allow_html=True)


# ==========================================
# 5. HALAMAN UTAMA (RESPONSIF)
# ==========================================
st.markdown("""
<div class="hero-container">
    <div class="hero-badge">⚡ Professional AI Suite for Teachers</div>
    <div class="hero-title">Generator 22 Perangkat Pembelajaran</div>
    <div class="hero-subtitle">Terbitkan berkas administrasi dan perangkat pembelajaran Kurikulum Merdeka baku, otomatis, dan siap ekspor langsung ke Microsoft Word.</div>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🏛️ 1. Data Sekolah & Dinas", "✍️ 2. Data Penandatangan"])

with tab1:
    col_kop1, col_kop2 = st.columns([1, 1])
    with col_kop1:
        dinas_pendidikan = st.text_input("Dinas Pendidikan Pembina:", value="DINAS PENDIDIKAN PROVINSI KALIMANTAN BARAT")
        nama_sekolah = st.text_input("Nama Satuan Pendidikan:", value="SMAS NUSA HARAPAN")
    with col_kop2:
        alamat_sekolah = st.text_input("Alamat Lengkap & Kontak:", value="Jl. Pancasila No. 10, Telp. (0561) 734567")
        kota_sekolah = st.text_input("Kota / Kabupaten Domisili:", value="Pontianak")

with tab2:
    col_staf1, col_staf2, col_staf3 = st.columns([1, 1, 1])
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
    
    col_p1, col_p2 = st.columns([1, 1])
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
            index=1
        )

    semester = st.radio(
        "Semester Berjalan:",
        options=["Semester Ganjil", "Semester Genap", "Semester Ganjil & Genap (1 Tahun Penuh)"],
        horizontal=True
    )

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
    tombol_proses = st.button("✨ Terbitkan Dokumen Administrasi Resmi", use_container_width=True)


# ==========================================
# 6. LOGIKA GENERASI (DENGAN FALLBACK MODEL)
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
                client = genai.Client(api_key=api_key_input.strip())
                
                # Coba model terbaru, jika gagal fallback otomatis ke model flash stabil
                daftar_model = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-pro']
                response = None
                error_terakhir = None
                
                for m in daftar_model:
                    try:
                        response = client.models.generate_content(
                            model=m,
                            contents=prompt_final,
                            config=types.GenerateContentConfig(
                                temperature=0.2,
                            )
                        )
                        if response and response.text:
                            break
                    except Exception as err:
                        error_terakhir = err
                        continue

                if response and response.text:
                    st.session_state.hasil_teks = response.text
                    nama_file_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', f"{jenis_perangkat[:2]}_{mapel}_{fase_kelas[:6]}")
                    st.session_state.nama_file_base = nama_file_clean
                    st.toast("Dokumen resmi berhasil diterbitkan!", icon="✅")
                else:
                    raise error_terakhir if error_terakhir else Exception("Gagal memproses respons AI.")

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
    
    col_dl1, col_dl2 = st.columns([1, 1])
    
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
    Di Buat Oleh Seorang Guru Yang Gabut<br>
    © 2026 Engine AI Perangkat Pembelajaran
</div>
""", unsafe_allow_html=True)
