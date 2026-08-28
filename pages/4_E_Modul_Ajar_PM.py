from datetime import datetime
from io import BytesIO
import json
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import google.generativeai as genai
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="GEMA: GENERATOR MODUL AJAR PEMBELAJARAN MENDALAM",
    page_icon="📚",
    layout="wide",
)

# Cek apakah guru sudah login melalui portal utama (app.py)
if "logged_in" not in st.session_state or not st.session_state.logged_in:
  st.warning(
      "⚠️ Anda belum login. Silakan kembali ke Halaman Utama (`app.py`) untuk"
      " masuk ke portal terlebih dahulu."
  )
  st.stop()

# Konfigurasi API Gemini (Mengambil dari st.secrets atau input manual)
try:
  api_key_default = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
  api_key_default = ""

# ===================================
# Custom CSS untuk tampilan UI yang modern dan profesional
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 16px 20px;
        border-radius: 12px;
        border: 1px solid #334155;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .header-title {
        color: #f8fafc;
        font-size: 15px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.3px;
    }
    @keyframes blink-animation {
        0% { opacity: 1; color: #facc15; }
        50% { opacity: 0.35; color: #38bdf8; }
        100% { opacity: 1; color: #facc15; }
    }
    .header-subtitle {
        font-size: 11.5px;
        margin-top: 6px;
        margin-bottom: 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        animation: blink-animation 1.6s infinite ease-in-out;
        font-weight: 600;
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 0.65rem 1rem;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #4338ca 0%, #2563eb 100%);
        box-shadow: 0 6px 18px rgba(59, 130, 246, 0.5);
        transform: translateY(-2px);
    }
    [data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1f2937;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Tampilan Header Modern dalam Card (Diperbarui)
st.markdown(
    """
    <div class="header-card">
        <h2 class="header-title" style="color: white; text-align: center;">
            E-Modul Ajar Pembelajaran Mendalam
        </h2>
        <div class="header-subtitle" style="text-align: center;">
            <b>Pengembang:</b> Yustinus Budi Setyanta - PS Cabdin Bangkalan &nbsp;&nbsp;
            <em>Aplikasi Otomatisasi Perancangan Pembelajaran Deep Learning</em>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# Input Pengguna di Sidebar
with st.sidebar:
  st.success(f"👤 **{st.session_state.get('guru_nama', 'Guru')}**")
  st.markdown("---")
  st.header("⚙️ Parameter Pembelajaran")
  api_key = st.text_input(
      "Masukkan Google Gemini API Key",
      value=api_key_default,
      type="password",
  )

  jenjang_pendidikan = st.selectbox(
      "Pilih Jenjang Pendidikan",
      ["SD / MI", "SMP / MTs", "SMA / MA", "SMK / MAK"],
  )

  if jenjang_pendidikan == "SD / MI":
    default_mapel = "Tematik / Kelas"
    jp_guidance = "Panduan: 1 JP = 35 Menit"
    fase_options = [
        "Fase A / Kelas 1 SD",
        "Fase A / Kelas 2 SD",
        "Fase B / Kelas 3 SD",
        "Fase B / Kelas 4 SD",
        "Fase C / Kelas 5 SD",
        "Fase C / Kelas 6 SD",
    ]
  elif jenjang_pendidikan == "SMP / MTs":
    default_mapel = "Matematika / IPA / IPS"
    jp_guidance = "Panduan: 1 JP = 40 Menit"
    fase_options = [
        "Fase D / Kelas 7 SMP",
        "Fase D / Kelas 8 SMP",
        "Fase D / Kelas 9 SMP",
    ]
  elif jenjang_pendidikan == "SMA / MA":
    default_mapel = "Bahasa Indonesia / Matematika"
    jp_guidance = "Panduan: 1 JP = 45 Menit"
    fase_options = [
        "Fase E / Kelas X SMA",
        "Fase F / Kelas XI SMA",
        "Fase F / Kelas XII SMA",
    ]
  else:
    default_mapel = (
        "Dasar-dasar Teknik Otomotif / Produk Kreatif dan Kewirausahaan"
    )
    jp_guidance = "Panduan: 1 JP = 45 Menit"
    fase_options = [
        "Fase E / Kelas X SMK (Program Dasar Keahlian)",
        "Fase F / Kelas XI SMK (Konsentrasi Keahlian)",
        "Fase F / Kelas XII SMK (Konsentrasi Keahlian)",
    ]

  mata_pelajaran = st.text_input(
      "Mata Pelajaran / Program Kejuruan", default_mapel
  )
  fase_kelas = st.selectbox("Fase / Kelas", fase_options)

  topik = st.text_input(
      "Topik / Materi Pokok / Elemen",
      (
          "Contoh: Pemeliharaan Sistem Rem Kendaraan Ringan"
          if jenjang_pendidikan == "SMK / MAK"
          else "Contoh: Menyimak Teks Laporan Observasi Secara Kritis"
      ),
  )

  st.caption(jp_guidance)
  alokasi_waktu = st.text_input("Alokasi Waktu", "2 JP (2 x 45 Menit)")
  pertemuan_ke = st.text_input("Pertemuan Ke-", "1 (Pertemuan Pertama)")

  st.markdown("---")
  st.header("🏫 Identitas Satuan Pendidikan")
  nama_sekolah = st.text_input(
      "Nama Sekolah",
      st.session_state.get("sekolah", "SMK Negeri 2 Bangkalan"),
  )
  semester = st.selectbox("Semester", ["Ganjil", "Genap"])
  tahun_pelajaran = st.text_input("Tahun Pelajaran", "2026/2027")

  st.markdown("---")
  st.header("✍️ Identitas Pengesahan Dokumen")
  nama_kota = st.text_input("Nama Kota", "Bangkalan")
  tanggal_pembuatan = st.text_input(
      "Tanggal / Bulan / Tahun", datetime.today().strftime("%d %B %Y")
  )
  nama_penulis = st.text_input(
      "Nama Penulis Modul",
      st.session_state.get("guru_nama", "Yustinus Budi Setyanta, S.Pd., M.Pd."),
  )
  nip_penulis = st.text_input("NIP Penulis", "196908302005011003")


# Fungsi pembantu warna latar belakang sel tabel Word
def set_cell_background(cell, fill_color):
  shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>')
  cell._tc.get_or_add_tcPr().append(shading_elm)


# Fungsi pembentuk dokumen Word lengkap berbasis Tabel Matriks sesuai permintaan
def generate_docx(
    data_ai,
    nama_sekolah,
    semester,
    tahun_pelajaran,
    mata_pelajaran,
    fase_kelas,
    topik,
    alokasi_waktu,
    pertemuan_ke,
    nama_penulis,
    nama_kota,
    tanggal_pembuatan,
    nip_penulis,
):
  doc = docx.Document()

  # Pengaturan margin halaman
  for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

  style = doc.styles["Normal"]
  font = style.font
  font.name = "Arial"
  font.size = Pt(10)
  font.color.rgb = RGBColor(51, 51, 51)

  # Judul Utama Dokumen
  p_title = doc.add_paragraph()
  p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_title.paragraph_format.space_before = Pt(0)
  p_title.paragraph_format.space_after = Pt(12)
  run_title = p_title.add_run("MODUL AJAR PEMBELAJARAN MENDALAM")
  run_title.font.name = "Arial"
  run_title.font.size = Pt(15)
  run_title.font.bold = True
  run_title.font.color.rgb = RGBColor(74, 46, 33)  # Coklat Tua

  def add_section_table(title_text, rows_data):
    table = doc.add_table(rows=len(rows_data) + 1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    hdr_cells[0].merge(hdr_cells[1])
    hdr_cells[0].text = title_text
    set_cell_background(hdr_cells[0], "5A3825")
    for p in hdr_cells[0].paragraphs:
      p.alignment = WD_ALIGN_PARAGRAPH.LEFT
      p.paragraph_format.space_before = Pt(4)
      p.paragraph_format.space_after = Pt(4)
      for run in p.runs:
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)

    for idx, (label, val) in enumerate(rows_data):
      row_cells = table.rows[idx + 1].cells
      row_cells[0].text = label
      row_cells[0].width = Inches(2.3)
      row_cells[1].width = Inches(4.2)

      set_cell_background(row_cells[0], "F5EBE0")

      for p in row_cells[0].paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
          run.font.size = Pt(10)
          run.font.bold = True
          run.font.color.rgb = RGBColor(51, 51, 51)

      val_str = str(val).replace("LKPD", "LKM").replace(
          "Lembar Kegiatan Murid", "Lembar Kerja Murid"
      )
      row_cells[1].text = ""

      lines = val_str.split("\n")
      for line_idx, line in enumerate(lines):
        if line_idx == 0:
          p_right = row_cells[1].paragraphs[0]
        else:
          p_right = row_cells[1].add_paragraph()

        p_right.paragraph_format.space_before = Pt(4)
        p_right.paragraph_format.space_after = Pt(4)
        p_right.paragraph_format.line_spacing = 1.15
        p_right.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        if ":" in line:
          parts = line.split(":", 1)
          prefix = parts[0].strip() + ":"
          content = parts[1].strip()

          r_prefix = p_right.add_run(prefix + " ")
          r_prefix.font.size = Pt(10)
          r_prefix.font.bold = True
          r_prefix.font.color.rgb = RGBColor(51, 51, 51)

          r_content = p_right.add_run(content)
          r_content.font.size = Pt(10)
          r_content.font.bold = False
          r_content.font.color.rgb = RGBColor(51, 51, 51)
        else:
          r_normal = p_right.add_run(line)
          r_normal.font.size = Pt(10)
          r_normal.font.bold = False
          r_normal.font.color.rgb = RGBColor(51, 51, 51)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

  # 1. Identifikasi dan Informasi Umum
  tabel_identifikasi = [
      ("Penulis Modul", nama_penulis),
      ("Satuan Pendidikan", nama_sekolah),
      ("Mata Pelajaran", mata_pelajaran),
      ("Fase / Kelas", fase_kelas),
      (
          "Semester / Tahun Pelajaran",
          f"{semester} / {tahun_pelajaran}",
      ),
      ("Materi / Topik", topik),
      ("Alokasi Waktu", alokasi_waktu),
      ("Pertemuan Ke-", pertemuan_ke),
  ]
  add_section_table("IDENTIFIKASI DAN INFORMASI UMUM", tabel_identifikasi)

  # 2. Dimensi Profil Lulusan
  tabel_dpl = [
      (
          "Dimensi Profil Lulusan",
          data_ai.get(
              "dimensi_profil_lulusan",
              "☑ Penalaran Kritis: Peserta didik dilatih menganalisis masalah"
              " secara logis.\n☑ Kolaborasi: Bekerja sama dalam kelompok"
              " investigasi.\n☑ Kemandirian: Bertanggung jawab atas tugas"
              " mandiri.\n☑ Komunikasi: Mempresentasikan hasil kerja.",
          ),
      ),
  ]
  add_section_table("DIMENSI PROFIL LULUSAN", tabel_dpl)

  # 3. Tujuan Pembelajaran
  tabel_tujuan = [
      (
          "Tujuan Pembelajaran",
          data_ai.get(
              "tujuan_pembelajaran",
              "Peserta didik mampu menguasai kompetensi sesuai materi.",
          ),
      ),
  ]
  add_section_table("TUJUAN PEMBELAJARAN", tabel_tujuan)

  # 4. Pemahaman Bermakna & Pertanyaan Pemantik
  tabel_pemahaman = [
      (
          "Pemahaman Bermakna",
          data_ai.get(
              "pemahaman_bermakna",
              "Manfaat praktis dan esensi pembelajaran bagi kehidupan.",
          ),
      ),
      (
          "Pertanyaan Pemantik",
          data_ai.get(
              "pertanyaan_pemantik",
              "Pertanyaan kritis untuk menstimulasi rasa ingin tahu peserta"
              " didik.",
          ),
      ),
  ]
  add_section_table(
      "PEMAHAMAN BERMAKNA & PERTANYAAN PEMANTIK", tabel_pemahaman
  )

  # 5. Kerangka Pembelajaran
  tabel_kerangka = [
      (
          "Praktik Pedagogis",
          data_ai.get(
              "praktik_pedagogis",
              "Model Pembelajaran: Problem Based Learning\nMetode"
              " Pembelajaran Pendukung: Diskusi, Tanya Jawab, Analisis Teks",
          ),
      ),
      (
          "Kemitraan Pembelajaran",
          data_ai.get(
              "kemitraan_pembelajaran",
              "Kemitraan Lingkungan Sekolah: Kolaborasi guru mapel"
              " produktif.\nKemitraan Lingkungan Luar Sekolah: Pemanfaatan"
              " data/narasumber instansi terkait.",
          ),
      ),
      (
          "Lingkungan Belajar",
          data_ai.get(
              "lingkungan_belajar",
              "Ruang Fisik: Kelas fleksibel dan kolaboratif.\nRuang Virtual:"
              " Google Drive / LMS Sekolah.\nBudaya Belajar: Kolaboratif,"
              " Berpikir Kritis, Keterbukaan.",
          ),
      ),
      (
          "Pemanfaatan Digital",
          data_ai.get(
              "pemanfaatan_digital",
              "Tahap Perencanaan: AI & Cloud Storage.\nTahap Pelaksanaan: QR"
              " Code & Audio/Video Digital.\nTahap Asesmen: Google Form /"
              " Menti.",
          ),
      ),
  ]
  add_section_table("KERANGKA PEMBELAJARAN", tabel_kerangka)

  # 6. Pengalaman Belajar
  tabel_pengalaman = [
      (
          "Kegiatan Pendahuluan",
          data_ai.get(
              "kegiatan_pendahuluan",
              "Orientasi, Apersepsi, Motivasi, dan Asesmen Diagnostik awal.",
          ),
      ),
      (
          "Kegiatan Inti (Memahami)",
          data_ai.get(
              "kegiatan_memahami",
              "Eksplorasi konsep dan penyajian masalah autentik.",
          ),
      ),
      (
          "Kegiatan Inti (Mengaplikasi)",
          data_ai.get(
              "kegiatan_mengaplikasi",
              "Penyelidikan kolaboratif dan penerapan konsep dalam LKM.",
          ),
      ),
      (
          "Kegiatan Inti (Merefleksi)",
          data_ai.get(
              "kegiatan_merefleksi",
              "Presentasi kelompok, umpan balik konstruktif, dan penguatan.",
          ),
      ),
      (
          "Kegiatan Penutup",
          data_ai.get(
              "kegiatan_penutup",
              "Refleksi bersama yang menyenangkan (joyful) dan bermakna.",
          ),
      ),
  ]
  add_section_table("PENGALAMAN BELAJAR (LANGKAH-LANGKAH)", tabel_pengalaman)

  # 7. Asesmen Pembelajaran
  tabel_asesmen = [
      (
          "Asesmen Awal",
          data_ai.get(
              "asesmen_awal", "Cek kesiapan sebelum masuk topik pembelajaran."
          ),
      ),
      (
          "Asesmen Proses (Formatif)",
          data_ai.get(
              "asesmen_formatif",
              "Pemantauan partisipasi, keaktifan, dan pemahaman selama"
              " kegiatan.",
          ),
      ),
      (
          "Asesmen Akhir (Sumatif)",
          data_ai.get(
              "asesmen_sumatif",
              "Evaluasi hasil berbasis unjuk kerja atau refleksi kedalaman"
              " konsep.",
          ),
      ),
  ]
  add_section_table("ASESMEN PEMBELAJARAN", tabel_asesmen)

  # Tanda Tangan
  p_sign = doc.add_paragraph()
  p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  p_sign.paragraph_format.space_before = Pt(14)
  p_sign.paragraph_format.space_after = Pt(4)
  p_sign.add_run(f"{nama_kota}, {tanggal_pembuatan}\nPenyusun,\n\n\n")
  run_name = p_sign.add_run(f"{nama_penulis}")
  run_name.font.bold = True
  p_sign.add_run(f"\nNIP. {nip_penulis}")

  # ====================================================
  # HALAMAN TERPISAH 1: RUBRIK PENILAIAN & PEDOMAN PENSKORAN
  # ====================================================
  doc.add_page_break()

  p_rubrik_title = doc.add_paragraph()
  p_rubrik_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_rubrik_title.paragraph_format.space_before = Pt(0)
  p_rubrik_title.paragraph_format.space_after = Pt(12)
  r_rub_t = p_rubrik_title.add_run("RUBRIK PENILAIAN & PEDOMAN PENSKORAN")
  r_rub_t.font.name = "Arial"
  r_rub_t.font.size = Pt(15)
  r_rub_t.font.bold = True
  r_rub_t.font.color.rgb = RGBColor(74, 46, 33)

  table_id_rubrik = doc.add_table(rows=3, cols=2)
  table_id_rubrik.style = "Table Grid"
  table_id_rubrik.alignment = WD_TABLE_ALIGNMENT.CENTER

  table_id_rubrik.rows[0].cells[0].text = "Nama Guru / Pengamat:"
  table_id_rubrik.rows[0].cells[1].text = f"{nama_penulis}"
  table_id_rubrik.rows[1].cells[0].text = "Kelas / Fase:"
  table_id_rubrik.rows[1].cells[1].text = f"{fase_kelas}"
  table_id_rubrik.rows[2].cells[0].text = "Mata Pelajaran / Topik:"
  table_id_rubrik.rows[2].cells[1].text = f"{mata_pelajaran} - {topik}"

  for row in table_id_rubrik.rows:
    row.cells[0].width = Inches(2.3)
    row.cells[1].width = Inches(4.2)
    set_cell_background(row.cells[0], "F5EBE0")
    for cell in row.cells:
      for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
          run.font.size = Pt(10)
          run.font.bold = True

  doc.add_paragraph().paragraph_format.space_after = Pt(6)

  rubrik_data = data_ai.get("rubrik_penilaian", "")
  p_sub = doc.add_paragraph()
  p_sub.paragraph_format.space_before = Pt(6)
  p_sub.paragraph_format.space_after = Pt(4)
  run_sub = p_sub.add_run("A. Rubrik Penilaian Kinerja / Kompetensi")
  run_sub.font.bold = True
  run_sub.font.size = Pt(10)
  run_sub.font.color.rgb = RGBColor(51, 51, 51)

  if isinstance(rubrik_data, dict):
    for k, v in rubrik_data.items():
      if isinstance(v, dict):
        nama = v.get("nama_kriteria", k)
        pb = v.get("perlu_bimbingan", "")
        c = v.get("cukup", "")
        b = v.get("baik", "")
        sb = v.get("sangat_baik", "")

        p_crit = doc.add_paragraph()
        p_crit.paragraph_format.space_before = Pt(4)
        p_crit.paragraph_format.space_after = Pt(2)
        p_crit.paragraph_format.left_indent = Inches(0.2)
        r_nama = p_crit.add_run(f"• {nama}")
        r_nama.font.bold = True
        r_nama.font.size = Pt(10)

        for level_name, level_desc in [
            ("Perlu Bimbingan", pb),
            ("Cukup", c),
            ("Baik", b),
            ("Sangat Baik", sb),
        ]:
          if level_desc:
            p_lvl = doc.add_paragraph()
            p_lvl.paragraph_format.space_before = Pt(1)
            p_lvl.paragraph_format.space_after = Pt(2)
            p_lvl.paragraph_format.left_indent = Inches(0.4)
            p_lvl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r_l_name = p_lvl.add_run(f"- {level_name}: ")
            r_l_name.font.bold = True
            r_l_name.font.size = Pt(9.5)
            r_l_desc = p_lvl.add_run(
                f"{str(level_desc).replace('LKPD', 'LKM').replace('Lembar Kegiatan Murid', 'Lembar Kerja Murid')}"
            )
            r_l_desc.font.bold = False
            r_l_desc.font.size = Pt(9.5)
  else:
    p_rubrik = doc.add_paragraph()
    p_rubrik.paragraph_format.space_before = Pt(4)
    p_rubrik.paragraph_format.space_after = Pt(4)
    p_rubrik.paragraph_format.left_indent = Inches(0.2)
    p_rubrik.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_desc = p_rubrik.add_run(
        str(rubrik_data)
        .replace("LKPD", "LKM")
        .replace("Lembar Kegiatan Murid", "Lembar Kerja Murid")
    )
    r_desc.font.bold = False
    r_desc.font.size = Pt(10)

  p_score_title = doc.add_paragraph()
  p_score_title.paragraph_format.space_before = Pt(8)
  p_score_title.paragraph_format.space_after = Pt(4)
  r_stitle = p_score_title.add_run("B. Pedoman Penskoran & Perhitungan Nilai")
  r_stitle.font.bold = True
  r_stitle.font.size = Pt(10)
  r_stitle.font.color.rgb = RGBColor(51, 51, 51)

  p_rumus = doc.add_paragraph()
  p_rumus.paragraph_format.space_before = Pt(2)
  p_rumus.paragraph_format.space_after = Pt(4)
  p_rumus.paragraph_format.left_indent = Inches(0.2)
  r_r1 = p_rumus.add_run("• Rumus Nilai: ")
  r_r1.font.bold = True
  r_r1.font.size = Pt(9.5)
  r_r2 = p_rumus.add_run(
      "Nilai Akhir = ((Skor Kriteria 1 + Skor Kriteria 2) / 8) x 100"
  )
  r_r2.font.size = Pt(9.5)

  p_kat_title = doc.add_paragraph()
  p_kat_title.paragraph_format.space_before = Pt(2)
  p_kat_title.paragraph_format.space_after = Pt(4)
  p_kat_title.paragraph_format.left_indent = Inches(0.2)
  r_k1 = p_kat_title.add_run("• Kategori Predikat:")
  r_k1.font.bold = True
  r_k1.font.size = Pt(9.5)

  table_score = doc.add_table(rows=5, cols=2)
  table_score.style = "Table Grid"
  table_score.alignment = WD_TABLE_ALIGNMENT.CENTER

  hdr_score = table_score.rows[0].cells
  hdr_score[0].text = "Skor"
  hdr_score[1].text = "Kategori"
  set_cell_background(hdr_score[0], "5A3825")
  set_cell_background(hdr_score[1], "5A3825")
  for cell in hdr_score:
    for p in cell.paragraphs:
      p.alignment = WD_ALIGN_PARAGRAPH.CENTER
      for r in p.runs:
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9.5)

  score_rows_data = [
      ("90 - 100", "Sangat Baik (A)"),
      ("80 - 89", "Baik (B)"),
      ("70 - 79", "Cukup (C)"),
      ("< 70", "Perlu Bimbingan (D)"),
  ]
  for idx, (skor_val, kat_val) in enumerate(score_rows_data):
    row_c = table_score.rows[idx + 1].cells
    row_c[0].text = skor_val
    row_c[1].text = kat_val
    row_c[0].width = Inches(2.0)
    row_c[1].width = Inches(4.5)
    set_cell_background(row_c[0], "F5EBE0")
    for c_idx, cell in enumerate(row_c):
      for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        if c_idx == 0:
          p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
          p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs:
          r.font.size = Pt(9.5)

  p_catatan = doc.add_paragraph()
  p_catatan.paragraph_format.space_before = Pt(6)
  p_catatan.paragraph_format.space_after = Pt(6)
  p_catatan.paragraph_format.left_indent = Inches(0.2)
  r_c1 = p_catatan.add_run("Catatan:\n")
  r_c1.font.bold = True
  r_c1.font.size = Pt(9.5)
  r_c2 = p_catatan.add_run(
      "Murid dinyatakan tuntas/mencapai tujuan pembelajaran jika memperoleh"
      " nilai minimal 70 (Predikat Baik)."
  )
  r_c2.font.size = Pt(9.5)

  doc.add_paragraph().paragraph_format.space_after = Pt(6)

  # ====================================================
  # HALAMAN TERPISAH 2: INSTRUMEN ASESMEN PROSES (FORMATIF)
  # ====================================================
  doc.add_page_break()

  p_inst_title = doc.add_paragraph()
  p_inst_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_inst_title.paragraph_format.space_before = Pt(0)
  p_inst_title.paragraph_format.space_after = Pt(12)
  r_inst_t = p_inst_title.add_run("INSTRUMEN ASESMEN PROSES (FORMATIF)")
  r_inst_t.font.name = "Arial"
  r_inst_t.font.size = Pt(15)
  r_inst_t.font.bold = True
  r_inst_t.font.color.rgb = RGBColor(74, 46, 33)

  table_id_inst = doc.add_table(rows=3, cols=2)
  table_id_inst.style = "Table Grid"
  table_id_inst.alignment = WD_TABLE_ALIGNMENT.CENTER

  table_id_inst.rows[0].cells[0].text = "Nama Guru / Pengamat:"
  table_id_inst.rows[0].cells[1].text = f"{nama_penulis}"
  table_id_inst.rows[1].cells[0].text = "Kelas / Fase:"
  table_id_inst.rows[1].cells[1].text = f"{fase_kelas}"
  table_id_inst.rows[2].cells[0].text = "Mata Pelajaran / Topik:"
  table_id_inst.rows[2].cells[1].text = f"{mata_pelajaran} - {topik}"

  for row in table_id_inst.rows:
    row.cells[0].width = Inches(2.3)
    row.cells[1].width = Inches(4.2)
    set_cell_background(row.cells[0], "F5EBE0")
    for cell in row.cells:
      for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
          run.font.size = Pt(10)
          run.font.bold = True

  doc.add_paragraph().paragraph_format.space_after = Pt(6)

  instrumen_data = data_ai.get("instrumen_formatif", {})
  if isinstance(instrumen_data, dict) and instrumen_data:
    inst_rows = []
    for inst_k, inst_v in instrumen_data.items():
      label_text = (
          inst_k.replace("_", " ")
          .title()
          .replace("Instrumen", "Instrumen Asesmen")
          .replace("Tujuan", "Tujuan Asesmen")
      )
      inst_rows.append((label_text, str(inst_v)))

    add_section_table("LEMBAR OBSERVASI / FORMATIF KELAS", inst_rows)
  else:
    p_inst_text = doc.add_paragraph()
    p_inst_text.paragraph_format.space_before = Pt(4)
    p_inst_text.paragraph_format.space_after = Pt(4)
    p_inst_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_it = p_inst_text.add_run(
        str(instrumen_data)
        .replace("LKPD", "LKM")
        .replace("Lembar Kegiatan Murid", "Lembar Kerja Murid")
    )
    r_it.font.size = Pt(10)

  doc.add_paragraph().paragraph_format.space_after = Pt(6)

  # ====================================================
  # HALAMAN TERPISAH 3: LEMBAR KERJA MURID (LKM)
  # ====================================================
  doc.add_page_break()

  p_lkm_title = doc.add_paragraph()
  p_lkm_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  p_lkm_title.paragraph_format.space_before = Pt(0)
  p_lkm_title.paragraph_format.space_after = Pt(12)
  r_lkm_t = p_lkm_title.add_run("LEMBAR KERJA MURID (LKM)")
  r_lkm_t.font.name = "Arial"
  r_lkm_t.font.size = Pt(15)
  r_lkm_t.font.bold = True
  r_lkm_t.font.color.rgb = RGBColor(74, 46, 33)

  table_id_lkm = doc.add_table(rows=3, cols=2)
  table_id_lkm.style = "Table Grid"
  table_id_lkm.alignment = WD_TABLE_ALIGNMENT.CENTER

  table_id_lkm.rows[0].cells[0].text = "Nama Kelompok / Peserta Didik:"
  table_id_lkm.rows[0].cells[1].text = (
      "........................................................................"
  )
  table_id_lkm.rows[1].cells[0].text = "Kelas / Fase:"
  table_id_lkm.rows[1].cells[1].text = f"{fase_kelas}"
  table_id_lkm.rows[2].cells[0].text = "Mata Pelajaran / Topik:"
  table_id_lkm.rows[2].cells[1].text = f"{mata_pelajaran} - {topik}"

  for row in table_id_lkm.rows:
    row.cells[0].width = Inches(2.3)
    row.cells[1].width = Inches(4.2)
    set_cell_background(row.cells[0], "F5EBE0")
    for cell in row.cells:
      for p in cell.paragraphs:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        for run in p.runs:
          run.font.size = Pt(10)
          run.font.bold = True

  doc.add_paragraph().paragraph_format.space_after = Pt(6)

  lkm_data = data_ai.get("lkm_content", {})
  if isinstance(lkm_data, dict) and lkm_data:
    lkm_rows = []
    for lkm_k, lkm_v in lkm_data.items():
      label_text = (
          lkm_k.replace("_", " ")
          .title()
          .replace("Lkm", "LKM")
          .replace("Judul", "Judul LKM")
          .replace("Tujuan", "Tujuan Pembelajaran")
      )
      lkm_rows.append((label_text, str(lkm_v)))

    add_section_table("STRUKTUR LEMBAR KERJA MURID (LKM)", lkm_rows)
  else:
    p_lkm_text = doc.add_paragraph()
    p_lkm_text.paragraph_format.space_before = Pt(4)
    p_lkm_text.paragraph_format.space_after = Pt(4)
    p_lkm_text.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_lt = p_lkm_text.add_run(
        str(lkm_data)
        .replace("LKPD", "LKM")
        .replace("Lembar Kegiatan Murid", "Lembar Kerja Murid")
    )
    r_lt.font.size = Pt(10)

  bio = BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# Tombol Eksekusi Generator
if st.button("🚀 Buat Modul Ajar Pembelajaran Mendalam"):
  if not api_key:
    st.error("Mohon masukkan Google Gemini API Key terlebih dahulu.")
  elif not topik:
    st.warning("Mohon isi topik pembelajaran.")
  else:
    with st.spinner(
        f"{nama_penulis} sedang menyusun Modul Ajar Pembelajaran Mendalam..."
    ):
      genai.configure(api_key=api_key)
      model = genai.GenerativeModel("gemini-3.5-flash")

      prompt = f"""
            Bertindaklah sebagai pakar kurikulum profesional. Buatkan konten Modul Ajar Berbasis Pembelajaran Mendalam (Deep Learning) yang **SANGAT LENGKAP, DETAIL, DAN KOMPREHENSIF** untuk:
            - Jenjang: {jenjang_pendidikan} ({fase_kelas})
            - Mata Pelajaran: {mata_pelajaran}
            - Topik / Materi Pokok: {topik}
            - Alokasi Waktu: {alokasi_waktu}
            - Pertemuan Ke-: {pertemuan_ke}

            Ketentuan Penting:
            1. Dimensi Profil Lulusan: Pilih 2 hingga 4 dimensi yang PALING RELEVAN dari 8 dimensi berikut (Keimanan dan Ketaqwaan terhadap Tuhan Yang Maha Esa, Kewargaan, Penalaran Kritis, Kreativitas, Kolaborasi, Kemandirian, Kesehatan, Komunikasi). **SANGAT PENTING: Tuliskan dan tampilkan HANYA dimensi yang dipilih saja (dengan tanda centang ☑ dan uraian penjelasannya). JANGAN SAMA SEKALI menyebutkan atau menuliskan daftar dimensi lain yang tidak dipilih/tidak digunakan.**
            2. Praktik Pedagogis: Gunakan format label persis berikut (dengan tanda titik dua):
               - Model Pembelajaran: [Uraian model seperti Problem Based Learning / Discovery Learning / dll]
               - Metode Pembelajaran Pendukung: [Uraian metode, misal 1. Studi Kasus Riil: ... 2. Demonstrasi Interaktif: ... dst]
            3. Kemitraan Pembelajaran: Gunakan format label persis berikut:
               - Kemitraan Lingkungan Sekolah: [...]
               - Kemitraan Lingkungan Luar Sekolah: [...]
            4. Lingkungan Belajar: Gunakan format label persis berikut:
               - Ruang Fisik: [...]
               - Ruang Virtual: [...]
               - Budaya Belajar: [...]
            5. Pemanfaatan Digital: Gunakan format label persis berikut:
               - Tahap Perencanaan: [...]
               - Tahap Pelaksanaan: [...]
               - Tahap Asesmen: [...]
            6. Pengalaman Belajar harus terstruktur mencakup Kegiatan Pendahuluan, Kegiatan Inti (Memahami, Mengaplikasi, Merefleksi), dan Kegiatan Penutup (refleksi joyful dan bermakna). Gunakan istilah **LKM (Lembar Kerja Murid)** (BUKAN LKPD atau Lembar Kegiatan Murid) di seluruh uraian.
            7. Asesmen Pembelajaran mencakup Asesmen Awal, Asesmen Proses (Formatif), dan Asesmen Akhir (Sumatif) beserta Rubrik Penilaian dan Pedoman Penskorannya.
            8. **Instrumen Asesmen Proses (Formatif)**: Sediakan instrumen asesmen proses/formatif yang mendalam (seperti lembar observasi keaktifan, catatan anekdotal, atau rubrik formatif aktivitas siswa) pada kunci `instrumen_formatif` yang terstruktur dengan sub-bagian penting bertanda titik dua agar mudah disajikan dalam bentuk tabel rapi pada halaman khusus sebelum LKM.
            9. **LKM (Lembar Kerja Murid)**: Sediakan konten LKM yang mendalam pada kunci `lkm_content` yang mencakup judul, tujuan, petunjuk kerja, serta langkah-langkah tugas/investigasi peserta didik yang terstruktur rapi dengan sub-bagian penting bertanda titik dua agar mudah ditebalkan di halaman terpisah paling akhir.

            Berikan output HANYA dalam format JSON valid yang memuat kunci-kunci berikut:
            {{
              "dimensi_profil_lulusan": "Hanya tuliskan dimensi profil lulusan yang dipilih saja (gunakan tanda ☑) beserta uraian penerapannya. JANGAN menuliskan dimensi yang tidak dipilih.",
              "tujuan_pembelajaran": "Uraian tujuan pembelajaran yang spesifik, operasional, dan terukur sesuai materi.",
              "pemahaman_bermakna": "Uraian pemahaman bermakna yang mendalam terkait materi.",
              "pertanyaan_pemantik": "2 pertanyaan pemantik yang kontekstual dan menantang daya nalar kritis siswa.",
              "praktik_pedagogis": "Model Pembelajaran: [Isi model]\\nMetode Pembelajaran Pendukung: [Isi metode dengan penomoran]",
              "kemitraan_pembelajaran": "Kemitraan Lingkungan Sekolah: [Isi]\\nKemitraan Lingkungan Luar Sekolah: [Isi]",
              "lingkungan_belajar": "Ruang Fisik: [Isi]\\nRuang Virtual: [Isi]\\nBudaya Belajar: [Isi]",
              "pemanfaatan_digital": "Tahap Perencanaan: [Isi]\\nTahap Pelaksanaan: [Isi]\\nTahap Asesmen: [Isi]",
              "kegiatan_pendahuluan": "Langkah rinci kegiatan pendahuluan (orientasi, apersepsi, asesmen awal).",
              "kegiatan_memahami": "Langkah rinci kegiatan inti pada tahap Memahami.",
              "kegiatan_mengaplikasi": "Langkah rinci kegiatan inti pada tahap Mengaplikasi menggunakan LKM.",
              "kegiatan_merefleksi": "Langkah rinci kegiatan inti pada tahap Merefleksi dan presentasi.",
              "kegiatan_penutup": "Langkah rinci kegiatan penutup yang joyful dan bermakna.",
              "asesmen_awal": "Uraian asesmen awal untuk cek kesiapan belajar.",
              "asesmen_formatif": "Uraian asesmen proses/formatif pemantauan partisipasi.",
              "asesmen_sumatif": "Uraian asesmen akhir/sumatif evaluasi unjuk kerja.",
              "rubrik_penilaian": {{
                "kriteria_1": {{
                  "nama_kriteria": "Nama kriteria pertama sesuai kompetensi materi",
                  "perlu_bimbingan": "Deskripsi tingkat perlu bimbingan",
                  "cukup": "Deskripsi tingkat cukup",
                  "baik": "Deskripsi tingkat baik",
                  "sangat_baik": "Deskripsi tingkat sangat baik"
                }},
                "kriteria_2": {{
                  "nama_kriteria": "Nama kriteria kedua",
                  "perlu_bimbingan": "Deskripsi...",
                  "cukup": "Deskripsi...",
                  "baik": "Deskripsi...",
                  "sangat_baik": "Deskripsi..."
                }}
              }},
              "pedoman_penskoran": {{
                "rumus_nilai": "Rumus perhitungan nilai akhir",
                "kategori_predikat": "Interval nilai dan predikat kelulusan"
              }},
              "instrumen_formatif": {{
                "judul_instrumen": "Judul spesifik instrumen asesmen proses",
                "tujuan_asesmen": "Tujuan penggunaan lembar asesmen formatif",
                "aspek_yang_diamati": "Indikator atau aspek keaktifan/proses yang dinilai dengan format sub-bagian berlabel titik dua (misal: 1. Keaktifan Diskusi: ...\\n2. Kolaborasi Tim: ...)",
                "pedoman_pengamatan": "Petunjuk penilaian atau rubrik ceklis observasi singkat"
              }},
              "lkm_content": {{
                "judul_lkm": "Judul spesifik LKM",
                "tujuan_lkm": "Tujuan pengerjaan LKM bagi peserta didik",
                "petunjuk_kerja": "Langkah panduan keselamatan dan cara pengerjaan dengan format sub-bagian berlabel titik dua (misal: 1. Alat dan Bahan: ...\\n2. Langkah Kerja: ...)",
                "tugas_analisis": "Rincian tugas investigasi, pertanyaan kerja, atau tabel isian praktik dengan format sub-bagian jelas"
              }}
            }}
            """

      response = model.generate_content(prompt)
      text_resp = response.text.strip()

      if text_resp.startswith("```json"):
        text_resp = text_resp[7:]
      if text_resp.startswith("```"):
        text_resp = text_resp[3:]
      if text_resp.endswith("```"):
        text_resp = text_resp[:-3]
      text_resp = text_resp.strip()

      try:
        data_ai = json.loads(text_resp)
      except Exception:
        data_ai = {}

      st.success("🎉 Modul Ajar Sesuai Sistematika Baru Berhasil Disusun!")
      st.info(
          "Dokumen Word (.docx) siap diunduh lengkap dengan halaman terpisah"
          " untuk Rubrik & Pedoman Penskoran (dilengkapi tabel kategori predikat), Instrumen Asesmen Proses (Formatif), dan Lembar Kerja Murid (LKM)."
      )

      docx_file = generate_docx(
          data_ai,
          nama_sekolah,
          semester,
          tahun_pelajaran,
          mata_pelajaran,
          fase_kelas,
          topik,
          alokasi_waktu,
          pertemuan_ke,
          nama_penulis,
          nama_kota,
          tanggal_pembuatan,
          nip_penulis,
      )

      st.download_button(
          label="📥 Unduh Modul Ajar Pembelajaran Mendalam (.docx)",
          data=docx_file,
          file_name=f"Modul_Ajar_{topik.replace(' ', '_')}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )