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
    page_title="PASTI - Portal Akademik Siswa Terintegrasi",
    page_icon="📚",
    layout="wide",
)

# Atur jarak atas agar konten naik dan tidak terpotong di bawah
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 1rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# ===================================
# Custom CSS untuk tampilan UI yang modern dan kontras
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Kustomisasi Background Sidebar agar lebih kontras */
    [data-testid="stSidebar"] {
        background-color: #17223b;
        border-right: 1px solid #334155;
    }
    
    .header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 20px 25px;
        border-radius: 12px;
        border: 1px solid #334155;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .header-title {
        color: #f8fafc;
        font-size: 26px;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.5px;
        text-align: center;
    }
    .header-subtitle {
        color: #94a3b8;
        font-size: 13px;
        margin-top: 6px;
        margin-bottom: 0;
        text-align: center;
        font-weight: 500;
    }
    .sub-badge {
        display: inline-block;
        background-color: #0f172a;
        color: #38bdf8;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        border: 1px solid #334155;
        margin-top: 10px;
    }
    .section-header {
        font-size: 1.2rem;
        font-weight: 700;
        color: #f8fafc;
        margin-top: 1.2rem;
        margin-bottom: 0.8rem;
        white-space: nowrap;
        letter-spacing: 0.2px;
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 0.75rem 1rem;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #4338ca 0%, #2563eb 100%);
        box-shadow: 0 6px 18px rgba(59, 130, 246, 0.5);
        transform: translateY(-2px);
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Inisialisasi session state untuk login
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
if "guru_nama" not in st.session_state:
  st.session_state.guru_nama = ""

# =====================================================================
# HALAMAN 1: PORTAL UTAMA (PASTI) & LOGIN
# =====================================================================
st.title("PASTI - Portal Akademik Siswa Terintegrasi")
  st.markdown(
      """
        <div class="header-card" style="text-align: center;">
            <h1 class="header-title">PASTI</h1>
            <div class="header-subtitle" style="font-size: 15px; font-weight: 600; color: #cbd5e1; margin-top: 4px;">
                Portal Akademik Siswa Terintegrasi
            </div>
            <div style="font-size: 12px; color: #94a3b8; margin-top: 2px;">
                Pusat Kendali Aplikasi Pembelajaran dan Administrasi Guru
            </div>
            <div>
                <span class="sub-badge">Pengembang: Yustinus Budi Setyanta - PS Cabdin Bangkalan | Sistem Otomatisasi Terintegrasi</span>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )

  if not st.session_state.logged_in:
    st.markdown("### 🔐 Login Akses Portal")
    st.markdown(
        "Silakan masukkan **Email** terdaftar atau **Token Unik** Anda untuk"
        " masuk ke sistem."
    )

    with st.form("login_form"):
      email_token = st.text_input(
          "Email / Token Unik Guru",
          placeholder="Contoh: yustinussetyanta08@dinas.belajar.id atau"
          " TOKEN300869",
      )
      submitted = st.form_submit_button("🚀 Masuk Portal")

      if submitted:
        if email_token.strip():
          st.session_state.logged_in = True
          st.session_state.guru_nama = (
              "Yustinus Budi Setyanta, S.Pd., M.Pd."
              if "yustinus" in email_token.lower()
              else "Guru Pengampu"
          )
          st.success("✅ Berhasil masuk ke portal! Memuat ulang sistem...")
          st.rerun()
        else:
          st.warning("⚠️ Mohon masukkan Email atau Token Unik Anda terlebih dahulu.")
  else:
    st.success(
        f"👤 Selamat datang kembali, **{st.session_state.guru_nama}**!"
    )
    st.info(
        "Anda sudah masuk ke sistem. Silakan pilih menu di sidebar sebelah kiri"
        " untuk mengelola E Presensi, E Jurnal, E Asesmen, atau E Modul Ajar."
    )
    if st.button("🚪 Keluar (Logout)"):
      st.session_state.logged_in = False
      st.session_state.guru_nama = ""
      st.rerun()

# =====================================================================
# HALAMAN 2: E PRESENSI SISWA
# =====================================================================
elif menu == "E Presensi Siswa":
  if not st.session_state.logged_in:
    st.warning("⚠️ Silakan login terlebih dahulu melalui menu [Portal Utama].")
    st.stop()

  st.markdown(
      '<div class="section-header">📋 E Presensi Siswa</div>',
      unsafe_allow_html=True,
  )
  st.write(
      "Fitur pengelolaan kehadiran siswa secara digital dan real-time terintegrasi."
  )
  # Tambahkan logika presensi di sini sesuai kebutuhan

# =====================================================================
# HALAMAN 3: E JURNAL MENGAJAR
# =====================================================================
elif menu == "E Jurnal Mengajar":
  if not st.session_state.logged_in:
    st.warning("⚠️ Silakan login terlebih dahulu melalui menu [Portal Utama].")
    st.stop()

  st.markdown(
      '<div class="section-header">📖 E Jurnal Mengajar Guru</div>',
      unsafe_allow_html=True,
  )
  st.write("Catat agenda harian, capaian materi, dan catatan kelas di sini.")
  # Tambahkan logika jurnal mengajar di sini sesuai kebutuhan

# =====================================================================
# HALAMAN 4: E ASESMEN PM
# =====================================================================
elif menu == "E Asesmen PM":
  if not st.session_state.logged_in:
    st.warning("⚠️ Silakan login terlebih dahulu melalui menu [Portal Utama].")
    st.stop()

  st.markdown(
      '<div class="section-header">📝 E Asesmen Pembelajaran Mendalam</div>',
      unsafe_allow_html=True,
  )
  st.write("Kelola lembar evaluasi formatif dan sumatif berbasis deep learning.")
  # Tambahkan logika asesmen di sini sesuai kebutuhan

# =====================================================================
# HALAMAN 5: E MODUL AJAR PM
# =====================================================================
elif menu == "E Modul Ajar PM":
  if not st.session_state.logged_in:
    st.warning("⚠️ Silakan login terlebih dahulu melalui menu [Portal Utama].")
    st.stop()

  st.markdown(
      '<div class="section-header">⚙️ Parameter Pembelajaran Modul Ajar</div>',
      unsafe_allow_html=True,
  )

  try:
    api_key_default = st.secrets.get("GEMINI_API_KEY", "")
  except Exception:
    api_key_default = ""

  api_key = st.text_input(
      "Masukkan Google Gemini API Key", value=api_key_default, type="password"
  )

  col_param1, col_param2 = st.columns(2)

  with col_param1:
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

  with col_param2:
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

  col_id1, col_id2 = st.columns(2)

  with col_id1:
    st.markdown(
        '<div class="section-header">🏫 Identitas Satuan Pendidikan</div>',
        unsafe_allow_html=True,
    )
    nama_sekolah = st.text_input(
        "Nama Sekolah", st.session_state.get("sekolah", "SMK Negeri 2 Bangkalan")
    )
    semester = st.selectbox("Semester", ["Ganjil", "Genap"])
    tahun_pelajaran = st.text_input("Tahun Pelajaran", "2026/2027")

  with col_id2:
    st.markdown(
        '<div class="section-header">✍️ Identitas Pengesahan Dokumen</div>',
        unsafe_allow_html=True,
    )
    nama_kota = st.text_input("Nama Kota", "Bangkalan")
    tanggal_pembuatan = st.text_input(
        "Tanggal / Bulan / Tahun", datetime.today().strftime("%d %B %Y")
    )
    nama_penulis = st.text_input(
        "Nama Penulis Modul",
        st.session_state.get(
            "guru_nama", "Yustinus Budi Setyanta, S.Pd., M.Pd."
        ),
    )
    nip_penulis = st.text_input("NIP Penulis", "196908302005011003")

  st.markdown("---")
  st.markdown("### 🚀 Generator Modul Ajar Berbasis Pembelajaran Mendalam")

  if st.button(
      "🚀 Buat Modul Ajar Pembelajaran Mendalam", use_container_width=True
  ):
    if not api_key:
      st.error("Mohon masukkan Google Gemini API Key terlebih dahulu.")
    elif not topik:
      st.warning("Mohon isi topik pembelajaran.")
    else:
      st.success("Modul Ajar siap diproses menggunakan AI.")
