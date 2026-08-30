from datetime import datetime
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from styles import apply_global_styles

# --- KONFIGURASI HALAMAN UTAMA ---
st.set_page_config(
    page_title="PASTI - Portal Akademik Siswa Terintegrasi",
    page_icon="🚀",
    layout="wide",
)

# Terapkan styling global terpusat di sini
apply_global_styles()

# Atur jarak bawah agar tampilan tidak terpotong
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 0.8rem;
            padding-bottom: 6rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ID Master Registry Pusat
MASTER_REGISTRY_ID = "1mgN63xzrLt__5b9-gBw8dIWYP3RRgNdagUiTurFZdgg"


def get_gspread_client():
  scope = [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
  ]
  creds_dict = dict(st.secrets["gcp_service_account"])
  creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
  return gspread.authorize(creds)


@st.cache_resource
def load_master_registry():
  try:
    client = get_gspread_client()
    sh = client.open_by_key(MASTER_REGISTRY_ID)
    worksheet = sh.worksheet("DATABASE_MASTER_REGISTRY")
    return worksheet.get_all_records()
  except Exception as e:
    st.error(
        f"❌ Gagal terhubung ke Google Spreadsheet Master Registry. Detail"
        f" Error: {e}"
    )
    return None


# --- INISIALISASI SESSION STATE GLOBAL ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
if "guru_nama" not in st.session_state:
  st.session_state.guru_nama = ""
if "spreadsheet_id" not in st.session_state:
  st.session_state.spreadsheet_id = ""
if "gemini_api_key" not in st.session_state:
  st.session_state.gemini_api_key = ""

# --- STYLING CSS TAMBAHAN (DIPADATKAN AGAR TIDAK TERPOTONG) ---
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main-header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 14px 20px;
        border-radius: 10px;
        border: 1px solid #334155;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px -3px rgba(0, 0, 0, 0.4);
        text-align: center;
    }
    .login-container-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 14px 20px;
        border-radius: 10px;
        border: 1px solid #334155;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px -3px rgba(0, 0, 0, 0.4);
    }
    .main-title-text {
        color: #38bdf8;
        font-size: 22px;
        font-weight: 800;
        margin: 2px 0 2px 0;
        letter-spacing: 0.5px;
        text-shadow: 0 0 10px rgba(56, 189, 248, 0.4);
    }
    .sub-title-2 {
        color: #94a3b8;
        font-size: 12px;
        font-weight: 500;
        margin-top: 0;
        margin-bottom: 6px;
    }
    .dev-badge {
        display: inline-block;
        background: rgba(56, 189, 248, 0.1);
        color: #38bdf8;
        padding: 3px 12px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
        border: 1px solid rgba(56, 189, 248, 0.3);
    }
    .user-welcome-card {
        background: linear-gradient(135deg, #1e293b 0%, #111827 100%);
        padding: 14px 18px;
        border-radius: 10px;
        border-left: 4px solid #38bdf8;
        border: 1px solid #334155;
        color: #e2e8f0;
        font-size: 14px;
        margin-bottom: 12px;
    }
    .info-notice-card {
        background: #111827;
        padding: 12px 18px;
        border-radius: 10px;
        border: 1px solid #1f2937;
        color: #cbd5e1;
        font-size: 13px;
        margin-bottom: 15px;
    }
    .module-card {
        background: #111827;
        padding: 18px;
        border-radius: 10px;
        border: 1px solid #1f2937;
        height: 100%;
        transition: all 0.3s ease;
    }
    .module-card:hover {
        border-color: #38bdf8;
        box-shadow: 0 4px 15px rgba(56, 189, 248, 0.15);
        transform: translateY(-2px);
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        background: linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%);
        color: white;
        border: none;
        padding: 0.5rem 1rem;
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

# --- HEADER UTAMA PORTAL PASTI (LOGO + AKRONIM RAPI) ---
st.markdown(
    """
    <div class="main-header-card">
        <img src="https://lh3.googleusercontent.com/d/15rUWzaqM_86lF2ht8atJmmyPocUPxl_z" alt="Logo PASTI" style="width: 75px; height: auto; margin: 0 auto 4px auto; display: block; filter: drop-shadow(0 4px 10px rgba(0,0,0,0.4));">
        <div class="main-title-text">Portal Administrasi Siswa Terintegrasi</div>
        <div class="sub-title-2">Pusat Kendali Administrasi Guru</div>
        <div class="dev-badge">
            <b>Pengembang:</b> Yustinus Budi Setyanta - PS Cabdin Bangkalan &nbsp;|&nbsp; <em>Sistem Otomatisasi Terintegrasi</em>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- KONDISI 1: JIKA BELUM LOGIN (TAMPILKAN HALAMAN LOGIN) ---
if not st.session_state.logged_in:
  st.markdown(
      """
        <div class="login-container-card">
            <h4 style='color: #38bdf8; margin: 0 0 4px 0; font-size: 16px;'>🔐 Login Akses Portal</h4>
            <p style='color: #94a3b8; font-size: 13px; margin: 0;'>
                Silakan masukkan <b>Email</b> terdaftar atau <b>Token Unik</b> Anda untuk masuk ke sistem.
            </p>
        </div>
        """,
      unsafe_allow_html=True,
  )

  with st.form("form_login"):
    user_input = st.text_input(
        "Email / Token Unik Guru",
        placeholder=(
            "Contoh: yustinus_bkl@gmail.com atau TOKENPASTI12345"
        ),
    )
    btn_login = st.form_submit_button("🚀 Masuk Portal", use_container_width=True)

    if btn_login:
      if not user_input:
        st.warning("⚠️ Mohon masukkan Email atau Token Unik Anda.")
      else:
        with st.spinner("Memeriksa kredensial ke Master Registry..."):
          data_registry = load_master_registry()

        if data_registry:
          df_registry = pd.DataFrame(data_registry)
          df_registry.columns = df_registry.columns.str.strip()

          matched = df_registry[
              (
                  df_registry["Email"]
                  .astype(str)
                  .str.strip()
                  .str.lower()
                  == user_input.strip().lower()
              )
              | (
                  df_registry["Token_Unik"].astype(str).str.strip()
                  == user_input.strip()
              )
          ]

          if not matched.empty:
            st.session_state.logged_in = True
            st.session_state.guru_nama = matched.iloc[0]["Nama_Guru"]
            st.session_state.spreadsheet_id = matched.iloc[0]["Spreadsheet_ID"]
            st.success(
                f"🎉 Selamat datang, {st.session_state.guru_nama}! Berhasil"
                " masuk."
            )
            st.rerun()
          else:
            st.error(
                "❌ Email atau Token Unik tidak ditemukan di Master Registry."
            )
        else:
          st.error(
              "❌ Database Master Registry kosong atau gagal diakses. Periksa"
              " koneksi spreadsheet Anda."
          )

# --- KONDISI 2: JIKA SUDAH LOGIN (TAMPILKAN SIDEBAR & NAVIGASI MODUL) ---
else:
  with st.sidebar:
    st.markdown(
        f"""
            <div style="text-align: center; padding: 10px; background: #1e293b; border-radius: 8px; margin-bottom: 15px; border: 1px solid #334155;">
                <span style="font-size: 24px;">👨‍💻</span><br>
                <b style="color: #facc15; font-size: 14px;">{st.session_state.guru_nama}</b><br>
                <span style="color: #94a3b8; font-size: 11px;">Sesi Aktif & Terverifikasi</span>
            </div>
            """,
        unsafe_allow_html=True,
    )

    # Menu Navigasi Samping (Sidebar)
    menu_pilihan = st.radio(
        "📌 Navigasi Menu",
        [
            "🏠 Dashboard Utama",
            "📋 E-Presensi Siswa",
            "📖 E-Jurnal Mengajar",
            "📝 E-Asesmen PM",
            "🤖 E-Modul Ajar PM",
        ],
    )

    st.markdown("---")
    if st.button("🚪 Keluar / Logout"):
      st.session_state.logged_in = False
      st.session_state.guru_nama = ""
      st.session_state.spreadsheet_id = ""
      st.session_state.gemini_api_key = ""
      st.rerun()

  # --- KONTROL HALAMAN BERDASARKAN MENU SIDEBAR ---
  if menu_pilihan == "🏠 Dashboard Utama":
    st.markdown(
        f"""
        <div class="user-welcome-card">
            ✅ Anda sudah masuk sebagai <b>{st.session_state.guru_nama}</b>.
        </div>
        <div class="info-notice-card">
            👉 Silakan pilih modul aplikasi di menu sebelah kiri (Sidebar) untuk mulai bekerja <b>(E Presensi Siswa, E Jurnal Mengajar, E-Asesmen PM, E-Modul Ajar PM)</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### 📑 Modul Aplikasi Tersedia")

    coll1, coll2 = st.columns(2)

    with coll1:
      st.markdown(
          """
            <div class="module-card">
                <h4 style="color: #facc15; margin-top: 0; font-size: 16px;">Generator Modul Ajar (E-Modul Ajar)</h4>
                <p style="color: #94a3b8; font-size: 12px; margin-bottom: 0;">Otomatisasi perancangan Modul Ajar dengan kecerdasan buatan.</p>
            </div>
            """,
          unsafe_allow_html=True,
      )

    with coll2:
      st.markdown(
          """
            <div class="module-card">
                <h4 style="color: #38bdf8; margin-top: 0; font-size: 16px;">Modul Lainnya (E Presensi Siswa, E Jurnal, dll)</h4>
                <p style="color: #94a3b8; font-size: 12px; margin-bottom: 0;">Persiapan modul tambahan untuk sistem administrasi.</p>
            </div>
            """,
          unsafe_allow_html=True,
      )

  elif menu_pilihan == "🤖 E-Modul Ajar PM":
    st.markdown(
        "### 🤖 Otomatisasi Penyusunan Modul Ajar PM", unsafe_allow_html=True
    )

    # Menu/Expander Panduan Pembuatan Google Gemini API Key dengan Link Drive
    with st.expander(
        "📖 Panduan Pembuatan Kode Google Gemini API Key", expanded=True
    ):
      st.markdown(
          """
            <div style='background-color: #1e293b; padding: 15px; border-radius: 8px; border: 1px solid #334155;'>
                <p style='color: #cbd5e1; font-size: 13px; margin-bottom: 10px;'>
                    Untuk menggunakan modul generator berbasis AI ini, Anda memerlukan <b>Google Gemini API Key</b> pribadi yang bersifat gratis. 
                    Silakan pelajari panduan lengkap melalui tautan Google Drive di bawah ini:
                </p>
                <a href="https://drive.google.com/drive/folders/YOUR_DRIVE_FOLDER_ID" target="_blank" 
                   style='display: inline-block; background: #38bdf8; color: #0f172a; padding: 8px 16px; border-radius: 6px; font-weight: 600; text-decoration: none; font-size: 13px;'>
                   📂 Buka Dokumen Panduan Google Drive (PDF)
                </a>
                <hr style='border-color: #334155; margin: 12px 0;'>
                <p style='color: #94a3b8; font-size: 12px; margin-bottom: 4px;'><b>Ringkasan Singkat Langkah-langkah:</b></p>
                <ol style='color: #cbd5e1; font-size: 12px; margin: 0; padding-left: 20px;'>
                    <li>Akses situs <a href="https://aistudio.google.com/app/apikey" target="_blank" style='color: #38bdf8;'>Google AI Studio API Keys</a>.</li>
                    <li>Klik tombol <b>Create API key</b> di pojok kanan atas.</li>
                    <li>Pilih project yang tersedia, lalu klik <b>Create key</b>.</li>
                    <li>Salin kode kunci (<code>Copy key</code>) yang muncul dan tempelkan ke kolom di bawah ini.</li>
                </ol>
            </div>
            """,
          unsafe_allow_html=True,
      )

    # Input Kolom Google Gemini API Key Mandiri
    st.session_state.gemini_api_key = st.text_input(
        "🔑 Masukkan Google Gemini API Key Anda",
        value=st.session_state.gemini_api_key,
        type="password",
        placeholder="Contoh: AIzaSy...",
        help=(
            "Masukkan kunci API Gemini Anda. Kunci ini bersifat aman dan hanya"
            " digunakan selama sesi aktif Anda."
        ),
    )

    if not st.session_state.gemini_api_key:
      st.warning(
          "⚠️ Mohon masukkan Google Gemini API Key Anda terlebih dahulu untuk"
          " mengaktifkan fitur pembuatan Modul Ajar."
      )
    else:
      st.success(
          "✅ Google Gemini API Key berhasil disematkan! Silakan isi parameter"
          " pembelajaran di bawah."
      )

      # Form Parameter Pembelajaran Modul Ajar
      with st.form("form_parameter_modul"):
        st.markdown("#### ⚙️ Parameter Pembelajaran")
        col_a, col_b = st.columns(2)

        with col_a:
          jenjang = st.selectbox(
              "Pilih Jenjang Pendidikan", ["SMK / MAK", "SMA / MA", "SMP"]
          )
          mapel = st.text_input(
              "Mata Pelajaran / Program Kejuruan",
              value=(
                  "Dasar-dasar Teknik Otomotif / Produk Kreatif dan Kewirausahaan"
              ),
          )
          fase_kelas = st.selectbox(
              "Fase / Kelas",
              [
                  "Fase E / Kelas X SMK (Program Dasar Keahlian)",
                  "Fase F / Kelas XI SMK",
                  "Fase F / Kelas XII SMK",
              ],
          )

        with col_b:
          topik = st.text_input(
              "Topik / Materi Pokok / Elemen",
              value="Menyimak Teks Laporan Hasil Observasi (KIKL)",
          )
          alokasi = st.text_input("Alokasi Waktu", value="2 JP (2 x 45 Menit)")
          pertemuan = st.selectbox(
              "Pertemuan Ke-", ["1 (Pertemuan Pertama)", "2 (Pertemuan Kedua)"]
          )

        btn_generate = st.form_submit_button(
            "🚀 Buat Modul Ajar dengan AI", use_container_width=True
        )

        if btn_generate:
          if not st.session_state.gemini_api_key:
            st.error("❌ Google Gemini API Key belum dimasukkan!")
          else:
            st.info(
                "⏳ Menghubungkan ke Gemini AI menggunakan API Key mandiri..."
            )
            # Logika pemanggilan API atau proses generator selanjutnya dapat diletakkan di sini

  else:
    # Modul Lainnya (Placeholder)
    st.markdown(f"### 📋 {menu_pilihan}")
    st.info(
        "Modul ini sedang dalam tahap persiapan dan akan segera terintegrasi"
        " dengan sistem pusat."
    )
