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

# Atur jarak bawah agar tombol login tidak terpotong
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 10rem;
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

# --- STYLING CSS TAMBAHAN (HEADER & KOMPONEN KHUSUS) ---
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main-header-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        padding: 16px 20px;
        border-radius: 12px;
        border: 1px solid #334155;
        margin-bottom: 15px;
        box-shadow: 0 5px 15px -3px rgba(0, 0, 0, 0.4);
        text-align: center;
    }
    .main-title {
        color: #38bdf8;
        font-size: 32px;
        font-weight: 800;
        margin: 0;
        letter-spacing: 1.2px;
        text-shadow: 0 0 10px rgba(56, 189, 248, 0.4);
    }
    .sub-title-1 {
        color: #f8fafc;
        font-size: 16px;
        font-weight: 700;
        margin-top: 4px;
        margin-bottom: 2px;
    }
    .sub-title-2 {
        color: #94a3b8;
        font-size: 13px;
        font-weight: 500;
        margin-top: 0;
        margin-bottom: 8px;
    }
    .dev-badge {
        display: inline-block;
        background: rgba(56, 189, 248, 0.1);
        color: #38bdf8;
        padding: 4px 14px;
        border-radius: 15px;
        font-size: 12px;
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
        padding: 0.6rem 1rem;
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

# --- HEADER UTAMA PORTAL PASTI ---
st.markdown(
    """
    <div class="main-header-card">
        <h1 class="main-title">PASTI</h1>
        <div class="sub-title-1">Portal Akademik Siswa Terintegrasi</div>
        <div class="sub-title-2">Pusat Kendali Aplikasi Pembelajaran dan Administrasi Guru</div>
        <div class="dev-badge">
            <b>Pengembang:</b> Yustinus Budi Setyanta - PS Cabdin Bangkalan &nbsp;|&nbsp; <em>Sistem Otomatisasi Terintegrasi</em>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- HALAMAN LOGIN / VERIFIKASI ---
if not st.session_state.logged_in:
  with st.form("form_login"):
    st.markdown(
        "<h3 style='color: #38bdf8; margin-top:0;'>🔐 Login Akses Portal</h3>",
        unsafe_allow_html=True,
    )
    st.write(
        "Silakan masukkan **Email** terdaftar atau **Token Unik** Anda untuk"
        " masuk ke sistem."
    )
    user_input = st.text_input(
        "Email / Token Unik Guru",
        placeholder=(
            "Contoh: yustinussetyanta08@dinas.belajar.id atau TOKEN300869"
        ),
    )
    btn_login = st.form_submit_button("🚀 Masuk Portal")

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
else:
  # --- SIDEBAR PROFESIONAL SETELAH LOGIN ---
    with st.sidebar:
      st.markdown(
          f"""
            <div class="user-profile-box">
                <span style="font-size: 24px;">👨‍💻</span><br>
                <b style="color: #facc15; font-size: 14px;">{st.session_state.guru_nama}</b><br>
                <span style="color: #94a3b8; font-size: 11px;">Sesi Aktif & Terverifikasi</span>
            </div>
            """,
          unsafe_allow_html=True,
      )

      st.markdown("---")
      st.markdown(
          """
            <p style='color: #94a3b8; font-size: 11px; font-weight: 600;'
            letter-spacing: 0.5px;'>NAVIGASI MODUL</p>
            """,
          unsafe_allow_html=True,
      )

    if st.button("🚪 Keluar / Logout"):
      st.session_state.logged_in = False
      st.session_state.guru_nama = ""
      st.session_state.spreadsheet_id = ""
      st.rerun()

  # --- TAMPILAN BERANDA UTAMA SETELAH LOGIN ---
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

  st.markdown("### 📚 Modul Aplikasi Tersedia")

  col1, col2 = st.columns(2)
  with col1:
    st.markdown(
        """
            <div class="module-card">
                <h4 style="color: #facc15; margin-top: 0; font-size: 16px;">Generator Modul Ajar (E-Modul Ajar PM)</h4>
                <p style="color: #94a3b8; font-size: 12px; margin-bottom: 0;">Otomatisasi perancangan Modul Ajar Pembelajaran Mendalam (Deep Learning) lengkap dengan LKM, Rubrik, dan Instrumen Formatif berformat Word (.docx).</p>
            </div>
            """,
        unsafe_allow_html=True,
    )
  with col2:
    st.markdown(
        """
            <div class="module-card">
                <h4 style="color: #38bdf8; margin-top: 0; font-size: 16px;">Modul Lainnya (E Presensi Siswa, E Jurnal Mengajar, E-Asesmen PM)</h4>
                <p style="color: #94a3b8; font-size: 12px; margin-bottom: 0;">Persiapan modul tambahan untuk sistem penilaian siswa, manajemen data pembelajaran, dan administrasi pendukung profesional lainnya.</p>
            </div>
            """,
        unsafe_allow_html=True,
    )
