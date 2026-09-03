from datetime import datetime
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import streamlit as st
from styles import apply_global_styles

st.set_page_config(
    page_title="E Presensi Siswa - Sistem Informasi Presensi Harian",
    page_icon="📋",
    layout="wide",
)
apply_global_styles()

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 1rem;
        }
        .stApp {
            background-color: #0e1117;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .main-header-card {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            padding: 18px 22px;
            border-radius: 12px;
            border: 1px solid #334155;
            margin-bottom: 20px;
            box-shadow: 0 8px 20px -5px rgba(0, 0, 0, 0.4);
        }
        .main-title {
            color: #38bdf8;
            font-size: 16px;
            font-weight: 700;
            margin: 0 0 6px 0;
            line-height: 1.4;
        }
        .sub-desc {
            color: #94a3b8;
            font-size: 13px;
            margin: 0;
        }
        .stat-card {
            background: #111827;
            padding: 18px;
            border-radius: 10px;
            border: 1px solid #1f2937;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }
        [data-testid="stSidebar"] {
            background-color: #111827;
            border-right: 1px solid #1f2937;
        }
        
        /* CSS PERBAIKAN: Spasi Baris Padat & Line Border Jelas */
        div[data-testid="stForm"] div[data-testid="stHorizontalBlock"] {
            border-bottom: 1px solid #2d3748 !important;
            padding-top: 2px !important;
            padding-bottom: 2px !important;
            margin-bottom: 0px !important;
            align-items: center !important;
        }
        div[data-testid="stForm"] div[data-testid="stColumn"] {
            padding-top: 0px !important;
            padding-bottom: 0px !important;
        }
        div[data-testid="stForm"] div[data-testid="stRadio"] {
            margin-top: -10px !important;
            margin-bottom: -10px !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Cek Login
if "logged_in" not in st.session_state or not st.session_state.logged_in:
  st.warning("⚠️ Silakan login terlebih dahulu melalui halaman utama (app.py).")
  st.stop()

spreadsheet_id = st.session_state.get("spreadsheet_id", "")
guru_nama = st.session_state.get("guru_nama", "Yusbuset")

# Header Utama E Presensi Siswa
st.markdown(
    f"""
    <div class="main-header-card">
        <h2 class="main-title">📋 Digitalisasi Presensi Harian Siswa</h2>
        <p class="sub-desc">Selamat Datang, {guru_nama} di Modul E Presensi Siswa</p>
        <p class="sub-desc" style="margin-top: 4px;">Kelola dan pantau kehadiran siswa berdasarkan Sekolah, Kelas, dan Mata Pelajaran.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Sidebar Profil & Navigasi
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


# Koneksi Google Sheets
@st.cache_resource
def get_gspread_client():
  scope = [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
  ]
  creds_dict = dict(st.secrets["gcp_service_account"])
  creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
  return gspread.authorize(creds)


client = get_gspread_client()


def load_data_from_sheet(sheet_name):
  try:
    sh = client.open_by_key(spreadsheet_id)
    try:
      worksheet = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
      worksheet = sh.add_worksheet(title=sheet_name, rows=100, cols=20)
      if sheet_name == "Data Kelas-Siswa":
        worksheet.append_row(["Sekolah", "Kelas", "No_Absen", "Nama_Siswa"])
      elif sheet_name == "Absensi Siswa":
        worksheet.append_row([
            "Tanggal",
            "Sekolah",
            "Kelas",
            "Mapel",
            "Jam",
            "No_Absen",
            "Nama_Siswa",
            "Status",
        ])
      elif sheet_name in ["Rekap Semester Ganjil", "Rekap Semester Genap"]:
        worksheet.append_row([
            "Sekolah",
            "Kelas",
            "Mapel",
            "No_Absen",
            "Nama_Siswa",
            "Hadir",
            "Izin",
            "Sakit",
            "Alpa",
            "Dispensasi",
            "Jumlah_TH",
        ])

    if sheet_name == "Absensi Siswa":
      expected_header = [
          "Tanggal",
          "Sekolah",
          "Kelas",
          "Mapel",
          "Jam",
          "No_Absen",
          "Nama_Siswa",
          "Status",
      ]
      current_header = worksheet.row_values(1)
      if current_header != expected_header:
        if not current_header:
          worksheet.append_row(expected_header)
        else:
          worksheet.update(range_name="A1:H1", values=[expected_header])

    data = worksheet.get_all_records()
    return pd.DataFrame(data), worksheet
  except Exception as e:
    return pd.DataFrame(), None


# Navigasi Menu Dalam Halaman Presensi
menu = st.radio(
    "Navigasi Modul Presensi",
    [
        "📝 Pencatatan Presensi Harian",
        "📊 Rekap Semester Ganjil & Genap",
        "📂 Manajemen Data Sekolah & Siswa",
    ],
    horizontal=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# ----------------------------------------------------
# MENU 1: PENCATATAN PRESENSI HARIAN
# ----------------------------------------------------
if menu == "📝 Pencatatan Presensi Harian":
  st.markdown("### 📝 Pencatatan Presensi Harian Siswa")
  st.markdown(
      "Kelola dan pantau kehadiran siswa berdasarkan Sekolah, Kelas, dan Mata"
      " Pelajaran."
  )

  df_siswa, ws_siswa = load_data_from_sheet("Data Kelas-Siswa")

  if df_siswa.empty or "Sekolah" not in df_siswa.columns:
    st.warning(
        "⚠️ Belum ada data Sekolah dan Siswa. Silakan isi atau upload data"
        ' melalui menu **Manajemen Data Sekolah & Siswa** di bawah.'
    )
  else:
    list_sekolah = df_siswa["Sekolah"].dropna().unique().tolist()

    # BARIS PERTAMA: Tanggal Presensi - Kelas - Sekolah (3 Kolom)
    col1, col2, col3 = st.columns(3)
    with col1:
      tanggal = st.date_input("Tanggal Presensi", datetime.now())
    with col3:
      selected_sekolah = st.selectbox(
          "Sekolah", list_sekolah if list_sekolah else ["-"]
      )

    df_filtered_sekolah = df_siswa[df_siswa["Sekolah"] == selected_sekolah]
    list_kelas = (
        df_filtered_sekolah["Kelas"].dropna().unique().tolist()
        if not df_filtered_sekolah.empty
        else ["-"]
    )
    with col2:
      selected_kelas = st.selectbox("Kelas", list_kelas)

    # BARIS KEDUA: Mata Pelajaran - Jam Pelajaran / Keterangan Waktu (2 Kolom)
    col_m1, col_m2 = st.columns(2)
    with col_m1:
      mapel = st.text_input(
          "Mata Pelajaran", placeholder="Contoh: Informatika / Bahasa Indonesia"
      )
    with col_m2:
      jam = st.text_input(
          "Jam Pelajaran / Keterangan Waktu", placeholder="Contoh: Jam ke 1-2"
      )

    st.markdown("---")
    st.subheader(
        f"Daftar Kehadiran Siswa - {selected_sekolah} ({selected_kelas})"
    )

    df_current_students = df_filtered_sekolah[
        df_filtered_sekolah["Kelas"] == selected_kelas
    ]

    if df_current_students.empty:
      st.info(
          "ℹ️ Tidak ada data siswa untuk kombinasi Sekolah dan Kelas tersebut."
      )
    else:
      if "No_Absen" in df_current_students.columns:
        df_current_students = df_current_students.sort_values(by="No_Absen")

      form_key = f"form_presensi_{selected_sekolah}_{selected_kelas}"
      with st.form(form_key):
        attendance_results = {}

        header_cols = st.columns([1, 3, 5])
        header_cols[0].markdown("**No**")
        header_cols[1].markdown("**Nama Siswa**")
        header_cols[2].markdown("**Status Kehadiran**")
        st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)

        for idx, row in df_current_students.reset_index(drop=True).iterrows():
          no_absen = row.get("No_Absen", idx + 1)
          nama = row.get("Nama_Siswa", "Tanpa Nama")

          r_cols = st.columns([1, 3, 5])
          r_cols[0].write(str(no_absen))
          r_cols[1].write(str(nama))

          status = r_cols[2].radio(
              f"Status {nama}",
              ["Hadir", "Izin", "Sakit", "Alpa", "Dispensasi"],
              horizontal=True,
              key=f"status_{selected_sekolah}_{selected_kelas}_{idx}",
              label_visibility="collapsed",
          )
          attendance_results[nama] = {"no_absen": no_absen, "status": status}

        st.markdown("")
        submitted = st.form_submit_button("💾 Simpan Presensi ke Spreadsheet")

        if submitted:
          if not mapel:
            st.error("⚠️ Mohon isi Mata Pelajaran terlebih dahulu.")
          else:
            _, ws_absensi = load_data_from_sheet("Absensi Siswa")
            rows_to_add = []
            tgl_str = tanggal.strftime("%Y-%m-%d")

            for nama, info in attendance_results.items():
              rows_to_add.append([
                  tgl_str,
                  selected_sekolah,
                  selected_kelas,
                  mapel,
                  jam,
                  info["no_absen"],
                  nama,
                  info["status"],
              ])

            if ws_absensi and rows_to_add:
              ws_absensi.append_rows(rows_to_add)
              st.balloons()
              st.success(
                  "🎉 Data presensi berhasil disimpan ke Google Spreadsheet"
                  " Anda!"
              )
            else:
              st.error("❌ Gagal menyimpan ke Spreadsheet.")

# ----------------------------------------------------
# MENU 2: REKAP SEMESTER GANJIL & GENAP
# ----------------------------------------------------
elif menu == "📊 Rekap Semester Ganjil & Genap":
  st.markdown("### 📊 Rekapitulasi Kehadiran Semester")
  st.markdown(
      "Pantau rekap kehadiran siswa berdasarkan Sekolah, Kelas, dan Mata"
      " Pelajaran."
  )

  tab_ganjil, tab_genap = st.tabs([
      "📚 Rekap Semester Ganjil (Juli - Desember)",
      "📖 Rekap Semester Genap (Januari - Juni)",
  ])

  df_absensi, _ = load_data_from_sheet("Absensi Siswa")


  def process_rekap(semester_type):
    if df_absensi.empty or "Tanggal" not in df_absensi.columns:
      st.info("ℹ️ Belum ada data presensi tercatat di Spreadsheet.")
      return

    try:
      df = df_absensi.copy()
      df["Bulan"] = pd.to_datetime(df["Tanggal"], errors="coerce").dt.month

      if semester_type == "Ganjil":
        df_sem = df[df["Bulan"].isin([7, 8, 9, 10, 11, 12])]
        sheet_target = "Rekap Semester Ganjil"
      else:
        df_sem = df[df["Bulan"].isin([1, 2, 3, 4, 5, 6])]
        sheet_target = "Rekap Semester Genap"

      if df_sem.empty:
        st.warning(f"⚠️ Belum ada data presensi untuk Semester {semester_type}.")
        return

      col1, col2, col3 = st.columns(3)
      with col1:
        list_sek = df_sem["Sekolah"].dropna().unique().tolist()
        sel_sek = st.selectbox(
            f"Pilih Sekolah ({semester_type})",
            list_sek if list_sek else ["-"],
            key=f"sek_{semester_type}",
        )

      with col2:
        df_s = df_sem[df_sem["Sekolah"] == sel_sek]
        list_kls = df_s["Kelas"].dropna().unique().tolist()
        sel_kls = st.selectbox(
            f"Pilih Kelas ({semester_type})",
            list_kls if list_kls else ["-"],
            key=f"kls_{semester_type}",
        )

      with col3:
        df_k = df_s[df_s["Kelas"] == sel_kls]
        list_mapel = (
            df_k["Mapel"].dropna().unique().tolist()
            if "Mapel" in df_k.columns
            else []
        )
        sel_mapel = st.selectbox(
            f"Pilih Mapel ({semester_type})",
            list_mapel if list_mapel else ["-"],
            key=f"mapel_{semester_type}",
        )

      df_filtered = (
          df_k[df_k["Mapel"] == sel_mapel] if sel_mapel != "-" else df_k
      )

      if df_filtered.empty:
        st.info("ℹ️ Tidak ada data presensi untuk kombinasi tersebut.")
        return

      rekap_list = []
      grouped = df_filtered.groupby(
          ["Sekolah", "Kelas", "Mapel", "No_Absen", "Nama_Siswa"]
      )

      for name, group in grouped:
        sekolah, kelas, mapel_val, no_absen, nama = name
        hadir = len(group[group["Status"] == "Hadir"])
        izin = len(group[group["Status"] == "Izin"])
        sakit = len(group[group["Status"] == "Sakit"])
        alpa = len(group[group["Status"] == "Alpa"])
        dispensasi = len(group[group["Status"] == "Dispensasi"])
        jumlah = izin + sakit + alpa + dispensasi

        rekap_list.append({
            "Sekolah": sekolah,
            "Kelas": kelas,
            "Mapel": mapel_val,
            "No_Absen": no_absen,
            "Nama_Siswa": nama,
            "Hadir": hadir,
            "Izin": izin,
            "Sakit": sakit,
            "Alpa": alpa,
            "Dispensasi": dispensasi,
            "Jumlah": jumlah,
        })

      df_rekap = pd.DataFrame(rekap_list)
      if not df_rekap.empty:
        df_rekap = df_rekap.sort_values(by="No_Absen")
        st.dataframe(df_rekap, use_container_width=True)

        if st.button(
            f"🔄 Sinkronkan Rekap {semester_type} ke Spreadsheet",
            key=f"sync_{semester_type}",
        ):
          _, ws_rekap = load_data_from_sheet(sheet_target)
          if ws_rekap:
            ws_rekap.clear()
            headers = [
                "Sekolah",
                "Kelas",
                "Mapel",
                "No_Absen",
                "Nama_Siswa",
                "Hadir",
                "Izin",
                "Sakit",
                "Alpa",
                "Dispensasi",
                "Jumlah",
            ]
            ws_rekap.append_row(headers)
            ws_rekap.append_rows(df_rekap.values.tolist())
            st.balloons()
            st.success(
                f"✅ Rekap Semester {semester_type} berhasil disinkronkan ke tab"
                f" '{sheet_target}' di Spreadsheet!"
            )

    except Exception as e:
      st.error(f"Terjadi kesalahan memproses rekap: {e}")

  with tab_ganjil:
    st.subheader("Rekapitulasi Kehadiran Semester Ganjil")
    process_rekap("Ganjil")

  with tab_genap:
    st.subheader("Rekapitulasi Kehadiran Semester Genap")
    process_rekap("Genap")

# ----------------------------------------------------
# MENU 3: MANAJEMEN DATA SEKOLAH & SISWA
# ----------------------------------------------------
elif menu == "📂 Manajemen Data Sekolah & Siswa":
  st.markdown("### 📂 Manajemen Data Sekolah, Kelas, & Siswa")
  st.markdown(
      "Download template atau upload daftar siswa Anda agar tersinkronisasi"
      " dengan Google Spreadsheet."
  )

  tab_dl, tab_ul = st.tabs(["📥 Download Template", "📤 Upload Data Siswa"])

  with tab_dl:
    st.subheader("Template Data Siswa")
    st.markdown(
        "Silakan unduh atau salin template resmi di bawah ini melalui Google"
        " Drive. Template ini sudah disesuaikan agar langsung siap diisi"
        " dengan format: **Sekolah - Kelas - No_Absen - Nama_Siswa**."
    )

    st.markdown(
        """
            <a href="https://docs.google.com/spreadsheets/d/1ioL8wtVFwf4EbubJuY8ontu3k0pBLRxS/edit?usp=sharing" target="_blank">
                <button style="background-color: #FF4B4B; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">
                    🔗 Buka & Download Template (Google Sheets)
                </button>
            </a>
            """,
        unsafe_allow_html=True,
    )

  with tab_ul:
    st.subheader("Upload Data Siswa (.xlsx)")
    uploaded_file = st.file_uploader(
        "Pilih file data siswa (.xlsx)", type=["xlsx"]
    )

    if uploaded_file is not None:
      try:
        df_upload = pd.read_excel(uploaded_file)

        st.write("Preview Data yang akan diunggah:")
        st.dataframe(df_upload.head())

        if st.button("🚀 Unggah ke Google Spreadsheet"):
          required_cols = ["Sekolah", "Kelas", "No_Absen", "Nama_Siswa"]
          if all(col in df_upload.columns for col in required_cols):
            _, ws_siswa = load_data_from_sheet("Data Kelas-Siswa")
            if ws_siswa:
              ws_siswa.clear()
              ws_siswa.append_row(required_cols)
              rows = df_upload[required_cols].values.tolist()
              ws_siswa.append_rows(rows)
              st.balloons()
              st.success(
                  "✅ Data berhasil diunggah dan disinkronkan ke Spreadsheet Anda!"
              )
              st.rerun()
          else:
            st.error(f"❌ Kolom pada file harus lengkap: {required_cols}")
      except Exception as e:
        st.error(f"Terjadi kesalahan membaca file: {e}")

  st.markdown("---")
  st.subheader("📋 Data Siswa yang Tersimpan di Spreadsheet Anda")
  df_existing, _ = load_data_from_sheet("Data Kelas-Siswa")
  if not df_existing.empty:
    st.dataframe(df_existing, use_container_width=True)
  else:
    st.info("Belum ada data tersimpan.")
