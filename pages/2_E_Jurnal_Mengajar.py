from datetime import datetime
import io
import gspread
import pandas as pd
import streamlit as str_lit  # menggunakan alias atau langsung st
import streamlit as st
from google.oauth2.service_account import Credentials
from styles import apply_global_styles

# Konfigurasi Halaman
st.set_page_config(
    page_title="E Jurnal Mengajar - Digitalisasi Jurnal Mengajar Guru",
    page_icon="📖",
    layout="wide",
)
apply_global_styles()

st.markdown(
    "<style>.block-container { padding-top: 1.5rem; padding-bottom: 1rem;"
    " }</style>",
    unsafe_allow_html=True,
)

# Cek apakah guru sudah login melalui portal utama (app.py)
if "logged_in" not in st.session_state or not st.session_state.logged_in:
    st.warning(
        "⚠️ Anda belum login. Silakan kembali ke Halaman Utama (`app.py`) untuk"
        " masuk ke portal terlebih dahulu."
    )
    st.stop()


def get_gspread_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)


# Fungsi Border & Format Header Spreadsheet (7 Kolom)
def apply_sheet_formatting(ws, num_rows, num_cols=7):
    if num_rows < 1 or num_cols < 1:
        return
    try:
        col_letter = gspread.utils.rowcol_to_a1(num_rows, num_cols)
        cell_range = f"A1:{col_letter}"

        border_format = {
            "top": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
            "bottom": {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            },
            "left": {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            },
            "right": {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            },
            "innerHorizontal": {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            },
            "innerVertical": {
                "style": "SOLID",
                "color": {"red": 0, "green": 0, "blue": 0},
            },
        }

        header_format = {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.9, "green": 0.95, "blue": 1.0},
        }

        ws.format(cell_range, {"borders": border_format})
        ws.format(
            f"A1:{gspread.utils.rowcol_to_a1(1, num_cols)}", header_format
        )
    except Exception as e:
        print(f"Gagal menerapkan format sheet: {e}")


# Header standar (Tanpa Guru_Pengajar - total 7 kolom)
EXPECTED_JURNAL_HEADER = [
    "Tanggal",
    "Sekolah",
    "Kelas",
    "Jam_Pelajaran",
    "Mata_Pelajaran",
    "Materi_Pembelajaran",
    "Catatan_Kejadian",
]


def get_clean_jurnal_dataframe(ws):
    try:
        rows = ws.get_all_values()
    except Exception:
        rows = []

    if not rows:
        ws.append_row(EXPECTED_JURNAL_HEADER)
        return pd.DataFrame(columns=EXPECTED_JURNAL_HEADER)

    # Jika baris pertama bukan header yang benar, sisipkan header di baris 1
    if rows[0] != EXPECTED_JURNAL_HEADER:
        ws.insert_row(EXPECTED_JURNAL_HEADER, index=1)
        rows = ws.get_all_values()

    if len(rows) > 1:
        df = pd.DataFrame(rows[1:], columns=rows[0])
        # Bersihkan kolom jika ada kelebihan/kekurangan kolom di sheet lama
        for col in EXPECTED_JURNAL_HEADER:
            if col not in df.columns:
                df[col] = ""
        df = df[EXPECTED_JURNAL_HEADER]  # Batasi hanya 7 kolom utama
        return df
    else:
        return pd.DataFrame(columns=EXPECTED_JURNAL_HEADER)


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


@st.cache_resource
def load_guru_database(sheet_id):
    try:
        client = get_gspread_client()
        return client.open_by_key(sheet_id)
    except Exception:
        return None


spreadsheet_id = st.session_state.get("spreadsheet_id", "")

if not spreadsheet_id:
    st.warning(
        "⚠️ Spreadsheet ID belum dikonfigurasi. Silakan kembali ke Halaman"
        " Utama untuk mengatur konfigurasi terlebih dahulu."
    )
    st.stop()

sh_guru = load_guru_database(spreadsheet_id)

if sh_guru is None:
    st.error(
        "Gagal terhubung ke Database Guru Anda. Periksa kembali ID Spreadsheet"
        " di Master Registry."
    )
else:
    menu_jurnal = st.sidebar.selectbox(
        "**Pilih Menu Jurnal**",
        [
            "🏠 Beranda Jurnal",
            "✍️ Input Jurnal Mengajar",
            "📚 Riwayat & Rekap Jurnal",
        ],
    )

    guru_nama = st.session_state.get("guru_nama", "")

    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, #1e293b 0%, #111827 100%); padding: 25px 30px; border-radius: 14px; border: 1px solid #334155; margin-bottom: 25px;">
            <h2 style="color: #ffffff; margin: 0 0 10px 0; font-size: 30px;">📖 Digitalisasi Jurnal Pembelajaran</h2>
            <p style="color: #e2e8f0; font-size: 15px; margin: 0 0 8px 0;">Selamat Datang, {guru_nama} di Modul E Jurnal Mengajar</p>
            <p style="color: #94a3b8; font-size: 14px; margin: 0;">Catat kegiatan pembelajaran harian, ketercapaian, dan laporan dengan mudah.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if menu_jurnal == "🏠 Beranda Jurnal":
        try:
            ws_jurnal = sh_guru.worksheet("Jurnal Mengajar")
            df_jurnal = get_clean_jurnal_dataframe(ws_jurnal)
            total_jurnal = len(df_jurnal)
        except Exception:
            total_jurnal = 0

        col1, col2 = st.columns(2)
        with col1:
            with st.container(border=True):
                st.markdown(
                    f"""
                    <p style="color: #94a3b8; font-size: 12px; margin: 0 0 5px 0; text-transform: uppercase; font-weight: 600;">TOTAL JURNAL TERCATAT</p>
                    <p style="color: #ffffff; font-size: 22px; margin: 0; font-weight: 600;">{total_jurnal} Kegiatan</p>
                    """,
                    unsafe_allow_html=True,
                )
        with col2:
            with st.container(border=True):
                st.markdown(
                    """
                    <p style="color: #94a3b8; font-size: 12px; margin: 0 0 5px 0; text-transform: uppercase; font-weight: 600;">STATUS KONEKSI DATABASE</p>
                    <p style="color: #ffffff; font-size: 22px; margin: 0; font-weight: 600;">Terhubung Aktif ✅</p>
                    """,
                    unsafe_allow_html=True,
                )

    elif menu_jurnal == "✍️ Input Jurnal Mengajar":
        st.markdown(
            """
            <div class="digma-banner">
                <h3>✍️ Pencatatan Jurnal Kegiatan Mengajar</h3>
                <p>Silakan isi rincian pelaksanaan kegiatan belajar mengajar di kelas secara lengkap.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        try:
            sheet_siswa = sh_guru.worksheet("Data Kelas-Siswa")
            data_siswa = sheet_siswa.get_all_records()
            df_siswa = pd.DataFrame(data_siswa)
            list_sekolah = (
                df_siswa["Sekolah"].dropna().unique().tolist()
                if not df_siswa.empty and "Sekolah" in df_siswa.columns
                else ["SMK Negeri 2 Bangkalan"]
            )
            list_kelas = (
                df_siswa["Kelas"].dropna().unique().tolist()
                if not df_siswa.empty and "Kelas" in df_siswa.columns
                else ["X PPLG 1", "XI PPLG 1", "XII PPLG 1"]
            )
        except Exception:
            list_sekolah = ["SMK Negeri 2 Bangkalan"]
            list_kelas = ["X PPLG 1", "XI PPLG 1", "XII PPLG 1"]

        with st.container(border=True):
            with st.form("form_jurnal_mengajar"):
                col_1, col_2 = st.columns(2)
                with col_1:
                    tanggal_jurnal = st.date_input(
                        "**Tanggal Mengajar**", datetime.today()
                    )
                    pilih_sekolah = st.selectbox("**Sekolah**", list_sekolah)
                with col_2:
                    jam_pelajaran = st.text_input(
                        "**Jam Pelajaran / JP**",
                        placeholder="Contoh: Jam ke 1 - 3 (07.00 - 09.30)",
                    )
                    pilih_kelas = st.selectbox("**Kelas**", list_kelas)

                mata_pelajaran = st.text_input(
                    "**Mata Pelajaran**",
                    placeholder="Contoh: Pemrograman Berorientasi Objek",
                )
                kompetensi_dasar = st.text_area(
                    "**Materi / Kompetensi Dasar / Tujuan Pembelajaran**",
                    placeholder=(
                        "Tuliskan materi atau tujuan pembelajaran yang disampaikan..."
                    ),
                )
                catatan_kejadian = st.text_area(
                    "**Catatan Kejadian / Refleksi Kelas** *(Opsional)*",
                    placeholder=(
                        "Catatan penting siswa atau kendala selama KBM..."
                    ),
                )

                btn_simpan_jurnal = st.form_submit_button(
                    "💾 **Simpan Jurnal Mengajar**"
                )

                if btn_simpan_jurnal:
                    if not mata_pelajaran or not kompetensi_dasar:
                        st.error(
                            "⚠️ Mata Pelajaran dan Materi/Tujuan Pembelajaran wajib"
                            " diisi!"
                        )
                    else:
                        with st.spinner(
                            "Menyimpan jurnal ke Google Sheets..."
                        ):
                            try:
                                ws_jurnal = sh_guru.worksheet(
                                    "Jurnal Mengajar"
                                )
                            except Exception:
                                ws_jurnal = sh_guru.add_worksheet(
                                    title="Jurnal Mengajar",
                                    rows="1000",
                                    cols="7",
                                )

                            # Pastikan header bersih ada di baris 1
                            rows_check = ws_jurnal.get_all_values()
                            if (
                                not rows_check
                                or rows_check[0] != EXPECTED_JURNAL_HEADER
                            ):
                                ws_jurnal.insert_row(
                                    EXPECTED_JURNAL_HEADER, index=1
                                )

                            # Append baris data (7 kolom tanpa Guru_Pengajar)
                            ws_jurnal.append_row([
                                str(tanggal_jurnal),
                                str(pilih_sekolah),
                                str(pilih_kelas),
                                str(jam_pelajaran),
                                str(mata_pelajaran),
                                str(kompetensi_dasar),
                                str(catatan_kejadian),
                            ])

                            all_jurnal_data = ws_jurnal.get_all_values()
                            apply_sheet_formatting(
                                ws_jurnal, len(all_jurnal_data), 7
                            )

                            st.balloons()
                            st.success(
                                "🎉 Jurnal mengajar berhasil disimpan ke database"
                                " Anda!"
                            )

    elif menu_jurnal == "📚 Riwayat & Rekap Jurnal":
        st.markdown(
            """
            <div class="digma-banner">
                <h3>📚 Riwayat Jurnal Mengajar</h3>
                <p>Daftar seluruh jurnal kegiatan mengajar yang telah Anda catat beserta fitur filter dan rekapitulasi.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        try:
            ws_jurnal = sh_guru.worksheet("Jurnal Mengajar")
            df_jurnal = get_clean_jurnal_dataframe(ws_jurnal)

            if df_jurnal.empty or len(df_jurnal) == 0:
                st.info("Belum ada data jurnal mengajar yang tersimpan.")
            else:
                with st.container(border=True):
                    st.markdown("#### **🔍 Filter Riwayat Jurnal**")
                    col_f1, col_f2 = st.columns(2)

                    schools_j = (
                        df_jurnal["Sekolah"].unique().tolist()
                        if "Sekolah" in df_jurnal.columns
                        else []
                    )
                    sel_sch_j = col_f1.selectbox(
                        "**🏫 Filter Sekolah**", ["Semua Sekolah"] + schools_j
                    )

                    if sel_sch_j != "Semua Sekolah":
                        classes_j = (
                            df_jurnal[df_jurnal["Sekolah"] == sel_sch_j]["Kelas"]
                            .unique()
                            .tolist()
                            if "Kelas" in df_jurnal.columns
                            else []
                        )
                    else:
                        classes_j = (
                            df_jurnal["Kelas"].unique().tolist()
                            if "Kelas" in df_jurnal.columns
                            else []
                        )
                    sel_cls_j = col_f2.selectbox(
                        "**📚 Filter Kelas**", ["Semua Kelas"] + classes_j
                    )

                df_filtered_j = df_jurnal.copy()
                if sel_sch_j != "Semua Sekolah":
                    df_filtered_j = df_filtered_j[
                        df_filtered_j["Sekolah"] == sel_sch_j
                    ]
                if sel_cls_j != "Semua Kelas":
                    df_filtered_j = df_filtered_j[
                        df_filtered_j["Kelas"] == sel_cls_j
                    ]

                # Mengatur agar nomor urut/indeks tabel dimulai dari angka 1
                if not df_filtered_j.empty:
                    df_filtered_j.index = range(1, len(df_filtered_j) + 1)

                with st.container(border=True):
                    st.dataframe(df_filtered_j, use_container_width=True)

                # Tombol Download Excel Jurnal
                output_jurnal = io.BytesIO()
                with pd.ExcelWriter(output_jurnal, engine="openpyxl") as writer:
                    df_filtered_j.to_excel(
                        writer, index=True, sheet_name="Jurnal_Mengajar"
                    )
                excel_jurnal_data = output_jurnal.getvalue()

                st.download_button(
                    label="📥 **Download Riwayat Jurnal (Excel)**",
                    data=excel_jurnal_data,
                    file_name="Rekap_Jurnal_Mengajar.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )
        except Exception as e:
            st.info(
                "Tab 'Jurnal Mengajar' belum tersedia atau belum ada data di"
                f" Spreadsheet Anda. (Detail: {e})"
            )
