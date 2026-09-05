import streamlit as st


def apply_global_styles():
  st.markdown(
      """
        <style>
            /* --- SIDEBAR UTAMA --- */
            [data-testid="stSidebar"] {
                background-color: #07090e !important;
                border-right: 2px solid #1e293b !important;
            }
            
            /* Sembunyikan menu navigasi otomatis bawaan Streamlit */
            [data-testid="stSidebarNav"] {
                display: none !important;
            }

            /* --- KOTAK PROFIL / USER BOX --- */
            .user-profile-box {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border: 1px solid #334155 !important;
                padding: 14px !important;
                border-radius: 10px !important;
                text-align: center !important;
                margin: 15px 0 !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
            }
            .user-profile-box span { font-size: 22px !important; }
            .user-profile-box b { color: #facc15 !important; font-size: 14px !important; }
            .user-profile-box small { color: #94a3b8 !important; font-size: 11px !important; }

            /* --- KARTU MODUL & KONTEN --- */
            .module-card {
                background-color: #161b22;
                border: 1px solid #30363d;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }

            /* --- WARNA TEKS UTAMA --- */
            h1, h2, h3, h4 {
                color: #ffffff !important;
            }
        </style>
        """,
      unsafe_allow_html=True,
  )


def render_sidebar():
  """Fungsi terpusat untuk menampilkan sidebar di setiap halaman"""
  with st.sidebar:
    # 1. Logo + Teks "PASTI"
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 10px; padding-bottom: 10px;">
            <img src="https://lh3.googleusercontent.com/d/15rUWzaqM_86lF2ht8atJmmyPocUPxl_z" style="width: 32px; height: auto;">
            <span style="color: #38bdf8; font-size: 20px; font-weight: bold; letter-spacing: 1px;">PASTI</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        "<hr style='margin: 5px 0 15px 0; border-color: #1e293b;'>",
        unsafe_allow_html=True,
    )

    # 2. Urutan Halaman (Menu Navigasi Manual)
    st.page_link("app.py", label="App (Beranda)", icon="🏠")
    st.page_link(
        "pages/1_E_Presensi_Siswa.py", label="E Presensi Siswa", icon="📋"
    )
    st.page_link(
        "pages/2_E_Jurnal_Mengajar.py", label="E Jurnal Mengajar", icon="📖"
    )
    st.page_link("pages/3_E_Asesmen_PM.py", label="E Asesmen PM", icon="📊")
    st.page_link("pages/4_E_Modul_Ajar_PM.py", label="E Modul Ajar PM", icon="📑")

    st.markdown(
        "<hr style='margin: 15px 0; border-color: #1e293b;'>",
        unsafe_allow_html=True,
    )

    # 3. Profil & Tombol Log Out (Hanya tampil jika sudah login)
    if st.session_state.get("logged_in", False):
      st.markdown(
          f"""
            <div class="user-profile-box">
                <span style="font-size: 24px;">👨‍💻</span><br>
                <b style="color: #facc15; font-size: 14px;">{st.session_state.get('guru_nama', 'Guru')}</b><br>
                <span style="color: #94a3b8; font-size: 11px;">Sesi Aktif & Terverifikasi</span>
            </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.guru_nama = ""
        st.session_state.spreadsheet_id = ""
        st.switch_page("app.py")
