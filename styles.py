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
            
            /* Mengubah kontainer sidebar menjadi Flexbox vertikal */
            [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
                display: flex;
                flex-direction: column;
            }

            /* --- PENGATURAN URUTAN (ORDER) DI SIDEBAR --- */
            
            /* 1. Logo + Teks PASTI */
            .sidebar-logo-container {
                order: 1 !important;
            }

            /* 2. Urutan Halaman (Menu Navigasi Bawaan Streamlit) */
            [data-testid="stSidebarNav"] {
                order: 2 !important;
            }

            /* 3. Kotak Profil */
            .user-profile-box {
                order: 3 !important;
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border: 1px solid #334155 !important;
                padding: 14px !important;
                border-radius: 10px !important;
                text-align: center !important;
                margin-bottom: 10px !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
            }
            .user-profile-box span { font-size: 22px !important; }
            .user-profile-box b { color: #facc15 !important; font-size: 14px !important; }
            .user-profile-box small { color: #94a3b8 !important; font-size: 11px !important; }

            /* 4. Menu Navigasi Tambahan (Jika ada) */
            .extra-nav {
                order: 4 !important;
            }

            /* 5. Tombol Log Out */
            .logout-container {
                order: 5 !important;
                margin-top: 10px !important;
            }

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
