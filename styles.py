import streamlit as st


def apply_global_styles():
    st.markdown(
        """
        <style>
            /* --- SIDEBAR UTAMA (SERAGAM & KONTRAST) --- */
            [data-testid="stSidebar"] {
                background-color: #07090e !important;
                border-right: 2px solid #1e293b !important;
            }
            
            /* Mengatur kontainer sidebar menjadi Flexbox agar posisi elemen bisa diatur */
            [data-testid="stSidebar"] > div:first-child {
                display: flex;
                flex-direction: column;
            }
            
            /* Memindahkan menu navigasi bawaan Streamlit ke bagian paling bawah */
            [data-testid="stSidebarNav"] {
                order: 99 !important;
            }

            /* --- KOTAK PROFIL / USER BOX SERAGAM (GLOBAL) --- */
            .user-profile-box {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%) !important;
                border: 1px solid #334155 !important;
                padding: 14px !important;
                border-radius: 10px !important;
                text-align: center !important;
                margin-bottom: 20px !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3) !important;
            }
            .user-profile-box span {
                font-size: 22px !important;
            }
            .user-profile-box b {
                color: #facc15 !important;
                font-size: 14px !important;
            }
            .user-profile-box small {
                color: #94a3b8 !important;
                font-size: 11px !important;
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
