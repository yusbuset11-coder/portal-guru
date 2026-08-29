import streamlit as st


def apply_global_styles():
  st.markdown(
      """
        <style>
            /* --- KUSTOMISASI SIDEBAR SERAGAM & KONTRAS TINGGI --- */
            [data-testid="stSidebar"] {
                background-color: #07090e !important; /* Lebih gelap & kontras dari background utama */
                border-right: 2px solid #334155 !important; /* Garis pemisah tegas dengan area utama */
            }

            /* --- KOTAK PROFIL / USER BOX SERAGAM --- */
            .user-profile-box {
                background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
                border: 1px solid #334155;
                padding: 12px;
                border-radius: 8px;
                text-align: center;
                margin-bottom: 1rem;
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
                color: #f0f6fc !important;
            }
        </style>
        """,
      unsafe_allow_html=True,
  )
