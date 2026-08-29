import streamlit as st


def apply_global_styles():
  st.markdown(
      """
        <style>
        /* Sembunyikan footer bawaan Streamlit */
        footer {visibility: hidden;}
        
        /* Styling Sidebar profesional */
        [data-testid="stSidebar"] {
            background-color: #0b0f19;
            border-right: 1px solid #1f2937;
        }
        
        /* Kotak Profil Pengguna di Sidebar */
        .user-profile-box {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            padding: 14px;
            border-radius: 10px;
            border: 1px solid #334155;
            text-align: center;
            margin-bottom: 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
        }
        </style>
        """,
      unsafe_allow_html=True,
  )
