"""tabs/tab_manage.py — Project Manager tab"""

import os
import streamlit as st

from config import BASE_DIRECTORY
from project_manager import create_project_from_template


def render(tab):
    with tab:
        st.subheader("⚙️ Project Manager")
        with st.form("new_proj"):
            p_num  = st.text_input("Project Number")
            p_name = st.text_input("Project Name")
            clone  = st.checkbox("Clone Template", value=True)
            if st.form_submit_button("Create Project"):
                p_num_clean  = p_num.strip()
                p_name_clean = p_name.strip()
                if not p_num_clean or not p_name_clean:
                    st.error("Project number and name cannot be empty.")
                else:
                    dest = os.path.join(BASE_DIRECTORY, f"{p_num_clean}_{p_name_clean}")
                    if os.path.exists(dest):
                        st.error(f"Project `{p_num_clean}_{p_name_clean}` already exists.")
                    elif clone:
                        create_project_from_template(dest)
                        st.success(f"✅ Project `{p_num_clean}_{p_name_clean}` created from template!")
                    else:
                        os.makedirs(dest)
                        st.success(f"✅ Project `{p_num_clean}_{p_name_clean}` created!")

