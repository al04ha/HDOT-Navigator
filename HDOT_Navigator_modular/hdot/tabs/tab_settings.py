"""tabs/tab_settings.py — Settings tab"""

import os
from pathlib import Path
import streamlit as st

from config import BASE_DIRECTORY, DEFAULT_DEST_FOLDER, PROJECT_TEMPLATE_STRUCTURE
from chroma_store import clear_all_collections


def render(tab):
    with tab:
        st.subheader("⚙️ Settings")

        env_file = Path(__file__).resolve().parent / ".env"

        def read_env() -> dict:
            if not env_file.exists():
                return {}
            result = {}
            for line in env_file.read_text().splitlines():
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    result[k.strip()] = v.strip()
            return result

        def write_env(data: dict) -> None:
            existing_lines = env_file.read_text().splitlines() if env_file.exists() else []
            comment_lines  = [l for l in existing_lines if l.strip().startswith("#") or not l.strip()]
            env_file.write_text("\n".join(comment_lines + [f"{k}={v}" for k, v in data.items()]) + "\n")

        current = read_env()

        st.markdown("#### 📁 Project Directory")
        new_base  = st.text_input("Base Directory:", value=current.get("BASE_DIRECTORY", BASE_DIRECTORY), key="settings_base_dir")
        dir_valid = os.path.isdir(new_base.strip()) if new_base.strip() else False
        c1, c2 = st.columns([0.15, 0.85])
        (c1.success if dir_valid else c1.error)("✅ Valid" if dir_valid else "❌ Not found")
        if c2.button("💾 Save Directory", disabled=not dir_valid, key="save_base_dir"):
            current["BASE_DIRECTORY"] = new_base.strip()
            write_env(current)
            st.success("✅ Saved! Restart to apply.")

        st.markdown("---")
        st.markdown("#### 🔑 Gemini API Key")
        masked  = ("*" * 20 + current.get("GEMINI_API_KEY", "")[-4:]) if current.get("GEMINI_API_KEY") else "not set"
        st.caption(f"Current key: `{masked}`")
        new_key = st.text_input("New API Key:", type="password", placeholder="Paste new key here", key="settings_api_key")
        if st.button("💾 Save API Key", disabled=not new_key.strip(), key="save_api_key"):
            current["GEMINI_API_KEY"] = new_key.strip()
            write_env(current)
            st.success("✅ API key saved! Restart to apply.")

        st.markdown("---")
        st.markdown("#### 📂 Default Landing Folder")
        subfolder_options = [s.split("/")[0] for s in PROJECT_TEMPLATE_STRUCTURE if "/" not in s]
        current_default   = current.get("DEFAULT_DEST_FOLDER", DEFAULT_DEST_FOLDER)
        default_idx       = subfolder_options.index(current_default) if current_default in subfolder_options else 0
        new_default       = st.selectbox("Default subfolder:", options=subfolder_options, index=default_idx, key="settings_default_folder")
        if st.button("💾 Save Default Folder", key="save_default_folder"):
            current["DEFAULT_DEST_FOLDER"] = new_default
            write_env(current)
            st.success(f"✅ Default folder set to `{new_default}`")

        st.markdown("---")
        st.markdown("#### 🔗 OneDrive Shared Folder")
        st.caption(
            "Set a OneDrive folder path to share templates and AI corrections across your team. "
            "Each user's ChromaDB index stays local."
        )

        current_od = current.get("ONEDRIVE_SHARE_DIR", "")
        if current_od:
            if os.path.isdir(current_od):
                st.success(f"✅ Shared folder active: `{current_od}`")
                st.caption(
                    f"Templates: `{os.path.join(current_od, 'doc_templates')}` · "
                    f"Feedback: `{os.path.join(current_od, 'doc_type_feedback.json')}`"
                )
            else:
                st.error(f"❌ Folder not found: `{current_od}` — check the path and OneDrive sync status.")

        new_od = st.text_input(
            "OneDrive shared folder path:",
            value=current_od,
            placeholder=r"e.g. C:\Users\yourname\OneDrive - State of Hawaii\HDOT_Navigator_Shared",
            key="settings_onedrive",
        )
        od_valid = os.path.isdir(new_od.strip()) if new_od.strip() else True

        oc1, oc2 = st.columns([0.3, 0.7])
        with oc1:
            if new_od.strip():
                (st.success if od_valid else st.error)("✅ Valid" if od_valid else "❌ Not found")
            else:
                st.info("Blank = use local storage")
        with oc2:
            if st.button("💾 Save OneDrive Path", key="save_onedrive"):
                current["ONEDRIVE_SHARE_DIR"] = new_od.strip()
                write_env(current)
                if new_od.strip():
                    st.success("✅ Saved! Restart the app to apply.")
                else:
                    st.success("✅ Cleared — using local storage.")

        with st.expander("ℹ️ How to set up OneDrive sharing", expanded=False):
            st.markdown("""
        **Setup steps:**

        1. **Create a shared folder** in OneDrive, e.g. `HDOT_Navigator_Shared`
        2. **Share it** with your team (right-click → Share in File Explorer)
        3. **Make sure everyone** has it synced locally
        4. **In Settings**, set the OneDrive path on each person's machine
        5. **Restart the app** — templates and corrections are now shared

        **What's shared:** Document templates · AI feedback corrections · KB folder links

        **What stays local:** ChromaDB vector index · Project files · Gemini API key
            """)

        st.markdown("---")
        st.markdown("#### ⚠️ Danger Zone")
        with st.expander("🗑️ Clear Vector Index"):
            st.warning("This deletes all indexed embeddings. Files on disk are not affected.")
            confirm = st.text_input("Type DELETE to confirm:", key="confirm_clear_index")
            if st.button("🗑️ Clear Index", type="primary", key="clear_index_btn"):
                if confirm == "DELETE":
                    try:
                        for cname in [COLLECTION_NAME, MANUAL_COLLECTION]:
                            try: get_client().delete_collection(cname)
                            except Exception: pass
                        st.success("✅ Vector index cleared.")
                    except Exception as e:
                        st.error(f"Error: {e}")
                else:
                    st.error("Type DELETE exactly to confirm.")

        st.markdown("---")
        st.markdown("#### 🔍 Gemini API Request Log")
        st.caption("Live log of all prompts and queries sent to the Gemini API this session.")

        if "_gemini_log" not in st.session_state:
            st.session_state["_gemini_log"] = []

        if st.session_state["_gemini_log"]:
            col_log1, col_log2 = st.columns([0.7, 0.3])
            col_log1.caption(f"{len(st.session_state['_gemini_log'])} request(s) this session")
            if col_log2.button("🗑️ Clear Log", key="clear_gemini_log"):
                st.session_state["_gemini_log"] = []
                st.rerun()

            for i, entry in enumerate(reversed(st.session_state["_gemini_log"])):
                with st.expander(
                    f"#{len(st.session_state['_gemini_log']) - i} · {entry['model']} · "
                    f"{entry['status']} · {entry['timestamp']}",
                    expanded=False,
                ):
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Model", entry["model"])
                    col_b.metric("Status", entry["status"])
                    col_c.metric("Temp", entry["temperature"])

                    st.markdown("**Prompt sent:**")
                    st.code(entry["prompt"], language="text")

                    if entry.get("response"):
                        st.markdown("**Response received:**")
                        st.code(entry["response"][:2000] + ("…" if len(entry["response"]) > 2000 else ""), language="text")
                    if entry.get("error"):
                        st.error(f"Error: {entry['error']}")
        else:
            st.info("No API requests made yet this session.")
