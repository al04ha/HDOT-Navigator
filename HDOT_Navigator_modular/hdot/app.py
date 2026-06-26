"""
app.py
Entry point for the HDOT Project Navigator Streamlit app.

All logic lives in the modules below — this file only:
  1. Validates config and prompts for BASE_DIRECTORY on first run
  2. Initialises session state defaults
  3. Renders the tab shell and delegates to each tab module
"""

import os
import sys
import time
from pathlib import Path

# Ensure hdot/ is on sys.path so all modules can find each other
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

import streamlit as st

# ── Must be the very first Streamlit call ─────────────────────────────────────
st.set_page_config(page_title="HDOT Project Navigator", layout="wide", page_icon="📂")

# ── Config (loads .env files, sets all constants) ─────────────────────────────
import config  # noqa: E402  — import after set_page_config
from config import (
    API_KEY, BASE_DIRECTORY, DEFAULT_DEST_FOLDER, PROJECT_TEMPLATE_STRUCTURE,
)

# ── API key guard ─────────────────────────────────────────────────────────────
if not API_KEY:
    st.error("Error: GEMINI_API_KEY not found in configuration.")
    st.stop()

# ── First-run: ask for BASE_DIRECTORY if not set ──────────────────────────────
if not BASE_DIRECTORY:
    local_env_path = Path.home() / ".hdot_navigator" / ".env"
    st.markdown("## 👋 Welcome to HDOT Project Navigator")
    st.markdown("**One-time setup:** Tell the app where your project files are stored.")
    st.markdown("---")
    typed_path = st.text_input(
        "📁 Enter the full path to your project folder:",
        placeholder=r"e.g. C:\Users\yourname\Downloads\PDM\PDM",
        key="setup_base_dir",
    )
    if st.button("✅ Save & Launch", type="primary"):
        p = typed_path.strip()
        if p and os.path.isdir(p):
            local_env_path.parent.mkdir(parents=True, exist_ok=True)
            lines = local_env_path.read_text().splitlines() if local_env_path.exists() else []
            lines = [l for l in lines if not l.startswith("BASE_DIRECTORY=")]
            lines.append(f"BASE_DIRECTORY={p}")
            local_env_path.write_text("\n".join(lines) + "\n")
            st.success("✅ Saved! Restarting…")
            time.sleep(1)
            st.rerun()
        elif typed_path.strip():
            st.error(f"❌ Folder not found: `{typed_path.strip()}`.")
        else:
            st.warning("Please enter a folder path.")
    st.stop()

# ── Ensure trash directory exists ─────────────────────────────────────────────
os.makedirs(os.path.join(BASE_DIRECTORY, ".trash"), exist_ok=True)

# ── Session state defaults ────────────────────────────────────────────────────
_SS_DEFAULTS = {
    "folder_nav_path": (
        os.path.join(BASE_DIRECTORY, DEFAULT_DEST_FOLDER)
        if os.path.exists(os.path.join(BASE_DIRECTORY, DEFAULT_DEST_FOLDER))
        else BASE_DIRECTORY
    ),
    "move_file_path":        None,
    "ai_results":            {},
    "accepted_paths":        {},
    "chat_history":          [],
    "filed_results":         {},
    "selected_attachments":  {},
    "kb_index_results":      [],
    # ── Secretaries Manual ────────────────────────────────────────────────────
    "secman_filled_name":          None,
    "secman_filled_text":          None,
    "secman_filled_bytes":         None,
    "secman_filled_is_docx":       False,
    "secman_plain_text":           None,
    "secman_report":               None,
    "secman_report_parsed":        None,
    "secman_doc_type":             None,
    "secman_template_context":     "",
    "secman_auto_recheck":         False,
    "secman_fixed_docx_bytes":     None,
    "secman_fix_log":              None,
    "secman_pending_fills":        None,
    "secman_pending_labels":       None,
    "secman_pending_tmpl_map":     None,
    "secman_pending_placeholders": None,
    "secman_debug_fills":          None,
    "_secman_template_docx":       None,
    # ── Template upload ───────────────────────────────────────────────────────
    "_tmpl_upload_name":     None,
    "_tmpl_upload_bytes":    None,
    "_gemini_log":           [],
}
for _k, _v in _SS_DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── Tab imports (after config is fully loaded) ────────────────────────────────
from tabs import tab_pipeline, tab_search, tab_manage, tab_browse  # noqa: E402
from tabs import tab_index, tab_secmanual, tab_settings, tab_progress  # noqa: E402
from streamlit_option_menu import option_menu  # noqa: E402

# ── Vertical sidebar navigation ───────────────────────────────────────────────
st.title("📂 HDOT Project Navigator")

MENU_ITEMS = [
    "OutLook Exracter",
    "Document Assistant",
    "Project Manager",
    "Progress",
    "File System",
    "Index Documents",
    "Secretaries Manual",
    "Settings",
]
MENU_ICONS = [
    "rocket-takeoff",
    "search",
    "gear",
    "bar-chart-line",
    "folder",
    "file-earmark-text",
    "journal-text",
    "sliders",
]

with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=MENU_ITEMS,
        icons=MENU_ICONS,
        menu_icon="list",
        default_index=0,
    )

# Render only the selected page into a container
page = st.container()

if selected == "OutLook Exracter":
    tab_pipeline.render(page)
elif selected == "Document Assistant":
    tab_search.render(page)
elif selected == "Project Manager":
    tab_manage.render(page)
elif selected == "Progress":
    tab_progress.render(page)
elif selected == "File System":
    tab_browse.render(page)
elif selected == "Index Documents":
    tab_index.render(page)
elif selected == "Secretaries Manual":
    tab_secmanual.render(page)
elif selected == "Settings":
    tab_settings.render(page)