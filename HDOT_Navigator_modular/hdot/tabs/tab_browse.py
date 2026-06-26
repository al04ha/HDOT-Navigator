"""tabs/tab_browse.py — File System Viewer tab"""

import os
import shutil
import time
import streamlit as st

from config import BASE_DIRECTORY, MANUAL_COLLECTION
from chroma_store import get_client
from project_manager import get_all_folders, folder_display_name, open_file, trash_item

TRASH_DIRECTORY = os.path.join(BASE_DIRECTORY, ".trash")


def render(tab):
    with tab:
        st.subheader("🗂️ File System Viewer")

        with st.expander("🛠️ Admin: Knowledge Base Management", expanded=False):
            if st.button("🔍 List Indexed Files"):
                try:
                    kb_col = get_client().get_or_create_collection(MANUAL_COLLECTION)
                    results = kb_col.get(include=["metadatas"])
                    metas   = results.get("metadatas", [])
                    if metas:
                        counts: dict[str, int] = {}
                        for m in metas:
                            src = m.get("source", "Unknown")
                            counts[src] = counts.get(src, 0) + 1
                        st.write("Files currently in Knowledge Base:")
                        for src, count in sorted(counts.items()):
                            st.caption(f"📄 {src} — {count} chunk(s)")
                    else:
                        st.info("The Knowledge Base is currently empty.")
                except Exception as e:
                    st.error(f"Error reading index: {e}")

        if "last_deleted" in st.session_state:
            deleted_info = st.session_state["last_deleted"]
            if time.time() - deleted_info.get("deleted_at", time.time()) > 5:
                del st.session_state["last_deleted"]
                st.rerun()
            else:
                c1, c2, c3 = st.columns([0.7, 0.15, 0.15])
                c1.warning(f"🗑️ Deleted: **{deleted_info['name']}**")
                if c2.button("↩️ Undo"):
                    shutil.move(deleted_info["trash_path"], deleted_info["original_path"])
                    del st.session_state["last_deleted"]
                    st.rerun()
                if c3.button("✖ Dismiss"):
                    del st.session_state["last_deleted"]
                    st.rerun()

        moving = st.session_state.get("move_file_path")
        if moving and os.path.isfile(moving):
            fname = os.path.basename(moving)
            st.info(f"✂️ Moving **{fname}** (from `{folder_display_name(os.path.dirname(moving), BASE_DIRECTORY)}`)")
            all_folders   = get_all_folders(BASE_DIRECTORY)
            folder_labels = [folder_display_name(f, BASE_DIRECTORY) for f in all_folders]
            current_abs   = os.path.dirname(moving)
            default_idx   = all_folders.index(current_abs) if current_abs in all_folders else 0
            chosen_label  = st.selectbox("Move to folder:", options=folder_labels, index=default_idx, key="move_dest_select")
            chosen_abs    = all_folders[folder_labels.index(chosen_label)]
            col_confirm, col_cancel = st.columns([0.2, 0.8])
            with col_confirm:
                if st.button("✅ Move Here", type="primary", use_container_width=True):
                    dest_path = os.path.join(chosen_abs, fname)
                    if os.path.abspath(chosen_abs) == os.path.abspath(os.path.dirname(moving)):
                        st.warning("File is already in that folder.")
                    elif os.path.exists(dest_path):
                        st.error(f"A file named `{fname}` already exists there.")
                    else:
                        shutil.move(moving, dest_path)
                        st.success(f"✅ Moved **{fname}** → `{chosen_label}`")
                        st.session_state["move_file_path"] = None
                        st.rerun()
            with col_cancel:
                if st.button("✖ Cancel Move"):
                    st.session_state["move_file_path"] = None
                    st.rerun()
            st.markdown("---")

        def render_tree(path: str) -> None:
            for item in sorted(os.listdir(path)):
                if item == ".trash":
                    continue
                p = os.path.join(path, item)
                if os.path.isdir(p):
                    col_label, col_del = st.columns([0.92, 0.08])
                    with col_label:
                        with st.expander(f"📁 {item}"):
                            render_tree(p)
                    with col_del:
                        if st.button("❌", key=f"del_{p}", help=f"Delete {item}"):
                            if trash_item(p, TRASH_DIRECTORY): st.rerun()
                else:
                    is_moving = st.session_state.get("move_file_path") == p
                    c1, c2, c3, c4 = st.columns([0.70, 0.09, 0.11, 0.10])
                    c1.markdown(f"📄 **{item}**" if is_moving else f"📄 {item}")
                    if c2.button("📂", key=f"open_{p}", help="Open file"):
                        open_file(p)
                    if c3.button("➡️ …" if is_moving else "➡️", key=f"move_{p}",
                                 help="Cancel move" if is_moving else "Move file"):
                        st.session_state["move_file_path"] = None if is_moving else p
                        st.rerun()
                    if c4.button("❌", key=f"delf_{p}", help=f"Delete {item}"):
                        if trash_item(p, TRASH_DIRECTORY): st.rerun()

        render_tree(BASE_DIRECTORY)

