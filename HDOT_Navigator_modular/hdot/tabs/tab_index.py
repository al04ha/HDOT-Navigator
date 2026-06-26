"""tabs/tab_index.py — Index Documents (KB) tab"""

import os
import time
import tempfile
import streamlit as st

from config import MANUAL_COLLECTION
from ai_client import get_embedding
from chroma_store import (
    get_client, kb_delete_by_source_and_folder,
    kb_delete_by_folder, kb_move_source_to_folder,
)
from doc_extractor import extract_text_from_file, chunk_text


def render(tab):
    with tab:
        st.subheader("📑 Index Documents")
        st.caption("Manage your Secretaries Manual folder structure and indexed documents.")

        col_kb = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
        kb_count = col_kb.count()
        # Call with no include so IDs + metadatas are both returned
        _all = col_kb.get() if kb_count > 0 else {"metadatas": [], "ids": []}
        kb_meta = _all.get("metadatas", [])

        def get_kb_folders() -> dict:
            folders = {}
            for m in kb_meta:
                # skip placeholder entries
                if m.get("placeholder"):
                    continue
                folder = m.get("folder", "General")
                source = m.get("source", "unknown")
                if folder not in folders:
                    folders[folder] = set()
                folders[folder].add(source)
            return {k: sorted(v) for k, v in sorted(folders.items())}

        def get_chunk_counts(folder: str = None) -> dict:
            counts = {}
            for m in kb_meta:
                if m.get("placeholder"):
                    continue
                if folder and m.get("folder", "General") != folder:
                    continue
                src = m.get("source", "unknown")
                counts[src] = counts.get(src, 0) + 1
            return counts

        kb_folders = get_kb_folders()

        if kb_count > 0:
            st.success(f"✅ {kb_count} chunks indexed across {len(kb_folders)} folder(s)")
        else:
            st.warning("⚠️ No documents indexed yet.")

        # ── Debug / repair tool ───────────────────────────────────────────────
        with st.expander("🔧 Debug & Repair", expanded=False):
            st.caption("Use this if folders show 0 chunks or delete isn't working.")
            if st.button("🔍 Show raw ChromaDB data", key="debug_raw"):
                raw = col_kb.get()
                metas = raw.get("metadatas", [])
                ids   = raw.get("ids", [])
                st.write(f"Total records: {len(ids)}")
                for i, (mid, m) in enumerate(zip(ids, metas)):
                    st.caption(f"[{i}] id={mid[:30]}… | source={m.get('source','?')} | folder={m.get('folder','NO FOLDER')} | chunk={m.get('chunk','?')}")
            
            st.markdown("---")
            st.caption("If files show 'NO FOLDER' above, click below to assign them to a folder.")
            repair_folder = st.text_input("Assign unfoldered docs to:", value="General", key="repair_folder_name")
            if st.button("🔨 Fix unfoldered documents", key="fix_no_folder", type="primary"):
                raw = col_kb.get()
                metas = raw.get("metadatas", [])
                ids   = raw.get("ids", [])
                fixed = 0
                for uid, m in zip(ids, metas):
                    if not m.get("folder") and not m.get("placeholder"):
                        col_kb.update(ids=[uid], metadatas=[{**m, "folder": repair_folder.strip()}])
                        fixed += 1
                if fixed:
                    st.success(f"✅ Fixed {fixed} records → assigned to '{repair_folder.strip()}'")
                    st.rerun()
                else:
                    st.info("No unfoldered documents found.")

        st.markdown("---")
        st.markdown("#### ➕ Add Documents")

        folder_options = list(kb_folders.keys()) if kb_folders else []

        use_existing = st.radio(
            "Add to:",
            ["Existing folder", "New folder"],
            horizontal=True,
            key="upload_folder_mode",
        )

        if use_existing == "Existing folder":
            target_folder = st.selectbox(
                "Select folder:",
                options=folder_options if folder_options else ["(no folders yet)"],
                key="upload_target_folder",
            )
            if not folder_options:
                target_folder = None
        else:
            target_folder = st.text_input(
                "New folder name:",
                placeholder="e.g. Chapter 02 - Pre-Design",
                key="upload_new_folder_inline",
            ).strip() or None

        kb_upload = st.file_uploader(
            "Choose documents (PDF, DOCX, XLSX)",
            type=["pdf", "docx", "xlsx"],
            key="kb_uploader",
            accept_multiple_files=True,
        )

        if kb_upload and target_folder:
            if st.button("📥 Index Selected Files", type="primary", key="index_files_btn"):
                results = []
                idx_progress = st.progress(0, text="Starting indexing…")
                total_files = len(kb_upload)
                for file_num, uploaded in enumerate(kb_upload):
                    idx_progress.progress(
                        file_num / total_files,
                        text=f"Indexing {file_num+1}/{total_files}: {uploaded.name}",
                    )
                    suffix   = os.path.splitext(uploaded.name)[1].lower()
                    tmp_path = None
                    text     = ""
                    error    = None
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode="wb") as tmp:
                            tmp.write(uploaded.getbuffer())
                            tmp_path = tmp.name
                        text = extract_text_from_file(tmp_path)
                    except Exception as e:
                        error = str(e)
                    finally:
                        if tmp_path and os.path.exists(tmp_path):
                            try: os.unlink(tmp_path)
                            except Exception: pass

                    if error:
                        results.append(("error", f"❌ **{uploaded.name}** — extraction error: `{error}`"))
                        continue
                    if not text.strip():
                        results.append(("warning", f"⚠️ **{uploaded.name}** — no text extracted."))
                        continue

                    word_count = len(text.split())
                    chunks = chunk_text(text, chunk_size=500, overlap=100)
                    indexed = 0
                    for idx, chunk in enumerate(chunks):
                        vec = get_embedding(chunk)
                        if vec:
                            col_kb.add(
                                embeddings=[vec],
                                documents=[chunk],
                                metadatas=[{
                                    "source": uploaded.name,
                                    "folder": target_folder,
                                    "chunk": idx,
                                }],
                                ids=[f"kb_{uploaded.name}_{idx}_{int(time.time())}"],
                            )
                            indexed += 1

                    if indexed > 0:
                        results.append(("success",
                            f"✅ **{uploaded.name}** → `{target_folder}` — {indexed} chunks ({word_count:,} words)."
                        ))
                    else:
                        results.append(("warning",
                            f"⚠️ **{uploaded.name}** — embeddings failed. Check API key/quota."
                        ))

                idx_progress.progress(100, text="✅ Indexing complete!")
                idx_progress.empty()

                st.session_state["kb_index_results"] = results
                st.rerun()

        elif kb_upload and not target_folder:
            st.warning("Select or enter a folder name above before indexing.")

        if st.session_state.get("kb_index_results"):
            for kind, msg in st.session_state["kb_index_results"]:
                if kind == "success": st.success(msg)
                elif kind == "warning": st.warning(msg)
                else: st.error(msg)
            if st.button("✖ Dismiss", key="dismiss_kb_results"):
                st.session_state["kb_index_results"] = []
                st.rerun()

        st.markdown("---")
        st.markdown("#### 📂 Folder Structure")

        if not kb_folders:
            st.info("No folders yet. Create one above or upload a document.")
        else:
            for folder_name, sources in kb_folders.items():
                chunk_counts = get_chunk_counts(folder_name)
                real_sources = [s for s in sources if not s.startswith("__folder_")]
                total_chunks = sum(chunk_counts.get(s, 0) for s in real_sources)

                with st.expander(f"📁 {folder_name}  ({len(real_sources)} file(s) · {total_chunks} chunks)", expanded=False):

                    if not real_sources:
                        st.caption("_(empty folder)_")
                    else:
                        for src in real_sources:
                            c1, c2, c3 = st.columns([0.65, 0.2, 0.15])
                            c1.markdown(f"📄 **{src}**")
                            c1.caption(f"{chunk_counts.get(src, 0)} chunks")

                            other_folders = [f for f in kb_folders if f != folder_name]
                            if other_folders:
                                move_dest = c2.selectbox(
                                    "Move to:",
                                    options=["—"] + other_folders,
                                    key=f"move_{folder_name}_{src}",
                                    label_visibility="collapsed",
                                )
                                if move_dest != "—":
                                    if st.button("➡️ Move", key=f"do_move_{folder_name}_{src}"):
                                        try:
                                            kb_move_source_to_folder(src, folder_name, move_dest)
                                            st.success(f"Moved **{src}** → `{move_dest}`")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Move failed: {e}")
                            else:
                                c2.caption("(only folder)")

                            if c3.button("🗑️", key=f"del_kb_{folder_name}_{src}", help=f"Remove {src}"):
                                try:
                                    kb_delete_by_source_and_folder(src, folder_name)
                                    st.success(f"Removed {src}.")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Could not remove: {e}")

                    st.markdown("---")
                    if st.button(f"🗑️ Delete folder '{folder_name}'", key=f"del_folder_{folder_name}"):
                        try:
                            kb_delete_by_folder(folder_name)
                            st.success(f"Deleted folder '{folder_name}'.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Could not delete folder: {e}")