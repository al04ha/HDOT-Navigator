"""tabs/tab_pipeline.py — File Pipeline tab (Outlook + manual upload)"""

import os
import sys
import subprocess
import tempfile
import streamlit as st

from config import BASE_DIRECTORY, IMAGE_EXTENSIONS
from chroma_store import index_file_to_chroma, index_file_to_kb
from doc_extractor import extract_text_from_file
from project_manager import (
    analyze_document_with_ai, resolve_destination_folder,
    get_active_email_attachments, download_outlook_attachment,
    get_subfolders,
)
from ui_components import render_ai_suggestion_card, render_pdf_preview


def render(tab):
    with tab:
        st.subheader("✉️ Outlook Attachment Ingestion")
        st.caption("Select an email in Outlook, then click Inspect to pull its attachments.")

        if st.button("🔄 Inspect Active Email Attachments", type="primary", use_container_width=True):
            with st.spinner("Connecting to Outlook…"):
                subject, attachments, error = get_active_email_attachments()
                if error:
                    st.error(error)
                elif not attachments:
                    st.warning(f"Email found but contains 0 attachments: `{subject}`")
                else:
                    st.session_state.outlook_subject     = subject
                    st.session_state.outlook_attachments = attachments
                    for k in ("ai_results", "accepted_paths", "filed_results", "selected_attachments"):
                        st.session_state[k] = {}
                    st.toast(f"Found {len(attachments)} attachment(s)!")

        if "outlook_attachments" in st.session_state and "outlook_subject" in st.session_state:
            email_subject = st.session_state.outlook_subject
            st.markdown("---")
            st.info(f"📥 **Email:** `{email_subject}`")

            col_analyze, _ = st.columns([0.35, 0.65])
            with col_analyze:
                if st.button("🤖 AI Analyze All Attachments", use_container_width=True):
                    progress = st.progress(0, text="Analyzing…")
                    for i, item in enumerate(st.session_state.outlook_attachments):
                        key = f"att_{item['index']}_{item['name']}"
                        if key not in st.session_state["ai_results"]:
                            extracted = ""
                            if item.get("preview_path") and os.path.exists(item["preview_path"]):
                                extracted = extract_text_from_file(item["preview_path"])
                            ai_result = analyze_document_with_ai(item["name"], extracted, email_subject)
                            st.session_state["ai_results"][key] = {
                                "ai":   ai_result,
                                "dest": resolve_destination_folder(ai_result),
                            }
                        progress.progress(
                            (i + 1) / len(st.session_state.outlook_attachments),
                            text=f"Analyzed {i+1}/{len(st.session_state.outlook_attachments)}: {item['name']}",
                        )
                    progress.empty()
                    st.toast("✅ AI analysis complete!")

            st.markdown("---")
            st.markdown("#### 📁 Choose Destination Folder")
            nav_path = st.session_state.folder_nav_path
            rel      = os.path.relpath(nav_path, BASE_DIRECTORY)
            parts    = ["(Root)"] + (rel.split(os.sep) if rel != "." else [])
            st.caption(f"📍 {' › '.join(parts)}")

            if nav_path != BASE_DIRECTORY:
                if st.button("⬆️ Go Up", key="folder_nav_up"):
                    st.session_state.folder_nav_path = os.path.dirname(nav_path)
                    st.rerun()

            subfolders = get_subfolders(nav_path)
            if subfolders:
                cols = st.columns(3)
                for i, folder in enumerate(subfolders):
                    with cols[i % 3]:
                        if st.button(f"📂 {folder}", key=f"nav_{folder}_{i}"):
                            st.session_state.folder_nav_path = os.path.join(nav_path, folder)
                            st.rerun()
            else:
                st.caption("_(no subfolders — this is the target)_")

            rel_display = os.path.relpath(nav_path, BASE_DIRECTORY)
            if rel_display == ".":
                rel_display = "(Root)"
            st.success(f"✅ **Selected destination:** `{rel_display}`")
            st.markdown("---")

            all_keys     = [f"att_{item['index']}_{item['name']}" for item in st.session_state.outlook_attachments]
            unfiled_keys = [k for k in all_keys if k not in st.session_state.get("filed_results", {})]

            all_selected = (
                all(st.session_state["selected_attachments"].get(k, False) for k in unfiled_keys)
                if unfiled_keys else False
            )
            select_all_val = st.checkbox("Select all attachments", value=all_selected, key="chk_select_all")
            if select_all_val != all_selected:
                for k in unfiled_keys:
                    st.session_state["selected_attachments"][k] = select_all_val
                st.rerun()

            selected_unfiled = [
                item for item in st.session_state.outlook_attachments
                if st.session_state["selected_attachments"].get(f"att_{item['index']}_{item['name']}", False)
                and f"att_{item['index']}_{item['name']}" not in st.session_state.get("filed_results", {})
            ]
            file_all_label = f"📥 File All Selected ({len(selected_unfiled)})" if selected_unfiled else "📥 File All Selected"
            if st.button(file_all_label, type="primary", use_container_width=True, disabled=not selected_unfiled):
                progress_bar = st.progress(0, text="Filing selected attachments…")
                filed_count, errors = 0, []
                for i, item in enumerate(selected_unfiled):
                    att_key     = f"att_{item['index']}_{item['name']}"
                    custom_name = st.session_state.get(f"input_{item['index']}_{item['name']}", item["name"])
                    effective_dest = st.session_state["accepted_paths"].get(att_key, st.session_state.folder_nav_path)
                    rel_dest       = os.path.relpath(effective_dest, BASE_DIRECTORY) if effective_dest else "(Root)"
                    os.makedirs(effective_dest, exist_ok=True)
                    full_target    = os.path.join(effective_dest, custom_name)
                    success, err_msg = download_outlook_attachment(item["index"], full_target)
                    if success:
                        index_file_to_chroma(full_target, rel_dest)
                        st.session_state["filed_results"][att_key] = {"name": custom_name, "rel": rel_dest, "abs": effective_dest}
                        st.session_state["selected_attachments"][att_key] = False
                        filed_count += 1
                    else:
                        errors.append(f"`{item['name']}`: {err_msg}")
                    progress_bar.progress((i + 1) / len(selected_unfiled), text=f"Filing {i+1}/{len(selected_unfiled)}: {item['name']}")
                progress_bar.empty()
                if filed_count:
                    st.toast(f"✅ Filed {filed_count} file(s)!")
                for err in errors:
                    st.error(f"COM Error — {err}")
                st.rerun()

            st.markdown("##### Process File Ingestion List:")
            for item in st.session_state.outlook_attachments:
                att_key       = f"att_{item['index']}_{item['name']}"
                already_filed = st.session_state.get("filed_results", {}).get(att_key)
                col_check, col_att_info, col_att_btn = st.columns([0.04, 3, 1])

                with col_check:
                    st.write(" ")
                    if already_filed:
                        st.markdown("✅")
                    else:
                        checked = st.session_state["selected_attachments"].get(att_key, False)
                        new_val = st.checkbox(
                            label=f"select_{att_key}", value=checked,
                            key=f"chk_{att_key}", label_visibility="collapsed",
                        )
                        if new_val != checked:
                            st.session_state["selected_attachments"][att_key] = new_val
                            st.rerun()

                with col_att_info:
                    custom_name = st.text_input(
                        f"Filename on disk (Index {item['index']}):",
                        value=item["name"],
                        key=f"input_{item['index']}_{item['name']}",
                        disabled=bool(already_filed),
                    )
                    if not already_filed:
                        if att_key not in st.session_state["ai_results"]:
                            if st.button(f"🤖 AI Analyze", key=f"analyze_{att_key}"):
                                with st.spinner(f"Analyzing {item['name']}…"):
                                    extracted = ""
                                    if item.get("preview_path") and os.path.exists(item["preview_path"]):
                                        extracted = extract_text_from_file(item["preview_path"])
                                    ai_result = analyze_document_with_ai(item["name"], extracted, email_subject)
                                    st.session_state["ai_results"][att_key] = {
                                        "ai": ai_result, "dest": resolve_destination_folder(ai_result),
                                    }
                                    st.rerun()
                        if att_key in st.session_state["ai_results"]:
                            cached   = st.session_state["ai_results"][att_key]
                            accepted = st.session_state["accepted_paths"].get(att_key)
                            if accepted:
                                st.success(f"✅ AI path accepted: `{os.path.relpath(accepted, BASE_DIRECTORY)}`")
                            else:
                                def make_accept_callback(k):
                                    def callback(path):
                                        st.session_state["accepted_paths"][k] = path
                                        st.session_state.folder_nav_path = path
                                        st.rerun()
                                    return callback
                                render_ai_suggestion_card(cached["ai"], cached["dest"], att_key, make_accept_callback(att_key))

                    if item.get("preview_path") and os.path.exists(item["preview_path"]):
                        fname_lower = item["name"].lower()
                        if fname_lower.endswith(".pdf"):
                            with st.expander(f"👁️ Preview: {item['name']}", expanded=False):
                                try:
                                    render_pdf_preview(open(item["preview_path"], "rb").read(), height=500)
                                except Exception as e:
                                    st.warning(f"Preview unavailable: {e}")
                        elif fname_lower.endswith((".docx", ".doc")):
                            size_kb = round(os.path.getsize(item["preview_path"]) / 1024, 1)
                            st.markdown(f"<div style='background:#1e1e2e;border:1px solid #444;border-radius:8px;padding:10px;margin-top:6px;'><span style='font-size:24px'>📝</span> <span style='color:#ccc;font-size:13px'>{item['name']}</span><br><span style='color:#888;font-size:11px'>Word Document | {size_kb} KB</span></div>", unsafe_allow_html=True)
                        elif fname_lower.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
                            size_kb = round(os.path.getsize(item["preview_path"]) / 1024, 1)
                            st.markdown(f"<div style='background:#1e1e2e;border:1px solid #444;border-radius:8px;padding:10px;margin-top:6px;'><span style='font-size:24px'>📊</span> <span style='color:#ccc;font-size:13px'>{item['name']}</span><br><span style='color:#888;font-size:11px'>Excel Spreadsheet | {size_kb} KB</span></div>", unsafe_allow_html=True)
                        else:
                            try:
                                from PIL import Image
                                st.image(Image.open(item["preview_path"]), caption=item["name"], width=300)
                            except Exception:
                                pass

                with col_att_btn:
                    st.write(" ")
                    if already_filed:
                        st.success("✅ Filed!")
                        st.caption(f"`{already_filed['rel']}`")
                        if st.button("📂 Open", key=f"open_filed_{att_key}"):
                            try:
                                if sys.platform.startswith("win32"):
                                    subprocess.run(["explorer", os.path.normpath(already_filed["abs"])])
                                else:
                                    subprocess.run(["xdg-open", already_filed["abs"]])
                            except Exception as e:
                                st.error(f"Could not open: {e}")
                    else:
                        effective_dest = st.session_state["accepted_paths"].get(att_key, nav_path)
                        effective_rel  = os.path.relpath(effective_dest, BASE_DIRECTORY) if effective_dest else rel_display
                        also_kb = st.checkbox(
                            "📖 Also add to Knowledge Base",
                            value=False,
                            key=f"also_kb_{att_key}",
                        )
                        if st.button("📥 File It", key=f"btn_{item['index']}", type="primary"):
                            os.makedirs(effective_dest, exist_ok=True)
                            full_target = os.path.join(effective_dest, custom_name)
                            success, err_msg = download_outlook_attachment(item["index"], full_target)
                            if success:
                                with st.spinner("Indexing…"):
                                    index_file_to_chroma(full_target, effective_rel)
                                    if also_kb:
                                        index_file_to_kb(full_target, custom_name)
                                st.session_state["filed_results"][att_key] = {
                                    "name": custom_name, "rel": effective_rel, "abs": effective_dest,
                                }
                                st.session_state["selected_attachments"][att_key] = False
                                kb_note = " + Knowledge Base" if also_kb else ""
                                st.toast(f"✅ Filed `{custom_name}`{kb_note}!")
                                st.rerun()
                            else:
                                st.error(f"COM Error: {err_msg}")
                        st.caption(f"→ `{effective_rel}`")

        st.markdown("---")
        st.subheader("📤 Manual File Upload")
        uploaded_file = st.file_uploader("Upload a file", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"])

        if uploaded_file:
            st.caption(f"📄 {uploaded_file.name} — {round(uploaded_file.size / 1024, 1)} KB")
            custom_name    = st.text_input("Save as:", value=uploaded_file.name, key="upload_custom_name")
            upload_ai_key  = f"upload_{uploaded_file.name}_{uploaded_file.size}"

            col_up_analyze, _ = st.columns([0.35, 0.65])
            with col_up_analyze:
                if upload_ai_key not in st.session_state["ai_results"]:
                    if st.button("🤖 AI Analyze This File", key="upload_analyze_btn"):
                        with st.spinner("Analyzing…"):
                            suffix, tmp_path, extracted = os.path.splitext(uploaded_file.name)[1], None, ""
                            try:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                                    tmp.write(uploaded_file.getbuffer())
                                    tmp_path = tmp.name
                                extracted = extract_text_from_file(tmp_path)
                            except Exception as ex:
                                st.error(f"Temp file error: {ex}")
                            finally:
                                if tmp_path and os.path.exists(tmp_path):
                                    try: os.unlink(tmp_path)
                                    except Exception: pass
                            ai_result = analyze_document_with_ai(uploaded_file.name, extracted)
                            if ai_result is None:
                                st.error("AI analysis failed. Check GEMINI_API_KEY.")
                            st.session_state["ai_results"][upload_ai_key] = {
                                "ai": ai_result, "dest": resolve_destination_folder(ai_result),
                            }
                            st.rerun()

            if upload_ai_key in st.session_state["ai_results"]:
                cached   = st.session_state["ai_results"][upload_ai_key]
                accepted = st.session_state["accepted_paths"].get(upload_ai_key)
                if accepted:
                    st.success(f"✅ AI path accepted: `{os.path.relpath(accepted, BASE_DIRECTORY)}`")
                else:
                    def upload_accept_callback(path):
                        st.session_state["accepted_paths"][upload_ai_key] = path
                        st.session_state.folder_nav_path = path
                        st.rerun()
                    render_ai_suggestion_card(cached["ai"], cached["dest"], upload_ai_key, upload_accept_callback)

            also_index_kb = st.checkbox("📖 Also add to Knowledge Base (Index Documents tab)", value=False, key="upload_also_index_kb")
            if st.button("📥 File & Index It", key="manual_upload_btn", type="primary"):
                effective_dest = st.session_state["accepted_paths"].get(upload_ai_key, st.session_state.folder_nav_path)
                os.makedirs(effective_dest, exist_ok=True)
                dest_path = os.path.join(effective_dest, custom_name)
                with open(dest_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                folder_display = os.path.relpath(effective_dest, BASE_DIRECTORY)
                with st.spinner("Indexing into project files…"):
                    index_file_to_chroma(dest_path, folder_display)
                if also_index_kb:
                    with st.spinner("Adding to knowledge base…"):
                        n = index_file_to_kb(dest_path, custom_name)
                        msg = f"✅ Filed `{custom_name}` → `{folder_display}` + {n} chunks added to knowledge base" if n else "⚠️ No extractable text for knowledge base."
                        st.success(msg) if n else st.warning(msg)
                else:
                    st.success(f"✅ Filed & indexed `{custom_name}` → `{folder_display}`")