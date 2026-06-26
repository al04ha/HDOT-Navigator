"""tabs/tab_search.py — Document Assistant (chat) tab"""

import streamlit as st

from config import COLLECTION_NAME, MANUAL_COLLECTION
from ai_client import get_embedding, gemini_call
from chroma_store import get_client


def render(tab):
    with tab:
        col_title, col_actions = st.columns([0.7, 0.3])
        with col_title:
            st.subheader("💬 Document Assistant")
            st.caption("Ask anything about your projects, documents, procedures, or HDOT policies.")
        with col_actions:
            st.write(" ")
            try:
                _proj_count = get_client().get_or_create_collection(COLLECTION_NAME).count()
                _kb_count   = get_client().get_or_create_collection(MANUAL_COLLECTION).count()
                status_icon = "🟢" if (_proj_count + _kb_count) > 0 else "🔴"
                st.caption(f"{status_icon} {_proj_count} project chunks · {_kb_count} KB chunks")
            except Exception:
                pass
            if st.session_state["chat_history"]:
                if st.button("🗑️ Clear chat", key="clear_chat", use_container_width=True):
                    st.session_state["chat_history"] = []
                    st.rerun()

        st.markdown("---")
        if not st.session_state["chat_history"]:
            st.markdown("""
            <div style='background:#1a1a2e;border:1px solid #3a3a5c;border-radius:12px;padding:20px 24px;margin-bottom:16px;'>
                <div style='font-size:28px;margin-bottom:8px;'>👋</div>
                <div style='color:#e0e0e0;font-size:15px;font-weight:600;margin-bottom:6px;'>Hi! I'm your HDOT Project Assistant.</div>
                <div style='color:#aaa;font-size:13px;line-height:1.6;'>
                    I can help you find documents, explain project status, summarize what's been filed,
                    answer questions about procedures, and more.
                </div>
                <div style='margin-top:12px;display:flex;flex-wrap:wrap;gap:8px;'>
                    <span style='background:#2a2a4a;color:#90caf9;padding:5px 12px;border-radius:20px;font-size:12px;'>What documents are filed for project 12345?</span>
                    <span style='background:#2a2a4a;color:#90caf9;padding:5px 12px;border-radius:20px;font-size:12px;'>What's the process for submitting a change order?</span>
                    <span style='background:#2a2a4a;color:#90caf9;padding:5px 12px;border-radius:20px;font-size:12px;'>Summarize the construction submittals we have</span>
                    <span style='background:#2a2a4a;color:#90caf9;padding:5px 12px;border-radius:20px;font-size:12px;'>What ROW documents have been received?</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if msg["role"] == "assistant" and msg.get("sources"):
                    sources = msg["sources"]
                    if sources.get("project") or sources.get("kb"):
                        with st.expander("📎 Sources referenced", expanded=False):
                            if sources.get("project"):
                                st.markdown("**From project files:**")
                                for s in sources["project"]:
                                    st.caption(f"📄 **{s['filename']}** — `{s['folder']}`")
                            if sources.get("kb"):
                                st.markdown("**From knowledge base:**")
                                for s in sources["kb"]:
                                    st.caption(f"📖 {s['source']}")

        user_input = st.chat_input("Ask me anything about your projects or documents…")
        if user_input:
            st.session_state["chat_history"].append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.markdown(user_input)

            with st.chat_message("assistant"):
                with st.spinner("Thinking…"):
                    vec = get_embedding(user_input)
                    project_docs, project_metas, kb_docs, kb_metas = [], [], [], []

                    if vec:
                        try:
                            col_proj = get_client().get_or_create_collection(name=COLLECTION_NAME)
                            if col_proj.count() > 0:
                                res = col_proj.query(query_embeddings=[vec], n_results=6)
                                project_docs  = res.get("documents", [[]])[0]
                                project_metas = res.get("metadatas", [[]])[0]
                        except Exception:
                            pass
                        try:
                            col_kb_q = get_client().get_or_create_collection(name=MANUAL_COLLECTION)
                            if col_kb_q.count() > 0:
                                res_kb = col_kb_q.query(query_embeddings=[vec], n_results=4)
                                kb_docs  = res_kb.get("documents", [[]])[0]
                                kb_metas = res_kb.get("metadatas", [[]])[0]
                        except Exception:
                            pass

                    context_parts = []
                    if project_docs:
                        context_parts.append("=== FILED PROJECT DOCUMENTS ===")
                        for i, doc in enumerate(project_docs):
                            meta = project_metas[i] if i < len(project_metas) else {}
                            context_parts.append(
                                f"File: {meta.get('filename','unknown')}\n"
                                f"Location: {meta.get('folder','')}\n"
                                f"Content excerpt: {doc[:500]}"
                            )
                    if kb_docs:
                        context_parts.append("=== KNOWLEDGE BASE / PROCEDURES ===")
                        for i, doc in enumerate(kb_docs):
                            meta = kb_metas[i] if i < len(kb_metas) else {}
                            context_parts.append(f"Source: {meta.get('source','reference')}\nContent: {doc[:500]}")

                    context_str = "\n\n".join(context_parts)[:3000] if context_parts else ""
                    history_str = "".join(
                        f"{'User' if m['role']=='user' else 'Assistant'}: {m['content'][:400]}\n\n"
                        for m in st.session_state["chat_history"][:-1][-6:]
                    )

                    if context_str:
                        chat_prompt = f"""You are a knowledgeable, conversational AI assistant for the Hawaii Department of Transportation (HDOT) projects team.

        CONTEXT FROM INDEXED DOCUMENTS:
        {context_str}

        CONVERSATION SO FAR:
        {history_str}

        USER QUESTION: {user_input}

        Give a thorough, helpful answer. Reference specific filenames and folder paths naturally. Use markdown formatting."""
                    else:
                        chat_prompt = f"""You are a knowledgeable AI assistant for the HDOT projects team.

        CONVERSATION SO FAR:
        {history_str}

        USER QUESTION: {user_input}

        No matching documents were found in the index. Answer helpfully from general knowledge if applicable, otherwise guide the user to file and index the relevant documents."""

                    answer = gemini_call(chat_prompt, temperature=0.4, timeout=45)
                    if not answer:
                        if project_docs or kb_docs:
                            lines = ["I'm having trouble reaching the AI right now, but here's what I found:\n"]
                            for i, doc in enumerate(project_docs[:4]):
                                meta = project_metas[i] if i < len(project_metas) else {}
                                lines.append(f"- 📄 **{meta.get('filename','')}** in `{meta.get('folder','')}`: {doc[:200]}…")
                            for i, doc in enumerate(kb_docs[:2]):
                                meta = kb_metas[i] if i < len(kb_metas) else {}
                                lines.append(f"- 📖 **{meta.get('source','')}**: {doc[:200]}…")
                            answer = "\n".join(lines)
                        else:
                            answer = "I couldn't find any matching documents. Try rephrasing, or check that files have been indexed via the File Pipeline tab."

                    st.markdown(answer)
                    sources: dict = {}
                    if project_docs:
                        sources["project"] = [
                            {"filename": project_metas[i].get("filename", "?"), "folder": project_metas[i].get("folder", "")}
                            for i in range(len(project_docs)) if i < len(project_metas)
                        ]
                    if kb_docs:
                        sources["kb"] = [
                            {"source": kb_metas[i].get("source", "?")}
                            for i in range(len(kb_docs)) if i < len(kb_metas)
                        ]
                    if sources.get("project") or sources.get("kb"):
                        with st.expander("📎 Sources referenced", expanded=False):
                            if sources.get("project"):
                                st.markdown("**From project files:**")
                                for s in sources["project"]:
                                    st.caption(f"📄 **{s['filename']}** — `{s['folder']}`")
                            if sources.get("kb"):
                                st.markdown("**From knowledge base:**")
                                for s in sources["kb"]:
                                    st.caption(f"📖 {s['source']}")

            st.session_state["chat_history"].append({
                "role": "assistant",
                "content": answer,
                "sources": sources if (project_docs or kb_docs) else {},
            })

