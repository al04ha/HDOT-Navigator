"""tabs/tab_progress.py — Project Progress tracker tab"""

import streamlit as st
import pandas as pd

from progress_store import (
    list_projects, get_project_progress,
    get_status_summary, generate_status_summary,
)


def render(tab):
    with tab:
        st.subheader("📊 Project Progress")
        st.caption("Completion is based on how many phase folders contain documents.")

        projects = list_projects()
        if not projects:
            st.info("No projects yet. Create one in the Project Manager tab.")
            return

        selected = st.selectbox("Select a project:", options=projects, key="progress_project_select")
        if not selected:
            return

        prog = get_project_progress(selected)

        # ── Top metrics ───────────────────────────────────────────────────────
        m1, m2, m3 = st.columns(3)
        m1.metric("Overall Completion", f"{prog['percent']}%")
        m2.metric("Phases With Files", f"{prog['completed']} / {prog['total']}")
        m3.metric("Total Documents", prog["total_files"])

        st.progress(prog["percent"] / 100)
        st.markdown("---")

        # ── Most recent document ──────────────────────────────────────────────
        st.markdown("#### 🕒 Most Recent Document")
        if prog["recent"]:
            r = prog["recent"]
            st.markdown(
                f"<div style='background:#1a1a2e;border-left:4px solid #4caf50;"
                f"border-radius:8px;padding:12px 16px;'>"
                f"<span style='color:#e0e0e0;font-weight:600;'>📄 {r['name']}</span><br>"
                f"<span style='color:#aaa;font-size:12px;'>Phase: {r['phase']} · "
                f"Last modified: {r['modified']}</span></div>",
                unsafe_allow_html=True,
            )
        else:
            st.caption("No documents filed yet.")

        st.markdown("---")

        # ── Per-phase chart ───────────────────────────────────────────────────
        st.markdown("#### 📈 Phase Breakdown")
        df = pd.DataFrame([
            {"Phase": p["name"], "Documents": p["file_count"]}
            for p in prog["phases"]
        ])
        df = df.set_index("Phase")
        st.bar_chart(df, height=400, color="#4caf50")

        # ── Phase checklist ───────────────────────────────────────────────────
        st.markdown("#### ✅ Phase Checklist")
        for p in prog["phases"]:
            icon = "✅" if p["has_files"] else "⬜"
            count_txt = f" — {p['file_count']} file(s)" if p["has_files"] else " — empty"
            st.markdown(f"{icon} **{p['name']}**{count_txt}")

        st.markdown("---")

        # ── AI status summary ─────────────────────────────────────────────────
        st.markdown("#### 🤖 AI Status Summary")
        cached = get_status_summary(selected)
        if cached:
            st.info(cached)
        if st.button("🔄 Generate / Refresh Summary", key="gen_summary"):
            with st.spinner("Analyzing project status…"):
                summary = generate_status_summary(selected)
            if summary:
                st.rerun()
            else:
                st.error("Could not generate summary. Check your Gemini API key or quota.")