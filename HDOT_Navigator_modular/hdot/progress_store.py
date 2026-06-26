"""
progress_store.py
Tracks per-project progress based on which phase folders contain files.
Progress = (phase folders with >=1 file) / (total phase folders).
Also surfaces the most recently modified document and an optional AI summary.
"""

import os
import json
import time
from pathlib import Path

from config import BASE_DIRECTORY, PROJECT_TEMPLATE_STRUCTURE, USER_HOME

# Only the top-level phases count toward progress (skip "14 APPENDIX/..." subfolders)
PHASE_FOLDERS = [p for p in PROJECT_TEMPLATE_STRUCTURE if "/" not in p]

# Where we cache AI status summaries
_PROGRESS_FILE = str(USER_HOME / ".hdot_navigator" / "project_progress.json")

# File types we count as real documents
_DOC_EXTS = (".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".png", ".jpg", ".jpeg")


def _load_summaries() -> dict:
    try:
        if os.path.exists(_PROGRESS_FILE):
            with open(_PROGRESS_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_summaries(data: dict) -> None:
    try:
        os.makedirs(os.path.dirname(_PROGRESS_FILE), exist_ok=True)
        with open(_PROGRESS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"[progress] save failed: {e}")


def list_projects() -> list[str]:
    """Return the names of all project folders under BASE_DIRECTORY."""
    projects = []
    try:
        for item in sorted(os.listdir(BASE_DIRECTORY)):
            full = os.path.join(BASE_DIRECTORY, item)
            if os.path.isdir(full) and item != ".trash":
                projects.append(item)
    except Exception:
        pass
    return projects


def _folder_has_files(path: str) -> bool:
    """True if a folder contains at least one real document (recursively)."""
    if not os.path.isdir(path):
        return False
    for root, _, files in os.walk(path):
        for f in files:
            if f.lower().endswith(_DOC_EXTS) and not f.startswith("~$"):
                return True
    return False


def _count_files(path: str) -> int:
    n = 0
    if not os.path.isdir(path):
        return 0
    for root, _, files in os.walk(path):
        for f in files:
            if f.lower().endswith(_DOC_EXTS) and not f.startswith("~$"):
                n += 1
    return n


def get_project_progress(project_name: str) -> dict:
    """
    Return a progress report for one project:
      - phases: list of {name, has_files, file_count}
      - completed / total / percent
      - recent: {name, phase, modified} for the most recently changed file
      - total_files
    """
    project_root = os.path.join(BASE_DIRECTORY, project_name)
    phases = []
    completed = 0
    total_files = 0
    most_recent = None  # (mtime, filename, phase)

    for phase in PHASE_FOLDERS:
        phase_path = os.path.join(project_root, phase)
        has_files = False
        count = 0
        if os.path.isdir(phase_path):
            for root, _, files in os.walk(phase_path):
                for f in files:
                    if f.lower().endswith(_DOC_EXTS) and not f.startswith("~$"):
                        has_files = True
                        count += 1
                        total_files += 1
                        full = os.path.join(root, f)
                        try:
                            mtime = os.path.getmtime(full)
                        except Exception:
                            mtime = 0
                        if most_recent is None or mtime > most_recent[0]:
                            most_recent = (mtime, f, phase)
        if has_files:
            completed += 1
        phases.append({"name": phase, "has_files": has_files, "file_count": count})

    total = len(PHASE_FOLDERS)
    percent = round((completed / total) * 100) if total else 0

    recent = None
    if most_recent:
        recent = {
            "name": most_recent[1],
            "phase": most_recent[2],
            "modified": time.strftime("%Y-%m-%d %H:%M", time.localtime(most_recent[0])),
        }

    return {
        "project": project_name,
        "phases": phases,
        "completed": completed,
        "total": total,
        "percent": percent,
        "recent": recent,
        "total_files": total_files,
    }


def get_all_progress() -> list[dict]:
    """Progress report for every project."""
    return [get_project_progress(p) for p in list_projects()]


# ── AI status summary (cached) ────────────────────────────────────────────────

def get_status_summary(project_name: str) -> str | None:
    """Return the cached AI status summary for a project, if any."""
    return _load_summaries().get(project_name, {}).get("summary")


def set_status_summary(project_name: str, summary: str) -> None:
    data = _load_summaries()
    data[project_name] = {"summary": summary, "updated": time.time()}
    _save_summaries(data)


def generate_status_summary(project_name: str) -> str | None:
    """
    Use AI to write a short plain-English status summary for a project,
    based on which phases have files. Caches the result.
    """
    from ai_client import gemini_call

    prog = get_project_progress(project_name)
    done_phases = [p["name"] for p in prog["phases"] if p["has_files"]]
    todo_phases = [p["name"] for p in prog["phases"] if not p["has_files"]]

    prompt = (
        "You are an HDOT project assistant. Write a 2-3 sentence plain-English "
        "status update for this project. Be specific and practical.\n\n"
        f"Project: {project_name}\n"
        f"Overall completion: {prog['percent']}% ({prog['completed']} of {prog['total']} phases have documents)\n"
        f"Phases WITH documents: {', '.join(done_phases) or 'none yet'}\n"
        f"Phases still empty: {', '.join(todo_phases) or 'none'}\n"
        f"Total documents filed: {prog['total_files']}\n\n"
        "Summarize where the project stands and what the likely next phase is. "
        "Do not use bullet points or markdown headers."
    )
    summary = gemini_call(prompt, temperature=0.3, timeout=30)
    if summary:
        set_status_summary(project_name, summary.strip())
    return summary