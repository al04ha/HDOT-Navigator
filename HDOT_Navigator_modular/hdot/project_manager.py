"""
project_manager.py
Project folder management, file system utilities, Outlook COM integration,
and the AI document classification pipeline.
"""

import os
import re
import sys
import json
import shutil
import time
import hashlib
import zipfile
import tempfile
import subprocess
from pathlib import Path

import streamlit as st

from config import (
    BASE_DIRECTORY, CHROMA_STORAGE_DIR, IMAGE_EXTENSIONS,
    PROJECT_TEMPLATE_STRUCTURE,
)
from ai_client import gemini_call

try:
    import win32com.client
    import pythoncom
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


# ── File system helpers ───────────────────────────────────────────────────────

def get_subfolders(path: str) -> list:
    try:
        return sorted(d for d in os.listdir(path)
                      if os.path.isdir(os.path.join(path, d)) and d != ".trash")
    except Exception:
        return []


def get_all_folders(base: str) -> list:
    folders = [base]
    for root, dirs, _ in os.walk(base):
        dirs[:] = sorted(d for d in dirs if d != ".trash")
        for d in dirs:
            folders.append(os.path.join(root, d))
    return folders


def folder_display_name(abs_path: str, base: str) -> str:
    rel = os.path.relpath(abs_path, base)
    return "(Root)" if rel == "." else rel


def open_file(file_path: str) -> None:
    try:
        if sys.platform.startswith("win32"):
            os.startfile(os.path.abspath(file_path))
        else:
            subprocess.run(["xdg-open", file_path], check=True)
    except Exception as e:
        st.error(f"❌ Could not open: {e}")


def trash_item(path: str, trash_dir: str) -> bool:
    try:
        name = os.path.basename(path)
        dest = os.path.join(trash_dir, f"{name}_{int(time.time())}")
        shutil.move(path, dest)
        st.session_state["last_deleted"] = {
            "original_path": path,
            "trash_path":    dest,
            "name":          name,
            "deleted_at":    time.time(),
        }
        return True
    except Exception as e:
        st.error(f"Error: {e}")
        return False


# ── Project creation ──────────────────────────────────────────────────────────

def create_project_from_template(project_root: str) -> None:
    """Create a new project folder using template.zip or built-in structure."""
    template_zip = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.zip")
    if os.path.exists(template_zip):
        try:
            os.makedirs(project_root, exist_ok=True)
            with zipfile.ZipFile(template_zip, "r") as zf:
                members    = zf.namelist()
                root_prefix = members[0].split("/")[0] + "/" if "/" in members[0] else ""
                for member in members:
                    basename = os.path.basename(member)
                    if basename.startswith("~$") or not basename:
                        continue
                    relative = member[len(root_prefix):] if (root_prefix and member.startswith(root_prefix)) else member
                    if not relative:
                        continue
                    target = os.path.join(project_root, relative)
                    if member.endswith("/"):
                        os.makedirs(target, exist_ok=True)
                    else:
                        os.makedirs(os.path.dirname(target), exist_ok=True)
                        with zf.open(member) as src, open(target, "wb") as dst:
                            dst.write(src.read())
            return
        except Exception as e:
            print(f"[TEMPLATE] zip extraction failed: {e}")
    os.makedirs(project_root, exist_ok=True)
    for subfolder in PROJECT_TEMPLATE_STRUCTURE:
        os.makedirs(os.path.join(project_root, subfolder), exist_ok=True)


def get_existing_projects(base_dir: str) -> list:
    projects = []
    try:
        for item in os.listdir(base_dir):
            full = os.path.join(base_dir, item)
            if os.path.isdir(full) and item != ".trash":
                projects.append({"folder_name": item, "abs_path": full})
    except Exception:
        pass
    return projects


# ── AI document classification ────────────────────────────────────────────────

def _cache_key(filename: str, text_preview: str) -> str:
    h = hashlib.md5((filename + text_preview[:100]).encode()).hexdigest()[:12]
    return os.path.join(CHROMA_STORAGE_DIR, f".ai_cache_{h}.json")


def analyze_document_with_ai(
    filename: str,
    extracted_text: str,
    email_subject: str = "",
    base_directory: str = "",
) -> dict | None:
    """Classify a document using Gemini and return structured filing metadata."""
    base_dir   = base_directory or BASE_DIRECTORY
    cache_file = _cache_key(filename, extracted_text)
    try:
        if os.path.exists(cache_file):
            with open(cache_file) as f:
                return json.load(f)
    except Exception:
        pass

    preview_text = extracted_text[:400]
    all_existing = get_existing_projects(base_dir)
    search_blob  = (filename + " " + preview_text + " " + (email_subject or "")).lower()
    relevant     = [p for p in all_existing
                    if any(word.lower() in search_blob
                           for word in re.split(r"[_\-\s]+", p["folder_name"]) if len(word) > 3)]
    folder_candidates = relevant or all_existing
    existing_list = ", ".join(p["folder_name"] for p in folder_candidates[:20]) or "none"

    prompt = (
        f'HDOT document classifier. JSON only, no markdown.\n'
        f'Filename: {filename}\nSubject: {email_subject or ""}\n'
        f'Text: {preview_text}\nExisting folders: {existing_list}\n'
        f'Return: {{"project_number":string|null,"project_name":string|null,'
        f'"document_type":string,"suggested_subfolder":"01 SCOPING"|"02 ENVIRONMENTAL"|'
        f'"03 DESIGN"|"04 ROW"|"05 UTILITIES"|"06 BID"|"07 CONSTRUCTION"|'
        f'"08 CLOSEOUT"|"09 CORRESPONDENCE"|"10 PHOTOS",'
        f'"matched_existing_project_folder":string|null,"confidence":0-100,"reasoning":string}}'
    )

    text = gemini_call(prompt, temperature=0.1)
    if text is None:
        st.session_state["ai_last_error"] = "All Gemini models failed or quota exhausted."
        return None
    try:
        text   = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text   = re.sub(r"\s*```$", "", text).strip()
        result = json.loads(text)
        st.session_state.pop("ai_last_error", None)
        os.makedirs(CHROMA_STORAGE_DIR, exist_ok=True)
        with open(cache_file, "w") as f:
            json.dump(result, f)
        return result
    except Exception as e:
        st.session_state["ai_last_error"] = f"JSON parse error: {e}\nRaw: {text[:300]}"
        return None


def resolve_destination_folder(ai_result: dict | None, base_directory: str = "") -> dict:
    """Turn AI classification output into an absolute destination path."""
    base_dir = base_directory or BASE_DIRECTORY
    if not ai_result:
        return {"abs_path": None, "display": "(could not determine)", "exists": False, "new_folder_name": None}

    matched_folder = ai_result.get("matched_existing_project_folder")
    subfolder      = ai_result.get("suggested_subfolder", "01 SCOPING")
    proj_num       = ai_result.get("project_number") or ""
    proj_name      = ai_result.get("project_name") or ""

    if matched_folder:
        candidate = os.path.join(base_dir, matched_folder, subfolder)
        if os.path.isdir(candidate):
            return {"abs_path": candidate, "display": os.path.relpath(candidate, base_dir), "exists": True, "new_folder_name": None}
        project_root = os.path.join(base_dir, matched_folder)
        if os.path.isdir(project_root):
            return {"abs_path": candidate, "display": os.path.relpath(candidate, base_dir), "exists": False, "new_folder_name": None, "create_subfolder_only": True}

    if proj_num or proj_name:
        parts           = [p for p in [proj_num.strip(), proj_name.strip()] if p]
        new_folder_name = "_".join(parts)
        proposed_path   = os.path.join(base_dir, new_folder_name, subfolder)
        return {"abs_path": proposed_path, "display": os.path.relpath(proposed_path, base_dir), "exists": False, "new_folder_name": new_folder_name}

    fallback = os.path.join(base_dir, subfolder)
    return {"abs_path": fallback, "display": subfolder, "exists": os.path.isdir(fallback), "new_folder_name": None}


# ── Outlook COM ───────────────────────────────────────────────────────────────

def get_active_email_attachments():
    """Connect to active Outlook email and return (subject, attachments, error)."""
    if not COM_AVAILABLE:
        return None, [], "Windows COM layer unavailable."
    try:
        pythoncom.CoInitialize()
        outlook   = win32com.client.GetObject(Class="Outlook.Application")
        inspector = outlook.ActiveInspector()
        if not inspector:
            explorer = outlook.ActiveExplorer()
            if explorer and explorer.Selection.Count > 0:
                mail_item = explorer.Selection.Item(1)
            else:
                return None, [], "No active email detected in Outlook."
        else:
            mail_item = inspector.CurrentItem

        subject  = getattr(mail_item, "Subject", "No Subject")
        SAVEABLE = IMAGE_EXTENSIONS + (".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xltx", ".xltm")
        attachment_list = []
        for i in range(1, mail_item.Attachments.Count + 1):
            att   = mail_item.Attachments.Item(i)
            name  = att.FileName
            entry = {"name": name, "index": i, "preview_path": None}
            if name.lower().endswith(SAVEABLE):
                suffix = os.path.splitext(name)[1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode="wb") as tmp:
                    tmp_path = tmp.name
                att.SaveAsFile(tmp_path)
                entry["preview_path"] = tmp_path
            attachment_list.append(entry)
        return subject, attachment_list, None
    except Exception as ex:
        return None, [], f"Outlook error: {ex}"
    finally:
        pythoncom.CoUninitialize()


def download_outlook_attachment(attachment_index: int, absolute_file_path: str):
    """Save a specific attachment from the active Outlook email to disk."""
    if not COM_AVAILABLE:
        return False, "COM Layer Unavailable"
    try:
        pythoncom.CoInitialize()
        outlook   = win32com.client.GetObject(Class="Outlook.Application")
        inspector = outlook.ActiveInspector()
        mail_item = inspector.CurrentItem if inspector else outlook.ActiveExplorer().Selection.Item(1)
        mail_item.Attachments.Item(attachment_index).SaveAsFile(absolute_file_path)
        return True, None
    except Exception as ex:
        return False, str(ex)
    finally:
        pythoncom.CoUninitialize()
