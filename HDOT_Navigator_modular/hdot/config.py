"""
config.py
All environment loading, path constants, and app-wide settings.
Import this first in every module that needs these values.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# ── Load .env files (shared → local → system) ────────────────────────────────
shared_env_path = Path(__file__).resolve().parent / ".env"
if shared_env_path.exists():
    load_dotenv(dotenv_path=shared_env_path)

local_env_path = Path.home() / ".hdot_navigator" / ".env"
if local_env_path.exists():
    load_dotenv(dotenv_path=local_env_path, override=True)

load_dotenv()

# ── Gemini ────────────────────────────────────────────────────────────────────
API_KEY       = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro"]

# ── Storage paths ─────────────────────────────────────────────────────────────
USER_HOME          = Path.home()
CHROMA_STORAGE_DIR = str(USER_HOME / ".hdot_navigator" / "chroma_db")
os.makedirs(CHROMA_STORAGE_DIR, exist_ok=True)

TEMPLATE_STORE_DIR = str(USER_HOME / ".hdot_navigator" / "doc_templates")
os.makedirs(TEMPLATE_STORE_DIR, exist_ok=True)

FEEDBACK_FILE = str(USER_HOME / ".hdot_navigator" / "doc_type_feedback.json")

# ── OneDrive shared folder (optional) ────────────────────────────────────────
_onedrive_raw = os.getenv("ONEDRIVE_SHARE_DIR", "").strip()
if _onedrive_raw and os.path.isdir(_onedrive_raw):
    ONEDRIVE_SHARE_DIR = _onedrive_raw
    TEMPLATE_STORE_DIR = os.path.join(_onedrive_raw, "doc_templates")
    FEEDBACK_FILE      = os.path.join(_onedrive_raw, "doc_type_feedback.json")
    os.makedirs(TEMPLATE_STORE_DIR, exist_ok=True)
else:
    ONEDRIVE_SHARE_DIR = ""

# ── ChromaDB collection names ────────────────────────────────────────────────
COLLECTION_NAME   = "enterprise_pdm_pool"
MANUAL_COLLECTION = "secretaries_manual_v2"

# ── File handling ─────────────────────────────────────────────────────────────
IMAGE_EXTENSIONS    = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")
DEFAULT_DEST_FOLDER = os.getenv("DEFAULT_DEST_FOLDER", "01 SCOPING").strip()

# ── Project folder template structure ────────────────────────────────────────
PROJECT_TEMPLATE_STRUCTURE = [
    "00 TABLE OF CONTENTS", "00-1 INTRODUCTION", "01 SCOPING", "02 PRE-DESIGN",
    "03 60% DESIGN", "04 PUBLIC INVOLVEMENT", "05 COORDINATION", "06 90% DESIGN",
    "07 100% DESIGN", "08 ADVERTISING", "09 BID OPENING", "10 DRAFTING COORDINATION",
    "12 COUNTY SPECIFIC PROC", "13 CANCELLING OR DEFERRING PROJECTS", "14 APPENDIX",
    "14 APPENDIX/14 Apdx Ch00I", "14 APPENDIX/14 Apdx Ch01", "14 APPENDIX/14 Apdx Ch02",
    "14 APPENDIX/14 Apdx Ch03", "14 APPENDIX/14 Apdx Ch04", "14 APPENDIX/14 Apdx Ch05",
    "14 APPENDIX/14 Apdx Ch06", "14 APPENDIX/14 Apdx Ch07", "14 APPENDIX/14 Apdx Ch08",
    "14 APPENDIX/14 Apdx Ch09", "14 APPENDIX/14 Apdx Ch10", "14 APPENDIX/14 Apdx Ch12",
    "14 APPENDIX/14 Apdx Ch13",
]

# ── Base directory (set after env is loaded) ──────────────────────────────────
BASE_DIRECTORY = os.getenv("BASE_DIRECTORY", "").strip()
