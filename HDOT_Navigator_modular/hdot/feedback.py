"""
feedback.py
Stores and retrieves user corrections to the AI document classifier.
These corrections are fed back into future classification prompts.
"""

import json
import os
from datetime import datetime

from config import FEEDBACK_FILE


def load_feedback() -> list:
    try:
        if os.path.exists(FEEDBACK_FILE):
            with open(FEEDBACK_FILE) as f:
                return json.load(f)
    except Exception:
        pass
    return []


def save_feedback(entry: dict) -> None:
    """Save a correction entry, ignoring exact duplicates."""
    entries = load_feedback()
    for existing in entries:
        if (existing.get("doc_snippet") == entry.get("doc_snippet") and
                existing.get("correct_type") == entry.get("correct_type")):
            return
    entries.append(entry)
    try:
        with open(FEEDBACK_FILE, "w") as f:
            json.dump(entries, f, indent=2)
    except Exception as e:
        print(f"[feedback] Save failed: {e}")


def delete_feedback(index: int) -> None:
    entries = load_feedback()
    if 0 <= index < len(entries):
        entries.pop(index)
        with open(FEEDBACK_FILE, "w") as f:
            json.dump(entries, f, indent=2)


def get_relevant_feedback(doc_snippet: str, max_examples: int = 5) -> list:
    """Return the most relevant past corrections for a given document snippet."""
    entries = load_feedback()
    if not entries:
        return []
    snippet_words = set(doc_snippet.upper().split())
    scored = []
    for entry in entries:
        ex_words = set(entry.get("doc_snippet", "").upper().split())
        score    = len(snippet_words & ex_words)
        scored.append((score, entry))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [e for s, e in scored[:max_examples] if s > 0]


def build_feedback_context(doc_snippet: str) -> str:
    """Build a prompt section from relevant past corrections."""
    examples = get_relevant_feedback(doc_snippet)
    if not examples:
        return ""
    lines = ["PAST CORRECTIONS (learn from these mistakes):"]
    for ex in examples:
        reason_part = f" Reason: {ex['reason']}" if ex.get("reason") else ""
        lines.append(
            f"- Document starting with \"{ex.get('doc_snippet','')[:80]}\" "
            f"was INCORRECTLY identified as \"{ex.get('wrong_type','')}\" "
            f"but the CORRECT type is \"{ex.get('correct_type','')}\".{reason_part}"
        )
    return "\n".join(lines)


def new_feedback_entry(
    doc_snippet: str,
    wrong_type: str,
    correct_type: str,
    reason: str,
    filename: str,
) -> dict:
    """Build a feedback entry dict ready to pass to save_feedback()."""
    return {
        "doc_snippet":  doc_snippet[:200],
        "wrong_type":   wrong_type,
        "correct_type": correct_type,
        "reason":       reason.strip(),
        "timestamp":    datetime.now().isoformat(),
        "filename":     filename,
    }
