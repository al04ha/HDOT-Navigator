"""
ai_client.py
All communication with the Gemini API: text generation and embeddings.
"""

import hashlib
import time
import requests
import streamlit as st
from collections import OrderedDict
from datetime import datetime
from typing import Optional

from config import API_KEY, GEMINI_MODELS

# ── LRU Cache ────────────────────────────────────────────────────────────────
# Bounded to MAX_CACHE_SIZE entries — evicts the oldest entry when full.
# Using OrderedDict so move_to_end + popitem(last=False) gives O(1) LRU.
MAX_CACHE_SIZE = 128

class _LRUCache:
    def __init__(self, maxsize: int = MAX_CACHE_SIZE):
        self._store: OrderedDict = OrderedDict()
        self._maxsize = maxsize

    def get(self, key):
        if key not in self._store:
            return None
        self._store.move_to_end(key)
        return self._store[key]

    def set(self, key, value):
        if key in self._store:
            self._store.move_to_end(key)
        self._store[key] = value
        if len(self._store) > self._maxsize:
            self._store.popitem(last=False)

    def pop(self, key, default=None):
        return self._store.pop(key, default)

    def clear(self):
        self._store.clear()

    def __contains__(self, key):
        return key in self._store


_gemini_cache: _LRUCache = _LRUCache()

# ── Connection pool ───────────────────────────────────────────────────────────
# A single requests.Session reuses TCP connections across all calls in a run,
# saving a ~50-100ms TLS handshake on every successive call.
_session = requests.Session()
_session.headers.update({"Content-Type": "application/json"})


def gemini_call(
    prompt: str,
    temperature: float = 0.1,
    timeout: int = 30,
    max_tokens: Optional[int] = None,
) -> Optional[str]:
    """Send a prompt to Gemini and return the text response, or None on failure.

    `max_tokens` optionally caps the output length.  When None (the default),
    no maxOutputTokens is sent and Gemini uses its natural limit (~8192).
    Pass an explicit value for calls that need a short answer (e.g. 50 for
    doc-type classification) to reduce latency.
    """
    cache_key = (hashlib.md5(prompt.encode()).hexdigest()[:16], round(temperature, 2), max_tokens)
    cached = _gemini_cache.get(cache_key)
    if cached is not None:
        return cached

    if "_gemini_log" not in st.session_state:
        st.session_state["_gemini_log"] = []

    log_entry = {
        "timestamp": datetime.now().strftime("%H:%M:%S"),
        "prompt": prompt,
        "prompt_chars": len(prompt),
        "model": "—",
        "temperature": temperature,
        "max_tokens": max_tokens,
        "status": "pending",
        "response": None,
        "error": None,
    }

    for model in GEMINI_MODELS:
        log_entry["model"] = model
        url = (
            f"https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent?key={API_KEY}"
        )
        gen_config: dict = {"temperature": temperature}
        if max_tokens is not None:
            gen_config["maxOutputTokens"] = max_tokens
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": gen_config,
        }

        for attempt in range(2):
            try:
                resp = _session.post(url, json=payload, timeout=timeout)
                raw = resp.json()

                if resp.status_code == 200 and "candidates" in raw:
                    result = raw["candidates"][0]["content"]["parts"][0]["text"].strip()
                    _gemini_cache.set(cache_key, result)
                    log_entry["status"]   = "✅ success"
                    log_entry["response"] = result[:200]   # keep log compact
                    _append_log(log_entry)
                    return result

                if resp.status_code == 429:
                    violations = []
                    try:
                        for detail in raw.get("error", {}).get("details", []):
                            if detail.get("@type", "").endswith("QuotaFailure"):
                                violations = [v.get("quotaId", "") for v in detail.get("violations", [])]
                    except Exception:
                        pass
                    if any("PerDay" in v for v in violations):
                        log_entry["status"] = "❌ daily quota"
                        log_entry["error"]  = "Daily quota exceeded"
                        break
                    # Per-minute rate limit — brief pause then retry
                    time.sleep(8)
                    continue

                log_entry["status"] = f"❌ HTTP {resp.status_code}"
                log_entry["error"]  = str(raw.get("error", resp.text))[:200]
                break

            except requests.Timeout:
                log_entry["status"] = "❌ timeout"
                log_entry["error"]  = f"No response within {timeout}s"
                break
            except Exception as e:
                log_entry["status"] = "❌ exception"
                log_entry["error"]  = str(e)[:200]
                break

    _append_log(log_entry)
    return None


def _append_log(entry: dict) -> None:
    """Append to the in-session API log, capping at 200 entries to prevent
    unbounded session-state growth over long sessions."""
    log = st.session_state.setdefault("_gemini_log", [])
    log.append(entry)
    if len(log) > 200:
        # Drop the oldest 50 entries in one slice instead of popping one at a time
        st.session_state["_gemini_log"] = log[-150:]


# ── Embedding cache ───────────────────────────────────────────────────────────
# Embedding the same text repeatedly (e.g. the same KB query on re-renders)
# wastes API quota.  A simple dict keyed on the first 200 chars of text is
# sufficient since embeddings are deterministic.
_embed_cache: dict = {}


def get_embedding(text: str) -> Optional[list]:
    """Return an embedding vector for the given text, or None on failure."""
    key = text[:200]
    if key in _embed_cache:
        return _embed_cache[key]

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"gemini-embedding-2:embedContent?key={API_KEY}"
    )
    try:
        resp = _session.post(
            url,
            json={"content": {"parts": [{"text": text[:2000]}]}},
            timeout=15,
        ).json()
        result = resp["embedding"]["values"] if "embedding" in resp else None
        if result is not None:
            _embed_cache[key] = result
        return result
    except Exception:
        return None


def clear_cache() -> None:
    """Wipe both the Gemini response cache and the embedding cache."""
    _gemini_cache.clear()
    _embed_cache.clear()