import requests
import os
import json
import hashlib
from datetime import datetime


AI_CACHE_FILE = "cache/ai_cache.json"
AI_HISTORY_FILE = "cache/ai_history.json"
AI_MODEL = os.getenv("DOCUMIND_AI_MODEL", "llama3.2:1b")
AI_TIMEOUT = int(os.getenv("DOCUMIND_AI_TIMEOUT", "120"))
AI_KEEP_ALIVE = os.getenv("DOCUMIND_AI_KEEP_ALIVE", "10m")
AI_NUM_PREDICT_SUMMARY = int(os.getenv("DOCUMIND_AI_SUMMARY_TOKENS", "500"))
AI_NUM_PREDICT_GENERATE = int(os.getenv("DOCUMIND_AI_GENERATE_TOKENS", "900"))
AI_HISTORY_MAX_ITEMS = int(os.getenv("DOCUMIND_AI_HISTORY_MAX_ITEMS", "300"))
_SESSION = requests.Session()
_AI_CACHE = None


def _load_ai_cache():
    global _AI_CACHE
    if _AI_CACHE is not None:
        return _AI_CACHE

    if os.path.exists(AI_CACHE_FILE):
        try:
            with open(AI_CACHE_FILE, "r", encoding="utf-8") as f:
                _AI_CACHE = json.load(f)
                return _AI_CACHE
        except Exception:
            _AI_CACHE = {}
            return _AI_CACHE

    _AI_CACHE = {}
    return _AI_CACHE


def _save_ai_cache(cache):
    try:
        os.makedirs(os.path.dirname(AI_CACHE_FILE), exist_ok=True)
        with open(AI_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception as e:
        print(f"AI cache save error: {e}")


def _load_ai_history():
    if os.path.exists(AI_HISTORY_FILE):
        try:
            with open(AI_HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            return []
    return []


def _save_ai_history(history_items):
    try:
        os.makedirs(os.path.dirname(AI_HISTORY_FILE), exist_ok=True)
        with open(AI_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history_items, f, ensure_ascii=False)
    except Exception as e:
        print(f"AI history save error: {e}")


def _append_ai_history(task, payload, response_text, cache_key, source):
    history = _load_ai_history()
    history.append({
        "timestamp": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "task": task,
        "model": AI_MODEL,
        "source": source,  # cache_hit ან model_call
        "cache_key": cache_key,
        "payload_preview": payload[:300],
        "response_preview": response_text[:500]
    })
    if len(history) > AI_HISTORY_MAX_ITEMS:
        history = history[-AI_HISTORY_MAX_ITEMS:]
    _save_ai_history(history)


def _make_cache_key(task, payload):
    payload_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return f"{task}:{AI_MODEL}:{payload_hash}"


def _call_model(prompt, num_predict):
    response = _SESSION.post(
        "http://localhost:11434/api/generate",
        json={
            "model": AI_MODEL,
            "prompt": prompt,
            "stream": False,
            "keep_alive": AI_KEEP_ALIVE,
            "options": {
                "num_predict": num_predict
            }
        },
        timeout=AI_TIMEOUT
    )
    response.raise_for_status()
    result = response.json()
    return result.get("response", "").strip()

def summarize_text(text):
    if len(text.strip()) < 100:
        return "Text is too short to summarize."

    text = text[:3000]

    cache = _load_ai_cache()
    cache_key = _make_cache_key("summarize", text)
    if cache_key in cache:
        _append_ai_history("summarize", text, cache[cache_key], cache_key, "cache_hit")
        return cache[cache_key]

    prompt = f"""You are a professional legal and business document analyst.
Analyze the following document and provide a structured summary:

**Parties Involved:** (who are the parties?)
**Purpose:** (what is this document about?)
**Key Dates:** (start date, end date, deadlines)
**Main Obligations:** (what must each party do?)
**Termination:** (how can it be terminated?)
**Other Important Points:** (anything else significant?)

Document:
{text}"""

    try:
        summary = _call_model(prompt, AI_NUM_PREDICT_SUMMARY)
        cache[cache_key] = summary
        _save_ai_cache(cache)
        _append_ai_history("summarize", text, summary, cache_key, "model_call")
        return summary
    except Exception as e:
        return f"Summarization error: {e}"

def summarize_document(doc):
    print(f"\nSummarizing: {doc['filename']}...")
    summary = summarize_text(doc['text'])
    return {
        'filename': doc['filename'],
        'summary': summary
    }

def generate_document(doc_type, details):
    payload = f"{doc_type}\n{details.strip()}"
    cache = _load_ai_cache()
    cache_key = _make_cache_key("generate", payload)
    if cache_key in cache:
        _append_ai_history("generate", payload, cache[cache_key], cache_key, "cache_hit")
        return cache[cache_key]

    prompt = f"""Create a sample template for a {doc_type} document for educational and training purposes only.
This is a fictional template with placeholder information.

Details:
{details}

Write the complete template document with all standard sections.
Use [PARTY A], [PARTY B], [DATE] as placeholders where needed.
Include all typical clauses and sections for this document type.
Format it professionally."""

    try:
        generated = _call_model(prompt, AI_NUM_PREDICT_GENERATE)
        cache[cache_key] = generated
        _save_ai_cache(cache)
        _append_ai_history("generate", payload, generated, cache_key, "model_call")
        return generated
    except Exception as e:
        return f"Generation error: {e}"