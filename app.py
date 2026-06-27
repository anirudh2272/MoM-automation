"""
MoM Live AI - Multilingual Minutes of Meeting Automation
Whisper transcription + Gemini AI analysis + 52-language translation + DOCX export
"""

import gc
import html
import json
import os
import re
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from faster_whisper import WhisperModel

# ── Optional dependencies ──────────────────────────────────────
try:
    from deep_translator import GoogleTranslator
    TRANSLATE_OK = True
except Exception:
    GoogleTranslator = None
    TRANSLATE_OK = False

try:
    from google import genai as _genai
    GEMINI_OK = True
except Exception:
    _genai = None
    GEMINI_OK = False

# ══════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════
APP_VER      = "2.0"
MAX_MB       = 25
TMP          = "/tmp"
GEMINI_MODEL = "gemini-2.5-flash"

# 52 languages supported by Whisper + Google Translate
LANGS: Dict[str, str] = {
    "af":"Afrikaans",  "ar":"Arabic",      "az":"Azerbaijani", "bn":"Bengali",
    "cs":"Czech",      "de":"German",      "en":"English",     "es":"Spanish",
    "et":"Estonian",   "fa":"Persian",     "fi":"Finnish",     "fr":"French",
    "gl":"Galician",   "gu":"Gujarati",    "he":"Hebrew",      "hi":"Hindi",
    "hr":"Croatian",   "id":"Indonesian",  "it":"Italian",     "ja":"Japanese",
    "ka":"Georgian",   "kk":"Kazakh",      "km":"Khmer",       "ko":"Korean",
    "lt":"Lithuanian", "lv":"Latvian",     "mk":"Macedonian",  "ml":"Malayalam",
    "mn":"Mongolian",  "mr":"Marathi",     "my":"Burmese",     "ne":"Nepali",
    "nl":"Dutch",      "pl":"Polish",      "ps":"Pashto",      "pt":"Portuguese",
    "ro":"Romanian",   "ru":"Russian",     "si":"Sinhala",     "sl":"Slovenian",
    "sv":"Swedish",    "sw":"Swahili",     "ta":"Tamil",       "te":"Telugu",
    "th":"Thai",       "tl":"Filipino",    "tr":"Turkish",     "uk":"Ukrainian",
    "ur":"Urdu",       "vi":"Vietnamese",  "xh":"Xhosa",       "zh":"Chinese",
}

# Languages that need the "base" model (tiny fails on these)
HARD_LANGS = {
    "te","hi","ta","ml","mr","gu","bn","ur","si","ne",
    "ar","fa","ps","he","ka","my","km","th","mn","kk",
    "sw","af","xh","tl","az","vi","lv","lt","sl","gl","mk","hr",
}

# Google Translate code remapping
GT_REMAP = {"zh":"zh-CN", "he":"iw"}

# ══════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="MoM Live AI",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════
#  APPLE-STYLE CSS
# ══════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap');

:root {
    --sf:    -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Inter", "Helvetica Neue", sans-serif;
    --ink:   #1d1d1f;
    --soft:  #6e6e73;
    --line:  #d2d2d7;
    --bg:    #ffffff;
    --gray:  #f5f5f7;
    --blue:  #0071e3;
    --blue2: #0077ed;
    --r:     18px;
    --r-sm:  12px;
}

/* Base */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html, body, [class*="css"] {
    font-family: var(--sf) !important;
    -webkit-font-smoothing: antialiased;
    letter-spacing: -0.01em;
    color: var(--ink);
}
.stApp { background: var(--bg) !important; }
.main .block-container { max-width: 960px; padding: 0 1.5rem 5rem; }

/* Typography */
p, li            { color: var(--soft) !important; line-height: 1.55; font-size: .97rem; }
span, label      { color: var(--ink) !important; }
h1,h2,h3,h4     { color: var(--ink) !important; letter-spacing: -0.02em; }
strong, b        { color: var(--ink) !important; font-weight: 600; }
hr               { border: none !important; height: 1px !important; background: var(--line) !important; margin: 2.5rem 0 !important; }

/* Hero */
.hero            { text-align: center; padding: 3.5rem 0 2.5rem; }
.hero-eyebrow    { display: inline-block; font-size: .78rem; font-weight: 600; color: var(--blue); margin-bottom: .9rem; letter-spacing: .01em; }
.hero-title      { font-size: clamp(2.4rem, 5.5vw, 4rem); font-weight: 700; color: var(--ink); line-height: 1.05; letter-spacing: -0.035em; margin-bottom: 1rem; }
.hero-title em   { font-style: normal; color: var(--blue); }
.hero-sub        { font-size: 1.18rem; font-weight: 400; color: var(--soft) !important; max-width: 580px; margin: 0 auto; line-height: 1.45; letter-spacing: -.01em; }

/* Stat cards */
.stat-grid   { display: grid; grid-template-columns: repeat(4,1fr); gap: .85rem; margin: 2rem 0 2.5rem; }
.stat-card   { background: var(--gray); border-radius: var(--r); padding: 1.5rem 1rem; text-align: center; transition: transform .2s ease; border: none; }
.stat-card:hover { transform: translateY(-2px); }
.stat-icon   { font-size: 1.45rem; display: block; margin-bottom: .45rem; }
.stat-val    { font-size: 1.85rem; font-weight: 600; color: var(--ink); line-height: 1; letter-spacing: -.03em; }
.stat-lbl    { font-size: .72rem; font-weight: 500; color: var(--soft); margin-top: .45rem; }

/* Section headers */
.sec          { font-size: 1.75rem; font-weight: 600; color: var(--ink); margin: 2.8rem 0 1.1rem; letter-spacing: -.025em; }

/* Cards */
.card         { background: var(--gray); border-radius: var(--r); padding: 1.35rem 1.5rem; color: var(--ink); line-height: 1.6; font-size: .97rem; white-space: pre-wrap; word-break: break-word; }
.card-sm      { background: var(--gray); border-radius: var(--r-sm); padding: .85rem 1rem; margin-bottom: .5rem; font-size: .95rem; color: var(--ink); line-height: 1.5; }

/* Result boxes */
.res-grid     { display: grid; grid-template-columns: repeat(5,1fr); gap: .85rem; margin-bottom: 2rem; }
.res-box      { background: var(--gray); border-radius: var(--r); padding: 1.35rem 1rem; text-align: center; transition: transform .2s; }
.res-box:hover{ transform: translateY(-2px); }
.res-val      { font-size: 2.3rem; font-weight: 700; color: var(--ink); line-height: 1; letter-spacing: -.03em; }
.res-lbl      { font-size: .72rem; font-weight: 500; color: var(--soft); margin-top: .5rem; }

/* AI insight boxes */
.ai-grid      { display: grid; grid-template-columns: 1fr 1fr; gap: .85rem; margin-bottom: 1.2rem; }
.ai-box       { background: var(--gray); border-radius: var(--r); padding: 1.4rem; }
.ai-icon      { font-size: 1.4rem; margin-bottom: .5rem; display: block; }
.ai-val       { font-size: 1.7rem; font-weight: 700; color: var(--ink); letter-spacing: -.025em; line-height: 1.1; }
.ai-lbl       { font-size: .7rem; font-weight: 500; color: var(--soft); margin: .3rem 0 .5rem; text-transform: uppercase; letter-spacing: .04em; }
.ai-desc      { font-size: .85rem; color: var(--soft); line-height: 1.45; }

/* Severity chips */
.chip         { display: inline-block; border-radius: 6px; padding: .2rem .6rem; font-size: .75rem; font-weight: 600; margin-right: .5rem; }
.chip-high    { background: #fdecea; color: #b3261e; }
.chip-med     { background: #fff4e5; color: #9a5b00; }
.chip-low     { background: #e9f6ec; color: #1c7a37; }

/* Pills */
.pill         { display: inline-flex; align-items: center; border-radius: 980px; padding: .35rem .85rem; font-size: .8rem; font-weight: 500; background: var(--gray); color: var(--ink); margin: .2rem .2rem 0 0; }

/* Empty state */
.empty        { text-align: center; padding: 5rem 2rem; background: var(--gray); border-radius: 24px; }
.empty-icon   { font-size: 3.5rem; display: block; margin-bottom: 1rem; }
.empty-title  { font-size: 1.8rem; font-weight: 600; color: var(--ink); letter-spacing: -.025em; margin-bottom: .5rem; }
.empty-sub    { color: var(--soft) !important; font-size: 1.02rem; }

/* Sidebar */
section[data-testid="stSidebar"]                           { background: var(--gray) !important; border-right: 1px solid var(--line) !important; }
section[data-testid="stSidebar"] *                         { color: var(--ink) !important; }
section[data-testid="stSidebar"] .stMarkdown p            { color: var(--soft) !important; font-size: .85rem !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3                        { font-weight: 600 !important; letter-spacing: -.02em !important; }

/* Primary button */
.stButton > button {
    background: var(--blue) !important; color: #fff !important;
    border: none !important; border-radius: 980px !important;
    font-family: var(--sf) !important; font-weight: 500 !important;
    font-size: 1rem !important; padding: .85rem 2rem !important;
    width: 100% !important; box-shadow: none !important;
    letter-spacing: -.01em !important; transition: background .15s ease !important;
}
.stButton > button:hover  { background: var(--blue2) !important; }
.stButton > button:active { background: #005bbf !important; }

/* Download button */
.stDownloadButton > button {
    background: var(--blue) !important; color: #fff !important;
    border: none !important; border-radius: 980px !important;
    font-weight: 500 !important; font-size: 1rem !important;
    padding: .85rem 2rem !important; width: 100% !important;
    box-shadow: none !important; transition: background .15s !important;
}
.stDownloadButton > button:hover { background: var(--blue2) !important; }

/* File uploader */
.stFileUploader > div {
    background: var(--gray) !important;
    border: 1.5px dashed var(--line) !important;
    border-radius: var(--r) !important; transition: border-color .2s !important;
}
.stFileUploader > div:hover { border-color: var(--blue) !important; }

/* Inputs / Selects */
.stSelectbox > div > div,
.stTextInput > div > div {
    background: var(--bg) !important;
    border: 1px solid var(--line) !important;
    border-radius: var(--r-sm) !important; color: var(--ink) !important;
}
.stTextInput > div > div:focus-within,
.stSelectbox > div > div:focus-within {
    border-color: var(--blue) !important;
    box-shadow: 0 0 0 3px rgba(0,113,227,.15) !important;
    outline: none !important;
}

/* Tabs: segmented control */
.stTabs [data-baseweb="tab-list"] {
    background: var(--gray); border-radius: 980px;
    padding: 3px; border: none; gap: 3px;
}
.stTabs [data-baseweb="tab"] {
    color: var(--soft) !important; border-radius: 980px !important;
    padding: .45rem 1.35rem !important; font-weight: 500 !important;
    font-size: .9rem !important;
}
.stTabs [aria-selected="true"] {
    background: var(--bg) !important; color: var(--ink) !important;
    box-shadow: 0 1px 4px rgba(0,0,0,.1) !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: var(--gray) !important; border: none !important;
    border-radius: var(--r-sm) !important; color: var(--ink) !important;
    font-weight: 500 !important; font-size: .9rem !important;
}
.streamlit-expanderHeader:hover { background: #ececf0 !important; }
.streamlit-expanderContent {
    background: var(--bg) !important;
    border: 1px solid var(--line) !important;
    border-top: none !important; border-radius: 0 0 var(--r-sm) var(--r-sm) !important;
}

/* Alerts */
.stAlert {
    background: var(--gray) !important;
    border: 1px solid var(--line) !important;
    border-radius: var(--r-sm) !important; color: var(--ink) !important;
}

/* Progress bar */
.stProgress > div > div { background: var(--blue) !important; border-radius: 980px !important; }
.stProgress > div       { background: #e5e5ea !important; border-radius: 980px !important; }

/* Radio */
.stRadio > div { gap: .5rem !important; }
.stRadio label { background: var(--gray) !important; border-radius: 980px !important; padding: .35rem .9rem !important; font-weight: 500 !important; }

/* Slider */
.stSlider > div > div > div > div { background: var(--blue) !important; }

/* Scrollbar */
::-webkit-scrollbar       { width: 7px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #c7c7cc; border-radius: 980px; }

/* Audio */
audio { width: 100%; border-radius: var(--r-sm); }

/* Hide streamlit branding */
#MainMenu, footer, header { visibility: hidden !important; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════
def sf(name: str) -> str:
    """Safe filename."""
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)[:100] or "file"

def clean(text: str) -> str:
    """Strip DOCX-illegal control chars."""
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text or ""))

def esc(text: str) -> str:
    return html.escape(text or "")

def lname(code: str) -> str:
    return LANGS.get(code, code)


# ══════════════════════════════════════════════════════════════
#  AUDIO
# ══════════════════════════════════════════════════════════════
def save_uploaded_file(uploaded) -> str:
    path = os.path.join(TMP, f"{int(time.time())}_{sf(uploaded.name)}")
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return path


def convert_to_wav(src: str) -> Tuple[str, str]:
    """Convert to 16kHz mono WAV. Returns (path, warning)."""
    dst = src.rsplit(".", 1)[0] + "_16k.wav"
    cmd = ["ffmpeg", "-y", "-i", src,
           "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1",
           "-loglevel", "error", dst]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if r.returncode != 0 or not os.path.exists(dst) or os.path.getsize(dst) < 1000:
            return src, f"FFmpeg note: {r.stderr[:120]}. Using original file."
        return dst, ""
    except FileNotFoundError:
        return src, "FFmpeg missing — add 'ffmpeg' to packages.txt."
    except subprocess.TimeoutExpired:
        return src, "FFmpeg timed out. Using original file."
    except Exception as e:
        return src, f"Conversion error: {str(e)[:80]}"


# ══════════════════════════════════════════════════════════════
#  TRANSCRIPTION
# ══════════════════════════════════════════════════════════════
@st.cache_resource(show_spinner=False)
def load_whisper_model(size: str) -> WhisperModel:
    return WhisperModel(size, device="cpu", compute_type="int8")


def transcribe_audio(path: str, model_size: str,
                     lang_hint: str = "auto") -> Tuple[str, List[dict], str, float, float]:
    """
    Two-pass transcription.
    Pass 1 — standard (beam=5, vad off, threshold=0.1).
    Pass 2 — ultra-lenient fallback when pass 1 returns nothing.
    Returns: transcript, segments, detected_lang, lang_prob_pct, duration_sec
    """
    model = load_whisper_model(model_size)
    lang  = None if lang_hint == "auto" else lang_hint

    def _run(beam, nst, lpt=None):
        kw = dict(beam_size=beam, language=lang,
                  vad_filter=False,
                  condition_on_previous_text=False,
                  no_speech_threshold=nst)
        if lpt is not None:
            kw["log_prob_threshold"] = lpt
        raw, info = model.transcribe(path, **kw)
        segs = [{"start": round(float(s.start), 2),
                 "end":   round(float(s.end),   2),
                 "text":  s.text.strip()}
                for s in raw if (s.text or "").strip()]
        return segs, info

    segs, info = _run(beam=5, nst=0.1)
    if not segs:
        segs, info = _run(beam=1, nst=0.05, lpt=-2.0)

    transcript = " ".join(s["text"] for s in segs)
    det  = getattr(info, "language",             None) or lang_hint or "en"
    prob = float(getattr(info, "language_probability", 0.0) or 0.0) * 100
    dur  = float(getattr(info, "duration",             0.0) or 0.0)
    return transcript, segs, det, round(prob, 1), dur


# ══════════════════════════════════════════════════════════════
#  DIARIZATION
# ══════════════════════════════════════════════════════════════
def diarize_segments(segs: List[dict], n: int) -> List[dict]:
    """Pause-based heuristic speaker assignment."""
    if not segs:
        return []
    out, spk, last = [], 1, segs[0]["start"]
    for s in segs:
        if n > 1 and (s["start"] - last) >= 1.5:
            spk = (spk % n) + 1
        out.append({**s, "speaker": f"Speaker {spk}"})
        last = s["end"]
    return out


# ══════════════════════════════════════════════════════════════
#  RULE-BASED MEETING INTELLIGENCE (fallback when no Gemini key)
# ══════════════════════════════════════════════════════════════
def _sents(text: str) -> List[str]:
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", (text or "").strip())
            if len(s.strip()) > 4]


def generate_summary(transcript: str, n: int = 5) -> str:
    sents = _sents(transcript)
    if not sents:
        return "No transcript available."
    if len(sents) <= n:
        return " ".join(sents)
    kw = ["decided","agreed","action","task","deadline","next","will","should",
          "review","send","prepare","follow","plan","risk","blocker","approval"]
    scored = [(sum(2 for k in kw if k in s.lower()) + min(len(s.split())/18, 2)
               + (1.8 if i == 0 else 0) + (.8 if i == len(sents)-1 else 0), i, s)
              for i, s in enumerate(sents)]
    return " ".join(x[2] for x in sorted(sorted(scored, reverse=True)[:n], key=lambda x: x[1]))


def _deadline(text: str) -> str:
    for pat in [r"\bby\s+(tomorrow|today|monday|tuesday|wednesday|thursday|friday)\b",
                r"\b(next\s+week|this\s+week|end\s+of\s+week|EOW|EOD)\b",
                r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"]:
        m = re.search(pat, text, re.I)
        if m:
            return m.group(0).strip()
    return "Not specified"


def extract_action_items(segs: List[dict]) -> List[dict]:
    kw = ["will","should","need to","must","please","send","schedule","review",
          "prepare","follow up","complete","update","share","create","submit",
          "confirm","make sure","assign","finish","work on","deliver"]
    out, seen = [], set()
    for s in segs:
        t, lo = s.get("text","").strip(), s.get("text","").lower()
        if len(t.split()) < 4 or not any(k in lo for k in kw):
            continue
        key = re.sub(r"\W+","",lo)[:80]
        if key in seen:
            continue
        seen.add(key)
        out.append({"speaker": s.get("speaker","Unknown"), "text": t,
                    "time": f"{s.get('start',0)}s",
                    "owner": s.get("speaker","Unknown"),
                    "deadline": _deadline(t)})
        if len(out) >= 12:
            break
    return out


def extract_decisions(segs: List[dict]) -> List[dict]:
    kw = ["decided","agreed","approved","confirmed","finalized","accepted",
          "we will","moving forward","going with","resolved","settled","conclusion"]
    out, seen = [], set()
    for s in segs:
        t, lo = s.get("text","").strip(), s.get("text","").lower()
        if len(t.split()) < 4 or not any(k in lo for k in kw):
            continue
        key = re.sub(r"\W+","",lo)[:80]
        if key in seen:
            continue
        seen.add(key)
        out.append({"speaker": s.get("speaker","Unknown"), "text": t,
                    "time": f"{s.get('start',0)}s"})
        if len(out) >= 10:
            break
    return out


def extract_open_questions(segs: List[dict]) -> List[dict]:
    starters = ("what","why","how","when","where","who","can","could",
                "should","do","does","did","is","are")
    phrases  = ["not sure","pending","to be discussed","need confirmation",
                "we need to clarify","tbd","follow up on","still unclear"]
    out, seen = [], set()
    for s in segs:
        t, lo = s.get("text","").strip(), s.get("text","").lower()
        first = lo.split()[0] if lo.split() else ""
        if not ("?" in t or first in starters or any(p in lo for p in phrases)):
            continue
        if len(t.split()) < 4:
            continue
        key = re.sub(r"\W+","",lo)[:80]
        if key in seen:
            continue
        seen.add(key)
        out.append({"speaker": s.get("speaker","Unknown"), "text": t,
                    "time": f"{s.get('start',0)}s"})
        if len(out) >= 10:
            break
    return out


# ══════════════════════════════════════════════════════════════
#  GEMINI AI  (optional — user supplies their own free key)
# ══════════════════════════════════════════════════════════════
def get_gemini_client(api_key: str):
    if not api_key or not api_key.strip() or not GEMINI_OK:
        return None
    try:
        return _genai.Client(api_key=api_key.strip())
    except Exception:
        return None


def analyze_with_gemini(client, transcript: str) -> Optional[dict]:
    """Single Gemini call returns all meeting intelligence as JSON."""
    if client is None or not transcript.strip():
        return None
    text = transcript[:16000]
    prompt = f"""Analyze this meeting transcript and return ONLY a valid JSON object.

{{
  "summary": "3-5 sentence executive summary",
  "action_items": [{{"text":"task","owner":"person or Unspecified","deadline":"deadline or Not specified"}}],
  "decisions": [{{"text":"decision made"}}],
  "open_questions": [{{"text":"unresolved question or pending item"}}],
  "sentiment": {{"tone":"Positive|Neutral|Tense|Mixed","explanation":"one sentence"}},
  "effectiveness": {{"score":0-100,"rationale":"one sentence"}},
  "risks": [{{"text":"risk or blocker","severity":"High|Medium|Low"}}]
}}

Infer owners and deadlines from context. Return ONLY the JSON object.

TRANSCRIPT:
\"\"\"{text}\"\"\"
"""
    try:
        try:
            from google.genai import types as _t
            cfg = _t.GenerateContentConfig(
                temperature=0.2, max_output_tokens=3000,
                response_mime_type="application/json")
            resp = client.models.generate_content(
                model=GEMINI_MODEL, contents=prompt, config=cfg)
        except Exception:
            resp = client.models.generate_content(
                model=GEMINI_MODEL, contents=prompt)

        raw = (resp.text or "").strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        return data if "summary" in data else None
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
#  TRANSLATION  (Google Translate — all 52×52 pairs)
# ══════════════════════════════════════════════════════════════
def translate_text_safe(text: str, tgt: str, src: str = "auto") -> str:
    if not text or not TRANSLATE_OK:
        return text or ""
    if src != "auto" and src == tgt:
        return text
    tgt_code = GT_REMAP.get(tgt, tgt)
    src_code = "auto" if src == "auto" else GT_REMAP.get(src, src)
    try:
        chunks, buf = [], text
        while buf:
            chunk = buf[:4500]
            if len(buf) > 4500:
                cut = max(chunk.rfind(". "), chunk.rfind("? "), chunk.rfind("! "))
                if cut > 300:
                    chunk = chunk[:cut + 1]
            chunks.append(chunk.strip())
            buf = buf[len(chunk):].strip()
        parts = [GoogleTranslator(source=src_code, target=tgt_code).translate(c)
                 for c in chunks if c]
        return " ".join(p for p in parts if p)
    except Exception:
        return text


def translate_list(items: List[dict], tgt: str, src: str) -> List[dict]:
    for item in items:
        item["tr"] = translate_text_safe(item.get("text",""), tgt, src)
    return items


# ══════════════════════════════════════════════════════════════
#  DOCX EXPORT
# ══════════════════════════════════════════════════════════════
def _bg(cell, hex_color: str):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def _banner(doc, title: str, color: str = "1F497D"):
    t = doc.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    _bg(c, color)
    run = c.paragraphs[0].add_run("  " + clean(title))
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(255, 255, 255)
    doc.add_paragraph()


def _meta_table(doc, rows: Dict[str, str]):
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.style = "Table Grid"
    for i, (k, v) in enumerate(rows.items()):
        tbl.rows[i].cells[0].paragraphs[0].add_run(k).bold = True
        tbl.rows[i].cells[1].paragraphs[0].add_run(clean(v))


def _bullets(doc, items: List[dict], tgt_name: str, color: str, empty_msg: str):
    if not items:
        doc.add_paragraph(empty_msg)
        return
    for item in items:
        p  = doc.add_paragraph(style="List Number")
        r  = p.add_run(f"[{clean(item.get('speaker',''))}]  ")
        r.bold = True
        r.font.color.rgb = RGBColor.from_string(color)
        p.add_run(clean(item.get("text","")))
        p.add_run(f"  ({clean(item.get('time',''))})").italic = True
        if item.get("deadline") and item["deadline"] != "Not specified":
            d = doc.add_paragraph()
            d.paragraph_format.left_indent = Inches(.35)
            d.add_run("Deadline: ").bold = True
            d.add_run(clean(item["deadline"]))
        tr = item.get("tr","")
        if tr and tr != item.get("text",""):
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(.35)
            tp.add_run(f"🌐 {tgt_name}: ").font.size = Pt(9)
            tx = tp.add_run(clean(tr))
            tx.font.size = Pt(9)
            tx.italic = True


def export_docx(summary: str, summary_tr: str,
                actions: List[dict], decisions: List[dict],
                questions: List[dict], diarized: List[dict],
                src_lang: str, tgt_lang: str,
                speakers: int, words: int, segments: int,
                ai_sentiment=None, ai_effectiveness=None, ai_risks=None) -> str:

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(.85)
    sec.top_margin  = sec.bottom_margin = Inches(.75)

    src_n = lname(src_lang)
    tgt_n = lname(tgt_lang)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MINUTES OF MEETING")
    r.bold = True; r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0, 113, 227)   # Apple blue

    s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s.add_run(f"MoM Live AI  •  {src_n} → {tgt_n}").italic = True
    doc.add_paragraph()

    # Metadata
    _banner(doc, "Meeting Information", "374151")
    _meta_table(doc, {
        "Generated":       datetime.now().strftime("%B %d, %Y  ·  %I:%M %p"),
        "Source language": src_n,
        "Target language": tgt_n,
        "Speakers":        str(speakers),
        "Word count":      str(words),
        "Segments":        str(segments),
    })
    doc.add_paragraph()

    # AI Insights (Gemini only)
    if ai_sentiment or ai_effectiveness:
        _banner(doc, "AI Insights  (Gemini)", "5b21b6")
        if ai_sentiment:
            p = doc.add_paragraph()
            p.add_run("Tone: ").bold = True
            p.add_run(clean(ai_sentiment.get("tone","—")))
            p.add_run("  —  " + clean(ai_sentiment.get("explanation",""))).italic = True
        if ai_effectiveness:
            p = doc.add_paragraph()
            p.add_run("Effectiveness: ").bold = True
            p.add_run(f"{ai_effectiveness.get('score','—')}/100")
            p.add_run("  —  " + clean(ai_effectiveness.get("rationale",""))).italic = True
        if ai_risks:
            doc.add_paragraph().add_run("Risks & Blockers:").bold = True
            for risk in ai_risks:
                rp = doc.add_paragraph(style="List Bullet")
                rp.add_run(f"[{clean(risk.get('severity',''))}] ").bold = True
                rp.add_run(clean(risk.get("text","")))
        doc.add_paragraph()

    # 1. Summary
    _banner(doc, "1.  Executive Summary", "0071e3")
    doc.add_paragraph(clean(summary or "No summary available."))
    if summary_tr and summary_tr != summary:
        p = doc.add_paragraph(clean(summary_tr))
        if p.runs: p.runs[0].italic = True
    doc.add_paragraph()

    # 2. Decisions
    _banner(doc, "2.  Key Decisions", "1a5c2e")
    _bullets(doc, decisions, tgt_n, "1a5c2e", "No decisions detected.")
    doc.add_paragraph()

    # 3. Action Items
    _banner(doc, "3.  Action Items", "8B0000")
    _bullets(doc, actions, tgt_n, "8B0000", "No action items detected.")
    doc.add_paragraph()

    # 4. Open Questions
    _banner(doc, "4.  Open Questions", "5c3d00")
    _bullets(doc, questions, tgt_n, "5c3d00", "No open questions detected.")
    doc.add_paragraph()

    # 5. Transcript
    _banner(doc, "5.  Full Transcript", "374151")
    for seg in diarized:
        p = doc.add_paragraph()
        r = p.add_run(f"[{seg.get('start',0)}s]  {seg.get('speaker','')}: ")
        r.bold = True; r.font.size = Pt(9)
        p.add_run(clean(seg.get("text",""))).font.size = Pt(9)
        tr = seg.get("tr","")
        if tr:
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(.35)
            tp.add_run(f"🌐 {tgt_n}: ").font.size = Pt(8)
            tx = tp.add_run(clean(tr))
            tx.font.size = Pt(8); tx.italic = True

    # 6. Disclaimer
    _banner(doc, "6.  Notes", "6b7280")
    disc = doc.add_paragraph(
        "Speaker labels are heuristic (pause-based). "
        "Translation via Google Translate free tier. "
        "For enterprise accuracy, upgrade to managed speech/translation APIs."
    )
    if disc.runs: disc.runs[0].italic = True

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(TMP, f"MoM_Live_AI_{ts}.docx")
    doc.save(path)
    return path


# ══════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## MoM Live AI")
    st.caption(f"v{APP_VER}  ·  Multilingual Meeting Intelligence")
    st.markdown("---")

    # Gemini AI key
    st.markdown("### AI Engine")
    gemini_key = st.text_input(
        "Gemini API key",
        type="password",
        placeholder="AIza...",
        help="Free key from aistudio.google.com/apikey — unlocks smart analysis, "
             "sentiment, effectiveness score, and risk detection. "
             "Leave blank for rule-based mode."
    )
    use_gemini = bool(gemini_key.strip()) and GEMINI_OK
    if gemini_key.strip() and not GEMINI_OK:
        st.error("Install google-genai — add to requirements.txt")
    elif use_gemini:
        st.success("Gemini AI active")
    else:
        st.caption("No key — using built-in analysis")

    st.markdown("---")

    # Target language
    st.markdown("### Language")
    tgt_lang = st.selectbox(
        "Translate output to",
        options=list(LANGS.keys()),
        format_func=lambda c: f"{LANGS[c]} ({c})",
        index=list(LANGS.keys()).index("en"),
    )

    lang_hint = st.selectbox(
        "Audio language (hint)",
        options=["auto"] + list(LANGS.keys()),
        format_func=lambda c: "Auto-detect" if c == "auto" else f"{LANGS[c]} ({c})",
        index=0,
        help="Leave on Auto-detect for most files. Set manually if detection is wrong."
    )

    st.markdown("---")

    # Model mode
    st.markdown("### Transcription")
    mode = st.radio("Model mode", ["Fast", "Better"], index=0, horizontal=True,
                    help="Fast = tiny (quicker).  Better = base (more accurate, required for Telugu, Hindi, Tamil, Arabic, etc.)")
    model_size = "tiny" if mode == "Fast" else "base"

    if lang_hint in HARD_LANGS and mode == "Fast":
        st.warning(f"Switch to **Better** mode for {LANGS.get(lang_hint, lang_hint)}.")

    num_speakers = st.slider("Number of speakers", 1, 6, 3)
    st.caption("Speaker labels are heuristic in this version.")

    st.markdown("---")

    # Translation toggles
    st.markdown("### Translate")
    tr_summary   = st.checkbox("Summary",              value=True)
    tr_actions   = st.checkbox("Action items",         value=True)
    tr_decisions = st.checkbox("Decisions",            value=True)
    tr_questions = st.checkbox("Open questions",       value=True)
    tr_transcript= st.checkbox("First 10 transcript lines", value=False,
                               help="Slower — enable only if you need it.")

    st.markdown("---")
    st.caption(f"Supported: {len(LANGS)} languages  ·  {len(LANGS)**2:,} translation pairs")
    st.caption("Max file size: 25 MB")
    st.caption("Formats: mp3  wav  m4a  mp4")


# ══════════════════════════════════════════════════════════════
#  HERO
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div class="hero-eyebrow">MoM Live AI</div>
  <h1 class="hero-title">Minutes that <em>write themselves.</em></h1>
  <p class="hero-sub">Upload any meeting recording — get a bilingual report with decisions,
  action items, open questions, and AI insights in seconds.</p>
</div>
""", unsafe_allow_html=True)

# Stat cards
c1, c2, c3, c4 = st.columns(4)
for col, icon, val, lbl in [
    (c1, "🌐", "52×52", "Languages"),
    (c2, "🎙️", "Whisper", "Transcription"),
    (c3, "🤖", "Gemini", "AI Analysis"),
    (c4, "📄", "DOCX", "Export"),
]:
    col.markdown(
        f'<div class="stat-card"><span class="stat-icon">{icon}</span>'
        f'<div class="stat-val">{val}</div>'
        f'<div class="stat-lbl">{lbl}</div></div>',
        unsafe_allow_html=True
    )


# ══════════════════════════════════════════════════════════════
#  UPLOAD
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Upload your meeting</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your audio or video file here",
    type=["mp3", "wav", "m4a", "mp4"],
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown(f"""
    <div class="empty">
      <span class="empty-icon">🎙️</span>
      <div class="empty-title">Drop your audio here</div>
      <p class="empty-sub">Supports mp3 · wav · m4a · mp4 — any language</p>
      <div style="margin-top:1.2rem">
        <span class="pill">Auto language detection</span>
        <span class="pill">52 languages</span>
        <span class="pill">DOCX export</span>
        <span class="pill">Gemini AI</span>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Validate size
mb = uploaded.size / 1_048_576
if mb > MAX_MB:
    st.error(f"File is {mb:.1f} MB — maximum is {MAX_MB} MB. Please trim or compress it.")
    st.stop()

upload_path = save_uploaded_file(uploaded)
st.audio(upload_path)

fa, fb, fc = st.columns(3)
fa.markdown(f'<span class="pill">📄 {esc(uploaded.name)}</span>',          unsafe_allow_html=True)
fb.markdown(f'<span class="pill">💾 {mb:.2f} MB</span>',                   unsafe_allow_html=True)
fc.markdown(f'<span class="pill">🌐 → {lname(tgt_lang)} ({tgt_lang})</span>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  GENERATE
# ══════════════════════════════════════════════════════════════
if st.button("Generate Meeting Minutes"):

    note     = st.empty()
    progress = st.progress(0)
    status   = st.empty()
    t0       = time.time()
    note.info("Starting — first run downloads the Whisper model, which takes about a minute.")

    try:
        # ── 1. Convert audio ──────────────────────────────────
        status.markdown("**Preparing audio...**")
        wav, warn = convert_to_wav(upload_path)
        if warn:
            st.warning(warn)
        progress.progress(8)

        # ── 2. Transcribe ────────────────────────────────────
        status.markdown(f"**Transcribing with Whisper {model_size}...**")
        transcript, segs, det_lang, lang_prob, duration = transcribe_audio(
            wav, model_size, lang_hint
        )
        progress.progress(36)

        if not transcript or not segs:
            status.empty(); progress.empty(); note.empty()
            st.error(
                "**No speech detected.**\n\n"
                f"Language hint: `{lang_hint}` · Model: `{mode}` ({model_size})\n\n"
                "**Common fixes:**\n"
                "- Switch to **Better** mode for Telugu, Hindi, Tamil, Arabic, etc.\n"
                "- Make sure the file has an audio track (not just video)\n"
                "- Try converting to MP3 first at cloudconvert.com"
            )
            st.stop()

        det_name = lname(det_lang)
        tgt_name = lname(tgt_lang)
        src_code = det_lang if det_lang in LANGS else "auto"

        # ── 3. Diarize ───────────────────────────────────────
        status.markdown("**Assigning speaker labels...**")
        diarized = diarize_segments(segs, num_speakers)
        progress.progress(44)

        # ── 4. Meeting intelligence ──────────────────────────
        gemini_data     = None
        ai_sentiment    = None
        ai_effectiveness= None
        ai_risks        = None

        if use_gemini:
            status.markdown("**Analyzing with Gemini AI...**")
            client      = get_gemini_client(gemini_key)
            gemini_data = analyze_with_gemini(client, transcript)
            if gemini_data is None:
                st.warning("Gemini analysis unavailable — using built-in analysis instead.")
            progress.progress(68)

        if gemini_data:
            summary   = gemini_data.get("summary","") or generate_summary(transcript)
            actions   = [{"speaker": a.get("owner","Unspecified"), "text": a["text"],
                          "time":"—", "owner": a.get("owner","Unspecified"),
                          "deadline": a.get("deadline","Not specified")}
                         for a in gemini_data.get("action_items",[]) if a.get("text")]
            decisions = [{"speaker":"—","text":d["text"],"time":"—"}
                         for d in gemini_data.get("decisions",[]) if d.get("text")]
            questions = [{"speaker":"—","text":q["text"],"time":"—"}
                         for q in gemini_data.get("open_questions",[]) if q.get("text")]
            ai_sentiment     = gemini_data.get("sentiment")
            ai_effectiveness = gemini_data.get("effectiveness")
            ai_risks         = [r for r in gemini_data.get("risks",[]) if r.get("text")]
        else:
            status.markdown("**Extracting summary...**")
            summary   = generate_summary(transcript)
            progress.progress(52)
            status.markdown("**Extracting actions & decisions...**")
            actions   = extract_action_items(diarized)
            decisions = extract_decisions(diarized)
            questions = extract_open_questions(diarized)
            progress.progress(68)

        # ── 5. Translate ─────────────────────────────────────
        status.markdown(f"**Translating {det_name} → {tgt_name}...**")
        summary_tr  = translate_text_safe(summary, tgt_lang, src_code) if tr_summary   else summary
        if tr_actions:   actions   = translate_list(actions,   tgt_lang, src_code)
        if tr_decisions: decisions = translate_list(decisions, tgt_lang, src_code)
        if tr_questions: questions = translate_list(questions, tgt_lang, src_code)
        if tr_transcript:
            for i, seg in enumerate(diarized):
                seg["tr"] = translate_text_safe(seg["text"], tgt_lang, src_code) if i < 10 else ""

        # Check if translation changed anything
        if tr_summary and summary_tr == summary and det_lang != tgt_lang:
            st.warning("Translation output is the same as input — this language pair may be unavailable in the free tier.")

        progress.progress(86)

        # ── 6. DOCX ──────────────────────────────────────────
        status.markdown("**Generating DOCX report...**")
        docx_path = export_docx(
            summary, summary_tr, actions, decisions, questions, diarized,
            det_lang, tgt_lang, num_speakers,
            len(transcript.split()), len(diarized),
            ai_sentiment, ai_effectiveness, ai_risks,
        )
        progress.progress(100)
        status.empty(); note.empty()
        gc.collect()

        elapsed = time.time() - t0
        st.success(f"Done in {elapsed:.1f} seconds.")
        st.markdown(
            f'<span class="pill">🌐 {det_name}</span>'
            f'<span class="pill">{lang_prob:.0f}% confidence</span>'
            f'<span class="pill">{duration:.0f}s audio</span>'
            f'<span class="pill">{"Gemini AI" if gemini_data else "Built-in analysis"}</span>',
            unsafe_allow_html=True
        )

        # ══════════════════════════════════════════════════════
        #  RESULTS
        # ══════════════════════════════════════════════════════

        # ── Overview ─────────────────────────────────────────
        st.markdown('<div class="sec">Results</div>', unsafe_allow_html=True)
        r1,r2,r3,r4,r5 = st.columns(5)
        for col, val, lbl in [
            (r1, len(transcript.split()), "Words"),
            (r2, len(diarized),            "Segments"),
            (r3, len(actions),             "Actions"),
            (r4, len(decisions),           "Decisions"),
            (r5, len(questions),           "Questions"),
        ]:
            col.markdown(
                f'<div class="res-box"><div class="res-val">{val}</div>'
                f'<div class="res-lbl">{lbl}</div></div>',
                unsafe_allow_html=True
            )

        # ── AI Insights (Gemini only) ─────────────────────────
        if ai_sentiment or ai_effectiveness:
            st.markdown('<div class="sec">AI Insights <span style="font-size:1rem;font-weight:400;color:#6e6e73;">· Gemini</span></div>', unsafe_allow_html=True)

            tone  = (ai_sentiment or {}).get("tone","—")
            texpl = (ai_sentiment or {}).get("explanation","")
            temo  = {"Positive":"😊","Neutral":"😐","Tense":"😬","Mixed":"🤔"}.get(tone,"💬")
            score = (ai_effectiveness or {}).get("score","—")
            ratnl = (ai_effectiveness or {}).get("rationale","")

            gi1, gi2 = st.columns(2)
            gi1.markdown(
                f'<div class="ai-box"><span class="ai-icon">{temo}</span>'
                f'<div class="ai-val">{esc(tone)}</div>'
                f'<div class="ai-lbl">Meeting tone</div>'
                f'<div class="ai-desc">{esc(texpl)}</div></div>',
                unsafe_allow_html=True
            )
            gi2.markdown(
                f'<div class="ai-box"><span class="ai-icon">📊</span>'
                f'<div class="ai-val">{esc(str(score))}<span style="font-size:1rem;font-weight:500;">/100</span></div>'
                f'<div class="ai-lbl">Effectiveness score</div>'
                f'<div class="ai-desc">{esc(ratnl)}</div></div>',
                unsafe_allow_html=True
            )

            if ai_risks:
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("**Risks & blockers**")
                sev_cls = {"High":"chip-high","Medium":"chip-med","Low":"chip-low"}
                for risk in ai_risks:
                    sev = risk.get("severity","Medium")
                    st.markdown(
                        f'<div class="card-sm"><span class="chip {sev_cls.get(sev,"chip-med")}">'
                        f'{esc(sev)}</span>{esc(risk.get("text",""))}</div>',
                        unsafe_allow_html=True
                    )

        # ── Summary ───────────────────────────────────────────
        st.markdown('<div class="sec">Summary</div>', unsafe_allow_html=True)
        tab1, tab2 = st.tabs([f"Original ({det_name})", f"Translated ({tgt_name})"])
        with tab1:
            st.markdown(f'<div class="card">{esc(summary)}</div>', unsafe_allow_html=True)
        with tab2:
            st.markdown(f'<div class="card">{esc(summary_tr or summary)}</div>', unsafe_allow_html=True)

        # ── Decisions ────────────────────────────────────────
        st.markdown('<div class="sec">Key decisions</div>', unsafe_allow_html=True)
        if decisions:
            for i, d in enumerate(decisions, 1):
                with st.expander(f"Decision {i}  —  {d.get('speaker','')}  {d.get('time','')}"):
                    st.markdown(f"**Original:** {d.get('text','')}")
                    tr = d.get("tr","")
                    if tr and tr != d.get("text"):
                        st.caption(f"🌐 {tgt_name}: {tr}")
        else:
            st.info("No explicit decisions detected.")

        # ── Actions ──────────────────────────────────────────
        st.markdown('<div class="sec">Action items</div>', unsafe_allow_html=True)
        if actions:
            for i, a in enumerate(actions, 1):
                with st.expander(f"Action {i}  —  {a.get('owner','')}  {a.get('time','')}"):
                    st.markdown(f"**Task:** {a.get('text','')}")
                    st.markdown(f"**Owner:** {a.get('owner','Unknown')}")
                    st.markdown(f"**Deadline:** {a.get('deadline','Not specified')}")
                    tr = a.get("tr","")
                    if tr and tr != a.get("text"):
                        st.caption(f"🌐 {tgt_name}: {tr}")
        else:
            st.info("No action items detected.")

        # ── Open questions ────────────────────────────────────
        st.markdown('<div class="sec">Open questions</div>', unsafe_allow_html=True)
        if questions:
            for i, q in enumerate(questions, 1):
                with st.expander(f"Question {i}  —  {q.get('speaker','')}  {q.get('time','')}"):
                    st.markdown(f"**Original:** {q.get('text','')}")
                    tr = q.get("tr","")
                    if tr and tr != q.get("text"):
                        st.caption(f"🌐 {tgt_name}: {tr}")
        else:
            st.info("No open questions detected.")

        # ── Transcript ────────────────────────────────────────
        st.markdown('<div class="sec">Transcript</div>', unsafe_allow_html=True)
        with st.expander(f"View full transcript  ({len(diarized)} segments)"):
            for seg in diarized:
                st.markdown(
                    f"**[{seg.get('start',0)}s]  {seg.get('speaker','')}:** {seg.get('text','')}"
                )
                tr = seg.get("tr","")
                if tr:
                    st.caption(f"🌐 {tr}")

        # ── Download ──────────────────────────────────────────
        st.markdown('<div class="sec">Download</div>', unsafe_allow_html=True)
        with open(docx_path, "rb") as f:
            st.download_button(
                label=f"Download Meeting Minutes  ({det_name} → {tgt_name})",
                data=f,
                file_name=os.path.basename(docx_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as exc:
        status.empty(); progress.empty(); note.empty()
        st.error(f"Processing failed: {str(exc)[:400]}")
        st.info("Try a shorter MP3/WAV clip. For Telugu/Hindi/Tamil, switch to **Better** mode in the sidebar.")
    finally:
        gc.collect()
