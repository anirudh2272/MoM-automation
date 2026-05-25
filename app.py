"""
MoM Live AI — Stable Streamlit Portfolio App

Stable version designed for Streamlit Cloud:
- Uses faster-whisper base for transcription
- Uses lightweight extractive summarization to avoid Streamlit memory crashes
- Extracts action items and decisions with robust rule-based NLP
- Translates summary/actions/decisions only for speed
- Exports professional DOCX meeting minutes

This version intentionally avoids loading BART-large, BART-MNLI, and mBART-50
inside Streamlit Cloud because those models often exceed free-tier memory.
"""

import os
import re
import html
import time
import tempfile
import subprocess
from datetime import datetime
from pathlib import Path

import streamlit as st
from faster_whisper import WhisperModel
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─────────────────────────────────────────────
# App constants
# ─────────────────────────────────────────────
MAX_FILE_MB = 50
WHISPER_MODEL = "base"   # stable on Streamlit Cloud
SUPPORTED_EXTS = ["mp3", "wav", "m4a", "mp4"]

LANG_NAMES = {
    "en": "English", "hi": "Hindi", "te": "Telugu", "ta": "Tamil",
    "fr": "French", "de": "German", "es": "Spanish", "it": "Italian",
    "pt": "Portuguese", "nl": "Dutch", "ar": "Arabic", "zh-CN": "Chinese",
    "ja": "Japanese", "ko": "Korean", "ru": "Russian", "ur": "Urdu",
    "bn": "Bengali", "gu": "Gujarati", "mr": "Marathi", "ml": "Malayalam",
    "kn": "Kannada", "pa": "Punjabi", "ne": "Nepali", "si": "Sinhala",
    "tr": "Turkish", "pl": "Polish", "sv": "Swedish", "fi": "Finnish",
    "cs": "Czech", "ro": "Romanian", "uk": "Ukrainian", "id": "Indonesian",
    "vi": "Vietnamese", "th": "Thai", "fa": "Persian", "he": "Hebrew",
    "sw": "Swahili", "af": "Afrikaans"
}


# ─────────────────────────────────────────────
# Page config and CSS
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="MoM Live AI",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
*{box-sizing:border-box}
html,body,[class*="css"]{font-family:'Inter',-apple-system,sans-serif;-webkit-font-smoothing:antialiased}
.stApp{background:#07070b;min-height:100vh;position:relative;overflow-x:hidden}
.stApp::before{content:'';position:fixed;top:-20%;left:-10%;width:60%;height:60%;background:radial-gradient(circle,rgba(168,85,247,.18) 0%,transparent 60%);pointer-events:none;z-index:0;animation:float1 25s ease-in-out infinite}
.stApp::after{content:'';position:fixed;bottom:-20%;right:-10%;width:60%;height:60%;background:radial-gradient(circle,rgba(236,72,153,.15) 0%,transparent 60%);pointer-events:none;z-index:0;animation:float2 30s ease-in-out infinite}
@keyframes float1{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(40px,30px) scale(1.1)}}
@keyframes float2{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(-40px,-30px) scale(1.15)}}
.main .block-container{position:relative;z-index:1;padding-top:2rem}
.hero-wrap{text-align:center;margin:1rem 0 3rem}
.hero-eyebrow{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.7rem;letter-spacing:4px;text-transform:uppercase;color:#a78bfa;padding:.4rem 1rem;background:rgba(168,85,247,.08);border:1px solid rgba(168,85,247,.25);border-radius:100px;margin-bottom:1.2rem}
.hero-title{font-family:'Instrument Serif',serif;font-size:4.2rem;font-weight:400;line-height:1;letter-spacing:-2px;margin:0;color:#f5f5f7}
.hero-title em{font-style:italic;background:linear-gradient(135deg,#c084fc 0%,#ec4899 50%,#f59e0b 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hero-subtitle{font-size:1rem;font-weight:300;color:#a1a1aa;margin-top:1rem;max-width:680px;margin-left:auto;margin-right:auto;line-height:1.6}
.stat-card{background:linear-gradient(135deg,rgba(255,255,255,.04) 0%,rgba(255,255,255,.01) 100%);border:1px solid rgba(255,255,255,.08);border-radius:20px;padding:1.4rem 1rem;text-align:center;transition:all .3s cubic-bezier(.4,0,.2,1)}
.stat-card:hover{transform:translateY(-4px);border-color:rgba(168,85,247,.3)}
.stat-icon{font-size:1.5rem;margin-bottom:.4rem;display:block}
.stat-value{font-family:'Instrument Serif',serif;font-size:2.2rem;font-weight:400;color:#f5f5f7;line-height:1;margin-bottom:.3rem}
.stat-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#a1a1aa;text-transform:uppercase;letter-spacing:2px}
.sec-header{font-family:'Instrument Serif',serif;font-size:1.8rem;font-style:italic;color:#f5f5f7;margin:2.5rem 0 1.2rem;display:flex;align-items:center;gap:.8rem}
.sec-header::after{content:'';flex:1;height:1px;background:linear-gradient(90deg,rgba(168,85,247,.4),transparent)}
.sec-number{font-family:'JetBrains Mono',monospace;font-size:.7rem;color:#a78bfa;background:rgba(168,85,247,.1);padding:.3rem .6rem;border-radius:6px;border:1px solid rgba(168,85,247,.2)}
.glass-card{background:linear-gradient(135deg,rgba(255,255,255,.035) 0%,rgba(255,255,255,.012) 100%);border:1px solid rgba(255,255,255,.08);border-radius:16px;padding:1.5rem;margin-bottom:1rem;color:#d4d4d8;line-height:1.7;font-size:.95rem;white-space:pre-wrap}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0a0f 0%,#07070b 100%) !important;border-right:1px solid rgba(255,255,255,.06) !important}
section[data-testid="stSidebar"] *{color:#a1a1aa !important}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{color:#f5f5f7 !important;font-family:'Instrument Serif',serif !important;font-style:italic !important;font-weight:400 !important}
.stButton>button{background:linear-gradient(135deg,#a855f7 0%,#ec4899 50%,#f59e0b 100%) !important;color:white !important;border:none !important;border-radius:100px !important;font-family:'Inter',sans-serif !important;font-weight:700 !important;font-size:.98rem !important;padding:.9rem 2rem !important;width:100% !important;transition:all .3s !important;box-shadow:0 8px 32px rgba(168,85,247,.3) !important}
.stButton>button:hover{transform:translateY(-2px) !important;box-shadow:0 12px 40px rgba(168,85,247,.45) !important}
.stDownloadButton>button{background:linear-gradient(135deg,#10b981 0%,#06b6d4 100%) !important;color:white !important;border:none !important;border-radius:100px !important;font-weight:700 !important;width:100% !important;padding:.9rem 2rem !important;box-shadow:0 8px 32px rgba(16,185,129,.25) !important;transition:all .3s !important}
.stSelectbox>div>div{background:rgba(255,255,255,.03) !important;border:1px solid rgba(255,255,255,.08) !important;border-radius:12px !important;color:#e4e4e7 !important}
.stFileUploader>div{background:linear-gradient(135deg,rgba(168,85,247,.04) 0%,rgba(236,72,153,.03) 100%) !important;border:2px dashed rgba(168,85,247,.25) !important;border-radius:20px !important}
.stFileUploader>div:hover{border-color:rgba(168,85,247,.5) !important}
.stTabs [data-baseweb="tab-list"]{background:rgba(255,255,255,.03);border-radius:100px;padding:4px;border:1px solid rgba(255,255,255,.06);gap:4px}
.stTabs [data-baseweb="tab"]{color:#a1a1aa !important;border-radius:100px !important;padding:.5rem 1.5rem !important;font-weight:500 !important}
.stTabs [aria-selected="true"]{background:linear-gradient(135deg,#a855f7 0%,#ec4899 100%) !important;color:white !important}
.streamlit-expanderHeader{background:rgba(255,255,255,.03) !important;border:1px solid rgba(255,255,255,.06) !important;border-radius:12px !important;color:#d4d4d8 !important;font-weight:500 !important}
.stAlert{background:rgba(255,255,255,.035) !important;border:1px solid rgba(255,255,255,.08) !important;border-radius:14px !important;color:#d4d4d8 !important}
.stProgress>div>div{background:linear-gradient(90deg,#a855f7,#ec4899,#f59e0b) !important;border-radius:100px !important}
.stProgress>div{background:rgba(255,255,255,.05) !important;border-radius:100px !important}
audio{width:100%;border-radius:100px;filter:invert(.92) hue-rotate(180deg)}
.pill{display:inline-flex;align-items:center;gap:.4rem;padding:.4rem .9rem;border-radius:100px;font-size:.75rem;font-weight:600;margin-right:.4rem;margin-bottom:.4rem}
.pill-purple{background:rgba(168,85,247,.12);color:#c084fc;border:1px solid rgba(168,85,247,.3)}
.pill-pink{background:rgba(236,72,153,.12);color:#f472b6;border:1px solid rgba(236,72,153,.3)}
.pill-green{background:rgba(16,185,129,.12);color:#34d399;border:1px solid rgba(16,185,129,.3)}
.pill-amber{background:rgba(245,158,11,.12);color:#fbbf24;border:1px solid rgba(245,158,11,.3)}
.result-box{background:linear-gradient(135deg,rgba(168,85,247,.08) 0%,rgba(236,72,153,.04) 100%);border:1px solid rgba(168,85,247,.2);border-radius:18px;padding:1.4rem 1rem;text-align:center;transition:all .3s}
.result-value{font-family:'Instrument Serif',serif;font-size:2.5rem;font-weight:400;background:linear-gradient(135deg,#c084fc,#ec4899);-webkit-background-clip:text;-webkit-text-fill-color:transparent;line-height:1}
.result-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#a1a1aa;text-transform:uppercase;letter-spacing:2px;margin-top:.5rem}
.empty-state{text-align:center;padding:5rem 2rem;background:linear-gradient(135deg,rgba(168,85,247,.03) 0%,rgba(236,72,153,.02) 100%);border:1px dashed rgba(168,85,247,.2);border-radius:24px;margin-top:1rem}
.empty-icon{font-size:4.5rem;margin-bottom:1.2rem;display:inline-block;animation:pulse 3s ease-in-out infinite}
@keyframes pulse{0%,100%{transform:scale(1);opacity:1}50%{transform:scale(1.05);opacity:.85}}
.empty-title{font-family:'Instrument Serif',serif;font-style:italic;font-size:2rem;color:#f5f5f7;margin-bottom:.6rem}
.empty-text{color:#a1a1aa;margin-top:.5rem;font-size:.95rem;margin-bottom:1.5rem}
p,li{color:#a1a1aa !important;line-height:1.7}
span,label{color:#d4d4d8 !important}
h1,h2,h3,h4{color:#f5f5f7 !important}
strong,b{color:#f5f5f7 !important}
hr{border:none !important;height:1px !important;background:linear-gradient(90deg,transparent,rgba(168,85,247,.3),transparent) !important;margin:2rem 0 !important}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:#07070b}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#a855f7,#ec4899);border-radius:100px}
#MainMenu{visibility:hidden}footer{visibility:hidden}header{background:transparent !important}
</style>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# Cached Whisper model
# ─────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def load_whisper():
    return WhisperModel(WHISPER_MODEL, device="cpu", compute_type="int8")


# ─────────────────────────────────────────────
# Utility functions
# ─────────────────────────────────────────────
def safe_html(text: str) -> str:
    return html.escape(text or "")


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", name)
    return name[:120] if name else "uploaded_audio"


def convert_to_wav(input_path: str) -> str:
    """Convert any uploaded audio/video to clean 16kHz mono WAV."""
    output_path = str(Path(tempfile.gettempdir()) / (Path(input_path).stem + "_clean.wav"))
    cmd = [
        "ffmpeg", "-y", "-i", input_path,
        "-vn", "-acodec", "pcm_s16le",
        "-ar", "16000", "-ac", "1",
        "-loglevel", "error", output_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if result.returncode != 0:
        raise RuntimeError(result.stderr[:500] or "ffmpeg conversion failed")
    if not os.path.exists(output_path) or os.path.getsize(output_path) < 1000:
        raise RuntimeError("Converted audio is empty. The file may have no audio track.")
    return output_path


def transcribe(audio_path: str):
    """Transcribe audio with faster-whisper."""
    wav_path = convert_to_wav(audio_path)
    model = load_whisper()
    raw_segments, info = model.transcribe(
        wav_path,
        beam_size=3,
        language=None,
        condition_on_previous_text=False,
        vad_filter=False,
    )

    segments = []
    for seg in raw_segments:
        text = seg.text.strip()
        if text:
            segments.append({
                "start": round(seg.start, 2),
                "end": round(seg.end, 2),
                "text": text,
            })

    transcript = " ".join(s["text"] for s in segments)
    lang = getattr(info, "language", "en") or "en"
    prob = round(float(getattr(info, "language_probability", 0.0)) * 100, 1)
    duration = round(float(getattr(info, "duration", 0.0)), 1)

    return transcript, segments, lang, prob, duration


def split_sentences(text: str):
    text = re.sub(r"\s+", " ", text or "").strip()
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+", text)
    return [p.strip() for p in parts if len(p.strip()) > 2]


def summarize_text(text: str, max_sentences: int = 5) -> str:
    """Fast extractive summary that does not need large models."""
    sentences = split_sentences(text)
    if not sentences:
        return ""
    if len(sentences) <= max_sentences:
        return " ".join(sentences)

    keywords = [
        "decided", "agreed", "action", "deadline", "launch", "review",
        "client", "schedule", "complete", "next", "important", "issue",
        "risk", "plan", "update", "follow", "deliver", "task"
    ]
    scored = []
    for i, s in enumerate(sentences):
        score = 0
        lower = s.lower()
        score += sum(2 for k in keywords if k in lower)
        score += 1 if i < 2 else 0
        score += 1 if len(s.split()) >= 8 else 0
        scored.append((score, i, s))

    chosen = sorted(sorted(scored, reverse=True)[:max_sentences], key=lambda x: x[1])
    return " ".join(s for _, _, s in chosen)


def simple_diarize(segments, num_speakers: int = 3):
    out, spk, last_end = [], 1, 0.0
    for seg in segments:
        if seg["start"] - last_end >= 1.5:
            spk = (spk % num_speakers) + 1
        out.append({**seg, "speaker": f"Speaker {spk}"})
        last_end = seg["end"]
    return out


def extract_actions(diarized):
    keywords = [
        "will", "should", "need to", "must", "please", "send", "schedule",
        "review", "ensure", "prepare", "follow up", "make sure", "complete",
        "assign", "share", "submit", "finish", "email", "call"
    ]
    actions = []
    seen = set()
    for seg in diarized:
        txt = seg["text"].strip()
        low = txt.lower()
        if any(k in low for k in keywords) and len(txt.split()) >= 4:
            key = txt.lower()
            if key not in seen:
                seen.add(key)
                actions.append({
                    "speaker": seg["speaker"],
                    "time": f"{seg['start']}s",
                    "text": txt,
                })
    return actions[:15]


def extract_decisions(diarized):
    keywords = [
        "decided", "agreed", "approved", "confirmed", "finalized",
        "moving forward", "we will", "going with", "accepted", "resolved",
        "decision", "launch date", "deadline is"
    ]
    decisions = []
    seen = set()
    for seg in diarized:
        txt = seg["text"].strip()
        low = txt.lower()
        if any(k in low for k in keywords) and len(txt.split()) >= 4:
            key = txt.lower()
            if key not in seen:
                seen.add(key)
                decisions.append({
                    "speaker": seg["speaker"],
                    "time": f"{seg['start']}s",
                    "text": txt,
                })
    return decisions[:15]


def translate_text(text: str, target_lang: str, source_lang: str = "auto") -> str:
    """Lightweight translation. Falls back safely to original text."""
    if not text or target_lang == "en" and source_lang in ("en", "auto"):
        return text or ""
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source=source_lang, target=target_lang).translate(text)
    except Exception:
        return text or ""


def translate_items(items, target_lang: str, source_lang: str):
    if source_lang == target_lang:
        for item in items:
            item["text_tr"] = item["text"]
        return items
    for item in items:
        item["text_tr"] = translate_text(item["text"], target_lang, source_lang)
    return items


# ─────────────────────────────────────────────
# DOCX export
# ─────────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_banner(doc, title, color="1F497D"):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, color)
    p = cell.paragraphs[0]
    r = p.add_run("  " + title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(255, 255, 255)
    doc.add_paragraph()


def export_docx(summary, summary_tr, actions, decisions, diarized, src_lang, tgt_lang):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(1)
    section.top_margin = section.bottom_margin = Inches(0.75)

    src_name = LANG_NAMES.get(src_lang, src_lang)
    tgt_name = LANG_NAMES.get(tgt_lang, tgt_lang)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("MINUTES OF MEETING")
    r.bold = True
    r.font.size = Pt(17)
    r.font.color.rgb = RGBColor(31, 73, 125)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run(f"Original: {src_name} | Translated: {tgt_name}").italic = True

    doc.add_paragraph(datetime.now().strftime("%B %d, %Y | %I:%M %p"))
    doc.add_paragraph()

    add_banner(doc, "1. Executive Summary", "1F497D")
    doc.add_paragraph(summary or "No summary available.")
    if tgt_lang != src_lang:
        p = doc.add_paragraph(summary_tr or "Translation unavailable.")
        if p.runs:
            p.runs[0].italic = True
    doc.add_paragraph()

    add_banner(doc, "2. Key Decisions", "375623")
    if decisions:
        for d in decisions:
            p = doc.add_paragraph(style="List Number")
            speaker = p.add_run(f"[{d.get('speaker', 'Speaker')}] ")
            speaker.bold = True
            speaker.font.color.rgb = RGBColor(55, 86, 35)
            p.add_run(d.get("text", ""))
            p.add_run(f" ({d.get('time', '')})").italic = True
            if tgt_lang != src_lang:
                tp = doc.add_paragraph()
                tp.paragraph_format.left_indent = Inches(0.4)
                run = tp.add_run(f"🌐 {tgt_name}: ")
                run.bold = True
                run.font.size = Pt(9)
                tr = tp.add_run(d.get("text_tr", ""))
                tr.italic = True
                tr.font.size = Pt(9)
    else:
        doc.add_paragraph("No decisions extracted.")
    doc.add_paragraph()

    add_banner(doc, "3. Action Items", "C00000")
    if actions:
        for a in actions:
            p = doc.add_paragraph(style="List Number")
            speaker = p.add_run(f"[{a.get('speaker', 'Speaker')}] ")
            speaker.bold = True
            speaker.font.color.rgb = RGBColor(192, 0, 0)
            p.add_run(a.get("text", ""))
            p.add_run(f" ({a.get('time', '')})").italic = True
            if tgt_lang != src_lang:
                tp = doc.add_paragraph()
                tp.paragraph_format.left_indent = Inches(0.4)
                run = tp.add_run(f"🌐 {tgt_name}: ")
                run.bold = True
                run.font.size = Pt(9)
                tr = tp.add_run(a.get("text_tr", ""))
                tr.italic = True
                tr.font.size = Pt(9)
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    add_banner(doc, "4. Full Transcript", "595959")
    for seg in diarized:
        p = doc.add_paragraph()
        speaker = p.add_run(f"[{seg.get('start', 0)}s] {seg.get('speaker', 'Speaker')}: ")
        speaker.bold = True
        speaker.font.size = Pt(9)
        p.add_run(seg.get("text", "")).font.size = Pt(9)

    output_path = str(Path(tempfile.gettempdir()) / f"MoM_Report_{int(time.time())}.docx")
    doc.save(output_path)
    return output_path


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.markdown(
    """
<div class="hero-wrap">
    <span class="hero-eyebrow">⟡ AI MEETING INTELLIGENCE</span>
    <h1 class="hero-title">Minutes that <em>write themselves</em></h1>
    <p class="hero-subtitle">
        Upload meeting audio or video, extract decisions and action items,
        translate key sections, and export professional meeting minutes.
    </p>
</div>
""",
    unsafe_allow_html=True,
)

cols = st.columns(4)
stats = [
    ("🌐", str(len(LANG_NAMES)), "Languages"),
    ("🎙️", WHISPER_MODEL, "Whisper"),
    ("⚡", "Stable", "Cloud Mode"),
    ("📄", "DOCX", "Export"),
]
for col, (icon, value, label) in zip(cols, stats):
    col.markdown(
        f'<div class="stat-card"><span class="stat-icon">{icon}</span>'
        f'<div class="stat-value">{value}</div>'
        f'<div class="stat-label">{label}</div></div>',
        unsafe_allow_html=True,
    )

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    target_lang = st.selectbox(
        "Translate key sections to:",
        options=list(LANG_NAMES.keys()),
        format_func=lambda code: f"{LANG_NAMES[code]} ({code})",
        index=list(LANG_NAMES.keys()).index("en"),
    )
    num_speakers = st.slider("Estimated number of speakers", 1, 6, 3)
    translate_summary_only = st.checkbox(
        "Fast mode: translate only summary/actions/decisions",
        value=True,
        help="Keeps Streamlit Cloud stable. Full transcript remains in original language.",
    )
    st.markdown("---")
    st.markdown("### 📁 Supported")
    st.markdown("`.mp3` `.wav` `.m4a` `.mp4`")
    st.caption(f"Max file size: {MAX_FILE_MB} MB")
    st.markdown("---")
    st.caption("This stable portfolio version avoids huge local BART/mBART models so it runs reliably on Streamlit Cloud.")

st.markdown(
    '<div class="sec-header"><span class="sec-number">01</span>Upload your meeting</div>',
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Upload meeting audio/video",
    type=SUPPORTED_EXTS,
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown(
        """
<div class="empty-state">
    <div class="empty-icon">🎙️</div>
    <div class="empty-title">Drop your meeting file</div>
    <div class="empty-text">Supports MP3, WAV, M4A, and MP4 files.</div>
    <div>
        <span class="pill pill-purple">⟡ Transcription</span>
        <span class="pill pill-pink">⟡ Action Items</span>
        <span class="pill pill-green">⟡ DOCX Export</span>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )
else:
    file_bytes = uploaded.getvalue()
    file_size_mb = len(file_bytes) / (1024 * 1024)
    if file_size_mb > MAX_FILE_MB:
        st.error(f"File is {file_size_mb:.1f} MB. Please upload a file under {MAX_FILE_MB} MB.")
        st.stop()

    safe_name = sanitize_filename(uploaded.name)
    audio_path = str(Path(tempfile.gettempdir()) / f"upload_{int(time.time())}_{safe_name}")
    with open(audio_path, "wb") as f:
        f.write(file_bytes)

    st.audio(audio_path)
    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<span class="pill pill-purple">📄 {safe_html(uploaded.name)}</span>', unsafe_allow_html=True)
    c2.markdown(f'<span class="pill pill-pink">💾 {file_size_mb:.2f} MB</span>', unsafe_allow_html=True)
    c3.markdown(f'<span class="pill pill-green">🌐 → {LANG_NAMES[target_lang]}</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("✨ Generate Meeting Minutes"):
        st.info("✅ Button clicked. Processing started.")
        progress = st.progress(0)
        status = st.empty()
        started_at = time.time()

        try:
            status.markdown("🎙️ **Converting and transcribing audio...**")
            transcript, segments, det_lang, det_prob, duration = transcribe(audio_path)
            if not transcript or not segments:
                st.error("No speech was extracted. Try a clearer MP3/WAV file.")
                st.stop()
            progress.progress(25)

            src_name = LANG_NAMES.get(det_lang, det_lang)
            tgt_name = LANG_NAMES[target_lang]
            st.markdown(
                f'<span class="pill pill-amber">Detected: {src_name} · {det_prob}% · {duration}s</span>',
                unsafe_allow_html=True,
            )

            status.markdown("👥 **Assigning speaker labels...**")
            diarized = simple_diarize(segments, num_speakers)
            progress.progress(40)

            status.markdown("🧠 **Extracting summary, decisions, and actions...**")
            summary = summarize_text(transcript)
            actions = extract_actions(diarized)
            decisions = extract_decisions(diarized)
            progress.progress(60)

            status.markdown("🌐 **Translating key sections...**")
            summary_tr = translate_text(summary, target_lang, det_lang)
            actions = translate_items(actions, target_lang, det_lang)
            decisions = translate_items(decisions, target_lang, det_lang)
            progress.progress(80)

            status.markdown("📄 **Generating DOCX report...**")
            docx_path = export_docx(
                summary=summary,
                summary_tr=summary_tr,
                actions=actions,
                decisions=decisions,
                diarized=diarized,
                src_lang=det_lang,
                tgt_lang=target_lang,
            )
            progress.progress(100)
            status.empty()

            elapsed = time.time() - started_at
            st.success(f"✅ Processing complete in {elapsed:.1f} seconds.")

            st.markdown(
                '<div class="sec-header"><span class="sec-number">02</span>Results overview</div>',
                unsafe_allow_html=True,
            )
            r1, r2, r3, r4 = st.columns(4)
            result_stats = [
                (len(actions), "Action Items"),
                (len(decisions), "Key Decisions"),
                (len(diarized), "Segments"),
                (len(transcript.split()), "Words"),
            ]
            for col, (value, label) in zip([r1, r2, r3, r4], result_stats):
                col.markdown(
                    f'<div class="result-box"><div class="result-value">{value}</div>'
                    f'<div class="result-label">{label}</div></div>',
                    unsafe_allow_html=True,
                )

            st.markdown(
                '<div class="sec-header"><span class="sec-number">03</span>Executive summary</div>',
                unsafe_allow_html=True,
            )
            tab1, tab2 = st.tabs([f"Original ({src_name})", f"Translated ({tgt_name})"])
            with tab1:
                st.markdown(f'<div class="glass-card">{safe_html(summary or "No summary available.")}</div>', unsafe_allow_html=True)
            with tab2:
                st.markdown(f'<div class="glass-card">{safe_html(summary_tr or "Translation unavailable.")}</div>', unsafe_allow_html=True)

            st.markdown(
                '<div class="sec-header"><span class="sec-number">04</span>Key decisions</div>',
                unsafe_allow_html=True,
            )
            if decisions:
                for i, d in enumerate(decisions, 1):
                    with st.expander(f"Decision {i} — {d.get('speaker')} ({d.get('time')})"):
                        st.write(d.get("text"))
                        if target_lang != det_lang:
                            st.write(d.get("text_tr"))
            else:
                st.info("No decisions extracted.")

            st.markdown(
                '<div class="sec-header"><span class="sec-number">05</span>Action items</div>',
                unsafe_allow_html=True,
            )
            if actions:
                for i, a in enumerate(actions, 1):
                    with st.expander(f"Action {i} — {a.get('speaker')} ({a.get('time')})"):
                        st.write(a.get("text"))
                        if target_lang != det_lang:
                            st.write(a.get("text_tr"))
            else:
                st.info("No action items extracted.")

            st.markdown(
                '<div class="sec-header"><span class="sec-number">06</span>Transcript</div>',
                unsafe_allow_html=True,
            )
            with st.expander(f"View transcript ({len(diarized)} segments)"):
                for seg in diarized:
                    st.markdown(
                        f"**[{seg['start']}s] {seg['speaker']}:** {seg['text']}"
                    )

            st.markdown(
                '<div class="sec-header"><span class="sec-number">07</span>Download report</div>',
                unsafe_allow_html=True,
            )
            with open(docx_path, "rb") as f:
                st.download_button(
                    label="⤓ Download Meeting Minutes DOCX",
                    data=f,
                    file_name=f"MoM_{src_name}_to_{tgt_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

        except subprocess.TimeoutExpired:
            st.error("Audio conversion timed out. Try a shorter file.")
        except Exception as e:
            st.error(f"Something went wrong: {str(e)[:500]}")
            st.info("Try MP3/WAV first. MP4 files sometimes fail if they do not contain a clear audio track.")
