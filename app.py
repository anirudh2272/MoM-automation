"""
MoM Automation - Working Portfolio Version
Meeting transcription, summarization, action/decision extraction, translation, and DOCX export.

Built to run on Streamlit Cloud without paid APIs or large local BART/mBART models.
Core stack: Streamlit + faster-whisper + rule-based meeting intelligence + deep-translator + python-docx.
"""

import gc
import html
import os
import re
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from faster_whisper import WhisperModel

try:
    from deep_translator import GoogleTranslator
except Exception:
    GoogleTranslator = None

APP_NAME = "MoM Automation"
MAX_FILE_MB = 25
MAX_TRANSLATE_CHARS = 4200
SUPPORTED_EXTENSIONS = ["mp3", "wav", "m4a", "mp4"]

LANG_NAMES: Dict[str, str] = {
    "ar": "Arabic", "cs": "Czech", "de": "German", "en": "English",
    "es": "Spanish", "et": "Estonian", "fi": "Finnish", "fr": "French",
    "gu": "Gujarati", "hi": "Hindi", "it": "Italian", "ja": "Japanese",
    "kk": "Kazakh", "ko": "Korean", "lt": "Lithuanian", "lv": "Latvian",
    "my": "Burmese", "ne": "Nepali", "nl": "Dutch", "ro": "Romanian",
    "ru": "Russian", "si": "Sinhala", "tr": "Turkish", "vi": "Vietnamese",
    "zh": "Chinese", "af": "Afrikaans", "az": "Azerbaijani", "bn": "Bengali",
    "fa": "Persian", "he": "Hebrew", "hr": "Croatian", "id": "Indonesian",
    "ka": "Georgian", "km": "Khmer", "mk": "Macedonian", "ml": "Malayalam",
    "mn": "Mongolian", "mr": "Marathi", "pl": "Polish", "ps": "Pashto",
    "pt": "Portuguese", "sv": "Swedish", "sw": "Swahili", "ta": "Tamil",
    "te": "Telugu", "th": "Thai", "tl": "Filipino", "uk": "Ukrainian",
    "ur": "Urdu", "xh": "Xhosa", "gl": "Galician", "sl": "Slovenian"
}

TRANSLATION_CODE_MAP: Dict[str, str] = {"zh": "zh-CN", "he": "iw"}

st.set_page_config(page_title=APP_NAME, page_icon="🎙️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
*{box-sizing:border-box} html,body,[class*="css"]{font-family:'Inter',sans-serif;-webkit-font-smoothing:antialiased}
.stApp{background:#07070b;color:#f5f5f7;min-height:100vh;overflow-x:hidden}
.stApp::before{content:'';position:fixed;top:-18%;left:-10%;width:65%;height:65%;background:radial-gradient(circle,rgba(168,85,247,.18),transparent 60%);pointer-events:none;z-index:0}
.stApp::after{content:'';position:fixed;bottom:-18%;right:-10%;width:65%;height:65%;background:radial-gradient(circle,rgba(236,72,153,.14),transparent 60%);pointer-events:none;z-index:0}
.main .block-container{position:relative;z-index:1;padding-top:2rem;max-width:1180px}
.hero{text-align:center;padding:1rem 0 2.5rem}.eyebrow{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.72rem;letter-spacing:4px;text-transform:uppercase;color:#c084fc;padding:.45rem 1rem;background:rgba(168,85,247,.10);border:1px solid rgba(168,85,247,.28);border-radius:999px;margin-bottom:1rem}
.hero h1{font-family:'Instrument Serif',serif;font-size:4.3rem;font-weight:400;letter-spacing:-2px;line-height:1;margin:0;color:#f5f5f7}.hero h1 em{font-style:italic;background:linear-gradient(135deg,#c084fc 0%,#ec4899 55%,#f59e0b 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}.hero p{color:#a1a1aa;font-size:1.05rem;line-height:1.65;max-width:720px;margin:1rem auto 0}
.stat-card{background:linear-gradient(135deg,rgba(255,255,255,.05),rgba(255,255,255,.015));border:1px solid rgba(255,255,255,.085);border-radius:20px;padding:1.35rem 1rem;text-align:center;transition:.25s ease}.stat-card:hover{transform:translateY(-3px);border-color:rgba(168,85,247,.35)}.stat-icon{font-size:1.55rem;display:block;margin-bottom:.35rem}.stat-value{font-family:'Instrument Serif',serif;font-size:2.25rem;color:#f5f5f7;line-height:1}.stat-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#8b8b96;text-transform:uppercase;letter-spacing:2px;margin-top:.4rem}
.sec-header{font-family:'Instrument Serif',serif;font-size:1.85rem;font-style:italic;color:#f5f5f7;margin:2.2rem 0 1rem;display:flex;align-items:center;gap:.75rem}.sec-header:after{content:'';height:1px;flex:1;background:linear-gradient(90deg,rgba(168,85,247,.45),transparent)}.sec-number{font-family:'JetBrains Mono',monospace;font-size:.72rem;color:#c084fc;background:rgba(168,85,247,.10);padding:.35rem .6rem;border:1px solid rgba(168,85,247,.25);border-radius:8px}
.glass-card{background:linear-gradient(135deg,rgba(255,255,255,.045),rgba(255,255,255,.015));border:1px solid rgba(255,255,255,.09);border-radius:18px;padding:1.25rem;color:#d4d4d8;line-height:1.72;font-size:.98rem;margin-bottom:1rem;white-space:pre-wrap}.result-box{background:linear-gradient(135deg,rgba(168,85,247,.10),rgba(236,72,153,.045));border:1px solid rgba(168,85,247,.22);border-radius:18px;padding:1.25rem;text-align:center}.result-value{font-family:'Instrument Serif',serif;font-size:2.45rem;background:linear-gradient(135deg,#c084fc,#ec4899);-webkit-background-clip:text;-webkit-text-fill-color:transparent;line-height:1}.result-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#a1a1aa;text-transform:uppercase;letter-spacing:2px;margin-top:.5rem}.pill{display:inline-flex;align-items:center;gap:.35rem;border-radius:999px;padding:.42rem .85rem;font-size:.76rem;font-weight:600;margin:.2rem .25rem .2rem 0;border:1px solid rgba(255,255,255,.1)}.purple{background:rgba(168,85,247,.13);color:#d8b4fe;border-color:rgba(168,85,247,.32)}.pink{background:rgba(236,72,153,.12);color:#f9a8d4;border-color:rgba(236,72,153,.30)}.green{background:rgba(16,185,129,.12);color:#6ee7b7;border-color:rgba(16,185,129,.30)}.amber{background:rgba(245,158,11,.12);color:#fbbf24;border-color:rgba(245,158,11,.30)}.empty{padding:4rem 2rem;text-align:center;border:1px dashed rgba(168,85,247,.25);border-radius:24px;background:linear-gradient(135deg,rgba(168,85,247,.04),rgba(236,72,153,.025))}.empty .icon{font-size:4rem;margin-bottom:1rem}.empty .title{font-family:'Instrument Serif',serif;font-style:italic;font-size:2rem;color:#f5f5f7}.empty .text{color:#8b8b96;margin-top:.5rem}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0a0f,#07070b)!important;border-right:1px solid rgba(255,255,255,.07)!important}section[data-testid="stSidebar"] *{color:#a1a1aa!important}.stButton>button{background:linear-gradient(135deg,#a855f7,#ec4899,#f59e0b)!important;color:white!important;border:none!important;border-radius:999px!important;font-weight:700!important;width:100%!important;padding:.85rem 1.6rem!important;box-shadow:0 8px 34px rgba(168,85,247,.28)!important}.stDownloadButton>button{background:linear-gradient(135deg,#10b981,#06b6d4)!important;color:white!important;border:none!important;border-radius:999px!important;font-weight:700!important;width:100%!important;padding:.85rem 1.6rem!important}.stFileUploader>div{background:linear-gradient(135deg,rgba(168,85,247,.05),rgba(236,72,153,.035))!important;border:2px dashed rgba(168,85,247,.26)!important;border-radius:22px!important}p,li,span,label{color:#d4d4d8!important} h1,h2,h3,h4,strong,b{color:#f5f5f7!important}#MainMenu, footer{visibility:hidden}
</style>
""", unsafe_allow_html=True)

def safe_filename(name: str) -> str:
    return (re.sub(r"[^A-Za-z0-9_.-]+", "_", name)[:120] or "uploaded_audio")

def clean_docx_text(text: str) -> str:
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(text or ""))

def escape_html(text: str) -> str:
    return html.escape(text or "")

@st.cache_resource(show_spinner=False)
def load_whisper_model(model_size: str):
    return WhisperModel(model_size, device="cpu", compute_type="int8")

def convert_to_wav(input_path: str) -> Tuple[str, str]:
    output_path = str(Path(input_path).with_suffix("")) + "_16k.wav"
    cmd = ["ffmpeg", "-y", "-i", input_path, "-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1", "-loglevel", "error", output_path]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            return input_path, f"FFmpeg conversion failed; using original file. {result.stderr[:180]}"
        if not os.path.exists(output_path) or os.path.getsize(output_path) < 1000:
            return input_path, "Converted WAV was empty; using original file."
        return output_path, ""
    except FileNotFoundError:
        return input_path, "FFmpeg is not installed. Add ffmpeg to packages.txt. Using original file."
    except subprocess.TimeoutExpired:
        return input_path, "FFmpeg timed out. Using original file."
    except Exception as exc:
        return input_path, f"FFmpeg error: {str(exc)[:180]}. Using original file."

def transcribe_audio(audio_path: str, model_size: str, language_hint: str = "auto") -> Tuple[str, List[Dict], str, float, float]:
    model = load_whisper_model(model_size)
    language = None if language_hint == "auto" else language_hint
    raw, info = model.transcribe(audio_path, beam_size=1, language=language, vad_filter=True, condition_on_previous_text=False, no_speech_threshold=0.35)
    segments = []
    for seg in raw:
        text = (seg.text or "").strip()
        if text:
            segments.append({"start": round(float(seg.start), 2), "end": round(float(seg.end), 2), "text": text})
    transcript = " ".join(s["text"] for s in segments).strip()
    lang = getattr(info, "language", None) or language_hint or "en"
    prob = float(getattr(info, "language_probability", 0.0) or 0.0) * 100
    duration = float(getattr(info, "duration", 0.0) or 0.0)
    return transcript, segments, lang, round(prob, 1), duration

def simple_diarize(segments: List[Dict], num_speakers: int) -> List[Dict]:
    if not segments:
        return []
    diarized, speaker, last_end = [], 1, segments[0].get("start", 0.0)
    for seg in segments:
        if float(seg.get("start", 0.0)) - float(last_end) >= 1.6 and num_speakers > 1:
            speaker = (speaker % num_speakers) + 1
        diarized.append({**seg, "speaker": f"Speaker {speaker}"})
        last_end = float(seg.get("end", last_end))
    return diarized

def split_sentences(text: str) -> List[str]:
    return [p.strip() for p in re.split(r"(?<=[.!?])\s+", (text or "").strip()) if len(p.strip()) > 3]

def summarize_text(text: str, max_sentences: int = 5) -> str:
    sentences = split_sentences(text)
    if not sentences:
        return "No transcript was generated."
    if len(sentences) <= max_sentences:
        return " ".join(sentences)
    keywords = ["decided", "agreed", "approved", "confirmed", "action", "next", "follow", "deadline", "task", "need", "will", "should", "review", "prepare", "send", "project", "issue", "risk", "plan", "deliverable", "timeline"]
    scored = []
    for idx, sent in enumerate(sentences):
        lower = sent.lower()
        score = sum(2 for kw in keywords if kw in lower) + min(len(sent.split()) / 18, 2)
        if idx == 0: score += 1.8
        if idx == len(sentences) - 1: score += 0.8
        scored.append((score, idx, sent))
    selected = sorted(sorted(scored, reverse=True)[:max_sentences], key=lambda x: x[1])
    return " ".join(x[2] for x in selected)

def detect_deadline(text: str) -> str:
    patterns = [r"\bby\s+(tomorrow|today|monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b", r"\bby\s+([A-Z][a-z]+\s+\d{1,2})\b", r"\b(due\s+(?:on|by)?\s*[^,.!?]{3,30})", r"\b(next\s+week|this\s+week|end\s+of\s+week|EOW)\b", r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match: return match.group(0).strip()
    return "Not specified"

def extract_actions(segments: List[Dict]) -> List[Dict]:
    keywords = ["will", "should", "need to", "needs to", "must", "please", "send", "schedule", "review", "prepare", "follow up", "complete", "finish", "update", "share", "create", "submit", "check", "confirm", "make sure", "assign", "work on"]
    out, seen = [], set()
    for seg in segments:
        text, lower = seg.get("text", "").strip(), seg.get("text", "").lower()
        if len(text.split()) < 4 or not any(k in lower for k in keywords): continue
        key = re.sub(r"\W+", "", lower)[:80]
        if key in seen: continue
        seen.add(key)
        out.append({"speaker": seg.get("speaker", "Unknown"), "text": text, "time": f"{seg.get('start', 0)}s", "owner": seg.get("speaker", "Unknown"), "deadline": detect_deadline(text)})
        if len(out) >= 12: break
    return out

def extract_decisions(segments: List[Dict]) -> List[Dict]:
    keywords = ["decided", "agreed", "approved", "confirmed", "finalized", "accepted", "we will", "moving forward", "going with", "resolved", "the decision", "conclusion", "settled", "chosen"]
    out, seen = [], set()
    for seg in segments:
        text, lower = seg.get("text", "").strip(), seg.get("text", "").lower()
        if len(text.split()) < 4 or not any(k in lower for k in keywords): continue
        key = re.sub(r"\W+", "", lower)[:80]
        if key in seen: continue
        seen.add(key)
        out.append({"speaker": seg.get("speaker", "Unknown"), "text": text, "time": f"{seg.get('start', 0)}s"})
        if len(out) >= 10: break
    return out

def extract_questions(segments: List[Dict]) -> List[Dict]:
    starters = ("what", "why", "how", "when", "where", "who", "can", "could", "should", "do", "does", "did", "is", "are")
    out, seen = [], set()
    for seg in segments:
        text = seg.get("text", "").strip()
        lower = text.lower()
        first = lower.split()[0] if lower.split() else ""
        if "?" not in text and first not in starters: continue
        if len(text.split()) < 4: continue
        key = re.sub(r"\W+", "", lower)[:80]
        if key in seen: continue
        seen.add(key)
        out.append({"speaker": seg.get("speaker", "Unknown"), "text": text, "time": f"{seg.get('start', 0)}s"})
        if len(out) >= 10: break
    return out

def translate_text(text: str, target_lang: str, source_lang: str = "auto") -> str:
    if not text or target_lang == source_lang:
        return text or ""
    if GoogleTranslator is None:
        return text
    target = TRANSLATION_CODE_MAP.get(target_lang, target_lang)
    source = "auto" if source_lang == "auto" else TRANSLATION_CODE_MAP.get(source_lang, source_lang)
    try:
        chunks, remaining = [], text
        while remaining:
            chunk = remaining[:MAX_TRANSLATE_CHARS]
            if len(remaining) > MAX_TRANSLATE_CHARS:
                cut = max(chunk.rfind(". "), chunk.rfind("? "), chunk.rfind("! "))
                if cut > 500: chunk = chunk[:cut + 1]
            chunks.append(chunk)
            remaining = remaining[len(chunk):].strip()
        return " ".join(GoogleTranslator(source=source, target=target).translate(c) for c in chunks if c)
    except Exception:
        return text

def translate_list_items(items: List[Dict], target_lang: str, source_lang: str) -> List[Dict]:
    for item in items:
        item["text_tr"] = translate_text(item.get("text", ""), target_lang, source_lang)
    return items

def set_cell_background(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:val"), "clear"); shade.set(qn("w:color"), "auto"); shade.set(qn("w:fill"), hex_color)
    tc_pr.append(shade)

def add_banner(doc: Document, title: str, color: str = "1F497D"):
    table = doc.add_table(rows=1, cols=1); cell = table.cell(0, 0); set_cell_background(cell, color)
    run = cell.paragraphs[0].add_run("  " + clean_docx_text(title)); run.bold = True; run.font.size = Pt(11); run.font.color.rgb = RGBColor(255, 255, 255)
    doc.add_paragraph()

def add_bullet_section(doc: Document, items: List[Dict], translated_label: str, color: str, empty_text: str):
    if not items:
        doc.add_paragraph(empty_text); return
    for item in items:
        p = doc.add_paragraph(style="List Number")
        r = p.add_run(f"[{clean_docx_text(item.get('speaker', 'Unknown'))}]  "); r.bold = True; r.font.color.rgb = RGBColor.from_string(color)
        p.add_run(clean_docx_text(item.get("text", ""))); p.add_run(f"  ({clean_docx_text(item.get('time', ''))})").italic = True
        if item.get("deadline"):
            d = doc.add_paragraph(); d.paragraph_format.left_indent = Inches(0.4); d.add_run("Deadline: ").bold = True; d.add_run(clean_docx_text(item.get("deadline", "Not specified")))
        if item.get("text_tr") and item.get("text_tr") != item.get("text"):
            tp = doc.add_paragraph(); tp.paragraph_format.left_indent = Inches(0.4)
            lab = tp.add_run(f"🌐 {translated_label}: "); lab.font.size = Pt(9)
            tr = tp.add_run(clean_docx_text(item.get("text_tr", ""))); tr.font.size = Pt(9); tr.italic = True

def export_docx(summary, summary_tr, actions, decisions, questions, diarized, src_lang, tgt_lang) -> str:
    doc = Document(); section = doc.sections[0]; section.left_margin = section.right_margin = Inches(0.85); section.top_margin = section.bottom_margin = Inches(0.75)
    src_name, tgt_name = LANG_NAMES.get(src_lang, src_lang), LANG_NAMES.get(tgt_lang, tgt_lang)
    title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = title.add_run("MINUTES OF MEETING"); tr.bold = True; tr.font.size = Pt(17); tr.font.color.rgb = RGBColor(31, 73, 125)
    sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER; sub.add_run(f"Original: {src_name}   |   Translated: {tgt_name}").italic = True
    gen = doc.add_paragraph(); gen.alignment = WD_ALIGN_PARAGRAPH.CENTER; gen.add_run(datetime.now().strftime("Generated on %B %d, %Y at %I:%M %p")); doc.add_paragraph()
    add_banner(doc, "1. Executive Summary", "1F497D"); doc.add_paragraph(clean_docx_text(summary or "No summary available."))
    if summary_tr and summary_tr != summary:
        p = doc.add_paragraph(clean_docx_text(summary_tr));
        if p.runs: p.runs[0].italic = True
    doc.add_paragraph(); add_banner(doc, "2. Key Decisions", "375623"); add_bullet_section(doc, decisions, tgt_name, "375623", "No key decisions were detected.")
    doc.add_paragraph(); add_banner(doc, "3. Action Items", "C00000"); add_bullet_section(doc, actions, tgt_name, "C00000", "No action items were detected.")
    doc.add_paragraph(); add_banner(doc, "4. Open Questions", "806000"); add_bullet_section(doc, questions, tgt_name, "806000", "No open questions were detected.")
    doc.add_paragraph(); add_banner(doc, "5. Full Transcript", "595959")
    for seg in diarized:
        p = doc.add_paragraph(); r = p.add_run(f"[{seg.get('start', 0)}s] {seg.get('speaker', 'Speaker')}: "); r.bold = True; r.font.size = Pt(9)
        p.add_run(clean_docx_text(seg.get("text", ""))).font.size = Pt(9)
    output_path = f"/tmp/MoM_Report_{safe_filename(src_name)}_to_{safe_filename(tgt_name)}.docx"; doc.save(output_path); return output_path

st.markdown("""<div class="hero"><span class="eyebrow">AI MEETING INTELLIGENCE</span><h1>Minutes that <em>write themselves</em></h1><p>Upload a meeting recording and generate a clean report with transcript, summary, decisions, action items, open questions, translation, and DOCX export.</p></div>""", unsafe_allow_html=True)
cols = st.columns(4)
for col, icon, val, lab in [(cols[0], "🌐", "52", "Language options"), (cols[1], "🎙️", "Whisper", "Transcription"), (cols[2], "📌", "Actions", "Task extraction"), (cols[3], "📄", "DOCX", "Report export")]:
    col.markdown(f'<div class="stat-card"><span class="stat-icon">{icon}</span><div class="stat-value">{val}</div><div class="stat-label">{lab}</div></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    target_lang = st.selectbox("Translate key sections to", options=list(LANG_NAMES.keys()), format_func=lambda c: f"{LANG_NAMES[c]} ({c})", index=list(LANG_NAMES.keys()).index("en"))
    language_hint = st.selectbox("Audio language hint", options=["auto"] + list(LANG_NAMES.keys()), format_func=lambda c: "Auto detect" if c == "auto" else f"{LANG_NAMES[c]} ({c})", index=0)
    model_size = st.selectbox("Transcription model", options=["tiny", "base"], index=1, help="Use tiny for faster demos; base for better accuracy.")
    num_speakers = st.slider("Estimated speakers", min_value=1, max_value=6, value=3)
    translate_transcript_preview = st.checkbox("Translate first 10 transcript lines", value=False, help="Keep off for faster public demos.")
    st.markdown("---"); st.markdown("### Public demo limits"); st.caption(f"Max file size: {MAX_FILE_MB} MB"); st.caption("Recommended: 30 seconds to 5 minutes."); st.markdown("`mp3` `wav` `m4a` `mp4`")

st.markdown('<div class="sec-header"><span class="sec-number">01</span>Upload your meeting</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("Upload meeting audio/video", type=SUPPORTED_EXTENSIONS, label_visibility="collapsed")

if uploaded is None:
    st.markdown("""<div class="empty"><div class="icon">🎙️</div><div class="title">Drop your audio file</div><div class="text">The app will transcribe, summarize, extract meeting intelligence, translate key sections, and generate a DOCX report.</div><div style="margin-top:1rem"><span class="pill purple">Auto language detection</span><span class="pill pink">52 language options</span><span class="pill green">DOCX export</span></div></div>""", unsafe_allow_html=True)
    st.stop()

file_size_mb = uploaded.size / (1024 * 1024)
if file_size_mb > MAX_FILE_MB:
    st.error(f"This file is {file_size_mb:.1f} MB. Please upload a file under {MAX_FILE_MB} MB for the public demo."); st.stop()

upload_path = f"/tmp/{int(time.time())}_{safe_filename(uploaded.name)}"
with open(upload_path, "wb") as f: f.write(uploaded.getbuffer())
st.audio(upload_path)
mc = st.columns(3); mc[0].markdown(f'<span class="pill purple">📄 {escape_html(uploaded.name)}</span>', unsafe_allow_html=True); mc[1].markdown(f'<span class="pill pink">💾 {file_size_mb:.2f} MB</span>', unsafe_allow_html=True); mc[2].markdown(f'<span class="pill green">🌐 → {LANG_NAMES[target_lang]}</span>', unsafe_allow_html=True)

if st.button("✨ Generate Meeting Minutes"):
    note = st.empty(); note.info("Button clicked. Processing started. First run may take longer while Whisper downloads.")
    progress = st.progress(0); status = st.empty(); t0 = time.time()
    try:
        status.markdown("🎧 **Preparing audio...**"); wav_path, warning = convert_to_wav(upload_path)
        if warning: st.warning(warning)
        progress.progress(12)
        status.markdown("🎙️ **Transcribing speech...**"); transcript, segments, detected_lang, lang_prob, duration = transcribe_audio(wav_path, model_size, language_hint); progress.progress(38)
        if not transcript or not segments:
            status.empty(); progress.empty(); st.error("No speech was detected. Try a clearer or shorter MP3/WAV sample."); st.stop()
        status.markdown("👥 **Assigning speaker labels...**"); diarized = simple_diarize(segments, num_speakers); progress.progress(48)
        status.markdown("🧠 **Extracting meeting intelligence...**"); summary = summarize_text(transcript); actions = extract_actions(diarized); decisions = extract_decisions(diarized); questions = extract_questions(diarized); progress.progress(66)
        src = detected_lang if detected_lang in LANG_NAMES else "auto"; status.markdown(f"🌐 **Translating key sections to {LANG_NAMES[target_lang]}...**")
        summary_tr = translate_text(summary, target_lang, src); actions = translate_list_items(actions, target_lang, src); decisions = translate_list_items(decisions, target_lang, src); questions = translate_list_items(questions, target_lang, src)
        if translate_transcript_preview:
            for i, seg in enumerate(diarized): seg["text_tr"] = translate_text(seg.get("text", ""), target_lang, src) if i < 10 else ""
        progress.progress(82)
        status.markdown("📄 **Generating DOCX report...**"); docx_path = export_docx(summary, summary_tr, actions, decisions, questions, diarized, detected_lang, target_lang); progress.progress(100); status.empty(); note.empty()
        st.success(f"Done in {time.time() - t0:.1f} seconds."); st.markdown(f'<span class="pill amber">Detected: {LANG_NAMES.get(detected_lang, detected_lang)} · {lang_prob}% · {duration:.0f}s</span>', unsafe_allow_html=True)
        st.markdown('<div class="sec-header"><span class="sec-number">02</span>Results overview</div>', unsafe_allow_html=True)
        rc = st.columns(4)
        for col, val, lab in [(rc[0], len(actions), "Action Items"), (rc[1], len(decisions), "Decisions"), (rc[2], len(questions), "Questions"), (rc[3], len(transcript.split()), "Words")]: col.markdown(f'<div class="result-box"><div class="result-value">{val}</div><div class="result-label">{lab}</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="sec-header"><span class="sec-number">03</span>Executive summary</div>', unsafe_allow_html=True)
        t1, t2 = st.tabs([f"Original ({LANG_NAMES.get(detected_lang, detected_lang)})", f"Translated ({LANG_NAMES[target_lang]})"])
        with t1: st.markdown(f'<div class="glass-card">{escape_html(summary)}</div>', unsafe_allow_html=True)
        with t2: st.markdown(f'<div class="glass-card">{escape_html(summary_tr)}</div>', unsafe_allow_html=True)
        st.markdown('<div class="sec-header"><span class="sec-number">04</span>Key decisions</div>', unsafe_allow_html=True)
        if decisions:
            for i, item in enumerate(decisions, 1):
                with st.expander(f"Decision {i} - {item.get('speaker', 'Unknown')} ({item.get('time', '')})"): st.write(item.get("text", "")); st.caption(item.get("text_tr", ""))
        else: st.info("No explicit decisions were detected. This is normal for short or informal recordings.")
        st.markdown('<div class="sec-header"><span class="sec-number">05</span>Action items</div>', unsafe_allow_html=True)
        if actions:
            for i, item in enumerate(actions, 1):
                with st.expander(f"Action {i} - {item.get('speaker', 'Unknown')} ({item.get('time', '')})"): st.write(item.get("text", "")); st.write(f"Deadline: {item.get('deadline', 'Not specified')}"); st.caption(item.get("text_tr", ""))
        else: st.info("No clear action items were detected. Try a meeting sample with words like please, send, review, or follow up.")
        st.markdown('<div class="sec-header"><span class="sec-number">06</span>Open questions</div>', unsafe_allow_html=True)
        if questions:
            for i, item in enumerate(questions, 1):
                with st.expander(f"Question {i} - {item.get('speaker', 'Unknown')} ({item.get('time', '')})"): st.write(item.get("text", "")); st.caption(item.get("text_tr", ""))
        else: st.info("No open questions were detected.")
        st.markdown('<div class="sec-header"><span class="sec-number">07</span>Transcript</div>', unsafe_allow_html=True)
        with st.expander(f"View transcript ({len(diarized)} segments)"):
            for seg in diarized:
                st.markdown(f"**[{seg.get('start', 0)}s] {seg.get('speaker', '')}:** {seg.get('text', '')}")
                if seg.get("text_tr"): st.caption(seg.get("text_tr"))
        st.markdown('<div class="sec-header"><span class="sec-number">08</span>Download report</div>', unsafe_allow_html=True)
        with open(docx_path, "rb") as report: st.download_button("⤓ Download Meeting Minutes DOCX", report, file_name=f"MoM_Report_{LANG_NAMES.get(detected_lang, detected_lang)}_to_{LANG_NAMES[target_lang]}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as exc:
        status.empty(); progress.empty(); st.error(f"Processing failed: {str(exc)[:500]}"); st.info("Try a shorter MP3/WAV file first. For Streamlit Cloud, 30 seconds to 5 minutes is best for the public portfolio demo.")
    finally:
        gc.collect()
