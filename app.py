"""
MoM Automation — Meeting transcription, summarization, and translation
faster-whisper · BART · mBART-50 · Streamlit Cloud
"""

import os
import gc
import time
import subprocess
import html
import torch
import warnings
import streamlit as st
from datetime import datetime
from faster_whisper import WhisperModel
from transformers import (
    BartForConditionalGeneration, BartTokenizer,
    MBartForConditionalGeneration, MBart50TokenizerFast,
    pipeline
)
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

warnings.filterwarnings("ignore")

# ── Device setup ──────────────────────────────
DEVICE          = "cuda" if torch.cuda.is_available() else "cpu"
COMPUTE         = "float16" if DEVICE == "cuda" else "int8"
HF_DEVICE       = 0 if DEVICE == "cuda" else -1
MAX_FILE_MB     = 50
TRANSLATE_BATCH = 4
MAX_INPUT_LEN   = 512

OUTPUT_DIR      = "/tmp/mom_outputs"
os.makedirs(OUTPUT_DIR, exist_ok=True)


def safe_docx_text(value):
    """Remove XML-incompatible control characters before writing to DOCX."""
    if value is None:
        return ""
    value = str(value)
    return "".join(
        ch for ch in value
        if ch in "\t\n\r" or ord(ch) >= 32
    )


def safe_html_text(value):
    """Escape text before injecting it into unsafe_allow_html blocks."""
    return html.escape(str(value or "")).replace("\n", "<br>")


# ── All 52 supported languages ────────────────
LANG_NAMES = {
    "ar":"Arabic",   "cs":"Czech",      "de":"German",     "en":"English",
    "es":"Spanish",  "et":"Estonian",   "fi":"Finnish",    "fr":"French",
    "gu":"Gujarati", "hi":"Hindi",      "it":"Italian",    "ja":"Japanese",
    "kk":"Kazakh",   "ko":"Korean",     "lt":"Lithuanian", "lv":"Latvian",
    "my":"Burmese",  "ne":"Nepali",     "nl":"Dutch",      "ro":"Romanian",
    "ru":"Russian",  "si":"Sinhala",    "tr":"Turkish",    "vi":"Vietnamese",
    "zh":"Chinese",  "af":"Afrikaans",  "az":"Azerbaijani","bn":"Bengali",
    "fa":"Persian",  "he":"Hebrew",     "hr":"Croatian",   "id":"Indonesian",
    "ka":"Georgian", "km":"Khmer",      "mk":"Macedonian", "ml":"Malayalam",
    "mn":"Mongolian","mr":"Marathi",    "pl":"Polish",     "ps":"Pashto",
    "pt":"Portuguese","sv":"Swedish",   "sw":"Swahili",    "ta":"Tamil",
    "te":"Telugu",   "th":"Thai",       "tl":"Filipino",   "uk":"Ukrainian",
    "ur":"Urdu",     "xh":"Xhosa",      "gl":"Galician",   "sl":"Slovenian"
}

MBART_LANG_MAP = {
    "ar":"ar_AR","cs":"cs_CZ","de":"de_DE","en":"en_XX",
    "es":"es_XX","et":"et_EE","fi":"fi_FI","fr":"fr_XX",
    "gu":"gu_IN","hi":"hi_IN","it":"it_IT","ja":"ja_XX",
    "kk":"kk_KZ","ko":"ko_KR","lt":"lt_LT","lv":"lv_LV",
    "my":"my_MM","ne":"ne_NP","nl":"nl_XX","ro":"ro_RO",
    "ru":"ru_RU","si":"si_LK","tr":"tr_TR","vi":"vi_VN",
    "zh":"zh_CN","af":"af_ZA","az":"az_AZ","bn":"bn_IN",
    "fa":"fa_IR","he":"he_IL","hr":"hr_HR","id":"id_ID",
    "ka":"ka_GE","km":"km_KH","mk":"mk_MK","ml":"ml_IN",
    "mn":"mn_MN","mr":"mr_IN","pl":"pl_PL","ps":"ps_AF",
    "pt":"pt_XX","sv":"sv_SE","sw":"sw_KE","ta":"ta_IN",
    "te":"te_IN","th":"th_TH","tl":"tl_XX","uk":"uk_UA",
    "ur":"ur_PK","xh":"xh_ZA","gl":"gl_ES","sl":"sl_SI"
}


# ══════════════════════════════════════════════
#  PAGE CONFIG & STYLING
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="MoM Automation",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
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
.hero-title{font-family:'Instrument Serif',serif;font-size:4.5rem;font-weight:400;line-height:1;letter-spacing:-2px;margin:0;color:#f5f5f7}
.hero-title em{font-style:italic;background:linear-gradient(135deg,#c084fc 0%,#ec4899 50%,#f59e0b 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hero-subtitle{font-size:1rem;font-weight:300;color:#71717a;margin-top:1rem;max-width:540px;margin-left:auto;margin-right:auto;line-height:1.6}
.stat-card{background:linear-gradient(135deg,rgba(255,255,255,.04) 0%,rgba(255,255,255,.01) 100%);border:1px solid rgba(255,255,255,.08);border-radius:20px;padding:1.4rem 1rem;text-align:center;transition:all .3s cubic-bezier(.4,0,.2,1)}
.stat-card:hover{transform:translateY(-4px);border-color:rgba(168,85,247,.3)}
.stat-icon{font-size:1.5rem;margin-bottom:.4rem;display:block}
.stat-value{font-family:'Instrument Serif',serif;font-size:2.2rem;font-weight:400;color:#f5f5f7;line-height:1;margin-bottom:.3rem}
.stat-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#71717a;text-transform:uppercase;letter-spacing:2px}
.sec-header{font-family:'Instrument Serif',serif;font-size:1.8rem;font-style:italic;color:#f5f5f7;margin:2.5rem 0 1.2rem;display:flex;align-items:center;gap:.8rem}
.sec-header::after{content:'';flex:1;height:1px;background:linear-gradient(90deg,rgba(168,85,247,.4),transparent)}
.sec-number{font-family:'JetBrains Mono',monospace;font-size:.7rem;color:#a78bfa;background:rgba(168,85,247,.1);padding:.3rem .6rem;border-radius:6px;border:1px solid rgba(168,85,247,.2)}
.glass-card{background:linear-gradient(135deg,rgba(255,255,255,.03) 0%,rgba(255,255,255,.01) 100%);border:1px solid rgba(255,255,255,.08);border-radius:16px;padding:1.5rem;margin-bottom:1rem;color:#d4d4d8;line-height:1.7;font-size:.95rem}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#0a0a0f 0%,#07070b 100%) !important;border-right:1px solid rgba(255,255,255,.06) !important}
section[data-testid="stSidebar"] *{color:#a1a1aa !important}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{color:#f5f5f7 !important;font-family:'Instrument Serif',serif !important;font-style:italic !important;font-weight:400 !important}
section[data-testid="stSidebar"] strong{color:#f5f5f7 !important;font-weight:600 !important}
section[data-testid="stSidebar"] code{background:rgba(168,85,247,.1) !important;color:#c084fc !important;padding:.1rem .4rem !important;border-radius:4px !important;font-family:'JetBrains Mono',monospace !important;font-size:.75rem !important}
.stButton>button{background:linear-gradient(135deg,#a855f7 0%,#ec4899 50%,#f59e0b 100%) !important;color:white !important;border:none !important;border-radius:100px !important;font-family:'Inter',sans-serif !important;font-weight:600 !important;font-size:.95rem !important;padding:.85rem 2rem !important;width:100% !important;transition:all .3s !important;box-shadow:0 8px 32px rgba(168,85,247,.3) !important}
.stButton>button:hover{transform:translateY(-2px) !important;box-shadow:0 12px 40px rgba(168,85,247,.45) !important}
.stDownloadButton>button{background:linear-gradient(135deg,#10b981 0%,#06b6d4 100%) !important;color:white !important;border:none !important;border-radius:100px !important;font-weight:600 !important;width:100% !important;padding:.85rem 2rem !important;box-shadow:0 8px 32px rgba(16,185,129,.25) !important;transition:all .3s !important}
.stDownloadButton>button:hover{transform:translateY(-2px) !important}
.stSelectbox>div>div{background:rgba(255,255,255,.03) !important;border:1px solid rgba(255,255,255,.08) !important;border-radius:12px !important;color:#e4e4e7 !important}
.stFileUploader>div{background:linear-gradient(135deg,rgba(168,85,247,.04) 0%,rgba(236,72,153,.03) 100%) !important;border:2px dashed rgba(168,85,247,.25) !important;border-radius:20px !important}
.stFileUploader>div:hover{border-color:rgba(168,85,247,.5) !important}
.stTabs [data-baseweb="tab-list"]{background:rgba(255,255,255,.03);border-radius:100px;padding:4px;border:1px solid rgba(255,255,255,.06);gap:4px}
.stTabs [data-baseweb="tab"]{color:#71717a !important;border-radius:100px !important;padding:.5rem 1.5rem !important;font-weight:500 !important}
.stTabs [aria-selected="true"]{background:linear-gradient(135deg,#a855f7 0%,#ec4899 100%) !important;color:white !important;box-shadow:0 4px 16px rgba(168,85,247,.3) !important}
.streamlit-expanderHeader{background:rgba(255,255,255,.03) !important;border:1px solid rgba(255,255,255,.06) !important;border-radius:12px !important;color:#d4d4d8 !important;font-weight:500 !important}
.streamlit-expanderHeader:hover{background:rgba(168,85,247,.05) !important;border-color:rgba(168,85,247,.2) !important}
.streamlit-expanderContent{background:rgba(255,255,255,.02) !important;border:1px solid rgba(255,255,255,.06) !important;border-top:none !important;border-radius:0 0 12px 12px !important}
.stAlert{background:rgba(255,255,255,.03) !important;border:1px solid rgba(255,255,255,.08) !important;border-radius:14px !important;color:#d4d4d8 !important}
.stProgress>div>div{background:linear-gradient(90deg,#a855f7,#ec4899,#f59e0b) !important;border-radius:100px !important}
.stProgress>div{background:rgba(255,255,255,.05) !important;border-radius:100px !important}
audio{width:100%;border-radius:100px;filter:invert(.92) hue-rotate(180deg)}
.pill{display:inline-flex;align-items:center;gap:.4rem;padding:.4rem .9rem;border-radius:100px;font-size:.75rem;font-weight:500;margin-right:.4rem}
.pill-purple{background:rgba(168,85,247,.12);color:#c084fc;border:1px solid rgba(168,85,247,.3)}
.pill-pink{background:rgba(236,72,153,.12);color:#f472b6;border:1px solid rgba(236,72,153,.3)}
.pill-green{background:rgba(16,185,129,.12);color:#34d399;border:1px solid rgba(16,185,129,.3)}
.pill-amber{background:rgba(245,158,11,.12);color:#fbbf24;border:1px solid rgba(245,158,11,.3)}
.result-box{background:linear-gradient(135deg,rgba(168,85,247,.08) 0%,rgba(236,72,153,.04) 100%);border:1px solid rgba(168,85,247,.2);border-radius:18px;padding:1.4rem 1rem;text-align:center;transition:all .3s}
.result-box:hover{transform:translateY(-3px);border-color:rgba(168,85,247,.4);box-shadow:0 12px 40px rgba(168,85,247,.15)}
.result-value{font-family:'Instrument Serif',serif;font-size:2.5rem;font-weight:400;background:linear-gradient(135deg,#c084fc,#ec4899);-webkit-background-clip:text;-webkit-text-fill-color:transparent;line-height:1}
.result-label{font-family:'JetBrains Mono',monospace;font-size:.65rem;color:#a1a1aa;text-transform:uppercase;letter-spacing:2px;margin-top:.5rem}
.empty-state{text-align:center;padding:5rem 2rem;background:linear-gradient(135deg,rgba(168,85,247,.03) 0%,rgba(236,72,153,.02) 100%);border:1px dashed rgba(168,85,247,.2);border-radius:24px;margin-top:1rem}
.empty-icon{font-size:4.5rem;margin-bottom:1.2rem;display:inline-block;animation:pulse 3s ease-in-out infinite}
@keyframes pulse{0%,100%{transform:scale(1);opacity:1}50%{transform:scale(1.05);opacity:.85}}
.empty-title{font-family:'Instrument Serif',serif;font-style:italic;font-size:2rem;color:#f5f5f7;margin-bottom:.6rem}
.empty-text{color:#71717a;margin-top:.5rem;font-size:.95rem;margin-bottom:1.5rem}
p,li{color:#a1a1aa !important;line-height:1.7}
span,label{color:#d4d4d8 !important}
h1,h2,h3,h4{color:#f5f5f7 !important}
strong,b{color:#f5f5f7 !important}
hr{border:none !important;height:1px !important;background:linear-gradient(90deg,transparent,rgba(168,85,247,.3),transparent) !important;margin:2rem 0 !important}
.stSpinner>div{border-top-color:#a855f7 !important}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:#07070b}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#a855f7,#ec4899);border-radius:100px}
#MainMenu{visibility:hidden}footer{visibility:hidden}header{background:transparent !important}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  MODEL LOADERS  (cached — load once per session)
# ══════════════════════════════════════════════
@st.cache_resource(show_spinner=False)
def load_whisper():
    # 'small' works for English, Telugu, Hindi, Tamil, Arabic, etc.
    return WhisperModel("small", device=DEVICE, compute_type=COMPUTE)


@st.cache_resource(show_spinner=False)
def load_summarizer():
    tok = BartTokenizer.from_pretrained("facebook/bart-large-cnn")
    mdl = BartForConditionalGeneration.from_pretrained(
              "facebook/bart-large-cnn").to(DEVICE)
    mdl.eval()
    return tok, mdl


@st.cache_resource(show_spinner=False)
def load_classifier():
    return pipeline(
        "zero-shot-classification",
        model="facebook/bart-large-mnli",
        device=HF_DEVICE
    )


@st.cache_resource(show_spinner=False)
def load_translator():
    tok = MBart50TokenizerFast.from_pretrained(
              "facebook/mbart-large-50-many-to-many-mmt")
    mdl = MBartForConditionalGeneration.from_pretrained(
              "facebook/mbart-large-50-many-to-many-mmt").to(DEVICE)
    mdl.eval()
    return tok, mdl


# ══════════════════════════════════════════════
#  AUDIO CONVERSION
# ══════════════════════════════════════════════
def convert_to_wav(input_path):
    """
    Convert any audio/video → 16kHz mono WAV using ffmpeg.
    Returns (wav_path, error_str).  error_str is None on success.
    """
    if not os.path.exists(input_path):
        return None, f"File not found: {input_path}"

    output_path = input_path.rsplit(".", 1)[0] + "_conv.wav"

    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-vn",                    # strip video
        "-acodec", "pcm_s16le",   # uncompressed PCM
        "-ar",  "16000",          # 16 kHz  (Whisper's native rate)
        "-ac",  "1",              # mono
        "-loglevel", "error",
        output_path
    ]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=300
        )
        if result.returncode != 0:
            return None, f"ffmpeg: {result.stderr[:300]}"
        if not os.path.exists(output_path):
            return None, "ffmpeg produced no output file"
        if os.path.getsize(output_path) < 1000:
            return None, "Converted file is empty — MP4 may have no audio"
        return output_path, None
    except FileNotFoundError:
        return None, "ffmpeg not installed (add 'ffmpeg' to packages.txt)"
    except subprocess.TimeoutExpired:
        return None, "ffmpeg timed out"
    except Exception as e:
        return None, str(e)[:200]


# ══════════════════════════════════════════════
#  TRANSCRIPTION
# ══════════════════════════════════════════════
def _run_whisper(model, path, language=None, lenient=False):
    """
    Call faster-whisper with correct parameters.
    NOTE: faster-whisper does NOT support 'best_of' — only beam_size.
    Returns (segments_list, info_obj) or raises.
    """
    kwargs = {
        "beam_size"  : 1 if lenient else 5,
        "language"   : language,
        "condition_on_previous_text": False if lenient else True,
    }
    if lenient:
        # Accept speech with much lower confidence
        kwargs["no_speech_threshold"]        = 0.1
        kwargs["log_prob_threshold"]         = -2.0
        kwargs["compression_ratio_threshold"] = 3.0

    raw, info = model.transcribe(path, **kwargs)

    segments = []
    for s in raw:
        text = s.text.strip()
        if text:
            segments.append({
                "start": round(s.start, 2),
                "end"  : round(s.end,   2),
                "text" : text
            })
    return segments, info


def transcribe(audio_path):
    """
    Robust transcription pipeline:
      1. Convert input → clean 16kHz WAV
      2. Pass 1: Standard Whisper (beam=5)
      3. Pass 2: Lenient Whisper with explicit language hint
    Returns (transcript, segments, language, confidence, duration)
    """
    # ── Convert to WAV ────────────────────────
    wav_path, err = convert_to_wav(audio_path)
    if wav_path is None:
        st.warning(f"⚠️ Audio conversion: {err}. Trying original file...")
        wav_path = audio_path

    # ── Sanity checks ─────────────────────────
    if not os.path.exists(wav_path):
        st.error("❌ Audio file is missing. Please re-upload.")
        return "", [], "en", 0.0, 0

    file_size = os.path.getsize(wav_path)
    if file_size < 1000:
        st.error(
            f"❌ Audio file is empty ({file_size} bytes). "
            "The video may not have an audio track."
        )
        return "", [], "en", 0.0, 0

    # ── Load Whisper ──────────────────────────
    try:
        model = load_whisper()
    except Exception as e:
        st.error(f"❌ Could not load Whisper model: {str(e)[:200]}")
        return "", [], "en", 0.0, 0

    # ── Pass 1: Standard ─────────────────────
    info       = None
    det_lang   = "en"
    duration   = 0.0

    try:
        segments, info = _run_whisper(model, wav_path, language=None, lenient=False)
        det_lang = info.language or "en"
        duration = info.duration or 0.0
        if segments:
            return (
                " ".join(s["text"] for s in segments),
                segments, det_lang,
                round(info.language_probability * 100, 1),
                duration
            )
    except Exception as e:
        st.warning(f"⚠️ First transcription pass failed ({str(e)[:80]}). Retrying...")

    # ── Pass 2: Lenient with language hint ────
    try:
        segments, info = _run_whisper(
            model, wav_path,
            language=det_lang if det_lang != "en" else None,
            lenient=True
        )
        if info:
            det_lang = info.language or det_lang
            duration = info.duration or duration
        if segments:
            return (
                " ".join(s["text"] for s in segments),
                segments, det_lang,
                round((info.language_probability * 100) if info else 0, 1),
                duration
            )
    except Exception as e:
        st.warning(f"⚠️ Second transcription pass failed ({str(e)[:80]}).")

    # ── Both passes failed ────────────────────
    st.error(
        "❌ **Could not extract speech from this audio.**\n\n"
        f"File size: {file_size/1024:.1f} KB · "
        f"Detected language: {det_lang} · "
        f"Duration: {duration:.1f}s\n\n"
        "**Common causes:**\n"
        "- Video has no audio track (check in a media player first)\n"
        "- Audio is entirely music with no speech\n"
        "- Microphone was muted during recording\n\n"
        "**Fix:** Convert the file to MP3 using "
        "[cloudconvert.com](https://cloudconvert.com) and try again."
    )
    return "", [], "en", 0.0, 0


# ══════════════════════════════════════════════
#  NLP FUNCTIONS
# ══════════════════════════════════════════════
def simple_diarize(segments, num_speakers=3):
    """Assign speaker labels based on pauses between segments."""
    out, spk, last_end = [], 1, 0.0
    for seg in segments:
        if seg["start"] - last_end >= 1.5:
            spk = (spk % num_speakers) + 1
        out.append({**seg, "speaker": f"Speaker {spk}"})
        last_end = seg["end"]
    return out


@torch.inference_mode()
def get_summary(text):
    """BART-based extractive → abstractive summary."""
    if not text or len(text.split()) < 20:
        return text or ""
    try:
        tok, mdl = load_summarizer()
        words  = text.split()
        chunks = [" ".join(words[i:i+700]) for i in range(0, len(words), 700)]
        parts  = []
        for c in chunks:
            if len(c.split()) < 30:
                continue
            inp = tok(c, return_tensors="pt", max_length=1024, truncation=True)
            inp = {k: v.to(DEVICE) for k, v in inp.items()}
            ids = mdl.generate(
                inp["input_ids"],
                max_length=130, min_length=30,
                num_beams=4, no_repeat_ngram_size=3
            )
            parts.append(tok.decode(ids[0], skip_special_tokens=True))
        return " ".join(parts) if parts else text[:500]
    except Exception as e:
        return text[:500]


def get_actions(segs):
    """Keyword pre-filter + zero-shot classification for action items."""
    if not segs:
        return []
    try:
        clf  = load_classifier()
        kws  = ["will","should","need to","must","please","send","schedule",
                "review","ensure","prepare","follow up","make sure","complete"]
        lbls = ["action item","task assignment","general discussion"]
        out  = []
        for s in segs:
            t = s.get("text","").lower()
            if not any(k in t for k in kws):
                continue
            if len(s["text"].split()) < 4:
                continue
            r = clf(s["text"], candidate_labels=lbls)
            if r["labels"][0] != "general discussion" and r["scores"][0] > 0.40:
                out.append({
                    "speaker": s.get("speaker","Unknown"),
                    "text"   : s["text"],
                    "time"   : f"{s['start']}s"
                })
        return out
    except Exception:
        return []


def get_decisions(segs):
    """Keyword pre-filter + zero-shot classification for decisions."""
    if not segs:
        return []
    try:
        clf  = load_classifier()
        kws  = ["decided","agreed","approved","confirmed","moving forward",
                "we will","finalized","going with","accepted","resolved"]
        lbls = ["decision made","agreement reached","general statement"]
        out  = []
        for s in segs:
            t = s.get("text","").lower()
            if not any(k in t for k in kws):
                continue
            if len(s["text"].split()) < 4:
                continue
            r = clf(s["text"], candidate_labels=lbls)
            if r["labels"][0] != "general statement" and r["scores"][0] > 0.38:
                out.append({
                    "speaker": s.get("speaker","Unknown"),
                    "text"   : s["text"],
                    "time"   : f"{s['start']}s"
                })
        return out
    except Exception:
        return []


@torch.inference_mode()
def translate_batch(texts, src, tgt):
    """Batch-translate a list of strings. Falls back to originals on any error."""
    if not texts:
        return []
    if src == tgt:
        return list(texts)
    if src not in MBART_LANG_MAP or tgt not in MBART_LANG_MAP:
        return list(texts)   # unsupported language — return originals
    try:
        tok, mdl = load_translator()
        tok.src_lang = MBART_LANG_MAP[src]
        tgt_id = tok.lang_code_to_id[MBART_LANG_MAP[tgt]]
        inp = tok(
            list(texts),
            return_tensors="pt",
            max_length=MAX_INPUT_LEN,
            truncation=True,
            padding=True
        )
        inp = {k: v.to(DEVICE) for k, v in inp.items()}
        out = mdl.generate(
            **inp,
            forced_bos_token_id=tgt_id,
            max_length=MAX_INPUT_LEN,
            num_beams=2,
            no_repeat_ngram_size=3
        )
        return tok.batch_decode(out, skip_special_tokens=True)
    except Exception as e:
        st.warning(f"Translation batch error (using originals): {str(e)[:80]}")
        return list(texts)


# ══════════════════════════════════════════════
#  DOCX EXPORT
# ══════════════════════════════════════════════
def _set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _add_banner(doc, title, color="1F497D"):
    t = doc.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    _set_cell_bg(c, color)
    p = c.paragraphs[0]
    r = p.add_run("  " + title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()


def export_docx(summary, summary_tr, actions, decisions,
                diarized, src_lang, tgt_lang):
    """Build and save a bilingual DOCX report. Returns file path."""
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1.0)
    sec.top_margin  = sec.bottom_margin = Inches(0.8)

    src_name = LANG_NAMES.get(src_lang, src_lang)
    tgt_name = LANG_NAMES.get(tgt_lang, tgt_lang)

    # ── Title ─────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MINUTES OF MEETING")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s.add_run(
        "Original: " + src_name + "   |   Translated: " + tgt_name
    ).italic = True
    doc.add_paragraph(datetime.now().strftime("%B %d, %Y  |  %I:%M %p"))
    doc.add_paragraph()

    # ── 1. Summary ────────────────────────────
    _add_banner(doc, "1.  Executive Summary", "1F497D")
    doc.add_paragraph(safe_docx_text(summary or "No summary available."))
    p2 = doc.add_paragraph(safe_docx_text(summary_tr or "Translation unavailable."))
    if p2.runs:
        p2.runs[0].italic = True
    doc.add_paragraph()

    # ── 2. Decisions ──────────────────────────
    _add_banner(doc, "2.  Key Decisions", "375623")
    if decisions:
        for d in decisions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + d.get("speaker","") + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0x37, 0x56, 0x23)
            p.add_run(safe_docx_text(d.get("text","")))
            p.add_run("  (" + d.get("time","") + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tr_ = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_.font.size = Pt(9)
            tt = tp.add_run(safe_docx_text(d.get("text_tr") or ""))
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No decisions extracted.")
    doc.add_paragraph()

    # ── 3. Actions ────────────────────────────
    _add_banner(doc, "3.  Action Items", "C00000")
    if actions:
        for a in actions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + a.get("speaker","") + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            p.add_run(safe_docx_text(a.get("text","")))
            p.add_run("  (" + a.get("time","") + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tr_ = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_.font.size = Pt(9)
            tt = tp.add_run(safe_docx_text(a.get("text_tr") or ""))
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    # ── 4. Transcript ─────────────────────────
    _add_banner(
        doc,
        "4.  Full Transcript (" + src_name + " + " + tgt_name + ")",
        "595959"
    )
    for seg in diarized:
        p = doc.add_paragraph()
        r = p.add_run(
            "[" + str(seg.get("start", 0)) + "s]  " +
            seg.get("speaker", "") + ": "
        )
        r.bold = True
        r.font.size = Pt(9)
        p.add_run(safe_docx_text(seg.get("text", ""))).font.size = Pt(9)
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.4)
        tr_ = tp.add_run("  🌐 " + tgt_name + ": ")
        tr_.font.size = Pt(8)
        tt = tp.add_run(safe_docx_text(seg.get("text_tr") or ""))
        tt.font.size = Pt(8)
        tt.italic = True

    # Safe filename — remove spaces and special chars
    safe_src = src_name.replace(" ", "_")
    safe_tgt = tgt_name.replace(" ", "_")
    path = f"{OUTPUT_DIR}/MoM_{safe_src}_to_{safe_tgt}_{int(time.time())}.docx"
    doc.save(path)
    return path


# ══════════════════════════════════════════════
#  UI — HERO
# ══════════════════════════════════════════════
st.markdown("""
<div class="hero-wrap">
    <span class="hero-eyebrow">⟡  AI MEETING INTELLIGENCE</span>
    <h1 class="hero-title">Minutes that <em>write themselves</em></h1>
    <p class="hero-subtitle">Upload any meeting audio in any language — get a polished,
    bilingual, speaker-attributed report in under a minute.</p>
</div>
""", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
for col, icon, val, lbl in [
    (c1, "🌐", "52",   "Languages"),
    (c2, "🤖", "4",    "AI Models"),
    (c3, "⚡", "~60s", "Per File"),
    (c4, "📄", "DOCX", "Export"),
]:
    col.markdown(
        f'<div class="stat-card"><span class="stat-icon">{icon}</span>'
        f'<div class="stat-value">{val}</div>'
        f'<div class="stat-label">{lbl}</div></div>',
        unsafe_allow_html=True
    )

st.markdown("<br>", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")
    target_lang = st.selectbox(
        "🌐 Translate output to:",
        options=list(LANG_NAMES.keys()),
        format_func=lambda x: LANG_NAMES[x] + " (" + x + ")",
        index=list(LANG_NAMES.keys()).index("en")
    )
    num_speakers = st.slider(
        "👥 Number of speakers", min_value=1, max_value=6, value=3
    )
    st.markdown("---")
    st.markdown("### 📁 Supported Formats")
    st.markdown("`mp3`  `wav`  `m4a`  `mp4`")
    st.caption(f"Max file size: {MAX_FILE_MB} MB")
    st.markdown("---")
    st.markdown("### 🤖 Models")
    st.markdown("🎙️ **faster-whisper small**"); st.caption("Speech to Text — 52 languages")
    st.markdown("📝 **BART-large-CNN**");       st.caption("Summarization")
    st.markdown("🏷️ **BART-large-MNLI**");      st.caption("Classification")
    st.markdown("🌐 **mBART-50**");              st.caption("Translation — 52 languages")

# ── Upload ────────────────────────────────────
st.markdown(
    '<div class="sec-header"><span class="sec-number">01</span>Upload your meeting</div>',
    unsafe_allow_html=True
)

uploaded = st.file_uploader(
    "Drag and drop your audio file here",
    type=["mp3","wav","m4a","mp4"],
    label_visibility="collapsed"
)

if uploaded:
    # Validate size
    file_bytes = uploaded.getvalue()
    file_mb    = len(file_bytes) / (1024 * 1024)
    if file_mb > MAX_FILE_MB:
        st.error(
            f"❌ File too large ({file_mb:.1f} MB). "
            f"Maximum is {MAX_FILE_MB} MB."
        )
        st.stop()

    # Save to /tmp with sanitized name
    safe_name  = "".join(c if c.isalnum() or c in ".-_" else "_"
                         for c in uploaded.name)
    audio_path = f"/tmp/upload_{safe_name}"
    with open(audio_path, "wb") as f:
        f.write(file_bytes)

    st.audio(audio_path)

    ca, cb, cc = st.columns(3)
    ca.markdown(f'<span class="pill pill-purple">📄 {uploaded.name}</span>', unsafe_allow_html=True)
    cb.markdown(f'<span class="pill pill-pink">💾 {file_mb:.2f} MB</span>',   unsafe_allow_html=True)
    cc.markdown(f'<span class="pill pill-green">🌐 → {LANG_NAMES.get(target_lang)}</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("✨  Generate Meeting Minutes"):

        progress = st.progress(0)
        status   = st.empty()
        t0       = time.time()

        try:
            # 1. Transcribe
            status.markdown("🎙️ **Transcribing audio...**")
            transcript, segments, det_lang, det_prob, duration = transcribe(audio_path)

            if not segments:
                progress.empty()
                status.empty()
                st.stop()

            progress.progress(20)
            st.markdown(
                f'<span class="pill pill-amber">'
                f'⟡ {LANG_NAMES.get(det_lang, det_lang)} · '
                f'{det_prob}% · {duration:.0f}s</span>',
                unsafe_allow_html=True
            )

            # 2. Diarize
            status.markdown("👥 **Assigning speakers...**")
            diarized = simple_diarize(segments, num_speakers)
            progress.progress(30)

            # 3. Extract
            status.markdown("🧠 **Summarizing...**")
            summary = get_summary(transcript)
            progress.progress(45)

            status.markdown("✅ **Extracting actions & decisions...**")
            actions   = get_actions(diarized)
            decisions = get_decisions(diarized)
            progress.progress(60)

            # 4. Translate
            tgt_name  = LANG_NAMES.get(target_lang, target_lang)
            src_name  = LANG_NAMES.get(det_lang, det_lang)
            same_lang = (det_lang == target_lang)

            if same_lang:
                status.markdown("⚡ **Same language — skipping translation...**")
                summary_tr = summary
                for x in actions:   x["text_tr"] = x["text"]
                for x in decisions: x["text_tr"] = x["text"]
                for x in diarized:  x["text_tr"] = x["text"]
                progress.progress(90)
            else:
                # Translate summary + actions + decisions in one call
                all_texts = (
                    [summary] +
                    [a["text"] for a in actions] +
                    [d["text"] for d in decisions]
                )
                n_actions   = len(actions)
                n_decisions = len(decisions)

                status.markdown(
                    f"🌐 **Translating {src_name} → {tgt_name}** "
                    f"(summary + {n_actions} actions + {n_decisions} decisions)..."
                )
                translated_header = translate_batch(all_texts, det_lang, target_lang)
                progress.progress(70)

                summary_tr = translated_header[0] if translated_header else summary
                for i, a in enumerate(actions):
                    idx = 1 + i
                    a["text_tr"] = translated_header[idx] if idx < len(translated_header) else a["text"]
                for i, d in enumerate(decisions):
                    idx = 1 + n_actions + i
                    d["text_tr"] = translated_header[idx] if idx < len(translated_header) else d["text"]

                # Translate transcript segments in batches
                seg_texts = [s["text"] for s in diarized]
                n_total   = len(seg_texts)
                translated_segs = []

                for bi in range(0, n_total, TRANSLATE_BATCH):
                    chunk = seg_texts[bi:bi + TRANSLATE_BATCH]
                    translated_segs.extend(
                        translate_batch(chunk, det_lang, target_lang)
                    )
                    done = min(bi + TRANSLATE_BATCH, n_total)
                    pct  = 70 + int(done / n_total * 20)
                    progress.progress(min(90, pct))
                    status.markdown(
                        f"🌐 **Translating transcript** · "
                        f"segment {done} / {n_total}"
                    )

                for seg, tr in zip(diarized, translated_segs):
                    seg["text_tr"] = tr

                # Ensure every segment has text_tr
                for seg in diarized:
                    if not seg.get("text_tr"):
                        seg["text_tr"] = seg["text"]

            # 5. Export
            status.markdown("📄 **Generating DOCX...**")
            docx_path = export_docx(
                summary, summary_tr, actions, decisions,
                diarized, det_lang, target_lang
            )
            progress.progress(100)
            status.empty()

            elapsed = time.time() - t0
            st.success(f"✅ Done in {elapsed:.1f} seconds!")

            # Cleanup
            gc.collect()
            if DEVICE == "cuda":
                torch.cuda.empty_cache()

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Results ──────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">02</span>Results overview</div>',
                unsafe_allow_html=True
            )
            r1, r2, r3, r4 = st.columns(4)
            for col, val, lbl in [
                (r1, len(actions),          "Action Items"),
                (r2, len(decisions),         "Key Decisions"),
                (r3, len(diarized),          "Segments"),
                (r4, len(transcript.split()),"Words"),
            ]:
                col.markdown(
                    f'<div class="result-box">'
                    f'<div class="result-value">{val}</div>'
                    f'<div class="result-label">{lbl}</div></div>',
                    unsafe_allow_html=True
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Summary ──────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">03</span>Executive summary</div>',
                unsafe_allow_html=True
            )
            t1, t2 = st.tabs([
                "🗣️ Original (" + src_name + ")",
                "🌐 Translated (" + tgt_name + ")"
            ])
            with t1:
                st.markdown(
                    '<div class="glass-card">' + safe_html_text(summary or "No summary.") + '</div>',
                    unsafe_allow_html=True
                )
            with t2:
                st.markdown(
                    '<div class="glass-card">' + safe_html_text(summary_tr or "Translation unavailable.") + '</div>',
                    unsafe_allow_html=True
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Decisions ────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">04</span>Key decisions</div>',
                unsafe_allow_html=True
            )
            if decisions:
                for i, d in enumerate(decisions, 1):
                    with st.expander(
                        f"Decision {i} — {d.get('speaker','')} ({d.get('time','')})"
                    ):
                        st.markdown("**🗣️ Original:** " + d.get("text",""))
                        st.markdown("**🌐 " + tgt_name + ":** " + (d.get("text_tr") or d.get("text","")))
            else:
                st.info("No decisions extracted from this meeting.")

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Actions ──────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">05</span>Action items</div>',
                unsafe_allow_html=True
            )
            if actions:
                for i, a in enumerate(actions, 1):
                    with st.expander(
                        f"Action {i} — {a.get('speaker','')} ({a.get('time','')})"
                    ):
                        st.markdown("**🗣️ Original:** " + a.get("text",""))
                        st.markdown("**🌐 " + tgt_name + ":** " + (a.get("text_tr") or a.get("text","")))
            else:
                st.info("No action items extracted from this meeting.")

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Transcript ───────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">06</span>Full transcript</div>',
                unsafe_allow_html=True
            )
            with st.expander(
                f"View full bilingual transcript ({len(diarized)} segments)"
            ):
                for seg in diarized:
                    st.markdown(
                        f"**[{seg.get('start',0)}s] {seg.get('speaker','')}:** "
                        f"{seg.get('text','')}  \n"
                        f"*🌐 {seg.get('text_tr') or seg.get('text','')}*"
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Download ─────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">07</span>Download report</div>',
                unsafe_allow_html=True
            )
            with open(docx_path, "rb") as f:
                st.download_button(
                    label     = "⤓  Download Bilingual Meeting Minutes (DOCX)",
                    data      = f,
                    file_name = f"MoM_{src_name}_to_{tgt_name}.docx",
                    mime      = (
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.document"
                    )
                )

        except Exception as e:
            progress.empty()
            status.empty()
            st.error(
                f"❌ Unexpected error: {str(e)[:300]}\n\n"
                "Please refresh the page and try again."
            )

else:
    st.markdown("""
    <div class="empty-state">
        <div class="empty-icon">🎙️</div>
        <div class="empty-title">Drop your audio<br>and watch the magic</div>
        <div class="empty-text">Supports .mp3 · .wav · .m4a · .mp4</div>
        <div style="margin-top:1.5rem">
            <span class="pill pill-purple">⟡ Auto language detection</span>
            <span class="pill pill-pink">⟡ 52 languages</span>
            <span class="pill pill-green">⟡ DOCX export</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
