"""
MoM Automation — Production-grade meeting transcription, summarization, and translation
Built with: faster-whisper, BART, mBART-50, Streamlit
Optimized for: Streamlit Cloud (CPU-only, 1GB RAM, 60s request timeout)
"""

import os
import gc
import time
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

# ══════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════
DEVICE          = "cuda" if torch.cuda.is_available() else "cpu"
COMPUTE         = "float16" if DEVICE == "cuda" else "int8"
HF_DEVICE       = 0 if DEVICE == "cuda" else -1
MAX_FILE_MB     = 50         # Streamlit Cloud limit
MAX_DURATION_S  = 1800       # 30 min audio limit
TRANSLATE_BATCH = 6          # batch size for translation
MAX_INPUT_LEN   = 512        # tokens per translation chunk

LANG_NAMES = {
    "ar":"Arabic","cs":"Czech","de":"German","en":"English",
    "es":"Spanish","et":"Estonian","fi":"Finnish","fr":"French",
    "gu":"Gujarati","hi":"Hindi","it":"Italian","ja":"Japanese",
    "kk":"Kazakh","ko":"Korean","lt":"Lithuanian","lv":"Latvian",
    "my":"Burmese","ne":"Nepali","nl":"Dutch","ro":"Romanian",
    "ru":"Russian","si":"Sinhala","tr":"Turkish","vi":"Vietnamese",
    "zh":"Chinese","af":"Afrikaans","az":"Azerbaijani","bn":"Bengali",
    "fa":"Persian","he":"Hebrew","hr":"Croatian","id":"Indonesian",
    "ka":"Georgian","km":"Khmer","mk":"Macedonian","ml":"Malayalam",
    "mn":"Mongolian","mr":"Marathi","pl":"Polish","ps":"Pashto",
    "pt":"Portuguese","sv":"Swedish","sw":"Swahili","ta":"Tamil",
    "te":"Telugu","th":"Thai","tl":"Filipino","uk":"Ukrainian",
    "ur":"Urdu","xh":"Xhosa","gl":"Galician","sl":"Slovenian"
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
* { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; -webkit-font-smoothing: antialiased; }
.stApp { background: #07070b; min-height: 100vh; position: relative; overflow-x: hidden; }
.stApp::before { content: ''; position: fixed; top: -20%; left: -10%; width: 60%; height: 60%; background: radial-gradient(circle, rgba(168,85,247,0.18) 0%, transparent 60%); pointer-events: none; z-index: 0; animation: float1 25s ease-in-out infinite; }
.stApp::after { content: ''; position: fixed; bottom: -20%; right: -10%; width: 60%; height: 60%; background: radial-gradient(circle, rgba(236,72,153,0.15) 0%, transparent 60%); pointer-events: none; z-index: 0; animation: float2 30s ease-in-out infinite; }
@keyframes float1 { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(40px,30px) scale(1.1); } }
@keyframes float2 { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(-40px,-30px) scale(1.15); } }
.main .block-container { position: relative; z-index: 1; padding-top: 2rem; }
.hero-wrap { text-align: center; margin: 1rem 0 3rem 0; }
.hero-eyebrow { display: inline-block; font-family: 'JetBrains Mono', monospace; font-size: 0.7rem; letter-spacing: 4px; text-transform: uppercase; color: #a78bfa; padding: 0.4rem 1rem; background: rgba(168,85,247,0.08); border: 1px solid rgba(168,85,247,0.25); border-radius: 100px; margin-bottom: 1.2rem; }
.hero-title { font-family: 'Instrument Serif', serif; font-size: 4.5rem; font-weight: 400; line-height: 1; letter-spacing: -2px; margin: 0; color: #f5f5f7; }
.hero-title em { font-style: italic; background: linear-gradient(135deg, #c084fc 0%, #ec4899 50%, #f59e0b 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }
.hero-subtitle { font-size: 1rem; font-weight: 300; color: #71717a; margin-top: 1rem; max-width: 540px; margin-left: auto; margin-right: auto; line-height: 1.6; }
.stat-card { background: linear-gradient(135deg, rgba(255,255,255,0.04) 0%, rgba(255,255,255,0.01) 100%); border: 1px solid rgba(255,255,255,0.08); border-radius: 20px; padding: 1.4rem 1rem; text-align: center; transition: all 0.3s cubic-bezier(0.4,0,0.2,1); }
.stat-card:hover { transform: translateY(-4px); border-color: rgba(168,85,247,0.3); }
.stat-icon { font-size: 1.5rem; margin-bottom: 0.4rem; display: block; }
.stat-value { font-family: 'Instrument Serif', serif; font-size: 2.2rem; font-weight: 400; color: #f5f5f7; line-height: 1; margin-bottom: 0.3rem; }
.stat-label { font-family: 'JetBrains Mono', monospace; font-size: 0.65rem; color: #71717a; text-transform: uppercase; letter-spacing: 2px; }
.sec-header { font-family: 'Instrument Serif', serif; font-size: 1.8rem; font-style: italic; color: #f5f5f7; margin: 2.5rem 0 1.2rem 0; display: flex; align-items: center; gap: 0.8rem; }
.sec-header::after { content: ''; flex: 1; height: 1px; background: linear-gradient(90deg, rgba(168,85,247,0.4), transparent); }
.sec-number { font-family: 'JetBrains Mono', monospace; font-size: 0.7rem; color: #a78bfa; background: rgba(168,85,247,0.1); padding: 0.3rem 0.6rem; border-radius: 6px; border: 1px solid rgba(168,85,247,0.2); }
.glass-card { background: linear-gradient(135deg, rgba(255,255,255,0.03) 0%, rgba(255,255,255,0.01) 100%); border: 1px solid rgba(255,255,255,0.08); border-radius: 16px; padding: 1.5rem; margin-bottom: 1rem; color: #d4d4d8; line-height: 1.7; font-size: 0.95rem; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0a0a0f 0%, #07070b 100%) !important; border-right: 1px solid rgba(255,255,255,0.06) !important; }
section[data-testid="stSidebar"] * { color: #a1a1aa !important; }
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #f5f5f7 !important; font-family: 'Instrument Serif', serif !important; font-style: italic !important; font-weight: 400 !important; }
section[data-testid="stSidebar"] strong { color: #f5f5f7 !important; font-weight: 600 !important; }
section[data-testid="stSidebar"] code { background: rgba(168,85,247,0.1) !important; color: #c084fc !important; padding: 0.1rem 0.4rem !important; border-radius: 4px !important; font-family: 'JetBrains Mono', monospace !important; font-size: 0.75rem !important; }
.stButton > button { background: linear-gradient(135deg, #a855f7 0%, #ec4899 50%, #f59e0b 100%) !important; color: white !important; border: none !important; border-radius: 100px !important; font-family: 'Inter', sans-serif !important; font-weight: 600 !important; font-size: 0.95rem !important; padding: 0.85rem 2rem !important; width: 100% !important; transition: all 0.3s !important; box-shadow: 0 8px 32px rgba(168,85,247,0.3) !important; }
.stButton > button:hover { transform: translateY(-2px) !important; box-shadow: 0 12px 40px rgba(168,85,247,0.45) !important; }
.stDownloadButton > button { background: linear-gradient(135deg, #10b981 0%, #06b6d4 100%) !important; color: white !important; border: none !important; border-radius: 100px !important; font-weight: 600 !important; width: 100% !important; padding: 0.85rem 2rem !important; box-shadow: 0 8px 32px rgba(16,185,129,0.25) !important; }
.stDownloadButton > button:hover { transform: translateY(-2px) !important; }
.stTextArea textarea { background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(255,255,255,0.08) !important; border-radius: 14px !important; color: #e4e4e7 !important; }
.stSelectbox > div > div { background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(255,255,255,0.08) !important; border-radius: 12px !important; color: #e4e4e7 !important; }
.stFileUploader > div { background: linear-gradient(135deg, rgba(168,85,247,0.04) 0%, rgba(236,72,153,0.03) 100%) !important; border: 2px dashed rgba(168,85,247,0.25) !important; border-radius: 20px !important; }
.stFileUploader > div:hover { border-color: rgba(168,85,247,0.5) !important; }
.stTabs [data-baseweb="tab-list"] { background: rgba(255,255,255,0.03); border-radius: 100px; padding: 4px; border: 1px solid rgba(255,255,255,0.06); gap: 4px; }
.stTabs [data-baseweb="tab"] { color: #71717a !important; border-radius: 100px !important; padding: 0.5rem 1.5rem !important; font-weight: 500 !important; }
.stTabs [aria-selected="true"] { background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important; color: white !important; box-shadow: 0 4px 16px rgba(168,85,247,0.3) !important; }
.streamlit-expanderHeader { background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(255,255,255,0.06) !important; border-radius: 12px !important; color: #d4d4d8 !important; font-weight: 500 !important; }
.streamlit-expanderHeader:hover { background: rgba(168,85,247,0.05) !important; border-color: rgba(168,85,247,0.2) !important; }
.streamlit-expanderContent { background: rgba(255,255,255,0.02) !important; border: 1px solid rgba(255,255,255,0.06) !important; border-top: none !important; border-radius: 0 0 12px 12px !important; }
.stAlert { background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(255,255,255,0.08) !important; border-radius: 14px !important; color: #d4d4d8 !important; }
.stProgress > div > div { background: linear-gradient(90deg, #a855f7, #ec4899, #f59e0b) !important; border-radius: 100px !important; }
.stProgress > div { background: rgba(255,255,255,0.05) !important; border-radius: 100px !important; }
audio { width: 100%; border-radius: 100px; filter: invert(0.92) hue-rotate(180deg); }
.pill { display: inline-flex; align-items: center; gap: 0.4rem; padding: 0.4rem 0.9rem; border-radius: 100px; font-size: 0.75rem; font-weight: 500; backdrop-filter: blur(10px); margin-right: 0.4rem; }
.pill-purple { background: rgba(168,85,247,0.12); color: #c084fc; border: 1px solid rgba(168,85,247,0.3); }
.pill-pink { background: rgba(236,72,153,0.12); color: #f472b6; border: 1px solid rgba(236,72,153,0.3); }
.pill-green { background: rgba(16,185,129,0.12); color: #34d399; border: 1px solid rgba(16,185,129,0.3); }
.pill-amber { background: rgba(245,158,11,0.12); color: #fbbf24; border: 1px solid rgba(245,158,11,0.3); }
.pill-red { background: rgba(239,68,68,0.12); color: #f87171; border: 1px solid rgba(239,68,68,0.3); }
.result-box { background: linear-gradient(135deg, rgba(168,85,247,0.08) 0%, rgba(236,72,153,0.04) 100%); border: 1px solid rgba(168,85,247,0.2); border-radius: 18px; padding: 1.4rem 1rem; text-align: center; }
.result-box:hover { transform: translateY(-3px); border-color: rgba(168,85,247,0.4); }
.result-value { font-family: 'Instrument Serif', serif; font-size: 2.5rem; font-weight: 400; background: linear-gradient(135deg, #c084fc, #ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1; }
.result-label { font-family: 'JetBrains Mono', monospace; font-size: 0.65rem; color: #a1a1aa; text-transform: uppercase; letter-spacing: 2px; margin-top: 0.5rem; }
.empty-state { text-align: center; padding: 5rem 2rem; background: linear-gradient(135deg, rgba(168,85,247,0.03) 0%, rgba(236,72,153,0.02) 100%); border: 1px dashed rgba(168,85,247,0.2); border-radius: 24px; margin-top: 1rem; }
.empty-icon { font-size: 4.5rem; margin-bottom: 1.2rem; display: inline-block; animation: pulse 3s ease-in-out infinite; }
@keyframes pulse { 0%,100% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.05); opacity: 0.85; } }
.empty-title { font-family: 'Instrument Serif', serif; font-style: italic; font-size: 2rem; color: #f5f5f7; margin-bottom: 0.6rem; }
.empty-text { color: #71717a; margin-top: 0.5rem; font-size: 0.95rem; margin-bottom: 1.5rem; }
p, li { color: #a1a1aa !important; line-height: 1.7; }
span, label { color: #d4d4d8 !important; }
h1, h2, h3, h4 { color: #f5f5f7 !important; }
strong, b { color: #f5f5f7 !important; }
hr { border: none !important; height: 1px !important; background: linear-gradient(90deg, transparent, rgba(168,85,247,0.3), transparent) !important; margin: 2rem 0 !important; }
.stSpinner > div { border-top-color: #a855f7 !important; }
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: #07070b; }
::-webkit-scrollbar-thumb { background: linear-gradient(180deg, #a855f7, #ec4899); border-radius: 100px; }
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { background: transparent !important; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  MODEL LOADERS (cached, load once per session)
# ══════════════════════════════════════════════
@st.cache_resource(show_spinner=False)
def load_whisper():
    return WhisperModel("base", device=DEVICE, compute_type=COMPUTE)


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
#  CORE FUNCTIONS
# ══════════════════════════════════════════════
def transcribe(audio_path):
    """Transcribe audio with auto language detection. Returns transcript, segments, language, confidence, duration."""
    last_error = None

    # Try 1 — standard transcription (most reliable)
    try:
        model = load_whisper()
        raw, info = model.transcribe(
            audio_path,
            beam_size=5,
            language=None
        )
        segments = []
        for s in raw:
            text = s.text.strip()
            if text:  # keep any non-empty segment
                segments.append({
                    "start": round(s.start, 2),
                    "end"  : round(s.end, 2),
                    "text" : text
                })
        if segments:
            transcript = " ".join(s["text"] for s in segments)
            return (
                transcript,
                segments,
                info.language,
                round(info.language_probability * 100, 1),
                info.duration
            )
        last_error = "No speech detected in audio"
    except Exception as e:
        last_error = str(e)[:200]

    # Try 2 — with smaller beam, no language hints (more lenient)
    try:
        model = load_whisper()
        raw, info = model.transcribe(
            audio_path,
            beam_size=1,
            language=None,
            condition_on_previous_text=False
        )
        segments = []
        for s in raw:
            text = s.text.strip()
            if text:
                segments.append({
                    "start": round(s.start, 2),
                    "end"  : round(s.end, 2),
                    "text" : text
                })
        if segments:
            transcript = " ".join(s["text"] for s in segments)
            return (
                transcript,
                segments,
                info.language,
                round(info.language_probability * 100, 1),
                info.duration
            )
        last_error = "Transcription returned no segments after retry"
    except Exception as e:
        last_error = str(e)[:200]

    # Both failed — show diagnostic
    st.error(
        f"❌ Transcription failed: **{last_error}**\n\n"
        f"**Common causes:**\n"
        f"- File has no audio track (try MP3 instead of MP4)\n"
        f"- Audio is silent or too quiet\n"
        f"- Background music with no speech\n\n"
        f"**Tip:** Convert your file to MP3 first using a free tool like https://cloudconvert.com"
    )
    return "", [], "en", 0.0, 0


def simple_diarize(segments, num_speakers=3):
    """Heuristic speaker assignment based on pause length."""
    out, spk, last_end = [], 1, 0.0
    for seg in segments:
        if seg["start"] - last_end >= 1.5:
            spk = (spk % num_speakers) + 1
        out.append({**seg, "speaker": f"Speaker {spk}"})
        last_end = seg["end"]
    return out


@torch.inference_mode()
def get_summary(text):
    """Generate summary using BART. Handles long texts via chunking."""
    if not text or len(text.split()) < 20:
        return text
    try:
        tok, mdl = load_summarizer()
        words = text.split()
        chunks = [
            " ".join(words[i:i+700])
            for i in range(0, len(words), 700)
        ]
        parts = []
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
        st.warning(f"Summary fallback: {str(e)[:100]}")
        return text[:500]


@torch.inference_mode()
def translate_batch(texts, src, tgt):
    """Translate a list of texts in one batch. Returns same-length list."""
    if not texts:
        return []
    if src == tgt or tgt not in MBART_LANG_MAP:
        return list(texts)
    if src not in MBART_LANG_MAP:
        # Source language not supported — return originals unchanged
        return list(texts)
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
        st.warning(f"Translation batch failed, using originals: {str(e)[:100]}")
        return list(texts)


def get_actions(segs):
    """Extract action items via keyword + zero-shot classification."""
    if not segs:
        return []
    try:
        clf = load_classifier()
        kws = ["will","should","need to","must","please","send",
               "schedule","review","ensure","prepare","follow up",
               "make sure","take care","handle","complete"]
        lbls = ["action item","task assignment","general discussion"]
        out = []
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
    """Extract decisions via keyword + zero-shot classification."""
    if not segs:
        return []
    try:
        clf = load_classifier()
        kws = ["decided","agreed","approved","confirmed","moving forward",
               "we will","finalized","going with","accepted","resolved",
               "concluded"]
        lbls = ["decision made","agreement reached","general statement"]
        out = []
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


# ══════════════════════════════════════════════
#  DOCX EXPORT
# ══════════════════════════════════════════════
def set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def add_banner(doc, title, color="1F497D"):
    t = doc.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    set_cell_bg(c, color)
    p = c.paragraphs[0]
    r = p.add_run("  " + title)
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    doc.add_paragraph()


def export_docx(summary, summary_tr, actions, decisions,
                diarized, src_lang, tgt_lang):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(1.0)
    sec.top_margin  = sec.bottom_margin = Inches(0.8)

    src_name = LANG_NAMES.get(src_lang, src_lang)
    tgt_name = LANG_NAMES.get(tgt_lang, tgt_lang)

    # Title
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

    # 1. Summary
    add_banner(doc, "1.  Executive Summary", "1F497D")
    doc.add_paragraph(summary if summary else "No summary available.")
    p2 = doc.add_paragraph(
        summary_tr if summary_tr else "Translation unavailable."
    )
    if p2.runs:
        p2.runs[0].italic = True
    doc.add_paragraph()

    # 2. Decisions
    add_banner(doc, "2.  Key Decisions", "375623")
    if decisions:
        for d in decisions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + d.get("speaker","Unknown") + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0x37, 0x56, 0x23)
            p.add_run(d.get("text",""))
            p.add_run("  (" + d.get("time","") + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(9)
            tt = tp.add_run(d.get("text_tr") or "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No decisions extracted.")
    doc.add_paragraph()

    # 3. Actions
    add_banner(doc, "3.  Action Items", "C00000")
    if actions:
        for a in actions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + a.get("speaker","Unknown") + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            p.add_run(a.get("text",""))
            p.add_run("  (" + a.get("time","") + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(9)
            tt = tp.add_run(a.get("text_tr") or "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    # 4. Transcript
    add_banner(
        doc,
        "4.  Full Transcript (" + src_name + " + " + tgt_name + ")",
        "595959"
    )
    for seg in diarized:
        p = doc.add_paragraph()
        r = p.add_run(
            "[" + str(seg.get("start",0)) + "s]  " +
            seg.get("speaker","Unknown") + ": "
        )
        r.bold = True
        r.font.size = Pt(9)
        p.add_run(seg.get("text","")).font.size = Pt(9)
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.4)
        tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(8)
        tt = tp.add_run(seg.get("text_tr") or "")
        tt.font.size = Pt(8)
        tt.italic = True

    safe_src = src_name.replace(" ", "_")
    safe_tgt = tgt_name.replace(" ", "_")
    path = f"/tmp/Meeting_Minutes_{safe_src}_to_{safe_tgt}.docx"
    doc.save(path)
    return path


# ══════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════

# Hero
st.markdown("""
<div class="hero-wrap">
    <span class="hero-eyebrow">⟡  AI MEETING INTELLIGENCE</span>
    <h1 class="hero-title">Minutes that <em>write themselves</em></h1>
    <p class="hero-subtitle">Upload any meeting audio in any language — get a polished,
    bilingual, speaker-attributed report in under a minute.</p>
</div>
""", unsafe_allow_html=True)

# Stats
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown("""<div class="stat-card"><span class="stat-icon">🌐</span><div class="stat-value">52</div><div class="stat-label">Languages</div></div>""", unsafe_allow_html=True)
with c2:
    st.markdown("""<div class="stat-card"><span class="stat-icon">🤖</span><div class="stat-value">4</div><div class="stat-label">AI Models</div></div>""", unsafe_allow_html=True)
with c3:
    st.markdown("""<div class="stat-card"><span class="stat-icon">⚡</span><div class="stat-value">~60s</div><div class="stat-label">Per File</div></div>""", unsafe_allow_html=True)
with c4:
    st.markdown("""<div class="stat-card"><span class="stat-icon">📄</span><div class="stat-value">DOCX</div><div class="stat-label">Export</div></div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Sidebar
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
    st.caption(f"Max file size: {MAX_FILE_MB} MB · Max duration: {MAX_DURATION_S//60} min")
    st.markdown("---")
    st.markdown("### 🤖 Models Used")
    st.markdown("🎙️ **faster-whisper**"); st.caption("Speech to Text")
    st.markdown("📝 **BART-large-CNN**");  st.caption("Summarization")
    st.markdown("🏷️ **BART-large-MNLI**"); st.caption("Classification")
    st.markdown("🌐 **mBART-50**");         st.caption("Translation · 52 languages")
    st.markdown("---")
    st.markdown("### 🔬 Pipeline")
    for i, step in enumerate([
        "Audio Upload","Transcription","Diarization",
        "NLP Extraction","Translation","DOCX Export"
    ], 1):
        st.markdown(f"`{i}` {step}")

# Upload
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
    # ── Validate file size ────────────────────
    file_bytes = uploaded.getvalue()
    file_mb    = len(file_bytes) / (1024 * 1024)

    if file_mb > MAX_FILE_MB:
        st.error(
            f"⚠️ File too large: {file_mb:.1f} MB. "
            f"Maximum allowed is {MAX_FILE_MB} MB. "
            f"Try compressing your audio first."
        )
        st.stop()

    # Save to /tmp safely
    safe_name  = "".join(
        c if c.isalnum() or c in ".-_" else "_"
        for c in uploaded.name
    )
    audio_path = f"/tmp/uploaded_{safe_name}"
    with open(audio_path, "wb") as f:
        f.write(file_bytes)

    st.audio(audio_path)

    # File info pills
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            f'<span class="pill pill-purple">📄 {uploaded.name}</span>',
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f'<span class="pill pill-pink">💾 {file_mb:.2f} MB</span>',
            unsafe_allow_html=True
        )
    with col3:
        st.markdown(
            f'<span class="pill pill-green">🌐 → {LANG_NAMES.get(target_lang, target_lang)}</span>',
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("✨  Generate Meeting Minutes"):
        progress = st.progress(0)
        status   = st.empty()
        t_start  = time.time()

        try:
            # ── 1. Transcribe ────────────────────
            status.markdown("🎙️ **Transcribing audio...**")
            transcript, segments, det_lang, det_prob, duration = transcribe(audio_path)

            if not segments:
                # Error already shown by transcribe()
                progress.empty()
                status.empty()
                st.stop()

            if duration > MAX_DURATION_S:
                st.warning(
                    f"⚠️ Audio is {duration/60:.1f} min long. "
                    f"Recommended max is {MAX_DURATION_S//60} min. "
                    f"Processing may be slow."
                )

            progress.progress(20)
            st.markdown(
                f'<span class="pill pill-amber">⟡ Detected: {LANG_NAMES.get(det_lang, det_lang)} · {det_prob}% · {duration:.0f}s</span>',
                unsafe_allow_html=True
            )

            # Warn if Whisper detected a language mBART can't translate
            if det_lang not in MBART_LANG_MAP and det_lang != target_lang:
                st.warning(
                    f"⚠️ {LANG_NAMES.get(det_lang, det_lang)} not supported for translation. "
                    f"Output will use original text."
                )

            # ── 2. Diarize ───────────────────────
            status.markdown("👥 **Assigning speakers...**")
            diarized = simple_diarize(segments, num_speakers)
            progress.progress(30)

            # ── 3. Extract ───────────────────────
            status.markdown("🧠 **Extracting summary...**")
            summary = get_summary(transcript)
            progress.progress(45)

            status.markdown("✅ **Extracting action items & decisions...**")
            actions   = get_actions(diarized)
            decisions = get_decisions(diarized)
            progress.progress(60)

            # ── 4. Translate ─────────────────────
            tgt_name  = LANG_NAMES.get(target_lang, target_lang)
            src_name  = LANG_NAMES.get(det_lang, det_lang)
            same_lang = (det_lang == target_lang)

            if same_lang:
                status.markdown("⚡ **Same language — skipping translation**")
                summary_tr = summary
                for a in actions:
                    a["text_tr"] = a["text"]
                for d in decisions:
                    d["text_tr"] = d["text"]
                for s in diarized:
                    s["text_tr"] = s["text"]
                progress.progress(95)

            else:
                status.markdown(
                    f"🌐 **Translating {src_name} → {tgt_name}** "
                    f"(this may take 1–3 minutes for long meetings)"
                )

                # Translate summary
                summary_tr = ""
                if summary:
                    sum_result = translate_batch(
                        [summary], det_lang, target_lang
                    )
                    summary_tr = sum_result[0] if sum_result else ""
                progress.progress(65)

                # Translate actions in batch
                if actions:
                    action_translated = translate_batch(
                        [a["text"] for a in actions],
                        det_lang, target_lang
                    )
                    for a, t in zip(actions, action_translated):
                        a["text_tr"] = t
                progress.progress(70)

                # Translate decisions in batch
                if decisions:
                    decision_translated = translate_batch(
                        [d["text"] for d in decisions],
                        det_lang, target_lang
                    )
                    for d, t in zip(decisions, decision_translated):
                        d["text_tr"] = t
                progress.progress(75)

                # Translate full transcript in batches
                transcript_texts = [s["text"] for s in diarized]
                n_total          = len(transcript_texts)
                translated_segs  = []

                for batch_idx in range(0, n_total, TRANSLATE_BATCH):
                    chunk = transcript_texts[batch_idx:batch_idx+TRANSLATE_BATCH]
                    chunk_result = translate_batch(
                        chunk, det_lang, target_lang
                    )
                    translated_segs.extend(chunk_result)

                    # Live progress update between 75 and 92
                    completed = min(batch_idx + TRANSLATE_BATCH, n_total)
                    pct = 75 + int(completed / n_total * 17)
                    progress.progress(min(92, pct))
                    status.markdown(
                        f"🌐 **Translating {src_name} → {tgt_name}** · "
                        f"segment {completed} / {n_total}"
                    )

                # Apply to diarized segments
                for s, t in zip(diarized, translated_segs):
                    s["text_tr"] = t

                # Fill any missing translations with originals
                for s in diarized:
                    if "text_tr" not in s or not s.get("text_tr"):
                        s["text_tr"] = s["text"]

                progress.progress(95)

            # ── 5. Export DOCX ───────────────────
            status.markdown("📄 **Generating DOCX report...**")
            docx_path = export_docx(
                summary, summary_tr,
                actions, decisions, diarized,
                det_lang, target_lang
            )
            progress.progress(100)

            elapsed = time.time() - t_start
            status.empty()
            st.success(f"✅ Processing complete in {elapsed:.1f} seconds!")

            # Cleanup memory
            gc.collect()
            if DEVICE == "cuda":
                torch.cuda.empty_cache()

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Results overview ─────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">02</span>Results overview</div>',
                unsafe_allow_html=True
            )
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown(f"""<div class="result-box"><div class="result-value">{len(actions)}</div><div class="result-label">Action Items</div></div>""", unsafe_allow_html=True)
            with m2:
                st.markdown(f"""<div class="result-box"><div class="result-value">{len(decisions)}</div><div class="result-label">Key Decisions</div></div>""", unsafe_allow_html=True)
            with m3:
                st.markdown(f"""<div class="result-box"><div class="result-value">{len(diarized)}</div><div class="result-label">Segments</div></div>""", unsafe_allow_html=True)
            with m4:
                st.markdown(f"""<div class="result-box"><div class="result-value">{len(transcript.split())}</div><div class="result-label">Words</div></div>""", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Summary ──────────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">03</span>Executive summary</div>',
                unsafe_allow_html=True
            )
            tab1, tab2 = st.tabs([
                "🗣️ Original (" + src_name + ")",
                "🌐 Translated (" + tgt_name + ")"
            ])
            with tab1:
                st.markdown(
                    '<div class="glass-card">' +
                    (summary or "No summary extracted.") +
                    '</div>',
                    unsafe_allow_html=True
                )
            with tab2:
                st.markdown(
                    '<div class="glass-card">' +
                    (summary_tr or "Translation unavailable.") +
                    '</div>',
                    unsafe_allow_html=True
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Decisions ────────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">04</span>Key decisions</div>',
                unsafe_allow_html=True
            )
            if decisions:
                for i, d in enumerate(decisions, 1):
                    with st.expander(
                        "Decision " + str(i) + " — " +
                        d.get("speaker","Unknown") + " (" + d.get("time","") + ")"
                    ):
                        st.markdown("**🗣️ Original:** " + d.get("text",""))
                        st.markdown(
                            "**🌐 " + tgt_name + ":** " +
                            (d.get("text_tr") or d.get("text",""))
                        )
            else:
                st.info("No decisions extracted from this meeting.")

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Actions ──────────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">05</span>Action items</div>',
                unsafe_allow_html=True
            )
            if actions:
                for i, a in enumerate(actions, 1):
                    with st.expander(
                        "Action " + str(i) + " — " +
                        a.get("speaker","Unknown") + " (" + a.get("time","") + ")"
                    ):
                        st.markdown("**🗣️ Original:** " + a.get("text",""))
                        st.markdown(
                            "**🌐 " + tgt_name + ":** " +
                            (a.get("text_tr") or a.get("text",""))
                        )
            else:
                st.info("No action items extracted from this meeting.")

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Transcript ───────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">06</span>Full transcript</div>',
                unsafe_allow_html=True
            )
            with st.expander(
                "View full bilingual transcript (" +
                str(len(diarized)) + " segments)"
            ):
                for seg in diarized:
                    st.markdown(
                        "**[" + str(seg.get("start",0)) + "s] " +
                        seg.get("speaker","Unknown") + ":** " +
                        seg.get("text","") +
                        "  \n*🌐 " + (seg.get("text_tr") or seg.get("text","")) + "*"
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── Download ─────────────────────────
            st.markdown(
                '<div class="sec-header"><span class="sec-number">07</span>Download report</div>',
                unsafe_allow_html=True
            )
            with open(docx_path, "rb") as f:
                st.download_button(
                    label     = "⤓  Download Bilingual Meeting Minutes (DOCX)",
                    data      = f,
                    file_name = f"Meeting_Minutes_{src_name}_to_{tgt_name}.docx",
                    mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            progress.empty()
            status.empty()
            st.error(
                f"❌ Something went wrong: {str(e)[:200]}\n\n"
                f"Try a smaller file or refresh the page."
            )

else:
    st.markdown("""
    <div class="empty-state">
        <div class="empty-icon">🎙️</div>
        <div class="empty-title">Drop your audio<br>and watch the magic</div>
        <div class="empty-text">Supports .mp3 · .wav · .m4a · .mp4</div>
        <div style='margin-top:1.5rem;'>
            <span class='pill pill-purple'>⟡ Auto language detection</span>
            <span class='pill pill-pink'>⟡ 52 languages</span>
            <span class='pill pill-green'>⟡ DOCX export</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
