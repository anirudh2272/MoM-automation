import os
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

DEVICE    = "cuda" if torch.cuda.is_available() else "cpu"
COMPUTE   = "float16" if DEVICE == "cuda" else "int8"
HF_DEVICE = 0 if DEVICE == "cuda" else -1

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
#  PAGE CONFIG & CSS
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
#  MODEL LOADERS
# ══════════════════════════════════════════════
@st.cache_resource
def load_whisper():
    return WhisperModel("base", device=DEVICE, compute_type=COMPUTE)

@st.cache_resource
def load_summarizer():
    tok = BartTokenizer.from_pretrained("facebook/bart-large-cnn")
    mdl = BartForConditionalGeneration.from_pretrained(
              "facebook/bart-large-cnn").to(DEVICE)
    return tok, mdl

@st.cache_resource
def load_classifier():
    return pipeline("zero-shot-classification",
                    model="facebook/bart-large-mnli",
                    device=HF_DEVICE)

@st.cache_resource
def load_translator():
    tok = MBart50TokenizerFast.from_pretrained(
              "facebook/mbart-large-50-many-to-many-mmt")
    mdl = MBartForConditionalGeneration.from_pretrained(
              "facebook/mbart-large-50-many-to-many-mmt").to(DEVICE)
    return tok, mdl


# ══════════════════════════════════════════════
#  CORE FUNCTIONS
# ══════════════════════════════════════════════
def transcribe(audio_path):
    model = load_whisper()
    raw, info = model.transcribe(audio_path, beam_size=5, language=None)
    segments = [{
        "start": round(s.start, 2),
        "end"  : round(s.end,   2),
        "text" : s.text.strip()
    } for s in raw]
    transcript = " ".join(s["text"] for s in segments)
    return transcript, segments, info.language, round(info.language_probability * 100, 1)


def simple_diarize(segments, num_speakers=3):
    out, spk, last_end = [], 1, 0.0
    for seg in segments:
        if seg["start"] - last_end >= 1.5:
            spk = (spk % num_speakers) + 1
        out.append({**seg, "speaker": f"Speaker {spk}"})
        last_end = seg["end"]
    return out


def get_summary(text):
    if not text or len(text.split()) < 20:
        return text or ""
    tok, mdl = load_summarizer()
    words  = text.split()
    chunks = [" ".join(words[i:i+700]) for i in range(0, len(words), 700)]
    parts  = []
    for c in chunks:
        if len(c.split()) > 30:
            inp = tok(c, return_tensors="pt", max_length=1024, truncation=True)
            inp = {k: v.to(DEVICE) for k, v in inp.items()}
            ids = mdl.generate(inp["input_ids"], max_length=130,
                               min_length=30, num_beams=4)
            parts.append(tok.decode(ids[0], skip_special_tokens=True))
    return " ".join(parts) if parts else text[:500]


def get_actions(segs):
    clf  = load_classifier()
    kws  = ["will","should","need to","must","please","send",
            "schedule","review","ensure","prepare","follow up"]
    lbls = ["action item","task assignment","general discussion"]
    out  = []
    for s in segs:
        if any(k in s["text"].lower() for k in kws):
            r = clf(s["text"], candidate_labels=lbls)
            if r["labels"][0] != "general discussion" and r["scores"][0] > 0.40:
                out.append({
                    "speaker": s["speaker"],
                    "text"   : s["text"],
                    "time"   : str(s["start"]) + "s"
                })
    return out


def get_decisions(segs):
    clf  = load_classifier()
    kws  = ["decided","agreed","approved","confirmed","moving forward",
            "we will","finalized","going with","accepted"]
    lbls = ["decision made","agreement reached","general statement"]
    out  = []
    for s in segs:
        if any(k in s["text"].lower() for k in kws):
            r = clf(s["text"], candidate_labels=lbls)
            if r["labels"][0] != "general statement" and r["scores"][0] > 0.38:
                out.append({
                    "speaker": s["speaker"],
                    "text"   : s["text"],
                    "time"   : str(s["start"]) + "s"
                })
    return out


def translate_text(text, src, tgt):
    if not text or src == tgt or tgt not in MBART_LANG_MAP:
        return text or ""
    try:
        tok, mdl = load_translator()
        tok.src_lang = MBART_LANG_MAP.get(src, "en_XX")
        inp = tok(text, return_tensors="pt", max_length=512,
                  truncation=True, padding=True)
        inp = {k: v.to(DEVICE) for k, v in inp.items()}
        tgt_id = tok.lang_code_to_id[MBART_LANG_MAP[tgt]]
        out = mdl.generate(**inp, forced_bos_token_id=tgt_id,
                           max_length=512, num_beams=4)
        return tok.decode(out[0], skip_special_tokens=True)
    except Exception:
        return text or ""


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

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("MINUTES OF MEETING")
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    s = doc.add_paragraph()
    s.alignment = WD_ALIGN_PARAGRAPH.CENTER
    s.add_run("Original: " + src_name + "   |   Translated: " + tgt_name).italic = True
    doc.add_paragraph(datetime.now().strftime("%B %d, %Y  |  %I:%M %p"))
    doc.add_paragraph()

    add_banner(doc, "1.  Executive Summary", "1F497D")
    doc.add_paragraph(summary if summary else "No summary available.")
    p2 = doc.add_paragraph(summary_tr if summary_tr else "Translation unavailable.")
    if p2.runs:
        p2.runs[0].italic = True
    doc.add_paragraph()

    add_banner(doc, "2.  Key Decisions", "375623")
    if decisions:
        for d in decisions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + d["speaker"] + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0x37, 0x56, 0x23)
            p.add_run(d["text"])
            p.add_run("  (" + d["time"] + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(9)
            tt = tp.add_run(d.get("text_tr") or "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No decisions extracted.")
    doc.add_paragraph()

    add_banner(doc, "3.  Action Items", "C00000")
    if actions:
        for a in actions:
            p = doc.add_paragraph(style="List Number")
            r = p.add_run("[" + a["speaker"] + "]  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
            p.add_run(a["text"])
            p.add_run("  (" + a["time"] + ")").italic = True
            tp = doc.add_paragraph()
            tp.paragraph_format.left_indent = Inches(0.4)
            tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(9)
            tt = tp.add_run(a.get("text_tr") or "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    add_banner(doc, "4.  Full Transcript (" + src_name + " + " + tgt_name + ")", "595959")
    for seg in diarized:
        p = doc.add_paragraph()
        r = p.add_run("[" + str(seg["start"]) + "s]  " + seg["speaker"] + ": ")
        r.bold = True
        r.font.size = Pt(9)
        p.add_run(seg["text"]).font.size = Pt(9)
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.4)
        tp.add_run("  🌐 " + tgt_name + ": ").font.size = Pt(8)
        tt = tp.add_run(seg.get("text_tr") or "")
        tt.font.size = Pt(8)
        tt.italic = True

    path = "/tmp/Meeting_Minutes_" + src_name + "_to_" + tgt_name + ".docx"
    doc.save(path)
    return path


# ══════════════════════════════════════════════
#  UI
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
with c1:
    st.markdown("""<div class="stat-card"><span class="stat-icon">🌐</span><div class="stat-value">52</div><div class="stat-label">Languages</div></div>""", unsafe_allow_html=True)
with c2:
    st.markdown("""<div class="stat-card"><span class="stat-icon">🤖</span><div class="stat-value">4</div><div class="stat-label">AI Models</div></div>""", unsafe_allow_html=True)
with c3:
    st.markdown("""<div class="stat-card"><span class="stat-icon">⚡</span><div class="stat-value">~60s</div><div class="stat-label">Per File</div></div>""", unsafe_allow_html=True)
with c4:
    st.markdown("""<div class="stat-card"><span class="stat-icon">📄</span><div class="stat-value">DOCX</div><div class="stat-label">Export</div></div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

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
    st.markdown("---")
    st.markdown("### 🤖 Models")
    st.markdown("🎙️ **faster-whisper**"); st.caption("Speech to Text")
    st.markdown("📝 **BART-large-CNN**");  st.caption("Summarization")
    st.markdown("🏷️ **BART-large-MNLI**"); st.caption("Classification")
    st.markdown("🌐 **mBART-50**");         st.caption("Translation · 52 languages")

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
    audio_path = "/tmp/uploaded_" + uploaded.name
    with open(audio_path, "wb") as f:
        f.write(uploaded.read())

    st.audio(audio_path)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f'<span class="pill pill-purple">📄 {uploaded.name}</span>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<span class="pill pill-pink">💾 {round(os.path.getsize(audio_path)/1024, 1)} KB</span>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<span class="pill pill-green">🌐 → {LANG_NAMES.get(target_lang, target_lang)}</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("✨  Generate Meeting Minutes"):

        with st.spinner("🎙️ Transcribing audio..."):
            transcript, segments, det_lang, det_prob = transcribe(audio_path)

        st.markdown(
            f'<span class="pill pill-amber">⟡ Detected: {LANG_NAMES.get(det_lang, det_lang)} · {det_prob}%</span>',
            unsafe_allow_html=True
        )

        with st.spinner("👥 Assigning speakers..."):
            diarized = simple_diarize(segments, num_speakers)

        with st.spinner("🧠 Extracting summary, actions, decisions..."):
            summary   = get_summary(transcript)
            actions   = get_actions(diarized)
            decisions = get_decisions(diarized)

        tgt_name = LANG_NAMES.get(target_lang, target_lang)
        src_name = LANG_NAMES.get(det_lang, det_lang)

        with st.spinner("🌐 Translating to " + tgt_name + "..."):
            summary_tr = translate_text(summary, det_lang, target_lang)
            for a in actions:
                a["text_tr"] = translate_text(a["text"], det_lang, target_lang)
            for d in decisions:
                d["text_tr"] = translate_text(d["text"], det_lang, target_lang)
            for i, s in enumerate(diarized):
                if i < 15:
                    s["text_tr"] = translate_text(s["text"], det_lang, target_lang)
                else:
                    s["text_tr"] = "..."

        st.success("✅ Processing complete!")
        st.markdown("<br>", unsafe_allow_html=True)

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

        st.markdown(
            '<div class="sec-header"><span class="sec-number">03</span>Executive summary</div>',
            unsafe_allow_html=True
        )
        tab1, tab2 = st.tabs([
            "🗣️ Original (" + src_name + ")",
            "🌐 Translated (" + tgt_name + ")"
        ])
        with tab1:
            st.markdown('<div class="glass-card">' + (summary or "No summary extracted.") + '</div>', unsafe_allow_html=True)
        with tab2:
            st.markdown('<div class="glass-card">' + (summary_tr or "Translation unavailable.") + '</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sec-header"><span class="sec-number">04</span>Key decisions</div>',
            unsafe_allow_html=True
        )
        if decisions:
            for i, d in enumerate(decisions, 1):
                with st.expander("Decision " + str(i) + " — " + d["speaker"] + " (" + d["time"] + ")"):
                    st.markdown("**🗣️ Original:** " + d["text"])
                    st.markdown("**🌐 " + tgt_name + ":** " + (d.get("text_tr") or d["text"]))
        else:
            st.info("No decisions extracted from this meeting.")

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sec-header"><span class="sec-number">05</span>Action items</div>',
            unsafe_allow_html=True
        )
        if actions:
            for i, a in enumerate(actions, 1):
                with st.expander("Action " + str(i) + " — " + a["speaker"] + " (" + a["time"] + ")"):
                    st.markdown("**🗣️ Original:** " + a["text"])
                    st.markdown("**🌐 " + tgt_name + ":** " + (a.get("text_tr") or a["text"]))
        else:
            st.info("No action items extracted from this meeting.")

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sec-header"><span class="sec-number">06</span>Full transcript</div>',
            unsafe_allow_html=True
        )
        with st.expander("View full bilingual transcript (" + str(len(diarized)) + " segments)"):
            for seg in diarized:
                st.markdown(
                    "**[" + str(seg["start"]) + "s] " + seg["speaker"] + ":** " +
                    seg["text"] + "  \n*🌐 " + (seg.get("text_tr") or "") + "*"
                )

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            '<div class="sec-header"><span class="sec-number">07</span>Download report</div>',
            unsafe_allow_html=True
        )
        docx_path = export_docx(summary, summary_tr, actions, decisions, diarized, det_lang, target_lang)
        with open(docx_path, "rb") as f:
            st.download_button(
                label     = "⤓  Download Bilingual Meeting Minutes (DOCX)",
                data      = f,
                file_name = "Meeting_Minutes_" + src_name + "_to_" + tgt_name + ".docx",
                mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
