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
    "en":"English","fr":"French","de":"German","hi":"Hindi",
    "ar":"Arabic","es":"Spanish","zh":"Chinese","ja":"Japanese",
    "ru":"Russian","te":"Telugu","ta":"Tamil","ko":"Korean",
    "pt":"Portuguese","it":"Italian","nl":"Dutch","tr":"Turkish",
    "uk":"Ukrainian","pl":"Polish","sv":"Swedish","fi":"Finnish"
}

MBART_LANG_MAP = {
    "ar":"ar_AR","de":"de_DE","en":"en_XX","es":"es_XX",
    "fi":"fi_FI","fr":"fr_XX","hi":"hi_IN","id":"id_ID",
    "it":"it_IT","ja":"ja_XX","ko":"ko_KR","nl":"nl_XX",
    "pl":"pl_PL","pt":"pt_XX","ro":"ro_RO","ru":"ru_RU",
    "sv":"sv_SE","ta":"ta_IN","te":"te_IN","th":"th_TH",
    "tr":"tr_TR","uk":"uk_UA","vi":"vi_VN","zh":"zh_CN"
}

# ══════════════════════════════════════════════
#  PAGE CONFIG & CUSTOM CSS
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="MoM Automation",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Background */
.stApp {
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    min-height: 100vh;
}

/* Main title */
.main-title {
    font-family: 'Syne', sans-serif;
    font-size: 3rem;
    font-weight: 800;
    background: linear-gradient(90deg, #f093fb, #f5576c, #fda085);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    text-align: center;
    margin-bottom: 0.2rem;
    letter-spacing: -1px;
}

.sub-title {
    text-align: center;
    color: rgba(255,255,255,0.5);
    font-size: 1rem;
    margin-bottom: 2rem;
    font-weight: 300;
}

/* Cards */
.glass-card {
    background: rgba(255,255,255,0.05);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 1rem;
}

/* Section headers */
.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: white;
    padding: 0.6rem 1rem;
    border-radius: 8px;
    margin-bottom: 1rem;
    display: inline-block;
}

.header-purple { background: linear-gradient(90deg, #7c3aed, #a855f7); }
.header-pink   { background: linear-gradient(90deg, #db2777, #f472b6); }
.header-blue   { background: linear-gradient(90deg, #1d4ed8, #3b82f6); }
.header-green  { background: linear-gradient(90deg, #059669, #34d399); }
.header-orange { background: linear-gradient(90deg, #d97706, #fbbf24); }
.header-red    { background: linear-gradient(90deg, #dc2626, #f87171); }

/* Metric boxes */
.metric-box {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.15);
    border-radius: 12px;
    padding: 1rem;
    text-align: center;
    transition: transform 0.2s;
}
.metric-box:hover { transform: translateY(-2px); }
.metric-value {
    font-family: 'Syne', sans-serif;
    font-size: 1.8rem;
    font-weight: 800;
    color: white;
}
.metric-label {
    font-size: 0.75rem;
    color: rgba(255,255,255,0.5);
    text-transform: uppercase;
    letter-spacing: 1px;
}

/* Step items */
.step-correct {
    background: rgba(52, 211, 153, 0.1);
    border-left: 3px solid #34d399;
    padding: 0.6rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.5rem;
    color: rgba(255,255,255,0.9);
    font-size: 0.9rem;
}
.step-wrong {
    background: rgba(248, 113, 113, 0.1);
    border-left: 3px solid #f87171;
    padding: 0.6rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.5rem;
    color: rgba(255,255,255,0.9);
    font-size: 0.9rem;
}
.step-unknown {
    background: rgba(156, 163, 175, 0.1);
    border-left: 3px solid #9ca3af;
    padding: 0.6rem 1rem;
    border-radius: 0 8px 8px 0;
    margin-bottom: 0.5rem;
    color: rgba(255,255,255,0.9);
    font-size: 0.9rem;
}

/* Perturbation pills */
.pert-survived {
    display: inline-block;
    background: rgba(52, 211, 153, 0.15);
    border: 1px solid #34d399;
    color: #34d399;
    padding: 0.3rem 0.8rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 500;
    margin: 0.2rem;
}
.pert-broke {
    display: inline-block;
    background: rgba(248, 113, 113, 0.15);
    border: 1px solid #f87171;
    color: #f87171;
    padding: 0.3rem 0.8rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 500;
    margin: 0.2rem;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: rgba(15, 12, 41, 0.95);
    border-right: 1px solid rgba(255,255,255,0.08);
}
section[data-testid="stSidebar"] * {
    color: rgba(255,255,255,0.85) !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(90deg, #f093fb, #f5576c) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 0.6rem 2rem !important;
    width: 100% !important;
    transition: opacity 0.2s !important;
}
.stButton > button:hover { opacity: 0.85 !important; }

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(90deg, #059669, #34d399) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    width: 100% !important;
}

/* Text inputs */
.stTextArea textarea, .stSelectbox select {
    background: rgba(255,255,255,0.05) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    border-radius: 10px !important;
    color: white !important;
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(255,255,255,0.05);
    border-radius: 10px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    color: rgba(255,255,255,0.6) !important;
    border-radius: 8px;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(90deg, #7c3aed, #a855f7) !important;
    color: white !important;
}

/* Expander */
.streamlit-expanderHeader {
    background: rgba(255,255,255,0.05) !important;
    border-radius: 8px !important;
    color: white !important;
}

/* Info/success boxes */
.stAlert {
    border-radius: 10px !important;
    border: none !important;
}

/* Divider */
hr { border-color: rgba(255,255,255,0.1) !important; }

/* All text white */
p, li, span, label { color: rgba(255,255,255,0.85) !important; }
h1, h2, h3, h4 { color: white !important; }

/* Spinner */
.stSpinner > div { border-top-color: #f093fb !important; }

/* Badge */
.badge {
    display: inline-block;
    padding: 0.2rem 0.6rem;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
}
.badge-purple { background: rgba(124,58,237,0.3); color: #a78bfa; border: 1px solid #7c3aed; }
.badge-pink   { background: rgba(219,39,119,0.3); color: #f9a8d4; border: 1px solid #db2777; }
.badge-green  { background: rgba(5,150,105,0.3);  color: #6ee7b7; border: 1px solid #059669; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  CACHED MODEL LOADERS
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
#  CORE PIPELINE FUNCTIONS
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
    return " ".join(parts) if parts else ""

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

    # Title
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

    # Summary
    add_banner(doc, "1.  Executive Summary", "1F497D")
    doc.add_paragraph(summary if summary else "No summary available.")
    # Safe italic paragraph
    translated_summary = summary_tr if summary_tr else "Translation unavailable."
    p2 = doc.add_paragraph(translated_summary)
    if p2.runs:
        p2.runs[0].italic = True
    doc.add_paragraph()

    # Decisions
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
            tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_run.font.size = Pt(9)
            translated = d.get("text_tr", "")
            tt = tp.add_run(translated if translated else "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No decisions extracted.")
    doc.add_paragraph()

    # Actions
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
            tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_run.font.size = Pt(9)
            translated = a.get("text_tr", "")
            tt = tp.add_run(translated if translated else "")
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    # Transcript
    add_banner(doc, "4.  Full Transcript (" + src_name + " + " + tgt_name + ")", "595959")
    for seg in diarized:
        p = doc.add_paragraph()
        r = p.add_run("[" + str(seg["start"]) + "s]  " + seg["speaker"] + ": ")
        r.bold = True
        r.font.size = Pt(9)
        p.add_run(seg["text"]).font.size = Pt(9)
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.4)
        tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
        tr_run.font.size = Pt(8)
        translated = seg.get("text_tr", "")
        tt = tp.add_run(translated if translated else "")
        tt.font.size = Pt(8)
        tt.italic = True

    path = "Meeting_Minutes_" + src_name + "_to_" + tgt_name + ".docx"
    doc.save(path)
    return path


# ══════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════

# Header
st.markdown('<h1 class="main-title">🎙️ MoM Automation</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Upload any meeting audio → Get bilingual AI-powered minutes in seconds</p>', unsafe_allow_html=True)

# Stats row
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown("""<div class="metric-box">
        <div class="metric-value">50+</div>
        <div class="metric-label">Languages</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown("""<div class="metric-box">
        <div class="metric-value">4</div>
        <div class="metric-label">AI Models</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown("""<div class="metric-box">
        <div class="metric-value">100%</div>
        <div class="metric-label">Free</div>
    </div>""", unsafe_allow_html=True)
with c4:
    st.markdown("""<div class="metric-box">
        <div class="metric-value">DOCX</div>
        <div class="metric-label">Export</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")

    target_lang = st.selectbox(
        "🌐 Translate output to:",
        options=list(LANG_NAMES.keys()),
        format_func=lambda x: LANG_NAMES[x] + " (" + x + ")",
        index=list(LANG_NAMES.keys()).index("hi")
    )

    num_speakers = st.slider(
        "👥 Number of speakers",
        min_value=1, max_value=6, value=3
    )

    st.markdown("---")
    st.markdown("### 📁 Supported Formats")
    st.markdown("`mp3`  `wav`  `m4a`  `mp4`")

    st.markdown("---")
    st.markdown("### 🤖 Models Used")
    models = [
        ("🎙️", "faster-whisper", "Speech to Text"),
        ("📝", "BART-large-CNN", "Summarization"),
        ("🏷️", "BART-large-MNLI", "Classification"),
        ("🌐", "mBART-50", "Translation"),
    ]
    for icon, name, role in models:
        st.markdown(f"{icon} **{name}**")
        st.caption(role)

    st.markdown("---")
    st.markdown("### 🔬 Pipeline")
    steps = ["Audio Upload", "Transcription", "Diarization",
             "NLP Extraction", "Translation", "DOCX Export"]
    for i, step in enumerate(steps, 1):
        st.markdown(f"`{i}` {step}")

# Main upload area
st.markdown('<div class="section-header header-purple">📁 Upload Meeting Audio</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drag and drop your audio file here",
    type=["mp3", "wav", "m4a", "mp4"],
    label_visibility="collapsed"
)

if uploaded:
    audio_path = "uploaded_" + uploaded.name
    with open(audio_path, "wb") as f:
        f.write(uploaded.read())

    st.audio(uploaded)

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.markdown(f'<span class="badge badge-purple">📄 {uploaded.name}</span>', unsafe_allow_html=True)
    with col_info2:
        st.markdown(f'<span class="badge badge-pink">💾 {round(os.path.getsize(audio_path)/1024, 1)} KB</span>', unsafe_allow_html=True)
    with col_info3:
        st.markdown(f'<span class="badge badge-green">🌐 → {LANG_NAMES.get(target_lang, target_lang)}</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Generate Meeting Minutes"):

        progress = st.progress(0)
        status   = st.empty()

        # Step 1: Transcribe
        status.markdown("🎙️ **Transcribing audio...**")
        transcript, segments, det_lang, det_prob = transcribe(audio_path)
        progress.progress(20)

        st.markdown(
            f'<span class="badge badge-green">✅ Detected: {LANG_NAMES.get(det_lang, det_lang)} ({det_prob}%)</span>',
            unsafe_allow_html=True
        )

        # Step 2: Diarize
        status.markdown("👥 **Assigning speakers...**")
        diarized = simple_diarize(segments, num_speakers)
        progress.progress(35)

        # Step 3: Extract
        status.markdown("🧠 **Extracting summary, actions, decisions...**")
        summary   = get_summary(transcript)
        actions   = get_actions(diarized)
        decisions = get_decisions(diarized)
        progress.progress(60)

        # Step 4: Translate
        tgt_name = LANG_NAMES.get(target_lang, target_lang)
        status.markdown(f"🌐 **Translating to {tgt_name}...**")
        summary_tr = translate_text(summary, det_lang, target_lang)
        for a in actions:
            a["text_tr"] = translate_text(a["text"], det_lang, target_lang)
        for d in decisions:
            d["text_tr"] = translate_text(d["text"], det_lang, target_lang)
        for s in diarized:
            s["text_tr"] = translate_text(s["text"], det_lang, target_lang)
        progress.progress(85)

        # Step 5: Export
        status.markdown("📄 **Generating DOCX...**")
        docx_path = export_docx(
            summary, summary_tr, actions, decisions,
            diarized, det_lang, target_lang
        )
        progress.progress(100)
        status.empty()

        st.success("✅ Processing complete!")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Results metrics ───────────────────
        st.markdown('<div class="section-header header-pink">📊 Results Overview</div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-value">{len(actions)}</div>
                <div class="metric-label">Action Items</div>
            </div>""", unsafe_allow_html=True)
        with m2:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-value">{len(decisions)}</div>
                <div class="metric-label">Key Decisions</div>
            </div>""", unsafe_allow_html=True)
        with m3:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-value">{len(diarized)}</div>
                <div class="metric-label">Segments</div>
            </div>""", unsafe_allow_html=True)
        with m4:
            st.markdown(f"""<div class="metric-box">
                <div class="metric-value">{len(transcript.split())}</div>
                <div class="metric-label">Words</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Summary ───────────────────────────
        st.markdown('<div class="section-header header-blue">📝 Executive Summary</div>', unsafe_allow_html=True)
        src_name = LANG_NAMES.get(det_lang, det_lang)
        tab1, tab2 = st.tabs([
            "🗣️ Original (" + src_name + ")",
            "🌐 Translated (" + tgt_name + ")"
        ])
        with tab1:
            st.markdown('<div class="glass-card">' + (summary or "No summary extracted.") + '</div>', unsafe_allow_html=True)
        with tab2:
            st.markdown('<div class="glass-card">' + (summary_tr or "Translation unavailable.") + '</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Decisions ─────────────────────────
        st.markdown('<div class="section-header header-green">🏛️ Key Decisions</div>', unsafe_allow_html=True)
        if decisions:
            for i, d in enumerate(decisions, 1):
                with st.expander("Decision " + str(i) + " — " + d["speaker"] + " (" + d["time"] + ")"):
                    st.markdown("**🗣️ Original:** " + d["text"])
                    st.markdown("**🌐 " + tgt_name + ":** " + d.get("text_tr", ""))
        else:
            st.info("No decisions extracted from this meeting.")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Actions ───────────────────────────
        st.markdown('<div class="section-header header-red">✅ Action Items</div>', unsafe_allow_html=True)
        if actions:
            for i, a in enumerate(actions, 1):
                with st.expander("Action " + str(i) + " — " + a["speaker"] + " (" + a["time"] + ")"):
                    st.markdown("**🗣️ Original:** " + a["text"])
                    st.markdown("**🌐 " + tgt_name + ":** " + a.get("text_tr", ""))
        else:
            st.info("No action items extracted from this meeting.")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Transcript ────────────────────────
        st.markdown('<div class="section-header header-orange">🗒️ Full Transcript</div>', unsafe_allow_html=True)
        with st.expander("View full bilingual transcript (" + str(len(diarized)) + " segments)"):
            for seg in diarized:
                st.markdown(
                    "**[" + str(seg["start"]) + "s] " + seg["speaker"] + ":** " +
                    seg["text"] + "  \n*🌐 " + seg.get("text_tr", "") + "*"
                )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Download ──────────────────────────
        st.markdown('<div class="section-header header-green">📥 Download Report</div>', unsafe_allow_html=True)
        with open(docx_path, "rb") as f:
            st.download_button(
                label     = "📥 Download Bilingual Meeting Minutes (DOCX)",
                data      = f,
                file_name = docx_path,
                mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

else:
    # Empty state
    st.markdown("""
    <div style='text-align:center; padding: 4rem 2rem; 
                background: rgba(255,255,255,0.03); 
                border: 2px dashed rgba(255,255,255,0.1);
                border-radius: 20px; margin-top: 1rem;'>
        <div style='font-size:4rem; margin-bottom:1rem;'>🎙️</div>
        <div style='font-family: Syne, sans-serif; font-size:1.3rem; 
                    color:rgba(255,255,255,0.7); font-weight:700;'>
            Drop your meeting audio here
        </div>
        <div style='color:rgba(255,255,255,0.3); margin-top:0.5rem; font-size:0.9rem;'>
            Supports mp3 · wav · m4a · mp4
        </div>
        <div style='margin-top:1.5rem;'>
            <span class='badge badge-purple'>🎙️ Auto language detection</span>&nbsp;
            <span class='badge badge-pink'>🌐 50+ languages</span>&nbsp;
            <span class='badge badge-green'>📄 DOCX export</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
