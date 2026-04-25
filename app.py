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

def transcribe(audio_path):
    model = load_whisper()
    raw, info = model.transcribe(audio_path, beam_size=5, language=None)
    segments = [{
        "start": round(s.start, 2),
        "end"  : round(s.end,   2),
        "text" : s.text.strip()
    } for s in raw]
    transcript = " ".join(s["text"] for s in segments)
    return transcript, segments, info.language, round(info.language_probability*100, 1)

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
    return " ".join(parts)

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
    if src == tgt or tgt not in MBART_LANG_MAP:
        return text
    tok, mdl = load_translator()
    tok.src_lang = MBART_LANG_MAP.get(src, "en_XX")
    inp = tok(text, return_tensors="pt", max_length=512,
              truncation=True, padding=True)
    inp = {k: v.to(DEVICE) for k, v in inp.items()}
    tgt_id = tok.lang_code_to_id[MBART_LANG_MAP[tgt]]
    out = mdl.generate(**inp, forced_bos_token_id=tgt_id,
                       max_length=512, num_beams=4)
    return tok.decode(out[0], skip_special_tokens=True)

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
    doc.add_paragraph(summary)
    p2 = doc.add_paragraph(summary_tr)
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
            tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_run.font.size = Pt(9)
            tt = tp.add_run(d.get("text_tr", ""))
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
            tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
            tr_run.font.size = Pt(9)
            tt = tp.add_run(a.get("text_tr", ""))
            tt.font.size = Pt(9)
            tt.italic = True
    else:
        doc.add_paragraph("No action items extracted.")
    doc.add_paragraph()

    add_banner(doc, "4.  Full Transcript (" + src_name + " + " + tgt_name + ")", "595959")
    for s in diarized:
        p = doc.add_paragraph()
        r = p.add_run("[" + str(s["start"]) + "s]  " + s["speaker"] + ": ")
        r.bold = True
        r.font.size = Pt(9)
        p.add_run(s["text"]).font.size = Pt(9)
        tp = doc.add_paragraph()
        tp.paragraph_format.left_indent = Inches(0.4)
        tr_run = tp.add_run("  🌐 " + tgt_name + ": ")
        tr_run.font.size = Pt(8)
        tt = tp.add_run(s.get("text_tr", ""))
        tt.font.size = Pt(8)
        tt.italic = True

    path = "Meeting_Minutes_" + src_name + "_to_" + tgt_name + ".docx"
    doc.save(path)
    return path


# ══════════════════════════════════════════════
#  STREAMLIT UI
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="MoM Automation System",
    page_icon="📋",
    layout="wide"
)

st.markdown("""
    <h1 style='text-align:center; color:#1F497D;'>
        📋 Minutes of Meeting Automation
    </h1>
    <p style='text-align:center; color:gray;'>
        Upload any meeting audio → Get bilingual formatted minutes
    </p>
    <hr>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Settings")
    target_lang = st.selectbox(
        "Translate output to:",
        options=list(LANG_NAMES.keys()),
        format_func=lambda x: LANG_NAMES[x] + " (" + x + ")",
        index=list(LANG_NAMES.keys()).index("hi")
    )
    num_speakers = st.slider(
        "Estimated number of speakers",
        min_value=1, max_value=6, value=3
    )
    st.markdown("---")
    st.markdown("**Supported formats:** `.mp3` `.wav` `.m4a` `.mp4`")
    st.markdown("---")
    st.markdown("**Models**")
    st.markdown("- faster-whisper")
    st.markdown("- BART-large-CNN")
    st.markdown("- BART-large-MNLI")
    st.markdown("- mBART-50")

uploaded = st.file_uploader(
    "📁 Upload your meeting audio",
    type=["mp3","wav","m4a","mp4"]
)

if uploaded:
    audio_path = "uploaded_" + uploaded.name
    with open(audio_path, "wb") as f:
        f.write(uploaded.read())

    st.audio(uploaded)
    st.success("✅ Uploaded: " + uploaded.name)

    if st.button("🚀 Generate Meeting Minutes", type="primary"):

        with st.spinner("🎙️ Transcribing audio..."):
            transcript, segments, det_lang, det_prob = transcribe(audio_path)

        st.info("🌐 Detected: **" + LANG_NAMES.get(det_lang, det_lang) + "** (" + str(det_prob) + "% confidence)")

        with st.spinner("👥 Assigning speakers..."):
            diarized = simple_diarize(segments, num_speakers)

        with st.spinner("🧠 Extracting summary, actions, decisions..."):
            summary   = get_summary(transcript)
            actions   = get_actions(diarized)
            decisions = get_decisions(diarized)

        tgt_name = LANG_NAMES.get(target_lang, target_lang)
        with st.spinner("🌐 Translating to " + tgt_name + "..."):
            summary_tr = translate_text(summary, det_lang, target_lang)
            for a in actions:
                a["text_tr"] = translate_text(a["text"], det_lang, target_lang)
            for d in decisions:
                d["text_tr"] = translate_text(d["text"], det_lang, target_lang)
            for s in diarized:
                s["text_tr"] = translate_text(s["text"], det_lang, target_lang)

        st.success("✅ Processing complete!")
        st.markdown("---")

        col1, col2, col3 = st.columns(3)
        col1.metric("Action Items",        len(actions))
        col2.metric("Key Decisions",       len(decisions))
        col3.metric("Transcript Segments", len(diarized))
        st.markdown("---")

        st.subheader("📝 Executive Summary")
        tab1, tab2 = st.tabs([
            "Original (" + LANG_NAMES.get(det_lang, det_lang) + ")",
            "Translated (" + tgt_name + ")"
        ])
        with tab1:
            st.write(summary)
        with tab2:
            st.write(summary_tr)

        st.subheader("🏛️ Key Decisions")
        if decisions:
            for i, d in enumerate(decisions, 1):
                with st.expander("Decision " + str(i) + " — " + d["speaker"] + " (" + d["time"] + ")"):
                    st.write("**Original:** " + d["text"])
                    st.write("**" + tgt_name + ":** " + d["text_tr"])
        else:
            st.info("No decisions extracted.")

        st.subheader("✅ Action Items")
        if actions:
            for i, a in enumerate(actions, 1):
                with st.expander("Action " + str(i) + " — " + a["speaker"] + " (" + a["time"] + ")"):
                    st.write("**Original:** " + a["text"])
                    st.write("**" + tgt_name + ":** " + a["text_tr"])
        else:
            st.info("No action items extracted.")

        st.subheader("🗒️ Full Transcript")
        with st.expander("View full transcript"):
            for s in diarized:
                st.markdown(
                    "**[" + str(s["start"]) + "s] " + s["speaker"] + ":** " +
                    s["text"] + "  \n*🌐 " + s.get("text_tr","") + "*"
                )

        st.markdown("---")
        with st.spinner("📄 Generating DOCX..."):
            docx_path = export_docx(
                summary, summary_tr, actions, decisions,
                diarized, det_lang, target_lang
            )

        with open(docx_path, "rb") as f:
            st.download_button(
                label="📥 Download Meeting Minutes (DOCX)",
                data=f,
                file_name=docx_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
