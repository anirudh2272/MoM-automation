# 🎙️ MoM Live AI — Multilingual Minutes of Meeting Automation System

> Turn any meeting recording into a polished, bilingual report — with **Gemini-powered analysis**, speaker attribution, action items, decisions, open questions, sentiment, an effectiveness score, risk detection, and a downloadable DOCX. Built as a portfolio-ready Streamlit app that runs on the free Streamlit Community Cloud tier.

**Live demo:** [animomautomation.streamlit.app](https://animomautomation.streamlit.app)

---

## 📋 Project Overview

Manually writing minutes of meeting is slow, error-prone, and language-bound. **MoM Live AI** is a free, open, end-to-end pipeline that:

1. Accepts an audio or video upload (mp3 / wav / m4a / mp4)
2. Transcribes the meeting speech-to-text using **faster-whisper**
3. Assigns heuristic speaker labels based on pause patterns
4. **Analyzes the meeting with Google Gemini** (optional, user-supplied key) — or a built-in rule-based engine as fallback
5. Extracts an executive summary, action items, decisions, open questions, **sentiment, an effectiveness score, and risks/blockers**
6. Translates the key sections into any of **52 languages**
7. Generates a professional, formatted **DOCX** report

---

## ✨ Features

- **🎙️ Audio & video upload** — drag-and-drop, 25 MB limit
- **🌐 Automatic language detection** with optional manual hint (52 languages)
- **🤖 Gemini AI analysis** (bring-your-own-key, free from Google AI Studio):
  - Abstractive executive summary that reads naturally
  - Smart action items with **inferred owners and deadlines**
  - Key decisions extracted by understanding context, not keywords
  - Open questions / unresolved items
  - **Meeting sentiment** (Positive / Neutral / Tense / Mixed)
  - **Meeting effectiveness score** (0–100 with rationale)
  - **Risk & blocker detection** with severity levels
- **🧠 Rule-based fallback** — works with zero API keys, fully offline-friendly
- **👥 Heuristic speaker diarization** — pause-based, 1–6 speakers
- **🌐 52 × 52 translation** — Google Translate via `deep-translator`
- **📄 Professional DOCX export** — colored banners, metadata table, AI insights, numbered lists, italic translations
- **⚡ Two transcription modes** — Fast (tiny) and Better (base) for Indic/Asian languages
- **🎨 Premium dark UI** — aurora gradients, Instrument Serif + Inter typography, glass cards

---

## 🛠️ Tech Stack

| Layer | Technology |
| --- | --- |
| Frontend / UI | Streamlit + custom CSS |
| Transcription | faster-whisper (CPU, int8) |
| Audio conversion | FFmpeg (system package) |
| **AI analysis** | **Google Gemini 2.5 Flash via `google-genai`** |
| Translation | deep-translator → Google Translate |
| Fallback extraction | Rule-based Python (no heavy LLMs) |
| Report export | python-docx |
| Deployment | Streamlit Community Cloud |

---

## 🔄 How It Works

```
Upload (.mp3/.wav/.m4a/.mp4)
        │
        ▼
   FFmpeg → 16kHz mono WAV
        │
        ▼
   faster-whisper transcription
        │
        ▼
   Pause-based speaker labels
        │
        ▼
   ┌─────────────────────────────┐
   │  Gemini API key provided?   │
   └─────────────┬───────────────┘
        yes      │      no
        ▼        │       ▼
   Gemini 2.5    │   Rule-based
   Flash         │   extraction
   ├ summary     │   ├ summary
   ├ actions     │   ├ actions
   ├ decisions   │   ├ decisions
   ├ questions   │   └ questions
   ├ sentiment   │
   ├ score       │
   └ risks       │
        │        │
        └────┬───┘
             ▼
   Google Translate → 52 languages
             │
             ▼
   python-docx → polished report
```

---

## 🤖 Getting a Free Gemini API Key

The AI features are **optional** and use **your own free key** (so deployment costs you nothing):

1. Go to [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
2. Sign in with a Google account (no credit card required)
3. Click **Create API key**
4. Copy it and paste it into the app's sidebar under **AI Engine**

The free tier covers Gemini Flash models with generous limits — plenty for meeting analysis. Without a key, the app automatically uses its built-in rule-based engine.

---

## 🚀 Local Setup

### 1. Clone

```bash
git clone https://github.com/anirudh2272/MoM-automation.git
cd MoM-automation
```

### 2. Install FFmpeg

```bash
# macOS
brew install ffmpeg
# Ubuntu / Debian
sudo apt-get install ffmpeg
# Windows (Chocolatey)
choco install ffmpeg
```

### 3. Install Python deps

```bash
python -m venv .venv
source .venv/bin/activate     # macOS / Linux
pip install -r requirements.txt
```

### 4. Run

```bash
streamlit run app.py
```

Opens at `http://localhost:8501`.

---

## ☁️ Streamlit Cloud Deployment

1. Push this repo to your GitHub account.
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
3. **New app** → Repository `your-username/MoM-automation` → Branch `main` → Main file `app.py`.
4. **Deploy**. First build takes ~5 minutes.

Required files:

```
MoM-automation/
├── app.py              # Streamlit app
├── requirements.txt    # Python deps (includes google-genai)
├── packages.txt        # System deps (ffmpeg)
└── README.md
```

---

## ⚠️ Limitations

> **This is a portfolio-friendly version.** It uses lightweight local processing and best-effort translation, with optional Gemini AI for higher-quality analysis. For enterprise-grade accuracy, future versions can integrate OpenAI, Google Cloud Speech-to-Text, Google Cloud Translation, AssemblyAI, Deepgram, or GPU-hosted models.

- **Transcription accuracy** scales with model size. Whisper tiny/base handle European languages well; Telugu, Hindi, Tamil, Arabic and other Indic/Asian languages do better with larger models on GPU.
- **Speaker diarization is heuristic** — labels rotate on pauses ≥ 1.5s.
- **Gemini free tier has rate limits** — on heavy use you may hit a limit and fall back to rule-based.
- **Translation uses Google Translate's free endpoint** — rate limits and outages are possible.
- **Max upload size: 25 MB** to stay within Streamlit Cloud limits.

---

## 🚧 Future Improvements

- True speaker diarization with `pyannote.audio` (GPU)
- Gemini-powered follow-up email draft and meeting chaptering
- Real-time streaming transcription (WebSocket + OpenAI Realtime / AssemblyAI Streaming)
- Calendar integration — push action items to Google Calendar / Outlook
- Email auto-distribution to attendees
- Multi-meeting search and analytics dashboard
- PDF export alongside DOCX

---

## 💼 Resume Bullet

> Built **MoM Live AI**, a multilingual meeting-intelligence web app that transcribes meeting audio, runs **LLM-powered analysis with Google Gemini** (summary, action items, sentiment, effectiveness scoring, and risk detection) with a rule-based fallback, translates output into 52 languages, and exports professional DOCX reports — using Streamlit, faster-whisper, google-genai, deep-translator, and python-docx, deployed on Streamlit Cloud.

---

## 🔗 LinkedIn Description

> What if your meetings analyzed themselves?
>
> I built **MoM Live AI** — a multilingual meeting-intelligence app that turns any recording into a polished bilingual report. Upload your standup, a Hindi product call, or a French client review and walk away with a clean DOCX containing an AI-generated summary, action items with owners and deadlines, key decisions, open questions, meeting sentiment, an effectiveness score, and detected risks — translated into any of 52 languages.
>
> The AI layer is powered by Google Gemini (bring-your-own-key) with a rule-based fallback so it always works. Built with Streamlit + faster-whisper + google-genai + deep-translator + python-docx. Deployed free on Streamlit Cloud.
>
> 🔗 Live: https://animomautomation.streamlit.app
> 🔗 Code: https://github.com/anirudh2272/MoM-automation

---

## 📄 License

MIT — fork and adapt freely.

---

## 🙏 Acknowledgments

- [faster-whisper](https://github.com/SYSTRAN/faster-whisper) by SYSTRAN
- [Google Gen AI SDK](https://github.com/googleapis/python-genai)
- [deep-translator](https://github.com/nidhaloff/deep-translator)
- [python-docx](https://python-docx.readthedocs.io)
- [Streamlit](https://streamlit.io)
