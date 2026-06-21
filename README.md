# 🎙️ MoM Live AI — Multilingual Minutes of Meeting Automation System

> Turn any meeting recording into a polished, bilingual report — with speaker attribution, action items, decisions, open questions, and a downloadable DOCX. Built as a portfolio-ready Streamlit app that runs on the free Streamlit Community Cloud tier.

**Live demo:** [animomautomation.streamlit.app](https://animomautomation.streamlit.app)

---

## 📋 Project Overview

Manually writing minutes of meeting is slow, error-prone, and language-bound. Most existing tools only support English and require enterprise budgets. **MoM Live AI** is a free, open, end-to-end pipeline that:

1. Accepts an audio or video upload (mp3 / wav / m4a / mp4)
2. Transcribes the meeting speech-to-text using **faster-whisper**
3. Assigns heuristic speaker labels based on pause patterns
4. Extracts an executive summary, action items, decisions, and open questions
5. Translates the key sections into any of **52 languages** via Google Translate
6. Generates a professional, formatted **DOCX** report you can share with the team

---

## ✨ Features

- **🎙️ Audio & video upload** — drag-and-drop interface, 25 MB limit
- **🌐 Automatic language detection** with optional manual hint (52 languages)
- **🧠 Lightweight meeting intelligence** — no heavy BART/mBART models, runs on Streamlit Cloud's free tier
  - Executive summary (extractive, keyword-weighted scoring)
  - Action items (with owner and deadline detection)
  - Key decisions
  - Open questions / pending items
- **👥 Heuristic speaker diarization** — pause-based, 1–6 speakers
- **🌐 52 × 52 translation** — Google Translate via `deep-translator`
- **📄 Professional DOCX export** — colored section banners, metadata table, numbered lists, speaker labels, italic translations
- **⚡ Two model modes** — Fast (tiny) for speed, Better (base) for accuracy on Indic/Asian languages
- **🎨 Premium dark UI** — aurora gradient backgrounds, Instrument Serif + Inter typography, glass cards

---

## 🛠️ Tech Stack

| Layer | Technology |
| --- | --- |
| Frontend / UI | Streamlit + custom CSS |
| Transcription | faster-whisper (CPU, int8) |
| Audio conversion | FFmpeg (system package) |
| Translation | deep-translator → Google Translate |
| Summarization & extraction | Rule-based Python (no heavy LLMs) |
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
   Meeting intelligence extraction
   ├─ Executive summary
   ├─ Action items + deadlines
   ├─ Key decisions
   └─ Open questions
        │
        ▼
   Google Translate → 52 languages
        │
        ▼
   python-docx → polished report
```

---

## 🚀 Local Setup

### 1. Clone the repository

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

# Windows (via Chocolatey)
choco install ffmpeg
```

### 3. Install Python dependencies

```bash
python -m venv .venv
source .venv/bin/activate   # macOS / Linux
# .venv\Scripts\activate    # Windows

pip install -r requirements.txt
```

### 4. Run the app

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## ☁️ Streamlit Cloud Deployment

1. Fork or push this repository to your own GitHub account.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app** and select:
   - **Repository:** `your-username/MoM-automation`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**. The first build takes ~5 minutes (downloads Whisper on first request).

Required files in the repo:

```
MoM-automation/
├── app.py              # Streamlit app
├── requirements.txt    # Python deps
├── packages.txt        # System deps (ffmpeg)
└── README.md
```

---

## ⚠️ Limitations

> **This is a portfolio-friendly version.** It uses lightweight local processing and best-effort free translation. For enterprise-grade accuracy, future versions can integrate OpenAI, Google Cloud Speech-to-Text, Google Cloud Translation, AssemblyAI, Deepgram, or GPU-hosted models.

Specifically:

- **Transcription accuracy** scales with model size. Whisper "tiny" and "base" handle most European languages well; Telugu, Hindi, Tamil, Arabic and other Indic/Asian languages produce better results with larger models on a GPU host.
- **Speaker diarization is heuristic** — labels rotate on pauses ≥ 1.5s and may not match real speakers in overlapping speech.
- **Summarization and extraction are rule-based** to keep the app deployable on free Streamlit Cloud RAM. Quality is below LLM-based methods.
- **Translation uses Google Translate's free endpoint** via `deep-translator`. Rate limits and outages are possible.
- **Max upload size: 25 MB** to stay within Streamlit Cloud limits.

---

## 🚧 Future Improvements

- True speaker diarization with `pyannote.audio` (needs GPU host)
- LLM-based summarization and action extraction (OpenAI / Anthropic / open-source)
- Real-time streaming transcription via WebSocket + OpenAI Realtime / AssemblyAI Streaming
- Calendar integration — push action items to Google Calendar / Outlook
- Email auto-distribution to meeting attendees
- Confidence scores per extracted item
- Multi-meeting search and analytics dashboard
- PDF export alongside DOCX

---

## 💼 Resume Bullet

> Built **MoM Live AI**, a multilingual meeting-intelligence web app that transcribes meeting audio, generates speaker-attributed transcripts, extracts decisions and action items, translates key sections into 52 languages, and exports professional DOCX reports — using Streamlit, faster-whisper, deep-translator, and python-docx, deployed on Streamlit Cloud.


---

## 📄 License

MIT — feel free to fork and adapt.

---

- [deep-translator](https://github.com/nidhaloff/deep-translator)
- [python-docx](https://python-docx.readthedocs.io)
- [Streamlit](https://streamlit.io) for the cleanest Python UI framework around
