# 🎧 PodcastLM Studio

**The open-source alternative to Google's NotebookLM Audio Overview.**

Turn any article, PDF, YouTube video, or text into a fully voiced, two-host podcast — with studio-quality mixing, background music, and caller segments. Runs entirely on your own API keys. No Google account required. No data leaves your control.

**[Live Demo →](https://your-app-url.streamlit.app)** · **[Deploy Your Own ↓](#deploy-in-60-seconds)**

---

## Why This Exists

Google's NotebookLM showed the world that AI-generated podcasts can sound surprisingly good. But it's a closed system — your documents go to Google, you get what Google gives you back, and you have zero control over the output.

PodcastLM Studio takes the same idea and rips it wide open:

| | Google NotebookLM | PodcastLM Studio |
|---|---|---|
| **Source code** | Closed, proprietary | Fully open source (MIT) |
| **Your data** | Processed on Google servers | Processed via API keys *you* own |
| **LLM provider** | Google Gemini only | DeepSeek, OpenAI, xAI Grok — your choice |
| **Script editing** | None — take what you get | Full line-by-line editing, rewriting, reordering |
| **Voice control** | No customization | 6 voices, per-speaker speed, HD toggle |
| **Background music** | None | 4 presets + custom upload, intro/outro support |
| **Caller segments** | Not available | Built-in phone-effect caller with custom questions |
| **Language support** | English only | 20+ languages including Urdu, Arabic, Hebrew, Hindi |
| **Cost transparency** | Hidden (bundled into Google One) | Live cost calculator — most podcasts cost $0.15–$1.00 |
| **Privacy mode** | Not available | One-click source text wipe after generation |
| **Export options** | Audio only | Audio + Markdown script + plain text + SRT subtitles |
| **Deployment** | Google Cloud only | One-click deploy to Streamlit Community Cloud (free) |

**You own every piece of this pipeline.** Fork it, modify it, host it wherever you want.

---

## What It Does

1. **Ingest anything** — PDFs, DOCX, PPTX, web articles, YouTube videos, audio files, or raw text
2. **Research your source** — AI chat assistant that answers questions using *only* your uploaded material
3. **Generate a script** — Natural two-host conversation with humor, tangents, and back-and-forth
4. **Edit and rehearse** — Line-by-line script editor with per-line audio preview
5. **Produce the final podcast** — Parallel TTS generation, background music mixing, phone-effect caller segments, loudness normalization, intro/outro support
6. **Export everything** — Download the podcast MP3, the script (Markdown or plain text), and SRT subtitles

All of this runs in your browser. No installation required.

---

## Features

### 🧠 Multi-Provider Intelligence
Choose your LLM:
- **DeepSeek-V3** (default) — best quality-to-cost ratio (~$0.03/script)
- **DeepSeek-R1** — reasoning model for complex source material
- **GPT-4o-mini** — fast, reliable, supports longest outputs
- **xAI Grok 4.1** — latest reasoning model from xAI

Budget Mode forces GPT-4o-mini regardless of selection (~90% cheaper).

### 🎙️ Studio-Quality Audio
- 6 OpenAI TTS voices in 3 curated pairs (Dynamic, Calm, Formal)
- **TTS-1-HD toggle** for noticeably better voice quality
- **Per-speaker speed control** — Host 1, Host 2, and Caller each get independent sliders
- Phone-effect processing on caller segments (bandpass filter + echo)
- Loudness normalization (EBU R128: -16 LUFS)
- Background music with 4 presets or custom upload
- Optional 5-second music ramp-up before dialogue
- Custom intro/outro clip support

### 🌐 20+ Languages
English (US/UK) · Spanish · French · German · Italian · Portuguese · Hindi · Urdu · Arabic · Hebrew · Russian · Turkish · Japanese · Korean · Chinese (Mandarin) · Polish · Dutch · Swedish · Indonesian · Thai

Script generation and TTS work natively in all supported languages.

### 🔒 Privacy-First Design
- Your documents are processed in-session only — nothing is stored server-side
- **Privacy Mode** wipes source text from memory immediately after script generation
- All API keys live in Streamlit Secrets — never committed to the repo
- No analytics, no tracking, no telemetry

### 💾 Session Persistence
- Save your entire session (source text, script, chat history) as a JSON file
- Restore any saved session later — no work lost if your browser tab closes

### 📦 Export Options
- **MP3 podcast** with timestamped filename
- **Markdown script** for blog posts or show notes
- **Plain text script** for accessibility
- **SRT subtitles** with estimated timestamps — ready for video overlay

---

## Deploy in 60 Seconds

### Option 1: One-Click Deploy (Recommended)

[![Deploy to Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/deploy)

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. **New app** → connect your fork → set the main file to `p0dcaster_appv2.py`
4. In **Settings → Secrets**, paste:

```toml
APP_PASSWORD = "your-chosen-password"
OPENAI_API_KEY = "sk-..."
DEEPSEEK_API_KEY = "sk-..."
XAI_API_KEY = "xai-..."

    Deploy. That's it.

    Note: Only OPENAI_API_KEY is required (for TTS/Whisper). DEEPSEEK_API_KEY is strongly recommended (default LLM). XAI_API_KEY is optional.

Option 2: Run Locally

Bash

git clone https://github.com/your-username/podcastlm-studio.git
cd podcastlm-studio
pip install -r requirements.txt
# Create .streamlit/secrets.toml with your keys (see above)
streamlit run p0dcaster_appv2.py

Requires ffmpeg installed on your system (brew install ffmpeg / apt install ffmpeg).
Cost Breakdown

The app shows a live cost estimate in the sidebar — before and after script generation.
Podcast Length	LLM (DeepSeek-V3)	TTS (Standard)	TTS (HD)	Total
Short (2 min)	~$0.03	~$0.07	~$0.13	$0.10–$0.16
Medium (5 min)	~$0.03	~$0.18	~$0.36	$0.21–$0.39
Long (15 min)	~$0.03	~$0.50	~$1.00	$0.53–$1.03
Extra Long (30 min)	~$0.03	~$1.00	~$2.00	$1.03–$2.03

Using GPT-4o-mini adds ~$0.07. Using Grok adds ~$0.27. Still cheaper than a coffee.
Architecture

text

┌─────────────────────────────────────────────────┐
│                   Browser (Streamlit)            │
├──────────┬──────────┬──────────┬────────────────┤
│ 📄 Source │ 💬 Chat  │ 📝 Script│  🎚️ Produce   │
│          │          │          │                │
│ PDF/DOCX │ Research │ LLM Gen  │ Parallel TTS   │
│ Web/YT   │ Q&A      │ Edit     │ Phone FX       │
│ Audio    │          │ Rehearse │ Music Mix       │
│ Text     │          │ Export   │ Loudnorm        │
└────┬─────┴────┬─────┴────┬─────┴───────┬────────┘
     │          │          │             │
     ▼          ▼          ▼             ▼
  Whisper   DeepSeek    DeepSeek     OpenAI TTS
  (OpenAI)  /OpenAI     /OpenAI      + ffmpeg
            /Grok       /Grok

All processing happens through standard API calls. The only binary dependency is ffmpeg for audio mixing (installed via packages.txt on Streamlit Cloud).
Tech Stack
Component	Technology
Frontend & hosting	Streamlit (Community Cloud — free)
Script generation	DeepSeek-V3 / OpenAI GPT-4o-mini / xAI Grok
Text-to-speech	OpenAI TTS-1 / TTS-1-HD
Transcription	OpenAI Whisper
Audio processing	ffmpeg via ffmpeg-python
Video download	yt-dlp
Document parsing	PyPDF2, python-docx, python-pptx
Web scraping	BeautifulSoup4

No Docker. No microservices. No GPU. Just Python.
Project Structure

text

podcastlm-studio/
├── p0dcaster_appv2.py        # Main application (single file)
├── requirements.txt           # Python dependencies
├── packages.txt               # System packages (ffmpeg)
├── .gitignore                 # Keeps secrets out of the repo
└── .streamlit/
    └── config.toml            # Theme and server settings

Contributing

Open an issue or PR if you want to:

    🎵 Add new background music presets
    🗣️ Integrate ElevenLabs or Azure TTS voices
    🎬 Add export-to-video with burned-in subtitles
    🌍 Improve non-Latin script rendering
    🧪 Add automated tests
    📱 Improve mobile layout

All contributions welcome. This is a community project.
FAQ

Q: Do I need all three API keys?
A: Only OPENAI_API_KEY is required — it powers TTS and Whisper transcription. DEEPSEEK_API_KEY is strongly recommended as the default script engine. XAI_API_KEY is entirely optional.

Q: How is this different from NotebookLM?
A: NotebookLM is a closed Google product — your data goes to Google, you get a fixed output, and you can't customize anything. PodcastLM Studio is open source, runs on API keys you control, and gives you full editorial control over the script, voices, pacing, and production.

Q: Can I use this commercially?
A: The code is MIT-licensed — do whatever you want with it. Check the terms of service for whichever API providers you use (OpenAI, DeepSeek, xAI).

Q: Why DeepSeek as the default?
A: DeepSeek-V3 offers the best quality-to-cost ratio for script generation — roughly 10× cheaper than Grok and 3× cheaper than GPT-4o-mini, with comparable or better output quality for creative writing tasks.

Q: What if the script gets cut off?
A: The app automatically detects truncated output and recovers all complete dialogue lines. You'll see a warning with the number of recovered lines. Try a shorter duration setting or switch to GPT-4o-mini (which supports longer outputs) if this happens frequently.

Q: Is my data private?
A: Your documents are processed in-session via API calls to the providers you choose. Nothing is stored on the Streamlit server. Enable Privacy Mode to wipe source text from memory immediately after script generation.
License

MIT — use it however you want.
Acknowledgments

Built with late-night caffeine and the belief that podcast creation should be open, affordable, and controllable.

Inspired by the AI podcasting wave of 2025 — but built for people who want to own their tools.

⭐ Star this repo if it saved you hours of editing.

Made with 🎧 and open-source conviction.
