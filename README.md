# 🎧 PodcastLM Studio

**The open-source, privacy-first alternative to Google's NotebookLM.**

Turn any article, PDF, YouTube video, or text into a **fully voiced, two-host podcast**—complete with studio mixing, background music, and caller segments. Plus, generate comprehensive **Study Guides, Flashcards, and Briefings** from your source material.

Runs entirely on your own API keys. No Google account required. No data leaves your control.

---

### Why This Exists

Google's NotebookLM showed the world that AI-generated podcasts can sound surprisingly good. But it's a closed system—your documents go to Google, you get what Google gives you back, and you have zero control over the output.

**PodcastLM Studio** takes the same idea and rips it wide open:

| Feature | Google NotebookLM | PodcastLM Studio |
| :--- | :--- | :--- |
| **Source Code** | Closed, proprietary | **Fully Open Source (MIT)** |
| **Data Privacy** | Processed on Google servers | **Your keys, your control** (Privacy Mode available) |
| **LLM Provider** | Gemini only | **DeepSeek, OpenAI, Grok** — your choice |
| **Audio Cost** | Free (bundled) | **Free (Edge TTS)** or Paid (OpenAI HD) |
| **Script Editing** | None — take what you get | **Full line-by-line editing**, rewriting, reordering |
| **Study Tools** | On-screen only | **Downloadable** (PDF/DOCX/Markdown) Study Guides, Flashcards & more |
| **Voices** | No customization | **300+ Voices** (Edge) or 6 HD Voices (OpenAI) |
| **Caller Segments** | Not available | **Built-in phone-effect caller** with custom questions |
| **Languages** | Limited | **20+ Languages** (Urdu, Arabic, Hindi, Japanese, etc.) |
| **Deployment** | Google Cloud only | **Run Locally** or on Streamlit Cloud |

You own every piece of this pipeline. Fork it, modify it, host it wherever you want.

---

### What It Does

1.  **Ingest Anything** — PDFs, DOCX, PPTX, web articles, YouTube videos, audio files, or raw text.
2.  **Research** — Chat with your source material using a RAG-style assistant that never hallucinates outside your docs.
3.  **Generate Script** — Creates a natural two-host conversation with humor, tangents, and "human" interruptions.
4.  **Produce Audio** — Parallellized TTS generation + background music mixing + loudness normalization.
5.  **Study** — Generate academic-quality study materials (Flashcards, Timelines, Glossaries).

---

### Features

#### 📚 Intelligent Study Tools (New!)
Go beyond the audio. Generate deeply structured study materials from your source text and download them in **Markdown, Word (.docx), or Plain Text**:
*   **Study Guide:** Summaries, key takeaways, and review questions.
*   **Briefing Document:** Executive-style findings and implications.
*   **Flashcards:** Front/Back concepts (formatted for Anki/Quizlet).
*   **Timeline:** Chronological breakdown of events.
*   **FAQ & Glossary:** Key definitions and common questions.

#### 🎙️ Studio-Quality Audio
*   **Edge TTS (Free):** Use Microsoft's neural voices for $0 cost.
*   **OpenAI HD:** Toggle `TTS-1-HD` for ultra-realistic breathing and intonation.
*   **Fine-Grained Control:** Adjust speaking speed per host (e.g., Host 1 @ 1.0x, Caller @ 0.95x).
*   **Production Suite:** Automatic background music ducking, intro/outro clips, and "phone-call" EQ filters for guest callers.

#### 🧠 Multi-Provider Intelligence
*   **DeepSeek-V3:** The default engine. Incredible creative writing at 1/10th the cost of GPT-4.
*   **GPT-4o-mini:** Fast, cheap, and handles massive contexts.
*   **xAI Grok:** Optional integration for high-speed reasoning.

#### 🔒 Privacy-First
*   **Client-Side Logic:** Your documents are processed in memory during the session only.
*   **Privacy Mode:** One-click wipe of source text from memory immediately after script generation.
*   **Session State:** Save your work as a `.json` file to your local machine and restore it later. Nothing is stored in a database.

---

### Quick Start

#### Option 1: Run Locally (Recommended)

1.  **Clone the repo:**
    ```bash
    git clone https://github.com/your-username/podcastlm-studio.git
    cd podcastlm-studio
    ```

2.  **Install dependencies:**
    ```bash
    pip install streamlit openai edge-tts yt-dlp ffmpeg-python beautifulsoup4 python-docx python-pptx pypdf2 requests
    ```

3.  **Run the app:**
    ```bash
    streamlit run p0dcaster_appv2.py
    ```

*Note: You must have [FFmpeg](https://ffmpeg.org/download.html) installed on your system path for audio mixing.*

#### Option 2: Deploy to Cloud
1.  Fork this repo.
2.  Go to [Streamlit Community Cloud](https://share.streamlit.io).
3.  Connect your fork and set the main file to `p0dcaster_appv2.py`.
4.  Add your API keys in the dashboard under "Secrets".

---

### Cost Breakdown

The app includes a live cost estimator.

| Engine | Audio Provider | Typical Cost (15 min podcast) |
| :--- | :--- | :--- |
| **DeepSeek** | **Edge TTS** | **~$0.03** (Total) |
| **DeepSeek** | **OpenAI Standard** | **~$0.53** |
| **GPT-4o** | **OpenAI HD** | **~$2.50** |

**Cheapest Setup:** DeepSeek + Edge TTS = Practically free.

---

### Tech Stack

*   **Frontend:** Streamlit
*   **LLMs:** DeepSeek / OpenAI / Grok
*   **Audio:** Edge TTS (free) / OpenAI TTS (paid)
*   **Processing:** FFmpeg (via `ffmpeg-python`)
*   **Parsing:** `PyPDF2`, `python-docx`, `BeautifulSoup4`, `yt-dlp`

---

### Contributing

Open an issue or PR! We are looking for:
- 🗣️ **ElevenLabs Integration** (For premium voice cloning)
- ☁️ **Azure / Google TTS** (Mid-tier pricing options)
- 📱 **Mobile UI improvements**

---

### License

**MIT License.** Use it, fork it, sell it, change it. It's yours.

*Made with 🎧 and open-source conviction.*
