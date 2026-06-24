# 📒 Production Note Studio

Transform textbook PDFs and lecture decks into **premium handwritten-style HTML notes** — like a topper's revision notebook.

## Features

- 🎨 Handwritten notebook aesthetic (Kalam font, ruled paper, red margin line)
- 📑 Three outputs: **Detailed Notes · Quick Highlights · Formula Sheet**
- 📦 One-click ZIP download (offline HTML + CSS bundle)
- 📱 Mobile responsive
- 🤖 Multiple AI backends:
  | Provider | Cost | Notes |
  |---|---|---|
  | Ollama | Free (local) | Runs offline, no API key |
  | Gemini | Free tier | `gemini-2.0-flash-lite` |
  | Claude | Paid | `claude-haiku-4-5` — best quality |
  | GPT-4.5 | Paid | `gpt-4.5-preview` |

## Quick Start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501).

## API Keys

Enter your API key directly in the sidebar when the app is running — no config files needed.

Alternatively, set environment variables before launching:

```bash
set CLAUDE_API_KEYS=sk-ant-...
set GEMINI_API_KEYS=AIza...
set OPENAI_API_KEYS=sk-...
```

## Deploy on Streamlit Cloud

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Point to `app.py`
4. Add your keys under **Settings → Secrets**:
   ```toml
   CLAUDE_API_KEYS = "sk-ant-..."
   GEMINI_API_KEYS = "AIza..."
   ```

## Supported File Types

- **PDF** — text-based (scanned image PDFs not supported)
- **PPTX** — PowerPoint lecture decks
