import streamlit as st
import pypdf
from pptx import Presentation
import zipfile
import io
import re
import os

# ── API Keys — loaded from environment variables (sidebar input is the primary method) ──
GEMINI_KEYS  = [k.strip() for k in os.environ.get("GEMINI_API_KEYS",  "").split(",") if k.strip()]
CLAUDE_KEYS  = [k.strip() for k in os.environ.get("CLAUDE_API_KEYS",  "").split(",") if k.strip()]
OPENAI_KEYS  = [k.strip() for k in os.environ.get("OPENAI_API_KEYS",  "").split(",") if k.strip()]

# ── CSS ──────────────────────────────────────────────────────────────────────
CSS = """
@import url('https://fonts.googleapis.com/css2?family=Kalam:wght@300;400;700&family=Patrick+Hand&display=swap');

:root {
    --ink:        #1c1c2e;
    --accent:     #e74c3c;
    --blue:       #1a6fc4;
    --green:      #1e8449;
    --orange:     #d35400;
    --purple:     #6c3483;
    --teal:       #0e7c7b;
    --pink:       #c0392b;
    --paper:      #fffdf4;
    --paper2:     #f5f0e8;
    --rule:       #c8c0b0;
    --margin:     #f4b9b2;
    --shadow:     rgba(0,0,0,0.10);
}

*, *::before, *::after { box-sizing: border-box; }

body, .stMarkdown, .stTabs [data-baseweb="tab-panel"] {
    background: var(--paper);
    font-family: 'Kalam', 'Comic Sans MS', 'Chalkboard SE', cursive;
    font-size: 17px;
    line-height: 2.0;
    color: var(--ink);
}

/* ── Notebook page ── */
.note-paper {
    position: relative;
    background: var(--paper);
    border-radius: 6px;
    padding: 2.5rem 3rem 2.5rem 4.5rem;
    box-shadow: 3px 4px 18px var(--shadow), -1px 0 6px rgba(0,0,0,0.04);
    background-image:
        linear-gradient(90deg, var(--margin) 56px, transparent 56px),
        repeating-linear-gradient(
            transparent, transparent 31px,
            var(--rule) 31px, var(--rule) 32px
        );
    min-height: 500px;
}

/* Red margin line gutter ring holes */
.note-paper::before {
    content: "◉\A◉\A◉";
    white-space: pre;
    position: absolute;
    left: 14px; top: 60px;
    font-size: 18px;
    color: #bbb;
    line-height: 120px;
}

/* ── Headings ── */
h1 {
    font-family: 'Patrick Hand', 'Kalam', cursive;
    color: var(--accent);
    text-align: center;
    font-size: 2rem;
    border-bottom: 3px double var(--accent);
    padding-bottom: 8px;
    margin-bottom: 1.2rem;
    letter-spacing: 1px;
    text-shadow: 1px 1px 0 #ffd6d6;
}
h2 {
    color: var(--blue);
    font-size: 1.35rem;
    border-bottom: 2.5px solid var(--blue);
    padding-bottom: 3px;
    margin-top: 1.8rem;
    display: flex;
    align-items: center;
    gap: 6px;
}
h3 {
    color: var(--green);
    font-size: 1.1rem;
    margin-top: 1rem;
    text-decoration: underline;
    text-decoration-style: wavy;
    text-decoration-color: var(--green);
}

/* ── Inline two-column layout ── */
.two-col {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin: 14px 0;
}
.three-col {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 14px;
    margin: 14px 0;
}

/* ── Callout boxes ── */
.highlight-box {
    background: linear-gradient(135deg, #fef9e7 80%, #fdebd0);
    border-left: 6px solid #f39c12;
    border-radius: 0 10px 10px 0;
    padding: 12px 16px;
    margin: 14px 0;
    box-shadow: 2px 2px 6px rgba(243,156,18,0.15);
    position: relative;
}
.highlight-box::before {
    content: "✏️ Key Concept";
    font-weight: 700;
    display: block;
    margin-bottom: 5px;
    color: #b7770d;
    font-size: 0.85rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.activity-box {
    background: linear-gradient(135deg, #eafaf1 80%, #d5f5e3);
    border-left: 6px solid var(--green);
    border-radius: 0 10px 10px 0;
    padding: 12px 16px;
    margin: 14px 0;
    box-shadow: 2px 2px 6px rgba(30,132,73,0.12);
}
.activity-box::before {
    content: "🧪 Activity / Experiment";
    font-weight: 700;
    display: block;
    margin-bottom: 5px;
    color: var(--green);
    font-size: 0.85rem;
    text-transform: uppercase;
}

.case-study-box {
    background: linear-gradient(135deg, #eaf4fb 80%, #d6eaf8);
    border-left: 6px solid var(--blue);
    border-radius: 0 10px 10px 0;
    padding: 12px 16px;
    margin: 14px 0;
    box-shadow: 2px 2px 6px rgba(26,111,196,0.12);
}
.case-study-box::before {
    content: "🌍 Real-World Application";
    font-weight: 700;
    display: block;
    margin-bottom: 5px;
    color: var(--blue);
    font-size: 0.85rem;
    text-transform: uppercase;
}

.exam-box {
    background: linear-gradient(135deg, #fdf2f8 80%, #f9ebf7);
    border: 2px dashed #8e44ad;
    border-radius: 10px;
    padding: 12px 16px;
    margin: 14px 0;
}
.exam-box::before {
    content: "🎯 Exam Tip";
    font-weight: 700;
    display: block;
    margin-bottom: 5px;
    color: var(--purple);
    font-size: 0.85rem;
    text-transform: uppercase;
}

.memory-box {
    background: linear-gradient(135deg, #e8f8f5 80%, #d1f2eb);
    border: 2px solid var(--teal);
    border-radius: 10px;
    padding: 12px 16px;
    margin: 14px 0;
    font-style: italic;
}
.memory-box::before {
    content: "🧠 Memory Trick";
    font-weight: 700;
    font-style: normal;
    display: block;
    margin-bottom: 5px;
    color: var(--teal);
    font-size: 0.85rem;
    text-transform: uppercase;
}

/* ── Image placeholder ── */
.img-placeholder {
    border: 3px dashed #aaa;
    border-radius: 10px;
    background: #f9f9f9;
    min-height: 160px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    color: #999;
    font-size: 0.9rem;
    text-align: center;
    padding: 16px;
    margin: 14px 0;
    gap: 8px;
}
.img-placeholder::before {
    content: "🖼️";
    font-size: 2.5rem;
}

/* ── Sticky note ── */
.sticky {
    background: #fff9c4;
    border-radius: 4px 4px 4px 20px;
    padding: 12px 16px;
    margin: 14px 0;
    box-shadow: 3px 4px 10px rgba(0,0,0,0.15), inset 0 -2px 4px rgba(0,0,0,0.05);
    transform: rotate(-1deg);
    display: inline-block;
    min-width: 180px;
    font-size: 0.95rem;
}

/* ── Tables ── */
table {
    width: 100%;
    border-collapse: collapse;
    margin: 16px 0;
    font-size: 0.95rem;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 2px 8px var(--shadow);
}
th {
    background: var(--blue);
    color: white;
    padding: 10px 14px;
    text-align: left;
    font-family: 'Patrick Hand', cursive;
}
td {
    padding: 8px 14px;
    border-bottom: 1px solid var(--rule);
    background: var(--paper2);
}
tr:nth-child(even) td { background: var(--paper); }

/* ── Bullet list ── */
ul.bullet-list { padding-left: 1.2rem; margin: 8px 0; }
ul.bullet-list li {
    margin-bottom: 6px;
    list-style: none;
    padding-left: 1.4rem;
    position: relative;
}
ul.bullet-list li::before {
    content: "✦";
    position: absolute;
    left: 0;
    color: var(--blue);
    font-size: 0.85rem;
    top: 2px;
}

/* ── pre / flowchart ── */
pre {
    background: var(--paper2);
    border: 1.5px solid var(--rule);
    border-radius: 8px;
    padding: 14px 18px;
    font-family: 'Kalam', monospace;
    font-size: 0.95rem;
    overflow-x: auto;
    line-height: 1.7;
}

/* ── Formula box ── */
.formula-block {
    background: linear-gradient(135deg, #f4f0ff 80%, #e8daff);
    border: 2px dashed var(--purple);
    border-radius: 10px;
    padding: 16px 20px;
    margin: 14px 0;
    text-align: center;
    font-size: 1.1rem;
    box-shadow: 2px 2px 8px rgba(108,52,131,0.10);
    overflow-x: auto;
}

/* ── Mobile responsive ── */
@media (max-width: 768px) {
    body, .stMarkdown, .stTabs [data-baseweb="tab-panel"] {
        font-size: 15px;
        line-height: 1.8;
    }

    .note-paper {
        padding: 1.2rem 1rem 1.2rem 1.2rem;
        background-image: repeating-linear-gradient(
            transparent, transparent 31px,
            var(--rule) 31px, var(--rule) 32px
        );
    }

    /* Hide ring holes on mobile — too cramped */
    .note-paper::before { display: none; }

    h1 { font-size: 1.4rem; }
    h2 { font-size: 1.1rem; }
    h3 { font-size: 1rem; }

    /* Stack columns on mobile */
    .two-col, .three-col {
        grid-template-columns: 1fr;
        gap: 10px;
    }

    .sticky {
        display: block;
        transform: none;
        min-width: unset;
        width: 100%;
    }

    table { font-size: 0.85rem; }
    th, td { padding: 6px 8px; }

    pre {
        font-size: 0.82rem;
        padding: 10px 12px;
    }

    .formula-block { font-size: 0.95rem; padding: 12px 10px; }

    .highlight-box, .activity-box, .case-study-box,
    .exam-box, .memory-box { padding: 10px 12px; }

    .img-placeholder { min-height: 120px; font-size: 0.82rem; }
}

@media (max-width: 480px) {
    body, .stMarkdown { font-size: 14px; }
    h1 { font-size: 1.2rem; }
    .note-paper { padding: 1rem 0.8rem; }
}
"""

HTML_WRAPPER = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
  <title>{title}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Kalam:wght@300;400;700&family=Patrick+Hand&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="style.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
  <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js"
          onload="renderMathInElement(document.body,{{delimiters:[{{left:'$$',right:'$$',display:true}},{{left:'$',right:'$',display:false}}]}});"></script>
</head>
<body>
  <div class="note-paper">
    {body}
  </div>
</body>
</html>"""

# ── Extraction helpers ────────────────────────────────────────────────────────
def extract_pdf(buf: io.BytesIO) -> str:
    reader = pypdf.PdfReader(buf)
    pages = []
    for page in reader.pages:
        t = page.extract_text()
        if t:
            pages.append(t)
    return "\n\n".join(pages)


def extract_pptx(buf: io.BytesIO) -> str:
    prs = Presentation(buf)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        parts = [f"<h3>Slide {i}</h3>"]
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                parts.append(shape.text.strip())
        slides.append("\n".join(parts))
    return "\n\n".join(slides)


# ── Markdown remnant sanitiser ────────────────────────────────────────────────
def sanitise(html: str) -> str:
    """Convert stray Markdown into safe HTML tags."""
    # **bold** → <b>
    html = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", html)
    # *italic* → <em>
    html = re.sub(r"\*(.+?)\*", r"<em>\1</em>", html)
    # Leading dash list items → <li>  (only bare lines, not inside existing tags)
    html = re.sub(r"(?m)^[ \t]*[-•]\s+(.+)$", r"<li>\1</li>", html)
    # Wrap consecutive <li> runs in a <ul class="bullet-list">
    html = re.sub(r"(<li>.*?</li>)(\s*<li>)", r"\1\2", html, flags=re.S)
    html = re.sub(r"((?:<li>.*?</li>\s*)+)", r'<ul class="bullet-list">\1</ul>', html, flags=re.S)
    # Remove accidental triple-backtick fences
    html = re.sub(r"```[a-z]*\n?", "", html)
    return html


# ── Shared prompt builder ─────────────────────────────────────────────────────
def build_prompt(raw_text: str, subject: str) -> str:
    return f"""You are an expert Science educator, NCERT content analyst, and topper-level student note-maker for subject "{subject}".

Your task is to convert the provided chapter into premium handwritten-style short notes exactly like the revision notebook of a rank-holding student.

OUTPUT RULES — follow exactly:
1. Return ONLY raw HTML. No markdown code fences (no ```html). No preamble or trailing commentary.
2. Divide your output using exactly these three HTML div markers (each on its own line):
   <div id="visual-notes">   ... </div>
   <div id="summary-page">   ... </div>
   <div id="formula-matrix"> ... </div>

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION 1 — <div id="visual-notes">
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Generate PREMIUM HANDWRITTEN TOPPER NOTES for the ENTIRE chapter.

⚠️ STRICT COVERAGE RULE — THIS IS MANDATORY:
• Go through the source material LINE BY LINE.
• Every heading, subheading, paragraph, example, activity, box, table, and figure caption MUST appear in the notes.
• Do NOT summarise multiple topics into one sentence. Each topic gets its OWN section.
• Do NOT write "and more..." or "etc." — spell out every item explicitly.
• Do NOT skip any topic even if it seems minor — completeness is the #1 priority.
• If the source has 10 topics, your notes must have 10 sections. If it has 20, write 20.
• Write MORE, not less. Long detailed notes are correct. Short summaries are wrong.

Structure to follow INSIDE this div:

A) CHAPTER MAP (tree format using <pre> tags):
   CHAPTER NAME
   │
   ├── Topic 1
   ├── Topic 2
   └── Topic N

B) TOPIC-WISE NOTES — for every topic and subtopic:
   • Use <h2> for main topics, <h3> for subtopics
   • Wrap key definitions/laws in:        <div class="highlight-box">
   • Wrap activities/experiments in:      <div class="activity-box">
   • Wrap real-world applications in:     <div class="case-study-box">
   • Wrap exam tips in:                   <div class="exam-box">
   • Wrap memory tricks/mnemonics in:     <div class="memory-box">
   • Use bullet points (<ul class="bullet-list"><li>) for concise revision points
   • Add emoji markers inline:
       ⭐ Important  🔥 Very Important  ⚠️ Common Mistake
       🎯 Exam Point  📝 Remember  💡 Shortcut  🧠 Memory Trick

   INLINE LAYOUT — use two-column and three-column grids often:
   • Side-by-side comparisons:  <div class="two-col"><div>...</div><div>...</div></div>
   • Three quick facts:         <div class="three-col"><div>...</div>...</div>
   • Sticky notes for tips:     <div class="sticky">short tip text</div>

   IMAGE PLACEHOLDERS — add a placeholder wherever a diagram/figure would help:
   <div class="img-placeholder">Diagram: [describe what the diagram shows, e.g. "Life cycle of a cell"]</div>
   Place these inline within the relevant topic section.

   For each topic include: Definition · Key idea · Real-life example · Exam relevance · Common misconception

C) SCIENTIFIC TERMS — wrap in a highlight-box:
   Term → Meaning  (one per line)

D) IMPORTANT EXAMPLES — for each example:
   🎯 Example N — Situation | Variables | Scientific Learning

E) ACTIVITIES & EXPERIMENTS — for each activity:
   🧪 Activity — Aim | Observation | Conclusion | Exam Question

F) SCIENCE THINKING BOXES (if applicable) — separate highlight-boxes for:
   🔍 Scientific Method · Modelling · Prediction · Estimation · Evidence · Limitations

G) SPECIAL FEATURES — convert "Meet a Scientist", "Threads of Curiosity", "Pause and Ponder" into:
   💡 Beyond NCERT  boxes using case-study-box class

H) MEMORY TOOLS — at the end:
   🧠 Mnemonics | Recall Tricks | One-line Memory Keys | Last Minute Tips

Use <pre> blocks for tree/flowchart structures. Use <table> for comparison tables.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION 2 — <div id="summary-page">
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ONE-PAGE REVISION SHEET. Include ALL of the following:
• All definitions  • All laws & principles  • All theories
• All symbols      • All important examples  • All scientific skills
• Common mistakes  • 10 Most Important Exam Points

Use <h2> for each category. Use <ul class="bullet-list"><li> for all points.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION 3 — <div id="formula-matrix">
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
For every formula/equation/constant/unit:
   Wrap in: <div class="formula-block">
   Include: Formula (LaTeX inside $$ ... $$) · Symbol meanings · Units · Conditions · Exam Tip

If the chapter has NO mathematical formulas, create a "📌 Important Relationships" section instead,
listing conceptual relationships in the same formula-block style.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SOURCE MATERIAL (subject: {subject}):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{raw_text}
"""


# ── Slice helper (shared) ─────────────────────────────────────────────────────
def slice_div(div_id: str, source: str) -> str:
    open_tag = f'<div id="{div_id}">'
    start = source.find(open_tag)
    if start == -1:
        return f"<p><em>Section '{div_id}' not found. Try re-processing.</em></p>"
    start += len(open_tag)
    depth = 1
    i = start
    while i < len(source) and depth > 0:
        open_pos = source.find("<div", i)
        close_pos = source.find("</div>", i)
        if close_pos == -1:
            break
        if open_pos != -1 and open_pos < close_pos:
            depth += 1
            i = open_pos + 4
        else:
            depth -= 1
            if depth == 0:
                return sanitise(source[start:close_pos].strip())
            i = close_pos + 6
    return sanitise(source[start:].strip())


def parse_sections(raw: str) -> dict:
    return {
        "notes":    slice_div("visual-notes",   raw),
        "bullets":  slice_div("summary-page",   raw),
        "formulas": slice_div("formula-matrix", raw),
    }


# ── Claude call ───────────────────────────────────────────────────────────────
def call_claude(raw_text: str, api_keys: list, subject: str) -> dict:
    prompt = build_prompt(raw_text, subject)
    try:
        import anthropic
    except ImportError:
        st.error("Install the Anthropic SDK:  pip install anthropic")
        st.stop()
    last_err = None
    raw = None
    for key in api_keys:
        try:
            client = anthropic.Anthropic(api_key=key)
            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=16000,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip()
            break
        except Exception as e:
            last_err = e
    if raw is None:
        raise last_err
    return parse_sections(raw)


# ── OpenAI call ───────────────────────────────────────────────────────────────
def call_openai(raw_text: str, api_keys: list, subject: str) -> dict:
    prompt = build_prompt(raw_text, subject)
    try:
        from openai import OpenAI
    except ImportError:
        st.error("Install the OpenAI SDK:  pip install openai")
        st.stop()
    last_err = None
    raw = None
    for key in api_keys:
        try:
            client = OpenAI(api_key=key)
            response = client.chat.completions.create(
                model="gpt-4.5-preview",
                max_tokens=16000,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.choices[0].message.content.strip()
            break
        except Exception as e:
            last_err = e
    if raw is None:
        raise last_err
    return parse_sections(raw)


# ── Ollama call (no API key needed) ──────────────────────────────────────────
def call_ollama(raw_text: str, model: str, subject: str) -> dict:
    import urllib.request, json
    prompt = build_prompt(raw_text, subject)
    payload = json.dumps({"model": model, "prompt": prompt, "stream": False, "options": {"num_predict": 6000, "num_ctx": 16384}}).encode()
    req = urllib.request.Request(
        "http://localhost:11434/api/generate",
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=900) as resp:
            raw = json.loads(resp.read())["response"].strip()
    except Exception as e:
        raise RuntimeError(
            f"Ollama error: {e}. Make sure Ollama is running (`ollama serve`) "
            f"and model '{model}' is pulled (`ollama pull {model}`)."
        )
    return parse_sections(raw)


# ── Gemini call ───────────────────────────────────────────────────────────────
def call_gemini(raw_text: str, api_keys: list, subject: str) -> dict:
    prompt = build_prompt(raw_text, subject)
    last_err = None
    raw = None
    for key in api_keys:
        try:
            try:
                from google import genai as _g
                client = _g.Client(api_key=key)
                try:
                    from google.genai import types as gt
                    cfg = gt.GenerateContentConfig(max_output_tokens=16000)
                    resp = client.models.generate_content(
                        model="gemini-2.0-flash-lite", contents=prompt, config=cfg)
                except Exception:
                    resp = client.models.generate_content(
                        model="gemini-2.0-flash-lite", contents=prompt)
                raw = resp.text.strip()
            except ImportError:
                import google.generativeai as genai_legacy
                genai_legacy.configure(api_key=key)
                resp = genai_legacy.GenerativeModel("gemini-2.0-flash-lite").generate_content(
                    prompt, generation_config={"max_output_tokens": 16000})
                raw = resp.text.strip()
            break
        except Exception as e:
            last_err = e
    if raw is None:
        raise last_err
    return parse_sections(raw)


# ── ZIP builder ───────────────────────────────────────────────────────────────
def build_zip(data: dict, base_name: str) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("style.css", CSS)
        zf.writestr("index.html",    HTML_WRAPPER.format(title="Detailed Notes",    body=data["notes"]))
        zf.writestr("summary.html",  HTML_WRAPPER.format(title="Quick Highlights",  body=data["bullets"]))
        zf.writestr("formulas.html", HTML_WRAPPER.format(title="Formula Sheet",     body=data["formulas"]))
        zf.writestr("README.txt",
            "Open index.html, summary.html, or formulas.html in any browser.\n"
            "KaTeX formulas render when connected to the internet.\n"
            f"Generated from: {base_name}\n")
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Production Note Studio", page_icon="📒", layout="wide")

# Inject sketchbook CSS globally
st.markdown(f"<style>{CSS}</style>", unsafe_allow_html=True)

st.title("📒 Production Note Studio")
st.caption("Transform textbook PDFs & lecture decks into premium digital sketchbook notebooks.")

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuration")

    provider = st.radio(
        "🤖 AI Provider",
        ["🏠 Ollama (Free, No API Key)", "🆓 Gemini (Free API Key)", "💎 Claude (Paid — Best Quality)", "🧠 GPT-4.5 (OpenAI)"],
        index=0,
        help="Ollama runs 100% locally — no API key ever. Gemini has a free tier. Claude/GPT need paid keys."
    )

    use_ollama = provider.startswith("🏠")
    use_gemini = provider.startswith("🆓")
    use_openai = provider.startswith("🧠")

    if use_ollama:
        api_keys = []
        ollama_model = st.selectbox(
            "Local Model",
            ["llama3.2", "mistral", "gemma2:2b", "phi3", "qwen2.5"],
            help="Run `ollama pull <model>` once to download it."
        )
        st.info("▶ Make sure Ollama is running: `ollama serve`")
    elif use_gemini:
        key_input = st.text_input(
            "Gemini API Key",
            type="password",
            placeholder="AIza...",
            help="Get a free key at https://aistudio.google.com/apikey",
        )
        api_keys = [key_input] if key_input.strip() else GEMINI_KEYS
        ollama_model = None
    elif use_openai:
        key_input = st.text_input(
            "OpenAI API Key",
            type="password",
            placeholder="sk-...",
            help="Get your key at https://platform.openai.com/api-keys",
        )
        api_keys = [key_input] if key_input.strip() else OPENAI_KEYS
        ollama_model = None
    else:
        key_input = st.text_input(
            "Claude API Key",
            type="password",
            placeholder="sk-ant-...",
            help="Get your key at https://console.anthropic.com/",
        )
        api_keys = [key_input] if key_input.strip() else CLAUDE_KEYS
        ollama_model = None

    subject = st.selectbox(
        "Subject",
        ["Biology 🔬", "Physics ⚡", "Chemistry 🧪", "Mathematics 📐",
         "History 📜", "Geography 🌍", "Economics 📊", "Computer Science 💻"],
    )
    st.divider()
    st.markdown("**Supported formats:** PDF, PPTX")
    if use_ollama:
        st.markdown("**Runs locally — 100% free, no internet needed after model download.**")
        st.markdown("[Install Ollama](https://ollama.com/download)")
    elif use_gemini:
        st.markdown("**Model:** gemini-2.0-flash-lite · [Get free key](https://aistudio.google.com/apikey)")
    elif use_openai:
        st.markdown("**Model:** gpt-4.5-preview · [Get key](https://platform.openai.com/api-keys)")
    else:
        st.markdown("**Model:** claude-haiku-4-5 (cost-efficient) · [Get key](https://console.anthropic.com/)")

# Upload
uploaded = st.file_uploader("Upload a Chapter PDF or Lecture Deck (.pptx)", type=["pdf", "pptx"])

if uploaded:
    if not use_ollama and not api_keys:
        provider_name = "Gemini" if use_gemini else ("OpenAI" if use_openai else "Anthropic")
        var_name = "GEMINI_API_KEYS" if use_gemini else ("OPENAI_API_KEYS" if use_openai else "CLAUDE_API_KEYS")
        st.error(f"No {provider_name} API key configured. Set the `{var_name}` environment variable or add the key in the config block at the top of `app.py`.")
        st.stop()


    ext = uploaded.name.rsplit(".", 1)[-1].lower()
    base_name = uploaded.name

    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"📄 **{base_name}** — ready to process")
    with col2:
        run_btn = st.button("⚡ Generate Notes", type="primary", use_container_width=True)

    if run_btn:
        with st.status("Processing your document…", expanded=True) as status:
            st.write("📖 Extracting text from file…")
            file_buf = io.BytesIO(uploaded.read())

            if ext == "pdf":
                raw_text = extract_pdf(file_buf)
            else:
                raw_text = extract_pptx(file_buf)

            word_count = len(raw_text.split())
            st.write(f"✅ Extracted **{word_count:,} words** from {base_name}")

            if word_count < 50:
                status.update(label="Extraction failed — too little text found.", state="error")
                st.error("The file appears to have very little extractable text. Scanned image PDFs are not supported.")
                st.stop()

            MAX_WORDS = 4000 if use_ollama else (5000 if use_gemini else 12000)
            words = raw_text.split()
            if len(words) > MAX_WORDS:
                raw_text = " ".join(words[:MAX_WORDS])
                st.warning(f"⚠️ Chapter has {word_count:,} words — processing first {MAX_WORDS:,} words. For full coverage use Claude or GPT-4.5.")

            if use_ollama:
                provider_label = f"Ollama ({ollama_model})"
            elif use_gemini:
                provider_label = "Gemini"
            elif use_openai:
                provider_label = "GPT-4.5"
            else:
                provider_label = "Claude"
            st.write(f"🤖 Sending to {provider_label} for note compilation…")
            try:
                if use_ollama:
                    sections = call_ollama(raw_text, ollama_model, subject)
                elif use_gemini:
                    sections = call_gemini(raw_text, api_keys, subject)
                elif use_openai:
                    sections = call_openai(raw_text, api_keys, subject)
                else:
                    sections = call_claude(raw_text, api_keys, subject)
            except Exception as e:
                status.update(label="Error — see details below.", state="error")
                st.error(f"**{provider_label} error:** {e}")
                if use_ollama:
                    st.info("Install Ollama from https://ollama.com/download, then run: ollama serve && ollama pull llama3.2")
                elif use_gemini:
                    st.info("Get a free Gemini key at https://aistudio.google.com/apikey — install SDK: pip install google-genai")
                elif use_openai:
                    st.info("Get an OpenAI key at https://platform.openai.com/api-keys — install SDK: pip install openai")
                else:
                    st.info("Get a Claude key at https://console.anthropic.com/ — install SDK: pip install anthropic")
                st.stop()

            st.write("📦 Building ZIP export package…")
            zip_buf = build_zip(sections, base_name)

            status.update(label="Notes ready!", state="complete", expanded=False)

        # Store in session state so tabs persist without re-running
        st.session_state["sections"] = sections
        st.session_state["zip_buf"]  = zip_buf
        st.session_state["base_name"] = base_name

# Render results if available
if "sections" in st.session_state:
    sections  = st.session_state["sections"]
    zip_buf   = st.session_state["zip_buf"]
    base_name = st.session_state["base_name"]

    st.success(f"✅ Notebook generated for **{base_name}**")

    tab1, tab2, tab3 = st.tabs(["🎨 Plain Paper Notes", "📌 Quick Highlights", "📐 Formula Sheet"])

    with tab1:
        st.markdown('<div class="note-paper">', unsafe_allow_html=True)
        st.markdown(sections["notes"], unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="note-paper">', unsafe_allow_html=True)
        st.markdown(sections["bullets"], unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="note-paper">', unsafe_allow_html=True)
        st.markdown(sections["formulas"], unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.download_button(
        label="⬇️ Download Offline HTML + CSS Bundle (.zip)",
        data=zip_buf,
        file_name=f"{base_name.rsplit('.', 1)[0]}_notes.zip",
        mime="application/zip",
        type="primary",
    )
