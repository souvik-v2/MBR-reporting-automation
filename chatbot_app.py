"""
chatbot_app.py -- browser-based chat interface for the MBR data chatbot.

USAGE
-----
    python3 chatbot_app.py
    # then open http://127.0.0.1:5000 in a browser

Every question rebuilds the knowledge base fresh from sample_data/ (see
src/mbr_pipeline/knowledge_base.py) -- there is no caching, so an answer
always reflects whatever is currently on disk, including files added or
edited after the server started.

No external AI API is used. See src/mbr_pipeline/chat_engine.py for how
questions are matched to answers.
"""
from __future__ import annotations

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(_HERE))

from flask import Flask, request, jsonify, Response

from src.mbr_pipeline.knowledge_base import build_knowledge_base
from src.mbr_pipeline.chat_engine import answer_question

app = Flask(__name__)

SAMPLE_DATA_ROOT = _HERE / "sample_data"

PAGE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>MBR Data Assistant</title>
<style>
  :root {
    --bg: #0f1419;
    --panel: #1a2129;
    --accent: #4da3ff;
    --accent-dim: #2c5d8a;
    --text: #e8edf2;
    --text-dim: #8b97a3;
    --bubble-user: #2c5d8a;
    --bubble-bot: #232b33;
    --border: #2a323b;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); height: 100vh; display: flex; flex-direction: column;
  }
  header {
    padding: 16px 24px; border-bottom: 1px solid var(--border); background: var(--panel);
    display: flex; align-items: center; justify-content: space-between;
  }
  header h1 { font-size: 16px; margin: 0; font-weight: 600; }
  header .month { font-size: 13px; color: var(--text-dim); }
  #chat { flex: 1; overflow-y: auto; padding: 20px 24px; display: flex; flex-direction: column; gap: 14px; }
  .row { display: flex; }
  .row.user { justify-content: flex-end; }
  .bubble {
    max-width: 70%; padding: 10px 14px; border-radius: 12px; white-space: pre-wrap;
    line-height: 1.45; font-size: 14.5px;
  }
  .row.user .bubble { background: var(--bubble-user); color: white; border-bottom-right-radius: 3px; }
  .row.bot .bubble { background: var(--bubble-bot); border-bottom-left-radius: 3px; border: 1px solid var(--border); }
  .suggestions { padding: 0 24px 12px; display: flex; flex-wrap: wrap; gap: 8px; }
  .chip {
    border: 1px solid var(--border); color: var(--text-dim); background: transparent;
    border-radius: 999px; padding: 6px 12px; font-size: 12.5px; cursor: pointer;
  }
  .chip:hover { border-color: var(--accent); color: var(--accent); }
  form { display: flex; gap: 10px; padding: 16px 24px; border-top: 1px solid var(--border); background: var(--panel); }
  input[type=text] {
    flex: 1; background: var(--bg); border: 1px solid var(--border); color: var(--text);
    padding: 12px 14px; border-radius: 8px; font-size: 14.5px; outline: none;
  }
  input[type=text]:focus { border-color: var(--accent); }
  button {
    background: var(--accent); color: #08121a; border: none; border-radius: 8px;
    padding: 12px 20px; font-size: 14.5px; font-weight: 600; cursor: pointer;
  }
  button:hover { background: #6cb4ff; }
  button:disabled { background: var(--accent-dim); cursor: default; }
</style>
</head>
<body>
  <header>
    <h1>MBR Data Assistant</h1>
    <span class="month" id="month-label">loading...</span>
  </header>
  <div id="chat"></div>
  <div class="suggestions" id="suggestions"></div>
  <form id="form">
    <input type="text" id="input" placeholder="Ask about a project... e.g. Who is the LT Manager for SEOD?" autocomplete="off">
    <button type="submit" id="send">Send</button>
  </form>

<script>
const chat = document.getElementById('chat');
const form = document.getElementById('form');
const input = document.getElementById('input');
const sendBtn = document.getElementById('send');
const monthLabel = document.getElementById('month-label');
const suggestionsEl = document.getElementById('suggestions');

function addBubble(text, who) {
  const row = document.createElement('div');
  row.className = 'row ' + who;
  const bubble = document.createElement('div');
  bubble.className = 'bubble';
  bubble.textContent = text;
  row.appendChild(bubble);
  chat.appendChild(row);
  chat.scrollTop = chat.scrollHeight;
}

const SUGGESTIONS = [
  "List all projects",
  "What month is this report for?",
  "Who is the LT Manager for SEOD?",
  "What are the key achievements for FLEX?",
  "What is the project roadmap for SLE?",
  "What are the in-progress tickets for Consumer Web?",
];

function renderSuggestions() {
  suggestionsEl.innerHTML = '';
  SUGGESTIONS.forEach(s => {
    const chip = document.createElement('button');
    chip.className = 'chip';
    chip.type = 'button';
    chip.textContent = s;
    chip.onclick = () => { input.value = s; form.requestSubmit(); };
    suggestionsEl.appendChild(chip);
  });
}

async function loadMonth() {
  try {
    const res = await fetch('/api/info');
    const data = await res.json();
    monthLabel.textContent = data.month ? ('Reporting month: ' + data.month) : '';
  } catch (e) {
    monthLabel.textContent = '';
  }
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const text = input.value.trim();
  if (!text) return;
  addBubble(text, 'user');
  input.value = '';
  sendBtn.disabled = true;
  try {
    const res = await fetch('/api/ask', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: text })
    });
    const data = await res.json();
    addBubble(data.answer, 'bot');
  } catch (err) {
    addBubble('Something went wrong reaching the server. Is chatbot_app.py still running?', 'bot');
  } finally {
    sendBtn.disabled = false;
    input.focus();
  }
});

addBubble("Hi! I can answer questions about this month's MBR data -- projects, managers, achievements, roadmap, and in-progress tickets. Try one of the suggestions below, or ask your own.", 'bot');
renderSuggestions();
loadMonth();
</script>
</body>
</html>
"""


@app.route("/")
def index() -> Response:
    return Response(PAGE_HTML, mimetype="text/html")


@app.route("/api/info")
def info():
    try:
        kb = build_knowledge_base(SAMPLE_DATA_ROOT)
        return jsonify({"month": kb.month, "project_count": len(kb.projects)})
    except Exception as exc:
        return jsonify({"month": None, "error": str(exc)})


@app.route("/api/ask", methods=["POST"])
def ask():
    payload = request.get_json(silent=True) or {}
    question = (payload.get("question") or "").strip()
    if not question:
        return jsonify({"answer": "Ask me something -- e.g. \"Who is the LT Manager for SEOD?\""})

    try:
        # Rebuilt fresh on every request -- no caching, so an edit to any
        # file in sample_data/ is reflected on the very next question
        # with no server restart needed.
        kb = build_knowledge_base(SAMPLE_DATA_ROOT)
    except Exception as exc:
        return jsonify({"answer": f"Could not read sample_data/ right now: {exc}"})

    result = answer_question(question, kb)
    return jsonify({"answer": result.text, "intent": result.matched_intent, "project": result.matched_project})


if __name__ == "__main__":
    print("MBR Data Assistant starting...")
    print("Open http://127.0.0.1:5000 in your browser.")
    app.run(host="127.0.0.1", port=5000, debug=False)
