# Doc Highlighter

A desktop app that automatically finds and highlights text in your Google Docs — controllable via GUI or HTTP API (great for Claude Code).

**[→ Check out the website](https://aadithkk.github.io/doc-highlighter/)**

---

## Features

- Highlight text in Google Docs automatically via keyboard automation
- 7 color options: Yellow, Blue, Green, Pink, Red, Purple, Teal
- Built-in HTTP server (port 7798) for scripting and Claude Code integration
- Batch import — paste dozens of terms at once
- Runs 100% locally, no cloud, no accounts

## Requirements

- Python 3.8+
- Google Docs open in a browser window

## Install & Run

```bash
git clone https://github.com/AadithKK/doc-highlighter.git
cd doc-highlighter
pip install customtkinter pyautogui Pillow
python doc_highlighter.py
```

## Usage

1. Launch the app
2. Open your Google Doc and click into it
3. Add highlight terms via the GUI (or use the API below)
4. Set your delay, then press **Start**

### Text Format

Each line is one highlight entry. Optionally specify a color with `::`:

```
neural network
transformer architecture::Blue
gradient descent::Green
```

## HTTP API

The app starts an HTTP server on `http://localhost:7798` automatically.

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/add` | Add a single highlight entry |
| POST | `/add-and-start` | Add an entry and start the queue |
| POST | `/batch` | Add multiple entries (newline-separated) |
| POST | `/start` | Start the queue |
| POST | `/stop` | Stop the queue |
| GET | `/status` | Get running state and queue info |

### Examples

```bash
# Add a single highlight
curl -X POST http://localhost:7798/add \
  -H "Content-Type: application/json" \
  -d '{"text": "hello world", "color": "Yellow"}'

# Add and immediately start
curl -X POST http://localhost:7798/add-and-start \
  -H "Content-Type: application/json" \
  -d '{"text": "neural network", "color": "Blue"}'

# Batch add multiple terms
curl -X POST http://localhost:7798/batch \
  -H "Content-Type: application/json" \
  -d '{"lines": "transformer::Yellow\nattention mechanism::Blue\ngradient descent::Green", "color": "Yellow"}'

# Check status
curl http://localhost:7798/status

# Stop
curl -X POST http://localhost:7798/stop
```

## Claude Code Integration

With the HTTP server running, Claude Code can control the app directly:

```python
import subprocess

subprocess.run([
    "curl", "-s", "-X", "POST", "http://localhost:7798/batch",
    "-H", "Content-Type: application/json",
    "-d", '{"lines": "key term 1::Yellow\nkey term 2::Blue", "color": "Yellow"}'
])
```

Or use the curl commands directly in a Claude Code session — the Settings tab in the app shows ready-to-use examples.

## License

MIT
