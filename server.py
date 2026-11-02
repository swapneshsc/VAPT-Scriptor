#!/usr/bin/env python3
"""
VAPT Report Tool — Web Interface
================================
Simple web UI for the VAPT report generator.
Uses only Python stdlib + python-docx (already available).

Run:
    python3 server.py
Then open http://127.0.0.1:8080
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import traceback
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

# Make the generator importable
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from generate_vapt_report import generate_report  # noqa: E402

HOST = "0.0.0.0"
PORT = 8080


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>VAPT Report Generator</title>
  <style>
    :root {
      --bg: #0f1419;
      --card: #1a2332;
      --border: #2d3a4f;
      --text: #e7ecf3;
      --muted: #8b9bb4;
      --accent: #3b82f6;
      --accent-hover: #2563eb;
      --success: #22c55e;
      --danger: #ef4444;
      --warning: #f59e0b;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
      line-height: 1.5;
    }
    .container {
      max-width: 960px;
      margin: 0 auto;
      padding: 2rem 1.25rem 4rem;
    }
    header {
      text-align: center;
      margin-bottom: 2rem;
    }
    header h1 {
      font-size: 1.75rem;
      font-weight: 700;
      letter-spacing: -0.02em;
      background: linear-gradient(135deg, #60a5fa, #a78bfa);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
    }
    header p {
      color: var(--muted);
      margin-top: 0.5rem;
      font-size: 0.95rem;
    }
    .card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 1.5rem;
      margin-bottom: 1.25rem;
    }
    .card h2 {
      font-size: 1.05rem;
      margin-bottom: 0.75rem;
      color: #93c5fd;
    }
    .row {
      display: flex;
      flex-wrap: wrap;
      gap: 0.75rem;
      margin-bottom: 1rem;
    }
    button, .btn {
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      padding: 0.55rem 1.1rem;
      border: none;
      border-radius: 8px;
      font-size: 0.9rem;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s, transform 0.1s;
    }
    button:active { transform: scale(0.98); }
    .btn-primary {
      background: var(--accent);
      color: #fff;
    }
    .btn-primary:hover { background: var(--accent-hover); }
    .btn-secondary {
      background: #334155;
      color: var(--text);
    }
    .btn-secondary:hover { background: #475569; }
    .btn-success {
      background: var(--success);
      color: #052e16;
    }
    .btn-success:hover { background: #16a34a; }
    .btn-danger {
      background: #7f1d1d;
      color: #fecaca;
    }
    textarea {
      width: 100%;
      min-height: 320px;
      background: #0c1220;
      border: 1px solid var(--border);
      border-radius: 8px;
      color: #e2e8f0;
      font-family: 'Cascadia Code', 'Fira Code', 'Consolas', monospace;
      font-size: 0.8rem;
      padding: 1rem;
      resize: vertical;
      line-height: 1.45;
    }
    textarea:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.25);
    }
    .status {
      margin-top: 1rem;
      padding: 0.75rem 1rem;
      border-radius: 8px;
      font-size: 0.9rem;
      display: none;
    }
    .status.show { display: block; }
    .status.info  { background: #1e3a5f; color: #93c5fd; }
    .status.ok    { background: #14532d; color: #86efac; }
    .status.err   { background: #7f1d1d; color: #fecaca; }
    .file-input-wrap {
      position: relative;
      overflow: hidden;
      display: inline-block;
    }
    .file-input-wrap input[type=file] {
      position: absolute;
      left: 0; top: 0;
      opacity: 0;
      width: 100%; height: 100%;
      cursor: pointer;
    }
    footer {
      text-align: center;
      color: var(--muted);
      font-size: 0.8rem;
      margin-top: 2rem;
    }
    .hint {
      color: var(--muted);
      font-size: 0.85rem;
      margin-bottom: 0.75rem;
    }
    #downloadLink {
      display: none;
      margin-top: 1rem;
    }
    #downloadLink a {
      color: #86efac;
      font-weight: 600;
      text-decoration: none;
    }
    #downloadLink a:hover { text-decoration: underline; }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>VAPT Report Generator</h1>
      <p>Upload or edit a JSON description → generate a professional Word report</p>
    </header>

    <div class="card">
      <h2>1. Load data</h2>
      <p class="hint">Start from a sample, upload your own JSON, or paste/edit below.</p>
      <div class="row">
        <button class="btn-secondary" onclick="loadSample('full')">Load Full Sample (GoldenEye)</button>
        <button class="btn-secondary" onclick="loadSample('minimal')">Load Minimal Template</button>
        <div class="file-input-wrap">
          <button class="btn-secondary">Upload JSON file</button>
          <input type="file" id="fileInput" accept=".json,application/json" />
        </div>
        <button class="btn-danger" onclick="clearEditor()">Clear</button>
      </div>
    </div>

    <div class="card">
      <h2>2. JSON content</h2>
      <textarea id="jsonEditor" placeholder="Paste or edit your engagement JSON here…" spellcheck="false"></textarea>
    </div>

    <div class="card">
      <h2>3. Generate report</h2>
      <div class="row">
        <button class="btn-primary" id="generateBtn" onclick="generateReport()">Generate DOCX Report</button>
      </div>
      <div id="status" class="status"></div>
      <div id="downloadLink">
        <a id="dlAnchor" href="#" download>⬇ Download report</a>
      </div>
    </div>

    <footer>
      VAPT Report Tool · Web UI · Uses the same engine as the CLI generator
    </footer>
  </div>

  <script>
    const editor = document.getElementById('jsonEditor');
    const statusEl = document.getElementById('status');
    const downloadLink = document.getElementById('downloadLink');
    const dlAnchor = document.getElementById('dlAnchor');
    const generateBtn = document.getElementById('generateBtn');

    function showStatus(msg, type) {
      statusEl.textContent = msg;
      statusEl.className = 'status show ' + type;
    }

    function clearStatus() {
      statusEl.className = 'status';
      downloadLink.style.display = 'none';
    }

    async function loadSample(kind) {
      clearStatus();
      showStatus('Loading sample…', 'info');
      try {
        const res = await fetch('/api/sample?kind=' + kind);
        if (!res.ok) throw new Error(await res.text());
        const text = await res.text();
        editor.value = text;
        showStatus('Sample loaded. You can edit it, then click Generate.', 'ok');
      } catch (e) {
        showStatus('Failed to load sample: ' + e.message, 'err');
      }
    }

    document.getElementById('fileInput').addEventListener('change', function (e) {
      const file = e.target.files[0];
      if (!file) return;
      clearStatus();
      const reader = new FileReader();
      reader.onload = function (ev) {
        editor.value = ev.target.result;
        showStatus('File loaded: ' + file.name, 'ok');
      };
      reader.onerror = function () {
        showStatus('Could not read file', 'err');
      };
      reader.readAsText(file);
    });

    function clearEditor() {
      editor.value = '';
      clearStatus();
    }

    async function generateReport() {
      clearStatus();
      const raw = editor.value.trim();
      if (!raw) {
        showStatus('Please load or paste JSON first.', 'err');
        return;
      }

      // Validate JSON client-side
      try {
        JSON.parse(raw);
      } catch (e) {
        showStatus('Invalid JSON: ' + e.message, 'err');
        return;
      }

      generateBtn.disabled = true;
      generateBtn.textContent = 'Generating…';
      showStatus('Generating report on server… this may take a few seconds.', 'info');

      try {
        const res = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: raw
        });

        if (!res.ok) {
          const errText = await res.text();
          throw new Error(errText || res.statusText);
        }

        const blob = await res.blob();
        const filename = res.headers.get('X-Filename') || 'VAPT_Report.docx';
        const url = URL.createObjectURL(blob);

        dlAnchor.href = url;
        dlAnchor.download = filename;
        dlAnchor.textContent = '⬇ Download ' + filename;
        downloadLink.style.display = 'block';
        showStatus('Report generated successfully!', 'ok');
      } catch (e) {
        showStatus('Generation failed: ' + e.message, 'err');
      } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generate DOCX Report';
      }
    }

    // Load minimal sample on first visit
    loadSample('minimal');
  </script>
</body>
</html>
"""


class VAPTHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        # Quieter logs
        sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(HTML_PAGE.encode("utf-8"))))
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))
            return

        if path == "/api/sample":
            qs = parse_qs(parsed.query)
            kind = (qs.get("kind") or ["minimal"])[0]
            if kind == "full":
                fname = "sample_input.json"
            else:
                fname = "minimal_template.json"
            fpath = ROOT / fname
            if not fpath.exists():
                self.send_error(404, "Sample not found")
                return
            data = fpath.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self.send_error(404, "Not found")

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/generate":
            self.send_error(404, "Not found")
            return

        length = int(self.headers.get("Content-Length", 0))
        if length <= 0 or length > 5_000_000:
            self.send_error(400, "Invalid or too large body")
            return

        body = self.rfile.read(length)
        try:
            data = json.loads(body.decode("utf-8"))
        except Exception as e:
            self.send_error(400, f"Invalid JSON: {e}")
            return

        # Generate into a temp file
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                out_path = Path(tmpdir) / "report.docx"
                generate_report(data, out_path)
                content = out_path.read_bytes()

            target = data.get("meta", {}).get("target_system", "VAPT")
            safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in target)
            safe_name = safe_name.replace(" ", "_")[:40] or "VAPT"
            filename = f"{safe_name}_VAPT_Report.docx"

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("X-Filename", filename)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            tb = traceback.format_exc()
            msg = f"Generation error: {e}\n\n{tb}"
            encoded = msg.encode("utf-8")
            self.send_response(500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(encoded)))
            self.end_headers()
            self.wfile.write(encoded)


def main():
    os.chdir(ROOT)
    server = HTTPServer((HOST, PORT), VAPTHandler)
    print(f"VAPT Report Web UI running at http://127.0.0.1:{PORT}")
    print("Press Ctrl+C to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.server_close()


if __name__ == "__main__":
    main()
