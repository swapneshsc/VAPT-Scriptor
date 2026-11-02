# VAPT Report Generator — Web UI

A browser-based interface for the VAPT report tool.

## Requirements

- Python 3.8+
- `python-docx` (`pip install python-docx`)

## Run

```bash
cd vapt_web_app
python3 server.py
```

Then open in your browser:

```
http://127.0.0.1:8080
```

## Features

1. **Load Full Sample** — GoldenEye-style complete example
2. **Load Minimal Template** — smaller starter JSON
3. **Upload JSON** — use your own engagement file
4. **Edit in browser** — modify the JSON in the text area
5. **Generate DOCX** — server builds the report and offers download

## How it works

- Frontend: single-page HTML + JS
- Backend: Python `http.server` + the same `generate_vapt_report.py` engine
- No extra frameworks required (Flask/Streamlit not needed)

## Tips

- Keep severity values as `CRITICAL`, `HIGH`, `MEDIUM`, `LOW`, `INFO` for colour coding.
- Large reports generate in a few seconds.
- The downloaded filename is derived from `meta.target_system`.

## Stop the server

Press `Ctrl+C` in the terminal where `server.py` is running.
