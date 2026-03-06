from __future__ import annotations

import json
import os
from pathlib import Path

from flask import Flask, Response, jsonify, send_file

BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", BASE_DIR / "output"))

app = Flask(__name__)


def _serve_html(name: str):
    path = OUTPUT_DIR / name
    if not path.exists():
        return Response(f"Missing {name}. Run: python scripts/regenerate.py", status=500)
    return send_file(path)


@app.get("/")
def home():
    return _serve_html("index.html")


@app.get("/calendar")
def calendar():
    return _serve_html("calendar.html")


@app.get("/personal")
def personal():
    return _serve_html("personal.html")


@app.get("/info")
def info():
    return _serve_html("info.html")


@app.get("/todos_pdf/<path:filename>")
def todo_pdf(filename: str):
    path = OUTPUT_DIR / "todos_pdf" / filename
    if not path.exists():
        return Response(f"Missing PDF: {filename}. Run: python scripts/regenerate.py", status=404)
    return send_file(path, mimetype="application/pdf")


@app.get("/api/data")
def data():
    path = OUTPUT_DIR / "all_todos.json"
    if not path.exists():
        return jsonify({"error": "Missing all_todos.json"}), 500
    return jsonify(json.loads(path.read_text()))


@app.get("/healthz")
def healthz():
    return jsonify({"ok": True})
