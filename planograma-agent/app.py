import os
import io
import uuid
import logging
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file
from dotenv import load_dotenv

import analyzer
import claude_client
import report_excel
import report_pdf

load_dotenv()

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s — %(message)s",
    level=logging.INFO,
)
log = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))

MAX_MB = int(os.environ.get("MAX_FILE_SIZE_MB", 10))
MAX_BYTES = MAX_MB * 1024 * 1024

# In-memory session store: {session_id: {"result": ..., "ts": datetime}}
_sessions: dict = {}
SESSION_TTL = timedelta(hours=1)


def _evict_old():
    cutoff = datetime.utcnow() - SESSION_TTL
    old = [k for k, v in _sessions.items() if v["ts"] < cutoff]
    for k in old:
        del _sessions[k]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze_route():
    _evict_old()

    if "file" not in request.files:
        return jsonify({"success": False, "error": "Niciun fișier trimis."}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".xlsx"):
        return jsonify({"success": False, "error": "Fișierul trebuie să fie în format .xlsx"}), 400

    raw = f.read()
    if len(raw) > MAX_BYTES:
        return jsonify({"success": False, "error": f"Fișierul depășește limita de {MAX_MB} MB."}), 413

    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
    except Exception as exc:
        log.error("Excel parse error: %s", exc)
        return jsonify({"success": False, "error": "Fișierul Excel nu poate fi citit. Verificați formatul."}), 400

    try:
        result = analyzer.analyze(xl)
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)}), 422
    except Exception as exc:
        log.exception("Analyzer error")
        return jsonify({"success": False, "error": f"Eroare la analiză: {exc}"}), 500

    narrative = ""
    try:
        narrative = claude_client.get_narrative(result["stats"])
    except Exception as exc:
        log.error("Claude API error: %s", exc)
        narrative = ""

    result["narrative"] = narrative

    session_id = str(uuid.uuid4())
    _sessions[session_id] = {"result": result, "ts": datetime.utcnow()}

    return jsonify({
        "success": True,
        "session_id": session_id,
        "stats": result["stats"],
        "chart_data": result["chart_data"],
        "recommendations": result["recommendations"],
        "narrative": narrative or "Analiza AI indisponibilă temporar.",
    })


@app.route("/download/excel/<session_id>")
def download_excel(session_id: str):
    entry = _sessions.get(session_id)
    if not entry:
        return jsonify({"error": "Sesiune expirată. Reîncărcați analiza."}), 404
    data = report_excel.generate(entry["result"])
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"planograma_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    )


@app.route("/download/pdf/<session_id>")
def download_pdf(session_id: str):
    entry = _sessions.get(session_id)
    if not entry:
        return jsonify({"error": "Sesiune expirată. Reîncărcați analiza."}), 404
    data = report_pdf.generate(entry["result"])
    return send_file(
        io.BytesIO(data),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"planograma_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
    )


@app.route("/clear", methods=["POST"])
def clear():
    sid = request.json.get("session_id") if request.is_json else None
    if sid and sid in _sessions:
        del _sessions[sid]
    return jsonify({"success": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
