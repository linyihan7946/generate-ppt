from __future__ import annotations

import os
import socket
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from generate_ppt.pipeline import generate_ppt
from generate_ppt.templates import list_templates


ROOT = Path(__file__).resolve().parents[1]
UPLOAD_DIR = ROOT / "workspace" / "uploads"
OUTPUT_DIR = ROOT / "workspace" / "outputs"
ALLOWED_EXTENSIONS = {".pdf", ".docx", ".md", ".markdown", ".txt"}


def create_app() -> Flask:
    app = Flask(__name__, template_folder=str(ROOT / "web_templates"), static_folder=str(ROOT / "static"))
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    @app.get("/")
    def index():
        return render_template("index.html", templates=list_templates())

    @app.post("/api/generate")
    def api_generate():
        upload = request.files.get("document")
        template_id = request.form.get("template_id", "technical-no-image")
        if upload is None or not upload.filename:
            return jsonify({"error": "请先选择一个技术文档。"}), 400

        original_name = upload.filename
        suffix = Path(original_name).suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            return jsonify({"error": f"暂不支持 {suffix} 文件。"}), 400

        filename = secure_filename(original_name) or f"document{suffix}"
        input_path = UPLOAD_DIR / filename
        upload.save(input_path)

        try:
            output_path = generate_ppt(input_path, template_id, OUTPUT_DIR)
        except Exception as exc:
            return jsonify({"error": str(exc)}), 500

        return jsonify(
            {
                "filename": output_path.name,
                "download_url": f"/download/{output_path.name}",
            }
        )

    @app.get("/download/<filename>")
    def download(filename: str):
        output_path = OUTPUT_DIR / filename
        if not output_path.exists():
            return jsonify({"error": "文件不存在。"}), 404
        return send_file(output_path, as_attachment=True, download_name=filename)

    return app


def open_browser_later(url: str) -> None:
    if os.environ.get("NO_BROWSER") == "1":
        return

    def _open() -> None:
        webbrowser.open(url)

    timer = threading.Timer(1.0, _open)
    timer.daemon = True
    timer.start()


def find_free_port(start: int = 8765, host: str = "127.0.0.1") -> int:
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex((host, port)) != 0:
                return port
    raise RuntimeError("没有找到可用端口。")
