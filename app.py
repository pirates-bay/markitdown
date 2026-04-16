import os
import socket
import tempfile
import webbrowser
from flask import Flask, request, jsonify, send_from_directory
from markitdown import MarkItDown

app = Flask(__name__, static_folder="static")
md = MarkItDown()

ALLOWED_EXTENSIONS = {
    "pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls",
    "html", "htm", "csv", "json", "xml", "txt", "md",
    "jpg", "jpeg", "png", "gif", "webp",
    "mp3", "wav", "m4a",
    "epub", "zip",
}

def allowed(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return send_from_directory("static", "index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not allowed(file.filename):
        return jsonify({"error": f"Unsupported file type"}), 400

    suffix = "." + file.filename.rsplit(".", 1)[1].lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        result = md.convert(tmp_path)
        return jsonify({"markdown": result.text_content, "title": result.title or file.filename})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)

def free_port(start=8080):
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start

if __name__ == "__main__":
    os.makedirs("static", exist_ok=True)
    port = free_port()
    print(f"\n  Open: http://localhost:{port}\n")
    webbrowser.open(f"http://localhost:{port}")
    app.run(debug=False, port=port)
