from pathlib import Path
import os

from flask import Flask, send_from_directory


BASE_DIR = Path(__file__).parent.resolve()

app = Flask(
    __name__,
    static_folder=str(BASE_DIR),
    static_url_path="",  # Serve static files (including index.html) from the repo root
)


@app.route("/")
def serve_index():
    return app.send_static_file("index.html")


@app.route("/quicken_export.xlsx")
def serve_quicken_workbook():
    return send_from_directory(str(BASE_DIR), "quicken_export.xlsx")


@app.route("/sample_export.xlsx")
def serve_sample_workbook():
    return send_from_directory(str(BASE_DIR), "sample_export.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
