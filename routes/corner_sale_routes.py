import os
import sys
import time
import uuid
import zipfile
import subprocess
from pathlib import Path

from flask import (
    Blueprint, render_template, send_file, abort, current_app
)
from werkzeug.utils import secure_filename
from config import Config

corner_sale_bp = Blueprint(
    "corner_sale", __name__, url_prefix="/corner-sale"
)


@corner_sale_bp.route("/")
def index():
    return render_template("corner_sale.html")


@corner_sale_bp.route("/run", methods=["POST"])
def run():
    script = Path(Config.CORNER_SALE_SCRIPT)
    output_dir = Path(Config.CORNER_SALE_OUTPUT_FOLDER)

    if not script.is_file():
        return render_template(
            "corner_sale.html",
            error=f"Script tidak ditemukan: {script}"
        ), 500

    output_dir.mkdir(parents=True, exist_ok=True)

    def snapshot():
        result = {}
        for file in output_dir.iterdir():
            if file.is_file() and file.suffix.lower() in {
                ".xlsx", ".xlsm"
            }:
                result[str(file.resolve())] = file.stat().st_mtime_ns
        return result

    before = snapshot()

    try:
        process = subprocess.run(
            [sys.executable, str(script)],
            cwd=str(script.parent),
            capture_output=True,
            text=True,
            timeout=1800,
            check=False
        )
    except subprocess.TimeoutExpired:
        return render_template(
            "corner_sale.html",
            error="Proses melebihi batas waktu 30 menit."
        ), 504
    except Exception:
        current_app.logger.exception(
            "Gagal menjalankan Corner Sale"
        )
        return render_template(
            "corner_sale.html",
            error="Gagal menjalankan script Corner Sale."
        ), 500

    after = snapshot()

    changed = [
        Path(path)
        for path, modified in after.items()
        if path not in before or before[path] != modified
    ]

    if process.returncode != 0:
        details = (process.stderr or process.stdout or "")[-8000:]
        return render_template(
            "corner_sale.html",
            error="Script Corner Sale gagal. Periksa detail log.",
            log=details
        ), 500

    if not changed:
        return render_template(
            "corner_sale.html",
            error=(
                "Script selesai, tetapi tidak ditemukan file Excel "
                "baru atau yang diperbarui di folder output."
            ),
            log=(process.stdout or "")[-8000:]
        )

    zip_name = (
        f"corner_sale_{time.strftime('%Y%m%d_%H%M%S')}_"
        f"{uuid.uuid4().hex[:6]}.zip"
    )
    zip_path = output_dir / zip_name

    with zipfile.ZipFile(
        zip_path, "w", zipfile.ZIP_DEFLATED
    ) as archive:
        for file in changed:
            if file.is_file():
                archive.write(file, arcname=file.name)

    return render_template(
        "corner_sale.html",
        success="Proses Corner Sale berhasil.",
        download_name=zip_name,
        output_names=[f.name for f in changed],
        log=(process.stdout or "")[-8000:]
    )


@corner_sale_bp.route("/download/<filename>")
def download(filename):
    filename = secure_filename(filename)

    if not filename.endswith(".zip"):
        abort(404)

    output_dir = Path(
        Config.CORNER_SALE_OUTPUT_FOLDER
    ).resolve()
    file_path = (output_dir / filename).resolve()

    if file_path.parent != output_dir or not file_path.is_file():
        abort(404)

    return send_file(
        file_path,
        as_attachment=True,
        download_name=filename
    )
