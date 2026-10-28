import os
import sys
import time
import uuid
import zipfile
import subprocess
import tempfile
from pathlib import Path

from flask import (
    Blueprint,
    render_template,
    send_file,
    abort,
    current_app,
    request,
)
from werkzeug.utils import secure_filename
from config import Config


corner_sale_bp = Blueprint(
    "corner_sale", __name__, url_prefix="/corner-sale"
)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


def allowed_excel(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@corner_sale_bp.route("/")
def index():
    return render_template("corner_sale.html")


@corner_sale_bp.route("/run", methods=["POST"])
def run():
    template_file = request.files.get("template_file")
    input_file = request.files.get("input_file")

    if not template_file or not template_file.filename:
        return render_template(
            "corner_sale.html",
            error="Silakan pilih file template Excel terlebih dahulu.",
        ), 400

    if not input_file or not input_file.filename:
        return render_template(
            "corner_sale.html",
            error="Silakan pilih file input promosi terlebih dahulu.",
        ), 400

    if not allowed_excel(template_file.filename):
        return render_template(
            "corner_sale.html",
            error="Template harus berformat .xlsx atau .xlsm.",
        ), 400

    if not allowed_excel(input_file.filename):
        return render_template(
            "corner_sale.html",
            error="File input harus berformat .xlsx atau .xlsm.",
        ), 400

    script = Path(Config.CORNER_SALE_SCRIPT).resolve()
    output_dir = Path(Config.CORNER_SALE_OUTPUT_FOLDER).resolve()

    if not script.is_file():
        return render_template(
            "corner_sale.html",
            error=f"Script tidak ditemukan: {script}",
        ), 500

    output_dir.mkdir(parents=True, exist_ok=True)

    template_name = secure_filename(template_file.filename)
    input_name = secure_filename(input_file.filename)

    if not template_name or not input_name:
        return render_template(
            "corner_sale.html",
            error="Nama salah satu file tidak valid.",
        ), 400

    job_id = uuid.uuid4().hex[:10]
    zip_name = (
        f"corner_sale_{time.strftime('%Y%m%d_%H%M%S')}_"
        f"{job_id}.zip"
    )
    zip_path = output_dir / zip_name

    try:
        # Folder unik per request agar file antar-user tidak tercampur.
        with tempfile.TemporaryDirectory(
            prefix=f"corner_sale_{job_id}_",
            dir=str(output_dir),
        ) as temp_dir:
            work_dir = Path(temp_dir)
            uploaded_template = work_dir / template_name
            uploaded_input = work_dir / input_name
            result_dir = work_dir / "result"
            result_dir.mkdir()

            template_file.save(str(uploaded_template))
            input_file.save(str(uploaded_input))

            process = subprocess.run(
                [
                    sys.executable,
                    str(script),
                    "--template-file",
                    str(uploaded_template),
                    "--input-file",
                    str(uploaded_input),
                    "--output-dir",
                    str(result_dir),
                ],
                cwd=str(script.parent),
                capture_output=True,
                text=True,
                timeout=1800,
                check=False,
            )

            log = (
                (process.stdout or "")
                + "\n"
                + (process.stderr or "")
            )[-12000:]

            if process.returncode != 0:
                return render_template(
                    "corner_sale.html",
                    error=(
                        "Proses Corner Sale gagal. "
                        "Periksa log untuk detailnya."
                    ),
                    log=log,
                ), 500

            generated_files = [
                path for path in result_dir.iterdir()
                if path.is_file()
                and path.suffix.lower() in {".xlsx", ".xlsm"}
                and not path.name.startswith("~$")
            ]

            if not generated_files:
                return render_template(
                    "corner_sale.html",
                    error=(
                        "Script selesai, tetapi tidak menghasilkan "
                        "file Excel."
                    ),
                    log=log,
                ), 500

            # ZIP disimpan di output_dir, di luar folder sementara.
            with zipfile.ZipFile(
                zip_path, "w", zipfile.ZIP_DEFLATED
            ) as archive:
                for file in generated_files:
                    archive.write(file, arcname=file.name)

            output_names = [f.name for f in generated_files]
            
        return render_template(
            "corner_sale.html",
            success="Proses Corner Sale berhasil.",
            download_name=zip_name,
            output_names=output_names,
            log=log,
        )
        
    except subprocess.TimeoutExpired:
        return render_template(
            "corner_sale.html",
            error="Proses melebihi batas waktu 30 menit.",
        ), 504

    except Exception:
        current_app.logger.exception(
            "Gagal menjalankan Corner Sale"
        )

        # Hapus ZIP parsial jika proses gagal saat menulis file.
        try:
            if zip_path.exists():
                zip_path.unlink()
        except OSError:
            current_app.logger.exception(
                "Gagal membersihkan ZIP parsial Corner Sale"
            )

        return render_template(
            "corner_sale.html",
            error=(
                "Terjadi kesalahan saat memproses Corner Sale. "
                "Periksa log aplikasi."
            ),
        ), 500


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
        download_name=filename,
    )
