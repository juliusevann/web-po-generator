
import os
import re
from uuid import uuid4

from flask import (
    Blueprint,
    abort,
    current_app,
    render_template,
    request,
    send_file,
)
from werkzeug.utils import secure_filename

from promo_result.generate_promo_result import run_promo_result


promo_result_bp = Blueprint(
    "promo_result",
    __name__,
    url_prefix="/promo-result",
)

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}


def get_output_files(output_dir):
    """Ambil hanya file Excel hasil yang ada di folder output."""
    if not os.path.isdir(output_dir):
        return []

    return sorted(
        os.path.join(output_dir, name)
        for name in os.listdir(output_dir)
        if name.lower().endswith(".xlsx")
        and os.path.isfile(os.path.join(output_dir, name))
    )


@promo_result_bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("promo_result.html")

    uploaded = request.files.get("input_file")

    if uploaded is None or not uploaded.filename:
        return render_template(
            "promo_result.html",
            error="Silakan pilih file Excel input.",
        ), 400

    filename = secure_filename(uploaded.filename)
    extension = os.path.splitext(filename)[1].lower()

    if extension not in ALLOWED_EXTENSIONS:
        return render_template(
            "promo_result.html",
            error="Format file harus .xlsx atau .xlsm.",
        ), 400

    job_id = uuid4().hex
    job_root = os.path.abspath(
        os.path.join(
            current_app.config["PROMO_RESULT_OUTPUT_FOLDER"],
            job_id,
        )
    )
    input_dir = os.path.join(job_root, "input")
    output_dir = os.path.join(job_root, "output")

    os.makedirs(input_dir, exist_ok=False)
    os.makedirs(output_dir, exist_ok=False)

    input_path = os.path.join(input_dir, filename)
    uploaded.save(input_path)

    try:
        generated_files = run_promo_result(input_path, output_dir)

        # Gunakan urutan yang sama untuk tampilan dan endpoint download.
        generated_files = sorted(
            os.path.abspath(path) for path in generated_files
        )

        # Pastikan seluruh hasil benar-benar berada di output job ini.
        safe_results = []
        for path in generated_files:
            if (
                os.path.commonpath([output_dir, path]) != output_dir
                or not os.path.isfile(path)
                or not path.lower().endswith(".xlsx")
            ):
                raise ValueError("Lokasi file output tidak valid.")

            safe_results.append({
                "name": os.path.basename(path),
                "url": (
                    f"/promo-result/download/{job_id}/"
                    f"{len(safe_results)}"
                ),
            })

        if not safe_results:
            raise RuntimeError("Generator tidak menghasilkan file Excel.")

        return render_template(
            "promo_result.html",
            results=safe_results,
        )

    except Exception:
        current_app.logger.exception(
            "Promo Result gagal diproses."
        )
        return render_template(
            "promo_result.html",
            error=(
                "Promo Result gagal diproses. "
                "Periksa log aplikasi untuk detail error."
            ),
        ), 500


@promo_result_bp.route(
    "/download/<job_id>/<int:index>",
    methods=["GET"],
)
def download(job_id, index):
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        abort(404)

    if index < 0:
        abort(404)

    output_root = os.path.abspath(
        current_app.config["PROMO_RESULT_OUTPUT_FOLDER"]
    )
    job_root = os.path.abspath(
        os.path.join(output_root, job_id)
    )
    output_dir = os.path.join(job_root, "output")

    # Pastikan job tetap berada di bawah direktori output.
    if os.path.commonpath([output_root, job_root]) != output_root:
        abort(404)

    files = get_output_files(output_dir)

    if index >= len(files):
        abort(404)

    path = files[index]

    return send_file(
        path,
        as_attachment=True,
        download_name=os.path.basename(path),
    )
