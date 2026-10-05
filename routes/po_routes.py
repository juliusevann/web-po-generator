import os
import time
import uuid

from flask import Blueprint, render_template, request, jsonify, send_file, abort
from werkzeug.utils import secure_filename

from config import Config
from core import job_manager

po_bp = Blueprint("po", __name__, url_prefix="/po")

ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def _allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@po_bp.route("/")
def index():
    return render_template("po.html")


@po_bp.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file yang dikirim."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Tidak ada file yang dipilih."}), 400

    if not _allowed_file(file.filename):
        return jsonify({"error": "Format file harus .xlsx atau .xls"}), 400

    filename = secure_filename(file.filename)
    unique_name = f"{int(time.time() * 1000)}_{uuid.uuid4().hex[:8]}_{filename}"
    save_path = os.path.join(Config.UPLOAD_FOLDER, unique_name)
    file.save(save_path)

    try:
        job_id = job_manager.create_job(save_path, filename)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membaca file: {e}"}), 400

    return jsonify({"job_id": job_id})


@po_bp.route("/status/<job_id>")
def status(job_id):
    job = job_manager.get_job(job_id)
    if job is None:
        return jsonify({"error": "Job tidak ditemukan"}), 404

    queue_position = job_manager.get_queue_position(job_id)

    return jsonify({
        "status": job["status"],
        "total": job["total"],
        "completed": job["completed"],
        "current_po_code": job["current_po_code"],
        "queue_position": queue_position,
        "results": job["results"],
        "error": job["error"],
        "download_ready": job["status"] == "done",
    })


@po_bp.route("/download/<job_id>")
def download(job_id):
    job = job_manager.get_job(job_id)
    if job is None or job["status"] != "done" or not job["zip_path"]:
        abort(404)

    return send_file(job["zip_path"], as_attachment=True, download_name="output.zip")
