import os
import time
import uuid

from flask import Blueprint, render_template, request, jsonify, send_file, abort
from werkzeug.utils import secure_filename

from config import Config
from core import barcode_job_manager

barcode_bp = Blueprint("barcode", __name__, url_prefix="/barcode")

# CONCERN: xlrd versi 2.0+ SUDAH TIDAK BISA baca file .xlsx, cuma .xls.
# Kalau server ini pakai xlrd >= 2.0 dan user upload .xlsx, create_job()
# di bawah akan gagal dengan pesan error yang jelas (bukan crash diam-diam) -
# tapi kalau mau strict dari awal, ubah baris di bawah jadi {"xls"} saja.
ALLOWED_EXTENSIONS = {"xls", "xlsx"}


def _allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@barcode_bp.route("/")
def index():
    return render_template("barcode.html")


@barcode_bp.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file yang dikirim."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Tidak ada file yang dipilih."}), 400

    if not _allowed_file(file.filename):
        return jsonify({"error": "Format file harus .xls atau .xlsx"}), 400

    filename = secure_filename(file.filename)
    unique_name = f"{int(time.time() * 1000)}_{uuid.uuid4().hex[:8]}_{filename}"
    save_path = os.path.join(Config.UPLOAD_FOLDER, unique_name)
    file.save(save_path)

    try:
        job_id = barcode_job_manager.create_job(save_path, filename)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membaca file: {e}"}), 400

    return jsonify({"job_id": job_id})


@barcode_bp.route("/status/<job_id>")
def status(job_id):
    job = barcode_job_manager.get_job(job_id)
    if job is None:
        return jsonify({"error": "Job tidak ditemukan"}), 404

    queue_position = barcode_job_manager.get_queue_position(job_id)

    return jsonify({
        "status": job["status"],
        "total": job["total"],
        "completed": job["completed"],
        "current_group": job["current_group"],
        "queue_position": queue_position,
        "results": job["results"],
        "error": job["error"],
        "download_ready": job["status"] == "done",
    })


@barcode_bp.route("/download/<job_id>")
def download(job_id):
    job = barcode_job_manager.get_job(job_id)
    if job is None or job["status"] != "done" or not job["zip_path"]:
        abort(404)

    return send_file(job["zip_path"], as_attachment=True, download_name="output_barcode.zip")
