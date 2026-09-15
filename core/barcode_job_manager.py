import os
import shutil
import threading
import time
import zipfile
import multiprocessing
from datetime import datetime, timedelta

from openpyxl import Workbook

from config import Config
from core.job_common import new_job_id, status_path, write_json_atomic, read_json

_workers_started = False
_init_lock = threading.Lock()

POLL_INTERVAL_SECONDS = 1


def create_job(input_path, original_filename):
    """
    Hitung dulu berapa grup di file yang diupload (cepat, tanpa DB/foto),
    buat job baru (status "queued"), lalu pastikan worker jalan.
    Return job_id.
    """
    cleanup_old_outputs()

    from core.barcode_generator import count_groups
    total = count_groups(input_path)

    if total == 0:
        raise ValueError(
            "File tidak menghasilkan grup apapun. Pastikan file punya baris "
            "data yang valid (kolom PO, item code, dan store terisi)."
        )

    job_id = new_job_id()
    job_output_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, job_id)
    os.makedirs(job_output_dir, exist_ok=True)

    status = {
        "id": job_id,
        "original_filename": original_filename,
        "input_path": input_path,
        "created_at": datetime.now().isoformat(),
        "status": "queued",
        "total": total,
        "completed": 0,
        "current_group": None,
        "results": [],
        "zip_path": None,
        "error": None,
    }
    write_json_atomic(status_path(job_output_dir), status)

    print(f"[BARCODE_JOB_MANAGER] Job {job_id} dibuat, status queued ({total} grup)", flush=True)
    _ensure_workers_started()
    return job_id


def get_job(job_id):
    output_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, job_id)
    return read_json(status_path(output_dir))


def get_queue_position(job_id):
    """Sama logic-nya seperti punya PO: posisi di antara job lain yang masih 'queued'."""
    this_status = get_job(job_id)
    if this_status is None or this_status.get("status") != "queued":
        return 0

    this_created_at = this_status.get("created_at", "")
    queued_before = 0

    if not os.path.isdir(Config.BARCODE_OUTPUT_FOLDER):
        return 1

    for name in os.listdir(Config.BARCODE_OUTPUT_FOLDER):
        job_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, name)
        if not os.path.isdir(job_dir):
            continue
        status = read_json(status_path(job_dir))
        if not status or status.get("status") != "queued":
            continue
        if status.get("created_at", "") < this_created_at:
            queued_before += 1

    return queued_before + 1


def _ensure_workers_started():
    global _workers_started
    with _init_lock:
        if _workers_started:
            return
        n = max(1, Config.BARCODE_MAX_CONCURRENT_JOBS)
        print(f"[BARCODE_JOB_MANAGER] Menjalankan {n} worker process (mode polling filesystem)...", flush=True)
        for i in range(n):
            p = multiprocessing.Process(target=_worker_loop, args=(i + 1,), daemon=True)
            p.start()
            print(f"[BARCODE_JOB_MANAGER] Worker #{i+1} dimulai, PID={p.pid}", flush=True)
        _workers_started = True


def _worker_loop(worker_id):

    print(f"[BARCODE_WORKER #{worker_id} pid={os.getpid()}] Worker aktif, polling job...", flush=True)

    while True:
        job_id = _claim_next_queued_job(worker_id)
        if job_id is None:
            time.sleep(POLL_INTERVAL_SECONDS)
            continue

        print(f"[BARCODE_WORKER #{worker_id} pid={os.getpid()}] Mengklaim & memproses job {job_id}", flush=True)
        try:
            _process_job(job_id)
            print(f"[BARCODE_WORKER #{worker_id} pid={os.getpid()}] Selesai job {job_id}", flush=True)
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            print(f"[BARCODE_WORKER #{worker_id} pid={os.getpid()}] ERROR job {job_id}:\n{tb}", flush=True)
            output_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, job_id)
            status = read_json(status_path(output_dir)) or {}
            status["status"] = "error"
            status["error"] = str(e)
            write_json_atomic(status_path(output_dir), status)


def _claim_next_queued_job(worker_id):
    if not os.path.isdir(Config.BARCODE_OUTPUT_FOLDER):
        return None

    candidates = []
    for name in os.listdir(Config.BARCODE_OUTPUT_FOLDER):
        job_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, name)
        if not os.path.isdir(job_dir):
            continue
        status = read_json(status_path(job_dir))
        if status and status.get("status") == "queued":
            candidates.append((status.get("created_at", ""), name))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])

    for _, job_id in candidates:
        output_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, job_id)
        claim_lock_path = os.path.join(output_dir, "claim.lock")

        try:
            fd = os.open(claim_lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(worker_id).encode())
            os.close(fd)
        except FileExistsError:
            continue

        status = read_json(status_path(output_dir))
        if status is None or status.get("status") != "queued":
            continue

        status["status"] = "processing"
        status["claimed_by_worker"] = worker_id
        write_json_atomic(status_path(output_dir), status)
        return job_id

    return None


def _process_job(job_id):
    # Import di dalam fungsi (bukan top-level), konsisten dengan pola PO -
    # supaya pyodbc/xlrd cuma ke-load di proses WORKER.
    from core.barcode_generator import process_barcode_excel

    output_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, job_id)
    status = read_json(status_path(output_dir))
    if status is None:
        return

    input_path = status["input_path"]
    barcode_dir = os.path.join(output_dir, "BARCODE")
    os.makedirs(barcode_dir, exist_ok=True)

    def on_group_done(result):
        status["current_group"] = result["group"]
        status["results"].append(result)
        status["completed"] += 1
        write_json_atomic(status_path(output_dir), status)

    results = process_barcode_excel(input_path, barcode_dir, on_group_done=on_group_done)

    failed = [r for r in results if not r["ok"]]
    if failed:
        wb = Workbook()
        ws = wb.active
        ws.title = "FAILED"
        ws.append(["GROUP", "ERROR"])
        for r in failed:
            ws.append([r["group"], r["error"]])
        wb.save(os.path.join(output_dir, "FAILED_LOG.xlsx"))

    zip_path = output_dir + ".zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                if file in ("status.json", "claim.lock"):
                    continue
                file_path = os.path.join(root, file)
                arcname = os.path.join("output", os.path.relpath(file_path, output_dir))
                zf.write(file_path, arcname)

    status["current_group"] = None
    status["zip_path"] = zip_path
    status["status"] = "done"
    write_json_atomic(status_path(output_dir), status)

    try:
        os.remove(input_path)
    except Exception:
        pass
    
def cleanup_old_outputs():
    cutoff = datetime.now() - timedelta(hours=Config.RETENTION_HOURS)
    if not os.path.isdir(Config.BARCODE_OUTPUT_FOLDER):
        return

    for name in os.listdir(Config.BARCODE_OUTPUT_FOLDER):
        job_dir = os.path.join(Config.BARCODE_OUTPUT_FOLDER, name)
        if not os.path.isdir(job_dir):
            continue

        status = read_json(status_path(job_dir))
        created_at_str = status.get("created_at") if status else None

        try:
            created_at = datetime.fromisoformat(created_at_str) if created_at_str else None
        except Exception:
            created_at = None

        if created_at is None or created_at >= cutoff:
            continue

        try:
            shutil.rmtree(job_dir, ignore_errors=True)
            zip_path = job_dir + ".zip"
            if os.path.exists(zip_path):
                os.remove(zip_path)
        except Exception:
            pass
