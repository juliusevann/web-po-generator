import os
import shutil
import threading
import time
import zipfile
import multiprocessing
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook

from config import Config
from core.job_common import new_job_id, status_path, write_json_atomic, read_json

_workers_started = False
_init_lock = threading.Lock()

POLL_INTERVAL_SECONDS = 1


def _rows_path(output_dir):
    return os.path.join(output_dir, "rows.json")


def create_job(excel_path, original_filename):
    """
    Baca file excel upload user (kolom KODE_PO, REF), buat job baru
    (status "queued" di status.json), lalu pastikan worker jalan.
    Return job_id.
    """
    cleanup_old_outputs()

    df = pd.read_excel(excel_path)

    required_cols = {"KODE_PO", "REF"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(
            f"Kolom wajib tidak ditemukan di file excel: {', '.join(missing)}. "
            f"Pastikan file punya kolom KODE_PO dan REF."
        )

    rows = []
    for _, row in df.iterrows():
        po_code = row.get("KODE_PO")
        ref = row.get("REF")
        if pd.isna(po_code) or pd.isna(ref):
            continue
        rows.append({"po_code": str(po_code).strip(), "ref": str(ref).strip()})

    if not rows:
        raise ValueError("File excel tidak berisi baris data yang valid.")

    job_id = new_job_id()
    job_output_dir = os.path.join(Config.OUTPUT_FOLDER, job_id)
    os.makedirs(os.path.join(job_output_dir, "PDF"), exist_ok=True)
    os.makedirs(os.path.join(job_output_dir, "PO"), exist_ok=True)

    write_json_atomic(_rows_path(job_output_dir), rows)

    status = {
        "id": job_id,
        "original_filename": original_filename,
        "created_at": datetime.now().isoformat(),
        "status": "queued",
        "total": len(rows),
        "completed": 0,
        "current_po_code": None,
        "results": [],
        "zip_path": None,
        "error": None,
    }
    write_json_atomic(status_path(job_output_dir), status)

    print(f"[JOB_MANAGER] Job {job_id} dibuat, status queued ({len(rows)} baris)", flush=True)
    _ensure_workers_started()
    return job_id


def get_job(job_id):
    output_dir = os.path.join(Config.OUTPUT_FOLDER, job_id)
    return read_json(status_path(output_dir))


def get_queue_position(job_id):
    """
    Posisi job yang MASIH MENUNGGU (status "queued"), dihitung dari urutan
    created_at di antara job lain yang juga masih "queued". 0 kalau job ini
    sudah dapat worker (statusnya bukan "queued" lagi).
    """
    this_status = get_job(job_id)
    if this_status is None or this_status.get("status") != "queued":
        return 0

    this_created_at = this_status.get("created_at", "")
    queued_before = 0

    if not os.path.isdir(Config.OUTPUT_FOLDER):
        return 1

    for name in os.listdir(Config.OUTPUT_FOLDER):
        job_dir = os.path.join(Config.OUTPUT_FOLDER, name)
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
        n = max(1, Config.MAX_CONCURRENT_JOBS)
        print(f"[JOB_MANAGER] Menjalankan {n} worker process (mode polling filesystem)...", flush=True)
        for i in range(n):
            p = multiprocessing.Process(target=_worker_loop, args=(i + 1,), daemon=True)
            p.start()
            print(f"[JOB_MANAGER] Worker #{i+1} dimulai, PID={p.pid}", flush=True)
        _workers_started = True


def _worker_loop(worker_id):
    # Tiap worker jalan di PROSES OS sendiri (bukan thread, bukan pakai
    # Manager antar-proses) - polling filesystem tiap POLL_INTERVAL_SECONDS.
    print(f"[WORKER #{worker_id} pid={os.getpid()}] Worker aktif, polling job...", flush=True)
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception as e:
        print(f"[WORKER #{worker_id} pid={os.getpid()}] CoInitialize gagal (mungkin bukan Windows): {e}", flush=True)

    while True:
        job_id = _claim_next_queued_job(worker_id)
        if job_id is None:
            time.sleep(POLL_INTERVAL_SECONDS)
            continue

        print(f"[WORKER #{worker_id} pid={os.getpid()}] Mengklaim & memproses job {job_id}", flush=True)
        try:
            _process_job(job_id)
            print(f"[WORKER #{worker_id} pid={os.getpid()}] Selesai job {job_id}", flush=True)
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            print(f"[WORKER #{worker_id} pid={os.getpid()}] ERROR job {job_id}:\n{tb}", flush=True)
            output_dir = os.path.join(Config.OUTPUT_FOLDER, job_id)
            status = read_json(status_path(output_dir)) or {}
            status["status"] = "error"
            status["error"] = str(e)
            write_json_atomic(status_path(output_dir), status)


def _claim_next_queued_job(worker_id):
    if not os.path.isdir(Config.OUTPUT_FOLDER):
        return None

    candidates = []
    for name in os.listdir(Config.OUTPUT_FOLDER):
        job_dir = os.path.join(Config.OUTPUT_FOLDER, name)
        if not os.path.isdir(job_dir):
            continue
        status = read_json(status_path(job_dir))
        if status and status.get("status") == "queued":
            candidates.append((status.get("created_at", ""), name))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])

    for _, job_id in candidates:
        output_dir = os.path.join(Config.OUTPUT_FOLDER, job_id)
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
    # Import di dalam fungsi (bukan top-level) supaya xlwings/win32com cuma
    # ke-load di proses WORKER, tidak ikut ke-load di proses Flask utama.
    from core.excel_generator import generate_po_excel
    from core.pdf_generator import generate_po_pdf

    output_dir = os.path.join(Config.OUTPUT_FOLDER, job_id)
    rows = read_json(_rows_path(output_dir), default=[])
    status = read_json(status_path(output_dir))
    if status is None:
        return

    pdf_dir = os.path.join(output_dir, "PDF")
    excel_dir = os.path.join(output_dir, "PO")

    failed_rows = []

    for row in rows:
        po_code = row["po_code"]
        ref = row["ref"]
        status["current_po_code"] = po_code
        write_json_atomic(status_path(output_dir), status)

        excel_ok, excel_path, excel_err = generate_po_excel(po_code, ref, excel_dir)
        pdf_ok, pdf_path, pdf_err = generate_po_pdf(po_code, ref, pdf_dir)

        result = {
            "po_code": po_code,
            "ref": ref,
            "excel_ok": excel_ok,
            "pdf_ok": pdf_ok,
            "error": None,
        }

        errors = []
        if not excel_ok:
            errors.append(f"Excel: {excel_err}")
        if not pdf_ok:
            errors.append(f"PDF: {pdf_err}")

        if errors:
            result["error"] = " | ".join(errors)
            failed_rows.append((po_code, ref, result["error"]))

        status["results"].append(result)
        status["completed"] += 1
        write_json_atomic(status_path(output_dir), status)

    if failed_rows:
        wb = Workbook()
        ws = wb.active
        ws.title = "FAILED"
        ws.append(["KODE_PO", "REF", "ERROR"])
        for po_code, ref, err in failed_rows:
            ws.append([po_code, ref, err])
        wb.save(os.path.join(output_dir, "FAILED_LOG.xlsx"))

    zip_path = output_dir + ".zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.join(
                    "output", os.path.relpath(file_path, output_dir)
                )
                zf.write(file_path, arcname)

    status["zip_path"] = zip_path
    status["current_po_code"] = None
    status["status"] = "done"
    write_json_atomic(status_path(output_dir), status)


def cleanup_old_outputs():
    cutoff = datetime.now() - timedelta(hours=Config.RETENTION_HOURS)
    if not os.path.isdir(Config.OUTPUT_FOLDER):
        return

    for name in os.listdir(Config.OUTPUT_FOLDER):
        job_dir = os.path.join(Config.OUTPUT_FOLDER, name)
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
