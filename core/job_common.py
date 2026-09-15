import json
import os
import uuid


def new_job_id():
    return uuid.uuid4().hex[:12]


def status_path(output_dir):
    return os.path.join(output_dir, "status.json")


def write_json_atomic(path, data):
    tmp_path = f"{path}.{os.getpid()}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, default=str)
    os.replace(tmp_path, path)


def read_json(path, default=None):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default
