const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const dzFilename = document.getElementById("dz-filename");
const btnProcess = document.getElementById("btn-process");
const uploadError = document.getElementById("upload-error");

const stageUpload = document.getElementById("stage-upload");
const stageProgress = document.getElementById("stage-progress");
const stageDone = document.getElementById("stage-done");

const progressFill = document.getElementById("progress-fill");
const progressCount = document.getElementById("progress-count");
const progressCurrent = document.getElementById("progress-current");
const progressSubtitle = document.getElementById("progress-subtitle");
const rowList = document.getElementById("row-list");
const rowListDone = document.getElementById("row-list-done");

const doneTitle = document.getElementById("done-title");
const doneSummary = document.getElementById("done-summary");
const btnDownload = document.getElementById("btn-download");
const btnReset = document.getElementById("btn-reset");
const failedLogLine = document.getElementById("failed-log-line");

let selectedFile = null;
let pollTimer = null;

function showStage(stage) {
  [stageUpload, stageProgress, stageDone].forEach((s) =>
    s.classList.add("hidden"),
  );
  stage.classList.remove("hidden");
}

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("drag-over");
});
dropzone.addEventListener("dragleave", () =>
  dropzone.classList.remove("drag-over"),
);
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("drag-over");
  if (e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
    handleFileSelect();
  }
});
fileInput.addEventListener("change", handleFileSelect);

function handleFileSelect() {
  const file = fileInput.files[0];
  uploadError.textContent = "";
  if (!file) return;
  selectedFile = file;
  dzFilename.textContent = file.name;
  btnProcess.disabled = false;
}

btnProcess.addEventListener("click", async () => {
  if (!selectedFile) return;
  uploadError.textContent = "";
  btnProcess.disabled = true;
  btnProcess.textContent = "Mengupload...";

  const formData = new FormData();
  formData.append("file", selectedFile);

  try {
    const res = await fetch("/po/upload", { method: "POST", body: formData });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Gagal upload file.");
    }

    startPolling(data.job_id);
    showStage(stageProgress);
  } catch (err) {
    uploadError.textContent = err.message;
    btnProcess.disabled = false;
    btnProcess.textContent = "Proses File";
  }
});

function startPolling(jobId) {
  poll(jobId);
  pollTimer = setInterval(() => poll(jobId), 1500);
}

async function poll(jobId) {
  try {
    const res = await fetch(`/po/status/${jobId}`);
    const data = await res.json();

    if (!res.ok) {
      clearInterval(pollTimer);
      progressSubtitle.textContent = data.error || "Terjadi kesalahan.";
      return;
    }

    renderProgress(data, jobId);

    if (data.status === "done") {
      clearInterval(pollTimer);
      renderDone(data, jobId);
    } else if (data.status === "error") {
      clearInterval(pollTimer);
      progressSubtitle.textContent =
        "Gagal: " + (data.error || "Terjadi kesalahan tidak diketahui.");
    }
  } catch (err) {
  }
}

function renderProgress(data, jobId) {
  const pct = data.total ? Math.round((data.completed / data.total) * 100) : 0;
  progressFill.style.width = pct + "%";
  progressCount.textContent = `${data.completed} / ${data.total}`;

  progressSubtitle.textContent = "Sedang memproses...";
  progressCurrent.textContent = data.current_po_code
    ? `PO ${data.current_po_code}`
    : "";

  rowList.innerHTML = "";
  data.results.forEach((r) => renderRow(rowList, r));
}

function renderRow(container, r) {
  const wrap = document.createElement("div");
  const ok = r.excel_ok && r.pdf_ok;
  wrap.innerHTML = `
    <div class="row-item">
      <span class="row-status ${ok ? "ok" : "fail"}">${ok ? "✓" : "✕"}</span>
      <span class="row-po">${escapeHtml(r.po_code)}</span>
      <span class="row-ref">${escapeHtml(r.ref)}</span>
    </div>
    ${!ok ? `<div class="row-err">${escapeHtml(r.error || "")}</div>` : ""}
  `;
  container.appendChild(wrap);
}

function renderDone(data, jobId) {
  showStage(stageDone);
  const failCount = data.results.filter(
    (r) => !(r.excel_ok && r.pdf_ok),
  ).length;
  const okCount = data.results.length - failCount;

  doneTitle.textContent =
    failCount === 0
      ? "Semua PO berhasil diproses"
      : "Selesai dengan beberapa PO gagal";
  doneSummary.textContent = `${okCount} berhasil, ${failCount} gagal dari total ${data.results.length} baris.`;
  failedLogLine.style.display = failCount > 0 ? "block" : "none";

  btnDownload.href = `/po/download/${jobId}`;

  rowListDone.innerHTML = "";
  data.results.forEach((r) => renderRow(rowListDone, r));
}

btnReset.addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  dzFilename.textContent = "";
  btnProcess.disabled = true;
  btnProcess.textContent = "Proses File";
  uploadError.textContent = "";
  showStage(stageUpload);
});

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
