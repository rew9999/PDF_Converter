/**
 * PDF変換ツール - フロントエンドアプリケーション
 */

// ---------- 状態管理 ----------
const state = {
    files: [], // { id, name, size, pageCount }
    converting: false,
};

// ---------- DOM要素 ----------
const $ = (sel) => document.querySelector(sel);
const uploadArea = $("#uploadArea");
const fileInput = $("#fileInput");
const fileList = $("#fileList");
const convertBtn = $("#convertBtn");
const progressSection = $("#progressSection");
const progressBar = $("#progressBar");
const progressText = $("#progressText");
const resultSection = $("#resultSection");
const resultSummary = $("#resultSummary");
const resultFiles = $("#resultFiles");

// ---------- ユーティリティ ----------
function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function getSelectedFormat() {
    return document.querySelector('input[name="format"]:checked').value;
}

function getSelectedDpi() {
    return parseInt(document.querySelector('input[name="dpi"]:checked').value, 10);
}

function updateConvertBtn() {
    convertBtn.disabled = state.files.length === 0 || state.converting;
}

// ---------- ファイルリスト表示 ----------
function renderFileList() {
    if (state.files.length === 0) {
        fileList.innerHTML = "";
        updateConvertBtn();
        return;
    }

    fileList.innerHTML = state.files
        .map(
            (f) => `
        <div class="file-item" data-id="${f.id}">
            <div class="file-info">
                <span class="file-name">${escapeHtml(f.name)}</span>
                <span class="file-meta">${formatSize(f.size)} / ${f.pageCount}ページ</span>
            </div>
            <button class="file-remove" title="削除" data-id="${f.id}">&times;</button>
        </div>
    `
        )
        .join("");

    // 削除ボタンのイベント
    fileList.querySelectorAll(".file-remove").forEach((btn) => {
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            removeFile(btn.dataset.id);
        });
    });

    updateConvertBtn();
}

function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

// ---------- ファイルアップロード ----------
async function uploadFiles(files) {
    for (const file of files) {
        if (!file.name.toLowerCase().endsWith(".pdf")) {
            alert(`${file.name} はPDFファイルではありません`);
            continue;
        }

        const formData = new FormData();
        formData.append("file", file);

        try {
            const res = await fetch("/upload", { method: "POST", body: formData });

            if (!res.ok) {
                const err = await res.json();
                alert(`${file.name}: ${err.detail}`);
                continue;
            }

            const data = await res.json();
            state.files.push({
                id: data.file_id,
                name: data.filename,
                size: data.size,
                pageCount: data.page_count,
            });
        } catch (e) {
            alert(`${file.name} のアップロードに失敗しました`);
        }
    }

    renderFileList();
}

async function removeFile(fileId) {
    try {
        await fetch(`/upload/${fileId}`, { method: "DELETE" });
    } catch {
        // サーバー側で消えていても問題なし
    }
    state.files = state.files.filter((f) => f.id !== fileId);
    renderFileList();
}

// ---------- 変換実行 ----------
async function startConversion() {
    if (state.files.length === 0 || state.converting) return;

    state.converting = true;
    updateConvertBtn();

    // UI更新
    const btnText = convertBtn.querySelector(".convert-btn-text");
    const btnSpinner = convertBtn.querySelector(".convert-btn-spinner");
    btnText.textContent = "変換中...";
    btnSpinner.hidden = false;

    progressSection.hidden = false;
    resultSection.hidden = true;
    progressBar.style.width = "0%";
    progressText.textContent = "変換を開始しています...";

    try {
        // 変換リクエスト
        const res = await fetch("/convert", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                file_ids: state.files.map((f) => f.id),
                format: getSelectedFormat(),
                dpi: getSelectedDpi(),
            }),
        });

        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail);
        }

        const { job_id } = await res.json();

        // ポーリングで進捗確認
        await pollStatus(job_id);
    } catch (e) {
        alert("変換エラー: " + e.message);
        progressSection.hidden = true;
    } finally {
        state.converting = false;
        state.files = [];
        renderFileList();
        btnText.textContent = "変換開始";
        btnSpinner.hidden = true;
        updateConvertBtn();
    }
}

async function pollStatus(jobId) {
    while (true) {
        try {
            const res = await fetch(`/status/${jobId}`);
            const data = await res.json();

            // プログレスバー更新
            progressBar.style.width = data.progress + "%";

            if (data.status === "processing") {
                progressText.textContent = `${data.current_file} を変換中... (${data.completed}/${data.total}ファイル)`;
                await sleep(300);
            } else {
                // 完了
                showResult(data);
                return;
            }
        } catch {
            await sleep(500);
        }
    }
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

// ---------- 結果表示 ----------
function showResult(data) {
    progressSection.hidden = true;
    resultSection.hidden = false;

    const totalFiles = data.output_files ? data.output_files.length : 0;
    const errorCount = data.errors ? data.errors.length : 0;

    if (errorCount === 0) {
        resultSummary.className = "result-summary success";
        resultSummary.textContent = `変換が完了しました（${totalFiles}ファイル生成）`;
    } else if (totalFiles > 0) {
        resultSummary.className = "result-summary partial";
        resultSummary.textContent = `一部完了: ${totalFiles}ファイル成功 / ${errorCount}件エラー`;
    } else {
        resultSummary.className = "result-summary error";
        resultSummary.textContent = `変換に失敗しました（${errorCount}件のエラー）`;
    }

    let html = "";

    if (data.output_files && data.output_files.length > 0) {
        // ファイル名のみ表示（最大20件）
        const displayFiles = data.output_files.slice(0, 20);
        html += displayFiles
            .map((f) => {
                const name = f.split(/[/\\]/).pop();
                return `<div class="result-file-item">${escapeHtml(name)}</div>`;
            })
            .join("");

        if (data.output_files.length > 20) {
            html += `<div class="result-file-item">...他 ${data.output_files.length - 20} ファイル</div>`;
        }
    }

    if (data.errors && data.errors.length > 0) {
        html += data.errors
            .map(
                (e) =>
                    `<div class="result-file-item" style="color:#dc2626">${escapeHtml(e.file)}: ${escapeHtml(e.error)}</div>`
            )
            .join("");
    }

    if (data.output_dir) {
        html += `<div class="result-output-dir"><strong>出力先:</strong> ${escapeHtml(data.output_dir)}</div>`;
    }

    resultFiles.innerHTML = html;
}

// ---------- イベントリスナー ----------

// ファイル選択
fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
        uploadFiles(Array.from(e.target.files));
        e.target.value = "";
    }
});

// ドラッグ＆ドロップ
uploadArea.addEventListener("click", () => fileInput.click());

uploadArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadArea.classList.add("drag-over");
});

uploadArea.addEventListener("dragleave", () => {
    uploadArea.classList.remove("drag-over");
});

uploadArea.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadArea.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) {
        uploadFiles(Array.from(e.dataTransfer.files));
    }
});

// 変換ボタン
convertBtn.addEventListener("click", startConversion);
