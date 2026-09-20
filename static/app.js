// 全画面共通: トースト、fetch（ragFetch）、1画面の中の段（ragSections）、
// ファイルのドロップ、確認ダイアログ、ジョブ進捗
"use strict";

// ---- トースト -------------------------------------------------------------------
function toast(message, kind = "ok") {
  const el = document.getElementById("toast");
  if (!el) return;
  el.textContent = message;
  el.className = `toast toast-${kind}`;
  el.hidden = false;
  clearTimeout(el._timer);
  el._timer = setTimeout(() => (el.hidden = true), kind === "err" ? 6000 : 3500);
}

// ---- fetch ---------------------------------------------------------------------
// JSON を POST して JSON を受け取る。失敗時はサーバの error 文で例外
async function postJson(url, body, options = {}) {
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify(body || {}),
    ...options,
  });
  if (res.status === 204) return {};
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    // 問題が複数あるとき（列の対応づけの保存など）は errors に全部入っている。1件だけ見せて残りを隠さない
    const error = new Error(data.error || `通信に失敗しました（HTTP ${res.status}）`);
    error.status = res.status;
    error.errors = Array.isArray(data.errors) ? data.errors : null;
    throw error;
  }
  return data;
}

async function getJson(url) {
  const res = await fetch(url, { headers: { Accept: "application/json" } });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    // 呼び出し側が 404（もうそのデータが無い）を通信エラーと区別できるように status を載せる
    const error = new Error(data.error || `通信に失敗しました（HTTP ${res.status}）`);
    error.status = res.status;
    throw error;
  }
  return data;
}

// ---- ジョブの進捗 ------------------------------------------------------------------
// url を一定間隔で取得し onUpdate(job) を呼ぶ。終わった状態になったら止める。戻り値の stop() で中止
const JOB_FINISHED = ["done", "failed", "cancelled", "interrupted"];

function pollJob(url, onUpdate, { interval = 1500, maxInterval = 10000 } = {}) {
  let stopped = false;
  let wait = interval;
  let timer = null;
  const tick = async () => {
    if (stopped) return;
    try {
      const job = await getJson(url);
      wait = interval;
      onUpdate(job);
      if (JOB_FINISHED.includes(job.status)) return;
      if (job.status === "paused") wait = Math.min(interval * 3, maxInterval);
    } catch (e) {
      // ジョブが消えている（ダウンロードや削除でデータごと片付いた）なら、待っても終わらないので画面を読み直す
      if (e.status === 404) {
        stopped = true;
        clearTimeout(timer);
        window.location.reload();
        return;
      }
      wait = Math.min(wait * 2, maxInterval); // 通信エラーは間隔を広げて続ける
    }
    if (!stopped) timer = setTimeout(tick, wait);
  };
  tick();
  return { stop() { stopped = true; clearTimeout(timer); } };
}

const JOB_LABELS = {
  queued: "待ち", running: "実行中", paused: "一時停止", done: "完了",
  failed: "失敗", cancelled: "中止", interrupted: "中断",
};

// components/_ui.html の progress(job, url) を自動で更新する
function bindProgressBox(box) {
  const url = box.dataset.jobUrl;
  if (!url || JOB_FINISHED.includes(box.dataset.jobStatus)) return;
  pollJob(url, (job) => {
    const prog = job.progress || job.progress_json || {};
    const done = Number(prog.done || 0);
    const total = Number(prog.total || 0);
    const pct = total ? Math.floor((done * 100) / total) : 0;
    const bar = box.querySelector(".progress-bar");
    const meter = box.querySelector(".progress");
    if (bar) {
      bar.style.width = `${pct}%`;
      bar.classList.toggle("indeterminate", !total && job.status === "running");
    }
    meter?.setAttribute("aria-valuenow", String(pct));
    const count = box.querySelector("[data-job-count]");
    if (count) count.textContent = total ? `${done} / ${total}件` : "";
    const msg = box.querySelector("[data-job-message]");
    if (msg) msg.textContent = job.message || "";
    const badge = box.querySelector(".progress-state .badge");
    if (badge && JOB_LABELS[job.status]) badge.textContent = JOB_LABELS[job.status];
    box.dataset.jobStatus = job.status;
    box.dispatchEvent(new CustomEvent("job:update", { detail: job, bubbles: true }));
    if (JOB_FINISHED.includes(job.status)) {
      box.dispatchEvent(new CustomEvent("job:finished", { detail: job, bubbles: true }));
      if (box.dataset.reloadOnDone !== undefined) window.location.reload();
    }
  });
}

// ---- ファイルのドロップ ---------------------------------------------------------------
function acceptsFile(input, file) {
  const accept = (input.getAttribute("accept") || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  if (!accept.length) return true;
  const name = file.name.toLowerCase();
  return accept.some((a) => (a.startsWith(".") ? name.endsWith(a) : true));
}

function bindFileDrop(root) {
  const input = root.querySelector("input[type=file]");
  const nameEl = root.querySelector("[data-file-name]");
  if (!input) return;
  const show = () => {
    const files = Array.from(input.files || []);
    const bad = files.filter((f) => !acceptsFile(input, f));
    root.classList.toggle("has-file", files.length > 0 && !bad.length);
    root.classList.toggle("is-invalid", bad.length > 0);
    if (nameEl) {
      nameEl.textContent = bad.length
        ? `${bad.map((f) => f.name).join("、")} は選べません（${input.getAttribute("accept")}）`
        : files.map((f) => f.name).join("、");
    }
    input.setCustomValidity(bad.length ? "この形式のファイルは選べません" : "");
  };
  input.addEventListener("change", show);
  ["dragenter", "dragover"].forEach((type) =>
    root.addEventListener(type, (e) => {
      e.preventDefault();
      root.classList.add("is-over");
    }));
  ["dragleave", "dragend"].forEach((type) => root.addEventListener(type, () => root.classList.remove("is-over")));
  root.addEventListener("drop", (e) => {
    e.preventDefault();
    root.classList.remove("is-over");
    const files = e.dataTransfer?.files;
    if (!files || !files.length) return;
    const dt = new DataTransfer();
    Array.from(files).slice(0, input.multiple ? files.length : 1).forEach((f) => dt.items.add(f));
    input.files = dt.files;
    show();
  });
}

// ---- 確認ダイアログ（data-confirm） -------------------------------------------------------
// <form data-confirm="文言"> は送信前、<a/button data-confirm="文言"> はクリック前に確認する
function confirmDialog(message, okLabel) {
  const dialog = document.getElementById("confirmDialog");
  if (!dialog || typeof dialog.showModal !== "function") return Promise.resolve(window.confirm(message));
  dialog.querySelector("#confirmDialogText").textContent = message;
  dialog.querySelector("#confirmDialogOk").textContent = okLabel || "実行する";
  dialog.returnValue = "";
  dialog.showModal();
  dialog.querySelector("button[value=cancel]")?.focus();
  return new Promise((resolve) => {
    dialog.addEventListener("close", () => resolve(dialog.returnValue === "ok"), { once: true });
  });
}

document.addEventListener("submit", async (event) => {
  const form = event.target;
  if (!form.matches("form[data-confirm]") || form.dataset.confirmed === "1") return;
  event.preventDefault();
  const submitter = event.submitter;
  if (await confirmDialog(form.dataset.confirm, form.dataset.confirmOk)) {
    form.dataset.confirmed = "1";
    if (submitter && submitter.name) {
      const hidden = document.createElement("input");
      hidden.type = "hidden";
      hidden.name = submitter.name;
      hidden.value = submitter.value;
      form.appendChild(hidden);
    }
    form.submit();
  }
});

document.addEventListener("click", async (event) => {
  const target = event.target.closest("a[data-confirm], button[data-confirm]:not([type=submit])");
  if (target && target.dataset.confirmed !== "1") {
    event.preventDefault();
    if (await confirmDialog(target.dataset.confirm, target.dataset.confirmOk)) {
      target.dataset.confirmed = "1";
      target.click();
      target.dataset.confirmed = "";
    }
    return;
  }

  // メッセージを閉じる
  const dismiss = event.target.closest("[data-dismiss]");
  if (dismiss) dismiss.parentElement.remove();

  // Markdown / JSON のコピー
  const copy = event.target.closest("[data-copy]");
  if (copy) {
    const source = document.querySelector(copy.dataset.copy);
    if (!source) return;
    navigator.clipboard.writeText(source.value ?? source.textContent).then(() => {
      toast("コピーしました。");
    }, () => toast("コピーできませんでした。", "err"));
  }

  // 全選択チェック（data-check-all="name"）
  const all = event.target.closest("input[data-check-all]");
  if (all) {
    document.querySelectorAll(`input[type=checkbox][name="${all.dataset.checkAll}"]`).forEach((box) => {
      if (!box.disabled) box.checked = all.checked;
    });
    document.dispatchEvent(new Event("selection:change"));
  }
});

// 選択数の表示（data-count-of="name" の要素に選択件数を入れる）
function updateSelectionCounts() {
  document.querySelectorAll("[data-count-of]").forEach((el) => {
    const n = document.querySelectorAll(`input[type=checkbox][name="${el.dataset.countOf}"]:checked`).length;
    el.textContent = String(n);
    const button = el.closest("button");
    if (button) button.disabled = n === 0;
  });
}
document.addEventListener("change", (event) => {
  if (event.target.matches("input[type=checkbox]")) updateSelectionCounts();
});
document.addEventListener("selection:change", updateSelectionCounts);

// ---- ragFetch: 画面を移らずにサーバへ送る -------------------------------------------------
// 1画面で全部やるので、送信はすべてここを通る（利用者の指示 2026-09-20）。
// 使い方:
//   await ragFetch("/tables/12/read")                        … GET
//   await ragFetch(url, { json: {...} })                     … JSON を POST（既定は POST）
//   await ragFetch(url, { form: formElementOrFormData })     … フォーム（ファイルもそのまま送れる）
//   await ragFetch(url, { html: true })                      … 画面の一部の HTML を受け取る
// 戻り値: サーバが JSON を返せばそのオブジェクト、HTML を返せば { html: "…" }。
// 失敗（HTTP エラー・通信できない）は、ここでトーストに出してから例外を投げる。
// 呼び出し側で自分でメッセージを出すときは { quiet: true }。
async function ragFetch(url, options = {}) {
  const { json, form, html, quiet = false, method, ...rest } = options;
  const init = { headers: { Accept: html ? "text/html" : "application/json" }, ...rest };
  if (json !== undefined) {
    init.method = method || "POST";
    init.headers["Content-Type"] = "application/json";
    init.body = JSON.stringify(json || {});
  } else if (form !== undefined) {
    init.method = method || "POST";
    // FormData のときは Content-Type を付けない（境界文字列はブラウザが付ける）
    init.body = form instanceof FormData ? form : new FormData(form);
  } else {
    init.method = method || "GET";
  }

  let res;
  try {
    res = await fetch(url, init);
  } catch (e) {
    const error = new Error("このPCのアプリにつながりませんでした。アプリが動いているか確かめてください");
    error.status = 0;
    if (!quiet) toast(error.message, "err");
    throw error;
  }

  const type = (res.headers.get("Content-Type") || "").toLowerCase();
  let data = {};
  let text = "";
  if (type.includes("application/json")) {
    data = await res.json().catch(() => ({}));
  } else {
    text = await res.text().catch(() => "");
    data = { html: text };
  }
  if (!res.ok) {
    // 問題が複数あるとき（列の対応づけの保存など）は errors に全部入っている。1件だけ見せて残りを隠さない
    const error = new Error(data.error || `うまくいきませんでした（HTTP ${res.status}）`);
    error.status = res.status;
    error.errors = Array.isArray(data.errors) ? data.errors : null;
    error.data = data;
    if (!quiet) toast(error.errors ? error.errors.slice(0, 3).join(" / ") : error.message, "err");
    throw error;
  }
  return data;
}

// ---- ragSections: 1画面の中の「段」の開け閉め -------------------------------------------
// 画面を移らず、下に段が増えていく形にする（利用者の指示 2026-09-20）。
// 段の書き方（templates 側）:
//   <section class="step" data-step="sheet">
//     <div class="step-head">
//       <span class="step-no"><span class="step-no-text">2</span></span>
//       <h2 class="step-title">どのシートを使うか</h2>
//       <p class="step-summary" data-step-summary></p>
//       <span class="step-toggle" data-step-toggle hidden>開く</span>
//     </div>
//     <div class="step-body"> … </div>
//   </section>
// 最初の段だけ class="step is-open" にしておけば、あとは JS が進める。
const ragSections = (() => {
  const el = (id) => (id instanceof Element ? id : document.querySelector(`.step[data-step="${id}"]`));
  const list = () => Array.from(document.querySelectorAll(".step[data-step]"));

  function setToggle(section) {
    const toggle = section.querySelector("[data-step-toggle]");
    if (!toggle) return;
    const done = section.classList.contains("is-done");
    toggle.hidden = !done;
    toggle.textContent = section.classList.contains("is-open") ? "閉じる" : "開く";
  }

  /** その段を開く（他の開いている段は、済んだものだけ畳む）。 */
  function open(id, { scroll = true } = {}) {
    const section = el(id);
    if (!section) return null;
    list().forEach((s) => {
      if (s !== section && s.classList.contains("is-done")) s.classList.remove("is-open");
      setToggle(s);
    });
    section.classList.add("is-open");
    setToggle(section);
    section.dispatchEvent(new CustomEvent("step:open", { bubbles: true }));
    if (scroll) section.scrollIntoView({ behavior: "smooth", block: "start" });
    return section;
  }

  /** その段を「済み」にして畳み、要約を見出しに出す。next を渡すとその段を開く。 */
  function done(id, summary, next) {
    const section = el(id);
    if (!section) return null;
    section.classList.add("is-done");
    section.classList.remove("is-open");
    const label = section.querySelector("[data-step-summary]");
    if (label && summary !== undefined && summary !== null) {
      label.textContent = String(summary);
      label.title = String(summary);
    }
    setToggle(section);
    section.dispatchEvent(new CustomEvent("step:done", { bubbles: true }));
    if (next) open(next);
    return section;
  }

  /** その段を開き直し、後ろの段を「まだ」に戻す（やり直すとき）。 */
  function reset(id) {
    const all = list();
    const index = all.indexOf(el(id));
    if (index < 0) return null;
    all.slice(index + 1).forEach((s) => {
      s.classList.remove("is-open", "is-done");
      const label = s.querySelector("[data-step-summary]");
      if (label) label.textContent = "";
      setToggle(s);
      s.dispatchEvent(new CustomEvent("step:reset", { bubbles: true }));
    });
    return open(all[index]);
  }

  /** 段の中身の要素（ここに受け取った HTML を入れる）。 */
  const body = (id) => el(id)?.querySelector(".step-body") || null;

  /** 段の中に「処理中」を出す（text が空なら消す）。画面は移らない。 */
  function working(id, text) {
    const section = el(id);
    if (!section) return;
    let box = section.querySelector("[data-step-working]");
    if (!box) {
      box = document.createElement("p");
      box.className = "step-working";
      box.setAttribute("data-step-working", "");
      box.setAttribute("role", "status");
      box.innerHTML = '<span class="spinner" aria-hidden="true"></span><span data-step-working-text></span>';
      (body(id) || section).prepend(box);
    }
    box.querySelector("[data-step-working-text]").textContent = text || "";
    box.hidden = !text;
  }

  // 済んだ段の見出しをクリックすると開き直せる
  document.addEventListener("click", (event) => {
    const head = event.target.closest(".step.is-done > .step-head");
    if (!head) return;
    const section = head.parentElement;
    if (section.classList.contains("is-open")) {
      section.classList.remove("is-open");
      setToggle(section);
    } else {
      open(section, { scroll: false });
    }
  });

  return { el, open, done, reset, body, working, refresh: () => list().forEach(setToggle) };
})();

// ---- 初期化 -------------------------------------------------------------------------
document.querySelectorAll("[data-file-drop]").forEach(bindFileDrop);
document.querySelectorAll(".progress-box[data-job-url]").forEach(bindProgressBox);
updateSelectionCounts();
ragSections.refresh();
// 他のスクリプトから使う
window.ragFetch = ragFetch;
window.ragSections = ragSections;
window.App = { toast, postJson, getJson, pollJob, confirmDialog, bindFileDrop, bindProgressBox, ragFetch, ragSections };
