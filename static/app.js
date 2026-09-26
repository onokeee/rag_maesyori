// 画面の JavaScript（1ファイル）。
//   1. 全画面共通: トースト、fetch（ragFetch）、1画面の中の段（ragSections）、ファイルのドロップ、確認ダイアログ、ジョブ進捗
//   2. 表の取り込み（/tables）
// 2 は自分の画面（[data-tables-page]）が無ければ何もしない。
// tests/test_cross.py は「// ==== 」の見出しでこのファイルを区切って、その画面の部分だけを node で動かす。
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

// base.html の ui_progress(job, url) を自動で更新する
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
    // クリックで選んだときと同じにする（change を見ている画面がファイル名を使う）
    input.dispatchEvent(new Event("change", { bubbles: true }));
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

});

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
    const error = new Error("アプリのサーバーにつながりませんでした。ネットワークとアプリが動いているか確かめてください");
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

// ---- ragDiscard: ダウンロードしないまま画面を離れたら捨てる --------------------------------
// 「その場でダウンロードしない限り、その場ですぐ捨てる」（利用者の指示 2026-09-20）。
// 画面が隠れた・閉じられた時点で、その人がまだダウンロードしていない取り込みをサーバに捨てさせる。
//
// 使い方（各画面の JS から）:
//   const guard = ragDiscard.watch("/tables/discard", () => ({ import_ids: ids }));
//   guard.now();     … いま持っている分を今すぐ捨てる（新しいファイルを置く前。await できる）
//
// 閉じる合図は pagehide（閉じる・別のページへ移る）と visibilitychange（タブを隠す・スリープ）の両方で見る。
// どちらも「必ず呼ばれる」ものではない（強制終了・LANの切断）ので、これだけに頼らない
// （サーバ側は IDLE_HOURS さわられていないものを捨て、起動時にも捨てる。core/purge.py）。
// 送るのは navigator.sendBeacon（閉じる最中でも届く。応答は読めない）。使えないブラウザでは
// keepalive を付けた fetch にする。同じものが2回届いても、サーバ側は2回目に何もしない。
const ragDiscard = (() => {
  function send(url, payload) {
    const text = JSON.stringify(payload || {});
    // sendBeacon は Content-Type を選べるが、text/plain 以外だと事前確認（preflight）が要る。
    // 同じサイト宛てなので preflight は起きないが、素直に text/plain で送る（サーバは中身で判断する）
    try {
      if (navigator.sendBeacon) {
        const blob = new Blob([text], { type: "text/plain;charset=UTF-8" });
        if (navigator.sendBeacon(url, blob)) return true;
      }
    } catch (e) { /* 使えなければ下の fetch にする */ }
    try {
      fetch(url, {
        method: "POST", keepalive: true, headers: { "Content-Type": "application/json" }, body: text,
      }).catch(() => {});
      return true;
    } catch (e) {
      return false;
    }
  }

  // タブを隠しただけ（別のタブを見に行った・画面を最小化した）で捨てると、戻ってきたときに
  // 取り込みが消えていて困る。閉じた（pagehide）ならすぐ捨て、隠れただけなら この時間だけ待つ。
  // 待っている間に戻ってくれば捨てない。隠れたままのタブでは timer が遅れるが、遅れる分には困らない。
  const HIDDEN_GRACE_MS = 5 * 60 * 1000;

  /** url へ「これを捨てて」と送る見張りを付ける。payloadFn() は {doc_ids:[…]} / {import_ids:[…]} を返す。 */
  function watch(url, payloadFn) {
    let hiddenTimer = null;
    const has = () => {
      const payload = (payloadFn && payloadFn()) || {};
      const ids = [].concat(payload.doc_ids || [], payload.import_ids || []);
      return ids.length ? payload : null;
    };
    const leave = () => {
      const payload = has();
      if (payload) send(url, payload);
    };
    const stopTimer = () => { clearTimeout(hiddenTimer); hiddenTimer = null; };

    window.addEventListener("pagehide", () => { stopTimer(); leave(); });
    document.addEventListener("visibilitychange", () => {
      stopTimer();
      if (document.visibilityState === "hidden") hiddenTimer = setTimeout(leave, HIDDEN_GRACE_MS);
    });
    return {
      /** いま持っている分を今すぐ捨てる（新しいファイルを置く前）。失敗しても取り込みは続けられる。 */
      async now() {
        const payload = has();
        if (!payload) return;
        try {
          await ragFetch(url, { json: payload, quiet: true });
        } catch (e) { /* 捨て損ねても、サーバ側の時間切れで片付く */ }
      },
    };
  }

  return { watch, send };
})();

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
    const open = section.classList.contains("is-open");
    toggle.hidden = !done;
    toggle.textContent = open ? "閉じる" : "開く";
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
  }

  /** その段を開く（3画面とも data-steps-open なので、ほかの段は畳まない）。 */
  function open(id, { scroll = true } = {}) {
    const section = el(id);
    if (!section) return null;
    list().forEach(setToggle);
    section.classList.add("is-open");
    setToggle(section);
    if (scroll) section.scrollIntoView({ behavior: "smooth", block: "start" });
    return section;
  }

  /** その段を「済み」にして要約を見出しに出す。next を渡すとその段を開く。 */
  function done(id, summary, next) {
    const section = el(id);
    if (!section) return null;
    section.classList.add("is-done");
    note(section, summary);
    setToggle(section);
    if (next) open(next);
    return section;
  }

  /** その段を「まだ」に戻す（やり直すとき。要約も消す）。 */
  function close(id) {
    const section = el(id);
    if (!section) return null;
    section.classList.remove("is-open", "is-done");
    note(section, "");
    setToggle(section);
    return section;
  }

  /** 段の見出しに出す要約（空なら消す）。 */
  function note(id, text) {
    const label = el(id)?.querySelector("[data-step-summary]");
    if (!label || text === undefined || text === null) return;
    label.textContent = String(text);
    label.title = String(text);
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

  return { el, open, done, close, note, body, working, refresh: () => list().forEach(setToggle) };
})();

// ---- 初期化 -------------------------------------------------------------------------
document.querySelectorAll("[data-file-drop]").forEach(bindFileDrop);
document.querySelectorAll(".progress-box[data-job-url]").forEach(bindProgressBox);
ragSections.refresh();
// 他のスクリプトから使う
window.ragFetch = ragFetch;
window.ragSections = ragSections;
window.ragDiscard = ragDiscard;
window.App = { toast, getJson, pollJob, confirmDialog, bindFileDrop, bindProgressBox, ragFetch,
               ragSections, ragDiscard };


// ---- AI接続（ヘッダー右上。どの画面にもある） ----------------------------------------------------
// 利用者の指示（2026-09-21）:「AI接続の設定は…ヘッダーの画面右上『AI接続』に移動」「接続中 か 未接続 一目で分かるように」。
// 状態はサーバーが覚えている結果（templates/base.html の ai_header）を出すだけ。確かめに行くのは
// 「パネルを開いたとき」「保存したとき」だけ（AI整形を始めるときはサーバー側で確かめ、答えに status が付く）。
// 画面を開くたびには確かめない（お金がかかり、画面が待たされる）。
window.ragAiHeader = (() => {
  const header = document.querySelector("[data-ai-header]");
  const dialog = document.getElementById("aiDialog");
  if (!header || !dialog) return { render() {} };
  const q = (selector) => dialog.querySelector(selector);
  const make = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (v !== null && v !== undefined) node.setAttribute(k, v);
    }
    node.append(...children.filter((c) => c != null));
    return node;
  };

  /** サーバーが返す status（llm.connection_status）でヘッダーとパネルを描き直す。 */
  function render(status) {
    if (!status) return;
    header.dataset.state = status.state;
    dialog.dataset.state = status.state;
    document.querySelectorAll("[data-ai-state]").forEach((n) => (n.textContent = status.state_label || ""));
    document.querySelectorAll("[data-ai-sub]").forEach((n) => (n.textContent = status.sub_text || ""));
    document.querySelectorAll("[data-ai-panel-sub]").forEach((n) => (n.textContent = status.panel_sub_text || ""));
    const detail = q("[data-ai-detail]");
    if (detail) {
      detail.hidden = !status.check_detail;
      detail.firstElementChild.textContent = status.check_detail || "";
    }
    const key = q("[data-ai-field=api_key]");
    if (key) key.placeholder = status.api_key_saved ? "保存済み（変えるときだけ入力）" : "sk-… を貼り付け";
    const keyNote = q("[data-ai-key-note]");
    if (keyNote) {
      keyNote.textContent = status.api_key_saved ? "このブラウザのキーを保存済みです。"
        : (status.api_key_source ? "このブラウザのキーは未保存です（サーバー共通のキーを使います）。" : "まだ保存していません。");
    }
    const clear = q("[data-ai-clear]");
    if (clear) clear.disabled = !status.api_key_saved;
    const list = q("[data-ai-model-list]");
    if (list && Array.isArray(status.models)) {
      list.replaceChildren(...status.models.map((m) => make("option", { value: m })));
      const note = q("[data-ai-models-note]");
      if (note) note.textContent = status.models.length ? `候補 ${status.models.length}件（入力欄で選べます）` : "取得すると入力欄の候補に出ます。";
    }
    const eff = status.effective || {};
    const effUrl = q("[data-ai-effective-url]");
    if (effUrl) effUrl.textContent = eff.chat_url || "（未設定）";
    const effModel = q("[data-ai-effective-model]");
    if (effModel) effModel.textContent = eff.model || "（未設定）";
  }

  function showSteps(steps) {
    const list = q("[data-ai-test-result]");
    if (!list) return;
    list.replaceChildren(...(steps || []).map((step) => make("li", {},
      make("div", { class: "item-main" },
        make("span", { class: "item-title", text: `${step.ok ? "OK" : "NG"}　${step.name}` }),
        make("span", { class: step.ok ? "muted" : "warn", text: step.detail || "" })))));
    list.hidden = !steps || !steps.length;
  }

  let checking = false;
  async function check() {
    if (checking) return;
    checking = true;
    const busy = q("[data-ai-busy]");
    const button = q("[data-ai-test]");
    if (busy) busy.hidden = false;
    if (button) button.disabled = true;
    try {
      const res = await ragFetch(dialog.dataset.testUrl, { json: {}, quiet: true });
      render(res.status);
      showSteps(res.steps);
    } catch (e) {
      toast(e.message, "err");
    } finally {
      checking = false;
      if (busy) busy.hidden = true;
      if (button) button.disabled = false;
    }
  }

  function open() {
    if (typeof dialog.showModal !== "function") return;
    if (!dialog.open) dialog.showModal();
    // 何も入っていなければ確かめに行っても「未設定」と返るだけなので、行かない
    if (header.dataset.state !== "off") check();
  }

  async function save(button) {
    const data = {};
    dialog.querySelectorAll("[data-ai-field]").forEach((input) => { data[input.dataset.aiField] = input.value; });
    button.disabled = true;
    try {
      const res = await ragFetch(dialog.dataset.saveUrl, { json: data, quiet: true });
      const key = q("[data-ai-field=api_key]");
      if (key) key.value = "";   // キーは画面に置いたままにしない
      render(res.status);
      showSteps(res.steps);
      toast(res.message || "保存しました", res.connected === false && res.status && res.status.ready ? "err" : "ok");
    } catch (e) {
      toast(e.message, "err");
    } finally {
      button.disabled = false;
    }
  }

  async function clearKey(button) {
    button.disabled = true;
    try {
      const res = await ragFetch(dialog.dataset.clearUrl, { json: {}, quiet: true });
      const key = q("[data-ai-field=api_key]");
      if (key) key.value = "";
      render(res.status);
      showSteps([]);
      toast(res.message || "キーを消しました");
    } catch (e) {
      toast(e.message, "err");
      button.disabled = false;
    }
  }

  async function fetchModels(button) {
    button.disabled = true;
    try {
      const res = await ragFetch(dialog.dataset.modelsUrl, { json: {}, quiet: true });
      render(res.status);
      toast(res.message || "取得しました");
    } catch (e) {
      toast(e.message, "err");
    } finally {
      button.disabled = false;
    }
  }

  document.addEventListener("click", (event) => {
    if (event.target.closest("[data-ai-open]")) {
      open();
      return;
    }
    if (!dialog.contains(event.target)) return;
    const hit = (selector) => event.target.closest(selector);
    if (hit("[data-ai-close]")) dialog.close();
    else if (hit("[data-ai-test]")) check();
    else if (hit("[data-ai-save]")) save(hit("[data-ai-save]"));
    else if (hit("[data-ai-models]")) fetchModels(hit("[data-ai-models]"));
    else if (hit("[data-ai-clear]")) {
      const button = hit("[data-ai-clear]");
      // data-confirm の付いたボタンは、上の確認ダイアログの処理が先に走り、確認後にもう一度クリックが来る
      if (!button.dataset.confirm || button.dataset.confirmed === "1") clearKey(button);
    }
  });
  // Enter で保存（パネルは <form> ではないので自分で拾う）
  dialog.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && event.target.matches("input[data-ai-field]")) {
      event.preventDefault();
      const button = q("[data-ai-save]");
      if (button && !button.disabled) save(button);
    }
  });

  return { render, open };
})();

// ==== tables: 表の取り込み（/tables。旧 tables.js） ====
// 表の取り込み（/tables）: 1画面で全部やる（利用者の指示 2026-09-20）。
// 画面は移らない。段（.step）の中身はサーバが HTML の断片で返し、保存・実行は JSON でやりとりする。
// 段: file → source → layout → columns → ai → preview → done
(() => {
  "use strict";
  const page = document.querySelector("[data-tables-page]");
  if (!page) return;
  const { toast } = window.App;
  const sections = window.ragSections;
  const rf = window.ragFetch;
  const JOB_FINISHED = ["done", "failed", "cancelled", "interrupted"];

  let urls = null;                 // サーバが返す各 URL（取り込みごと）
  let importId = null;             // いま画面で作業している取り込みの番号（捨てるときに使う）
  const pollers = {};              // 段ごとの進捗の見張り
  const notes = {};                // 段ごとの要約（見出しに出す）

  // ダウンロードしないまま画面を離れたら、この取り込みは捨てる（利用者の指示 2026-09-20）。
  // 番号はサーバが返す URL から取る（res.import_id があればそれを使う）。
  const idFrom = (url) => {
    const hit = /\/imports\/(\d+)\//.exec(url || "");
    return hit ? Number(hit[1]) : null;
  };
  const guard = (window.ragDiscard || { watch: () => ({ now: async () => {} }) })
    .watch(page.dataset.discardUrl || "/tables/discard", () => ({ import_ids: importId ? [importId] : [] }));

  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else if (v !== null && v !== undefined) node.setAttribute(k, v);
    }
    for (const child of children) {
      if (child == null) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  };

  const panelUrl = (name) => urls.panel.replace(/NAME$/, name);
  const body = (name) => sections.body(name);

  // data-confirm の付いたボタンは app.js が先に確認ダイアログを出す。確認前のクリックでは何もしない
  const confirmed = (node) => !node.dataset.confirm || node.dataset.confirmed === "1";

  const stopPoller = (name) => {
    if (pollers[name]) {
      pollers[name].stop();
      delete pollers[name];
    }
  };

  // ---- 進み具合（同じ画面の中に出す） ----------------------------------------------------
  function poll(url, onUpdate) {
    let stopped = false;
    let wait = 1000;
    let timer = null;
    const tick = async () => {
      if (stopped) return;
      try {
        const job = await rf(url, { quiet: true });
        wait = 1200;
        onUpdate(job);
        if (JOB_FINISHED.includes(job.status)) return;
      } catch (e) {
        if (e.status === 404) {
          // 取り込みごと無くなった（ダウンロードした・自分で消した・画面を離れて捨てた）。
          // 失敗ではないので何も言わず止め、画面を最初の状態に戻す
          stopped = true;
          clearTimeout(timer);
          window.location.reload();
          return;
        }
        wait = Math.min(wait * 2, 8000);   // 通信エラーは間隔を広げて続ける
      }
      if (!stopped) timer = setTimeout(tick, wait);
    };
    tick();
    return { stop() { stopped = true; clearTimeout(timer); } };
  }

  const JOB_TITLES = { table_read: "表を読み込んでいます", table_preview: "Markdownの下書きを作っています",
                       table_render: "Markdownを作っています", ai_format: "AI整形を実行しています" };

  /** 段の中に進み具合を出し、終わったら done(job) を呼ぶ。 */
  function runJob(name, job, onFinish) {
    stopPoller(name);
    const target = body(name);
    const bar = el("div", { class: "progress-bar" });
    const count = el("span", {});
    const message = el("span", { class: "muted" });
    // 中止できるジョブのときだけボタンを出す（サーバーが止める口を教えてくれる）
    const cancelUrl = job.cancel_url || "";
    const cancel = cancelUrl
      ? el("button", { type: "button", class: "btn danger-outline small", text: "中止" })
      : el("span", { class: "hint", text: "この処理は終わるまでお待ちください（中止はできません）。" });
    if (cancelUrl) {
      cancel.addEventListener("click", async () => {
        cancel.disabled = true;
        try {
          const res = await rf(cancelUrl, { json: {} });
          toast(res.message || "処理を中止しました");
        } catch (e) {
          cancel.disabled = false;
        }
      });
    }
    const box = el("div", { class: "progress-box" },
      el("div", { class: "progress-head" },
        el("span", { class: "progress-label", text: JOB_TITLES[job.kind] || "処理しています" }),
        el("span", { class: "progress-state" })),
      el("div", { class: "progress", role: "progressbar", "aria-label": "進み具合" }, bar),
      el("p", { class: "progress-text" }, count, " ", message),
      el("p", { class: "hint", text: "この画面のままお待ちください（大きなファイルは数十秒かかることがあります）。" }),
      el("div", { class: "btn-row" }, cancel));
    target.replaceChildren(box);
    sections.open(name, { scroll: false });

    const update = (j) => {
      const p = j.progress || {};
      const done = Number(p.done || 0);
      const total = Number(p.total || 0);
      bar.style.width = `${total ? Math.floor((done * 100) / total) : 0}%`;
      bar.classList.toggle("indeterminate", !total && j.status === "running");
      count.textContent = total ? `${done} / ${total}件` : "";
      message.textContent = j.message || "";
    };
    update(job);
    pollers[name] = poll(job.url, (j) => {
      update(j);
      if (JOB_FINISHED.includes(j.status)) {
        stopPoller(name);
        onFinish(j);
      }
    });
  }

  // ---- 段の中身を取りに行く -------------------------------------------------------------
  function setNote(name, text) {
    notes[name] = text || "";
    const label = sections.el(name)?.querySelector("[data-step-summary]");
    if (label) {
      label.textContent = notes[name];
      label.title = notes[name];
    }
  }

  function lock(name, data) {
    const section = sections.el(name);
    if (!section) return;
    stopPoller(name);
    setNote(name, data.locked || "");
    if (data.failed) {
      // 読み込みに失敗したときは、やり直せるように開いたままにする
      const box = body(name);
      box.replaceChildren(
        el("div", { class: "flash flash-error", role: "alert", text: data.locked || "" }),
        el("div", { class: "form-actions" },
          el("button", { type: "button", class: "btn primary", "data-reread": "" , text: "もう一度読み込む" })));
      section.classList.remove("is-done");
      sections.open(name, { scroll: false });
      return;
    }
    section.classList.remove("is-open", "is-done");
    body(name).replaceChildren();
    sections.refresh();   // 「開く」の文字が残らないように印を付け直す
  }

  async function loadPanel(name, { open = true, scroll = true, query = "" } = {}) {
    const target = body(name);
    if (!target || !urls) return null;
    stopPoller(name);
    target.replaceChildren(el("p", { class: "muted", text: "読み込んでいます…" }));
    if (open) sections.open(name, { scroll });
    let data;
    try {
      data = await rf(panelUrl(name) + query, { quiet: true });
    } catch (e) {
      target.replaceChildren(el("div", { class: "flash flash-error", role: "alert", text: e.message }));
      return null;
    }
    if (data.locked) {
      lock(name, data);
      return data;
    }
    if (data.job && !data.job.finished) {
      setNote(name, "");
      runJob(name, data.job, () => loadPanel(name, { open, scroll: false }));
      return data;
    }
    target.innerHTML = data.html || "";
    setNote(name, data.note || "");
    afterRender(name, target);
    if (open) sections.open(name, { scroll });
    return data;
  }

  function clearStep(name) {
    stopPoller(name);
    const section = sections.el(name);
    if (!section) return;
    section.classList.remove("is-open", "is-done");
    setNote(name, "");
    body(name).replaceChildren();
  }

  // 段の並び。飛ばした段にも「まだできない理由」を出すために使う
  const STEP_ORDER = ["file", "source", "layout", "columns", "ai", "preview", "done"];

  /** 読み込み・確定などのジョブを始めたあとの共通処理。進み具合は次の段の中に出す。 */
  function afterAction(res) {
    (res.reset || []).forEach(clearStep);
    const next = res.next;
    if (!next) return;
    // 飛ばした段（例: 経過の記録の列が無いときの AI整形）は開かずに取りに行く。
    // 取りに行かないと灰色のまま何も書かれず、なぜ使えないのかが分からない
    const skipped = (res.reset || [])
      .filter((name) => name !== next && STEP_ORDER.indexOf(name) < STEP_ORDER.indexOf(next));
    const loadSkipped = () => skipped.forEach((name) => loadPanel(name, { open: false, scroll: false }));
    if (res.job && !res.job.finished) {
      // 処理が終わってから取りに行く（動いている間は、どの段も同じジョブの進み具合を映してしまう）
      runJob(next, res.job, () => { loadSkipped(); loadPanel(next, { open: true, scroll: false }); });
      return;
    }
    loadSkipped();
    loadPanel(next);
  }

  // ---- 段ごとの後始末（断片を入れたあとの初期化） ------------------------------------------
  function afterRender(name, target) {
    if (name === "preview") {
      const file = target.querySelector("[data-file-open]");
      if (file) file.click();
    }
    if (name === "ai") {
      // AI整形のジョブの進み具合（断片の中の progress-box は自分で動かす）
      target.querySelectorAll(".progress-box[data-job-url]").forEach((box) => {
        window.App.bindProgressBox(box);
        let last = box.dataset.jobStatus;
        box.addEventListener("job:update", (event) => {
          const status = event.detail.status;
          const paused = (s) => s === "paused";
          // 実行中⇔一時停止が変わったら段を出し直す（ボタンを［再開］／［一時停止］に合わせる）
          if (last && status && paused(last) !== paused(status) && ["running", "paused"].includes(status)) {
            loadPanel("ai", { open: true, scroll: false });
            return;
          }
          last = status || last;
          const p = event.detail.progress || {};
          const detail = target.querySelector("[data-ai-detail]");
          if (detail) {
            const rest = !p.remaining_sec ? "" : p.remaining_sec < 60 ? "／残り1分未満" : `／残り約${Math.ceil(p.remaining_sec / 60)}分`;
            detail.textContent = `OK ${p.ok || 0}／要確認 ${p.flagged || 0}／エラー ${p.error || 0}／ルールのみ ${p.rule_only || 0}${rest}`;
          }
        });
        box.addEventListener("job:finished", () => loadPanel("ai", { open: true, scroll: false }), { once: true });
      });
    }
  }

  // ---- 1 ファイルを置く ---------------------------------------------------------------
  const uploadForm = page.querySelector("[data-upload-form]");
  // 置いた（またはクリックで選んだ）時点でそのまま読み取る。ボタンは押さなくてよい
  uploadForm?.querySelector("input[type=file]")?.addEventListener("change", (event) => {
    if (event.target.files && event.target.files.length) uploadForm.requestSubmit();
  });
  uploadForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const input = uploadForm.querySelector("input[type=file]");
    if (!input?.files?.length) {
      toast("ファイルを選んでください。", "err");
      return;
    }
    sections.working("file", "ファイルを読み取っています…");
    try {
      // 新しいファイルを置いたら、前の取り込み（ダウンロードしていない分）はその場で捨てる
      if (importId) { await guard.now(); importId = null; urls = null; }
      const res = await rf(page.dataset.uploadUrl, { form: new FormData(uploadForm), quiet: true });
      urls = res.urls;
      importId = Number(res.import_id) || idFrom(urls && urls.panel);
      // 前の取り込みの中身を残さない（残すと✓付きのまま押せて、消えた取り込みを取りに行って404になる）
      ["columns", "ai", "preview", "done"].forEach(clearStep);
      sections.done("file", res.file_name);
      await loadPanel("layout", { open: true, scroll: false });
      await loadPanel("source", { open: true });
    } catch (e) {
      toast(e.message, "err");
    } finally {
      sections.working("file", "");
    }
  });

  // ---- 共通のクリック -------------------------------------------------------------------
  page.addEventListener("click", async (event) => {
    const hit = (selector) => event.target.closest(selector);

    // この取り込みを削除
    const delImport = hit("[data-import-delete]");
    if (delImport) {
      if (!confirmed(delImport)) return;
      try {
        const res = await rf(urls.delete, { json: {}, quiet: true });
        importId = null;                      // もう消えているので、画面を離れるときに捨てるものは無い
        toast(res.message || "削除しました");
        window.location.href = page.dataset.newUrl;
      } catch (e) {
        toast(e.message, "err");
      }
      return;
    }

    // もう一度読み込む
    const reread = hit("[data-reread]");
    if (reread) {
      reread.disabled = true;
      try {
        afterAction(await rf(urls.read, { json: {}, quiet: true }));
      } catch (e) {
        toast(e.message, "err");
        reread.disabled = false;
      }
      return;
    }

    // 表の範囲: 行番号をクリックして見出し行・データの終わりを指定
    const rownum = hit("[data-layout-page] th.rownum");
    if (rownum) {
      pickRow(rownum, event.shiftKey);
      return;
    }

    // 表の範囲: この範囲で読み込む
    const layoutSubmit = hit("[data-layout-submit]");
    if (layoutSubmit) {
      const root = layoutSubmit.closest("[data-layout-page]");
      layoutSubmit.disabled = true;
      try {
        const res = await rf(urls.layout, {
          json: { header_rows: root.querySelector("[data-header-input]").value,
                  data_end_row: root.querySelector("[data-end-input]").value },
          quiet: true });
        sections.done("source", notes.source);
        sections.done("layout", notes.layout);
        afterAction(res);
      } catch (e) {
        toast(e.message, "err");
      } finally {
        layoutSubmit.disabled = false;
      }
      return;
    }

    // 列の対応づけ: ［変更する］で要約をたたんで表を出す
    const openTable = hit("[data-columns-open]");
    if (openTable) {
      openColumnsTable(openTable.closest("[data-columns-editor]"));
      return;
    }

    // 列の対応づけ: 保存して読み込む
    const save = hit("[data-editor-save]");
    if (save) {
      await saveColumns(save);
      return;
    }

    // 内容の確認: ファイルの中身、ページ送り、確定
    const fileOpen = hit("[data-file-open]");
    if (fileOpen) {
      await showFile(fileOpen);
      return;
    }
    const rowsPage = hit("[data-rows-page]");
    if (rowsPage) {
      await loadPanel("preview", { open: true, scroll: false, query: `?page=${rowsPage.dataset.rowsPage}` });
      return;
    }
    const previewRetry = hit("[data-preview-retry]");
    if (previewRetry) {
      previewRetry.disabled = true;
      try {
        const res = await rf(urls.preview_start, { json: {}, quiet: true });
        runJob("preview", res.job, () => loadPanel("preview", { open: true, scroll: false }));
      } catch (e) {
        toast(e.message, "err");
        previewRetry.disabled = false;
      }
      return;
    }
    const confirmRun = hit("[data-confirm-run]");
    if (confirmRun) {
      if (!confirmed(confirmRun)) return;
      confirmRun.disabled = true;
      try {
        const res = await rf(urls.confirm, { json: {}, quiet: true });
        sections.done("preview", notes.preview);
        // ⑥も描き直す。描き直さないと、確定前の中身と押せない［確定してMarkdownを作成］が居座る
        runJob("done", res.job, async () => {
          await loadPanel("done", { open: true });
          await loadPanel("preview", { open: false, scroll: false });
        });
      } catch (e) {
        toast(e.message, "err");
        confirmRun.disabled = false;
      }
      return;
    }

    // ダウンロード（押すとこの取り込みのデータは消える。design.md 3.3）
    const download = hit("[data-done-page] a[href$='download.zip']");
    if (download) {
      if (!confirmed(download)) return;
      // 渡した時点でサーバ側は消える（purge_after_send）ので、画面を離れるときに捨てるものはもう無い
      importId = null;
      // 2回押すと、2回目は消えたあとの取得になって画面ごと404へ飛ぶ。1回目で無効にする。
      // ただし href を今すぐ外すと、この click の「リンクを辿る処理」自体が取り消されて
      // zip が1バイトも落ちてこない（a の activation behavior は dispatch のあとに href を見る）。
      // 次の一拍に回して、ダウンロードが始まってから外す（2026-09-23 のレビューで実測）
      setTimeout(() => {
        download.removeAttribute("href");          // href が無い a はリンクとして押せなくなる
        download.setAttribute("aria-disabled", "true");   // 見た目は .btn[aria-disabled] で薄くなる
      }, 0);
      setTimeout(() => {
        body("done").replaceChildren(
          el("p", { text: "ダウンロードしました。この取り込みのデータ（元のファイル・読み込んだ内容・作った Markdown）は"
                          + "サーバーから消えています。" }),
          el("p", { class: "hint", text: "zip を開いて、中の .md を LightRAG の画面にドラッグしてください。" }),
          el("div", { class: "form-actions" },
            el("button", { type: "button", class: "btn primary", "data-restart-done": "", text: "別の表を取り込む" })));
        setNote("done", "ダウンロード済み");
      }, 4000);
      return;
    }
    if (hit("[data-restart-done]")) {
      window.location.href = page.dataset.newUrl;
      return;
    }

    // AI整形へ / 内容の確認へ
    if (hit("[data-ai-next]")) {
      sections.done("ai", notes.ai);
      await loadPanel("preview");
    }
  });

  // ---- 2 読み取り方（変えるたびに保存し、下の段をやり直す） ---------------------------------
  async function saveSource() {
    const form = page.querySelector("[data-source-form]");
    if (!form) return;
    const data = {};
    form.querySelectorAll("[data-source-field]").forEach((input) => {
      if (input.type === "radio") {
        if (input.checked) data[input.name] = input.value;
      } else if (input.type === "checkbox") {
        data[input.name] = input.checked;
      } else {
        data[input.name] = input.value;
      }
    });
    sections.working("source", "読み取り方を保存しています…");
    try {
      const res = await rf(urls.source, { json: data, quiet: true });
      if (res.note) setNote("source", res.note);
      // 「表の範囲」はこのあとすぐ読み込み直すので、ここでは消さない
      (res.reset || []).forEach((name) => { if (name !== "layout") clearStep(name); });
      await loadPanel("layout", { open: true, scroll: false });
    } catch (e) {
      toast(e.message, "err");
    } finally {
      sections.working("source", "");
    }
  }

  page.addEventListener("change", (event) => {
    if (event.target.closest("[data-source-field]")) {
      saveSource();
      return;
    }
    const editorRow = event.target.closest("[data-columns-editor] tr[data-col]");
    if (editorRow) onEditorChange(editorRow, event.target);
  });

  // ---- 3 表の範囲 ---------------------------------------------------------------------
  const KINDS = ["header", "data", "subtotal", "note", "continuation", "excluded", "title", "blank"];
  let detectTimer = null;
  let detectSeq = 0;

  // 行番号の読み方はサーバー（views/tables.py の _int_list・_row_no）と同じ: 全角は半角にそろえ、
  // 数字だけ・「3-4」（10行まで）だけを読む。それ以外（「3a」など）は読まない
  const parseEnd = (text) => {
    const x = String(text ?? "").normalize("NFKC").trim();
    return /^[0-9]+$/.test(x) && Number(x) > 0 ? Number(x) : null;
  };
  const parseRows = (text) => {
    const out = new Set();
    String(text || "").normalize("NFKC").split(/[,、\s]+/).forEach((x) => {
      const m = /^([0-9]+)-([0-9]+)$/.exec(x);
      if (m) {
        const a = Number(m[1]);
        const b = Number(m[2]);
        if (a > 0 && a <= b && b - a < 10) for (let n = a; n <= b; n += 1) out.add(n);
        return;
      }
      const n = parseEnd(x);
      if (n) out.add(n);
    });
    return [...out].sort((a, b) => a - b);
  };

  function applyLayout(root, info) {
    root.querySelectorAll("tr[data-row]").forEach((tr) => {
      const rc = info.rows[tr.dataset.row];
      KINDS.forEach((k) => tr.classList.remove(`row-${k}`));
      if (rc) {
        tr.classList.add(`row-${rc.kind}`);
        tr.title = rc.reason || "";
      } else {
        tr.classList.add("row-blank");
        tr.title = "";
      }
      tr.classList.toggle("is-picked",
        info.header_rows.includes(Number(tr.dataset.row)) || Number(tr.dataset.row) === info.data_end);
    });
    root.querySelector("[data-kind]").textContent = info.table_kind_label;
    // 見出し行が見つからないと data_end < data_start になる（存在しない行番号を出さない）
    const range = info.data_end < info.data_start ? "データの行が見つかりません" : `${info.data_start}〜${info.data_end}行目`;
    root.querySelector("[data-range]").textContent = range;
    root.querySelector("[data-counts]").textContent =
      Object.entries(info.counts).map(([k, v]) => `${k} ${v}行`).join("／");
    root.querySelector("[data-headers]").textContent = info.headers.join("、");
    const warnings = root.querySelector("[data-warnings]");
    warnings.replaceChildren(...info.warnings.map((w) => el("li", { text: w })));
    const crosstab = info.table_kind === "crosstab";
    root.querySelector("[data-crosstab]").hidden = !crosstab;
    const errors = info.errors || [];
    const errorBox = root.querySelector("[data-layout-error]");
    errorBox.textContent = errors.join("\u3000");
    errorBox.hidden = errors.length === 0;
    root.querySelector("[data-layout-submit]").disabled = crosstab || errors.length > 0;
    const headerInput = root.querySelector("[data-header-input]");
    if (!headerInput.value.trim()) headerInput.value = info.header_rows.join(",");
    root.querySelector("[data-end-input]").placeholder = `自動（${info.data_end}行目）`;
    setNote("layout", range);
  }

  function detectLayout() {
    const root = page.querySelector("[data-layout-page]");
    if (!root) return;
    clearTimeout(detectTimer);
    detectTimer = setTimeout(async () => {
      const mine = ++detectSeq;
      root.setAttribute("aria-busy", "true");
      // 大きい表は判定に数秒かかる。古い表示のままだと「効いていない」ように見えるので、その場に出す
      const rangeEl = root.querySelector("[data-range]");
      if (rangeEl) rangeEl.textContent = "判定しています…";
      try {
        const info = await rf(urls.detect, {
          json: { header_rows: parseRows(root.querySelector("[data-header-input]").value),
                  data_end: parseEnd(root.querySelector("[data-end-input]").value) },
          quiet: true });
        if (mine === detectSeq) applyLayout(root, info);
      } catch (e) {
        if (rangeEl) rangeEl.textContent = "判定できませんでした";
        toast(e.message, "err");
      } finally {
        root.removeAttribute("aria-busy");
      }
    }, 250);
  }

  function pickRow(th, shift) {
    const root = th.closest("[data-layout-page]");
    const n = Number(th.parentElement.dataset.row);
    const headerInput = root.querySelector("[data-header-input]");
    // 行番号のクリックは見出し行の指定（Shift で2段）。データの終わりは入力欄で決める
    if (shift) {
      const all = [...new Set([...parseRows(headerInput.value), n])].sort((a, b) => a - b);
      headerInput.value = [all[0], all[all.length - 1]].filter((v, i, a) => a.indexOf(v) === i).join(",");
    } else {
      headerInput.value = String(n);
    }
    detectLayout();
  }

  page.addEventListener("input", (event) => {
    if (event.target.closest("[data-header-input], [data-end-input]")) detectLayout();
  });

  // ---- 4 列の対応づけ -------------------------------------------------------------------
  // ［変更する］: 要約を隠して表を出す。表は要約のときも DOM にあるので、保存で送る中身は変わらない
  function openColumnsTable(editor) {
    if (!editor) return;
    const summary = editor.querySelector("[data-columns-summary]");
    const table = editor.querySelector("[data-columns-table]");
    if (summary) summary.hidden = true;
    if (table) table.hidden = false;
  }

  function onEditorChange(tr, input) {
    const editor = tr.closest("[data-columns-editor]");
    const field = input.dataset.field;
    if (field === "use") tr.classList.toggle("is-unused", !input.checked);
    // 経過の記録（log）は1列だけ。ほかの行が選んでいたら「その他」に戻し、この行は「使う」にする
    if (field === "role" && input.value === "log") {
      editor.querySelectorAll("tr[data-col]").forEach((other) => {
        if (other === tr) return;
        const role = other.querySelector("[data-field=role]");
        if (role && role.value === "log") role.value = "attribute";
      });
      tr.querySelector("[data-field=use]").checked = true;
      tr.classList.remove("is-unused");
    }
  }

  async function saveColumns(button) {
    const editor = page.querySelector("[data-columns-editor]");
    if (!editor) return;
    const errors = editor.querySelector("[data-editor-errors]");
    const name = editor.querySelector("[data-setting=name]")?.value || "";
    // 送るのは「使う・役割」だけ。キー・型・単位・出し方はサーバーが見出しと値から決める
    const columns = [...editor.querySelectorAll("tr[data-col]")].map((tr) => ({
      index: Number(tr.dataset.index),
      use: tr.querySelector("[data-field=use]").checked,
      role: tr.querySelector("[data-field=role]").value,
    }));
    // ファイルの分け方（チェックした列の位置。1つも無ければ「分けない」として空で送る）
    const group_by = [...editor.querySelectorAll("[data-group-col]")]
      .filter((box) => box.checked).map((box) => Number(box.dataset.index));
    button.disabled = true;
    errors.replaceChildren();
    try {
      const res = await rf(editor.dataset.saveUrl, { json: { name, columns, group_by }, quiet: true });
      setNote("columns", name);
      sections.done("columns", name);
      afterAction(res);
    } catch (e) {
      // 問題が複数あれば全部並べる（1件だけ直して保存し直す、を繰り返さずに済む）
      const messages = e.errors && e.errors.length ? e.errors : [e.message];
      errors.replaceChildren(...messages.map((m) => el("li", { text: m })));
      toast(messages.length > 1 ? `${messages.length}件の問題があります` : messages[0], "err");
    } finally {
      button.disabled = false;
    }
  }

  // ---- 5 AI整形 -----------------------------------------------------------------------
  const trials = [];

  // 分割プレビュー: 左に原文（セルのまま）、右に「順番1、順番2…」のかたまり。同じ色が対応する
  // （利用者の指示 2026-09-21「順番1、順番2の表示にしてほしい。また、原文と並べて見れるようにしたい」）。
  // s.id（s1, s2…）は AI との受け渡し用で画面には出さない。位置は s.start / s.end（data.text の中）。
  const SEG_COLORS = 8;

  const segmentCards = (data) => {
    const list = el("ol", { class: "seg-list" });
    data.segments.forEach((s, i) => {
      const meta = el("div", { class: "seg-meta" },
        el("strong", { class: "seg-no", text: `順番${s.no || i + 1}` }),
        el("span", { class: s.when_estimated ? "est" : "", text: `日付: ${s.when}` }),
        el("span", { class: s.author_estimated ? "est" : "", text: `記入者: ${s.author || "なし"}` }),
        s.marks.length ? el("span", { class: "muted", text: `印: ${s.marks.join("・")}` }) : null,
        s.identifiers.length ? el("span", { class: "muted", text: `識別子: ${s.identifiers.join("、")}` }) : null);
      list.append(el("li", { class: `seg-card seg-c${i % SEG_COLORS}`, "data-seg": s.id, tabindex: "0",
                             title: "クリックすると原文のその部分を示します" },
        meta, el("div", { class: "pre", text: s.body })));
    });
    return list;
  };

  // 原文（改行もそのまま）。分けた部分ごとに順番の色で塗る。位置が無い・重なる部分は塗らずに残す
  const sourceText = (data) => {
    const pre = el("pre", { class: "split-source", "data-split-source": "" });
    const text = data.text || "";
    const spans = data.segments
      .map((s, i) => ({ id: s.id, no: s.no || i + 1, start: Number(s.start), end: Number(s.end), i }))
      .filter((s) => Number.isFinite(s.start) && Number.isFinite(s.end) && s.end > s.start)
      .sort((a, b) => a.start - b.start);
    let pos = 0;
    for (const s of spans) {
      const start = Math.max(s.start, pos);
      if (start >= s.end) continue;
      if (start > pos) pre.append(text.slice(pos, start));
      pre.append(el("mark", { class: `seg-c${s.i % SEG_COLORS}`, "data-seg": s.id, title: `順番${s.no}`, tabindex: "0" },
        text.slice(start, s.end)));
      pos = s.end;
    }
    if (pos < text.length) pre.append(text.slice(pos));
    return pre;
  };

  const splitCompare = (data) => el("div", { class: "split-compare", "data-split-compare": "" },
    el("div", { class: "split-col" }, el("h4", { text: "原文（セルのまま）" }), sourceText(data)),
    el("div", { class: "split-col" }, el("h4", { text: "分けた結果" }), segmentCards(data)));

  // 順番のカード ⇄ 原文の色の部分を行き来する（クリック・Enter）
  const jumpSegment = (node) => {
    const box = node.closest("[data-split-compare]");
    if (!box) return;
    const id = node.dataset.seg;
    const inSource = node.matches("mark");
    box.querySelectorAll("[data-seg].is-active").forEach((n) => n.classList.remove("is-active"));
    const other = box.querySelector(inSource ? `li[data-seg="${id}"]` : `mark[data-seg="${id}"]`);
    node.classList.add("is-active");
    if (other) {
      other.classList.add("is-active");
      other.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
  };
  page.addEventListener("keydown", (event) => {
    const node = event.target.closest("[data-split-compare] [data-seg]");
    if (node && (event.key === "Enter" || event.key === " ")) {
      event.preventDefault();
      jumpSegment(node);
    }
  });

  const trialCard = (res) => {
    const labels = { ok: "OK", flagged: "要確認", error: "エラー", rule_only: "ルールのみ", skipped: "対象外" };
    const card = el("div", { class: "trial-card" },
      el("div", { class: "seg-meta" }, el("strong", { text: res.row_key }),
        el("span", { text: labels[res.status] || res.status || "" }),
        res.reason ? el("span", { class: "muted", text: res.reason }) : null,
        res.cached ? el("span", { class: "muted", text: "（保存済みの応答を使用）" }) : null));
    if (res.error) card.append(el("p", { class: "warn", text: res.error }));
    if (res.points.length) card.append(el("p", { text: "抜き出した項目:" }), el("pre", { class: "md-view", text: res.points.join("\n") }));
    if (res.issues.length) card.append(el("ul", { class: "list-plain warn" }, ...res.issues.map((m) => el("li", { text: m }))));
    card.append(el("details", {}, el("summary", { text: "Markdown のプレビュー" }), el("pre", { class: "md-view", text: res.markdown })));
    return card;
  };

  const runOptions = (root) => ({
    scope: root.querySelector("[data-run-scope]")?.value || "pending",
    concurrency: Number(root.querySelector("[data-run-concurrency]")?.value || 1),
  });

  page.addEventListener("click", async (event) => {
    const root = page.querySelector("[data-ai-page]");
    if (!root) return;
    const hit = (selector) => event.target.closest(selector);

    const segNode = hit("[data-split-compare] [data-seg]");
    if (segNode) {
      jumpSegment(segNode);
      return;
    }

    if (hit("[data-split-run]")) {
      const box = root.querySelector("[data-split-result]");
      const rowKey = root.querySelector("[data-split-row]").value;
      box.textContent = "分割しています…";
      try {
        const data = await rf(root.dataset.splitUrl, { json: { row_key: rowKey }, quiet: true });
        // AI接続が未設定のときは「送る」だけだと今この操作で送ったように読めるので、送られないことを書く
        const routeText = data.route === "ai"
          ? (data.ai_ready === false ? "送る（画面右上の「AI接続」の設定後。今は送っていません）" : "送る")
          : `送らない（${data.reason}）`;
        const count = data.segments.length;
        const lead = count
          ? `${count}つに分かれました（日付や記入者がオレンジのものは、書き方から推定したものです）。`
          : "分けられませんでした（日付などの区切りが見つかりません）。";
        const parts = [
          el("p", { class: "hint", text: `${lead}AIに送るか: ${routeText}` }),
          splitCompare(data),
        ];
        if (data.notes.length) parts.push(el("ul", { class: "list-plain warn" }, ...data.notes.map((n) => el("li", { text: n }))));
        parts.push(el("h4", { text: "Markdown での時系列" }), el("pre", { class: "md-view", text: data.timeline.join("\n") }));
        if (data.sent_text) {
          parts.push(el("details", {}, el("summary", { text: "送信内容を表示" }), el("pre", { class: "md-view", text: data.sent_text })));
        }
        box.replaceChildren(...parts);
      } catch (e) {
        box.textContent = "";
        toast(e.message, "err");
      }
      return;
    }

    const trialRun = hit("[data-trial-run]");
    if (trialRun) {
      const external = root.querySelector("[data-trial-external]");
      if (external && !external.checked) {
        toast("外部に送信されることを確認して、チェックを入れてください。", "err");
        return;
      }
      const keys = JSON.parse(root.dataset.trialKeys || "[]");
      const status = root.querySelector("[data-trial-status]");
      const result = root.querySelector("[data-trial-result]");
      trialRun.disabled = true;
      result.replaceChildren();
      trials.length = 0;
      for (let i = 0; i < keys.length; i += 1) {
        status.textContent = `${i + 1} / ${keys.length}行目を処理しています…`;
        try {
          const res = await rf(root.dataset.trialUrl,
            { json: { row_key: keys[i], confirm_external: Boolean(external?.checked) }, quiet: true });
          trials.push(...(res.stats || []));
          result.append(trialCard(res));
        } catch (e) {
          toast(e.message, "err");
          status.textContent = `止まりました: ${e.message}`;
          trialRun.disabled = false;
          return;
        }
      }
      status.textContent = `${keys.length}行の試し実行が終わりました。`;
      trialRun.disabled = false;
      return;
    }

    const estimateRun = hit("[data-estimate-run]");
    if (estimateRun) {
      const out = root.querySelector("[data-estimate]");
      estimateRun.disabled = true;
      out.textContent = "見積もっています…";
      try {
        const r = await rf(root.dataset.estimateUrl, { json: { ...runOptions(root), trials }, quiet: true });
        out.textContent = `AIに送る行 ${r.ai_rows}件（呼び出し ${r.calls}回、保存済み ${r.cached}件、ルールのみ ${r.rule_only_rows}件）`
          + `／入力 約${r.tokens_in}トークン・出力 約${r.tokens_out}トークン／${r.duration_text || `約${Math.ceil(r.minutes || 0)}分`}`
          + (r.basis === "trial" ? "（試し実行の実測から）" : "（目安）");
      } catch (e) {
        out.textContent = "";
        toast(e.message, "err");
      } finally {
        estimateRun.disabled = false;
      }
      return;
    }

    const aiRun = hit("[data-ai-run]");
    if (aiRun) {
      const external = root.querySelector("[data-run-external]");
      if (external && !external.checked) {
        toast("外部に送信されることを確認して、チェックを入れてください。", "err");
        return;
      }
      aiRun.disabled = true;
      try {
        const res = await rf(root.dataset.runUrl, { json: { ...runOptions(root), confirm_external: Boolean(external?.checked) }, quiet: true });
        window.ragAiHeader?.render(res.ai_status);   // 始める前に接続を確かめた結果をヘッダーにも映す
        await loadPanel("ai", { open: true, scroll: false });
      } catch (e) {
        window.ragAiHeader?.render(e.data?.ai_status);
        toast(e.message, "err");
        aiRun.disabled = false;
      }
      return;
    }

    const control = hit("[data-ai-control]");
    if (control) {
      if (!confirmed(control)) return;
      const action = control.dataset.aiControl;
      control.disabled = true;
      try {
        await rf(root.dataset[`${action}Url`], { json: {}, quiet: true });
      } catch (e) {
        toast(e.message, "err");
      }
      await loadPanel("ai", { open: true, scroll: false });
    }
    // AI接続の設定はこの段には無い（ヘッダー右上の「AI接続」。上の window.ragAiHeader）
  });

  // ---- 6 内容の確認: md の中身 ------------------------------------------------------------
  async function showFile(button) {
    const root = button.closest("[data-preview-page]");
    const title = root.querySelector("[data-file-title]");
    const text = root.querySelector("[data-file-text]");
    const name = button.dataset.fileOpen;
    root.querySelectorAll("tr[data-file]").forEach((tr) => tr.classList.toggle("is-current", tr.dataset.file === name));
    title.textContent = `${name}（読み込み中…）`;
    try {
      const data = await rf(`${root.dataset.fileUrl}?name=${encodeURIComponent(name)}`, { quiet: true });
      title.textContent = name;
      text.textContent = data.text;
    } catch (e) {
      title.textContent = name;
      toast(e.message, "err");
    }
  }
})();
