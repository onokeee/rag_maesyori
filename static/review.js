// 帳票取り込み（/forms）: 1枚の画面で ファイルを置く → 種類とシート → 読み取り結果 → 確定してダウンロード。
// 画面は移動しない。どの操作も fetch でルートを呼び、返ってきた HTML の断片をその場に入れ替える。
(function () {
  "use strict";

  const page = document.getElementById("formsPage");
  if (!page) return;

  const toast = (msg, kind) => (window.App && window.App.toast ? window.App.toast(msg, kind) : null);

  // ---- 段の開け閉て（app.js の window.ragSections。無ければ自前で同じことをする） ----
  const sections = {
    el(step) {
      return typeof step === "string" ? page.querySelector('[data-step="' + step + '"]') : step;
    },
    open(step, scroll) {
      const el = sections.el(step);
      if (!el) return null;
      if (window.ragSections) return window.ragSections.open(el, { scroll: !!scroll });
      el.classList.add("is-open");
      if (scroll) el.scrollIntoView({ behavior: "smooth", block: "start" });
      return el;
    },
    // 済んだ段。見出しに要約（summary）を出して畳み、クリックで開き直せるようにする
    done(step, summary, next) {
      const el = sections.el(step);
      if (!el) return null;
      if (window.ragSections) return window.ragSections.done(el, summary, next);
      el.classList.add("is-done");
      el.classList.remove("is-open");
      return el;
    },
    // まだの段に戻す（この段より下も全部まっさらにする）
    close(step) {
      const el = sections.el(step);
      if (!el) return null;
      el.classList.remove("is-open", "is-done");
      const label = el.querySelector("[data-step-summary]");
      if (label) label.textContent = "";
      if (window.ragSections) window.ragSections.refresh();
      return el;
    },
    note(step, text) {
      const el = sections.el(step);
      const label = el && el.querySelector("[data-step-summary]");
      if (label) { label.textContent = text || ""; label.title = text || ""; }
    },
    show(step) {
      return sections.open(step, true);
    },
  };

  // ---- 通信（shell の window.ragFetch。無ければ自前。エラーは1か所でトースト） ----
  async function send(url, options) {
    if (window.ragFetch) return window.ragFetch(url, options);
    const res = await fetch(url, Object.assign({ headers: { Accept: "application/json" } }, options || {}));
    if (res.status === 204) return {};
    const text = await res.text();
    let data = {};
    try { data = text ? JSON.parse(text) : {}; } catch (e) { data = { html: text }; }
    if (!res.ok) {
      const error = new Error(data.error || "うまくいきませんでした（HTTP " + res.status + "）");
      error.status = res.status;
      error.data = data;
      throw error;
    }
    return data;
  }

  function postJson(url, body) {
    return send(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify(body || {}),
    });
  }

  // ---- 作業中の表示（同じ画面の中に出す。閉じたら消えてよい） ----
  function work(name, text) {
    const line = page.querySelector('[data-work="' + name + '"]');
    if (!line) return;
    const span = line.querySelector("[data-work-text]");
    if (span) span.textContent = text || "";
    line.hidden = !text;
  }

  // ---- 画面の状態 ----
  let docs = [];        // [{id, file_name, state}]
  let currentId = null;
  const el = {
    typeBody: page.querySelector("[data-type-body]"),
    reviewBody: page.querySelector("[data-review-body]"),
    finishBody: page.querySelector("[data-finish-body]"),
    docTabs: page.querySelector("[data-doc-tabs]"),
    docName: page.querySelector('#step-type [data-step-summary]'),
    fileNote: page.querySelector('#step-file [data-step-summary]'),
    saveStatus: page.querySelector('#step-review [data-step-summary]'),
  };
  const ids = () => docs.map((d) => d.id).join(",");
  const setStatus = (text) => { if (el.saveStatus) el.saveStatus.textContent = text; };

  // ダウンロードしないまま画面を離れたら、この画面で取り込んだ分は捨てる（利用者の指示 2026-09-20）。
  // docs に残っているものが「まだダウンロードしていない分」そのもの（ダウンロード・削除で docs から抜ける）。
  const guard = (window.ragDiscard || { watch: () => ({ now: async () => {}, clear() {}, arm() {} }) })
    .watch(page.dataset.discardUrl || "/forms/discard", () => ({ doc_ids: docs.map((d) => d.id) }));

  // ---- 1 ファイルを置く ----
  const uploadForm = document.getElementById("uploadForm");
  if (uploadForm) {
    // 置いた（またはクリックで選んだ）時点でそのまま読み取る。ボタンは押さなくてよい
    uploadForm.querySelector("input[type=file]")?.addEventListener("change", (event) => {
      if (event.target.files && event.target.files.length) uploadForm.requestSubmit();
    });
    uploadForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const input = uploadForm.querySelector("input[type=file]");
      const files = input && input.files ? Array.from(input.files) : [];
      if (!files.length) { toast("ファイルを置いてください", "err"); return; }
      const button = uploadForm.querySelector("button[type=submit]");
      if (button) button.disabled = true;
      work("upload", files.length > 1 ? files.length + "件のファイルを取り込んでいます…" : "ファイルを取り込んでいます…");
      try {
        // 新しいファイルを置いたら、前の分（ダウンロードしていない帳票）はその場で捨てる
        if (docs.length) { await guard.now(); resetPage(); }
        const data = new FormData();
        files.forEach((f) => data.append("file", f));
        const res = await send(page.dataset.uploadUrl, { method: "POST", body: data });
        (res.errors || []).forEach((m) => toast(m, "err"));
        docs = (res.docs || []).map((d) => Object.assign({ state: "unread" }, d));
        if (el.fileNote) {
          el.fileNote.textContent = docs.length > 1
            ? docs.length + "件を取り込みました（1件ずつ読み取ります）"
            : (docs[0] ? docs[0].file_name : "");
        }
        if (input) { input.value = ""; input.dispatchEvent(new Event("change", { bubbles: true })); }
        renderDocTabs();
        await openDoc(docs[0].id);
      } catch (e) {
        toast(e.message, "err");
      } finally {
        work("upload", "");
        if (button) button.disabled = false;
      }
    });
  }

  // ---- 取り込む帳票の切り替え（まとめて置いたとき） ----
  function renderDocTabs() {
    if (!el.docTabs) return;
    el.docTabs.hidden = docs.length < 2;
    el.docTabs.textContent = "";
    if (docs.length < 2) return;
    docs.forEach((d, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "doc-tab" + (d.id === currentId ? " is-current" : "") +
        (d.state === "confirmed" || d.state === "modified" ? " is-done" : "");
      b.setAttribute("role", "tab");
      b.setAttribute("aria-selected", d.id === currentId ? "true" : "false");
      b.dataset.openDoc = String(d.id);
      b.textContent = (i + 1) + ". " + d.file_name;
      el.docTabs.appendChild(b);
    });
  }

  function docOf(id) { return docs.find((d) => d.id === id) || null; }

  async function openDoc(id) {
    currentId = id;
    const doc = docOf(id);
    if (el.docName) el.docName.textContent = doc ? doc.file_name : "";
    renderDocTabs();
    clearReview();
    setStatus("");
    sections.done("file");
    sections.open("type");
    el.typeBody.innerHTML = '<p class="muted">読み込んでいます…</p>';
    try {
      const res = await send("/forms/" + id + "/type");
      el.typeBody.innerHTML = res.html || "";
      bindTypeForm();
    } catch (e) {
      el.typeBody.innerHTML = '<p class="warn"></p>';
      el.typeBody.querySelector(".warn").textContent = e.message;
    }
    await refreshFinish();
    sections.show("type");
  }

  // ---- 2 帳票の種類とシート ----
  function bindTypeForm() {
    const form = el.typeBody.querySelector("[data-type-form]");
    if (!form) return;
    let suggested = {};
    try { suggested = JSON.parse(form.dataset.suggested || "{}"); } catch (e) { suggested = {}; }
    form.querySelectorAll("input[name=pattern_id]").forEach((radio) => {
      radio.addEventListener("change", () => {
        const sheets = suggested[radio.value];
        if (!sheets || !sheets.length) return;
        form.querySelectorAll("input[name=sheets]").forEach((box) => { box.checked = sheets.includes(box.value); });
      });
    });
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = form.querySelector("button[type=submit]");
      if (button) button.disabled = true;
      const doc = docOf(Number(form.dataset.doc));
      work("read", (doc ? doc.file_name : "帳票") + " を読み取っています…");
      try {
        const res = await send(form.dataset.readUrl, { method: "POST", body: new FormData(form) });
        showReview(res.html);
        await refreshFinish();
        sections.done("type");
        sections.show("review");
      } catch (e) {
        toast(e.message, "err");
      } finally {
        work("read", "");
        if (button) button.disabled = false;
      }
    });
  }

  // 取り込みをやめる（確認ダイアログは app.js が先に出す。ここは「はい」のあとの本番）
  page.addEventListener("click", async (event) => {
    const drop = event.target.closest("[data-drop-doc][data-confirmed='1']");
    if (!drop) return;
    const id = Number(drop.dataset.dropDoc);
    try {
      await postJson("/forms/" + id + "/delete", {});
      docs = docs.filter((d) => d.id !== id);
      toast("この帳票の取り込みをやめました", "ok");
      if (docs.length) {
        renderDocTabs();
        await openDoc(docs[0].id);
      } else {
        resetPage();
      }
    } catch (e) { toast(e.message, "err"); }
  });

  function resetPage() {
    docs = [];
    currentId = null;
    el.typeBody.textContent = "";
    el.finishBody.textContent = "";
    clearReview();
    if (el.fileNote) el.fileNote.textContent = "";
    if (el.docName) el.docName.textContent = "";
    if (el.docTabs) { el.docTabs.hidden = true; el.docTabs.textContent = ""; }
    ["type", "review", "finish"].forEach(sections.close);
    sections.close("file");
    sections.show("file");
  }

  // ---- 3 読み取り結果 ----
  const SOURCES = {
    auto: ["自動で読み取り", "blue"], ai: ["AIが入力（要確認）", "violet"],
    manual: ["手で修正", "teal"], blank: ["空欄", "gray"],
  };
  let root = null;        // #review
  let form = null;        // #reviewForm
  let fieldBoxes = [];
  let activeBox = null;
  let sheetGrids = [];
  let sheetTabs = [];
  let paneTabs = [];
  let dirty = false;
  let timer = null;
  let saving = null;
  let pageToken = "";

  const inputOf = (box) => box._activeCell || box.querySelector("input.input, textarea.input, textarea.cell-input");
  const valueOf = (box) => box.querySelector("[name^='value-']");
  const currentVersion = () => (root ? root.dataset.version || "" : "");

  // 別の帳票に切り替えるときは、前の帳票の入力欄・版を持ち越さない
  function clearReview() {
    clearTimeout(timer);
    dirty = false;
    saving = null;
    activeBox = null;
    root = null;
    form = null;
    fieldBoxes = [];
    sheetGrids = [];
    sheetTabs = [];
    paneTabs = [];
    el.reviewBody.textContent = "";
    sections.close("review");
    sections.close("finish");
  }

  function showReview(html) {
    clearReview();
    el.reviewBody.innerHTML = html || "";
    sections.open("review");
    bindReview();
    setStatus("");
  }

  function bindReview() {
    root = el.reviewBody.querySelector("#review");
    form = el.reviewBody.querySelector("#reviewForm");
    if (!root || !form) { fieldBoxes = []; return; }
    pageToken = (window.crypto && crypto.randomUUID)
      ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2);
    fieldBoxes = Array.from(form.querySelectorAll("[data-field]"));
    sheetGrids = Array.from(root.querySelectorAll(".sheet-grid"));
    sheetTabs = Array.from(root.querySelectorAll("[data-sheet-tab]"));
    paneTabs = Array.from(root.querySelectorAll("[data-pane-tab]"));

    sheetTabs.forEach((tab) => tab.addEventListener("click", () => {
      const g = sheetGrids[parseInt(tab.dataset.sheetTab, 10)];
      if (g) showSheet(g.dataset.sheet);
    }));

    paneTabs.forEach((tab) => tab.addEventListener("click", () => {
      paneTabs.forEach((t) => t.setAttribute("aria-selected", t === tab ? "true" : "false"));
      form.querySelectorAll("[data-pane]").forEach((p) => { p.hidden = p.dataset.pane !== tab.dataset.paneTab; });
    }));

    fieldBoxes.forEach((box) => {
      const input = box.querySelector("input.input, textarea.input");
      if (input) input.addEventListener("focus", () => { activeBox = box; highlight(box); });
      box.addEventListener("focusin", (event) => {
        if (!event.target.classList.contains("cell-input")) return;
        box._activeCell = event.target;
        activeBox = box;
        highlight(box);
      });
      const table = box.querySelector("[data-table-editor]");
      if (table) {
        table.addEventListener("input", () => syncTable(box));
        box.addEventListener("click", (event) => onTableClick(box, table, event));
      }
    });

    // セルをクリック → 選んでいる項目にその値を入れる
    root.addEventListener("click", (event) => {
      const td = event.target.closest("td[data-cell]");
      if (!td || !activeBox) return;
      const input = inputOf(activeBox);
      if (!input) return;
      input.value = td.textContent.trim();
      input.dispatchEvent(new Event("input", { bubbles: true }));
      root.querySelectorAll("td.hl-value").forEach((c) => c.classList.remove("hl-value"));
      td.classList.add("hl-value");
      input.focus();
      toast("セル " + td.dataset.cell + " の値を入れました", "info");
    });

    form.addEventListener("input", (event) => {
      if (!event.target.name || !event.target.name.startsWith("value-")) return;
      dirty = true;
      setStatus("未保存の変更があります");
      clearTimeout(timer);
      timer = setTimeout(saveNow, 700);
    });

    const nextBtn = el.reviewBody.querySelector("[data-next-issue]");
    if (nextBtn) nextBtn.addEventListener("click", gotoNextIssue);
  }

  function showSheet(name) {
    let shown = null;
    sheetGrids.forEach((g, i) => {
      const on = g.dataset.sheet === name;
      g.hidden = !on;
      if (sheetTabs[i]) sheetTabs[i].setAttribute("aria-selected", on ? "true" : "false");
      if (on) shown = g;
    });
    return shown;
  }

  const topLeft = (coord) => (coord || "").split(":")[0];

  function highlight(box) {
    root.querySelectorAll("td.hl-label, td.hl-value").forEach((td) => td.classList.remove("hl-label", "hl-value"));
    const sheet = box.dataset.sheet;
    let grid = sheet ? sheetGrids.find((g) => g.dataset.sheet === sheet) : null;
    if (!grid) return;
    if (grid.hidden) grid = showSheet(sheet);
    const label = grid.querySelector("td[data-cell='" + topLeft(box.dataset.labelCell) + "']");
    const value = grid.querySelector("td[data-cell='" + topLeft(box.dataset.valueCell) + "']");
    if (label) label.classList.add("hl-label");
    if (value) value.classList.add("hl-value");
    const target = value || label;
    if (target) target.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  // 明細表: セルの編集・行の追加と削除 → hidden の JSON に戻す
  function syncTable(box) {
    const hidden = box.querySelector("[data-table-value]");
    const table = box.querySelector("[data-table-editor]");
    if (!hidden || !table) return;
    const columns = Array.from(table.querySelectorAll("thead th[scope=col]")).map((th) => th.textContent.trim());
    const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) =>
      Array.from(tr.querySelectorAll(".cell-input")).map((e) => e.value));
    hidden.value = JSON.stringify({ columns, rows });
    hidden.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function onTableClick(box, table, event) {
    const remove = event.target.closest("[data-remove-table-row]");
    if (remove) {
      const tr = remove.closest("tr");
      if (box._activeCell && tr.contains(box._activeCell)) {
        box._activeCell = null;
        if (activeBox === box) activeBox = null;
      }
      tr.remove();
      syncTable(box);
      return;
    }
    if (!event.target.closest("[data-add-table-row]")) return;
    const body = table.querySelector("tbody");
    const n = table.querySelectorAll("thead th[scope=col]").length;
    const tr = document.createElement("tr");
    for (let i = 0; i < n; i += 1) {
      const td = document.createElement("td");
      const ta = document.createElement("textarea");
      ta.className = "cell-input";
      ta.rows = 1;
      td.appendChild(ta);
      tr.appendChild(td);
    }
    const td = document.createElement("td");
    td.className = "center";
    td.innerHTML = '<button type="button" class="btn small ghost" data-remove-table-row>行を削除</button>';
    tr.appendChild(td);
    body.appendChild(tr);
    const first = tr.querySelector(".cell-input");
    if (first) first.focus();
  }

  function values() {
    const out = {};
    fieldBoxes.forEach((box) => {
      const input = valueOf(box);
      if (input) out[box.dataset.field] = input.value;
    });
    return out;
  }

  function setVersion(v) {
    if (!v || !root) return;
    root.dataset.version = v;
    const input = form && form.querySelector("[data-version-input]");
    if (input) input.value = v;
  }

  async function saveNow() {
    if (saving) await saving;
    if (!dirty || !root) return;
    dirty = false;
    setStatus("保存中…");
    saving = (async () => {
      try {
        const body = { values: values(), version: currentVersion(), page_token: pageToken };
        const res = await fetch(root.dataset.draftUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(body),
        });
        if (!res.ok && res.status !== 204) {
          let msg = "保存できませんでした（" + res.status + "）";
          try { const d = await res.json(); if (d.error) msg = d.error; } catch (e) { /* 本文なし */ }
          throw new Error(msg);
        }
        setVersion(res.headers.get("X-Doc-Version"));
        const summary = await postJson(root.dataset.previewUrl,
          { values: values(), version: currentVersion(), page_token: pageToken });
        if (summary) applySummary(summary);
        setStatus("保存しました");
      } catch (err) {
        dirty = true;
        setStatus("保存できませんでした");
        toast(err.message, "err");
      }
    })();
    await saving;
    saving = null;
  }

  function applySummary(s) {
    Object.entries(s.counts || {}).forEach(([key, n]) => {
      const node = el.reviewBody.querySelector("[data-count='" + key + "']");
      if (!node) return;
      node.textContent = n;
      const chip = node.closest(".chip");
      if (chip) chip.classList.toggle("is-zero", !n);
    });
    fieldBoxes.forEach((box) => {
      const st = (s.fields || {})[box.dataset.field];
      if (!st) return;
      box.dataset.issue = st.issue ? "1" : "0";
      box.classList.toggle("is-issue", !!st.issue);
      const badge = box.querySelector("[data-source-badge]");
      const src = SOURCES[st.source] || SOURCES.auto;
      if (badge) { badge.textContent = src[0]; badge.className = "badge badge-" + src[1]; }
      const warn = box.querySelector("[data-warning]");
      if (warn) {
        const text = (st.missing_required ? "必須の項目が空欄です。" : "") + (st.warning || "");
        warn.textContent = text;
        warn.hidden = !text;
      }
    });
    const pre = el.reviewBody.querySelector("[data-md-preview]");
    if (pre && typeof s.markdown === "string") pre.textContent = s.markdown;
    const name = el.reviewBody.querySelector("[data-md-name]");
    if (name && s.file_name) name.textContent = s.file_name;
    // 確定してダウンロードの欄の「必須が空欄」の案内も合わせる
    const box = el.finishBody.querySelector("[data-missing-box]");
    if (box) {
      const list = s.missing_required || [];
      box.hidden = !list.length;
      const span = box.querySelector("[data-missing-list]");
      if (span) span.textContent = list.join("、");
    }
  }

  function gotoNextIssue() {
    const issues = fieldBoxes.filter((b) => b.dataset.issue === "1");
    if (!issues.length) { toast("要確認の項目はありません", "ok"); return; }
    paneTabs.forEach((t) => { if (t.dataset.paneTab === "fields") t.click(); });
    const idx = activeBox ? issues.findIndex((b) => fieldBoxes.indexOf(b) > fieldBoxes.indexOf(activeBox)) : 0;
    const box = issues[idx >= 0 ? idx : 0];
    box.scrollIntoView({ block: "center" });
    const input = inputOf(box);
    if (input) input.focus();
  }

  // 画面を閉じるときは、未保存の変更を送っておく
  window.addEventListener("beforeunload", () => {
    if (!dirty || !root) return;
    try {
      const blob = new Blob([JSON.stringify({ values: values(), version: currentVersion(), page_token: pageToken })],
        { type: "application/json" });
      navigator.sendBeacon(root.dataset.draftUrl, blob);
    } catch (e) { /* 送れなくても次の入力で保存される */ }
  });

  // ---- 4 確定してダウンロード ----
  async function refreshFinish() {
    if (!docs.length) { sections.close("finish"); return; }
    try {
      const url = page.dataset.finishUrl + "?ids=" + encodeURIComponent(ids()) +
        (currentId ? "&current=" + currentId : "");
      const res = await send(url);
      el.finishBody.innerHTML = res.html || "";
      // 読み取る前は灰色のままにして、見出しに理由を1行だけ出す（「読み取り結果」より先に開かない）
      if (res.read_yet) {
        sections.note("finish", "");
        sections.open("finish");
      } else {
        sections.close("finish");
        sections.note("finish", "読み取りが終わると確定できます");
      }
    } catch (e) {
      toast(e.message, "err");
    }
  }

  page.addEventListener("click", async (event) => {
    const confirmBtn = event.target.closest("[data-confirm-doc]");
    if (confirmBtn) {
      event.preventDefault();
      await confirmDoc(Number(confirmBtn.dataset.confirmDoc), confirmBtn);
      return;
    }
    const open = event.target.closest("[data-open-doc]");
    if (open) {
      event.preventDefault();
      await openDoc(Number(open.dataset.openDoc));
    }
  });

  async function confirmDoc(id, button) {
    const allow = el.finishBody.querySelector("[data-allow-missing]");
    if (button) button.disabled = true;
    work("confirm", "Markdown を作っています…");
    try {
      clearTimeout(timer);
      if (dirty) await saveNow();
      if (saving) await saving;
      await postJson("/forms/" + id + "/confirm",
        { version: currentVersion(), allow_missing: !!(allow && allow.checked) });
      const doc = docOf(id);
      if (doc) doc.state = "confirmed";
      renderDocTabs();
      await refreshFinish();
      sections.done("review", "確定しました");
      const next = docs.find((d) => d.state !== "confirmed" && d.state !== "modified");
      if (next && docs.length > 1) {
        toast("確定しました。次の帳票を読み取ります", "ok");
        await openDoc(next.id);
      } else {
        toast("確定しました", "ok");
        sections.show("finish");
      }
    } catch (e) {
      if (e.data && e.data.missing_required) {
        const box = el.finishBody.querySelector("[data-missing-box]");
        if (box) {
          box.hidden = false;
          const span = box.querySelector("[data-missing-list]");
          if (span) span.textContent = e.data.missing_required.join("、");
        }
      }
      toast(e.message, "err");
    } finally {
      work("confirm", "");
      if (button) button.disabled = false;
    }
  }

  // ダウンロードすると、渡した帳票のデータはサーバーから消えるので、画面も片付ける
  page.addEventListener("click", (event) => {
    const link = event.target.closest("a[data-download][data-confirmed='1']");
    if (!link) return;
    const rest = docs.filter((d) => d.state !== "confirmed" && d.state !== "modified");
    setTimeout(async () => {
      toast("ダウンロードしました。渡した分のデータはサーバーから消えました", "ok");
      if (!rest.length) { resetPage(); return; }
      docs = rest;                   // 未確定の帳票はサーバーに残っているので、続けて読み取れる
      renderDocTabs();
      await openDoc(docs[0].id);
    }, 1500);
  });
})();
