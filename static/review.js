// 帳票取り込み（/forms）: 1枚の画面で ファイルを置く → 種類とシート → 読み取り結果 → 確定してダウンロード。
// 画面は移動しない。どの操作も fetch でルートを呼び、返ってきた HTML の断片をその場に入れ替える。
//
// まとめて置いた帳票（同じフォーム）は、②で1回だけ種類とシートを決め、③に全部の読み取り結果を
// 縦に並べる（利用者の指示 2026-09-20）。重くならないよう、元のシートの表はその帳票が画面に
// 近づいた時点で読み込む（スクロールのたびに位置を見る）。
(function () {
  "use strict";

  const page = document.getElementById("formsPage");
  if (!page) return;

  const toast = (msg, kind) => (window.App && window.App.toast ? window.App.toast(msg, kind) : null);

  // ---- 段の開け閉て・通信は app.js（window.ragSections / window.ragFetch）を使う ----
  const rag = window.ragSections;
  const sections = {
    el: rag.el,
    open: (step, scroll) => rag.open(step, { scroll: !!scroll }),
    done: rag.done,
    close: rag.close,
    note: rag.note,
    show: (step) => rag.open(step, { scroll: true }),
  };
  // quiet: エラーの言い方はこの画面の側で決める（ragFetch の自動トーストと二重に出さない）
  const get = (url) => window.ragFetch(url, { quiet: true });
  const postJson = (url, body) => window.ragFetch(url, { json: body, quiet: true });
  const postForm = (url, data) => window.ragFetch(url, { form: data, quiet: true });
  const uuid = () => ((window.crypto && crypto.randomUUID)
    ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2));

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
  const el = {
    typeBody: page.querySelector("[data-type-body]"),
    reviewBody: page.querySelector("[data-review-body]"),
    finishBody: page.querySelector("[data-finish-body]"),
    docName: page.querySelector('#step-type [data-step-summary]'),
    fileNote: page.querySelector('#step-file [data-step-summary]'),
    saveStatus: page.querySelector('#step-review [data-step-summary]'),
  };
  const ids = () => docs.map((d) => d.id).join(",");
  const setStatus = (text) => { if (el.saveStatus) el.saveStatus.textContent = text; };
  const isDone = (d) => d.state === "confirmed" || d.state === "modified";
  const docOf = (id) => docs.find((d) => d.id === id) || null;

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
        const res = await postForm(page.dataset.uploadUrl, data);
        (res.errors || []).forEach((m) => toast(m, "err"));
        docs = (res.docs || []).map((d) => Object.assign({ state: "unread" }, d));
        if (el.fileNote) {
          el.fileNote.textContent = docs.length > 1
            ? docs.length + "件を取り込みました（まとめて読み取ります）"
            : (docs[0] ? docs[0].file_name : "");
        }
        if (input) { input.value = ""; input.dispatchEvent(new Event("change", { bubbles: true })); }
        await openType();
      } catch (e) {
        toast(e.message, "err");
      } finally {
        work("upload", "");
        if (button) button.disabled = false;
      }
    });
  }

  // ---- 2 帳票の種類とシート（置かれた分すべてに同じ設定を使う） ----
  async function openType() {
    if (!docs.length) { resetPage(); return; }
    clearReview();
    setStatus("");
    sections.done("file");
    sections.open("type");
    if (el.docName) {
      el.docName.textContent = docs.length > 1 ? docs.length + "件のファイル" : (docs[0] ? docs[0].file_name : "");
    }
    el.typeBody.innerHTML = '<p class="muted">読み込んでいます…</p>';
    try {
      const url = (page.dataset.typeUrl || "/forms/type") + "?ids=" + encodeURIComponent(ids());
      const res = await get(url);
      el.typeBody.innerHTML = res.html || "";
      bindTypeForm();
    } catch (e) {
      el.typeBody.innerHTML = '<p class="warn"></p>';
      el.typeBody.querySelector(".warn").textContent = e.message;
    }
    await refreshFinish();
    sections.show("type");
  }

  function bindTypeForm() {
    const form = el.typeBody.querySelector("[data-type-form]");
    if (!form) return;
    const parse = (text) => { try { return JSON.parse(text || "{}"); } catch (e) { return {}; } };
    const suggested = parse(form.dataset.suggested);
    const fileCounts = parse(form.dataset.fileCounts);

    // 種類を選び直したら、その種類で見つかったシートとファイルごとの項目数を出し直す
    function showCounts(patternId) {
      const counts = fileCounts[patternId] || {};
      form.querySelectorAll("[data-file-count]").forEach((node) => {
        const n = counts[node.dataset.fileCount];
        if (n === undefined) return;
        node.textContent = n;
        const chip = node.closest("[data-file-row]");
        if (chip) chip.classList.toggle("is-weak", n === 0);
      });
    }
    showCounts(String(form.querySelector("input[name=pattern_id]:checked")?.value || ""));

    form.querySelectorAll("input[name=pattern_id]").forEach((radio) => {
      radio.addEventListener("change", () => {
        showCounts(radio.value);
        const sheets = suggested[radio.value];
        if (!sheets || !sheets.length) return;
        form.querySelectorAll("input[name=sheets]").forEach((box) => { box.checked = sheets.includes(box.value); });
      });
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = form.querySelector("button[type=submit]");
      if (button) button.disabled = true;
      work("read", docs.length > 1 ? docs.length + "件の帳票を読み取っています…" : "帳票を読み取っています…");
      try {
        const res = await postForm(form.dataset.readUrl, new FormData(form));
        (res.errors || []).forEach((m) => toast(m, "err"));
        showReview(res);
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
      toast(docs.length ? "このファイルを外しました" : "この帳票の取り込みをやめました", "ok");
      if (docs.length) await openType();
      else resetPage();
    } catch (e) { toast(e.message, "err"); }
  });

  function resetPage() {
    docs = [];
    el.typeBody.textContent = "";
    el.finishBody.textContent = "";
    clearReview();
    if (el.fileNote) el.fileNote.textContent = "";
    if (el.docName) el.docName.textContent = "";
    ["type", "review", "finish"].forEach(sections.close);
    sections.close("file");
    sections.show("file");
  }

  // ---- 3 読み取り結果（帳票を縦に全部並べる。その場で直す・途中保存） ----
  const SOURCES = {
    auto: ["自動で読み取り", "blue"], manual: ["手で修正", "teal"], blank: ["空欄", "gray"],
  };
  const STATE_COLORS = { "確定済み": "green", "修正中": "violet", "未確定": "gray" };
  const blocks = new Map();   // 帳票ID → その帳票の塊（入力欄・版・保存の状態）
  let watcher = null;         // 元のシートを読み込むための、スクロールを見る役

  function clearReview() {
    blocks.forEach((b) => clearTimeout(b.timer));
    blocks.clear();
    stopWatching();
    el.reviewBody.textContent = "";
    sections.close("review");
    sections.close("finish");
  }

  function showReview(res) {
    clearReview();
    el.reviewBody.innerHTML = (res && res.html) || "";
    // 読み取れた帳票の状態をサーバの返事にそろえる（読み取れなかったファイルは並ばない）
    if (res && res.docs) {
      const states = new Map(res.docs.map((d) => [d.id, d.state]));
      docs.forEach((d) => { if (states.has(d.id)) d.state = states.get(d.id); });
    }
    sections.open("review");
    el.reviewBody.querySelectorAll(".review-doc").forEach(bindBlock);
    watchGrids();
    updateBar();
    setStatus("");
  }

  function bindBlock(root) {
    const b = {
      id: Number(root.dataset.doc),
      root,
      form: root.querySelector("[data-review-form]"),
      fieldBoxes: [],
      sheetGrids: [],
      sheetTabs: [],
      paneTabs: [],
      activeBox: null,
      dirty: false,
      timer: null,
      saving: null,
      gridLoaded: !root.querySelector("[data-grid-wait]"),
      gridLoading: null,
      pageToken: uuid(),
    };
    if (!b.form) return;
    blocks.set(b.id, b);
    b.fieldBoxes = Array.from(b.form.querySelectorAll("[data-field]"));
    b.sheetGrids = Array.from(root.querySelectorAll(".sheet-grid"));
    b.sheetTabs = Array.from(root.querySelectorAll("[data-sheet-tab]"));
    b.paneTabs = Array.from(root.querySelectorAll("[data-pane-tab]"));

    b.sheetTabs.forEach((tab) => tab.addEventListener("click", async () => {
      await loadGrid(b);
      showSheet(b, tab.dataset.sheetTab);
    }));

    b.paneTabs.forEach((tab) => tab.addEventListener("click", () => {
      b.paneTabs.forEach((t) => t.setAttribute("aria-selected", t === tab ? "true" : "false"));
      b.form.querySelectorAll("[data-pane]").forEach((p) => { p.hidden = p.dataset.pane !== tab.dataset.paneTab; });
    }));

    b.fieldBoxes.forEach((box) => {
      const input = box.querySelector("input.input, textarea.input");
      if (input) input.addEventListener("focus", () => { b.activeBox = box; highlight(b, box); });
      box.addEventListener("focusin", (event) => {
        if (!event.target.classList.contains("cell-input")) return;
        box._activeCell = event.target;
        b.activeBox = box;
        highlight(b, box);
      });
      const table = box.querySelector("[data-table-editor]");
      if (table) {
        table.addEventListener("input", () => syncTable(box));
        box.addEventListener("click", (event) => onTableClick(b, box, table, event));
      }
    });

    // セルをクリック → 選んでいる項目にその値を入れる（あとから読み込む表にも効く）
    root.addEventListener("click", (event) => {
      const td = event.target.closest("td[data-cell]");
      if (!td || !b.activeBox) return;
      const input = inputOf(b.activeBox);
      if (!input) return;
      input.value = td.textContent.trim();
      input.dispatchEvent(new Event("input", { bubbles: true }));
      root.querySelectorAll("td.hl-value").forEach((c) => c.classList.remove("hl-value"));
      td.classList.add("hl-value");
      input.focus();
      toast("セル " + td.dataset.cell + " の値を入れました", "info");
    });

    b.form.addEventListener("input", (event) => {
      if (!event.target.name || !event.target.name.startsWith("value-")) return;
      b.dirty = true;
      setSaveText(b, "未保存の変更があります");
      setStatus("未保存の変更があります");
      clearTimeout(b.timer);
      b.timer = setTimeout(() => saveNow(b), 700);
    });

    const nextBtn = root.querySelector("[data-next-issue]");
    if (nextBtn) nextBtn.addEventListener("click", () => gotoNextIssue(b));
  }

  const inputOf = (box) => box._activeCell || box.querySelector("input.input, textarea.input, textarea.cell-input");
  const valueOf = (box) => box.querySelector("[name^='value-']");
  const versionOf = (b) => b.root.dataset.version || "";

  function setSaveText(b, text) {
    const node = b.root.querySelector("[data-doc-save]");
    if (node) node.textContent = text || "";
  }

  // ---- 元のシートの表は、その帳票が画面に近づいてから読み込む ----
  // 位置を自分で見る（IntersectionObserver は、画面に出していない窓では動かないことがある）。
  const GRID_MARGIN = 600;   // 画面の上下 600px 手前から読み込む

  function checkGrids() {
    const waiting = Array.from(blocks.values()).filter((b) => !b.gridLoaded && !b.gridLoading);
    if (!waiting.length) { stopWatching(); return; }
    const height = window.innerHeight || document.documentElement.clientHeight || 0;
    waiting.forEach((b) => {
      const box = b.root.getBoundingClientRect();
      if (box.bottom > -GRID_MARGIN && box.top < height + GRID_MARGIN) loadGrid(b);
    });
  }

  function stopWatching() {
    if (!watcher) return;
    window.removeEventListener("scroll", watcher, true);
    window.removeEventListener("resize", watcher);
    watcher = null;
  }

  function watchGrids() {
    stopWatching();
    if (!blocks.size) return;
    let timer = null;      // スクロール中に何度も測らない
    watcher = () => {
      if (timer) return;
      timer = setTimeout(() => { timer = null; checkGrids(); }, 120);
    };
    window.addEventListener("scroll", watcher, true);   // 中の枠のスクロールも拾う
    window.addEventListener("resize", watcher);
    checkGrids();
  }

  function loadGrid(b) {
    if (b.gridLoaded) return Promise.resolve();
    if (b.gridLoading) return b.gridLoading;
    const slot = b.root.querySelector("[data-grid-slot]");
    b.gridLoading = (async () => {
      try {
        const res = await get(b.root.dataset.gridUrl);
        if (slot) slot.innerHTML = res.html || "";
        b.sheetGrids = Array.from(b.root.querySelectorAll(".sheet-grid"));
      } catch (e) {
        if (slot) {
          slot.textContent = "";
          const p = document.createElement("p");
          p.className = "warn";
          p.textContent = "元のシートを読み込めませんでした（" + e.message + "）";
          slot.appendChild(p);
        }
      } finally {
        b.gridLoaded = true;      // 失敗しても何度も取りに行かない
        b.gridLoading = null;
      }
    })();
    return b.gridLoading;
  }

  function showSheet(b, name) {
    let shown = null;
    b.sheetGrids.forEach((g) => {
      const on = g.dataset.sheet === name;
      g.hidden = !on;
      if (on) shown = g;
    });
    b.sheetTabs.forEach((t) => t.setAttribute("aria-selected", t.dataset.sheetTab === name ? "true" : "false"));
    return shown;
  }

  const topLeft = (coord) => (coord || "").split(":")[0];

  function highlight(b, box) {
    if (!b.gridLoaded) { loadGrid(b).then(() => highlight(b, box)); return; }
    b.root.querySelectorAll("td.hl-label, td.hl-value").forEach((td) => td.classList.remove("hl-label", "hl-value"));
    const sheet = box.dataset.sheet;
    let grid = sheet ? b.sheetGrids.find((g) => g.dataset.sheet === sheet) : null;
    if (!grid) return;
    if (grid.hidden) grid = showSheet(b, sheet);
    if (!grid) return;
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

  function onTableClick(b, box, table, event) {
    const remove = event.target.closest("[data-remove-table-row]");
    if (remove) {
      const tr = remove.closest("tr");
      if (box._activeCell && tr.contains(box._activeCell)) {
        box._activeCell = null;
        if (b.activeBox === box) b.activeBox = null;
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

  function values(b) {
    const out = {};
    b.fieldBoxes.forEach((box) => {
      const input = valueOf(box);
      if (input) out[box.dataset.field] = input.value;
    });
    return out;
  }

  function setVersion(b, v) {
    if (!v) return;
    b.root.dataset.version = v;
    const input = b.form.querySelector("[data-version-input]");
    if (input) input.value = v;
  }

  async function saveNow(b) {
    if (b.saving) await b.saving;
    if (!b.dirty) return;
    b.dirty = false;
    setSaveText(b, "保存中…");
    setStatus("保存中…");
    b.saving = (async () => {
      try {
        const body = { values: values(b), version: versionOf(b), page_token: b.pageToken };
        const res = await fetch(b.root.dataset.draftUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(body),
        });
        if (!res.ok && res.status !== 204) {
          let msg = "保存できませんでした（" + res.status + "）";
          try { const d = await res.json(); if (d.error) msg = d.error; } catch (e) { /* 本文なし */ }
          throw new Error(msg);
        }
        setVersion(b, res.headers.get("X-Doc-Version"));
        const summary = await postJson(b.root.dataset.previewUrl,
          { values: values(b), version: versionOf(b), page_token: b.pageToken });
        if (summary) applySummary(b, summary);
        setSaveText(b, "保存しました");
        setStatus("保存しました");
      } catch (err) {
        b.dirty = true;
        setSaveText(b, "保存できませんでした");
        setStatus("保存できませんでした");
        toast(err.message, "err");
      }
    })();
    await b.saving;
    b.saving = null;
  }

  function applySummary(b, s) {
    Object.entries(s.counts || {}).forEach(([key, n]) => {
      const node = b.root.querySelector("[data-count='" + key + "']");
      if (!node) return;
      node.textContent = n;
      const chip = node.closest(".chip");
      if (chip) chip.classList.toggle("is-zero", !n);
    });
    b.fieldBoxes.forEach((box) => {
      const st = (s.fields || {})[box.dataset.field];
      if (!st) return;
      box.dataset.issue = st.issue ? "1" : "0";
      box.classList.toggle("is-issue", !!st.issue);
      const badge = box.querySelector("[data-source-badge]");
      const src = SOURCES[st.source] || SOURCES.auto;
      if (badge) { badge.textContent = src[0]; badge.className = "badge badge-" + src[1]; }
      const warn = box.querySelector("[data-warning]");
      if (warn) {
        const text = st.warning || "";
        warn.textContent = text;
        warn.hidden = !text;
      }
    });
    const pre = b.root.querySelector("[data-md-preview]");
    if (pre && typeof s.markdown === "string") pre.textContent = s.markdown;
    const name = b.root.querySelector("[data-md-name]");
    if (name && s.file_name) name.textContent = s.file_name;
    if (s.state) {
      const doc = docOf(b.id);
      if (doc) doc.state = s.state;
      b.root.dataset.state = s.state;
    }
    setState(b, (s.counts || {}).issue || 0);
    updateBar();
  }

  /** 見出しの状態（要確認 n件 / 確定済み / 修正中）。 */
  function setState(b, issues) {
    const badge = b.root.querySelector("[data-doc-state]");
    if (!badge) return;
    const state = b.root.dataset.state;
    let text = issues ? "要確認 " + issues + "件" : "未確定";
    if (state === "confirmed") text = "確定済み";
    else if (state === "modified") text = "修正中";
    badge.textContent = text;
    badge.className = "badge badge-" + (STATE_COLORS[text] || "amber");
  }

  const issueCount = (b) => b.fieldBoxes.filter((box) => box.dataset.issue === "1").length;

  function gotoNextIssue(b) {
    const issues = b.fieldBoxes.filter((x) => x.dataset.issue === "1");
    if (!issues.length) { toast("この帳票に要確認の項目はありません", "ok"); return; }
    b.paneTabs.forEach((t) => { if (t.dataset.paneTab === "fields") t.click(); });
    const idx = b.activeBox ? issues.findIndex((x) => b.fieldBoxes.indexOf(x) > b.fieldBoxes.indexOf(b.activeBox)) : 0;
    const box = issues[idx >= 0 ? idx : 0];
    box.scrollIntoView({ block: "center" });
    const input = inputOf(box);
    if (input) input.focus();
  }

  // ---- まとめて置いたときの進み具合（③の上の1行） ----
  function updateBar() {
    const text = el.reviewBody.querySelector("[data-batch-text]");
    if (!text) return;
    const shown = docs.filter((d) => blocks.has(d.id));   // 読み取れなかったファイルは並んでいない
    text.textContent = shown.length + "件中" + shown.filter(isDone).length + "件を確定しました";
    const button = el.reviewBody.querySelector("[data-next-doc]");
    if (!button) return;
    const target = nextTarget();
    button.disabled = !target;
    button.textContent = !target ? "すべて確定しました"
      : (isDone(docOf(target.id) || {}) ? "要確認の残る帳票へ" : "次の未確定の帳票へ");
  }

  /** まだ手が要る帳票（未確定 → 要確認の残る帳票の順）。 */
  function nextTarget() {
    for (const d of docs) {
      const b = blocks.get(d.id);
      if (b && !isDone(d)) return b;
    }
    for (const d of docs) {
      const b = blocks.get(d.id);
      if (b && issueCount(b)) return b;
    }
    return null;
  }

  el.reviewBody.addEventListener("click", (event) => {
    if (!event.target.closest("[data-next-doc]")) return;
    const target = nextTarget();
    if (target) gotoBlock(target.id);
  });

  function gotoBlock(id) {
    const b = blocks.get(id);
    if (!b) return;
    sections.open("review", false);
    b.root.scrollIntoView({ behavior: "smooth", block: "start" });
    b.root.classList.add("is-jumped");
    setTimeout(() => b.root.classList.remove("is-jumped"), 1200);
  }

  // 画面を閉じるときは、未保存の変更を送っておく
  window.addEventListener("beforeunload", () => {
    blocks.forEach((b) => {
      if (!b.dirty) return;
      try {
        const blob = new Blob([JSON.stringify({ values: values(b), version: versionOf(b), page_token: b.pageToken })],
          { type: "application/json" });
        navigator.sendBeacon(b.root.dataset.draftUrl, blob);
      } catch (e) { /* 送れなくても次の入力で保存される */ }
    });
  });

  // ---- 4 確定してダウンロード ----
  async function refreshFinish() {
    if (!docs.length) { sections.close("finish"); return; }
    try {
      const url = page.dataset.finishUrl + "?ids=" + encodeURIComponent(ids());
      const res = await get(url);
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
      confirmBtn.disabled = true;
      work("confirm", "Markdown を作っています…");
      try {
        await confirmDoc(Number(confirmBtn.dataset.confirmDoc));
        await refreshFinish();
        sections.done("review", "確定しました");
        toast("確定しました", "ok");
        sections.show("finish");
      } catch (e) {
        toast(e.message, "err");
      } finally {
        work("confirm", "");
        confirmBtn.disabled = false;
      }
      return;
    }
    const goto = event.target.closest("[data-goto-doc]");
    if (goto) {
      event.preventDefault();
      gotoBlock(Number(goto.dataset.gotoDoc));
    }
  });

  /** 1件を確定する（入力中の値は先に保存してから）。 */
  async function confirmDoc(id) {
    const b = blocks.get(id);
    if (b) {
      clearTimeout(b.timer);
      if (b.dirty) await saveNow(b);
      if (b.saving) await b.saving;
    }
    await postJson("/forms/" + id + "/confirm", { version: b ? versionOf(b) : "" });
    const doc = docOf(id);
    if (doc) doc.state = "confirmed";
    if (b) {
      b.root.dataset.state = "confirmed";
      setState(b, issueCount(b));
    }
    updateBar();
  }

  // ダウンロード（zip も1件の .md も）: 確定していない帳票をこの場で全部確定してから受け取る
  page.addEventListener("click", async (event) => {
    const link = event.target.closest("a[data-confirm-all][data-confirmed='1']");
    if (!link) return;
    event.preventDefault();
    event.stopPropagation();
    // 修正中（確定したあとに直した）帳票も確定し直す。そうしないと直した値がファイルに入らない
    const pending = docs.filter((d) => blocks.has(d.id) && d.state !== "confirmed");
    work("confirm", pending.length ? pending.length + "件を確定しています…" : "ダウンロードの用意をしています…");
    try {
      for (const d of pending) await confirmDoc(d.id);
      await refreshFinish();
      window.location.assign(link.getAttribute("href"));
      // 読み取れなかった帳票はサーバーに残る（渡していないので消えない）ので、続けて読み取れるようにする
      const rest = docs.filter((d) => !isDone(d));
      setTimeout(async () => {
        toast("ダウンロードしました。渡した分のデータはサーバーから消えました", "ok");
        if (!rest.length) { resetPage(); return; }
        docs = rest;
        await openType();
      }, 1500);
    } catch (e) {
      toast(e.message, "err");
      await refreshFinish();
    } finally {
      work("confirm", "");
    }
  }, true);
})();
