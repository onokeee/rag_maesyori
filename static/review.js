// 帳票の画面用: 種類とシートの選択、読み取り結果の確認・修正、帳票の種類の編集（行の追加）
(function () {
  "use strict";

  const toast = (msg, kind) => (window.App && window.App.toast ? window.App.toast(msg, kind) : null);

  // ---- 種類とシートを確認: 種類を選ぶと、その種類で見つかったシートにチェックを付け直す ----
  const typeForm = document.querySelector("[data-type-select]");
  if (typeForm) {
    let suggested = {};
    try { suggested = JSON.parse(typeForm.dataset.suggested || "{}"); } catch (e) { suggested = {}; }
    typeForm.querySelectorAll("input[name=pattern_id]").forEach((radio) => {
      radio.addEventListener("change", () => {
        const sheets = suggested[radio.value];
        if (!sheets || !sheets.length) return;
        typeForm.querySelectorAll("input[name=sheets]").forEach((box) => { box.checked = sheets.includes(box.value); });
      });
    });
  }

  // ---- 帳票の種類の編集: 型が「明細表」のときだけ列見出しの欄を出す ----
  document.addEventListener("change", (event) => {
    const select = event.target.closest && event.target.closest("[data-type-select]");
    if (!select) return;
    const row = select.closest("tr");
    if (!row) return;
    const hide = select.value !== "table";
    row.querySelectorAll("[data-table-columns], [data-table-columns-label]").forEach((el) => { el.hidden = hide; });
  });

  // ---- 帳票の種類の編集: シート行・項目行の追加 ----
  document.querySelectorAll("[data-add-row]").forEach((button) => {
    button.addEventListener("click", () => {
      const kind = button.dataset.addRow;  // sheet | field
      const tpl = document.getElementById(kind + "-row-template");
      const body = document.querySelector(kind === "sheet" ? "[data-sheet-rows]" : "[data-field-rows]");
      if (!tpl || !body) return;
      const prefix = kind === "sheet" ? "sheets-" : "fields-";
      let max = -1;
      body.querySelectorAll("[name^='" + prefix + "']").forEach((el) => {
        const n = parseInt(el.name.split("-")[1], 10);
        if (!Number.isNaN(n)) max = Math.max(max, n);
      });
      const html = tpl.innerHTML.replace(/__INDEX__/g, String(max + 1));
      body.insertAdjacentHTML("beforeend", html);
      const first = body.lastElementChild && body.lastElementChild.querySelector("input:not([type=checkbox]), textarea");
      if (first) first.focus();
    });
  });

  // ---- 読み取り結果の確認・修正 ----
  const root = document.getElementById("review");
  const form = document.getElementById("reviewForm");
  if (!root || !form) return;

  const SOURCES = {
    auto: ["自動で読み取り", "blue"], ai: ["AIが入力（要確認）", "violet"],
    manual: ["手で修正", "teal"], blank: ["空欄", "gray"],
  };
  // 文書の状態タグ（components/_ui.html の STATE_LABELS と同じ語・色）
  const STATES = {
    unread: ["読み取り前", "gray"], reviewing: ["確認中", "amber"],
    modified: ["修正中", "violet"], confirmed: ["確定済み", "green"],
  };
  const fieldBoxes = Array.from(form.querySelectorAll("[data-field]"));
  // 明細表の項目は、最後に選んだセルの入力欄を「入力欄」とする（セルのクリックで値を入れる先）
  const inputOf = (box) => box._activeCell || box.querySelector("input.input, textarea.input, textarea.cell-input");
  const valueOf = (box) => box.querySelector("[name^='value-']");
  let activeBox = null;

  // シートのタブ
  const sheetGrids = Array.from(root.querySelectorAll(".sheet-grid"));
  const sheetTabs = Array.from(root.querySelectorAll("[data-sheet-tab]"));
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
  sheetTabs.forEach((tab) => tab.addEventListener("click", () => {
    const g = sheetGrids[parseInt(tab.dataset.sheetTab, 10)];
    if (g) showSheet(g.dataset.sheet);
  }));

  // 項目にフォーカス → 見出しセル・値セルをハイライト
  const topLeft = (coord) => (coord || "").split(":")[0];
  function clearHighlight() {
    root.querySelectorAll("td.hl-label, td.hl-value").forEach((td) => td.classList.remove("hl-label", "hl-value"));
  }
  function highlight(box) {
    clearHighlight();
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
  fieldBoxes.forEach((box) => {
    const input = box.querySelector("input.input, textarea.input");
    if (input) input.addEventListener("focus", () => { activeBox = box; highlight(box); });
    box.addEventListener("focusin", (event) => {
      if (!event.target.classList.contains("cell-input")) return;
      box._activeCell = event.target;
      activeBox = box;
      highlight(box);
    });
  });

  // 明細表: セルの編集・行の追加と削除 → hidden の JSON に戻す（入力イベントで途中保存される）
  function syncTable(box) {
    const hidden = box.querySelector("[data-table-value]");
    const table = box.querySelector("[data-table-editor]");
    if (!hidden || !table) return;
    const columns = Array.from(table.querySelectorAll("thead th[scope=col]")).map((th) => th.textContent.trim());
    const rows = Array.from(table.querySelectorAll("tbody tr")).map((tr) =>
      Array.from(tr.querySelectorAll(".cell-input")).map((el) => el.value));
    hidden.value = JSON.stringify({ columns, rows });
    hidden.dispatchEvent(new Event("input", { bubbles: true }));
  }
  fieldBoxes.forEach((box) => {
    const table = box.querySelector("[data-table-editor]");
    if (!table) return;
    table.addEventListener("input", () => syncTable(box));
    box.addEventListener("click", (event) => {
      const remove = event.target.closest("[data-remove-table-row]");
      if (remove) {
        const tr = remove.closest("tr");
        if (box._activeCell && tr.contains(box._activeCell)) {
          // 選んでいたセルの行を消したら、この表への入力先も外す（シートのクリックで先頭のセルを上書きしない）
          box._activeCell = null;
          if (activeBox === box) activeBox = null;
        }
        tr.remove();
        syncTable(box);
        return;
      }
      if (event.target.closest("[data-add-table-row]")) {
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
    });
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

  // 右ペインのタブ（項目 / Markdownプレビュー）
  const paneTabs = Array.from(root.querySelectorAll("[data-pane-tab]"));
  paneTabs.forEach((tab) => tab.addEventListener("click", () => {
    paneTabs.forEach((t) => t.setAttribute("aria-selected", t === tab ? "true" : "false"));
    form.querySelectorAll("[data-pane]").forEach((p) => { p.hidden = p.dataset.pane !== tab.dataset.paneTab; });
  }));

  // 入力値の収集
  function values() {
    const out = {};
    fieldBoxes.forEach((box) => {
      const input = valueOf(box);
      if (input) out[box.dataset.field] = input.value;
    });
    return out;
  }

  // 状態の反映（チップ・タグ・警告・必須欄・Markdown）
  function applySummary(s) {
    Object.entries(s.counts || {}).forEach(([key, n]) => {
      const el = document.querySelector("[data-count='" + key + "']");
      if (!el) return;
      el.textContent = n;
      const chip = el.closest(".chip");
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
    const missingBox = form.querySelector("[data-missing-box]");
    if (missingBox) {
      const list = s.missing_required || [];
      missingBox.hidden = !list.length;
      const span = missingBox.querySelector("[data-missing-list]");
      if (span) span.textContent = list.join("、");
    }
    // 確定済みの帳票を直すと状態が「修正中」に変わるので、見出しの状態タグも書き替える
    const stateBadge = document.querySelector("[data-doc-state] .badge");
    if (stateBadge && s.state && STATES[s.state]) {
      stateBadge.textContent = STATES[s.state][0];
      stateBadge.className = "badge badge-" + STATES[s.state][1] + " badge-" + s.state;
    }
    const pre = form.querySelector("[data-md-preview]");
    if (pre && typeof s.markdown === "string") pre.textContent = s.markdown;
    const name = form.querySelector("[data-md-name]");
    if (name && s.file_name) name.textContent = s.file_name;
    if (s.batch) applyBatch(s.batch);
  }

  // まとめ取り込みの欄: 途中保存で「修正中」になった帳票を、zip の確認文と修正中の一覧に出す
  function applyBatch(b) {
    document.querySelectorAll("[data-batch-zip]").forEach((a) => { a.dataset.confirm = b.zip_confirm; });
    // 保存先フォルダに保存するフォームも、確認文を今の状態に合わせる（保存でも渡した分のデータが消える）
    if (b.save_confirm) document.querySelectorAll("[data-batch-save]").forEach((f) => { f.dataset.confirm = b.save_confirm; });
    const box = document.querySelector("[data-batch-modified]");
    if (!box) return;
    const list = b.modified || [];
    box.hidden = !list.length;
    const count = box.querySelector("[data-batch-modified-count]");
    if (count) count.textContent = list.length;
    const span = box.querySelector("[data-batch-modified-list]");
    if (!span) return;
    span.textContent = "";
    list.forEach((d, i) => {
      if (i) span.append("、");
      const a = document.createElement("a");
      a.href = d.href;
      a.textContent = d.name;
      span.append(a);
    });
  }

  // 途中保存（204）→ プレビュー更新
  const statusEl = document.querySelector("[data-save-status]");
  const setStatus = (text) => { if (statusEl) statusEl.textContent = text; };
  let timer = null;
  let dirty = false;
  let saving = null;

  // 画面を開いたときの作業データの版。保存・確定で送り、別の画面で変わっていたら 409 で止まる
  const versionInput = form.querySelector("[data-version-input]");
  const currentVersion = () => root.dataset.version || "";
  // この画面の目印。途中保存の応答が届く前に画面を離れたとき、自分の保存とぶつからないように送る
  const pageToken = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now()) + Math.random().toString(16).slice(2);
  function setVersion(v) {
    if (!v) return;
    root.dataset.version = v;
    if (versionInput) versionInput.value = v;
  }

  async function postJson(url, body) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok && res.status !== 204) {
      let msg = "保存できませんでした（" + res.status + "）";
      try { const data = await res.json(); if (data.error) msg = data.error; } catch (e) { /* 本文なし */ }
      throw new Error(msg);
    }
    setVersion(res.headers.get("X-Doc-Version"));
    return res.status === 204 ? null : res.json();
  }

  async function saveNow() {
    if (saving) await saving;  // 前の保存が終わってから（同じ版を2回送って自分の保存とぶつからないように）
    if (!dirty) return;
    dirty = false;
    const body = { values: values(), version: currentVersion(), page_token: pageToken };
    setStatus("保存中…");
    saving = (async () => {
      try {
        await postJson(root.dataset.draftUrl, body);
        body.version = currentVersion();  // 保存で版が進んだので、プレビューには新しい版を送る
        const summary = await postJson(root.dataset.previewUrl, body);
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

  form.addEventListener("input", (event) => {
    if (!event.target.name || !event.target.name.startsWith("value-")) return;
    dirty = true;
    setStatus("未保存の変更があります");
    clearTimeout(timer);
    timer = setTimeout(saveNow, 700);
  });

  // 画面を離れる前に保存（確定・AIボタンの送信は値をフォームで送るので不要）
  let submitting = false;
  let waited = false;
  form.addEventListener("submit", (event) => {
    if (saving && !waited) {
      // 途中保存の応答を待ってから送る（新しい版を送らないと、自分の保存を「別の画面の変更」と取り違える）
      event.preventDefault();
      const submitter = event.submitter;
      waited = true;
      saving.finally(() => (submitter ? form.requestSubmit(submitter) : form.requestSubmit()));
      return;
    }
    waited = false;
    submitting = true;
    clearTimeout(timer);
  });
  window.addEventListener("beforeunload", () => {
    if (submitting || !dirty) return;
    try {
      const blob = new Blob([JSON.stringify({ values: values(), version: currentVersion(), page_token: pageToken })], { type: "application/json" });
      navigator.sendBeacon(root.dataset.draftUrl, blob);
    } catch (e) { /* 送れなくても次回の入力で保存される */ }
  });

  // 次の要確認へ
  const nextBtn = document.querySelector("[data-next-issue]");
  if (nextBtn) {
    nextBtn.addEventListener("click", () => {
      const issues = fieldBoxes.filter((b) => b.dataset.issue === "1");
      if (!issues.length) { toast("要確認の項目はありません", "ok"); return; }
      paneTabs.forEach((t) => { if (t.dataset.paneTab === "fields") t.click(); });
      const idx = activeBox ? issues.findIndex((b) => fieldBoxes.indexOf(b) > fieldBoxes.indexOf(activeBox)) : 0;
      const box = issues[idx >= 0 ? idx : 0];
      box.scrollIntoView({ block: "center" });
      const input = inputOf(box);
      if (input) input.focus();
    });
  }
})();
