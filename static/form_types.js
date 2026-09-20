// 帳票登録（/form-types）: 1枚の画面で 登録済みの一覧 → Excel を置く → セルをクリックして項目を作る → 使用開始。
// 画面は移動しない。どの操作も fetch でルートを呼び、返ってきた HTML の断片をその場に入れ替える。
(function () {
  "use strict";

  const page = document.getElementById("typesPage");
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

  function work(name, text) {
    const line = page.querySelector('[data-work="' + name + '"]');
    if (!line) return;
    const span = line.querySelector("[data-work-text]");
    if (span) span.textContent = text || "";
    line.hidden = !text;
  }

  const listBody = page.querySelector("[data-list-body]");
  const buildBody = page.querySelector("[data-build-body]");
  const buildName = page.querySelector('#step-build [data-step-summary]');
  let patternId = null;

  // ---- 返ってきた断片を画面に入れる ----
  function apply(res, opts) {
    if (res.list_html !== undefined && listBody) listBody.innerHTML = res.list_html;
    if (res.html !== undefined) {
      buildBody.innerHTML = res.html;
      const root = buildBody.querySelector("#cellBuilder");
      patternId = root ? Number(root.dataset.pattern) : patternId;
      const nameInput = buildBody.querySelector("[data-rename]");
      if (buildName) buildName.textContent = nameInput ? nameInput.value : "";
      bindBuilder();
      sections.done("new", nameInput ? nameInput.value : "");
      sections.open("build", !!(opts && opts.scroll));
    }
    if (res.message) toast(res.message, "ok");
    (res.errors || []).forEach((m) => toast(m, "err"));
  }

  // ---- 新しく登録する ----
  const newForm = document.getElementById("newTypeForm");
  if (newForm) {
    const file = newForm.querySelector("input[type=file]");
    const name = newForm.querySelector("[data-type-name]");
    if (file && name) {
      file.addEventListener("change", () => {
        if (!file.files || !file.files.length) return;
        if (!name.value.trim()) name.value = file.files[0].name.replace(/\.[^.]+$/, "");
        // 置いたらそのままシートを開く（ボタンを押さなくてよい）
        newForm.requestSubmit();
      });
    }
    newForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (!file || !file.files || !file.files.length) { toast("帳票のExcelを置いてください", "err"); return; }
      const button = newForm.querySelector("button[type=submit]");
      if (button) button.disabled = true;
      work("create", "Excel を読み込んでいます…");
      try {
        const res = await postForm(page.dataset.createUrl, new FormData(newForm));
        apply(res, { scroll: true });
        newForm.reset();
        file.dispatchEvent(new Event("change", { bubbles: true }));
      } catch (e) {
        toast(e.message, "err");
      } finally {
        work("create", "");
        if (button) button.disabled = false;
      }
    });
  }

  // ---- 一覧・登録中の欄のボタン（どちらも同じ操作） ----
  page.addEventListener("click", async (event) => {
    const open = event.target.closest("[data-open-type]");
    if (open) {
      event.preventDefault();
      await openType(Number(open.dataset.openType));
      return;
    }
    const sample = event.target.closest("[data-open-sample]");
    if (sample) {
      event.preventDefault();
      await openType(patternId, Number(sample.dataset.openSample));
      return;
    }
    const status = event.target.closest("[data-set-status]");
    if (status) {
      event.preventDefault();
      status.disabled = true;
      try {
        const id = Number(status.dataset.pattern);
        const res = await postJson("/form-types/" + id + "/status", { status: status.dataset.setStatus });
        patternId = id;
        apply(res);
      } catch (e) { toast(e.message, "err"); }
      status.disabled = false;
      return;
    }
    // 削除（確認ダイアログは app.js が先に出す。ここは「はい」のあとの本番）
    const del = event.target.closest("[data-delete-type][data-confirmed='1']");
    if (del) {
      try {
        const res = await postJson("/form-types/" + Number(del.dataset.deleteType) + "/delete", {});
        if (Number(del.dataset.deleteType) === patternId) {
          buildBody.textContent = "";
          patternId = null;
          if (buildName) buildName.textContent = "";
          sections.close("build");
          sections.close("new");
          sections.open("new");
        }
        apply(res);
      } catch (e) { toast(e.message, "err"); }
      return;
    }
    const field = event.target.closest("[data-delete-field]");
    if (field) {
      event.preventDefault();
      try { apply(await postForm(field.dataset.deleteField, new FormData())); }
      catch (e) { toast(e.message, "err"); }
      return;
    }
    const sampleDel = event.target.closest("[data-delete-sample][data-confirmed='1']");
    if (sampleDel) {
      try { apply(await postForm(sampleDel.dataset.deleteSample, new FormData())); }
      catch (e) { toast(e.message, "err"); }
    }
  });

  async function openType(id, sampleId) {
    if (!id) return;
    patternId = id;
    sections.open("build");
    buildBody.innerHTML = '<p class="muted">読み込んでいます…</p>';
    try {
      const url = "/form-types/" + id + "/panel" + (sampleId ? "?sample=" + sampleId : "");
      apply(await get(url), { scroll: true });
    } catch (e) {
      buildBody.textContent = "";
      toast(e.message, "err");
    }
  }

  // ---- 見本を足す ----
  page.addEventListener("submit", async (event) => {
    const form = event.target.closest("[data-add-samples]");
    if (!form) return;
    event.preventDefault();
    try { apply(await postForm(form.dataset.addSamples, new FormData(form))); }
    catch (e) { toast(e.message, "err"); }
  });

  // ---- 読み取る項目の見出しを直す（入力欄から離れる／Enter で保存、Escape で戻す） ----
  // 直るのは Markdown に書く名前だけ。探す見出しと読み取るセルはクリックしたときのまま
  page.addEventListener("keydown", (event) => {
    const input = event.target.closest("[data-field-label]");
    if (!input) return;
    if (event.key === "Enter" || event.key === "Escape") {
      event.preventDefault();
      if (event.key === "Escape") input.value = input.dataset.saved || "";
      input.blur();
    }
  });

  page.addEventListener("focusout", async (event) => {
    const input = event.target.closest("[data-field-label]");
    if (!input) return;
    const was = input.dataset.saved || "";
    const value = input.value.trim();
    if (!value || value === was) { input.value = was; return; }
    const root = input.closest("#cellBuilder");
    const data = new FormData();
    data.append("name", value);
    data.append("sample", (root && root.dataset.sample) || "");
    try { apply(await postForm(input.dataset.fieldLabel, data)); }
    catch (e) { input.value = was; toast(e.message, "err"); }
  });

  // ---- 名前を直す（入力欄から離れたら保存） ----
  page.addEventListener("focusout", async (event) => {
    const input = event.target.closest("[data-rename]");
    if (!input) return;
    const value = input.value.trim();
    if (!value || value === input.dataset.saved) return;
    try {
      const res = await postJson(input.dataset.rename, { name: value });
      input.dataset.saved = value;
      if (buildName) buildName.textContent = value;
      if (res.list_html !== undefined && listBody) listBody.innerHTML = res.list_html;
    } catch (e) { toast(e.message, "err"); }
  });

  // ---- セルをクリックして項目を作る ----
  let pending = null;   // {sheet, cell}

  function bindBuilder() {
    const root = buildBody.querySelector("#cellBuilder");
    if (!root) return;
    pending = null;
    const hint = root.querySelector("[data-click-hint]");
    const actions = root.querySelector("[data-click-actions]");
    const tabs = Array.from(root.querySelectorAll("[data-sheet-tab]"));
    const panes = Array.from(root.querySelectorAll(".sheet-grid"));

    const clearHighlight = () => root.querySelectorAll("td.hl-label, td.hl-value")
      .forEach((td) => td.classList.remove("hl-label", "hl-value"));

    function reset() {
      pending = null;
      clearHighlight();
      if (actions) actions.hidden = true;
      if (hint) hint.textContent = "見出しのセルをクリックしてください。";
    }

    async function addField(sheet, labelCell, valueCell) {
      const data = new FormData();
      data.append("sample", root.dataset.sample || "");
      data.append("sheet", sheet);
      data.append("label_cell", labelCell);
      data.append("value_cell", valueCell || "");
      if (hint) hint.textContent = "項目を作っています…";
      try { apply(await postForm(root.dataset.addUrl, data)); }
      catch (e) { toast(e.message, "err"); reset(); }
    }

    tabs.forEach((tab, i) => tab.addEventListener("click", () => {
      tabs.forEach((t, j) => t.setAttribute("aria-selected", String(i === j)));
      panes.forEach((p, j) => { p.hidden = i !== j; });
      reset();
    }));

    root.addEventListener("click", (event) => {
      if (event.target.closest("[data-click-cancel]")) { reset(); return; }
      if (event.target.closest("[data-click-empty]")) {
        if (pending) addField(pending.sheet, pending.cell, "");
        return;
      }
      const td = event.target.closest("td[data-cell]");
      if (!td) return;
      const pane = td.closest(".sheet-grid");
      if (!pane) return;
      const sheet = pane.dataset.sheet;
      const cell = td.dataset.cell;
      if (!pending) {
        if (td.dataset.tableHead) { addField(sheet, cell, ""); return; }  // 表は1回のクリックで項目にする
        if (!td.textContent.trim()) return;
        pending = { sheet: sheet, cell: cell };
        clearHighlight();
        td.classList.add("hl-label");
        if (actions) actions.hidden = false;
        if (hint) hint.textContent = "「" + td.textContent.trim().slice(0, 20) +
          "」の値のセルをクリックしてください（同じセルに値も入っているときは、もう一度このセルを）。";
        return;
      }
      if (pending.sheet !== sheet) { reset(); return; }
      td.classList.add("hl-value");
      addField(pending.sheet, pending.cell, cell);
    });

    // 項目の行を押すと、そのセルを光らせる
    root.querySelectorAll("tr[data-field]").forEach((tr) => {
      tr.addEventListener("click", (event) => {
        if (event.target.closest("button")) return;
        if (pending) return;
        clearHighlight();
        const pane = root.querySelector('.sheet-grid[data-sheet="' + (tr.dataset.sheet || "") + '"]');
        if (!pane) return;
        const i = panes.indexOf(pane);
        if (i >= 0 && tabs.length) tabs[i].click();
        [["labelCell", "hl-label"], ["valueCell", "hl-value"]].forEach((pair) => {
          const coord = (tr.dataset[pair[0]] || "").split(":")[0];
          const cell = coord && pane.querySelector('td[data-cell="' + coord + '"]');
          if (cell) cell.classList.add(pair[1]);
        });
      });
    });
  }
})();
