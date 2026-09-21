// 帳票登録（/form-types）: 1枚の画面で 登録済みの一覧 → Excel を置く → セルをクリックして項目を作る → 使用開始。
// 画面は移動しない。どの操作も fetch でルートを呼び、返ってきた HTML の断片をその場に入れ替える。
//
// 置いた Excel はサーバーに残らない（利用者の指示 2026-09-21「見本のExcelは置かずに、設定だけ保持する」）。
// ブラウザ側で選んだファイル（bookFile）を持ち続け、セルのクリック・項目の削除・見出しの手直し・
// 使用開始のたびに一緒に送る。サーバーは受け取った Excel を読み取るだけで保存しない。
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
  let bookFile = null;     // いま画面で見ている帳票の Excel（ブラウザの中だけ。サーバーには残らない）

  // ---- サーバーへ送るとき、いまの Excel を一緒に持たせる ----
  // 送るものが無いときは、さっき読んでもらったブックの合図（sha256）だけを送る
  function withBook(data) {
    const form = data || new FormData();
    if (bookFile) {
      form.append("book", bookFile, bookFile.name);
      return form;
    }
    const root = buildBody.querySelector("#cellBuilder");
    const hash = root ? root.dataset.book : "";
    if (hash) form.append("book_hash", hash);
    return form;
  }

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
        if (!file.files || !file.files.length) { name.dataset.autofill = ""; return; }
        const stem = file.files[0].name.replace(/\.[^.]+$/, "");
        // 前のファイルから入れた名前が残っているとき（登録に失敗したあと）は、置き直した
        // ファイルの名前に入れ替える。手で書いた名前はそのままにする
        if (!name.value.trim() || name.value === name.dataset.autofill) {
          name.value = stem;
          name.dataset.autofill = stem;
        }
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
      // 置かれたファイルはブラウザ側で持ち続ける（サーバーには残らないので、次の操作でまた送る）
      bookFile = file.files[0];
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
    const status = event.target.closest("[data-set-status]");
    if (status) {
      event.preventDefault();
      status.disabled = true;
      try {
        const id = Number(status.dataset.pattern);
        const same = id === patternId;
        const data = new FormData();
        data.append("status", status.dataset.setStatus);
        // 開いている種類なら、いまの Excel も送ってシートを出したままにする
        const res = await postForm("/form-types/" + id + "/status", same ? withBook(data) : data);
        if (!same) bookFile = null;   // 別の種類に切り替わるので、前の種類の Excel は持ち越さない
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
          bookFile = null;
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
      // いま見ている Excel を送る（送らないとシートが消えてしまう）
      try { apply(await postForm(field.dataset.deleteField, withBook())); }
      catch (e) { toast(e.message, "err"); }
    }
  });

  async function openType(id) {
    if (!id) return;
    if (id !== patternId) bookFile = null;   // 別の種類を開いたら、前の種類の Excel は持ち越さない
    patternId = id;
    sections.open("build");
    buildBody.innerHTML = '<p class="muted">読み込んでいます…</p>';
    try {
      apply(await get("/form-types/" + id + "/panel"), { scroll: true });
    } catch (e) {
      buildBody.textContent = "";
      toast(e.message, "err");
    }
  }

  // ---- この帳票の Excel を置く（登録済みの種類を開き直したとき・別の書き方の帳票に替えるとき） ----
  page.addEventListener("submit", async (event) => {
    const form = event.target.closest("[data-book-form]");
    if (!form) return;
    event.preventDefault();
    const input = form.querySelector("input[type=file]");
    if (!input || !input.files || !input.files.length) { toast("この帳票のExcelを置いてください", "err"); return; }
    bookFile = input.files[0];
    const data = new FormData();
    data.append("book", bookFile, bookFile.name);
    try { apply(await postForm(form.dataset.bookForm, data), { scroll: true }); }
    catch (e) { bookFile = null; toast(e.message, "err"); }
  });

  // 選んだらそのまま読み込む（ボタンを押さなくてよい）
  page.addEventListener("change", (event) => {
    const input = event.target.closest("[data-book-form] input[type=file]");
    if (input && input.files && input.files.length && input.form) input.form.requestSubmit();
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
    const data = new FormData();
    data.append("name", value);
    try { apply(await postForm(input.dataset.fieldLabel, withBook(data))); }
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
      data.append("sheet", sheet);
      data.append("label_cell", labelCell);
      data.append("value_cell", valueCell || "");
      if (hint) hint.textContent = "項目を作っています…";
      try { apply(await postForm(root.dataset.addUrl, withBook(data))); }
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
