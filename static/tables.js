// 表の取り込み（/tables）: 1画面で全部やる（利用者の指示 2026-09-20）。
// 画面は移らない。段（.step）の中身はサーバが HTML の断片で返し、保存・実行は JSON でやりとりする。
// 段: file → source → layout → columns → ai → preview → done
"use strict";

(() => {
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
    const cancel = el("button", { type: "button", class: "btn danger-outline small", text: "中止" });
    cancel.addEventListener("click", async () => {
      cancel.disabled = true;
      try {
        const res = await rf(urls.cancel, { json: {} });
        toast(res.message || "処理を中止しました");
      } catch (e) {
        cancel.disabled = false;
      }
    });
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
    // 飛ばした段（例: 追記ログの列が無いときの AI整形）は開かずに取りに行く。
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
    const button = uploadForm.querySelector("[data-upload-run]");
    const input = uploadForm.querySelector("input[type=file]");
    if (!input?.files?.length) {
      toast("ファイルを選んでください。", "err");
      return;
    }
    if (button) button.disabled = true;
    sections.working("file", "ファイルを読み取っています…");
    try {
      // 新しいファイルを置いたら、前の取り込み（ダウンロードしていない分）はその場で捨てる
      if (importId) { await guard.now(); importId = null; urls = null; }
      const res = await rf(page.dataset.uploadUrl, { form: new FormData(uploadForm), quiet: true });
      urls = res.urls;
      importId = Number(res.import_id) || idFrom(urls && urls.panel);
      sections.done("file", res.file_name);
      await loadPanel("layout", { open: true, scroll: false });
      await loadPanel("source", { open: true });
    } catch (e) {
      toast(e.message, "err");
    } finally {
      if (button) button.disabled = false;
      sections.working("file", "");
    }
  });

  page.querySelector("[data-restart]")?.addEventListener("click", async () => {
    // 別のファイルにするときも、いまの取り込み（ダウンロードしていない分）はその場で捨てる
    await guard.now();
    importId = null;
    window.location.href = page.dataset.newUrl;
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
        runJob("done", res.job, () => loadPanel("done", { open: true }));
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
    root.querySelector("[data-layout-submit]").disabled = crosstab;
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
    // AI整形の対象（追記ログ）は1列だけ。ほかの行が選んでいたら「その他」に戻し、この行は「使う」にする
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
    button.disabled = true;
    errors.replaceChildren();
    try {
      const res = await rf(editor.dataset.saveUrl, { json: { name, columns }, quiet: true });
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

  const segmentList = (data) => {
    const list = el("ol", { class: "seg-list" });
    data.segments.forEach((s) => {
      const meta = el("div", { class: "seg-meta" },
        el("strong", { text: s.id }),
        el("span", { class: s.when_estimated ? "est" : "", text: `日付: ${s.when}` }),
        el("span", { class: s.author_estimated ? "est" : "", text: `記入者: ${s.author || "なし"}` }),
        s.marks.length ? el("span", { class: "muted", text: `印: ${s.marks.join("・")}` }) : null,
        s.identifiers.length ? el("span", { class: "muted", text: `識別子: ${s.identifiers.join("、")}` }) : null);
      list.append(el("li", {}, meta, el("div", { class: "pre", text: s.body })));
    });
    return list;
  };

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

    if (hit("[data-split-run]")) {
      const box = root.querySelector("[data-split-result]");
      const rowKey = root.querySelector("[data-split-row]").value;
      box.textContent = "分割しています…";
      try {
        const data = await rf(root.dataset.splitUrl, { json: { row_key: rowKey }, quiet: true });
        // AI接続が未設定のときは「送る」だけだと今この操作で送ったように読めるので、送られないことを書く
        const routeText = data.route === "ai"
          ? (data.ai_ready === false ? "送る（AI接続の設定後。今は送っていません）" : "送る")
          : `送らない（${data.reason}）`;
        const parts = [
          el("p", { class: "hint", text: `区切り ${data.segments.length}件（オレンジは推定）。AIに送るか: ${routeText}` }),
          segmentList(data),
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
        await rf(root.dataset.runUrl, { json: { ...runOptions(root), confirm_external: Boolean(external?.checked) }, quiet: true });
        await loadPanel("ai", { open: true, scroll: false });
      } catch (e) {
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
      return;
    }

    // AI接続（この段の中）
    const panel = hit("[data-ai-connection]");
    if (!panel) return;
    if (hit("[data-ai-test]")) {
      const button = hit("[data-ai-test]");
      const list = panel.querySelector("[data-ai-test-result]");
      button.disabled = true;
      button.textContent = "テスト中…";
      list.hidden = true;
      try {
        const result = await rf(panel.dataset.testUrl, { json: {}, quiet: true });
        list.replaceChildren(...result.steps.map((step) => el("li", {},
          el("div", { class: "item-main" },
            el("span", { class: "item-title", text: `${step.ok ? "OK" : "NG"}　${step.name}` }),
            el("span", { class: step.ok ? "muted" : "warn", text: step.detail })))));
        list.hidden = false;
        toast(result.ok ? "接続できました。" : "接続できませんでした。", result.ok ? "ok" : "err");
      } catch (e) {
        toast(e.message, "err");
      } finally {
        button.disabled = false;
        button.textContent = "接続をテストする";
      }
      return;
    }
    if (hit("[data-ai-models]")) {
      const button = hit("[data-ai-models]");
      button.disabled = true;
      try {
        const res = await rf(panel.dataset.modelsUrl, { json: {}, quiet: true });
        const list = panel.querySelector("[data-ai-model-list]");
        const known = new Set([...list.querySelectorAll("[data-ai-model]")].map((i) => i.value));
        res.models.filter((m) => !known.has(m)).forEach((m) => {
          list.append(el("label", { class: "check" },
            el("input", { type: "checkbox", "data-ai-model": "", value: m }),
            el("span", { class: "mono", text: m })));
        });
        toast(res.message || "取得しました");
      } catch (e) {
        toast(e.message, "err");
      } finally {
        button.disabled = false;
      }
      return;
    }
    if (hit("[data-ai-save]")) {
      const button = hit("[data-ai-save]");
      const data = { models: [...panel.querySelectorAll("[data-ai-model]:checked")].map((i) => i.value) };
      panel.querySelectorAll("[data-ai-field]").forEach((input) => {
        data[input.dataset.aiField] = input.type === "checkbox" ? input.checked : input.value;
      });
      button.disabled = true;
      try {
        const res = await rf(panel.dataset.saveUrl, { json: data, quiet: true });
        toast(res.message || "保存しました");
        await loadPanel("ai", { open: true, scroll: false });
      } catch (e) {
        toast(e.message, "err");
      } finally {
        button.disabled = false;
      }
    }
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
