// 一覧表を取り込む: 範囲の再判定、列の対応づけの保存、AI整形（分割プレビュー・試し実行・全件実行）、md の中身表示
"use strict";

(() => {
  const { toast, postJson, getJson } = window.App;

  const el = (tag, attrs = {}, ...children) => {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === "class") node.className = v;
      else if (k === "text") node.textContent = v;
      else node.setAttribute(k, v);
    }
    for (const child of children) {
      if (child == null) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  };

  // ---- 3 表の範囲と見出しを確認 -------------------------------------------------------------
  const layoutPage = document.querySelector("[data-layout-page]");
  if (layoutPage) {
    const headerInput = layoutPage.querySelector("[data-header-input]");
    const endInput = layoutPage.querySelector("[data-end-input]");
    const submit = layoutPage.querySelector("[data-layout-submit]");
    const KINDS = ["header", "data", "subtotal", "note", "continuation", "excluded", "title", "blank"];
    let timer = null;
    let seq = 0;

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

    const apply = (info) => {
      layoutPage.querySelectorAll("tr[data-row]").forEach((tr) => {
        const rc = info.rows[tr.dataset.row];
        KINDS.forEach((k) => tr.classList.remove(`row-${k}`));
        if (rc) {
          tr.classList.add(`row-${rc.kind}`);
          tr.title = rc.reason || "";
        } else {
          tr.classList.add("row-blank");
          tr.title = "";
        }
        tr.classList.toggle("is-picked", info.header_rows.includes(Number(tr.dataset.row)) || Number(tr.dataset.row) === info.data_end);
      });
      layoutPage.querySelector("[data-kind]").textContent = info.table_kind_label;
      // 見出し行が見つからないと data_end < data_start になる（存在しない行番号を出さない）
      layoutPage.querySelector("[data-range]").textContent =
        info.data_end < info.data_start ? "データの行が見つかりません" : `${info.data_start}〜${info.data_end}行目`;
      layoutPage.querySelector("[data-counts]").textContent =
        Object.entries(info.counts).map(([k, v]) => `${k} ${v}行`).join("／");
      layoutPage.querySelector("[data-headers]").textContent = info.headers.join("、");
      const warnings = layoutPage.querySelector("[data-warnings]");
      warnings.replaceChildren(...info.warnings.map((w) => el("li", { text: w })));
      const crosstab = info.table_kind === "crosstab";
      layoutPage.querySelector("[data-crosstab]").hidden = !crosstab;
      submit.disabled = crosstab;
      if (!headerInput.value.trim()) headerInput.value = info.header_rows.join(",");
      endInput.placeholder = `自動（${info.data_end}行目）`;
    };

    const detect = () => {
      clearTimeout(timer);
      timer = setTimeout(async () => {
        const mine = ++seq;
        layoutPage.setAttribute("aria-busy", "true");
        try {
          const info = await postJson(layoutPage.dataset.detectUrl, {
            header_rows: parseRows(headerInput.value), data_end: parseEnd(endInput.value),
          });
          if (mine === seq) apply(info);
        } catch (e) {
          toast(e.message, "err");
        } finally {
          layoutPage.removeAttribute("aria-busy");
        }
      }, 250);
    };

    headerInput.addEventListener("input", detect);
    endInput.addEventListener("input", detect);
    layoutPage.addEventListener("click", (event) => {
      const th = event.target.closest("th.rownum");
      if (!th) return;
      const n = Number(th.parentElement.dataset.row);
      const mode = layoutPage.querySelector("input[name=pick]:checked")?.value || "header";
      if (mode === "end") {
        endInput.value = String(n);
      } else if (event.shiftKey) {
        const current = parseRows(headerInput.value);
        const all = [...new Set([...current, n])].sort((a, b) => a - b);
        headerInput.value = [all[0], all[all.length - 1]].filter((v, i, a) => a.indexOf(v) === i).join(",");
      } else {
        headerInput.value = String(n);
      }
      detect();
    });
  }

  // ---- 4 列の対応づけ（取り込み設定の編集も同じ） ----------------------------------------------------
  const editor = document.querySelector("[data-columns-editor]");
  if (editor) {
    const errors = editor.querySelector("[data-editor-errors]");

    editor.addEventListener("change", (event) => {
      const tr = event.target.closest("tr[data-col]");
      if (!tr) return;
      const field = event.target.dataset.field;
      if (field === "use") tr.classList.toggle("is-unused", !event.target.checked);
      // AI整形の対象（役割＝追記ログ）は1列だけ。ほかの行のチェックを外し、追記ログの役割は長文に戻す
      const onlyThisLog = () => {
        editor.querySelectorAll("tr[data-col]").forEach((other) => {
          if (other === tr) return;
          const box = other.querySelector("input[data-field=ai]");
          if (box) box.checked = false;
          const role = other.querySelector("[data-field=role]");
          if (role && role.value === "log") role.value = "text";
        });
      };
      if (field === "ai" && event.target.checked) {
        // 役割を追記ログ・型を長文にする
        onlyThisLog();
        tr.querySelector("[data-field=role]").value = "log";
        tr.querySelector("[data-field=type]").value = "text";
        tr.querySelector("[data-field=use]").checked = true;
        tr.classList.remove("is-unused");
      }
      if (field === "role" && event.target.value === "log") {
        onlyThisLog();
        tr.querySelector("[data-field=ai]").checked = true;
        tr.querySelector("[data-field=type]").value = "text";
      }
    });

    const collect = () => {
      const settings = {};
      editor.querySelectorAll("[data-setting]").forEach((input) => {
        settings[input.dataset.setting] = input.type === "checkbox" ? input.checked : input.value;
      });
      const columns = [...editor.querySelectorAll("tr[data-col]")].map((tr) => {
        const row = { index: Number(tr.dataset.index), header: tr.dataset.header };
        tr.querySelectorAll("[data-field]").forEach((input) => {
          row[input.dataset.field] = input.type === "checkbox" ? input.checked : input.value;
        });
        return row;
      });
      return { ...settings, header_rows_count: Number(editor.dataset.headerRowsCount || 1), columns };
    };

    document.querySelectorAll("[data-editor-save]").forEach((button) => {
      button.addEventListener("click", async () => {
        button.disabled = true;
        errors.replaceChildren();
        try {
          const res = await postJson(editor.dataset.saveUrl, collect());
          window.location.href = res.redirect;
        } catch (e) {
          // 問題が複数あれば全部並べる（1件だけ直して保存し直す、を繰り返さずに済む）
          const messages = e.errors && e.errors.length ? e.errors : [e.message];
          errors.replaceChildren(...messages.map((m) => el("li", { text: m })));
          toast(messages.length > 1 ? `${messages.length}件の問題があります` : messages[0], "err");
          button.disabled = false;
        }
      });
    });
  }

  // ---- 5 AI整形 -------------------------------------------------------------------------------
  const aiPage = document.querySelector("[data-ai-page]");
  if (aiPage) {
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

    aiPage.querySelector("[data-split-run]")?.addEventListener("click", async () => {
      const box = aiPage.querySelector("[data-split-result]");
      const rowKey = aiPage.querySelector("[data-split-row]").value;
      box.textContent = "分割しています…";
      try {
        const data = await postJson(aiPage.dataset.splitUrl, { row_key: rowKey });
        // AI接続が未設定のときは「送る」だけだと今この操作で送ったように読めるので、送られないことを書く
        const routeText = data.route === "ai"
          ? (data.ai_ready === false ? "送る（AI接続の設定後。今は送っていません）" : "送る")
          : `送らない（${data.reason}）`;
        const parts = [
          el("p", { class: "hint", text: `区切り ${data.segments.length}件（オレンジは推定）。AIに送るか: ${routeText}` }),
          segmentList(data),
        ];
        if (data.notes.length) parts.push(el("ul", { class: "list-plain warn" }, ...data.notes.map((n) => el("li", { text: n }))));
        parts.push(el("h3", { text: "Markdown での時系列" }), el("pre", { class: "md-view", text: data.timeline.join("\n") }));
        if (data.sent_text) {
          parts.push(el("details", {}, el("summary", { text: "送信内容を表示" }), el("pre", { class: "md-view", text: data.sent_text })));
        }
        box.replaceChildren(...parts);
      } catch (e) {
        box.textContent = "";
        toast(e.message, "err");
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

    aiPage.querySelector("[data-trial-run]")?.addEventListener("click", async (event) => {
      const button = event.currentTarget;
      const external = aiPage.querySelector("[data-trial-external]");
      if (external && !external.checked) {
        toast("外部に送信されることを確認して、チェックを入れてください。", "err");
        return;
      }
      const keys = JSON.parse(aiPage.dataset.trialKeys || "[]");
      const status = aiPage.querySelector("[data-trial-status]");
      const result = aiPage.querySelector("[data-trial-result]");
      button.disabled = true;
      result.replaceChildren();
      trials.length = 0;
      for (let i = 0; i < keys.length; i += 1) {
        status.textContent = `${i + 1} / ${keys.length}行目を処理しています…`;
        try {
          const res = await postJson(aiPage.dataset.trialUrl, { row_key: keys[i], confirm_external: Boolean(external?.checked) });
          trials.push(...(res.stats || []));
          result.append(trialCard(res));
        } catch (e) {
          toast(e.message, "err");
          status.textContent = `止まりました: ${e.message}`;
          button.disabled = false;
          return;
        }
      }
      status.textContent = `${keys.length}行の試し実行が終わりました。`;
      button.disabled = false;
    });

    const runOptions = () => ({
      scope: aiPage.querySelector("[data-run-scope]")?.value || "pending",
      concurrency: Number(aiPage.querySelector("[data-run-concurrency]")?.value || 1),
    });

    aiPage.querySelector("[data-estimate-run]")?.addEventListener("click", async (event) => {
      const button = event.currentTarget;  // await のあとでは event.currentTarget が null になる
      const out = aiPage.querySelector("[data-estimate]");
      button.disabled = true;
      out.textContent = "見積もっています…";
      try {
        const r = await postJson(aiPage.dataset.estimateUrl, { ...runOptions(), trials });
        out.textContent = `AIに送る行 ${r.ai_rows}件（呼び出し ${r.calls}回、保存済み ${r.cached}件、ルールのみ ${r.rule_only_rows}件）`
          + `／入力 約${r.tokens_in}トークン・出力 約${r.tokens_out}トークン／${r.duration_text || `約${Math.ceil(r.minutes || 0)}分`}`
          + (r.basis === "trial" ? "（試し実行の実測から）" : "（目安）");
      } catch (e) {
        out.textContent = "";
        toast(e.message, "err");
      } finally {
        button.disabled = false;
      }
    });

    aiPage.querySelector("[data-ai-run]")?.addEventListener("click", async (event) => {
      const external = aiPage.querySelector("[data-run-external]");
      if (external && !external.checked) {
        toast("外部に送信されることを確認して、チェックを入れてください。", "err");
        return;
      }
      const button = event.currentTarget;  // await のあとでは event.currentTarget が null になる
      button.disabled = true;
      try {
        await postJson(aiPage.dataset.runUrl, { ...runOptions(), confirm_external: Boolean(external?.checked) });
        window.location.reload();
      } catch (e) {
        toast(e.message, "err");
        button.disabled = false;
      }
    });

    // 進み具合の内訳、終わったら再表示
    // 実行中⇔一時停止が切り替わったら再表示する（ボタンを［再開］／［一時停止］に合わせる。
    // レート制限が続いてジョブが自分で一時停止したときも）
    let lastStatus = aiPage.querySelector("[data-job-status]")?.dataset.jobStatus;
    aiPage.addEventListener("job:update", (event) => {
      const status = event.detail.status;
      const paused = (s) => s === "paused";
      if (lastStatus && status && paused(lastStatus) !== paused(status) && ["running", "paused"].includes(status)) {
        window.location.reload();
        return;
      }
      lastStatus = status || lastStatus;
      const p = event.detail.progress || {};
      const detail = aiPage.querySelector("[data-ai-detail]");
      if (detail) {
        const rest = !p.remaining_sec ? "" : p.remaining_sec < 60 ? "／残り1分未満" : `／残り約${Math.ceil(p.remaining_sec / 60)}分`;
        detail.textContent = `OK ${p.ok || 0}／要確認 ${p.flagged || 0}／エラー ${p.error || 0}／ルールのみ ${p.rule_only || 0}${rest}`;
      }
    });
    aiPage.addEventListener("job:finished", () => window.location.reload());
  }

  // ---- 6 内容とファイルの確認: md の中身 --------------------------------------------------------------
  const previewPage = document.querySelector("[data-preview-page]");
  if (previewPage) {
    const title = previewPage.querySelector("[data-file-title]");
    const text = previewPage.querySelector("[data-file-text]");
    previewPage.addEventListener("click", async (event) => {
      const button = event.target.closest("[data-file-open]");
      if (!button) return;
      const name = button.dataset.fileOpen;
      previewPage.querySelectorAll("tr[data-file]").forEach((tr) => tr.classList.toggle("is-current", tr.dataset.file === name));
      title.textContent = `${name}（読み込み中…）`;
      try {
        const data = await getJson(`${previewPage.dataset.fileUrl}?name=${encodeURIComponent(name)}`);
        title.textContent = name;
        text.textContent = data.text;
      } catch (e) {
        title.textContent = name;
        toast(e.message, "err");
      }
    });
    const first = previewPage.querySelector("[data-file-open]");
    if (first) first.click();
  }
})();
