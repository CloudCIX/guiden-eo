(function (global, document) {
  "use strict";
  const G = global.Guiden,
    C = G.core;
  const $ = id => document.getElementById(id);
  const STORAGE = {
    key: "cloudcix-guiden.apiKey",
    remember: "cloudcix-guiden.rememberKey",
    prompt: "cloudcix-guiden.prompt",
    tone: "cloudcix-guiden.tone",
    length: "cloudcix-guiden.length",
    structure: "cloudcix-guiden.structure",
    citations: "cloudcix-guiden.keepCitations",
    track: "cloudcix-guiden.trackChanges",
    history: "cloudcix-guiden.history"
  };
  const state = {
    initialized: false,
    editor: "word",
    actionId: "writer.ask",
    context: null,
    suggestion: null,
    busy: false,
    controller: null,
    refreshTimer: null,
    saveTimer: null,
    managed: null,
    // A command from the toolbar or right-click menu waiting for a key (pendingRun) or for the current request to end.
    pendingRun: null,
    queuedRun: null,
    // Position in the instruction history while Up and Down step through it; -1 is the text being typed.
    historyIndex: -1,
    draftPrompt: "",
    // The follow-up to send again when Try again is pressed; empty to run the prompt again.
    retry: "",
    // Whether any of Guiden's reply has arrived for the request in progress.
    streamed: false
  };
  const defaultActions = { word: "writer.draft", cell: "cells.ask", slide: "slides.ask" };
  const ACTION_ORDER = {
    word: [
      "writer.draft",
      "writer.improve",
      "writer.rewrite",
      "writer.fix",
      "writer.summarise",
      "writer.ask",
      "writer.review",
      "knowledge.word",
      "writer.convert"
    ]
  };
  G.actionOrder = ACTION_ORDER;
  G.diagnostics = G.diagnostics || [];
  function diagnosticReference() {
    return "GDN-" + Date.now().toString(36).slice(-6).toUpperCase();
  }
  function logFailure(stage, error, details, reference) {
    const entry = {
      reference: reference || diagnosticReference(),
      time: new Date().toISOString(),
      stage,
      message: (error && error.message) || String(error),
      details: Object.assign({}, (error && error.guidenDetails) || {}, details || {})
    };
    G.diagnostics.push(entry);
    if (G.diagnostics.length > 50) G.diagnostics.shift();
    if (global.console && console.groupCollapsed) {
      console.groupCollapsed("[Guiden " + entry.reference + "] " + stage + " failed");
      console.error(error);
      console.log(entry);
      console.groupEnd();
    } else if (global.console) console.error("[Guiden]", entry, error);
    return entry.reference;
  }

  function textHtml(value) {
    return (
      "<p>" +
      C.escapeHtml(String(value || ""))
        .replace(/\n\n+/g, "</p><p>")
        .replace(/\n/g, "<br>") +
      "</p>"
    );
  }
  function sanitizeHtml(html, trustedUrls) {
    const allowedTags = new Set([
      "H1",
      "H2",
      "H3",
      "P",
      "STRONG",
      "EM",
      "A",
      "UL",
      "OL",
      "LI",
      "BLOCKQUOTE",
      "TABLE",
      "THEAD",
      "TBODY",
      "TR",
      "TH",
      "TD",
      "BR"
    ]);
    const template = document.createElement("template");
    template.innerHTML = String(html || "");
    const trustedLinks = Array.isArray(trustedUrls) ? new Set(trustedUrls.filter(C.isSafeLink)) : null;
    function clean(node) {
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType !== 1) return;
        if (!allowedTags.has(child.tagName)) {
          const fragment = document.createDocumentFragment();
          while (child.firstChild) fragment.appendChild(child.firstChild);
          child.replaceWith(fragment);
          clean(node);
          return;
        }
        const href = child.tagName === "A" ? child.getAttribute("href") || "" : "";
        Array.from(child.attributes).forEach(attribute => child.removeAttribute(attribute.name));
        if (child.tagName === "A" && C.isSafeLink(href) && (!trustedLinks || trustedLinks.has(href))) {
          child.setAttribute("href", href);
          child.setAttribute("rel", "noopener noreferrer");
          child.setAttribute("target", "_blank");
        }
        clean(child);
      });
    }
    clean(template.content);
    return template.innerHTML;
  }
  // A detail Guiden was told not to make up, such as [City] or [Media contact]. It starts with a letter, so citation
  // numbers like [12] are not placeholders.
  const PLACEHOLDER = /\[[A-Za-zÀ-ÿ][^\]\n]{1,59}\]/g;
  function placeholdersIn(text) {
    return Array.from(new Set(String(text || "").match(PLACEHOLDER) || []));
  }
  function markPlaceholders(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT),
      nodes = [];
    while (walker.nextNode()) if (walker.currentNode.nodeValue.match(PLACEHOLDER)) nodes.push(walker.currentNode);
    nodes.forEach(node => {
      if (node.parentNode.closest && node.parentNode.closest("mark, textarea, input")) return;
      const fragment = document.createDocumentFragment(),
        text = node.nodeValue;
      let last = 0;
      text.replace(PLACEHOLDER, (match, offset) => {
        if (offset > last) fragment.appendChild(document.createTextNode(text.slice(last, offset)));
        const mark = document.createElement("mark");
        mark.className = "placeholder";
        mark.textContent = match;
        fragment.appendChild(mark);
        last = offset + match.length;
      });
      if (last < text.length) fragment.appendChild(document.createTextNode(text.slice(last)));
      node.parentNode.replaceChild(fragment, node);
    });
  }
  function writerNativePayload(html) {
    const template = document.createElement("template");
    template.innerHTML = String(html || "");
    if (template.content.querySelector("table")) return { supported: false, blocks: [], linkCount: 0 };
    const blocks = [];
    let linkCount = 0;
    function collect(node, style, output) {
      if (node.nodeType === 3) {
        const text = node.textContent || "",
          bold = Boolean(style.bold),
          italic = Boolean(style.italic);
        let last = 0;
        text.replace(PLACEHOLDER, (match, offset) => {
          if (offset > last) output.push({ text: text.slice(last, offset), bold, italic });
          output.push({ text: match, bold, italic, highlight: true });
          last = offset + match.length;
        });
        if (last < text.length) output.push({ text: text.slice(last), bold, italic });
        return;
      }
      if (node.nodeType !== 1) return;
      const tag = node.tagName;
      if (tag === "BR") {
        output.push({ text: "\n", bold: Boolean(style.bold), italic: Boolean(style.italic) });
        return;
      }
      if (tag === "A") {
        const href = node.getAttribute("href") || "",
          text = node.textContent || href;
        if (C.isSafeLink(href) && text) {
          output.push({ text, href, bold: Boolean(style.bold), italic: Boolean(style.italic) });
          linkCount += 1;
        } else Array.from(node.childNodes).forEach(child => collect(child, style, output));
        return;
      }
      const next = { bold: style.bold || tag === "STRONG", italic: style.italic || tag === "EM" };
      Array.from(node.childNodes).forEach(child => collect(child, next, output));
    }
    // list is { kind: "bullet" | "numbered", id, marker }: one id per <ul>/<ol>, so each numbered list starts at 1, and
    // a marker for editors that cannot make real lists.
    function add(node, paragraphStyle, list) {
      const output = [];
      collect(node, {}, output);
      if (output.some(item => item.text))
        blocks.push({ segments: output, paragraphStyle: paragraphStyle || "", list: list || null });
    }
    let lists = 0;
    Array.from(template.content.childNodes).forEach(node => {
      if (node.nodeType === 3) {
        if (String(node.textContent || "").trim()) add(node, "");
        return;
      }
      const tag = node.tagName,
        styles = { H1: "Heading 1", H2: "Heading 2", H3: "Heading 3", BLOCKQUOTE: "Quote" };
      if (tag === "UL" || tag === "OL") {
        const kind = tag === "OL" ? "numbered" : "bullet",
          id = (lists += 1);
        Array.from(node.children)
          .filter(child => child.tagName === "LI")
          .forEach((item, index) =>
            add(item, "List Paragraph", { kind, id, marker: kind === "numbered" ? index + 1 + ". " : "• " })
          );
      } else add(node, styles[tag] || "");
    });
    return { supported: true, blocks, linkCount };
  }

  function applyManagedSettings() {
    const plugin = global.Asc && global.Asc.plugin,
      managed = C.managedSettings(plugin && plugin.info);
    state.managed = managed.key ? managed : null;
    G.api.configure(managed);
    $("apiKeyField").hidden = Boolean(state.managed);
    $("rememberKeyField").hidden = Boolean(state.managed);
    $("managedKeyNote").hidden = !state.managed;
    if (state.managed) showConnect(false);
  }
  function currentKey() {
    return state.managed ? state.managed.key : $("apiKey").value.trim();
  }
  function plugin() {
    return (global.Asc && global.Asc.plugin) || null;
  }
  // True when the panel runs as a window of the background plugin; editor events then come from the background.
  function inWindow() {
    const p = plugin();
    return Boolean(p && p.windowID);
  }

  // First run: a card that checks the key with CloudCIX once and remembers it, then carries on with what was asked.
  function showConnect(show, message) {
    $("connectCard").hidden = !show;
    document.body.classList.toggle("needs-key", show);
    if (!show) return;
    const pending = state.pendingRun && G.actions[state.pendingRun.actionId];
    $("connectLead").textContent = pending
      ? "Paste your CloudCIX API key and Guiden will " + pending.label.toLowerCase() + " straight after."
      : "Paste your CloudCIX API key to start. Guiden checks it once and remembers it on this browser.";
    connectMessage(message || "", message ? "error" : "");
    setTimeout(() => $("connectKey").focus(), 0);
  }
  function connectMessage(text, mode) {
    $("connectStatus").textContent = text;
    $("connectStatus").className = "status " + (mode || "");
  }
  function storeKey(key, remember) {
    try {
      localStorage.setItem(STORAGE.remember, remember ? "1" : "0");
      if (remember) localStorage.setItem(STORAGE.key, key);
      else localStorage.removeItem(STORAGE.key);
    } catch (_) {}
  }
  async function connect() {
    const key = $("connectKey").value.trim(),
      reference = diagnosticReference(),
      button = $("connectButton");
    if (!key) {
      connectMessage("Paste your CloudCIX API key first.", "error");
      return;
    }
    button.disabled = true;
    button.textContent = "Checking…";
    connectMessage("Checking the key with CloudCIX…", "loading");
    try {
      await G.api.test(key, new AbortController().signal);
      $("apiKey").value = key;
      $("rememberKey").checked = $("connectRemember").checked;
      storeKey(key, $("connectRemember").checked);
      $("connectKey").value = "";
      showConnect(false);
      setStatus("Connected. Guiden is ready.", "done");
      const next = state.pendingRun;
      state.pendingRun = null;
      if (next) handleRun(next);
      else $("prompt").focus();
    } catch (error) {
      logFailure("connect", error, {}, reference);
      connectMessage((error.message || "Connection failed.") + " Reference: " + reference, "error");
    } finally {
      button.disabled = false;
      button.textContent = "Connect";
    }
  }

  // A toolbar or right-click command: { actionId, run, instruction }. "run" commands start straight away; the others
  // open the action and wait for an instruction.
  function handleRun(command) {
    const action = command && G.actions[command.actionId];
    if (!action || action.editor !== state.editor) return;
    if (state.busy) {
      state.queuedRun = command;
      if (state.controller) state.controller.abort();
      return;
    }
    selectAction(command.actionId, { focus: !command.run });
    $("prompt").value = command.instruction || "";
    if (!command.run) {
      setStatus("Type an instruction, then press " + shortcutName() + ".", "");
      return;
    }
    if (!currentKey()) {
      state.pendingRun = command;
      showConnect(true);
      return;
    }
    generate();
  }
  function shortcutName() {
    return /Mac|iPhone|iPad/.test((global.navigator && global.navigator.platform) || "") ? "⌘+Enter" : "Ctrl+Enter";
  }

  // Without a theme from the editor (older builds), fall back to the operating system's preference.
  function applyTheme(theme) {
    let dark = C.isDarkTheme(theme);
    if (!theme || typeof theme !== "object") {
      try {
        dark = global.matchMedia("(prefers-color-scheme: dark)").matches;
      } catch (_) {
        dark = false;
      }
    }
    document.documentElement.classList.toggle("theme-dark", dark);
  }

  G.ui = { sanitizeHtml, writerNativePayload };
  function safeLinksFromHtml(html) {
    const template = document.createElement("template");
    template.innerHTML = String(html || "");
    return Array.from(template.content.querySelectorAll("a"))
      .map(link => link.getAttribute("href") || "")
      .filter(C.isSafeLink);
  }
  function safeRichHtml(html) {
    return sanitizeHtml(html);
  }
  function hasVisibleHtml(html) {
    const template = document.createElement("template");
    template.innerHTML = String(html || "");
    return Boolean(String(template.content.textContent || "").trim());
  }
  function setStatus(text, mode) {
    $("status").textContent = text;
    $("status").className = "status " + (mode || "");
  }
  function setBusy(value, label) {
    state.busy = value;
    document.body.classList.toggle("is-busy", value);
    $("generate").disabled = value;
    $("generate").textContent = value
      ? label || "Working…"
      : currentAction() && currentAction().mutates
        ? "Generate preview"
        : "Ask Guiden";
    $("stop").disabled = !value;
    // Stop only matters while Guiden is working, and Clear only while it is not.
    $("stop").hidden = !value;
    $("clear").hidden = value;
    if (value) $("retry").hidden = true;
    renderQuick();
    if (!value && state.queuedRun) {
      const next = state.queuedRun;
      state.queuedRun = null;
      setTimeout(() => handleRun(next), 0);
    }
    $("clear").disabled = value;
    $("apply").disabled = value;
    $("copy").disabled = value;
    $("testConnection").disabled = value;
    document.querySelectorAll(".action-card").forEach(b => (b.disabled = value));
  }
  function currentAction() {
    return G.actions[state.actionId];
  }
  function editorActions() {
    const order = ACTION_ORDER[state.editor] || [];
    return Object.values(G.actions)
      .filter(a => a.editor === state.editor)
      .sort((a, b) => {
        const ai = order.indexOf(a.id),
          bi = order.indexOf(b.id);
        return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi);
      });
  }
  function registerKnowledge(editor) {
    const id = "knowledge." + editor;
    if (G.actions[id]) return;
    G.registerAction({
      id,
      editor,
      label: "Ask Company Knowledge",
      description: "Search trusted company sources",
      mutates: false,
      canInsert: editor !== "cell",
      defaultInstruction: "Answer using relevant company knowledge.",
      placeholder: "Ask a question about your company knowledge.",
      supportedTargets: ["selection", "full_document", "current_slide", "selected_text"],
      buildRequest(snapshot, instruction) {
        return [
          {
            role: "system",
            content:
              'Use only the supplied Guiden retrieval context. Return JSON only: {"answer":"...","sources":[{"title":"...","url":"https://exact-url-from-context"}]}. Copy source URLs only when they appear exactly in the supplied retrieval context; never invent, repair, or complete a URL. Use an empty sources array when no source URL is present. If the retrieval context does not answer the question, say no relevant company source was found.'
          },
          { role: "user", content: instruction }
        ];
      },
      parseResponse(content) {
        const d = C.extractJson(content);
        if (typeof d.answer !== "string" || !d.answer.trim())
          throw new Error("Guiden returned an invalid company-knowledge answer.");
        const reportedSources = C.safeSources(d.sources);
        return {
          answer: d.answer,
          sources: reportedSources.slice(),
          reportedSources,
          foundRelevantSource: false,
          insertAllowed: false,
          retrievalUsed: false,
          provenanceUnavailable: false
        };
      },
      buildPreview(result) {
        return {
          kind: "knowledge",
          answer: result.answer,
          sources: result.sources,
          foundRelevantSource: result.foundRelevantSource,
          insertAllowed: result.insertAllowed,
          retrievalUsed: result.retrievalUsed,
          provenanceUnavailable: result.provenanceUnavailable
        };
      },
      revalidateTarget: async () => ({ ok: true }),
      apply: async () => {
        throw new Error("Company knowledge answers are read-only until you choose Insert answer.");
      }
    });
  }
  ["word", "cell", "slide"].forEach(registerKnowledge);

  // What Guiden will work on, in a few words: the panel is narrow, and the editor is obvious.
  function contextLabel(snapshot) {
    if (!snapshot) return "Selection unavailable";
    if (snapshot.editor === "word") {
      if (snapshot.target === "cursor") return "At the cursor";
      const words = snapshot.words + (snapshot.words === 1 ? " word" : " words");
      return (snapshot.target === "selection" ? "Selection" : "Whole document") + " · " + words;
    }
    if (snapshot.editor === "cell") {
      const cells = snapshot.rows * snapshot.cols;
      return snapshot.sheetName + " · " + snapshot.address + " · " + cells + (cells === 1 ? " cell" : " cells");
    }
    return (
      "Slide " +
      (snapshot.slideIndex + 1) +
      " · " +
      (snapshot.target === "selected_text" ? "selected text" : snapshot.shapes.length + " text areas")
    );
  }
  // The one-click actions for what is selected now. Rebuilt only when that changes, so a click is never lost.
  function hasSelection(snapshot) {
    return Boolean(snapshot && (snapshot.target === "selection" || snapshot.target === "selected_text"));
  }
  function renderQuick() {
    const bar = $("quickBar");
    if (!bar || !G.menus || !G.menus.quickActions) return;
    const selected = hasSelection(state.context),
      key = state.editor + ":" + selected;
    if (bar.dataset.key !== key) {
      bar.dataset.key = key;
      bar.innerHTML = "";
      G.menus.quickActions(state.editor, selected).forEach(item => {
        if (item.choices) {
          const select = document.createElement("select");
          select.className = "quick-choice";
          select.setAttribute("aria-label", item.text);
          select.innerHTML =
            '<option value="">' +
            C.escapeHtml(item.text) +
            " ▾</option>" +
            item.choices
              .map(c => '<option value="' + C.escapeHtml(c.id) + '">' + C.escapeHtml(c.text) + "</option>")
              .join("");
          bar.appendChild(select);
        } else {
          const button = document.createElement("button");
          button.type = "button";
          button.dataset.command = item.id;
          button.textContent = item.text;
          bar.appendChild(button);
        }
      });
    }
    bar.querySelectorAll("button, select").forEach(el => (el.disabled = state.busy));
  }
  function runQuick(id) {
    const command = G.menus && G.menus.commandFor(id);
    if (command && command.actionId) handleRun(command);
  }
  async function refreshContext(quiet) {
    try {
      const actionId = state.actionId || defaultActions[state.editor];
      const snapshot = await G.host.readSnapshot(state.editor, actionId);
      state.context = snapshot;
      $("contextText").textContent = contextLabel(snapshot);
      renderQuick();
      if (!quiet) setStatus("Selection refreshed.", "done");
    } catch (error) {
      $("contextText").textContent = "Selection unavailable";
      if (!quiet) setStatus(error.message || "Unable to read the current selection.", "error");
    }
  }
  function scheduleRefresh() {
    clearTimeout(state.refreshTimer);
    state.refreshTimer = setTimeout(() => refreshContext(true), 250);
  }
  // The editor's own events only reach the background frame, which forwards selection changes as guidenSelection.
  // A slow poll covers editors that send neither; it pauses while the panel is closed.
  function attachSelectionEvents() {
    const p = plugin();
    if (!p) return;
    if (!inWindow() && typeof p.attachEditorEvent === "function")
      ["onTargetPositionChanged", "onSelectionChanged", "onWorksheetSelectionChanged"].forEach(name => {
        try {
          p.attachEditorEvent(name, scheduleRefresh);
        } catch (_) {}
      });
    global.setInterval(() => {
      if (!state.busy && !state.suggestion && panelVisible()) refreshContext(true);
    }, 5000);
  }
  function panelVisible() {
    return document.visibilityState !== "hidden" && global.innerWidth > 0;
  }
  function renderActions() {
    const grid = $("actionGrid");
    grid.innerHTML = "";
    editorActions().forEach(action => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "action-card" + (action.id === state.actionId ? " selected" : "");
      button.dataset.action = action.id;
      button.innerHTML =
        "<span>" + C.escapeHtml(action.label) + "</span><small>" + C.escapeHtml(action.description) + "</small>";
      grid.appendChild(button);
    });
    document.body.dataset.editor = state.editor;
    updateActionUI();
  }
  function updateActionUI() {
    const action = currentAction();
    if (!action) return;
    document
      .querySelectorAll(".action-card")
      .forEach(b => b.classList.toggle("selected", b.dataset.action === action.id));
    $("selectedActionLabel").textContent = action.label;
    $("selectedActionDescription").textContent = action.description;
    $("prompt").placeholder = action.placeholder || "Add an instruction.";
    $("promptLabel").textContent = action.mutates ? "Tell Guiden what to change" : "What would you like to know?";
    $("promptHint").textContent = action.mutates
      ? "You'll see a preview before anything changes."
      : "Read-only: your file won't change.";
    $("generate").textContent = action.mutates ? "Generate preview" : "Ask Guiden";
    $("outputOptions").hidden =
      action.id.startsWith("knowledge.") ||
      action.id.endsWith(".ask") ||
      action.id === "cells.edit" ||
      action.id === "cells.format" ||
      action.id === "cells.formula";
    $("destination").closest("label").hidden = ![
      "cells.formula",
      "cells.add_column",
      "cells.classify",
      "cells.table"
    ].includes(action.id);
    $("tableTheme").closest("label").hidden = action.id !== "cells.table";
    if (action.id === "cells.formula") $("destination").value = "selection";
    if (["cells.add_column", "cells.classify"].includes(action.id)) $("destination").value = "right";
    if (action.id === "cells.table") $("destination").value = "start";
    $("destination").disabled = true;
  }
  function selectAction(id, settings) {
    if (!G.actions[id] || G.actions[id].editor !== state.editor) return;
    state.actionId = id;
    state.suggestion = null;
    $("resultGroup").hidden = true;
    $("insertKnowledge").hidden = true;
    renderActions();
    $("actionPicker").open = false;
    refreshContext(true);
    setStatus("", "");
    if (!settings || settings.focus !== false) $("prompt").focus();
  }
  function options() {
    return {
      tone: $("tone").value,
      length: $("length").value,
      structure: $("structure").value,
      destination: $("destination").value,
      tableTheme: $("tableTheme").value,
      slideCount: $("slideCount").value,
      keepCitations: $("keepCitations").checked
    };
  }
  function renderSources(sources, unverified) {
    const box = $("sources"),
      safe = C.safeSources(sources);
    box.innerHTML = "";
    box.hidden = !safe.length;
    if (!safe.length) return;
    const heading = document.createElement("div");
    heading.className = "sources-title";
    heading.textContent = unverified ? "Links reported by Guiden — verify before use" : "Sources";
    box.appendChild(heading);
    safe.forEach(source => {
      const a = document.createElement("a");
      a.href = source.url;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      a.textContent = source.title;
      box.appendChild(a);
    });
  }
  function tableHtml(table) {
    if (!Array.isArray(table) || !table.length) return "";
    return (
      '<div class="table-scroll"><table>' +
      table
        .slice(0, 25)
        .map(
          (row, i) =>
            "<tr>" +
            (Array.isArray(row) ? row : [row])
              .map(v => "<" + (i ? "td" : "th") + ">" + C.escapeHtml(v) + "</" + (i ? "td" : "th") + ">")
              .join("") +
            "</tr>"
        )
        .join("") +
      "</table>" +
      (table.length > 25 ? "<p>Preview shows the first 25 rows.</p>" : "") +
      "</div>"
    );
  }
  function beginStreaming(action, snapshot) {
    const preview = $("preview"),
      resultActions = document.querySelector(".result-actions");
    $("operationSummary").textContent = action.label + " · " + snapshot.targetLabel;
    $("previewBadge").textContent = "Generating";
    $("previewBadge").className = "badge loading";
    $("targetWarning").hidden = true;
    renderSources([]);
    preview.classList.add("streaming");
    preview.setAttribute("aria-busy", "true");
    state.streamed = false;
    preview.innerHTML =
      "<div class=stream-lead><span></span>Guiden is writing…</div><div class=stream-output>Waiting for the first words…</div>";
    if (resultActions) resultActions.hidden = true;
    $("refine").hidden = true;
    $("trackOption").hidden = true;
    $("resultGroup").hidden = false;
    try {
      $("resultGroup").scrollIntoView({ behavior: "smooth", block: "nearest" });
    } catch (_) {}
  }
  function updateStreaming(content) {
    const preview = $("preview"),
      output = preview.querySelector(".stream-output"),
      readable = C.streamingText(content);
    if (!output) return;
    if (content) state.streamed = true;
    output.textContent = readable || "Structuring the response…";
    preview.scrollTop = preview.scrollHeight;
  }
  function finishStreaming() {
    const preview = $("preview"),
      resultActions = document.querySelector(".result-actions");
    preview.classList.remove("streaming");
    preview.removeAttribute("aria-busy");
    if (resultActions) resultActions.hidden = false;
  }
  function failStreaming(label) {
    const preview = $("preview"),
      lead = preview.querySelector(".stream-lead");
    // Nothing arrived (a server error, or Stop before the first words): the status says what happened, and an
    // empty "Failed" preview would only get in the way.
    if (!state.streamed) {
      $("resultGroup").hidden = true;
      return;
    }
    preview.classList.remove("streaming");
    preview.removeAttribute("aria-busy");
    $("previewBadge").textContent = label;
    $("previewBadge").className = "badge failed";
    if (lead) lead.textContent = label + " — this incomplete response cannot be applied.";
  }
  // The words a change deletes and adds, marked up, and how much of the original it changes (0 to 1). null when the
  // texts are too long to compare.
  function changesHtml(before, after) {
    const paragraphs = text =>
      String(text || "")
        .replace(/\r\n?/g, "\n")
        .split(/\n+/)
        .map(line => line.trim())
        .filter(Boolean)
        .join("\n");
    const a = paragraphs(before),
      b = paragraphs(after),
      edits = a && b ? G.host.wordEdits(a, b) : null;
    if (!edits) return null;
    let html = "",
      offset = 0,
      removed = 0;
    edits.forEach(edit => {
      html += C.escapeHtml(a.slice(offset, edit[0]));
      if (edit[1] > edit[0]) html += "<del>" + C.escapeHtml(a.slice(edit[0], edit[1])) + "</del>";
      if (edit[2]) html += "<ins>" + C.escapeHtml(edit[2]) + "</ins>";
      removed += edit[1] - edit[0];
      offset = edit[1];
    });
    html += C.escapeHtml(a.slice(offset));
    return { html, ratio: removed / a.length };
  }
  // A Changes view (deleted and added words) next to the result. It opens first when the change is a light edit,
  // such as Fix or Improve; a rewrite or summary opens on the result.
  function withChanges(before, after, resultHtml) {
    const changes = changesHtml(before, after);
    if (!changes) return resultHtml;
    const first = changes.ratio < 0.6 ? "changes" : "result";
    return (
      '<div class="view-toggle" role="group" aria-label="Preview view">' +
      '<button type="button" data-view="changes" aria-pressed="' +
      (first === "changes") +
      '">Changes</button><button type="button" data-view="result" aria-pressed="' +
      (first === "result") +
      '">Result</button></div><article class="response-flow diff-view" data-view-panel="changes"' +
      (first === "changes" ? "" : " hidden") +
      ">" +
      changes.html +
      '</article><div data-view-panel="result"' +
      (first === "result" ? "" : " hidden") +
      ">" +
      resultHtml +
      "</div>"
    );
  }
  function showView(view) {
    $("preview")
      .querySelectorAll("[data-view]")
      .forEach(b => b.setAttribute("aria-pressed", String(b.dataset.view === view)));
    $("preview")
      .querySelectorAll("[data-view-panel]")
      .forEach(panel => (panel.hidden = panel.dataset.viewPanel !== view));
  }
  // Follow-ups on a result. Chips suit text results; the box also answers a needs-input question or asks a
  // follow-up about an answer.
  const TEXT_KINDS = ["writer-diff", "writer-new", "text-diff", "answer", "knowledge", "notes"];
  function showRefine(preview) {
    const chips = TEXT_KINDS.includes(preview.kind) && !preview.provenanceUnavailable;
    $("refineChips").hidden = !chips;
    $("refineText").value = "";
    $("refineText").placeholder =
      preview.kind === "needs-input"
        ? "Answer Guiden's question…"
        : ["answer", "knowledge", "analysis"].includes(preview.kind)
          ? "Ask a follow-up…"
          : "Ask Guiden to change this…";
    $("refine").hidden = false;
  }
  function refine(text) {
    const followUp = String(text || "").trim();
    if (!followUp || !state.suggestion || state.busy) return;
    generate(followUp);
  }

  // Earlier instructions, newest first, for the Up and Down keys in the prompt.
  function history() {
    try {
      const list = JSON.parse(localStorage.getItem(STORAGE.history) || "[]");
      return Array.isArray(list) ? list.filter(item => typeof item === "string") : [];
    } catch (_) {
      return [];
    }
  }
  function remember(instruction) {
    const text = String(instruction || "").trim();
    if (!text) return;
    try {
      const list = [text].concat(history().filter(item => item !== text)).slice(0, 20);
      localStorage.setItem(STORAGE.history, JSON.stringify(list));
    } catch (_) {}
    state.historyIndex = -1;
  }
  function recall(step) {
    const list = history();
    if (!list.length) return false;
    const next = Math.max(-1, Math.min(list.length - 1, state.historyIndex + step));
    if (next === state.historyIndex) return false;
    if (state.historyIndex === -1) state.draftPrompt = $("prompt").value;
    state.historyIndex = next;
    $("prompt").value = next === -1 ? state.draftPrompt || "" : list[next];
    return true;
  }

  function renderPreview(preview) {
    finishStreaming();
    const parts = [];
    $("targetWarning").hidden = !preview.warning;
    $("targetWarning").textContent = preview.warning || "";
    if (["answer", "knowledge", "analysis"].includes(preview.kind)) {
      parts.push('<article class="response-flow">' + textHtml(preview.answer) + "</article>");
      if (preview.table) parts.push(tableHtml(preview.table));
      if (preview.kind === "knowledge" && preview.provenanceUnavailable)
        parts.push(
          "<div class=empty-source>This Guiden server did not provide retrieval verification. The answer may still use your configured company corpora; review the content and any reported links before inserting.</div>"
        );
      else if (preview.kind === "knowledge" && !preview.retrievalUsed)
        parts.push(
          "<div class=empty-source>No relevant company source was found. Guiden did not use an unverified answer.</div>"
        );
      else if (preview.kind === "knowledge" && preview.retrievalUsed && !(preview.sources || []).length)
        parts.push(
          "<div class=empty-source>This answer used company knowledge, but that source has no clickable URL.</div>"
        );
    }
    if (preview.kind === "needs-input")
      parts.push(
        "<div class=needs-input><strong>Nothing will be changed.</strong> " + C.escapeHtml(preview.message) + "</div>"
      );
    if (preview.kind === "comments")
      parts.push(
        preview.comments.length
          ? '<div class="change-count">' +
              preview.comments.length +
              (preview.comments.length === 1 ? " comment" : " comments") +
              ' will be added</div><ol class="review-comments">' +
              preview.comments
                .map(c => "<li><q>" + C.escapeHtml(c.quote) + "</q><span>" + C.escapeHtml(c.comment) + "</span></li>")
                .join("") +
              "</ol>"
          : '<div class="empty-source">Guiden found nothing that needs a comment.</div>'
      );
    if (preview.kind === "writer-new")
      parts.push('<article class="response-flow">' + safeRichHtml(preview.afterHtml) + "</article>");
    if (preview.kind === "writer-diff")
      parts.push(
        withChanges(
          preview.before,
          C.htmlText(preview.afterHtml),
          '<details class="before-disclosure"><summary>View original text</summary><div>' +
            textHtml(preview.before) +
            '</div></details><div class="suggestion-label">Suggested version</div><article class="response-flow">' +
            safeRichHtml(preview.afterHtml) +
            "</article>"
        )
      );
    if (preview.kind === "text-diff")
      parts.push(
        withChanges(
          preview.before,
          preview.after,
          '<details class="before-disclosure"><summary>View original text</summary><div>' +
            textHtml(preview.before) +
            '</div></details><div class="suggestion-label">Suggested version</div><article class="response-flow">' +
            textHtml(preview.after) +
            "</article>"
        )
      );
    if (preview.kind === "cell-updates") {
      if (preview.narrative) parts.push(textHtml(preview.narrative));
      parts.push(
        '<div class="change-count">' +
          preview.updates.length +
          " cells will change</div>" +
          tableHtml(
            [["Cell", "Before", "After"]].concat(
              preview.updates.map(u => [
                C.rangeAddress(
                  {
                    startRow: preview.snapshot.start.startRow + u.row,
                    startCol: preview.snapshot.start.startCol + u.column
                  },
                  1,
                  1
                ),
                u.before,
                u.after
              ])
            )
          )
      );
    }
    if (preview.kind === "formula")
      parts.push(
        '<div class="destination">Destination: ' +
          C.escapeHtml(preview.address) +
          '</div><code class="formula">' +
          C.escapeHtml(preview.formula) +
          "</code>" +
          textHtml(preview.explanation)
      );
    if (preview.kind === "new-column")
      parts.push(
        textHtml(preview.narrative) +
          '<div class="destination">Destination: ' +
          C.escapeHtml(preview.address) +
          "</div>" +
          tableHtml([[preview.header]].concat(preview.values.map(v => [v])))
      );
    if (preview.kind === "formatting") {
      const cells = Array.from(
        new Set(
          preview.cells.map(item =>
            C.rangeAddress(
              {
                startRow: preview.snapshot.start.startRow + item.row,
                startCol: preview.snapshot.start.startCol + item.column
              },
              1,
              1
            )
          )
        )
      );
      parts.push(
        textHtml(preview.narrative) +
          '<div class="change-count">' +
          cells.length +
          ' matching cells</div><div class="swatches">' +
          preview.cells
            .slice(0, 8)
            .map(
              item =>
                '<span><i style="background:' +
                C.escapeHtml(item.style.fill || item.style.fontColor || "#e5e7eb") +
                '"></i>' +
                C.escapeHtml(
                  C.rangeAddress(
                    {
                      startRow: preview.snapshot.start.startRow + item.row,
                      startCol: preview.snapshot.start.startCol + item.column
                    },
                    1,
                    1
                  )
                ) +
                "</span>"
            )
            .join("") +
          "</div>"
      );
    }
    if (preview.kind === "table")
      parts.push(
        textHtml(preview.narrative) +
          '<div class="destination">Destination: ' +
          C.escapeHtml(preview.address) +
          " · " +
          preview.table.length +
          "×" +
          preview.table[0].length +
          "</div>" +
          tableHtml(preview.table)
      );
    if (preview.kind === "shape-updates")
      parts.push(
        textHtml(preview.narrative) +
          preview.updates
            .map(
              u =>
                '<div class="shape-change"><strong>Text area ' +
                (u.shapeIndex + 1) +
                '</strong><div class="diff before">' +
                textHtml(u.before) +
                '</div><div class="diff after">' +
                textHtml(u.after) +
                "</div></div>"
            )
            .join("")
      );
    if (preview.kind === "new-slides") {
      const names = { title: "Title slide", section: "Section", content: "Title and bullets", two: "Two columns" },
        list = items => "<ul>" + items.map(x => "<li>" + C.escapeHtml(x) + "</li>").join("") + "</ul>",
        column = part =>
          '<div class="slide-col">' +
          (part.heading ? "<strong>" + C.escapeHtml(part.heading) + "</strong>" : "") +
          list(part.bullets) +
          "</div>";
      parts.push(
        '<div class="diff-label">Uses this deck\'s own layouts, fonts and colours</div>' +
          preview.slides
            .map(
              (slide, i) =>
                '<div class="slide-card slide-' +
                slide.layout +
                '"><div class="slide-tag">' +
                (preview.slides.length > 1 ? i + 1 + " · " : "") +
                names[slide.layout] +
                "</div><h2>" +
                C.escapeHtml(slide.title) +
                "</h2>" +
                (slide.subtitle && (slide.layout === "title" || slide.layout === "section")
                  ? '<p class="slide-sub">' + C.escapeHtml(slide.subtitle) + "</p>"
                  : "") +
                (slide.layout === "two"
                  ? '<div class="slide-cols">' + column(slide.left) + column(slide.right) + "</div>"
                  : slide.bullets.length
                    ? list(slide.bullets)
                    : "") +
                "</div>"
            )
            .join("")
      );
    }
    if (preview.kind === "design") parts.push(designHtml(preview));
    if (preview.kind === "outline")
      parts.push(
        textHtml(preview.answer) +
          preview.slides
            .map(
              (s, i) =>
                '<div class="outline-slide"><strong>' +
                (i + 1) +
                ". " +
                C.escapeHtml(s.title || "Untitled") +
                "</strong><ul>" +
                (s.bullets || []).map(x => "<li>" + C.escapeHtml(x) + "</li>").join("") +
                "</ul></div>"
            )
            .join("")
      );
    if (preview.kind === "notes") parts.push('<div class="diff-label">Speaker notes</div>' + textHtml(preview.notes));
    $("preview").innerHTML = parts.join("");
    markPlaceholders($("preview"));
    renderSources(preview.sources || [], preview.provenanceUnavailable);
  }
  // The chosen design as a small slide in its own colours and fonts, with the other designs to switch to.
  function designHtml(preview) {
    const font = name =>
      /caladea|serif/i.test(name)
        ? "'" + name + "', Cambria, Georgia, serif"
        : "'" + name + "', 'Open Sans', Ubuntu, 'Segoe UI', sans-serif";
    const chosen = preview.designs.filter(d => d.id === preview.design)[0] || preview.designs[0],
      c = chosen.colors,
      swatches = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
        .map(k => '<span class="swatch" style="background:' + c[k] + '"></span>')
        .join("");
    return (
      '<div class="design-slide" style="background:' +
      c.lt1 +
      ";color:" +
      c.dk1 +
      ";border-top-color:" +
      c.accent1 +
      '"><div class="design-title" style="font-family:' +
      font(chosen.fonts.heading) +
      '">' +
      C.escapeHtml(chosen.name) +
      '</div><div class="design-body" style="font-family:' +
      font(chosen.fonts.body) +
      '">' +
      C.escapeHtml(
        chosen.fonts.heading === chosen.fonts.body
          ? chosen.fonts.body + " throughout"
          : chosen.fonts.heading + " headings, " + chosen.fonts.body + " text"
      ) +
      '</div><div class="design-swatches">' +
      swatches +
      '</div><span class="design-foot" style="background:' +
      c.accent2 +
      '"></span></div>' +
      (preview.reason ? '<p class="design-reason">' + C.escapeHtml(preview.reason) + "</p>" : "") +
      '<div class="diff-label">Or choose another design</div><div class="design-choices">' +
      preview.designs
        .map(
          d =>
            '<button type="button" class="design-choice' +
            (d.id === chosen.id ? " selected" : "") +
            '" data-design="' +
            d.id +
            '" aria-pressed="' +
            (d.id === chosen.id) +
            '"><span class="design-dot" style="background:' +
            d.colors.lt1 +
            ";border-color:" +
            d.colors.accent1 +
            '"><span style="background:' +
            d.colors.accent1 +
            '"></span></span>' +
            C.escapeHtml(d.name) +
            "</button>"
        )
        .join("") +
      '</div><div class="prompt-hint">Changes the theme colours, fonts and background of every slide. Undo in the editor puts the old design back.</div>'
    );
  }
  // The Apply button says what it will do: "Replace selection", "Insert at cursor", "Add 3 comments".
  function applyLabel(suggestion) {
    const action = G.actions[suggestion.actionId],
      snapshot = suggestion.snapshot || {},
      preview = suggestion.preview || {};
    const count = (n, one, many) => n + " " + (n === 1 ? one : many);
    if (action.id === "writer.review") return "Add " + count((preview.comments || []).length, "comment", "comments");
    if (action.id === "slides.notes") return "Add speaker notes";
    if (action.id === "slides.new")
      return (preview.slides || []).length > 1 ? "Create " + preview.slides.length + " slides" : "Create slide";
    if (action.id === "slides.design") return "Apply design";
    if (action.editor === "word") {
      if (snapshot.target === "cursor") return "Insert at cursor";
      if ($("trackChanges").checked) return "Suggest changes";
      return snapshot.target === "selection" ? "Replace selection" : "Replace document";
    }
    if (action.editor === "slide") return snapshot.target === "selected_text" ? "Replace text" : "Update slide";
    if (preview.kind === "cell-updates" || preview.kind === "formatting") {
      const n = (preview.updates || preview.cells || []).length;
      return n ? (preview.kind === "formatting" ? "Format " : "Update ") + count(n, "cell", "cells") : "Apply";
    }
    if (preview.kind === "formula") return "Insert formula";
    if (preview.kind === "table") return "Insert table";
    if (preview.kind === "new-column") return "Fill column";
    return "Apply";
  }
  // Plain words for the failures people actually hit, and whether trying again could help.
  function friendlyError(error) {
    const details = (error && error.guidenDetails) || {},
      status = details.status;
    if (details.timeout) return error.message;
    if (details.phase === "fetch")
      return "Guiden can't be reached. Check your internet connection, then try again. Nothing was changed.";
    if (status === 429) return "Guiden is busy right now. Wait a moment, then try again. Nothing was changed.";
    if (status >= 500)
      return "Guiden is unavailable right now (HTTP " + status + "). Try again in a moment. Nothing was changed.";
    if (details.phase === "stream")
      return "Guiden's answer was cut off before it finished. Try again. Nothing was changed.";
    return (error && error.message) || "Unable to generate a result.";
  }
  function appliedMessage(action, result) {
    if (action.editor === "cell")
      return (
        (action.id === "cells.format" ? "Formatted " : "Updated ") +
        result.changed +
        (result.changed === 1 ? " cell" : " cells") +
        " in " +
        result.target +
        "."
      );
    if (action.id === "slides.notes") return "Added speaker notes to " + result.target.replace(/ notes$/, "") + ".";
    if (action.id === "slides.new")
      return (
        (result.changed === 1 ? "Created 1 new slide." : "Created " + result.changed + " new slides.") +
        (result.plainDesign ? " This deck has the plain default design: try Design deck to give it a look." : "")
      );
    if (action.id === "slides.design") return "Applied the " + result.target + " to every slide.";
    if (action.id === "writer.review")
      return (
        "Added " +
        result.changed +
        (result.changed === 1 ? " comment" : " comments") +
        (result.missing ? " (" + result.missing + " could not be placed)" : "") +
        ". They are in the Comments pane."
      );
    const tracked = result.tracked ? " as tracked changes; accept or reject them under Collaboration" : "";
    if (action.editor === "word" && /cursor/i.test(String(result.target || "")))
      return tracked ? "Added the draft" + tracked + "." : "Added the draft at the cursor.";
    if (action.editor === "slide" && /selected/i.test(String(result.target || "")))
      return "Updated the selected text on the slide.";
    if (action.editor === "slide")
      return (
        "Updated " +
        result.changed +
        (result.changed === 1 ? " text area" : " text areas") +
        " on " +
        result.target +
        "."
      );
    const target = String(result.target || "the document");
    return (
      (tracked ? "Suggested changes to " : "Updated ") +
      target.charAt(0).toLowerCase() +
      target.slice(1) +
      tracked +
      "."
    );
  }
  // followUp continues the conversation behind the current result ("Make it shorter", or an answer to a question
  // Guiden asked): the same target, the earlier messages, Guiden's last reply and the follow-up.
  async function generate(followUp) {
    if (state.busy) return;
    const previous = typeof followUp === "string" && followUp.trim() && state.suggestion ? state.suggestion : null;
    const action = previous ? G.actions[previous.actionId] : currentAction(),
      key = currentKey(),
      instruction = previous ? followUp.trim() : $("prompt").value.trim() || action.defaultInstruction,
      reference = diagnosticReference();
    let snapshot = null,
      rawResponse = "";
    if (!key && previous) {
      showConnect(true);
      return;
    }
    if (!key) {
      state.pendingRun = { actionId: action.id, run: true, instruction: $("prompt").value.trim() };
      showConnect(true);
      return;
    }
    if (!instruction) {
      setStatus("Add an instruction or question.", "error");
      return;
    }
    state.retry = previous ? instruction : "";
    setBusy(true, action.mutates ? "Preparing preview…" : "Asking…");
    setStatus(action.mutates ? "Reading the current selection…" : "Reading context for your question…", "loading");
    state.controller = new AbortController();
    state.suggestion = null;
    $("resultGroup").hidden = true;
    try {
      if (previous) snapshot = previous.snapshot;
      else {
        if ($("prompt").value.trim()) remember($("prompt").value);
        snapshot = await G.host.readSnapshot(state.editor, action.id);
        snapshot.actionId = action.id;
        snapshot.options = options();
        if (action.supportedTargets && !action.supportedTargets.includes(snapshot.target))
          throw new Error("This action is not available for the current target. Update your selection and try again.");
        state.context = snapshot;
        $("contextText").textContent = contextLabel(snapshot);
      }
      setStatus(
        action.mutates
          ? "Guiden is generating a preview for " + snapshot.targetLabel + "…"
          : "Guiden is preparing an answer for " + snapshot.targetLabel + "…",
        "loading"
      );
      const fitted = previous
        ? { truncated: previous.truncated, notice: previous.notice }
        : C.fitContext(snapshot, action.mutates);
      const requestMessages = previous
        ? previous.messages.concat([
            { role: "assistant", content: previous.raw },
            { role: "user", content: instruction }
          ])
        : action.buildRequest(fitted.snapshot, instruction, snapshot.options);
      beginStreaming(action, snapshot);
      rawResponse = await G.api.complete(key, requestMessages, state.controller.signal, updateStreaming);
      let result;
      try {
        result = action.parseResponse(rawResponse, snapshot);
      } catch (schemaError) {
        setStatus("Guiden returned the content in the wrong format. Repairing the preview…", "loading");
        const lead = $("preview").querySelector(".stream-lead");
        if (lead) lead.textContent = "Guiden is repairing the response format…";
        rawResponse = await G.api.complete(
          key,
          C.schemaRetryMessages(requestMessages),
          state.controller.signal,
          updateStreaming
        );
        result = action.parseResponse(rawResponse, snapshot);
      }
      const metadata = G.api.lastMetadata || { available: false, retrievalUsed: false, sources: [] };
      if (Object.prototype.hasOwnProperty.call(result, "sources"))
        result.sources = metadata.available
          ? metadata.sources.slice()
          : C.safeSources(result.reportedSources || result.sources);
      if (typeof result.replacementHtml === "string") {
        const proposedLinks = safeLinksFromHtml(result.replacementHtml),
          allowedLinks = safeLinksFromHtml(snapshot.sourceHtml).concat(
            metadata.available ? metadata.sources.map(source => source.url) : proposedLinks
          );
        result.unverifiedLinks = !metadata.available && proposedLinks.length > 0;
        result.replacementHtml = sanitizeHtml(result.replacementHtml, allowedLinks);
        if (!hasVisibleHtml(result.replacementHtml))
          throw new Error("Guiden returned empty or unsupported document content. Nothing can be applied.");
      }
      if (
        typeof result.replacementHtml === "string" &&
        action.editor === "word" &&
        action.mutates &&
        $("keepCitations").checked
      ) {
        const citationSources = metadata.available ? metadata.sources : C.safeSources(result.sources),
          existingUrls = new Set(safeLinksFromHtml(result.replacementHtml)),
          missingSources = citationSources.filter(source => !existingUrls.has(source.url));
        if (missingSources.length)
          result.replacementHtml +=
            "<h3>Sources</h3><ul>" +
            missingSources
              .map(
                source => '<li><a href="' + C.escapeHtml(source.url) + '">' + C.escapeHtml(source.title) + "</a></li>"
              )
              .join("") +
            "</ul>";
      }
      if (action.id.indexOf("knowledge.") === 0) {
        if (!metadata.available) {
          result.provenanceUnavailable = true;
          result.retrievalUsed = null;
          result.foundRelevantSource = false;
          result.insertAllowed = true;
        } else {
          result.provenanceUnavailable = false;
          result.retrievalUsed = metadata.retrievalUsed;
          result.foundRelevantSource = metadata.retrievalUsed;
          result.insertAllowed = metadata.retrievalUsed;
          if (!metadata.retrievalUsed) {
            result.answer =
              "No relevant company source was found for this question. Guiden has withheld an unverified answer.";
            result.sources = [];
          }
        }
      }
      const preview = action.buildPreview(result, snapshot),
        applicable = action.mutates && !result.needsInput;
      state.suggestion = {
        actionId: action.id,
        snapshot,
        result,
        preview,
        raw: rawResponse,
        messages: requestMessages,
        truncated: fitted.truncated,
        notice: fitted.notice
      };
      if (fitted.truncated) preview.warning = (preview.warning ? preview.warning + " " : "") + fitted.notice;
      if (result.unverifiedLinks)
        preview.warning =
          (preview.warning ? preview.warning + " " : "") +
          "Check where the links go before applying: Guiden did not say which sources they came from.";
      const placeholders =
        typeof result.replacementHtml === "string" ? placeholdersIn(C.htmlText(result.replacementHtml)) : [];
      state.suggestion.placeholders = placeholders;
      if (placeholders.length)
        preview.warning =
          (preview.warning ? preview.warning + " " : "") +
          (placeholders.length === 1 ? "1 placeholder" : placeholders.length + " placeholders") +
          " to fill in, such as " +
          placeholders[0] +
          (state.editor === "word" ? ". They stay highlighted in the document." : ".");
      $("operationSummary").textContent = action.label + " · " + snapshot.targetLabel;
      $("previewBadge").textContent = result.needsInput
        ? "Needs information"
        : preview.provenanceUnavailable
          ? "Compatibility"
          : action.mutates
            ? "Will change file"
            : "Read-only";
      $("previewBadge").className =
        "badge " + (result.needsInput ? "attention" : action.mutates ? "mutation" : "readonly");
      renderPreview(preview);
      showRefine(preview);
      $("apply").hidden = !applicable;
      $("trackOption").hidden = !(applicable && action.editor === "word" && action.id !== "writer.review");
      if (preview.kind === "comments" && !preview.comments.length) $("apply").hidden = true;
      $("apply").textContent = applyLabel(state.suggestion);
      $("copy").hidden = Boolean(result.needsInput);
      $("insertKnowledge").hidden = !(
        action.canInsert &&
        preview.insertAllowed &&
        (state.editor !== "slide" || snapshot.target === "selected_text")
      );
      $("insertKnowledge").textContent = state.editor === "slide" ? "Insert on slide" : "Insert in document";
      $("resultGroup").hidden = false;
      setStatus(
        result.needsInput
          ? "Guiden needs more information. Nothing can be applied yet."
          : action.mutates
            ? "Preview ready. Nothing changes until you apply it."
            : "Answer ready. Your file has not changed.",
        result.needsInput ? "attention" : "done"
      );
      if (!previous) localStorage.setItem(STORAGE.prompt, $("prompt").value);
    } catch (error) {
      if (error.name === "AbortError") {
        if (snapshot) failStreaming("Stopped");
        setStatus("Stopped. The incomplete response was not applied.", "error");
      } else {
        if (snapshot) failStreaming("Failed");
        logFailure(
          "generate",
          error,
          {
            actionId: action.id,
            editor: state.editor,
            target: (snapshot && snapshot.targetLabel) || null,
            responseLength: rawResponse.length,
            rawResponse: rawResponse.slice(0, 12000)
          },
          reference
        );
        setStatus(friendlyError(error) + " Reference: " + reference, "error");
        const status = error.guidenDetails && error.guidenDetails.status;
        $("retry").hidden = status === 401 || status === 403;
        if ((status === 401 || status === 403) && !state.managed) {
          state.pendingRun = { actionId: action.id, run: true, instruction: $("prompt").value.trim() };
          showConnect(true, "CloudCIX did not accept the saved API key. Paste a new one to carry on.");
        }
      }
      // A follow-up that fails leaves the result it started from.
      if (previous && !state.suggestion) {
        state.suggestion = previous;
        renderPreview(previous.preview);
        showRefine(previous.preview);
        $("refineText").value = instruction;
        $("resultGroup").hidden = false;
        const resultActions = document.querySelector(".result-actions");
        if (resultActions) resultActions.hidden = false;
      }
    } finally {
      state.controller = null;
      setBusy(false);
    }
  }
  async function applySuggestion() {
    const suggestion = state.suggestion;
    if (!suggestion || state.busy) return;
    const action = G.actions[suggestion.actionId];
    if (!action || !action.mutates || suggestion.result.needsInput) return;
    const reference = diagnosticReference();
    setBusy(true, "Applying…");
    setStatus("Confirming the target has not changed…", "loading");
    try {
      // Re-read the target rather than trusting selection events: the editor also sends them when nothing that
      // matters changed, such as when the panel opens.
      const valid = await action.revalidateTarget(suggestion.snapshot);
      if (!valid.ok) throw new Error("The target changed. Generate again for the current selection.");
      const result = await action.apply(suggestion.result, suggestion.snapshot, {
        track: action.editor === "word" && $("trackChanges").checked
      });
      C.validateApplyResult(result);
      const holders = (suggestion.placeholders || []).length;
      setStatus(
        appliedMessage(action, result) +
          (holders && action.editor === "word"
            ? holders === 1
              ? " Fill in the highlighted placeholder."
              : " Fill in the " + holders + " highlighted placeholders."
            : ""),
        "done"
      );
      $("resultGroup").hidden = true;
      state.suggestion = null;
      const canUndo = await G.host.canUndo();
      $("undo").hidden = !canUndo;
      await refreshContext(true);
    } catch (error) {
      logFailure(
        "apply",
        error,
        { actionId: action.id, editor: state.editor, target: suggestion.snapshot.targetLabel },
        reference
      );
      setStatus((error.message || "Unable to apply the changes.") + " Reference: " + reference, "error");
    } finally {
      setBusy(false);
    }
  }
  async function insertKnowledge() {
    const s = state.suggestion;
    if (!s || !s.preview || !s.preview.insertAllowed) return;
    const reference = diagnosticReference();
    setBusy(true, "Inserting…");
    try {
      if (state.editor === "slide") {
        const valid = await G.host.revalidate(s.snapshot);
        if (!valid.ok) throw new Error("The target changed. Generate again for the current selection.");
      }
      let result;
      const html = textHtml(s.result.answer) + ($("keepCitations").checked ? C.sourcesHtml(s.result.sources) : "");
      if (state.editor === "word") result = await G.host.insertWriter(sanitizeHtml(html));
      else result = await G.host.applySelectedSlideText(s.result.answer, s.snapshot);
      C.validateApplyResult(result);
      setStatus("Inserted the company-knowledge answer into " + result.target + ".", "done");
      $("resultGroup").hidden = true;
      state.suggestion = null;
      $("undo").hidden = !(await G.host.canUndo());
    } catch (error) {
      logFailure(
        "insert-knowledge",
        error,
        { actionId: s.actionId, editor: state.editor, target: s.snapshot && s.snapshot.targetLabel },
        reference
      );
      setStatus((error.message || "Unable to insert the answer.") + " Reference: " + reference, "error");
    } finally {
      setBusy(false);
    }
  }
  function discard() {
    state.suggestion = null;
    $("resultGroup").hidden = true;
    $("insertKnowledge").hidden = true;
    setStatus("Preview discarded. Nothing was changed.", "");
  }
  function saveSettings() {
    try {
      const remember = $("rememberKey").checked;
      localStorage.setItem(STORAGE.remember, remember ? "1" : "0");
      ["tone", "length", "structure"].forEach(id => localStorage.setItem(STORAGE[id], $(id).value));
      localStorage.setItem(STORAGE.citations, $("keepCitations").checked ? "1" : "0");
      if (remember && !state.managed) localStorage.setItem(STORAGE.key, $("apiKey").value);
      else localStorage.removeItem(STORAGE.key);
      feedback("Saved ✓", "done");
      setStatus(
        state.managed
          ? "Settings saved."
          : remember
            ? "Settings and API key saved."
            : "Settings saved. API key is not stored.",
        "done"
      );
    } catch (_) {
      feedback("Save failed", "error");
      setStatus("Settings could not be saved in this browser.", "error");
    }
  }
  function feedback(label, mode) {
    clearTimeout(state.saveTimer);
    const b = $("saveSettings"),
      old = "Save settings";
    b.textContent = label;
    b.className = mode;
    state.saveTimer = setTimeout(() => {
      b.textContent = old;
      b.className = "";
    }, 2200);
  }
  function loadSettings() {
    try {
      const remember = localStorage.getItem(STORAGE.remember) === "1";
      $("rememberKey").checked = remember;
      $("apiKey").value = remember ? localStorage.getItem(STORAGE.key) || "" : "";
      $("prompt").value = localStorage.getItem(STORAGE.prompt) || "";
      ["tone", "length", "structure"].forEach(id => {
        const v = localStorage.getItem(STORAGE[id]);
        if (v) $(id).value = v;
      });
      if (localStorage.getItem(STORAGE.citations) !== null)
        $("keepCitations").checked = localStorage.getItem(STORAGE.citations) === "1";
      $("trackChanges").checked = localStorage.getItem(STORAGE.track) === "1";
    } catch (_) {}
  }
  async function testConnection() {
    const key = currentKey(),
      reference = diagnosticReference();
    if (!key) {
      setStatus("Enter an API key before testing.", "error");
      return;
    }
    setBusy(true, "Testing…");
    const controller = new AbortController();
    try {
      await G.api.test(key, controller.signal);
      setStatus("Connection successful. Guiden is ready.", "done");
    } catch (error) {
      logFailure("connection-test", error, {}, reference);
      setStatus((error.message || "Connection failed.") + " Reference: " + reference, "error");
    } finally {
      setBusy(false);
    }
  }
  async function init() {
    if (state.initialized) return;
    // Inside the editor, wait for the runtime's init (it brings the editor type, options and theme).
    const p = plugin();
    if (p && !p.info) return;
    state.initialized = true;
    loadSettings();
    applyManagedSettings();
    applyTheme(
      global.Asc &&
        global.Asc.plugin &&
        (global.Asc.plugin.theme || (global.Asc.plugin.info && global.Asc.plugin.info.theme))
    );
    state.editor = await G.host.editorType();
    state.actionId = defaultActions[state.editor];
    renderActions();
    $("shortcutHint").textContent = shortcutName() + " to run · Esc to stop";
    if (!currentKey()) showConnect(true);
    await refreshContext(true);
    attachSelectionEvents();
    if (!state.busy) setStatus("", "");
    if (inWindow() && typeof p.sendToPlugin === "function") p.sendToPlugin("guidenReady", {});
  }
  document.addEventListener("DOMContentLoaded", () => {
    // Key boxes are masked text inputs (see .secret in main.css). A browser that cannot mask them gets real password
    // inputs instead, so a key is never shown in plain text.
    const canMask = window.CSS && CSS.supports && CSS.supports("-webkit-text-security", "disc");
    if (!canMask) document.querySelectorAll("input.secret").forEach(input => (input.type = "password"));
    $("actionGrid").addEventListener("click", e => {
      const b = e.target.closest("button[data-action]");
      if (b) selectAction(b.dataset.action);
    });
    $("generate").addEventListener("click", () => generate());
    $("retry").addEventListener("click", () => {
      $("retry").hidden = true;
      generate(state.retry || undefined);
    });
    $("quickBar").addEventListener("click", e => {
      const b = e.target.closest("button[data-command]");
      if (b) runQuick(b.dataset.command);
    });
    $("quickBar").addEventListener("change", e => {
      const select = e.target.closest("select");
      if (!select || !select.value) return;
      const id = select.value;
      select.value = "";
      runQuick(id);
    });
    $("stop").addEventListener("click", () => state.controller && state.controller.abort());
    $("clear").addEventListener("click", () => {
      $("prompt").value = "";
      localStorage.removeItem(STORAGE.prompt);
      discard();
    });
    $("apply").addEventListener("click", applySuggestion);
    $("trackChanges").addEventListener("change", () => {
      if (state.suggestion) $("apply").textContent = applyLabel(state.suggestion);
      try {
        localStorage.setItem(STORAGE.track, $("trackChanges").checked ? "1" : "0");
      } catch (_) {}
    });
    $("refineChips").addEventListener("click", e => {
      const b = e.target.closest("button[data-refine]");
      if (b) refine(b.dataset.refine);
    });
    $("refineSend").addEventListener("click", () => refine($("refineText").value));
    $("refineText").addEventListener("keydown", e => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        e.stopPropagation();
        refine($("refineText").value);
      }
    });
    $("preview").addEventListener("click", e => {
      const b = e.target.closest("button[data-view]");
      if (b) showView(b.dataset.view);
      const choice = e.target.closest("button[data-design]"),
        s = state.suggestion;
      if (choice && s && s.preview && s.preview.kind === "design" && !state.busy) {
        // Switching design is local: no new request, and the reason no longer applies.
        s.result.design = s.preview.design = choice.dataset.design;
        s.preview.reason = "";
        renderPreview(s.preview);
        $("apply").textContent = applyLabel(s);
      }
    });
    $("prompt").addEventListener("keydown", e => {
      if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
      const box = $("prompt"),
        atStart = box.selectionStart === 0 && box.selectionEnd === 0,
        atEnd = box.selectionStart === box.value.length;
      if (e.key === "ArrowUp" && (atStart || !box.value) && recall(1)) e.preventDefault();
      else if (e.key === "ArrowDown" && atEnd && state.historyIndex >= 0 && recall(-1)) e.preventDefault();
    });
    $("prompt").addEventListener("input", () => (state.historyIndex = -1));
    $("insertKnowledge").addEventListener("click", insertKnowledge);
    $("discard").addEventListener("click", discard);
    $("copy").addEventListener("click", async () => {
      if (!state.suggestion) return;
      try {
        await navigator.clipboard.writeText(C.previewText(state.suggestion.preview, state.suggestion.raw));
        setStatus("Copied.", "done");
      } catch (_) {
        setStatus("Unable to copy.", "error");
      }
    });
    $("refreshContext").addEventListener("click", () => refreshContext(false));
    $("connectButton").addEventListener("click", connect);
    $("connectKey").addEventListener("keydown", e => {
      if (e.key === "Enter") {
        e.preventDefault();
        connect();
      }
    });
    document.addEventListener("keydown", e => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        if (!$("connectCard").hidden) connect();
        else generate();
      } else if (e.key === "Escape" && state.busy && state.controller) {
        e.preventDefault();
        state.controller.abort();
      }
    });
    $("saveSettings").addEventListener("click", saveSettings);
    $("testConnection").addEventListener("click", testConnection);
    $("undo").addEventListener("click", async () => {
      try {
        await G.host.undo();
        $("undo").hidden = true;
        setStatus("Last Guiden action undone.", "done");
        await refreshContext(true);
      } catch (_) {
        setStatus("Undo is not available.", "error");
      }
    });
    const p = plugin();
    if (p) {
      p.init = function () {
        applyManagedSettings();
        applyTheme(p.theme || (p.info && p.info.theme));
        init();
      };
      p.button = function () {};
      p.onThemeChanged = applyTheme;
      p.onUpdateOptions = applyManagedSettings;
      // Messages from the background frame (background.js).
      if (typeof p.attachEvent === "function") {
        p.attachEvent("guidenRun", handleRun);
        p.attachEvent("guidenSelection", scheduleRefresh);
        p.attachEvent("guidenTheme", applyTheme);
        p.attachEvent("guidenOptions", options => {
          if (p.info) p.info.options = options || {};
          applyManagedSettings();
        });
      }
    }
    init();
  });
})(window, document);
