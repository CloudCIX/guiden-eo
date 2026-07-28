(function (global, document) {
  "use strict";
  const G = global.Guiden, C = G.core;
  const $ = id => document.getElementById(id);
  const STORAGE = { key: "cloudcix-guiden.apiKey", remember: "cloudcix-guiden.rememberKey", prompt: "cloudcix-guiden.prompt", tone: "cloudcix-guiden.tone", length: "cloudcix-guiden.length", structure: "cloudcix-guiden.structure", citations: "cloudcix-guiden.keepCitations" };
  const state = { initialized: false, editor: "word", actionId: "writer.ask", context: null, suggestion: null, busy: false, controller: null, refreshTimer: null, selectionVersion: 0, saveTimer: null };
  const editorNames = { word: "Writer", cell: "Cells", slide: "Slides" };
  const defaultActions = { word: "writer.draft", cell: "cells.ask", slide: "slides.ask" };
  const ACTION_ORDER = { word: ["writer.draft", "writer.improve", "writer.rewrite", "writer.fix", "writer.summarise", "writer.ask", "knowledge.word", "writer.convert"] };
  G.actionOrder = ACTION_ORDER;
  G.diagnostics = G.diagnostics || [];
  function diagnosticReference() { return "GDN-" + Date.now().toString(36).slice(-6).toUpperCase(); }
  function logFailure(stage, error, details, reference) {
    const entry = { reference: reference || diagnosticReference(), time: new Date().toISOString(), stage, message: error && error.message || String(error), details: Object.assign({}, error && error.guidenDetails || {}, details || {}) };
    G.diagnostics.push(entry); if (G.diagnostics.length > 50) G.diagnostics.shift();
    if (global.console && console.groupCollapsed) { console.groupCollapsed("[Guiden " + entry.reference + "] " + stage + " failed"); console.error(error); console.log(entry); console.groupEnd(); }
    else if (global.console) console.error("[Guiden]", entry, error);
    return entry.reference;
  }

  function textHtml(value) { return "<p>" + C.escapeHtml(String(value || "")).replace(/\n\n+/g, "</p><p>").replace(/\n/g, "<br>") + "</p>"; }
  function sanitizeHtml(html, trustedUrls) {
    const allowedTags = new Set(["H1","H2","H3","P","STRONG","EM","A","UL","OL","LI","BLOCKQUOTE","TABLE","THEAD","TBODY","TR","TH","TD","BR"]);
    const template = document.createElement("template"); template.innerHTML = String(html || ""); const trustedLinks = Array.isArray(trustedUrls) ? new Set(trustedUrls.filter(C.isSafeLink)) : null;
    function clean(node) {
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType !== 1) return;
        if (!allowedTags.has(child.tagName)) { const fragment = document.createDocumentFragment(); while (child.firstChild) fragment.appendChild(child.firstChild); child.replaceWith(fragment); clean(node); return; }
        const href = child.tagName === "A" ? child.getAttribute("href") || "" : "";
        Array.from(child.attributes).forEach(attribute => child.removeAttribute(attribute.name));
        if (child.tagName === "A" && C.isSafeLink(href) && (!trustedLinks || trustedLinks.has(href))) { child.setAttribute("href", href); child.setAttribute("rel", "noopener noreferrer"); child.setAttribute("target", "_blank"); }
        clean(child);
      });
    }
    clean(template.content); return template.innerHTML;
  }
  function writerNativePayload(html) {
    const template = document.createElement("template"); template.innerHTML = String(html || "");
    if (template.content.querySelector("table")) return { supported: false, blocks: [], linkCount: 0 };
    const blocks = []; let linkCount = 0;
    function collect(node, style, output) {
      if (node.nodeType === 3) { if (node.textContent) output.push({ text: node.textContent, bold: Boolean(style.bold), italic: Boolean(style.italic) }); return; }
      if (node.nodeType !== 1) return;
      const tag = node.tagName;
      if (tag === "BR") { output.push({ text: "\n", bold: Boolean(style.bold), italic: Boolean(style.italic) }); return; }
      if (tag === "A") { const href = node.getAttribute("href") || "", text = node.textContent || href; if (C.isSafeLink(href) && text) { output.push({ text, href, bold: Boolean(style.bold), italic: Boolean(style.italic) }); linkCount += 1; } else Array.from(node.childNodes).forEach(child => collect(child, style, output)); return; }
      const next = { bold: style.bold || tag === "STRONG", italic: style.italic || tag === "EM" }; Array.from(node.childNodes).forEach(child => collect(child, next, output));
    }
    function add(node, prefix, paragraphStyle) { const output = []; if (prefix) output.push({ text: prefix, bold: false, italic: false }); collect(node, {}, output); if (output.some(item => item.text)) blocks.push({ segments: output, paragraphStyle: paragraphStyle || "" }); }
    Array.from(template.content.childNodes).forEach(node => {
      if (node.nodeType === 3) { if (String(node.textContent || "").trim()) add(node, "", ""); return; }
      const tag = node.tagName, styles = { H1: "Heading 1", H2: "Heading 2", H3: "Heading 3", BLOCKQUOTE: "Quote" }; if (tag === "UL" || tag === "OL") { Array.from(node.children).filter(child => child.tagName === "LI").forEach((item, index) => add(item, tag === "OL" ? (index + 1) + ". " : "• ", "List Paragraph")); } else add(node, "", styles[tag] || "");
    });
    return { supported: true, blocks, linkCount };
  }

  G.ui = { sanitizeHtml, writerNativePayload };
  function safeLinksFromHtml(html) { const template = document.createElement("template"); template.innerHTML = String(html || ""); return Array.from(template.content.querySelectorAll("a")).map(link => link.getAttribute("href") || "").filter(C.isSafeLink); }
  function safeRichHtml(html) { return sanitizeHtml(html); }
  function hasVisibleHtml(html) { const template = document.createElement("template"); template.innerHTML = String(html || ""); return Boolean(String(template.content.textContent || "").trim()); }
  function setStatus(text, mode) { $("status").textContent = text; $("status").className = "status " + (mode || ""); }
  function setBusy(value, label) {
    state.busy = value; document.body.classList.toggle("is-busy", value); $("generate").disabled = value; $("generate").textContent = value ? label || "Working…" : (currentAction() && currentAction().mutates ? "Generate preview" : "Ask Guiden"); $("stop").disabled = !value; $("clear").disabled = value; $("apply").disabled = value; $("copy").disabled = value; $("testConnection").disabled = value; document.querySelectorAll(".action-card").forEach(b => b.disabled = value);
  }
  function currentAction() { return G.actions[state.actionId]; }
  function editorActions() { const order = ACTION_ORDER[state.editor] || []; return Object.values(G.actions).filter(a => a.editor === state.editor).sort((a, b) => { const ai = order.indexOf(a.id), bi = order.indexOf(b.id); return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi); }); }
  function registerKnowledge(editor) {
    const id = "knowledge." + editor; if (G.actions[id]) return;
    G.registerAction({ id, editor, label: "Ask Company Knowledge", description: "Search trusted company sources", mutates: false, canInsert: editor !== "cell", defaultInstruction: "Answer using relevant company knowledge.", placeholder: "Ask a question about your company knowledge.", supportedTargets: ["selection", "full_document", "current_slide", "selected_text"],
      buildRequest(snapshot, instruction) { return [{ role: "system", content: "Use only the supplied Guiden retrieval context. Return JSON only: {\"answer\":\"...\",\"sources\":[{\"title\":\"...\",\"url\":\"https://exact-url-from-context\"}]}. Copy source URLs only when they appear exactly in the supplied retrieval context; never invent, repair, or complete a URL. Use an empty sources array when no source URL is present. If the retrieval context does not answer the question, say no relevant company source was found." }, { role: "user", content: instruction }]; },
      parseResponse(content) { const d = C.extractJson(content); if (typeof d.answer !== "string" || !d.answer.trim()) throw new Error("Guiden returned an invalid company-knowledge answer."); const reportedSources = C.safeSources(d.sources); return { answer: d.answer, sources: reportedSources.slice(), reportedSources, foundRelevantSource: false, insertAllowed: false, retrievalUsed: false, provenanceUnavailable: false }; },
      buildPreview(result) { return { kind: "knowledge", answer: result.answer, sources: result.sources, foundRelevantSource: result.foundRelevantSource, insertAllowed: result.insertAllowed, retrievalUsed: result.retrievalUsed, provenanceUnavailable: result.provenanceUnavailable }; }, revalidateTarget: async () => ({ ok: true }), apply: async () => { throw new Error("Company knowledge answers are read-only until you choose Insert answer."); }
    });
  }
  ["word", "cell", "slide"].forEach(registerKnowledge);

  function contextLabel(snapshot) {
    if (!snapshot) return editorNames[state.editor] + " · Selection unavailable";
    if (snapshot.editor === "word") return "Writer · " + snapshot.targetLabel + " · " + snapshot.words + " words";
    if (snapshot.editor === "cell") return "Cells · " + snapshot.sheetName + " · " + snapshot.address + " · " + (snapshot.rows * snapshot.cols) + " cells";
    return "Slides · Slide " + (snapshot.slideIndex + 1) + " · " + (snapshot.target === "selected_text" ? "Selected text" : snapshot.shapes.length + " text areas");
  }
  async function refreshContext(quiet) {
    try { const actionId = state.actionId || defaultActions[state.editor]; const snapshot = await G.host.readSnapshot(state.editor, actionId); state.context = snapshot; $("contextText").textContent = contextLabel(snapshot); if (!quiet) setStatus("Selection refreshed.", "done"); }
    catch (error) { $("contextText").textContent = editorNames[state.editor] + " · Selection unavailable"; if (!quiet) setStatus(error.message || "Unable to read the current selection.", "error"); }
  }
  function scheduleRefresh() { state.selectionVersion += 1; clearTimeout(state.refreshTimer); state.refreshTimer = setTimeout(() => refreshContext(true), 250); }
  function attachSelectionEvents() {
    const plugin = global.Asc && global.Asc.plugin; if (!plugin) return;
    const attach = plugin.attachEditorEvent || plugin.attachEvent;
    if (typeof attach === "function") ["onTargetPositionChanged", "onSelectionChanged", "onWorksheetSelectionChanged"].forEach(name => { try { attach.call(plugin, name, scheduleRefresh); } catch (_) {} });
    global.setInterval(() => { if (!state.busy && !state.suggestion) refreshContext(true); }, 5000);
  }
  function renderActions() {
    const grid = $("actionGrid"); grid.innerHTML = "";
    editorActions().forEach(action => { const button = document.createElement("button"); button.type = "button"; button.className = "action-card" + (action.id === state.actionId ? " selected" : ""); button.dataset.action = action.id; button.innerHTML = "<span>" + C.escapeHtml(action.label) + "</span><small>" + C.escapeHtml(action.description) + "</small>"; grid.appendChild(button); });
    document.body.dataset.editor = state.editor; updateActionUI();
  }
  function updateActionUI() {
    const action = currentAction(); if (!action) return;
    document.querySelectorAll(".action-card").forEach(b => b.classList.toggle("selected", b.dataset.action === action.id));
    $("selectedActionLabel").textContent = action.label; $("selectedActionDescription").textContent = action.description;
    $("prompt").placeholder = action.placeholder || "Add an instruction."; $("promptLabel").textContent = action.mutates ? "Tell Guiden what to change" : "What would you like to know?";
    $("promptHint").textContent = action.mutates ? "Guiden will show a preview and confirm the target before anything changes." : "This action is read-only and will not change your file.";
    $("generate").textContent = action.mutates ? "Generate preview" : "Ask Guiden";
    $("outputOptions").hidden = action.id.startsWith("knowledge.") || action.id.endsWith(".ask") || action.id === "cells.edit" || action.id === "cells.format" || action.id === "cells.formula";
    $("destination").closest("label").hidden = !["cells.formula", "cells.add_column", "cells.classify", "cells.table"].includes(action.id);
    $("tableTheme").closest("label").hidden = action.id !== "cells.table";
    if (action.id === "cells.formula") $("destination").value = "selection";
    if (["cells.add_column", "cells.classify"].includes(action.id)) $("destination").value = "right";
    if (action.id === "cells.table") $("destination").value = "start";
    $("destination").disabled = true;
  }
  function selectAction(id) {
    if (!G.actions[id] || G.actions[id].editor !== state.editor) return; state.actionId = id; state.suggestion = null; $("resultGroup").hidden = true; $("insertKnowledge").hidden = true; renderActions(); $("actionPicker").open = false; refreshContext(true); setStatus("", ""); $("prompt").focus();
  }
  function options() { return { tone: $("tone").value, length: $("length").value, structure: $("structure").value, destination: $("destination").value, tableTheme: $("tableTheme").value, slideCount: $("slideCount").value, keepCitations: $("keepCitations").checked }; }
  function renderSources(sources, unverified) {
    const box = $("sources"), safe = C.safeSources(sources); box.innerHTML = ""; box.hidden = !safe.length; if (!safe.length) return;
    const heading = document.createElement("div"); heading.className = "sources-title"; heading.textContent = unverified ? "Links reported by Guiden — verify before use" : "Sources"; box.appendChild(heading);
    safe.forEach(source => { const a = document.createElement("a"); a.href = source.url; a.target = "_blank"; a.rel = "noopener noreferrer"; a.textContent = source.title; box.appendChild(a); });
  }
  function tableHtml(table) { if (!Array.isArray(table) || !table.length) return ""; return "<div class=\"table-scroll\"><table>" + table.slice(0, 25).map((row, i) => "<tr>" + (Array.isArray(row) ? row : [row]).map(v => "<" + (i ? "td" : "th") + ">" + C.escapeHtml(v) + "</" + (i ? "td" : "th") + ">").join("") + "</tr>").join("") + "</table>" + (table.length > 25 ? "<p>Preview shows the first 25 rows.</p>" : "") + "</div>"; }
  function beginStreaming(action, snapshot) {
    const preview = $("preview"), resultActions = document.querySelector(".result-actions");
    $("operationSummary").textContent = action.label + " · " + snapshot.targetLabel; $("previewBadge").textContent = "Generating"; $("previewBadge").className = "badge loading"; $("targetWarning").hidden = true; renderSources([]);
    preview.classList.add("streaming"); preview.setAttribute("aria-busy", "true"); preview.innerHTML = "<div class=stream-lead><span></span>Guiden is writing…</div><div class=stream-output>Waiting for the first words…</div>";
    if (resultActions) resultActions.hidden = true; $("resultGroup").hidden = false; try { $("resultGroup").scrollIntoView({ behavior: "smooth", block: "nearest" }); } catch (_) {}
  }
  function updateStreaming(content) {
    const preview = $("preview"), output = preview.querySelector(".stream-output"), readable = C.streamingText(content); if (!output) return;
    output.textContent = readable || "Structuring the response…"; preview.scrollTop = preview.scrollHeight;
  }
  function finishStreaming() { const preview = $("preview"), resultActions = document.querySelector(".result-actions"); preview.classList.remove("streaming"); preview.removeAttribute("aria-busy"); if (resultActions) resultActions.hidden = false; }
  function failStreaming(label) { const preview = $("preview"), lead = preview.querySelector(".stream-lead"); preview.classList.remove("streaming"); preview.removeAttribute("aria-busy"); $("previewBadge").textContent = label; $("previewBadge").className = "badge failed"; if (lead) lead.textContent = label + " — this incomplete response cannot be applied."; }
  function renderPreview(preview) { finishStreaming();
    const parts = []; $("targetWarning").hidden = !preview.warning; $("targetWarning").textContent = preview.warning || "";
    if (["answer", "knowledge", "analysis"].includes(preview.kind)) { parts.push("<article class=\"response-flow\">" + textHtml(preview.answer) + "</article>"); if (preview.table) parts.push(tableHtml(preview.table)); if (preview.kind === "knowledge" && preview.provenanceUnavailable) parts.push("<div class=empty-source>This Guiden server did not provide retrieval verification. The answer may still use your configured company corpora; review the content and any reported links before inserting.</div>"); else if (preview.kind === "knowledge" && !preview.retrievalUsed) parts.push("<div class=empty-source>No relevant company source was found. Guiden did not use an unverified answer.</div>"); else if (preview.kind === "knowledge" && preview.retrievalUsed && !(preview.sources || []).length) parts.push("<div class=empty-source>This answer used company knowledge, but that source has no clickable URL.</div>"); }
    if (preview.kind === "needs-input") parts.push("<div class=needs-input><strong>Nothing will be changed.</strong> " + C.escapeHtml(preview.message) + "</div>");
    if (preview.kind === "writer-new") parts.push("<article class=\"response-flow\">" + safeRichHtml(preview.afterHtml) + "</article>");
    if (preview.kind === "writer-diff") parts.push("<details class=\"before-disclosure\"><summary>View original text</summary><div>" + textHtml(preview.before) + "</div></details><div class=\"suggestion-label\">Suggested version</div><article class=\"response-flow\">" + safeRichHtml(preview.afterHtml) + "</article>");
    if (preview.kind === "text-diff") parts.push("<details class=\"before-disclosure\"><summary>View original text</summary><div>" + textHtml(preview.before) + "</div></details><div class=\"suggestion-label\">Suggested version</div><article class=\"response-flow\">" + textHtml(preview.after) + "</article>");
    if (preview.kind === "cell-updates") { if (preview.narrative) parts.push(textHtml(preview.narrative)); parts.push("<div class=\"change-count\">" + preview.updates.length + " cells will change</div>" + tableHtml([["Cell","Before","After"]].concat(preview.updates.map(u => [C.rangeAddress({ startRow: preview.snapshot.start.startRow + u.row, startCol: preview.snapshot.start.startCol + u.column }, 1, 1), u.before, u.after])))); }
    if (preview.kind === "formula") parts.push("<div class=\"destination\">Destination: " + C.escapeHtml(preview.address) + "</div><code class=\"formula\">" + C.escapeHtml(preview.formula) + "</code>" + textHtml(preview.explanation));
    if (preview.kind === "new-column") parts.push(textHtml(preview.narrative) + "<div class=\"destination\">Destination: " + C.escapeHtml(preview.address) + "</div>" + tableHtml([[preview.header]].concat(preview.values.map(v => [v]))));
    if (preview.kind === "formatting") { const cells = Array.from(new Set(preview.cells.map(item => C.rangeAddress({ startRow: preview.snapshot.start.startRow + item.row, startCol: preview.snapshot.start.startCol + item.column }, 1, 1)))); parts.push(textHtml(preview.narrative) + "<div class=\"change-count\">" + cells.length + " matching cells</div><div class=\"swatches\">" + preview.cells.slice(0, 8).map(item => "<span><i style=\"background:" + C.escapeHtml(item.style.fill || item.style.fontColor || "#e5e7eb") + "\"></i>" + C.escapeHtml(C.rangeAddress({ startRow: preview.snapshot.start.startRow + item.row, startCol: preview.snapshot.start.startCol + item.column }, 1, 1)) + "</span>").join("") + "</div>"); }
    if (preview.kind === "table") parts.push(textHtml(preview.narrative) + "<div class=\"destination\">Destination: " + C.escapeHtml(preview.address) + " · " + preview.table.length + "×" + preview.table[0].length + "</div>" + tableHtml(preview.table));
    if (preview.kind === "shape-updates") parts.push(textHtml(preview.narrative) + preview.updates.map(u => "<div class=\"shape-change\"><strong>Text area " + (u.shapeIndex + 1) + "</strong><div class=\"diff before\">" + textHtml(u.before) + "</div><div class=\"diff after\">" + textHtml(u.after) + "</div></div>").join(""));
    if (preview.kind === "new-slide") parts.push("<div class=\"slide-preview\"><h2>" + C.escapeHtml(preview.title) + "</h2><ul>" + preview.bullets.map(x => "<li>" + C.escapeHtml(x) + "</li>").join("") + "</ul></div>");
    if (preview.kind === "outline") parts.push(textHtml(preview.answer) + preview.slides.map((s, i) => "<div class=\"outline-slide\"><strong>" + (i + 1) + ". " + C.escapeHtml(s.title || "Untitled") + "</strong><ul>" + (s.bullets || []).map(x => "<li>" + C.escapeHtml(x) + "</li>").join("") + "</ul></div>").join(""));
    if (preview.kind === "notes") parts.push("<div class=\"diff-label\">Speaker notes</div>" + textHtml(preview.notes));
    $("preview").innerHTML = parts.join(""); renderSources(preview.sources || [], preview.provenanceUnavailable);
  }
  function appliedMessage(action, result) {
    if (action.editor === "cell") return (action.id === "cells.format" ? "Formatted " : "Updated ") + result.changed + (result.changed === 1 ? " cell" : " cells") + " in " + result.target + ".";
    if (action.id === "slides.notes") return "Added speaker notes to " + result.target.replace(/ notes$/, "") + ".";
    if (action.id === "slides.new") return "Created 1 new slide.";
    if (action.editor === "slide") return "Updated " + result.changed + (result.changed === 1 ? " text area" : " text areas") + " on " + result.target + ".";
    return "Updated " + result.target + ".";
  }
  async function generate() {
    if (state.busy) return;
    const action = currentAction(), key = $("apiKey").value.trim(), instruction = $("prompt").value.trim() || action.defaultInstruction, reference = diagnosticReference();
    let snapshot = null, rawResponse = "";
    if (!key) { setStatus("Add your CloudCIX API key in Settings first.", "error"); return; }
    if (!instruction) { setStatus("Add an instruction or question.", "error"); return; }
    setBusy(true, action.mutates ? "Preparing preview…" : "Asking…"); setStatus(action.mutates ? "Reading the current selection…" : "Reading context for your question…", "loading"); state.controller = new AbortController(); state.suggestion = null; $("resultGroup").hidden = true;
    try {
      snapshot = await G.host.readSnapshot(state.editor, action.id); snapshot.actionId = action.id;
      if (action.supportedTargets && !action.supportedTargets.includes(snapshot.target)) throw new Error("This action is not available for the current target. Update your selection and try again.");
      state.context = snapshot; $("contextText").textContent = contextLabel(snapshot); setStatus(action.mutates ? "Guiden is generating a preview for " + snapshot.targetLabel + "…" : "Guiden is preparing an answer for " + snapshot.targetLabel + "…", "loading");
      const requestMessages = action.buildRequest(snapshot, instruction, options());
      beginStreaming(action, snapshot); rawResponse = await G.api.complete(key, requestMessages, state.controller.signal, updateStreaming);
      let result;
      try { result = action.parseResponse(rawResponse, snapshot); }
      catch (schemaError) {
        setStatus("Guiden returned the content in the wrong format. Repairing the preview…", "loading");
        const lead = $("preview").querySelector(".stream-lead"); if (lead) lead.textContent = "Guiden is repairing the response format…";
        rawResponse = await G.api.complete(key, C.schemaRetryMessages(requestMessages), state.controller.signal, updateStreaming);
        result = action.parseResponse(rawResponse, snapshot);
      }
      const metadata = G.api.lastMetadata || { available: false, retrievalUsed: false, sources: [] };
      if (Object.prototype.hasOwnProperty.call(result, "sources")) result.sources = metadata.available ? metadata.sources.slice() : C.safeSources(result.reportedSources || result.sources);
      if (typeof result.replacementHtml === "string") { const proposedLinks = safeLinksFromHtml(result.replacementHtml), allowedLinks = safeLinksFromHtml(snapshot.sourceHtml).concat(metadata.available ? metadata.sources.map(source => source.url) : proposedLinks); result.unverifiedLinks = !metadata.available && proposedLinks.length > 0; result.replacementHtml = sanitizeHtml(result.replacementHtml, allowedLinks); if (!hasVisibleHtml(result.replacementHtml)) throw new Error("Guiden returned empty or unsupported document content. Nothing can be applied."); }
      if (typeof result.replacementHtml === "string" && action.editor === "word" && action.mutates && $("keepCitations").checked) { const citationSources = metadata.available ? metadata.sources : C.safeSources(result.sources), existingUrls = new Set(safeLinksFromHtml(result.replacementHtml)), missingSources = citationSources.filter(source => !existingUrls.has(source.url)); if (missingSources.length) result.replacementHtml += "<h3>Sources</h3><ul>" + missingSources.map(source => "<li><a href=\"" + C.escapeHtml(source.url) + "\">" + C.escapeHtml(source.title) + "</a></li>").join("") + "</ul>"; }
      if (action.id.indexOf("knowledge.") === 0) { if (!metadata.available) { result.provenanceUnavailable = true; result.retrievalUsed = null; result.foundRelevantSource = false; result.insertAllowed = true; } else { result.provenanceUnavailable = false; result.retrievalUsed = metadata.retrievalUsed; result.foundRelevantSource = metadata.retrievalUsed; result.insertAllowed = metadata.retrievalUsed; if (!metadata.retrievalUsed) { result.answer = "No relevant company source was found for this question. Guiden has withheld an unverified answer."; result.sources = []; } } }
      const preview = action.buildPreview(result, snapshot), applicable = action.mutates && !result.needsInput; state.suggestion = { actionId: action.id, selectionVersion: state.selectionVersion, snapshot, result, preview, raw: rawResponse };
      if (result.unverifiedLinks) preview.warning = (preview.warning ? preview.warning + " " : "") + "This legacy Guiden response contains unverified links. Review their destinations before applying.";
      $("operationSummary").textContent = action.label + " · " + snapshot.targetLabel; $("previewBadge").textContent = result.needsInput ? "Needs information" : preview.provenanceUnavailable ? "Compatibility" : action.mutates ? "Will change file" : "Read-only"; $("previewBadge").className = "badge " + (result.needsInput ? "attention" : action.mutates ? "mutation" : "readonly");
      renderPreview(preview); $("apply").hidden = !applicable; $("apply").textContent = action.id === "slides.notes" ? "Add notes" : action.id === "slides.new" ? "Create slide" : "Apply changes"; $("copy").hidden = Boolean(result.needsInput); $("insertKnowledge").hidden = !(action.canInsert && preview.insertAllowed && (state.editor !== "slide" || snapshot.target === "selected_text")); $("insertKnowledge").textContent = state.editor === "slide" ? "Insert on slide" : "Insert in document"; $("resultGroup").hidden = false;
      setStatus(result.needsInput ? "Guiden needs more information. Nothing can be applied yet." : action.mutates ? "Preview ready. Review the target and changes before applying." : "Answer ready. Nothing in your file has changed.", result.needsInput ? "attention" : "done"); localStorage.setItem(STORAGE.prompt, $("prompt").value);
    } catch (error) {
      if (error.name === "AbortError") { if (snapshot) failStreaming("Stopped"); setStatus("Stopped. The incomplete response was not applied.", "error"); }
      else { if (snapshot) failStreaming("Failed"); logFailure("generate", error, { actionId: action.id, editor: state.editor, target: snapshot && snapshot.targetLabel || null, responseLength: rawResponse.length, rawResponse: rawResponse.slice(0, 12000) }, reference); setStatus((error.message || "Unable to generate a result.") + " Reference: " + reference, "error"); }
    } finally { state.controller = null; setBusy(false); }
  }
  async function applySuggestion() {
    const suggestion = state.suggestion; if (!suggestion || state.busy) return; const action = G.actions[suggestion.actionId]; if (!action || !action.mutates || suggestion.result.needsInput) return;
    const reference = diagnosticReference(); setBusy(true, "Applying…"); setStatus("Confirming the target has not changed…", "loading");
    try { if (suggestion.selectionVersion !== state.selectionVersion) throw new Error("The target changed. Generate again for the current selection."); const valid = await action.revalidateTarget(suggestion.snapshot); if (!valid.ok) throw new Error("The target changed. Generate again for the current selection."); const result = await action.apply(suggestion.result, suggestion.snapshot); C.validateApplyResult(result); setStatus(appliedMessage(action, result), "done"); $("resultGroup").hidden = true; state.suggestion = null; const canUndo = await G.host.canUndo(); $("undo").hidden = !canUndo; await refreshContext(true); }
    catch (error) { logFailure("apply", error, { actionId: action.id, editor: state.editor, target: suggestion.snapshot.targetLabel }, reference); setStatus((error.message || "Unable to apply the changes.") + " Reference: " + reference, "error"); }
    finally { setBusy(false); }
  }
  async function insertKnowledge() {
    const s = state.suggestion; if (!s || !s.preview || !s.preview.insertAllowed) return; const reference = diagnosticReference(); setBusy(true, "Inserting…");
    try { if (state.editor === "slide") { if (s.selectionVersion !== state.selectionVersion) throw new Error("The target changed. Generate again for the current selection."); const valid = await G.host.revalidate(s.snapshot); if (!valid.ok) throw new Error("The target changed. Generate again for the current selection."); } let result; const html = textHtml(s.result.answer) + ($("keepCitations").checked ? "<h3>Sources</h3><ul>" + s.result.sources.map(x => "<li><a href=\"" + C.escapeHtml(x.url) + "\">" + C.escapeHtml(x.title) + "</a></li>").join("") + "</ul>" : ""); if (state.editor === "word") result = await G.host.insertWriter(sanitizeHtml(html)); else result = await G.host.applySelectedSlideText(s.result.answer, s.snapshot); C.validateApplyResult(result); setStatus("Inserted the company-knowledge answer into " + result.target + ".", "done"); $("resultGroup").hidden = true; state.suggestion = null; $("undo").hidden = !(await G.host.canUndo()); }
    catch (error) { logFailure("insert-knowledge", error, { actionId: s.actionId, editor: state.editor, target: s.snapshot && s.snapshot.targetLabel }, reference); setStatus((error.message || "Unable to insert the answer.") + " Reference: " + reference, "error"); } finally { setBusy(false); }
  }
  function discard() { state.suggestion = null; $("resultGroup").hidden = true; $("insertKnowledge").hidden = true; setStatus("Preview discarded. Nothing was changed.", ""); }
  function saveSettings() { try { const remember = $("rememberKey").checked; localStorage.setItem(STORAGE.remember, remember ? "1" : "0"); ["tone","length","structure"].forEach(id => localStorage.setItem(STORAGE[id], $(id).value)); localStorage.setItem(STORAGE.citations, $("keepCitations").checked ? "1" : "0"); if (remember) localStorage.setItem(STORAGE.key, $("apiKey").value); else localStorage.removeItem(STORAGE.key); feedback("Saved ✓", "done"); setStatus(remember ? "Settings and API key saved." : "Settings saved. API key is not stored.", "done"); } catch (_) { feedback("Save failed", "error"); setStatus("Settings could not be saved in this browser.", "error"); } }
  function feedback(label, mode) { clearTimeout(state.saveTimer); const b = $("saveSettings"), old = "Save settings"; b.textContent = label; b.className = mode; state.saveTimer = setTimeout(() => { b.textContent = old; b.className = ""; }, 2200); }
  function loadSettings() { try { const remember = localStorage.getItem(STORAGE.remember) === "1"; $("rememberKey").checked = remember; $("apiKey").value = remember ? localStorage.getItem(STORAGE.key) || "" : ""; $("prompt").value = localStorage.getItem(STORAGE.prompt) || ""; ["tone","length","structure"].forEach(id => { const v = localStorage.getItem(STORAGE[id]); if (v) $(id).value = v; }); if (localStorage.getItem(STORAGE.citations) !== null) $("keepCitations").checked = localStorage.getItem(STORAGE.citations) === "1"; } catch (_) {} }
  async function testConnection() { const key = $("apiKey").value.trim(), reference = diagnosticReference(); if (!key) { setStatus("Enter an API key before testing.", "error"); return; } setBusy(true, "Testing…"); const controller = new AbortController(); try { await G.api.test(key, controller.signal); setStatus("Connection successful. Guiden is ready.", "done"); } catch (error) { logFailure("connection-test", error, {}, reference); setStatus((error.message || "Connection failed.") + " Reference: " + reference, "error"); } finally { setBusy(false); } }
  async function init() { if (state.initialized) return; state.initialized = true;
    loadSettings(); state.editor = await G.host.editorType(); state.actionId = defaultActions[state.editor]; renderActions(); await refreshContext(true); attachSelectionEvents(); setStatus("", "");
  }
  document.addEventListener("DOMContentLoaded", () => {
    $("actionGrid").addEventListener("click", e => { const b = e.target.closest("button[data-action]"); if (b) selectAction(b.dataset.action); }); $("generate").addEventListener("click", generate); $("stop").addEventListener("click", () => state.controller && state.controller.abort()); $("clear").addEventListener("click", () => { $("prompt").value = ""; localStorage.removeItem(STORAGE.prompt); discard(); }); $("apply").addEventListener("click", applySuggestion); $("insertKnowledge").addEventListener("click", insertKnowledge); $("discard").addEventListener("click", discard); $("copy").addEventListener("click", async () => { if (!state.suggestion) return; try { await navigator.clipboard.writeText(state.suggestion.preview.answer || state.suggestion.raw); setStatus("Copied.", "done"); } catch (_) { setStatus("Unable to copy.", "error"); } }); $("refreshContext").addEventListener("click", () => refreshContext(false)); $("saveSettings").addEventListener("click", saveSettings); $("testConnection").addEventListener("click", testConnection); $("undo").addEventListener("click", async () => { try { await G.host.undo(); $("undo").hidden = true; setStatus("Last Guiden action undone.", "done"); await refreshContext(true); } catch (_) { setStatus("Undo is not available.", "error"); } });
    if (global.Asc && global.Asc.plugin) { global.Asc.plugin.init = init; global.Asc.plugin.button = function () {}; } init();
  });
})(window, document);
