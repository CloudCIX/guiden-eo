(function (global) {
  "use strict";
  const G = global.Guiden, C = G.core;
  const host = {};
  host.execute = function (name, args) { return new Promise((resolve, reject) => { try { global.Asc.plugin.executeMethod(name, args || [], resolve); } catch (e) { reject(e); } }); };
  host.command = function (fn) { return new Promise((resolve, reject) => { try { global.Asc.plugin.callCommand(fn, false, true, resolve); } catch (e) { reject(e); } }); };
  host.dynamicCommand = function (body, data) {
    const fn = new Function("return function(){const input=" + JSON.stringify(data || {}) + ";" + body + "}")();
    return host.command(fn);
  };
  host.editorType = async function () {
    const info = global.Asc && global.Asc.plugin && global.Asc.plugin.info;
    const type = info && (info.editorType || info.editorTypeName || info.type);
    if (["word", "cell", "slide"].includes(type)) return type;
    return host.command(function () {
      try { if (Api.GetActiveSheet && Api.GetActiveSheet()) return "cell"; } catch (_) {}
      try { if (Api.GetPresentation && Api.GetPresentation()) return "slide"; } catch (_) {}
      return "word";
    }).catch(() => "word");
  };
  host.readWriter = async function (actionId) {
    let selected = await host.execute("GetSelectedText", [{}]).catch(() => ""); selected = selected || "";
    const selectedHtml = selected.trim() ? await host.execute("GetSelectedHtml", []).catch(() => "") : "";
    const insertsAtCursor = actionId === "writer.draft";
    let target = selected.trim() ? "selection" : insertsAtCursor ? "cursor" : "full_document", text = selected;
    if (target === "full_document") text = await host.command(function () {
      const doc = Api.GetDocument(), options = { Numbering: true, Math: true, TableCellSeparator: "\t", TableRowSeparator: "\n", ParaSeparator: "\n\n", NewLineSeparator: "\n" };
      if (doc && typeof doc.GetText === "function") return doc.GetText(options) || "";
      const out = [], paragraphs = doc && doc.GetAllParagraphs ? doc.GetAllParagraphs() : [];
      for (let i = 0; i < paragraphs.length; i += 1) if (paragraphs[i] && paragraphs[i].GetText) out.push(paragraphs[i].GetText(options) || "");
      return out.join("\n\n");
    }).catch(() => "");
    const source = { target, text: text || "", html: selectedHtml || "" };
    return { editor: "word", actionId, target, targetLabel: target === "selection" ? "Selected text" : target === "cursor" ? "Current cursor" : "Whole document", text: source.text, sourceHtml: source.html, words: C.wordCount(source.text), fingerprint: C.fingerprint(source) };
  };
  host.readCells = async function (actionId) {
    const result = await host.command(function () {
      const sheet = Api.GetActiveSheet(), range = Api.GetSelection ? Api.GetSelection() : sheet.GetSelection();
      if (!range) return { ok: false, error: "Unable to read the selected cells." };
      return { ok: true, sheet: sheet.GetName ? sheet.GetName() : "Sheet", address: range.GetAddress ? range.GetAddress(true, true, "xlA1", false) : "", values: range.GetValue ? range.GetValue() : "", formulas: range.GetFormula ? range.GetFormula() : "", rows: range.GetRowsCount ? range.GetRowsCount() : 0, cols: range.GetColumnsCount ? range.GetColumnsCount() : 0 };
    });
    if (!result || !result.ok) throw new Error(result && result.error || "Unable to read the selected cells.");
    const values = C.normalizeMatrix(result.values), formulas = C.normalizeMatrix(result.formulas), parsed = C.parseA1Address(result.address), d = C.matrixDimensions(values);
    const snapshot = { editor: "cell", actionId, sheetName: result.sheet || "Sheet", address: result.address || "selected range", start: parsed, rows: Number(result.rows) || parsed && parsed.rows || d.rows, cols: Number(result.cols) || parsed && parsed.cols || d.cols, values, formulas };
    snapshot.target = snapshot.address; snapshot.targetLabel = snapshot.sheetName + " · " + snapshot.address; snapshot.text = values.map(r => r.join("\t")).join("\n");
    snapshot.fingerprint = C.fingerprint({ sheetName: snapshot.sheetName, address: snapshot.address, rows: snapshot.rows, cols: snapshot.cols, values, formulas });
    return snapshot;
  };
  host.readSlides = async function (actionId) {
    const selected = await host.execute("GetSelectedText", [{}]).catch(() => "");
    const result = await host.command(function () {
      const presentation = Api.GetPresentation();
      const slide = presentation.GetCurrentVisibleSlide ? presentation.GetCurrentVisibleSlide() : presentation.GetCurrentSlide();
      const all = presentation.GetAllSlides ? presentation.GetAllSlides() : [];
      let slideIndex = 0; for (let i = 0; i < all.length; i += 1) if (all[i] === slide) slideIndex = i;
      const drawings = slide && slide.GetAllDrawings ? slide.GetAllDrawings() : [], shapes = [];
      for (let i = 0; i < drawings.length; i += 1) {
        const drawing = drawings[i]; if (!drawing || !drawing.GetDocContent) continue;
        const content = drawing.GetDocContent(); let text = "";
        if (content && content.GetText) text = content.GetText({ ParaSeparator: "\n", NewLineSeparator: "\n" }) || "";
        else if (content && content.GetElementsCount && content.GetElement) { const lines = []; for (let j = 0; j < content.GetElementsCount(); j += 1) { const p = content.GetElement(j); if (p && p.GetText) lines.push(p.GetText()); } text = lines.join("\n"); }
        shapes.push({ index: i, text: text });
      }
      return { ok: true, slideIndex: slideIndex, shapes: shapes };
    });
    if (!result || !result.ok) throw new Error("Unable to read the current slide.");
    const selectedPreferred = actionId === "slides.improve" || String(actionId || "").indexOf("knowledge.") === 0;
    const target = selectedPreferred && selected && String(selected).trim() ? "selected_text" : "current_slide";
    const source = { slideIndex: result.slideIndex, target, selected: selected || "", shapes: result.shapes || [] };
    return { editor: "slide", actionId, target, targetLabel: target === "selected_text" ? "Selected slide text" : "Slide " + (result.slideIndex + 1), slideIndex: result.slideIndex, selectedText: selected || "", shapes: result.shapes || [], text: target === "selected_text" ? selected : (result.shapes || []).map(s => "[" + s.index + "] " + s.text).join("\n\n"), fingerprint: C.fingerprint(source) };
  };
  host.readSnapshot = async function (editor, actionId) { return editor === "cell" ? host.readCells(actionId) : editor === "slide" ? host.readSlides(actionId) : host.readWriter(actionId); };
  host.revalidate = async function (snapshot) {
    const fresh = await host.readSnapshot(snapshot.editor, snapshot.actionId);
    const same = fresh.fingerprint === snapshot.fingerprint && fresh.target === snapshot.target && (snapshot.editor !== "cell" || fresh.address === snapshot.address) && (snapshot.editor !== "slide" || fresh.slideIndex === snapshot.slideIndex);
    return { ok: same, fresh };
  };
  host.pasteWriter = async function (html) {
    const payload = G.ui && typeof G.ui.writerNativePayload === "function" ? G.ui.writerNativePayload(html) : null;
    if (payload && payload.supported && payload.linkCount > 0 && payload.blocks.length) {
      const result = await host.dynamicCommand('const doc=Api.GetDocument();if(!doc||typeof doc.InsertContent!=="function"||typeof Api.CreateParagraph!=="function"||typeof Api.CreateHyperlink!=="function")return {ok:false,unsupported:true};const content=[];for(let i=0;i<input.blocks.length;i+=1){const block=input.blocks[i],para=Api.CreateParagraph(),segments=block.segments;if(block.paragraphStyle&&typeof para.SetStyle==="function"){try{para.SetStyle(block.paragraphStyle);}catch(_){}}for(let j=0;j<segments.length;j+=1){const segment=segments[j],text=String(segment.text||"");if(!text)continue;if(segment.href){const link=Api.CreateHyperlink(segment.href,text,"");if(!link)return {ok:false,error:"A native hyperlink could not be created."};const added=typeof para.AddElement==="function"?para.AddElement(link):para.Push(link);if(added===false)return {ok:false,error:"A native hyperlink could not be inserted."};continue;}const parts=text.split(String.fromCharCode(10));for(let k=0;k<parts.length;k+=1){const run=Api.CreateRun();if(parts[k])run.AddText(parts[k]);if(segment.bold&&typeof run.SetBold==="function")run.SetBold(true);if(segment.italic&&typeof run.SetItalic==="function")run.SetItalic(true);if(k<parts.length-1&&typeof run.AddLineBreak==="function")run.AddLineBreak();if(para.Push(run)===false)return {ok:false,error:"Writer text could not be inserted."};}}content.push(para);}const ok=doc.InsertContent(content,false);return {ok:ok!==false,changed:ok===false?0:1,links:input.linkCount};', payload).catch(error => ({ ok: false, error: error && error.message || String(error) }));
      if (result && result.ok && result.links === payload.linkCount) return result;
      if (global.console && console.warn) console.warn("[Guiden] Native Writer hyperlinks unavailable; using HTML import.", result || { stage: "no-result" });
      await host.execute("PasteHtml", [String(html || "")]);
      return { ok: true, changed: 1, links: payload.linkCount, warnings: ["Euro-Office used its HTML compatibility importer for document links."] };
    }
    await host.execute("PasteHtml", [String(html || "")]); return { ok: true, changed: 1, links: 0, warnings: [] };
  };

  host.writerHtmlDocument = function (html) {
    const body = String(html || "").trim().replace(/\s+(?:target|rel)="[^"]*"/gi, "");
    return "<!doctype html><html><head><meta charset=\"utf-8\"></head><body>" + body + "</body></html>";
  };
  host.applyWriter = async function (html, snapshot) {
    if (snapshot.target === "full_document" && snapshot.text.trim()) {
      let selected = false;
      try { await host.execute("SelectAll", []); await new Promise(resolve => global.setTimeout(resolve, 120)); selected = Boolean(String(await host.execute("GetSelectedText", [{}]).catch(() => "")).trim()); } catch (_) {}
      if (!selected) throw new Error("This Euro-Office build cannot safely select the whole document.");
    }
    const pasted = await host.pasteWriter(html);
    return { ok: true, changed: 1, target: snapshot.targetLabel, warnings: pasted.warnings || [] };
  };
  host.insertWriter = async function (html) { const pasted = await host.pasteWriter(html); return { ok: true, changed: 1, target: "current cursor or selection", warnings: pasted.warnings || [] }; };
  host.applyCellUpdates = async function (updates, snapshot) {
    const cells = updates.map(u => ({ address: C.rangeAddress({ startRow: snapshot.start.startRow + u.row, startCol: snapshot.start.startCol + u.column }, 1, 1), value: u.after }));
    const result = await host.dynamicCommand('const sheet=Api.GetActiveSheet();let changed=0;for(let i=0;i<input.cells.length;i+=1){const item=input.cells[i],range=sheet&&sheet.GetRange(item.address);if(!range||typeof range.SetValue!=="function")return {ok:false,changed:changed,error:"A target cell cannot be edited."};const ok=range.SetValue(item.value);if(ok===false)return {ok:false,changed:changed,error:"Euro-Office rejected a cell change."};changed+=1;}return {ok:true,changed:changed};', { cells });
    if (!result || !result.ok || result.changed !== cells.length) throw new Error(result && result.error || "Only part of the cell update was applied.");
    return { ok: true, changed: result.changed, target: snapshot.address, warnings: [] };
  };
  host.applyFormula = async function (targetAddress, formula) {
    const result = await host.dynamicCommand('const sheet=Api.GetActiveSheet(),range=sheet&&sheet.GetRange(input.address);if(!range)return {ok:false,error:"Formula destination is unavailable."};let ok;if(typeof range.SetFormula==="function")ok=range.SetFormula(input.formula);else if(typeof range.SetValue==="function")ok=range.SetValue(input.formula);else return {ok:false,error:"Formulas are not supported."};return {ok:ok!==false,changed:ok===false?0:1};', { address: targetAddress, formula });
    if (!result || !result.ok || result.changed !== 1) throw new Error(result && result.error || "The formula was not applied.");
    return { ok: true, changed: 1, target: targetAddress, warnings: [] };
  };
  host.applyStyles = async function (cells, snapshot) {
    const payload = cells.map(item => ({ address: C.rangeAddress({ startRow: snapshot.start.startRow + item.row, startCol: snapshot.start.startCol + item.column }, 1, 1), style: item.style }));
    const result = await host.dynamicCommand('const sheet=Api.GetActiveSheet();function color(hex){return Api.CreateColorFromRGB(parseInt(hex.slice(1,3),16),parseInt(hex.slice(3,5),16),parseInt(hex.slice(5,7),16));}let changed=0;for(let i=0;i<input.cells.length;i+=1){const item=input.cells[i],r=sheet&&sheet.GetRange(item.address),s=item.style;if(!r)return {ok:false,changed:changed,error:"A formatting target is unavailable."};let supported=false;if(s.fill&&r.SetFillColor){r.SetFillColor(color(s.fill));supported=true;}if(s.fontColor&&r.SetFontColor){r.SetFontColor(color(s.fontColor));supported=true;}if(typeof s.bold==="boolean"&&r.SetBold){r.SetBold(s.bold);supported=true;}if(typeof s.italic==="boolean"&&r.SetItalic){r.SetItalic(s.italic);supported=true;}if(s.align&&r.SetAlignHorizontal){r.SetAlignHorizontal(s.align);supported=true;}if(s.numberFormat&&r.SetNumberFormat){r.SetNumberFormat(s.numberFormat);supported=true;}if(!supported)return {ok:false,changed:changed,error:"Formatting is not supported by this Euro-Office build."};changed+=1;}return {ok:true,changed:changed};', { cells: payload });
    if (!result || !result.ok || result.changed !== payload.length) throw new Error(result && result.error || "Only part of the formatting was applied.");
    return { ok: true, changed: result.changed, target: snapshot.address, warnings: [] };
  };
  host.readRange = async function (address) { return host.dynamicCommand('const s=Api.GetActiveSheet(),r=s&&s.GetRange(input.address);if(!r)return {ok:false};return {ok:true,values:r.GetValue?r.GetValue():"",formulas:r.GetFormula?r.GetFormula():""};', { address }); };
  host.applyTable = async function (table, address, styles) {
    const result = await host.dynamicCommand('const s=Api.GetActiveSheet(),r=s&&s.GetRange(input.address);if(!r||!r.SetValue)return {ok:false,error:"Table destination cannot be edited."};const ok=r.SetValue(input.table);if(ok===false)return {ok:false,error:"Euro-Office rejected the table."};return {ok:true,changed:input.table.length*input.table[0].length};', { table, address });
    if (!result || !result.ok) throw new Error(result && result.error || "The table was not applied.");
    if (styles && styles.length) await host.applyStyles(styles, { start: C.parseA1Address(address), address });
    return { ok: true, changed: result.changed, target: address, warnings: [] };
  };
  host.applySelectedSlideText = async function (text, snapshot) { await host.execute("PasteText", [text]); return { ok: true, changed: 1, target: snapshot.targetLabel, warnings: [] }; };
  host.applyShapeUpdates = async function (updates, snapshot) {
    const result = await host.dynamicCommand('const p=Api.GetPresentation(),slide=p.GetCurrentVisibleSlide?p.GetCurrentVisibleSlide():p.GetCurrentSlide(),drawings=slide&&slide.GetAllDrawings?slide.GetAllDrawings():[];let changed=0;for(let i=0;i<input.updates.length;i+=1){const u=input.updates[i],d=drawings[u.shapeIndex],c=d&&d.GetDocContent?d.GetDocContent():null;if(!c||!c.RemoveAllElements)return {ok:false,changed:changed,error:"The intended text area cannot be edited."};c.RemoveAllElements();const lines=String(u.after).split(/\\r?\\n/);for(let j=0;j<lines.length;j+=1){const para=Api.CreateParagraph();para.AddText(lines[j]);if(c.Push)c.Push(para);else if(c.AddElement)c.AddElement(j,para);else return {ok:false,changed:changed,error:"Text cannot be written to this area."};}changed+=1;}return {ok:true,changed:changed};', { updates });
    if (!result || !result.ok || result.changed !== updates.length) throw new Error(result && result.error || "Only part of the slide was updated.");
    return { ok: true, changed: result.changed, target: snapshot.targetLabel, warnings: [] };
  };
  host.createSlide = async function (slideData) {
    const result = await host.dynamicCommand('const p=Api.GetPresentation();if(!p||!p.AddSlide||!Api.CreateSlide)return {ok:false,error:"Adding slides is not supported."};const current=p.GetCurrentVisibleSlide?p.GetCurrentVisibleSlide():p.GetCurrentSlide(),slide=Api.CreateSlide();if(!slide)return {ok:false,error:"A new slide could not be created."};if(current&&current.GetLayout&&slide.SetLayout){const layout=current.GetLayout();if(layout)slide.SetLayout(layout);}const added=p.AddSlide(slide);if(added===false)return {ok:false,error:"Euro-Office rejected the new slide."};function add(text,y,size){if(!Api.CreateShape||!slide.AddObject)return false;const fill=Api.CreateNoFill(),stroke=Api.CreateStroke(0,Api.CreateNoFill()),shape=Api.CreateShape("rect",8000000,size,fill,stroke);shape.SetPosition(700000,y);const c=shape.GetDocContent&&shape.GetDocContent();if(!c)return false;if(c.RemoveAllElements)c.RemoveAllElements();const lines=String(text).split(/\\r?\\n/);for(let i=0;i<lines.length;i+=1){const para=Api.CreateParagraph();para.AddText(lines[i]);if(c.Push)c.Push(para);else if(c.AddElement)c.AddElement(i,para);else return false;}slide.AddObject(shape);return true;}if(!add(input.title,500000,1000000))return {ok:false,error:"Slide text shapes are not supported."};if(!add(input.bullets.map(x=>"• "+x).join("\\n"),1700000,3500000))return {ok:false,error:"Slide body could not be created."};return {ok:true,changed:1};', slideData);
    if (!result || !result.ok) throw new Error(result && result.error || "The new slide was not created.");
    return { ok: true, changed: 1, target: "New slide", warnings: [] };
  };
  host.addNotes = async function (text, snapshot) {
    const result = await host.dynamicCommand('const p=Api.GetPresentation(),slide=p.GetCurrentVisibleSlide?p.GetCurrentVisibleSlide():p.GetCurrentSlide();if(!slide||!slide.AddNotesText)return {ok:false,error:"Speaker notes are not supported by this Euro-Office build."};slide.AddNotesText(input.text);return {ok:true,changed:1};', { text });
    if (!result || !result.ok) throw new Error(result && result.error || "Speaker notes were not added.");
    return { ok: true, changed: 1, target: snapshot.targetLabel + " notes", warnings: [] };
  };
  host.canUndo = async function () { const value = await host.execute("CanUndo", []).catch(() => false); return value === true; };
  host.undo = async function () { await host.execute("Undo", []); return true; };
  G.host = host;
})(window);
