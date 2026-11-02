(function (global) {
  "use strict";
  const G = global.Guiden,
    C = G.core;
  const host = {};
  // Runs inside the editor (it is sent with the commands that use it): applies each item's style to its cell.
  function styleCells(sheet, cells) {
    function color(hex) {
      return Api.CreateColorFromRGB(
        parseInt(hex.slice(1, 3), 16),
        parseInt(hex.slice(3, 5), 16),
        parseInt(hex.slice(5, 7), 16)
      );
    }
    let changed = 0;
    for (let i = 0; i < cells.length; i += 1) {
      const item = cells[i],
        r = sheet && sheet.GetRange(item.address),
        s = item.style;
      if (!r) return { ok: false, changed: changed, error: "A formatting target is unavailable." };
      let supported = false;
      if (s.fill && r.SetFillColor) {
        r.SetFillColor(color(s.fill));
        supported = true;
      }
      if (s.fontColor && r.SetFontColor) {
        r.SetFontColor(color(s.fontColor));
        supported = true;
      }
      if (typeof s.bold === "boolean" && r.SetBold) {
        r.SetBold(s.bold);
        supported = true;
      }
      if (typeof s.italic === "boolean" && r.SetItalic) {
        r.SetItalic(s.italic);
        supported = true;
      }
      if (s.align && r.SetAlignHorizontal) {
        r.SetAlignHorizontal(s.align);
        supported = true;
      }
      if (s.numberFormat && r.SetNumberFormat) {
        r.SetNumberFormat(s.numberFormat);
        supported = true;
      }
      if (!supported)
        return { ok: false, changed: changed, error: "Formatting is not supported by this Euro-Office build." };
      changed += 1;
    }
    return { ok: true, changed: changed };
  }

  // Runs inside the editor as well as in the panel: the smallest word-level edits that turn before into after, as
  // [start, end, text] spans of before (replace before.slice(start, end) with text), in order. A single space between
  // two changes joins them, so a rewritten phrase reads as one change. null when the texts are too long to compare.
  function wordEdits(before, after) {
    const a = String(before).match(/\s+|[^\s]+/g) || [],
      b = String(after).match(/\s+|[^\s]+/g) || [];
    if ((a.length + 1) * (b.length + 1) > 400000) return null;
    const n = a.length,
      m = b.length,
      width = m + 1,
      common = new Uint32Array((n + 1) * width);
    for (let i = n - 1; i >= 0; i -= 1)
      for (let j = m - 1; j >= 0; j -= 1)
        common[i * width + j] =
          a[i] === b[j]
            ? common[(i + 1) * width + j + 1] + 1
            : Math.max(common[(i + 1) * width + j], common[i * width + j + 1]);
    const steps = [];
    let i = 0,
      j = 0;
    while (i < n || j < m) {
      if (i < n && j < m && a[i] === b[j]) {
        steps.push(["=", a[i]]);
        i += 1;
        j += 1;
      } else if (j < m && (i === n || common[i * width + j + 1] >= common[(i + 1) * width + j])) {
        steps.push(["+", b[j]]);
        j += 1;
      } else {
        steps.push(["-", a[i]]);
        i += 1;
      }
    }
    const edits = [];
    let offset = 0,
      current = null;
    for (let k = 0; k < steps.length; k += 1) {
      const kind = steps[k][0],
        token = steps[k][1];
      const joins = kind === "=" && current && /^\s+$/.test(token) && k + 1 < steps.length && steps[k + 1][0] !== "=";
      if (kind === "=" && !joins) {
        if (current) edits.push([current.start, current.end, current.text]);
        current = null;
        offset += token.length;
        continue;
      }
      if (!current) current = { start: offset, end: offset, text: "" };
      if (kind !== "+") {
        offset += token.length;
        current.end = offset;
      }
      if (kind !== "-") current.text += token;
    }
    if (current) edits.push([current.start, current.end, current.text]);
    return edits;
  }
  host.wordEdits = wordEdits;
  host.trackedEdits = trackedEdits;

  // Runs inside the editor: makes the selection read input.after, one line per selected paragraph, as word-level
  // edits. Used with Track changes on, so each edit shows as a deletion and an insertion someone can accept or reject.
  // Nothing changes unless every paragraph still reads as input.before.
  function trackedEdits(input) {
    const doc = Api.GetDocument();
    if (!doc || typeof doc.GetRangeBySelect !== "function" || typeof doc.IsTrackRevisions !== "function")
      return { ok: false, unsupported: true };
    if (!doc.IsTrackRevisions()) return { ok: false, error: "Track changes is off." };
    const range = doc.GetRangeBySelect(),
      lines = input.before,
      at = [],
      plans = [];
    if (!range || typeof range.GetRange !== "function") return { ok: false, unsupported: true };
    function strip(text) {
      return String(text || "").replace(/[\r\n]+$/, "");
    }
    if (lines.length === 1) {
      if (strip(range.GetText()) !== lines[0]) return { ok: false, error: "The selection changed." };
      at.push(function (start, end) {
        return range.GetRange(start, end);
      });
    } else {
      const paragraphs = typeof range.GetAllParagraphs === "function" ? range.GetAllParagraphs() : [];
      if (paragraphs.length !== lines.length) return { ok: false, error: "The selection changed." };
      for (let i = 0; i < lines.length; i += 1) {
        const paragraph = paragraphs[i],
          text = strip(paragraph.GetText()),
          offset = i === 0 ? text.length - lines[i].length : 0;
        if (offset < 0 || text.substr(offset, lines[i].length) !== lines[i])
          return { ok: false, error: "The selection changed." };
        at.push(function (start, end) {
          return paragraph.GetRange(offset + start, offset + end);
        });
      }
    }
    for (let i = 0; i < lines.length; i += 1) {
      const edits = wordEdits(lines[i], input.after[i]);
      if (!edits) return { ok: false, error: "The text is too long to compare." };
      plans.push(edits);
    }
    // Editor positions are not character offsets: each run starts with an empty position. So map every character to
    // its position first, then edit from the end backwards, deleting the old words and then adding the new ones after
    // the character in front of them. Each step changes only positions after the ones it uses.
    const maps = [];
    for (let i = 0; i < lines.length; i += 1) {
      const map = [];
      let position = 0;
      while (map.length < lines[i].length && position < lines[i].length * 3 + 50) {
        const piece = at[i](position, position + 1),
          text = piece ? String(piece.GetText() || "") : "";
        for (let c = 0; c < text.length; c += 1) map.push(position);
        position += 1;
      }
      if (map.length !== lines[i].length) return { ok: false, error: "The selection could not be measured." };
      maps.push(map);
    }
    let changed = 0;
    for (let i = plans.length - 1; i >= 0; i -= 1)
      for (let k = plans[i].length - 1; k >= 0; k -= 1) {
        const start = plans[i][k][0],
          end = plans[i][k][1],
          text = plans[i][k][2],
          map = maps[i];
        if (end > start) at[i](map[start], map[end - 1] + 1).Delete();
        if (text) {
          if (start > 0) at[i](map[start - 1], map[start - 1] + 1).AddText(text, "after");
          else if (map.length) at[i](map[0], map[0] + 1).AddText(text, "before");
          else return { ok: false, error: "An empty paragraph cannot be edited this way." };
        }
        changed += 1;
      }
    return { ok: true, changed: changed };
  }

  // One editor call at a time. The panel's calls travel through the background frame and their answers are matched
  // to callers in order, and the runtime keeps a single callback for commands, so overlapping calls can get each
  // other's answers. A call the editor never answers fails after callTimeoutMs instead of holding the queue.
  host.callTimeoutMs = 60000;
  let queue = Promise.resolve();
  function serial(call) {
    const task = () =>
      new Promise((resolve, reject) => {
        const timer = setTimeout(
          () => reject(new Error("Euro-Office did not answer in time. Try again.")),
          host.callTimeoutMs
        );
        const done = value => {
          clearTimeout(timer);
          resolve(value);
        };
        try {
          call(done);
        } catch (e) {
          clearTimeout(timer);
          reject(e);
        }
      });
    const result = queue.then(task, task);
    queue = result.catch(() => {});
    return result;
  }
  host.execute = function (name, args) {
    return serial(done => global.Asc.plugin.executeMethod(name, args || [], done));
  };
  host.command = function (fn) {
    return serial(done => global.Asc.plugin.callCommand(fn, false, true, done));
  };
  // Runs fn(input) inside the editor. The editor receives the command as source text, so fn can't use anything from
  // this file: input travels as JSON written into the source, and helpers are functions sent along with it.
  host.dynamicCommand = function (fn, data, helpers) {
    const source =
      "return function () {" +
      (helpers || []).map(String).join("\n") +
      "\nconst input = " +
      JSON.stringify(data || {}) +
      ";\nreturn (" +
      String(fn) +
      ")(input);\n}";
    return host.command(new Function(source)());
  };
  host.editorType = async function () {
    const info = global.Asc && global.Asc.plugin && global.Asc.plugin.info;
    const type = info && (info.editorType || info.editorTypeName || info.type);
    if (["word", "cell", "slide"].includes(type)) return type;
    return host
      .command(function () {
        try {
          if (Api.GetActiveSheet && Api.GetActiveSheet()) return "cell";
        } catch (_) {}
        try {
          if (Api.GetPresentation && Api.GetPresentation()) return "slide";
        } catch (_) {}
        return "word";
      })
      .catch(() => "word");
  };
  const documentText = function () {
    const doc = Api.GetDocument(),
      options = {
        Numbering: true,
        Math: true,
        TableCellSeparator: "\t",
        TableRowSeparator: "\n",
        ParaSeparator: "\n\n",
        NewLineSeparator: "\n"
      };
    if (doc && typeof doc.GetText === "function") return doc.GetText(options) || "";
    const out = [],
      paragraphs = doc && doc.GetAllParagraphs ? doc.GetAllParagraphs() : [];
    for (let i = 0; i < paragraphs.length; i += 1)
      if (paragraphs[i] && paragraphs[i].GetText) out.push(paragraphs[i].GetText(options) || "");
    return out.join("\n\n");
  };
  host.readWriter = async function (actionId) {
    let selected = await host.execute("GetSelectedText", [{}]).catch(() => "");
    selected = selected || "";
    const selectedHtml = selected.trim() ? await host.execute("GetSelectedHtml", []).catch(() => "") : "";
    const insertsAtCursor = actionId === "writer.draft";
    let target = selected.trim() ? "selection" : insertsAtCursor ? "cursor" : "full_document",
      text = selected;
    if (target === "full_document") text = await host.command(documentText).catch(() => "");
    // A draft at the cursor reads the document too, so new content fits its subject and wording. It is context only:
    // it is not part of the fingerprint, so typing elsewhere in the document does not block adding the draft.
    const context =
      target === "cursor"
        ? String((await host.command(documentText).catch(() => "")) || "").slice(0, C.contextLimits.draftContextChars)
        : "";
    const source = { target, text: text || "", html: selectedHtml || "" };
    return {
      editor: "word",
      actionId,
      target,
      targetLabel: target === "selection" ? "Selected text" : target === "cursor" ? "Current cursor" : "Whole document",
      text: source.text,
      contextText: context.trim() ? context : "",
      sourceHtml: source.html,
      words: C.wordCount(source.text),
      fingerprint: C.fingerprint(source)
    };
  };
  host.readCells = async function (actionId) {
    const result = await host.command(function () {
      const sheet = Api.GetActiveSheet(),
        range = Api.GetSelection ? Api.GetSelection() : sheet.GetSelection();
      if (!range) return { ok: false, error: "Unable to read the selected cells." };
      return {
        ok: true,
        sheet: sheet.GetName ? sheet.GetName() : "Sheet",
        address: range.GetAddress ? range.GetAddress(true, true, "xlA1", false) : "",
        values: range.GetValue ? range.GetValue() : "",
        formulas: range.GetFormula ? range.GetFormula() : "",
        rows: range.GetRowsCount ? range.GetRowsCount() : 0,
        cols: range.GetColumnsCount ? range.GetColumnsCount() : 0
      };
    });
    if (!result || !result.ok) throw new Error((result && result.error) || "Unable to read the selected cells.");
    const values = C.normalizeMatrix(result.values),
      // GetFormula returns a plain cell's value too; only text that starts with "=" is a formula.
      formulas = C.normalizeMatrix(result.formulas).map(row => row.map(f => (C.isFormula(f) ? f : ""))),
      address = String(result.address || "").replace(/\$/g, ""),
      parsed = C.parseA1Address(address),
      d = C.matrixDimensions(values);
    const snapshot = {
      editor: "cell",
      actionId,
      sheetName: result.sheet || "Sheet",
      address: address || "selected range",
      start: parsed,
      rows: Number(result.rows) || (parsed && parsed.rows) || d.rows,
      cols: Number(result.cols) || (parsed && parsed.cols) || d.cols,
      values,
      formulas
    };
    snapshot.target = "selection";
    snapshot.targetLabel = snapshot.sheetName + " · " + snapshot.address;
    snapshot.text = values.map(r => r.join("\t")).join("\n");
    snapshot.fingerprint = C.fingerprint({
      sheetName: snapshot.sheetName,
      address: snapshot.address,
      rows: snapshot.rows,
      cols: snapshot.cols,
      values,
      formulas
    });
    return snapshot;
  };
  host.readSlides = async function (actionId) {
    const selected = await host.execute("GetSelectedText", [{}]).catch(() => "");
    const result = await host.command(function () {
      const presentation = Api.GetPresentation();
      const slide = presentation.GetCurrentVisibleSlide
        ? presentation.GetCurrentVisibleSlide()
        : presentation.GetCurrentSlide();
      const all = presentation.GetAllSlides ? presentation.GetAllSlides() : [];
      let slideIndex = 0;
      for (let i = 0; i < all.length; i += 1) if (all[i] === slide) slideIndex = i;
      const drawings = slide && slide.GetAllDrawings ? slide.GetAllDrawings() : [],
        shapes = [];
      for (let i = 0; i < drawings.length; i += 1) {
        const drawing = drawings[i];
        if (!drawing || !drawing.GetDocContent) continue;
        const content = drawing.GetDocContent();
        let text = "";
        if (content && content.GetText) text = content.GetText({ ParaSeparator: "\n", NewLineSeparator: "\n" }) || "";
        else if (content && content.GetElementsCount && content.GetElement) {
          const lines = [];
          for (let j = 0; j < content.GetElementsCount(); j += 1) {
            const p = content.GetElement(j);
            if (p && p.GetText) lines.push(p.GetText());
          }
          text = lines.join("\n");
        }
        shapes.push({ index: i, text: text });
      }
      return { ok: true, slideIndex: slideIndex, shapes: shapes };
    });
    if (!result || !result.ok) throw new Error("Unable to read the current slide.");
    const selectedPreferred = actionId === "slides.improve" || String(actionId || "").indexOf("knowledge.") === 0;
    const target = selectedPreferred && selected && String(selected).trim() ? "selected_text" : "current_slide";
    const source = { slideIndex: result.slideIndex, target, selected: selected || "", shapes: result.shapes || [] };
    return {
      editor: "slide",
      actionId,
      target,
      targetLabel: target === "selected_text" ? "Selected slide text" : "Slide " + (result.slideIndex + 1),
      slideIndex: result.slideIndex,
      selectedText: selected || "",
      shapes: result.shapes || [],
      text:
        target === "selected_text"
          ? selected
          : (result.shapes || []).map(s => "[" + s.index + "] " + s.text).join("\n\n"),
      fingerprint: C.fingerprint(source)
    };
  };
  host.readSnapshot = async function (editor, actionId) {
    return editor === "cell"
      ? host.readCells(actionId)
      : editor === "slide"
        ? host.readSlides(actionId)
        : host.readWriter(actionId);
  };
  host.revalidate = async function (snapshot) {
    const fresh = await host.readSnapshot(snapshot.editor, snapshot.actionId);
    const same =
      fresh.fingerprint === snapshot.fingerprint &&
      fresh.target === snapshot.target &&
      (snapshot.editor !== "cell" || fresh.address === snapshot.address) &&
      (snapshot.editor !== "slide" || fresh.slideIndex === snapshot.slideIndex);
    return { ok: same, fresh };
  };
  // Builds Writer paragraphs from payload blocks and puts them in the document. Runs carry no font of their own, so the
  // text takes the document's formatting. mode is:
  // - "inline": a single paragraph's runs go into the paragraph being edited, so it keeps that paragraph's style and
  //   leaves no empty paragraph behind;
  // - "replace": the paragraphs replace the selection;
  // - "after": a draft at the cursor. The paragraphs go after the paragraph holding the cursor (or replace it when it
  //   is empty), so they keep their own styles and never merge into the text around the cursor, and the cursor ends
  //   after the new text.
  function insertNative(payload, mode) {
    return host
      .dynamicCommand(
        function (input) {
          const doc = Api.GetDocument();
          if (
            !doc ||
            typeof doc.InsertContent !== "function" ||
            typeof Api.CreateParagraph !== "function" ||
            typeof Api.CreateHyperlink !== "function"
          )
            return { ok: false, unsupported: true };
          // One real list per <ul>/<ol>; null when this build cannot make one, and the item then starts with its marker.
          const numberings = {};
          function numberingFor(list) {
            if (!(list.id in numberings)) {
              let level = null;
              try {
                const numbering = typeof doc.CreateNumbering === "function" ? doc.CreateNumbering(list.kind) : null;
                level = numbering && typeof numbering.GetLevel === "function" ? numbering.GetLevel(0) : null;
              } catch (_) {}
              numberings[list.id] = level;
            }
            return numberings[list.id];
          }
          const content = [];
          for (let i = 0; i < input.blocks.length; i += 1) {
            const block = input.blocks[i],
              para = Api.CreateParagraph(),
              segments = block.segments;
            // SetStyle takes a style object, not a name; a name the document lacks leaves the paragraph as it is.
            if (block.paragraphStyle && typeof doc.GetStyle === "function" && typeof para.SetStyle === "function") {
              try {
                const style = doc.GetStyle(block.paragraphStyle);
                if (style) para.SetStyle(style);
              } catch (_) {}
            }
            if (block.list) {
              const level = numberingFor(block.list);
              let numbered = false;
              if (level && typeof para.SetNumbering === "function") {
                try {
                  numbered = para.SetNumbering(level) !== false;
                } catch (_) {}
              }
              if (!numbered) {
                const marker = Api.CreateRun();
                marker.AddText(block.list.marker);
                para.Push(marker);
              }
            }
            for (let j = 0; j < segments.length; j += 1) {
              const segment = segments[j],
                text = String(segment.text || "");
              if (!text) continue;
              if (segment.href) {
                const link = Api.CreateHyperlink(segment.href, text, "");
                if (!link) return { ok: false, error: "A native hyperlink could not be created." };
                const added = typeof para.AddElement === "function" ? para.AddElement(link) : para.Push(link);
                if (added === false) return { ok: false, error: "A native hyperlink could not be inserted." };
                continue;
              }
              const parts = text.split(String.fromCharCode(10));
              for (let k = 0; k < parts.length; k += 1) {
                const run = Api.CreateRun();
                if (parts[k]) run.AddText(parts[k]);
                if (segment.bold && typeof run.SetBold === "function") run.SetBold(true);
                if (segment.italic && typeof run.SetItalic === "function") run.SetItalic(true);
                // A [placeholder] Guiden could not fill in is highlighted, so it is easy to find and replace.
                if (segment.highlight && parts[k] && typeof run.SetHighlight === "function") run.SetHighlight("yellow");
                if (k < parts.length - 1 && typeof run.AddLineBreak === "function") run.AddLineBreak();
                if (para.Push(run) === false) return { ok: false, error: "Writer text could not be inserted." };
              }
            }
            content.push(para);
          }
          let current = null;
          if (input.mode === "after" && typeof doc.GetCurrentParagraph === "function") {
            try {
              current = doc.GetCurrentParagraph();
            } catch (_) {}
          }
          if (current && typeof current.InsertParagraph === "function") {
            const empty = !String(current.GetText() || "").replace(/[\r\n]+$/, "");
            let at = current;
            for (let i = 0; i < content.length; i += 1) {
              const next = at.InsertParagraph(content[i], "after", true);
              if (next === false) return { ok: false, error: "Writer text could not be inserted." };
              at = next || content[i];
            }
            if (empty && typeof current.Delete === "function") current.Delete();
            try {
              const count = at.GetElementsCount(),
                last = count ? at.GetElement(count - 1) : null;
              if (last && typeof last.MoveCursorToPos === "function") last.MoveCursorToPos(last.GetText().length);
            } catch (_) {}
            return { ok: true, changed: 1, links: input.linkCount };
          }
          const ok = doc.InsertContent(content, input.mode === "inline");
          return { ok: ok !== false, changed: ok === false ? 0 : 1, links: input.linkCount };
        },
        { blocks: payload.blocks, linkCount: payload.linkCount, mode: mode || "replace" }
      )
      .catch(error => ({ ok: false, error: (error && error.message) || String(error) }));
  }
  function blockText(block) {
    return block.segments.map(segment => String(segment.text || "")).join("");
  }
  // Paragraphs in a Writer selection, as GetSelectedText returns it (one per line, usually with a final line break).
  function lines(text) {
    const all = String(text || "")
      .replace(/\r\n?/g, "\n")
      .replace(/\n+$/, "")
      .split("\n");
    return all.length === 1 && !all[0] ? [] : all;
  }
  host.paragraphCount = function (text) {
    return lines(text).length;
  };

  // Puts Guiden's HTML into the document in the way that best keeps the document's own look:
  // - an edit of several selected paragraphs that returns the same number of plain paragraphs (Improve, Fix,
  //   Translate) replaces each paragraph's text in place, so headings, fonts and spacing stay as they were;
  // - anything else without tables (headings, paragraphs, real bullet and numbered lists, links) is built natively in
  //   the document's formatting, and a draft at the cursor becomes its own paragraphs after the cursor's paragraph;
  // - tables go through the editor's HTML importer, which sets its own font.
  // With tracked=true (Track changes is on), an edit that keeps the paragraphs as they are is made as word-level
  // tracked edits, so reviewers see only the words that changed; anything else shows as the old text deleted and the
  // new text inserted.
  host.pasteWriter = async function (html, snapshot, tracked) {
    const payload = G.ui && typeof G.ui.writerNativePayload === "function" ? G.ui.writerNativePayload(html) : null;
    if (payload && payload.supported && payload.blocks.length) {
      const selected = snapshot && snapshot.target !== "cursor" ? host.paragraphCount(snapshot.text) : 0;
      const sameShape = selected > 0 && payload.linkCount === 0 && payload.blocks.length === selected;
      if (tracked && sameShape) {
        const before = lines(snapshot.text).slice(0, selected);
        const edited = await host
          .dynamicCommand(trackedEdits, { before, after: payload.blocks.map(blockText) }, [wordEdits])
          .catch(error => ({ ok: false, error: error && error.message }));
        if (edited && edited.ok) return { ok: true, changed: 1, links: 0, warnings: [], tracked: true };
        if (global.console && console.warn) console.warn("[Guiden] Word-level tracked edit unavailable.", edited);
      }
      if (!tracked && selected > 1 && sameShape) {
        const replaced = await host.execute("ReplaceTextSmart", [payload.blocks.map(blockText)]).catch(() => false);
        if (replaced === true) return { ok: true, changed: 1, links: 0, warnings: [] };
      }
      const single = payload.blocks.length === 1 && !payload.blocks[0].paragraphStyle;
      const mode = snapshot && snapshot.target === "cursor" ? "after" : single ? "inline" : "replace";
      const result = await insertNative(payload, mode);
      if (result && result.ok && result.links === payload.linkCount)
        return { ok: true, changed: 1, links: payload.linkCount, warnings: [] };
      if (global.console && console.warn)
        console.warn("[Guiden] Native Writer insert unavailable; using HTML import.", result || { stage: "no-result" });
      await host.execute("PasteHtml", [String(html || "")]);
      return {
        ok: true,
        changed: 1,
        links: payload.linkCount,
        warnings: payload.linkCount ? ["Euro-Office used its HTML compatibility importer for document links."] : []
      };
    }
    await host.execute("PasteHtml", [String(html || "")]);
    return { ok: true, changed: 1, links: 0, warnings: [] };
  };

  host.writerHtmlDocument = function (html) {
    const body = String(html || "")
      .trim()
      .replace(/\s+(?:target|rel)="[^"]*"/gi, "");
    return '<!doctype html><html><head><meta charset="utf-8"></head><body>' + body + "</body></html>";
  };
  host.applyWriter = async function (html, snapshot, options) {
    if (snapshot.target === "full_document" && snapshot.text.trim()) {
      let selected = false;
      try {
        await host.execute("SelectAll", []);
        await new Promise(resolve => global.setTimeout(resolve, 120));
        selected = Boolean(String(await host.execute("GetSelectedText", [{}]).catch(() => "")).trim());
      } catch (_) {}
      if (!selected) throw new Error("This Euro-Office build cannot safely select the whole document.");
    }
    const tracking = await host.trackChanges(Boolean(options && options.track));
    let pasted;
    try {
      pasted = await host.pasteWriter(html, snapshot, tracking.on);
    } finally {
      if (tracking.turnedOn) await host.trackChanges(false, true);
    }
    return {
      ok: true,
      changed: 1,
      target: snapshot.targetLabel,
      tracked: tracking.on,
      warnings: pasted.warnings || []
    };
  };
  // Whether Track changes is on, after turning it on when on=true. turnedOn says this call switched it on, so the
  // caller can switch it back off with force=true; Guiden never turns off tracking someone else turned on.
  host.trackChanges = async function (on, force) {
    const result = await host
      .dynamicCommand(
        function (input) {
          const doc = Api.GetDocument();
          if (!doc || typeof doc.IsTrackRevisions !== "function" || typeof doc.SetTrackRevisions !== "function")
            return { supported: false, on: false, turnedOn: false };
          const was = Boolean(doc.IsTrackRevisions());
          if (input.force) doc.SetTrackRevisions(input.on);
          else if (input.on && !was) doc.SetTrackRevisions(true);
          const now = Boolean(doc.IsTrackRevisions());
          return { supported: true, on: now, turnedOn: !was && now };
        },
        { on: Boolean(on), force: Boolean(force) }
      )
      .catch(() => null);
    return result || { supported: false, on: false, turnedOn: false };
  };
  // Inserts at the cursor, or in place of the selection when there is one.
  // Adds each comment to the first unused place its quote appears, inside the selection when there is one. Comments
  // whose text has gone are skipped and counted.
  host.addComments = async function (comments, snapshot) {
    const result = await host.dynamicCommand(
      function (input) {
        const doc = Api.GetDocument();
        if (!doc || typeof doc.Search !== "function") return { ok: false, error: "Comments are not supported here." };
        let scope = doc;
        if (input.selection && typeof doc.GetRangeBySelect === "function") {
          const range = doc.GetRangeBySelect();
          if (range && typeof range.Search === "function") scope = range;
        }
        const used = {};
        let added = 0,
          missing = 0;
        for (let i = 0; i < input.comments.length; i += 1) {
          const item = input.comments[i],
            found = scope.Search(item.quote, true) || [],
            index = used[item.quote] || 0,
            range = found[index] || found[0];
          if (!range || typeof range.AddComment !== "function") {
            missing += 1;
            continue;
          }
          used[item.quote] = index + 1;
          if (range.AddComment(item.comment, input.author)) added += 1;
          else missing += 1;
        }
        return { ok: added > 0, changed: added, missing: missing, error: added ? "" : "No comment could be added." };
      },
      { comments, selection: snapshot && snapshot.target === "selection", author: "Guiden" }
    );
    return Object.assign({ target: snapshot ? snapshot.targetLabel : "document" }, result || { ok: false });
  };
  host.insertWriter = async function (html) {
    const selected = String((await host.execute("GetSelectedText", [{}]).catch(() => "")) || "").trim();
    const pasted = await host.pasteWriter(html, { target: selected ? "selection" : "cursor", text: "" });
    return { ok: true, changed: 1, target: "current cursor or selection", warnings: pasted.warnings || [] };
  };
  host.applyCellUpdates = async function (updates, snapshot) {
    const cells = updates.map(u => ({
      address: C.rangeAddress(
        { startRow: snapshot.start.startRow + u.row, startCol: snapshot.start.startCol + u.column },
        1,
        1
      ),
      value: u.after
    }));
    const result = await host.dynamicCommand(
      function (input) {
        const sheet = Api.GetActiveSheet();
        let changed = 0;
        for (let i = 0; i < input.cells.length; i += 1) {
          const item = input.cells[i],
            range = sheet && sheet.GetRange(item.address);
          if (!range || typeof range.SetValue !== "function")
            return { ok: false, changed: changed, error: "A target cell cannot be edited." };
          const ok = range.SetValue(item.value);
          if (ok === false) return { ok: false, changed: changed, error: "Euro-Office rejected a cell change." };
          changed += 1;
        }
        return { ok: true, changed: changed };
      },
      { cells }
    );
    if (!result || !result.ok || result.changed !== cells.length)
      throw new Error((result && result.error) || "Only part of the cell update was applied.");
    return { ok: true, changed: result.changed, target: snapshot.address, warnings: [] };
  };
  host.applyFormula = async function (targetAddress, formula) {
    const result = await host.dynamicCommand(
      function (input) {
        const sheet = Api.GetActiveSheet(),
          range = sheet && sheet.GetRange(input.address);
        if (!range) return { ok: false, error: "Formula destination is unavailable." };
        let ok;
        if (typeof range.SetFormula === "function") ok = range.SetFormula(input.formula);
        else if (typeof range.SetValue === "function") ok = range.SetValue(input.formula);
        else return { ok: false, error: "Formulas are not supported." };
        return { ok: ok !== false, changed: ok === false ? 0 : 1 };
      },
      { address: targetAddress, formula }
    );
    if (!result || !result.ok || result.changed !== 1)
      throw new Error((result && result.error) || "The formula was not applied.");
    return { ok: true, changed: 1, target: targetAddress, warnings: [] };
  };
  host.applyStyles = async function (cells, snapshot) {
    const payload = cells.map(item => ({
      address: C.rangeAddress(
        { startRow: snapshot.start.startRow + item.row, startCol: snapshot.start.startCol + item.column },
        1,
        1
      ),
      style: item.style
    }));
    const result = await host.dynamicCommand(
      function (input) {
        const sheet = Api.GetActiveSheet();
        const styled = styleCells(sheet, input.cells);
        if (!styled.ok) return styled;
        return { ok: true, changed: styled.changed };
      },
      { cells: payload },
      [styleCells]
    );
    if (!result || !result.ok || result.changed !== payload.length)
      throw new Error((result && result.error) || "Only part of the formatting was applied.");
    return { ok: true, changed: result.changed, target: snapshot.address, warnings: [] };
  };
  // Writes values (or formulas, when an item has formula) to single cells in one editor command, so a single Undo
  // reverts the whole change.
  host.applyCells = async function (cells, target) {
    const result = await host.dynamicCommand(
      function (input) {
        const sheet = Api.GetActiveSheet();
        let changed = 0;
        for (let i = 0; i < input.cells.length; i += 1) {
          const item = input.cells[i],
            range = sheet && sheet.GetRange(item.address);
          if (!range) return { ok: false, changed: changed, error: "A target cell cannot be edited." };
          let ok;
          if (typeof item.formula === "string") {
            if (typeof range.SetFormula === "function") ok = range.SetFormula(item.formula);
            else if (typeof range.SetValue === "function") ok = range.SetValue(item.formula);
            else return { ok: false, changed: changed, error: "Formulas are not supported." };
          } else {
            if (typeof range.SetValue !== "function")
              return { ok: false, changed: changed, error: "A target cell cannot be edited." };
            ok = range.SetValue(item.value);
          }
          if (ok === false) return { ok: false, changed: changed, error: "Euro-Office rejected a cell change." };
          changed += 1;
        }
        return { ok: true, changed: changed };
      },
      { cells }
    );
    if (!result || !result.ok || result.changed !== cells.length)
      throw new Error((result && result.error) || "Only part of the change was applied.");
    return { ok: true, changed: result.changed, target, warnings: [] };
  };
  host.readRange = async function (address) {
    return host.dynamicCommand(
      function (input) {
        const s = Api.GetActiveSheet(),
          r = s && s.GetRange(input.address);
        if (!r) return { ok: false };
        return { ok: true, values: r.GetValue ? r.GetValue() : "", formulas: r.GetFormula ? r.GetFormula() : "" };
      },
      { address }
    );
  };
  host.applyTable = async function (table, address, styles) {
    const start = C.parseA1Address(address),
      cells = (styles || []).map(item => ({
        address: C.rangeAddress({ startRow: start.startRow + item.row, startCol: start.startCol + item.column }, 1, 1),
        style: item.style
      }));
    // Values and header styling go in one command, so a single Undo removes the whole table.
    const result = await host.dynamicCommand(
      function (input) {
        const sheet = Api.GetActiveSheet(),
          r = sheet && sheet.GetRange(input.address);
        if (!r || !r.SetValue) return { ok: false, error: "Table destination cannot be edited." };
        const ok = r.SetValue(input.table);
        if (ok === false) return { ok: false, error: "Euro-Office rejected the table." };
        const styled = styleCells(sheet, input.cells);
        if (!styled.ok) return { ok: false, error: styled.error };
        return { ok: true, changed: input.table.length * input.table[0].length };
      },
      { table, address, cells },
      [styleCells]
    );
    if (!result || !result.ok) throw new Error((result && result.error) || "The table was not applied.");
    return { ok: true, changed: result.changed, target: address, warnings: [] };
  };
  host.applySelectedSlideText = async function (text, snapshot) {
    await host.execute("PasteText", [text]);
    return { ok: true, changed: 1, target: snapshot.targetLabel, warnings: [] };
  };
  host.applyShapeUpdates = async function (updates, snapshot) {
    const result = await host.dynamicCommand(
      function (input) {
        const p = Api.GetPresentation(),
          slide = p.GetCurrentVisibleSlide ? p.GetCurrentVisibleSlide() : p.GetCurrentSlide(),
          drawings = slide && slide.GetAllDrawings ? slide.GetAllDrawings() : [];
        let changed = 0;
        for (let i = 0; i < input.updates.length; i += 1) {
          const u = input.updates[i],
            d = drawings[u.shapeIndex],
            c = d && d.GetDocContent ? d.GetDocContent() : null;
          if (!c || !c.RemoveAllElements)
            return { ok: false, changed: changed, error: "The intended text area cannot be edited." };
          c.RemoveAllElements();
          const lines = String(u.after).split(/\r?\n/);
          for (let j = 0; j < lines.length; j += 1) {
            const para = Api.CreateParagraph();
            para.AddText(lines[j]);
            if (c.Push) c.Push(para);
            else if (c.AddElement) c.AddElement(j, para);
            else return { ok: false, changed: changed, error: "Text cannot be written to this area." };
          }
          changed += 1;
        }
        return { ok: true, changed: changed };
      },
      { updates }
    );
    if (!result || !result.ok || result.changed !== updates.length)
      throw new Error((result && result.error) || "Only part of the slide was updated.");
    return { ok: true, changed: result.changed, target: snapshot.targetLabel, warnings: [] };
  };
  host.createSlide = async function (slideData) {
    const result = await host.dynamicCommand(function (input) {
      const p = Api.GetPresentation();
      if (!p || !p.AddSlide || !Api.CreateSlide) return { ok: false, error: "Adding slides is not supported." };
      const current = p.GetCurrentVisibleSlide ? p.GetCurrentVisibleSlide() : p.GetCurrentSlide(),
        slide = Api.CreateSlide();
      if (!slide) return { ok: false, error: "A new slide could not be created." };
      if (current && current.GetLayout && slide.SetLayout) {
        const layout = current.GetLayout();
        if (layout) slide.SetLayout(layout);
      }
      const added = p.AddSlide(slide);
      if (added === false) return { ok: false, error: "Euro-Office rejected the new slide." };
      function add(text, y, size) {
        if (!Api.CreateShape || !slide.AddObject) return false;
        const fill = Api.CreateNoFill(),
          stroke = Api.CreateStroke(0, Api.CreateNoFill()),
          shape = Api.CreateShape("rect", 8000000, size, fill, stroke);
        shape.SetPosition(700000, y);
        const c = shape.GetDocContent && shape.GetDocContent();
        if (!c) return false;
        if (c.RemoveAllElements) c.RemoveAllElements();
        const lines = String(text).split(/\r?\n/);
        for (let i = 0; i < lines.length; i += 1) {
          const para = Api.CreateParagraph();
          para.AddText(lines[i]);
          if (c.Push) c.Push(para);
          else if (c.AddElement) c.AddElement(i, para);
          else return false;
        }
        slide.AddObject(shape);
        return true;
      }
      if (!add(input.title, 500000, 1000000)) return { ok: false, error: "Slide text shapes are not supported." };
      if (!add(input.bullets.map(x => "• " + x).join("\n"), 1700000, 3500000))
        return { ok: false, error: "Slide body could not be created." };
      return { ok: true, changed: 1 };
    }, slideData);
    if (!result || !result.ok) throw new Error((result && result.error) || "The new slide was not created.");
    return { ok: true, changed: 1, target: "New slide", warnings: [] };
  };
  host.addNotes = async function (text, snapshot) {
    const result = await host.dynamicCommand(
      function (input) {
        const p = Api.GetPresentation(),
          slide = p.GetCurrentVisibleSlide ? p.GetCurrentVisibleSlide() : p.GetCurrentSlide();
        if (!slide || !slide.AddNotesText)
          return { ok: false, error: "Speaker notes are not supported by this Euro-Office build." };
        slide.AddNotesText(input.text);
        return { ok: true, changed: 1 };
      },
      { text }
    );
    if (!result || !result.ok) throw new Error((result && result.error) || "Speaker notes were not added.");
    return { ok: true, changed: 1, target: snapshot.targetLabel + " notes", warnings: [] };
  };
  host.canUndo = async function () {
    const value = await host.execute("CanUndo", []).catch(() => false);
    return value === true;
  };
  host.undo = async function () {
    await host.execute("Undo", []);
    return true;
  };
  G.host = host;
})(window);
