(function (window, document) {
  "use strict";

  const STORAGE = {
    apiKey: "cloudcix-guiden.apiKey",
    rememberKey: "cloudcix-guiden.rememberKey",
    prompt: "cloudcix-guiden.prompt",
    format: "cloudcix-guiden.format",
    tone: "cloudcix-guiden.tone",
    length: "cloudcix-guiden.length",
    workflow: "cloudcix-guiden.workflow",
    keepCitations: "cloudcix-guiden.keepCitations"
  };

  const GUIDEN_API_URL = "https://inference.cloudcix.com/v1/guiden/chat/completions";
  const MODEL = "Mistral-Medium-3.5";
  const MAX_TOKENS = 15000;
  const CHARACTER_LIMIT = 110000;
  const SAFE_TAGS = new Set([
    "h1",
    "h2",
    "h3",
    "p",
    "strong",
    "em",
    "ul",
    "ol",
    "li",
    "blockquote",
    "a",
    "br",
    "table",
    "thead",
    "tbody",
    "tr",
    "th",
    "td"
  ]);
  const TAG_ALIASES = { b: "strong", i: "em" };
  const state = {
    controller: null,
    busy: false,
    suggestion: null,
    snapshot: null,
    adapter: null,
    saveFeedbackTimer: null
  };

  function $(id) {
    return document.getElementById(id);
  }

  function setStatus(text, mode) {
    const status = $("status");
    status.textContent = text;
    status.className = `status ${mode || ""}`.trim();
  }

  function hideRedundantPluginCloseButton() {
    let hostDocument = null;
    try {
      hostDocument = window.parent && window.parent.document;
    } catch (error) {
      return;
    }
    if (!hostDocument) {
      return;
    }

    const selector = "#panel-plugins-cloudcix-guiden .current-plugin-header .tools .plugin-close.close";

    function hideButton() {
      hostDocument.querySelectorAll(selector).forEach(function (element) {
        element.style.display = "none";
        element.setAttribute("aria-hidden", "true");
      });
    }

    hideButton();
    try {
      const observer = new window.MutationObserver(hideButton);
      observer.observe(hostDocument.body, { childList: true, subtree: true });
    } catch (error) {
      // The button is cosmetic; ignore hosts that block observation.
    }
  }

  function executeMethod(name, args) {
    return new Promise((resolve, reject) => {
      try {
        window.Asc.plugin.executeMethod(name, args || [], resolve);
      } catch (error) {
        reject(error);
      }
    });
  }

  function callCommand(command) {
    return new Promise((resolve, reject) => {
      try {
        window.Asc.plugin.callCommand(command, false, true, resolve);
      } catch (error) {
        reject(error);
      }
    });
  }

  function getRuntimeEditorType() {
    const info = window.Asc && window.Asc.plugin ? window.Asc.plugin.info : null;
    const type = info && (info.editorType || info.editorTypeName || info.type);
    if (type === "cell" || type === "slide" || type === "word") {
      return type;
    }
    return "";
  }

  async function detectEditorType() {
    const runtimeType = getRuntimeEditorType();
    if (runtimeType) {
      return runtimeType;
    }

    return callCommand(function () {
      try {
        if (typeof Api !== "undefined" && Api.GetActiveSheet && Api.GetActiveSheet()) {
          return "cell";
        }
      } catch (error) {
        // Probe the next editor type.
      }
      try {
        if (typeof Api !== "undefined" && Api.GetPresentation && Api.GetPresentation()) {
          return "slide";
        }
      } catch (error) {
        // Probe the next editor type.
      }
      return "word";
    }).catch(function () {
      return "word";
    });
  }

  function normalizeMatrix(value) {
    if (Array.isArray(value)) {
      if (Array.isArray(value[0])) {
        return value.map(function (row) {
          return row.map(function (cell) {
            return cell == null ? "" : cell;
          });
        });
      }
      return [value.map(function (cell) {
        return cell == null ? "" : cell;
      })];
    }
    return [[value == null ? "" : value]];
  }

  function matrixDimensions(matrix) {
    const rows = Math.max(matrix.length, 1);
    const cols = matrix.reduce(function (max, row) {
      return Math.max(max, Array.isArray(row) ? row.length : 1);
    }, 1);
    return { rows, cols };
  }

  function columnNameToNumber(name) {
    return String(name || "").toUpperCase().split("").reduce(function (total, char) {
      const code = char.charCodeAt(0);
      if (code < 65 || code > 90) {
        return total;
      }
      return total * 26 + code - 64;
    }, 0);
  }

  function columnNumberToName(number) {
    let value = Number(number) || 1;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name || "A";
  }

  function parseA1Address(address) {
    const value = String(address || "")
      .replace(/'/g, "")
      .split("!")
      .pop()
      .replace(/\$/g, "")
      .toUpperCase();
    const match = value.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!match) {
      return null;
    }
    const startCol = columnNameToNumber(match[1]);
    const startRow = Number(match[2]);
    const endCol = match[3] ? columnNameToNumber(match[3]) : startCol;
    const endRow = match[4] ? Number(match[4]) : startRow;
    return {
      startCol,
      startRow,
      endCol,
      endRow,
      rows: Math.max(1, endRow - startRow + 1),
      cols: Math.max(1, endCol - startCol + 1)
    };
  }

  function rangeAddressFromStart(start, rows, cols) {
    const endRow = start.startRow + rows - 1;
    const endCol = start.startCol + cols - 1;
    const first = `${columnNumberToName(start.startCol)}${start.startRow}`;
    const last = `${columnNumberToName(endCol)}${endRow}`;
    return first === last ? first : `${first}:${last}`;
  }

  function matrixToDelimitedText(matrix) {
    return normalizeMatrix(matrix).map(function (row) {
      return row.map(function (cell) {
        return String(cell == null ? "" : cell);
      }).join("\t");
    }).join("\n");
  }

  function isEmptyMatrix(matrix) {
    return normalizeMatrix(matrix).every(function (row) {
      return row.every(function (cell) {
        return String(cell == null ? "" : cell).trim() === "";
      });
    });
  }

  function isProtectedCell(value, formula) {
    if (String(formula == null ? "" : formula).trim()) {
      return true;
    }
    return typeof value === "number" || typeof value === "boolean";
  }

  async function readSelectionText() {
    const text = await executeMethod("GetSelectedText", [{}]);
    return text || "";
  }

  async function readWholeDocumentText() {
    const text = await callCommand(function () {
      const documentApi = Api.GetDocument();
      const options = {
        Numbering: true,
        Math: true,
        TableCellSeparator: "\t",
        TableRowSeparator: "\n",
        ParaSeparator: "\n\n",
        TabSymbol: "\t",
        NewLineSeparator: "\n"
      };

      if (documentApi && typeof documentApi.GetText === "function") {
        return documentApi.GetText(options) || "";
      }

      if (documentApi && typeof documentApi.GetAllParagraphs === "function") {
        const paragraphs = documentApi.GetAllParagraphs();
        const parts = [];
        for (let i = 0; i < paragraphs.length; i += 1) {
          if (paragraphs[i] && typeof paragraphs[i].GetText === "function") {
            parts.push(paragraphs[i].GetText(options) || "");
          }
        }
        return parts.join("\n\n");
      }

      return "";
    });
    return text || "";
  }

  async function detectTargetText() {
    setStatus("Checking selected text...", "loading");
    const selectionText = await readSelectionText().catch(function () {
      return "";
    });
    if (selectionText.trim()) {
      return { target: "selection", text: selectionText };
    }

    setStatus("Reading whole document...", "loading");
    const documentText = await readWholeDocumentText().catch(function () {
      return "";
    });
    return { target: "full_document", text: documentText };
  }

  function stripHtml(html) {
    const template = document.createElement("template");
    template.innerHTML = html || "";
    return template.content.textContent || "";
  }

  function wait(ms) {
    return new Promise((resolve) => {
      window.setTimeout(resolve, ms);
    });
  }

  function htmlToPlainText(html) {
    const template = document.createElement("template");
    template.innerHTML = html || "";
    const blockTags = new Set(["H1", "H2", "H3", "P", "LI", "TR", "BLOCKQUOTE", "TABLE", "UL", "OL"]);
    const parts = [];

    function append(text) {
      const value = String(text || "").replace(/\s+/g, " ").trim();
      if (value) {
        parts.push(value);
      }
    }

    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        append(node.textContent);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) {
        return;
      }
      if (node.tagName === "BR") {
        parts.push("\n");
        return;
      }
      if (node.tagName === "A") {
        append(node.textContent);
        const href = node.getAttribute("href") || "";
        if (isSafeLink(href)) {
          append("(" + href + ")");
        }
        return;
      }
      if (node.tagName === "TH" || node.tagName === "TD") {
        append(node.textContent);
        parts.push("\t");
        return;
      }
      Array.from(node.childNodes).forEach(walk);
      if (blockTags.has(node.tagName)) {
        parts.push("\n");
      }
    }

    Array.from(template.content.childNodes).forEach(walk);
    return parts.join(" ").replace(/[ \t]*\n[ \t]*/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function hasRichStructure(html) {
    return /<\s*(h1|h2|h3|ul|ol|li|table|blockquote|a)\b/i.test(html || "");
  }

  async function pasteSelectionReplacement(suggestion) {
    if (suggestion && suggestion.html && hasRichStructure(suggestion.html)) {
      await executeMethod("PasteHtml", [suggestion.html]);
      return;
    }

    const text = suggestion && suggestion.html
      ? htmlToPlainText(suggestion.html)
      : suggestion && suggestion.content
        ? suggestion.content
        : "";
    await executeMethod("PasteText", [text]);
  }

  async function pasteWholeDocumentReplacement(suggestion) {
    if (!suggestion || !suggestion.sourceEmpty) {
      const selected = await selectWholeDocument();
      if (!selected) {
        throw new Error("Unable to select the whole document in this Euro-Office build.");
      }
      await wait(120);
    }
    if (suggestion && suggestion.html) {
      await executeMethod("PasteHtml", [suggestion.html]);
      return;
    }
    await executeMethod("PasteText", [suggestion && suggestion.content ? suggestion.content : ""]);
  }

  async function selectWholeDocument() {
    async function selectionHasText() {
      const selectedText = await readSelectionText().catch(function () {
        return "";
      });
      return Boolean(selectedText.trim());
    }

    try {
      await executeMethod("SelectAll", []);
      await wait(120);
      if (await selectionHasText()) {
        return true;
      }
    } catch (error) {
      // Try macro-side selection below.
    }

    try {
      await executeMethod("SelectAll", [{}]);
      await wait(120);
      if (await selectionHasText()) {
        return true;
      }
    } catch (error) {
      // Try macro-side selection below.
    }

    try {
      await callCommand(function () {
        if (typeof Asc !== "undefined" && Asc.editor && typeof Asc.editor.asc_EditSelectAll === "function") {
          Asc.editor.asc_EditSelectAll();
          return;
        }

        const documentApi = Api.GetDocument();
        if (documentApi && typeof documentApi.SelectAll === "function") {
          documentApi.SelectAll();
        }
      });
      await wait(160);
      if (await selectionHasText()) {
        return true;
      }
    } catch (error) {
      // Selection is verified below by returning false.
    }

    return false;
  }

  function extractJsonObject(text) {
    const cleaned = cleanModelContent(text);
    try {
      return JSON.parse(cleaned);
    } catch (error) {
      const first = cleaned.indexOf("{");
      const last = cleaned.lastIndexOf("}");
      if (first !== -1 && last > first) {
        try {
          return JSON.parse(cleaned.slice(first, last + 1));
        } catch (innerError) {
          return null;
        }
      }
    }
    return null;
  }

  function normalizeFormula(text) {
    const value = String(text || "").trim();
    if (!value) {
      return "";
    }
    return value.startsWith("=") ? value : `=${value}`;
  }

  function hasIsoTimestamp(values) {
    return normalizeMatrix(values).some(function (row) {
      return row.some(function (value) {
        return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(String(value || ""));
      });
    });
  }

  function makeFormulaSafeForIsoTimestamps(formula, values) {
    if (!hasIsoTimestamp(values)) {
      return formula;
    }
    return String(formula).replace(/(\$?[A-Z]{1,3}\$?\d+)\s*-\s*(\$?[A-Z]{1,3}\$?\d+)/g, function (match, endCell, startCell) {
      function asDateTime(cell) {
        return "(DATEVALUE(LEFT(" + cell + ",10))+TIMEVALUE(MID(" + cell + ",12,8)))";
      }
      return "(" + asDateTime(endCell) + "-" + asDateTime(startCell) + ")";
    });
  }

  function parseSheetJson(content, expectedMessage) {
    const parsed = extractJsonObject(content);
    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") {
      throw new Error(expectedMessage || "Guiden® expected structured spreadsheet JSON.");
    }
    return parsed;
  }

  function ensureRectangularTable(table, message) {
    if (!Array.isArray(table) || !table.length || !Array.isArray(table[0]) || !table[0].length) {
      throw new Error(message || "Guiden® expected a rectangular table.");
    }
    const width = table[0].length;
    if (!table.every(function (row) {
      return Array.isArray(row) && row.length === width;
    })) {
      throw new Error(message || "Guiden® expected a rectangular table.");
    }
    return table.map(function (row) {
      return row.map(function (cell) {
        return cell == null ? "" : cell;
      });
    });
  }

  function renderSheetResult(parsed) {
    const parts = [];
    if (parsed.narrative) {
      parts.push(renderHtml(parsed.narrative));
    }
    if (parsed.table.length) {
      parts.push(renderMarkdownTable(parsed.table.map(function (row) {
        return `| ${row.map(function (cell) {
          return String(cell == null ? "" : cell).replace(/\|/g, "\\|");
        }).join(" | ")} |`;
      })));
    }
    return parts.length ? htmlDocument(parts.map(htmlBodyFragment).join("")) : htmlDocument("");
  }

  function sheetCommonUserPrompt(prompt, snapshot, operationName) {
    return [
      `Operation: ${operationName}`,
      `Range: ${snapshot.sheet.address}`,
      `Rows: ${snapshot.sheet.rows}`,
      `Columns: ${snapshot.sheet.cols}`,
      `Values (tab-delimited):\n${snapshot.text}`,
      `Formulas (tab-delimited, non-empty means protected):\n${matrixToDelimitedText(snapshot.sheet.formulas)}`,
      `Instruction:\n${prompt}\n${formatInstruction()}\n${$("tone").value}\n${$("length").value}`
    ].join("\n\n");
  }

  const sheetOperations = {
    analysis: {
      key: "analysis",
      label: "analyzing",
      applyLabel: "Preview only",
      previewOnly: true,
      detect(prompt) {
        return /\b(analy[sz]e|summari[sz]e|explain|review|insight|trend|compare|report|describe)\b/i.test(prompt || "");
      },
      buildMessages(prompt, snapshot) {
        return [
          {
            role: "system",
            content:
              "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. The operation is analysis. Return JSON only with keys narrative and table. narrative must be a concise string. table may be an empty array or a rectangular two-dimensional array for useful summary data. Do not return formulas. Do not include Markdown or code fences."
          },
          {
            role: "user",
            content: sheetCommonUserPrompt(prompt, snapshot, "analysis")
          }
        ];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected analysis JSON.");
        const table = parsed.table && parsed.table.length
          ? ensureRectangularTable(parsed.table, "Guiden® expected a rectangular analysis table.")
          : [];
        return {
          narrative: typeof parsed.narrative === "string" ? parsed.narrative : "",
          table
        };
      },
      formatPreview(content) {
        try {
          return renderSheetResult(this.parseOutput(content));
        } catch (error) {
          return renderHtml(content || "");
        }
      },
      async apply() {
        throw new Error("This response is preview-only.");
      }
    },
    formula: {
      key: "formula",
      label: "building a formula for",
      applyLabel: "Apply formula",
      detect(prompt) {
        return /\b(formula|function|calculate|calculation|sumproduct|sumif|sumifs|lookup|xlookup|vlookup|index|match)\b/i.test(prompt || "");
      },
      buildMessages(prompt, snapshot) {
        return [
          {
            role: "system",
            content:
              "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. The operation is formula generation. Return JSON only with keys narrative and formula. formula must be exactly one complete Euro-Office/Excel-compatible formula string beginning with =. Do not return tables, Markdown, code fences, or a prose-only answer."
          },
          {
            role: "user",
            content: sheetCommonUserPrompt(prompt, snapshot, "formula")
          }
        ];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected one formula but received analysis text.");
        if (typeof parsed.formula !== "string" || !parsed.formula.trim()) {
          throw new Error("Guiden® expected one formula but received analysis text.");
        }
        return {
          narrative: typeof parsed.narrative === "string" ? parsed.narrative : "",
          table: [[normalizeFormula(parsed.formula)]],
          formula: normalizeFormula(parsed.formula)
        };
      },
      formatPreview(content) {
        try {
          return renderSheetResult(this.parseOutput(content));
        } catch (error) {
          return renderHtml(content || "");
        }
      },
      async apply(suggestion, snapshot) {
        const parsed = this.parseOutput(suggestion.content);
        const targetAddress = snapshot.sheet.start ? rangeAddressFromStart(snapshot.sheet.start, 1, 1) : "";
        await applySheetMatrix([[parsed.formula]], targetAddress, true);
      }
    },
    classify_rows: {
      key: "classify_rows",
      label: "classifying rows in",
      applyLabel: "Add category column",
      detect(prompt) {
        return /\b(classif(?:y|ication)|categori[sz]e|tag each row)\b/i.test(prompt || "");
      },
      buildMessages(prompt, snapshot) {
        return [{
          role: "system",
          content: "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. Classify each data row in the selected table and add one category column immediately to its right. Return JSON only with narrative, header, and values. values must be an array with exactly one plain-text category per data row, excluding the header row. Do not return formulas, Markdown, or code fences."
        }, {
          role: "user",
          content: sheetCommonUserPrompt(prompt, snapshot, "classify rows")
        }];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected category-column JSON.");
        if (typeof parsed.header !== "string" || !parsed.header.trim() || !Array.isArray(parsed.values)) {
          throw new Error("Guiden® must return a category header and one value per data row.");
        }
        return { narrative: typeof parsed.narrative === "string" ? parsed.narrative : "", header: parsed.header.trim(), values: parsed.values.map(function (value) { return value == null ? "" : String(value); }), table: [[parsed.header.trim()]].concat(parsed.values.map(function (value) { return [value == null ? "" : String(value)]; })) };
      },
      formatPreview(content) {
        try { return renderSheetResult(this.parseOutput(content)); } catch (error) { return renderHtml(content || ""); }
      },
      async apply(suggestion, snapshot) {
        const parsed = this.parseOutput(suggestion.content);
        if (snapshot.sheet.rows < 2 || !snapshot.sheet.start) {
          throw new Error("Select a table with a header row and at least one data row.");
        }
        if (parsed.values.length !== snapshot.sheet.rows - 1) {
          throw new Error("Guiden® must return one category for every data row in the selection.");
        }
        const target = { startCol: snapshot.sheet.start.endCol + 1, startRow: snapshot.sheet.start.startRow };
        await applySheetMatrix([[parsed.header]].concat(parsed.values.map(function (value) { return [value]; })), rangeAddressFromStart(target, snapshot.sheet.rows, 1));
      }
    },
    add_column: {
      key: "add_column",
      label: "adding a calculated column beside",
      applyLabel: "Add column",
      detect(prompt) {
        const instruction = String(prompt || "");
        // A derived value is often requested as a calculation "per user" or
        // "for each row", without explicitly saying "add a column". Treat
        // those requests as a calculated-column operation so we create one
        // formula and fill it down instead of rewriting the whole table.
        const explicitColumn = /\b(?:add|create|insert|make|build|append|include)\b[\s\S]{0,80}\bcolumn\b|\b(?:new|additional|another|extra|calculated|derived)\s+column\b|\bcolumn\b[\s\S]{0,80}\b(?:add|create|insert|make|build|append|include)\b/i;
        const perRecordCalculation = /\b(?:calculate|compute|derive|work\s*out|determine)\b[\s\S]{0,100}\b(?:per|for\s+each)\s+(?:row|user|record|item|entry|job)\b|\b(?:cost|total|amount|rate|duration)\b[\s\S]{0,100}\b(?:per|for\s+each)\s+(?:row|user|record|item|entry|job)\b/i;
        return explicitColumn.test(instruction) || perRecordCalculation.test(instruction);
      },
      buildMessages(prompt, snapshot) {
        return [
          {
            role: "system",
            content:
              "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. The operation is adding one calculated column immediately to the right of the selected table. Return JSON only with keys narrative, header, and formula. header is the new column heading. formula is exactly one complete Euro-Office/Excel-compatible formula beginning with = for the FIRST data row below the header. Use relative row references so the editor can fill it down. If timestamp values use ISO text such as 2026-05-01T14:11:13, do not subtract those cells directly; convert each with DATEVALUE(LEFT(cell,10))+TIMEVALUE(MID(cell,12,8)) first. When the request asks for a total per user (or another grouped total), calculate across all matching rows with SUMPRODUCT using same-sized absolute ranges and the current-row criterion. Never use SUMIF or SUMIFS with a calculated expression as its sum_range: Euro-Office returns #VALUE! for that pattern. Do not generate one formula per row. Do not return Markdown, code fences, or any cells outside this one new column."
          },
          {
            role: "user",
            content: sheetCommonUserPrompt(prompt, snapshot, "add calculated column")
          }
        ];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected calculated-column JSON.");
        if (typeof parsed.header !== "string" || !parsed.header.trim() || typeof parsed.formula !== "string" || !parsed.formula.trim()) {
          throw new Error("Guiden® must return a column header and one first-row formula.");
        }
        return {
          narrative: typeof parsed.narrative === "string" ? parsed.narrative : "",
          header: parsed.header.trim(),
          formula: normalizeFormula(parsed.formula),
          table: [[parsed.header.trim()], [normalizeFormula(parsed.formula)]]
        };
      },
      formatPreview(content) {
        try {
          return renderSheetResult(this.parseOutput(content));
        } catch (error) {
          return renderHtml(content || "");
        }
      },
      async apply(suggestion, snapshot) {
        const parsed = this.parseOutput(suggestion.content);
        if (snapshot.sheet.rows < 2 || !snapshot.sheet.start) {
          throw new Error("Select a table with a header row and at least one data row.");
        }
        const target = {
          startCol: snapshot.sheet.start.endCol + 1,
          startRow: snapshot.sheet.start.startRow
        };
        const headerAddress = rangeAddressFromStart(target, 1, 1);
        const formulaStart = { startCol: target.startCol, startRow: target.startRow + 1 };
        const formulaAddress = rangeAddressFromStart(formulaStart, snapshot.sheet.rows - 1, 1);
        const formula = makeFormulaSafeForIsoTimestamps(parsed.formula, snapshot.sheet.values);
        await applyCalculatedColumn(parsed.header, formula, headerAddress, formulaAddress, formulaStart.startRow, snapshot.sheet.rows - 1, columnNumberToName(target.startCol));
      }
    },
    table: {
      key: "table",
      label: "building a table for",
      applyLabel: "Apply table",
      detect(prompt) {
        return /\b(generate|create|build|make|draft)\b/i.test(prompt || "") && /\b(table|range|spreadsheet|sheet|rows?|columns?)\b/i.test(prompt || "");
      },
      buildMessages(prompt, snapshot) {
        return [
          {
            role: "system",
            content:
              "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. The operation is table generation. Return JSON only with keys narrative and table. table must be a non-empty rectangular two-dimensional array of plain strings, numbers, or booleans. Do not return formulas unless the user explicitly asks for formula generation. Do not include Markdown or code fences."
          },
          {
            role: "user",
            content: sheetCommonUserPrompt(prompt, snapshot, "table")
          }
        ];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected table JSON.");
        return {
          narrative: typeof parsed.narrative === "string" ? parsed.narrative : "",
          table: ensureRectangularTable(parsed.table, "Guiden® expected a rectangular table.")
        };
      },
      formatPreview(content) {
        try {
          return renderSheetResult(this.parseOutput(content));
        } catch (error) {
          return renderHtml(content || "");
        }
      },
      async apply(suggestion, snapshot) {
        const parsed = this.parseOutput(suggestion.content);
        const outputDimensions = matrixDimensions(parsed.table);
        const selectedRows = snapshot.sheet.rows;
        const selectedCols = snapshot.sheet.cols;
        const writeFromSingleEmptyCell = snapshot.sheet.empty && selectedRows === 1 && selectedCols === 1;
        const writeInsideEmptySelection = snapshot.sheet.empty
          && outputDimensions.rows <= selectedRows
          && outputDimensions.cols <= selectedCols;

        if (!writeFromSingleEmptyCell && !writeInsideEmptySelection) {
          throw new Error(`Output is ${outputDimensions.rows}x${outputDimensions.cols}, but the selected range is ${selectedRows}x${selectedCols}. Select an empty range large enough for the output or a single empty starting cell.`);
        }

        const targetAddress = snapshot.sheet.start
          ? rangeAddressFromStart(snapshot.sheet.start, outputDimensions.rows, outputDimensions.cols)
          : "";
        await applySheetMatrix(parsed.table, targetAddress);
      }
    },
    rewrite_cells: {
      key: "rewrite_cells",
      label: "editing",
      applyLabel: "Apply to cells",
      detect() {
        return true;
      },
      buildMessages(prompt, snapshot) {
        return [
          {
            role: "system",
            content:
              "You are CloudCIX Guiden® inside Euro-Office Spreadsheets. The operation is cell rewriting. Return JSON only with keys narrative and table. table must be a rectangular two-dimensional array with exactly the same rows and columns as the selected range. Rewrite only editable text cells. For formula cells, numeric cells, boolean cells, and blank cells that should remain blank, return the original value. Do not return formulas, Markdown, code fences, or prose-only answers."
          },
          {
            role: "user",
            content: sheetCommonUserPrompt(prompt, snapshot, "rewrite_cells")
          }
        ];
      },
      parseOutput(content) {
        const parsed = parseSheetJson(content, "Guiden® expected cell rewrite JSON.");
        return {
          narrative: typeof parsed.narrative === "string" ? parsed.narrative : "",
          table: ensureRectangularTable(parsed.table, "Guiden® expected a rectangular table.")
        };
      },
      formatPreview(content) {
        try {
          return renderSheetResult(this.parseOutput(content));
        } catch (error) {
          return renderHtml(content || "");
        }
      },
      async apply(suggestion, snapshot) {
        const parsed = this.parseOutput(suggestion.content);
        const outputDimensions = matrixDimensions(parsed.table);
        if (outputDimensions.rows !== snapshot.sheet.rows || outputDimensions.cols !== snapshot.sheet.cols) {
          throw new Error(`Output is ${outputDimensions.rows}x${outputDimensions.cols}, but the selected range is ${snapshot.sheet.rows}x${snapshot.sheet.cols}.`);
        }

        const matrix = parsed.table.map(function (row, rowIndex) {
          return row.map(function (cell, colIndex) {
            const original = snapshot.sheet.values[rowIndex] && snapshot.sheet.values[rowIndex][colIndex] != null
              ? snapshot.sheet.values[rowIndex][colIndex]
              : "";
            const formula = snapshot.sheet.formulas[rowIndex] && snapshot.sheet.formulas[rowIndex][colIndex] != null
              ? snapshot.sheet.formulas[rowIndex][colIndex]
              : "";
            return isProtectedCell(original, formula) ? original : cell;
          });
        });

        const targetAddress = snapshot.sheet.start
          ? rangeAddressFromStart(snapshot.sheet.start, outputDimensions.rows, outputDimensions.cols)
          : "";
        await applySheetMatrix(matrix, targetAddress);
      }
    }
  };

  function detectSheetOperation(prompt) {
    return [
      sheetOperations.classify_rows,
      sheetOperations.add_column,
      sheetOperations.formula,
      sheetOperations.table,
      sheetOperations.analysis,
      sheetOperations.rewrite_cells
    ].find(function (operation) {
      return operation.detect(prompt);
    }) || sheetOperations.rewrite_cells;
  }

  function plainTextFromSuggestion(suggestion) {
    if (!suggestion) {
      return "";
    }
    return suggestion.html ? htmlToPlainText(suggestion.html) : suggestion.content || "";
  }

  async function applySheetMatrix(matrix, address, formulas) {
    const command = new Function(`return function () {
      const data = ${JSON.stringify(matrix)};
      const targetAddress = ${JSON.stringify(address || "")};
      const useFormula = ${JSON.stringify(Boolean(formulas))};
      const worksheet = Api.GetActiveSheet();
      const range = targetAddress && worksheet && typeof worksheet.GetRange === "function"
        ? worksheet.GetRange(targetAddress)
        : Api.GetSelection
          ? Api.GetSelection()
          : worksheet.GetSelection();
      if (!range || (typeof range.SetValue !== "function" && typeof range.SetFormula !== "function")) {
        return { ok: false, error: "Selected cells cannot be edited in this Euro-Office build." };
      }
      const ok = useFormula && typeof range.SetFormula === "function"
        ? range.SetFormula(data)
        : range.SetValue(data);
      return { ok: ok !== false };
    };`)();
    const result = await callCommand(command);
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : "Unable to write to the selected cells.");
    }
  }

  async function applyCalculatedColumn(header, formula, headerAddress, formulaAddress, firstFormulaRow, formulaRows, targetColumn) {
    const command = new Function(`return function () {
      const worksheet = Api.GetActiveSheet();
      const headerRange = worksheet && worksheet.GetRange(${JSON.stringify(headerAddress)});
      const formulaRange = worksheet && worksheet.GetRange(${JSON.stringify(formulaAddress)});
      const firstFormulaRow = ${JSON.stringify(firstFormulaRow)};
      const formulaRows = ${JSON.stringify(formulaRows)};
      const targetColumn = ${JSON.stringify(targetColumn)};
      const formulaTemplate = ${JSON.stringify(formula)};
      if (!headerRange || !formulaRange || typeof headerRange.SetValue !== "function") {
        return { ok: false, error: "The new column cannot be written in this Euro-Office build." };
      }
      headerRange.SetValue(${JSON.stringify(header)});
      const firstFormulaCell = worksheet.GetRange(${JSON.stringify(formulaAddress.split(":")[0])});
      if (!firstFormulaCell || (typeof firstFormulaCell.SetFormula !== "function" && typeof firstFormulaCell.SetValue !== "function")) {
        return { ok: false, error: "Formulas cannot be written in this Euro-Office build." };
      }
      if (typeof firstFormulaCell.SetFormula === "function") firstFormulaCell.SetFormula(formulaTemplate);
      else firstFormulaCell.SetValue(formulaTemplate);
      if (typeof formulaRange.FillDown === "function") {
        formulaRange.FillDown();
      } else {
        function shiftRelativeRows(formula, offset) {
          return String(formula).replace(/(\\$?)([A-Z]{1,3})(\\$?)(\\d+)/g, function (match, columnAbsolute, column, rowAbsolute, row) {
            if (rowAbsolute === "$") return match;
            return columnAbsolute + column + rowAbsolute + String(Number(row) + offset);
          });
        }
        for (let index = 1; index < formulaRows; index += 1) {
          const cell = worksheet.GetRange(targetColumn + String(firstFormulaRow + index));
          if (!cell) return { ok: false, error: "Unable to access a cell while filling the calculated column." };
          const rowFormula = shiftRelativeRows(formulaTemplate, index);
          if (typeof cell.SetFormula === "function") cell.SetFormula(rowFormula);
          else if (typeof cell.SetValue === "function") cell.SetValue(rowFormula);
          else return { ok: false, error: "Formulas cannot be written in this Euro-Office build." };
        }
      }
      return { ok: true };
    };`)();
    const result = await callCommand(command);
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : "Unable to create the calculated column.");
    }
  }

  async function replaceFirstSlideText(text) {
    const command = new Function(`return function () {
      const replacement = ${JSON.stringify(text)};
      const presentation = Api.GetPresentation();
      const slide = presentation.GetCurrentVisibleSlide ? presentation.GetCurrentVisibleSlide() : presentation.GetCurrentSlide();
      const drawings = slide && typeof slide.GetAllDrawings === "function"
        ? slide.GetAllDrawings()
        : presentation && typeof presentation.GetAllDrawings === "function"
          ? presentation.GetAllDrawings()
          : [];

      function writeContent(drawing) {
        if (!drawing || typeof drawing.GetDocContent !== "function") {
          return false;
        }
        const content = drawing.GetDocContent();
        if (!content || typeof content.RemoveAllElements !== "function") {
          return false;
        }
        content.RemoveAllElements();
        const lines = String(replacement || "").split(/\\r?\\n/).filter(function (line) {
          return line.trim();
        });
        if (!lines.length) {
          lines.push("");
        }
        for (let i = 0; i < lines.length; i += 1) {
          const paragraph = Api.CreateParagraph();
          paragraph.AddText(lines[i].replace(/^[-*]\\s+/, ""));
          if (i === 0 && typeof content.GetElement === "function") {
            const first = content.GetElement(0);
            if (first && typeof first.RemoveAllElements === "function") {
              first.RemoveAllElements();
              first.AddText(lines[i].replace(/^[-*]\\s+/, ""));
              continue;
            }
          }
          if (typeof content.Push === "function") {
            content.Push(paragraph);
          } else if (typeof content.AddElement === "function") {
            content.AddElement(i, paragraph);
          }
        }
        return true;
      }

      for (let i = 0; i < drawings.length; i += 1) {
        if (writeContent(drawings[i])) {
          return { ok: true, updated: "existing" };
        }
      }

      if (slide && Api.CreateShape && Api.CreateSolidFill && Api.CreateStroke && Api.CreateNoFill && slide.AddObject) {
        const fill = Api.CreateSolidFill(Api.RGB(255, 255, 255));
        const stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = Api.CreateShape("rect", 7200000, 2400000, fill, stroke);
        shape.SetPosition(720000, 900000);
        slide.AddObject(shape);
        if (writeContent(shape)) {
          return { ok: true, updated: "created" };
        }
      }
      return { ok: false, error: "Unable to find or create an editable text area on this slide." };
    };`)();
    const result = await callCommand(command);
    if (!result || !result.ok) {
      throw new Error(result && result.error ? result.error : "Unable to apply text to the current slide.");
    }
    return result;
  }

  const writerAdapter = {
    app: "word",
    async readTarget() {
      const detected = await detectTargetText();
      return {
        app: "word",
        target: detected.target,
        label: detected.target === "full_document" ? "whole document" : "selected text",
        text: detected.text
      };
    },
    buildMessages(prompt, snapshot) {
      return [
        {
          role: "system",
          content:
            "You are CloudCIX Guiden®, a writing assistant inside Euro-Office Writer. Return only clean HTML that can be pasted directly into a word-processing document. Do not explain your changes. Do not use Markdown. Do not wrap output in code fences. Preserve facts, names, numbers, and intent. Create professional document structure with headings, paragraphs, lists, and tables where appropriate."
        },
        {
          role: "user",
          content: `Target: ${snapshot.target}\n\nSource text begins:\n${snapshot.text}\nSource text ends.\n\nInstruction:\n${composeInstruction(prompt, snapshot.target)}`
        }
      ];
    },
    formatPreview(content) {
      return formatStructuredOutput(content) || renderHtml(content);
    },
    async applySuggestion(suggestion) {
      if (suggestion.target === "full_document") {
        await pasteWholeDocumentReplacement(suggestion);
      } else {
        await pasteSelectionReplacement(suggestion);
      }
    }
  };

  const sheetAdapter = {
    app: "cell",
    async readTarget(prompt) {
      setStatus("Reading selected cells...", "loading");
      const result = await callCommand(function () {
        const range = Api.GetSelection ? Api.GetSelection() : Api.GetActiveSheet().GetSelection();
        if (!range) {
          return { ok: false, error: "Unable to read the selected cells in this Euro-Office build." };
        }
        const address = typeof range.GetAddress === "function" ? range.GetAddress(true, true, "xlA1", false) : "selected range";
        const values = typeof range.GetValue === "function" ? range.GetValue() : "";
        const formulas = typeof range.GetFormula === "function" ? range.GetFormula() : "";
        const rows = typeof range.GetRowsCount === "function" ? range.GetRowsCount() : 0;
        const cols = typeof range.GetColumnsCount === "function" ? range.GetColumnsCount() : 0;
        return { ok: true, address, values, formulas, rows, cols };
      });
      if (!result || !result.ok) {
        throw new Error(result && result.error ? result.error : "Unable to read the selected cells.");
      }
      const values = normalizeMatrix(result.values);
      const formulas = normalizeMatrix(result.formulas);
      const addressInfo = parseA1Address(result.address);
      const valueDimensions = matrixDimensions(values);
      const dimensions = {
        rows: Number(result.rows) || (addressInfo && addressInfo.rows) || valueDimensions.rows,
        cols: Number(result.cols) || (addressInfo && addressInfo.cols) || valueDimensions.cols
      };
      const operation = detectSheetOperation(prompt);
      return {
        app: "cell",
        target: operation.key,
        operation,
        label: `${operation.label} ${result.address || "selected cells"}`,
        text: matrixToDelimitedText(values),
        sheet: {
          address: result.address || "selected range",
          values,
          formulas,
          start: addressInfo,
          rows: dimensions.rows,
          cols: dimensions.cols,
          empty: isEmptyMatrix(values)
        }
      };
    },
    buildMessages(prompt, snapshot) {
      return snapshot.operation.buildMessages(prompt, snapshot);
    },
    formatPreview(content, snapshot) {
      return snapshot && snapshot.operation
        ? snapshot.operation.formatPreview(content)
        : renderHtml(content || "");
    },
    async applySuggestion(suggestion, snapshot) {
      await snapshot.operation.apply(suggestion, snapshot);
    }
  };

  const slideAdapter = {
    app: "slide",
    async readTarget() {
      setStatus("Checking selected slide text...", "loading");
      const selected = await readSelectionText().catch(function () {
        return "";
      });
      if (selected.trim()) {
        return {
          app: "slide",
          target: "selected_text",
          label: "selected slide text",
          text: selected
        };
      }

      setStatus("Reading current slide...", "loading");
      const result = await callCommand(function () {
        const presentation = Api.GetPresentation();
        const slide = presentation.GetCurrentVisibleSlide ? presentation.GetCurrentVisibleSlide() : presentation.GetCurrentSlide();
        const drawings = slide && typeof slide.GetAllDrawings === "function"
          ? slide.GetAllDrawings()
          : presentation && typeof presentation.GetAllDrawings === "function"
            ? presentation.GetAllDrawings()
            : [];
        const parts = [];
        for (let i = 0; i < drawings.length; i += 1) {
          const drawing = drawings[i];
          if (!drawing || typeof drawing.GetDocContent !== "function") {
            continue;
          }
          const content = drawing.GetDocContent();
          if (content && typeof content.GetText === "function") {
            const text = content.GetText({ ParaSeparator: "\n", NewLineSeparator: "\n" });
            if (text) {
              parts.push(text);
            }
          } else if (content && typeof content.GetElementsCount === "function" && typeof content.GetElement === "function") {
            const paragraphs = [];
            for (let index = 0; index < content.GetElementsCount(); index += 1) {
              const element = content.GetElement(index);
              if (element && typeof element.GetText === "function") {
                paragraphs.push(element.GetText());
              }
            }
            if (paragraphs.length) {
              parts.push(paragraphs.join("\n"));
            }
          }
        }
        return { ok: true, text: parts.join("\n\n") };
      });

      if (!result || !result.ok) {
        throw new Error("Unable to read the current slide in this Euro-Office build.");
      }
      return {
        app: "slide",
        target: "current_slide",
        label: result.text && result.text.trim() ? "current slide" : "empty current slide",
        text: result.text || ""
      };
    },
    buildMessages(prompt, snapshot) {
      return [
        {
          role: "system",
          content:
            "You are CloudCIX Guiden® inside Euro-Office Presentations. Preserve slide layout. Return concise slide-ready text only. Use a short title and clear bullet-style lines when useful. Do not include code fences or explanations. " + structuredOutputInstruction()
        },
        {
          role: "user",
          content: `Target: ${snapshot.target}\n\nCurrent slide text begins:\n${snapshot.text}\nCurrent slide text ends.\n\nInstruction:\n${prompt}\n${formatInstruction()}\n${$("tone").value}\n${$("length").value}`
        }
      ];
    },
    formatPreview(content) {
      return formatStructuredOutput(content) || renderHtml(content);
    },
    async applySuggestion(suggestion, snapshot) {
      const text = plainTextFromSuggestion(suggestion);
      if (snapshot.target === "selected_text") {
        await executeMethod("PasteText", [text]);
        return;
      }
      await replaceFirstSlideText(text);
    }
  };

  async function detectContext(prompt) {
    const app = await detectEditorType();
    const adapter = app === "cell" ? sheetAdapter : app === "slide" ? slideAdapter : writerAdapter;
    const snapshot = await adapter.readTarget(prompt);
    state.adapter = adapter;
    state.snapshot = snapshot;
    return { adapter, snapshot };
  }

  function setBusy(busy, label) {
    state.busy = busy;
    document.body.classList.toggle("is-busy", busy);
    $("generate").disabled = busy;
    $("generate").textContent = busy && label ? label : "Generate";
    $("stop").disabled = !busy;
    $("clear").disabled = busy;
    document.getElementById("saveSettings").disabled = busy;
    document.getElementById("workflow").disabled = busy;
    document.getElementById("keepCitations").disabled = busy;
    $("apply").disabled = busy || !state.suggestion || state.suggestion.previewOnly;
    $("copy").disabled = busy || !state.suggestion;
    document.querySelectorAll(".quick-actions button").forEach((button) => {
      button.disabled = busy;
    });
  }

  function saveSettings() {
    const remember = $("rememberKey").checked;
    const apiKey = $("apiKey").value;

    try {
      localStorage.setItem(STORAGE.rememberKey, remember ? "1" : "0");
      localStorage.setItem(STORAGE.format, $("format").value);
      localStorage.setItem(STORAGE.tone, $("tone").value);
      localStorage.setItem(STORAGE.length, document.getElementById("length").value);
      localStorage.setItem(STORAGE.workflow, document.getElementById("workflow").value);
      localStorage.setItem(STORAGE.keepCitations, document.getElementById("keepCitations").checked ? "1" : "0");
      if (remember) {
        localStorage.setItem(STORAGE.apiKey, apiKey);
      } else {
        localStorage.removeItem(STORAGE.apiKey);
      }
    } catch (error) {
      setStatus("Settings could not be saved in this browser.", "error");
      showSaveSettingsFeedback("Save failed", "error");
      return;
    }

    const message = remember ? "Settings and API key saved." : "Settings saved. API key is not stored.";
    setStatus(message, "done");
    showSaveSettingsFeedback("Saved ✓", "done");
  }

  function showSaveSettingsFeedback(label, mode) {
    const button = $("saveSettings");
    window.clearTimeout(state.saveFeedbackTimer);
    button.textContent = label;
    button.classList.remove("done", "error");
    button.classList.add(mode);
    state.saveFeedbackTimer = window.setTimeout(function () {
      button.textContent = "Save settings";
      button.classList.remove("done", "error");
    }, 2500);
  }

  function loadSettings() {
    const remember = localStorage.getItem(STORAGE.rememberKey) === "1";
    $("rememberKey").checked = remember;
    $("apiKey").value = remember ? localStorage.getItem(STORAGE.apiKey) || "" : "";
    $("prompt").value = localStorage.getItem(STORAGE.prompt) || "";
    if (localStorage.getItem(STORAGE.format)) {
      $("format").value = localStorage.getItem(STORAGE.format);
    }
    if (localStorage.getItem(STORAGE.tone)) {
      $("tone").value = localStorage.getItem(STORAGE.tone);
    }
    if (localStorage.getItem(STORAGE.length)) {
      document.getElementById("length").value = localStorage.getItem(STORAGE.length);
    }
    if (localStorage.getItem(STORAGE.workflow)) {
      document.getElementById("workflow").value = localStorage.getItem(STORAGE.workflow);
    }
    if (localStorage.getItem(STORAGE.keepCitations) !== null) {
      document.getElementById("keepCitations").checked = localStorage.getItem(STORAGE.keepCitations) === "1";
    }
  }

  function escapeHtml(text) {
    return String(text).replace(/[&<>"']/g, function (char) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "\"": "&quot;",
        "'": "&#39;"
      }[char];
    });
  }

  function renderInlineMarkdown(text) {
    let output = "";
    let index = 0;
    while (index < text.length) {
      const linkMatch = text.slice(index).match(/^\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/i);
      if (linkMatch) {
        output += "<a href=\"" + escapeHtml(linkMatch[2]) + "\">" + escapeHtml(linkMatch[1]) + "</a>";
        index += linkMatch[0].length;
        continue;
      }
      if (text.startsWith("**", index)) {
        const end = text.indexOf("**", index + 2);
        if (end !== -1) {
          output += `<strong>${escapeHtml(text.slice(index + 2, end))}</strong>`;
          index = end + 2;
          continue;
        }
      }
      if (text[index] === "*") {
        const end = text.indexOf("*", index + 1);
        if (end !== -1) {
          output += `<em>${escapeHtml(text.slice(index + 1, end))}</em>`;
          index = end + 1;
          continue;
        }
      }
      output += escapeHtml(text[index]);
      index += 1;
    }
    return output;
  }

  function isMarkdownTableRow(line) {
    const trimmed = String(line || "").trim();
    return trimmed.startsWith("|") && trimmed.endsWith("|") && trimmed.slice(1, -1).includes("|");
  }

  function isMarkdownTableSeparator(line) {
    if (!isMarkdownTableRow(line)) {
      return false;
    }
    return splitMarkdownTableRow(line).every(function (cell) {
      return /^:?-{3,}:?$/.test(cell.trim());
    });
  }

  function splitMarkdownTableRow(line) {
    return String(line || "")
      .trim()
      .replace(/^\|/, "")
      .replace(/\|$/, "")
      .split("|")
      .map(function (cell) {
        return cell.trim();
      });
  }

  function renderMarkdownTable(rows) {
    if (!rows.length) {
      return "";
    }

    let header = null;
    let bodyRows = rows;
    if (rows.length > 1 && isMarkdownTableSeparator(rows[1])) {
      header = splitMarkdownTableRow(rows[0]);
      bodyRows = rows.slice(2);
    }

    const headerHtml = header
      ? `<thead><tr>${header.map(function (cell) {
        return styledTag("th", renderInlineMarkdown(cell));
      }).join("")}</tr></thead>`
      : "";
    const bodyHtml = bodyRows.map(function (row) {
      return `<tr>${splitMarkdownTableRow(row).map(function (cell) {
        return styledTag("td", renderInlineMarkdown(cell));
      }).join("")}</tr>`;
    }).join("");

    return styledTag("table", `${headerHtml}<tbody>${bodyHtml}</tbody>`);
  }

  function hasMarkdownTable(lines) {
    for (let index = 0; index < lines.length - 1; index += 1) {
      if (isMarkdownTableRow(lines[index]) && isMarkdownTableSeparator(lines[index + 1])) {
        return true;
      }
    }
    return false;
  }

  function convertMarkdownTablesInMixedContent(text) {
    const lines = String(text || "").split(/\r?\n/);
    const parts = [];
    let tableRows = [];

    function closeTable() {
      if (tableRows.length) {
        parts.push(renderMarkdownTable(tableRows));
        tableRows = [];
      }
    }

    for (const line of lines) {
      if (isMarkdownTableRow(line)) {
        tableRows.push(line);
        continue;
      }
      closeTable();
      parts.push(line);
    }

    closeTable();
    return parts.join("\n");
  }

  function looksLikeHtml(text) {
    return /<\/?\s*(h1|h2|h3|p|strong|em|ul|ol|li|blockquote|a|br|table|thead|tbody|tr|th|td|b|i)\b/i.test(text || "");
  }

  function cleanModelContent(text) {
    return String(text || "")
      .replace(/^\s*```[\t ]*[A-Za-z0-9_-]*[\t ]*\r?\n?/, "")
      .replace(/\s*```\s*$/, "")
      .trim();
  }

  function styledTag(tag, children) {
    if (tag === "table") {
      return `<table border="1" cellspacing="0" cellpadding="4">${children}</table>`;
    }
    return `<${tag}>${children}</${tag}>`;
  }

  function htmlDocument(body) {
    return `<!doctype html><html><head><meta charset="utf-8"></head><body>${body}</body></html>`;
  }

  function htmlBodyFragment(html) {
    const template = document.createElement("template");
    template.innerHTML = html || "";
    const body = template.content.querySelector("body");
    return body ? body.innerHTML : template.innerHTML;
  }

  function sanitizeHtml(html) {
    const template = document.createElement("template");
    template.innerHTML = html || "";

    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        return escapeHtml(node.textContent || "");
      }
      if (node.nodeType !== Node.ELEMENT_NODE) {
        return "";
      }
      const alias = TAG_ALIASES[node.tagName.toLowerCase()];
      const tag = alias || node.tagName.toLowerCase();
      const children = Array.from(node.childNodes).map(walk).join("");
      if (!SAFE_TAGS.has(tag)) {
        return children;
      }
      if (tag === "br") {
        return "<br>";
      }
      if (tag === "a") {
        const href = node.getAttribute("href") || "";
        if (!/^https?:\/\//i.test(href)) {
          return children;
        }
        return "<a href=\"" + escapeHtml(href) + "\">" + children + "</a>";
      }
      return styledTag(tag, children);
    }

    return htmlDocument(Array.from(template.content.childNodes).map(walk).join(""));
  }

  function isStructuredDocumentFormat() {
    return ["report", "briefing", "minutes"].includes($("format").value);
  }

  function isProbablyPlainHeading(line, blockCount) {
    if (!isStructuredDocumentFormat()) {
      return false;
    }
    if (!line || line.length > 90) {
      return false;
    }
    if (/^(dear|hi|hello|kind regards|best regards|regards|yours sincerely|yours faithfully|sincerely)[, ]/i.test(line)) {
      return false;
    }
    if (/[.!?;:]$/.test(line)) {
      return false;
    }
    if (blockCount === 0) {
      return true;
    }
    const words = line.split(/\s+/).filter(Boolean);
    if (words.length > 9) {
      return false;
    }
    const meaningfulWords = words.filter((word) => !/^(and|or|of|in|the|to|for|with|on|a|an)$/i.test(word));
    return meaningfulWords.length > 0 && meaningfulWords.every((word) => /^[A-Z0-9]/.test(word));
  }

  function renderHtml(text) {
    const cleaned = cleanModelContent(text);
    const lines = cleaned.split(/\r?\n/);
    if (looksLikeHtml(cleaned)) {
      return sanitizeHtml(hasMarkdownTable(lines) ? convertMarkdownTablesInMixedContent(cleaned) : cleaned);
    }

    const parts = [];
    let listTag = "";
    let blockCount = 0;
    let tableRows = [];

    function closeList() {
      if (listTag) {
        parts.push(`</${listTag}>`);
        listTag = "";
      }
    }

    function closeTable() {
      if (tableRows.length) {
        parts.push(renderMarkdownTable(tableRows));
        tableRows = [];
        blockCount += 1;
      }
    }

    for (const rawLine of lines) {
      let line = rawLine.trim();
      if (!line) {
        closeTable();
        closeList();
        continue;
      }

      if (isMarkdownTableRow(line)) {
        closeList();
        tableRows.push(line);
        continue;
      }

      closeTable();

      let heading = "";
      if (line.startsWith("### ")) {
        heading = "h3";
        line = line.slice(4).trim();
      } else if (line.startsWith("## ")) {
        heading = "h2";
        line = line.slice(3).trim();
      } else if (line.startsWith("# ")) {
        heading = "h1";
        line = line.slice(2).trim();
      }

      if (heading) {
        closeList();
        parts.push(styledTag(heading, renderInlineMarkdown(line)));
        blockCount += 1;
      } else if (line.startsWith("- ") || line.startsWith("* ")) {
        if (listTag !== "ul") {
          closeList();
          listTag = "ul";
          parts.push("<ul>");
        }
        parts.push(styledTag("li", renderInlineMarkdown(line.slice(2).trim())));
        blockCount += 1;
      } else if (/^\d+[.)]\s+/.test(line)) {
        if (listTag !== "ol") {
          closeList();
          listTag = "ol";
          parts.push("<ol>");
        }
        parts.push(styledTag("li", renderInlineMarkdown(line.replace(/^\d+[.)]\s+/, "").trim())));
        blockCount += 1;
      } else if (line.startsWith("> ")) {
        closeList();
        parts.push(styledTag("blockquote", renderInlineMarkdown(line.slice(2).trim())));
        blockCount += 1;
      } else if (isProbablyPlainHeading(line, blockCount)) {
        closeList();
        parts.push(styledTag(blockCount === 0 ? "h1" : "h2", renderInlineMarkdown(line)));
        blockCount += 1;
      } else {
        closeList();
        const styleName = /^(dear|hi|hello|kind regards|best regards|regards|yours sincerely|yours faithfully|sincerely)[, ]/i.test(line)
          ? "compactP"
          : "p";
        parts.push(styledTag("p", renderInlineMarkdown(line), styleName));
        blockCount += 1;
      }
    }

    closeList();
    closeTable();
    return htmlDocument(parts.join("\n"));
  }

  function isSafeLink(url) {
    return /^https?:\/\//i.test(String(url || "").trim());
  }

  function structuredList(values) {
    if (!Array.isArray(values) || !values.length) {
      return "";
    }
    return styledTag("ul", values.map(function (value) {
      return styledTag("li", renderInlineMarkdown(String(value == null ? "" : value)));
    }).join(""));
  }

  function structuredSources(sources) {
    if (!keepCitations() || !Array.isArray(sources)) {
      return "";
    }
    const items = sources.filter(function (source) {
      return source && isSafeLink(source.url);
    }).map(function (source) {
      const title = String(source.title || source.url);
      return styledTag("li", "<a href=\"" + escapeHtml(String(source.url)) + "\">" + renderInlineMarkdown(title) + "</a>");
    });
    return items.length ? styledTag("h3", "Sources") + styledTag("ul", items.join("")) : "";
  }

  function formatStructuredOutput(content) {
    const workflow = selectedWorkflow();
    if (workflow === "auto") {
      return "";
    }
    const data = extractJsonObject(content);
    if (!data || typeof data !== "object" || Array.isArray(data)) {
      return "";
    }
    const parts = [];
    if (data.title) {
      parts.push(styledTag("h1", renderInlineMarkdown(String(data.title))));
    }
    if (workflow === "minutes") {
      [["Attendees", data.attendees], ["Decisions", data.decisions], ["Actions", data.actions], ["Next steps", data.next_steps]].forEach(function (section) {
        const list = structuredList(section[1]);
        if (list) {
          parts.push(styledTag("h2", section[0]) + list);
        }
      });
    } else if (workflow === "table") {
      if (Array.isArray(data.columns) && Array.isArray(data.rows)) {
        const headers = data.columns.map(function (cell) { return styledTag("th", renderInlineMarkdown(String(cell == null ? "" : cell))); }).join("");
        const rows = data.rows.filter(Array.isArray).map(function (row) {
          return "<tr>" + row.map(function (cell) { return styledTag("td", renderInlineMarkdown(String(cell == null ? "" : cell))); }).join("") + "</tr>";
        }).join("");
        parts.push(styledTag("table", "<thead><tr>" + headers + "</tr></thead><tbody>" + rows + "</tbody>"));
      }
    } else if (workflow === "slides" && Array.isArray(data.slides)) {
      data.slides.forEach(function (slide, index) {
        const heading = slide && slide.title ? String(slide.title) : "Slide " + String(index + 1);
        parts.push(styledTag("h2", renderInlineMarkdown(heading)));
        const list = structuredList(slide && slide.bullets);
        if (list) {
          parts.push(list);
        }
      });
    }
    parts.push(structuredSources(data.sources));
    return parts.length ? htmlDocument(parts.join("")) : "";
  }

  function stripLinksForApply(html) {
    if (keepCitations()) {
      return html;
    }
    const template = document.createElement("template");
    template.innerHTML = html || "";
    template.content.querySelectorAll("a").forEach(function (link) {
      link.replaceWith(document.createTextNode(link.textContent || ""));
    });
    return htmlDocument(template.innerHTML);
  }

  function formatInstruction() {
    const format = $("format").value;
    const map = {
      auto: "Choose the best document structure for the user's instruction and source text.",
      email: "Format as a polished email with greeting, short paragraphs, and a professional sign-off. Use placeholders only when the source lacks a real name or detail.",
      report: "Format as a report with a clear title, section headings, concise paragraphs, and bullet lists where useful.",
      briefing: "Format as a briefing note with a title, context, key points, recommendation/next steps, and concise bullets.",
      bullets: "Format as a clean bullet list with short, parallel bullet points.",
      minutes: "Format as meeting minutes with headings for attendees, decisions, actions, and next steps where appropriate."
    };
    return map[format] || map.auto;
  }

  function selectedWorkflow() {
    const control = document.getElementById("workflow");
    return control ? control.value : "auto";
  }

  function keepCitations() {
    const control = document.getElementById("keepCitations");
    return !control || control.checked;
  }

  function structuredOutputInstruction() {
    const workflow = selectedWorkflow();
    const citationRule = keepCitations()
      ? "When sources are available, include them only in a sources array with title and https URL. Never invent a source or URL."
      : "Do not include a sources section or links.";
    if (workflow === "minutes") {
      return "Return JSON only with title, attendees, decisions, actions, next_steps, and sources. attendees, decisions, actions, next_steps, and sources are arrays. " + citationRule;
    }
    if (workflow === "table") {
      return "Return JSON only with title, columns, rows, and sources. columns is an array, rows is a rectangular array, and sources is an array. " + citationRule;
    }
    if (workflow === "slides") {
      return "Return JSON only with title, slides, and sources. slides is an array of objects with title and bullets arrays. " + citationRule;
    }
    return citationRule;
  }

  function composeInstruction(prompt, target) {
    return [
      prompt,
      formatInstruction(),
      document.getElementById("tone").value,
      document.getElementById("length").value,
      "Return valid HTML only unless a structured workflow is selected. Use only these tags: h1, h2, h3, p, strong, em, a, ul, ol, li, blockquote, table, thead, tbody, tr, th, td, br.",
      "Use real paragraph tags for separate paragraphs. Use lists for lists. Use tables only when tabular information is clearly useful.",
      structuredOutputInstruction(),
      target === "full_document"
        ? "You are editing the whole document. Return the complete revised document, not commentary."
        : "You are editing selected text. Return only the replacement text for that selection. Match the original structure and do not introduce new headings, tables, or document-level formatting unless the user explicitly asks for them."
    ].filter(Boolean).join("\n");
  }

  function appendStreamChunk(payload) {
    if (!payload || !payload.choices || !payload.choices[0]) {
      return "";
    }
    const choice = payload.choices[0];
    if (choice.delta && typeof choice.delta.content === "string") {
      return choice.delta.content;
    }
    if (choice.message && typeof choice.message.content === "string") {
      return choice.message.content;
    }
    if (typeof choice.text === "string") {
      return choice.text;
    }
    return "";
  }

  function updateStreamingSuggestion(content, snapshot, adapter) {
    const html = adapter.formatPreview(content || " ", snapshot);
    state.suggestion = {
      content,
      html,
      target: snapshot.target,
      app: snapshot.app,
      previewOnly: Boolean(snapshot.operation && snapshot.operation.previewOnly),
      applyLabel: snapshot.operation && snapshot.operation.applyLabel ? snapshot.operation.applyLabel : "Apply"
    };
    $("preview").innerHTML = htmlBodyFragment(html);
    $("preview").scrollTop = $("preview").scrollHeight;
  }

  async function readStreamingResponse(response, snapshot, adapter, onUpdate) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    let buffer = "";
    let content = "";

    function processEvent(eventText) {
      const dataLines = eventText.split(/\r?\n/)
        .filter((line) => line.startsWith("data:"))
        .map((line) => line.slice(5).trim());
      if (!dataLines.length) {
        return false;
      }
      const data = dataLines.join("\n");
      if (data === "[DONE]") {
        return true;
      }
      try {
        const payload = JSON.parse(data);
        const chunk = appendStreamChunk(payload);
        if (chunk) {
          content += chunk;
          onUpdate(content, snapshot, adapter);
        }
      } catch (error) {
        // Ignore malformed partial events; the next chunk may complete them.
      }
      return false;
    }

    while (true) {
      const result = await reader.read();
      if (result.done) {
        break;
      }
      buffer += decoder.decode(result.value, { stream: true });
      const events = buffer.split(/\r?\n\r?\n/);
      buffer = events.pop() || "";
      for (const eventText of events) {
        if (processEvent(eventText)) {
          return content;
        }
      }
    }

    buffer += decoder.decode();
    if (buffer.trim()) {
      processEvent(buffer);
    }
    return content;
  }

  function createSuggestion(content, snapshot, adapter) {
    if (snapshot.operation && typeof snapshot.operation.parseOutput === "function") {
      snapshot.operation.parseOutput(content);
    }
    return {
      content,
      html: adapter.formatPreview(content, snapshot),
      target: snapshot.target,
      app: snapshot.app,
      previewOnly: Boolean(snapshot.operation && snapshot.operation.previewOnly),
      applyLabel: snapshot.operation && snapshot.operation.applyLabel ? snapshot.operation.applyLabel : "Apply"
    };
  }

  async function callCloudCIX(apiKey, prompt, snapshot, adapter, signal) {
    const response = await fetch(GUIDEN_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: MODEL,
        stream: true,
        max_tokens: MAX_TOKENS,
        messages: adapter.buildMessages(prompt, snapshot)
      }),
      signal
    });

    if (response.ok && response.body) {
      const contentType = response.headers.get("Content-Type") || "";
      if (contentType.includes("text/event-stream")) {
        const content = await readStreamingResponse(response, snapshot, adapter, updateStreamingSuggestion);
        if (!content.trim()) {
          throw new Error("CloudCIX returned an empty response.");
        }
        return createSuggestion(content, snapshot, adapter);
      }
    }

    const responseText = await response.text();
    let payload = {};
    try {
      payload = responseText ? JSON.parse(responseText) : {};
    } catch (error) {
      if (!response.ok) {
        const preview = responseText.replace(/\s+/g, " ").trim().slice(0, 160);
        throw new Error(`Guiden® API returned HTTP ${response.status}: ${preview || response.statusText}`);
      }
      throw new Error("CloudCIX returned an invalid JSON response.");
    }

    if (!response.ok) {
      const message = payload.error && typeof payload.error === "object" ? payload.error.message : responseText;
      throw new Error(`CloudCIX HTTP ${response.status}: ${message || response.statusText}`);
    }

    const content = payload.choices && payload.choices[0] && payload.choices[0].message
      ? payload.choices[0].message.content || ""
      : "";
    if (!content.trim()) {
      throw new Error("CloudCIX returned an empty response.");
    }
    return createSuggestion(content, snapshot, adapter);
  }

  function showSuggestion(suggestion) {
    state.suggestion = suggestion;
    $("preview").classList.remove("streaming");
    $("preview").innerHTML = htmlBodyFragment(suggestion.html);
    $("resultGroup").hidden = false;
    $("apply").textContent = suggestion.applyLabel || "Apply";
    if (suggestion.previewOnly) {
      setStatus("Preview ready. This spreadsheet response is not applied to cells.", "done");
    }
    setBusy(false);
  }

  async function generate() {
    if (state.busy) {
      return;
    }

    const prompt = $("prompt").value.trim();
    const apiKey = $("apiKey").value;

    if (!prompt) {
      setStatus("Choose an action or type an instruction.", "error");
      return;
    }
    if (!apiKey) {
      setStatus("CloudCIX API key is required in Settings.", "error");
      return;
    }
    if (prompt.length > CHARACTER_LIMIT) {
      setStatus(`Instruction is too long. Keep input below ${CHARACTER_LIMIT} characters.`, "error");
      return;
    }

    setBusy(true, "Generating...");
    state.controller = new AbortController();

    try {
      localStorage.setItem(STORAGE.prompt, prompt);
      localStorage.setItem(STORAGE.format, $("format").value);
      localStorage.setItem(STORAGE.tone, $("tone").value);
      localStorage.setItem(STORAGE.length, $("length").value);
      if ($("rememberKey").checked) {
        localStorage.setItem(STORAGE.apiKey, apiKey);
      }

      const context = await detectContext(prompt);
      const target = context.snapshot.target;
      const text = context.snapshot.text;
      if (text.length + prompt.length > CHARACTER_LIMIT) {
        throw new Error(`Input text is too long. Reduce it below ${CHARACTER_LIMIT} characters.`);
      }

      setStatus(`Drafting suggestion for ${context.snapshot.label}...`, "loading");
      state.suggestion = null;
      $("preview").innerHTML = "";
      $("preview").classList.add("streaming");
      $("resultGroup").hidden = false;
      const suggestion = await callCloudCIX(apiKey, prompt, context.snapshot, context.adapter, state.controller.signal);
      suggestion.sourceEmpty = context.snapshot.app === "word" && target === "full_document" && !text.trim();
      showSuggestion(suggestion);
      setStatus(
        suggestion.previewOnly ? "Preview ready. Review it, copy it, or discard it." : "Suggestion ready. Review it, then apply or discard.",
        "done"
      );
    } catch (error) {
      $("preview").classList.remove("streaming");
      if (error.name === "AbortError") {
        setStatus("Stopped.", "error");
      } else {
        setStatus(error.message || "Unable to generate a suggestion.", "error");
      }
      setBusy(false);
    } finally {
      state.controller = null;
    }
  }

  async function applySuggestion() {
    if (!state.suggestion || state.busy) {
      return;
    }

    setBusy(true, "Applying...");
    setStatus("Applying suggestion...", "loading");
    try {
      const adapter = state.adapter || writerAdapter;
      state.suggestion.html = stripLinksForApply(state.suggestion.html);
      await adapter.applySuggestion(state.suggestion, state.snapshot);
      setStatus("Applied.", "done");
      discardSuggestion(false);
    } catch (error) {
      setStatus(error.message || "Unable to apply suggestion.", "error");
    } finally {
      setBusy(false);
    }
  }

  async function copySuggestion() {
    if (!state.suggestion) {
      return;
    }
    try {
      await navigator.clipboard.writeText(state.suggestion.content);
      setStatus("Copied suggestion.", "done");
    } catch (error) {
      setStatus("Unable to copy suggestion.", "error");
    }
  }

  function discardSuggestion(updateStatus = true) {
    state.suggestion = null;
    state.snapshot = null;
    state.adapter = null;
    $("preview").classList.remove("streaming");
    $("preview").innerHTML = "";
    $("resultGroup").hidden = true;
    $("apply").textContent = "Apply";
    setBusy(false);
    if (updateStatus) {
      setStatus("Suggestion discarded.", "");
    }
  }

  function stop() {
    if (state.controller) {
      state.controller.abort();
    }
  }

  function clearPrompt() {
    $("prompt").value = "";
    localStorage.removeItem(STORAGE.prompt);
    discardSuggestion(false);
    setStatus("Ready.", "");
  }

  function useQuickAction(event) {
    const button = event.target.closest("button[data-prompt]");
    if (!button) {
      return;
    }
    $("prompt").value = button.dataset.prompt || "";
    generate();
  }

  window.Asc.plugin.init = function () {
    hideRedundantPluginCloseButton();
    loadSettings();
    setBusy(false);
  };

  window.Asc.plugin.button = function () {};

  document.addEventListener("DOMContentLoaded", function () {
    hideRedundantPluginCloseButton();
    loadSettings();
    setBusy(false);
    document.querySelectorAll(".quick-actions").forEach(function (actions) {
      actions.addEventListener("click", useQuickAction);
    });
    $("generate").addEventListener("click", generate);
    $("stop").addEventListener("click", stop);
    $("clear").addEventListener("click", clearPrompt);
    $("apply").addEventListener("click", applySuggestion);
    $("copy").addEventListener("click", copySuggestion);
    $("discard").addEventListener("click", function () {
      discardSuggestion(true);
    });
    $("saveSettings").addEventListener("click", saveSettings);
  });
})(window, document);
