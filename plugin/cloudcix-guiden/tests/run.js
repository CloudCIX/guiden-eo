"use strict";
const assert = require("assert");
global.window = global;
const C = require("../scripts/core.js");
global.Asc = {
  plugin: {
    // Like the editor, run the command from its source text, so it can't reach anything in the plugin's scope.
    callCommand(fn, _close, _calc, callback) {
      try {
        callback(new Function("return (" + String(fn) + ")();")());
      } catch (error) {
        throw error;
      }
    },
    executeMethod(_name, _args, callback) {
      callback(false);
    }
  }
};
require("../scripts/host.js");
require("../scripts/writer.js");
require("../scripts/cells.js");
require("../scripts/slides.js");
require("../scripts/api.js");
require("../scripts/menus.js");
let passed = 0;
function test(name, fn) {
  try {
    const value = fn();
    if (value && typeof value.then === "function")
      return value.then(() => {
        passed += 1;
        process.stdout.write("✓ " + name + "\n");
      });
    passed += 1;
    process.stdout.write("✓ " + name + "\n");
  } catch (error) {
    error.message = name + ": " + error.message;
    throw error;
  }
}
// Async tests share globals such as Api and fetch, so they run one after another.
const queued = [];
function queue(name, fn) {
  queued.push([name, fn]);
}

test("action registry keeps explicit editor actions", () => {
  assert.strictEqual(global.Guiden.actions["cells.format"].mutates, true);
  assert.strictEqual(global.Guiden.actions["cells.ask"].mutates, false);
  assert.strictEqual(global.Guiden.actions["slides.rewrite"].editor, "slide");
  assert.strictEqual(global.Guiden.actions["writer.improve"].editor, "word");
});
test("A1 parsing and range expansion", () => {
  assert.deepStrictEqual(C.parseA1Address("'Sheet 1'!$E$2:$E$10"), {
    startCol: 5,
    startRow: 2,
    endCol: 5,
    endRow: 10,
    rows: 9,
    cols: 1
  });
  assert.strictEqual(C.rangeAddress({ startCol: 5, startRow: 2 }, 9, 1), "E2:E10");
});
test("fingerprints are deterministic and change with source", () => {
  assert.strictEqual(C.fingerprint({ b: 2, a: 1 }), C.fingerprint({ a: 1, b: 2 }));
  assert.notStrictEqual(C.fingerprint([1, 2]), C.fingerprint([1, 3]));
});
test("all condition operators evaluate locally", () => {
  assert(C.conditionMatches(82, { operator: "lt", value: 83 }));
  assert(C.conditionMatches(83, { operator: "lte", value: 83 }));
  assert(C.conditionMatches(84, { operator: "gt", value: 83 }));
  assert(C.conditionMatches(83, { operator: "gte", value: 83 }));
  assert(C.conditionMatches("Active", { operator: "eq", value: "Active" }));
  assert(C.conditionMatches("A", { operator: "neq", value: "B" }));
  assert(C.conditionMatches("Finance Team", { operator: "contains", value: "finance" }));
});
test("sparse updates validate coordinates, before values, and formulas", () => {
  const s = {
    rows: 2,
    cols: 2,
    values: [
      ["alice", 10],
      ["bob", 20]
    ],
    formulas: [
      ["", ""],
      ["", "=A1"]
    ]
  };
  assert.deepStrictEqual(
    C.validateUpdates([{ row: 0, column: 0, before: "alice", after: "Alice" }], s)[0].after,
    "Alice"
  );
  assert.throws(() => C.validateUpdates([{ row: 9, column: 0, before: "", after: "x" }], s));
  assert.throws(() => C.validateUpdates([{ row: 0, column: 0, before: "wrong", after: "x" }], s));
  assert.throws(() => C.validateUpdates([{ row: 1, column: 1, before: 20, after: 21 }], s));
});
test("conditional formatting expands only matching E2:E10 cells", () => {
  const values = [[95.5], [82.3], [77.8], [91], [88.6], [76.2], [93.4], [85.1], [89.7]],
    snapshot = { rows: 9, cols: 1, values };
  const styles = C.validateStyles(
    [{ target: { type: "condition", column: 0, operator: "lt", value: 83 }, fill: "#FF0000" }],
    snapshot
  );
  assert.deepStrictEqual(
    C.expandStyleRules(styles, snapshot).map(x => x.row),
    [1, 2, 5]
  );
});
test("unsafe links are rejected without repair", () => {
  assert(C.isSafeLink("https://example.com/source"));
  assert(!C.isSafeLink("javascript:alert(1)"));
  assert.deepStrictEqual(
    C.safeSources([
      { title: "bad", url: "example.com" },
      { title: "ok", url: "https://example.com" }
    ]),
    [{ title: "ok", url: "https://example.com" }]
  );
});
test("Writer paste HTML preserves safe hrefs as native-import anchors", () => {
  const html = global.Guiden.host.writerHtmlDocument(
    '<p>Read <a href="https://docs.example.com/podnet" target="_blank" rel="noopener noreferrer">PodNet docs</a>.</p>'
  );
  assert(html.startsWith('<!doctype html><html><head><meta charset="utf-8">'));
  assert(html.includes('<a href="https://docs.example.com/podnet">PodNet docs</a>'));
  assert(!html.includes("target="));
  assert(!html.includes("rel="));
});
// The mocked editor records what Guiden asked it to do. options.current is the text of the paragraph holding the
// cursor (GetCurrentParagraph is missing without it), and options.lists=false leaves out CreateNumbering.
function mockWriter(withHyperlinks, methodAnswers, options) {
  options = options || {};
  const calls = [],
    inserted = [],
    inline = [],
    after = [];
  const run = () => ({
    parts: [],
    AddText(t) {
      this.parts.push(t);
    },
    SetBold() {
      this.bold = true;
    },
    SetItalic() {},
    AddLineBreak() {
      this.parts.push("\n");
    }
  });
  // Each paragraph inserted after the cursor's paragraph lands after the one before it, so recording order is enough.
  function InsertParagraph(para, position, returnNew) {
    after.push(para);
    return returnNew ? para : true;
  }
  const current = {
    deleted: false,
    GetText: () => options.current,
    InsertParagraph,
    Delete() {
      this.deleted = true;
      return true;
    }
  };
  const doc = {
    InsertContent(content, isInline) {
      inserted.push(...content);
      inline.push(isInline);
      return true;
    },
    GetStyle: name => ({ name })
  };
  if (options.current !== undefined) doc.GetCurrentParagraph = () => current;
  if (options.lists !== false) doc.CreateNumbering = kind => ({ GetLevel: level => ({ kind, level }) });
  global.Api = {
    GetDocument: () => doc,
    CreateParagraph: () => ({
      items: [],
      InsertParagraph,
      SetStyle(style) {
        this.style = style.name;
      },
      SetNumbering(level) {
        this.numbering = level;
      },
      Push(x) {
        this.items.push(x);
      },
      AddElement(x) {
        this.items.push(x);
      }
    }),
    CreateRun: run
  };
  if (withHyperlinks) global.Api.CreateHyperlink = (href, text) => ({ href, text });
  const executeMethod = global.Asc.plugin.executeMethod;
  global.Asc.plugin.executeMethod = (name, args, callback) => {
    calls.push([name, args]);
    callback((methodAnswers || {})[name]);
  };
  return {
    calls,
    inserted,
    inline,
    after,
    current,
    restore() {
      global.Asc.plugin.executeMethod = executeMethod;
    }
  };
}
const text = para => para.items.map(item => (item.parts ? item.parts.join("") : item.text)).join("");
// writerNativePayload parses HTML with the browser's DOM, so it is covered by tests/browser.js; here it is stubbed.
function withPayload(payload, fn) {
  const ui = global.Guiden.ui;
  global.Guiden.ui = { writerNativePayload: () => payload };
  return Promise.resolve(fn()).finally(() => (global.Guiden.ui = ui));
}
queue("Writer inserts native hyperlinks, keeping bold text and line breaks", async () => {
  const editor = mockWriter(true);
  const payload = {
    supported: true,
    linkCount: 1,
    blocks: [
      { paragraphStyle: "Heading 2", segments: [{ text: "Title" }] },
      {
        paragraphStyle: "",
        segments: [{ text: "One", bold: true }, { text: "\ntwo " }, { text: "Doc", href: "https://example.com/doc" }]
      }
    ]
  };
  try {
    await withPayload(payload, async () => {
      const result = await global.Guiden.host.pasteWriter("<p>ignored by the stub</p>");
      assert.strictEqual(result.links, 1);
      assert.deepStrictEqual(editor.calls, [], "the HTML importer was not needed");
      assert.strictEqual(editor.inserted.length, 2);
      assert.strictEqual(editor.inserted[0].style, "Heading 2");
      const body = editor.inserted[1].items;
      assert(body[0].bold);
      assert.deepStrictEqual(
        body.find(item => item.href),
        { href: "https://example.com/doc", text: "Doc" }
      );
      assert(
        body.some(item => item.parts && item.parts.includes("\n")),
        "the newline became a line break"
      );
    });
  } finally {
    editor.restore();
  }
});
queue("Writer falls back to the HTML importer when native hyperlinks are unavailable", async () => {
  const editor = mockWriter(false),
    warn = console.warn;
  console.warn = () => {};
  try {
    const html = '<p>See <a href="https://example.com/doc">Doc</a></p>';
    const result = await withPayload(
      {
        supported: true,
        linkCount: 1,
        blocks: [{ paragraphStyle: "", segments: [{ text: "Doc", href: "https://example.com/doc" }] }]
      },
      () => global.Guiden.host.pasteWriter(html)
    );
    assert.deepStrictEqual(editor.calls, [["PasteHtml", [html]]]);
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.warnings.length, 1);
  } finally {
    editor.restore();
    console.warn = warn;
  }
});
queue("API asks for strict JSON with the configured model and key", async () => {
  const originalFetch = global.fetch;
  let request;
  global.fetch = async (url, options) => {
    request = { url, options, body: JSON.parse(options.body) };
    return {
      ok: true,
      headers: { get: () => "application/json" },
      json: async () => ({ choices: [{ message: { content: "{}" } }] })
    };
  };
  try {
    global.Guiden.api.configure({ model: "Test-Model" });
    await global.Guiden.api.complete("secret", [{ role: "user", content: "hi" }], new AbortController().signal);
    assert.deepStrictEqual(request.body.response_format, { type: "json_object" });
    assert.strictEqual(request.body.model, "Test-Model");
    assert.strictEqual(request.options.headers.Authorization, "Bearer secret");
  } finally {
    global.fetch = originalFetch;
    global.Guiden.api.configure({});
  }
});
test("streaming text decodes partial structured responses", () => {
  assert.strictEqual(C.streamingText('{"answer":"Hello\\nworld'), "Hello\nworld");
  assert.strictEqual(C.streamingText('{"replacementHtml":"<p>Writing now</p>'), "Writing now");
  assert.strictEqual(C.streamingText("plain streamed answer"), "plain streamed answer");
  assert.strictEqual(C.streamingText('{"updates":['), "");
});
test("streaming Writer HTML keeps a readable prose flow", () => {
  assert.strictEqual(
    C.streamingText('{"replacementHtml":"<h2>Title</h2><p>First.</p><p>Second.</p>'),
    "Title\n\nFirst.\n\nSecond."
  );
});
test("shape updates require exact indexed source text", () => {
  const snapshot = { shapes: [{ text: "Title" }, { text: "Body" }] };
  assert.strictEqual(
    C.validateShapeUpdates([{ shapeIndex: 1, before: "Body", after: "Better" }], snapshot)[0].after,
    "Better"
  );
  assert.throws(() => C.validateShapeUpdates([{ shapeIndex: 1, before: "Wrong", after: "Better" }], snapshot));
});
test("malformed model JSON is rejected", () => assert.throws(() => C.extractJson("not json")));
test("zero-change and partial apply results are not success", () => {
  assert.throws(() => C.validateApplyResult({ ok: true, changed: 0, target: "A1" }));
  assert.throws(() => C.validateApplyResult({ ok: false, changed: 1, target: "A1", error: "partial" }));
  assert.strictEqual(C.validateApplyResult({ ok: true, changed: 2, target: "A1:A2" }).changed, 2);
});
test("every registered action exposes the V2 contract", () => {
  Object.values(global.Guiden.actions).forEach(action =>
    [
      "id",
      "editor",
      "label",
      "description",
      "mutates",
      "supportedTargets",
      "buildRequest",
      "parseResponse",
      "buildPreview",
      "revalidateTarget",
      "apply"
    ].forEach(key => assert.notStrictEqual(action[key], undefined, action.id + " missing " + key))
  );
});
test("cell edit parser rejects full-matrix responses", () => {
  const action = global.Guiden.actions["cells.edit"],
    snapshot = { rows: 1, cols: 1, values: [["a"]], formulas: [[""]] };
  assert.throws(() => action.parseResponse('{"table":[["b"]]}', snapshot));
});
test("formula parser rejects prose", () =>
  assert.throws(() => global.Guiden.actions["cells.formula"].parseResponse('{"formula":"SUM(A1:A2)"}')));
test("Writer mutations reject raw HTML without the response contract", () => {
  assert.throws(
    () => global.Guiden.actions["writer.improve"].parseResponse("<p>Clearer <strong>content</strong>.</p>"),
    /invalid response/i
  );
});
test("Writer mutations reject truncated JSON even when it contains HTML", () => {
  assert.throws(
    () => global.Guiden.actions["writer.draft"].parseResponse('{"status":"ready","replacementHtml":"<p>Only part'),
    /invalid response/i
  );
});
test("Writer refusals can never become applicable document content", () => {
  const action = global.Guiden.actions["writer.draft"],
    screenshot =
      "<p>This action cannot be completed as there is no source text or context provided. Please provide the relevant source material.</p>",
    refusal = action.parseResponse(JSON.stringify({ replacementHtml: screenshot })),
    explicit = action.parseResponse(
      JSON.stringify({ status: "needs_input", message: "Provide the intended audience." })
    );
  assert.strictEqual(refusal.needsInput, true);
  assert.strictEqual(explicit.needsInput, true);
  assert.strictEqual(action.buildPreview(refusal, { target: "cursor" }).kind, "needs-input");
});
test("Draft at an empty cursor requests new content rather than source text", () => {
  const action = global.Guiden.actions["writer.draft"],
    snapshot = { target: "cursor", targetLabel: "Current cursor", text: "", sourceHtml: "" },
    request = action.buildRequest(snapshot, "Draft an introduction.", {
      tone: "auto",
      length: "auto",
      structure: "auto"
    });
  assert(request[0].content.includes("Drafting at an empty cursor is normal"));
  const result = action.parseResponse(
    JSON.stringify({ status: "ready", replacementHtml: "<p>A new introduction.</p>" })
  );
  assert.strictEqual(action.buildPreview(result, snapshot).kind, "writer-new");
});
test("plain-text replacements become one paragraph per line, so paragraph breaks survive", () => {
  const rewrite = global.Guiden.actions["writer.rewrite"],
    parse = html => rewrite.parseResponse(JSON.stringify({ status: "ready", replacementHtml: html })).replacementHtml;
  assert.strictEqual(parse("Le prochain trimestre."), "<p>Le prochain trimestre.</p>");
  assert.strictEqual(
    parse("Premier & deux.\n\nSecond <em>ici</em>."),
    "<p>Premier & deux.</p><p>Second <em>ici</em>.</p>"
  );
  assert.strictEqual(parse("Un.<br><br>Deux."), "<p>Un.</p><p>Deux.</p>");
  assert.strictEqual(parse("<h2>Titre</h2>\nTexte"), "<h2>Titre</h2>\nTexte");
  const system = rewrite.buildRequest({ target: "selection", targetLabel: "Selection", text: "x" }, "", {
    tone: "auto",
    length: "auto",
    structure: "auto"
  })[0].content;
  assert.match(system, /Irish English/);
  assert.match(system, /placeholder in square brackets/);
  assert.match(system, /tense/);
});
test("a draft at the cursor sees the rest of the document as context, and other actions do not", () => {
  const options = { tone: "auto", length: "auto", structure: "auto" },
    snapshot = { target: "cursor", targetLabel: "Current cursor", text: "", contextText: "Our new software platform." },
    draft = global.Guiden.actions["writer.draft"].buildRequest(snapshot, "Write next steps.", options)[0].content,
    improve = global.Guiden.actions["writer.improve"].buildRequest(
      Object.assign({}, snapshot, { target: "selection", text: "x" }),
      "",
      options
    )[0].content;
  assert.match(draft, /rest of the document, for context only[^]*Our new software platform\./);
  assert.doesNotMatch(improve, /Our new software platform/);
});
test("Ordinary document prose is not mistaken for an assistant refusal", () => {
  const result = global.Guiden.actions["writer.draft"].parseResponse(
    JSON.stringify({ replacementHtml: "<p>This report cannot be completed until Phase 2 begins.</p>" })
  );
  assert.strictEqual(result.needsInput, undefined);
});
test("Writer accepts a plain-text answer but still rejects malformed JSON", () => {
  const warn = console.warn;
  console.warn = () => {};
  try {
    assert.strictEqual(
      global.Guiden.actions["writer.ask"].parseResponse("A useful answer.").answer,
      "A useful answer."
    );
    assert.throws(() => global.Guiden.actions["writer.ask"].parseResponse("{broken json"));
  } finally {
    console.warn = warn;
  }
});
test("RAG retrieval queries contain only the user instruction", () => {
  const options = { tone: "auto", length: "auto", structure: "auto", tableTheme: "business", slideCount: "3" },
    question = "Draft a PodNet overview from company documentation.";
  const writer = global.Guiden.actions["writer.draft"].buildRequest(
    { target: "cursor", targetLabel: "Current cursor", text: "", sourceHtml: "" },
    question,
    options
  );
  const cells = global.Guiden.actions["cells.ask"].buildRequest(
    {
      targetLabel: "Sheet1 · A1:B2",
      rows: 2,
      cols: 2,
      text: "PodNet\tStatus",
      formulas: [
        ["", ""],
        ["", ""]
      ]
    },
    question,
    options
  );
  const slides = global.Guiden.actions["slides.new"].buildRequest(
    { targetLabel: "Slide 1", text: "Existing slide" },
    question,
    options
  );
  [writer, cells, slides].forEach(messages => {
    assert.strictEqual(messages[messages.length - 1].role, "user");
    assert.strictEqual(messages[messages.length - 1].content, question);
    assert(messages.slice(0, -1).every(message => message.role === "system"));
  });
  assert(writer[0].content.includes("Current cursor"));
  assert(cells[0].content.includes("PodNet"));
  assert(slides[0].content.includes("Existing slide"));
});
test("schema retry preserves the exact user query", () => {
  const question = "Draft a PodNet overview from company documentation.",
    messages = C.groundedMessages("Return JSON.", "Current cursor", question),
    retry = C.schemaRetryMessages(messages);
  assert.notStrictEqual(retry, messages);
  assert(retry[0].content.includes("RESPONSE FORMAT RETRY"));
  assert.strictEqual(retry[retry.length - 1].content, question);
  assert.strictEqual(messages[0].content.includes("RESPONSE FORMAT RETRY"), false);
});
test("writer links remain structured sources", () => {
  const parsed = global.Guiden.actions["writer.ask"].parseResponse(
    '{"answer":"See source","sources":[{"title":"Doc","url":"https://example.com/doc"}]}'
  );
  assert.strictEqual(parsed.sources[0].url, "https://example.com/doc");
});
queue("dynamic slide commands compile and keep indexed targets", async () => {
  const contents = [
    {
      text: "Title",
      RemoveAllElements() {
        this.text = "";
      },
      Push(p) {
        this.text += p.text;
      }
    },
    {
      text: "Body",
      RemoveAllElements() {
        this.text = "";
      },
      Push(p) {
        this.text += p.text;
      }
    }
  ];
  const drawings = contents.map(content => ({
    GetDocContent() {
      return content;
    }
  }));
  const slide = {
    objects: [],
    GetAllDrawings() {
      return drawings;
    },
    GetLayout() {
      return { id: 1 };
    },
    AddObject(shape) {
      this.objects.push(shape);
    }
  };
  const presentation = {
    GetCurrentSlide() {
      return slide;
    },
    AddSlide(newSlide) {
      this.added = newSlide;
      return true;
    }
  };
  global.Api = {
    GetPresentation() {
      return presentation;
    },
    CreateParagraph() {
      return {
        text: "",
        AddText(value) {
          this.text += value;
        }
      };
    },
    CreateSlide() {
      return {
        objects: [],
        SetLayout(layout) {
          this.layout = layout;
        },
        AddObject(shape) {
          this.objects.push(shape);
        }
      };
    },
    CreateNoFill() {
      return {};
    },
    CreateStroke() {
      return {};
    },
    CreateShape() {
      const content = {
        items: [],
        RemoveAllElements() {
          this.items = [];
        },
        Push(p) {
          this.items.push(p);
        }
      };
      return {
        SetPosition() {},
        GetDocContent() {
          return content;
        }
      };
    }
  };
  const changed = await global.Guiden.host.applyShapeUpdates(
    [{ shapeIndex: 1, before: "Body", after: "Clear\nBody" }],
    { targetLabel: "Slide 1" }
  );
  assert.strictEqual(changed.changed, 1);
  assert.strictEqual(contents[1].text, "ClearBody");
  const created = await global.Guiden.host.createSlide({ title: "New title", bullets: ["One", "Two"] });
  assert.strictEqual(created.changed, 1);
  assert(presentation.added);
});
queue("mocked host integration applies only sparse coordinates", async () => {
  let received;
  global.Guiden.host.applyCellUpdates = async updates => {
    received = updates;
    return { ok: true, changed: updates.length, target: "A1" };
  };
  const action = global.Guiden.actions["cells.edit"],
    snapshot = {
      rows: 1,
      cols: 1,
      values: [["alice"]],
      formulas: [[""]],
      start: { startRow: 1, startCol: 1 },
      address: "A1"
    },
    result = action.parseResponse(
      '{"narrative":"Cleaned","updates":[{"row":0,"column":0,"before":"alice","after":"Alice"}]}',
      snapshot
    );
  const applied = await action.apply(result, snapshot);
  assert.strictEqual(applied.changed, 1);
  assert.deepStrictEqual(received, result.updates);
});
queue("mocked slide integration preserves shape identity", async () => {
  let received;
  global.Guiden.host.applyShapeUpdates = async updates => {
    received = updates;
    return { ok: true, changed: 1, target: "Slide 1" };
  };
  const action = global.Guiden.actions["slides.rewrite"],
    snapshot = { target: "current_slide", shapes: [{ text: "Title" }, { text: "Body" }] },
    result = action.parseResponse('{"updates":[{"shapeIndex":1,"before":"Body","after":"Clear body"}]}', snapshot);
  await action.apply(result, snapshot);
  assert.strictEqual(received[0].shapeIndex, 1);
});
queue("API exposes only exact server retrieval metadata", async () => {
  const originalFetch = global.fetch;
  global.fetch = async () => ({
    ok: true,
    headers: {
      get() {
        return "application/json";
      }
    },
    json: async () => ({
      choices: [{ message: { content: '{"answer":"Grounded"}' } }],
      guiden: {
        retrieval_used: true,
        sources: [
          { title: "Trusted", url: "https://docs.example.com/source" },
          { title: "Unsafe", url: "javascript:alert(1)" },
          { title: "Invented shape", url: "docs.example.com/missing-scheme" }
        ]
      }
    })
  });
  try {
    const content = await global.Guiden.api.complete("test-key", [], new AbortController().signal);
    assert.strictEqual(content, '{"answer":"Grounded"}');
    assert.deepStrictEqual(global.Guiden.api.lastMetadata, {
      available: true,
      retrievalUsed: true,
      sources: [{ title: "Trusted", url: "https://docs.example.com/source" }]
    });
  } finally {
    global.fetch = originalFetch;
  }
});
queue("streaming API captures authoritative RAG metadata before model tokens", async () => {
  const originalFetch = global.fetch;
  const encoder = new TextEncoder();
  const events =
    'data: {"choices":[],"guiden":{"retrieval_used":true,"sources":[{"title":"RAG document","url":"https://docs.example.com/rag"}]}}\n\ndata: {"choices":[{"delta":{"content":"Grounded answer"}}]}\n\ndata: [DONE]\n\n';
  global.fetch = async () => ({
    ok: true,
    headers: {
      get() {
        return "text/event-stream";
      }
    },
    body: new ReadableStream({
      start(controller) {
        controller.enqueue(encoder.encode(events));
        controller.close();
      }
    })
  });
  try {
    const content = await global.Guiden.api.complete("test-key", [], new AbortController().signal);
    assert.strictEqual(content, "Grounded answer");
    assert.deepStrictEqual(global.Guiden.api.lastMetadata.sources, [
      { title: "RAG document", url: "https://docs.example.com/rag" }
    ]);
  } finally {
    global.fetch = originalFetch;
  }
});
queue("Writer apply defensively rejects needs-input results", async () => {
  const action = global.Guiden.actions["writer.draft"];
  await assert.rejects(
    action.apply({ needsInput: true, message: "More detail needed." }, { target: "cursor" }),
    /no validated document change/i
  );
});
queue("truncated streams without DONE are rejected before preview", async () => {
  const originalFetch = global.fetch,
    encoder = new TextEncoder(),
    events =
      "data: " +
      JSON.stringify({ choices: [{ delta: { content: '{"status":"ready","replacementHtml":"<p>Only part' } }] }) +
      "\n\n";
  global.fetch = async () => ({
    ok: true,
    headers: {
      get() {
        return "text/event-stream";
      }
    },
    body: new ReadableStream({
      start(controller) {
        controller.enqueue(encoder.encode(events));
        controller.close();
      }
    })
  });
  try {
    await assert.rejects(
      global.Guiden.api.complete("test-key", [], new AbortController().signal),
      /ended before completion/i
    );
  } finally {
    global.fetch = originalFetch;
  }
});
// Managed settings, timeouts, context limits, Copy text, sources, table themes and single-step undo.
test("managed settings come from the editor's plugin options", () => {
  assert.deepStrictEqual(
    C.managedSettings({
      options: { guidenApiKey: " k1 ", guidenUrl: "https://g.example.com/v1/chat", guidenModel: "M" }
    }),
    { key: "k1", url: "https://g.example.com/v1/chat", model: "M" }
  );
  assert.strictEqual(C.managedSettings({}).key, "");
  assert.strictEqual(C.managedSettings(undefined).key, "");
});
test("API endpoint accepts only https and falls back to the defaults", () => {
  const api = global.Guiden.api,
    defaults = api.configure({});
  assert.strictEqual(api.configure({ url: "http://insecure.example.com", model: "" }).url, defaults.url);
  assert.deepStrictEqual(api.configure({ url: "https://g.example.com/v1/chat", model: "Other" }), {
    url: "https://g.example.com/v1/chat",
    model: "Other"
  });
  api.configure({});
});
test("dark theme detection follows the editor theme", () => {
  assert(C.isDarkTheme({ type: "dark" }));
  assert(!C.isDarkTheme({ type: "light", "background-normal": "#000000" }));
  assert(C.isDarkTheme({ "background-normal": "#222222" }));
  assert(!C.isDarkTheme({ "background-normal": "#ffffff" }));
  assert(!C.isDarkTheme(null));
});
test("Copy gives the preview's text, not the raw JSON reply", () => {
  const raw = '{"status":"ready","replacementHtml":"<p>x</p>"}';
  assert.strictEqual(
    C.previewText(
      { kind: "writer-new", afterHtml: "<h2>Title</h2><p>Body <strong>text</strong>.</p><ul><li>One</li></ul>" },
      raw
    ),
    "Title\n\nBody text.\n\n• One"
  );
  assert.strictEqual(C.previewText({ kind: "formula", formula: "=SUM(A1:A3)", explanation: "x" }, raw), "=SUM(A1:A3)");
  assert.strictEqual(
    C.previewText(
      {
        kind: "table",
        table: [
          ["A", "B"],
          [1, 2]
        ]
      },
      raw
    ),
    "A\tB\n1\t2"
  );
  assert.strictEqual(C.previewText({ kind: "new-slide", title: "T", bullets: ["a", "b"] }, raw), "T\n\n• a\n\n• b");
  assert.strictEqual(C.previewText({ kind: "answer", answer: "Hello" }, raw), "Hello");
  assert.strictEqual(C.previewText({ kind: "unknown" }, raw), raw);
});
test("inserting an answer with no sources adds no empty Sources heading", () => {
  assert.strictEqual(C.sourcesHtml([]), "");
  assert.strictEqual(C.sourcesHtml([{ title: "bad", url: "javascript:x" }]), "");
  assert(C.sourcesHtml([{ title: "Doc", url: "https://example.com/d" }]).startsWith("<h3>Sources</h3>"));
});
test("long documents are cut for questions and refused for rewrites", () => {
  const limit = C.contextLimits.chars,
    long = "a".repeat(limit + 1000),
    snapshot = { editor: "word", target: "full_document", text: long, sourceHtml: "<p>x</p>" };
  const fitted = C.fitContext(snapshot, false);
  assert(fitted.truncated);
  assert(fitted.snapshot.text.length < long.length);
  assert.strictEqual(fitted.snapshot.sourceHtml, "");
  assert.strictEqual(snapshot.text, long, "the original snapshot is unchanged for revalidation");
  assert.throws(() => C.fitContext(snapshot, true), /too long/);
  assert.strictEqual(C.fitContext({ editor: "word", target: "full_document", text: "short" }, true).truncated, false);
});
test("large cell selections are cut for analysis and refused for edits", () => {
  const cols = 10,
    rows = Math.ceil(C.contextLimits.cells / cols) + 5,
    values = Array.from({ length: rows }, () => Array(cols).fill(1)),
    snapshot = { editor: "cell", rows, cols, values, formulas: values.map(r => r.map(() => "")), text: "" };
  const fitted = C.fitContext(snapshot, false);
  assert(fitted.truncated);
  assert(fitted.snapshot.rows * cols <= C.contextLimits.cells);
  assert.throws(() => C.fitContext(snapshot, true), /smaller range/);
});
function mockSheet() {
  const cells = {};
  const range = address => {
    const c = (cells[address] = cells[address] || { address, style: {} });
    return {
      SetValue(v) {
        c.value = v;
        return true;
      },
      SetFormula(f) {
        c.formula = f;
        return true;
      },
      SetFillColor(x) {
        c.style.fill = x;
      },
      SetFontColor(x) {
        c.style.fontColor = x;
      },
      SetBold(b) {
        c.style.bold = b;
      }
    };
  };
  const commands = { count: 0 };
  global.Api = {
    GetActiveSheet() {
      return { GetRange: range };
    },
    CreateColorFromRGB(r, g, b) {
      return [r, g, b]
        .map(n => n.toString(16).padStart(2, "0"))
        .join("")
        .toUpperCase();
    }
  };
  const callCommand = global.Asc.plugin.callCommand;
  global.Asc.plugin.callCommand = function (fn, a, b, cb) {
    commands.count += 1;
    return callCommand.call(this, fn, a, b, cb);
  };
  return {
    cells,
    commands,
    restore() {
      global.Asc.plugin.callCommand = callCommand;
    }
  };
}
queue("a calculated column is written in one command, so one Undo removes it", async () => {
  const sheet = mockSheet();
  try {
    global.Guiden.host.readRange = async () => ({ ok: true, values: [[""], [""], [""]], formulas: [[""], [""], [""]] });
    const action = global.Guiden.actions["cells.add_column"],
      snapshot = { rows: 3, cols: 1, start: { startRow: 1, startCol: 1, endCol: 1 } },
      result = action.parseResponse('{"header":"Total","formulas":["=A2*2","=A3*2"]}', snapshot);
    const applied = await action.apply(result, snapshot);
    assert.strictEqual(sheet.commands.count, 1);
    assert.strictEqual(applied.changed, 3);
    assert.strictEqual(sheet.cells.B1.value, "Total");
    assert.strictEqual(sheet.cells.B2.formula, "=A2*2");
    assert.strictEqual(sheet.cells.B3.formula, "=A3*2");
  } finally {
    sheet.restore();
  }
});
queue("a new table uses the chosen theme and is written in one command", async () => {
  const sheet = mockSheet();
  try {
    global.Guiden.host.readRange = async () => ({
      ok: true,
      values: [
        ["", ""],
        ["", ""]
      ],
      formulas: [
        ["", ""],
        ["", ""]
      ]
    });
    const action = global.Guiden.actions["cells.table"],
      snapshot = { start: { startRow: 1, startCol: 1 }, options: { tableTheme: "neutral" } },
      result = action.parseResponse('{"table":[["Name","Score"],["A",1]]}');
    await action.apply(result, snapshot);
    assert.strictEqual(sheet.commands.count, 1);
    assert.deepStrictEqual(sheet.cells.A1.style, { fill: "E5E7EB", fontColor: "111827", bold: true });
    assert.deepStrictEqual(sheet.cells.B1.style, { fill: "E5E7EB", fontColor: "111827", bold: true });
    const minimal = mockSheet();
    try {
      await action.apply(result, { start: { startRow: 1, startCol: 1 }, options: { tableTheme: "minimal" } });
      assert.deepStrictEqual(minimal.cells.A1.style, { bold: true });
    } finally {
      minimal.restore();
    }
  } finally {
    sheet.restore();
  }
});
queue("a request that never answers times out instead of hanging", async () => {
  const api = global.Guiden.api,
    originalFetch = global.fetch,
    first = api.firstByteMs;
  api.firstByteMs = 30;
  global.fetch = (_url, options) =>
    new Promise((_resolve, reject) =>
      options.signal.addEventListener("abort", () => {
        const e = new Error("aborted");
        e.name = "AbortError";
        reject(e);
      })
    );
  try {
    await assert.rejects(api.complete("k", [], new AbortController().signal), /stopped responding/);
  } finally {
    global.fetch = originalFetch;
    api.firstByteMs = first;
  }
});
queue("a stream that goes quiet times out, and Stop is still reported as Stop", async () => {
  const api = global.Guiden.api,
    originalFetch = global.fetch,
    idle = api.idleMs,
    encoder = new TextEncoder();
  api.idleMs = 30;
  global.fetch = async (_url, options) => ({
    ok: true,
    headers: {
      get() {
        return "text/event-stream";
      }
    },
    body: new ReadableStream({
      start(controller) {
        controller.enqueue(encoder.encode('data: {"choices":[{"delta":{"content":"{\\"answer\\":\\"Hal"}}]}\n\n'));
        options.signal.addEventListener("abort", () => {
          const e = new Error("aborted");
          e.name = "AbortError";
          controller.error(e);
        });
      }
    })
  });
  try {
    await assert.rejects(api.complete("k", [], new AbortController().signal), /stopped responding/);
    api.idleMs = 10000;
    const stop = new AbortController();
    const pending = api.complete("k", [], stop.signal);
    setTimeout(() => stop.abort(), 20);
    await assert.rejects(pending, error => error.name === "AbortError");
  } finally {
    global.fetch = originalFetch;
    api.idleMs = idle;
  }
});

const plain = text => ({ paragraphStyle: "", segments: [{ text }] });
queue("an edit of several selected paragraphs keeps each paragraph's own formatting", async () => {
  const editor = mockWriter(true, { ReplaceTextSmart: true });
  try {
    const payload = { supported: true, linkCount: 0, blocks: [plain("Un."), plain("Deux.")] };
    const snapshot = { target: "selection", text: "One.\r\nTwo.\r\n" };
    await withPayload(payload, () => global.Guiden.host.pasteWriter("<p>Un.</p><p>Deux.</p>", snapshot));
    assert.deepStrictEqual(editor.calls, [["ReplaceTextSmart", [["Un.", "Deux."]]]]);
    assert.strictEqual(editor.inserted.length, 0);
  } finally {
    editor.restore();
  }
});
queue("one new paragraph goes into the edited paragraph, in the document's formatting", async () => {
  const editor = mockWriter(true);
  try {
    const payload = { supported: true, linkCount: 0, blocks: [plain("Shorter.")] };
    const snapshot = { target: "selection", text: "One.\r\nTwo.\r\n" };
    await withPayload(payload, () => global.Guiden.host.pasteWriter("<p>Shorter.</p>", snapshot));
    assert.deepStrictEqual(editor.calls, [], "neither ReplaceTextSmart nor the HTML importer");
    assert.deepStrictEqual(editor.inline, [true]);
  } finally {
    editor.restore();
  }
});
queue("a draft at the cursor becomes its own paragraphs after the cursor's paragraph, with their styles", async () => {
  const editor = mockWriter(true, {}, { current: "Some text.\r\n" });
  try {
    const payload = {
      supported: true,
      linkCount: 0,
      blocks: [{ paragraphStyle: "Heading 2", segments: [{ text: "Next steps" }] }, plain("We ship in Q3.")]
    };
    await withPayload(payload, () => global.Guiden.host.pasteWriter("<h2>…</h2><p>…</p>", { target: "cursor" }));
    assert.deepStrictEqual(editor.calls, [], "no HTML importer");
    assert.strictEqual(editor.inserted.length, 0, "nothing merged into the text around the cursor");
    assert.deepStrictEqual(editor.after.map(text), ["Next steps", "We ship in Q3."]);
    assert.strictEqual(editor.after[0].style, "Heading 2");
    assert.strictEqual(editor.current.deleted, false, "the paragraph with text stays");
  } finally {
    editor.restore();
  }
});
queue("a draft in an empty paragraph takes that paragraph's place", async () => {
  const editor = mockWriter(true, {}, { current: "\r\n" });
  try {
    const payload = { supported: true, linkCount: 0, blocks: [plain("One line.")] };
    await withPayload(payload, () => global.Guiden.host.pasteWriter("<p>One line.</p>", { target: "cursor" }));
    assert.deepStrictEqual(editor.after.map(text), ["One line."]);
    assert.strictEqual(editor.current.deleted, true);
  } finally {
    editor.restore();
  }
});
queue("an answer inserted with nothing selected goes after the cursor's paragraph", async () => {
  const editor = mockWriter(true, { GetSelectedText: "" }, { current: "Intro.\r\n" });
  try {
    const payload = { supported: true, linkCount: 0, blocks: [plain("The answer.")] };
    await withPayload(payload, () => global.Guiden.host.insertWriter("<p>The answer.</p>"));
    assert.deepStrictEqual(editor.after.map(text), ["The answer."]);
    assert.strictEqual(editor.inserted.length, 0);
  } finally {
    editor.restore();
  }
});
const item = (words, kind, id, marker) => ({
  paragraphStyle: "List Paragraph",
  segments: [{ text: words }],
  list: { kind, id, marker }
});
const listPayload = {
  supported: true,
  linkCount: 0,
  blocks: [item("Apples", "bullet", 1, "• "), item("Pears", "bullet", 1, "• "), item("Step", "numbered", 2, "1. ")]
};
queue("lists become real bullet and numbered lists in the document's formatting", async () => {
  const editor = mockWriter(true);
  try {
    await withPayload(listPayload, () =>
      global.Guiden.host.pasteWriter("<ul>…</ul><ol>…</ol>", { target: "selection", text: "x" })
    );
    assert.deepStrictEqual(editor.calls, [], "no HTML importer");
    assert.deepStrictEqual(editor.inserted.map(text), ["Apples", "Pears", "Step"], "no typed bullet characters");
    assert.deepStrictEqual(
      editor.inserted.map(para => para.numbering),
      [
        { kind: "bullet", level: 0 },
        { kind: "bullet", level: 0 },
        { kind: "numbered", level: 0 }
      ]
    );
    assert.strictEqual(editor.inserted[0].numbering, editor.inserted[1].numbering, "one list per <ul>");
    assert.strictEqual(editor.inserted[0].style, "List Paragraph");
  } finally {
    editor.restore();
  }
});
queue("without real lists, list items start with a bullet or number", async () => {
  const editor = mockWriter(true, {}, { lists: false });
  try {
    await withPayload(listPayload, () =>
      global.Guiden.host.pasteWriter("<ul>…</ul><ol>…</ol>", { target: "selection", text: "x" })
    );
    assert.deepStrictEqual(editor.inserted.map(text), ["• Apples", "• Pears", "1. Step"]);
  } finally {
    editor.restore();
  }
});
queue("tables still go through the HTML importer", async () => {
  const editor = mockWriter(true);
  try {
    const html = "<table><tr><td>1</td></tr></table>";
    await withPayload({ supported: false, linkCount: 0, blocks: [] }, () =>
      global.Guiden.host.pasteWriter(html, { target: "selection", text: "x" })
    );
    assert.deepStrictEqual(editor.calls, [["PasteHtml", [html]]]);
  } finally {
    editor.restore();
  }
});
test("paragraphs in a selection are counted the way GetSelectedText returns them", () => {
  const count = global.Guiden.host.paragraphCount;
  assert.strictEqual(count("One.\r\nTwo.\r\n"), 2);
  assert.strictEqual(count("Part of a sentence"), 1);
  assert.strictEqual(count(""), 0);
});
queue("editor calls never overlap, so answers can't reach the wrong caller", async () => {
  const executeMethod = global.Asc.plugin.executeMethod;
  let active = 0,
    most = 0;
  global.Asc.plugin.executeMethod = (name, args, callback) => {
    active += 1;
    most = Math.max(most, active);
    setTimeout(() => {
      active -= 1;
      callback(name);
    }, 5);
  };
  try {
    const answers = await Promise.all(["A", "B", "C"].map(name => global.Guiden.host.execute(name, [])));
    assert.deepStrictEqual(answers, ["A", "B", "C"]);
    assert.strictEqual(most, 1);
  } finally {
    global.Asc.plugin.executeMethod = executeMethod;
  }
});
queue("an editor call that never answers fails instead of blocking the ones after it", async () => {
  const executeMethod = global.Asc.plugin.executeMethod,
    host = global.Guiden.host,
    timeout = host.callTimeoutMs;
  host.callTimeoutMs = 20;
  global.Asc.plugin.executeMethod = (name, args, callback) => {
    if (name === "Answered") callback("ok");
  };
  try {
    await assert.rejects(host.execute("Lost", []), /did not answer/);
    assert.strictEqual(await host.execute("Answered", []), "ok");
  } finally {
    global.Asc.plugin.executeMethod = executeMethod;
    host.callTimeoutMs = timeout;
  }
});
queue("Cells reads the selection as a target every Cells action accepts", async () => {
  const range = {
    GetAddress: () => "$A$1:$B$2",
    GetValue: () => [
      ["Region", "Q1"],
      ["north", 120]
    ],
    // Like the editor, GetFormula returns plain values too.
    GetFormula: () => [
      ["Region", "Q1"],
      ["north", "=B1*2"]
    ],
    GetRowsCount: () => 2,
    GetColumnsCount: () => 2
  };
  global.Api = { GetActiveSheet: () => ({ GetName: () => "Sales", GetSelection: () => range }) };
  const snapshot = await global.Guiden.host.readCells("cells.edit");
  assert.strictEqual(snapshot.target, "selection");
  assert.strictEqual(snapshot.address, "A1:B2");
  assert.deepStrictEqual(snapshot.formulas, [
    ["", ""],
    ["", "=B1*2"]
  ]);
  Object.values(global.Guiden.actions)
    .filter(action => action.editor === "cell")
    .forEach(action => assert(action.supportedTargets.includes(snapshot.target), action.id));
  const updates = C.validateUpdates([{ row: 1, column: 0, before: "north", after: "North" }], snapshot);
  assert.strictEqual(updates[0].after, "North", "a plain value can be cleaned");
  assert.throws(() => C.validateUpdates([{ row: 1, column: 1, before: 120, after: 1 }], snapshot), /replace a formula/);
});

const M = global.Guiden.menus;
test("every toolbar and right-click command opens an action of its own editor", () => {
  const ids = [];
  const collect = items =>
    items.forEach(item => {
      if (item.items) collect(item.items);
      else ids.push(item.id);
    });
  ["word", "cell", "slide"].forEach(editor => {
    collect(M.toolbar(editor)[0].items);
    ["Selection", "Target"].forEach(type => collect(M.contextMenu(editor, type)));
    ids.forEach(id => {
      const command = M.commandFor(id);
      assert(command, "unknown command " + id);
      if (command.open) return;
      const action = global.Guiden.actions[command.actionId];
      if (command.actionId.startsWith("knowledge.")) return; // registered by the panel (app.js)
      assert(action, id + " -> " + command.actionId);
      assert.strictEqual(action.editor, editor, id);
    });
    ids.length = 0;
  });
});
test("every one-click action in the panel resolves to an action for its editor", () => {
  const M = global.Guiden.menus;
  [
    ["word", true],
    ["word", false],
    ["cell", true],
    ["slide", true],
    ["slide", false]
  ].forEach(([editor, selected]) => {
    const items = M.quickActions(editor, selected);
    assert(items.length >= 4 && items.length <= 6, editor + " has a short bar");
    items.forEach(item =>
      (item.choices || [item]).forEach(choice => {
        const command = M.commandFor(choice.id);
        assert(command, choice.id);
        if (command.actionId.startsWith("knowledge.")) return; // registered by the panel (app.js)
        assert(global.Guiden.actions[command.actionId], choice.id);
        assert.strictEqual(global.Guiden.actions[command.actionId].editor, editor, choice.id);
      })
    );
  });
  assert.deepStrictEqual(
    M.quickActions("slide", false).map(i => i.text),
    ["Shorten slide", "Speaker notes", "New slide…", "Ask…"]
  );
});
test("the right-click menu fits what was clicked", () => {
  const labels = (editor, type) => {
    const menu = M.contextMenu(editor, type);
    return menu.length ? menu[0].items.map(item => item.text) : [];
  };
  assert.deepStrictEqual(labels("word", "Selection").slice(0, 4), [
    "Improve writing",
    "Fix spelling & grammar",
    "Summarise selection",
    "Translate"
  ]);
  assert.deepStrictEqual(labels("word", "Target"), [
    "Draft here…",
    "Review document with comments",
    "Ask about this document…",
    "Open Guiden"
  ]);
  assert(labels("word", "Selection").includes("Change tone"));
  assert(labels("slide", "Selection").includes("Change tone"));
  assert(labels("cell", "Selection").includes("Explain formulas"));
  assert.deepStrictEqual(M.commandFor("word.tone.formal"), {
    actionId: "writer.rewrite",
    run: true,
    instruction:
      "Rewrite it in a more formal, professional register: replace casual words, phrases and contractions with formal ones. Keep the meaning, names and numbers."
  });
  assert.strictEqual(M.commandFor("word.review").actionId, "writer.review");
  assert.strictEqual(M.commandFor("cell.explain").actionId, "cells.ask");
  assert.deepStrictEqual(labels("word", "Image"), [], "nothing for an image");
  assert(labels("cell", "Selection").includes("Clean up values"));
  assert(labels("slide", "Selection").includes("Improve text"));
  const french = M.commandFor("word.translate.French");
  assert.deepStrictEqual(french, {
    actionId: "writer.rewrite",
    run: true,
    instruction: "Translate into French. Keep the meaning, names, numbers, links and formatting."
  });
  assert.deepStrictEqual(M.commandFor("guiden.open"), { open: true });
  assert.strictEqual(M.commandFor("not.a.command"), null);
});

// background.js runs in the editor's background frame; this stands in for the runtime around it.
function loadBackground(editorType) {
  const calls = [],
    windows = [];
  function PluginWindow() {
    this.events = {};
    this.sent = [];
    windows.push(this);
  }
  PluginWindow.prototype.show = function (variation) {
    this.variation = variation;
  };
  PluginWindow.prototype.activate = function () {
    this.activated = (this.activated || 0) + 1;
  };
  PluginWindow.prototype.command = function (name, data) {
    this.sent.push([name, data]);
  };
  PluginWindow.prototype.attachEvent = function (name, fn) {
    this.events[name] = fn;
  };
  const saved = global.Asc;
  global.Asc = {
    PluginWindow,
    plugin: {
      guid: "asc.{guiden}",
      info: { editorType },
      executeMethod: (name, args) => calls.push([name, args[0]]),
      attachEditorEvent: name => calls.push(["attach", name])
    }
  };
  delete require.cache[require.resolve("../scripts/background.js")];
  require("../scripts/background.js");
  const plugin = global.Asc.plugin;
  global.Asc = saved;
  plugin.init();
  return { plugin, calls, windows };
}
test("in the background, Guiden adds its tab and a closed side panel when the editor starts", () => {
  const bg = loadBackground("word");
  const toolbar = bg.calls.find(call => call[0] === "AddToolbarMenuItem");
  assert.strictEqual(toolbar[1].guid, "asc.{guiden}");
  assert.strictEqual(toolbar[1].tabs[0].text, "Guiden");
  assert.strictEqual(bg.windows.length, 1);
  assert.strictEqual(bg.windows[0].variation.type, "panelRight");
  assert.strictEqual(bg.windows[0].variation.isActivated, false);
  assert(bg.calls.some(call => call[0] === "attach" && call[1] === "onSelectionChanged"));
});
test("a menu click opens the panel and the command waits until the panel is ready", () => {
  const bg = loadBackground("word"),
    panel = bg.windows[0];
  bg.plugin.event_onContextMenuClick("word.translate.French");
  assert.strictEqual(panel.activated, 1);
  assert.deepStrictEqual(panel.sent, [], "nothing is sent before the panel has loaded");
  panel.events.guidenReady();
  assert.strictEqual(panel.sent[0][0], "guidenRun");
  assert.strictEqual(panel.sent[0][1].actionId, "writer.rewrite");
  bg.plugin.event_onToolbarMenuClick("word.improve");
  assert.strictEqual(panel.sent[1][1].actionId, "writer.improve");
  bg.plugin.event_onToolbarMenuClick("guiden.open");
  assert.strictEqual(panel.sent.length, 2, "Open only opens the panel");
  assert.strictEqual(panel.activated, 3);
});
test("the background always answers a right-click, even with nothing to add", () => {
  const bg = loadBackground("word");
  bg.plugin.event_onContextMenuShow({ type: "Image" });
  bg.plugin.event_onContextMenuShow({ type: "Selection" });
  const answers = bg.calls.filter(call => call[0] === "AddContextMenuItem");
  assert.strictEqual(answers.length, 2);
  assert.deepStrictEqual(answers[0][1].items, []);
  assert.strictEqual(answers[1][1].items[0].text, "Guiden");
});
test("the background passes theme changes to the panel", () => {
  const bg = loadBackground("cell"),
    panel = bg.windows[0];
  panel.events.guidenReady();
  bg.plugin.onThemeChanged({ type: "dark", name: "theme-dark", "background-normal": "#333" });
  assert.deepStrictEqual(panel.sent.pop(), ["guidenTheme", { type: "dark", name: "theme-dark" }]);
});
// Word-level edits, applied the way the editor applies them.
function applyEdits(text, edits) {
  let out = text;
  for (let k = edits.length - 1; k >= 0; k -= 1) out = out.slice(0, edits[k][0]) + edits[k][2] + out.slice(edits[k][1]);
  return out;
}
test("word-level edits turn the old text into the new, joining a rewritten phrase into one change", () => {
  const W = global.Guiden.host.wordEdits;
  const pairs = [
    ["customers is happy with it", "customers are happy with it"],
    ["We will focus on reducing the cost of running it.", "We will focus on lower running costs."],
    ["", "All new."],
    ["All gone.", ""],
    ["Same text.", "Same text."],
    ["a b c", "x a b c y"]
  ];
  pairs.forEach(([before, after]) => assert.strictEqual(applyEdits(before, W(before, after)), after));
  assert.deepStrictEqual(W("customers is happy", "customers are happy"), [[10, 12, "are"]]);
  assert.deepStrictEqual(
    W("on reducing the cost", "on lower costs"),
    [[3, 20, "lower costs"]],
    "one change, not three"
  );
  assert.deepStrictEqual(W("Same.", "Same."), []);
  assert.strictEqual(W("x ".repeat(1000), "y ".repeat(1000)), null, "too long to compare");
});

// A tracked document: every character keeps its place, deleted or not, and inserted characters are marked.
function trackedDocument(paragraphTexts, selection) {
  const paragraphs = paragraphTexts.map(text => Array.from(text).map(ch => ({ ch })));
  const visible = chars =>
    chars
      .filter(c => !c.deleted)
      .map(c => c.ch)
      .join("");
  const original = chars =>
    chars
      .filter(c => !c.inserted)
      .map(c => c.ch)
      .join("");
  function rangeOver(chars, from) {
    return (start, end) => ({
      GetText: () =>
        chars
          .slice(from + start, from + end)
          .map(c => c.ch)
          .join(""),
      Delete() {
        chars.slice(from + start, from + end).forEach(c => (c.deleted = true));
        return true;
      },
      AddText(text, where) {
        const at = from + (where === "before" ? start : end);
        chars.splice(at, 0, ...Array.from(text).map(ch => ({ ch, inserted: true })));
        return true;
      }
    });
  }
  let tracking = false;
  const doc = {
    IsTrackRevisions: () => tracking,
    SetTrackRevisions(on) {
      tracking = on;
    },
    GetRangeBySelect() {
      if (selection.length === 1) {
        const p = paragraphs[selection[0].paragraph],
          from = selection[0].from,
          sub = rangeOver(p, from);
        return {
          GetText: () =>
            p
              .slice(from, from + selection[0].length)
              .map(c => c.ch)
              .join(""),
          GetRange: sub
        };
      }
      return {
        GetText: () => "",
        GetRange: () => null,
        GetAllParagraphs: () =>
          selection.map(item => {
            const p = paragraphs[item.paragraph];
            return { GetText: () => p.map(c => c.ch).join("") + "\r\n", GetRange: rangeOver(p, 0) };
          })
      };
    }
  };
  global.Api = { GetDocument: () => doc };
  return {
    doc,
    accepted: () => paragraphs.map(visible),
    rejected: () => paragraphs.map(original),
    tracking: () => tracking
  };
}
queue("tracked edits mark only the changed words, so rejecting them restores the original", async () => {
  const before = ["Our team have made progress.", "We will focus on reducing the cost of running it."],
    after = ["Our team has made progress.", "We will focus on lower running costs."];
  const tracked = trackedDocument(before, [{ paragraph: 0 }, { paragraph: 1 }]);
  tracked.doc.SetTrackRevisions(true);
  const result = await global.Guiden.host.dynamicCommand(global.Guiden.host.trackedEdits, { before, after }, [
    global.Guiden.host.wordEdits
  ]);
  assert.strictEqual(result.ok, true, result.error);
  assert.deepStrictEqual(tracked.accepted(), after);
  assert.deepStrictEqual(tracked.rejected(), before);
});
queue("tracked edits work on part of one paragraph, and refuse a selection that changed", async () => {
  const tracked = trackedDocument(["The figures is final."], [{ paragraph: 0, from: 4, length: 16 }]);
  tracked.doc.SetTrackRevisions(true);
  const H = global.Guiden.host;
  const ok = await H.dynamicCommand(H.trackedEdits, { before: ["figures is final"], after: ["figures are final"] }, [
    H.wordEdits
  ]);
  assert.strictEqual(ok.ok, true);
  assert.deepStrictEqual(tracked.accepted(), ["The figures are final."]);
  const stale = await H.dynamicCommand(H.trackedEdits, { before: ["something else"], after: ["x"] }, [H.wordEdits]);
  assert.strictEqual(stale.ok, false);
});
queue("Suggest as tracked changes turns tracking on for the edit and back off after", async () => {
  const tracked = trackedDocument(["The figures is final."], [{ paragraph: 0, from: 0, length: 21 }]);
  const doc = tracked.doc;
  doc.InsertContent = () => {
    throw new Error("the in-place tracked edit should have been used");
  };
  const ui = global.Guiden.ui;
  global.Guiden.ui = {
    writerNativePayload: () => ({
      supported: true,
      linkCount: 0,
      blocks: [{ paragraphStyle: "", segments: [{ text: "The figures are final." }] }]
    })
  };
  try {
    const snapshot = { target: "selection", text: "The figures is final.", targetLabel: "Selected text" };
    const result = await global.Guiden.host.applyWriter("<p>The figures are final.</p>", snapshot, { track: true });
    assert.strictEqual(result.tracked, true);
    assert.deepStrictEqual(tracked.accepted(), ["The figures are final."]);
    assert.deepStrictEqual(tracked.rejected(), ["The figures is final."]);
    assert.strictEqual(tracked.tracking(), false, "tracking is back off, as it was");
    doc.SetTrackRevisions(true);
    const warn = console.warn;
    console.warn = () => {};
    await global.Guiden.host
      .applyWriter("<p>x</p>", snapshot, { track: false })
      .catch(() => {})
      .finally(() => (console.warn = warn));
    assert.strictEqual(tracked.tracking(), true, "Guiden never turns off tracking someone else turned on");
  } finally {
    global.Guiden.ui = ui;
  }
});

test("a review keeps only comments that quote the text exactly", () => {
  const review = global.Guiden.actions["writer.review"];
  const snapshot = { text: "Our team have made progress. Costs is high.", targetLabel: "Whole document" };
  const result = review.parseResponse(
    JSON.stringify({
      status: "ready",
      comments: [
        { quote: "team have made", comment: "Use “has”." },
        { quote: "Costs is high", comment: "Use “are”." },
        { quote: "not in the text", comment: "Dropped." },
        { quote: "team have made", comment: "Use “has”." }
      ]
    }),
    snapshot
  );
  assert.deepStrictEqual(
    result.comments.map(c => c.quote),
    ["team have made", "Costs is high"]
  );
  assert.strictEqual(result.unmatched, 1);
  assert.match(review.buildPreview(result).warning, /1 comment was left out/);
  const overlapping = review.parseResponse(
    JSON.stringify({
      comments: [
        { quote: "Costs is high", comment: "Use “are”." },
        { quote: "Costs is high.", comment: "Say how high." },
        { quote: "team have", comment: "Fine in Irish English." }
      ]
    }),
    snapshot
  );
  assert.deepStrictEqual(overlapping.comments, [
    { quote: "Costs is high", comment: "Use “are”. Say how high." },
    { quote: "team have", comment: "Fine in Irish English." }
  ]);
  assert.deepStrictEqual(review.parseResponse('{"status":"ready","comments":[]}', snapshot).comments, []);
  assert.throws(
    () => review.parseResponse('{"comments":[{"quote":"nowhere","comment":"x"}]}', snapshot),
    /did not quote the document exactly/
  );
  const messages = review.buildRequest(snapshot, "");
  assert.match(messages[0].content, /copied exactly/);
  assert.strictEqual(messages[messages.length - 1].content, review.defaultInstruction);
});
queue("review comments go on the quoted text, within the selection when there is one", async () => {
  const added = [];
  const place = (scope, quote) => ({
    AddComment(text, author) {
      added.push([scope, quote, text, author]);
      return {};
    }
  });
  const search = scope => (quote, matchCase) => {
    assert.strictEqual(matchCase, true);
    if (quote === "twice") return [place(scope, "twice#1"), place(scope, "twice#2")];
    return quote === "gone" ? [] : [place(scope, quote)];
  };
  global.Api = {
    GetDocument: () => ({ Search: search("document"), GetRangeBySelect: () => ({ Search: search("selection") }) })
  };
  const comments = [
    { quote: "twice", comment: "First" },
    { quote: "twice", comment: "Second" },
    { quote: "gone", comment: "Lost" }
  ];
  const result = await global.Guiden.host.addComments(comments, { target: "selection", targetLabel: "Selected text" });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.changed, 2);
  assert.strictEqual(result.missing, 1);
  assert.deepStrictEqual(
    added.map(a => a.slice(0, 2).concat(a[3])),
    [
      ["selection", "twice#1", "Guiden"],
      ["selection", "twice#2", "Guiden"]
    ]
  );
  added.length = 0;
  await global.Guiden.host.addComments([{ quote: "x", comment: "y" }], { target: "full_document" });
  assert.strictEqual(added[0][0], "document");
});

queued
  .reduce((chain, [name, fn]) => chain.then(() => test(name, fn)), Promise.resolve())
  .then(() => process.stdout.write("\n" + passed + " tests passed.\n"))
  .catch(error => {
    console.error(error.stack || error);
    process.exit(1);
  });
