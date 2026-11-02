"use strict";
// Drives the real panel (index.html and every script) in Chromium against a fake Writer editor and a fake Guiden API.
// Run from the repository root after `npm ci && npx playwright install chromium`:
//   node plugin/cloudcix-guiden/tests/browser.js
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const PLUGIN = path.join(__dirname, "..");
const ORIGIN = "https://guiden.test";
const API = "https://inference.cloudcix.com/v1/guiden/chat/completions";
const TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png"
};

// Stands in for the Euro-Office plugin runtime (../v1/plugins.js) and a Writer document. Everything Guiden asks the
// editor to do is recorded in window.editorCalls. With setup.window the panel runs as the background plugin's side
// panel: window.emit(name, data) sends it a message from the background, and window.sentToPlugin records its replies.
const FAKE_RUNTIME = `
window.editorCalls = [];
window.sentToPlugin = [];
window.nativeText = [];
window.nativeLinks = [];
const setup = window.__setup || {};
const doc = { text: setup.documentText || "", inserted: [] };
window.Api = {
  GetDocument: () => ({
    GetText: () => doc.text,
    InsertContent(content) { doc.inserted.push(content); window.editorCalls.push(["InsertContent", content.length]); return true; }
  }),
  CreateParagraph: () => ({ SetStyle() {}, Push() {}, AddElement() {} }),
  CreateRun: () => ({ AddText(text) { window.nativeText.push(text); }, SetBold() {}, SetItalic() {}, AddLineBreak() {} }),
  CreateHyperlink: (href, text) => { window.nativeLinks.push(href); window.nativeText.push(text); return {}; }
};
window.Asc = {
  scope: {},
  plugin: {
    info: { editorType: "word", options: setup.options || {} },
    theme: setup.theme,
    windowID: setup.window ? "w1" : undefined,
    events: {},
    attachEvent(name, fn) { this.events[name] = fn; },
    sendToPlugin(name) { window.sentToPlugin.push(name); },
    attachEditorEvent() {},
    executeMethod(name, args, callback) {
      window.editorCalls.push([name, args]);
      const value = name === "GetSelectedText" ? setup.selection || "" : name === "GetSelectedHtml" ? setup.selectionHtml || "" : name === "CanUndo" ? true : undefined;
      setTimeout(() => callback && callback(value), 0);
    },
    callCommand(fn, _close, _recalc, callback) { const value = new Function("return (" + String(fn) + ")();")(); setTimeout(() => callback && callback(value), 0); }
  }
};
window.emit = (name, data) => window.Asc.plugin.events[name](data);`;

function sse(events, done) {
  return events.map(e => "data: " + JSON.stringify(e) + "\n\n").join("") + (done === false ? "" : "data: [DONE]\n\n");
}
function delta(content) {
  return { choices: [{ delta: { content } }] };
}

let browser;
async function open(setup, reply) {
  const context = await browser.newContext({ permissions: ["clipboard-read", "clipboard-write"] });
  const page = await context.newPage();
  const errors = [],
    requests = [];
  page.on("pageerror", error => errors.push(error.message));
  await page.addInitScript(value => {
    window.__setup = value;
    // A returning user: the key was connected and remembered earlier.
    if (value.storedKey && !localStorage.getItem("cloudcix-guiden.apiKey")) {
      localStorage.setItem("cloudcix-guiden.rememberKey", "1");
      localStorage.setItem("cloudcix-guiden.apiKey", value.storedKey);
    }
  }, setup || {});
  await page.route(ORIGIN + "/**", route => {
    const url = new URL(route.request().url());
    if (url.pathname === "/v1/plugins.js") return route.fulfill({ contentType: "text/javascript", body: FAKE_RUNTIME });
    const file = path.join(PLUGIN, url.pathname.replace(/^\/cloudcix-guiden\//, ""));
    if (!file.startsWith(PLUGIN) || !fs.existsSync(file)) return route.fulfill({ status: 404, body: "" });
    return route.fulfill({
      contentType: TYPES[path.extname(file)] || "application/octet-stream",
      body: fs.readFileSync(file)
    });
  });
  await page.route(API, async route => {
    const request = route.request();
    requests.push({ body: JSON.parse(request.postData()), auth: request.headers().authorization });
    // A reply is the stream body, or { status } for an HTTP error; a function may answer later with a Promise.
    const answer = await (typeof reply === "function" ? reply(requests.length, requests[requests.length - 1]) : reply);
    const status = answer && typeof answer === "object" ? answer.status : 200;
    return route
      .fulfill({
        status,
        headers: { "Content-Type": "text/event-stream", "Access-Control-Allow-Origin": "*" },
        body: status === 200 ? answer : ""
      })
      .catch(() => {}); // the panel may have stopped the request
  });
  await page.goto(ORIGIN + "/cloudcix-guiden/index.html");
  await page.waitForFunction(() => !/Connecting/.test(document.getElementById("contextText").textContent));
  return { page, errors, requests, close: () => context.close() };
}
async function chooseAction(page, id) {
  await page.click("#actionPicker summary");
  await page.click('button[data-action="' + id + '"]');
}
async function generate(page, prompt) {
  await page.fill("#prompt", prompt);
  await page.click("#generate");
  await page.waitForFunction(() => !document.body.classList.contains("is-busy"));
}
const text = (page, selector) => page.$eval(selector, el => el.textContent.trim());
const hidden = (page, selector) => page.$eval(selector, el => el.hidden || el.closest("[hidden]") !== null);
const typedKey = { options: {}, storedKey: "typed-key" };
const busy = page => page.evaluate(() => document.body.classList.contains("is-busy"));
const idle = page => page.waitForFunction(() => !document.body.classList.contains("is-busy"));
const later = (ms, value) => new Promise(resolve => setTimeout(() => resolve(value), ms));

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

test("Writer opens on Draft New Content with the action list collapsed, in the agreed order", async () => {
  const t = await open(typedKey, "");
  try {
    assert.strictEqual(await text(t.page, "#selectedActionLabel"), "Draft New Content");
    assert.strictEqual(await t.page.$eval("#actionPicker", el => el.open), false);
    const order = await t.page.$$eval("#actionGrid button", buttons => buttons.map(b => b.dataset.action));
    assert.deepStrictEqual(order, [
      "writer.draft",
      "writer.improve",
      "writer.rewrite",
      "writer.fix",
      "writer.summarise",
      "writer.ask",
      "knowledge.word",
      "writer.convert"
    ]);
    await chooseAction(t.page, "writer.ask");
    assert.strictEqual(await t.page.$eval("#actionPicker", el => el.open), false, "choosing an action closes the list");
    assert.deepStrictEqual(t.errors, []);
  } finally {
    await t.close();
  }
});

test("without a key, the connect card checks it, remembers it and runs the command that was waiting", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Better text.</p>" }))]);
  const t = await open({ window: true, selection: "Worse text." }, reply);
  try {
    assert.strictEqual(await hidden(t.page, "#connectCard"), false);
    assert.strictEqual(await t.page.isVisible("#generate"), false, "the work area waits for a key");
    assert.deepStrictEqual(await t.page.evaluate(() => window.sentToPlugin), ["guidenReady"]);
    await t.page.evaluate(() => window.emit("guidenRun", { actionId: "writer.improve", run: true, instruction: "" }));
    assert.match(await text(t.page, "#connectLead"), /improve writing straight after/);
    assert.strictEqual(t.requests.length, 0);
    await t.page.fill("#connectKey", "new-key");
    await t.page.press("#connectKey", "Enter");
    await t.page.waitForFunction(() => !document.getElementById("resultGroup").hidden);
    await idle(t.page);
    assert.strictEqual(await hidden(t.page, "#connectCard"), true);
    assert.strictEqual(t.requests[0].body.stream, false, "the key was checked first");
    assert.strictEqual(t.requests[1].auth, "Bearer new-key");
    assert.strictEqual(await text(t.page, "#selectedActionLabel"), "Improve Writing");
    assert.strictEqual(await t.page.evaluate(() => localStorage.getItem("cloudcix-guiden.apiKey")), "new-key");
    assert.deepStrictEqual(t.errors, []);
  } finally {
    await t.close();
  }
});

test("a rejected key brings the connect card back, and the request runs again once a new key works", async () => {
  const reply = sse([delta(JSON.stringify({ answer: "Fine." }))]);
  const t = await open(Object.assign({ window: true }, typedKey), (n, request) =>
    request.auth === "Bearer typed-key" ? { status: 401 } : reply
  );
  try {
    await chooseAction(t.page, "writer.ask");
    await generate(t.page, "Is this clear?");
    assert.strictEqual(await hidden(t.page, "#connectCard"), false);
    assert.match(await text(t.page, "#connectStatus"), /did not accept the saved API key/);
    await t.page.fill("#connectKey", "fresh-key");
    await t.page.click("#connectButton");
    await t.page.waitForFunction(() => /Fine\./.test(document.getElementById("preview").textContent));
    await idle(t.page);
    const last = t.requests[t.requests.length - 1];
    assert.strictEqual(last.auth, "Bearer fresh-key");
    assert.deepStrictEqual(last.body.messages[last.body.messages.length - 1].content, "Is this clear?");
  } finally {
    await t.close();
  }
});

test("Translate from the menu fills in the instruction and runs straight away", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Bonjour.</p>" }))]);
  const t = await open(Object.assign({ window: true, selection: "Hello." }, typedKey), reply);
  try {
    const instruction = "Translate into French. Keep the meaning, names, numbers, links and formatting.";
    await t.page.evaluate(
      i => window.emit("guidenRun", { actionId: "writer.rewrite", run: true, instruction: i }),
      instruction
    );
    await t.page.waitForFunction(() => !document.getElementById("resultGroup").hidden);
    await idle(t.page);
    assert.strictEqual(await t.page.inputValue("#prompt"), instruction);
    const messages = t.requests[0].body.messages;
    assert.strictEqual(messages[messages.length - 1].content, instruction);
    assert.strictEqual(await text(t.page, "#previewBadge"), "Will change file");
  } finally {
    await t.close();
  }
});

test("Ask from the menu waits for a question; Ctrl+Enter runs it and Esc stops it", async () => {
  const reply = sse([delta(JSON.stringify({ answer: "Done." }))]);
  const t = await open(Object.assign({ window: true }, typedKey), n => (n === 1 ? later(3000, reply) : reply));
  try {
    await t.page.evaluate(() => window.emit("guidenRun", { actionId: "writer.ask", run: false, instruction: "" }));
    assert.strictEqual(await text(t.page, "#selectedActionLabel"), "Ask");
    assert.match(await text(t.page, "#status"), /press (Ctrl|⌘)\+Enter/);
    assert.strictEqual(t.requests.length, 0);
    await t.page.fill("#prompt", "What is this about?");
    await t.page.press("#prompt", "Control+Enter");
    await t.page.waitForFunction(() => document.body.classList.contains("is-busy"));
    await t.page.press("#prompt", "Escape");
    await idle(t.page);
    assert.match(await text(t.page, "#status"), /Stopped/);
    await t.page.press("#prompt", "Control+Enter");
    await t.page.waitForFunction(() => !document.getElementById("resultGroup").hidden);
    await idle(t.page);
    assert.match(await text(t.page, "#preview"), /Done\./);
  } finally {
    await t.close();
  }
});

test("a command that arrives while Guiden is busy replaces the running request", async () => {
  const answer = text => sse([delta(JSON.stringify({ answer: text }))]);
  const asked = request => request.body.messages[request.body.messages.length - 1].content;
  const t = await open(Object.assign({ window: true }, typedKey), (n, request) =>
    asked(request) === "Second" ? answer("Second.") : later(3000, answer("First."))
  );
  try {
    await chooseAction(t.page, "writer.ask");
    await t.page.fill("#prompt", "First question");
    await t.page.click("#generate");
    await t.page.waitForFunction(() => document.body.classList.contains("is-busy"));
    await t.page.evaluate(() => window.emit("guidenRun", { actionId: "writer.ask", run: true, instruction: "Second" }));
    await t.page.waitForFunction(() => /Second\./.test(document.getElementById("preview").textContent));
    await idle(t.page);
    assert.strictEqual(await t.page.inputValue("#prompt"), "Second");
  } finally {
    await t.close();
  }
});

test("theme changes from the editor reach the panel", async () => {
  const t = await open(Object.assign({ window: true }, typedKey), "");
  try {
    assert.strictEqual(await t.page.evaluate(() => document.documentElement.classList.contains("theme-dark")), false);
    await t.page.evaluate(() => window.emit("guidenTheme", { type: "dark", name: "theme-dark" }));
    assert.strictEqual(await t.page.evaluate(() => document.documentElement.classList.contains("theme-dark")), true);
  } finally {
    await t.close();
  }
});

test("a key, endpoint model and theme from the editor are used, and the key box is hidden", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Hi</p>" }))]);
  const t = await open(
    { options: { guidenApiKey: "managed-key", guidenModel: "Managed-Model" }, theme: { type: "dark" } },
    reply
  );
  try {
    assert.strictEqual(await hidden(t.page, "#apiKeyField"), true);
    assert.strictEqual(await hidden(t.page, "#managedKeyNote"), false);
    assert.strictEqual(await t.page.evaluate(() => document.documentElement.classList.contains("theme-dark")), true);
    await generate(t.page, "Say hi");
    assert.strictEqual(t.requests[0].auth, "Bearer managed-key");
    assert.strictEqual(t.requests[0].body.model, "Managed-Model");
    await t.page.click("#settings summary");
    await t.page.click("#saveSettings");
    assert.strictEqual(
      await t.page.evaluate(() => localStorage.getItem("cloudcix-guiden.apiKey")),
      null,
      "a provided key is never stored"
    );
  } finally {
    await t.close();
  }
});

test("a streamed draft previews as prose, copies as text and applies at the cursor", async () => {
  const json = JSON.stringify({ status: "ready", replacementHtml: "<h2>Plan</h2><p>Step <strong>one</strong>.</p>" });
  const t = await open(typedKey, sse([delta(json.slice(0, 20)), delta(json.slice(20))]));
  try {
    await generate(t.page, "Draft a plan");
    assert.strictEqual(await t.page.$$eval("#preview article.response-flow", els => els.length), 1);
    assert.strictEqual(await hidden(t.page, "#apply"), false);
    assert.strictEqual(await text(t.page, "#previewBadge"), "Will change file");
    // The user message is only the instruction; editor context stays in the system message.
    const messages = t.requests[0].body.messages;
    assert.deepStrictEqual(messages[messages.length - 1], { role: "user", content: "Draft a plan" });
    assert.deepStrictEqual(t.requests[0].body.response_format, { type: "json_object" });
    await t.page.click("#copy");
    assert.strictEqual(await t.page.evaluate(() => navigator.clipboard.readText()), "Plan\n\nStep one.");
    await t.page.click("#apply");
    await t.page.waitForFunction(() => !document.body.classList.contains("is-busy"));
    // Built natively, so it takes the document's fonts instead of the HTML importer's.
    const calls = await t.page.evaluate(() => window.editorCalls);
    assert(!calls.some(call => call[0] === "PasteHtml"), "the HTML importer was not used");
    assert.deepStrictEqual(
      calls.find(call => call[0] === "InsertContent"),
      ["InsertContent", 2]
    );
    assert.strictEqual(await t.page.evaluate(() => window.nativeText.join("")), "PlanStep one.");
    assert.match(await text(t.page, "#status"), /Added the draft at the cursor/);
    assert.strictEqual(await hidden(t.page, "#undo"), false);
    // Lists are built natively too: one list per <ul>/<ol>, with a marker only for editors without real lists.
    const lists = await t.page.evaluate(() =>
      Guiden.ui
        .writerNativePayload("<ul><li>Apples</li><li>Pears</li></ul><ol><li>Wash</li></ol>")
        .blocks.map(block => [block.segments[0].text, block.paragraphStyle, block.list])
    );
    assert.deepStrictEqual(lists, [
      ["Apples", "List Paragraph", { kind: "bullet", id: 1, marker: "• " }],
      ["Pears", "List Paragraph", { kind: "bullet", id: 1, marker: "• " }],
      ["Wash", "List Paragraph", { kind: "numbered", id: 2, marker: "1. " }]
    ]);
  } finally {
    await t.close();
  }
});

test("changing a selection shows the original text behind a disclosure", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Better text.</p>" }))]);
  const t = await open(
    Object.assign({ selection: "Worse text.", selectionHtml: "<p>Worse text.</p>" }, typedKey),
    reply
  );
  try {
    await chooseAction(t.page, "writer.improve");
    await generate(t.page, "Improve it");
    assert.strictEqual(await text(t.page, "#preview .before-disclosure summary"), "View original text");
    assert.match(await text(t.page, "#preview .before-disclosure div"), /Worse text\./);
  } finally {
    await t.close();
  }
});

test("a needs_input reply and an incomplete stream can never be applied", async () => {
  const needsInput = sse([delta(JSON.stringify({ status: "needs_input", message: "Who is the audience?" }))]);
  const truncated = sse([delta('{"status":"ready","replacementHtml":"<p>Only part')], false);
  const t = await open(typedKey, n => (n === 1 ? needsInput : truncated));
  try {
    await generate(t.page, "Write an intro");
    assert.strictEqual(await text(t.page, "#previewBadge"), "Needs information");
    assert.strictEqual(await hidden(t.page, "#apply"), true);
    await generate(t.page, "Write an intro");
    assert.strictEqual(await text(t.page, "#previewBadge"), "Failed");
    assert.strictEqual(await hidden(t.page, "#apply"), true);
    assert.match(await text(t.page, "#status"), /ended before completion/);
  } finally {
    await t.close();
  }
});

test("a draft keeps only links Guiden's retrieval returned, and adds them as Sources", async () => {
  const reply = sse([
    {
      choices: [],
      guiden: { retrieval_used: true, sources: [{ title: "Handbook", url: "https://docs.example.com/handbook" }] }
    },
    delta(
      JSON.stringify({
        status: "ready",
        replacementHtml: '<p>See <a href="https://invented.example.com">this</a>.</p>'
      })
    )
  ]);
  const t = await open(typedKey, reply);
  try {
    await generate(t.page, "Summarise the handbook");
    await t.page.click("#apply");
    await t.page.waitForFunction(() => !document.body.classList.contains("is-busy"));
    // Links go in as native hyperlinks, so check what the editor was asked to create.
    const inserted = await t.page.evaluate(() => ({ links: window.nativeLinks, text: window.nativeText.join("") }));
    assert.deepStrictEqual(inserted.links, ["https://docs.example.com/handbook"], "a model-authored link is removed");
    assert(
      inserted.text.includes("See this.") && inserted.text.includes("Sources") && inserted.text.includes("Handbook")
    );
  } finally {
    await t.close();
  }
});

test("Company Knowledge withholds unverified answers and inserts verified ones", async () => {
  const none = sse([
    { choices: [], guiden: { retrieval_used: false, sources: [] } },
    delta(JSON.stringify({ answer: "Made up." }))
  ]);
  const found = sse([
    { choices: [], guiden: { retrieval_used: true, sources: [] } },
    delta(JSON.stringify({ answer: "From the corpus." }))
  ]);
  const legacy = sse([
    delta(JSON.stringify({ answer: "Legacy answer.", sources: [{ title: "Doc", url: "https://example.com/doc" }] }))
  ]);
  const t = await open(typedKey, n => [none, found, legacy][n - 1]);
  try {
    await chooseAction(t.page, "knowledge.word");
    await generate(t.page, "What is PodNet?");
    assert.match(await text(t.page, "#preview"), /withheld an unverified answer/);
    assert.strictEqual(await hidden(t.page, "#insertKnowledge"), true);

    await generate(t.page, "What is PodNet?");
    assert.match(await text(t.page, "#preview"), /From the corpus\./);
    assert.strictEqual(await hidden(t.page, "#insertKnowledge"), false);
    await t.page.click("#insertKnowledge");
    await t.page.waitForFunction(() => !document.body.classList.contains("is-busy"));
    const inserted = await t.page.evaluate(() => window.nativeText.join(""));
    assert(inserted.includes("From the corpus."));
    assert(!inserted.includes("Sources"), "no empty Sources heading");

    await generate(t.page, "What is PodNet?");
    assert.strictEqual(await text(t.page, "#previewBadge"), "Compatibility");
    assert.match(await text(t.page, "#sources"), /Links reported by Guiden/);
    assert.strictEqual(await hidden(t.page, "#insertKnowledge"), false);
  } finally {
    await t.close();
  }
});

(async () => {
  browser = await chromium.launch();
  let failed = 0;
  // GUIDEN_TEST=<words> runs only the tests whose name includes them.
  const only = process.env.GUIDEN_TEST || "";
  for (const [name, fn] of tests.filter(([name]) => name.includes(only))) {
    try {
      await fn();
      process.stdout.write("✓ " + name + "\n");
    } catch (error) {
      failed += 1;
      process.stdout.write("✗ " + name + "\n  " + (error.stack || error) + "\n");
    }
  }
  await browser.close();
  const ran = tests.filter(([name]) => name.includes(only)).length;
  process.stdout.write("\n" + (ran - failed) + " of " + ran + " browser tests passed.\n");
  process.exit(failed ? 1 : 0);
})();
