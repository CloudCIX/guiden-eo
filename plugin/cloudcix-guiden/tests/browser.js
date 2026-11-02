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
const doc = { text: setup.documentText || "", inserted: [], tracking: false };
const search = quote => ((setup.selection || doc.text).includes(quote)
  ? [{ AddComment(text, author) { window.editorCalls.push(["AddComment", quote, text, author]); return {}; } }]
  : []);
window.Api = {
  GetDocument: () => ({
    GetText: () => doc.text,
    InsertContent(content) { doc.inserted.push(content); window.editorCalls.push(["InsertContent", content.length]); return true; },
    Search: search,
    GetRangeBySelect: () => ({ Search: search, GetText: () => setup.selection || "", GetRange: () => null }),
    IsTrackRevisions: () => doc.tracking,
    SetTrackRevisions(on) { doc.tracking = on; window.editorCalls.push(["SetTrackRevisions", on]); }
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
      "writer.review",
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
  const setup = Object.assign({}, typedKey, { documentText: "Release notes for our software platform." });
  const t = await open(setup, sse([delta(json.slice(0, 20)), delta(json.slice(20))]));
  try {
    await generate(t.page, "Draft a plan");
    // The rest of the document goes along as context, so the draft fits what the document is about.
    assert.match(t.requests[0].body.messages[0].content, /for context only[^]*Release notes for our software platform/);
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

test("a light edit previews as marked changes, and Result shows the clean version", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>The figures are final.</p>" }))]);
  const t = await open(Object.assign({ selection: "The figures is final." }, typedKey), reply);
  try {
    await chooseAction(t.page, "writer.fix");
    await generate(t.page, "");
    assert.strictEqual(await text(t.page, "#preview del"), "is");
    assert.strictEqual(await text(t.page, "#preview ins"), "are");
    assert.strictEqual(await hidden(t.page, "#preview [data-view-panel=result]"), true);
    await t.page.click('#preview button[data-view="result"]');
    assert.strictEqual(await hidden(t.page, "#preview [data-view-panel=changes]"), true);
    assert.match(await text(t.page, "#preview [data-view-panel=result] .response-flow"), /The figures are final\./);
  } finally {
    await t.close();
  }
});

test("one click refines a result, carrying on the same conversation, and Apply uses the refined text", async () => {
  const first = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>A long, detailed intro.</p>" }))]);
  const second = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Short intro.</p>" }))]);
  const t = await open(Object.assign({ selection: "Intro." }, typedKey), n => (n === 1 ? first : second));
  try {
    await chooseAction(t.page, "writer.improve");
    await generate(t.page, "Improve it");
    assert.strictEqual(await hidden(t.page, "#refine"), false);
    await t.page.click('#refineChips button[data-refine="Make it shorter."]');
    await idle(t.page);
    const messages = t.requests[1].body.messages;
    assert.deepStrictEqual(
      messages.slice(-3).map(m => m.role),
      ["user", "assistant", "user"]
    );
    assert.strictEqual(messages[messages.length - 1].content, "Make it shorter.");
    assert.match(messages[messages.length - 2].content, /A long, detailed intro/);
    assert.strictEqual(messages[messages.length - 3].content, "Improve it");
    assert.match(await text(t.page, "#preview [data-view-panel=result] .response-flow"), /Short intro\./);
    await t.page.click("#apply");
    await idle(t.page);
    assert.strictEqual(await t.page.evaluate(() => window.nativeText.join("")), "Short intro.");
  } finally {
    await t.close();
  }
});

test("Guiden's question is answered in the follow-up box", async () => {
  const question = sse([delta(JSON.stringify({ status: "needs_input", message: "Who is the audience?" }))]);
  const draft = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Hello, partners.</p>" }))]);
  const t = await open(typedKey, n => (n === 1 ? question : draft));
  try {
    await generate(t.page, "Write an intro");
    assert.strictEqual(await hidden(t.page, "#refineChips"), true, "no tone chips for a question");
    assert.strictEqual(await t.page.$eval("#refineText", el => el.placeholder), "Answer Guiden's question…");
    await t.page.fill("#refineText", "Our partners");
    await t.page.press("#refineText", "Enter");
    await idle(t.page);
    assert.strictEqual(t.requests[1].body.messages.pop().content, "Our partners");
    assert.strictEqual(await hidden(t.page, "#apply"), false);
  } finally {
    await t.close();
  }
});

test("Up and Down in the instruction box step through earlier instructions", async () => {
  const reply = sse([delta(JSON.stringify({ answer: "Fine." }))]);
  const t = await open(typedKey, reply);
  try {
    await chooseAction(t.page, "writer.ask");
    await generate(t.page, "first question");
    await generate(t.page, "second question");
    await t.page.fill("#prompt", "");
    await t.page.focus("#prompt");
    await t.page.keyboard.press("ArrowUp");
    assert.strictEqual(await t.page.inputValue("#prompt"), "second question");
    await t.page.keyboard.press("Home");
    await t.page.keyboard.press("ArrowUp");
    assert.strictEqual(await t.page.inputValue("#prompt"), "first question");
    await t.page.keyboard.press("End");
    await t.page.keyboard.press("ArrowDown");
    assert.strictEqual(await t.page.inputValue("#prompt"), "second question");
  } finally {
    await t.close();
  }
});

test("Review with Comments lists the comments and adds them to the quoted text", async () => {
  const reply = sse([
    delta(
      JSON.stringify({
        status: "ready",
        comments: [
          { quote: "figures is final", comment: "Use “are”: figures is plural." },
          { quote: "made up", comment: "Dropped: not in the text." }
        ]
      })
    )
  ]);
  const t = await open(Object.assign({ selection: "The figures is final." }, typedKey), reply);
  try {
    await chooseAction(t.page, "writer.review");
    await generate(t.page, "");
    assert.match(await text(t.page, "#preview .change-count"), /^1 comment will be added/);
    assert.match(await text(t.page, "#targetWarning"), /1 comment was left out/);
    assert.strictEqual(await text(t.page, "#apply"), "Add comments");
    assert.strictEqual(await hidden(t.page, "#trackOption"), true);
    await t.page.click("#apply");
    await idle(t.page);
    const calls = await t.page.evaluate(() => window.editorCalls.filter(call => call[0] === "AddComment"));
    assert.deepStrictEqual(calls, [["AddComment", "figures is final", "Use “are”: figures is plural.", "Guiden"]]);
    assert.match(await text(t.page, "#status"), /Added 1 comment\. They are in the Comments pane\./);
  } finally {
    await t.close();
  }
});

test("Suggest as tracked changes is remembered and turns tracking on only for the edit", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>The figures are final.</p>" }))]);
  const t = await open(Object.assign({ selection: "The figures is final." }, typedKey), reply);
  try {
    await chooseAction(t.page, "writer.fix");
    await generate(t.page, "");
    assert.strictEqual(await hidden(t.page, "#trackOption"), false);
    await t.page.check("#trackChanges");
    await t.page.click("#apply");
    await idle(t.page);
    const tracking = await t.page.evaluate(() =>
      window.editorCalls.filter(call => call[0] === "SetTrackRevisions").map(call => call[1])
    );
    assert.deepStrictEqual(tracking, [true, false]);
    assert.match(await text(t.page, "#status"), /as tracked changes/);
    assert.strictEqual(await t.page.evaluate(() => localStorage.getItem("cloudcix-guiden.trackChanges")), "1");
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
