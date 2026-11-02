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
// editor to do is recorded in window.editorCalls.
const FAKE_RUNTIME = `
window.editorCalls = [];
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
    attachEditorEvent() {},
    executeMethod(name, args, callback) {
      window.editorCalls.push([name, args]);
      const value = name === "GetSelectedText" ? setup.selection || "" : name === "GetSelectedHtml" ? setup.selectionHtml || "" : name === "CanUndo" ? true : undefined;
      setTimeout(() => callback && callback(value), 0);
    },
    callCommand(fn, _close, _recalc, callback) { const value = new Function("return (" + String(fn) + ")();")(); setTimeout(() => callback && callback(value), 0); }
  }
};`;

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
  await page.route(API, route => {
    const request = route.request();
    requests.push({ body: JSON.parse(request.postData()), auth: request.headers().authorization });
    const body = typeof reply === "function" ? reply(requests.length) : reply;
    return route.fulfill({
      status: 200,
      headers: { "Content-Type": "text/event-stream", "Access-Control-Allow-Origin": "*" },
      body
    });
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
const typedKey = { options: {} };
async function typeKey(page) {
  await page.click("#settings summary");
  await page.fill("#apiKey", "typed-key");
}

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

test("without a key, Guiden asks for one and opens Settings", async () => {
  const t = await open(typedKey, "");
  try {
    await generate(t.page, "Write something");
    assert.match(await text(t.page, "#status"), /API key/);
    assert.strictEqual(await t.page.$eval("#settings", el => el.open), true);
    assert.strictEqual(t.requests.length, 0);
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
    await typeKey(t.page);
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
    const paste = (await t.page.evaluate(() => window.editorCalls)).find(call => call[0] === "PasteHtml");
    assert(paste && paste[1][0].includes("<h2>Plan</h2>"), "the draft was pasted");
    assert.match(await text(t.page, "#status"), /Updated Current cursor/);
    assert.strictEqual(await hidden(t.page, "#undo"), false);
  } finally {
    await t.close();
  }
});

test("changing a selection shows the original text behind a disclosure", async () => {
  const reply = sse([delta(JSON.stringify({ status: "ready", replacementHtml: "<p>Better text.</p>" }))]);
  const t = await open({ selection: "Worse text.", selectionHtml: "<p>Worse text.</p>" }, reply);
  try {
    await typeKey(t.page);
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
    await typeKey(t.page);
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
    await typeKey(t.page);
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
    await typeKey(t.page);
    await chooseAction(t.page, "knowledge.word");
    await generate(t.page, "What is PodNet?");
    assert.match(await text(t.page, "#preview"), /withheld an unverified answer/);
    assert.strictEqual(await hidden(t.page, "#insertKnowledge"), true);

    await generate(t.page, "What is PodNet?");
    assert.match(await text(t.page, "#preview"), /From the corpus\./);
    assert.strictEqual(await hidden(t.page, "#insertKnowledge"), false);
    await t.page.click("#insertKnowledge");
    await t.page.waitForFunction(() => !document.body.classList.contains("is-busy"));
    const paste = (await t.page.evaluate(() => window.editorCalls)).filter(call => call[0] === "PasteHtml").pop();
    assert(paste && paste[1][0].includes("From the corpus."));
    assert(!paste[1][0].includes("Sources"), "no empty Sources heading");

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
  for (const [name, fn] of tests) {
    try {
      await fn();
      process.stdout.write("✓ " + name + "\n");
    } catch (error) {
      failed += 1;
      process.stdout.write("✗ " + name + "\n  " + (error.stack || error) + "\n");
    }
  }
  await browser.close();
  process.stdout.write("\n" + (tests.length - failed) + " of " + tests.length + " browser tests passed.\n");
  process.exit(failed ? 1 : 0);
})();
