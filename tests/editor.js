"use strict";
// Smoke test in a real Euro-Office: opens a new document, spreadsheet and presentation in the image's example app and
// checks that Guiden loads in each, adds its toolbar tab, and opens its panel. CI runs it against the image it has just
// built, so a new Euro-Office version is checked before it is released. To run it by hand:
//   docker run -d --name guiden-eo-test -p 80:80 -e JWT_ENABLED=false -e EXAMPLE_ENABLED=true \
//     -e ALLOW_PRIVATE_IP_ADDRESS=true <image>
//   EDITOR_URL=http://localhost node tests/editor.js
// Use port 80 on both sides: the editor fetches the document from the address the browser uses.
// On failure it saves a screenshot and the page's console for each editor under EDITOR_ARTIFACTS (default
// editor-test-results/).
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

const BASE = (process.env.EDITOR_URL || "http://localhost").replace(/\/$/, "");
const OUT = process.env.EDITOR_ARTIFACTS || "editor-test-results";
const EDITORS = [
  ["Writer", "docx"],
  ["Cells", "xlsx"],
  ["Slides", "pptx"]
];

async function waitForHealth() {
  const until = Date.now() + 5 * 60 * 1000;
  while (Date.now() < until) {
    try {
      const res = await fetch(BASE + "/healthcheck");
      if (res.ok && (await res.text()).trim() === "true") return;
    } catch (e) {
      // not up yet
    }
    await new Promise(r => setTimeout(r, 3000));
  }
  throw new Error("The editor at " + BASE + " did not become healthy within 5 minutes");
}

// The frame of the editor itself, inside the example app's page.
async function editorFrame(page) {
  const until = Date.now() + 120000;
  while (Date.now() < until) {
    const frame = page
      .frames()
      .find(f => /\/web-apps\/apps\/(documenteditor|spreadsheeteditor|presentationeditor)\//.test(f.url()));
    if (frame) return frame;
    await page.waitForTimeout(1000);
  }
  throw new Error("The editor never loaded in " + page.url());
}

async function check(browser, name, ext) {
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
  const log = [];
  page.on("console", m => log.push(`[${m.type()}] ${m.text()}`));
  page.on("pageerror", e => log.push(`[pageerror] ${e.message}`));
  try {
    await page.goto(`${BASE}/example/editor?fileExt=${ext}&userid=uid-1&lang=en`, {
      waitUntil: "load",
      timeout: 120000
    });
    const frame = await editorFrame(page);

    // Guiden is a background plugin: its frame loads with the editor.
    const until = Date.now() + 120000;
    while (!page.frames().some(f => f.url().includes("/sdkjs-plugins/cloudcix-guiden/background.html"))) {
      if (Date.now() > until) throw new Error("Guiden's background frame never loaded");
      await page.waitForTimeout(1000);
    }

    // Its toolbar tab, then the panel from the tab's Guiden button.
    const tab = frame
      .locator(".ribtab, .tabs li, [data-tab]")
      .filter({ hasText: /^\s*Guiden\s*$/ })
      .first();
    await tab.waitFor({ state: "visible", timeout: 60000 });
    await tab.click();
    const open = frame
      .getByRole("button", { name: /^Guiden$/ })
      .or(frame.locator("[data-hint] :text-is('Guiden')"))
      .first();
    await open.click({ timeout: 30000 });
    const panelUntil = Date.now() + 60000;
    let panel;
    while (!(panel = page.frames().find(f => f.url().includes("/sdkjs-plugins/cloudcix-guiden/index.html")))) {
      if (Date.now() > panelUntil) throw new Error("The Guiden panel never opened");
      await page.waitForTimeout(1000);
    }
    await panel.locator("body").waitFor({ state: "visible", timeout: 30000 });
    const pluginErrors = log.filter(l => /\[pageerror\]|\[error\]/.test(l) && /cloudcix-guiden/.test(l));
    if (pluginErrors.length) throw new Error("Guiden logged errors:\n" + pluginErrors.join("\n"));
    console.log(`✓ ${name}: Guiden loads, adds its toolbar tab and opens its panel`);
  } catch (e) {
    fs.mkdirSync(OUT, { recursive: true });
    await page.screenshot({ path: path.join(OUT, `${name}.png`), fullPage: true }).catch(() => {});
    fs.writeFileSync(
      path.join(OUT, `${name}.log`),
      [`url: ${page.url()}`, "frames:", ...page.frames().map(f => "  " + f.url()), "console:", ...log].join("\n")
    );
    // CI logs are easier to reach than its artifacts, so print what the page held too.
    console.error(
      `--- ${name}: ${page.url()}\nframes:\n${page
        .frames()
        .map(f => "  " + f.url())
        .join("\n")}`
    );
    console.error(`console (last 40):\n${log.slice(-40).join("\n")}`);
    throw new Error(`${name}: ${e.message}`);
  } finally {
    await page.close();
  }
}

(async () => {
  await waitForHealth();
  const browser = await chromium.launch();
  const failures = [];
  try {
    for (const [name, ext] of EDITORS) {
      try {
        await check(browser, name, ext);
      } catch (e) {
        console.error("✗ " + e.message);
        failures.push(name);
      }
    }
  } finally {
    await browser.close();
  }
  if (failures.length) {
    console.error(`${failures.length} of ${EDITORS.length} editors failed; see ${OUT}/`);
    process.exit(1);
  }
  console.log(`${EDITORS.length} editors passed.`);
})().catch(e => {
  console.error(e);
  process.exit(1);
});
