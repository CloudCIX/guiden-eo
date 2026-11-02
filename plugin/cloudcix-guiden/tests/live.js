"use strict";
// Sends Guiden's real requests to the real Guiden API and checks the replies the way the panel does, so a prompt or
// model change that breaks an action shows up here. It needs a CloudCIX API key and costs a few requests.
//
//   GUIDEN_API_KEY=... node plugin/cloudcix-guiden/tests/live.js [report.json]
//
// GUIDEN_URL and GUIDEN_MODEL override the endpoint and model. GUIDEN_TEST=<words> runs only the matching cases.
// GUIDEN_RUNS=<n> runs every case n times, because Guiden's replies vary from one request to the next. The report
// holds each request's messages, the raw reply, whether Guiden used retrieval, and every check, and never the key.
const assert = require("assert");
const fs = require("fs");
const path = require("path");

global.window = global;
global.Asc = { plugin: { callCommand() {}, executeMethod() {} } };
const C = require("../scripts/core.js");
["host", "writer", "cells", "slides", "api", "menus"].forEach(name => require("../scripts/" + name + ".js"));
const G = global.Guiden;

const key = process.env.GUIDEN_API_KEY || "";
if (!key) {
  console.error("Set GUIDEN_API_KEY to a CloudCIX API key.");
  process.exit(2);
}
G.api.configure({ url: process.env.GUIDEN_URL, model: process.env.GUIDEN_MODEL });
const reportPath = path.resolve(process.argv[2] || "guiden-live-report.json");
const OPTIONS = {
  tone: "auto",
  length: "auto",
  structure: "auto",
  destination: "selection",
  tableTheme: "business",
  slideCount: "1",
  keepCitations: true
};

// Snapshots shaped like the ones host.js reads from the editor.
function writer(actionId, text, target, contextText) {
  return {
    editor: "word",
    actionId,
    target,
    targetLabel: target === "selection" ? "Selected text" : target === "cursor" ? "Current cursor" : "Whole document",
    text,
    contextText: contextText || "",
    sourceHtml: "",
    options: OPTIONS
  };
}
function cells(actionId, values, formulas) {
  const rows = values.length,
    cols = values[0].length;
  return {
    editor: "cell",
    actionId,
    sheetName: "Sheet1",
    address: C.rangeAddress({ startRow: 1, startCol: 1 }, rows, cols),
    start: C.parseA1Address(C.rangeAddress({ startRow: 1, startCol: 1 }, rows, cols)),
    rows,
    cols,
    values,
    formulas: formulas || values.map(row => row.map(() => "")),
    target: "selection",
    targetLabel: "Sheet1 · " + C.rangeAddress({ startRow: 1, startCol: 1 }, rows, cols),
    text: values.map(row => row.join("\t")).join("\n"),
    options: OPTIONS
  };
}
function slide(actionId, shapes, selected) {
  const target = selected ? "selected_text" : "current_slide";
  return {
    editor: "slide",
    actionId,
    target,
    targetLabel: selected ? "Selected slide text" : "Slide 1",
    slideIndex: 0,
    selectedText: selected || "",
    shapes: shapes.map((text, index) => ({ index, text })),
    text: selected || shapes.map((text, index) => "[" + index + "] " + text).join("\n\n"),
    options: OPTIONS
  };
}
// Paragraph-level blocks in replacement HTML: the number of paragraphs the edit would leave.
function blocks(html) {
  return (String(html || "").match(/<(p|h[1-6]|li|blockquote)\b/gi) || []).length;
}
// -ize spellings such as "prioritize"; Guiden is asked for Irish English, which spells them "prioritise".
const AMERICAN = /\b\w{3,}iz(?:e|es|ed|ing|ation)\b/i;
// The source says what "we will" do next quarter; an edit that keeps the meaning keeps the future tense.
const FUTURE = /\bwill\b|['’]ll\b/i;
function words(text) {
  return String(text || "")
    .replace(/<[^>]+>/g, " ")
    .split(/\s+/)
    .filter(Boolean).length;
}

const DOC = [
  "Our team have made good progress this quarter, alot of work was done on the new platfrom and customers is happy with it.",
  "Next quarter we will focus on reliability and on reducing the cost of running the service."
];
const CLEAN = ["Our team have made good progress this quarter.", DOC[1]];
const SALES = [
  ["Region", "Q1", "Q2"],
  ["north", "120", "135"],
  ["South ", "98", "101"],
  ["EAST", "143", "150"],
  ["west", "87", "90"]
];

// Each case runs one action. check(result, context) throws when the reply would not work in the editor; notes it
// returns are written to the report.
const cases = [
  {
    name: "Writer Fix keeps the two paragraphs, so tracked changes are word by word",
    actionId: "writer.fix",
    snapshot: writer("writer.fix", DOC.join("\r\n") + "\r\n", "selection"),
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert.strictEqual(blocks(result.replacementHtml), 2, "paragraph count changed: " + result.replacementHtml);
      assert(!/\balot\b|platfrom/.test(result.replacementHtml), "left a spelling mistake");
      return {
        blocks: 2,
        edits: G.host.wordEdits(DOC.join(" "), C.htmlText(result.replacementHtml)).length,
        keptIrishTeamHave: /team have/.test(result.replacementHtml)
      };
    }
  },
  {
    name: "Writer Improve returns one paragraph for one",
    actionId: "writer.improve",
    snapshot: writer("writer.improve", DOC[1], "selection"),
    followUp: "Make it shorter.",
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert.strictEqual(blocks(result.replacementHtml), 1, "not one paragraph: " + result.replacementHtml);
      assert(FUTURE.test(C.htmlText(result.replacementHtml)), "changed the tense: " + result.replacementHtml);
      return { words: words(result.replacementHtml) };
    },
    checkFollowUp(result, first) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert(
        words(result.replacementHtml) < words(first.replacementHtml),
        "not shorter: " + result.replacementHtml + " vs " + first.replacementHtml
      );
      return { words: words(result.replacementHtml) };
    }
  },
  {
    name: "Writer Change tone (More formal) rewrites the selection",
    actionId: "writer.rewrite",
    snapshot: writer("writer.rewrite", DOC[0], "selection"),
    instruction: G.menus.commandFor("word.tone.formal").instruction,
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert(blocks(result.replacementHtml) >= 1, "empty");
      // "a lot of" and "happy" are casual; a more formal rewrite replaces them rather than only fixing the grammar.
      assert(!/\ba ?lot of\b|\bhappy\b/i.test(result.replacementHtml), "still casual: " + result.replacementHtml);
      // The source says "good progress"; a tone change should not turn it into a bigger claim.
      assert(
        !/\b(?:significant|substantial|excellent|remarkable|outstanding|considerable|major|great|strong) progress/i.test(
          result.replacementHtml
        ),
        "overstated the progress: " + result.replacementHtml
      );
      return { html: result.replacementHtml };
    }
  },
  {
    name: "Writer Translate into French keeps the two paragraphs",
    actionId: "writer.rewrite",
    snapshot: writer("writer.rewrite", CLEAN.join("\r\n") + "\r\n", "selection"),
    instruction: G.menus.commandFor("word.translate.French").instruction,
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert(/\b(le|la|les|nous|des|du|de)\b/i.test(C.htmlText(result.replacementHtml)), "not French");
      assert.strictEqual(blocks(result.replacementHtml), 2, "paragraph count changed: " + result.replacementHtml);
      return { html: result.replacementHtml };
    }
  },
  {
    name: "Writer Review with Comments quotes the text exactly",
    actionId: "writer.review",
    snapshot: writer("writer.review", "Quarterly update\r\n" + DOC.join("\r\n") + "\r\n", "full_document"),
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert(result.comments.length >= 2, "fewer than 2 comments for a text with obvious mistakes");
      const irish = result.comments.filter(c => /team have/.test(c.quote) && /\bhas\b|singular/.test(c.comment));
      assert(!irish.length, "flagged Irish English 'team have': " + JSON.stringify(irish));
      const passive = result.comments.filter(c => /passive/i.test(c.comment) && !/was done/.test(c.quote));
      assert(!passive.length, "called active text passive: " + JSON.stringify(passive));
      const overlapping = result.comments.filter((c, i) =>
        result.comments.some((d, j) => i !== j && d.quote !== c.quote && d.quote.includes(c.quote))
      );
      assert(!overlapping.length, "overlapping quotes: " + JSON.stringify(overlapping));
      const long = result.comments.filter(c => words(c.quote) > 12);
      return {
        comments: result.comments.length,
        unmatched: result.unmatched,
        longQuotes: long.length,
        sample: result.comments.slice(0, 3)
      };
    }
  },
  {
    name: "Writer Draft here writes headings and paragraphs that fit the document",
    actionId: "writer.draft",
    // The document around the cursor is about a software service, so "shipping the platform" means releasing it.
    snapshot: writer(
      "writer.draft",
      "",
      "cursor",
      "Quarterly update\n\n" +
        CLEAN.join("\n\n") +
        "\n\nThe platform is the new version of our hosted software service."
    ),
    instruction:
      "Write a short 'Next steps' section with a heading and three bullet points about shipping the platform.",
    check(result) {
      assert(!result.needsInput, "asked for input: " + result.message);
      assert(/<h[1-3]\b/i.test(result.replacementHtml), "no heading");
      assert(/<li\b/i.test(result.replacementHtml), "no list");
      assert(!/logistic|freight|carrier|warehouse|shipment/i.test(result.replacementHtml), "read as physical shipping");
      return { html: result.replacementHtml };
    }
  },
  {
    name: "Writer Draft with missing facts asks a question instead of inventing them",
    actionId: "writer.draft",
    snapshot: writer("writer.draft", "", "cursor"),
    instruction: "Write the press release announcing our new product, with its name, price and launch date.",
    followUp: "The product is CloudCIX Guiden for Euro-Office, it is free for members, and it launches on 1 December.",
    check(result) {
      // Either answer is acceptable; the report shows which the model chose.
      return result.needsInput ? { asked: result.message } : { drafted: words(result.replacementHtml) + " words" };
    },
    checkFollowUp(result) {
      assert(!result.needsInput, "asked again: " + result.message);
      assert(/Guiden/.test(result.replacementHtml), "did not use the answer");
      const text = C.htmlText(result.replacementHtml);
      assert(!/mailto:|[\w.+-]+@[\w-]+\.[a-z]{2,}/i.test(result.replacementHtml), "made up an email contact");
      assert(!/\+\d{2,3}[\s\d()-]{6,}|\b0\d{1,2}[\s-]?\d{3}[\s-]?\d{4}\b/.test(text), "made up a phone number");
      assert(!/[“"][^”"]{15,}[”"],?\s+(?:said|says|added)\b/i.test(text), "made up a quote");
      assert(!/\b(?:Dublin|Cork|Galway|Limerick)\b/.test(text), "made up a dateline city");
      assert(!/credentials|log ?in with|sign ?in with/i.test(text), "made up how people sign in");
      assert(!/\bimmediately\b/i.test(text), "said available immediately, but it launches on 1 December");
      return { words: words(result.replacementHtml), placeholders: text.match(/\[[^\]]{3,60}\]/g) || [] };
    }
  },
  {
    name: "Writer Ask answers about the document",
    actionId: "writer.ask",
    snapshot: writer("writer.ask", DOC.join("\r\n"), "full_document"),
    instruction: "What will the team focus on next quarter?",
    followUp: "And what went well this quarter?",
    check(result) {
      assert(/reliab/i.test(result.answer), "missed reliability: " + result.answer);
      return { answer: result.answer };
    },
    checkFollowUp(result) {
      assert(/platform|progress|customer/i.test(result.answer), "missed the progress: " + result.answer);
      return { answer: result.answer };
    }
  },
  {
    name: "Cells Clean up values updates only the messy cells",
    actionId: "cells.edit",
    snapshot: cells("cells.edit", SALES),
    instruction: "Clean up the region names: trim spaces and use title case.",
    check(result) {
      assert(result.updates.length >= 3, "too few updates: " + JSON.stringify(result.updates));
      return { updates: result.updates.map(u => u.before + " → " + u.after) };
    }
  },
  {
    name: "Cells Explain formulas",
    actionId: "cells.ask",
    snapshot: cells(
      "cells.ask",
      [
        ["Net", "VAT", "Gross"],
        ["100", "0.23", "123"]
      ],
      [
        ["", "", ""],
        ["", "", "=A2*(1+B2)"]
      ]
    ),
    instruction: G.menus.commandFor("cell.explain").instruction,
    check(result) {
      assert(/VAT|tax|23|B2/i.test(result.answer), "did not explain the formula: " + result.answer);
      assert(/1\.23/.test(result.answer) && /\b123\b/.test(result.answer), "did not work it through: " + result.answer);
      return { answer: result.answer };
    }
  },
  {
    name: "Slides Improve selected text",
    actionId: "slides.improve",
    snapshot: slide(
      "slides.improve",
      ["Q3 plan", "we will focus on reliablity and lower costs"],
      "we will focus on reliablity and lower costs"
    ),
    check(result) {
      assert(result.replacementText.trim(), "empty");
      assert(!/reliablity/.test(result.replacementText), "left the typo");
      assert(/\bwill\b|['’]ll\b/i.test(result.replacementText), "changed the tense: " + result.replacementText);
      assert(/cost/i.test(result.replacementText), "dropped the costs point: " + result.replacementText);
      assert(!AMERICAN.test(result.replacementText), "American spelling: " + result.replacementText);
      return { text: result.replacementText };
    }
  },
  {
    name: "Slides Write speaker notes",
    actionId: "slides.notes",
    snapshot: slide("slides.notes", ["Q3 plan", "Reliability first\nLower running costs\nShip the platform"]),
    check(result) {
      const notes = result.notes;
      assert(!/<[a-z][^>]*>|\*\*|^#/im.test(notes), "Markdown or HTML in the notes: " + notes);
      assert(/reliab/i.test(notes) && /cost/i.test(notes) && /platform/i.test(notes), "missed a point: " + notes);
      assert(words(notes) >= 20 && words(notes) <= 250, "odd length: " + words(notes) + " words");
      assert(!/\d+\s?%|€\s?\d/.test(notes), "made up a figure: " + notes);
      assert(!AMERICAN.test(notes), "American spelling: " + notes);
      return { words: words(notes), notes };
    }
  }
];

async function ask(action, snapshot, messages) {
  const started = Date.now();
  let raw = await G.api.complete(key, messages, new AbortController().signal);
  let result,
    retried = false;
  try {
    result = action.parseResponse(raw, snapshot);
  } catch (error) {
    // The panel asks once more with a stricter format reminder; so does this test, and the report says so.
    retried = error.message;
    raw = await G.api.complete(key, C.schemaRetryMessages(messages), new AbortController().signal);
    result = action.parseResponse(raw, snapshot);
  }
  const metadata = G.api.lastMetadata || {};
  return {
    result,
    raw,
    retried,
    ms: Date.now() - started,
    // available is false when the server sent no retrieval metadata at all, which is different from "not used".
    retrieval: {
      available: Boolean(metadata.available),
      used: Boolean(metadata.retrievalUsed),
      sources: metadata.sources || []
    }
  };
}

function label(item, run, runs) {
  return runs > 1 ? item.name + " [run " + run + "]" : item.name;
}

(async () => {
  const only = process.env.GUIDEN_TEST || "",
    runs = Math.max(1, parseInt(process.env.GUIDEN_RUNS || "1", 10) || 1);
  const report = { when: new Date().toISOString(), url: G.api.url, model: G.api.model, runs, cases: [] };
  let failed = 0;
  const selected = cases.filter(c => c.name.includes(only));
  for (let run = 1; run <= runs; run += 1)
    for (const item of selected) {
      const action = G.actions[item.actionId],
        entry = { name: item.name, run, actionId: item.actionId, steps: [] };
      report.cases.push(entry);
      try {
        // An empty prompt sends the action's default instruction, as the panel does.
        const messages = action.buildRequest(item.snapshot, item.instruction || action.defaultInstruction, OPTIONS);
        const first = await ask(action, item.snapshot, messages);
        entry.steps.push({
          messages,
          raw: first.raw,
          retried: first.retried,
          ms: first.ms,
          retrieval: first.retrieval
        });
        entry.notes = item.check(first.result);
        if (item.followUp && item.checkFollowUp) {
          const conversation = messages.concat([
            { role: "assistant", content: first.raw },
            { role: "user", content: item.followUp }
          ]);
          const second = await ask(action, item.snapshot, conversation);
          entry.steps.push({
            followUp: item.followUp,
            raw: second.raw,
            retried: second.retried,
            ms: second.ms,
            retrieval: second.retrieval
          });
          entry.followUpNotes = item.checkFollowUp(second.result, first.result);
        }
        entry.ok = true;
        process.stdout.write(
          "✓ " + label(item, run, runs) + " (" + entry.steps.map(s => s.ms + " ms").join(", ") + ")\n"
        );
      } catch (error) {
        failed += 1;
        entry.ok = false;
        entry.error = error.message;
        process.stdout.write("✗ " + label(item, run, runs) + "\n  " + error.message.slice(0, 400) + "\n");
      }
    }
  fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
  process.stdout.write(
    "\n" +
      (report.cases.length - failed) +
      " of " +
      report.cases.length +
      " live cases passed" +
      (runs > 1 ? " over " + runs + " runs" : "") +
      ". Report: " +
      reportPath +
      "\n"
  );
  process.exit(failed ? 1 : 0);
})();
