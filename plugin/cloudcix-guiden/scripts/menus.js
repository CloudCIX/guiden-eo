// What Guiden adds to the editor: its toolbar tab and its right-click menu. Shared by background.js (which builds the
// menus) and the tests. Each command opens the Guiden panel on one action; "run" commands start straight away and
// "open" commands wait for the person to type an instruction. Mutating actions still stop at a preview with Apply.
(function (global) {
  "use strict";
  const G = (global.Guiden = global.Guiden || {});
  const LANGUAGES = ["English", "Irish", "French", "German", "Spanish", "Italian", "Dutch", "Polish", "Portuguese"];
  const translate = language =>
    "Translate into " + language + ". Keep the meaning, names, numbers, links and formatting.";
  // [id, label, instruction] for Change tone.
  const TONES = [
    [
      "formal",
      "More formal",
      "Rewrite it in a more formal, professional register: replace casual words, phrases and contractions with formal ones. Keep the meaning, names and numbers, and keep every claim exactly as strong as it is: do not turn 'good' into 'significant' or 'a lot' into 'substantial'."
    ],
    [
      "friendly",
      "Friendlier",
      "Rewrite it to sound friendlier and warmer while staying professional. Keep the meaning."
    ],
    ["concise", "More concise", "Rewrite it more concisely without losing any information."],
    ["simple", "Simpler", "Rewrite it in plain English: simpler words and shorter sentences. Keep the meaning."],
    ["confident", "More confident", "Rewrite it to sound more confident and direct. Keep the meaning."]
  ];

  // id -> what the panel should do. The id is what the editor sends back when the item is clicked.
  const COMMANDS = {
    "word.improve": { actionId: "writer.improve", run: true, label: "Improve writing" },
    "word.fix": { actionId: "writer.fix", run: true, label: "Fix spelling & grammar" },
    "word.summarise": { actionId: "writer.summarise", run: true, label: "Summarise" },
    "word.rewrite": { actionId: "writer.rewrite", run: false, label: "Rewrite…" },
    "word.ask": { actionId: "writer.ask", run: false, label: "Ask about this…" },
    "word.draft": { actionId: "writer.draft", run: false, label: "Draft here…" },
    "word.knowledge": { actionId: "knowledge.word", run: false, label: "Ask Company Knowledge…" },
    "word.review": { actionId: "writer.review", run: true, label: "Review with comments" },
    "cell.ask": { actionId: "cells.ask", run: true, label: "Analyse selection" },
    "cell.edit": { actionId: "cells.edit", run: true, label: "Clean up values" },
    "cell.explain": {
      actionId: "cells.ask",
      run: true,
      label: "Explain formulas",
      instruction:
        "Explain what each formula in the selection calculates, in plain language, then work it through step by step with the cells' actual values. If there are no formulas, explain what the values show."
    },
    "cell.formula": { actionId: "cells.formula", run: false, label: "Create formula…" },
    "cell.format": { actionId: "cells.format", run: false, label: "Format cells…" },
    "cell.classify": { actionId: "cells.classify", run: false, label: "Classify rows…" },
    "cell.table": { actionId: "cells.table", run: false, label: "Create table…" },
    "cell.knowledge": { actionId: "knowledge.cell", run: false, label: "Ask Company Knowledge…" },
    "slide.improve": { actionId: "slides.improve", run: true, label: "Improve text" },
    "slide.shorten": { actionId: "slides.shorten", run: true, label: "Shorten slide" },
    "slide.rewrite": { actionId: "slides.rewrite", run: false, label: "Rewrite slide…" },
    "slide.notes": { actionId: "slides.notes", run: true, label: "Write speaker notes" },
    "slide.new": { actionId: "slides.new", run: false, label: "New slide…" },
    "slide.ask": { actionId: "slides.ask", run: false, label: "Ask about this slide…" },
    "slide.knowledge": { actionId: "knowledge.slide", run: false, label: "Ask Company Knowledge…" }
  };
  TONES.forEach(tone => {
    COMMANDS["word.tone." + tone[0]] = { actionId: "writer.rewrite", run: true, label: tone[1], instruction: tone[2] };
    COMMANDS["slide.tone." + tone[0]] = { actionId: "slides.improve", run: true, label: tone[1], instruction: tone[2] };
  });
  LANGUAGES.forEach(language => {
    COMMANDS["word.translate." + language] = {
      actionId: "writer.rewrite",
      run: true,
      label: language,
      instruction: translate(language)
    };
    COMMANDS["slide.translate." + language] = {
      actionId: "slides.improve",
      run: true,
      label: language,
      instruction: translate(language)
    };
  });
  const OPEN = "guiden.open";

  // The message the background sends the panel for a clicked id, or null for an id Guiden doesn't know.
  function commandFor(id) {
    if (id === OPEN) return { open: true };
    const command = COMMANDS[id];
    if (!command) return null;
    return { actionId: command.actionId, run: command.run, instruction: command.instruction || "" };
  }

  // Icon templates the editor expands for the light and dark themes.
  function icon(name, folder) {
    return "resources/" + (folder || "toolbar") + "/%theme-type%(light|dark)/" + name + ".%extension%(svg)";
  }
  function item(id, extra) {
    return Object.assign({ id, text: COMMANDS[id].label }, extra || {});
  }
  function languageItems(prefix) {
    return LANGUAGES.map(language => ({ id: prefix + language, text: language }));
  }
  function toneItems(prefix) {
    return TONES.map(tone => ({ id: prefix + tone[0], text: tone[1] }));
  }

  // Right-click menu. type is the editor's menu type: "Selection" when text is selected, "Target" for a plain
  // cursor, and "Shape", "Image" and so on for objects.
  function contextMenu(editor, type) {
    const selected = type === "Selection";
    let items = [];
    if (editor === "word") {
      if (type !== "Selection" && type !== "Target") return [];
      items = selected
        ? [
            item("word.improve"),
            item("word.fix"),
            item("word.summarise", { text: "Summarise selection" }),
            { id: "word.translate", text: "Translate", items: languageItems("word.translate.") },
            { id: "word.tone", text: "Change tone", items: toneItems("word.tone.") },
            item("word.rewrite"),
            item("word.review"),
            item("word.ask", { text: "Ask about selection…" })
          ]
        : [
            item("word.draft"),
            item("word.review", { text: "Review document with comments" }),
            item("word.ask", { text: "Ask about this document…" })
          ];
    } else if (editor === "cell") {
      items = [
        item("cell.ask"),
        item("cell.explain"),
        item("cell.edit"),
        item("cell.formula"),
        item("cell.format"),
        item("cell.classify")
      ];
    } else if (editor === "slide") {
      items = selected
        ? [
            item("slide.improve"),
            { id: "slide.translate", text: "Translate", items: languageItems("slide.translate.") },
            { id: "slide.tone", text: "Change tone", items: toneItems("slide.tone.") },
            item("slide.ask")
          ]
        : [item("slide.shorten"), item("slide.notes"), item("slide.rewrite"), item("slide.ask")];
    } else return [];
    items.push({ id: OPEN, text: "Open Guiden", separator: true });
    return [{ id: "guiden", text: "Guiden", icons: icon("guiden", "menu"), items }];
  }

  function button(id, iconName, text, hint, extra) {
    return Object.assign(
      { id, type: "big-button", text, hint, icons: icon(iconName), lockInViewMode: true },
      extra || {}
    );
  }
  // The "Guiden" toolbar tab.
  function toolbar(editor) {
    const open = button(OPEN, "guiden", "Guiden", "Open the Guiden panel");
    let items = [];
    if (editor === "word")
      items = [
        open,
        button("word.improve", "improve", "Improve", "Improve the selected text", { separator: true }),
        button("word.fix", "fix", "Fix", "Fix spelling and grammar in the selected text"),
        button("word.summarise", "summarise", "Summarise", "Summarise the selected text"),
        button("word.translate", "translate", "Translate", "Translate the selected text", {
          items: languageItems("word.translate.")
        }),
        button("word.tone", "tone", "Tone", "Change the tone of the selected text", {
          items: toneItems("word.tone.")
        }),
        button("word.ask", "ask", "Ask", "Ask a question about the document", { separator: true }),
        button("word.review", "review", "Review", "Review the selection or document and leave comments"),
        button("word.knowledge", "knowledge", "Company Knowledge", "Ask your company's knowledge")
      ];
    else if (editor === "cell")
      items = [
        open,
        button("cell.ask", "ask", "Analyse", "Analyse the selected cells", { separator: true }),
        button("cell.explain", "explain", "Explain", "Explain the formulas in the selected cells"),
        button("cell.edit", "improve", "Clean up", "Clean up the selected values"),
        button("cell.formula", "formula", "Formula", "Create a formula for the selected cell"),
        button("cell.format", "format", "Format", "Format the selected cells"),
        button("cell.table", "table", "Table", "Create a table from the selected cell"),
        button("cell.knowledge", "knowledge", "Company Knowledge", "Ask your company's knowledge", {
          separator: true
        })
      ];
    else if (editor === "slide")
      items = [
        open,
        button("slide.improve", "improve", "Improve", "Improve the selected text", { separator: true }),
        button("slide.shorten", "summarise", "Shorten", "Shorten the current slide"),
        button("slide.notes", "notes", "Notes", "Write speaker notes for the current slide"),
        button("slide.new", "slide", "New slide", "Create a new slide"),
        button("slide.ask", "ask", "Ask", "Ask a question about the current slide", { separator: true }),
        button("slide.knowledge", "knowledge", "Company Knowledge", "Ask your company's knowledge")
      ];
    return [{ id: "guiden", text: "Guiden", items }];
  }

  G.menus = { LANGUAGES, TONES, COMMANDS, OPEN, commandFor, contextMenu, toolbar, icon };
})(typeof window !== "undefined" ? window : globalThis);
