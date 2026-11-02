# CloudCIX Guiden® for Euro-Office

CloudCIX Guiden® is an editor-aware Euro-Office/ONLYOFFICE assistant that calls
the CloudCIX OpenAI-compatible Guiden® API directly.

```text
Browser / Euro-Office plugin
  -> https://inference.cloudcix.com/v1/guiden/chat/completions
```

## Layout

```text
plugin/cloudcix-guiden/   ONLYOFFICE-style plugin folder
scripts/                  installation and runtime preflight helpers
```

## Install The Plugin

Run these commands on the host that runs the Euro-Office/ONLYOFFICE Document Server container. The host needs Docker access; use `sudo` when your account is not in the Docker group. Copy or clone this repository onto that host first.

`euro-office` is the default container name. Replace it with the actual container name when it differs.

```sh
cd /path/to/guiden-eo
sudo ./scripts/preflight-runtime.sh euro-office
sudo ./scripts/install-plugin.sh euro-office
```

The install script checks for the local plugin runtime files first:

```text
sdkjs-plugins/v1/plugins.js
sdkjs-plugins/v1/plugins-ui.js
sdkjs-plugins/v1/plugins.css
```

If those files are missing, the script stops because current Euro-Office builds may otherwise show a blank/inert plugin panel.

If the runtime files are missing, install the official ONLYOFFICE plugin runtime files once:

```sh
sudo ./scripts/install-runtime.sh euro-office
sudo ./scripts/preflight-runtime.sh euro-office
sudo ./scripts/install-plugin.sh euro-office
```

## Build It Into The Image

For Kubernetes, or anywhere the editor's pods are replaced rather than patched, build the plugin into the editor image
instead of copying it into a running container. The `Dockerfile` adds the plugin and its runtime files to Euro-Office
v9.3.7, pinned by digest.

GitHub Actions (`.github/workflows/image.yml`) validates, tests and builds every pull request, and publishes
`ghcr.io/cloudcix/guiden-eo` from `main` (tags `main` and `sha-<commit>`) and from version tags such as `v2.1.0`.
Deploy by the digest in the job's summary, so a pod never picks up a different build.

To build by hand:

```sh
docker build -t guiden-eo .
```

Euro-Office lists the plugins it finds when it starts, so new pods show Guiden with no editor config change.

The editor tells browsers to keep plugin files for a year, under a path that changes only with the Euro-Office
version. After upgrading Guiden on the same Euro-Office version, someone who used the old one may still see it until
they reload the editor with Ctrl+Shift+R (⌘+Shift+R on a Mac) or clear the site's cached files.

## Use

Guiden starts with the editor. It adds:

- a **Guiden** toolbar tab with one-click actions for the editor in use (Improve, Fix, Summarise, Translate, Ask and
  Company Knowledge in documents; Analyse, Clean up, Formula, Format and Table in spreadsheets; Improve, Shorten,
  Notes and New slide in presentations);
- a **Guiden** submenu on the right-click menu, with actions for what is under the cursor: the selected text, the
  cursor position in a document, the selected cells, or the current slide;
- the Guiden panel, with a button in the right sidebar. Any action opens it. Its top row has one-click actions for
  what is selected right now: Improve, Fix, Shorter, Translate, Tone and Ask for selected text; Draft, Review
  document, Ask and Company knowledge with nothing selected; Analyse, Explain formulas, Clean up and Formula in
  spreadsheets; Shorten slide, Speaker notes, New slide and Ask on a slide.

- **Change tone** (more formal, friendlier, more concise, simpler, more confident) for selected text in documents and
  presentations, **Review with comments** for a selection or a whole document, and **Explain formulas** for cells.

Actions such as Improve, Fix and Translate start straight away. Actions that need an instruction (Rewrite, Ask, Draft
here, Formula and so on) open the panel ready for it; press Ctrl+Enter (⌘+Enter on a Mac) to run and Esc to stop.
Every change still stops at a preview, and nothing in the file changes until the Apply button is clicked. The button
says what it will do: **Replace selection**, **Insert at cursor**, **Suggest changes**, **Add 3 comments**, **Update 4
cells** and so on.

Working with a result:

- **Changes** shows the words an edit deletes and adds; **Result** shows the clean text. Light edits such as Fix open
  on Changes.
- **Refine** in one click (Shorter, Longer, More formal, Friendlier, Simpler) or type a follow-up. Guiden continues
  the same conversation, so it changes its last answer instead of starting again. When Guiden asks a question, the
  answer goes in the same box.
- **Suggest as tracked changes** (documents) makes the edit with Track changes on, word by word, so anyone can accept
  or reject each change under Collaboration. Guiden turns tracking back off afterwards unless it was already on;
  when the document already tracks changes, Guiden's edits are always tracked.
- **Review with comments** adds Guiden's comments to the exact phrases they are about, in the editor's own Comments
  pane. Comments that do not quote the text exactly are left out.
- Up and Down in the instruction box step through earlier instructions.
- Details Guiden was not given, such as a dateline city or a spokesperson quote, come back as placeholders like
  `[City]`. The preview marks them, and in documents they stay highlighted in yellow until they are filled in.
- When Guiden cannot be reached, is busy or fails, the panel says so in plain words and offers **Try again**.

The first time, the panel asks for a CloudCIX API key, checks it, remembers it in that browser (untick **Remember
on this browser** to keep it for the session only), and then carries on with the action that was clicked. If CloudCIX
later turns the key down, the panel asks for a new one the same way. The host can also provide the key (see below),
and then nobody is asked.

Documents keep their own look: edits replace each paragraph's text in place where they can, so headings, fonts and
spacing stay; new text takes the document's font and real heading styles, and lists become real bullet and numbered
lists. A draft at the cursor becomes its own paragraphs after the paragraph holding the cursor, and the cursor ends
after it. Only tables go through the editor's HTML importer.

The panel follows the editor's light or dark theme, including when it is switched while the editor is open.

## Provide The Key From The Host

The application that opens the editor can hand Guiden each person's key, so nobody types one in. Add it to the
editor config, under `editorConfig.plugins.options`, keyed by the plugin's GUID, before the config is signed:

```json
"editorConfig": {
  "plugins": {
    "options": {
      "asc.{7D4A7F73-2BB0-4C0F-82F9-0D0CDA18C621}": {
        "guidenApiKey": "the person's CloudCIX API key",
        "guidenUrl": "https://inference.cloudcix.com/v1/guiden/chat/completions",
        "guidenModel": "Mistral-Medium-3.5"
      }
    }
  }
}
```

Only `guidenApiKey` is needed; `guidenUrl` (https only) and `guidenModel` override the defaults. When a key is
provided, Settings hides the key field and the key is never saved in the browser. The key is visible to the person
it belongs to, as it is when they type it, so give each person their own key and never a shared one.

## Notes

- Guiden® is a background plugin (`config.json` variation `type: "background"`) for Writer, Spreadsheet and
  Presentation. `background.html` adds the toolbar tab and right-click menu (both defined in `scripts/menus.js`) and
  opens the panel, `index.html`, as a right-sidebar window. Editor events reach only the background frame, so it passes
  clicks, selection changes, theme changes and host options to the panel.
- Editor calls from the panel are made one at a time and time out after 60 seconds, so an answer can never reach the
  wrong caller and a call the editor never answers cannot block the ones after it.
- Writer provides explicit read, rewrite, correction, summary, drafting, translation, and document-conversion actions. Mutations show a before/after preview and revalidate the selected text before applying.
- Writer opens with **Draft New Content** selected by default; users can switch to Ask or another operation from the compact action picker.
- Cells uses sparse value updates, protected formulas, validated formatting rules, exact destinations, and empty-range checks for new tables and columns. Analysis remains read-only.
- Slides reads indexed text areas and updates only returned shape indexes. It can also create a new slide, preview an outline, or add speaker notes without replacing unrelated objects.
- The plugin uses the local Document Server plugin runtime path `../v1/plugins.js`; it does not depend on `onlyoffice.github.io` at runtime.
- The plugin calls `https://inference.cloudcix.com/v1/guiden/chat/completions` with `model: Mistral-Medium-3.5`, `stream: true`, `response_format: {"type":"json_object"}`, and `max_tokens: 15000`. Responses stream from the service and are validated against the selected action before a preview is shown.
- While a response is streaming, the result panel shows readable text as it arrives. Mutation controls remain hidden until the complete response passes schema and target validation.
- Writer and read-only responses use a prose-first preview: streamed HTML retains paragraph and list spacing, new drafts are not boxed as diffs, and original replacement text is available through a compact disclosure.
- Writer mutation responses distinguish validated content from `needs_input`. Refusals and missing-information messages never expose Apply, and cursor drafts do not show a misleading empty Before panel.
- Writer mutations require a complete structured response and the streaming `[DONE]` marker. Raw HTML, truncated JSON, empty sanitized output, and partial streams are never applicable.
- If a complete response fails its action schema, Guiden automatically regenerates it once with stricter formatting instructions while preserving the exact original user query for retrieval. A second invalid response fails closed and never exposes Apply.
- Writer converts verified links—and exact legacy links accepted through the compatibility warning—into native Euro-Office `Api.CreateHyperlink` objects. The generated native command uses parser-safe line splitting so multiline content cannot break Apply; If the native hyperlink command is unavailable or rejected, Guiden falls back to the proven raw HTML importer so the answer is still inserted with its anchor tags instead of failing the entire action.
- Returned content is sanitized in the browser to a small HTML allowlist for preview and whole-document apply.
- The editor action list is collapsed behind a compact current-action control so the prompt remains the primary interface; choosing Change opens the full list temporarily.
- Free-form use defaults to read-only Ask. Any document mutation requires an explicit action, preview, target revalidation, and Apply.
- Company-knowledge questions are sent to retrieval without unrelated editor text. Answers fail closed when no relevant chunk is retrieved, and citations are copied only from exact HTTP(S) URLs returned by the retrieval service; model-authored links are discarded.
- Company Knowledge uses verified server metadata when available. On legacy servers it enters a labelled compatibility mode, returns the Guiden answer, preserves only exact HTTP(S) URLs reported in the response without repairing them, and allows insertion after user review.
- All editor actions keep document, sheet, and slide context in the system message while the final user message contains only the user instruction. This preserves the clean query used by Guiden retrieval.
- A request is cancelled if Guiden sends nothing for 60 seconds, so a stalled service never leaves the panel waiting.
- Very long documents (over 120,000 characters) or selections (over 5,000 cells) are cut for read-only questions, with a notice in the preview, and refused for changes, which would otherwise lose the part that was not sent.
- Multi-cell results such as a calculated column or a new table are written in one editor command, so one Undo removes all of it.
- No Nextcloud config change is required.

## Tests

The plugin has no dependencies. The unit and mocked-editor suite runs on Node alone:

```sh
node plugin/cloudcix-guiden/tests/run.js
```

The browser suite opens the real panel in Chromium against a fake Writer editor and a fake Guiden API, and checks what
a person sees: the connect card, toolbar and right-click commands arriving from the editor, previews, Apply, Copy,
Company Knowledge, a provided key and the dark theme. It and the formatting check
need the development tools in `package.json`:

```sh
npm ci && npx playwright install chromium
npm test && npm run test:browser
npm run format:check   # npm run format rewrites the files
```

GitHub Actions runs all three on every pull request. Manual release QA is still required in Writer, Cells, and Slides
because host API support varies between Euro-Office builds.
