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

## Use

1. Open a document, spreadsheet, or presentation in Euro-Office from Nextcloud.
2. Open **Plugins > CloudCIX Guiden®**.
3. Enter the CloudCIX API key.
4. Select the text/cells/object you want to work with, or leave the current Writer document/current slide active for a broader suggestion.
5. Choose an editor-specific action, then add an optional instruction.
6. Click **Generate preview**, review the exact target and changes, then click **Apply changes**.

The API key is not stored unless **Remember key on this browser** is checked.

## Notes

- Guiden® registers for Writer, Spreadsheet, and Presentation editors via `EditorsSupport: ["word", "cell", "slide"]`.
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
- No Nextcloud config change is required.

## Tests

The V2 validation and mocked-host suite has no package dependencies:

```sh
node plugin/cloudcix-guiden/tests/run.js
```

Manual release QA is still required in Writer, Cells, and Slides because host API support varies between Euro-Office builds.
