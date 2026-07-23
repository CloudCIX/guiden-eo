# CloudCIX Guiden® for Euro-Office

CloudCIX Guiden® is a Writer-first Euro-Office/ONLYOFFICE editor plugin that calls
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
5. Choose a quick action or type a custom instruction.
6. Click **Generate**, review the suggestion, then click **Apply**.

The API key is not stored unless **Remember key on this browser** is checked.

## Notes

- Guiden® registers for Writer, Spreadsheet, and Presentation editors via `EditorsSupport: ["word", "cell", "slide"]`.
- Writer automatically edits selected text when text is selected, otherwise it works on the whole document. Simple selected-text edits use `PasteText` to preserve local formatting; structured selected output and whole-document output use semantic HTML paste without forced plugin fonts or margins.
- Spreadsheets use guided operations for analysis, formula generation, table generation, and cell rewrites. Analysis is preview-only; formula generation writes one formula to the selected starting cell; tables write only into safe empty ranges; rewrites protect formulas and non-text numeric cells.
- Presentations read selected text first, otherwise the current slide text. Applies preserve slide layout by updating selected text or an existing text-bearing shape where possible.
- The plugin uses the local Document Server plugin runtime path `../v1/plugins.js`; it does not depend on `onlyoffice.github.io` at runtime.
- The plugin calls `https://inference.cloudcix.com/v1/guiden/chat/completions` with `model: Mistral-Medium-3.5`, `stream: true`, and `max_tokens: 15000`, so the suggestion panel can update while the model is drafting.
- Returned content is sanitized in the browser to a small HTML allowlist for preview and whole-document apply.
- No Nextcloud config change is required.
