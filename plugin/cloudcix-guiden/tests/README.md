# Guiden tests

From the repository root:

```sh
node plugin/cloudcix-guiden/tests/run.js       # helpers, menus, the background frame and mocked-editor apply
npm ci && npx playwright install chromium
node plugin/cloudcix-guiden/tests/browser.js   # the real panel in Chromium, with a fake editor and a fake Guiden API
```

`run.js` needs only Node. `browser.js` needs Playwright from the repository's `package.json`. Manual host QA remains
required for Euro-Office API compatibility.

`GUIDEN_TEST="part of a test name" node plugin/cloudcix-guiden/tests/browser.js` runs only the matching browser tests.
