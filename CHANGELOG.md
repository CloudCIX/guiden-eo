# Changelog

Each release is the version in `plugin/cloudcix-guiden/config.json`, tagged `v<version>` on `main`. CI refuses a tag
that does not match `config.json` or has no section here, and the release notes on GitHub are this section.

Earlier versions (before 2.5.2) were never tagged; their entries are reconstructed from the commit history.

## 2.5.2 (17 Nov 2026)

- No change to what Guiden does. The first release made by the new release process: merging a new version into
  `main` now tags it, publishes the image as `ghcr.io/cloudcix/guiden-eo:<version>` and makes a GitHub release.
- `scripts/install-runtime.sh` installs the pinned plugin runtime in `runtime/v1` instead of upstream's latest.

## 2.5.1 (3 Nov 2026)

- The toolbar button, right-click menu and editor sidebar use the CloudCIX mark again.
- iCloud Passwords no longer offers saved logins on the instruction box.

## 2.5.0 (3 Nov 2026)

- **New slide** builds slides in the deck's own layouts, fonts and colours.
- **Design deck** applies one of six ready-made designs to every slide master.
- Password managers no longer pop up on the panel's text boxes.

## 2.4.0 (2 Nov 2026)

- One-click actions in the panel's top row, clearer Apply buttons, and placeholders such as `[City]` for details
  Guiden was not given.

## 2.3.0 to 2.3.4 (2 Nov 2026)

- Tracked-change suggestions, review comments, follow-ups and a Changes view.
- Prompt tuning from live runs: drafts get the document as context, no made-up datelines, no example words leaking
  into replies, speaker notes stay on the slide.

## 2.2.0 (2 Nov 2026)

- Guiden is in the editor's toolbar and right-click menu, and asks for the key only when it is needed.

## 2.1.0 (2 Nov 2026)

- The host can provide each person's key through the editor config.
- Fixes to Copy, Sources, the table theme and undo; request timeouts; dark theme.
- The plugin is built into a Euro-Office v9.3.7 image, published from GitHub Actions to `ghcr.io/cloudcix/guiden-eo`.

## 2.0.17 (28 Jul 2026)

- Version 2: a rewrite of the panel and the Writer, Cells and Slides actions.

## 1.4.7 (23 Jul 2026)

- First public release.
