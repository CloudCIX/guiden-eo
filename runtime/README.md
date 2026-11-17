# ONLYOFFICE plugin runtime

`v1/plugins.js`, `v1/plugins-ui.js` and `v1/plugins.css` are the ONLYOFFICE plugin runtime, copied unchanged from
[ONLYOFFICE/onlyoffice.github.io](https://github.com/ONLYOFFICE/onlyoffice.github.io), `sdkjs-plugins/v1/`, at commit
`496e5520dd5ec57fc337d8f85fa8c8e2c0871cfc` (22 Jul 2026). They are © Ascensio System SIA and licensed under the GNU
AGPL v3, as their headers say.

Euro-Office builds may ship without them, and then the plugin panel stays blank, so the image adds them
(`Dockerfile`) and `scripts/install-runtime.sh` copies these same files into a running container.

To update them, copy the three files from a newer commit, put that commit here, and test in Writer, Cells and Slides.
To check the copies still match, compare `git hash-object runtime/v1/<file>` with
`git rev-parse <commit>:sdkjs-plugins/v1/<file>` in a clone of that repository.
