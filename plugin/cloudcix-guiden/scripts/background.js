// Starts with the editor (config.json makes Guiden a background plugin). It adds the Guiden toolbar tab, the Guiden
// right-click menu and the Guiden side panel (index.html), and passes clicks, selection changes and theme changes to
// the panel. The panel does all the work; this frame only runs editor methods for it.
(function (global) {
  "use strict";
  // The runtime adds PluginWindow to Asc when the plugin starts, so keep the object rather than the class.
  const Asc = global.Asc,
    plugin = Asc.plugin,
    M = global.Guiden.menus;
  const state = { editor: "word", panel: null, ready: false, queue: [], selectionTimer: null };

  function execute(name, args) {
    plugin.executeMethod(name, args || []);
  }

  // The panel is a window frame in the right sidebar. It is created once, closed (sidebar button only) at start, and
  // opened when a command needs it.
  function panelVariation(activate) {
    return {
      url: "index.html",
      description: "Guiden",
      icons: "resources/%theme-type%(light|dark)/panel.%extension%(svg)",
      type: "panelRight",
      isVisual: true,
      isModal: false,
      isViewer: false,
      isActivated: activate,
      EditorsSupport: ["word", "cell", "slide"],
      buttons: []
    };
  }
  function createPanel(activate) {
    const panel = new Asc.PluginWindow();
    state.panel = panel;
    state.ready = false;
    panel.attachEvent("guidenReady", () => {
      state.ready = true;
      const waiting = state.queue.splice(0);
      waiting.forEach(([name, data]) => panel.command(name, data));
    });
    panel.show(panelVariation(activate));
  }
  function send(name, data) {
    if (state.panel && state.ready) state.panel.command(name, data);
    else state.queue.push([name, data]);
  }
  function openPanel() {
    if (!state.panel) createPanel(true);
    else state.panel.activate(true);
  }
  function run(id) {
    const command = M.commandFor(id);
    if (!command) return;
    openPanel();
    if (!command.open) send("guidenRun", command);
  }

  // Editor events reach only this frame, so the panel hears about them from here.
  function selectionChanged() {
    if (!state.ready) return;
    clearTimeout(state.selectionTimer);
    state.selectionTimer = setTimeout(() => send("guidenSelection", {}), 250);
  }
  function themeOf(theme) {
    return theme && typeof theme === "object" ? { type: theme.type, name: theme.name } : null;
  }

  plugin.init = function () {
    const info = plugin.info || {};
    state.editor = ["word", "cell", "slide"].includes(info.editorType) ? info.editorType : "word";
    execute("AddToolbarMenuItem", [{ guid: plugin.guid, tabs: M.toolbar(state.editor) }]);
    createPanel(false);
    ["onSelectionChanged", "onTargetPositionChanged"].forEach(name => {
      try {
        plugin.attachEditorEvent(name, selectionChanged);
      } catch (_) {}
    });
  };
  plugin.button = function () {};
  // The editor waits for every plugin that listens for onContextMenuShow, so always answer, even with nothing.
  plugin.event_onContextMenuShow = function (options) {
    const type = options && (options.type || options.Type);
    execute("AddContextMenuItem", [{ guid: plugin.guid, items: M.contextMenu(state.editor, type) }]);
  };
  plugin.event_onContextMenuClick = function (id) {
    run(String(id || "").split("_oo_sep_")[0]);
  };
  plugin.event_onToolbarMenuClick = function (id) {
    run(String(id || "").split("_oo_sep_")[0]);
  };
  plugin.onThemeChanged = function (theme) {
    if (plugin.onThemeChangedBase) plugin.onThemeChangedBase(theme);
    send("guidenTheme", themeOf(theme));
  };
  plugin.onUpdateOptions = function () {
    send("guidenOptions", (plugin.info && plugin.info.options) || {});
  };
  global.Guiden.background = state;
})(window);
