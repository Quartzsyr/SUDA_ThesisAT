let bridge = null;
const fieldTimers = new Map();
const $ = (id) => document.getElementById(id);

function debounceField(name, value) {
  clearTimeout(fieldTimers.get(name));
  fieldTimers.set(name, setTimeout(() => {
    if (bridge) bridge.setField(name, value);
  }, 180));
}

function syncInput(el, value) {
  if (!el || document.activeElement === el) return;
  const next = value || "";
  if (el.value !== next) el.value = next;
}

function renderPreview(previewUrl) {
  const frame = $("preview-frame");
  const shell = $("preview-shell");
  const empty = $("preview-empty");
  if (previewUrl) {
    const src = `${previewUrl}#toolbar=0&navpanes=0&scrollbar=1&view=FitH`;
    if (frame.dataset.src !== src) {
      frame.src = src;
      frame.dataset.src = src;
    }
    shell.style.display = "block";
    empty.style.display = "none";
    return;
  }
  frame.removeAttribute("src");
  frame.dataset.src = "";
  shell.style.display = "none";
  empty.style.display = "grid";
}

function renderWindowState(state) {
  document.body.classList.toggle("maximized", !!state.windowMaximized);
  $("app-name").textContent = state.appName || "ThesisFlow";
  $("window-max").textContent = state.windowMaximized ? "❐" : "□";
}

function renderState(state) {
  const select = $("template-select");
  if (select.options.length === 0) {
    state.templateOptions.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      select.appendChild(option);
    });
  }

  if (select.value !== state.template) select.value = state.template;
  $("status-pill").textContent = state.previewStatus;
  $("source-path").value = state.sourcePath || "";
  $("output-path").value = state.outputPath || "";
  $("doc-metrics").textContent = state.docMetrics;
  $("doc-preview").textContent = state.docPreview || "这里会显示正文摘要。";
  $("file-pill").textContent = state.fileName;
  $("metric-template").textContent = state.metricTemplate;
  $("metric-cover").textContent = state.metricCover;
  $("metric-output").textContent = state.metricOutput;
  $("preview-path").textContent = state.previewPathText;
  $("include-cover").checked = !!state.includeCover;
  document.querySelectorAll("[data-field]").forEach((input) => syncInput(input, state.fields[input.dataset.field]));
  renderPreview(state.previewUrl);
  renderWindowState(state);
}

function showToast(kind, message) {
  const host = $("toast-host");
  const node = document.createElement("div");
  node.className = `toast ${kind}`;
  node.textContent = message;
  host.appendChild(node);
  requestAnimationFrame(() => node.classList.add("visible"));
  setTimeout(() => {
    node.classList.remove("visible");
    setTimeout(() => node.remove(), 180);
  }, 3200);
}

function isInteractive(target) {
  return Boolean(target.closest("button, input, select, textarea, a, label.switch"));
}

function bindTitlebar() {
  const titlebar = $("titlebar");
  titlebar.addEventListener("mousedown", (event) => {
    if (event.button !== 0 || isInteractive(event.target)) return;
    if (bridge) bridge.startWindowDrag();
  });
  titlebar.addEventListener("dblclick", (event) => {
    if (!isInteractive(event.target) && bridge) bridge.toggleMaximize();
  });
}

function bindEvents() {
  $("template-select").addEventListener("change", (event) => bridge && bridge.setTemplate(event.target.value));
  $("pick-source").addEventListener("click", () => bridge && bridge.chooseSourceFile());
  $("pick-output").addEventListener("click", () => bridge && bridge.chooseOutputFile());
  $("open-editor").addEventListener("click", () => bridge && bridge.openEditor());
  $("export-files").addEventListener("click", () => bridge && bridge.saveOutput());
  $("refresh-preview").addEventListener("click", () => bridge && bridge.refreshPreview());
  $("about-btn").addEventListener("click", () => bridge && bridge.showAbout());
  $("include-cover").addEventListener("change", (event) => bridge && bridge.setIncludeCover(event.target.checked));
  $("window-min").addEventListener("click", () => bridge && bridge.minimizeWindow());
  $("window-max").addEventListener("click", () => bridge && bridge.toggleMaximize());
  $("window-close").addEventListener("click", () => bridge && bridge.closeWindow());
  document.querySelectorAll("[data-field]").forEach((input) => {
    input.addEventListener("input", (event) => debounceField(event.target.dataset.field, event.target.value));
  });
  $("preview-frame").addEventListener("error", () => showToast("error", "PDF 预览加载失败，请尝试刷新预览。"));
  bindTitlebar();
}

new QWebChannel(qt.webChannelTransport, (channel) => {
  bridge = channel.objects.bridge;
  bridge.stateChanged.connect((payload) => renderState(JSON.parse(payload)));
  bridge.toast.connect((kind, message) => showToast(kind, message));
  bindEvents();
  bridge.requestState();
});


