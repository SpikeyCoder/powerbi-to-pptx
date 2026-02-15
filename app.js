const models = window["powerbi-client"]?.models;

const state = {
  report: null,
  pages: [],
  selectedVisualKeys: new Set(),
  visualIndex: new Map(),
  msalApp: null,
  msalConfigKey: "",
  account: null,
  msalToken: "",
  tokenExpiresOn: null,
  demoMode: false,
  busy: false,
};

const layoutDimensions = {
  LAYOUT_WIDE: { w: 13.333, h: 7.5 },
  LAYOUT_STANDARD: { w: 10, h: 7.5 },
};

const DECK_FONT_FAMILY = "D-DIN";
const DECK_CANVAS_FONT_STACK = '"D-DIN", Arial, sans-serif';

const cloudProfiles = {
  commercial: {
    id: "commercial",
    authorityBase: "https://login.microsoftonline.com",
    defaultScope: "https://analysis.windows.net/powerbi/api/Report.Read.All",
  },
  gcc: {
    id: "gcc",
    authorityBase: "https://login.microsoftonline.com",
    defaultScope: "https://analysis.usgovcloudapi.net/powerbi/api/Report.Read.All",
  },
  gccHigh: {
    id: "gccHigh",
    authorityBase: "https://login.microsoftonline.us",
    defaultScope: "https://high.analysis.usgovcloudapi.net/powerbi/api/Report.Read.All",
  },
  dod: {
    id: "dod",
    authorityBase: "https://login.microsoftonline.us",
    defaultScope: "https://mil.analysis.usgovcloudapi.net/powerbi/api/Report.Read.All",
  },
  china: {
    id: "china",
    authorityBase: "https://login.chinacloudapi.cn",
    defaultScope: "https://analysis.chinacloudapi.cn/powerbi/api/Report.Read.All",
  },
};

const allDefaultScopes = new Set(Object.values(cloudProfiles).map((profile) => profile.defaultScope));

const dom = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  wireEvents();
  toggleAuthModeSections();
  applyCloudDefaults(true);
  ensureSdkAvailability();
  logStatus("Ready. Provide auth + embed settings, then click Embed Report.", "success");
});

function cacheDom() {
  dom.authModeRadios = document.querySelectorAll("input[name='authMode']");
  dom.manualTokenSection = document.getElementById("manualTokenSection");
  dom.msalSection = document.getElementById("msalSection");

  dom.accessTokenInput = document.getElementById("accessTokenInput");
  dom.tokenTypeInput = document.getElementById("tokenTypeInput");

  dom.tenantIdInput = document.getElementById("tenantIdInput");
  dom.clientIdInput = document.getElementById("clientIdInput");
  dom.cloudEnvironmentInput = document.getElementById("cloudEnvironmentInput");
  dom.scopesInput = document.getElementById("scopesInput");
  dom.signInBtn = document.getElementById("signInBtn");
  dom.signOutBtn = document.getElementById("signOutBtn");
  dom.msalAccountInfo = document.getElementById("msalAccountInfo");

  dom.embedUrlInput = document.getElementById("embedUrlInput");
  dom.reportIdInput = document.getElementById("reportIdInput");
  dom.startingPageInput = document.getElementById("startingPageInput");

  dom.embedReportBtn = document.getElementById("embedReportBtn");
  dom.loadVisualsBtn = document.getElementById("loadVisualsBtn");
  dom.loadDemoBtn = document.getElementById("loadDemoBtn");

  dom.deckTitleInput = document.getElementById("deckTitleInput");
  dom.slideLayoutInput = document.getElementById("slideLayoutInput");
  dom.imageScaleInput = document.getElementById("imageScaleInput");
  dom.includePageNameInTitleInput = document.getElementById("includePageNameInTitleInput");

  dom.selectAllBtn = document.getElementById("selectAllBtn");
  dom.clearSelectionBtn = document.getElementById("clearSelectionBtn");
  dom.thumbnailBtn = document.getElementById("thumbnailBtn");
  dom.generatePptBtn = document.getElementById("generatePptBtn");

  dom.selectionCountLabel = document.getElementById("selectionCountLabel");
  dom.embedContainer = document.getElementById("embedContainer");
  dom.visualSelection = document.getElementById("visualSelection");
  dom.statusLog = document.getElementById("statusLog");
}

function wireEvents() {
  for (const radio of dom.authModeRadios) {
    radio.addEventListener("change", toggleAuthModeSections);
  }

  if (dom.cloudEnvironmentInput) {
    dom.cloudEnvironmentInput.addEventListener("change", () => applyCloudDefaults(false));
  }

  dom.signInBtn.addEventListener("click", () => runAction("MSAL sign in", signInWithMsal));
  dom.signOutBtn.addEventListener("click", () => runAction("MSAL sign out", signOutMsal));

  dom.embedReportBtn.addEventListener("click", () => runAction("Embed report", embedReport));
  dom.loadVisualsBtn.addEventListener("click", () => runAction("Load pages + visuals", loadPagesAndVisuals));
  dom.loadDemoBtn.addEventListener("click", () => runAction("Load demo visuals", loadDemoData));

  dom.selectAllBtn.addEventListener("click", selectAllVisuals);
  dom.clearSelectionBtn.addEventListener("click", clearSelections);
  dom.thumbnailBtn.addEventListener("click", () => runAction("Load thumbnails", loadThumbnailsForSelection));
  dom.generatePptBtn.addEventListener("click", () => runAction("Generate PPTX", generateDeck));

  dom.visualSelection.addEventListener("change", onVisualSelectionChanged);
}

function ensureSdkAvailability() {
  if (!window.powerbi || !models) {
    logStatus("Power BI JS SDK is unavailable. Check CDN loading.", "error");
  }
  if (!window.PptxGenJS) {
    logStatus("PptxGenJS is unavailable. Check CDN loading.", "error");
  }
}

function getSelectedCloudProfile() {
  const cloudKey = dom.cloudEnvironmentInput?.value || "commercial";
  return cloudProfiles[cloudKey] || cloudProfiles.commercial;
}

function applyCloudDefaults(force) {
  const profile = getSelectedCloudProfile();
  const currentValue = (dom.scopesInput?.value || "").trim();

  if (force || !currentValue || allDefaultScopes.has(currentValue)) {
    dom.scopesInput.value = profile.defaultScope;
  }
}

function getAuthMode() {
  const checked = document.querySelector("input[name='authMode']:checked");
  return checked?.value || "manual";
}

function toggleAuthModeSections() {
  const mode = getAuthMode();
  dom.manualTokenSection.classList.toggle("hidden", mode !== "manual");
  dom.msalSection.classList.toggle("hidden", mode !== "msal");
}

async function runAction(label, fn) {
  if (state.busy) {
    logStatus("Another operation is running. Please wait.");
    return;
  }

  state.busy = true;
  setButtonsDisabled(true);
  try {
    await fn();
  } catch (error) {
    logStatus(`${label} failed: ${extractErrorMessage(error)}`, "error");
    console.error(error);
  } finally {
    setButtonsDisabled(false);
    state.busy = false;
  }
}

function setButtonsDisabled(disabled) {
  const buttons = [
    dom.signInBtn,
    dom.signOutBtn,
    dom.embedReportBtn,
    dom.loadVisualsBtn,
    dom.loadDemoBtn,
    dom.selectAllBtn,
    dom.clearSelectionBtn,
    dom.thumbnailBtn,
    dom.generatePptBtn,
  ];
  buttons.forEach((button) => {
    if (button) {
      button.disabled = disabled;
    }
  });
}

async function ensureMsalApp() {
  if (!window.msal) {
    throw new Error("MSAL SDK is not loaded.");
  }

  const tenantId = dom.tenantIdInput.value.trim();
  const clientId = dom.clientIdInput.value.trim();

  if (!tenantId || !clientId) {
    throw new Error("Tenant ID and Client ID are required for MSAL auth.");
  }

  const cloudProfile = getSelectedCloudProfile();
  const configKey = `${tenantId}|${clientId}|${cloudProfile.id}|${cloudProfile.authorityBase}`;
  if (!state.msalApp || state.msalConfigKey !== configKey) {
    state.msalApp = new window.msal.PublicClientApplication({
      auth: {
        clientId,
        authority: `${cloudProfile.authorityBase}/${tenantId}`,
        redirectUri: window.location.href.split("#")[0],
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    });

    await state.msalApp.initialize();
    state.msalConfigKey = configKey;

    const redirectResult = await state.msalApp.handleRedirectPromise();
    if (redirectResult?.account) {
      state.account = redirectResult.account;
    }
  }

  if (!state.account) {
    const accounts = state.msalApp.getAllAccounts();
    if (accounts.length > 0) {
      state.account = accounts[0];
    }
  }

  updateMsalAccountInfo();
  return state.msalApp;
}

function readMsalScopes() {
  const raw = dom.scopesInput.value.trim();
  if (!raw) {
    return [getSelectedCloudProfile().defaultScope];
  }
  return raw
    .split(/[\s,]+/)
    .map((scope) => scope.trim())
    .filter(Boolean);
}

async function signInWithMsal() {
  const app = await ensureMsalApp();
  const scopes = readMsalScopes();

  const loginResponse = await app.loginPopup({
    scopes,
    prompt: "select_account",
  });

  if (loginResponse?.account) {
    state.account = loginResponse.account;
  }

  await acquireMsalToken(scopes);
  dom.tokenTypeInput.value = "Aad";

  logStatus("MSAL sign-in complete. Token copied into token field.", "success");
}

async function signOutMsal() {
  const app = await ensureMsalApp();

  if (!state.account) {
    logStatus("No signed-in MSAL account found.");
    return;
  }

  await app.logoutPopup({
    account: state.account,
    postLogoutRedirectUri: window.location.href.split("#")[0],
  });

  state.account = null;
  state.msalToken = "";
  state.tokenExpiresOn = null;
  dom.accessTokenInput.value = "";
  updateMsalAccountInfo();

  logStatus("Signed out from MSAL.", "success");
}

async function acquireMsalToken(scopes) {
  const app = await ensureMsalApp();

  if (!state.account) {
    throw new Error("No MSAL account available. Sign in first.");
  }

  let tokenResponse;
  try {
    tokenResponse = await app.acquireTokenSilent({
      account: state.account,
      scopes,
    });
  } catch (silentError) {
    tokenResponse = await app.acquireTokenPopup({
      account: state.account,
      scopes,
    });
  }

  state.msalToken = tokenResponse.accessToken;
  state.tokenExpiresOn = tokenResponse.expiresOn || null;
  dom.accessTokenInput.value = tokenResponse.accessToken;
  updateMsalAccountInfo();

  return tokenResponse.accessToken;
}

function updateMsalAccountInfo() {
  if (!state.account) {
    dom.msalAccountInfo.textContent = "Not signed in.";
    return;
  }

  const expires = state.tokenExpiresOn
    ? `Token exp: ${state.tokenExpiresOn.toLocaleTimeString()}`
    : "Token exp: unknown";
  dom.msalAccountInfo.textContent = `${state.account.username || state.account.homeAccountId} | ${expires}`;
}

function getTokenTypeValue() {
  if (!models?.TokenType) {
    throw new Error("Power BI models.TokenType is unavailable.");
  }

  return dom.tokenTypeInput.value === "Embed" ? models.TokenType.Embed : models.TokenType.Aad;
}

async function resolveAccessToken() {
  const mode = getAuthMode();

  if (mode === "msal") {
    const scopes = readMsalScopes();
    if (!state.msalToken || !state.tokenExpiresOn || Date.now() > state.tokenExpiresOn.getTime() - 2 * 60 * 1000) {
      await acquireMsalToken(scopes);
    }
    return state.msalToken;
  }

  const manualToken = dom.accessTokenInput.value.trim();
  if (!manualToken) {
    throw new Error("Access token is required.");
  }

  return manualToken;
}

async function embedReport() {
  if (!window.powerbi) {
    throw new Error("Power BI SDK is unavailable.");
  }

  const accessToken = await resolveAccessToken();
  const embedUrl = dom.embedUrlInput.value.trim();
  const reportId = dom.reportIdInput.value.trim();

  if (!embedUrl || !reportId) {
    throw new Error("Embed URL and Report ID are required.");
  }

  state.demoMode = false;
  resetSelectionState();
  window.powerbi.reset(dom.embedContainer);

  const config = {
    type: "report",
    tokenType: getTokenTypeValue(),
    accessToken,
    embedUrl,
    id: reportId,
    settings: {
      panes: {
        filters: { visible: true },
        pageNavigation: { visible: true },
      },
      background: models.BackgroundType.Transparent,
    },
  };

  state.report = window.powerbi.embed(dom.embedContainer, config);
  attachReportErrorListener(state.report);

  await waitForReportEvent(state.report, "loaded", 45000);

  const requestedPage = dom.startingPageInput.value.trim();
  if (requestedPage) {
    await state.report.setPage(requestedPage);
  }

  logStatus("Report embedded successfully.", "success");
}

function attachReportErrorListener(report) {
  report.off("error");
  report.on("error", (event) => {
    const message = event?.detail?.message || event?.detail || "Unknown Power BI error";
    logStatus(`Power BI error: ${message}`, "error");
  });
}

function waitForReportEvent(report, eventName, timeoutMs) {
  return new Promise((resolve, reject) => {
    let timeoutId;

    const onSuccess = (event) => {
      cleanup();
      resolve(event);
    };

    const onError = (event) => {
      cleanup();
      reject(new Error(event?.detail?.message || "Power BI returned an error event."));
    };

    const cleanup = () => {
      clearTimeout(timeoutId);
      report.off(eventName);
      report.off("error", onError);
    };

    report.on(eventName, onSuccess);
    report.on("error", onError);

    timeoutId = setTimeout(() => {
      cleanup();
      reject(new Error(`Timed out waiting for report event "${eventName}".`));
    }, timeoutMs);
  });
}

async function loadPagesAndVisuals() {
  if (!state.report) {
    throw new Error("Embed a report first.");
  }

  state.demoMode = false;
  const pages = await state.report.getPages();
  const loadedPages = [];

  for (const page of pages) {
    const visuals = await page.getVisuals();
    loadedPages.push({
      page,
      visuals: visuals.filter((visual) => visual?.name),
    });
  }

  state.pages = loadedPages;
  rebuildVisualIndex();
  renderVisualSelection();
  updateSelectionCount();

  logStatus(`Loaded ${loadedPages.length} pages and ${state.visualIndex.size} visuals.`, "success");
}

async function loadDemoData() {
  state.demoMode = true;
  state.report = null;

  if (window.powerbi) {
    window.powerbi.reset(dom.embedContainer);
  }

  setDemoEmbedPlaceholder();
  resetSelectionState("Demo visuals loaded. Adjust selection and generate the deck.");

  state.pages = createDemoPages();
  rebuildVisualIndex();
  renderVisualSelection();
  selectAllVisuals();

  logStatus("Demo mode loaded with sample pages and visuals.", "success");
}

function createDemoPages() {
  return [
    {
      page: { name: "DemoPageExecutive", displayName: "Executive Overview" },
      visuals: [
        { name: "visualRevenueKpi", title: "Revenue KPI", type: "card", layout: { width: 420, height: 260 } },
        { name: "visualRevenueTrend", title: "Monthly Revenue Trend", type: "lineChart", layout: { width: 880, height: 420 } },
        { name: "visualMarginBySegment", title: "Margin by Segment", type: "barChart", layout: { width: 760, height: 430 } },
      ],
    },
    {
      page: { name: "DemoPageRegional", displayName: "Regional Performance" },
      visuals: [
        { name: "visualGeoMap", title: "Revenue by Region", type: "map", layout: { width: 860, height: 500 } },
        { name: "visualDealPipeline", title: "Deal Pipeline", type: "columnChart", layout: { width: 780, height: 420 } },
        { name: "visualTopAccounts", title: "Top Accounts", type: "tableEx", layout: { width: 920, height: 380 } },
      ],
    },
    {
      page: { name: "DemoPageProfitability", displayName: "Profitability Insights" },
      visuals: [
        { name: "visualWaterfall", title: "Variance Waterfall", type: "waterfall", layout: { width: 860, height: 430 } },
        { name: "visualProfitScatter", title: "Profitability Scatter", type: "scatterChart", layout: { width: 840, height: 430 } },
        { name: "visualForecast", title: "Forecast Snapshot", type: "areaChart", layout: { width: 820, height: 390 } },
      ],
    },
  ];
}

function setDemoEmbedPlaceholder() {
  dom.embedContainer.innerHTML =
    '<div class="empty-state" style="height: 100%; min-height: 340px;">' +
    'Demo Mode enabled. No Power BI authentication is required.<br />' +
    'Use the generated sample visuals to preview PPTX output.' +
    '</div>';
}

function rebuildVisualIndex() {
  state.visualIndex.clear();

  for (const pageGroup of state.pages) {
    for (const visual of pageGroup.visuals) {
      const key = makeVisualKey(pageGroup.page.name, visual.name);
      state.visualIndex.set(key, {
        page: pageGroup.page,
        visual,
      });
    }
  }
}

function renderVisualSelection() {
  if (!state.pages.length) {
    dom.visualSelection.className = "visual-selection empty-state";
    dom.visualSelection.textContent = "No visuals available. Ensure your report has accessible pages.";
    return;
  }

  dom.visualSelection.className = "visual-selection";

  const groups = state.pages.map((pageGroup) => {
    const pageName = pageGroup.page.name;
    const pageDisplayName = pageGroup.page.displayName || pageName;

    const cards = pageGroup.visuals.map((visual) => {
      const key = makeVisualKey(pageName, visual.name);
      const domKey = encodeURIComponent(key);
      const visualTitle = visual.title || visual.name;
      const dimensions = describeVisualDimensions(visual);
      const checked = state.selectedVisualKeys.has(key) ? "checked" : "";

      return `
        <article class="visual-card" data-card-key="${domKey}">
          <div class="visual-card-top">
            <input type="checkbox" data-visual-checkbox data-visual-dom-key="${domKey}" ${checked} />
            <div>
              <div class="visual-title">${escapeHtml(visualTitle)}</div>
              <div class="visual-subtitle">${escapeHtml(visual.name)}</div>
            </div>
          </div>
          <div class="visual-meta">
            <span>${escapeHtml(visual.type || "unknown")}</span>
            <span>${escapeHtml(dimensions)}</span>
          </div>
          <div class="thumb-wrap" data-thumb-wrap="${domKey}">Thumbnail not loaded</div>
        </article>
      `;
    });

    const pageDomKey = encodeURIComponent(pageName);

    return `
      <section class="page-group" data-page-dom-key="${pageDomKey}">
        <div class="page-header">
          <label class="inline-option">
            <input type="checkbox" data-page-checkbox data-page-dom-key="${pageDomKey}" />
            <span class="page-name">${escapeHtml(pageDisplayName)}</span>
          </label>
          <span class="muted">${pageGroup.visuals.length} visuals</span>
        </div>
        <div class="visual-list">
          ${cards.join("")}
        </div>
      </section>
    `;
  });

  dom.visualSelection.innerHTML = groups.join("");
  refreshPageCheckboxStates();
}

function onVisualSelectionChanged(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  if (target.hasAttribute("data-visual-checkbox")) {
    const domKey = target.getAttribute("data-visual-dom-key");
    if (!domKey) {
      return;
    }

    const key = decodeURIComponent(domKey);
    if (target.checked) {
      state.selectedVisualKeys.add(key);
    } else {
      state.selectedVisualKeys.delete(key);
    }

    refreshPageCheckboxStates();
    updateSelectionCount();
    return;
  }

  if (target.hasAttribute("data-page-checkbox")) {
    const pageDomKey = target.getAttribute("data-page-dom-key");
    if (!pageDomKey) {
      return;
    }

    const pageGroup = dom.visualSelection.querySelector(`[data-page-dom-key="${escapeSelector(pageDomKey)}"]`);
    if (!pageGroup) {
      return;
    }

    const visualChecks = pageGroup.querySelectorAll("input[data-visual-checkbox]");
    visualChecks.forEach((checkbox) => {
      checkbox.checked = target.checked;
      const visualDomKey = checkbox.getAttribute("data-visual-dom-key");
      if (!visualDomKey) {
        return;
      }
      const key = decodeURIComponent(visualDomKey);
      if (target.checked) {
        state.selectedVisualKeys.add(key);
      } else {
        state.selectedVisualKeys.delete(key);
      }
    });

    refreshPageCheckboxStates();
    updateSelectionCount();
  }
}

function refreshPageCheckboxStates() {
  const pageGroups = dom.visualSelection.querySelectorAll("[data-page-dom-key]");

  pageGroups.forEach((pageGroup) => {
    const pageCheckbox = pageGroup.querySelector("input[data-page-checkbox]");
    const visuals = pageGroup.querySelectorAll("input[data-visual-checkbox]");
    const checkedCount = Array.from(visuals).filter((checkbox) => checkbox.checked).length;

    pageCheckbox.checked = checkedCount === visuals.length && visuals.length > 0;
    pageCheckbox.indeterminate = checkedCount > 0 && checkedCount < visuals.length;
  });
}

function selectAllVisuals() {
  const checks = dom.visualSelection.querySelectorAll("input[data-visual-checkbox]");
  checks.forEach((checkbox) => {
    checkbox.checked = true;
    const domKey = checkbox.getAttribute("data-visual-dom-key");
    if (domKey) {
      state.selectedVisualKeys.add(decodeURIComponent(domKey));
    }
  });
  refreshPageCheckboxStates();
  updateSelectionCount();
}

function clearSelections() {
  state.selectedVisualKeys.clear();
  const checks = dom.visualSelection.querySelectorAll("input[data-visual-checkbox]");
  checks.forEach((checkbox) => {
    checkbox.checked = false;
  });
  refreshPageCheckboxStates();
  updateSelectionCount();
}

async function loadThumbnailsForSelection() {
  if (!state.report && !state.demoMode) {
    throw new Error("Embed and load visuals first, or load Demo Mode.");
  }

  const selected = collectSelectedVisualsInOrder();
  if (!selected.length) {
    throw new Error("Select at least one visual.");
  }

  let processed = 0;
  const scale = clampNumber(parseFloat(dom.imageScaleInput.value), 1, 4, 2);

  for (const item of selected) {
    processed += 1;
    const key = makeVisualKey(item.page.name, item.visual.name);
    const domKey = encodeURIComponent(key);

    logStatus(`Thumbnail ${processed}/${selected.length}: ${item.visual.title || item.visual.name}`);

    const thumbWrap = dom.visualSelection.querySelector(`[data-thumb-wrap="${escapeSelector(domKey)}"]`);
    if (!thumbWrap) {
      continue;
    }

    const imageData = await getVisualImageData(item.page, item.visual, Math.max(1, scale - 0.5));
    thumbWrap.innerHTML = `<img alt="${escapeHtmlAttr(item.visual.title || item.visual.name)}" src="${imageData}" />`;
  }

  logStatus(`Loaded ${selected.length} thumbnails.`, "success");
}

async function generateDeck() {
  if (!window.PptxGenJS) {
    throw new Error("PptxGenJS is unavailable.");
  }

  if (!state.report && !state.demoMode) {
    throw new Error("Embed and load visuals first, or load Demo Mode.");
  }

  const selected = collectSelectedVisualsInOrder();
  if (!selected.length) {
    throw new Error("Select at least one visual.");
  }

  const pptx = new window.PptxGenJS();
  const layout = dom.slideLayoutInput.value || "LAYOUT_WIDE";
  const dimensions = layoutDimensions[layout] || layoutDimensions.LAYOUT_WIDE;
  const imageScale = clampNumber(parseFloat(dom.imageScaleInput.value), 1, 4, 2);
  const deckTitle = dom.deckTitleInput.value.trim() || "Power BI Executive Brief";

  pptx.layout = layout;
  pptx.author = "Power BI PPTX Generator";
  pptx.company = "Power BI";
  pptx.subject = "Executive briefing";
  pptx.title = deckTitle;

  addExecutiveCoverSlide(pptx, dimensions, deckTitle, selected.length, state.demoMode);
  addExecutiveSummarySlide(pptx, dimensions, selected);

  let activePageName = "";

  for (let index = 0; index < selected.length; index += 1) {
    const item = selected[index];
    const visualTitle = item.visual.title || item.visual.name;

    logStatus(`Exporting ${index + 1}/${selected.length}: ${visualTitle}`);

    if (!state.demoMode && item.page.name !== activePageName) {
      activePageName = item.page.name;
      await state.report.setPage(activePageName);
      await sleep(300);
    }

    const imageData = await getVisualImageData(item.page, item.visual, imageScale);
    addSlideForVisual(
      pptx,
      dimensions,
      item,
      imageData,
      dom.includePageNameInTitleInput.checked,
      index + 1,
      selected.length
    );
  }

  const prefix = sanitizeFileName(deckTitle || "powerbi-executive-brief");
  const fileName = `${prefix}-${createTimestampSlug()}.pptx`;

  await pptx.writeFile({ fileName });
  logStatus(`Deck generated: ${fileName}`, "success");
}

function addExecutiveCoverSlide(pptx, dimensions, deckTitle, selectedCount, demoMode) {
  const slide = pptx.addSlide();

  slide.background = { color: "F4F8FC" };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: dimensions.w,
    h: 0.48,
    fill: { color: "10324A" },
    line: { color: "10324A", pt: 0 },
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 1.02,
    w: dimensions.w - 1.1,
    h: 4.95,
    fill: { color: "FFFFFF", transparency: 0 },
    line: { color: "D3E0EB", pt: 1 },
    rectRadius: 0.08,
  });

  slide.addText("Executive Briefing Deck", {
    x: 0.8,
    y: 1.36,
    w: dimensions.w - 1.6,
    h: 0.35,
    fontFace: DECK_FONT_FAMILY,
    color: "5A738B",
    fontSize: 16,
    bold: true,
    charSpace: 1,
  });

  slide.addText(deckTitle, {
    x: 0.8,
    y: 1.88,
    w: dimensions.w - 1.6,
    h: 1.12,
    fontFace: DECK_FONT_FAMILY,
    color: "152F46",
    bold: true,
    fontSize: 35,
    fit: "resize",
    valign: "top",
  });

  const modeLabel = demoMode ? "DEMO MODE" : "LIVE POWER BI";
  const generatedOn = new Date().toLocaleString();
  const summaryLine = `Slides with visuals: ${selectedCount} | Source mode: ${modeLabel}`;

  slide.addText(summaryLine, {
    x: 0.8,
    y: 3.35,
    w: dimensions.w - 1.6,
    h: 0.35,
    fontFace: DECK_FONT_FAMILY,
    color: "3E5B74",
    fontSize: 13,
    bold: true,
  });

  slide.addText(`Generated ${generatedOn}`, {
    x: 0.8,
    y: 3.76,
    w: dimensions.w - 1.6,
    h: 0.3,
    fontFace: DECK_FONT_FAMILY,
    color: "5F7890",
    fontSize: 11,
  });

  slide.addText("Confidential - Executive audience", {
    x: 0.8,
    y: dimensions.h - 0.56,
    w: dimensions.w - 1.6,
    h: 0.2,
    fontFace: DECK_FONT_FAMILY,
    color: "7B8FA4",
    fontSize: 10,
    italic: true,
  });
}

function addExecutiveSummarySlide(pptx, dimensions, selected) {
  const slide = pptx.addSlide();

  slide.background = { color: "F8FBFE" };

  slide.addText("Executive Summary", {
    x: 0.52,
    y: 0.28,
    w: dimensions.w - 1.04,
    h: 0.42,
    fontFace: DECK_FONT_FAMILY,
    color: "1C3750",
    fontSize: 24,
    bold: true,
  });

  const pageCounts = new Map();
  const typeCounts = new Map();
  selected.forEach((item) => {
    const pageName = item.page.displayName || item.page.name;
    const typeName = item.visual.type || "visual";
    pageCounts.set(pageName, (pageCounts.get(pageName) || 0) + 1);
    typeCounts.set(typeName, (typeCounts.get(typeName) || 0) + 1);
  });

  const pageHighlights = [...pageCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 4)
    .map(([name, count]) => `${name}: ${count}`)
    .join(" | ");

  const topTypes = [...typeCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 4)
    .map(([name, count]) => `${name} (${count})`)
    .join(", ");

  const summaryPoints = [
    `Portfolio includes ${selected.length} selected visuals across ${pageCounts.size} report pages.`,
    `Primary visual mix: ${topTypes || "n/a"}.`,
    `Highest concentration pages: ${pageHighlights || "n/a"}.`,
    "Each following slide includes a single visual, context metadata, and an executive takeaway.",
  ];

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.52,
    y: 0.92,
    w: dimensions.w - 1.04,
    h: 5.98,
    fill: { color: "FFFFFF", transparency: 0 },
    line: { color: "D3E1ED", pt: 1 },
    rectRadius: 0.06,
  });

  summaryPoints.forEach((line, idx) => {
    slide.addText(`- ${line}`, {
      x: 0.8,
      y: 1.26 + idx * 0.88,
      w: dimensions.w - 1.6,
      h: 0.52,
      fontFace: DECK_FONT_FAMILY,
      color: "29455F",
      fontSize: 14,
      bold: idx < 2,
      fit: "shrink",
    });
  });
}

function buildExecutiveTakeaway(item) {
  const type = String(item.visual.type || "").toLowerCase();

  if (type.includes("line") || type.includes("area")) {
    return "Trend direction and inflection points should guide near-term planning and resourcing decisions.";
  }

  if (type.includes("bar") || type.includes("column") || type.includes("waterfall")) {
    return "Ranking and contribution differences highlight where leadership attention can improve outcomes fastest.";
  }

  if (type.includes("map")) {
    return "Regional dispersion suggests localized performance variance; prioritize top and underperforming territories.";
  }

  if (type.includes("table")) {
    return "Detailed record-level view supports risk validation and targeted follow-up on highest-value accounts.";
  }

  if (type.includes("scatter")) {
    return "Outliers indicate potential opportunity and risk clusters that merit executive review and mitigation plans.";
  }

  if (type.includes("card") || type.includes("kpi")) {
    return "KPI headline should be tracked against target and variance drivers in the immediate decision cycle.";
  }

  return "Visual signals a meaningful performance pattern; align owners and actions to confirm underlying drivers.";
}

function addSlideForVisual(pptx, dimensions, item, imageData, includePageName, slideIndex, totalSlides) {
  const slide = pptx.addSlide();
  const margin = 0.35;
  const titleBandHeight = 0.96;
  const calloutHeight = 0.9;
  const framePadding = 0.1;

  const titleText = includePageName
    ? `${item.page.displayName || item.page.name} - ${item.visual.title || item.visual.name}`
    : item.visual.title || item.visual.name;

  slide.background = { color: "F7FAFD" };

  slide.addShape(pptx.ShapeType.rect, {
    x: margin,
    y: 0.18,
    w: dimensions.w - margin * 2,
    h: 0.56,
    fill: { color: "E8F1F8", transparency: 8 },
    line: { color: "D2E0EC", pt: 1 },
  });

  slide.addText(titleText, {
    x: margin + 0.12,
    y: 0.28,
    w: dimensions.w - margin * 2 - 0.24,
    h: 0.24,
    fontFace: DECK_FONT_FAMILY,
    color: "17324A",
    bold: true,
    fontSize: 14,
    fit: "shrink",
  });

  slide.addText(`${item.visual.type || "visual"} | ${item.visual.name}`, {
    x: margin + 0.12,
    y: 0.52,
    w: dimensions.w - margin * 2 - 1.9,
    h: 0.13,
    fontFace: DECK_FONT_FAMILY,
    color: "43607A",
    fontSize: 9,
  });

  slide.addText(`Slide ${slideIndex + 2} of ${totalSlides + 2}`, {
    x: dimensions.w - margin - 1.6,
    y: 0.52,
    w: 1.45,
    h: 0.13,
    align: "right",
    fontFace: DECK_FONT_FAMILY,
    color: "5B7792",
    fontSize: 9,
    bold: true,
  });

  const imageContainer = {
    x: margin,
    y: titleBandHeight,
    w: dimensions.w - margin * 2,
    h: dimensions.h - titleBandHeight - margin - calloutHeight,
  };

  const visualAspectRatio = getVisualAspectRatio(item.visual);
  const fitted = fitRect(imageContainer, visualAspectRatio);

  slide.addShape(pptx.ShapeType.rect, {
    x: fitted.x - framePadding,
    y: fitted.y - framePadding,
    w: fitted.w + framePadding * 2,
    h: fitted.h + framePadding * 2,
    fill: { color: "FFFFFF", transparency: 0 },
    line: { color: "D5E1ED", pt: 1 },
  });

  slide.addImage({
    data: imageData,
    x: fitted.x,
    y: fitted.y,
    w: fitted.w,
    h: fitted.h,
  });

  const takeaway = buildExecutiveTakeaway(item);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: margin,
    y: dimensions.h - margin - calloutHeight,
    w: dimensions.w - margin * 2,
    h: calloutHeight,
    fill: { color: "EEF5FB", transparency: 0 },
    line: { color: "CDDFED", pt: 1 },
    rectRadius: 0.05,
  });

  slide.addText("Executive takeaway", {
    x: margin + 0.14,
    y: dimensions.h - margin - calloutHeight + 0.1,
    w: dimensions.w - margin * 2 - 0.28,
    h: 0.18,
    fontFace: DECK_FONT_FAMILY,
    color: "1A3B57",
    bold: true,
    fontSize: 10,
    charSpace: 0.5,
  });

  slide.addText(takeaway, {
    x: margin + 0.14,
    y: dimensions.h - margin - calloutHeight + 0.31,
    w: dimensions.w - margin * 2 - 0.28,
    h: 0.44,
    fontFace: DECK_FONT_FAMILY,
    color: "35506A",
    fontSize: 10,
    fit: "shrink",
  });
}

async function exportVisualAsImage(page, visual, scaleMultiplier) {
  const width = Math.max(Math.round((Number(visual.layout?.width) || 1280) * scaleMultiplier), 640);
  const height = Math.max(Math.round((Number(visual.layout?.height) || 720) * scaleMultiplier), 360);

  const calls = [];

  if (typeof state.report.exportVisualAsImage === "function") {
    calls.push(() => state.report.exportVisualAsImage(page.name, visual.name, width, height));
    calls.push(() => state.report.exportVisualAsImage(page.name, visual.name, { width, height }));
    calls.push(() => state.report.exportVisualAsImage({ pageName: page.name, visualName: visual.name, width, height }));
    calls.push(() => state.report.exportVisualAsImage(page.name, visual.name));
  }

  if (typeof visual.exportVisualAsImage === "function") {
    calls.push(() => visual.exportVisualAsImage(width, height));
    calls.push(() => visual.exportVisualAsImage({ width, height }));
    calls.push(() => visual.exportVisualAsImage());
  }

  if (typeof page.exportVisualAsImage === "function") {
    calls.push(() => page.exportVisualAsImage(visual.name, width, height));
    calls.push(() => page.exportVisualAsImage(visual.name, { width, height }));
  }

  if (!calls.length) {
    throw new Error(
      "exportVisualAsImage is not exposed by the embedded report in this environment. Confirm tenant feature support, permissions, and SDK capability."
    );
  }

  let lastError = null;
  for (const call of calls) {
    try {
      const result = await call();
      const dataUrl = await coerceToDataUrl(result);
      if (dataUrl) {
        return dataUrl;
      }
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`Failed to export visual as image: ${extractErrorMessage(lastError)}`);
}

async function coerceToDataUrl(value) {
  if (!value) {
    return "";
  }

  if (typeof value === "string") {
    return normalizeImageString(value);
  }

  if (value instanceof Blob) {
    return blobToDataUrl(value);
  }

  if (value instanceof ArrayBuffer) {
    return normalizeImageString(arrayBufferToBase64(value));
  }

  const candidates = [
    value.data,
    value.image,
    value.imageData,
    value.base64Image,
    value.base64,
    value.value,
    value.url,
    value.body?.data,
    value.body?.image,
    value.body?.imageData,
    value.body?.base64Image,
    value.body?.base64,
    value.body?.value,
    value.body?.url,
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (typeof candidate === "string") {
      if (/^https?:\/\//i.test(candidate)) {
        const fetched = await fetch(candidate);
        if (!fetched.ok) {
          continue;
        }
        const blob = await fetched.blob();
        return blobToDataUrl(blob);
      }
      const normalized = normalizeImageString(candidate);
      if (normalized) {
        return normalized;
      }
    }

    if (candidate instanceof Blob) {
      return blobToDataUrl(candidate);
    }

    if (candidate instanceof ArrayBuffer) {
      return normalizeImageString(arrayBufferToBase64(candidate));
    }
  }

  return "";
}

function normalizeImageString(raw) {
  const value = String(raw || "").trim();
  if (!value) {
    return "";
  }

  if (value.startsWith("data:image")) {
    return value;
  }

  const cleaned = value.replace(/\s+/g, "");
  if (/^[A-Za-z0-9+/=]+$/.test(cleaned)) {
    return `data:image/png;base64,${cleaned}`;
  }

  return "";
}

function arrayBufferToBase64(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function collectSelectedVisualsInOrder() {
  const selected = [];

  for (const pageGroup of state.pages) {
    for (const visual of pageGroup.visuals) {
      const key = makeVisualKey(pageGroup.page.name, visual.name);
      if (state.selectedVisualKeys.has(key)) {
        selected.push({
          page: pageGroup.page,
          visual,
        });
      }
    }
  }

  return selected;
}

function makeVisualKey(pageName, visualName) {
  return `${pageName}::${visualName}`;
}

function updateSelectionCount() {
  dom.selectionCountLabel.textContent = `Selected visuals: ${state.selectedVisualKeys.size}`;
}

function describeVisualDimensions(visual) {
  const width = Number(visual.layout?.width);
  const height = Number(visual.layout?.height);
  if (!width || !height) {
    return "size unknown";
  }
  return `${Math.round(width)}x${Math.round(height)}`;
}

function getVisualAspectRatio(visual) {
  const width = Number(visual.layout?.width);
  const height = Number(visual.layout?.height);
  if (!width || !height) {
    return 16 / 9;
  }
  return width / height;
}

function fitRect(bounds, aspectRatio) {
  let width = bounds.w;
  let height = width / aspectRatio;

  if (height > bounds.h) {
    height = bounds.h;
    width = height * aspectRatio;
  }

  return {
    x: bounds.x + (bounds.w - width) / 2,
    y: bounds.y + (bounds.h - height) / 2,
    w: width,
    h: height,
  };
}

function clampNumber(value, min, max, fallback) {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return Math.min(max, Math.max(min, value));
}

function sanitizeFileName(value) {
  return value.replace(/[^a-zA-Z0-9-_]+/g, "-").replace(/^-+|-+$/g, "");
}

function createTimestampSlug() {
  const now = new Date();
  const parts = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ];

  return `${parts[0]}${parts[1]}${parts[2]}-${parts[3]}${parts[4]}${parts[5]}`;
}

function escapeSelector(value) {
  if (window.CSS?.escape) {
    return window.CSS.escape(value);
  }
  return value.replace(/["\\#.:\[\]=]/g, "\\$&");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeHtmlAttr(value) {
  return escapeHtml(value).replaceAll("`", "&#96;");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resetSelectionState(message = "Report embedded. Click \"Load Pages and Visuals\".") {
  state.pages = [];
  state.visualIndex.clear();
  state.selectedVisualKeys.clear();
  dom.visualSelection.className = "visual-selection empty-state";
  dom.visualSelection.textContent = message;
  updateSelectionCount();
}

function extractErrorMessage(error) {
  if (!error) {
    return "Unknown error";
  }

  if (typeof error === "string") {
    return error;
  }

  return (
    error.message ||
    error.detailedMessage ||
    error.error?.message ||
    error.body?.message ||
    JSON.stringify(error)
  );
}

function logStatus(message, type = "info") {
  const entry = document.createElement("div");
  entry.className = `status-item ${type}`;
  const stamp = new Date().toLocaleTimeString();
  entry.textContent = `[${stamp}] ${message}`;
  dom.statusLog.prepend(entry);
}

async function getVisualImageData(page, visual, scaleMultiplier) {
  if (state.demoMode) {
    return generateDemoVisualAsImage(page, visual, scaleMultiplier);
  }

  return exportVisualAsImage(page, visual, scaleMultiplier);
}

function generateDemoVisualAsImage(page, visual, scaleMultiplier) {
  const width = Math.max(Math.round((Number(visual.layout?.width) || 1280) * scaleMultiplier), 640);
  const height = Math.max(Math.round((Number(visual.layout?.height) || 720) * scaleMultiplier), 360);

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Canvas context is unavailable for demo rendering.");
  }

  const palette = getDemoPalette(`${page.name}-${visual.name}`);

  const gradient = ctx.createLinearGradient(0, 0, width, height);
  gradient.addColorStop(0, palette.bgStart);
  gradient.addColorStop(1, palette.bgEnd);
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, width, height);

  const pad = Math.round(width * 0.04);
  const titleArea = Math.max(Math.round(height * 0.2), 90);

  ctx.fillStyle = "rgba(255,255,255,0.93)";
  ctx.fillRect(pad, pad, width - pad * 2, titleArea);

  ctx.fillStyle = "#17324a";
  ctx.font = `700 ${Math.max(18, Math.round(width * 0.024))}px ${DECK_CANVAS_FONT_STACK}`;
  ctx.fillText(visual.title || visual.name, pad + 16, pad + Math.round(titleArea * 0.45));

  ctx.fillStyle = "#4a6178";
  ctx.font = `500 ${Math.max(12, Math.round(width * 0.014))}px ${DECK_CANVAS_FONT_STACK}`;
  ctx.fillText(`${page.displayName || page.name} - ${visual.type || "visual"}`, pad + 16, pad + Math.round(titleArea * 0.75));

  const chartX = pad;
  const chartY = pad + titleArea + 18;
  const chartW = width - pad * 2;
  const chartH = height - chartY - pad;

  ctx.fillStyle = "rgba(255,255,255,0.9)";
  ctx.fillRect(chartX, chartY, chartW, chartH);

  const type = String(visual.type || "").toLowerCase();
  if (type.includes("line") || type.includes("area")) {
    drawDemoLineChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else if (type.includes("bar") || type.includes("column") || type.includes("waterfall")) {
    drawDemoBarChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else if (type.includes("pie") || type.includes("donut")) {
    drawDemoDonutChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else if (type.includes("map")) {
    drawDemoMapChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else if (type.includes("table")) {
    drawDemoTableChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else if (type.includes("scatter")) {
    drawDemoScatterChart(ctx, chartX, chartY, chartW, chartH, palette);
  } else {
    drawDemoKpiCard(ctx, chartX, chartY, chartW, chartH, palette);
  }

  return canvas.toDataURL("image/png");
}

function drawDemoLineChart(ctx, x, y, w, h, palette) {
  const baseY = y + h - 28;
  ctx.strokeStyle = palette.axis;
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(x + 20, y + 16);
  ctx.lineTo(x + 20, baseY);
  ctx.lineTo(x + w - 16, baseY);
  ctx.stroke();

  const points = [0.12, 0.28, 0.2, 0.44, 0.4, 0.62, 0.55, 0.75, 0.68, 0.82, 0.9, 0.7];
  ctx.strokeStyle = palette.primary;
  ctx.lineWidth = 4;
  ctx.beginPath();

  points.forEach((ratio, idx) => {
    const px = x + 30 + (idx * (w - 54)) / (points.length - 1);
    const py = baseY - ratio * (h - 58);
    if (idx === 0) {
      ctx.moveTo(px, py);
    } else {
      ctx.lineTo(px, py);
    }
  });
  ctx.stroke();

  ctx.fillStyle = palette.primarySoft;
  ctx.beginPath();
  ctx.moveTo(x + 30, baseY);
  points.forEach((ratio, idx) => {
    const px = x + 30 + (idx * (w - 54)) / (points.length - 1);
    const py = baseY - ratio * (h - 58);
    ctx.lineTo(px, py);
  });
  ctx.lineTo(x + w - 24, baseY);
  ctx.closePath();
  ctx.fill();
}

function drawDemoBarChart(ctx, x, y, w, h, palette) {
  const baseY = y + h - 26;
  const barCount = 8;
  const gap = 12;
  const barW = (w - 36 - gap * (barCount - 1)) / barCount;

  ctx.strokeStyle = palette.axis;
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(x + 18, y + 16);
  ctx.lineTo(x + 18, baseY);
  ctx.lineTo(x + w - 12, baseY);
  ctx.stroke();

  for (let i = 0; i < barCount; i += 1) {
    const magnitude = 0.25 + ((i * 37) % 61) / 100;
    const barH = magnitude * (h - 62);
    const bx = x + 24 + i * (barW + gap);
    const by = baseY - barH;

    ctx.fillStyle = i % 2 === 0 ? palette.primary : palette.secondary;
    ctx.fillRect(bx, by, barW, barH);
  }
}

function drawDemoDonutChart(ctx, x, y, w, h, palette) {
  const cx = x + w * 0.35;
  const cy = y + h * 0.5;
  const r = Math.min(w, h) * 0.28;
  const ring = r * 0.44;
  const slices = [0.28, 0.22, 0.18, 0.32];
  const colors = [palette.primary, palette.secondary, palette.accent, palette.primarySoft];

  let start = -Math.PI / 2;
  slices.forEach((part, idx) => {
    const end = start + part * Math.PI * 2;
    ctx.beginPath();
    ctx.strokeStyle = colors[idx % colors.length];
    ctx.lineWidth = ring;
    ctx.arc(cx, cy, r, start, end);
    ctx.stroke();
    start = end;
  });

  ctx.fillStyle = "#334e68";
  ctx.font = `700 ${Math.max(16, Math.round(w * 0.06))}px ${DECK_CANVAS_FONT_STACK}`;
  ctx.fillText("62%", cx - r * 0.3, cy + 8);
}

function drawDemoMapChart(ctx, x, y, w, h, palette) {
  ctx.fillStyle = "#edf4fa";
  ctx.fillRect(x + 20, y + 20, w - 40, h - 40);

  const regions = [
    [0.2, 0.35, 14],
    [0.35, 0.55, 10],
    [0.55, 0.42, 12],
    [0.68, 0.58, 15],
    [0.78, 0.33, 11],
  ];

  regions.forEach(([rx, ry, size], idx) => {
    const px = x + rx * w;
    const py = y + ry * h;
    ctx.beginPath();
    ctx.fillStyle = idx % 2 === 0 ? palette.primary : palette.secondary;
    ctx.arc(px, py, size, 0, Math.PI * 2);
    ctx.fill();

    ctx.strokeStyle = "rgba(51,78,104,0.25)";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(px, py);
    ctx.lineTo(x + w * 0.5, y + h * 0.5);
    ctx.stroke();
  });
}

function drawDemoTableChart(ctx, x, y, w, h, palette) {
  const rows = 7;
  const cols = 4;
  const cellW = (w - 24) / cols;
  const cellH = (h - 24) / rows;

  for (let r = 0; r < rows; r += 1) {
    for (let c = 0; c < cols; c += 1) {
      const cx = x + 12 + c * cellW;
      const cy = y + 12 + r * cellH;
      ctx.fillStyle = r === 0 ? "#d7e7f3" : r % 2 === 0 ? "#f8fbfe" : "#eef4fa";
      ctx.fillRect(cx, cy, cellW - 4, cellH - 4);

      if (r > 0 && c > 0) {
        ctx.fillStyle = (r + c) % 2 === 0 ? palette.primary : palette.secondary;
        ctx.fillRect(cx + 10, cy + Math.max(6, cellH * 0.28), Math.max(18, (cellW - 24) * ((r + c) % 3 + 1) * 0.25), 6);
      }
    }
  }
}

function drawDemoScatterChart(ctx, x, y, w, h, palette) {
  const baseX = x + 24;
  const baseY = y + h - 24;

  ctx.strokeStyle = palette.axis;
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(baseX, y + 18);
  ctx.lineTo(baseX, baseY);
  ctx.lineTo(x + w - 18, baseY);
  ctx.stroke();

  for (let i = 0; i < 24; i += 1) {
    const px = baseX + ((i * 37) % 91) / 100 * (w - 56);
    const py = baseY - ((i * 29) % 87) / 100 * (h - 52);
    const radius = 4 + ((i * 11) % 9);

    ctx.fillStyle = i % 2 === 0 ? palette.primary : palette.secondary;
    ctx.beginPath();
    ctx.arc(px, py, radius, 0, Math.PI * 2);
    ctx.fill();
  }
}

function drawDemoKpiCard(ctx, x, y, w, h, palette) {
  ctx.fillStyle = "#eef5fb";
  ctx.fillRect(x + 18, y + 18, w - 36, h - 36);

  ctx.fillStyle = palette.primary;
  ctx.font = `700 ${Math.max(26, Math.round(w * 0.1))}px ${DECK_CANVAS_FONT_STACK}`;
  ctx.fillText("$12.4M", x + 36, y + h * 0.52);

  ctx.fillStyle = "#4f6a83";
  ctx.font = `600 ${Math.max(14, Math.round(w * 0.035))}px ${DECK_CANVAS_FONT_STACK}`;
  ctx.fillText("Year-to-date revenue", x + 36, y + h * 0.68);
}

function getDemoPalette(seedValue) {
  const palettes = [
    {
      bgStart: "#f4f8ff",
      bgEnd: "#dce9f6",
      primary: "#1f6f8b",
      primarySoft: "rgba(31,111,139,0.24)",
      secondary: "#3f8fb8",
      accent: "#57b4ba",
      axis: "#7a95ad",
    },
    {
      bgStart: "#f7fbf5",
      bgEnd: "#e0f0e3",
      primary: "#2d7b55",
      primarySoft: "rgba(45,123,85,0.24)",
      secondary: "#4ea06f",
      accent: "#86bc7b",
      axis: "#7c9788",
    },
    {
      bgStart: "#fff8f2",
      bgEnd: "#f6e7d8",
      primary: "#ad5a2c",
      primarySoft: "rgba(173,90,44,0.24)",
      secondary: "#d1864b",
      accent: "#ebb26c",
      axis: "#a88b73",
    },
  ];

  const idx = Math.abs(hashString(seedValue)) % palettes.length;
  return palettes[idx];
}

function hashString(value) {
  let hash = 0;
  const text = String(value || "");
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash << 5) - hash + text.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}
