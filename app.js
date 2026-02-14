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
  busy: false,
};

const layoutDimensions = {
  LAYOUT_WIDE: { w: 13.333, h: 7.5 },
  LAYOUT_STANDARD: { w: 10, h: 7.5 },
};

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
  if (!state.report) {
    throw new Error("Embed and load visuals first.");
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

    const imageData = await exportVisualAsImage(item.page, item.visual, Math.max(1, scale - 0.5));
    thumbWrap.innerHTML = `<img alt="${escapeHtmlAttr(item.visual.title || item.visual.name)}" src="${imageData}" />`;
  }

  logStatus(`Loaded ${selected.length} thumbnails.`, "success");
}

async function generateDeck() {
  if (!window.PptxGenJS) {
    throw new Error("PptxGenJS is unavailable.");
  }

  if (!state.report) {
    throw new Error("Embed and load visuals first.");
  }

  const selected = collectSelectedVisualsInOrder();
  if (!selected.length) {
    throw new Error("Select at least one visual.");
  }

  const pptx = new window.PptxGenJS();
  const layout = dom.slideLayoutInput.value || "LAYOUT_WIDE";
  const dimensions = layoutDimensions[layout] || layoutDimensions.LAYOUT_WIDE;
  const imageScale = clampNumber(parseFloat(dom.imageScaleInput.value), 1, 4, 2);

  pptx.layout = layout;
  pptx.author = "Power BI PPTX Generator";
  pptx.company = "Power BI";
  pptx.subject = "Automated visual export";
  pptx.title = dom.deckTitleInput.value.trim() || "Power BI Deck";

  let activePageName = "";

  for (let index = 0; index < selected.length; index += 1) {
    const item = selected[index];
    const visualTitle = item.visual.title || item.visual.name;

    logStatus(`Exporting ${index + 1}/${selected.length}: ${visualTitle}`);

    if (item.page.name !== activePageName) {
      activePageName = item.page.name;
      await state.report.setPage(activePageName);
      await sleep(300);
    }

    const imageData = await exportVisualAsImage(item.page, item.visual, imageScale);
    addSlideForVisual(pptx, dimensions, item, imageData, dom.includePageNameInTitleInput.checked);
  }

  const prefix = sanitizeFileName(dom.deckTitleInput.value.trim() || "powerbi-export");
  const fileName = `${prefix}-${createTimestampSlug()}.pptx`;

  await pptx.writeFile({ fileName });
  logStatus(`Deck generated: ${fileName}`, "success");
}

function addSlideForVisual(pptx, dimensions, item, imageData, includePageName) {
  const slide = pptx.addSlide();
  const margin = 0.35;
  const titleBandHeight = 0.9;
  const framePadding = 0.1;

  const titleText = includePageName
    ? `${item.page.displayName || item.page.name} - ${item.visual.title || item.visual.name}`
    : item.visual.title || item.visual.name;

  slide.background = { color: "F7FAFD" };

  slide.addShape(pptx.ShapeType.rect, {
    x: margin,
    y: 0.18,
    w: dimensions.w - margin * 2,
    h: 0.52,
    fill: { color: "E8F1F8", transparency: 8 },
    line: { color: "D2E0EC", pt: 1 },
  });

  slide.addText(titleText, {
    x: margin + 0.12,
    y: 0.29,
    w: dimensions.w - margin * 2 - 0.24,
    h: 0.22,
    fontFace: "Aptos",
    color: "17324A",
    bold: true,
    fontSize: 14,
    fit: "shrink",
  });

  slide.addText(`${item.visual.type || "visual"} - ${item.visual.name}`, {
    x: margin + 0.12,
    y: 0.52,
    w: dimensions.w - margin * 2 - 0.24,
    h: 0.13,
    fontFace: "Aptos",
    color: "43607A",
    fontSize: 9,
  });

  const imageContainer = {
    x: margin,
    y: titleBandHeight,
    w: dimensions.w - margin * 2,
    h: dimensions.h - titleBandHeight - margin,
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

function resetSelectionState() {
  state.pages = [];
  state.visualIndex.clear();
  state.selectedVisualKeys.clear();
  dom.visualSelection.className = "visual-selection empty-state";
  dom.visualSelection.textContent = "Report embedded. Click \"Load Pages and Visuals\".";
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
