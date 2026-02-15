# Power BI PPTX Generator (Bitbucket Pages)

Client-side JavaScript app that embeds a Power BI report, lets users pick report pages/visuals, and generates a downloadable PowerPoint deck with one visual per slide.

## What this app does

- Embeds a Power BI report with the **Power BI JavaScript API**.
- Loads selectable visuals via:
  - `report.getPages()`
  - `page.getVisuals()`
- Supports checklist-based selection (page-level and visual-level).
- Exports selected visuals as images via **`exportVisualAsImage`** and inserts each image into a slide.
- Builds `.pptx` in-browser with **PptxGenJS**.
- Includes a **Demo Mode** that generates mock visuals for deck layout testing without authentication.
- Uses a SpaceX-inspired executive template (dark space background, icon cards, and branded footer treatment) for cover, summary, and visual slides.
- Applies `D-DIN` as the presentation font for all generated slide text.
- Preserves visual aspect ratio from Power BI layout metadata and auto-generates slide titles from visual/page metadata.

## Project files

- `/Users/kevinarmstrong/powerbi-pptx-generator/index.html` - UI and CDN dependencies
- `/Users/kevinarmstrong/powerbi-pptx-generator/styles.css` - styling
- `/Users/kevinarmstrong/powerbi-pptx-generator/app.js` - embed, selection, export logic
- `/Users/kevinarmstrong/powerbi-pptx-generator/assets/template-spacex/` - template images and icon assets extracted from `Presentation2.pptx`

## Run locally

Because this is a static app, use any static file server:

```bash
cd /Users/kevinarmstrong/powerbi-pptx-generator
python3 -m http.server 8080
```

Open [http://localhost:8080](http://localhost:8080).

## Authentication options (static-host friendly)

### Option A: Paste token manually

1. Obtain a short-lived Azure AD token or embed token from your secure auth flow.
2. Paste the token into **Power BI Access Token**.
3. Choose token type:
   - **Azure AD** for user-owns-data
   - **Embed** for app-owns-data token

### Option B: Azure AD sign-in with MSAL (browser)

1. Create an Azure App Registration (SPA/public client).
2. Add your URL as a redirect URI:
   - local: `http://localhost:8080`
   - Bitbucket Pages: `https://powerbi-to-pptx.bitbucket.io/`
3. Grant delegated Power BI permissions (for example `Report.Read.All`) and consent.
4. In the app, choose **Sign in with Azure AD (MSAL)** and provide:
   - Tenant ID
   - Client ID
   - Cloud Environment
   - Scopes
5. Click **Sign in**.

### Sovereign cloud values

Use the **Cloud Environment** selector to apply the correct authority/scopes for your cloud:

- Commercial: `https://analysis.windows.net/powerbi/api/Report.Read.All`
- US Gov GCC: `https://analysis.usgovcloudapi.net/powerbi/api/Report.Read.All`
- US Gov GCC High: `https://high.analysis.usgovcloudapi.net/powerbi/api/Report.Read.All`
- US Gov DoD: `https://mil.analysis.usgovcloudapi.net/powerbi/api/Report.Read.All`
- China: `https://analysis.chinacloudapi.cn/powerbi/api/Report.Read.All`

### Important security note

Bitbucket Pages has no backend runtime. Do **not** place client secrets in this app. If you need service principal tokens, mint them in a secure backend/token broker and pass short-lived tokens to the frontend.

## Deploy to Bitbucket Pages

1. Create a public Bitbucket repository named `powerbi-to-pptx.bitbucket.io` in workspace `powerbi-to-pptx`.
2. Push this repository content to Bitbucket (keep GitHub live during transition).
3. In Bitbucket repository settings, enable **Pages** from `main` branch and root path.
4. Confirm publish at `https://powerbi-to-pptx.bitbucket.io/`.
5. Add `https://powerbi-to-pptx.bitbucket.io/` to your Azure App redirect URIs (if MSAL login is used).
6. During dual-host transition, keep existing GitHub Pages redirect URI configured until Bitbucket validation is complete.

## Usage workflow

### Live Power BI mode

1. Fill token + report embed settings.
2. Click **Embed Report**.
3. Click **Load Pages and Visuals**.
4. Select desired visuals (or Select All).
5. Optional: click **Load Thumbnails** to preview selected visuals.
6. Click **Generate PPTX** to download the deck.

### Demo mode (no login)

1. Click **Load Demo Mode (No Login)**.
2. Keep all demo visuals selected or adjust selection.
3. Optional: click **Load Thumbnails** to preview rendered mock visuals.
4. Click **Generate PPTX** to download a sample deck for layout review.

## Troubleshooting

- **`exportVisualAsImage` not exposed**:
  - Your tenant/capabilities or embed context may not expose this API.
  - Confirm report permissions and that your embedding scenario supports visual image export.
- **Auth popup errors**:
  - Validate redirect URI and API permissions in Azure App Registration.
- **No visuals listed**:
  - Verify the report has accessible pages and loaded successfully.

## References

- [Power BI JavaScript SDK](https://learn.microsoft.com/javascript/api/overview/powerbi/)
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
- [MSAL Browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)

## D-DIN font note

The generated PPTX sets slide text to `D-DIN`. For exact rendering, `D-DIN` must be installed on the machine opening the deck in PowerPoint. Visual images exported from Power BI keep the font styling from the source report image.
