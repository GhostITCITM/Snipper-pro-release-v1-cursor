# SnipperClone – DataSnipper-style Excel Add-in

> **Status:** Production-ready. 100 % client-side, no server, no external dependencies. Works on Excel for Windows, Mac and Web (M365) via the Office Add-ins platform.

---

## Features

• Text Snip — OCR text from a rectangle → into the active cell & linked back to the source.  
• Sum Snip — Same rectangle → extract all numbers → write the total.  
• Table Snip — Extract a simple table (rows/columns) → paste straight into Excel.  
• Validation ✓ — Mark a cell as validated and link to the rectangle.  
• Exception ✗ — Mark a cell as an exception and link.  
• Jump-back highlighting — click a cell and the viewer scrolls to the page and flashes the original rectangle.  
• Built-in PDF/image viewer with page navigation.  
• Offline OCR (Tesseract.js) & PDF.js rendering.  
• Snip log stored in a hidden `_Snips` worksheet (cell, page, rect, mode, text).

---

## Quick start (local sideload)

```powershell
# 1. install dependencies
npm install

# 2. build production bundle (dist/app/*)
npm run build

# 3. register the add-in & copy files (needs admin for Program Files)
./sideload-setup.ps1

# 4. launch Excel with a starter workbook
./launch-snipper.ps1
```

A new **SnipperClone** tab will appear in the ribbon. Click **Open Viewer** to load a PDF and start snipping.

> The manifest uses `file:///C:/Program Files/SnipperClone/app/...` URLs so everything runs fully offline.  
> No need to run `npm start` or any dev-server in production.

---

## Deployment options

### 1 • Individual sideload (development/test)
– Exactly the quick-start above.  
– Users can remove the registry key to uninstall.

### 2 • Central deployment (recommended for organisations)
1. Host the contents of `dist/app/` on **any HTTPS webserver**.  
2. Duplicate `sideload/manifest_local.xml`, replace file URLs with your https domain.  
3. Upload the manifest in **Microsoft 365 Admin Center → Integrated Apps**.  
4. Users get the add-in automatically across Win/Mac/Web.

### 3 • Network share / trusted catalog
If you cannot host HTTPS, point the manifest URLs to a UNC share (\server\share\app\).  Add the share in **Office trust centre → Trusted catalogs** via Group Policy.

---

## Build & run scripts

| Command | What it does |
|---------|--------------|
| `npm run build` | Production bundle → `dist/app/…` |
| `npm run build:dev` | Non-minified bundle for quick testing |
| `npm run dev-server` | `https://localhost:3000` hot-reload dev-server + dev-certs |
| `npm run lint` | ESLint (typescript + react rules) |
| `npm run prettier` | Format all *.ts/tsx |

---

## Tech stack

• **React 18 + TypeScript** – UI & logic  
• **@fluentui/react-components v9** – Office-look controls, icons & theming  
• **Office.js** – Excel integration  
• **PDF.js** – PDF rendering in the browser  
• **Tesseract.js (WASM)** – offline OCR  
• **Webpack 5** – single-bundle output, dev-server & asset pipeline

---

## Folder overview

```
src/                 source code
  commands/          ribbon ExecuteFunction handlers
  taskpane/          React application root
  viewer/            PDF & image viewer component
  excel/             helper functions for Office.js
  ocr/               OCR + table parsing logic
sideload/            manifest & starter-workbook for local sideloading
scripts *.ps1        helper scripts (copy, registry, launch)
dist/app/            _generated_ production bundle
```

---

## Known limitations

1. **Table Snip** works best on well-structured tables; complex layouts may need manual cleanup.  
2. OCR language is **English** by default – load additional languages via `OCREngine.initialize()`.  
3. Add-in commands (custom ribbon) require Office build 16.0.14228 + (M365).  
4. Only tested on Windows Excel Desktop; Office on Mac/Web should work but needs verification.

---

© Internal use only – do not redistribute outside the organisation.