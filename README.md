# Pi4Word

**Pi4Word** is a Microsoft Word task pane add-in that embeds a Pi Agent-powered assistant using [`@mariozechner/pi-agent-core`](https://www.npmjs.com/package/@mariozechner/pi-agent-core).

<p align="center">
  <img src="./docs/assets/screenshot-panel.jpg" alt="Pi4Word task pane" width="400"/>
</p>

For architecture, runtime behavior, tools, storage, and streaming details, see [SPECS.md](./SPECS.md).

## Manifests

| File | Purpose | Task pane URL |
| --- | --- | --- |
| `manifest.production.xml` | Production usage in Word | `https://damianofalcioni.github.io/pi-for-word/public/index.html` |
| `manifest.xml` | Local development and Word sideloading | `https://localhost:3000/public/index.html` |

## Requirements

- Node.js and npm for local development
- Desktop Microsoft Word for local sideloading
- HTTPS on `https://localhost:3000` for local development

## Production Usage

Use [`manifest.production.xml`](https://damianofalcioni.github.io/pi-for-word/manifest.production.xml) to load the production add-in in Word. This manifest points Word to the public GitHub Pages build at `https://damianofalcioni.github.io/pi-for-word/public/`.

### Word on the Web

1. Open a document in Word on the web.
2. Open **Insert > Add-ins > Advanced...**.
3. Choose **Upload My Add-in**.
4. Select [`manifest.production.xml`](https://damianofalcioni.github.io/pi-for-word/manifest.production.xml).

After the add-in is loaded, use **Open Pi4Word** from the Word ribbon.

### Word on Windows (Workaround)

1. Run `npx -y office-addin-debugging start manifest.production.xml desktop --app word`

Pi4Word is available under **Home > Add-ins > Developer Add-ins**.

### Word on Windows (Official)

1. Create a local folder for the add-in catalog, for example `C:\OfficeAddinCatalog`.
2. Copy [`manifest.production.xml`](https://damianofalcioni.github.io/pi-for-word/manifest.production.xml) into that folder.
3. Right-click the folder, open **Properties > Sharing > Advanced Sharing**, enable **Share this folder**, and make sure your Windows user has at least read access.
4. Copy the folder's network path from the **Sharing** tab, for example `\\YOUR-PC\OfficeAddinCatalog`. Use this network path in Word, not the local `C:\...` path.
5. In Word, open **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**.
6. Paste the network path into **Catalog Url**, select **Add catalog**, enable **Show in Menu** for that catalog, select **OK**, then restart Word.
7. Open **Home > Add-ins > Shared Folder**, select **Pi4Word**, and add it to the document.

After the add-in is loaded, use **Open Pi4Word** from the Word ribbon.

## Local Development

```bash
npm install
npm run build
npm run cert        (windows only)
npm run word
```

`npm run word` starts the HTTPS dev server and sideloads `manifest.xml` into Word.

## Build

```bash
npm run build
```

The build runs tests, linting, and `scripts/esbuild.mjs`. Output is written to `public/`, including:

- `public/index.html`
- `public/index.min.js`
- `public/index.min.css`
- copied runtime assets under `public/assets/` and `public/pdfjs-dist/`

## Scripts

| Command | Description |
| --- | --- |
| `npm run cert` | Creates local HTTPS certificate material on Windows. |
| `npm run serve` | Serves the repository over HTTPS on port `3000`. |
| `npm run word` | Starts Word sideloading with `manifest.xml`. |
| `npm run test` | Runs Node tests. |
| `npm run eslint` | Runs ESLint. |
| `npm run esbuild` | Bundles the task pane into `public/`. |
| `npm run build` | Runs tests, linting, and bundling. |

## License

MIT
