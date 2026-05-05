# Pi4Word

**Pi4Word** is a Microsoft Word task pane add-in that embeds a Pi Agent–powered assistant ([`@mariozechner/pi-agent-core`](https://www.npmjs.com/package/@mariozechner/pi-agent-core)).

**Architecture, behavior, tools, storage, and streaming:** see **[SPECS.md](./SPECS.md)**.

## Local development

You need Node.js, desktop **Word**, and HTTPS on `**https://localhost:3000`** (manifest default). On Windows, `**npm run cert`** helps create dev TLS material; `**npm run serve**` expects `**certs/cert.pem**` — check `**scripts/setup-cert.ps1**` and the `**serve**` script in `**package.json**` for how those fit together.

```bash
npm install
npm run build
npm run cert    # once, if needed (windows only)
npm run word    # sideloads manifest into Word
```

Other npm scripts (`**esbuild**`, `**test**`, `**eslint**`, …) are listed in `**package.json**`.

## LICENSE

MIT

