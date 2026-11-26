/**
 * build_standalone.mjs
 * Genera dist/Optativas_Boletin.html: versión independiente "solo Excel" para colegas.
 *
 *   - Un único archivo HTML; se abre con doble clic (file://), sin servidor ni internet.
 *   - JSX precompilado (no usa Babel en el navegador).
 *   - React, ReactDOM, JSZip y SheetJS incrustados (versiones fijas en package.json).
 *   - CSS y tipografías (Crimson Pro, DM Mono) incrustadas.
 *   - window.OPTI_EXCEL_ONLY = true → oculta la lectura de documentos Word.
 *
 * Uso (desde la raíz del repo):
 *   npm install
 *   npm run build:standalone
 *
 * Fuentes: js/optativas_boletin.js, js/docx_parser.js, css/optativas_boletin.css
 * (las mismas que usa src/optativas_boletin.html, que no se modifica).
 */
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const OUT_DIR = path.join(ROOT, "dist");
const OUT = path.join(OUT_DIR, "Optativas_Boletin.html");

const read = p => fs.readFileSync(p, "utf8");
const pkgFile = (pkg, rel) => path.join(path.dirname(require.resolve(`${pkg}/package.json`)), rel);
const version = pkg => JSON.parse(read(pkgFile(pkg, "package.json"))).version;

/** Evita que un "</script>" dentro del código cierre la etiqueta antes de tiempo. */
const safeJs = s => s.replace(/<\/script/gi, "<\\/script");

// ── 1. Compilar el JSX de la app ───────────────────────────────────────────
const Babel = require("@babel/standalone");
const appSrc = read(path.join(ROOT, "js", "optativas_boletin.js"));
const appJs = Babel.transform(appSrc, { presets: ["react"], comments: false }).code;

// ── 2. Librerías ───────────────────────────────────────────────────────────
const libs = [
  ["react",     "umd/react.production.min.js"],
  ["react-dom", "umd/react-dom.production.min.js"],
  ["jszip",     "dist/jszip.min.js"],
  ["xlsx",      "dist/xlsx.full.min.js"],
].map(([pkg, rel]) => ({ name: `${pkg}@${version(pkg)}`, code: read(pkgFile(pkg, rel)) }));

// ── 3. Tipografías (subconjunto latin, cubre español) ──────────────────────
const fonts = [
  ["@fontsource/crimson-pro", "Crimson Pro", "crimson-pro", [300, 400, 500, 600]],
  ["@fontsource/dm-mono",     "DM Mono",     "dm-mono",     [400, 500]],
].flatMap(([pkg, family, slug, weights]) => weights.map(w => {
  const b64 = fs.readFileSync(pkgFile(pkg, `files/${slug}-latin-${w}-normal.woff2`)).toString("base64");
  return `@font-face{font-family:'${family}';font-style:normal;font-weight:${w};font-display:swap;` +
         `src:url(data:font/woff2;base64,${b64}) format('woff2');}`;
})).join("\n");

const css = read(path.join(ROOT, "css", "optativas_boletin.css"));
const parserJs = read(path.join(ROOT, "js", "docx_parser.js"));

// ── 4. Ensamblar ───────────────────────────────────────────────────────────
const built = new Date().toISOString().slice(0, 10);
const html = `<!DOCTYPE html>
<!--
  Generador de Boletín de Optativas — ITAM (versión independiente, solo Excel)
  Generado el ${built} por build/build_standalone.mjs — no editar a mano.
  Abrir con doble clic en Chrome, Edge o Firefox. No requiere internet ni instalación.
  Incluye: ${libs.map(l => l.name).join(", ")}
-->
<html lang="es">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Boletín de Optativas — ITAM</title>
<style>
${fonts}
${css}
</style>
</head>
<body>
<div id="root"></div>
<noscript>Esta herramienta necesita JavaScript habilitado.</noscript>
${libs.map(l => `<script>/* ${l.name} */\n${safeJs(l.code)}\n</script>`).join("\n")}
<script>window.OPTI_EXCEL_ONLY = true;</script>
<script>/* docx_parser.js */\n${safeJs(parserJs)}\n</script>
<script>/* optativas_boletin.js (compilado) */\n${safeJs(appJs)}\n</script>
</body>
</html>
`;

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(OUT, html, "utf8");
console.log(`✓ ${path.relative(ROOT, OUT)}  (${(Buffer.byteLength(html) / 1024).toFixed(0)} KB)`);
