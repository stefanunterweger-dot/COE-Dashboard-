import * as esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

const here = path.dirname(new URL(import.meta.url).pathname);
const nm = path.join(here, 'node_modules');

/** Alle Imports aus dem Repo über die node_modules der Vorschau auflösen; SPFx-Pakete durch Shims ersetzen. */
const resolvePlugin = {
  name: 'preview-resolve',
  setup(build) {
    build.onResolve({ filter: /^@microsoft\/sp-http$/ }, () => ({ path: path.join(here, 'shims/sp-http.ts') }));
    build.onResolve({ filter: /^@microsoft\/sp-component-base$/ }, () => ({ path: path.join(here, 'shims/sp-component-base.ts') }));
    build.onResolve({ filter: /^(react|react-dom|@fluentui\/react)(\/.*)?$/ }, async args => {
      if (args.pluginData === 'x') { return undefined; }
      const r = await build.resolve(args.path, { resolveDir: here, kind: args.kind, pluginData: 'x' });
      return r;
    });
  }
};

/**
 * Styles genau so verwenden, wie SharePoint sie ausliefert: statt die .scss selbst zu übersetzen,
 * wird das vom SPFx-Build kompilierte CSS-Modul aus lib/ genommen (vorher „npx gulp bundle“ im Projekt).
 * So zeigt die Vorschau dieselben Klassennamen und dieselben CSS-Optimierungen wie in SharePoint.
 */
const spfxCssPlugin = {
  name: 'spfx-compiled-css',
  setup(build) {
    build.onResolve({ filter: /\.module\.scss$/ }, args => {
      const srcFile = path.resolve(args.resolveDir, args.path);
      const libFile = srcFile.replace(`${path.sep}src${path.sep}`, `${path.sep}lib${path.sep}`) + '.js';
      if (!fs.existsSync(libFile)) {
        throw new Error(`Kompiliertes CSS fehlt: ${libFile} – zuerst „npx gulp bundle“ im Projektordner ausführen.`);
      }
      return { path: libFile };
    });
  }
};

/** Icon-Schriften der Fluent UI (MDL2) als data-URI einbetten statt vom CDN zu laden. */
const fontPlugin = {
  name: 'inline-icon-fonts',
  setup(build) {
    build.onLoad({ filter: /font-icons-mdl2[\\/]lib[\\/]fabric-icons(-\d+)?\.js$/ }, args => {
      let src = fs.readFileSync(args.path, 'utf8');
      const fontsDir = path.join(path.dirname(args.path), '..', 'fonts');
      src = src.replace(/"url\('"\.concat\(baseUrl, "([^"]+\.woff)'\) format\('woff'\)"\)/, (_m, file) => {
        const b64 = fs.readFileSync(path.join(fontsDir, file)).toString('base64');
        return `"url('data:font/woff;base64,${b64}') format('woff')"`;
      });
      return { contents: src, loader: 'js' };
    });
  }
};

await esbuild.build({
  entryPoints: [path.join(here, 'main.tsx')],
  bundle: true,
  minify: true,
  format: 'iife',
  target: 'es2019',
  outfile: path.join(here, 'dist/app.js'),
  define: { 'process.env.NODE_ENV': '"production"' },
  nodePaths: [nm],
  tsconfigRaw: { compilerOptions: { jsx: 'react' } },
  // SPFx hat die CSS-Klassen bereits eindeutig benannt – unverändert übernehmen
  loader: { '.module.css': 'global-css' },
  plugins: [resolvePlugin, fontPlugin, spfxCssPlugin],
  logLevel: 'warning'
});

const js = fs.readFileSync(path.join(here, 'dist/app.js'), 'utf8');
const cssFile = path.join(here, 'dist/app.css');
const css = fs.existsSync(cssFile) ? fs.readFileSync(cssFile, 'utf8') : '';
const shell = fs.readFileSync(path.join(here, 'shell.css'), 'utf8');
const html = `<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CoE Cockpit Vorschau</title>
<style>${shell}</style>
<style>${css}</style>
</head>
<body>
<div id="root"></div>
<script>${js.replace(/<\/script/gi, '<\\/script')}</script>
</body>
</html>`;
fs.writeFileSync(path.join(here, 'dist/coe-cockpit-vorschau.html'), html);
console.log('OK', (html.length / 1024).toFixed(0), 'KB');
