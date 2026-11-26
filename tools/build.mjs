// Builds the hosted task pane and the manifests.
//
//   node tools/build.mjs                    # dist/: the bundle for https://converter.studio/integrations/office/
//                                           # (addin.config.json), its manifests, and manifest/*.xml
//   node tools/build.mjs --base URL --out D # the same for another address (a local test server)
//   node tools/build.mjs --manifests        # only manifest/outlook.xml and manifest/office.xml
//
// dist/ is what is deployed at the base: taskpane.html, commands.html, taskpane.css, taskpane.js,
// commands.js, legacy.js, assets/icon-*.png, vendor/libheif-js-1.23.5/ (the HEIC decoder, as published
// on npm), and manifest/ (the two manifests for that address).
// manifest/ in this folder holds the production manifests, committed: they are what Partner
// Center gets, and the tests check they match this generator.
import { build } from 'esbuild';
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { normalBase, officeManifest, outlookManifest } from './manifest.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const config = JSON.parse(readFileSync(join(root, 'addin.config.json'), 'utf8'));
const { version } = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

function arg(name) {
  const index = process.argv.indexOf(name);
  return index === -1 ? undefined : process.argv[index + 1];
}

/** Sizes the manifests point at (ribbon 16/32/80, task pane 32/64, Outlook 64/128). */
const ICON_SIZES = [16, 32, 64, 80, 128];

function writeManifests(folder, base) {
  mkdirSync(folder, { recursive: true });
  writeFileSync(join(folder, 'outlook.xml'), outlookManifest(config, version, base));
  writeFileSync(join(folder, 'office.xml'), officeManifest(config, version, base));
}

// The production manifests, always from the configured base.
writeManifests(join(root, 'manifest'), config.baseUrl);
if (process.argv.includes('--manifests')) process.exit(0);

const base = normalBase(arg('--base') ?? config.baseUrl);
const out = resolve(root, arg('--out') ?? 'dist');
rmSync(out, { recursive: true, force: true });
mkdirSync(join(out, 'assets'), { recursive: true });

await build({
  entryPoints: { taskpane: join(root, 'src/main.ts'), commands: join(root, 'src/commands.ts') },
  outdir: out,
  bundle: true,
  format: 'iife',
  // Microsoft 365's webviews (WebView2, Safari) and the browsers Office on the web supports;
  // es2019 also keeps Office 2019's EdgeHTML webview working.
  target: 'es2019',
  minify: !process.argv.includes('--dev'),
  legalComments: 'none',
  logLevel: 'warning',
});
for (const file of ['taskpane.html', 'commands.html', 'taskpane.css', 'legacy.js'])
  copyFileSync(join(root, 'src', file), join(out, file));
for (const size of ICON_SIZES)
  copyFileSync(join(root, 'assets', `icon-${size}.png`), join(out, 'assets', `icon-${size}.png`));

// libheif-js (LGPL-3.0), the HEIC decoder for "Fit attachments for sending": its WebAssembly build
// as published on npm, unchanged, in its own folder and loaded only when a HEIC photo needs it
// (src/fit/heic.ts). The version is pinned here and in package.json.
const LIBHEIF = { version: '1.23.5', folder: join(root, 'node_modules', 'libheif-js') };
const libheifVersion = JSON.parse(
  readFileSync(join(LIBHEIF.folder, 'package.json'), 'utf8'),
).version;
if (libheifVersion !== LIBHEIF.version)
  throw new Error(`libheif-js ${libheifVersion} installed, ${LIBHEIF.version} expected (npm ci)`);
const vendor = join(out, 'vendor', `libheif-js-${LIBHEIF.version}`);
mkdirSync(join(vendor, 'libheif-wasm'), { recursive: true });
for (const file of [
  'LICENSE',
  'README.md',
  'package.json',
  'libheif-wasm/LICENSE',
  'libheif-wasm/libheif.js',
  'libheif-wasm/libheif.wasm',
])
  copyFileSync(join(LIBHEIF.folder, file), join(vendor, file));
writeManifests(join(out, 'manifest'), base);
console.log(`built ${out} for ${base}`);
