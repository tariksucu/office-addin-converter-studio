// Builds the hosted task pane and the manifests.
//
//   node tools/build.mjs                    # dist/: the bundle for https://converter.studio/integrations/office/
//                                           # (addin.config.json), its manifests, and manifest/*.xml
//   node tools/build.mjs --base URL --out D # the same for another address (a local test server)
//   node tools/build.mjs --manifests        # only manifest/outlook.xml and manifest/office.xml
//
// dist/ is what is deployed at the base: taskpane.html, commands.html, taskpane.css, taskpane.js,
// commands.js, legacy.js, assets/icon-*.png, and manifest/ (the two manifests for that address).
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
writeManifests(join(out, 'manifest'), base);
console.log(`built ${out} for ${base}`);
