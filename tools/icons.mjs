// Draws the add-ins' icons from Converter Studio's mark (an open square over an orange one, on
// its cream tile, which reads on light and dark ribbons alike) at every size the manifests and
// Partner Center ask for, with rsvg-convert and optipng.
//
//   node tools/icons.mjs
import { execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const assets = join(dirname(fileURLToPath(import.meta.url)), '..', 'assets');
const MARK = `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 32 32">
  <rect width="32" height="32" rx="3" fill="#f2efe8"/>
  <rect x="5" y="5" width="15" height="15" fill="none" stroke="#16150f" stroke-width="2.2"/>
  <rect x="12" y="12" width="15" height="15" fill="#ff5418"/>
</svg>
`;
// Ribbon buttons (16, 32, 80), the task pane add-in's icons (32, 64), the Outlook add-in's (64,
// 128) and the Marketplace listing's logo (300).
const SIZES = [16, 32, 64, 80, 128, 300];

mkdirSync(assets, { recursive: true });
writeFileSync(join(assets, 'icon.svg'), MARK);
for (const size of SIZES) {
  const out = join(assets, `icon-${size}.png`);
  execFileSync('rsvg-convert', [
    '-w',
    String(size),
    '-h',
    String(size),
    '-o',
    out,
    join(assets, 'icon.svg'),
  ]);
  execFileSync('optipng', ['-quiet', '-o5', '-strip', 'all', out]);
}
console.log(`icons written to ${assets}`);
