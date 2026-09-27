#!/usr/bin/env node
/**
 * Export the banner's inline SVG icons and body silhouettes as transparent PNGs
 * (canva-assets/) so build_canva_pptx.py can place them as movable images.
 *
 * Usage:  npm i puppeteer-core && CHROME_PATH=/path/to/chrome node export-icons.js
 */
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');

const DIR = __dirname;
const OUT = path.join(DIR, 'canva-assets');
const html = fs.readFileSync(path.join(DIR, 'giant-axonal-neuropathy-banner-visual.html'), 'utf8');
const defs = html.match(/<svg width="0" height="0"[\s\S]*?<\/svg>/)[0];

const PALETTE = { teal: '#1A9C8E', gold: '#B7801E', coral: '#E4572E', ink: '#4A5570', ghost: 'rgba(20,33,61,0.09)', tealSoft: '#DDEFEC' };
const cascade = [['gene', 'teal'], ['missing', 'teal'], ['tag', 'gold'], ['pile', 'gold'], ['axon', 'coral'], ['signal', 'coral']];
const therapy = ['capsid', 'cargo', 'inject', 'neuron'];
const bodies = [['teal', 78], ['gold', 50], ['coral', 22], ['ink', 0]];

// Each tile is rendered at 4 css px per mm; icons are 66 mm circles, bodies 66 x 160 mm.
const PX = 4;
let items = '';
for (const [name, col] of cascade) {
  items += `<div class="tile"><div class="ring" style="border-color:${PALETTE[col]};color:${PALETTE[col]}"><svg viewBox="0 0 100 100"><use href="#ic-${name}"/></svg></div></div>`;
}
for (const name of therapy) {
  items += `<div class="tile"><div class="ring soft"><svg viewBox="0 0 100 100"><use href="#ic-${name}"/></svg></div></div>`;
}
for (const [col, pct] of bodies) {
  items += `<div class="tile body"><svg viewBox="0 0 100 236"><use href="#body" fill="${PALETTE.ghost}"/><use href="#body" fill="${PALETTE[col]}" style="clip-path: inset(${pct}% 0 0 0)"/></svg></div>`;
}
const page_html = `<!doctype html><html><head><meta charset="utf-8"><style>
  body { margin: 0; background: transparent; display: flex; flex-wrap: wrap; gap: ${PX * 10}px; padding: ${PX * 10}px; width: ${PX * 1000}px; }
  .tile { width: ${66 * PX}px; height: ${66 * PX}px; }
  .tile.body { width: ${66 * PX}px; height: ${160 * PX}px; }
  .ring { width: 100%; height: 100%; border-radius: 50%; background: #fff; border: ${2.6 * PX}px solid; display: grid; place-items: center; }
  .ring.soft { background: ${PALETTE.tealSoft}; border-color: ${PALETTE.tealSoft}; color: ${PALETTE.teal}; }
  .ring svg { width: ${38 * PX}px; height: ${38 * PX}px; }
  .tile.body svg { width: 100%; height: 100%; display: block; }
</style></head><body>${defs}${items}</body></html>`;

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROME_PATH || '/usr/local/bin/google-chrome',
    headless: true, args: ['--no-sandbox', '--disable-gpu'],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: PX * 1000, height: PX * 600, deviceScaleFactor: 3 });
  await page.setContent(page_html, { waitUntil: 'load' });
  const tiles = await page.$$('.tile');
  const names = [...cascade.map(([n]) => `cascade-${n}`), ...therapy.map(n => `therapy-${n}`), ...bodies.map(([c]) => `body-${c}`)];
  for (let i = 0; i < tiles.length; i++) {
    await tiles[i].screenshot({ path: path.join(OUT, `${names[i]}.png`), omitBackground: true });
  }
  await browser.close();
  console.log(`exported ${tiles.length} PNGs to ${OUT}`);
})();
