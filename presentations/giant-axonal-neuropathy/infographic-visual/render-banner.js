#!/usr/bin/env node
/**
 * Render the roll-up banner HTML to a print-ready PDF (850 × 2000 mm, vector) and a PNG preview.
 *
 * Usage:  npm i puppeteer-core && CHROME_PATH=/path/to/chrome node render-banner.js
 * (Defaults to the Google Chrome path used in the build environment.)
 */
const path = require('path');
const puppeteer = require('puppeteer-core');

const DIR = __dirname;
const FILE = `file://${path.join(DIR, 'giant-axonal-neuropathy-banner-visual.html')}`;
const PX_PER_MM = 96 / 25.4;
const WIDTH_MM = 850;
const HEIGHT_MM = 2000;

(async () => {
  const browser = await puppeteer.launch({
    executablePath: process.env.CHROME_PATH || '/usr/local/bin/google-chrome',
    headless: true,
    args: ['--no-sandbox', '--disable-gpu', '--allow-file-access-from-files'],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: Math.round(WIDTH_MM * PX_PER_MM), height: Math.round(HEIGHT_MM * PX_PER_MM) });
  await page.goto(FILE, { waitUntil: 'networkidle0' });
  await page.evaluate(() => document.fonts.ready);

  // Guard: the banner must not exceed its trim height.
  const overflowMm = await page.evaluate((pxPerMm) => {
    const b = document.getElementById('banner');
    return (b.scrollHeight - b.clientHeight) / pxPerMm;
  }, PX_PER_MM);
  if (overflowMm > 0.5) {
    console.error(`Content overflows the 2000 mm trim by ${overflowMm.toFixed(1)} mm — shorten a section before exporting.`);
    await browser.close();
    process.exit(1);
  }

  await page.pdf({
    path: path.join(DIR, 'giant-axonal-neuropathy-banner-visual.pdf'),
    width: `${WIDTH_MM}mm`, height: `${HEIGHT_MM}mm`,
    printBackground: true, preferCSSPageSize: true,
    margin: { top: 0, right: 0, bottom: 0, left: 0 },
  });

  await page.setViewport({ width: Math.round(WIDTH_MM * PX_PER_MM), height: Math.round(HEIGHT_MM * PX_PER_MM), deviceScaleFactor: 0.31 });
  await page.screenshot({ path: path.join(DIR, 'giant-axonal-neuropathy-banner-visual-preview.png') });

  await browser.close();
  console.log('Exported giant-axonal-neuropathy-banner-visual.pdf and giant-axonal-neuropathy-banner-visual-preview.png');
})();
