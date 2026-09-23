// HTML -> common slide IR using Puppeteer.
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const { findChrome } = require('./html_to_pptx');
const { createDeck, createSlide, normalizeElement } = require('./slide_ir');

async function htmlFilesToIr(htmlFiles, options = {}) {
  const width = options.width || 1920, height = options.height || 1080;
  const browser = await puppeteer.launch({
    executablePath: options.chromePath || findChrome(),
    headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
  const deck = createDeck({ source: 'html' });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width, height });
    for (let i = 0; i < htmlFiles.length; i++) {
      const file = path.resolve(htmlFiles[i]);
      await page.goto('file:///' + file.replace(/\\/g, '/'), { waitUntil: 'networkidle0' });
      const elements = await page.evaluate(() => {
        const candidates = [...document.body.querySelectorAll('h1,h2,h3,p,li,img,svg,table,[data-slide-element]')];
        return candidates.map((el, n) => {
          const r = el.getBoundingClientRect(), cs = getComputedStyle(el);
          return {
            id: el.id || `el-${n + 1}`,
            type: el.tagName.toLowerCase(),
            x: r.x, y: r.y, width: r.width, height: r.height,
            text: (el.innerText || el.getAttribute('alt') || '').trim(),
            style: {
              fontSize: parseFloat(cs.fontSize) || null,
              fontFamily: cs.fontFamily,
              fontWeight: cs.fontWeight,
              textAlign: cs.textAlign,
              color: cs.color,
              backgroundColor: cs.backgroundColor,
            }
          };
        }).filter(x => x.width > 0 && x.height > 0);
      });
      deck.slides.push(createSlide({
        id: `slide-${i + 1}`, width, height,
        elements: elements.map(normalizeElement),
      }));
    }
    return deck;
  } finally { await browser.close(); }
}

async function main() {
  const [out, ...files] = process.argv.slice(2);
  if (!out || !files.length) throw new Error('Usage: node src/html_to_ir.js output.json slide1.html [...]');
  const deck = await htmlFilesToIr(files);
  fs.mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
  fs.writeFileSync(out, JSON.stringify(deck, null, 2));
}
if (require.main === module) main().catch(e => { console.error(e); process.exitCode = 1; });
module.exports = { htmlFilesToIr };
