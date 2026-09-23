// Generic HTML -> PNG -> PPTX renderer.
// Extracted and generalized from presentation-generation code used in
// graduation-research-motion-analysis. No research-specific content is included.

const fs = require('fs');
const path = require('path');
const os = require('os');
const puppeteer = require('puppeteer-core');
const pptxgen = require('pptxgenjs');

function findChrome() {
  const candidates = [
    process.env.CHROME_PATH,
    process.platform === 'win32' ? 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' : null,
    process.platform === 'win32' ? 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe' : null,
    process.platform === 'darwin' ? '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' : null,
    process.platform === 'linux' ? '/usr/bin/google-chrome' : null,
    process.platform === 'linux' ? '/usr/bin/chromium' : null,
    process.platform === 'linux' ? '/usr/bin/chromium-browser' : null,
  ].filter(Boolean);

  const found = candidates.find(p => fs.existsSync(p));
  if (!found) {
    throw new Error('Chrome/Chromium was not found. Set CHROME_PATH to the executable path.');
  }
  return found;
}

async function renderHtmlFilesToPptx(htmlFiles, outputFile, options = {}) {
  if (!Array.isArray(htmlFiles) || htmlFiles.length === 0) {
    throw new Error('htmlFiles must contain at least one HTML file.');
  }

  const width = options.width || 1920;
  const height = options.height || 1080;
  const deviceScaleFactor = options.deviceScaleFactor || 2;
  const tmpDir = options.tmpDir || fs.mkdtempSync(path.join(os.tmpdir(), 'presentation-render-'));

  const browser = await puppeteer.launch({
    executablePath: options.chromePath || findChrome(),
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  try {
    const page = await browser.newPage();
    await page.setViewport({ width, height, deviceScaleFactor });

    for (let i = 0; i < htmlFiles.length; i += 1) {
      const htmlPath = path.resolve(htmlFiles[i]);
      if (!fs.existsSync(htmlPath)) {
        throw new Error(`HTML file not found: ${htmlPath}`);
      }

      await page.goto('file:///' + htmlPath.replace(/\\/g, '/'), { waitUntil: 'networkidle0' });

      const pngPath = path.join(tmpDir, `slide_${String(i + 1).padStart(2, '0')}.png`);
      await page.screenshot({ path: pngPath, fullPage: false });

      const slide = pptx.addSlide();
      slide.addImage({ path: pngPath, x: 0, y: 0, w: 13.333, h: 7.5 });
    }

    fs.mkdirSync(path.dirname(path.resolve(outputFile)), { recursive: true });
    await pptx.writeFile({ fileName: path.resolve(outputFile) });
    return path.resolve(outputFile);
  } finally {
    await browser.close();
  }
}

async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.log('Usage: node src/html_to_pptx.js <output.pptx> <slide1.html> [slide2.html ...]');
    process.exitCode = 1;
    return;
  }

  const [outputFile, ...htmlFiles] = args;
  const result = await renderHtmlFilesToPptx(htmlFiles, outputFile);
  console.log(`Created: ${result}`);
}

if (require.main === module) {
  main().catch(err => {
    console.error(err);
    process.exitCode = 1;
  });
}

module.exports = { renderHtmlFilesToPptx, findChrome };
