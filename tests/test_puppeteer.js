const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const { findChrome } = require('../src/html_to_pptx');

async function test() {
  const browser = await puppeteer.launch({
    executablePath: findChrome(),
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080, deviceScaleFactor: 2 });

    const htmlPath = path.join(__dirname, 'fixtures', 'test_slide.html');
    const outputDir = path.join(__dirname, '..', 'output');
    const outputPath = path.join(outputDir, 'test_slide.png');

    fs.mkdirSync(outputDir, { recursive: true });
    await page.goto('file:///' + htmlPath.replace(/\\/g, '/'), { waitUntil: 'networkidle0' });
    await page.screenshot({ path: outputPath });

    console.log('Rendered:', outputPath);
  } finally {
    await browser.close();
  }
}

test().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
