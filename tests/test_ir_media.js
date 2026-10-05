const fs = require('fs');
const os = require('os');
const path = require('path');
const assert = require('assert');
const { irToPptx, EMU_PER_INCH } = require('../src/ir_to_pptx');

async function main() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'presentation-media-'));
  const gif = path.join(dir, 'tiny.gif');
  // 1x1 transparent GIF89a
  fs.writeFileSync(
    gif,
    Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==', 'base64')
  );

  const out = path.join(dir, 'media.pptx');
  const deck = {
    version: 2,
    slideSize: { width: Math.round(13.333 * EMU_PER_INCH), height: Math.round(7.5 * EMU_PER_INCH) },
    slides: [{
      id: 's1',
      background: 'FFFFFF',
      elements: [{
        id: 'gif1',
        type: 'pic',
        x: 1 * EMU_PER_INCH,
        y: 1 * EMU_PER_INCH,
        width: 3 * EMU_PER_INCH,
        height: 2 * EMU_PER_INCH,
        source: { path: gif, mediaType: 'gif' },
        animation: { entrance: 'fade', order: 1, trigger: 'click' }
      }]
    }]
  };

  const result = await irToPptx(deck, out);
  assert.ok(fs.existsSync(out), 'pptx should be generated');
  assert.ok(fs.statSync(out).size > 0, 'pptx should not be empty');
  assert.deepStrictEqual(result.unsupported, []);
  assert.strictEqual(result.animationReservations.length, 1);
  assert.strictEqual(result.animationReservations[0].applied, false);
  assert.strictEqual(result.shapeMap.s1.gif1, 1);
  console.log('media writer ok:', out);
}

main().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
