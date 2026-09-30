// Presentation lint checks over common IR.
// Works with native EMU geometry and pixel_layout_engine.py output.
const fs = require('fs');

function intersects(a, b, pad = 0) {
  return (
    a.x < b.x + b.width + pad &&
    a.x + a.width + pad > b.x &&
    a.y < b.y + b.height + pad &&
    a.y + a.height + pad > b.y
  );
}

function isForeground(el) {
  if (!el) return false;
  if (el.style && el.style.allowOverlap) return false;
  return el.type === 'sp' || el.type === 'pic' || !!el.table;
}

function lintDeck(deck, options = {}) {
  const minFont = options.minFontSize || 14;
  const warnFont = options.warnFontSize || 18;
  const maxTextChars = options.maxTextChars || 900;
  const overlapPad = Number(options.overlapPad || 0);
  const issues = [];

  if (deck.slideSize && deck.slideSize.width && deck.slideSize.height) {
    const ratio = deck.slideSize.width / deck.slideSize.height;
    const is16x9 = Math.abs(ratio - 16 / 9) < 0.05;
    const is4x3 = Math.abs(ratio - 4 / 3) < 0.05;
    if (!is16x9 && !is4x3) {
      issues.push({
        rule: 'aspect-ratio',
        severity: 'warning',
        message: `Unusual aspect ratio: ${ratio.toFixed(3)}. Expected 16:9 or 4:3.`,
      });
    }
  }

  for (const slide of deck.slides || []) {
    const slideW = slide.width || deck.slideSize?.width || 12192000;
    const slideH = slide.height || deck.slideSize?.height || 6858000;
    const els = slide.elements || [];

    for (const el of els) {
      if (
        el.x < 0 ||
        el.y < 0 ||
        el.x + el.width > slideW ||
        el.y + el.height > slideH
      ) {
        issues.push({
          slide: slide.id,
          element: el.id,
          rule: 'overflow',
          severity: 'error',
        });
      }

      const checkRuns = (runs, elId) => {
        for (const r of runs || []) {
          const size = Number(r.fontSizePt || r.fontSize);
          if (!size) continue;
          if (size < minFont) {
            issues.push({
              slide: slide.id,
              element: elId,
              rule: 'tiny-text',
              severity: 'error',
              value: size,
              message: `Text below ${minFont}pt: ${size}pt`,
            });
          } else if (size < warnFont) {
            issues.push({
              slide: slide.id,
              element: elId,
              rule: 'small-text',
              severity: 'warning',
              value: size,
              message: `Text below ${warnFont}pt guideline: ${size}pt`,
            });
          }
        }
      };

      const isFooterLike =
        String(el.id || '').endsWith('-footer') ||
        String(el.id || '').endsWith('-page');

      if (!isFooterLike) {
        if (el.runs && el.runs.length) {
          checkRuns(el.runs, el.id);
        } else if (el.style?.fontSizePt || el.style?.fontSize) {
          checkRuns(
            [{ fontSizePt: el.style.fontSizePt || el.style.fontSize }],
            el.id
          );
        }
      }

      if (el.table && el.table.rows) {
        for (const row of el.table.rows) {
          for (const cell of row.cells || []) {
            checkRuns(cell.runs, el.id);
          }
        }
      }
    }

    const chars = els.reduce((n, e) => n + (e.text || '').length, 0);
    if (chars > maxTextChars) {
      issues.push({
        slide: slide.id,
        rule: 'text-density',
        severity: 'warning',
        value: chars,
      });
    }

    const fg = els.filter(isForeground);
    for (let i = 0; i < fg.length; i++) {
      for (let j = i + 1; j < fg.length; j++) {
        const a = fg[i];
        const b = fg[j];

        if (
          String(a.id || '').endsWith('-footer') ||
          String(a.id || '').endsWith('-page') ||
          String(b.id || '').endsWith('-footer') ||
          String(b.id || '').endsWith('-page')
        ) {
          continue;
        }

        if (intersects(a, b, overlapPad)) {
          issues.push({
            slide: slide.id,
            elements: [a.id, b.id],
            rule: 'overlap',
            severity: 'error',
          });
        }
      }
    }
  }

  return issues;
}

if (require.main === module) {
  const file = process.argv[2];
  if (!file) throw new Error('Usage: node src/lint_ir.js deck.json');
  const deck = JSON.parse(fs.readFileSync(file, 'utf8'));
  const issues = lintDeck(deck);
  console.log(JSON.stringify(issues, null, 2));
  if (issues.some(x => x.severity === 'error')) process.exitCode = 1;
}

module.exports = { lintDeck, intersects, isForeground };
