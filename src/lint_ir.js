// Presentation lint checks over common IR.
// Enforces AGENTS.md rules: 20pt minimum font, 10-11pt forbidden, 16:9 ratio, no-overflow.
const fs = require('fs');

function intersects(a,b) {
  return a.x < b.x+b.width && a.x+a.width > b.x && a.y < b.y+b.height && a.y+a.height > b.y;
}

function lintDeck(deck, options = {}) {
  const minFont = options.minFontSize || 20;
  const maxTextChars = options.maxTextChars || 900;
  const issues = [];

  // Aspect ratio check (16:9 preferred)
  if (deck.slideSize && deck.slideSize.width && deck.slideSize.height) {
    const ratio = deck.slideSize.width / deck.slideSize.height;
    const is16x9 = Math.abs(ratio - (16 / 9)) < 0.05;
    const is4x3 = Math.abs(ratio - (4 / 3)) < 0.05;
    if (!is16x9 && !is4x3) {
      issues.push({ rule: 'aspect-ratio', severity: 'warning', message: `Unusual aspect ratio: ${ratio.toFixed(3)}. Preferred is 16:9.` });
    }
  }

  for (const slide of deck.slides || []) {
    const slideW = slide.width || deck.slideSize?.width || 12192000;
    const slideH = slide.height || deck.slideSize?.height || 6858000;
    const els = slide.elements || [];
    for (const el of els) {
      if (el.x < 0 || el.y < 0 || el.x+el.width > slideW || el.y+el.height > slideH)
        issues.push({slide:slide.id, element:el.id, rule:'overflow', severity:'error'});

      const checkRuns = (runs, elId) => {
        for (const r of runs || []) {
          const sz = r.fontSizePt || Number(r.fontSize);
          if (sz) {
            if (sz <= 11) {
              issues.push({ slide: slide.id, element: elId, rule: 'forbidden-small-text', severity: 'error', value: sz, message: `10-11pt text is strictly forbidden: ${sz}pt found` });
            } else if (sz < minFont) {
              issues.push({ slide: slide.id, element: elId, rule: 'sub-20pt-text', severity: 'warning', value: sz, message: `Text below ${minFont}pt guideline: ${sz}pt found` });
            }
          }
        }
      };

      if (el.runs && el.runs.length) {
        checkRuns(el.runs, el.id);
      } else if (el.style?.fontSize) {
        const fsiz = Number(el.style.fontSize);
        if (fsiz <= 11) {
          issues.push({ slide: slide.id, element: el.id, rule: 'forbidden-small-text', severity: 'error', value: fsiz });
        } else if (fsiz < minFont) {
          issues.push({ slide: slide.id, element: el.id, rule: 'sub-20pt-text', severity: 'warning', value: fsiz });
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
    const chars = els.reduce((n,e)=>n+(e.text||'').length,0);
    if (chars > maxTextChars) issues.push({slide:slide.id, rule:'text-density', severity:'warning', value:chars});
    for (let i=0;i<els.length;i++) for (let j=i+1;j<els.length;j++) {
      if (intersects(els[i],els[j])) {
        const a=els[i], b=els[j];
        issues.push({slide:slide.id, elements:[a.id,b.id], rule:'overlap', severity:'info'});
      }
    }
  }
  return issues;
}

if (require.main === module) {
  const file=process.argv[2]; if(!file) throw new Error('Usage: node src/lint_ir.js deck.json');
  const deck=JSON.parse(fs.readFileSync(file,'utf8'));
  console.log(JSON.stringify(lintDeck(deck),null,2));
}
module.exports={lintDeck,intersects};
