// Presentation lint checks over common IR.
const fs = require('fs');

function intersects(a,b) {
  return a.x < b.x+b.width && a.x+a.width > b.x && a.y < b.y+b.height && a.y+a.height > b.y;
}
function lintDeck(deck, options = {}) {
  const minFont = options.minFontSize || 18;
  const maxTextChars = options.maxTextChars || 900;
  const issues = [];
  for (const slide of deck.slides || []) {
    const els = slide.elements || [];
    for (const el of els) {
      if (el.x < 0 || el.y < 0 || el.x+el.width > slide.width || el.y+el.height > slide.height)
        issues.push({slide:slide.id, element:el.id, rule:'overflow', severity:'error'});
      const fsiz = Number(el.style?.fontSize);
      if (fsiz && fsiz < minFont)
        issues.push({slide:slide.id, element:el.id, rule:'tiny-text', severity:'warning', value:fsiz});
    }
    const chars = els.reduce((n,e)=>n+(e.text||'').length,0);
    if (chars > maxTextChars) issues.push({slide:slide.id, rule:'text-density', severity:'warning', value:chars});
    for (let i=0;i<els.length;i++) for (let j=i+1;j<els.length;j++) {
      if (intersects(els[i],els[j])) {
        const a=els[i], b=els[j];
        // nested HTML elements often overlap intentionally; source/type-specific suppression can be added later.
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
