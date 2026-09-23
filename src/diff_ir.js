// Compare generated and human-edited/common IR decks.
const fs=require('fs');
function diffDecks(before, after, tolerance=1) {
  const changes=[];
  const bSlides=new Map((before.slides||[]).map(s=>[s.id,s]));
  for (const aSlide of after.slides||[]) {
    const bSlide=bSlides.get(aSlide.id); if(!bSlide){changes.push({type:'slide-added',slide:aSlide.id});continue;}
    const bEls=new Map((bSlide.elements||[]).map(e=>[e.id,e]));
    for(const a of aSlide.elements||[]){
      const b=bEls.get(a.id); if(!b){changes.push({type:'element-added',slide:aSlide.id,element:a.id});continue;}
      for(const key of ['x','y','width','height']){
        if(Math.abs(Number(a[key]||0)-Number(b[key]||0))>tolerance)
          changes.push({type:'geometry',slide:aSlide.id,element:a.id,field:key,before:b[key],after:a[key]});
      }
      if((a.text||'')!==(b.text||'')) changes.push({type:'text',slide:aSlide.id,element:a.id,before:b.text||'',after:a.text||''});
      if(JSON.stringify(a.style||{})!==JSON.stringify(b.style||{}))
        changes.push({type:'style',slide:aSlide.id,element:a.id,before:b.style||{},after:a.style||{}});
      bEls.delete(a.id);
    }
    for(const id of bEls.keys()) changes.push({type:'element-removed',slide:aSlide.id,element:id});
    bSlides.delete(aSlide.id);
  }
  for(const id of bSlides.keys()) changes.push({type:'slide-removed',slide:id});
  return changes;
}
if(require.main===module){
 const [b,a]=process.argv.slice(2); if(!b||!a) throw new Error('Usage: node src/diff_ir.js before.json after.json');
 console.log(JSON.stringify(diffDecks(JSON.parse(fs.readFileSync(b)),JSON.parse(fs.readFileSync(a))),null,2));
}
module.exports={diffDecks};
