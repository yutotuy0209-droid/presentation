// Slide IR v2 (EMU) -> editable PPTX.
// Supports common text boxes and preserves geometry. Unsupported objects are
// reported instead of silently pretending the round trip is lossless.
const fs=require('fs'), path=require('path'), pptxgen=require('pptxgenjs');
const EMU_PER_INCH=914400;
const inch=v=>Number(v||0)/EMU_PER_INCH;

async function irToPptx(deck, outputFile){
  const pptx=new pptxgen();
  const sw=inch(deck.slideSize?.width||12192000), sh=inch(deck.slideSize?.height||6858000);
  pptx.defineLayout({name:'IR_LAYOUT',width:sw,height:sh}); pptx.layout='IR_LAYOUT';
  const unsupported=[];
  for(const s of deck.slides||[]){
    const slide=pptx.addSlide();
    const elements=[...(s.elements||[])].sort((a,b)=>(a.z||0)-(b.z||0));
    for(const el of elements){
      const box={x:inch(el.x),y:inch(el.y),w:inch(el.width),h:inch(el.height),rotate:Number(el.rotation||0)};
      if(el.type==='sp' && (el.text||el.runs?.length)){
        const first=el.runs?.find(r=>r.text)||{};
        slide.addText(el.text||'',{
          ...box,margin:0,
          fontFace:first.fontFamily||'Arial',
          fontSize:first.fontSizePt||18,
          bold:!!first.bold,italic:!!first.italic,
          breakLine:false,
        });
      } else {
        unsupported.push({slide:s.id,element:el.id,name:el.name,type:el.type,reason:'not-yet-lossless'});
      }
    }
  }
  fs.mkdirSync(path.dirname(path.resolve(outputFile)),{recursive:true});
  await pptx.writeFile({fileName:path.resolve(outputFile)});
  return {output:path.resolve(outputFile),unsupported};
}
async function main(){
 const [input,output,report]=process.argv.slice(2);
 if(!input||!output) throw new Error('Usage: node src/ir_to_pptx.js deck.json output.pptx [report.json]');
 const result=await irToPptx(JSON.parse(fs.readFileSync(input,'utf8')),output);
 if(report) fs.writeFileSync(report,JSON.stringify(result,null,2));
 console.log(JSON.stringify(result,null,2));
}
if(require.main===module) main().catch(e=>{console.error(e);process.exitCode=1});
module.exports={irToPptx,EMU_PER_INCH};
