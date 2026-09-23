// Append presentation corrections as JSONL experience without bloating presets automatically.
const fs=require('fs'), path=require('path');
function appendExperience(file, record){
  fs.mkdirSync(path.dirname(path.resolve(file)),{recursive:true});
  fs.appendFileSync(file, JSON.stringify({timestamp:new Date().toISOString(),...record})+'\n');
}
if(require.main===module){
 const [file,diffFile]=process.argv.slice(2); if(!file||!diffFile) throw new Error('Usage: node src/record_experience.js history.jsonl diff.json');
 appendExperience(file,{kind:'human-correction',changes:JSON.parse(fs.readFileSync(diffFile,'utf8'))});
}
module.exports={appendExperience};
