// Anonymised replay fixture for the shift-roster parser.
//
// The parser reads nothing from a PDF but numPages and, per page, a list of
// {str, x, y}. The fixture is that list with the names swapped, so no glyphs,
// fonts, document metadata or incremental-save history come along.
//
// The one rule: a substitute must be indistinguishable TO THE PARSER. The
// parser decides on the whole text item (after stripping [ ] (T) (Psy)), and
// its predicates key off shape and length — isNameTok caps at 15 characters,
// COMPOUND_RE at 21, NOISE_RE matches on a prefix. So substitutes preserve
// length, capitalisation and hyphen positions exactly, and every item is then
// VERIFIED to classify identically before the fixture is written.
const { chromium } = require('playwright');
const fs=require('fs'), path=require('path');

const KEEP=[/^\d{1,2}[h:;]\d{2}$/i,/^\d{1,2}[h:;]\d{2}\s*[-–]\s*\d{1,2}[h:;]\d{2}$/i,
 /^[\d.,()\[\]\/&+*%:;#-]+$/,/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Tues|Wed|Thur|Thurs|Fri|Sat|Sun)$/i,
 /^(January|February|March|April|May|June|July|August|September|October|November|December|Sept)$/i,
 /^(WEEK|CONSULTANT|REGISTRAR|INTERN|COSMO|MO|EC|PH|Psych|Psy|O\/T|T|S&T|SHIFTS?|HOURS?|TOTAL|TOT|NUMBER|No|Nights?|Leave|LEAVE|worked|WORKED|per|Shift|Hours|W\/E|and|of|the|x|Date|Day|Off|Good|Family|Freedom|Worker|Heritage|Youth|Reconciliation|Christmas|Goodwill|Week|PERSAL|MDHS)$/i];
const PREFIX=new Set(['Van','De','Du','Von','Le']);
const C='bcdfgklmnprstvw', V='aeiou';
// Same length, same capitalisation, same hyphens — only the letters change.
function shaped(word, seed){
  let out=''; let k=seed;
  for(let i=0;i<word.length;i++){
    const ch=word[i];
    if(ch==='-'||ch==="'"){ out+=ch; continue; }
    k=(k*1103515245+12345)>>>0;
    let c = (i%2===0) ? C[k%C.length] : V[k%V.length];
    if(i===0) c = ch===ch.toUpperCase() ? c.toUpperCase() : c;
    else if(ch===ch.toUpperCase()&&ch!==ch.toLowerCase()) c=c.toUpperCase();
    out+=c;
  }
  return out;
}
// A surname can be noise by accident: "Dayar" matches the NOISE_RE branch
// "Day", so the parser never counts it. The substitute has to inherit that or
// the fixture gains a doctor the real file does not have.
const NOISE_STEMS=['Christmas','Reconciliation','Freedom','Heritage','Family','Goodwill','Worker',
 'Nights','Shifts','Hours','Youth','Week','Good','TOT','Day','No'];
function noiseShaped(word, seed){
  const stem=NOISE_STEMS.find(st=>st.length<=word.length) || 'No';
  const tail=shaped('x'.repeat(Math.max(0,word.length-stem.length)), seed);
  return stem+tail;
}
(async()=>{
 const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome'});
 const q=await b.newContext().then(c=>c.newPage());
 await q.route('**/dspemkdjifgpaswjwoij.supabase.co/**',r=>r.abort());
 await q.goto('file:///home/user/VHWRosterTranslator/index.html');
 await q.waitForTimeout(900);
 // The parser's verdict on an item, as the parser itself computes it.
 const verdict = items => q.evaluate(list=>list.map(s=>{
   const raw=s.replace(/\(T\)/gi,'').replace(/\(Psy\)/gi,'').replace(/\[|\]/g,'').trim();
   const COMPOUND_RE=/^(Van|De|Du|Von|Le)\s+([A-Z][a-zA-Z\-]{1,20})$/;
   return [raw.length, isNoise(raw), isNameTok(raw), !!raw.match(COMPOUND_RE), NAME_PREFIXES.has(raw)].join('|');
 }), items);

 for (const f of process.argv.slice(2)) {
   const buf=fs.readFileSync(path.join(__dirname,f));
   const raw=await q.evaluate(async bytes=>{
     const pdf=await pdfjsLib.getDocument({data:new Uint8Array(bytes)}).promise; const pages=[];
     for(let p=1;p<=pdf.numPages;p++){const page=await pdf.getPage(p);
       pages.push((await page.getTextContent()).items.map(i=>[i.str,Math.round(i.transform[4]),Math.round(i.transform[5])]));}
     return pages;
   }, Array.from(buf));

   const split=s=>s.split(/(\s+)/), parts=w=>w.match(/^([^A-Za-z]*)([A-Za-z][A-Za-z'\-]*)(.*)$/);
   const words=new Set();
   for(const pg of raw) for(const [s] of pg) for(const w of split(s)){
     if(/^\s*$/.test(w)||KEEP.some(re=>re.test(w))||PREFIX.has(w)) continue;
     const m=parts(w); if(!m||KEEP.some(re=>re.test(m[2]))) continue;
     words.add(m[2]);
   }
   const list=[...words];
   const isNoisy=await q.evaluate(ws=>ws.map(w=>isNoise(w)), list);
   const map=new Map(); const used=new Set();
   list.forEach((w,i)=>{ const gen=isNoisy[i]?noiseShaped:shaped;
     let s=1,f; do{ f=gen(w,i*7919+s*104729); s++; }while(used.has(f)&&s<50); used.add(f); map.set(w,f); });
   const sub = s => split(s).map(w=>{ if(/^\s*$/.test(w)) return w;
     const m=parts(w); return (!m||!map.has(m[2])) ? w : m[1]+map.get(m[2])+m[3]; }).join('');

   const pages=raw.map(items=>items.map(([s,x,y])=>[sub(s),x,y]));
   if(process.env.DUMP_MAP) fs.writeFileSync(path.join(__dirname,f+'.map.debug.json'), JSON.stringify([...map]));
   // verify: every item must earn the same verdict as the original
   const flatA=raw.flat().map(i=>i[0]), flatB=pages.flat().map(i=>i[0]);
   const [vA,vB]=[await verdict(flatA), await verdict(flatB)];
   const bad=vA.map((v,i)=>v===vB[i]?null:i).filter(i=>i!==null);
   // Items classify one at a time; the date grammar lives on the ROW. A
   // substitute can keep every per-item verdict and still stop a row looking
   // like a date — which is precisely what "Wednesday" did.
   const rowVerdict = pgs => q.evaluate(ps=>ps.map(items=>{
     const m=new Map();
     for(const [t,x,y] of items){ if(!t.trim()) continue; let hit=false;
       for(const [k] of m) if(Math.abs(k-y)<=6){ m.get(k).push([t,x]); hit=true; break; }
       if(!hit) m.set(y,[[t,x]]); }
     return [...m.keys()].sort((a,b)=>b-a).map(y=>{
       const line=m.get(y).sort((a,b)=>a[1]-b[1]).map(w=>w[0].trim()).join(' ').trim();
       return (DATE_RE.test(line)?'D':'-')+(PARTIAL_DATE_RE.test(line)?'P':'-')+(HAS_DATE.test(line)?'H':'-');
     }).join('');
   }), pgs);
   const [rA,rB]=[await rowVerdict(raw), await rowVerdict(pages)];
   const badRows=rA.map((v,i)=>v===rB[i]?null:i+1).filter(i=>i!==null);
   const leak=[...words].filter(w=>flatB.join(' ').includes(w));
   const out=f.replace(/\.pdf$/,'.fixture.json');
   fs.writeFileSync(path.join(__dirname,out), JSON.stringify({note:'anonymised replay fixture', pages}));
   console.log(`${f} → ${out}  ${pages.length}p ${flatA.length} items, ${words.size} names swapped`);
   console.log(`    items classifying differently: ${bad.length}${bad.length?'  e.g. '+bad.slice(0,3).map(i=>`"${flatA[i]}"→"${flatB[i]}"`).join(', '):''}`);
   console.log(`    pages whose row grammar moved: ${badRows.length?badRows.join(','):'none'}`);
   console.log(`    original words still present : ${leak.length?leak.join(', '):'none'}`);
 }
 await b.close();
})();
