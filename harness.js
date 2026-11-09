const fs=require('fs');
const stub=`
const mk=()=>{const f=function(){};const p=new Proxy(f,{get:(t,k)=>{if(k===Symbol.toPrimitive)return ()=>"0";if(k==="dataset")return {};if(k==="classList")return {toggle(){}};if(k==="length")return 0;if(k==="value")return "0";return p},set:()=>true,apply:()=>p});return p};
global.document={querySelector:()=>mk(),querySelectorAll:()=>[],addEventListener(){},activeElement:null};
global.window={addEventListener(){}};global.location={hash:""};global.history={replaceState(){}};
global.localStorage={getItem(){return null},setItem(){},removeItem(){}};global.navigator={};
`;
let js=fs.readFileSync(require('path').join(__dirname,'..','index.html'),'utf8').match(/<script>([\s\S]*)<\/script>/)[1];
if(!js.includes('renderBrand(); load();')) throw new Error('hook missing');
js=js.replace('renderBrand(); load();','global.__t={getS:()=>S,freshState:freshState,calc:calc,compCredits:compCredits,recFor:recFor,bandOf:bandOf,LINES:LINES,GROUPS:GROUPS};renderBrand(); load();');
eval(stub+js);
const T=global.__t, scs=JSON.parse(fs.readFileSync('scenarios.json','utf8'));
const out={lines:T.LINES.map(L=>({id:L.id,b:L.b,r:L.r,free:L.free||0,per:L.per||'mo',noramp:!!L.noramp,kind:L.kind||''})),res:[]};
for(const sc of scs){
  T.freshState(); const S=T.getS();
  S.emp=sc.emp;S.base=sc.base;S.sku={};sc.skus.forEach(k=>S.sku[k]=true);S.preset=sc.preset;
  S.compOn=!!sc.compOn;S.compM=sc.compM;S.buf=sc.buf;S.price=sc.price;S.cur=sc.cur;S.fx=sc.fx;S.apiTotal=sc.api;S.grace=sc.grace;
  T.GROUPS.forEach(g=>{S.groups[g]={on:true,level:'custom'}});
  T.LINES.forEach(L=>{const st=S.lines[L.id];st.on=true;st.vol=sc.vols[L.id]||0;st.per=(L.id==='rc')?'yr':(L.per||'mo');st.runs=sc.runs;});
  const r=T.calc(sc.m), comp=T.compCredits(), net=Math.max(0,r.total-comp), rec=T.recFor(r.total);
  const ap=r.apiInfo;
  out.res.push({lines:r.lines,linesTotal:r.total-r.api,ent:ap.ent,sum:ap.sum,over:ap.over,api:r.api,total:r.total,compAnnual:S.base==='planning'?10000:T.bandOf(S.emp).c,comp:comp,net:net,rec:rec,cost:rec*S.price*(S.cur==='AUD'?S.fx:1)});
}
fs.writeFileSync('results.json',JSON.stringify(out));
console.log('HARNESS_OK',out.res.map(x=>Math.round(x.total)).join(','));
