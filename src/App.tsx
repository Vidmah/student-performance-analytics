import {useState,useMemo,useRef} from 'react';
import * as XLSX from 'xlsx';
type R=Record<string,any>;
/* ---------- theme ---------- */
const CSS=`
:root{--bg:#f3f5f9;--card:#ffffff;--fg:#0f172a;--mu:#64748b;--bd:#e2e8f0;--grid:#e8ecf2;--hover:#f1f5f9;--ac:#1e3a8a;--ac2:#0f766e}
html,body{background:var(--bg);color:var(--fg);font-family:Inter,"Segoe UI",system-ui,-apple-system,sans-serif;-webkit-font-smoothing:antialiased;font-variant-numeric:tabular-nums}
.mu{color:var(--mu)}
.glass{background:var(--card);border:1px solid var(--bd);border-radius:10px;box-shadow:0 1px 2px rgba(15,23,42,.05)}
.fade{animation:fade .35s ease both}@keyframes fade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
.ct{font-size:.7rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--mu);margin-bottom:.9rem}
.ph{padding-bottom:1rem;border-bottom:1px solid var(--bd)}
.eyebrow{font-size:.68rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--ac);margin-bottom:.35rem}
.kpi{background:var(--card);border:1px solid var(--bd);border-left:3px solid var(--ac);border-radius:10px;padding:1rem;box-shadow:0 1px 2px rgba(15,23,42,.05)}
.kl{font-size:.68rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--mu)}
.kv{font-size:1.6rem;font-weight:700;margin-top:.4rem;line-height:1.1;letter-spacing:-.01em}
.ks{font-size:.74rem;color:var(--mu);margin-top:.3rem}
.btn{border:1px solid var(--bd);background:var(--card);color:var(--fg);padding:.4rem .8rem;border-radius:6px;font-size:.78rem;font-weight:600;cursor:pointer;transition:background .15s,border-color .15s}
.btn:hover:not(:disabled){background:var(--hover);border-color:var(--ac)}
.btn.on{background:var(--ac);color:#fff;border-color:var(--ac)}
.btn:disabled{opacity:.45;cursor:not-allowed}
.inp{border:1px solid var(--bd);background:var(--card);color:var(--fg);border-radius:6px;padding:.45rem .65rem;font-size:.82rem}
.inp:focus,.btn:focus-visible,.nav:focus-visible,input[type=range]:focus-visible{outline:2px solid var(--ac);outline-offset:1px}
input[type=range]{width:100%;accent-color:var(--ac)}
table{width:100%;border-collapse:collapse;font-size:.8rem}
th{text-align:left;font-size:.66rem;letter-spacing:.07em;text-transform:uppercase;color:var(--mu);font-weight:700;padding:.55rem .7rem;border-bottom:1px solid var(--bd);background:var(--hover);white-space:nowrap}
td{padding:.5rem .7rem;border-bottom:1px solid var(--bd)}
tbody tr:hover td{background:var(--hover)}
.nav{display:flex;align-items:center;justify-content:center;gap:.7rem;padding:.55rem .6rem;border-radius:6px;font-size:.82rem;font-weight:500;color:var(--mu);border-left:3px solid transparent;cursor:pointer;text-align:left;width:100%}
@media(min-width:768px){.nav{justify-content:flex-start}}
.nav:hover{background:var(--hover);color:var(--fg)}
.nav.on{background:var(--hover);color:var(--ac);font-weight:700;border-left-color:var(--ac)}
.grp{font-size:.62rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--mu);padding:.9rem .6rem .3rem}
.brand{width:2.5rem;height:2.5rem;border-radius:8px;display:grid;place-items:center;background:var(--ac);color:#fff;font-weight:800;font-size:.85rem;letter-spacing:.04em}
.dot{transition:r .12s}.dot:hover{r:6}
.bar{transition:opacity .15s}.bar:hover{opacity:.8}
.drop{border:1.5px dashed var(--bd);border-radius:8px;padding:1.5rem;text-align:center;cursor:pointer;display:block;transition:background .15s,border-color .15s}
.drop:hover{background:var(--hover);border-color:var(--ac)}
.note{border-left:3px solid var(--ac2);background:var(--hover);padding:.75rem .95rem;border-radius:6px;font-size:.85rem;line-height:1.5}
.topbar{display:flex;justify-content:space-between;gap:1rem;flex-wrap:wrap;font-size:.74rem;color:var(--mu);padding-bottom:.8rem;margin-bottom:1.2rem;border-bottom:1px solid var(--bd)}
.findings{list-style:none;padding:0;margin:0;display:grid;gap:.6rem}
.findings li{font-size:.85rem;line-height:1.5;padding-left:1rem;position:relative}
.findings li:before{content:"";position:absolute;left:0;top:.55em;width:5px;height:5px;background:var(--ac);border-radius:1px}
`;
/* ---------- utils ---------- */
const sum=(a:number[])=>a.reduce((x,y)=>x+y,0),mean=(a:number[])=>a.length?sum(a)/a.length:0;
const sdv=(a:number[])=>{const m=mean(a);return Math.sqrt(mean(a.map(v=>(v-m)**2)))||1};
const f2=(v:number,d=2)=>Number.isFinite(v)?v.toFixed(d):'—',pc=(v:number)=>(v*100).toFixed(1)+'%';
const lbl=(s:string)=>s.replace(/_/g,' ');
const CL=['Low','Medium','High'],CC:R={Low:'#b91c1c',Medium:'#b45309',High:'#047857'};
const rng=(s:number)=>()=>{s=(s+0x6D2B79F5)|0;let t=Math.imul(s^(s>>>15),1|s);t=(t+Math.imul(t^(t>>>7),61|t))^t;return((t^(t>>>14))>>>0)/4294967296};
const mode=(v:any[])=>{const o:R={};v.forEach(x=>o[x]=(o[x]||0)+1);return Object.keys(o).sort((a,b)=>o[b]-o[a])[0]??''};
const toCSV=(rs:R[])=>{const c=Object.keys(rs[0]||{}),q=(v:any)=>/[",\n]/.test(String(v))?'"'+String(v).replace(/"/g,'""')+'"':String(v);return[c.join(','),...rs.map(r=>c.map(k=>q(r[k])).join(','))].join('\n')};
const download=(rs:R[],name:string)=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([toCSV(rs)],{type:'text/csv'}));a.download=name.replace(/\.[^.]+$/,'').replace(/[^\w-]+/g,'_')+'_cleaned.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};
/* ---------- sample data ---------- */
function sample():R[]{const r=rng(11),g=()=>Math.sqrt(-2*Math.log(r()+1e-9))*Math.cos(6.283*r()),cl=(v:number,a:number,b:number)=>Math.min(b,Math.max(a,v));
return Array.from({length:240},(_,i)=>{const att=cl(78+12*g(),40,100),st=cl(3.6+1.8*g(),.5,9),pv=cl(62+14*g(),25,98),as=cl(.5*pv+.15*att+25+9*g(),20,100),z=att+6*st+10*g(),pa=z>115?'High':z>90?'Medium':'Low';
const fin=cl(.35*pv+.2*att+3.2*st+.15*as+{High:5,Medium:0,Low:-4}[pa]+5*g(),12,99);
return {Student_ID:'S'+(1001+i),Gender:r()<.5?'Male':'Female',Attendance:+att.toFixed(1),Study_Hours:+st.toFixed(1),Previous_Marks:+pv.toFixed(1),Assignment_Score:i%47===5?'':+as.toFixed(1),Participation:pa,Final_Marks:+fin.toFixed(1)}})}
/* ---------- file parsing & preprocessing ---------- */
function parseCSV(t:string):R[]{const rows:string[][]=[];let c='',row:string[]=[],q=false;
for(let i=0;i<t.length;i++){const ch=t[i];if(q){if(ch=='"'&&t[i+1]=='"'){c+='"';i++}else if(ch=='"')q=false;else c+=ch}else if(ch=='"')q=true;else if(ch==','){row.push(c);c=''}else if(ch=='\n'||ch=='\r'){if(ch=='\r'&&t[i+1]=='\n')i++;row.push(c);rows.push(row);row=[];c=''}else c+=ch}
if(c||row.length){row.push(c);rows.push(row)}const [h,...b]=rows.filter(r=>r.some(x=>x.trim()));return b.map(r=>Object.fromEntries(h.map((k,i)=>[k.trim(),(r[i]??'').trim()])))}
const norm=(rs:R[])=>rs.map(r=>Object.fromEntries(Object.entries(r).map(([k,v])=>{const s=String(v??'').trim();return[k,/^(|na|n\/a|nan|null|\?)$/i.test(s)?'':(!isNaN(+s)?+s:s)]})));
function prep(raw:R[],o:{missing:string,dedupe:boolean}){let rs=raw,dup=0;
if(o.dedupe){const seen=new Set();rs=rs.filter(r=>{const k=JSON.stringify(Object.values(r));if(seen.has(k)){dup++;return false}seen.add(k);return true})}
const cols=Object.keys(rs[0]||{}),missing=sum(rs.map(r=>cols.filter(c=>r[c]==='').length));
if(o.missing==='drop')rs=rs.filter(r=>cols.every(c=>r[c]!==''));
else if(o.missing==='fill'){const fill:R={};cols.forEach(c=>{const v=rs.map(r=>r[c]).filter(x=>x!=='');fill[c]=v.length&&v.every(x=>typeof x==='number')?+mean(v).toFixed(1):mode(v)});rs=rs.map(r=>({...r,...Object.fromEntries(cols.filter(c=>r[c]==='').map(c=>[c,fill[c]]))}))}
return{rs,dup,missing}}
function schema(rs:R[]){const cols=Object.keys(rs[0]||{}).filter(c=>!/^(student.?id|id|name|roll.*)$/i.test(c));
const isNum=(c:string)=>rs.filter(r=>r[c]!=='').every(r=>typeof r[c]==='number');const nums=cols.filter(isNum);
const target=nums.find(c=>/final/i.test(c))||nums.filter(c=>/mark|score|grade/i.test(c)).pop()||nums[nums.length-1];
return{num:nums.filter(c=>c!==target),cat:cols.filter(c=>!isNum(c)&&new Set(rs.map(r=>r[c])).size<=12),target}}
/* ---------- algorithms ---------- */
const dc=(rows:any[])=>{const o:R={};rows.forEach(r=>o[r.y]=(o[r.y]||0)+1);return o};
const ent=(o:R)=>{const n=sum(Object.values(o) as number[]);return-sum((Object.values(o) as number[]).map(c=>c/n*Math.log2(c/n)))};
function apriori(T:string[][],minSup:number,minConf:number,onlyY:boolean){const N=T.length,S=T.map(t=>new Set(t)),col=(i:string)=>i.split('=')[0];
const sup=(is:string[])=>S.filter(s=>is.every(i=>s.has(i))).length/N;
let L=[...new Set(T.flat())].sort().map(i=>[i]).filter(i=>sup(i)>=minSup);const all:string[][]=[...L];
for(let k=2;k<=4;k++){const nx:string[][]=[];for(const a of L)for(const b of L){const p=a.slice(0,-1).join(),q=b.slice(0,-1).join();if(p===q&&a[a.length-1]<b[b.length-1]&&col(a[a.length-1])!==col(b[b.length-1])){const c=[...a,b[b.length-1]];if(sup(c)>=minSup)nx.push(c)}}if(!nx.length)break;all.push(...nx);L=nx}
const rules:R[]=[];all.filter(s=>s.length>1).forEach(s=>{const ss=sup(s);s.forEach(c=>{if(onlyY&&col(c)!=='Result')return;const ant=s.filter(x=>x!==c),conf=ss/sup(ant);if(conf>=minConf)rules.push({rule:ant.join(' ∧ ')+'  →  '+c,sup:ss,conf,lift:conf/sup([c]),n:s.length})})});return{rules,n:all.length}}
function j48(rows:any[],attrs:string[],d=0):any{const dist=dc(rows),n=rows.length,maj=Object.keys(dist).sort((a,b)=>dist[b]-dist[a])[0],leaf={maj,dist,n};
if(Object.keys(dist).length<2||d>=4||n<12||!attrs.length)return leaf;const H=ent(dist);let best:any=null;
for(const a of attrs){const g:R={};rows.forEach(r=>(g[r.x[a]]=g[r.x[a]]||[]).push(r));const gs=Object.values(g) as any[][];if(gs.filter(x=>x.length>=5).length<2)continue;
const gain=H-sum(gs.map(x=>x.length/n*ent(dc(x)))),si=-sum(gs.map(x=>x.length/n*Math.log2(x.length/n))),gr=gain/(si||1);if(gain>.02&&(!best||gr>best.gr))best={a,g,gr}}
if(!best)return leaf;return{...leaf,attr:best.a,kids:Object.fromEntries(Object.entries(best.g).map(([k,v]:any)=>[k,j48(v,attrs.filter(x=>x!==best.a),d+1)]))}}
const walk=(n:any,x:R)=>{let c=n;while(c.attr){const k=c.kids[x[c.attr]];if(!k)break;c=k}return c};
function naive(tr:any[],attrs:string[]){const pri:R={},lik:R={},vals:R={};attrs.forEach(a=>vals[a]=[...new Set(tr.map(r=>r.x[a]))]);
CL.forEach(c=>{const rs=tr.filter(r=>r.y===c);pri[c]=(rs.length+1)/(tr.length+3);lik[c]={};attrs.forEach(a=>{lik[c][a]={};vals[a].forEach((v:string)=>lik[c][a][v]=(rs.filter(r=>r.x[a]===v).length+1)/(rs.length+vals[a].length))})});
return{pri,pred:(x:R)=>{const lp=CL.map(c=>Math.log(pri[c])+sum(attrs.map(a=>Math.log(lik[c][a][x[a]]??1e-3)))),m=Math.max(...lp),e=lp.map(v=>Math.exp(v-m)),s=sum(e);return e.map(v=>v/s)}}}
function metrics(a:string[],p:string[]){const cm=CL.map(x=>CL.map(y=>a.filter((v,i)=>v===x&&p[i]===y).length));
const per=CL.map((c,i)=>{const tp=cm[i][i],fp=sum(CL.map((_,j)=>cm[j][i]))-tp,fn=sum(cm[i])-tp,pr=tp/(tp+fp)||0,rc=tp/(tp+fn)||0;return{c,pr,rc,f1:2*pr*rc/(pr+rc)||0,n:sum(cm[i])}});
return{cm,per,acc:sum(CL.map((_,i)=>cm[i][i]))/a.length,pr:mean(per.map(x=>x.pr)),rc:mean(per.map(x=>x.rc)),f1:mean(per.map(x=>x.f1))}}
function solve(A:number[][],b:number[]){const n=b.length,M=A.map((r,i)=>[...r,b[i]]);for(let i=0;i<n;i++){let p=i;for(let r=i+1;r<n;r++)if(Math.abs(M[r][i])>Math.abs(M[p][i]))p=r;[M[i],M[p]]=[M[p],M[i]];for(let r=0;r<n;r++)if(r!==i){const f=M[r][i]/M[i][i];for(let c=i;c<=n;c++)M[r][c]-=f*M[i][c]}}return M.map((r,i)=>r[n]/r[i])}
function kmeans(X:number[][],k:number){const r=rng(3),dist=(a:number[],b:number[])=>sum(a.map((v,i)=>(v-b[i])**2));let C=[X[Math.floor(r()*X.length)]];
while(C.length<k){const d=X.map(x=>Math.min(...C.map(c=>dist(x,c)))),tot=sum(d);let t=r()*tot,i=0;for(;i<d.length-1&&(t-=d[i])>0;i++);C.push(X[i])}
let a=new Array(X.length).fill(0);for(let it=0;it<60;it++){const na=X.map(x=>{const ds=C.map(c=>dist(x,c));return ds.indexOf(Math.min(...ds))}),ch=na.some((v,i)=>v!==a[i]);a=na;
C=C.map((c,j)=>{const m=X.filter((_,i)=>a[i]===j);return m.length?c.map((_,d)=>mean(m.map(x=>x[d]))):c});if(!ch&&it>0)break}
return{a,C,inertia:sum(X.map((x,i)=>dist(x,C[a[i]])))}}
/* ---------- model pipeline ---------- */
function build(rs:R[]):any{const S=schema(rs);if(!S.target||rs.length<20)return null;const{num,cat,target}=S,feats=[...num,...cat];
const data=rs.filter(r=>[...feats,target].every(c=>r[c]!==''&&r[c]!=null));if(data.length<20||!feats.length)return null;
const cuts:R={};num.forEach(f=>{const v=data.map(r=>+r[f]).sort((a,b)=>a-b);cuts[f]=[v[Math.floor(v.length/3)],v[Math.floor(2*v.length/3)]]});
const mx=Math.max(...data.map(r=>+r[target])),top=mx>100?mx:mx>20?100:mx>10?20:10,cls=(v:number)=>v<.5*top?'Low':v<.65*top?'Medium':'High';
const disc=(r:R)=>{const d:R={};num.forEach(f=>d[f]=+r[f]<cuts[f][0]?'Low':+r[f]<cuts[f][1]?'Medium':'High');cat.forEach(f=>d[f]=String(r[f]));return d};
const D=data.map(r=>({x:disc(r),y:cls(+r[target]),r})),rr=rng(42),idx=D.map((_,i)=>i).sort(()=>rr()-.5),cut=Math.floor(D.length*.7),tr=idx.slice(0,cut).map(i=>D[i]),te=idx.slice(cut).map(i=>D[i]);
const tree=j48(tr,feats),jm=metrics(te.map(r=>r.y),te.map(r=>walk(tree,r.x).maj)),nbm=naive(tr,feats),nm=metrics(te.map(r=>r.y),te.map(r=>CL[nbm.pred(r.x).indexOf(Math.max(...nbm.pred(r.x)))]));
const lev:R={};cat.forEach(c=>lev[c]=[...new Set(data.map(r=>r[c]))].sort());
const enc=(r:R)=>[1,...num.map(f=>+r[f]),...cat.flatMap(c=>lev[c].slice(1).map((l:string)=>r[c]===l?1:0))];
const fit=(rows:R[])=>{const X=rows.map(enc),y=rows.map(r=>+r[target]),p=X[0].length;const A=Array.from({length:p},(_,i)=>Array.from({length:p},(_,j)=>sum(X.map(x=>x[i]*x[j]))+(i===j?1e-6:0)));return solve(A,Array.from({length:p},(_,i)=>sum(X.map((x,k)=>x[i]*y[k]))))};
const beta=fit(tr.map(t=>t.r)),pr=(r:R)=>Math.min(top,Math.max(0,sum(enc(r).map((v,i)=>v*beta[i])))),act=te.map(t=>+t.r[target]),pd=te.map(t=>pr(t.r)),mm=mean(act);
const reg={beta,pred:pr,mae:mean(act.map((a,i)=>Math.abs(a-pd[i]))),rmse:Math.sqrt(mean(act.map((a,i)=>(a-pd[i])**2))),r2:1-sum(act.map((a,i)=>(a-pd[i])**2))/sum(act.map(a=>(a-mm)**2)),act,pd,
 impact:num.map((f,i)=>({f,v:beta[i+1]*sdv(data.map(r=>+r[f]))}))};
const mu=num.map(f=>mean(data.map(r=>+r[f]))),sd=num.map(f=>sdv(data.map(r=>+r[f]))),Z=(r:R)=>num.map((f,i)=>(+r[f]-mu[i])/sd[i]);
const med:R={};num.forEach(f=>{const v=data.map(r=>+r[f]).sort((a,b)=>a-b);med[f]=v[Math.floor(v.length/2)]});
const rngs:R={};num.forEach(f=>{const v=data.map(r=>+r[f]);rngs[f]=[Math.min(...v),Math.max(...v)]});
return{S,data,top,cls,disc,tree,jm,nb:nbm,nm,nbPred:nbm.pred,reg,Z,mu,sd,med,rngs,lev,feats,T:D.map(d=>[...feats.map(f=>f+'='+d.x[f]),'Result='+d.y]),
 corr:num.map(f=>{const a=data.map(r=>+r[f]),b=data.map(r=>+r[target]),ma=mean(a),mb=mean(b);return{f,v:sum(a.map((x,i)=>(x-ma)*(b[i]-mb)))/(a.length*sdv(a)*sdv(b))}}),best:Math.max(jm.acc,nm.acc)}}
function runKM(M:any,k:number){const{a,C,inertia}=kmeans(M.data.map(M.Z),k),{num,target}=M.S;
const st=Array.from({length:k},(_,j)=>{const m=M.data.filter((_:any,i:number)=>a[i]===j);return{j,size:m.length,avg:mean(m.map((r:R)=>+r[target])),pass:mean(m.map((r:R)=>+(+r[target]>=.5*M.top))),means:num.map((f:string)=>mean(m.map((r:R)=>+r[f])))}});
const ord=[...st].sort((x,y)=>x.avg-y.avg).map(s=>s.j),info=st.map(s=>{const rk=ord.indexOf(s.j);return{...s,rk,name:k===3?['At-risk','Average','High achievers'][rk]:k===2?['Lower','Higher'][rk]:'Tier '+(rk+1),color:`hsl(${k>1?rk/(k-1)*140:70} 62% 40%)`}});
return{a,C,inertia,info,nearest:(r:R)=>{const z=M.Z(r),ds=C.map((c:number[])=>sum(c.map((v,i)=>(v-z[i])**2)));return info[ds.indexOf(Math.min(...ds))]}}}
/* ---------- icons ---------- */
const IC:R={dash:'M3 13h4v8H3zM10 3h4v18h-4zM17 9h4v12h-4z',
data:'M4 6c0-1.7 3.6-3 8-3s8 1.3 8 3-3.6 3-8 3-8-1.3-8-3zM4 6v6c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6',
assoc:'M10 13a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 11a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1',
j48:'M12 3v6M12 9H6v5M12 9h6v5M6 14v4M18 14v4',
nb:'M5 5h14v14H5zM9 9h.01M15 9h.01M12 12h.01M9 15h.01M15 15h.01',
reg:'M3 3v18h18M7 15l4-4 3 3 5-6',
clu:'M8 8m-3 0a3 3 0 1 0 6 0a3 3 0 1 0-6 0M17 9m-3 0a3 3 0 1 0 6 0a3 3 0 1 0-6 0M12 17m-3 0a3 3 0 1 0 6 0a3 3 0 1 0-6 0',
cmp:'M12 3v18M5 21h14M5 7h14M5 7l-3 7a3 3 0 0 0 6 0zM19 7l-3 7a3 3 0 0 0 6 0z',
pred:'M12 12m-9 0a9 9 0 1 0 18 0a9 9 0 1 0-18 0M12 12m-4 0a4 4 0 1 0 8 0a4 4 0 1 0-8 0M12 12h.01',
about:'M12 12m-9 0a9 9 0 1 0 18 0a9 9 0 1 0-18 0M12 11v5M12 8h.01'};
const Icon=({n}:any)=><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={IC[n]}/></svg>;
/* ---------- chart primitives ---------- */
function useTip(){const ref=useRef<HTMLDivElement>(null),[tip,set]=useState<any>(null);
return{ref,show:(s:string)=>(e:any)=>{const b=ref.current!.getBoundingClientRect();set({x:e.clientX-b.left,y:e.clientY-b.top,s})},hide:()=>set(null),
el:tip&&<div className="pointer-events-none absolute z-20 text-xs px-2.5 py-1.5 rounded-md text-white shadow-lg whitespace-nowrap" style={{left:Math.min(tip.x+12,300),top:tip.y+12,background:'#0f172a'}}>{tip.s}</div>}}
function Scatter({pts,xl,yl,diag}:any){const T=useTip(),W=560,H=300,L=44,B=34;const xs=pts.map((p:any)=>p.x),ys=pts.map((p:any)=>p.y);
let[x0,x1,y0,y1]=[Math.min(...xs),Math.max(...xs),Math.min(...ys),Math.max(...ys)];if(diag){x0=y0=Math.min(x0,y0);x1=y1=Math.max(x1,y1)}
const px=(v:number)=>L+(v-x0)/(x1-x0||1)*(W-L-10),py=(v:number)=>H-B-(v-y0)/(y1-y0||1)*(H-B-10),tk=(a:number,b:number)=>[0,1,2,3,4].map(i=>a+(b-a)*i/4);
return<div ref={T.ref} className="relative"><svg viewBox={`0 0 ${W} ${H}`} className="w-full">
{tk(x0,x1).map((v,i)=><g key={i}><line x1={px(v)} x2={px(v)} y1={10} y2={H-B} stroke="var(--grid)"/><text x={px(v)} y={H-B+14} fontSize="10" textAnchor="middle" fill="var(--mu)">{f2(v,0)}</text></g>)}
{tk(y0,y1).map((v,i)=><g key={i}><line x1={L} x2={W-10} y1={py(v)} y2={py(v)} stroke="var(--grid)"/><text x={L-6} y={py(v)+3} fontSize="10" textAnchor="end" fill="var(--mu)">{f2(v,0)}</text></g>)}
{diag&&<line x1={px(x0)} y1={py(y0)} x2={px(x1)} y2={py(y1)} stroke="#1e3a8a" strokeDasharray="5 4" strokeWidth="1.5"/>}
{pts.map((p:any,i:number)=><circle key={i} className="dot" cx={px(p.x)} cy={py(p.y)} r={4} fill={p.c} fillOpacity=".7" stroke="#fff" strokeWidth=".5" onMouseMove={T.show(p.t)} onMouseLeave={T.hide}/>)}
<text x={(L+W)/2} y={H-4} fontSize="11" textAnchor="middle" fill="var(--mu)">{xl}</text><text transform={`translate(11 ${H/2}) rotate(-90)`} fontSize="11" textAnchor="middle" fill="var(--mu)">{yl}</text></svg>{T.el}</div>}
function Bars({g,s,max,fmt=(v:number)=>f2(v,0)}:any){const T=useTip(),W=560,H=260,B=30,L=34,M=max??Math.max(...g.flatMap((x:any)=>x.v))*1.15,bw=(W-L-10)/g.length;
return<div ref={T.ref} className="relative"><svg viewBox={`0 0 ${W} ${H}`} className="w-full">
{[0,.25,.5,.75,1].map(t=><g key={t}><line x1={L} x2={W-10} y1={H-B-t*(H-B-14)} y2={H-B-t*(H-B-14)} stroke="var(--grid)"/><text x={L-5} y={H-B-t*(H-B-14)+3} fontSize="10" textAnchor="end" fill="var(--mu)">{fmt(M*t)}</text></g>)}
{g.map((x:any,i:number)=><g key={i}>{x.v.map((v:number,j:number)=>{const w=bw*.7/x.v.length,h=Math.max(0,v/M*(H-B-14));return<rect key={j} className="bar" x={L+i*bw+bw*.15+j*w} y={H-B-h} width={w-2} height={h} rx="2" fill={s[j].c} onMouseMove={T.show(`${x.l}${s.length>1?' · '+s[j].n:''}: ${fmt(v)}`)} onMouseLeave={T.hide}/>})}
<text x={L+i*bw+bw/2} y={H-B+15} fontSize="10" textAnchor="middle" fill="var(--mu)">{x.l}</text></g>)}</svg>
{s.length>1&&<div className="flex gap-4 justify-center text-xs mu">{s.map((x:any)=><span key={x.n}><i className="inline-block w-2.5 h-2.5 rounded-sm mr-1" style={{background:x.c}}/>{x.n}</span>)}</div>}{T.el}</div>}
function Donut({d}:any){const T=useTip(),tot=sum(d.map((x:any)=>x.v)),Rr=70,C=2*Math.PI*Rr;let o=0;
return<div ref={T.ref} className="relative flex items-center gap-5"><svg viewBox="0 0 200 200" className="w-40"><g transform="rotate(-90 100 100)">{d.map((x:any)=>{const len=x.v/tot*C,el=<circle key={x.l} cx="100" cy="100" r={Rr} fill="none" stroke={x.c} strokeWidth="24" strokeDasharray={`${len} ${C-len}`} strokeDashoffset={-o} onMouseMove={T.show(`${x.l}: ${x.v} (${pc(x.v/tot)})`)} onMouseLeave={T.hide}/>;o+=len;return el})}</g>
<text x="100" y="106" textAnchor="middle" fontSize="26" fontWeight="700" fill="var(--fg)">{tot}</text></svg>
<div className="text-sm space-y-1.5">{d.map((x:any)=><div key={x.l}><i className="inline-block w-2.5 h-2.5 rounded-sm mr-2" style={{background:x.c}}/>{x.l} <b>{pc(x.v/tot)}</b> <span className="mu text-xs">({x.v})</span></div>)}</div>{T.el}</div>}
/* ---------- shared UI ---------- */
const Card=({t,children,cls=''}:any)=><section className={`glass p-4 md:p-5 fade ${cls}`}>{t&&<h3 className="ct">{t}</h3>}{children}</section>;
const PH=({t,s,e}:any)=><header className="ph mb-5 fade">{e&&<div className="eyebrow">{e}</div>}<h1 className="text-2xl md:text-[1.7rem] font-bold tracking-tight">{t}</h1><p className="mu text-sm mt-1.5 max-w-3xl leading-relaxed">{s}</p></header>;
const KPI=({l,v,s,c}:any)=><div className="kpi fade" style={{borderLeftColor:c}} title={s}><div className="kl">{l}</div><div className="kv">{v}</div>{s&&<div className="ks">{s}</div>}</div>;
const Ctl=({l,v,children}:any)=><label className="block text-xs font-semibold mu">{l} {v!==undefined&&<b className="text-[var(--fg)]">{v}</b>}<div className="mt-1.5">{children}</div></label>;
function STable({cols,rows,lim=40}:any){const[k,sk]=useState(cols[0].k),[d,sd]=useState(-1);const s=[...rows].sort((a:R,b:R)=>(a[k]>b[k]?1:a[k]<b[k]?-1:0)*d);
return<div className="overflow-x-auto"><table><thead><tr>{cols.map((c:R)=><th key={c.k} className="cursor-pointer select-none" title="Click to sort" onClick={()=>k===c.k?sd(-d):(sk(c.k),sd(-1))}>{c.l}{k===c.k?(d>0?' ▲':' ▼'):''}</th>)}</tr></thead>
<tbody>{s.slice(0,lim).map((r:R,i:number)=><tr key={i}>{cols.map((c:R)=><td key={c.k}>{c.f?c.f(r[c.k]):r[c.k]}</td>)}</tr>)}</tbody></table>{rows.length>lim&&<p className="mu text-xs mt-2">Showing top {lim} of {rows.length}</p>}</div>}
function Matrix({m}:any){const mx=Math.max(...m.cm.flat());return<div><table style={{width:'auto'}}><thead><tr><th>Actual ↓ / Predicted →</th>{CL.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{m.cm.map((r:number[],i:number)=><tr key={i}><td className="font-semibold">{CL[i]}</td>{r.map((v,j)=><td key={j} title={`${CL[i]} predicted as ${CL[j]}`} className="text-center font-bold" style={{background:`rgba(${i===j?'4,120,87':'185,28,28'},${v/mx*.5})`}}>{v}</td>)}</tr>)}</tbody></table></div>}
function Tree({n,lab,d=0}:any):any{if(!n.attr)return<div className="text-xs py-0.5 ml-3"><span className="mu">{lab}</span> → <b style={{color:CC[n.maj]}}>{n.maj}</b> <span className="mu">({n.n} students)</span></div>;
return<details open={d<2} className="ml-3 pl-3 border-l" style={{borderColor:'var(--bd)'}}><summary className="text-xs cursor-pointer py-0.5">{lab&&<span className="mu">{lab} → </span>}<b>{n.attr}?</b></summary>{Object.entries(n.kids).map(([k,v])=><Tree key={k} n={v} lab={`${n.attr} = ${k}`} d={d+1}/>)}</details>}
/* ---------- pages ---------- */
function Dashboard({M}:any){const{S,data,top,cls,corr}=M,att=S.num.find((c:string)=>/attend/i.test(c))||S.num[0],T=S.target;
const marks=data.map((r:R)=>+r[T]),lo=Math.min(...marks),hi=Math.max(...marks),nb=8,h=Array(nb).fill(0);marks.forEach((m:number)=>h[Math.min(nb-1,Math.floor((m-lo)/(hi-lo||1)*nb))]++);
const cnt=CL.map(c=>({l:c,v:data.filter((r:R)=>cls(+r[T])===c).length,c:CC[c]})),lead=[...corr].sort((a:R,b:R)=>Math.abs(b.v)-Math.abs(a.v))[0],bm=M.jm.acc>=M.nm.acc?'J48 decision tree':'Naive Bayes';
return<div><PH e="Overview" t="Executive Dashboard" s="Summary of cohort performance, the factors associated with it, and the predictive accuracy of the fitted models."/>
<div className="grid grid-cols-2 lg:grid-cols-5 gap-3 mb-4"><KPI l="Total students" v={data.length} c="#1e3a8a" s="Complete records used"/><KPI l="Average marks" v={f2(mean(marks),1)} c="#0e7490" s={`Out of ${top}`}/>
<KPI l="Average attendance" v={f2(mean(data.map((r:R)=>+r[att])),1)+(top?'%':'')} c="#0f766e" s={lbl(att)}/><KPI l="Pass rate" v={pc(mean(marks.map((m:number)=>+(m>=.5*top))))} c="#047857" s="Marks at or above 50% of maximum"/><KPI l="Best model accuracy" v={pc(M.best)} c="#b45309" s={bm}/></div>
<Card t="Key findings" cls="mb-4"><ul className="findings">
{lead&&<li><b>{lbl(lead.f)}</b> has the strongest linear association with {lbl(T)} (r = {f2(lead.v)}).</li>}
<li>{pc(cnt[0].v/data.length)} of students ({cnt[0].v}) fall in the Low performance band and may warrant early academic support.</li>
<li>{bm} delivers the highest hold-out accuracy at {pc(M.best)}; linear regression explains {pc(Math.max(0,M.reg.r2))} of the variance in {lbl(T)}.</li></ul></Card>
<div className="grid lg:grid-cols-2 gap-4"><Card t="Performance levels"><Donut d={cnt}/></Card>
<Card t="Marks distribution"><Bars g={h.map((v:number,i:number)=>({l:f2(lo+(hi-lo)*i/nb,0),v:[v]}))} s={[{n:'Students',c:'#1e3a8a'}]}/></Card>
<Card t={`${lbl(att)} vs ${lbl(T)}`}><Scatter xl={lbl(att)} yl={lbl(T)} pts={data.map((r:R)=>({x:+r[att],y:+r[T],c:CC[cls(+r[T])],t:`${r.Student_ID||''} ${att} ${r[att]} · ${T} ${r[T]}`}))}/></Card>
<Card t={`Correlation with ${lbl(T)}`}><Bars g={corr.map((c:R)=>({l:lbl(c.f).slice(0,12),v:[c.v]}))} s={[{n:'r',c:'#0f766e'}]} max={1} fmt={(v:number)=>f2(v,2)}/></Card></div></div>}
function Dataset({raw,rs,info,name,opts,setOpts,onFile,reset,M}:any){const[q,sq]=useState(''),[lv,slv]=useState('All'),[pg,spg]=useState(0);const cols=Object.keys(rs[0]||{}),T=M?.S.target;
const fl=rs.filter((r:R)=>(!q||Object.values(r).join(' ').toLowerCase().includes(q.toLowerCase()))&&(lv==='All'||!M||(typeof r[T]==='number'&&M.cls(r[T])===lv))),pages=Math.max(1,Math.ceil(fl.length/10));
return<div><PH e="Overview" t="Dataset Management" s="Import a CSV or Excel file, configure data cleaning, and inspect the records used for modelling. A sample cohort is preloaded."/>
<div className="grid lg:grid-cols-3 gap-4 mb-4"><Card t="Data source" cls="lg:col-span-1"><label className="drop" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();e.dataTransfer.files[0]&&onFile(e.dataTransfer.files[0])}}>
<div className="text-sm font-semibold">Select a file or drag it here</div><div className="mu text-xs mt-1">Accepted: .csv, .xlsx, .xls with a numeric marks column</div><input type="file" accept=".csv,.xlsx,.xls" hidden onChange={e=>e.target.files?.[0]&&onFile(e.target.files[0])}/></label>
<p className="text-xs mu mt-3">Active file: <b className="text-[var(--fg)]">{name}</b><br/>{raw.length} rows · {cols.length} columns</p><div className="flex gap-2 flex-wrap mt-3"><button className="btn" onClick={reset}>Restore sample data</button><button className="btn" onClick={()=>download(rs,name)}>Export cleaned CSV</button></div></Card>
<Card t="Preprocessing" cls="lg:col-span-2"><div className="grid sm:grid-cols-2 gap-4"><div><div className="text-xs font-semibold mu mb-2">Missing values ({info.missing} detected)</div><div className="flex gap-2 flex-wrap">{[['fill','Impute (mean / mode)'],['drop','Remove rows'],['keep','Retain as-is']].map(([k,l])=><button key={k} className={'btn '+(opts.missing===k?'on':'')} onClick={()=>setOpts({...opts,missing:k})}>{l}</button>)}</div></div>
<div><div className="text-xs font-semibold mu mb-2">Duplicate records ({info.dup} removed)</div><button className={'btn '+(opts.dedupe?'on':'')} onClick={()=>setOpts({...opts,dedupe:!opts.dedupe})}>{opts.dedupe?'Duplicates removed':'Duplicates retained'}</button></div></div>
<p className="text-xs mu mt-4">Models are trained on complete rows only. Detected target: <b className="text-[var(--fg)]">{T||'none found'}</b> · numeric features: {M?.S.num.length??0} · categorical features: {M?.S.cat.length??0}</p></Card></div>
<Card t={`Data preview (${fl.length} rows)`}><div className="flex gap-2 mb-3 flex-wrap"><input className="inp flex-1 min-w-40" placeholder="Search any value" value={q} onChange={e=>{sq(e.target.value);spg(0)}}/><select className="inp" value={lv} onChange={e=>{slv(e.target.value);spg(0)}}>{['All',...CL].map(c=>c==='All'?<option key={c} value={c}>All performance levels</option>:<option key={c} value={c}>{c} performance</option>)}</select></div>
<div className="overflow-x-auto"><table><thead><tr>{cols.map(c=><th key={c}>{lbl(c)}</th>)}</tr></thead><tbody>{fl.slice(pg*10,pg*10+10).map((r:R,i:number)=><tr key={i}>{cols.map(c=><td key={c} style={r[c]===''?{background:'rgba(185,28,28,.15)'}:{}}>{r[c]===''?'missing':String(r[c])}</td>)}</tr>)}</tbody></table></div>
<div className="flex items-center justify-between mt-3 text-xs"><span className="mu">Page {pg+1} of {pages}</span><div className="flex gap-2"><button className="btn" disabled={pg===0} onClick={()=>spg(pg-1)}>Previous</button><button className="btn" disabled={pg+1>=pages} onClick={()=>spg(pg+1)}>Next</button></div></div></Card></div>}
function Assoc({M}:any){const[ms,sms]=useState(.15),[mc,smc]=useState(.7),[oy,soy]=useState(true);const A=useMemo(()=>apriori(M.T,ms,mc,oy),[M,ms,mc,oy]);
return<div><PH e="Analysis" t="Association Rule Mining" s="The Apriori algorithm identifies combinations of student attributes that occur together frequently. Numeric columns are discretised into Low, Medium and High bands."/>
<Card cls="mb-4"><div className="grid sm:grid-cols-3 gap-5 items-end"><Ctl l="Minimum support" v={pc(ms)}><input type="range" min=".05" max=".5" step=".01" value={ms} onChange={e=>sms(+e.target.value)}/></Ctl><Ctl l="Minimum confidence" v={pc(mc)}><input type="range" min=".4" max="1" step=".01" value={mc} onChange={e=>smc(+e.target.value)}/></Ctl>
<label className="text-xs font-semibold flex items-center gap-2"><input type="checkbox" checked={oy} onChange={e=>soy(e.target.checked)}/> Restrict to rules that predict the result</label></div>
<p className="text-xs mu mt-3">{A.n} frequent itemsets produced <b className="text-[var(--fg)]">{A.rules.length}</b> rules. A lift above 1 indicates that the attributes co-occur more often than expected by chance.</p></Card>
<Card t="Discovered rules">{A.rules.length?<STable rows={A.rules} cols={[{k:'rule',l:'Rule'},{k:'sup',l:'Support',f:pc},{k:'conf',l:'Confidence',f:pc},{k:'lift',l:'Lift',f:(v:number)=>f2(v)}]}/>:<p className="mu text-sm">No rules meet the current thresholds. Reduce the minimum support or confidence.</p>}</Card></div>}
function ModelView({m,title,sub,extra}:any){return<div><PH e="Models" t={title} s={sub}/><div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4"><KPI l="Accuracy" v={pc(m.acc)} c="#1e3a8a" s="Held-out 30% test set"/><KPI l="Precision" v={pc(m.pr)} c="#0e7490" s="Macro average"/><KPI l="Recall" v={pc(m.rc)} c="#0f766e" s="Macro average"/><KPI l="F1-score" v={pc(m.f1)} c="#b45309" s="Macro average"/></div>
<div className="grid lg:grid-cols-2 gap-4"><Card t="Confusion matrix"><Matrix m={m}/><p className="text-xs mu mt-3">Green cells denote correct predictions; red cells denote misclassifications.</p></Card>
<Card t="Per-class results"><STable rows={m.per.map((p:R)=>({...p}))} cols={[{k:'c',l:'Class'},{k:'n',l:'Test students'},{k:'pr',l:'Precision',f:pc},{k:'rc',l:'Recall',f:pc},{k:'f1',l:'F1',f:pc}]}/></Card>{extra}</div></div>}
function Regression({M}:any){const g=M.reg,T=M.S.target;return<div><PH e="Models" t="Linear Regression" s={`Multiple linear regression estimates ${lbl(T)} from all numeric and categorical features.`}/>
<div className="grid grid-cols-3 gap-3 mb-4"><KPI l="MAE" v={f2(g.mae)} c="#0e7490" s="Mean absolute error"/><KPI l="RMSE" v={f2(g.rmse)} c="#b45309" s="Root mean squared error"/><KPI l="R²" v={f2(g.r2,3)} c="#1e3a8a" s="Variance explained"/></div>
<div className="grid lg:grid-cols-2 gap-4"><Card t="Actual vs predicted marks (test set)"><Scatter diag xl="Actual" yl="Predicted" pts={g.act.map((a:number,i:number)=>({x:a,y:g.pd[i],c:'#1e3a8a',t:`Actual ${f2(a,1)} · Predicted ${f2(g.pd[i],1)}`}))}/><p className="text-xs mu mt-2">Points close to the dashed line indicate accurate predictions.</p></Card>
<Card t="Feature impact (marks per +1 standard deviation)"><Bars g={g.impact.map((c:R)=>({l:lbl(c.f).slice(0,12),v:[Math.max(0,c.v)]}))} s={[{n:'Impact',c:'#0f766e'}]} fmt={(v:number)=>f2(v,1)}/></Card></div></div>}
function Clusters({M,K,k,sk}:any){const{num}=M.S,[xi,sx]=useState(0),[yi,sy]=useState(Math.min(1,num.length-1));
return<div><PH e="Analysis" t="Cluster Analysis" s="K-Means groups students with similar profiles. Features are standardised so that each contributes equally."/>
<Card cls="mb-4"><div className="grid sm:grid-cols-3 gap-5"><Ctl l="Number of clusters (k)" v={k}><input type="range" min="2" max="6" value={k} onChange={e=>sk(+e.target.value)}/></Ctl>{[['X axis',xi,sx],['Y axis',yi,sy]].map(([l,v,f]:any)=><Ctl key={l} l={l}><select className="inp w-full" value={v} onChange={e=>f(+e.target.value)}>{num.map((n:string,i:number)=><option key={n} value={i}>{lbl(n)}</option>)}</select></Ctl>)}</div></Card>
<div className="grid lg:grid-cols-5 gap-4"><Card t="Student groups" cls="lg:col-span-3"><Scatter xl={lbl(num[xi])} yl={lbl(num[yi])} pts={M.data.map((r:R,i:number)=>({x:+r[num[xi]],y:+r[num[yi]],c:K.info[K.a[i]].color,t:`${K.info[K.a[i]].name} · ${num[xi]} ${r[num[xi]]} · ${num[yi]} ${r[num[yi]]}`}))}/>
<div className="flex gap-4 flex-wrap text-xs mu justify-center">{[...K.info].sort((a:R,b:R)=>a.rk-b.rk).map((c:R)=><span key={c.j}><i className="inline-block w-2.5 h-2.5 rounded-sm mr-1" style={{background:c.color}}/>{c.name}</span>)}</div></Card>
<Card t="Cluster statistics" cls="lg:col-span-2"><div className="space-y-3">{[...K.info].sort((a:R,b:R)=>a.rk-b.rk).map((c:R)=><div key={c.j} className="rounded-lg p-3" style={{border:'1px solid var(--bd)',borderLeft:`4px solid ${c.color}`}}><div className="flex justify-between text-sm font-bold"><span>{c.name}</span><span>{c.size} students</span></div>
<div className="text-xs mu mt-1">Avg {lbl(M.S.target)}: <b className="text-[var(--fg)]">{f2(c.avg,1)}</b> · Pass rate: <b className="text-[var(--fg)]">{pc(c.pass)}</b></div><div className="text-xs mu mt-1">{num.map((n:string,i:number)=>`${lbl(n)} ${f2(c.means[i],1)}`).join(' · ')}</div></div>)}</div>
<p className="text-xs mu mt-3">Within-cluster variance (inertia): {f2(K.inertia,1)}</p></Card></div></div>}
function Compare({M}:any){const rows=[{m:'J48 decision tree',a:M.jm.acc,p:M.jm.pr,r:M.jm.rc,f:M.jm.f1},{m:'Naive Bayes',a:M.nm.acc,p:M.nm.pr,r:M.nm.rc,f:M.nm.f1}],w=M.jm.f1>=M.nm.f1?rows[0]:rows[1],gap=Math.abs(M.jm.f1-M.nm.f1);
return<div><PH e="Models" t="Model Comparison" s="Both classifiers are trained and evaluated on the same train/test split, so their scores are directly comparable."/>
<div className="grid lg:grid-cols-2 gap-4"><Card t="Classification metrics (%)"><Bars max={100} g={['Accuracy','Precision','Recall','F1'].map((l,i)=>({l,v:rows.map(r=>[r.a,r.p,r.r,r.f][i]*100)}))} s={[{n:'J48',c:'#1e3a8a'},{n:'Naive Bayes',c:'#0f766e'}]}/></Card>
<Card t="Summary"><STable rows={rows} cols={[{k:'m',l:'Model'},{k:'a',l:'Accuracy',f:pc},{k:'p',l:'Precision',f:pc},{k:'r',l:'Recall',f:pc},{k:'f',l:'F1',f:pc}]}/>
<div className="note mt-4"><b>{w.m}</b> achieves the highest F1-score ({pc(w.f)}), ahead by {f2(gap*100,1)} percentage points. {gap<.02?'The difference is small and may not be material.':'The difference is meaningful for this split.'}</div>
<p className="mt-3 text-sm mu">Regression: R² <b className="text-[var(--fg)]">{f2(M.reg.r2,3)}</b>, MAE <b className="text-[var(--fg)]">{f2(M.reg.mae)}</b>, RMSE <b className="text-[var(--fg)]">{f2(M.reg.rmse)}</b> marks. Association rules and clustering are descriptive techniques and do not have an accuracy score.</p></Card></div></div>}
function Predict({M,K}:any){const{S}=M,[v,sv]=useState<R>({});const r:R={...M.med,...Object.fromEntries(S.cat.map((c:string)=>[c,M.lev[c][0]])),...v},x=M.disc(r),jn=walk(M.tree,x),pb=M.nbPred(x),nbc=CL[pb.indexOf(Math.max(...pb))],pm=M.reg.pred(r),cu=K.nearest(r),weak=S.num.filter((f:string)=>x[f]==='Low');
return<div><PH e="Tools" t="Student Performance Prediction" s="Enter a student's attributes to obtain the predicted outcome from each model."/>
<div className="grid lg:grid-cols-5 gap-4"><Card t="Student attributes" cls="lg:col-span-2"><div className="space-y-4">{S.num.map((f:string)=><Ctl key={f} l={lbl(f)} v={f2(r[f],1)}><input type="range" min={M.rngs[f][0]} max={M.rngs[f][1]} step={(M.rngs[f][1]-M.rngs[f][0])/100} value={r[f]} onChange={e=>sv({...v,[f]:+e.target.value})}/></Ctl>)}
{S.cat.map((c:string)=><Ctl key={c} l={lbl(c)}><select className="inp w-full" value={r[c]} onChange={e=>sv({...v,[c]:e.target.value})}>{M.lev[c].map((l:string)=><option key={l}>{l}</option>)}</select></Ctl>)}</div></Card>
<div className="lg:col-span-3 space-y-4"><Card><div className="flex items-center gap-5"><svg viewBox="0 0 120 120" className="w-32 shrink-0"><circle cx="60" cy="60" r="50" fill="none" stroke="var(--grid)" strokeWidth="10"/><circle cx="60" cy="60" r="50" fill="none" stroke={CC[M.cls(pm)]} strokeWidth="10" strokeLinecap="butt" strokeDasharray={`${pm/M.top*314} 314`} transform="rotate(-90 60 60)" style={{transition:'stroke-dasharray .5s'}}/><text x="60" y="66" textAnchor="middle" fontSize="24" fontWeight="700" fill="var(--fg)">{f2(pm,1)}</text></svg>
<div><div className="kl">Regression: predicted {lbl(S.target)}</div><div className="text-xl font-bold mt-1" style={{color:CC[M.cls(pm)]}}>{M.cls(pm)} performance</div><div className="text-xs mu mt-1">Expected error of approximately ±{f2(M.reg.mae,1)} marks</div></div></div></Card>
<div className="grid sm:grid-cols-3 gap-3"><KPI l="J48 prediction" v={<span style={{color:CC[jn.maj]}}>{jn.maj}</span>} c="#1e3a8a" s={`Leaf confidence ${pc(jn.dist[jn.maj]/jn.n)}`}/><KPI l="Naive Bayes" v={<span style={{color:CC[nbc]}}>{nbc}</span>} c="#0f766e" s={`Probability ${pc(Math.max(...pb))}`}/><KPI l="Cluster" v={<span style={{color:cu.color}}>{cu.name}</span>} c="#b45309" s={`${cu.size} similar students`}/></div>
<Card t="Naive Bayes class probabilities">{CL.map((c,i)=><div key={c} className="mb-2"><div className="flex justify-between text-xs"><span>{c}</span><b>{pc(pb[i])}</b></div><div className="h-2 rounded-sm" style={{background:'var(--grid)'}}><div className="h-full rounded-sm" style={{width:pb[i]*100+'%',background:CC[c],transition:'width .4s'}}/></div></div>)}
<p className="text-xs mu mt-3">{weak.length?`Areas for improvement: ${weak.map(lbl).join(', ')} (in the lowest third of the cohort).`:'No numeric factor falls in the lowest third of the cohort.'}</p></Card></div></div></div>}
function About(){const it=[['Association rules (Apriori)','Identifies frequent attribute combinations and scores them by support, confidence and lift.'],['J48 decision tree','A C4.5-style tree using gain ratio, with depth and minimum-leaf limits to reduce overfitting.'],['Naive Bayes','A probabilistic classifier with Laplace smoothing applied to discretised features.'],['Linear regression','Solves the normal equations to estimate final marks; evaluated using MAE, RMSE and R².'],['K-Means clustering','K-means++ initialisation on standardised features, with per-cluster summary statistics.']];
return<div><PH e="Tools" t="About This Application" s="Intelligent Student Performance Analytics combines five data-mining techniques in a single workspace. All processing occurs in the browser; uploaded files are never transmitted."/>
<div className="grid md:grid-cols-2 gap-4">{it.map(([t,d],i)=><Card key={t}><div className="eyebrow">Method {i+1}</div><div className="font-bold text-sm">{t}</div><p className="mu text-sm mt-1.5 leading-relaxed">{d}</p></Card>)}<Card t="Data handling"><p className="mu text-sm leading-relaxed">Performance levels are defined as Low (below 50% of maximum marks), Medium (50% to 65%) and High (65% and above). Models are trained on 70% of rows and evaluated on the remaining 30%.</p></Card></div></div>}
/* ---------- app ---------- */
const NAV=[['dash','Dashboard','dash','Overview'],['data','Dataset','data','Overview'],['assoc','Association Rules','assoc','Analysis'],['clu','Clustering','clu','Analysis'],['j48','J48 Classification','j48','Models'],['nb','Naive Bayes','nb','Models'],['reg','Regression','reg','Models'],['cmp','Model Comparison','cmp','Models'],['pred','Student Prediction','pred','Tools'],['about','About','about','Tools']];
function App(){const[pg,setPg]=useState('dash'),[raw,setRaw]=useState<R[]>(()=>sample()),[name,setName]=useState('Sample cohort (240 students)'),[opts,setOpts]=useState({missing:'fill',dedupe:true}),[ver,setVer]=useState(0),[k,sk]=useState(3);
const info=useMemo(()=>prep(raw,opts),[raw,opts]),M=useMemo(()=>build(info.rs),[info]),K=useMemo(()=>M&&runKM(M,k),[M,k]);
const onFile=async(f:File)=>{try{let rows:R[];if(/\.csv$/i.test(f.name))rows=parseCSV(await f.text());else{const X=XLSX,wb=X.read(await f.arrayBuffer());rows=X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''})}
if(!rows.length)throw 0;setRaw(norm(rows));setName(f.name);setVer(v=>v+1)}catch(e){alert('The file could not be read. Please use a CSV or Excel sheet with a header row.')}};
const reset=()=>{setRaw(sample());setName('Sample cohort (240 students)');setVer(v=>v+1)};
let body:any;
if(pg==='data')body=<Dataset {...{raw,rs:info.rs,info,name,opts,setOpts,onFile,reset,M}}/>;
else if(pg==='about')body=<About/>;
else if(!M)body=<Card><h3 className="font-bold">This dataset cannot be modelled yet</h3><p className="mu text-sm mt-2">Models require at least 20 complete rows and a numeric marks column (for example, Final_Marks). Open the Dataset page to impute missing values or upload a different file.</p></Card>;
else body={dash:<Dashboard M={M}/>,assoc:<Assoc M={M}/>,
j48:<ModelView m={M.jm} title="J48 Decision Tree" sub="A C4.5-style decision tree that classifies students into performance levels." extra={<Card t="Decision tree" cls="lg:col-span-2"><Tree n={M.tree}/></Card>}/>,
nb:<ModelView m={M.nm} title="Naive Bayes Classifier" sub="Classifies performance level using the conditional probability of each attribute given each class." extra={<Card t="Class priors" cls="lg:col-span-2"><Bars max={100} g={CL.map(c=>({l:c,v:[M.nb.pri[c]*100]}))} s={[{n:'Prior',c:'#0f766e'}]} fmt={(v:number)=>f2(v,0)+'%'}/></Card>}/>,
reg:<Regression M={M}/>,clu:<Clusters M={M} K={K} k={k} sk={sk}/>,cmp:<Compare M={M}/>,pred:<Predict key={ver} M={M} K={K}/>}[pg];
return<div className="min-h-screen flex"><style>{CSS}</style>
<aside className="flex w-[4.25rem] md:w-64 shrink-0 flex-col gap-0.5 p-2 md:p-4 sticky top-0 h-screen overflow-y-auto z-30" style={{background:'var(--card)',borderRight:'1px solid var(--bd)'}}>
<div className="flex items-center justify-center md:justify-start gap-3 md:px-2 py-3 mb-1"><div className="brand">SP</div><div className="hidden md:block leading-tight"><div className="font-bold text-sm">Student Performance<br/>Analytics</div><div className="mu text-[0.68rem] mt-0.5">Decision-support workspace</div></div></div>
{NAV.map(([id,l,ic,g],i)=><div key={id} className="contents">{(i===0||NAV[i-1][3]!==g)&&<div className="grp hidden md:block">{g}</div>}<button className={'nav '+(pg===id?'on':'')} title={l} aria-label={l} aria-current={pg===id?'page':undefined} onClick={()=>setPg(id)}><Icon n={ic}/><span className="hidden md:inline">{l}</span></button></div>)}
<div className="hidden md:block mt-auto mu text-xs px-2 pt-4 leading-relaxed">{info.rs.length} students loaded<br/>{M?`Best accuracy ${pc(M.best)}`:'Insufficient data for modelling'}</div></aside>
<main className="flex-1 min-w-0 p-3 md:p-8 max-w-7xl w-full mx-auto" key={pg}>
<div className="topbar"><span>Dataset: <b className="text-[var(--fg)]">{name}</b></span><span>{info.rs.length} records · {NAV.find(n=>n[0]===pg)?.[1]}</span></div>{body}</main></div>}
export default App;
