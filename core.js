// Inti perhitungan SMKP (dipakai browser & node). Tanpa dependensi.
(function(root){
const W={I:10,II:15,III:17,IV:35,V:15,VI:3,VII:5};
function mandays(table,n,kelas,faktorPct){
  const r=table.find(x=>n>=x.min&&n<=x.max); if(!r) return null;
  const base=r[kelas], f=Math.min(faktorPct||0,10);
  const adj=Math.floor(base*f/100), total=base+adj, t1=Math.floor(total*0.10);
  return {base,adj,total,tahap1:t1,tahap2:total-t1};
}
// faktor penyesuaian: tiap penambah ya = +; pengurang (tersertifikasi) = -; maks 10%
function faktorPct(adders,certified){const n=adders.filter(Boolean).length; return Math.max(0,Math.min(10,(n>0?10:0)-(certified?10:0)));}
// skor: items [{id,element,max,nilai(null=belum, 'NA')}]
function skor(items){
  const el={};
  for(const it of items){ if(it.max==null) continue; if(it.nilai==='NA') continue;
    const e=el[it.element]=el[it.element]||{sum:0,max:0,n:0,filled:0};
    e.max+=it.max; e.n++; if(it.nilai!=null&&it.nilai!==''){e.sum+=Number(it.nilai);e.filled++;} }
  let total=0; for(const k in el){const e=el[k]; e.pct=e.max?e.sum/e.max*(W[k]||0):0; total+=e.pct; e.weight=W[k];}
  return {el,total};
}
function kategori(nilai,max){ if(nilai==='NA'||nilai==null||nilai==='') return null; nilai=Number(nilai);
  if(nilai>=max) return 'POSITIF'; return nilai<max/2?'MAYOR':'MINOR'; }
function cekKeterkaitan(links,nilai,maxMap,th){ th=th||50; const out=[];
  for(const l of links){const a=nilai[l.up],b=nilai[l.down]; if(a==null||b==null||a==='NA'||b==='NA') continue;
    const pa=a/maxMap[l.up]*100,pb=b/maxMap[l.down]*100;
    if(pb-pa>th||(pa<50&&pb>=75)) out.push({up:l.up,down:l.down,pa:Math.round(pa),pb:Math.round(pb),why:l.why}); }
  return out; }
// Slovin & stratified
function slovin(N,e){e=e||0.05; return Math.ceil(N/(1+N*e*e));}
function stratified(strata,n){ // strata [{key,N}] -> alokasi proporsional min 1
  const T=strata.reduce((s,x)=>s+x.N,0); if(!T) return []; let a=strata.map(x=>({key:x.key,N:x.N,n:Math.max(1,Math.round(n*x.N/T))}));
  return a;}
function rankArea(areas){return areas.slice().sort((x,y)=>(y.risiko||0)-(x.risiko||0));}
const api={W,mandays,faktorPct,skor,kategori,cekKeterkaitan,slovin,stratified,rankArea};
if(typeof module!=='undefined') module.exports=api; else root.SMKP=api;
})(this);
