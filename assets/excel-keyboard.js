/* Ethan Excel keyboard and selection repair v2 — cell-first navigation */
(function(){
'use strict';
const $=s=>document.querySelector(s);
const book=()=>window.ethanGetState?.().sheets.find(x=>x.id===window.ethanGetCurrentSheet());
const pos=()=>window.getXlCell?.()||{r:0,c:0};
const COLS='ABCDEFGHIJKLMNOPQRST';
const MAX_R=59,MAX_C=19;
const clamp=(n,min,max)=>Math.max(min,Math.min(max,n));
const cell=(r,c)=>$('#sheetGrid td[data-r="'+r+'"][data-c="'+c+'"]');
const address=(r,c)=>COLS[c]+(r+1);
let anchor=null,focusPos=null,dragging=false,dragMoved=false,clipboard='';
const current=()=>({r:pos().r||0,c:pos().c||0});
function range(){
 const a=anchor||current(),b=focusPos||current(),out=[];
 for(let r=Math.min(a.r,b.r);r<=Math.max(a.r,b.r);r++)for(let c=Math.min(a.c,b.c);c<=Math.max(a.c,b.c);c++)out.push({r,c});
 return out;
}
function paint(){
 const host=$('#sheetGrid');if(!host)return;
 host.querySelectorAll('td.sel,td.rangeSelected').forEach(x=>x.classList.remove('sel','rangeSelected'));
 const p=current(),items=range();
 items.forEach(x=>{const el=cell(x.r,x.c);if(el)el.classList.add(x.r===p.r&&x.c===p.c?'sel':'rangeSelected')});
 const b=book(),vals=b?items.map(x=>b.grid[x.r]?.[x.c]??''):[];
 const nums=vals.map(v=>Number(v)).filter((v,i)=>String(vals[i]).trim()!==''&&Number.isFinite(v));
 const sum=nums.reduce((a,v)=>a+v,0);
 const name=$('#cellName');if(name)name.textContent=address(p.r,p.c);
 const formula=$('#formulaBar');if(formula&&document.activeElement!==formula)formula.value=b?.grid[p.r]?.[p.c]??'';
 const status=$('#excelStatus');if(status)status.textContent=address(p.r,p.c)+' • '+(items.length>1?items.length+' cells selected':'Ready');
 const summary=$('#excelSummary');if(summary)summary.textContent='Sum: '+sum.toLocaleString()+' • Average: '+(nums.length?(sum/nums.length).toFixed(2):0)+' • Count: '+nums.length;
}
function commit(el){
 const b=book();if(!b||!el)return;
 const r=+el.dataset.r,c=+el.dataset.c;
 // The core grid's blur handler commits the cell. Preserve the underlying formula
 // in data-raw before focus moves, otherwise it mistakes the displayed result for input.
 if(!Object.prototype.hasOwnProperty.call(el.dataset,'raw')&&String(b.grid[r]?.[c]??'').startsWith('=')){
  el.dataset.raw=String(b.grid[r][c]);
 }
}
function setSelection(r,c,extend){
 const before=current(),next={r:clamp(r,0,MAX_R),c:clamp(c,0,MAX_C)};
 const p=pos();p.r=next.r;p.c=next.c;
 if(extend){if(!anchor)anchor={r:before.r,c:before.c};focusPos=next}
 else{anchor={...next};focusPos={...next}}
 paint();
 const el=cell(next.r,next.c);
 if(el){el.focus({preventScroll:true});try{const sel=getSelection();sel?.removeAllRanges()}catch{}}
}
function moveFrom(el,r,c,extend){
 commit(el);
 setSelection(r,c,extend);
}
function copyRange(cut){
 const b=book();if(!b)return;
 const items=range(),a=anchor||current(),z=focusPos||current();
 const minR=Math.min(a.r,z.r),minC=Math.min(a.c,z.c),maxR=Math.max(a.r,z.r),maxC=Math.max(a.c,z.c);
 clipboard=Array.from({length:maxR-minR+1},(_,ri)=>Array.from({length:maxC-minC+1},(_,ci)=>String(b.grid[minR+ri]?.[minC+ci]??'')).join('\t')).join('\n');
 try{navigator.clipboard?.writeText(clipboard).catch(()=>{})}catch{}
 if(cut){items.forEach(x=>b.grid[x.r][x.c]='');window.ethanSave?.();window.drawGrid?.(b);setSelection(minR,minC,false)}
}
function pasteText(text){
 const b=book();if(!b)return;
 const p=current(),rows=String(text??'').replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));
 rows.forEach((row,ri)=>row.forEach((v,ci)=>{const r=p.r+ri,c=p.c+ci;if(r<=MAX_R&&c<=MAX_C)b.grid[r][c]=v}));
 window.ethanSave?.();window.drawGrid?.(b);setSelection(p.r,p.c,false);
}
function keydown(e){
 const el=e.target.closest?.('#sheetGrid td[data-r][data-c]');if(!el)return;
 const r=+el.dataset.r,c=+el.dataset.c,key=e.key;
 const mod=e.ctrlKey||e.metaKey;
 if(mod&&['c','x','v'].includes(key.toLowerCase())){
  e.preventDefault();e.stopImmediatePropagation();
  if(key.toLowerCase()==='v'){
   try{navigator.clipboard?.readText().then(pasteText).catch(()=>pasteText(clipboard))}catch{pasteText(clipboard)}
  }else copyRange(key.toLowerCase()==='x');
  return;
 }
 if(key==='F2'){e.preventDefault();e.stopImmediatePropagation();return}
 if(key==='Escape'){e.preventDefault();e.stopImmediatePropagation();const b=book();if(b){const p=current();el.textContent=String(window.displayCell?.(b.grid[p.r]?.[p.c],b.grid)??b.grid[p.r]?.[p.c]??'');delete el.dataset.raw}return}
 if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End'].includes(key)){
  e.preventDefault();e.stopImmediatePropagation();
  let nr=r,nc=c,extend=e.shiftKey;
  if(key==='ArrowUp')nr--;
  if(key==='ArrowDown'||(key==='Enter'&&!e.shiftKey))nr++;
  if(key==='Enter'&&e.shiftKey)nr--;
  if(key==='ArrowLeft'||(key==='Tab'&&e.shiftKey))nc--;
  if(key==='ArrowRight'||(key==='Tab'&&!e.shiftKey))nc++;
  if(key==='Home'){nc=mod?0:0;if(mod)nr=0}
  if(key==='End'){nc=MAX_C;if(mod)nr=MAX_R}
  moveFrom(el,nr,nc,extend);return;
 }
}
function bind(){
 const host=$('#sheetGrid');if(!host||host.dataset.keyboardRepairV2==='1')return;
 host.dataset.keyboardRepairV2='1';
 host.addEventListener('keydown',keydown,true);
 host.addEventListener('input',e=>{const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;el.dataset.raw=el.innerText;const f=$('#formulaBar');if(f)f.value=el.innerText;const s=$('#excelSaveState');if(s)s.textContent='Editing…'},true);
 host.addEventListener('pointerdown',e=>{
  const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;
  const p={r:+el.dataset.r,c:+el.dataset.c};
  if(e.shiftKey){anchor=anchor||current();focusPos=p}else{anchor=p;focusPos=p}
  const q=pos();q.r=p.r;q.c=p.c;dragging=true;dragMoved=false;paint();
 },true);
 host.addEventListener('pointerover',e=>{
  if(!dragging)return;const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;
  const p={r:+el.dataset.r,c:+el.dataset.c};if(!focusPos||p.r!==focusPos.r||p.c!==focusPos.c)dragMoved=true;
  focusPos=p;const q=pos();q.r=p.r;q.c=p.c;paint();
 },true);
 window.addEventListener('pointerup',()=>{dragging=false});
 host.addEventListener('click',e=>{
  const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;
  const p={r:+el.dataset.r,c:+el.dataset.c};
  if(e.shiftKey){anchor=anchor||current();focusPos=p}else if(!dragMoved){anchor=p;focusPos=p}
  const q=pos();q.r=p.r;q.c=p.c;dragMoved=false;setTimeout(paint,0);
 },false);
}
const draw=window.drawGrid;
if(typeof draw==='function')window.drawGrid=function(s){
 const result=draw(s);bind();
 if(!anchor){const p=current();anchor={...p};focusPos={...p}}
 setTimeout(paint,0);return result;
};
window.xlAutoSum=function(){
 const b=book();if(!b)return;
 const a=anchor||current(),z=focusPos||current();
 const minR=Math.min(a.r,z.r),maxR=Math.max(a.r,z.r),minC=Math.min(a.c,z.c),maxC=Math.max(a.c,z.c);
 const p=current();let targetR=maxR+1,targetC=clamp(p.c,minC,maxC);
 if(targetR>MAX_R)targetR=minR-1;
 if(targetR<0){targetR=clamp(p.r,0,MAX_R);targetC=clamp(p.c,0,MAX_C)}
 const formula='=SUM('+address(minR,minC)+':'+address(maxR,maxC)+')';
 b.grid[targetR][targetC]=formula;window.ethanSave?.();
 const q=pos();q.r=targetR;q.c=targetC;anchor={r:targetR,c:targetC};focusPos={r:targetR,c:targetC};
 window.drawGrid?.(b);const el=cell(targetR,targetC);el?.focus();paint();
 if(window.toast)window.toast('AutoSum inserted: '+formula);
};
window.ethanSelectRange=function(spec){
 const m=/^([A-T])(\d+):([A-T])(\d+)$/i.exec(String(spec||''));if(!m)return false;
 const a={r:clamp(+m[2]-1,0,MAX_R),c:m[1].toUpperCase().charCodeAt(0)-65};
 const z={r:clamp(+m[4]-1,0,MAX_R),c:m[3].toUpperCase().charCodeAt(0)-65};
 anchor=a;focusPos=z;const p=pos();p.r=z.r;p.c=z.c;paint();cell(z.r,z.c)?.focus();return true;
};
document.addEventListener('DOMContentLoaded',bind);setTimeout(bind,0);
})();
