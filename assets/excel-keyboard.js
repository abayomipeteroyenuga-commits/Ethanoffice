/* Ethan Excel keyboard repair v3 — document-level Excel navigation */
(function(){
'use strict';
const $=s=>document.querySelector(s), book=()=>window.ethanGetState?.().sheets.find(x=>x.id===window.ethanGetCurrentSheet());
const pos=()=>window.getXlCell?.()||{r:0,c:0}, letters='ABCDEFGHIJKLMNOPQRST', maxR=59,maxC=19;
const cell=(r,c)=>$('#sheetGrid td[data-r="'+r+'"][data-c="'+c+'"]');
const clamp=(v,a,b)=>Math.max(a,Math.min(b,v)), addr=(r,c)=>letters[c]+(r+1);
let anchor=null,edge=null,dragging=false,dragMoved=false,clip='';
const point=()=>({r:clamp(Number(pos().r)||0,0,maxR),c:clamp(Number(pos().c)||0,0,maxC)});
function bounds(){const a=anchor||point(),b=edge||point();return{r1:Math.min(a.r,b.r),r2:Math.max(a.r,b.r),c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c)}}
function cells(){const b=bounds(),a=[];for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)a.push({r,c});return a}
function paint(){
 const host=$('#sheetGrid');if(!host)return;host.querySelectorAll('.sel,.rangeSelected').forEach(x=>x.classList.remove('sel','rangeSelected'));
 const p=point();cells().forEach(x=>{const el=cell(x.r,x.c);if(el)el.classList.add(x.r===p.r&&x.c===p.c?'sel':'rangeSelected')});
 const b=book(),items=cells(),values=b?items.map(x=>b.grid[x.r]?.[x.c]??''):[];
 const nums=values.map(Number).filter((v,i)=>String(values[i]).trim()!==''&&Number.isFinite(v)),sum=nums.reduce((a,v)=>a+v,0);
 const name=$('#cellName');if(name)name.textContent=addr(p.r,p.c);
 const bar=$('#formulaBar');if(bar&&document.activeElement!==bar)bar.value=b?.grid[p.r]?.[p.c]??'';
 const st=$('#excelStatus');if(st)st.textContent=addr(p.r,p.c)+' • '+(items.length>1?items.length+' cells selected':'Ready');
 const sm=$('#excelSummary');if(sm)sm.textContent='Sum: '+sum.toLocaleString()+' • Average: '+(nums.length?(sum/nums.length).toFixed(2):0)+' • Count: '+nums.length;
}
function commit(el){
 const b=book();if(!b||!el)return;const r=+el.dataset.r,c=+el.dataset.c;
 // Core blur handler commits data-raw when present; retain formulas instead of their rendered value.
 if(!Object.prototype.hasOwnProperty.call(el.dataset,'raw')&&String(b.grid[r]?.[c]??'').startsWith('='))el.dataset.raw=String(b.grid[r][c]);
}
function focusCell(r,c,extend,old){
 r=clamp(r,0,maxR);c=clamp(c,0,maxC);
 if(old)commit(old);
 const p=pos();p.r=r;p.c=c;
 if(extend){anchor=anchor||{r:old?+old.dataset.r:p.r,c:old?+old.dataset.c:p.c};edge={r,c}}
 else{anchor={r,c};edge={r,c}}
 const el=cell(r,c);if(el){el.focus({preventScroll:true});try{getSelection()?.removeAllRanges()}catch{}}
 paint();
}
function activeCell(e){
 let el=e.target?.closest?.('#sheetGrid td[data-r][data-c]');
 if(!el)el=document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]');
 if(!el){const p=point();el=cell(p.r,p.c)}
 return el;
}
function onKey(e){
 const el=activeCell(e);if(!el||!$('#sheetWorkspace')||$('#sheetWorkspace').classList.contains('hidden'))return;
 const r=+el.dataset.r,c=+el.dataset.c,k=e.key,mod=e.ctrlKey||e.metaKey;
 if(mod&&['c','x','v'].includes(k.toLowerCase())){
  e.preventDefault();e.stopImmediatePropagation();
  if(k.toLowerCase()==='v'){
   try{navigator.clipboard.readText().then(paste).catch(()=>paste(clip))}catch{paste(clip)}
  }else{
   const b=book(),q=bounds();clip=Array.from({length:q.r2-q.r1+1},(_,ri)=>Array.from({length:q.c2-q.c1+1},(_,ci)=>String(b.grid[q.r1+ri]?.[q.c1+ci]??'')).join('\t')).join('\n');
   try{navigator.clipboard?.writeText(clip).catch(()=>{})}catch{}
   if(k.toLowerCase()==='x'){cells().forEach(p=>b.grid[p.r][p.c]='');window.ethanSave?.();window.drawGrid?.(b);focusCell(q.r1,q.c1,false,null)}
  }return;
 }
 if(k==='F2'){e.preventDefault();e.stopImmediatePropagation();return}
 if(k==='Escape'){e.preventDefault();e.stopImmediatePropagation();const b=book();if(b){el.textContent=String(window.displayCell?.(b.grid[r][c],b.grid)??b.grid[r][c]??'');delete el.dataset.raw}return}
 if(!['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End'].includes(k))return;
 e.preventDefault();e.stopImmediatePropagation();
 let nr=r,nc=c,extend=e.shiftKey;
 if(k==='ArrowUp'||(k==='Enter'&&e.shiftKey))nr--;
 if(k==='ArrowDown'||(k==='Enter'&&!e.shiftKey))nr++;
 if(k==='ArrowLeft'||(k==='Tab'&&e.shiftKey))nc--;
 if(k==='ArrowRight'||(k==='Tab'&&!e.shiftKey))nc++;
 if(k==='Home'){nc=0;if(mod)nr=0}
 if(k==='End'){nc=maxC;if(mod)nr=maxR}
 focusCell(nr,nc,extend,el);
}
function paste(value){
 const b=book();if(!b)return;const p=point(),rows=String(value??'').replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));
 rows.forEach((row,ri)=>row.forEach((v,ci)=>{if(p.r+ri<=maxR&&p.c+ci<=maxC)b.grid[p.r+ri][p.c+ci]=v}));
 window.ethanSave?.();window.drawGrid?.(b);focusCell(p.r,p.c,false,null);
}
function bind(){
 const host=$('#sheetGrid');if(!host)return;
 host.querySelectorAll('td[data-r][data-c]').forEach(td=>{if(!td.hasAttribute('tabindex'))td.tabIndex=0});
 if(host.dataset.keyboardRepairV3==='1')return;host.dataset.keyboardRepairV3='1';
 host.addEventListener('pointerdown',e=>{
  const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;const p={r:+el.dataset.r,c:+el.dataset.c};
  if(e.shiftKey){anchor=anchor||point();edge=p}else{anchor=p;edge=p}
  const q=pos();q.r=p.r;q.c=p.c;dragging=true;dragMoved=false;paint();
 },true);
 host.addEventListener('pointerover',e=>{
  if(!dragging)return;const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;const p={r:+el.dataset.r,c:+el.dataset.c};
  if(!edge||p.r!==edge.r||p.c!==edge.c)dragMoved=true;edge=p;const q=pos();q.r=p.r;q.c=p.c;paint();
 },true);
 window.addEventListener('pointerup',()=>{dragging=false});
 host.addEventListener('click',e=>{
  const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;const p={r:+el.dataset.r,c:+el.dataset.c};
  if(e.shiftKey){anchor=anchor||point();edge=p}else if(!dragMoved){anchor=p;edge=p}
  const q=pos();q.r=p.r;q.c=p.c;dragMoved=false;setTimeout(paint,0);
 });
}
document.addEventListener('keydown',onKey,true);
const oldDraw=window.drawGrid;
if(typeof oldDraw==='function')window.drawGrid=function(s){
 const result=oldDraw(s);bind();if(!anchor){const p=point();anchor=p;edge=p}
 const p=point();setTimeout(()=>{bind();paint();if(!document.activeElement?.matches?.('#sheetGrid td[data-r][data-c]'))cell(p.r,p.c)?.focus({preventScroll:true})},0);
 return result;
};
window.xlAutoSum=function(){
 const b=book();if(!b)return;const q=bounds(),p=point();let r=q.r2+1,c=clamp(p.c,q.c1,q.c2);
 if(r>maxR)r=q.r1-1;if(r<0){r=clamp(p.r,0,maxR);c=clamp(p.c,0,maxC)}
 const formula='=SUM('+addr(q.r1,q.c1)+':'+addr(q.r2,q.c2)+')';b.grid[r][c]=formula;window.ethanSave?.();
 const a=pos();a.r=r;a.c=c;anchor={r,c};edge={r,c};window.drawGrid?.(b);cell(r,c)?.focus();paint();
 if(window.toast)window.toast('AutoSum inserted: '+formula);
};
window.ethanSelectRange=spec=>{
 const m=/^([A-T])(\d+):([A-T])(\d+)$/i.exec(String(spec||''));if(!m)return false;
 anchor={r:clamp(+m[2]-1,0,maxR),c:m[1].toUpperCase().charCodeAt(0)-65};
 edge={r:clamp(+m[4]-1,0,maxR),c:m[3].toUpperCase().charCodeAt(0)-65};
 const p=pos();p.r=edge.r;p.c=edge.c;paint();cell(edge.r,edge.c)?.focus();return true;
};
document.addEventListener('DOMContentLoaded',bind);setTimeout(bind,0);
})();
