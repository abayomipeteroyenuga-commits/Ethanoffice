/* Ethan Excel keyboard and selection repair v1 */
(function(){
'use strict';
const $=s=>document.querySelector(s);
const book=()=>window.ethanGetState?.().sheets.find(x=>x.id===window.ethanGetCurrentSheet());
const pos=()=>window.getXlCell?window.getXlCell():{r:0,c:0};
const cols='ABCDEFGHIJKLMNOPQRST';
let start=null,end=null,dragging=false,copied='';
const clamp=(n,min,max)=>Math.max(min,Math.min(max,n));
const cell=(r,c)=>$('#sheetGrid td[data-r="'+r+'"][data-c="'+c+'"]');
const addr=(r,c)=>cols[c]+(r+1);
function selected(){
 const a=start||pos(),z=end||a,out=[];
 for(let r=Math.min(a.r,z.r);r<=Math.max(a.r,z.r);r++)
  for(let c=Math.min(a.c,z.c);c<=Math.max(a.c,z.c);c++)out.push({r,c});
 return out;
}
function paint(){
 const g=$('#sheetGrid');if(!g)return;
 g.querySelectorAll('td.sel,td.rangeSelected').forEach(x=>x.classList.remove('sel','rangeSelected'));
 const a=pos();
 selected().forEach(p=>{const t=cell(p.r,p.c);if(t)t.classList.add(p.r===a.r&&p.c===a.c?'sel':'rangeSelected')});
 const b=book();if(!b)return;
 const ps=selected(),vals=ps.map(p=>b.grid[p.r]?.[p.c]??'');
 const nums=vals.map(Number).filter((n,i)=>String(vals[i]).trim()!==''&&Number.isFinite(n));
 const sum=nums.reduce((x,y)=>x+y,0);
 const status=$('#excelStatus');if(status)status.textContent=addr(a.r,a.c)+' • '+(ps.length>1?ps.length+' cells selected':'Ready');
 const summary=$('#excelSummary');if(summary)summary.textContent='Sum: '+sum.toLocaleString()+' • Average: '+(nums.length?(sum/nums.length).toFixed(2):0)+' • Count: '+nums.length;
}
function commit(t){
 const b=book();if(!b||!t)return;
 const r=+t.dataset.r,c=+t.dataset.c;
 b.grid[r][c]=String(t.dataset.raw!==undefined?t.dataset.raw:t.innerText).trim();
 delete t.dataset.raw;window.ethanSave?.();
}
function activate(r,c,extend){
 const b=book();if(!b)return;
 r=clamp(r,0,59);c=clamp(c,0,19);
 const p=pos();p.r=r;p.c=c;
 if(!extend){start={r,c};end={r,c}}else{start=start||{r:p.r,c:p.c};end={r,c}}
 const t=cell(r,c);if(t){paint();t.focus({preventScroll:true})}
}
function caretAtEdge(t,key){
 const sel=window.getSelection?.();if(!sel||!sel.isCollapsed||!sel.anchorNode)return true;
 if(!t.contains(sel.anchorNode))return true;
 const range=document.createRange();range.selectNodeContents(t);
 if(key==='ArrowLeft'||key==='ArrowUp'){range.setEnd(sel.anchorNode,sel.anchorOffset);return range.toString().length===0}
 if(key==='ArrowRight'||key==='ArrowDown'){range.setStart(sel.anchorNode,sel.anchorOffset);return range.toString().length===0}
 return true;
}
function paste(text){
 const b=book();if(!b)return;
 const ps=selected(),a=pos(),rows=String(text).replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));
 rows.forEach((row,ri)=>row.forEach((v,ci)=>{const r=a.r+ri,c=a.c+ci;if(r<60&&c<20)b.grid[r][c]=v}));
 window.ethanSave?.();window.drawGrid?.(b);activate(a.r,a.c,false);toast('Pasted into worksheet');
}
function toast(s){if(window.toast)window.toast(s)}
function bind(){
 const g=$('#sheetGrid');if(!g||g.dataset.keyboardRepair==='1')return;
 g.dataset.keyboardRepair='1';
 g.addEventListener('pointerdown',e=>{
  const t=e.target.closest('td[data-r][data-c]');if(!t)return;
  dragging=true;const p={r:+t.dataset.r,c:+t.dataset.c};start=p;end=p;
  const a=pos();a.r=p.r;a.c=p.c;paint();
 });
 g.addEventListener('pointerover',e=>{
  if(!dragging)return;const t=e.target.closest('td[data-r][data-c]');if(!t)return;
  end={r:+t.dataset.r,c:+t.dataset.c};paint();
 });
 window.addEventListener('pointerup',()=>{dragging=false});
 g.addEventListener('click',e=>{
  const t=e.target.closest('td[data-r][data-c]');if(!t)return;
  const p={r:+t.dataset.r,c:+t.dataset.c};
  if(e.shiftKey){end=p}else if(!dragging){start=p;end=p}
  const a=pos();a.r=p.r;a.c=p.c;setTimeout(paint,0);
 });
 g.addEventListener('keydown',e=>{
  const t=e.target.closest('td[data-r][data-c]');if(!t)return;
  const r=+t.dataset.r,c=+t.dataset.c;
  const key=e.key;
  if((e.ctrlKey||e.metaKey)&&['c','x','v'].includes(key.toLowerCase())){
   e.preventDefault();e.stopImmediatePropagation();
   if(key.toLowerCase()==='c'||key.toLowerCase()==='x'){
    const b=book();copied=selected().map(p=>b.grid[p.r]?.[p.c]??'').join('\t');
    navigator.clipboard?.writeText(copied).catch(()=>{});
    if(key.toLowerCase()==='x'){selected().forEach(p=>b.grid[p.r][p.c]='');window.ethanSave?.();window.drawGrid(b)}
   }else navigator.clipboard?.readText().then(paste).catch(()=>paste(copied));
   return;
  }
  if(key==='F2'){e.preventDefault();e.stopImmediatePropagation();t.focus();return}
  if(key==='Enter'||key==='Tab'){
   e.preventDefault();e.stopImmediatePropagation();commit(t);
   activate(r+(key==='Enter'?(e.shiftKey?-1:1):0),c+(key==='Tab'?(e.shiftKey?-1:1):0),false);return;
  }
  if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(key)){
   if(!e.shiftKey&&!e.ctrlKey&&!e.metaKey&&!caretAtEdge(t,key))return;
   e.preventDefault();e.stopImmediatePropagation();commit(t);
   const dr=key==='ArrowUp'?-1:key==='ArrowDown'?1:0,dc=key==='ArrowLeft'?-1:key==='ArrowRight'?1:0;
   activate(e.ctrlKey||e.metaKey?(dr<0?0:dr>0?59:0):r+dr,e.ctrlKey||e.metaKey?(dc<0?0:dc>0?19:0):c+dc,e.shiftKey);return;
  }
  if(key==='Home'||key==='End'){
   e.preventDefault();e.stopImmediatePropagation();commit(t);
   activate(e.ctrlKey?(key==='Home'?0:59):r,e.ctrlKey?(key==='Home'?0:19):(key==='Home'?0:19),e.shiftKey);
  }
 },true);
}
const originalDraw=window.drawGrid;
if(typeof originalDraw==='function')window.drawGrid=function(s){const result=originalDraw(s);bind();if(!start){const p=pos();start={r:p.r,c:p.c};end={r:p.r,c:p.c}}setTimeout(paint,0);return result};
window.xlAutoSum=function(){
 const b=book();if(!b)return;const ps=selected(),a=pos();
 let minR=Math.min(...ps.map(p=>p.r)),maxR=Math.max(...ps.map(p=>p.r)),minC=Math.min(...ps.map(p=>p.c)),maxC=Math.max(...ps.map(p=>p.c));
 let targetR=maxR+1,targetC=clamp(a.c,minC,maxC);
 if(targetR>59){targetR=minR-1}
 if(targetR<0){targetR=a.r;targetC=a.c}
 const formula='=SUM('+addr(minR,minC)+':'+addr(maxR,maxC)+')';
 b.grid[targetR][targetC]=formula;window.ethanSave?.();const p=pos();p.r=targetR;p.c=targetC;start={r:targetR,c:targetC};end={r:targetR,c:targetC};window.drawGrid(b);const t=cell(targetR,targetC);if(t)t.focus();toast('AutoSum inserted: '+formula);
};
window.ethanSelectRange=(range)=>{
 const m=/^([A-T])(\d+):([A-T])(\d+)$/i.exec(String(range||''));if(!m)return;
 start={r:+m[2]-1,c:m[1].toUpperCase().charCodeAt(0)-65};end={r:+m[4]-1,c:m[3].toUpperCase().charCodeAt(0)-65};const p=pos();p.r=end.r;p.c=end.c;paint();
};
document.addEventListener('DOMContentLoaded',()=>{bind();setTimeout(bind,100)});
setTimeout(bind,0);
})();
