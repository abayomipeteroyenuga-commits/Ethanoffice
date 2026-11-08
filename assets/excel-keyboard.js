/* Ethan Office Excel keyboard v5 — bind directly to the live worksheet and active cell */
(function(){
'use strict';
if(window.__ethanExcelKeyboardV5)return;window.__ethanExcelKeyboardV5=true;
const $=s=>document.querySelector(s);
const getSheet=()=>{const st=window.ethanGetState?.(),id=window.ethanGetCurrentSheet?.();return st?.sheets?.find(s=>s.id===id)};
const getPos=()=>window.getXlCell?.()||{r:0,c:0};
const getGrid=()=>$('#sheetGrid');
const getCell=(r,c)=>getGrid()?.querySelector('td[data-r="'+r+'"][data-c="'+c+'"]');
const isOpen=()=>document.querySelector('.page.active')?.id==='page-excel'&&!$('#sheetWorkspace')?.classList.contains('hidden');
let start=null,finish=null,copy='';
function set(r,c){const p=getPos(),s=getSheet();if(!s)return;r=Math.max(0,Math.min((s.grid?.length||60)-1,r));c=Math.max(0,Math.min((s.grid?.[0]?.length||20)-1,c));p.r=r;p.c=c;const td=getCell(r,c);if(!td)return;td.focus({preventScroll:true});td.click();td.classList.add('sel');$('#cellName').textContent=String.fromCharCode(65+c)+(r+1);$('#formulaBar').value=s.grid[r]?.[c]??'';const a=start||{r,c},b=finish||{r,c};getGrid()?.querySelectorAll('.rangeSelected').forEach(x=>x.classList.remove('rangeSelected'));if(start&&finish){for(let y=Math.min(a.r,b.r);y<=Math.max(a.r,b.r);y++)for(let x=Math.min(a.c,b.c);x<=Math.max(a.c,b.c);x++){const el=getCell(y,x);if(el&&!(y===r&&x===c))el.classList.add('rangeSelected')}}}
function onKey(e){
 if(!isOpen()||e.altKey||e.isComposing)return;
 const grid=getGrid();if(!grid)return;
 let td=e.target?.closest?.('#sheetGrid td[data-r][data-c]')||document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]')||grid.querySelector('td.sel[data-r][data-c]');
 if(!td)return;
 const s=getSheet();if(!s)return;
 const r=+td.dataset.r,c=+td.dataset.c,k=e.key,mod=e.ctrlKey||e.metaKey;
 if(mod&&['c','x','v'].includes(k.toLowerCase())){
  e.preventDefault();e.stopImmediatePropagation();
  if(k.toLowerCase()==='v'){const paste=t=>{const rows=String(t||'').replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));rows.forEach((row,i)=>row.forEach((v,j)=>{if(s.grid[r+i]&&c+j<20)s.grid[r+i][c+j]=v}));window.ethanSave?.();window.drawGrid?.(s);set(r,c)};if(navigator.clipboard?.readText)navigator.clipboard.readText().then(paste).catch(()=>paste(copy));else paste(copy);return}
  const a=start||{r,c},b=finish||{r,c};copy=Array.from({length:Math.abs(b.r-a.r)+1},(_,i)=>Array.from({length:Math.abs(b.c-a.c)+1},(_,j)=>String(s.grid[Math.min(a.r,b.r)+i]?.[Math.min(a.c,b.c)+j]??'')).join('\t')).join('\n');try{navigator.clipboard?.writeText(copy).catch(()=>{})}catch{}if(k.toLowerCase()==='x'){for(let y=Math.min(a.r,b.r);y<=Math.max(a.r,b.r);y++)for(let x=Math.min(a.c,b.c);x<=Math.max(a.c,b.c);x++)s.grid[y][x]='';window.ethanSave?.();window.drawGrid?.(s);set(r,c)}return;
 }
 const keys=['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End','PageUp','PageDown'];if(!keys.includes(k))return;
 e.preventDefault();e.stopImmediatePropagation();
 let nr=r,nc=c,extend=e.shiftKey;
 if(k==='ArrowUp')nr--;if(k==='ArrowDown')nr++;if(k==='ArrowLeft')nc--;if(k==='ArrowRight')nc++;
 if(k==='Enter')nr+=e.shiftKey?-1:1;if(k==='Tab')nc+=e.shiftKey?-1:1;
 if(k==='Home'){nc=0;if(mod)nr=0}if(k==='End'){nc=(s.grid?.[0]?.length||20)-1;if(mod)nr=(s.grid?.length||60)-1}
 if(k==='PageUp')nr-=28;if(k==='PageDown')nr+=28;
 if(extend){start=start||{r,c};finish={r:nr,c:nc}}else{start={r:nr,c:nc};finish={r:nr,c:nc}}
 set(nr,nc);
}
document.addEventListener('keydown',onKey,true);
document.addEventListener('click',e=>{const td=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(!td||!isOpen())return;const p={r:+td.dataset.r,c:+td.dataset.c};if(!e.shiftKey){start=p;finish=p}else{finish=p}},true);
window.addEventListener('focus',()=>{if(!isOpen())return;const p=getPos();getCell(p.r,p.c)?.focus({preventScroll:true)},true);
})();