/* Ethan Spreadsheet interaction engine v7 */
(function(){
'use strict';
if(window.__ethanSpreadsheetEngineV7)return;
window.__ethanSpreadsheetEngineV7=true;
const $=s=>document.querySelector(s);
const sheet=()=>{const st=window.ethanGetState?.(),id=window.ethanGetCurrentSheet?.();return st?.sheets?.find(x=>x.id===id)};
const grid=()=>$('#sheetGrid');
const cell=(r,c)=>grid()?.querySelector('td[data-r="'+r+'"][data-c="'+c+'"]');
const pos=()=>window.getXlCell?.()||{r:0,c:0};
const open=()=>document.querySelector('.page.active')?.id==='page-excel'&&!$('#sheetWorkspace')?.classList.contains('hidden');
const cols='ABCDEFGHIJKLMNOPQRST';
let anchor=null,edge=null,clipboard='',undo=[],redo=[],lastFocused=null,editStart=null;
const key=()=>{const s=sheet();return s?.id||'none'};
function bounds(){const p=pos(),a=anchor||p,b=edge||p;return {r1:Math.min(a.r,b.r),r2:Math.max(a.r,b.r),c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c)}}
function snapshot(){const s=sheet();return s?JSON.stringify(s.grid):null}
function pushUndo(before){if(!before)return;const stack=undo.filter(x=>x.id===key());stack.push({id:key(),grid:before});undo=undo.filter(x=>x.id!==key()).concat(stack).slice(-80);redo=redo.filter(x=>x.id!==key())}
function save(){window.ethanSave?.();const s=sheet();if(s){s.updated=Date.now();if(Array.isArray(s.worksheets)){const w=s.worksheets.find(x=>x.id===s.activeWorksheetId);if(w){w.grid=s.grid;w.styles=s.styles||{};w.formats=s.formats||{};w.rules=s.rules||[]}}}}
function render(){const s=sheet();if(!s)return;window.drawGrid?.(s)}
function markRange(){const g=grid();if(!g)return;g.querySelectorAll('.rangeSelected').forEach(x=>x.classList.remove('rangeSelected'));g.querySelectorAll('.sel').forEach(x=>x.classList.remove('sel'));const b=bounds();for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++){const t=cell(r,c);if(t){if(r===pos().r&&c===pos().c)t.classList.add('sel');else t.classList.add('rangeSelected')}}}
function setActive(r,c,extend=false){const s=sheet();if(!s)return;const rows=s.grid?.length||60,columns=s.grid?.[0]?.length||20;r=Math.max(0,Math.min(rows-1,r));c=Math.max(0,Math.min(columns-1,c));if(!extend){anchor={r,c};edge={r,c}}else{anchor=anchor||pos();edge={r,c}}const t=cell(r,c);if(!t)return;t.focus({preventScroll:true});const p=pos();p.r=r;p.c=c;$('#cellName')&&($('#cellName').textContent=cols[c]+(r+1));$('#formulaBar')&&($('#formulaBar').value=s.grid[r]?.[c]??'');markRange();$('#excelStatus')&&($('#excelStatus').textContent=cols[c]+(r+1)+' • Ready')}
function commit(t){if(!t)return;const s=sheet();if(!s)return;const r=+t.dataset.r,c=+t.dataset.c;if(!Number.isInteger(r)||!Number.isInteger(c))return;const value=t.innerText.replace(/\u00a0/g,' ').trim();const old=String(s.grid[r]?.[c]??'');const original=editStart?.cell===t?editStart.text:null;if(original!==null&&value===original)return;if(old.startsWith('=')&&original===null)return;if(value!==old){pushUndo(snapshot());s.grid[r][c]=value;save();$('#formulaBar')&&($('#formulaBar').value=value)}}
function selectedValues(){const s=sheet(),b=bounds(),out=[];for(let r=b.r1;r<=b.r2;r++){const row=[];for(let c=b.c1;c<=b.c2;c++)row.push(String(s?.grid?.[r]?.[c]??''));out.push(row)}return out}
function copyRange(cut){const s=sheet();if(!s)return;const vals=selectedValues();clipboard=vals.map(r=>r.join('\t')).join('\n');try{navigator.clipboard?.writeText(clipboard).catch(()=>{})}catch{}if(cut){const before=snapshot(),b=bounds();for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)s.grid[r][c]='';pushUndo(before);save();render();setActive(b.r1,b.c1,true)}}
function pasteText(text){const s=sheet();if(!s)return;const start=pos(),rows=String(text??'').replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));if(!rows.length)return;const before=snapshot();rows.forEach((row,i)=>row.forEach((v,j)=>{const r=start.r+i,c=start.c+j;if(s.grid[r]&&c<(s.grid[r].length||20))s.grid[r][c]=v}));pushUndo(before);save();render();setActive(start.r,start.c)}
function restore(stack,other){const s=sheet();if(!s)return;const i=stack.map(x=>x.id).lastIndexOf(key());if(i<0)return;const current=snapshot(),entry=stack.splice(i,1)[0];other.push({id:key(),grid:current});s.grid=JSON.parse(entry.grid);save();render();setActive(pos().r,pos().c)}
function addToolbar(){const host=$('#excelRibbonBody');if(!host||$('#ethanSpreadsheetCoreTools'))return;const el=document.createElement('div');el.id='ethanSpreadsheetCoreTools';el.className='eo-tools';el.innerHTML='<div><b>EDITING</b><button type="button" data-action="undo">↶ Undo</button><button type="button" data-action="redo">↷ Redo</button><button type="button" data-action="copy">Copy</button><button type="button" data-action="cut">Cut</button><button type="button" data-action="paste">Paste</button><button type="button" data-action="clear">Clear cells</button></div><div><b>SELECTION</b><button type="button" data-action="all">Select all</button><button type="button" data-action="top">Go to A1</button><button type="button" data-action="end">Go to last cell</button></div>';host.appendChild(el);el.addEventListener('click',e=>{const b=e.target.closest('button[data-action]');if(!b)return;const a=b.dataset.action;if(a==='undo')restore(undo,redo);if(a==='redo')restore(redo,undo);if(a==='copy')copyRange(false);if(a==='cut')copyRange(true);if(a==='paste'){if(navigator.clipboard?.readText)navigator.clipboard.readText().then(pasteText).catch(()=>pasteText(clipboard));else pasteText(clipboard)}if(a==='clear'){const s=sheet(),before=snapshot(),q=bounds();for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++)s.grid[r][c]='';pushUndo(before);save();render();setActive(q.r1,q.c1)}if(a==='all'){anchor={r:0,c:0};edge={r:59,c:19};markRange()}if(a==='top')setActive(0,0);if(a==='end'){const s=sheet();let r=0,c=0;for(let y=0;y<s.grid.length;y++)for(let x=0;x<s.grid[y].length;x++)if(String(s.grid[y][x]??'').trim()){r=y;c=x}setActive(r,c)}})}
document.addEventListener('focusin',e=>{const t=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(!t)return;lastFocused=t;editStart={cell:t,text:t.innerText};const p={r:+t.dataset.r,c:+t.dataset.c};if(!anchor){anchor=p;edge=p}const s=sheet();if(s){const v=s.grid[p.r]?.[p.c]??'';$('#cellName')&&($('#cellName').textContent=cols[p.c]+(p.r+1));$('#formulaBar')&&($('#formulaBar').value=v)}},true);
document.addEventListener('input',e=>{const t=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(t&&$('#formulaBar'))$('#formulaBar').value=t.innerText},true);
document.addEventListener('focusout',e=>{const t=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(t)commit(t)},true);
document.addEventListener('click',e=>{const t=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(!t||!open())return;const p={r:+t.dataset.r,c:+t.dataset.c};if(e.shiftKey){anchor=anchor||pos();edge=p}else{anchor=p;edge=p}setTimeout(markRange,0)},true);
document.addEventListener('keydown',e=>{
 if(!open()||e.altKey||e.isComposing)return;
 const t=e.target?.closest?.('#sheetGrid td[data-r][data-c]')||document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]')||grid()?.querySelector('td.sel[data-r][data-c]');
 if(!t)return;
 const r=+t.dataset.r,c=+t.dataset.c,k=e.key,mod=e.ctrlKey||e.metaKey;
 if(mod&&k.toLowerCase()==='z'){e.preventDefault();e.stopImmediatePropagation();restore(undo,redo);return}
 if(mod&&k.toLowerCase()==='y'){e.preventDefault();e.stopImmediatePropagation();restore(redo,undo);return}
 if(mod&&k.toLowerCase()==='c'){e.preventDefault();e.stopImmediatePropagation();copyRange(false);return}
 if(mod&&k.toLowerCase()==='x'){e.preventDefault();e.stopImmediatePropagation();copyRange(true);return}
 if(mod&&k.toLowerCase()==='v'){e.preventDefault();e.stopImmediatePropagation();if(navigator.clipboard?.readText)navigator.clipboard.readText().then(pasteText).catch(()=>pasteText(clipboard));else pasteText(clipboard);return}
 if(k==='Delete'||k==='Backspace'){e.preventDefault();e.stopImmediatePropagation();const s=sheet(),before=snapshot(),b=bounds();for(let y=b.r1;y<=b.r2;y++)for(let x=b.c1;x<=b.c2;x++)s.grid[y][x]='';pushUndo(before);save();render();setActive(r,c);return}
 if(k==='F2'){e.preventDefault();e.stopImmediatePropagation();t.focus();const range=document.createRange();range.selectNodeContents(t);range.collapse(false);window.getSelection()?.removeAllRanges();window.getSelection()?.addRange(range);return}
 const nav=['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End','PageUp','PageDown'];if(!nav.includes(k))return;
 e.preventDefault();e.stopImmediatePropagation();commit(t);
 let nr=r,nc=c,extend=e.shiftKey;
 if(k==='ArrowUp')nr--;if(k==='ArrowDown')nr++;if(k==='ArrowLeft')nc--;if(k==='ArrowRight')nc++;
 if(k==='Enter')nr+=extend?-1:1;if(k==='Tab')nc+=extend?-1:1;
 if(k==='Home'){nc=0;if(mod)nr=0}if(k==='End'){nc=19;if(mod)nr=59}
 if(k==='PageUp')nr-=25;if(k==='PageDown')nr+=25;
 setActive(nr,nc,extend);
},true);
window.addEventListener('beforeunload',()=>{const t=lastFocused;if(t)commit(t)});
const oldApply=window.applyFormula;
if(typeof oldApply==='function')window.applyFormula=function(){const s=sheet();if(!s)return;const before=snapshot();oldApply();if(before!==snapshot())pushUndo(before)};
document.addEventListener('DOMContentLoaded',addToolbar);setTimeout(addToolbar,0);
const oldDraw=window.drawGrid;if(typeof oldDraw==='function'){window.drawGrid=function(s){oldDraw(s);addToolbar()}}
})();
