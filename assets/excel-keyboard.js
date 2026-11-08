/* Ethan Excel keyboard engine v4 — unified grid navigation */
(function(){
'use strict';
if(window.__ethanExcelKeyboardV4)return;window.__ethanExcelKeyboardV4=true;
const $=s=>document.querySelector(s),clamp=(n,a,b)=>Math.max(a,Math.min(b,n));
const state=()=>{try{return window.ethanGetState?.()}catch{return null}};
const sheet=()=>{const s=state(),id=window.ethanGetCurrentSheet?.();return s?.sheets?.find(x=>x.id===id)||null};
const pos=()=>{try{return window.getXlCell?.()||{r:0,c:0}}catch{return{r:0,c:0}}};
const grid=()=>$('#sheetGrid'),cell=(r,c)=>grid()?.querySelector('td[data-r="'+r+'"][data-c="'+c+'"]');
const active=()=>grid()?.querySelector('td.sel[data-r][data-c]')||document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]')||null;
let anchor=null,edge=null,clipboard='';
const xy=el=>({r:+el.dataset.r,c:+el.dataset.c});
const lim=()=>{const s=sheet();return{r:Math.max(0,(s?.grid?.length||60)-1),c:Math.max(0,(s?.grid?.[0]?.length||20)-1)}};
function setPos(r,c){const p=pos();p.r=r;p.c=c}
function commit(el){if(!el)return;const s=sheet();if(!s)return;const p=xy(el);if(el.dataset.raw!==undefined){s.grid[p.r][p.c]=el.dataset.raw;delete el.dataset.raw}else if(String(s.grid[p.r]?.[p.c]??'').startsWith('=')){el.dataset.raw=String(s.grid[p.r][p.c])}}
function paint(){const g=grid();if(!g)return;g.querySelectorAll('.sel,.rangeSelected').forEach(x=>x.classList.remove('sel','rangeSelected'));const p=pos(),a=anchor||p,z=edge||p;
for(let r=Math.min(a.r,z.r);r<=Math.max(a.r,z.r);r++)for(let c=Math.min(a.c,z.c);c<=Math.max(a.c,z.c);c++){const el=cell(r,c);if(el)el.classList.add(r===p.r&&c===p.c?'sel':'rangeSelected')}
const name=$('#cellName');if(name)name.textContent=String.fromCharCode(65+p.c)+(p.r+1);const s=sheet(),f=$('#formulaBar');if(f&&document.activeElement!==f)f.value=s?.grid?.[p.r]?.[p.c]??''}
function select(r,c,extend,old){const l=lim();r=clamp(r,0,l.r);c=clamp(c,0,l.c);commit(old);const p=pos();if(extend){anchor=anchor||{r:old?+old.dataset.r:p.r,c:old?+old.dataset.c:p.c};edge={r,c}}else{anchor={r,c};edge={r,c}}setPos(r,c);const el=cell(r,c);if(el){el.tabIndex=0;el.focus({preventScroll:true})}paint()}
function visible(){return !!$('#sheetWorkspace')&&!$('#sheetWorkspace').classList.contains('hidden')&&!!grid()?.querySelector('td[data-r][data-c]')}
function onKey(e){if(!visible()||e.isComposing||e.altKey)return;const el=active();if(!el)return;const k=e.key,mod=e.ctrlKey||e.metaKey,s=sheet();
if(mod&&['c','x','v'].includes(k.toLowerCase())&&s){const p=pos(),op=k.toLowerCase();e.preventDefault();e.stopImmediatePropagation();if(op==='v'){const put=t=>{const rows=String(t||'').replace(/\r/g,'').replace(/\n$/,'').split('\n').map(x=>x.split('\t'));rows.forEach((row,i)=>row.forEach((v,j)=>{if(s.grid[p.r+i]&&p.c+j<20)s.grid[p.r+i][p.c+j]=v}));window.ethanSave?.();window.drawGrid?.(s);select(p.r,p.c,false,null)};if(navigator.clipboard?.readText)navigator.clipboard.readText().then(put).catch(()=>put(clipboard));else put(clipboard);return}
const a=anchor||p,z=edge||p,r1=Math.min(a.r,z.r),r2=Math.max(a.r,z.r),c1=Math.min(a.c,z.c),c2=Math.max(a.c,z.c);clipboard=Array.from({length:r2-r1+1},(_,i)=>Array.from({length:c2-c1+1},(_,j)=>String(s.grid[r1+i]?.[c1+j]??'')).join('\t')).join('\n');try{navigator.clipboard?.writeText(clipboard).catch(()=>{})}catch{}if(op==='x'){for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)s.grid[r][c]='';window.ethanSave?.();window.drawGrid?.(s);select(r1,c1,false,null)}return}
if(!['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End','PageUp','PageDown'].includes(k))return;
e.preventDefault();e.stopImmediatePropagation();const p=xy(el);let r=p.r,c=p.c;if(k==='ArrowUp')r--;if(k==='ArrowDown')r++;if(k==='ArrowLeft')c--;if(k==='ArrowRight')c++;if(k==='Enter')r+=e.shiftKey?-1:1;if(k==='Tab')c+=e.shiftKey?-1:1;if(k==='Home'){c=0;if(mod)r=0}if(k==='End'){c=lim().c;if(mod)r=lim().r}if(k==='PageUp')r-=28;if(k==='PageDown')r+=28;select(r,c,e.shiftKey,el)}
document.addEventListener('keydown',onKey,true);
function bind(){const g=grid();if(!g)return;g.querySelectorAll('td[data-r][data-c]').forEach(x=>x.tabIndex=0);if(g.dataset.ethanKeyboardV4==='1')return;g.dataset.ethanKeyboardV4='1';g.addEventListener('pointerdown',e=>{const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;const p=xy(el);if(e.shiftKey){anchor=anchor||pos();edge=p}else{anchor=p;edge=p}setPos(p.r,p.c);paint()},true);g.addEventListener('click',e=>{const el=e.target.closest?.('td[data-r][data-c]');if(!el)return;const p=xy(el);if(!e.shiftKey){anchor=p;edge=p}else edge=p;setPos(p.r,p.c);paint()},true)}
const draw=window.drawGrid;if(typeof draw==='function')window.drawGrid=function(s){const out=draw(s);bind();const p=pos();setTimeout(()=>{bind();if(!grid()?.contains(document.activeElement))cell(p.r,p.c)?.focus({preventScroll:true});paint()},0);return out};
window.xlAutoSum=function(){const s=sheet();if(!s)return;const p=pos(),col=String.fromCharCode(65+p.c),row=Math.min(lim().r,p.r+1);s.grid[row][p.c]='=SUM('+col+(Math.max(1,p.r))+':'+col+(p.r+1)+')';window.ethanSave?.();window.drawGrid?.(s);select(row,p.c,false,null)};
window.ethanSelectRange=function(v){const m=/^([A-T])(\d+):([A-T])(\d+)$/i.exec(String(v||''));if(!m)return false;anchor={r:+m[2]-1,c:m[1].toUpperCase().charCodeAt(0)-65};edge={r:+m[4]-1,c:m[3].toUpperCase().charCodeAt(0)-65};select(edge.r,edge.c,true,null);return true};
document.addEventListener('DOMContentLoaded',bind);setTimeout(bind,0);
})();