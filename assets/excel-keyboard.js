/* Ethan Office Excel keyboard navigation v12 — DOM-first cell navigation */
(function () {
  'use strict';
  if (window.__ethanExcelKeyboardV12) return;
  window.__ethanExcelKeyboardV12 = true;
  const $ = (s) => document.querySelector(s);
  const grid = () => $('#sheetGrid');
  const cell = (r,c) => grid()?.querySelector('td[contenteditable="true"][data-r="'+r+'"][data-c="'+c+'"]') ||
    grid()?.querySelector('td[data-r="'+r+'"][data-c="'+c+'"]');
  const isExcelOpen = () => {
    const w=$('#sheetWorkspace'), p=$('#page-excel');
    return !!(w && grid() && !w.classList.contains('hidden') && (!p || p.classList.contains('active')));
  };
  let anchor=null, edge=null, clipboard='';
  const currentCell = (e) => e.target?.closest?.('#sheetGrid td[data-r][data-c]') ||
    document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]') ||
    grid()?.querySelector('td.sel[data-r][data-c]');
  function bounds(){
    const g=grid(), rows=g?.querySelectorAll('tbody tr').length||60;
    const cols=g?.querySelectorAll('tbody tr:first-child td[data-c]').length||20;
    return {rows,cols};
  }
  function select(td, extend){
    if(!td)return;
    const p={r:+td.dataset.r,c:+td.dataset.c};
    if(!extend||!anchor)anchor=p;
    edge=p;
    // Use the application's own selection handler; it updates its real cursor,
    // formula bar, selected-cell class and status in one operation.
    td.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,shiftKey:!!extend}));
    td.focus({preventScroll:true});
    const scroller=td.closest('.sheet-wrap,.table-wrap,.excel-grid-wrap')||td.parentElement;
    if(scroller?.scrollIntoView)td.scrollIntoView({block:'nearest',inline:'nearest'});
    paint();
  }
  function paint(){
    const g=grid();if(!g)return;
    g.querySelectorAll('.rangeSelected').forEach(x=>x.classList.remove('rangeSelected'));
    if(!anchor||!edge)return;
    const p=window.getXlCell?.()||edge;
    for(let r=Math.min(anchor.r,edge.r);r<=Math.max(anchor.r,edge.r);r++)
      for(let c=Math.min(anchor.c,edge.c);c<=Math.max(anchor.c,edge.c);c++){
        if(r===p.r&&c===p.c)continue;
        cell(r,c)?.classList.add('rangeSelected');
      }
  }
  function selectedRange(){
    const p=window.getXlCell?.()||{r:0,c:0},a=anchor||p,b=edge||p;
    return {top:Math.min(a.r,b.r),bottom:Math.max(a.r,b.r),left:Math.min(a.c,b.c),right:Math.max(a.c,b.c)};
  }
  function copy(cut){
    const b=selectedRange(), rows=[];
    for(let r=b.top;r<=b.bottom;r++){
      const row=[];for(let c=b.left;c<=b.right;c++)row.push(cell(r,c)?.innerText||'');
      rows.push(row);
    }
    clipboard=rows.map(x=>x.join('\t')).join('\n');
    try{navigator.clipboard?.writeText(clipboard).catch(()=>{});}catch(_){}
    if(cut){for(let r=b.top;r<=b.bottom;r++)for(let c=b.left;c<=b.right;c++){const t=cell(r,c);if(t)t.textContent='';}select(cell(b.top,b.left),false);}
  }
  function paste(value){
    const p=window.getXlCell?.()||{r:0,c:0};
    String(value||'').replace(/\r/g,'').replace(/\n$/,'').split('\n').forEach((line,ri)=>
      line.split('\t').forEach((v,ci)=>{const t=cell(p.r+ri,p.c+ci);if(t)t.textContent=v;}));
    select(cell(p.r,p.c),false);
  }
  document.addEventListener('focusin',e=>{
    const td=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(!td)return;
    const p={r:+td.dataset.r,c:+td.dataset.c};if(!anchor)anchor=p;edge=p;
  },true);
  document.addEventListener('keydown',e=>{
    if(!isExcelOpen()||e.altKey||e.isComposing)return;
    const td=currentCell(e);if(!td)return;
    const r=+td.dataset.r,c=+td.dataset.c,k=e.key,mod=e.ctrlKey||e.metaKey;
    const stop=()=>{e.preventDefault();e.stopPropagation();e.stopImmediatePropagation();};
    if(mod&&k.toLowerCase()==='c'){stop();copy(false);return;}
    if(mod&&k.toLowerCase()==='x'){stop();copy(true);return;}
    if(mod&&k.toLowerCase()==='v'){
      stop();if(navigator.clipboard?.readText)navigator.clipboard.readText().then(paste).catch(()=>paste(clipboard));else paste(clipboard);return;
    }
    if(k==='Delete'||k==='Backspace'){
      stop();const b=selectedRange();
      for(let y=b.top;y<=b.bottom;y++)for(let x=b.left;x<=b.right;x++){const t=cell(y,x);if(t)t.textContent='';}
      select(cell(r,c),false);return;
    }
    if(mod&&k.toLowerCase()==='a'){stop();const z=bounds();anchor={r:0,c:0};edge={r:z.rows-1,c:z.cols-1};paint();return;}
    const nav=['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End','PageUp','PageDown'];
    if(!nav.includes(k))return;
    // Let normal arrow keys edit text only when the user is actively selecting text.
    if(td.isContentEditable&&window.getSelection?.()?.toString()&&['ArrowLeft','ArrowRight'].includes(k))return;
    stop();
    let nr=r,nc=c;const shift=e.shiftKey,z=bounds();
    if(k==='ArrowUp')nr--; if(k==='ArrowDown')nr++;
    if(k==='ArrowLeft')nc--; if(k==='ArrowRight')nc++;
    if(k==='Enter')nr+=shift?-1:1;
    if(k==='Tab')nc+=shift?-1:1;
    if(k==='Home'){nc=0;if(mod)nr=0;}
    if(k==='End'){nc=z.cols-1;if(mod)nr=z.rows-1;}
    if(k==='PageUp')nr-=20;if(k==='PageDown')nr+=20;
    nr=Math.max(0,Math.min(z.rows-1,nr));nc=Math.max(0,Math.min(z.cols-1,nc));
    // Commit the edited cell through the app's normal blur handler before moving.
    td.blur();
    select(cell(nr,nc),shift);
  },true);
  document.addEventListener('click',e=>{
    const td=e.target?.closest?.('#sheetGrid td[data-r][data-c]');if(!td)return;
    const p={r:+td.dataset.r,c:+td.dataset.c};
    if(e.shiftKey){if(!anchor)anchor=window.getXlCell?.()||p;edge=p;}
    else{anchor=p;edge=p;}
    paint();
  },true);
})();
