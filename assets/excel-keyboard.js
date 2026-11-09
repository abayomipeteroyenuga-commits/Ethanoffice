/* Ethan Office Excel keyboard navigation v11 — authoritative capture-phase navigation */
(function () {
  'use strict';
  if (window.__ethanExcelKeyboardV11) return;
  window.__ethanExcelKeyboardV11 = true;

  const $ = (s) => document.querySelector(s);
  const grid = () => $('#sheetGrid');
  const activePage = () => document.querySelector('.page.active');
  const isExcelOpen = () => { const w = $('#sheetWorkspace'); return !!grid() && !!w && !w.classList.contains('hidden') && w.offsetParent !== null; };
  const getCell = (r, c) => grid()?.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
  const getPos = () => {
    const p = window.getXlCell?.();
    return p && Number.isInteger(p.r) && Number.isInteger(p.c) ? p : { r: 0, c: 0 };
  };
  let anchor = null;
  let edge = null;
  let clipboard = '';

  function range() {
    const p = getPos(), a = anchor || p, b = edge || p;
    return {
      top: Math.min(a.r, b.r), bottom: Math.max(a.r, b.r),
      left: Math.min(a.c, b.c), right: Math.max(a.c, b.c)
    };
  }
  function paintRange() {
    const g = grid();
    if (!g) return;
    g.querySelectorAll('.rangeSelected').forEach((x) => x.classList.remove('rangeSelected'));
    const p = getPos(), b = range();
    for (let r = b.top; r <= b.bottom; r++) {
      for (let c = b.left; c <= b.right; c++) {
        const td = getCell(r, c);
        if (td && (r !== p.r || c !== p.c)) td.classList.add('rangeSelected');
      }
    }
  }
  function activate(r, c, extend) {
    const g = grid(); if (!g) return;
    const rows = g.querySelectorAll('tbody tr').length || 60, cols = g.querySelectorAll('tbody tr:first-child td[data-c]').length || 20;
    r = Math.max(0, Math.min(rows - 1, r));
    c = Math.max(0, Math.min(cols - 1, c));
    const p = getPos();
    // Update the actual shared cell cursor before focusing; getXlCell returns the live cursor.
    p.r = r; p.c = c;
    if (!extend || !anchor) anchor = { r, c };
    edge = { r, c };
    const td = getCell(r, c);
    if (!td) return;
    td.focus({ preventScroll: true });
    // Reassert after focus handlers, which may update the cursor.
    const live = getPos(); live.r = r; live.c = c;
    gScroll(td);
    $('#cellName') && ($('#cellName').textContent = 'ABCDEFGHIJKLMNOPQRST'[c] + (r + 1));
    $('#formulaBar') && ($('#formulaBar').value = td.innerText ?? '');
    $('#excelStatus') && ($('#excelStatus').textContent = 'ABCDEFGHIJKLMNOPQRST'[c] + (r + 1) + ' • Ready');
    paintRange();
  }
  function gScroll(td) {
    const g = grid();
    const scroller = g?.closest('.sheet-wrap, .table-wrap, .excel-grid-wrap') || g?.parentElement;
    if (!scroller || scroller === document.body) {
      td.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
      return;
    }
    const a = td.getBoundingClientRect(), b = scroller.getBoundingClientRect();
    if (a.top < b.top) scroller.scrollTop -= b.top - a.top;
    else if (a.bottom > b.bottom) scroller.scrollTop += a.bottom - b.bottom;
    if (a.left < b.left) scroller.scrollLeft -= b.left - a.left;
    else if (a.right > b.right) scroller.scrollLeft += a.right - b.right;
  }
  function values() {
    const b = range(), out = [];
    for (let r=b.top;r<=b.bottom;r++){const row=[];for(let c=b.left;c<=b.right;c++)row.push(getCell(r,c)?.innerText??'');out.push(row)}
    return out;
  }
  function copy(cut) {
    const b = range(), rows = values();
    clipboard = rows.map((r) => r.join('\t')).join('\n');
    try { navigator.clipboard?.writeText(clipboard).catch(() => {}); } catch (_) {}
    if (cut) {
      for(let r=b.top;r<=b.bottom;r++)for(let c=b.left;c<=b.right;c++){const t=getCell(r,c);if(t)t.innerText=''}
      activate(b.top,b.left,false);
    }
  }
  function paste(text) {
    const p = getPos();
    const rows = String(text ?? '').replace(/\r/g, '').replace(/\n$/, '').split('\n').map((x) => x.split('\t'));
    rows.forEach((row, ri) => row.forEach((v, ci) => {
      const t=getCell(p.r+ri,p.c+ci);if(t)t.innerText=v;
    }));
    activate(p.r,p.c,false);
  }
  document.addEventListener('focusin', (e) => {
    const td = e.target?.closest?.('#sheetGrid td[data-r][data-c]');
    if (!td) return;
    const p = getPos();
    p.r = Number(td.dataset.r); p.c = Number(td.dataset.c);
    if (!anchor) anchor = { r: p.r, c: p.c };
    edge = { r: p.r, c: p.c };
  }, true);
  document.addEventListener('click', (e) => {
    const td = e.target?.closest?.('#sheetGrid td[data-r][data-c]');
    if (!td || !isExcelOpen()) return;
    const p = { r: Number(td.dataset.r), c: Number(td.dataset.c) };
    if (e.shiftKey) edge = p;
    else { anchor = p; edge = p; }
    paintRange();
  }, true);
  document.addEventListener('keydown', (e) => {
    if (!isExcelOpen() || e.altKey || e.isComposing) return;
    const td = e.target?.closest?.('#sheetGrid td[data-r][data-c]') ||
      document.activeElement?.closest?.('#sheetGrid td[data-r][data-c]') ||
      grid()?.querySelector('td.sel[data-r][data-c]');
    if (!td) return;
    const r = Number(td.dataset.r), c = Number(td.dataset.c), k = e.key;
    const mod = e.ctrlKey || e.metaKey;
    const stop = () => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); };
    if (mod && k.toLowerCase() === 'c') { stop(); copy(false); return; }
    if (mod && k.toLowerCase() === 'x') { stop(); copy(true); return; }
    if (mod && k.toLowerCase() === 'v') {
      stop();
      if (navigator.clipboard?.readText) navigator.clipboard.readText().then(paste).catch(() => paste(clipboard));
      else paste(clipboard);
      return;
    }
    if (k === 'Delete' || k === 'Backspace') {
      stop();
      const b=range();for(let y=b.top;y<=b.bottom;y++)for(let x=b.left;x<=b.right;x++){const t=getCell(y,x);if(t)t.innerText=''}activate(r,c,false);
      return;
    }
    if (mod && k.toLowerCase() === 'a') {
      stop(); anchor = { r: 0, c: 0 }; edge = { r: 59, c: 19 }; paintRange(); return;
    }
    const nav = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Enter','Tab','Home','End','PageUp','PageDown'];
    if (!nav.includes(k)) return;
    stop();
    let nr = r, nc = c;
    const shift = e.shiftKey;
    if (k === 'ArrowUp') nr--;
    if (k === 'ArrowDown') nr++;
    if (k === 'ArrowLeft') nc--;
    if (k === 'ArrowRight') nc++;
    if (k === 'Enter') nr += shift ? -1 : 1;
    if (k === 'Tab') nc += shift ? -1 : 1;
    if (k === 'Home') { nc = 0; if (mod) nr = 0; }
    if (k === 'End') { nc = 19; if (mod) nr = 59; }
    if (k === 'PageUp') nr -= 20;
    if (k === 'PageDown') nr += 20;
    activate(nr, nc, shift);
  }, true);

  // Rebind after the application replaces the grid DOM.
  const observer = new MutationObserver(() => {
    const g = grid();
    if (g) g.querySelectorAll('td[data-r][data-c]').forEach((td) => {
      td.setAttribute('aria-label', 'Cell ' + 'ABCDEFGHIJKLMNOPQRST'[Number(td.dataset.c)] + (Number(td.dataset.r) + 1));
    });
  });
  if (document.documentElement) observer.observe(document.documentElement, { childList: true, subtree: true });
})();