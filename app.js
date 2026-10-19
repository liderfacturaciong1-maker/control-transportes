/* Liquidación Financiera — retraso por días calendario, hora Colombia.
 * Instalación: <script src="./retraso-liquidacion.js" defer></script>
 * No modifica la base de datos ni las otras columnas.
 */
(() => {
  'use strict';
  const STYLE_ID = 'lf-retraso-v2-style';
  const TARGET = ['retraso en liquidacion', 'retraso liquidacion'];
  const normalize = (s) => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const colombiaToday = () => {
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/Bogota', year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(new Date());
    const map = Object.fromEntries(parts.map(x => [x.type, x.value]));
    return Date.UTC(Number(map.year), Number(map.month) - 1, Number(map.day));
  };
  function parseDate(raw) {
    const value = String(raw || '').trim();
    let match = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s|$)/);
    let d, m, y;
    if (match) { [, d, m, y] = match.map((v, i) => i ? Number(v) : v); }
    else {
      match = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]|$)/);
      if (!match) return null;
      y = Number(match[1]); m = Number(match[2]); d = Number(match[3]);
    }
    if (!Number.isInteger(y) || y < 1900 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const utc = Date.UTC(y, m - 1, d);
    const date = new Date(utc);
    if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return null;
    return utc;
  }
  function hoursSince(dateText) {
    const utc = parseDate(dateText);
    return utc == null ? null : Math.max(0, Math.round((colombiaToday() - utc) / 86400000) * 24);
  }
  function injectStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const el = document.createElement('style'); el.id = STYLE_ID;
    el.textContent = `
      td.lf-retraso-col { vertical-align:middle !important; padding:4px 7px !important; }
      .lf-retraso-badge { box-sizing:border-box; display:inline-flex; align-items:center; justify-content:center;
        min-width:84px; min-height:30px; padding:4px 8px; border-radius:7px; font-size:13px;
        line-height:1.15; font-weight:700; white-space:nowrap; color:#34495e; background:#edf3f8;
        border:1px solid #c8d9e7; }
      .lf-retraso-badge.lf-alerta { color:#981b28; background:#ffe1e5; border-color:#e85b69;
        animation: lf-parpadeo 1.8s ease-in-out infinite; }
      @keyframes lf-parpadeo { 0%,100% { background:#ffe1e5; border-color:#e85b69; }
        50% { background:#ffc0c9; border-color:#db354b; } }
      @media (prefers-reduced-motion:reduce) { .lf-retraso-badge.lf-alerta { animation:none; } }
    `;
    (document.head || document.documentElement).appendChild(el);
  }
  function update() {
    document.querySelectorAll('table').forEach(table => {
      const rows = Array.from(table.rows || []);
      if (!rows.length) return;
      const header = rows.find(row => Array.from(row.cells).some(c => TARGET.includes(normalize(c.textContent))));
      if (!header) return;
      const labels = Array.from(header.cells).map(c => normalize(c.textContent));
      const dateIndex = labels.findIndex(x => x === 'fecha');
      const hoursIndex = labels.findIndex(x => TARGET.includes(x));
      if (dateIndex < 0 || hoursIndex < 0) return;
      const dateIsFirst = dateIndex === 0;
      if (!dateIsFirst) return;
      rows.slice(rows.indexOf(header) + 1).forEach(row => {
        if (!row.cells || row.cells.length <= Math.max(dateIndex, hoursIndex)) return;
        const dateText = row.cells[dateIndex].textContent;
        const hours = hoursSince(dateText);
        if (hours == null) return;
        const cell = row.cells[hoursIndex];
        const label = `${hours} horas`;
        const existing = cell.querySelector('.lf-retraso-badge');
        if (existing && existing.textContent === label && existing.classList.contains('lf-alerta') === (hours >= 48)) return;
        cell.classList.add('lf-retraso-col');
        // This is an output-only table cell: rebuild only this last display cell.
        const badge = document.createElement('span');
        badge.className = 'lf-retraso-badge' + (hours >= 48 ? ' lf-alerta' : '');
        badge.textContent = label;
        badge.setAttribute('aria-label', `Retraso en liquidación: ${label}`);
        cell.replaceChildren(badge);
      });
    });
  }
  let scheduled = false;
  function scheduleUpdate() {
    if (scheduled) return;
    scheduled = true;
    setTimeout(() => { scheduled = false; update(); }, 50);
  }
  function start() {
    injectStyle();
    update();
    const observer = new MutationObserver(scheduleUpdate);
    observer.observe(document.body, { subtree:true, childList:true, characterData:true });
    setInterval(update, 60000);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) scheduleUpdate(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, {once:true});
  else start();
})();
