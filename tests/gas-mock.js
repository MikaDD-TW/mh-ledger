// Minimal Google Apps Script mock: SpreadsheetApp / Utilities / LockService / PropertiesService / Logger.
(function (root) {
  const chain = (target) => new Proxy(target, {
    get(t, k) {
      if (k in t) return t[k];
      if (typeof k === 'symbol') return undefined;
      return function () { return chain(t); }; // formatting no-ops stay chainable
    }
  });
  const colToNum = (s) => { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  function parseA1(a1, maxRows, maxCols) {
    const m = a1.match(/^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/);
    if (!m) throw new Error('bad a1 ' + a1);
    const c1 = m[1] ? colToNum(m[1]) : 1, r1 = m[2] ? +m[2] : 1;
    let c2 = c1, r2 = r1;
    if (a1.includes(':')) { c2 = m[3] ? colToNum(m[3]) : maxCols; r2 = m[4] ? +m[4] : maxRows; }
    return { r: r1, c: c1, nr: r2 - r1 + 1, nc: c2 - c1 + 1 };
  }
  // Values the real Sheets would parse from strings typed into a non-text cell.
  function convert(sh, r, c, v) {
    if (typeof v === 'string' && !sh.textCols.has(c) && /^\d{4}-\d{2}-\d{2}$/.test(v) && c === 1 && sh.name === '明細') {
      const [y, mo, d] = v.split('-').map(Number); return new Date(Date.UTC(y, mo - 1, d) - 8 * 3600e3); // midnight Taipei
    }
    return v;
  }
  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh._get(this.r + i, this.c + j)); out.push(row); } return out; }
    getValue() { return this.sh._get(this.r, this.c); }
    setValues(v) {
      if (v.length !== this.nr || v.some((row) => row.length !== this.nc)) throw new Error(`setValues dims ${v.length}x${v[0] && v[0].length} vs ${this.nr}x${this.nc}`);
      v.forEach((row, i) => row.forEach((x, j) => this.sh._set(this.r + i, this.c + j, convert(this.sh, this.r + i, this.c + j, x))));
      return chain(this);
    }
    setValue(x) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh._set(this.r + i, this.c + j, convert(this.sh, this.r + i, this.c + j, x)); return chain(this); }
    setFormula(f) { this.sh._set(this.r, this.c, f); return chain(this); }
    setNumberFormat(f) { if (f === '@') for (let j = 0; j < this.nc; j++) this.sh.textCols.add(this.c + j); return chain(this); }
    getRow() { return this.r; }
  }
  class Sheet {
    constructor(name) { this.name = name; this.cells = []; this.maxRows = 1000; this.maxCols = 26; this.textCols = new Set(); }
    getName() { return this.name; }
    _get(r, c) { const row = this.cells[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined || v === null ? '' : v; }
    _set(r, c, v) { if (r > this.maxRows || c > this.maxCols) throw new Error(`Range outside sheet (${r},${c}) max ${this.maxRows}`); while (this.cells.length < r) this.cells.push([]); this.cells[r - 1][c - 1] = v; }
    getLastRow() { for (let r = this.cells.length; r >= 1; r--) { const row = this.cells[r - 1] || []; if (row.some((v) => v !== '' && v !== undefined && v !== null)) return r; } return 0; }
    getLastColumn() { let m = 0; this.cells.forEach((row) => row.forEach((v, i) => { if (v !== '' && v != null) m = Math.max(m, i + 1); })); return m; }
    getMaxRows() { return this.maxRows; }
    insertRowsAfter(after, n) { this.maxRows += n; return chain(this); }
    deleteRow(r) { this.cells.splice(r - 1, 1); this.maxRows--; return chain(this); }
    getRange(a, b, c, d) {
      if (typeof a === 'string') { const p = parseA1(a, this.maxRows, this.maxCols); return chain(new Range(this, p.r, p.c, p.nr, p.nc)); }
      if (a < 1 || b < 1) throw new Error('bad coords');
      return chain(new Range(this, a, b, c || 1, d || 1));
    }
    getRangeList(list) { const rs = list.map((x) => this.getRange(x)); return chain({ setValue(v) { rs.forEach((r) => r.setValue(v)); return this; } }); }
  }
  class SS {
    constructor() { this.sheets = [chain(new Sheet('工作表1'))]; this.tz = 'America/New_York'; }
    getSheetByName(n) { return this.sheets.find((s) => s.getName() === n) || null; }
    insertSheet(n, idx) { const s = chain(new Sheet(n)); if (idx === undefined) this.sheets.push(s); else this.sheets.splice(idx, 0, s); return s; }
    getSheets() { return this.sheets.slice(); }
    deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
    setSpreadsheetTimeZone(t) { this.tz = t; }
    getSpreadsheetTimeZone() { return this.tz; }
  }
  const ss = chain(new SS());
  const props = {};
  root.__ss = ss;
  root.SpreadsheetApp = {
    getActiveSpreadsheet: () => ss,
    newDataValidation: () => chain({ build: () => ({}) }),
  };
  root.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) };
  root.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  root.Logger = { log: (...a) => (root.__log = (root.__log || []).concat(a)) };
  root.Utilities = {
    getUuid: () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (ch) => { const r = (Math.random() * 16) | 0; return (ch === 'x' ? r : (r & 3) | 8).toString(16); }),
    formatDate(d, tz, fmt) {
      const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(d).map((p) => [p.type, p.value]));
      return fmt.replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day).replace('HH', parts.hour).replace('mm', parts.minute);
    },
  };
  root.HtmlService = { createTemplateFromFile: () => chain({ evaluate: () => chain({}) }) };
})(typeof globalThis !== 'undefined' ? globalThis : window);
