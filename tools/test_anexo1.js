// Prueba de apps-script-anexo1/POI_Codigo.gs con un Sheet simulado.  Uso: node tools/test_anexo1.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const dir = path.join(__dirname, '..', 'apps-script-anexo1');

class Range {
  constructor(sh, r, c, nr = 1, nc = 1) { Object.assign(this, { sh, r, c, nr, nc }); }
  getSheet() { return this.sh; } getRow() { return this.r; } getColumn() { return this.c; }
  getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getValue() { return this.sh.v(this.r, this.c); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.v(this.r + i, this.c + j)); o.push(row); } return o; }
  setValue(x) { this.sh.cells[this.r + ',' + this.c] = x; return this; }
  setValues(a) { a.forEach((row, i) => row.forEach((x, j) => this.sh.cells[(this.r + i) + ',' + (this.c + j)] = x)); return this; }
  clearContent() { delete this.sh.cells[this.r + ',' + this.c]; return this; }
  setDataValidation(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.val[(this.r + i) + ',' + (this.c + j)] = v; return this; }
  getDataValidation() { return this.sh.val[this.r + ',' + this.c] || null; }
  getDataValidations() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.val[(this.r + i) + ',' + (this.c + j)] || null); o.push(row); } return o; }
  setFontWeight() { return this; }
}
class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.cells = {}; this.val = {}; this.rows = 1000; }
  v(r, c) { const x = this.cells[r + ',' + c]; return x === undefined ? '' : x; }
  getName() { return this.name; } getParent() { return this.ss; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
  getMaxRows() { return this.rows; } insertRowsAfter(_, n) { this.rows += n; }
  setColumnWidth() {} hideSheet() { this.hidden = true; } clear() { this.cells = {}; }
  setFrozenRows() {} autoResizeColumns() {}
  allowed(r, c) { const v = this.val[r + ',' + c]; return v ? v.range.getValues().map(x => x[0]) : null; }
}
class SS {
  constructor() { this.sheets = []; }
  getSheets() { return this.sheets; } getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  toast() {}
}
const ss = new SS();
const triggers = [];
const builder = () => { const o = {}; const b = { requireValueInRange(r) { o.range = r; return b; }, setAllowInvalid() { return b; }, setHelpText() { return b; }, build() { o.getCriteriaValues = () => [o.range, true]; return o; } }; return b; };
const menus = [];
const ui = { createMenu: n => { const m = { n, items: [], addItem(l, f) { m.items.push(f); return m; }, addSeparator() { return m; }, addToUi() { menus.push(m); } }; return m; }, alert() {}, ButtonSet: { OK: 0 } };
const ctx = {
  SpreadsheetApp: { newDataValidation: builder, getActive: () => ss, getUi: () => ui },
  ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
    newTrigger: fn => { const t = { fn, getHandlerFunction: () => fn }; const b = { forSpreadsheet() { return b; }, onEdit() { t.ev = 'edit'; return b; }, onOpen() { t.ev = 'open'; return b; }, create() { triggers.push(t); return t; } }; return b; } },
  console };
vm.createContext(ctx);
const code = ['POI_Datos.gs', 'POI_Codigo.gs'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') +
  '\n;this.__x = { POI_FACULTADES, POI_AEI, POI_CFG, poiFacultadDeHoja_, poiAplicarHoja_, poiAlEditar, poiInstalar, poiHojas_ };';
vm.runInContext(code, ctx);
const X = ctx.__x;

// 0) no define disparadores simples ni nombres que puedan chocar con las automatizaciones del libro
assert.strictEqual(vm.runInContext('typeof onEdit + typeof onOpen + typeof CFG + typeof indice_', ctx), 'undefinedundefinedundefinedundefined');
assert.ok(!/OEI/.test(fs.readFileSync(path.join(dir, 'POI_Datos.gs'), 'utf8').replace('Sin OEI', '')), 'Datos sin OEI');

// 1) pestañas reales del libro Anexo 1
const nombres = ['F01_FM','F02_FDCP','F03_FLCH','F04_FFB','F05_FO','F06_FE','F07_FQIQ','F08_FMV','F09_FCA','F10_FCB','Copia de F10_FCB','F11_FCC','Copia de 11. FCC','F12_FCE','F13_FCF','F14_FCM','F15_FCCSS','F16_FIGMMG','F17_FII','F18_FPSIC','F19_FIEE','F20_FISI','MAPA DE PROCESOS'];
const hojas = {}; nombres.forEach(n => hojas[n] = ss.insertSheet(n));
const esperado = { F01_FM: 'FM', F05_FO: 'FO', F09_FCA: 'FCA', F15_FCCSS: 'FCCSS', F17_FII: 'FII', F18_FPSIC: 'FPSIC', F19_FIEE: 'FIEE', F20_FISI: 'FISI', 'Copia de F10_FCB': 'FCB', 'Copia de 11. FCC': 'FCC' };
for (const n in esperado) assert.strictEqual(X.poiFacultadDeHoja_(n)?.sigla, esperado[n], n);
assert.strictEqual(X.poiFacultadDeHoja_('MAPA DE PROCESOS'), null);
assert.strictEqual(X.poiFacultadDeHoja_('POI_LISTAS'), null);
assert.strictEqual(X.poiHojas_(ss).length, 20);                       // las 20 facultades, sin copias ni MAPA

// 2) poiInstalar: crea 2 disparadores instalables, menu, listas y valida SOLO D7/E7
hojas['F09_FCA'].getRange(7, 4).setValue('texto previo del libro');
hojas['F09_FCA'].getRange(8, 4).setValue('AE 04.01 lista propia del libro');
X.poiInstalar();
assert.deepStrictEqual(triggers.map(t => t.fn + ':' + t.ev).sort(), ['poiAlEditar:edit', 'poiCrearMenu:open']);
X.poiInstalar();                                                       // idempotente: no duplica
assert.strictEqual(triggers.length, 2);
assert.strictEqual(menus[0].n, 'POI AEI-AOI');
const fca = hojas['F09_FCA'];
assert.strictEqual(fca.v(7, 4), 'texto previo del libro');            // no toca valores
assert.strictEqual(fca.v(8, 4), 'AE 04.01 lista propia del libro');
assert.strictEqual(fca.allowed(7, 4).length, 12);                      // AEI de FCA
assert.deepStrictEqual(fca.allowed(7, 5), [X.POI_CFG.PH_AOI]);
assert.strictEqual(fca.allowed(8, 4), null); assert.strictEqual(fca.allowed(7, 3), null); assert.strictEqual(fca.allowed(6, 4), null);
assert.strictEqual(hojas['Copia de F10_FCB'].allowed(7, 4), null);     // copias omitidas
assert.strictEqual(hojas['MAPA DE PROCESOS'].allowed(7, 4), null);
['F01_FM','F05_FO','F17_FII','F18_FPSIC','F19_FIEE','F20_FISI'].forEach(n => assert.ok(hojas[n].allowed(7, 4).length > 0, n));
assert.ok(ss.getSheetByName('POI_LISTAS').hidden);
fca.clearContent && 0;
delete fca.cells['7,4'];

// helpers
const edit = (sh, r, c, nr = 1, nc = 1) => ctx.poiAlEditar ? vm.runInContext('poiAlEditar', ctx)({ range: sh.getRange(r, c, nr, nc) }) : null;
const copiar = (sh, f, c, nf, nc, f2, c2) => { for (let i = 0; i < nf; i++) for (let j = 0; j < nc; j++) {
  const k = (f + i) + ',' + (c + j), k2 = (f2 + i) + ',' + (c2 + j);
  if (sh.cells[k] !== undefined) sh.cells[k2] = sh.cells[k]; else delete sh.cells[k2];
  if (sh.val[k]) sh.val[k2] = sh.val[k]; else delete sh.val[k2]; } };

// 3) elegir AEI en D7 -> E7 solo con las AOI de ese AEI (codigo - denominacion)
fca.getRange(7, 4).setValue(X.POI_AEI['AEI.01.01']); edit(fca, 7, 4);
let aois = fca.allowed(7, 5);
assert.strictEqual(aois.length, 7);
assert.strictEqual(aois[0], 'AOI00008800777 - EJECUCIÓN DE LOS PROGRAMAS DE POSGRADO');
// 4) elegir AOI; cambiar el AEI -> se limpia y cambia la lista
fca.getRange(7, 5).setValue(aois[0]); edit(fca, 7, 5);
assert.strictEqual(fca.v(7, 5), aois[0]);
fca.getRange(7, 4).setValue(X.POI_AEI['AEI.01.02']); edit(fca, 7, 4);
assert.strictEqual(fca.v(7, 5), ''); assert.strictEqual(fca.allowed(7, 5).length, 5);
// 5) AEI borrado -> vuelve el aviso
fca.getRange(7, 5).setValue(fca.allowed(7, 5)[0]);
delete fca.cells['7,4']; edit(fca, 7, 4);
assert.strictEqual(fca.v(7, 5), ''); assert.deepStrictEqual(fca.allowed(7, 5), [X.POI_CFG.PH_AOI]);

// 6) COPIAR D7:E7 (AEI + AOI) a otra fila y a otras columnas: se conserva si es consistente
fca.getRange(7, 4).setValue(X.POI_AEI['AEI.01.01']); edit(fca, 7, 4);
fca.getRange(7, 5).setValue(fca.allowed(7, 5)[2]); edit(fca, 7, 5);
copiar(fca, 7, 4, 1, 2, 20, 4); edit(fca, 20, 4, 1, 2);
assert.strictEqual(fca.v(20, 5), fca.v(7, 5)); assert.strictEqual(fca.allowed(20, 5).length, 7);
copiar(fca, 7, 4, 1, 2, 30, 9); edit(fca, 30, 9, 1, 2);                // columnas I:J
assert.strictEqual(fca.v(30, 10), fca.v(7, 5));
// 7) COPIAR SOLO el AEI (caso reportado): la celda de la derecha tenia el aviso
delete fca.cells['7,5']; fca.getRange(7, 4).setValue(X.POI_AEI['AEI.01.05']); edit(fca, 7, 4);
copiar(fca, 7, 4, 1, 1, 40, 4);
fca.val['40,5'] = fca.val['7,5'];
fca.val['40,5'] = hojas['F09_FCA'].val['20,5'];                         // vecina con lista de otro AEI
edit(fca, 40, 4);
assert.strictEqual(fca.allowed(40, 5).length, 4, 'la AOI pegada debe listar las actividades del AEI');
// 8) pegar SOLO VALOR (sin validacion) en cualquier columna
fca.getRange(50, 12).setValue(X.POI_AEI['AEI.01.01']); edit(fca, 50, 12);
assert.strictEqual(fca.allowed(50, 12).length, 12); assert.strictEqual(fca.allowed(50, 13).length, 7);
// 9) AOI pegada en fila con otro AEI -> se limpia
fca.getRange(60, 4).setValue(X.POI_AEI['AEI.01.02']); edit(fca, 60, 4);
copiar(fca, 20, 5, 1, 1, 60, 5); edit(fca, 60, 5);
assert.strictEqual(fca.v(60, 5), ''); assert.strictEqual(fca.allowed(60, 5).length, 5);
// 10) texto ajeno a la derecha del AEI pegado: no se borra, pero recibe la lista
fca.getRange(70, 4).setValue(X.POI_AEI['AEI.01.01']); fca.getRange(70, 5).setValue('GESTIÓN DEL PLANEAMIENTO (propio del libro)');
fca.val['70,5'] = { range: { getSheet: () => ({ getName: () => 'OtraLista' }), getRow: () => 1 }, getCriteriaValues() { return [this.range]; } };
edit(fca, 70, 4);
assert.strictEqual(fca.v(70, 5), 'GESTIÓN DEL PLANEAMIENTO (propio del libro)');
assert.strictEqual(fca.allowed(70, 5).length, 7);
// 11) ediciones normales del libro (texto cualquiera, otras hojas) no hacen nada
fca.getRange(9, 2).setValue('Facultad de Ciencias Administrativas'); edit(fca, 9, 2); assert.strictEqual(fca.allowed(9, 2), null);
edit(hojas['MAPA DE PROCESOS'], 7, 4); edit(ss.getSheetByName('POI_LISTAS'), 2, 1);
// 12) copia de trabajo "Copia de F10_FCB": no se aplica en bloque, pero si se pega alli funciona
const cp = hojas['Copia de F10_FCB']; cp.getRange(7, 4).setValue(X.POI_AEI['AEI.01.01']); edit(cp, 7, 4);
assert.ok(cp.allowed(7, 5).length > 0);
console.log('OK: todas las pruebas del script Anexo 1 pasaron');
