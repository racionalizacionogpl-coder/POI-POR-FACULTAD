// Prueba de la logica de apps-script/Codigo.gs con un Sheet simulado.  Uso: node tools/test_codigo.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const dir = path.join(__dirname, '..', 'apps-script');

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
  setFontWeight() { return this; } setWrap() { return this; } setVerticalAlignment() { return this; }
}
class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.cells = {}; this.val = {}; this.rows = 1000; }
  v(r, c) { const x = this.cells[r + ',' + c]; return x === undefined ? '' : x; }
  getName() { return this.name; } getParent() { return this.ss; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
  getMaxRows() { return this.rows; } insertRowsAfter(_, n) { this.rows += n; }
  setColumnWidth() {} hideSheet() { this.hidden = true; } clear() { this.cells = {}; }
  setFrozenRows() {} autoResizeColumns() {}
  // Valores permitidos de la celda segun su validacion (rango en POI_LISTAS)
  allowed(r, c) { const v = this.val[r + ',' + c]; if (!v) return null; const g = v.range; return g.getValues().map(x => x[0]); }
}
class SS {
  constructor() { this.sheets = []; }
  getSheets() { return this.sheets; } getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
  toast() {}
}
const ss = new SS();
const builder = () => { const o = {}; const b = { requireValueInRange(r) { o.range = r; return b; }, setAllowInvalid() { return b; }, setHelpText() { return b; }, build() { return o; } }; return b; };
const ctx = { SpreadsheetApp: { newDataValidation: builder, getActive: () => ss, getUi: () => ({}) }, console };
vm.createContext(ctx);
const code = ['Datos.gs', 'Codigo.gs'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') +
  '\n;this.__x = { POI_FACULTADES, POI_OEI, POI_AEI, CFG, indice_, facultadDeHoja_, aplicarHoja_, generarListas_, onEdit, hojasDeFacultad_ };';
vm.runInContext(code, ctx);
const X = ctx.__x;

// 1) deteccion de hojas
for (const [n, s] of [['F01_FM', 'FM'], ['F09_FCA', 'FCA'], ['F15_FCCSS', 'FCCSS'], ['F06_FE', 'FE'], ['FPSIC', 'FPSIC'], ['F21_FIEE', 'FIEE'], ['FISI', 'FISI'], ['F20_FII', 'FII']])
  assert.strictEqual(X.facultadDeHoja_(n)?.sigla, s, n);
assert.strictEqual(X.facultadDeHoja_('Anexo 1'), null);
assert.strictEqual(X.facultadDeHoja_('POI_LISTAS'), null);
assert.strictEqual(X.facultadDeHoja_('POI_CATALOGO'), null);

// 2) aplicar a la hoja FCA
const fca = ss.insertSheet('F09_FCA'); const fac = X.facultadDeHoja_('F09_FCA');
const listas = X.generarListas_(ss);
X.aplicarHoja_(fca, fac, listas);
const r0 = X.CFG.FILA_INICIO;
assert.strictEqual(fca.v(r0 - 1, 1), 'OEI');
const oeiList = fca.allowed(r0, 1);
assert.strictEqual(oeiList.length, 4);
assert.deepStrictEqual(fca.allowed(r0, 2), [X.CFG.PH_AEI]);
assert.deepStrictEqual(fca.allowed(r0, 3), [X.CFG.PH_AOI]);

const edit = (r, c) => X.onEdit({ range: fca.getRange(r, c) });
// 3) elegir OEI.01 -> solo AEI.01.xx de FCA
fca.getRange(r0, 1).setValue(X.POI_OEI['OEI.01']); edit(r0, 1);
const aeis = fca.allowed(r0, 2);
assert.deepStrictEqual(aeis.map(a => a.slice(0, 9)), ['AEI.01.01', 'AEI.01.02', 'AEI.01.03', 'AEI.01.04', 'AEI.01.05']);
// 4) elegir AEI.01.01 -> exactamente 7 AOI (segun el PDF), con codigo y denominacion
fca.getRange(r0, 2).setValue(X.POI_AEI['AEI.01.01']); edit(r0, 2);
let aois = fca.allowed(r0, 3);
assert.strictEqual(aois.length, 7);
assert.strictEqual(aois[0], 'AOI00008800777 - EJECUCIÓN DE LOS PROGRAMAS DE POSGRADO');
// 5) elegir AOI, luego cambiar AEI -> AOI se limpia y la lista cambia
fca.getRange(r0, 3).setValue(aois[0]); edit(r0, 3);
assert.strictEqual(fca.v(r0, 3), aois[0]);
fca.getRange(r0, 2).setValue(X.POI_AEI['AEI.01.02']); edit(r0, 2);
assert.strictEqual(fca.v(r0, 3), '');
aois = fca.allowed(r0, 3);
assert.strictEqual(aois.length, 5);
assert.ok(aois.every(a => !a.startsWith('AOI00008800777')));
// 6) cambiar OEI a OEI.04 -> AEI y AOI se limpian; AEI de OEI.04
fca.getRange(r0, 3).setValue(aois[0]);
fca.getRange(r0, 1).setValue(X.POI_OEI['OEI.04']); edit(r0, 1);
assert.strictEqual(fca.v(r0, 2), ''); assert.strictEqual(fca.v(r0, 3), '');
assert.deepStrictEqual(fca.allowed(r0, 3), [X.CFG.PH_AOI]);
assert.deepStrictEqual(fca.allowed(r0, 2).map(a => a.slice(0, 9)), ['AEI.04.02', 'AEI.04.04', 'AEI.04.05']);
// 7) fila independiente
assert.deepStrictEqual(fca.allowed(r0 + 1, 2), [X.CFG.PH_AEI]);
// 8) AEI pegado con OEI distinto: el OEI se alinea
fca.getRange(r0 + 1, 2).setValue(X.POI_AEI['AEI.02.04']); edit(r0 + 1, 2);
assert.strictEqual(fca.v(r0 + 1, 1), X.POI_OEI['OEI.02']);
assert.strictEqual(fca.allowed(r0 + 1, 3).length, 2);
// 9) hoja que no es facultad: no hace nada
const otra = ss.insertSheet('Hoja 1'); otra.getRange(2, 1).setValue('x');
X.onEdit({ range: otra.getRange(2, 1) });
// 10) todas las facultades: cuadran con los datos esperados
let tot = 0; X.POI_FACULTADES.forEach(f => { const i = X.indice_(f); i.oeis.forEach(o => i.aeisPorOei[o].forEach(a => { assert.ok(X.POI_AEI[a], a); tot += i.aoisPorAei[a].length; })); });
assert.strictEqual(tot, 649);
console.log('OK: todas las pruebas pasaron (649 actividades, 20 facultades)');
