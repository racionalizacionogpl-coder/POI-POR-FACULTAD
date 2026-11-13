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
  deleteColumn(c) { const mv = (o) => { const n = {}; for (const k in o) { const [r, cc] = k.split(',').map(Number); if (cc === c) continue; n[r + ',' + (cc > c ? cc - 1 : cc)] = o[k]; } return n; }; this.cells = mv(this.cells); this.val = mv(this.val); }
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
const ctx = { SpreadsheetApp: { newDataValidation: builder, getActive: () => ss, getUi: () => ({ alert: () => 'YES', Button: { YES: 'YES' }, ButtonSet: { YES_NO: 0 } }) }, console };
vm.createContext(ctx);
const code = ['Datos.gs', 'Codigo.gs'].map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') +
  '\n;this.__x = { POI_FACULTADES, POI_OEI, POI_AEI, CFG, indice_, facultadDeHoja_, aplicarHoja_, generarListas_, onEdit, hojasDeFacultad_, quitarColumnaOei };';
vm.runInContext(code, ctx);
const X = ctx.__x;

// 1) deteccion de hojas
for (const [n, s] of [['F01_FM', 'FM'], ['F09_FCA', 'FCA'], ['F15_FCCSS', 'FCCSS'], ['F06_FE', 'FE'], ['FPSIC', 'FPSIC'], ['F21_FIEE', 'FIEE'], ['FISI', 'FISI'], ['F20_FII', 'FII'],
  ['F20_FISI', 'FISI'], ['F17_FII', 'FII'], ['F18_FPSIC', 'FPSIC'], ['F19_FIEE', 'FIEE'], ['F05', 'FO'], ['F05_FO', 'FO']])
  assert.strictEqual(X.facultadDeHoja_(n)?.sigla, s, n);
assert.strictEqual(X.facultadDeHoja_('Anexo 1'), null);
assert.strictEqual(X.facultadDeHoja_('POI_LISTAS'), null);
assert.strictEqual(X.facultadDeHoja_('POI_CATALOGO'), null);

// 2) aplicar a la hoja FCA: solo AEI (col 1) y AOI (col 2)
const fca = ss.insertSheet('F09_FCA'); const fac = X.facultadDeHoja_('F09_FCA');
const listas = X.generarListas_(ss);
X.aplicarHoja_(fca, fac, listas);
const r0 = X.CFG.FILA_INICIO;
assert.strictEqual(fca.v(r0 - 1, 1), 'ACCIÓN ESTRATÉGICA (AEI)');
assert.strictEqual(fca.v(r0 - 1, 2), 'ACTIVIDAD OPERATIVA (CÓDIGO - DENOMINACIÓN)');
assert.strictEqual(fca.v(r0 - 1, 3), '');                       // no hay columna OEI ni tercera columna
const aeiList = fca.allowed(r0, 1);
assert.strictEqual(aeiList.length, 12);                         // AEI que FCA tiene en el PDF
assert.ok(aeiList.every(a => a.startsWith('AEI.')));
assert.deepStrictEqual(fca.allowed(r0, 2), [X.CFG.PH_AOI]);
assert.strictEqual(fca.allowed(r0, 3), null);
const edit = (r, c, nc = 1) => X.onEdit({ range: fca.getRange(r, c, 1, nc) });

// 3) elegir AEI.01.01 -> exactamente 7 AOI (segun el PDF), con codigo y denominacion
fca.getRange(r0, 1).setValue(X.POI_AEI['AEI.01.01']); edit(r0, 1);
let aois = fca.allowed(r0, 2);
assert.strictEqual(aois.length, 7);
assert.strictEqual(aois[0], 'AOI00008800777 - EJECUCIÓN DE LOS PROGRAMAS DE POSGRADO');
// 4) elegir AOI, luego cambiar AEI -> AOI se limpia y la lista cambia
fca.getRange(r0, 2).setValue(aois[0]); edit(r0, 2);
assert.strictEqual(fca.v(r0, 2), aois[0]);
fca.getRange(r0, 1).setValue(X.POI_AEI['AEI.01.02']); edit(r0, 1);
assert.strictEqual(fca.v(r0, 2), '');
aois = fca.allowed(r0, 2);
assert.strictEqual(aois.length, 5);
assert.ok(aois.every(a => !a.startsWith('AOI00008800777')));
// 5) AEI vaciado -> AOI se limpia y vuelve el placeholder
fca.getRange(r0, 2).setValue(aois[0]);
fca.getRange(r0, 1).clearContent(); edit(r0, 1);
assert.strictEqual(fca.v(r0, 2), ''); assert.deepStrictEqual(fca.allowed(r0, 2), [X.CFG.PH_AOI]);
// 6) fila independiente
assert.deepStrictEqual(fca.allowed(r0 + 1, 2), [X.CFG.PH_AOI]);
// 7) PEGAR AEI+AOI juntos (A:B) y consistentes: se conservan
fca.getRange(r0 + 1, 1).setValue(X.POI_AEI['AEI.02.04']);
fca.getRange(r0 + 1, 2).setValue('AOI00008800971 - DESARROLLO DE EVENTOS PARA FINES DE INVESTIGACIÓN');
edit(r0 + 1, 1, 2);
assert.strictEqual(fca.v(r0 + 1, 2), 'AOI00008800971 - DESARROLLO DE EVENTOS PARA FINES DE INVESTIGACIÓN');
assert.strictEqual(fca.allowed(r0 + 1, 2).length, 2);
// 8) PEGAR AEI+AOI inconsistentes (AOI de otro AEI): la AOI se limpia
fca.getRange(r0 + 2, 1).setValue(X.POI_AEI['AEI.01.01']);
fca.getRange(r0 + 2, 2).setValue('AOI00008800971 - DESARROLLO DE EVENTOS PARA FINES DE INVESTIGACIÓN');
edit(r0 + 2, 1, 2);
assert.strictEqual(fca.v(r0 + 2, 2), '');
// 9) PEGAR un AEI de otra facultad (inexistente en FCA): queda marcado, AOI limpia y sin lista
fca.getRange(r0 + 3, 1).setValue('AEI.99.99 Inexistente'); fca.getRange(r0 + 3, 2).setValue('x'); edit(r0 + 3, 1, 2);
assert.strictEqual(fca.v(r0 + 3, 2), ''); assert.deepStrictEqual(fca.allowed(r0 + 3, 2), [X.CFG.PH_AOI]);
// 9b) hoja que no es facultad: no hace nada
const otra = ss.insertSheet('Hoja 1'); otra.getRange(2, 1).setValue('x');
X.onEdit({ range: otra.getRange(2, 1) });
// 9c) migracion desde la version con OEI: columnas A=OEI, B=AEI, C=AOI
const fo = ss.insertSheet('F05_FO'); const facFo = X.facultadDeHoja_('F05_FO');
fo.getRange(1, 1).setValue('OEI'); fo.getRange(1, 2).setValue('AEI'); fo.getRange(1, 3).setValue('AOI (CÓDIGO - DENOMINACIÓN)');
const aeiFo = X.POI_FACULTADES.find(f => f.sigla === 'FO').items[0];
fo.getRange(2, 1).setValue(X.POI_OEI['OEI.0' + aeiFo[0][4]]);
fo.getRange(2, 2).setValue(X.POI_AEI[aeiFo[0]]);
fo.getRange(2, 3).setValue(aeiFo[1] + ' - ' + aeiFo[2]);
X.quitarColumnaOei();
assert.strictEqual(fo.v(1, 1), 'ACCIÓN ESTRATÉGICA (AEI)'); assert.strictEqual(fo.v(1, 2), 'ACTIVIDAD OPERATIVA (CÓDIGO - DENOMINACIÓN)');
assert.strictEqual(fo.v(2, 1), X.POI_AEI[aeiFo[0]]); assert.strictEqual(fo.v(2, 2), aeiFo[1] + ' - ' + aeiFo[2]);
assert.strictEqual(fo.v(2, 3), ''); assert.strictEqual(fo.v(1, 3), '');
assert.ok(fo.allowed(2, 2).includes(aeiFo[1] + ' - ' + aeiFo[2]));
assert.strictEqual(fo.allowed(2, 3), null);
// y una segunda ejecucion ya no encuentra columna OEI
X.quitarColumnaOei(); assert.strictEqual(fo.v(2, 1), X.POI_AEI[aeiFo[0]]);
// 10) todas las facultades: cuadran con los datos esperados
let tot = 0; X.POI_FACULTADES.forEach(f => { const i = X.indice_(f); i.aeis.forEach(a => { assert.ok(X.POI_AEI[a], a); tot += i.aoisPorAei[a].length; }); });
assert.strictEqual(tot, 649);
console.log('OK: todas las pruebas pasaron (649 actividades, 20 facultades)');
