/**
 * POI 2026 - Desplegables dependientes AEI -> AOI para el libro
 * "ANEXO 1 - Inventario de productos y procesos".
 * =====================================================================
 * Requiere el archivo POI_Datos.gs (POI_AEI, POI_FACULTADES).
 *
 * Diseñado para convivir con las automatizaciones que ya tiene el libro:
 *  - NO define onEdit ni onOpen (no pisa los existentes). Todo lo que crea
 *    lleva el prefijo "poi"/"POI_".
 *  - Usa disparadores instalables (se crean con poiInstalar()).
 *
 * Qué hace
 *  1. Detecta la facultad de cada pestaña por su nombre (F01_FM, F09_FCA, ...).
 *  2. En cada pestaña de facultad coloca los desplegables SOLO en dos celdas:
 *       D7 = ACCIÓN ESTRATÉGICA (AEI)   ->   E7 = ACTIVIDAD OPERATIVA (AOI)
 *     La AOI se muestra como "AOI0000XXXXXXX - DENOMINACIÓN" y solo ofrece las
 *     actividades del AEI elegido.
 *  3. Cada persona puede copiar D7:E7 y pegarlo donde necesite (incluso solo el AEI):
 *     el AEI pegado en CUALQUIER celda recupera la lista de actividades de la celda
 *     de su derecha.
 *
 * Instalación: pegar POI_Datos.gs y POI_Codigo.gs, y ejecutar UNA vez poiInstalar().
 */

const POI_CFG = {
  FILA: 7,            // fila de las celdas con desplegable
  COL_AEI: 4,         // D  (A=1, B=2, C=3, D=4, ...)
  COL_AOI: 5,         // E
  SEP: ' - ',
  HOJA_LISTAS: 'POI_LISTAS',
  HOJA_CATALOGO: 'POI_CATALOGO',
  PH_AOI: '(Seleccione primero el AEI)',
  MAX_FILAS_POR_EDICION: 200,
  MAX_COLS_POR_EDICION: 30,
  IGNORAR_HOJAS: /^copia de/i,   // al "Aplicar a todas" se omiten las copias de trabajo
};

/* ------------------------------------------------------------------ */
/*  Instalación y menú                                                 */
/* ------------------------------------------------------------------ */

/** EJECUTAR UNA VEZ desde el editor: crea los disparadores, las listas y los desplegables. */
function poiInstalar() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['poiAlEditar', 'poiCrearMenu'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('poiAlEditar').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('poiCrearMenu').forSpreadsheet(ss).onOpen().create();
  poiCrearMenu();
  const n = poiAplicarATodas_(ss);
  ss.toast('Instalado. Desplegables en D7/E7 de ' + n + ' hojas.', 'POI 2026', 10);
}

function poiCrearMenu() {
  SpreadsheetApp.getUi().createMenu('POI AEI-AOI')
    .addItem('1. Aplicar en D7/E7 a TODAS las facultades', 'poiAplicarATodas')
    .addItem('2. Aplicar solo a la hoja activa', 'poiAplicarAHojaActiva')
    .addSeparator()
    .addItem('Diagnóstico (hoja ↔ facultad)', 'poiDiagnostico')
    .addItem('Generar hoja POI_CATALOGO (consulta)', 'poiGenerarCatalogo')
    .addItem('Regenerar listas (POI_LISTAS)', 'poiRegenerarListas')
    .addToUi();
}

function poiAplicarATodas() {
  const ss = SpreadsheetApp.getActive();
  const n = poiAplicarATodas_(ss);
  const sin = poiFacultadesSinHoja_(poiHojas_(ss));
  ss.toast('Desplegables aplicados en ' + n + ' hojas.', 'POI 2026', 8);
  if (sin.length) SpreadsheetApp.getUi().alert('Facultades sin hoja detectada: ' + sin.join(', '));
}

function poiAplicarATodas_(ss) {
  const hojas = poiHojas_(ss);
  const listas = poiGenerarListas_(ss);
  hojas.forEach(h => poiAplicarHoja_(h.hoja, h.fac, listas));
  return hojas.length;
}

function poiAplicarAHojaActiva() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getActiveSheet();
  const fac = poiFacultadDeHoja_(hoja.getName());
  if (!fac) {
    SpreadsheetApp.getUi().alert('La hoja "' + hoja.getName() + '" no coincide con ninguna facultad.');
    return;
  }
  poiAplicarHoja_(hoja, fac, poiGenerarListas_(ss));
  ss.toast('Desplegables aplicados en ' + hoja.getName(), 'POI 2026', 5);
}

function poiDiagnostico() {
  const ss = SpreadsheetApp.getActive();
  const hojas = poiHojas_(ss);
  const lineas = hojas.map(h => h.hoja.getName() + '  ->  ' + h.fac.sigla + ' (' + h.fac.cc + ', ' +
    h.fac.items.length + ' actividades)');
  const sin = poiFacultadesSinHoja_(hojas);
  SpreadsheetApp.getUi().alert('Diagnóstico POI',
    (lineas.join('\n') || '(ninguna hoja reconocida)') +
    (sin.length ? '\n\nFacultades SIN hoja: ' + sin.join(', ') : '\n\nTodas las facultades tienen hoja.') +
    '\n\n(Se omiten las pestañas "Copia de ...")',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function poiRegenerarListas() {
  poiGenerarListas_(SpreadsheetApp.getActive());
  SpreadsheetApp.getActive().toast('Listas regeneradas.', 'POI 2026', 5);
}

/* ------------------------------------------------------------------ */
/*  Detección de hojas                                                 */
/* ------------------------------------------------------------------ */

const POI_CACHE_HOJA = {};

/**
 * F09_FCA -> facultad FCA. La sigla manda; el número de formulario solo se usa si
 * ninguna sigla coincide. También reconoce copias ("Copia de F10_FCB").
 */
function poiFacultadDeHoja_(nombre) {
  if (nombre in POI_CACHE_HOJA) return POI_CACHE_HOJA[nombre];
  const tokens = String(nombre).toUpperCase().split(/[^A-Z0-9]+/).filter(String);
  let hits = POI_FACULTADES.filter(f => tokens.indexOf(f.sigla) >= 0);
  if (!hits.length) hits = POI_FACULTADES.filter(f => f.formulario && tokens.indexOf(f.formulario) >= 0);
  return (POI_CACHE_HOJA[nombre] = hits.length === 1 ? hits[0] : null);
}

function poiHojas_(ss) {
  const out = [];
  ss.getSheets().forEach(h => {
    if (POI_CFG.IGNORAR_HOJAS.test(h.getName())) return;
    const fac = poiFacultadDeHoja_(h.getName());
    if (fac) out.push({ hoja: h, fac: fac });
  });
  return out;
}

function poiFacultadesSinHoja_(hojas) {
  const usadas = {};
  hojas.forEach(h => usadas[h.fac.sigla] = true);
  return POI_FACULTADES.filter(f => !usadas[f.sigla]).map(f => f.sigla);
}

/* ------------------------------------------------------------------ */
/*  Índices por facultad (AEI -> AOI)                                  */
/* ------------------------------------------------------------------ */

function poiEtiquetaAoi_(cod, den) { return cod + POI_CFG.SEP + den; }

function poiIndice_(fac) {
  if (fac._idx) return fac._idx;
  const idx = { aeis: [], aoisPorAei: {}, aeiPorEtiqueta: {}, aoiPorEtiqueta: {} };
  fac.items.forEach(it => {
    const codAei = it[0];
    if (idx.aeis.indexOf(codAei) < 0) idx.aeis.push(codAei);
    (idx.aoisPorAei[codAei] = idx.aoisPorAei[codAei] || []).push({ cod: it[1], den: it[2] });
  });
  idx.aeis.sort();
  idx.aeis.forEach(a => {
    idx.aeiPorEtiqueta[POI_AEI[a]] = a;
    idx.aoisPorAei[a].forEach(x => idx.aoiPorEtiqueta[poiEtiquetaAoi_(x.cod, x.den)] = { aei: a, cod: x.cod, den: x.den });
  });
  return (fac._idx = idx);
}

/* ------------------------------------------------------------------ */
/*  Hoja de listas (origen de los desplegables, oculta)                */
/*  Se usan rangos y no listas "en línea" para evitar límites de       */
/*  tamaño de las validaciones.                                        */
/* ------------------------------------------------------------------ */

/** Disposición determinista: { filas: [[valor, tipo, sigla, padre]], mapa: { clave: {fila, n} } } */
function poiDisposicionListas_() {
  const filas = [], mapa = {};
  const FILA0 = 2;                                  // fila 1 = encabezado
  const add = (clave, tipo, sigla, padre, valores) => {
    mapa[clave] = { fila: FILA0 + filas.length, n: valores.length };
    valores.forEach(v => filas.push([v, tipo, sigla, padre]));
  };
  add('PH_AOI', 'PLACEHOLDER', '', '', [POI_CFG.PH_AOI]);
  POI_FACULTADES.forEach(f => {
    const idx = poiIndice_(f);
    add(f.sigla + '|AEI', 'AEI', f.sigla, '', idx.aeis.map(a => POI_AEI[a]));
    idx.aeis.forEach(a =>
      add(f.sigla + '|' + a, 'AOI', f.sigla, a, idx.aoisPorAei[a].map(x => poiEtiquetaAoi_(x.cod, x.den))));
  });
  return { filas: filas, mapa: mapa };
}

function poiGenerarListas_(ss) {
  const disp = poiDisposicionListas_();
  let sh = ss.getSheetByName(POI_CFG.HOJA_LISTAS);
  if (!sh) sh = ss.insertSheet(POI_CFG.HOJA_LISTAS);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['VALOR (no editar)', 'TIPO', 'FACULTAD', 'PADRE']]).setFontWeight('bold');
  if (sh.getMaxRows() < disp.filas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), disp.filas.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, disp.filas.length, 4).setValues(disp.filas);
  sh.setColumnWidth(1, 600);
  sh.hideSheet();
  return { hoja: sh, mapa: disp.mapa };
}

function poiObtenerListas_(ss) {
  const sh = ss.getSheetByName(POI_CFG.HOJA_LISTAS);
  if (!sh) return poiGenerarListas_(ss);
  return { hoja: sh, mapa: poiDisposicionListas_().mapa };
}

function poiRegla_(listas, clave, ayuda) {
  const e = listas.mapa[clave];
  if (!e) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(listas.hoja.getRange(e.fila, 1, e.n, 1), true)
    .setAllowInvalid(false)
    .setHelpText(ayuda)
    .build();
}

function poiReglaAei_(listas, fac) {
  return poiRegla_(listas, fac.sigla + '|AEI', 'Seleccione la acción estratégica de la lista');
}

/** Fija la lista de AOI de la celda (fila, col) según el AEI indicado. */
function poiRefrescarAoi_(hoja, fac, idx, listas, fila, col, aeiTxt) {
  const codAei = idx.aeiPorEtiqueta[aeiTxt];
  const regla = codAei
    ? poiRegla_(listas, fac.sigla + '|' + codAei, 'Seleccione la actividad operativa (código - denominación)')
    : poiRegla_(listas, 'PH_AOI', POI_CFG.PH_AOI);
  hoja.getRange(fila, col).setDataValidation(regla);
}

/* ------------------------------------------------------------------ */
/*  Aplicar a una hoja: SOLO D7 (AEI) y E7 (AOI)                       */
/*  No modifica ningún valor, únicamente la validación de esas celdas. */
/* ------------------------------------------------------------------ */

function poiAplicarHoja_(hoja, fac, listas) {
  const idx = poiIndice_(fac);
  const f = POI_CFG.FILA;
  hoja.getRange(f, POI_CFG.COL_AEI).setDataValidation(poiReglaAei_(listas, fac));
  poiRefrescarAoi_(hoja, fac, idx, listas, f, POI_CFG.COL_AOI, String(hoja.getRange(f, POI_CFG.COL_AEI).getValue()).trim());
}

/* ------------------------------------------------------------------ */
/*  Disparador onEdit INSTALABLE (poiAlEditar)                         */
/*                                                                     */
/*  Funciona en cualquier celda de una pestaña de facultad, para que   */
/*  copiar y pegar el AEI en otra zona siga dando su lista de          */
/*  actividades. Convención: la AOI está en la celda de la derecha.    */
/* ------------------------------------------------------------------ */

/** 'AEI' / 'AOI' según la lista a la que apunta una regla de validación, o '' si no es nuestra. */
function poiTipoDeRegla_(regla, listas) {
  try {
    if (!regla) return '';
    const v = regla.getCriteriaValues();
    const r = v && v[0];
    if (!r || typeof r.getSheet !== 'function' || r.getSheet().getName() !== POI_CFG.HOJA_LISTAS) return '';
    const inicio = r.getRow();
    for (const clave in listas.mapa) {
      if (listas.mapa[clave].fila === inicio) return /\|AEI$/.test(clave) ? 'AEI' : 'AOI';
    }
  } catch (err) { /* regla de otro tipo (listas propias del libro) */ }
  return '';
}

function poiAlEditar(e) {
  try {
    if (!e || !e.range) return;
    const hoja = e.range.getSheet();
    const fac = poiFacultadDeHoja_(hoja.getName());
    if (!fac) return;                                 // MAPA DE PROCESOS, POI_LISTAS, etc.

    const rg = e.range;
    const f0 = rg.getRow(), c0 = rg.getColumn();
    const nf = Math.min(rg.getNumRows(), POI_CFG.MAX_FILAS_POR_EDICION);
    const nc = Math.min(rg.getNumColumns(), POI_CFG.MAX_COLS_POR_EDICION);
    const zona = hoja.getRange(f0, c0, nf, nc);
    const valores = zona.getValues();
    const reglas = zona.getDataValidations();

    const ss = hoja.getParent();
    const listas = poiObtenerListas_(ss);
    const idx = poiIndice_(fac);

    // Celdas AEI: la celda D7, una celda con la lista AEI (se copia al pegar) o con un texto
    // que coincide con un AEI de la facultad (pegado como valor). Una AOI editada remite al
    // AEI de su izquierda.
    const aeiCeldas = {};
    for (let i = 0; i < nf; i++) {
      for (let j = 0; j < nc; j++) {
        const fila = f0 + i, col = c0 + j;
        const v = String(valores[i][j]).trim();
        const tipo = poiTipoDeRegla_(reglas[i][j], listas);
        const cfgAei = fila === POI_CFG.FILA && col === POI_CFG.COL_AEI;
        const cfgAoi = fila === POI_CFG.FILA && col === POI_CFG.COL_AOI;
        if (cfgAei || tipo === 'AEI' || idx.aeiPorEtiqueta[v]) {
          aeiCeldas[fila + ',' + col] = { fila: fila, col: col, cfg: cfgAei };
        } else if (cfgAoi || tipo === 'AOI' || idx.aoiPorEtiqueta[v] || v === POI_CFG.PH_AOI) {
          const colAei = cfgAoi ? POI_CFG.COL_AEI : col - 1;
          if (colAei < 1) continue;
          const izq = hoja.getRange(fila, colAei);
          if (cfgAoi || idx.aeiPorEtiqueta[String(izq.getValue()).trim()] ||
              poiTipoDeRegla_(izq.getDataValidation(), listas) === 'AEI') {
            aeiCeldas[fila + ',' + colAei] = { fila: fila, col: colAei, cfg: cfgAoi };
          }
        }
      }
    }
    Object.keys(aeiCeldas).forEach(k => poiProcesarAei_(hoja, fac, idx, listas, aeiCeldas[k]));
  } catch (err) {
    SpreadsheetApp.getActive().toast('Error en desplegables POI: ' + err.message, 'POI 2026', 10);
  }
}

/** Sincroniza la AOI (celda a la derecha) con el AEI de la celda indicada. */
function poiProcesarAei_(hoja, fac, idx, listas, c) {
  const cA = hoja.getRange(c.fila, c.col);
  const colAoi = c.col === POI_CFG.COL_AEI && c.fila === POI_CFG.FILA ? POI_CFG.COL_AOI : c.col + 1;
  const cP = hoja.getRange(c.fila, colAoi);
  const aei = String(cA.getValue()).trim();
  let aoi = String(cP.getValue()).trim();
  const codAei = idx.aeiPorEtiqueta[aei];
  const datosAoi = idx.aoiPorEtiqueta[aoi];

  // Solo se limpia la AOI si es nuestra (celda D7/E7, aviso, AOI conocida o con lista nuestra).
  // Un texto ajeno en esa celda no se borra, pero igualmente recibe la lista de AOI.
  const nuestra = c.cfg || aoi === '' || aoi === POI_CFG.PH_AOI || datosAoi ||
                  poiTipoDeRegla_(cP.getDataValidation(), listas) !== '';
  if (nuestra) {
    if (aoi === POI_CFG.PH_AOI) { cP.clearContent(); aoi = ''; }
    if (aoi !== '' && !(codAei && datosAoi && datosAoi.aei === codAei)) cP.clearContent();
  }

  // Se restablece la lista del AEI (por si se pegó con otra validación) y la de la AOI.
  cA.setDataValidation(poiReglaAei_(listas, fac));
  poiRefrescarAoi_(hoja, fac, idx, listas, c.fila, colAoi, aei);
}

/* ------------------------------------------------------------------ */
/*  Hoja de consulta con todo el catálogo (opcional)                   */
/* ------------------------------------------------------------------ */

function poiGenerarCatalogo() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(POI_CFG.HOJA_CATALOGO);
  if (!sh) sh = ss.insertSheet(POI_CFG.HOJA_CATALOGO);
  sh.clear();
  const filas = [['FORMULARIO', 'SIGLA', 'CENTRO DE COSTO', 'FACULTAD', 'AEI', 'COD. AOI', 'ACTIVIDAD OPERATIVA']];
  POI_FACULTADES.forEach(f => f.items.forEach(it => filas.push([
    f.formulario || '', f.sigla, f.cc, f.nombre, POI_AEI[it[0]], it[1], it[2]])));
  sh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  sh.getRange(1, 1, 1, filas[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, filas[0].length);
  ss.toast('Catálogo generado: ' + (filas.length - 1) + ' actividades.', 'POI 2026', 5);
}
