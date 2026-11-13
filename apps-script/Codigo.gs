/**
 * POI 2026 por facultad - Desplegables dependientes AEI -> AOI
 * ====================================================================
 * Requiere el archivo Datos.gs (POI_OEI, POI_AEI, POI_FACULTADES).
 *
 * Qué hace
 *  1. Detecta la facultad de cada hoja por su nombre (F01_FM, F09_FCA, ...).
 *  2. En dos columnas de cada hoja crea desplegables encadenados:
 *       ACCIÓN ESTRATÉGICA (AEI)  ->  ACTIVIDAD OPERATIVA (AOI)
 *     La AOI se muestra como "AOI0000XXXXXXX - DENOMINACIÓN" y solo ofrece las
 *     actividades del AEI elegido.
 *  3. Al cambiar el AEI, el onEdit limpia y recalcula la AOI.
 *
 * Uso: menú "POI 2026" (aparece al abrir el Sheet) -> "1. Aplicar a TODAS las
 * facultades". Antes ajuste CFG si las columnas no son A y B.
 */

const CFG = {
  FILA_INICIO: 2,   // primera fila de datos (la fila anterior es el encabezado)
  FILAS: 300,       // cuántas filas quedan con desplegables
  COL_AEI: 1,       // A  (número de columna: A=1, B=2, C=3, ...)  Acción estratégica
  COL_AOI: 2,       // B  Actividad operativa (código - denominación)
  COL_COD: 0,       // opcional: columna donde se copia SOLO el código AOI (0 = no usar)
  COL_DENOM: 0,     // opcional: columna donde se copia SOLO la denominación (0 = no usar)
  ESCRIBIR_ENCABEZADOS: true,  // solo escribe si la celda de encabezado está vacía
  ENCABEZADOS: ['ACCIÓN ESTRATÉGICA (AEI)', 'ACTIVIDAD OPERATIVA (CÓDIGO - DENOMINACIÓN)'],
  FORMATO: true,    // ajusta ancho y ajuste de texto de esas columnas
  ANCHOS: [330, 420],
  SEP: ' - ',
  HOJA_LISTAS: 'POI_LISTAS',
  HOJA_CATALOGO: 'POI_CATALOGO',
  PH_AOI: '(Seleccione primero el AEI)',
  MAX_FILAS_POR_EDICION: 200,
};

/* ------------------------------------------------------------------ */
/*  Menú                                                               */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('POI 2026')
    .addItem('1. Aplicar a TODAS las facultades', 'aplicarATodas')
    .addItem('2. Aplicar solo a la hoja activa', 'aplicarAHojaActiva')
    .addItem('3. Quitar la columna OEI (hojas hechas con la versión anterior)', 'quitarColumnaOei')
    .addSeparator()
    .addItem('Diagnóstico (qué hoja corresponde a qué facultad)', 'diagnostico')
    .addItem('Generar hoja POI_CATALOGO (consulta)', 'generarCatalogo')
    .addItem('Regenerar listas (POI_LISTAS)', 'generarListas')
    .addToUi();
}

function aplicarATodas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const hojas = hojasDeFacultad_(ss);
  if (!hojas.length) {
    ui.alert('No se encontró ninguna hoja de facultad. Ejecute "Diagnóstico" para ver los nombres esperados.');
    return;
  }
  const conDatos = hojas.filter(h => contarOcupadas_(h.hoja) > 0).map(h => h.hoja.getName());
  if (conDatos.length) {
    const r = ui.alert('Columnas con contenido',
      'Estas hojas ya tienen datos en las columnas configuradas (AEI/AOI):\n\n' +
      conDatos.join(', ') + '\n\nSe reemplazarán SOLO las validaciones (desplegables); los valores no se borran.\n¿Continuar?',
      ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  const listas = generarListas_(ss);
  hojas.forEach(h => aplicarHoja_(h.hoja, h.fac, listas));
  const sin = FACULTADES_SIN_HOJA_(hojas);
  ss.toast('Desplegables aplicados en ' + hojas.length + ' hojas.', 'POI 2026', 8);
  if (sin.length) {
    ui.alert('Facultades sin hoja detectada: ' + sin.join(', ') +
      '\n\nRevise el nombre de la hoja (debe contener la sigla, ej. F09_FCA) o use "Diagnóstico".');
  }
}

function aplicarAHojaActiva() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getActiveSheet();
  const fac = facultadDeHoja_(hoja.getName());
  if (!fac) {
    SpreadsheetApp.getUi().alert('La hoja "' + hoja.getName() + '" no coincide con ninguna facultad.');
    return;
  }
  aplicarHoja_(hoja, fac, generarListas_(ss));
  ss.toast('Desplegables aplicados en ' + hoja.getName(), 'POI 2026', 5);
}

function diagnostico() {
  const ss = SpreadsheetApp.getActive();
  const hojas = hojasDeFacultad_(ss);
  const lineas = hojas.map(h => h.hoja.getName() + '  ->  ' + h.fac.sigla + ' (' + h.fac.cc + ' ' +
    h.fac.nombre + ', ' + h.fac.items.length + ' actividades)');
  const sin = FACULTADES_SIN_HOJA_(hojas);
  SpreadsheetApp.getUi().alert('Diagnóstico POI',
    (lineas.join('\n') || '(ninguna hoja reconocida)') +
    (sin.length ? '\n\nFacultades SIN hoja: ' + sin.join(', ') : '\n\nTodas las facultades tienen hoja.'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ------------------------------------------------------------------ */
/*  Detección de hojas                                                 */
/* ------------------------------------------------------------------ */

const CACHE_HOJA_ = {};

/**
 * F09_FCA -> facultad FCA. Compara cada "palabra" del nombre con la sigla y, solo si
 * ninguna sigla coincide, con el formulario. La sigla manda porque el número de
 * formulario no es confiable (ej. F20_FISI: F20 es el formulario de FII, pero la hoja es FISI).
 */
function facultadDeHoja_(nombre) {
  if (nombre in CACHE_HOJA_) return CACHE_HOJA_[nombre];
  const tokens = String(nombre).toUpperCase().split(/[^A-Z0-9]+/).filter(String);
  let hits = POI_FACULTADES.filter(f => tokens.indexOf(f.sigla) >= 0);
  if (!hits.length) hits = POI_FACULTADES.filter(f => f.formulario && tokens.indexOf(f.formulario) >= 0);
  return (CACHE_HOJA_[nombre] = hits.length === 1 ? hits[0] : null);
}

function hojasDeFacultad_(ss) {
  const out = [];
  ss.getSheets().forEach(h => {
    const fac = facultadDeHoja_(h.getName());
    if (fac) out.push({ hoja: h, fac: fac });
  });
  return out;
}

function FACULTADES_SIN_HOJA_(hojas) {
  const usadas = {};
  hojas.forEach(h => usadas[h.fac.sigla] = true);
  return POI_FACULTADES.filter(f => !usadas[f.sigla]).map(f => f.sigla);
}

function contarOcupadas_(hoja) {
  let n = 0;
  [CFG.COL_AEI, CFG.COL_AOI].forEach(c => {
    const max = hoja.getMaxRows();
    if (CFG.FILA_INICIO > max) return;
    const filas = Math.min(CFG.FILAS, max - CFG.FILA_INICIO + 1);
    hoja.getRange(CFG.FILA_INICIO, c, filas, 1).getValues()
      .forEach(r => { if (String(r[0]).trim() !== '') n++; });
  });
  return n;
}

/* ------------------------------------------------------------------ */
/*  Índices por facultad (AEI -> AOI)                                  */
/* ------------------------------------------------------------------ */

function oeiDeAei_(codAei) { return 'OEI.' + codAei.split('.')[1]; }
function etiquetaAoi_(cod, den) { return cod + CFG.SEP + den; }

function indice_(fac) {
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
    idx.aoisPorAei[a].forEach(x => idx.aoiPorEtiqueta[etiquetaAoi_(x.cod, x.den)] = { aei: a, cod: x.cod, den: x.den });
  });
  return (fac._idx = idx);
}

/* ------------------------------------------------------------------ */
/*  Hoja de listas (origen de los desplegables)                        */
/*  Se usan rangos y no listas "en línea" para evitar límites de       */
/*  tamaño de las validaciones.                                        */
/* ------------------------------------------------------------------ */

/** Disposición determinista: { filas: [[valor, tipo, sigla, padre]], mapa: { clave: {fila, n} } } */
function disposicionListas_() {
  const filas = [], mapa = {};
  const FILA0 = 2;                                  // fila 1 = encabezado
  const add = (clave, tipo, sigla, padre, valores) => {
    mapa[clave] = { fila: FILA0 + filas.length, n: valores.length };
    valores.forEach(v => filas.push([v, tipo, sigla, padre]));
  };
  add('PH_AOI', 'PLACEHOLDER', '', '', [CFG.PH_AOI]);
  POI_FACULTADES.forEach(f => {
    const idx = indice_(f);
    add(f.sigla + '|AEI', 'AEI', f.sigla, '', idx.aeis.map(a => POI_AEI[a]));
    idx.aeis.forEach(a =>
      add(f.sigla + '|' + a, 'AOI', f.sigla, a, idx.aoisPorAei[a].map(x => etiquetaAoi_(x.cod, x.den))));
  });
  return { filas: filas, mapa: mapa };
}

function generarListas() { generarListas_(SpreadsheetApp.getActive()); SpreadsheetApp.getActive().toast('Listas regeneradas.'); }

function generarListas_(ss) {
  const disp = disposicionListas_();
  let sh = ss.getSheetByName(CFG.HOJA_LISTAS);
  if (!sh) sh = ss.insertSheet(CFG.HOJA_LISTAS);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['VALOR (no editar)', 'TIPO', 'FACULTAD', 'PADRE']]).setFontWeight('bold');
  if (sh.getMaxRows() < disp.filas.length + 1) sh.insertRowsAfter(sh.getMaxRows(), disp.filas.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, disp.filas.length, 4).setValues(disp.filas);
  sh.setColumnWidth(1, 600);
  sh.hideSheet();
  return { hoja: sh, mapa: disp.mapa };
}

function obtenerListas_(ss) {
  const sh = ss.getSheetByName(CFG.HOJA_LISTAS);
  if (!sh) return generarListas_(ss);
  return { hoja: sh, mapa: disposicionListas_().mapa };
}

function reglaDeLista_(listas, clave, ayuda) {
  const e = listas.mapa[clave];
  if (!e) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(listas.hoja.getRange(e.fila, 1, e.n, 1), true)
    .setAllowInvalid(false)
    .setHelpText(ayuda)
    .build();
}

/* ------------------------------------------------------------------ */
/*  Aplicar desplegables a una hoja                                    */
/* ------------------------------------------------------------------ */

function aplicarHoja_(hoja, fac, listas) {
  const n = CFG.FILAS, r0 = CFG.FILA_INICIO;
  if (hoja.getMaxRows() < r0 + n - 1) hoja.insertRowsAfter(hoja.getMaxRows(), r0 + n - 1 - hoja.getMaxRows());

  if (CFG.ESCRIBIR_ENCABEZADOS && r0 > 1) {
    [CFG.COL_AEI, CFG.COL_AOI].forEach((c, i) => {
      const celda = hoja.getRange(r0 - 1, c);
      if (String(celda.getValue()).trim() === '') celda.setValue(CFG.ENCABEZADOS[i]).setFontWeight('bold');
    });
  }
  if (CFG.FORMATO) {
    [CFG.COL_AEI, CFG.COL_AOI].forEach((c, i) => {
      hoja.setColumnWidth(c, CFG.ANCHOS[i]);
      hoja.getRange(r0, c, n, 1).setWrap(true).setVerticalAlignment('middle');
    });
  }

  // Estado base: AEI libre dentro de la lista de la facultad; AOI bloqueada hasta elegir AEI.
  hoja.getRange(r0, CFG.COL_AEI, n, 1).setDataValidation(reglaAei_(listas, fac));
  hoja.getRange(r0, CFG.COL_AOI, n, 1).setDataValidation(
    reglaDeLista_(listas, 'PH_AOI', CFG.PH_AOI));

  // Filas que ya traen un AEI: dejar su lista de AOI consistente.
  const idx = indice_(fac);
  const va = hoja.getRange(r0, CFG.COL_AEI, n, 1).getValues().map(r => String(r[0]).trim());
  for (let i = 0; i < n; i++) {
    if (va[i]) refrescarListaAoi_(hoja, fac, idx, listas, r0 + i, CFG.COL_AOI, va[i]);
  }
}

function reglaAei_(listas, fac) {
  return reglaDeLista_(listas, fac.sigla + '|AEI', 'Seleccione la acción estratégica de la lista');
}

/** Fija la lista de AOI de la celda (fila, col) según el AEI indicado. */
function refrescarListaAoi_(hoja, fac, idx, listas, fila, col, aeiTxt) {
  const codAei = idx.aeiPorEtiqueta[aeiTxt];
  const regla = codAei
    ? reglaDeLista_(listas, fac.sigla + '|' + codAei, 'Seleccione la actividad operativa (código - denominación)')
    : reglaDeLista_(listas, 'PH_AOI', CFG.PH_AOI);
  hoja.getRange(fila, col).setDataValidation(regla);
}

/* ------------------------------------------------------------------ */
/*  Migración desde la versión con OEI                                 */
/* ------------------------------------------------------------------ */

/** Borra la columna OEI (columna A) de las hojas hechas con la versión anterior y reaplica. */
function quitarColumnaOei() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const hojas = hojasDeFacultad_(ss);
  const antiguas = hojas.filter(h =>
    String(h.hoja.getRange(CFG.FILA_INICIO - 1, 1).getValue()).trim().toUpperCase() === 'OEI');
  if (!antiguas.length) {
    ui.alert('No hay hojas con la columna "OEI" en A' + (CFG.FILA_INICIO - 1) + '. Nada que quitar.');
    return;
  }
  const r = ui.alert('Quitar columna OEI',
    'Se ELIMINARÁ la columna A (OEI) de ' + antiguas.length + ' hojas y las demás columnas se correrán a la izquierda.\n' +
    'El AEI quedará en A y la AOI en B.\n\n¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const listas = generarListas_(ss);
  antiguas.forEach(h => {
    h.hoja.deleteColumn(1);
    if (CFG.ESCRIBIR_ENCABEZADOS && CFG.FILA_INICIO > 1) {
      [CFG.COL_AEI, CFG.COL_AOI].forEach((c, i) =>
        h.hoja.getRange(CFG.FILA_INICIO - 1, c).setValue(CFG.ENCABEZADOS[i]).setFontWeight('bold'));
    }
    aplicarHoja_(h.hoja, h.fac, listas);
  });
  ss.toast('Columna OEI quitada en ' + antiguas.length + ' hojas.', 'POI 2026', 8);
}

/* ------------------------------------------------------------------ */
/*  Disparador onEdit (simple: no requiere autorización adicional)     */
/*                                                                     */
/*  Funciona en cualquier lugar de la hoja, no solo en las columnas    */
/*  configuradas, para que copiar y pegar un AEI en otra zona siga     */
/*  dando su lista de actividades. Convención: la AOI está en la       */
/*  celda inmediatamente a la derecha del AEI.                         */
/* ------------------------------------------------------------------ */

/** Tipo de lista (AEI / AOI) a la que apunta una regla de validación, o '' si no es nuestra. */
function tipoDeRegla_(regla, listas) {
  try {
    if (!regla) return '';
    const v = regla.getCriteriaValues();
    const r = v && v[0];
    if (!r || typeof r.getSheet !== 'function' || r.getSheet().getName() !== CFG.HOJA_LISTAS) return '';
    const inicio = r.getRow();
    for (const clave in listas.mapa) {
      if (listas.mapa[clave].fila === inicio) return /\|AEI$/.test(clave) ? 'AEI' : 'AOI';
    }
  } catch (err) { /* regla de otro tipo */ }
  return '';
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const hoja = e.range.getSheet();
    const fac = facultadDeHoja_(hoja.getName());
    if (!fac) return;

    const rg = e.range;
    const f0 = rg.getRow(), c0 = rg.getColumn();
    const nf = Math.min(rg.getNumRows(), CFG.MAX_FILAS_POR_EDICION);
    const nc = Math.min(rg.getNumColumns(), 30);
    const zona = hoja.getRange(f0, c0, nf, nc);
    const valores = zona.getValues();
    const reglas = zona.getDataValidations();

    const ss = hoja.getParent();
    const listas = obtenerListas_(ss);
    const idx = indice_(fac);
    const enFilasCfg = f => f >= CFG.FILA_INICIO && f < CFG.FILA_INICIO + CFG.FILAS;

    // 1) Qué celdas son AEI: columna configurada, regla de lista AEI (copiada al pegar)
    //    o valor que coincide con un AEI de la facultad (pegado solo como valor).
    //    Una AOI editada remite a su AEI de la izquierda.
    const aeiCeldas = {};
    for (let i = 0; i < nf; i++) {
      for (let j = 0; j < nc; j++) {
        const fila = f0 + i, col = c0 + j;
        const v = String(valores[i][j]).trim();
        const tipo = tipoDeRegla_(reglas[i][j], listas);
        const cfgAei = col === CFG.COL_AEI && enFilasCfg(fila);
        const cfgAoi = col === CFG.COL_AOI && enFilasCfg(fila);
        if (cfgAei || tipo === 'AEI' || idx.aeiPorEtiqueta[v]) {
          aeiCeldas[fila + ',' + col] = { fila: fila, col: col, cfg: cfgAei };
        } else if (cfgAoi || tipo === 'AOI' || idx.aoiPorEtiqueta[v] || v === CFG.PH_AOI) {
          const colAei = cfgAoi ? CFG.COL_AEI : col - 1;
          if (colAei < 1) continue;
          const izq = hoja.getRange(fila, colAei);
          const vi = String(izq.getValue()).trim();
          if (cfgAoi || idx.aeiPorEtiqueta[vi] || tipoDeRegla_(izq.getDataValidation(), listas) === 'AEI') {
            aeiCeldas[fila + ',' + colAei] = { fila: fila, col: colAei, cfg: cfgAoi };
          }
        }
      }
    }
    Object.keys(aeiCeldas).forEach(k => procesarAei_(hoja, fac, idx, listas, aeiCeldas[k]));
  } catch (err) {
    SpreadsheetApp.getActive().toast('Error en desplegables POI: ' + err.message, 'POI 2026', 10);
  }
}

/** Sincroniza la AOI (celda a la derecha) con el AEI de la celda (fila, col). */
function procesarAei_(hoja, fac, idx, listas, c) {
  const cA = hoja.getRange(c.fila, c.col);
  const colAoi = c.col === CFG.COL_AEI ? CFG.COL_AOI : c.col + 1;
  const cP = hoja.getRange(c.fila, colAoi);
  const aei = String(cA.getValue()).trim();
  let aoi = String(cP.getValue()).trim();
  const codAei = idx.aeiPorEtiqueta[aei];
  let datosAoi = idx.aoiPorEtiqueta[aoi];

  // Solo se toca la celda de la derecha si es de las columnas configuradas o si está vacía,
  // trae una AOI conocida / el aviso, o ya tiene una lista nuestra: así no se pisa texto ajeno.
  const reglaAoi = cP.getDataValidation();
  const nuestra = c.cfg || aoi === '' || aoi === CFG.PH_AOI || datosAoi ||
                  tipoDeRegla_(reglaAoi, listas) !== '';
  if (!nuestra) return;

  if (aoi === CFG.PH_AOI) { cP.clearContent(); aoi = ''; }
  // La AOI actual debe pertenecer al AEI; si no, se limpia.
  if (aoi !== '' && !(codAei && datosAoi && datosAoi.aei === codAei)) {
    cP.clearContent(); aoi = ''; datosAoi = undefined;
  }

  // Se restablece la lista del AEI (por si se pegó desde otra hoja con otra validación) y la de la AOI.
  cA.setDataValidation(reglaAei_(listas, fac));
  refrescarListaAoi_(hoja, fac, idx, listas, c.fila, colAoi, aei);

  // Columnas opcionales (solo zona configurada): código y denominación por separado.
  if (c.cfg && CFG.COL_COD > 0) hoja.getRange(c.fila, CFG.COL_COD).setValue(datosAoi ? datosAoi.cod : '');
  if (c.cfg && CFG.COL_DENOM > 0) hoja.getRange(c.fila, CFG.COL_DENOM).setValue(datosAoi ? datosAoi.den : '');
}

/* ------------------------------------------------------------------ */
/*  Hoja de consulta con todo el catálogo                              */
/* ------------------------------------------------------------------ */

function generarCatalogo() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.HOJA_CATALOGO);
  if (!sh) sh = ss.insertSheet(CFG.HOJA_CATALOGO);
  sh.clear();
  const filas = [['FORMULARIO', 'SIGLA', 'CENTRO DE COSTO', 'FACULTAD', 'OEI', 'AEI', 'COD. AOI', 'ACTIVIDAD OPERATIVA']];
  POI_FACULTADES.forEach(f => f.items.forEach(it => filas.push([
    f.formulario || '', f.sigla, f.cc, f.nombre, POI_OEI[oeiDeAei_(it[0])], POI_AEI[it[0]], it[1], it[2]])));
  sh.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  sh.getRange(1, 1, 1, filas[0].length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, filas[0].length);
  ss.toast('Catálogo generado: ' + (filas.length - 1) + ' actividades.', 'POI 2026', 5);
}
