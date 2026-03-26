/**
 * Formateador corporativo — Estilo minimalista
 * Función principal: formatearTodo()
 * v4.0 — fix total getV + rediseño Google/Notion style
 */

// ─────────────────────────────────────────────
// PALETA — Negro/blanco base + acentos corporativos
// ─────────────────────────────────────────────
const C = {
  // Fondo
  WHITE:      '#FFFFFF',
  BG_PAGE:    '#F9F9FB',   // fondo muy suave (tipo Notion)
  BG_ALT:     '#F3F2F8',   // fila alternada, casi blanco
  BG_HOVER:   '#EDEAF6',   // hover / highlight

  // Corporativos
  DARK:       '#1E1340',   // negro corporativo
  PURPLE:     '#9B35B5',
  PURPLE_D:   '#6B1E80',
  PURPLE_XL:  '#FAF5FF',
  RED:        '#E84040',
  RED_D:      '#A02828',
  RED_XL:     '#FFF5F5',
  INDIGO:     '#6355D4',
  INDIGO_D:   '#3B2D9A',
  INDIGO_XL:  '#F4F2FF',

  // Texto
  TEXT:       '#1E1340',   // texto principal, oscuro pero no negro puro
  TEXT_S:     '#4A3878',   // texto secundario
  TEXT_L:     '#9E94C0',   // texto terciario / notas

  // Bordes — muy sutiles
  BORDER:     '#E8E5F0',   // borde estándar
  BORDER_M:   '#C8BEE8',   // borde medio (solo exterior)

  // Gantt
  PT_BG:      '#1E1340',   // paquete: negro corp
  TASK_BG:    '#6B1E80',   // tarea: púrpura oscuro
  DEL_BG:     '#EAE6FF',   // entregable: fondo suave
  DEL_FG:     '#3B2D9A',

  PT_ON:      '#E84040',
  TASK_ON:    '#9B35B5',
  DEL_ON:     '#6355D4',
  PT_OFF:     '#F0EAEE',
  TASK_OFF:   '#EDE8F5',
  DEL_OFF:    '#F0EEF8',

  // Totales (gradiente índigo → suave)
  SUM_A:      '#1E1340',   // SUBTOTAL
  SUM_B:      '#3B2D9A',   // IVA
  SUM_C:      '#6355D4',   // TOTAL
  SUM_D:      '#9B35B5',   // BAJA

  CREAM:      '#FFFDF8',   // notas / comentarios
};

// ─────────────────────────────────────────────
// MENÚ
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Formato corporativo')
    .addItem('✦ Aplicar formato', 'formatearTodo')
    .addToUi();
}

// ─────────────────────────────────────────────
// FUNCIÓN PRINCIPAL
// ─────────────────────────────────────────────
function formatearTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    const n = norm(sheet.getName());
    if      (n.includes('datos previos'))  formatDatosPrevios(sheet);
    else if (n.includes('presupuesto'))    formatPresupuesto(sheet);
    else if (n.includes('gantt'))          formatGantt(sheet);
    else if (n.includes('sensibilidad'))   formatSensibilidad(sheet);
    else if (n.includes('licitacion'))     formatLicitaciones(sheet);
    else                                   formatGenerico(sheet);
  });
  SpreadsheetApp.getUi().alert(
    '✦ Formato aplicado',
    'El diseño corporativo se ha aplicado correctamente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─────────────────────────────────────────────
// UTILIDADES BASE
// ─────────────────────────────────────────────

function norm(s) {
  return String(s == null ? '' : s)
    .toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').trim();
}

function txt(v) {
  return String(v == null ? '' : v).trim();
}

/**
 * getV — Lee valor de celda. Usa getDisplayValue como fallback
 * para evitar la excepción diferida cuando la celda tiene tipo "Texto" (@).
 */
function getV(sheet, r, c) {
  try {
    const cell = sheet.getRange(r, c);
    let v;
    try { v = cell.getValue(); } catch(e1) { v = null; }
    // Si getValue lanzó null o excepción diferida, usamos displayValue
    if (v == null) {
      try { v = cell.getDisplayValue(); } catch(e2) { v = ''; }
    }
    return v;
  } catch(e) { return ''; }
}

function getT(sheet, r, c) {
  return txt(getV(sheet, r, c));
}

function hasAny(text, arr) {
  const t = norm(text);
  return arr.some(function(k) { return t.indexOf(norm(k)) !== -1; });
}

function isBlankRow(sheet, r, lc) {
  try {
    return sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].join('').trim() === '';
  } catch(e) { return false; }
}

// ─────────────────────────────────────────────
// PREPARACIÓN DE HOJA
// ─────────────────────────────────────────────

function breakMerges(sheet) {
  try {
    const lr = Math.max(sheet.getLastRow(), 1);
    const lc = Math.max(sheet.getLastColumn(), 1);
    sheet.getRange(1, 1, lr, lc).breakApart();
  } catch(e) {}
}

function resetSheet(sheet) {
  const lr = Math.max(sheet.getLastRow(), 1);
  const lc = Math.max(sheet.getLastColumn(), 1);
  const rg = sheet.getRange(1, 1, lr, lc);

  // Quitar tipo "Texto" (@) del xlsx — fallback columna a columna
  try {
    rg.setNumberFormat('General');
    SpreadsheetApp.flush();
  } catch(e1) {
    for (let c = 1; c <= lc; c++) {
      try {
        sheet.getRange(1, c, lr, 1).setNumberFormat('General');
        SpreadsheetApp.flush();
      } catch(e2) {}
    }
  }

  // Reset visual
  try { rg.setBackground(C.WHITE); }                                     catch(e) {}
  try { rg.setFontColor(C.TEXT); }                                        catch(e) {}
  try { rg.setFontFamily('Google Sans, Arial, sans-serif'); }             catch(e) {}
  try { rg.setFontSize(10); }                                             catch(e) {}
  try { rg.setFontWeight('normal'); }                                     catch(e) {}
  try { rg.setFontStyle('normal'); }                                      catch(e) {}
  try { rg.setHorizontalAlignment('left'); }                              catch(e) {}
  try { rg.setVerticalAlignment('middle'); }                              catch(e) {}
  try { rg.setWrap(false); }                                              catch(e) {}
  try { rg.setBorder(false,false,false,false,false,false); }              catch(e) {}
  try { SpreadsheetApp.flush(); }                                         catch(e) {}
}

function prepSheet(sheet) {
  breakMerges(sheet);
  resetSheet(sheet);
}

function freeze(sheet, rows, cols) {
  try { sheet.setFrozenRows(rows); }    catch(e) {}
  try { sheet.setFrozenColumns(cols); } catch(e) {}
}

// ─────────────────────────────────────────────
// FORMATO DE NÚMEROS — seguros
// ─────────────────────────────────────────────

function snf(cell, fmt) {
  try {
    let v;
    try { v = cell.getValue(); } catch(e) { return; }
    if (typeof v === 'number') {
      try { cell.setNumberFormat(fmt); } catch(e) {}
    }
  } catch(e) {}
}

function applyFmt(cell, label) {
  try {
    let v;
    try { v = cell.getValue(); } catch(e) { return; }
    if (typeof v !== 'number') return;

    const l = norm(label);
    let fmt = '€#,##0.00';

    if (hasAny(l, ['baja','beneficio','gastos generales','%','porcentaje','margen','peso criterio'])) {
      fmt = '0.00%';
    } else if (hasAny(l, ['pm','personas adscritas','precio hora','hora real','experiencia'])) {
      fmt = '#,##0.00';
    } else if (hasAny(l, ['horas','dias','días','duracion','duración','meses','licitadores','licitador','ofertas','mes_inicio','mes_final'])) {
      fmt = '#,##0';
    }

    try { cell.setNumberFormat(fmt); }          catch(e) {}
    try { cell.setHorizontalAlignment('right'); } catch(e) {}
  } catch(e) {}
}

// ─────────────────────────────────────────────
// PRIMITIVAS DE ESTILO — minimalista
// ─────────────────────────────────────────────

/**
 * Cabecera minimalista: fondo sólido, texto blanco, sin borde inferior llamativo.
 * Sin bordes interiores, solo un hairline inferior muy sutil.
 */
function hdr(range, bg) {
  try {
    range
      .setBackground(bg || C.DARK)
      .setFontColor(C.WHITE)
      .setFontWeight('bold')
      .setFontSize(9)
      .setLetterSpacing && range.setLetterSpacing(1);
  } catch(e) {}
  try { range.setFontSize(9); }                                           catch(e) {}
  try { range.setVerticalAlignment('middle'); }                           catch(e) {}
  try { range.setHorizontalAlignment('left'); }                           catch(e) {}
  // Borde inferior: solo una línea fina del color del fondo ligeramente más oscuro
  try {
    range.setBorder(false, false, true, false, false, false,
      C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}
}

/**
 * Fila alternada: fondo muy suave, borde inferior hairline.
 */
function zebra(range, idx) {
  try { range.setBackground(idx % 2 === 0 ? C.BG_ALT : C.WHITE); } catch(e) {}
  try { range.setFontColor(C.TEXT); }                                catch(e) {}
  divider(range);
}

/**
 * Línea divisoria: borde inferior de 1px muy tenue.
 */
function divider(range) {
  try {
    range.setBorder(false, false, true, false, false, false,
      C.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}
}

/**
 * Borde exterior: solo el perímetro, SOLID_MEDIUM.
 */
function outerBorder(range) {
  try {
    range.setBorder(true, true, true, true, false, false,
      C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch(e) {}
}

/**
 * Fila de total/resumen: fondo sólido, texto blanco, alta jerarquía visual.
 */
function sumRow(range, bg, fontSize) {
  try {
    range
      .setBackground(bg)
      .setFontColor(C.WHITE)
      .setFontWeight('bold')
      .setFontSize(fontSize || 10)
      .setVerticalAlignment('middle');
  } catch(e) {}
}

/**
 * Pill / badge para etiquetas de estado.
 */
function colorEstadoCell(cell, estado) {
  const e = norm(estado);
  try {
    if (e === 'resuelta') {
      cell.setBackground(C.RED_XL).setFontColor(C.RED_D).setFontWeight('bold').setHorizontalAlignment('center');
    } else if (e === 'adjudicada') {
      cell.setBackground(C.INDIGO_XL).setFontColor(C.INDIGO_D).setFontWeight('bold').setHorizontalAlignment('center');
    } else if (e === 'publicada') {
      cell.setBackground(C.PURPLE_XL).setFontColor(C.PURPLE_D).setFontWeight('bold').setHorizontalAlignment('center');
    } else if (e !== '') {
      cell.setFontWeight('bold').setHorizontalAlignment('center');
    }
  } catch(e) {}
}

function findHeaderRow(sheet, needles, maxRows) {
  const lr = Math.min(sheet.getLastRow(), maxRows || 10);
  const lc = Math.max(sheet.getLastColumn(), 1);
  for (let r = 1; r <= lr; r++) {
    try {
      const rowText = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].map(norm).join(' | ');
      if (needles.every(function(k) { return rowText.indexOf(norm(k)) !== -1; })) return r;
    } catch(e) {}
  }
  return 1;
}

// ─────────────────────────────────────────────
// HOJA 1 — DATOS PREVIOS
// ─────────────────────────────────────────────
function formatDatosPrevios(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.PURPLE);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  // Fondo de página suave
  try { sheet.getRange(1, 1, lr, lc).setBackground(C.WHITE); } catch(e) {}

  const leftHdr     = ['licitacion subtotal','costes personal'];
  const leftTypes   = ['plantilla ingeniero superior','plantilla ingeniero','plantilla otro perfil','subcontratado','dieta'];
  const leftMetrics = ['horas totales','pm','meses de trabajo','personas adscritas','coste personal total','coste no personal'];
  const rightHdr    = ['costes dietas'];
  const rightTypes  = ['alquiler coche','gasolina','desplazamiento'];
  const rightMetrics= ['precio ofertado','precio con iva','baja','beneficio industrial','gastos generales',
                       'baja promedio','baja segun convenio','baja según convenio','baja maxima','baja máxima','anormal','licitador'];

  for (let r = 1; r <= lr; r++) {
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 10); continue; }
    sheet.setRowHeight(r, 26);

    const a    = getT(sheet, r, 1);
    const d    = lc >= 4 ? getT(sheet, r, 4) : '';
    const e    = lc >= 5 ? getT(sheet, r, 5) : '';

    // Comentarios largos
    if (a.length > 60) {
      try {
        sheet.getRange(r, 1, 1, lc)
          .setBackground(C.CREAM).setFontColor(C.TEXT_L)
          .setFontStyle('italic').setFontSize(9).setWrap(true);
        sheet.setRowHeight(r, 48);
      } catch(ex) {}
      continue;
    }

    // Zona izquierda
    if (hasAny(a, leftHdr)) {
      hdr(sheet.getRange(r, 1, 1, Math.min(3, lc)), C.DARK);
      try { sheet.getRange(r,1).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
      sheet.setRowHeight(r, 32);
    } else if (hasAny(a, leftTypes)) {
      try {
        sheet.getRange(r, 1, 1, Math.min(3, lc)).setBackground(C.WHITE);
        sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.DARK).setHorizontalAlignment('right');
        if (lc>=2) applyFmt(sheet.getRange(r,2), a);
        if (lc>=3) { sheet.getRange(r,3).setFontColor(C.TEXT_S); applyFmt(sheet.getRange(r,3),'precio hora'); }
        divider(sheet.getRange(r, 1, 1, Math.min(3, lc)));
      } catch(ex) {}
    } else if (hasAny(a, leftMetrics)) {
      try {
        const rg = sheet.getRange(r, 1, 1, Math.min(3, lc));
        rg.setBackground(C.BG_HOVER);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.INDIGO_D).setHorizontalAlignment('right');
        if (lc>=2) { applyFmt(sheet.getRange(r,2), a); sheet.getRange(r,2).setFontWeight('bold').setFontColor(C.DARK); }
        divider(rg);
      } catch(ex) {}
    } else if (a) {
      try {
        sheet.getRange(r, 1).setFontColor(C.TEXT_S);
        divider(sheet.getRange(r, 1, 1, Math.min(3, lc)));
      } catch(ex) {}
    }

    // Zona derecha (col D+)
    if (lc >= 4) {
      if (hasAny(d, rightHdr)) {
        hdr(sheet.getRange(r, 4, 1, lc-3), C.DARK);
        try { sheet.getRange(r,4).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
        sheet.setRowHeight(r, 32);
      } else if (hasAny(d, rightTypes)) {
        try {
          sheet.getRange(r,4).setFontWeight('bold').setFontColor(C.DARK).setHorizontalAlignment('right');
          if (lc>=5) applyFmt(sheet.getRange(r,5), d);
          divider(sheet.getRange(r, 4, 1, Math.max(1, lc-3)));
        } catch(ex) {}
      } else if (hasAny(e, rightMetrics) || hasAny(d, rightMetrics)) {
        const lCol = hasAny(e, rightMetrics) ? 5 : 4;
        const vCol = lCol + 1;
        try {
          sheet.getRange(r, lCol, 1, lc - lCol + 1).setBackground(C.BG_HOVER);
          sheet.getRange(r, lCol).setFontWeight('bold').setFontColor(C.INDIGO_D).setHorizontalAlignment('right');
          if (vCol <= lc) {
            sheet.getRange(r, vCol).setFontWeight('bold').setFontColor(C.DARK);
            applyFmt(sheet.getRange(r, vCol), lCol===5 ? e : d);
          }
          if (vCol+1 <= lc) {
            sheet.getRange(r, vCol+1).setFontColor(C.TEXT_L).setFontStyle('italic').setFontSize(9).setWrap(true);
          }
          divider(sheet.getRange(r, lCol, 1, lc - lCol + 1));
        } catch(ex) {}
      } else if (d) {
        try {
          sheet.getRange(r,4).setFontColor(C.TEXT_S);
          if (lc>=5) applyFmt(sheet.getRange(r,5), d);
        } catch(ex) {}
      }
    }
  }

  // Anchos de columna
  if (lc>=1) sheet.setColumnWidth(1, 220);
  if (lc>=2) sheet.setColumnWidth(2, 110);
  if (lc>=3) sheet.setColumnWidth(3, 90);
  if (lc>=4) sheet.setColumnWidth(4, 210);
  if (lc>=5) sheet.setColumnWidth(5, 110);
  if (lc>=6) sheet.setColumnWidth(6, 110);
  if (lc>=7) sheet.setColumnWidth(7, 230);
  if (lc>=8) sheet.setColumnWidth(8, 180);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// HOJA 2 — PRESUPUESTO
// ─────────────────────────────────────────────
function formatPresupuesto(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.RED);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const hRow = findHeaderRow(sheet, ['actividad','descripcion'], 8);
  hdr(sheet.getRange(hRow, 1, 1, lc), C.DARK);
  sheet.setRowHeight(hRow, 36);
  // Alineaciones cabecera
  try {
    sheet.getRange(hRow,1).setHorizontalAlignment('left');
    sheet.getRange(hRow,2).setHorizontalAlignment('left');
    if (lc>=3) sheet.getRange(hRow,3).setHorizontalAlignment('center');
    if (lc>=4) sheet.getRange(hRow,4).setHorizontalAlignment('left');
    for (let c=5; c<=lc; c++) sheet.getRange(hRow,c).setHorizontalAlignment('right');
  } catch(ex) {}

  let zi = 0;
  for (let r = hRow+1; r <= lr; r++) {
    const a   = getT(sheet, r, 1);
    const n   = norm(a);
    const row = sheet.getRange(r, 1, 1, lc);

    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 10); continue; }
    sheet.setRowHeight(r, 26);

    if (n === 'subtotal') {
      sumRow(row, C.SUM_A); sheet.setRowHeight(r, 30);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) applyFmt(sheet.getRange(r,9), 'euros');
    } else if (n === 'iva') {
      sumRow(row, C.SUM_B); sheet.setRowHeight(r, 28);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) { try { sheet.getRange(r,9).setNumberFormat('0.00%'); } catch(ex){} }
    } else if (n === 'total') {
      sumRow(row, C.SUM_C, 11); sheet.setRowHeight(r, 32);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) applyFmt(sheet.getRange(r,9), 'euros');
    } else if (n.indexOf('baja') !== -1) {
      sumRow(row, C.SUM_D); sheet.setRowHeight(r, 28);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) applyFmt(sheet.getRange(r,9), '% baja');
    } else {
      zi++;
      try {
        row.setBackground(zi%2===0 ? C.BG_ALT : C.WHITE).setFontColor(C.TEXT);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.DARK).setWrap(true);
        sheet.getRange(r,2).setFontColor(C.TEXT_S).setFontStyle('italic').setFontSize(9).setWrap(true);
        if (lc>=3) sheet.getRange(r,3).setHorizontalAlignment('center').setFontColor(C.TEXT_L);
        if (lc>=4) sheet.getRange(r,4).setFontColor(C.TEXT_S);
        if (lc>=5) { sheet.getRange(r,5).setHorizontalAlignment('right'); snf(sheet.getRange(r,5),'#,##0'); }
        if (lc>=6) { sheet.getRange(r,6).setHorizontalAlignment('right'); snf(sheet.getRange(r,6),'#,##0'); }
        if (lc>=7) { sheet.getRange(r,7).setHorizontalAlignment('right'); snf(sheet.getRange(r,7),'€#,##0.00'); }
        if (lc>=8) { sheet.getRange(r,8).setHorizontalAlignment('right'); snf(sheet.getRange(r,8),'€#,##0.00'); }
        if (lc>=9) { sheet.getRange(r,9).setHorizontalAlignment('right').setFontWeight('bold'); snf(sheet.getRange(r,9),'€#,##0.00'); }
      } catch(ex) {}
    }
    divider(row);
  }

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 340);
  if (lc>=3) sheet.setColumnWidth(3, 120);
  if (lc>=4) sheet.setColumnWidth(4, 170);
  if (lc>=5) sheet.setColumnWidth(5, 90);
  if (lc>=6) sheet.setColumnWidth(6, 100);
  if (lc>=7) sheet.setColumnWidth(7, 110);
  if (lc>=8) sheet.setColumnWidth(8, 120);
  if (lc>=9) sheet.setColumnWidth(9, 130);

  outerBorder(sheet.getRange(hRow, 1, lr-hRow+1, lc));
  freeze(sheet, hRow, 2);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// HOJA 3 — GANTT
// ─────────────────────────────────────────────
function formatGantt(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.INDIGO);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const META = 8;
  const TL   = 9;

  // Detectar primera fila de datos reales
  let firstData = 4;
  for (let r = 1; r <= Math.min(lr, 20); r++) {
    const code = getT(sheet, r, 2);
    if (/^(PT\d+|A\d+(\.\d+)*|E\d+(\.\d+)*)$/i.test(code)) { firstData = r; break; }
  }

  // Cabeceras timeline
  for (let r = 1; r < firstData; r++) {
    try {
      const vals   = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].map(txt);
      const joined = vals.map(norm).join(' | ');
      const isYear  = vals.some(function(v){ return /^20\d{2}$/.test(v); });
      const isMonth = vals.some(function(v){ return /^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)$/i.test(v); });
      const isNum   = vals.some(function(v){ return /^\d+$/.test(v); });
      const isMeta  = joined.indexOf('lider')!== -1 || joined.indexOf('mes_inicio')!==-1;

      let bg=C.INDIGO_D, h=20, fs=8;
      if      (isYear)  { bg=C.DARK;     h=28; fs=10; }
      else if (isMeta)  { bg=C.DARK;     h=30; fs=9;  }
      else if (isMonth) { bg=C.INDIGO_D; h=20; fs=8;  }
      else if (isNum)   { bg=C.INDIGO;   h=16; fs=7;  }

      sheet.setRowHeight(r, h);
      hdr(sheet.getRange(r, 1, 1, Math.min(META,lc)), bg);
      try { sheet.getRange(r, 1, 1, Math.min(META,lc)).setFontSize(fs).setHorizontalAlignment('center'); } catch(ex){}
      if (lc>=TL) {
        hdr(sheet.getRange(r, TL, 1, lc-TL+1), bg);
        try { sheet.getRange(r, TL, 1, lc-TL+1).setFontSize(fs).setHorizontalAlignment('center'); } catch(ex){}
      }
    } catch(ex) {}
  }

  // Filas de datos
  for (let r = firstData; r <= lr; r++) {
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 8); continue; }

    const code = getT(sheet, r, 2);
    let metaBg=C.WHITE, metaFg=C.TEXT, metaW='normal', metaFs=10;
    let actColor=C.INDIGO, inactColor=C.BG_ALT;

    if (/^PT\d+$/i.test(code)) {
      metaBg=C.PT_BG;   metaFg=C.WHITE; metaW='bold'; metaFs=10;
      actColor=C.PT_ON; inactColor=C.PT_OFF;
      sheet.setRowHeight(r, 28);
    } else if (/^A\d+(\.\d+)*$/i.test(code)) {
      metaBg=C.TASK_BG; metaFg=C.WHITE; metaW='bold'; metaFs=10;
      actColor=C.TASK_ON; inactColor=C.TASK_OFF;
      sheet.setRowHeight(r, 24);
    } else if (/^E\d+(\.\d+)*$/i.test(code)) {
      metaBg=C.DEL_BG; metaFg=C.DEL_FG; metaW='normal'; metaFs=9;
      actColor=C.DEL_ON; inactColor=C.DEL_OFF;
      sheet.setRowHeight(r, 22);
    } else {
      sheet.setRowHeight(r, 22);
    }

    try {
      sheet.getRange(r, 1, 1, Math.min(META,lc))
        .setBackground(metaBg).setFontColor(metaFg)
        .setFontWeight(metaW).setFontSize(metaFs);
      if (lc>=1) sheet.getRange(r,1).setHorizontalAlignment('center');
      if (lc>=2) sheet.getRange(r,2).setHorizontalAlignment('center');
      if (lc>=3) sheet.getRange(r,3).setHorizontalAlignment('left').setWrap(true);
      if (lc>=4) sheet.getRange(r,4).setHorizontalAlignment('center');
      if (lc>=5) sheet.getRange(r,5).setHorizontalAlignment('left').setFontSize(8).setWrap(true);
      if (lc>=6) sheet.getRange(r,6).setHorizontalAlignment('center'); snf(sheet.getRange(r,6),'#,##0');
      if (lc>=7) sheet.getRange(r,7).setHorizontalAlignment('center'); snf(sheet.getRange(r,7),'#,##0');
      if (lc>=8) sheet.getRange(r,8).setHorizontalAlignment('center'); snf(sheet.getRange(r,8),'#,##0');
    } catch(ex) {}

    // Timeline
    if (lc >= TL) {
      for (let c = TL; c <= lc; c++) {
        try {
          const cell = sheet.getRange(r, c);
          let v; try { v = cell.getValue(); } catch(ex) { v = 0; }
          const on = (v===1||v==='1');
          cell.setBackground(on ? actColor : inactColor)
              .setFontColor(on ? actColor : inactColor)
              .setFontSize(6).setHorizontalAlignment('center')
              .setFontWeight(on ? 'bold' : 'normal');
        } catch(ex) {}
      }
    }
    divider(sheet.getRange(r, 1, 1, lc));
  }

  // Anchos
  if (lc>=1) sheet.setColumnWidth(1, 20);
  if (lc>=2) sheet.setColumnWidth(2, 65);
  if (lc>=3) sheet.setColumnWidth(3, 340);
  if (lc>=4) sheet.setColumnWidth(4, 90);
  if (lc>=5) sheet.setColumnWidth(5, 200);
  if (lc>=6) sheet.setColumnWidth(6, 65);
  if (lc>=7) sheet.setColumnWidth(7, 65);
  if (lc>=8) sheet.setColumnWidth(8, 65);
  for (let c=TL; c<=lc; c++) sheet.setColumnWidth(c, 22);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, Math.max(firstData-1, 1), META);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// HOJA 4 — ANÁLISIS DE SENSIBILIDAD
// ─────────────────────────────────────────────
function formatSensibilidad(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.INDIGO);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  for (let r = 1; r <= lr; r++) {
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 10); continue; }
    sheet.setRowHeight(r, 24);

    const a  = getT(sheet, r, 1);
    const b  = lc>=2 ? getT(sheet, r, 2) : '';
    const row = sheet.getRange(r, 1, 1, lc);
    const aN = norm(a);
    const bN = norm(b);

    // Sub-cabecera de bloque (col A vacía, col B tiene perfil)
    if (a==='' && (bN.indexOf('plantilla')!==-1 || bN.indexOf('ingeniero')!==-1)) {
      if (lc>=2) { hdr(sheet.getRange(r, 2, 1, Math.min(3, lc-1)), C.INDIGO_D); }
      sheet.setRowHeight(r, 30); continue;
    }

    // Cabecera de sección
    if (hasAny(a, ['categoria profesional','tabla salarial','nivel salarial','baja desproporcionada'])) {
      hdr(row, C.DARK);
      try { sheet.getRange(r,1).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
      sheet.setRowHeight(r, 32); continue;
    }

    // Año como título
    if (/^(ano|año)?\s*\d{4}$/.test(aN)) {
      try {
        row.setBackground(C.WHITE).setFontColor(C.INDIGO_D)
           .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left');
        sheet.setRowHeight(r, 36);
      } catch(ex){}
      continue;
    }

    // Notas
    if (aN.indexOf('nota:')===0 || aN[0]==='*') {
      try {
        row.setBackground(C.CREAM).setFontColor(C.TEXT_L)
           .setFontStyle('italic').setFontSize(8).setWrap(true);
        sheet.setRowHeight(r, 32);
      } catch(ex){}
      continue;
    }

    // Filas de licitadores
    if (/^\d+\s+licitador/.test(aN)) {
      try {
        row.setBackground(C.RED_XL);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.RED_D).setHorizontalAlignment('right');
        if (lc>=2) applyFmt(sheet.getRange(r,2),'euros');
        if (lc>=3) applyFmt(sheet.getRange(r,3),'% baja');
        divider(row);
      } catch(ex){}
      continue;
    }

    // Métricas resumen
    if (hasAny(a, ['coste personal total','coste no personal','total + bi','precio ofertado','precio con iva','baja'])) {
      try {
        row.setBackground(C.BG_HOVER).setFontColor(C.INDIGO_D).setFontWeight('bold');
        sheet.getRange(r,1).setHorizontalAlignment('right');
        for (let c=2; c<=lc; c++) applyFmt(sheet.getRange(r,c), a);
        divider(row);
      } catch(ex){}
      continue;
    }

    // Perfiles
    if (hasAny(a, ['plantilla ingeniero superior','plantilla ingeniero','plantilla otro perfil','subcontratado'])) {
      try {
        row.setBackground(r%2===0 ? C.BG_ALT : C.WHITE);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.DARK);
        for (let c=2; c<=lc; c++) {
          let v; try { v=sheet.getRange(r,c).getValue(); } catch(ex){ v=null; }
          if (typeof v==='number') applyFmt(sheet.getRange(r,c), a);
        }
        divider(row);
      } catch(ex){}
      continue;
    }

    // Fila genérica
    try {
      row.setBackground(r%2===0 ? C.BG_ALT : C.WHITE).setFontColor(C.TEXT);
      sheet.getRange(r,1).setFontColor(C.TEXT_S);
      for (let c=2; c<=lc; c++) applyFmt(sheet.getRange(r,c), a);
      divider(row);
    } catch(ex) {}
  }

  // Notas col J
  if (lc>=10) {
    for (let r=1; r<=lr; r++) {
      const j = getT(sheet, r, 10);
      if (j.length>20) {
        try {
          sheet.getRange(r,10).setBackground(C.CREAM).setFontColor(C.TEXT_L)
            .setFontStyle('italic').setFontSize(8).setWrap(true);
          sheet.setRowHeight(r, Math.max(sheet.getRowHeight(r), 36));
        } catch(ex){}
      }
    }
    sheet.setColumnWidth(10, 260);
  }

  if (lc>=1) sheet.setColumnWidth(1, 240);
  if (lc>=2) sheet.setColumnWidth(2, 130);
  if (lc>=3) sheet.setColumnWidth(3, 130);
  if (lc>=4) sheet.setColumnWidth(4, 130);
  if (lc>=5) sheet.setColumnWidth(5, 150);
  if (lc>=6) sheet.setColumnWidth(6, 160);
  if (lc>=7) sheet.setColumnWidth(7, 110);
  if (lc>=8) sheet.setColumnWidth(8, 110);
  if (lc>=9) sheet.setColumnWidth(9, 120);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 2, 1);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// HOJA 5 — LICITACIONES PREVIAS
// ─────────────────────────────────────────────
function formatLicitaciones(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.RED);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const hRow = findHeaderRow(sheet, ['expediente','objeto'], 6);
  hdr(sheet.getRange(hRow, 1, 1, lc), C.DARK);
  sheet.setRowHeight(hRow, 36);

  // Detectar columnas especiales
  const headers = sheet.getRange(hRow, 1, 1, lc).getDisplayValues()[0].map(norm);
  let colEstado=0;
  const moneyCols=[], pctCols=[];
  headers.forEach(function(h,i) {
    const c=i+1;
    if (h==='estado'||h.indexOf('estado')!==-1) colEstado=c;
    if (h.indexOf('presupuesto')!==-1||h.indexOf('valor estimado')!==-1||h.indexOf('importe adjudicado')!==-1) moneyCols.push(c);
    if (h.indexOf('%')!==-1||h.indexOf('baja')!==-1||h.indexOf('peso criterio')!==-1) pctCols.push(c);
  });

  let zi=0;
  for (let r=hRow+1; r<=lr; r++) {
    const row = sheet.getRange(r, 1, 1, lc);
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 10); continue; }
    zi++;
    try {
      row.setBackground(zi%2===0 ? C.BG_ALT : C.WHITE).setFontColor(C.TEXT);
      sheet.setRowHeight(r, 24);
      moneyCols.forEach(function(c){ applyFmt(sheet.getRange(r,c),'euros'); });
      pctCols.forEach(function(c){   applyFmt(sheet.getRange(r,c),'%');     });
      if (colEstado>0) colorEstadoCell(sheet.getRange(r,colEstado), getT(sheet,r,colEstado));
      divider(row);
    } catch(ex) {}
  }

  if (lc>=1)  sheet.setColumnWidth(1, 120);
  if (lc>=2)  sheet.setColumnWidth(2, 420);
  if (lc>=3)  sheet.setColumnWidth(3, 280);
  if (lc>=4)  sheet.setColumnWidth(4, 240);
  if (lc>=5)  sheet.setColumnWidth(5, 120);
  if (lc>=9)  sheet.setColumnWidth(9, 130);
  if (lc>=10) sheet.setColumnWidth(10, 130);
  if (lc>=16) sheet.setColumnWidth(16, 130);
  if (lc>=17) sheet.setColumnWidth(17, 100);
  if (lc>=18) sheet.setColumnWidth(18, 220);

  outerBorder(sheet.getRange(hRow, 1, lr-hRow+1, lc));
  freeze(sheet, hRow, 2);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// FALLBACK GENÉRICO
// ─────────────────────────────────────────────
function formatGenerico(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.PURPLE);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  hdr(sheet.getRange(1, 1, 1, lc), C.DARK);
  sheet.setRowHeight(1, 34);

  let zi=0;
  for (let r=2; r<=lr; r++) {
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 10); continue; }
    zi++;
    try {
      sheet.getRange(r,1,1,lc)
        .setBackground(zi%2===0 ? C.BG_ALT : C.WHITE)
        .setFontColor(C.TEXT);
      divider(sheet.getRange(r,1,1,lc));
      sheet.setRowHeight(r, 24);
    } catch(ex) {}
  }

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}
