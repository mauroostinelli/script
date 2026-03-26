/**
 * Formateador corporativo — Estilo minimalista
 * Función principal: formatearTodo()
 * v5.0 — Sistema celda editable (amarillo) vs calculada (fórmula) según SOP
 *
 * SISTEMA DE COLORES (basado en el SOP del equipo):
 *   🟡 INPUT_BG (#FFF9C4)  — Celda EDITABLE por el usuario (datos de entrada)
 *   🟣 CALC_BG  (#EDEAF6)  — Celda con FÓRMULA / calculada automáticamente (solo lectura visual)
 *   ⬛ DARK    (#1E1340)  — Cabeceras de sección
 *   ⚪ WHITE   (#FFFFFF)  — Filas normales / alternadas
 */

const C = {
  WHITE:      '#FFFFFF',
  BG_ALT:     '#F5F4FB',
  BG_HOVER:   '#EDEAF6',

  // Sistema de celdas según SOP
  INPUT_BG:   '#FFF9C4',   // 🟡 amarillo suave — celda EDITABLE (el usuario rellena)
  INPUT_FG:   '#3D3000',   // texto oscuro sobre amarillo
  INPUT_BD:   '#E6D400',   // borde amarillo intenso para resaltar el campo editable
  CALC_BG:    '#EDEAF6',   // 🟣 lila suave — celda CALCULADA (fórmula, no tocar)
  CALC_FG:    '#3B2D9A',   // texto índigo sobre calculada

  // Corporativos
  DARK:       '#1E1340',
  PURPLE:     '#9B35B5',
  PURPLE_D:   '#6B1E80',
  PURPLE_XL:  '#FAF5FF',
  RED:        '#E84040',
  RED_D:      '#A02828',
  RED_XL:     '#FFF5F5',
  INDIGO:     '#6355D4',
  INDIGO_D:   '#3B2D9A',
  INDIGO_XL:  '#F4F2FF',

  TEXT:       '#1E1340',
  TEXT_S:     '#4A3878',
  TEXT_L:     '#9E94C0',

  BORDER:     '#E8E5F0',
  BORDER_M:   '#C8BEE8',

  // Gantt
  PT_BG:      '#1E1340',
  TASK_BG:    '#6B1E80',
  DEL_BG:     '#EAE6FF',
  DEL_FG:     '#3B2D9A',
  PT_ON:      '#E84040',
  TASK_ON:    '#9B35B5',
  DEL_ON:     '#6355D4',
  PT_OFF:     '#F0EAEE',
  TASK_OFF:   '#EDE8F5',
  DEL_OFF:    '#F0EEF8',

  SUM_A:      '#1E1340',
  SUM_B:      '#3B2D9A',
  SUM_C:      '#6355D4',
  SUM_D:      '#9B35B5',

  CREAM:      '#FFFDF8',
};

// ─── MENÚ
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Formato corporativo')
    .addItem('✦ Aplicar formato', 'formatearTodo')
    .addToUi();
}

// ─── PRINCIPAL
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
    'El diseño corporativo se ha aplicado.\n\n' +
    '🟡 Amarillo = celda a rellenar por el usuario\n' +
    '🟣 Lila = celda calculada automáticamente',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─── UTILIDADES

function norm(s) {
  return String(s == null ? '' : s)
    .toLowerCase().normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '').trim();
}

function txt(v) { return String(v == null ? '' : v).trim(); }

/** Lee valor de celda con fallback a getDisplayValue para evitar excepción diferida */
function getV(sheet, r, c) {
  try {
    const cell = sheet.getRange(r, c);
    let v;
    try { v = cell.getValue(); } catch(e1) { v = null; }
    if (v == null) { try { v = cell.getDisplayValue(); } catch(e2) { v = ''; } }
    return v;
  } catch(e) { return ''; }
}

function getT(sheet, r, c) { return txt(getV(sheet, r, c)); }

/** Devuelve true si la celda contiene una fórmula (es calculada, no editable) */
function isFormula(sheet, r, c) {
  try {
    const f = sheet.getRange(r, c).getFormula();
    return f !== null && f.length > 0;
  } catch(e) { return false; }
}

function hasAny(text, arr) {
  const t = norm(text);
  return arr.some(function(k) { return t.indexOf(norm(k)) !== -1; });
}

function isBlankRow(sheet, r, lc) {
  try { return sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].join('').trim() === ''; }
  catch(e) { return false; }
}

// ─── PREPARACIÓN

function breakMerges(sheet) {
  try { sheet.getRange(1, 1, Math.max(sheet.getLastRow(),1), Math.max(sheet.getLastColumn(),1)).breakApart(); } catch(e) {}
}

function resetSheet(sheet) {
  const lr = Math.max(sheet.getLastRow(), 1);
  const lc = Math.max(sheet.getLastColumn(), 1);
  const rg = sheet.getRange(1, 1, lr, lc);

  try { rg.setNumberFormat('General'); SpreadsheetApp.flush(); } catch(e1) {
    for (let c = 1; c <= lc; c++) {
      try { sheet.getRange(1, c, lr, 1).setNumberFormat('General'); SpreadsheetApp.flush(); } catch(e2) {}
    }
  }
  try { rg.setBackground(C.WHITE); }                        catch(e) {}
  try { rg.setFontColor(C.TEXT); }                          catch(e) {}
  try { rg.setFontFamily('Arial'); }                        catch(e) {}
  try { rg.setFontSize(10); }                               catch(e) {}
  try { rg.setFontWeight('normal'); }                       catch(e) {}
  try { rg.setFontStyle('normal'); }                        catch(e) {}
  try { rg.setHorizontalAlignment('left'); }                catch(e) {}
  try { rg.setVerticalAlignment('middle'); }                catch(e) {}
  try { rg.setWrap(false); }                                catch(e) {}
  try { rg.setBorder(false,false,false,false,false,false); } catch(e) {}
  try { SpreadsheetApp.flush(); }                           catch(e) {}
}

function prepSheet(sheet) { breakMerges(sheet); resetSheet(sheet); }

function freeze(sheet, rows, cols) {
  try { sheet.setFrozenRows(rows); }    catch(e) {}
  try { sheet.setFrozenColumns(cols); } catch(e) {}
}

// ─── FORMATOS NUMÉRICOS

function snf(cell, fmt) {
  try {
    let v; try { v = cell.getValue(); } catch(e) { return; }
    if (typeof v === 'number') { try { cell.setNumberFormat(fmt); } catch(e) {} }
  } catch(e) {}
}

function applyFmt(cell, label) {
  try {
    let v; try { v = cell.getValue(); } catch(e) { return; }
    if (typeof v !== 'number') return;
    const l = norm(label);
    let fmt = '€#,##0.00';
    if (hasAny(l, ['baja','beneficio','gastos generales','%','porcentaje','margen','peso criterio'])) fmt = '0.00%';
    else if (hasAny(l, ['pm','personas adscritas','precio hora','hora real','experiencia'])) fmt = '#,##0.00';
    else if (hasAny(l, ['horas','dias','días','duracion','duración','meses','licitadores','licitador','mes_inicio','mes_final'])) fmt = '#,##0';
    try { cell.setNumberFormat(fmt); }           catch(e) {}
    try { cell.setHorizontalAlignment('right'); } catch(e) {}
  } catch(e) {}
}

// ─── PRIMITIVAS DE ESTILO

function hdr(range, bg) {
  try { range.setBackground(bg||C.DARK).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle').setHorizontalAlignment('left'); } catch(e) {}
  try { range.setBorder(false,false,true,false,false,false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
}

function divider(range) {
  try { range.setBorder(false,false,true,false,false,false, C.BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
}

function outerBorder(range) {
  try { range.setBorder(true,true,true,true,false,false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID_MEDIUM); } catch(e) {}
}

function sumRow(range, bg, fontSize) {
  try { range.setBackground(bg).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(fontSize||10).setVerticalAlignment('middle'); } catch(e) {}
}

/**
 * inputCell — Marca una celda como EDITABLE por el usuario.
 * Fondo amarillo suave + borde amarillo inferior para destacar el campo.
 */
function inputCell(cell) {
  try {
    cell.setBackground(C.INPUT_BG).setFontColor(C.INPUT_FG).setFontWeight('bold');
    cell.setBorder(false, false, true, false, false, false, C.INPUT_BD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch(e) {}
}

/**
 * calcCell — Marca una celda como CALCULADA (fórmula, no editar).
 * Fondo lila muy suave + texto índigo.
 */
function calcCell(cell) {
  try {
    cell.setBackground(C.CALC_BG).setFontColor(C.CALC_FG).setFontWeight('bold');
    cell.setBorder(false, false, true, false, false, false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}
}

/**
 * markValueCells — Recorre un rango de celdas de valor y aplica
 * inputCell si la celda tiene valor literal, o calcCell si tiene fórmula.
 * Acepta array de columnas o rango simple.
 */
function markValueCells(sheet, r, cols) {
  cols.forEach(function(c) {
    try {
      const cell = sheet.getRange(r, c);
      let v; try { v = cell.getValue(); } catch(ex) { return; }
      if (v === '' || v === null) return; // celda vacía: no tocar
      if (isFormula(sheet, r, c)) {
        calcCell(cell);
      } else {
        inputCell(cell);
      }
    } catch(ex) {}
  });
}

function colorEstadoCell(cell, estado) {
  const e = norm(estado);
  try {
    if (e === 'resuelta')   { cell.setBackground(C.RED_XL).setFontColor(C.RED_D).setFontWeight('bold').setHorizontalAlignment('center'); }
    else if (e === 'adjudicada') { cell.setBackground(C.INDIGO_XL).setFontColor(C.INDIGO_D).setFontWeight('bold').setHorizontalAlignment('center'); }
    else if (e === 'publicada')  { cell.setBackground(C.PURPLE_XL).setFontColor(C.PURPLE_D).setFontWeight('bold').setHorizontalAlignment('center'); }
    else if (e !== '')           { cell.setFontWeight('bold').setHorizontalAlignment('center'); }
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
//
// SEGÚN SOP, campos editables en esta hoja:
//   → Presupuesto licitación subtotal (B1)
//   → Costes por hora de cada perfil (C3:C6)
//   → Coste dieta por día (col E: alquiler coche, gasolina, desplazamiento)
//   → Valores para viajes (col F)
//   → Comentarios (fila 15+)
// Campos calculados (fórmula):
//   → Todo lo demás: horas totales, PM, coste personal total, precio ofertado, etc.
// ─────────────────────────────────────────────
function formatDatosPrevios(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.PURPLE);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

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

    const a = getT(sheet, r, 1);
    const d = lc >= 4 ? getT(sheet, r, 4) : '';
    const e = lc >= 5 ? getT(sheet, r, 5) : '';

    // Comentarios largos (texto libre — editable)
    if (a.length > 60) {
      try {
        sheet.getRange(r, 1, 1, lc)
          .setBackground(C.INPUT_BG).setFontColor(C.INPUT_FG)
          .setFontStyle('italic').setFontSize(9).setWrap(true);
        sheet.setRowHeight(r, 52);
      } catch(ex) {}
      continue;
    }

    // ─ Zona izquierda ─
    if (hasAny(a, leftHdr)) {
      // Cabecera de sección
      hdr(sheet.getRange(r, 1, 1, Math.min(3, lc)), C.DARK);
      try { sheet.getRange(r,1).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
      // B1 = Presupuesto licitación: el valor es EDITABLE (lo ingresa el usuario según SOP)
      if (lc >= 2 && a.toLowerCase().includes('licitacion') || a.toLowerCase().includes('licitación') || a.toLowerCase().includes('subtotal')) {
        try {
          const cellB = sheet.getRange(r, 2);
          inputCell(cellB);
          applyFmt(cellB, 'euros');
          cellB.setHorizontalAlignment('right');
        } catch(ex) {}
      }
      sheet.setRowHeight(r, 32);

    } else if (hasAny(a, leftTypes)) {
      // Filas de perfil: col B = coste/hora (EDITABLE), col C = precio hora (CALCULADO)
      try {
        sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.DARK).setHorizontalAlignment('right');
        if (lc >= 2) {
          const cellB = sheet.getRange(r, 2);
          if (isFormula(sheet, r, 2)) { calcCell(cellB); } else { inputCell(cellB); }
          applyFmt(cellB, a);
        }
        if (lc >= 3) {
          const cellC = sheet.getRange(r, 3);
          // Col C es generalmente el precio/hora calculado
          if (isFormula(sheet, r, 3)) { calcCell(cellC); } else { inputCell(cellC); }
          applyFmt(cellC, 'precio hora');
        }
        divider(sheet.getRange(r, 1, 1, Math.min(3, lc)));
      } catch(ex) {}

    } else if (hasAny(a, leftMetrics)) {
      // Métricas calculadas — siempre CALC (fondo lila)
      try {
        const rg = sheet.getRange(r, 1, 1, Math.min(3, lc));
        rg.setBackground(C.CALC_BG);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.INDIGO_D).setHorizontalAlignment('right');
        if (lc >= 2) {
          calcCell(sheet.getRange(r, 2));
          applyFmt(sheet.getRange(r, 2), a);
          sheet.getRange(r, 2).setHorizontalAlignment('right');
        }
        divider(rg);
      } catch(ex) {}

    } else if (a) {
      try {
        sheet.getRange(r, 1).setFontColor(C.TEXT_S);
        divider(sheet.getRange(r, 1, 1, Math.min(3, lc)));
      } catch(ex) {}
    }

    // ─ Zona derecha (col D+) ─
    if (lc >= 4) {
      if (hasAny(d, rightHdr)) {
        hdr(sheet.getRange(r, 4, 1, lc - 3), C.DARK);
        try { sheet.getRange(r,4).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
        sheet.setRowHeight(r, 32);

      } else if (hasAny(d, rightTypes)) {
        // Costes logísticos: col E = valor EDITABLE (el usuario introduce su coste real)
        try {
          sheet.getRange(r, 4).setFontWeight('bold').setFontColor(C.DARK).setHorizontalAlignment('right');
          if (lc >= 5) {
            const cellE = sheet.getRange(r, 5);
            if (isFormula(sheet, r, 5)) { calcCell(cellE); } else { inputCell(cellE); }
            applyFmt(cellE, d);
            cellE.setHorizontalAlignment('right');
          }
          divider(sheet.getRange(r, 4, 1, Math.max(1, lc - 3)));
        } catch(ex) {}

      } else if (hasAny(e, rightMetrics) || hasAny(d, rightMetrics)) {
        // Métricas resumen derechas: calculadas
        const lCol = hasAny(e, rightMetrics) ? 5 : 4;
        const vCol = lCol + 1;
        try {
          sheet.getRange(r, lCol).setFontWeight('bold').setFontColor(C.INDIGO_D).setHorizontalAlignment('right');
          if (vCol <= lc) {
            const cellV = sheet.getRange(r, vCol);
            if (isFormula(sheet, r, vCol)) { calcCell(cellV); }
            else { inputCell(cellV); }
            applyFmt(cellV, lCol === 5 ? e : d);
            cellV.setHorizontalAlignment('right');
          }
          // Col siguiente = descripción / ayuda (texto)
          if (vCol + 1 <= lc) {
            sheet.getRange(r, vCol+1)
              .setBackground(C.WHITE).setFontColor(C.TEXT_L)
              .setFontStyle('italic').setFontSize(9).setWrap(true);
          }
          divider(sheet.getRange(r, lCol, 1, lc - lCol + 1));
        } catch(ex) {}

      } else if (d) {
        try {
          sheet.getRange(r, 4).setFontColor(C.TEXT_S);
          if (lc >= 5) applyFmt(sheet.getRange(r, 5), d);
        } catch(ex) {}
      }
    }
  }

  if (lc>=1) sheet.setColumnWidth(1, 220);
  if (lc>=2) sheet.setColumnWidth(2, 110);
  if (lc>=3) sheet.setColumnWidth(3, 95);
  if (lc>=4) sheet.setColumnWidth(4, 210);
  if (lc>=5) sheet.setColumnWidth(5, 110);
  if (lc>=6) sheet.setColumnWidth(6, 110);
  if (lc>=7) sheet.setColumnWidth(7, 240);
  if (lc>=8) sheet.setColumnWidth(8, 185);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────
// HOJA 2 — PRESUPUESTO
//
// Según SOP, campos editables en esta hoja:
//   → Col A: Actividad (nombre de tarea — lo escribe el usuario)
//   → Col B: Descripción adicional
//   → Col C: Tipo de recurso/perfil (dropdown o texto)
//   → Col D: Responsable
//   → Col E/F: Días estimados (el usuario estima las horas)
// Campos calculados:
//   → Col G, H, I: Costes calculados automáticamente por fórmulas
//   → Filas SUBTOTAL / IVA / TOTAL / BAJA: siempre calculadas
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

    // Filas de totales: siempre calculadas, fondo corporativo sólido
    if (n === 'subtotal') {
      sumRow(row, C.SUM_A); sheet.setRowHeight(r, 30);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) { calcCell(sheet.getRange(r,9)); applyFmt(sheet.getRange(r,9),'euros'); }
    } else if (n === 'iva') {
      sumRow(row, C.SUM_B); sheet.setRowHeight(r, 28);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) { try { sheet.getRange(r,9).setNumberFormat('0.00%'); } catch(ex){} }
    } else if (n === 'total') {
      sumRow(row, C.SUM_C, 11); sheet.setRowHeight(r, 32);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) { calcCell(sheet.getRange(r,9)); applyFmt(sheet.getRange(r,9),'euros'); sheet.getRange(r,9).setFontSize(11); }
    } else if (n.indexOf('baja') !== -1) {
      sumRow(row, C.SUM_D); sheet.setRowHeight(r, 28);
      try { sheet.getRange(r,1).setHorizontalAlignment('left'); } catch(ex){}
      if (lc>=9) { calcCell(sheet.getRange(r,9)); applyFmt(sheet.getRange(r,9),'% baja'); }
    } else {
      // Filas de actividad: INPUT para cols A-F, CALC para cols G-I
      zi++;
      try {
        row.setBackground(zi%2===0 ? '#FAFAFA' : C.WHITE).setFontColor(C.TEXT);

        // Col A: nombre de actividad — EDITABLE
        const cellA = sheet.getRange(r, 1);
        if (isFormula(sheet, r, 1)) { calcCell(cellA); }
        else if (getT(sheet, r, 1) !== '') { inputCell(cellA); }
        try { cellA.setFontWeight('bold').setWrap(true); } catch(ex){}

        // Col B: descripción — EDITABLE
        if (lc >= 2) {
          const cellB = sheet.getRange(r, 2);
          if (isFormula(sheet, r, 2)) { calcCell(cellB); }
          else if (getT(sheet, r, 2) !== '') { inputCell(cellB); }
          try { cellB.setFontStyle('italic').setFontSize(9).setWrap(true); } catch(ex){}
        }

        // Col C: tipo/perfil — EDITABLE
        if (lc >= 3) {
          const cellC = sheet.getRange(r, 3);
          if (isFormula(sheet, r, 3)) { calcCell(cellC); }
          else if (getT(sheet, r, 3) !== '') { inputCell(cellC); }
          try { cellC.setHorizontalAlignment('center'); } catch(ex){}
        }

        // Col D: responsable/notas — EDITABLE
        if (lc >= 4) {
          const cellD = sheet.getRange(r, 4);
          if (isFormula(sheet, r, 4)) { calcCell(cellD); }
          else if (getT(sheet, r, 4) !== '') { inputCell(cellD); }
        }

        // Col E: días perfil 1 — EDITABLE (usuario estima)
        if (lc >= 5) {
          const cellE = sheet.getRange(r, 5);
          if (isFormula(sheet, r, 5)) { calcCell(cellE); }
          else { inputCell(cellE); }
          snf(cellE, '#,##0');
          try { cellE.setHorizontalAlignment('right'); } catch(ex){}
        }

        // Col F: días perfil 2 — EDITABLE (usuario estima)
        if (lc >= 6) {
          const cellF = sheet.getRange(r, 6);
          if (isFormula(sheet, r, 6)) { calcCell(cellF); }
          else { inputCell(cellF); }
          snf(cellF, '#,##0');
          try { cellF.setHorizontalAlignment('right'); } catch(ex){}
        }

        // Col G, H, I: costes calculados por fórmula — CALC
        if (lc >= 7) { calcCell(sheet.getRange(r,7)); snf(sheet.getRange(r,7),'€#,##0.00'); try { sheet.getRange(r,7).setHorizontalAlignment('right'); } catch(ex){} }
        if (lc >= 8) { calcCell(sheet.getRange(r,8)); snf(sheet.getRange(r,8),'€#,##0.00'); try { sheet.getRange(r,8).setHorizontalAlignment('right'); } catch(ex){} }
        if (lc >= 9) { calcCell(sheet.getRange(r,9)); snf(sheet.getRange(r,9),'€#,##0.00'); try { sheet.getRange(r,9).setHorizontalAlignment('right').setFontWeight('bold'); } catch(ex){} }

      } catch(ex) {}
    }
    divider(row);
  }

  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 340);
  if (lc>=3) sheet.setColumnWidth(3, 130);
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

  const META = 8, TL = 9;
  let firstData = 4;
  for (let r = 1; r <= Math.min(lr, 20); r++) {
    const code = getT(sheet, r, 2);
    if (/^(PT\d+|A\d+(\.\d+)*|E\d+(\.\d+)*)$/i.test(code)) { firstData = r; break; }
  }

  for (let r = 1; r < firstData; r++) {
    try {
      const vals   = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].map(txt);
      const joined = vals.map(norm).join(' | ');
      const isYear  = vals.some(function(v){ return /^20\d{2}$/.test(v); });
      const isMonth = vals.some(function(v){ return /^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)$/i.test(v); });
      const isNum   = vals.some(function(v){ return /^\d+$/.test(v); });
      const isMeta  = joined.indexOf('lider')!==-1 || joined.indexOf('mes_inicio')!==-1;

      let bg=C.INDIGO_D, h=20, fs=8;
      if (isYear)  { bg=C.DARK;     h=28; fs=10; }
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

  for (let r = firstData; r <= lr; r++) {
    if (isBlankRow(sheet, r, lc)) { sheet.setRowHeight(r, 8); continue; }
    const code = getT(sheet, r, 2);
    let metaBg=C.WHITE, metaFg=C.TEXT, metaW='normal', metaFs=10;
    let actColor=C.INDIGO, inactColor='#F0EEF8';

    if (/^PT\d+$/i.test(code)) {
      metaBg=C.PT_BG; metaFg=C.WHITE; metaW='bold';
      actColor=C.PT_ON; inactColor=C.PT_OFF; sheet.setRowHeight(r, 28);
    } else if (/^A\d+(\.\d+)*$/i.test(code)) {
      metaBg=C.TASK_BG; metaFg=C.WHITE; metaW='bold';
      actColor=C.TASK_ON; inactColor=C.TASK_OFF; sheet.setRowHeight(r, 24);
    } else if (/^E\d+(\.\d+)*$/i.test(code)) {
      metaBg=C.DEL_BG; metaFg=C.DEL_FG; metaW='normal'; metaFs=9;
      actColor=C.DEL_ON; inactColor=C.DEL_OFF; sheet.setRowHeight(r, 22);
    } else { sheet.setRowHeight(r, 22); }

    try {
      sheet.getRange(r, 1, 1, Math.min(META,lc))
        .setBackground(metaBg).setFontColor(metaFg).setFontWeight(metaW).setFontSize(metaFs);
      if (lc>=1) sheet.getRange(r,1).setHorizontalAlignment('center');
      if (lc>=2) sheet.getRange(r,2).setHorizontalAlignment('center');
      if (lc>=3) sheet.getRange(r,3).setHorizontalAlignment('left').setWrap(true);
      if (lc>=4) sheet.getRange(r,4).setHorizontalAlignment('center');
      if (lc>=5) sheet.getRange(r,5).setHorizontalAlignment('left').setFontSize(8).setWrap(true);
      if (lc>=6) { sheet.getRange(r,6).setHorizontalAlignment('center'); snf(sheet.getRange(r,6),'#,##0'); }
      if (lc>=7) { sheet.getRange(r,7).setHorizontalAlignment('center'); snf(sheet.getRange(r,7),'#,##0'); }
      if (lc>=8) { sheet.getRange(r,8).setHorizontalAlignment('center'); snf(sheet.getRange(r,8),'#,##0'); }
    } catch(ex) {}

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

    if (a==='' && (bN.indexOf('plantilla')!==-1 || bN.indexOf('ingeniero')!==-1)) {
      if (lc>=2) { hdr(sheet.getRange(r, 2, 1, Math.min(3, lc-1)), C.INDIGO_D); }
      sheet.setRowHeight(r, 30); continue;
    }
    if (hasAny(a, ['categoria profesional','tabla salarial','nivel salarial','baja desproporcionada'])) {
      hdr(row, C.DARK);
      try { sheet.getRange(r,1).setHorizontalAlignment('left').setFontSize(10); } catch(ex){}
      sheet.setRowHeight(r, 32); continue;
    }
    if (/^(ano|año)?\s*\d{4}$/.test(aN)) {
      try { row.setBackground(C.WHITE).setFontColor(C.INDIGO_D).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left'); sheet.setRowHeight(r, 36); } catch(ex){}
      continue;
    }
    if (aN.indexOf('nota:')===0 || aN[0]==='*') {
      try { row.setBackground(C.CREAM).setFontColor(C.TEXT_L).setFontStyle('italic').setFontSize(8).setWrap(true); sheet.setRowHeight(r, 32); } catch(ex){}
      continue;
    }
    if (/^\d+\s+licitador/.test(aN)) {
      try {
        row.setBackground(C.RED_XL);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.RED_D).setHorizontalAlignment('right');
        if (lc>=2) applyFmt(sheet.getRange(r,2),'euros');
        if (lc>=3) applyFmt(sheet.getRange(r,3),'% baja');
        divider(row);
      } catch(ex){} continue;
    }
    if (hasAny(a, ['coste personal total','coste no personal','total + bi','precio ofertado','precio con iva','baja'])) {
      try {
        row.setBackground(C.CALC_BG).setFontColor(C.CALC_FG).setFontWeight('bold');
        sheet.getRange(r,1).setHorizontalAlignment('right');
        for (let c=2; c<=lc; c++) { calcCell(sheet.getRange(r,c)); applyFmt(sheet.getRange(r,c), a); }
        divider(row);
      } catch(ex){} continue;
    }
    if (hasAny(a, ['plantilla ingeniero superior','plantilla ingeniero','plantilla otro perfil','subcontratado'])) {
      try {
        row.setBackground(r%2===0 ? '#F7F6FF' : C.WHITE);
        sheet.getRange(r,1).setFontWeight('bold').setFontColor(C.DARK);
        for (let c=2; c<=lc; c++) {
          let v; try { v=sheet.getRange(r,c).getValue(); } catch(ex){ v=null; }
          if (typeof v==='number') {
            if (isFormula(sheet,r,c)) { calcCell(sheet.getRange(r,c)); } else { inputCell(sheet.getRange(r,c)); }
            applyFmt(sheet.getRange(r,c), a);
          }
        }
        divider(row);
      } catch(ex){} continue;
    }
    try {
      row.setBackground(r%2===0 ? '#F7F6FF' : C.WHITE).setFontColor(C.TEXT);
      sheet.getRange(r,1).setFontColor(C.TEXT_S);
      for (let c=2; c<=lc; c++) applyFmt(sheet.getRange(r,c), a);
      divider(row);
    } catch(ex) {}
  }

  if (lc>=10) {
    for (let r=1; r<=lr; r++) {
      const j = getT(sheet, r, 10);
      if (j.length>20) {
        try { sheet.getRange(r,10).setBackground(C.CREAM).setFontColor(C.TEXT_L).setFontStyle('italic').setFontSize(8).setWrap(true); sheet.setRowHeight(r, Math.max(sheet.getRowHeight(r), 36)); } catch(ex){}
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

  const headers = sheet.getRange(hRow, 1, 1, lc).getDisplayValues()[0].map(norm);
  let colEstado=0; const moneyCols=[], pctCols=[];
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
      row.setBackground(zi%2===0 ? '#FAFAFA' : C.WHITE).setFontColor(C.TEXT);
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
      sheet.getRange(r,1,1,lc).setBackground(zi%2===0 ? '#FAFAFA' : C.WHITE).setFontColor(C.TEXT);
      divider(sheet.getRange(r,1,1,lc));
      sheet.setRowHeight(r, 24);
    } catch(ex) {}
  }
  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}
