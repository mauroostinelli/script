/**
 * Formateador corporativo para Google Sheets
 * Función principal: formatearTodo()
 */

const C = {
  RED: "#E84040",
  RED_D: "#A02828",
  RED_XL: "#FFF5F5",

  PURPLE: "#9B35B5",
  PURPLE_D: "#6B1E80",
  PURPLE_XL: "#FAF5FF",

  INDIGO: "#6355D4",
  INDIGO_D: "#3B2D9A",
  INDIGO_XL: "#F4F2FF",

  DARK: "#1E1340",
  WHITE: "#FFFFFF",
  ROW_ALT: "#F8F6FF",
  CREAM: "#FFFDF5",

  BORDER_M: "#C8BEE8",
  BORDER_L: "#E8E4F5",

  TEXT: "#1E1340",
  TEXT_M: "#4A3878",
  TEXT_L: "#7B6CA8",

  GANTT_PT_BG: "#A02828",
  GANTT_TASK_BG: "#6B1E80",
  GANTT_DEL_BG: "#EAE6FF",
  GANTT_DEL_FG: "#3B2D9A",

  GANTT_PT_ON: "#E84040",
  GANTT_TASK_ON: "#9B35B5",
  GANTT_DEL_ON: "#6355D4",

  GANTT_PT_OFF: "#F6DEDE",
  GANTT_TASK_OFF: "#F2E4F8",
  GANTT_DEL_OFF: "#F1EEFF",

  SUM_1: "#2F257E",
  SUM_2: "#3B2D9A",
  SUM_3: "#4D40B8",
  SUM_4: "#6355D4"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Formato corporativo')
    .addItem('Aplicar formato', 'formatearTodo')
    .addToUi();
}

function formatearTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    const n = norm(sheet.getName());
    if (n.includes('datos previos')) {
      formatDatosPrevios(sheet);
    } else if (n.includes('presupuesto')) {
      formatPresupuesto(sheet);
    } else if (n.includes('gantt')) {
      formatGantt(sheet);
    } else if (n.includes('sensibilidad')) {
      formatSensibilidad(sheet);
    } else if (n.includes('licitacion')) {
      formatLicitaciones(sheet);
    } else {
      formatGenerico(sheet);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Formato aplicado',
    'El formato corporativo se ha aplicado correctamente.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* =========================
   Utilidades base
   ========================= */

function norm(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function txt(v) {
  return String(v == null ? '' : v).trim();
}

function getV(sheet, r, c) {
  return sheet.getRange(r, c).getValue();
}

function getT(sheet, r, c) {
  return txt(getV(sheet, r, c));
}

function hasAny(text, arr) {
  const t = norm(text);
  return arr.some(function(k) { return t.indexOf(norm(k)) !== -1; });
}

function isBlankRow(sheet, r, lc) {
  const vals = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0];
  return vals.join('').trim() === '';
}

function prepSheet(sheet) {
  breakMerges(sheet);
  resetSheet(sheet);
}

function breakMerges(sheet) {
  const lr = Math.max(sheet.getLastRow(), 1);
  const lc = Math.max(sheet.getLastColumn(), 1);
  try {
    sheet.getRange(1, 1, lr, lc).breakApart();
  } catch (e) {}
}

function resetSheet(sheet) {
  const lr = Math.max(sheet.getLastRow(), 1);
  const lc = Math.max(sheet.getLastColumn(), 1);
  const rg = sheet.getRange(1, 1, lr, lc);

  try {
    rg.setNumberFormat('General');
  } catch (e) {}
  SpreadsheetApp.flush();

  rg.setBackground(C.WHITE)
    .setFontColor(C.TEXT)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('normal')
    .setFontStyle('normal')
    .setWrap(false)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);
}

function freeze(sheet, rows, cols) {
  try { sheet.setFrozenRows(rows); } catch (e) {}
  try { sheet.setFrozenColumns(cols); } catch (e) {}
}

function snf(cell, fmt) {
  try {
    if (typeof cell.getValue() === 'number') {
      cell.setNumberFormat(fmt);
    }
  } catch (e) {}
}

function applyFmt(cell, label) {
  try {
    const v = cell.getValue();
    if (typeof v !== 'number') return;

    const l = norm(label);
    let fmt = '€#,##0.00';

    if (
      hasAny(l, [
        'baja', 'beneficio', 'gastos generales', 'iva', '%',
        'peso criterio', 'porcentaje', 'margen'
      ])
    ) {
      fmt = '0.00%';
    } else if (
      hasAny(l, [
        'pm', 'personas adscritas', 'precio hora', 'hora real',
        'experiencia', 'coste empresa + experiencia'
      ])
    ) {
      fmt = '#,##0.00';
    } else if (
      hasAny(l, [
        'horas', 'dias', 'días', 'duracion', 'duración',
        'meses', 'licitadores', 'licitador', 'ofertas',
        'lotes', 'mes_inicio', 'mes_final'
      ])
    ) {
      fmt = '#,##0';
    }

    cell.setNumberFormat(fmt).setHorizontalAlignment('right');
  } catch (e) {}
}

function outerBorder(range) {
  try {
    range.setBorder(true, true, true, true, true, true, C.BORDER_L, SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {}
  try {
    range.setBorder(true, true, true, true, null, null, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch (e) {}
}

function bottomBorder(range) {
  try {
    range.setBorder(false, false, true, false, false, false, C.BORDER_L, SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {}
}

function hdr(range, bg) {
  range.setBackground(bg || C.DARK)
    .setFontColor(C.WHITE)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  try {
    range.setBorder(false, false, true, false, false, false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch (e) {}
}

function zebra(range, idx) {
  range.setBackground(idx % 2 === 0 ? C.ROW_ALT : C.WHITE)
    .setFontColor(C.TEXT);
  bottomBorder(range);
}

function findHeaderRow(sheet, needles, maxRows) {
  const lr = Math.min(sheet.getLastRow(), maxRows || 10);
  const lc = Math.max(sheet.getLastColumn(), 1);

  for (let r = 1; r <= lr; r++) {
    const rowText = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].map(norm).join(' | ');
    const ok = needles.every(function(k) { return rowText.indexOf(norm(k)) !== -1; });
    if (ok) return r;
  }
  return 1;
}

function colorEstadoCell(cell, estado) {
  const e = norm(estado);
  if (e === 'resuelta') {
    cell.setBackground(C.RED_XL).setFontColor(C.RED_D).setFontWeight('bold').setHorizontalAlignment('center');
  } else if (e === 'adjudicada') {
    cell.setBackground(C.INDIGO_XL).setFontColor(C.INDIGO_D).setFontWeight('bold').setHorizontalAlignment('center');
  } else if (e === 'publicada') {
    cell.setBackground(C.PURPLE_XL).setFontColor(C.PURPLE_D).setFontWeight('bold').setHorizontalAlignment('center');
  } else {
    cell.setFontWeight('bold').setHorizontalAlignment('center');
  }
}

/* =========================
   Hoja: Datos previos
   ========================= */

function formatDatosPrevios(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.PURPLE);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const leftHdr = ['licitacion subtotal', 'costes personal'];
  const leftTypes = [
    'plantilla ingeniero superior', 'plantilla ingeniero',
    'plantilla otro perfil', 'subcontratado', 'dieta'
  ];
  const leftMetrics = [
    'horas totales', 'pm', 'meses de trabajo', 'personas adscritas',
    'coste personal total', 'coste no personal'
  ];

  const rightHdr = ['costes dietas'];
  const rightTypes = ['alquiler coche', 'gasolina', 'desplazamiento'];
  const rightMetrics = [
    'precio ofertado', 'precio con iva', 'baja',
    'beneficio industrial', 'gastos generales',
    'baja promedio', 'baja segun convenio', 'baja según convenio',
    'baja maxima', 'baja máxima', 'anormal', 'licitador'
  ];

  for (let r = 1; r <= lr; r++) {
    sheet.setRowHeight(r, 22);

    const a = getT(sheet, r, 1);
    const d = lc >= 4 ? getT(sheet, r, 4) : '';
    const e = lc >= 5 ? getT(sheet, r, 5) : '';
    const rowAll = sheet.getRange(r, 1, 1, lc);

    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }

    if (a.length > 60) {
      rowAll.setBackground(C.CREAM).setFontColor(C.TEXT_L).setFontStyle('italic').setWrap(true);
      sheet.setRowHeight(r, 52);
    }

    if (hasAny(a, leftHdr)) {
      hdr(sheet.getRange(r, 1, 1, Math.min(3, lc)), C.PURPLE_D);
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      sheet.setRowHeight(r, 28);
    } else if (hasAny(a, leftTypes)) {
      const rg = sheet.getRange(r, 1, 1, Math.min(3, lc));
      rg.setBackground(C.ROW_ALT);
      sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.PURPLE_D).setHorizontalAlignment('right');
      if (lc >= 2) applyFmt(sheet.getRange(r, 2), a);
      if (lc >= 3) applyFmt(sheet.getRange(r, 3), 'precio hora');
      bottomBorder(rg);
    } else if (hasAny(a, leftMetrics)) {
      const rg = sheet.getRange(r, 1, 1, Math.min(3, lc));
      rg.setBackground(C.INDIGO_XL).setFontColor(C.INDIGO_D).setFontWeight('bold');
      sheet.getRange(r, 1).setHorizontalAlignment('right');
      if (lc >= 2) applyFmt(sheet.getRange(r, 2), a);
      if (lc >= 3) applyFmt(sheet.getRange(r, 3), a);
      bottomBorder(rg);
    } else if (a && a.length <= 60) {
      const rg = sheet.getRange(r, 1, 1, Math.min(3, lc));
      rg.setBackground(C.WHITE);
      bottomBorder(rg);
    }

    if (lc >= 4) {
      if (hasAny(d, rightHdr)) {
        hdr(sheet.getRange(r, 4, 1, lc - 3), C.INDIGO_D);
        sheet.getRange(r, 4).setHorizontalAlignment('left');
        sheet.setRowHeight(r, 28);
      } else if (hasAny(d, rightTypes)) {
        const rg = sheet.getRange(r, 4, 1, Math.min(2, lc - 3));
        rg.setBackground(C.ROW_ALT);
        sheet.getRange(r, 4).setFontWeight('bold').setFontColor(C.TEXT_M).setHorizontalAlignment('right');
        if (lc >= 5) applyFmt(sheet.getRange(r, 5), d);
        bottomBorder(sheet.getRange(r, 4, 1, Math.max(1, lc - 3)));
      } else if (hasAny(e, rightMetrics) || hasAny(d, rightMetrics)) {
        const labelCol = hasAny(e, rightMetrics) ? 5 : 4;
        const valueCol = labelCol + 1;
        if (labelCol <= lc) {
          sheet.getRange(r, labelCol)
            .setBackground(C.ROW_ALT)
            .setFontWeight('bold')
            .setFontColor(C.PURPLE_D)
            .setHorizontalAlignment('right');
        }
        if (valueCol <= lc) {
          sheet.getRange(r, valueCol)
            .setBackground(C.ROW_ALT)
            .setFontWeight('bold')
            .setFontColor(C.INDIGO_D);
          applyFmt(sheet.getRange(r, valueCol), labelCol === 5 ? e : d);
        }
        if (valueCol + 1 <= lc) {
          sheet.getRange(r, valueCol + 1)
            .setBackground(C.ROW_ALT)
            .setFontColor(C.TEXT_L)
            .setFontStyle('italic')
            .setFontSize(9)
            .setWrap(true);
        }
        bottomBorder(sheet.getRange(r, labelCol, 1, lc - labelCol + 1));
      } else if (d) {
        sheet.getRange(r, 4).setFontWeight('bold').setFontColor(C.TEXT_M);
        if (lc >= 5) applyFmt(sheet.getRange(r, 5), d);
      }
    }
  }

  if (lc >= 1) sheet.setColumnWidth(1, 220);
  if (lc >= 2) sheet.setColumnWidth(2, 110);
  if (lc >= 3) sheet.setColumnWidth(3, 100);
  if (lc >= 4) sheet.setColumnWidth(4, 200);
  if (lc >= 5) sheet.setColumnWidth(5, 110);
  if (lc >= 6) sheet.setColumnWidth(6, 120);
  if (lc >= 7) sheet.setColumnWidth(7, 220);
  if (lc >= 8) sheet.setColumnWidth(8, 180);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}

/* =========================
   Hoja: Presupuesto
   ========================= */

function formatPresupuesto(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.RED);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const hRow = findHeaderRow(sheet, ['actividad', 'descripcion'], 8);
  hdr(sheet.getRange(hRow, 1, 1, lc), C.DARK);
  sheet.setRowHeight(hRow, 38);

  if (lc >= 1) sheet.getRange(hRow, 1).setHorizontalAlignment('left');
  if (lc >= 2) sheet.getRange(hRow, 2).setHorizontalAlignment('left');
  if (lc >= 3) sheet.getRange(hRow, 3).setHorizontalAlignment('center');
  if (lc >= 4) sheet.getRange(hRow, 4).setHorizontalAlignment('left');
  for (let c = 5; c <= lc; c++) {
    sheet.getRange(hRow, c).setHorizontalAlignment('right');
  }

  let zebraIdx = 0;
  for (let r = hRow + 1; r <= lr; r++) {
    const a = getT(sheet, r, 1);
    const n = norm(a);
    const row = sheet.getRange(r, 1, 1, lc);

    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }

    sheet.setRowHeight(r, 24);

    if (n === 'subtotal') {
      row.setBackground(C.SUM_1).setFontColor(C.WHITE).setFontWeight('bold');
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      if (lc >= 9) applyFmt(sheet.getRange(r, 9), 'euros');
      sheet.setRowHeight(r, 28);
    } else if (n === 'iva') {
      row.setBackground(C.SUM_2).setFontColor(C.WHITE).setFontWeight('bold');
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      if (lc >= 9) applyFmt(sheet.getRange(r, 9), 'iva');
      sheet.setRowHeight(r, 27);
    } else if (n === 'total') {
      row.setBackground(C.SUM_3).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(11);
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      if (lc >= 9) applyFmt(sheet.getRange(r, 9), 'euros');
      sheet.setRowHeight(r, 30);
    } else if (n.indexOf('baja') !== -1) {
      row.setBackground(C.SUM_4).setFontColor(C.WHITE).setFontWeight('bold');
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      if (lc >= 9) applyFmt(sheet.getRange(r, 9), '% baja');
      sheet.setRowHeight(r, 27);
    } else {
      zebraIdx++;
      zebra(row, zebraIdx);

      if (lc >= 1) sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.DARK).setWrap(true);
      if (lc >= 2) sheet.getRange(r, 2).setFontColor(C.TEXT_M).setFontStyle('italic').setFontSize(9).setWrap(true);
      if (lc >= 3) sheet.getRange(r, 3).setHorizontalAlignment('center');
      if (lc >= 4) sheet.getRange(r, 4).setHorizontalAlignment('left');

      if (lc >= 5) { sheet.getRange(r, 5).setHorizontalAlignment('right'); snf(sheet.getRange(r, 5), '#,##0'); }
      if (lc >= 6) { sheet.getRange(r, 6).setHorizontalAlignment('right'); snf(sheet.getRange(r, 6), '#,##0'); }
      if (lc >= 7) { sheet.getRange(r, 7).setHorizontalAlignment('right'); snf(sheet.getRange(r, 7), '€#,##0.00'); }
      if (lc >= 8) { sheet.getRange(r, 8).setHorizontalAlignment('right'); snf(sheet.getRange(r, 8), '€#,##0.00'); }
      if (lc >= 9) { sheet.getRange(r, 9).setHorizontalAlignment('right'); snf(sheet.getRange(r, 9), '€#,##0.00'); }
    }

    bottomBorder(row);
  }

  if (lc >= 1) sheet.setColumnWidth(1, 240);
  if (lc >= 2) sheet.setColumnWidth(2, 340);
  if (lc >= 3) sheet.setColumnWidth(3, 120);
  if (lc >= 4) sheet.setColumnWidth(4, 170);
  if (lc >= 5) sheet.setColumnWidth(5, 100);
  if (lc >= 6) sheet.setColumnWidth(6, 110);
  if (lc >= 7) sheet.setColumnWidth(7, 110);
  if (lc >= 8) sheet.setColumnWidth(8, 120);
  if (lc >= 9) sheet.setColumnWidth(9, 130);

  outerBorder(sheet.getRange(hRow, 1, lr - hRow + 1, lc));
  freeze(sheet, hRow, 2);
  SpreadsheetApp.flush();
}

/* =========================
   Hoja: Gantt
   ========================= */

function formatGantt(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.INDIGO);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const META = 8;
  const TL = 9;

  let firstData = 4;
  for (let r = 1; r <= Math.min(lr, 20); r++) {
    const code = getT(sheet, r, 2);
    if (/^(PT\d+|A\d+(\.\d+)*|E\d+(\.\d+)*)$/i.test(code)) {
      firstData = r;
      break;
    }
  }

  for (let r = 1; r < firstData; r++) {
    const vals = sheet.getRange(r, 1, 1, lc).getDisplayValues()[0].map(txt);
    const joined = vals.map(norm).join(' | ');

    const isYear = vals.some(function(v) { return /^20\d{2}$/.test(v); });
    const isMonth = vals.some(function(v) { return /^(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)$/i.test(v); });
    const isNum = vals.some(function(v) { return /^\d+$/.test(v); });
    const isMeta = joined.indexOf('lider') !== -1 || joined.indexOf('mes_inicio') !== -1 || joined.indexOf('mes_final') !== -1;

    let bg = C.INDIGO;
    let h = 22;
    let fs = 9;

    if (isYear) {
      bg = C.DARK;
      h = 26;
      fs = 10;
    } else if (isMeta) {
      bg = C.PURPLE_D;
      h = 28;
      fs = 9;
    } else if (isMonth) {
      bg = C.INDIGO_D;
      h = 22;
      fs = 8;
    } else if (isNum) {
      bg = C.INDIGO;
      h = 18;
      fs = 7;
    }

    sheet.setRowHeight(r, h);
    hdr(sheet.getRange(r, 1, 1, Math.min(META, lc)), bg);
    sheet.getRange(r, 1, 1, Math.min(META, lc)).setFontSize(fs);

    if (lc >= TL) {
      hdr(sheet.getRange(r, TL, 1, lc - TL + 1), bg);
      sheet.getRange(r, TL, 1, lc - TL + 1).setFontSize(fs);
    }
  }

  for (let r = firstData; r <= lr; r++) {
    const code = getT(sheet, r, 2);
    const row = sheet.getRange(r, 1, 1, lc);

    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }

    let metaBg = C.WHITE;
    let metaFg = C.TEXT;
    let metaWeight = 'normal';
    let activeColor = C.INDIGO;
    let inactiveColor = C.WHITE;
    let metaFontSize = 10;

    if (/^PT\d+$/i.test(code)) {
      metaBg = C.GANTT_PT_BG;
      metaFg = C.WHITE;
      metaWeight = 'bold';
      activeColor = C.GANTT_PT_ON;
      inactiveColor = C.GANTT_PT_OFF;
      metaFontSize = 10;
      sheet.setRowHeight(r, 26);
    } else if (/^A\d+(\.\d+)*$/i.test(code)) {
      metaBg = C.GANTT_TASK_BG;
      metaFg = C.WHITE;
      metaWeight = 'bold';
      activeColor = C.GANTT_TASK_ON;
      inactiveColor = C.GANTT_TASK_OFF;
      metaFontSize = 10;
      sheet.setRowHeight(r, 23);
    } else if (/^E\d+(\.\d+)*$/i.test(code)) {
      metaBg = C.GANTT_DEL_BG;
      metaFg = C.GANTT_DEL_FG;
      metaWeight = 'normal';
      activeColor = C.GANTT_DEL_ON;
      inactiveColor = C.GANTT_DEL_OFF;
      metaFontSize = 9;
      sheet.setRowHeight(r, 22);
    } else {
      sheet.setRowHeight(r, 22);
    }

    const metaRange = sheet.getRange(r, 1, 1, Math.min(META, lc));
    metaRange.setBackground(metaBg)
      .setFontColor(metaFg)
      .setFontWeight(metaWeight)
      .setFontSize(metaFontSize);

    if (lc >= 1) sheet.getRange(r, 1).setHorizontalAlignment('center');
    if (lc >= 2) sheet.getRange(r, 2).setHorizontalAlignment('center');
    if (lc >= 3) sheet.getRange(r, 3).setHorizontalAlignment('left').setWrap(true);
    if (lc >= 4) sheet.getRange(r, 4).setHorizontalAlignment('center');
    if (lc >= 5) sheet.getRange(r, 5).setHorizontalAlignment('left').setFontSize(8).setWrap(true);
    if (lc >= 6) sheet.getRange(r, 6).setHorizontalAlignment('center'); 
    if (lc >= 7) sheet.getRange(r, 7).setHorizontalAlignment('center');
    if (lc >= 8) sheet.getRange(r, 8).setHorizontalAlignment('center');

    if (lc >= 6) snf(sheet.getRange(r, 6), '#,##0');
    if (lc >= 7) snf(sheet.getRange(r, 7), '#,##0');
    if (lc >= 8) snf(sheet.getRange(r, 8), '#,##0');

    if (lc >= TL) {
      for (let c = TL; c <= lc; c++) {
        const cell = sheet.getRange(r, c);
        const v = cell.getValue();
        const active = v === 1 || v === '1';

        if (active) {
          cell.setBackground(activeColor).setFontColor(activeColor).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(6);
        } else {
          cell.setBackground(inactiveColor).setFontColor(inactiveColor).setHorizontalAlignment('center').setFontSize(6);
        }
      }
    }

    bottomBorder(row);
  }

  if (lc >= 1) sheet.setColumnWidth(1, 24);
  if (lc >= 2) sheet.setColumnWidth(2, 68);
  if (lc >= 3) sheet.setColumnWidth(3, 340);
  if (lc >= 4) sheet.setColumnWidth(4, 100);
  if (lc >= 5) sheet.setColumnWidth(5, 220);
  if (lc >= 6) sheet.setColumnWidth(6, 70);
  if (lc >= 7) sheet.setColumnWidth(7, 70);
  if (lc >= 8) sheet.setColumnWidth(8, 70);
  for (let c = TL; c <= lc; c++) {
    sheet.setColumnWidth(c, 24);
  }

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, Math.max(firstData - 1, 1), META);
  SpreadsheetApp.flush();
}

/* =========================
   Hoja: Análisis de sensibilidad
   ========================= */

function formatSensibilidad(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.INDIGO);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  for (let r = 1; r <= lr; r++) {
    const a = getT(sheet, r, 1);
    const b = lc >= 2 ? getT(sheet, r, 2) : '';
    const row = sheet.getRange(r, 1, 1, lc);
    const aN = norm(a);
    const bN = norm(b);

    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }

    sheet.setRowHeight(r, 22);

    if (
      a === '' &&
      (bN.indexOf('plantilla ingeniero') !== -1 || bN.indexOf('plantilla otro') !== -1)
    ) {
      if (lc >= 2) {
        hdr(sheet.getRange(r, 2, 1, Math.min(3, lc - 1)), C.INDIGO_D);
      }
      sheet.setRowHeight(r, 28);
      continue;
    }

    if (
      hasAny(a, ['categoria profesional', 'tabla salarial', 'nivel salarial', 'baja desproporcionada']) ||
      hasAny(a, ['baja desproporcionada temeraria anormal'])
    ) {
      hdr(row, C.PURPLE_D);
      sheet.getRange(r, 1).setHorizontalAlignment('left');
      sheet.setRowHeight(r, 28);
      continue;
    }

    if (/^(ano|año)\s*\d{4}$/i.test(a) || /^\d{4}$/.test(a)) {
      row.setBackground(C.WHITE).setFontColor(C.PURPLE_D).setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      sheet.setRowHeight(r, 34);
      continue;
    }

    if (aN.indexOf('nota:') === 0 || aN.indexOf('*') === 0) {
      row.setBackground(C.CREAM).setFontColor(C.TEXT_L).setFontStyle('italic').setFontSize(8).setWrap(true);
      sheet.setRowHeight(r, 30);
      continue;
    }

    if (/^\d+\s+licitador/.test(aN) || /^\d+\s+licitadores/.test(aN)) {
      row.setBackground(C.RED_XL);
      sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.RED_D).setHorizontalAlignment('right');
      if (lc >= 2) applyFmt(sheet.getRange(r, 2), 'euros');
      if (lc >= 3) applyFmt(sheet.getRange(r, 3), '% baja');
      bottomBorder(row);
      continue;
    }

    if (
      hasAny(a, [
        'coste personal total', 'coste no personal',
        'total + bi', 'precio ofertado', 'precio con iva', 'baja'
      ])
    ) {
      row.setBackground(C.INDIGO_XL).setFontColor(C.INDIGO_D).setFontWeight('bold');
      sheet.getRange(r, 1).setHorizontalAlignment('right');
      for (let c = 2; c <= lc; c++) {
        applyFmt(sheet.getRange(r, c), a);
      }
      bottomBorder(row);
      continue;
    }

    if (
      hasAny(a, [
        'plantilla ingeniero superior', 'plantilla ingeniero',
        'plantilla otro perfil', 'subcontratado'
      ])
    ) {
      row.setBackground(r % 2 === 0 ? C.PURPLE_XL : C.INDIGO_XL);
      sheet.getRange(r, 1).setFontWeight('bold').setFontColor(C.PURPLE_D);
      for (let c = 2; c <= lc; c++) {
        const cell = sheet.getRange(r, c);
        const v = cell.getValue();
        if (typeof v === 'number') {
          applyFmt(cell, a);
        }
      }
      bottomBorder(row);
      continue;
    }

    zebra(row, r);
    sheet.getRange(r, 1).setFontColor(C.TEXT_M);
    for (let c = 2; c <= lc; c++) {
      applyFmt(sheet.getRange(r, c), a);
    }
  }

  if (lc >= 10) {
    for (let r = 1; r <= lr; r++) {
      const j = getT(sheet, r, 10);
      if (j.length > 20) {
        sheet.getRange(r, 10)
          .setBackground(C.CREAM)
          .setFontColor(C.TEXT_L)
          .setFontStyle('italic')
          .setFontSize(8)
          .setWrap(true);
        sheet.setRowHeight(r, Math.max(sheet.getRowHeight(r), 38));
      }
    }
    sheet.setColumnWidth(10, 260);
  }

  if (lc >= 1) sheet.setColumnWidth(1, 240);
  if (lc >= 2) sheet.setColumnWidth(2, 130);
  if (lc >= 3) sheet.setColumnWidth(3, 130);
  if (lc >= 4) sheet.setColumnWidth(4, 130);
  if (lc >= 5) sheet.setColumnWidth(5, 150);
  if (lc >= 6) sheet.setColumnWidth(6, 160);
  if (lc >= 7) sheet.setColumnWidth(7, 110);
  if (lc >= 8) sheet.setColumnWidth(8, 110);
  if (lc >= 9) sheet.setColumnWidth(9, 120);

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 2, 1);
  SpreadsheetApp.flush();
}

/* =========================
   Hoja: Licitaciones previas
   ========================= */

function formatLicitaciones(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.RED);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  const hRow = findHeaderRow(sheet, ['expediente', 'objeto'], 6);
  hdr(sheet.getRange(hRow, 1, 1, lc), C.DARK);
  sheet.setRowHeight(hRow, 38);

  const headers = sheet.getRange(hRow, 1, 1, lc).getDisplayValues()[0].map(norm);

  let colEstado = 0;
  const moneyCols = [];
  const pctCols = [];

  headers.forEach(function(h, i) {
    const c = i + 1;

    if (h === 'estado' || h.indexOf('estado') !== -1) colEstado = c;

    if (
      h.indexOf('presupuesto') !== -1 ||
      h.indexOf('valor estimado') !== -1 ||
      h.indexOf('importe adjudicado') !== -1
    ) {
      moneyCols.push(c);
    }

    if (
      h.indexOf('%') !== -1 ||
      h.indexOf('baja') !== -1 ||
      h.indexOf('peso criterio') !== -1
    ) {
      pctCols.push(c);
    }
  });

  let zebraIdx = 0;
  for (let r = hRow + 1; r <= lr; r++) {
    const row = sheet.getRange(r, 1, 1, lc);

    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }

    zebraIdx++;
    zebra(row, zebraIdx);
    sheet.setRowHeight(r, 22);

    moneyCols.forEach(function(c) {
      applyFmt(sheet.getRange(r, c), 'euros');
    });

    pctCols.forEach(function(c) {
      applyFmt(sheet.getRange(r, c), '%');
    });

    if (colEstado > 0) {
      colorEstadoCell(sheet.getRange(r, colEstado), getT(sheet, r, colEstado));
    }
  }

  if (lc >= 1) sheet.setColumnWidth(1, 120);
  if (lc >= 2) sheet.setColumnWidth(2, 420);
  if (lc >= 3) sheet.setColumnWidth(3, 280);
  if (lc >= 4) sheet.setColumnWidth(4, 240);
  if (lc >= 5) sheet.setColumnWidth(5, 120);
  if (lc >= 9) sheet.setColumnWidth(9, 130);
  if (lc >= 10) sheet.setColumnWidth(10, 130);
  if (lc >= 16) sheet.setColumnWidth(16, 130);
  if (lc >= 17) sheet.setColumnWidth(17, 100);
  if (lc >= 18) sheet.setColumnWidth(18, 220);

  outerBorder(sheet.getRange(hRow, 1, lr - hRow + 1, lc));
  freeze(sheet, hRow, 2);
  SpreadsheetApp.flush();
}

/* =========================
   Fallback genérico
   ========================= */

function formatGenerico(sheet) {
  prepSheet(sheet);
  sheet.setTabColor(C.PURPLE);

  const lr = sheet.getLastRow();
  const lc = sheet.getLastColumn();
  if (lr < 1 || lc < 1) return;

  hdr(sheet.getRange(1, 1, 1, lc), C.DARK);
  sheet.setRowHeight(1, 34);

  let zebraIdx = 0;
  for (let r = 2; r <= lr; r++) {
    if (isBlankRow(sheet, r, lc)) {
      sheet.setRowHeight(r, 8);
      continue;
    }
    zebraIdx++;
    zebra(sheet.getRange(r, 1, 1, lc), zebraIdx);
    sheet.setRowHeight(r, 22);
  }

  outerBorder(sheet.getRange(1, 1, lr, lc));
  freeze(sheet, 1, 0);
  SpreadsheetApp.flush();
}
