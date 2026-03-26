/**
 * PANEL DE SEGUIMIENTO DE LICITACIÓN Y COSTES DE PROYECTO
 * Genera la hoja Google Sheets a partir de la imagen del panel.
 *
 * Función principal: crearPanel()
 *
 * SISTEMA DE COLORES:
 *   🟡 INPUT_BG (#FFF9C4)  — Celda EDITABLE por el usuario
 *   🟣 CALC_BG  (#EDEAF6)  — Celda con FÓRMULA / calculada
 *   ⬛ DARK    (#1E1340)  — Cabeceras de sección
 *   🟦 BLUE    (#1A3A5C)  — Cabeceras azul oscuro (panel principal)
 *   🟨 YELLOW  (#FFD600)  — Resaltado especial (Meses de Trabajo)
 */

// ─── PALETA DE COLORES
const C = {
  WHITE:      '#FFFFFF',
  DARK:       '#1A3A5C',
  DARK2:      '#1E2B3A',
  MID:        '#2C4A6E',
  LIGHT_HDR:  '#2E5F8A',
  GRAY_HDR:   '#4A5568',
  BLUE_ROW:   '#D6E4F0',
  ALT:        '#EBF3FA',
  INPUT_BG:   '#FFF9C4',
  INPUT_FG:   '#3D3000',
  INPUT_BD:   '#E6D400',
  CALC_BG:    '#EDEAF6',
  CALC_FG:    '#3B2D9A',
  YELLOW_HL:  '#FFD600',
  YELLOW_FG:  '#1A1A00',
  RED:        '#E84040',
  GREEN:      '#2E7D32',
  TEXT:       '#1A1A2E',
  TEXT_L:     '#6B7280',
  BORDER:     '#B8C8D8',
  BORDER_M:   '#7A9AB8',
  ALERT_BG:   '#FFFDE7',
  ALERT_BD:   '#F9A825',
  LINK_BG:    '#E8F0FE',
  LINK_FG:    '#1A56DB',
};

// ─── MENÚ
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Panel Licitación')
    .addItem('✦ Crear Panel', 'crearPanel')
    .addItem('↺ Limpiar hoja', 'limpiarHoja')
    .addToUi();
}

// ─── UTILIDADES
function s(sheet, r, c, valor) {
  sheet.getRange(r, c).setValue(valor);
}

function fmt(sheet, r, c, opciones) {
  const cell = sheet.getRange(r, c);
  if (opciones.bg)        cell.setBackground(opciones.bg);
  if (opciones.fg)        cell.setFontColor(opciones.fg);
  if (opciones.bold)      cell.setFontWeight('bold');
  if (opciones.italic)    cell.setFontStyle('italic');
  if (opciones.size)      cell.setFontSize(opciones.size);
  if (opciones.align)     cell.setHorizontalAlignment(opciones.align);
  if (opciones.valign)    cell.setVerticalAlignment(opciones.valign);
  if (opciones.wrap)      cell.setWrap(true);
  if (opciones.numFmt)    cell.setNumberFormat(opciones.numFmt);
  if (opciones.border)    cell.setBorder(true, true, true, true, false, false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID);
}

function rng(sheet, r, c, numR, numC) {
  return sheet.getRange(r, c, numR, numC);
}

function fmtRange(sheet, r, c, numR, numC, opciones) {
  const range = rng(sheet, r, c, numR, numC);
  if (opciones.bg)     range.setBackground(opciones.bg);
  if (opciones.fg)     range.setFontColor(opciones.fg);
  if (opciones.bold)   range.setFontWeight('bold');
  if (opciones.size)   range.setFontSize(opciones.size);
  if (opciones.align)  range.setHorizontalAlignment(opciones.align);
  if (opciones.valign) range.setVerticalAlignment(opciones.valign);
  if (opciones.wrap)   range.setWrap(true);
  if (opciones.border) range.setBorder(true, true, true, true, false, false, C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID);
}

function merge(sheet, r, c, numR, numC) {
  try { sheet.getRange(r, c, numR, numC).merge(); } catch(e) {}
}

function outerBorder(sheet, r, c, numR, numC, color) {
  try {
    sheet.getRange(r, c, numR, numC)
      .setBorder(true, true, true, true, false, false,
        color || C.BORDER_M, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } catch(e) {}
}

function innerBorder(sheet, r, c, numR, numC) {
  try {
    sheet.getRange(r, c, numR, numC)
      .setBorder(true, true, true, true, true, true,
        C.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}
}

function cabeceraSec(sheet, r, c, numR, numC, texto, bg) {
  merge(sheet, r, c, numR, numC);
  s(sheet, r, c, texto);
  fmtRange(sheet, r, c, numR, numC, {
    bg: bg || C.DARK,
    fg: C.WHITE,
    bold: true,
    size: 9,
    align: 'left',
    valign: 'middle'
  });
}

function inputCell(sheet, r, c) {
  const cell = sheet.getRange(r, c);
  cell.setBackground(C.INPUT_BG)
      .setFontColor(C.INPUT_FG)
      .setFontWeight('bold');
  cell.setBorder(false, false, true, false, false, false,
    C.INPUT_BD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function calcCell(sheet, r, c) {
  const cell = sheet.getRange(r, c);
  cell.setBackground(C.CALC_BG)
      .setFontColor(C.CALC_FG)
      .setFontWeight('bold');
}

// ─── LIMPIAR HOJA
function limpiarHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Panel Licitación');
  if (sheet) {
    sheet.clear();
    try { sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart(); } catch(e) {}
  }
}

// ─── FUNCIÓN PRINCIPAL
function crearPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Obtener o crear hoja
  let sheet = ss.getSheetByName('Panel Licitación');
  if (!sheet) {
    sheet = ss.insertSheet('Panel Licitación');
  } else {
    sheet.clear();
    try { sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart(); } catch(e) {}
  }

  sheet.setTabColor(C.DARK);

  // ── Configurar columnas y filas
  const colWidths = [18, 140, 110, 110, 18, 140, 100, 100, 18, 130, 120, 120, 18, 140, 110, 110];
  colWidths.forEach(function(w, i) {
    try { sheet.setColumnWidth(i+1, w); } catch(e) {}
  });

  // Alturas de fila
  for (let r = 1; r <= 60; r++) {
    try { sheet.setRowHeight(r, 22); } catch(e) {}
  }
  try { sheet.setRowHeight(1, 36); } catch(e) {}
  try { sheet.setRowHeight(3, 28); } catch(e) {}
  try { sheet.setRowHeight(11, 28); } catch(e) {}
  try { sheet.setRowHeight(19, 28); } catch(e) {}
  try { sheet.setRowHeight(29, 28); } catch(e) {}
  try { sheet.setRowHeight(40, 28); } catch(e) {}
  try { sheet.setRowHeight(50, 28); } catch(e) {}

  // ════════════════════════════════════════════
  // FILA 1 — TÍTULO PRINCIPAL
  // ════════════════════════════════════════════
  merge(sheet, 1, 1, 1, 16);
  s(sheet, 1, 1, 'PANEL DE SEGUIMIENTO DE LICITACIÓN Y COSTES DE PROYECTO');
  fmtRange(sheet, 1, 1, 1, 16, {
    bg: C.DARK,
    fg: C.WHITE,
    bold: true,
    size: 13,
    align: 'center',
    valign: 'middle'
  });

  // ════════════════════════════════════════════
  // BLOQUE A — RESUMEN ECONÓMICO (cols 1-4, filas 3-8)
  // ════════════════════════════════════════════
  cabeceraSec(sheet, 3, 1, 1, 4, '📈  RESUMEN ECONÓMICO', C.DARK2);
  outerBorder(sheet, 3, 1, 6, 4);

  const resEco = [
    [4, 'Lic. Subtotal',    '€80.000,00',   false],
    [5, 'Precio Ofertado',  '€91.198,74',   true],
    [6, 'Precio con IVA',   '€110.350,47',  true],
    [7, 'Baja',             '12,28%',       true],
  ];
  resEco.forEach(function(row) {
    s(sheet, row[0], 2, row[1]);
    s(sheet, row[0], 3, row[2]);
    fmtRange(sheet, row[0], 2, 1, 1, {fg: C.WHITE, bg: C.DARK2, align: 'right', bold: true, size: 9});
    fmtRange(sheet, row[0], 3, 1, 2, {fg: C.WHITE, bg: C.DARK2, bold: true, align: 'right', size: 9, numFmt: row[3] ? '0.00%' : '€#,##0.00'});
    merge(sheet, row[0], 3, 1, 2);
    if (row[3]) calcCell(sheet, row[0], 3);
    else        inputCell(sheet, row[0], 3);
  });

  // ════════════════════════════════════════════
  // BLOQUE B — RESUMEN DE EJECUCIÓN (cols 5-8, filas 3-8)
  // ════════════════════════════════════════════
  cabeceraSec(sheet, 3, 5, 1, 4, '⏱  RESUMEN DE EJECUCIÓN', C.LIGHT_HDR);
  outerBorder(sheet, 3, 5, 6, 4);

  const resEje = [
    [4, 'Horas Totales',       '5.381',  false, false],
    [5, 'Meses de Trabajo',    '2,00',   true,  true],
    [6, 'Personas adscritas',  '20,23',  true,  false],
  ];
  resEje.forEach(function(row) {
    s(sheet, row[0], 6, row[1]);
    s(sheet, row[0], 8, row[2]);
    fmtRange(sheet, row[0], 6, 1, 2, {fg: C.TEXT, bg: row[3] ? C.YELLOW_HL : C.WHITE, bold: row[3], size: 9, align: 'left'});
    merge(sheet, row[0], 6, 1, 2);
    fmtRange(sheet, row[0], 8, 1, 1, {fg: row[3] ? C.YELLOW_FG : C.TEXT, bg: row[3] ? C.YELLOW_HL : C.WHITE, bold: true, size: 9, align: 'right'});
    if (row[4]) {
      // Meses de trabajo: celda destacada en amarillo
      sheet.getRange(row[0], 8).setBackground(C.YELLOW_HL).setFontColor(C.YELLOW_FG).setFontWeight('bold');
    } else if (row[2]) {
      calcCell(sheet, row[0], 8);
    } else {
      inputCell(sheet, row[0], 8);
    }
  });

  // ════════════════════════════════════════════
  // BLOQUE C — COSTES TOTALES (cols 9-12, filas 3-8)
  // ════════════════════════════════════════════
  cabeceraSec(sheet, 3, 9, 1, 4, 'COSTES TOTALES', C.MID);
  outerBorder(sheet, 3, 9, 6, 4);

  s(sheet, 4, 10, '👤 Coste personal total');
  s(sheet, 4, 12, '€91.198,74');
  fmtRange(sheet, 4, 10, 1, 2, {fg: C.WHITE, bg: C.MID, bold: true, size: 9, align: 'left'});
  merge(sheet, 4, 10, 1, 2);
  calcCell(sheet, 4, 12);
  sheet.getRange(4, 12).setValue('€91.198,74').setNumberFormat('€#,##0.00').setFontWeight('bold').setHorizontalAlignment('right');

  s(sheet, 6, 10, 'Coste no personal');
  s(sheet, 6, 12, '€0,00');
  fmtRange(sheet, 6, 10, 1, 2, {fg: C.WHITE, bg: C.MID, bold: true, size: 9, align: 'left'});
  merge(sheet, 6, 10, 1, 2);
  inputCell(sheet, 6, 12);
  sheet.getRange(6, 12).setValue(0).setNumberFormat('€#,##0.00').setHorizontalAlignment('right');

  // ════════════════════════════════════════════
  // BLOQUE D — MÁRGENES Y PARÁMETROS (cols 13-16, filas 3-8)
  // ════════════════════════════════════════════
  cabeceraSec(sheet, 3, 13, 1, 4, 'MÁRGENES Y PARÁMETROS', C.GRAY_HDR);
  outerBorder(sheet, 3, 13, 6, 4);

  const margenes = [
    [4, 'Beneficio industrial',     '6%',   'por defecto'],
    [5, 'Gastos generales',         '13%',  'por defecto'],
    [6, 'Baja promedio histórico',  '21%',  ''],
  ];
  margenes.forEach(function(row) {
    s(sheet, row[0], 14, row[1]);
    s(sheet, row[0], 15, row[2]);
    s(sheet, row[0], 16, row[3]);
    fmtRange(sheet, row[0], 14, 1, 1, {fg: C.TEXT, bg: C.WHITE, size: 9, align: 'left'});
    inputCell(sheet, row[0], 15);
    sheet.getRange(row[0], 15).setNumberFormat('0%').setHorizontalAlignment('right');
    fmtRange(sheet, row[0], 16, 1, 1, {fg: C.TEXT_L, bg: C.WHITE, italic: true, size: 8});
  });

  // ════════════════════════════════════════════
  // BLOQUE E — TARIFAS DE PERSONAL (cols 1-8, filas 11-17)
  // ════════════════════════════════════════════
  const R_TARIFAS = 11;
  cabeceraSec(sheet, R_TARIFAS, 1, 1, 8, 'TARIFAS DE PERSONAL', C.DARK);
  outerBorder(sheet, R_TARIFAS, 1, 8, 8);

  // Encabezados de tabla
  s(sheet, R_TARIFAS+1, 2, 'Perfil');
  s(sheet, R_TARIFAS+1, 5, 'Tarifa/Mes');
  s(sheet, R_TARIFAS+1, 8, 'Tarifa/Hora');
  fmtRange(sheet, R_TARIFAS+1, 1, 1, 8, {bg: C.DARK2, fg: C.WHITE, bold: true, size: 9, align: 'left'});
  sheet.getRange(R_TARIFAS+1, 5).setHorizontalAlignment('right');
  sheet.getRange(R_TARIFAS+1, 8).setHorizontalAlignment('right');
  merge(sheet, R_TARIFAS+1, 2, 1, 3);

  const tarifas = [
    ['Plantilla ingeniero superior', '€4.700,00', '35,34 €'],
    ['Plantilla ingeniero',          '€3.700,00', '27,82 €'],
    ['Plantilla otro perfil',        '€3.400,00', '25,56 €'],
    ['Subcontratado',                '€5.000,00', '37,59 €'],
    ['Dieta',                        '',          '109,00 €'],
  ];
  tarifas.forEach(function(tar, i) {
    const r = R_TARIFAS + 2 + i;
    s(sheet, r, 2, tar[0]);
    if (tar[1]) s(sheet, r, 5, tar[1]);
    s(sheet, r, 8, tar[2]);

    fmtRange(sheet, r, 1, 1, 8, {bg: i%2===0 ? C.WHITE : C.ALT, fg: C.TEXT, size: 9});
    merge(sheet, r, 2, 1, 3);
    sheet.getRange(r, 2).setFontWeight('bold').setHorizontalAlignment('left');
    if (tar[1]) {
      inputCell(sheet, r, 5);
      sheet.getRange(r, 5).setNumberFormat('€#,##0.00').setHorizontalAlignment('right');
    }
    calcCell(sheet, r, 8);
    sheet.getRange(r, 8).setHorizontalAlignment('right').setNumberFormat('€#,##0.00');
    try { sheet.getRange(r, 1, 1, 8).setBorder(false, false, true, false, false, false, C.BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
  });

  // Nota R2M/PENTA en fila Dieta
  const rDieta = R_TARIFAS + 2 + 4;
  s(sheet, rDieta, 4, 'R2M/PENTA?');
  sheet.getRange(rDieta, 4).setFontColor(C.LINK_FG).setFontStyle('italic').setFontSize(8);

  // ════════════════════════════════════════════
  // BLOQUE F — DETALLE DE DIETAS Y VIAJES (cols 1-8, filas 19-27)
  // ════════════════════════════════════════════
  const R_DIETAS = 19;
  cabeceraSec(sheet, R_DIETAS, 1, 1, 8, 'DETALLE DE DIETAS Y VIAJES', C.DARK);
  outerBorder(sheet, R_DIETAS, 1, 9, 8);

  // Sub-cabecera Dietas Diarias
  s(sheet, R_DIETAS+1, 1, 'Dietas Diarias');
  fmtRange(sheet, R_DIETAS+1, 1, 1, 4, {bg: C.DARK2, fg: C.WHITE, bold: true, size: 9, align: 'left'});
  merge(sheet, R_DIETAS+1, 1, 1, 4);

  const dietasDiarias = [
    ['Alquiler coche', '€50,00'],
    ['Gasolina',       '€27,00'],
    ['Desplazamiento', '€32,00'],
  ];
  dietasDiarias.forEach(function(d, i) {
    const r = R_DIETAS + 2 + i;
    s(sheet, r, 2, d[0]);
    s(sheet, r, 4, d[1]);
    fmtRange(sheet, r, 1, 1, 4, {bg: i%2===0 ? C.WHITE : C.ALT, fg: C.TEXT, size: 9});
    merge(sheet, r, 2, 1, 2);
    sheet.getRange(r, 2).setFontWeight('bold');
    inputCell(sheet, r, 4);
    sheet.getRange(r, 4).setNumberFormat('€#,##0.00').setHorizontalAlignment('right');
    try { sheet.getRange(r, 1, 1, 4).setBorder(false, false, true, false, false, false, C.BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
  });

  // Sub-cabecera Viajes
  s(sheet, R_DIETAS+1, 5, 'Viajes');
  fmtRange(sheet, R_DIETAS+1, 5, 1, 4, {bg: C.DARK2, fg: C.WHITE, bold: true, size: 9, align: 'left'});
  merge(sheet, R_DIETAS+1, 5, 1, 4);

  s(sheet, R_DIETAS+2, 5, 'Internacionales y Nacionales');
  fmtRange(sheet, R_DIETAS+2, 5, 1, 4, {bg: C.LIGHT_HDR, fg: C.WHITE, bold: true, size: 8, align: 'left'});
  merge(sheet, R_DIETAS+2, 5, 1, 4);

  const viajes = [
    ['Por días en pernoctar',    '48,08 €'],
    ['Por días en sin pernoctar','64,00 €'],
    ['Gasolina per km',          '0,27 €'],
  ];
  viajes.forEach(function(v, i) {
    const r = R_DIETAS + 3 + i;
    s(sheet, r, 6, v[0]);
    s(sheet, r, 8, v[1]);
    fmtRange(sheet, r, 5, 1, 4, {bg: i%2===0 ? C.WHITE : C.ALT, fg: C.TEXT, size: 9});
    merge(sheet, r, 6, 1, 2);
    sheet.getRange(r, 6).setFontWeight('bold').setHorizontalAlignment('left');
    inputCell(sheet, r, 8);
    sheet.getRange(r, 8).setHorizontalAlignment('right').setNumberFormat('€#,##0.00');
    try { sheet.getRange(r, 5, 1, 4).setBorder(false, false, true, false, false, false, C.BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
  });

  // ════════════════════════════════════════════
  // BLOQUE G — ANÁLISIS DE BAJAS Y CONTINGENCIAS (cols 9-16, filas 11-27)
  // ════════════════════════════════════════════
  const R_BAJAS = 11;
  cabeceraSec(sheet, R_BAJAS, 9, 1, 8, 'ANÁLISIS DE BAJAS Y CONTINGENCIAS', C.DARK);
  outerBorder(sheet, R_BAJAS, 9, 17, 8);

  const bajas = [
    ['Baja costes empresa + 9% GG + 5% BI', '-118,42%', '-118,42 €', C.RED,    true ],
    ['Baja según convenio',                 '-98,00%',  '-28.000 €', C.MID,    false],
    ['Baja máxima (convenio + 9% GG + 5% BI','-31,00%', '-31,00%',  C.MID,    false],
    ['Baja máxima anormal 1 licitador',     '-25,00%',  '-25.000 €', C.MID,    false],
    ['Baja máxima anormal 2 licitadores',   '-20,00%',  '-20.000 €', C.MID,    false],
    ['Baja máxima anormal 3+ licitadores',  '-28,58%',  '-28.580 €', C.MID,    false],
  ];

  // Encabezados
  s(sheet, R_BAJAS+1, 10, 'Concepto');
  s(sheet, R_BAJAS+1, 14, '%');
  s(sheet, R_BAJAS+1, 16, 'Valor €');
  fmtRange(sheet, R_BAJAS+1, 9, 1, 8, {bg: C.DARK2, fg: C.WHITE, bold: true, size: 9});
  merge(sheet, R_BAJAS+1, 10, 1, 4);
  sheet.getRange(R_BAJAS+1, 14).setHorizontalAlignment('right');
  sheet.getRange(R_BAJAS+1, 16).setHorizontalAlignment('right');

  bajas.forEach(function(b, i) {
    const r = R_BAJAS + 2 + i;
    s(sheet, r, 10, b[0]);
    s(sheet, r, 14, b[1]);
    s(sheet, r, 16, b[2]);

    // Barra de progreso simulada (col 15)
    const barBg = b[4] ? C.RED : C.MID;
    sheet.getRange(r, 15).setBackground(barBg);

    fmtRange(sheet, r, 9, 1, 8, {bg: i%2===0 ? C.WHITE : C.ALT, size: 9});
    merge(sheet, r, 10, 1, 4);
    sheet.getRange(r, 10).setFontWeight('bold').setFontColor(b[3]).setHorizontalAlignment('left').setWrap(true);
    calcCell(sheet, r, 14);
    sheet.getRange(r, 14).setHorizontalAlignment('right').setFontColor(C.RED).setFontWeight('bold');
    calcCell(sheet, r, 16);
    sheet.getRange(r, 16).setHorizontalAlignment('right').setFontColor(C.RED).setFontWeight('bold');
    try { sheet.getRange(r, 9, 1, 8).setBorder(false, false, true, false, false, false, C.BORDER, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
  });

  // Notas 6% / 13% por defecto
  s(sheet, R_BAJAS+9,  9, '6% por defecto');
  s(sheet, R_BAJAS+10, 9, '13% por defecto');
  fmtRange(sheet, R_BAJAS+9,  9, 1, 8, {bg: C.ALT, fg: C.TEXT_L, italic: true, size: 8});
  fmtRange(sheet, R_BAJAS+10, 9, 1, 8, {bg: C.ALT, fg: C.TEXT_L, italic: true, size: 8});
  merge(sheet, R_BAJAS+9,  9, 1, 8);
  merge(sheet, R_BAJAS+10, 9, 1, 8);

  // ════════════════════════════════════════════
  // BLOQUE H — SECCIÓN DE COMENTARIOS Y ALERTAS (cols 9-12, filas 29-38)
  // ════════════════════════════════════════════
  const R_COM = 29;
  cabeceraSec(sheet, R_COM, 9, 1, 4, 'SECCIÓN DE COMENTARIOS Y ALERTAS', C.ALERT_BD);
  sheet.getRange(R_COM, 9).setFontColor(C.TEXT);
  outerBorder(sheet, R_COM, 9, 10, 4, C.ALERT_BD);

  const comentarios = [
    'Explicar la dinámica general del proyecto.',
    '¿Cuántos trabajadores planeamos contratar?',
    '¿A quién se subcontrata?',
    '¿Quién apoya desde R2M/PENTA?',
  ];
  comentarios.forEach(function(com, i) {
    const r = R_COM + 1 + i;
    s(sheet, r, 9, '☐  ' + com);
    fmtRange(sheet, r, 9, 1, 4, {bg: C.ALERT_BG, fg: C.TEXT, size: 9, wrap: true});
    merge(sheet, r, 9, 1, 4);
    sheet.setRowHeight(r, 24);
    try { sheet.getRange(r, 9, 1, 4).setBorder(false, false, true, false, false, false, C.ALERT_BD, SpreadsheetApp.BorderStyle.SOLID); } catch(e) {}
  });

  // Filas adicionales de comentario libre
  for (let i = 0; i < 5; i++) {
    const r = R_COM + 5 + i;
    inputCell(sheet, r, 9);
    merge(sheet, r, 9, 1, 4);
    fmtRange(sheet, r, 9, 1, 4, {bg: C.INPUT_BG, size: 9, wrap: true});
    sheet.setRowHeight(r, 22);
  }

  // ════════════════════════════════════════════
  // BLOQUE I — ENLACES DE REFERENCIA (cols 13-16, filas 29-38)
  // ════════════════════════════════════════════
  const R_LINKS = 29;
  cabeceraSec(sheet, R_LINKS, 13, 1, 4, 'ENLACES DE REFERENCIA', C.DARK2);
  outerBorder(sheet, R_LINKS, 13, 10, 4);

  // Botón simulado: Gantt y presupuesto
  merge(sheet, R_LINKS+2, 13, 2, 4);
  s(sheet, R_LINKS+2, 13, '📊  Gantt y presupuesto...');
  fmtRange(sheet, R_LINKS+2, 13, 2, 4, {
    bg: C.LINK_BG,
    fg: C.LINK_FG,
    bold: true,
    size: 10,
    align: 'center',
    valign: 'middle',
    border: true
  });
  sheet.setRowHeight(R_LINKS+2, 34);
  sheet.setRowHeight(R_LINKS+3, 34);

  // Campos de URL editables
  const linksLabel = ['Convocatoria', 'Pliego técnico', 'Pliego económico', 'Documentación'];
  linksLabel.forEach(function(label, i) {
    const r = R_LINKS + 4 + i;
    s(sheet, r, 14, label);
    fmtRange(sheet, r, 14, 1, 1, {fg: C.TEXT, size: 9, bold: true});
    inputCell(sheet, r, 15);
    merge(sheet, r, 15, 1, 2);
    sheet.getRange(r, 15).setFontColor(C.LINK_FG).setFontSize(9);
    sheet.setRowHeight(r, 22);
  });

  // ════════════════════════════════════════════
  // BORDE EXTERIOR GENERAL
  // ════════════════════════════════════════════
  outerBorder(sheet, 1, 1, 38, 16, C.DARK);

  // ════════════════════════════════════════════
  // CONGELAR FILAS / COLUMNAS
  // ════════════════════════════════════════════
  try { sheet.setFrozenRows(2); }    catch(e) {}
  try { sheet.setFrozenColumns(1); } catch(e) {}

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    '✦ Panel creado',
    'El panel de seguimiento se ha generado correctamente.\n\n' +
    '🟡 Amarillo = celda a rellenar por el usuario\n' +
    '🟣 Lila     = celda calculada automáticamente',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
