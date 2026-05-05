// ============================================================
// La Fondiatta — Receptor de Presupuestos desde el Generador
// v2.1 — UPDATE support + lectura batch (perf)
// ============================================================
//
// v2.1: Fix de performance. El loop original hacía 1 getValue() por fila
// (965 round-trips ≈ 50s con el Sheet actual). Ahora lee column A en batch
// junto con column K → 2 llamadas vs N. Tiempo total: ~1s.
//
// REQUISITO: El Sheet debe tener una columna V con header "Internal ID".
// ============================================================

function doPost(e) {
  try {
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = JSON.parse(e.parameter.payload || '{}');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheets()[0];
    var lastRow = sh.getLastRow();

    var ID_PRES_COL = 11;     // K = ID Presupuesto (P-2026-XXX)
    var INTERNAL_ID_COL = 22; // V = Internal ID (UUID interno)

    // === DEDUP: si llega presupuestoId y matchea, ACTUALIZAR ===
    var targetRow = null;
    var isUpdate = false;
    if (data.presupuestoId && lastRow >= 2) {
      var internalIds = sh.getRange(2, INTERNAL_ID_COL, lastRow - 1, 1).getValues();
      for (var j = 0; j < internalIds.length; j++) {
        if (internalIds[j][0] === data.presupuestoId) {
          targetRow = j + 2;
          isUpdate = true;
          data.id = sh.getRange(targetRow, ID_PRES_COL).getValue();
          break;
        }
      }
    }

    // === Si no fue UPDATE: buscar fila pre-generada vacía o appender ===
    if (!targetRow) {
      var rangeRows = Math.max(lastRow - 1, 1);
      // Lectura BATCH de columnas K (IDs) y A (contactos) — evita N round-trips
      var idsValues = sh.getRange(2, ID_PRES_COL, rangeRows, 1).getValues();
      var contactosValues = sh.getRange(2, 1, rangeRows, 1).getValues();

      for (var i = 0; i < idsValues.length; i++) {
        var id = idsValues[i][0];
        var contacto = contactosValues[i][0];
        if (id && String(id).indexOf('P-2026-') === 0 && !contacto) {
          targetRow = i + 2;
          break;
        }
      }

      if (!targetRow) {
        targetRow = lastRow + 1;
        var newId = 'P-2026-' + String(targetRow - 1).padStart(3, '0');
        sh.getRange(targetRow, ID_PRES_COL).setValue(newId);
        data.id = newId;
      } else {
        data.id = idsValues[targetRow - 2][0];
      }
    }

    var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');

    // En UPDATE preservamos G (Fecha Envío original)
    var fechaEnvioFinal = isUpdate
      ? sh.getRange(targetRow, 7).getValue()
      : (data.fechaEnvio || hoy);

    // A-J (1-10)
    sh.getRange(targetRow, 1, 1, 10).setValues([[
      data.contacto || '',
      data.empresa || '',
      data.pax || '',
      data.estado || 'Pendiente de Enviar',
      data.vendedor || '',
      data.tipoEvento || '',
      fechaEnvioFinal,
      data.fechaEvento || '',
      data.fechaUltContacto || hoy,
      data.mes || ''
    ]]);

    // L-Q (12-17)
    sh.getRange(targetRow, 12, 1, 6).setValues([[
      data.telefono || '',
      data.email || '',
      data.nombreEvento || '',
      data.locacion || '',
      data.detalle || '',
      data.canal || ''
    ]]);

    // R-T (18-20) — solo en INSERT, en UPDATE preservamos lo manual
    if (!isUpdate) {
      sh.getRange(targetRow, 18, 1, 3).setValues([['', '', '']]);
    }

    // U Observaciones (21)
    sh.getRange(targetRow, 21).setValue(data.observaciones || '');

    // V Internal ID (22) — siempre se escribe, permite identificar futuras revisiones
    if (data.presupuestoId) {
      sh.getRange(targetRow, INTERNAL_ID_COL).setValue(data.presupuestoId);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ok: true, id: data.id, row: targetRow, updated: isUpdate}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, msg: 'La Fondiatta Apps Script online v2.1'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// === TEST DEBUG (no toca el Sheet, solo verifica acceso) ===
function test() {
  Logger.log('1. Empezó');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('2. Spreadsheet OK: ' + ss.getName());
  var sh = ss.getSheets()[0];
  Logger.log('3. Sheet OK: ' + sh.getName() + ', filas: ' + sh.getLastRow());
  Logger.log('4. Fin');
}

// === TEST POST (prueba doPost de verdad — INSERT/UPDATE) ===
function testPost() {
  var fakePayload = {
    postData: {
      contents: JSON.stringify({
        contacto: '🧪 TEST CLIENTE',
        empresa: 'TEST SA',
        pax: 50,
        vendedor: 'Colo',
        tipoEvento: 'Desayuno Corporativo',
        fechaEvento: '01/05/2026',
        telefono: '+54 11 1234-5678',
        email: 'test@test.com',
        nombreEvento: 'Prueba de integración',
        locacion: 'CABA',
        detalle: 'Desayuno Premium x 50 · Barra sin alcohol',
        canal: 'Vendedor',
        mes: 'mayo 26',
        observaciones: 'TOTAL c/IVA: $1.500.000',
        presupuestoId: 'p_test_123'
      })
    }
  };
  var t0 = Date.now();
  var result = doPost(fakePayload);
  Logger.log('Tiempo: ' + (Date.now() - t0) + 'ms');
  Logger.log(result.getContent());
}

// ============================================================
// SETUP DEL SHEET — look & feel La Fondiatta
// ============================================================
// Función para correr UNA SOLA VEZ desde el editor.
// Aplica:
//  - Header negro con texto crema (estilo fondiattero)
//  - Dropdowns en Estado (D), Vendedor (E), Tipo de Evento (F)
//  - Date pickers en G, H, I (las 3 columnas de fecha)
//  - Conditional formatting en Estado: verde/amarillo/azul/gris/rojo
//  - Columna V (Internal ID) discreta: gris monospace, ancho 120px
//  - Freeze de la fila 1 y de las 2 primeras columnas
//  - Filas alternadas (banding) crema
//
// Es IDEMPOTENTE: se puede correr varias veces sin romper nada.
// ============================================================
function setupSheetFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = Math.max(sh.getLastRow(), 2);
  var lastCol = 22;

  Logger.log('Empezando setup — ' + lastRow + ' filas, ' + lastCol + ' columnas');

  // === 1. HEADER ROW — fondiattero ===
  var headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground('#0a0a0a')
    .setFontColor('#f4efe6')
    .setFontWeight('bold')
    .setFontSize(11)
    .setFontFamily('Inter')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);
  Logger.log('1. Header listo');

  // === 2. DROPDOWNS ===
  // D — Estado
  var estados = ['Pendiente de Enviar', 'En Seguimiento', 'Avanzado', 'Confirmado', 'Perdido'];
  var ruleEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(estados, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 4, lastRow - 1, 1).setDataValidation(ruleEstado);

  // E — Vendedor
  var vendedores = ['JP', 'Colo', 'Ako', 'Camba', 'Chino', 'Zenon', 'Nahue'];
  var ruleVendedor = SpreadsheetApp.newDataValidation()
    .requireValueInList(vendedores, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 5, lastRow - 1, 1).setDataValidation(ruleVendedor);

  // F — Tipo de Evento
  var tiposEvento = [
    'Desayuno Corporativo', 'Almuerzo Corporativo', 'Merienda Corporativa',
    'After Corporativo', 'Cena Corporativa', 'Desayuno + Almuerzo Corporativo',
    'Evento Corporativo', 'Casamiento', 'Civil', 'Cumpleaños',
    'Evento Social', 'Catering', 'Propuesta General',
    'Entrega Corporativa', 'Entrega Particular'
  ];
  var ruleTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(tiposEvento, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 6, lastRow - 1, 1).setDataValidation(ruleTipo);

  // G, H, I — Date pickers (Fecha Envío, Fecha Evento, Fecha Último Contacto)
  var ruleFecha = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, 7, lastRow - 1, 3).setDataValidation(ruleFecha);

  Logger.log('2. Dropdowns y date pickers listos');

  // === 3. CONDITIONAL FORMATTING en Estado (D) ===
  var coloresEstado = [
    { state: 'Confirmado',         bg: '#d4edda', fg: '#155724' }, // verde
    { state: 'Avanzado',           bg: '#fff3cd', fg: '#856404' }, // amarillo
    { state: 'En Seguimiento',     bg: '#cce5ff', fg: '#004085' }, // azul
    { state: 'Pendiente de Enviar',bg: '#e2e3e5', fg: '#383d41' }, // gris
    { state: 'Perdido',            bg: '#f8d7da', fg: '#721c24' }  // rojo
  ];
  // Limpiar reglas viejas que apunten a columna D
  var existingRules = sh.getConditionalFormatRules();
  var keepRules = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) { return rg.getColumn() === 4; });
  });
  coloresEstado.forEach(function(c) {
    var newRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(c.state)
      .setBackground(c.bg)
      .setFontColor(c.fg)
      .setRanges([sh.getRange(2, 4, lastRow - 1, 1)])
      .build();
    keepRules.push(newRule);
  });
  sh.setConditionalFormatRules(keepRules);
  Logger.log('3. Conditional formatting en Estado listo');

  // === 4. COLUMNA V (Internal ID) — discreta ===
  sh.setColumnWidth(22, 120);
  var vRange = sh.getRange(2, 22, lastRow - 1, 1);
  vRange
    .setFontColor('#9aa0a6')
    .setFontFamily('Roboto Mono')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  // El header de V también más chico
  sh.getRange(1, 22).setFontSize(9);
  Logger.log('4. Columna V formateada');

  // === 5. FREEZE ===
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2); // A (Contacto) + B (Empresa) siempre visibles
  Logger.log('5. Freeze listo');

  // === 6. ROW BANDING (filas alternadas crema) ===
  // Borrar bandings viejos
  var bandings = sh.getBandings();
  bandings.forEach(function(b) { b.remove(); });
  // Aplicar nuevo
  var bandingRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  var banding = bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null);  // ya tenemos header custom
  banding.setFirstRowColor('#ffffff');
  banding.setSecondRowColor('#fbf9f4'); // crema sutil
  Logger.log('6. Banding listo');

  // === 7. ANCHOS DE COLUMNA RAZONABLES ===
  var widths = {
    1: 200,  // A Contacto
    2: 180,  // B Empresa
    3: 60,   // C Pax
    4: 140,  // D Estado
    5: 90,   // E Vendedor
    6: 200,  // F Tipo Evento
    7: 110,  // G Fecha Envío
    8: 110,  // H Fecha Evento
    9: 110,  // I Fecha Último Contacto
    10: 100, // J Mes
    11: 110, // K ID Presupuesto
    12: 130, // L Teléfono
    13: 200, // M Email
    14: 220, // N Nombre Evento
    15: 220, // O Locación
    16: 320, // P Detalle Cotizado
    17: 110, // Q Canal
    18: 110, // R Resultado
    19: 180, // S Motivo Pérdida
    20: 110, // T Calendar
    21: 280, // U Observaciones
    22: 120  // V Internal ID
  };
  Object.keys(widths).forEach(function(col) {
    sh.setColumnWidth(parseInt(col, 10), widths[col]);
  });
  Logger.log('7. Anchos de columna listos');

  Logger.log('✅ Setup completo.');
}
