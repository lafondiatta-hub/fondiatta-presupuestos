// ============================================================
// La Fondiatta — Receptor de Presupuestos desde el Generador
// v3.0 — UPDATE + listado para restore + items JSON en columna W
// ============================================================
//
// REQUISITOS DEL SHEET:
//   Columna V "Internal ID"   (UUID interno, agregar manual)
//   Columna W "Items JSON"    (ítems serializados, agregar manual)
//
// Cambios v3.0 vs v2.1:
// - doPost ahora guarda itemsJson (payload.itemsJson) en columna W.
// - doGet?action=list devuelve los últimos N presupuestos con todos los
//   campos necesarios para repopular el formulario (incluye items).
// - doGet?action=list&vendedor=Colo filtra por vendedor.
// - doGet sin action sigue siendo health check.
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
    var INTERNAL_ID_COL = 22; // V = Internal ID (UUID)
    var ITEMS_JSON_COL = 23;  // W = Items JSON

    // === DEDUP: si llega presupuestoId y matchea, ACTUALIZAR ===
    // Solo matchea si llega un presupuestoId con formato válido (empieza con 'p_').
    // Eso evita falsos positivos contra notas viejas que pueda tener la columna V.
    var targetRow = null;
    var isUpdate = false;
    var presuId = data.presupuestoId && String(data.presupuestoId).indexOf('p_') === 0
      ? data.presupuestoId : '';
    if (presuId && lastRow >= 2) {
      var internalIds = sh.getRange(2, INTERNAL_ID_COL, lastRow - 1, 1).getValues();
      for (var j = 0; j < internalIds.length; j++) {
        if (internalIds[j][0] === presuId) {
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

    var fechaEnvioFinal = isUpdate
      ? sh.getRange(targetRow, 7).getValue()
      : (data.fechaEnvio || hoy);

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

    sh.getRange(targetRow, 12, 1, 6).setValues([[
      data.telefono || '',
      data.email || '',
      data.nombreEvento || '',
      data.locacion || '',
      data.detalle || '',
      data.canal || ''
    ]]);

    if (!isUpdate) {
      sh.getRange(targetRow, 18, 1, 3).setValues([['', '', '']]);
    }

    sh.getRange(targetRow, 21).setValue(data.observaciones || '');

    // V — Internal ID
    if (data.presupuestoId) {
      sh.getRange(targetRow, INTERNAL_ID_COL).setValue(data.presupuestoId);
    }

    // W — Items JSON (para poder restaurar el presupuesto)
    if (data.itemsJson) {
      sh.getRange(targetRow, ITEMS_JSON_COL).setValue(data.itemsJson);
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

// ============================================================
// doGet — health check + listado de presupuestos para restore
// ============================================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'list') {
      return listPresupuestos(e);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, msg: 'La Fondiatta Apps Script online v3.0'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function listPresupuestos(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, items: []}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var vendedorFiltro = ((e.parameter.vendedor || '').trim()).toLowerCase();
  var limit = parseInt(e.parameter.limit, 10) || 50;

  // Lectura BATCH de las 23 columnas (A-W)
  var data = sh.getRange(2, 1, lastRow - 1, 23).getValues();

  var items = [];
  // Iteramos de la última fila hacia atrás (más recientes primero)
  for (var i = data.length - 1; i >= 0 && items.length < limit; i--) {
    var row = data[i];
    var internalId = row[21];   // V
    if (!internalId) continue;  // sin internal ID no se puede restaurar
    // Filtro: solo aceptar IDs con formato UUID generado (p_xxx_yyy).
    // Esto excluye notas/textos viejos que pudieran existir en la columna V.
    if (String(internalId).indexOf('p_') !== 0) continue;

    var vendedorRow = String(row[4] || '');
    if (vendedorFiltro && vendedorRow.toLowerCase() !== vendedorFiltro) continue;

    var itemsJson = row[22];    // W
    var itemsParsed = null;
    if (itemsJson) {
      try { itemsParsed = JSON.parse(itemsJson); } catch (err) { itemsParsed = null; }
    }

    items.push({
      internalId: internalId,
      idPresupuesto: row[10],   // K
      contacto: row[0],         // A
      empresa: row[1],          // B
      pax: row[2],              // C
      estado: row[3],           // D
      vendedor: vendedorRow,    // E
      tipoEvento: row[5],       // F
      fechaEnvio: formatDate(row[6]),    // G
      fechaEvento: formatDate(row[7]),   // H
      mes: row[9],              // J
      telefono: row[11],        // L
      email: row[12],           // M
      nombreEvento: row[13],    // N
      locacion: row[14],        // O
      canal: row[16],           // Q
      observaciones: row[20],   // U
      items: itemsParsed
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok: true, items: items}))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  }
  return String(v);
}

// === TEST DEBUG ===
function test() {
  Logger.log('1. Empezó');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('2. Spreadsheet OK: ' + ss.getName());
  var sh = ss.getSheets()[0];
  Logger.log('3. Sheet OK: ' + sh.getName() + ', filas: ' + sh.getLastRow());
  Logger.log('4. Fin');
}

// === TEST POST ===
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
        presupuestoId: 'p_test_123',
        itemsJson: JSON.stringify([
          {cat:'Corpo · Desayunos', name:'Desayuno Premium', price:25500, qty:50, unit:'persona', desc:'', custom:false, tipoServicio:'Entrada'}
        ])
      })
    }
  };
  var t0 = Date.now();
  var result = doPost(fakePayload);
  Logger.log('Tiempo: ' + (Date.now() - t0) + 'ms');
  Logger.log(result.getContent());
}

// === TEST LIST ===
function testList() {
  var fakeEvent = { parameter: { action: 'list', limit: 5 } };
  var t0 = Date.now();
  var result = doGet(fakeEvent);
  Logger.log('Tiempo: ' + (Date.now() - t0) + 'ms');
  Logger.log(result.getContent());
}

// ============================================================
// SETUP DEL SHEET — look & feel La Fondiatta (one-shot)
// ============================================================
function setupSheetFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = Math.max(sh.getLastRow(), 2);
  var lastCol = 23;

  Logger.log('Empezando setup — ' + lastRow + ' filas');

  // 1. HEADER
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

  // 2. DROPDOWNS
  var estados = ['Pendiente de Enviar', 'En Seguimiento', 'Avanzado', 'Confirmado', 'Perdido'];
  sh.getRange(2, 4, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(estados, true).setAllowInvalid(true).build()
  );

  var vendedores = ['JP', 'Colo', 'Ako', 'Camba', 'Chino', 'Zenon', 'Nahue'];
  sh.getRange(2, 5, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(vendedores, true).setAllowInvalid(true).build()
  );

  var tiposEvento = [
    'Desayuno Corporativo', 'Almuerzo Corporativo', 'Merienda Corporativa',
    'After Corporativo', 'Cena Corporativa', 'Desayuno + Almuerzo Corporativo',
    'Evento Corporativo', 'Casamiento', 'Civil', 'Cumpleaños',
    'Evento Social', 'Catering', 'Propuesta General',
    'Entrega Corporativa', 'Entrega Particular'
  ];
  sh.getRange(2, 6, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(tiposEvento, true).setAllowInvalid(true).build()
  );

  var ruleFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  sh.getRange(2, 7, lastRow - 1, 3).setDataValidation(ruleFecha);

  // 3. CONDITIONAL FORMATTING en Estado
  var coloresEstado = [
    { state: 'Confirmado',         bg: '#d4edda', fg: '#155724' },
    { state: 'Avanzado',           bg: '#fff3cd', fg: '#856404' },
    { state: 'En Seguimiento',     bg: '#cce5ff', fg: '#004085' },
    { state: 'Pendiente de Enviar',bg: '#e2e3e5', fg: '#383d41' },
    { state: 'Perdido',            bg: '#f8d7da', fg: '#721c24' }
  ];
  var existingRules = sh.getConditionalFormatRules();
  var keepRules = existingRules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rg) { return rg.getColumn() === 4; });
  });
  coloresEstado.forEach(function(c) {
    keepRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(c.state)
        .setBackground(c.bg)
        .setFontColor(c.fg)
        .setRanges([sh.getRange(2, 4, lastRow - 1, 1)])
        .build()
    );
  });
  sh.setConditionalFormatRules(keepRules);

  // 4. COLUMNA V (Internal ID) y W (Items JSON) discretas
  sh.setColumnWidth(22, 120);
  sh.setColumnWidth(23, 100);
  sh.getRange(2, 22, lastRow - 1, 2)
    .setFontColor('#9aa0a6')
    .setFontFamily('Roboto Mono')
    .setFontSize(9);
  sh.getRange(1, 22, 1, 2).setFontSize(9);

  // 5. FREEZE
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  // 6. BANDING
  sh.getBandings().forEach(function(b) { b.remove(); });
  var banding = sh.getRange(2, 1, lastRow - 1, lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null);
  banding.setFirstRowColor('#ffffff');
  banding.setSecondRowColor('#fbf9f4');

  // 7. ANCHOS
  var widths = {1:200,2:180,3:60,4:140,5:90,6:200,7:110,8:110,9:110,10:100,11:110,12:130,13:200,14:220,15:220,16:320,17:110,18:110,19:180,20:110,21:280,22:120,23:100};
  Object.keys(widths).forEach(function(col) {
    sh.setColumnWidth(parseInt(col, 10), widths[col]);
  });

  Logger.log('✅ Setup completo.');
}
