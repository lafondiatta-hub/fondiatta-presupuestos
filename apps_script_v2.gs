// ============================================================
// La Fondiatta — Receptor de Presupuestos desde el Generador
// v4.0 — Monto en col X + endpoint ?action=dashboard
// ============================================================
//
// COLUMNAS DEL SHEET:
//   A  Contacto            B  Cliente/Empresa      C  Pax
//   D  Estado              E  Vendedor              F  Tipo Evento
//   G  Fecha Envío         H  Fecha Evento          I  Fecha Último Contacto
//   J  Mes                 K  ID Presupuesto        L  Teléfono
//   M  Email               N  Nombre Evento         O  Locación
//   P  Detalle Cotizado     Q  Canal                 R  Resultado
//   S  Motivo Pérdida      T  Calendar?             U  Observaciones
//   V  Internal ID         W  Items JSON            X  Monto (NUEVO v4.0)
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
    // Usamos findLastWrittenRow_ — getLastRow() se infla cuando hay validación
    // o formato condicional aplicado a muchas filas abajo, y el append cae
    // muy lejos de la última fila con datos reales.
    var lastRow = findLastWrittenRow_(sh);

    var ID_PRES_COL    = 11; // K
    var INTERNAL_ID_COL = 22; // V
    var ITEMS_JSON_COL  = 23; // W
    var MONTO_COL       = 24; // X — nuevo

    // === DEDUP: si llega presupuestoId y matchea, ACTUALIZAR ===
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

    // W — Items JSON
    if (data.itemsJson) {
      sh.getRange(targetRow, ITEMS_JSON_COL).setValue(data.itemsJson);
    }

    // X — Monto (total calculado por la app)
    if (data.monto !== undefined && data.monto !== null && data.monto !== '') {
      sh.getRange(targetRow, MONTO_COL).setValue(Number(data.monto) || 0);
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
// doGet — health check + list (restore) + dashboard
// ============================================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'list') return listPresupuestos(e);
    if (action === 'dashboard') return dashboardData(e);
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, msg: 'La Fondiatta Apps Script online v4.0'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// dashboardData — todos los presupuestos con vendedor, para el
// dashboard de métricas. No filtra por internal ID.
// ============================================================
function dashboardData(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = findLastWrittenRow_(sh);
  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, items: []}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Leer todas las columnas A-X (24 cols)
  var data = sh.getRange(2, 1, lastRow - 1, 24).getValues();
  var items = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var vendedor = String(row[4] || '').trim();
    if (!vendedor) continue; // filas sin vendedor no aportan al dashboard

    items.push({
      contacto:      String(row[0]  || ''),
      empresa:       String(row[1]  || ''),
      pax:           row[2] || '',
      estado:        String(row[3]  || ''),
      vendedor:      vendedor,
      tipoEvento:    String(row[5]  || ''),
      fechaEnvio:    formatDate(row[6]),
      fechaEvento:   formatDate(row[7]),
      mes:           String(row[9]  || ''),
      idPresupuesto: String(row[10] || ''),
      canal:         String(row[16] || ''),
      resultado:     String(row[17] || ''),
      motivoPerdida: String(row[18] || ''),
      monto:         Number(row[23]) || 0
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ok: true, items: items}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// listPresupuestos — para restaurar presupuestos en la app
// ============================================================
function listPresupuestos(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = findLastWrittenRow_(sh);
  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, items: []}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var vendedorFiltro = ((e.parameter.vendedor || '').trim()).toLowerCase();
  var limit = parseInt(e.parameter.limit, 10) || 50;

  var data = sh.getRange(2, 1, lastRow - 1, 24).getValues();

  var items = [];
  for (var i = data.length - 1; i >= 0 && items.length < limit; i--) {
    var row = data[i];
    var internalId = row[21]; // V
    if (!internalId) continue;
    if (String(internalId).indexOf('p_') !== 0) continue;

    var vendedorRow = String(row[4] || '');
    if (vendedorFiltro && vendedorRow.toLowerCase() !== vendedorFiltro) continue;

    var itemsJson = row[22]; // W
    var itemsParsed = null;
    if (itemsJson) {
      try { itemsParsed = JSON.parse(itemsJson); } catch (err) { itemsParsed = null; }
    }

    items.push({
      internalId:    internalId,
      idPresupuesto: row[10],
      contacto:      row[0],
      empresa:       row[1],
      pax:           row[2],
      estado:        row[3],
      vendedor:      vendedorRow,
      tipoEvento:    row[5],
      fechaEnvio:    formatDate(row[6]),
      fechaEvento:   formatDate(row[7]),
      mes:           row[9],
      telefono:      row[11],
      email:         row[12],
      nombreEvento:  row[13],
      locacion:      row[14],
      canal:         row[16],
      observaciones: row[20],
      monto:         Number(row[23]) || 0,
      items:         itemsParsed
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

// Encuentra la última fila *escrita* mirando columnas de datos reales
// (A Contacto, B Cliente, K ID Presupuesto, V Internal ID). Evita el bug
// de getLastRow() que se "estira" cuando hay validación, CF o formulas
// vacías muchas filas abajo y el append cae lejos de la data real.
function findLastWrittenRow_(sh) {
  var primaryCols = [1, 2, 11, 22]; // A, B, K, V
  var maxRow = sh.getMaxRows();
  if (maxRow < 2) return 1;
  var scanTo = Math.min(maxRow, 10000);
  var minCol = Math.min.apply(null, primaryCols);
  var maxCol = Math.max.apply(null, primaryCols);
  var width = maxCol - minCol + 1;
  var values = sh.getRange(2, minCol, scanTo - 1, width).getValues();
  var offsets = primaryCols.map(function(c) { return c - minCol; });
  for (var i = values.length - 1; i >= 0; i--) {
    for (var k = 0; k < offsets.length; k++) {
      var v = values[i][offsets[k]];
      if (v !== '' && v !== null && v !== undefined) return i + 2;
    }
  }
  return 1;
}

// ============================================================
// SETUP DEL SHEET — one-shot, incluye columna X Monto
// ============================================================
function setupSheetFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[0];
  var lastRow = Math.max(sh.getLastRow(), 2);
  var lastCol = 24; // ahora incluye col X

  Logger.log('Empezando setup v4.0 — ' + lastRow + ' filas');

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

  // Header de col X
  sh.getRange(1, 24).setValue('Monto');

  // 2. DROPDOWNS
  var estados = ['Pendiente de Enviar', 'En Seguimiento', 'Avanzado', 'Confirmado', 'Perdido', 'Cancelado'];
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

  // 3. CONDITIONAL FORMATTING en Estado
  var coloresEstado = [
    { state: 'Confirmado',          bg: '#d4edda', fg: '#155724' },
    { state: 'Avanzado',            bg: '#fff3cd', fg: '#856404' },
    { state: 'En Seguimiento',      bg: '#cce5ff', fg: '#004085' },
    { state: 'Pendiente de Enviar', bg: '#e2e3e5', fg: '#383d41' },
    { state: 'Perdido',             bg: '#f8d7da', fg: '#721c24' },
    { state: 'Cancelado',           bg: '#f0f0f0', fg: '#555555' }
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

  // 4. Formato número para col X (Monto)
  sh.getRange(2, 24, lastRow - 1, 1)
    .setNumberFormat('$ #,##0');

  // 5. Columnas internas discretas
  sh.setColumnWidth(22, 120);
  sh.setColumnWidth(23, 100);
  sh.setColumnWidth(24, 130);
  sh.getRange(2, 22, lastRow - 1, 2)
    .setFontColor('#9aa0a6')
    .setFontFamily('Roboto Mono')
    .setFontSize(9);
  sh.getRange(1, 22, 1, 2).setFontSize(9);

  // 6. FREEZE
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  // 7. BANDING
  sh.getBandings().forEach(function(b) { b.remove(); });
  var banding = sh.getRange(2, 1, lastRow - 1, lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null);
  banding.setFirstRowColor('#ffffff');
  banding.setSecondRowColor('#fbf9f4');

  // 8. ANCHOS
  var widths = {
    1:200, 2:180, 3:60, 4:140, 5:90, 6:200, 7:110, 8:110, 9:110,
    10:100, 11:110, 12:130, 13:200, 14:220, 15:220, 16:320, 17:110,
    18:110, 19:180, 20:110, 21:280, 22:120, 23:100, 24:130
  };
  Object.keys(widths).forEach(function(col) {
    sh.setColumnWidth(parseInt(col, 10), widths[col]);
  });

  Logger.log('✅ Setup v4.0 completo — columna X (Monto) agregada.');
}

// === TESTS ===
function test() {
  Logger.log('1. Empezó');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('2. Spreadsheet OK: ' + ss.getName());
  var sh = ss.getSheets()[0];
  Logger.log('3. Sheet OK: ' + sh.getName() + ', filas: ' + sh.getLastRow());
  Logger.log('4. Fin');
}

function testPost() {
  var fakePayload = {
    postData: {
      contents: JSON.stringify({
        contacto: '🧪 TEST v4.0',
        empresa: 'TEST SA',
        pax: 50,
        vendedor: 'Colo',
        tipoEvento: 'Desayuno Corporativo',
        fechaEvento: '01/05/2026',
        telefono: '+54 11 1234-5678',
        email: 'test@test.com',
        nombreEvento: 'Prueba con Monto',
        locacion: 'CABA',
        detalle: 'Desayuno Premium x 50',
        canal: 'Vendedor',
        mes: 'mayo 26',
        monto: 1500000,
        observaciones: 'TOTAL s/IVA: $1.500.000',
        presupuestoId: 'p_test_v4_123',
        itemsJson: JSON.stringify([
          {cat:'Corpo · Desayunos', name:'Desayuno Premium', price:30000, qty:50, unit:'persona'}
        ])
      })
    }
  };
  var result = doPost(fakePayload);
  Logger.log(result.getContent());
}

function testDashboard() {
  var fakeEvent = { parameter: { action: 'dashboard' } };
  var result = doGet(fakeEvent);
  var parsed = JSON.parse(result.getContent());
  Logger.log('Items: ' + parsed.items.length);
  if (parsed.items.length > 0) Logger.log('Primero: ' + JSON.stringify(parsed.items[0]));
}
