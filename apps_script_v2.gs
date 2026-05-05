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
