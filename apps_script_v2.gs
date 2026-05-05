// ============================================================
// La Fondiatta — Receptor de Presupuestos desde el Generador
// v2 — soporta UPDATE de presupuestos via presupuestoId (UUID interno)
// ============================================================
//
// Cambios vs v1:
// - Si llega presupuestoId y matchea con la columna V, ACTUALIZA esa fila
//   en lugar de insertar una nueva. Soluciona el problema de presupuestos
//   duplicados cuando el cliente pide cambios y el vendedor re-exporta.
// - Preserva G (Fecha Envío original) y R/S/T (campos manuales) en UPDATE.
// - Siempre escribe el presupuestoId en columna V para futuras revisiones.
// - Devuelve `updated: true/false` en la respuesta para diagnóstico.
//
// REQUISITO: El Sheet debe tener una columna V con header "Internal ID".
// ============================================================

function doPost(e) {
  try {
    // Parsear el payload (viene como text/plain)
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      // Fallback: puede venir como form-encoded
      data = JSON.parse(e.parameter.payload || '{}');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheets()[0]; // primera hoja del Sheet
    var lastRow = sh.getLastRow();

    var ID_PRES_COL = 11;     // K = ID Presupuesto (P-2026-XXX)
    var INTERNAL_ID_COL = 22; // V = Internal ID (UUID interno)

    // === DEDUP: si llega presupuestoId y matchea, ACTUALIZAR fila existente ===
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

    // === Si no fue UPDATE: buscar fila pre-generada con contacto vacío ===
    if (!targetRow) {
      var ids = sh.getRange(2, ID_PRES_COL, Math.max(lastRow - 1, 1), 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        var id = ids[i][0];
        var contacto = sh.getRange(i + 2, 1).getValue();
        if (id && String(id).indexOf('P-2026-') === 0 && !contacto) {
          targetRow = i + 2;
          break;
        }
      }

      // Si no hay pre-generado libre, agregar al final
      if (!targetRow) {
        targetRow = lastRow + 1;
        var newId = 'P-2026-' + String(targetRow - 1).padStart(3, '0');
        sh.getRange(targetRow, ID_PRES_COL).setValue(newId);
        data.id = newId;
      } else {
        data.id = sh.getRange(targetRow, ID_PRES_COL).getValue();
      }
    }

    // Fecha hoy formato dd/MM/yyyy (Argentina)
    var hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');

    // En UPDATE: preservar Fecha Envío (G) original, refrescar Fecha Último Contacto (I).
    // En INSERT: usar fecha de hoy para ambas (o la que venga en el payload).
    var fechaEnvioFinal = isUpdate
      ? sh.getRange(targetRow, 7).getValue()
      : (data.fechaEnvio || hoy);

    // Columnas A-J (1-10)
    sh.getRange(targetRow, 1, 1, 10).setValues([[
      data.contacto || '',                      // A Contacto
      data.empresa || '',                       // B Cliente/Empresa
      data.pax || '',                           // C Cant Personas
      data.estado || 'Pendiente de Enviar',     // D Estado
      data.vendedor || '',                      // E Vendedor
      data.tipoEvento || '',                    // F Tipo Evento
      fechaEnvioFinal,                          // G Fecha Envío (preservada en UPDATE)
      data.fechaEvento || '',                   // H Fecha Evento
      data.fechaUltContacto || hoy,             // I Fecha Último Contacto
      data.mes || ''                            // J Mes
    ]]);

    // Columna K ya tiene el ID (no se toca)

    // Columnas L-Q (12-17) — datos del form
    sh.getRange(targetRow, 12, 1, 6).setValues([[
      data.telefono || '',                      // L Teléfono
      data.email || '',                         // M Email
      data.nombreEvento || '',                  // N Nombre Evento
      data.locacion || '',                      // O Locación
      data.detalle || '',                       // P Detalle Cotizado
      data.canal || ''                          // Q Canal
    ]]);

    // Columnas R-T (18-20) — Resultado, Motivo Pérdida, Calendar
    // Solo se inicializan en INSERT. En UPDATE se preservan los valores
    // que el vendedor pudo haber escrito a mano en el Sheet.
    if (!isUpdate) {
      sh.getRange(targetRow, 18, 1, 3).setValues([['', '', '']]);
    }

    // Columna U (21) — Observaciones (siempre se actualiza)
    sh.getRange(targetRow, 21).setValue(data.observaciones || '');

    // Columna V (22) — Internal ID (UUID), permite identificar futuras revisiones
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

// Permite chequear si el script está online (sirve para test desde el browser)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, msg: 'La Fondiatta Apps Script online v2'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// === TEST MANUAL ===
// Corré esta función desde el editor (Select function → test → Run)
// para verificar que escribe bien en el Sheet antes del Deploy.
//
// IMPORTANTE: el presupuestoId 'p_test_123' es fijo, así que:
//   - 1ra ejecución → INSERT (crea fila nueva)
//   - 2da ejecución → UPDATE (actualiza la misma fila)
// Cambiá el contacto a "🧪 TEST CLIENTE 2" antes de la 2da run para verificar.
function test() {
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
  var result = doPost(fakePayload);
  Logger.log(result.getContent());
}
