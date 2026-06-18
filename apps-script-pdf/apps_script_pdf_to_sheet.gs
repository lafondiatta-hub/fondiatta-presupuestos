/**
 * PDF → Seguimiento de Presupuestos LF
 *
 * Apps Script bound al Sheet:
 *   1MrczSE46x3ecpUCuHJsTYI5Gm3VEvp33GtTPyBzsg98
 *
 * Flujo:
 *   1. Web app recibe POST con { filename, pdfBase64 }
 *   2. Llama Gemini con el PDF y un prompt que devuelve JSON con los 22 campos
 *   3. Auto-asigna próximo ID Presupuesto (P-2026-XXX)
 *   4. Inserta nueva fila al final de la Sheet
 *   5. Devuelve a la web los datos extraídos (para preview/edición)
 *
 * Setup:
 *   - File → Project Settings → Script Properties:
 *       GEMINI_API_KEY = <tu key de aistudio.google.com>
 *   - Deploy → New Deployment → Web App
 *       Execute as: Me
 *       Who has access: Anyone (necesario para que el HTML llegue sin auth)
 *   - Copiar la /exec URL al index.html
 */

const SHEET_ID = '1MrczSE46x3ecpUCuHJsTYI5Gm3VEvp33GtTPyBzsg98';
const TAB_GID = 0;
const GEMINI_MODEL = 'gemini-2.5-flash-lite';
const GEMINI_MODEL_FALLBACK = 'gemini-2.5-flash'; // último recurso si flash-lite tira 503 (saturado)

const POSIBLES_CALENDAR_NAME = 'LF Posibles';
const POSIBLES_CALENDAR_COLOR = CalendarApp.EventColor.BLUE; // Azul Arándano (Blueberry, colorId 9)
const CALENDAR_SEARCH_WINDOW_DAYS = 3;

const HEADERS = [
  'Contacto', 'Cliente / Empresa', 'Cantidad Personas', 'Estado Presupuesto',
  'Vendedor', 'Tipo de Evento', 'Fecha Envío', 'Fecha Evento',
  'Fecha Último Contacto', 'Mes', 'ID Presupuesto', 'Teléfono / WhatsApp',
  'Email', 'Nombre Evento', 'Locación', 'Detalle Cotizado',
  'Canal de adquisición', 'Resultado', 'Motivo de Pérdida',
  'Se agendo en Calendar?', 'Observaciones', 'INTERNAL ID'
];

const VENDEDORES = ['Camba', 'Colo', 'JP', 'Ako'];
const TIPOS_EVENTO = [
  'Casamiento', 'Civil', 'Catering', 'Evento Social',
  'Almuerzo Corporativo', 'Cena Corporativa', 'Desayuno Corporativo',
  'After Corporativo', 'Desayuno + Almuerzo Corporativo',
  'Propuesta General', 'Merchansiding'
];

const ESTADOS = [
  '', 'Pendiente de Enviar', 'En Seguimiento', 'Avanzado',
  'Confirmado', 'Perdido', 'Cancelado'
];

const CANALES = [
  '', 'Web', 'Vendedor', 'Cliente Existente', 'Referido',
  'Mail Directo', 'Contagram', 'Vtas La Fondiatta', 'Mailing', 'IG'
];

// ==================================================================
// Web App entry points
// ==================================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Acción: buscar eventos en Calendar cerca de una fecha
    if (body.action === 'searchCalendar') {
      return jsonResponse({
        ok: true,
        result: searchCalendarMatches(body.fechaEvento, body.cliente)
      });
    }

    // Acción: agendar el presupuesto como "POSIBLE" en el Calendar LF Posibles
    if (body.action === 'createPosibleEvent') {
      return jsonResponse({
        ok: true,
        event: createPosibleEvent(body.fechaEvento, body.cliente, body.pax)
      });
    }

    // Acción multi-jornada: crea N POSIBLES (uno por fecha) para la misma
    // propuesta repetida en varias fechas (ej VIP Coliseo × 4 miércoles).
    // Devuelve un array de eventos en el mismo orden que `fechas`.
    // Si una fecha falla (no parsea, etc) el error queda en ese ítem y
    // el resto sigue — no aborta todo.
    if (body.action === 'createPosibleEvents') {
      return jsonResponse({
        ok: true,
        events: createPosibleEventsBulk(body.fechas, body.cliente, body.pax)
      });
    }

    // Acción: actualizar la columna "Se agendo en Calendar?" (T) de un presupuesto
    // ya escrito por el bound v4.0. Reintenta si la fila todavía no apareció
    // (race condition con el bound que recibe el POST en paralelo desde el cotizador).
    if (body.action === 'updateCalendarUrl') {
      return jsonResponse(updateCalendarUrlForPresupuesto(body.presupuestoId, body.calendarEventUrl));
    }

    // Caso 1: registro confirmado con datos editados por el usuario
    //         (no se vuelve a llamar a Gemini, se escribe directo)
    if (body.confirmedData && !body.dryRun) {
      // Multi-jornada: el frontend puede mandar `calendarEventUrls` (array)
      // en lugar de `calendarEventUrl` (singular). Los joineamos con salto
      // de línea para guardar las N URLs en la celda T del Sheet.
      let calendarUrlField = body.calendarEventUrl;
      if (Array.isArray(body.calendarEventUrls) && body.calendarEventUrls.length) {
        calendarUrlField = body.calendarEventUrls.filter(Boolean).join('\n');
      }
      return jsonResponse(registrarConfirmado(body.confirmedData, body.idempotencyKey, calendarUrlField));
    }

    // Caso 2: extracción inicial desde input (PDF / imagen / texto libre)
    const input = resolveInput(body);
    const extracted = extractFromInput(input);
    return jsonResponse({ ok: true, data: extracted, mode: 'preview' });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message, stack: err.stack });
  }
}

function doGet(e) {
  return jsonResponse({
    ok: true,
    service: 'PDF → Seguimiento Presupuestos LF',
    sheetId: SHEET_ID,
    geminiModel: GEMINI_MODEL,
    next_id: nextIdPresupuesto()
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================================================================
// Registro confirmado (con idempotencia + lock)
// ==================================================================

function registrarConfirmado(confirmedData, idempotencyKey, calendarEventUrl) {
  const cache = CacheService.getScriptCache();

  // Si ya escribimos esta key antes, devolver el resultado guardado (no duplicar).
  if (idempotencyKey) {
    const prev = cache.get('idemp_' + idempotencyKey);
    if (prev) {
      return Object.assign(JSON.parse(prev), { mode: 'idempotent_hit' });
    }
  }

  // Lock para evitar que dos uploads simultáneos saquen el mismo ID.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const dataToWrite = Object.assign({}, confirmedData);
    dataToWrite['ID Presupuesto'] = nextIdPresupuesto();
    if (!dataToWrite['Fecha Envío']) dataToWrite['Fecha Envío'] = todayDDMMYYYY();
    if (!dataToWrite['Fecha Último Contacto']) dataToWrite['Fecha Último Contacto'] = todayDDMMYYYY();
    if (!dataToWrite['Estado Presupuesto']) dataToWrite['Estado Presupuesto'] = 'En Seguimiento';
    if (calendarEventUrl) {
      dataToWrite['Se agendo en Calendar?'] = calendarEventUrl;
    }
    const rowNum = appendRow(dataToWrite);

    const result = { ok: true, data: dataToWrite, row: rowNum, mode: 'written' };
    if (idempotencyKey) {
      // 6 min de cache — sobra para cualquier retry razonable.
      cache.put('idemp_' + idempotencyKey, JSON.stringify(result), 360);
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

// ==================================================================
// Gemini extraction (con responseSchema → más rápido y predecible)
// ==================================================================

/**
 * Resuelve el input desde el body del request a un descriptor uniforme:
 *   { kind: 'pdf'|'image'|'text', base64?, mimeType?, content?, filename }
 * Mantiene compat con el frontend viejo que mandaba { pdfBase64, filename }.
 */
function resolveInput(body) {
  // Texto libre (lead pegado al toque)
  if (body.text && String(body.text).trim()) {
    return {
      kind: 'text',
      content: String(body.text).trim(),
      filename: body.filename || 'lead-texto.txt'
    };
  }
  // Imagen (JPG, PNG, WEBP, etc — screenshot de mail/WhatsApp)
  if (body.imageBase64) {
    return {
      kind: 'image',
      base64: body.imageBase64,
      mimeType: body.mimeType || 'image/jpeg',
      filename: body.filename || 'lead-imagen.jpg'
    };
  }
  // PDF (compat con frontend viejo)
  if (body.pdfBase64) {
    return {
      kind: 'pdf',
      base64: body.pdfBase64,
      mimeType: 'application/pdf',
      filename: body.filename || 'lead.pdf'
    };
  }
  throw new Error('Falta input: enviá pdfBase64, imageBase64 o text');
}

// Reintenta la llamada a Gemini ante errores transitorios (503 saturado,
// 429 rate-limit, 500 interno). Un 503 vuelve casi instantáneo, así que
// reintentar es barato. Backoff entre intentos; si flash-lite sigue caído
// tras agotar los reintentos, prueba el modelo de fallback (capacidad
// separada). Errores no transitorios (400/403: API key o payload) cortan
// al toque, sin reintentar al pedo.
function callGeminiWithRetry_(apiKey, payload) {
  const models = [GEMINI_MODEL, GEMINI_MODEL_FALLBACK];
  const backoffMs = [1500, 3000];
  let lastCode = 0, lastBody = '';

  for (let m = 0; m < models.length; m++) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${models[m]}:generateContent?key=${apiKey}`;
    for (let attempt = 0; attempt <= backoffMs.length; attempt++) {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code === 200) return res;

      lastCode = code;
      lastBody = res.getContentText();

      // Solo 503/429/500 son transitorios; el resto (400/403…) falla ya.
      if (code !== 503 && code !== 429 && code !== 500) {
        throw new Error(`Gemini ${code}: ${lastBody.slice(0, 500)}`);
      }
      if (attempt < backoffMs.length) Utilities.sleep(backoffMs[attempt]);
    }
    // Reintentos agotados con este modelo → pasamos al fallback.
  }

  if (lastCode === 503 || lastCode === 429) {
    throw new Error('Gemini está saturado en este momento. Reintentá en un minuto (el documento está OK).');
  }
  throw new Error(`Gemini ${lastCode}: ${String(lastBody).slice(0, 500)}`);
}

function extractFromInput(input) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Falta GEMINI_API_KEY en Script Properties');

  const prompt = buildPrompt(input.kind);

  // El primer "part" depende del tipo: archivo binario o texto plano.
  const firstPart = input.kind === 'text'
    ? { text: 'Lead recibido (texto libre, puede venir de mail/WhatsApp/notas):\n\n' + input.content }
    : { inline_data: { mime_type: input.mimeType, data: input.base64 } };

  const payload = {
    contents: [{
      parts: [firstPart, { text: prompt }]
    }],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: 'application/json',
      response_schema: buildResponseSchema()
    }
  };

  const res = callGeminiWithRetry_(apiKey, payload);
  const json = JSON.parse(res.getContentText());
  const text = json.candidates && json.candidates[0]
    && json.candidates[0].content && json.candidates[0].content.parts
    && json.candidates[0].content.parts[0] && json.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Gemini no devolvió texto');

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error(`Respuesta Gemini no es JSON válido: ${text.slice(0, 300)}`);
  }

  const out = {};
  HEADERS.forEach(h => { out[h] = (parsed[h] != null ? String(parsed[h]) : '').trim(); });
  // Guardia anti-alucinación: el schema obliga a Gemini a devolver algo en
  // "Fecha Evento" aunque el documento no traiga fecha, y a veces copia el
  // nombre del cliente. Si no parece una fecha, se blanquea para que el
  // campo quede vacío en el form y el usuario la cargue a mano.
  if (out['Fecha Evento'] && !looksLikeFecha_(out['Fecha Evento'])) {
    out['Fecha Evento'] = '';
  }
  // Campo extra para multi-jornada (no es columna del Sheet, lo consume el frontend
  // para mostrar chips de fechas y crear N POSIBLES en Calendar).
  out.fechas_individuales = Array.isArray(parsed.fechas_individuales)
    ? parsed.fechas_individuales.map(s => String(s || '').trim()).filter(Boolean)
    : [];
  out.INTERNAL_FILENAME = input.filename;
  out.INTERNAL_SOURCE = input.kind; // 'pdf' | 'image' | 'text'
  return out;
}

// true si el texto contiene algo "de fecha": un dígito, un mes o un día de semana.
function looksLikeFecha_(s) {
  const t = String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (/\d/.test(t)) return true;
  return /\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre|lunes|martes|miercoles|jueves|viernes|sabado|domingo)\b/.test(t);
}

function buildResponseSchema() {
  // Schema OpenAPI subset: garantiza que Gemini devuelva exactamente
  // los 22 campos como strings. No usamos `enum` porque Gemini rechaza
  // strings vacíos dentro de enum, y muchos de estos campos deben poder
  // quedar vacíos. El prompt ya instruye qué valores son válidos.
  //
  // Campo extra `fechas_individuales` (opcional, no es columna del Sheet):
  // se usa cuando el PDF cubre la misma propuesta repetida en N fechas
  // (ej "4 miércoles de junio"). El frontend lo usa para crear N POSIBLES
  // en Calendar manteniendo una sola fila en el Sheet.
  const properties = {};
  HEADERS.forEach(h => {
    properties[h] = { type: 'string' };
  });
  properties.fechas_individuales = {
    type: 'array',
    items: { type: 'string' }
  };
  return {
    type: 'object',
    properties: properties,
    required: HEADERS
  };
}

function buildPrompt(kind) {
  const sourceLabel = kind === 'text'
    ? 'el texto del lead (puede ser informal, abreviado, copiado de mail/WhatsApp/notas)'
    : kind === 'image'
      ? 'la imagen adjunta (screenshot de mail, WhatsApp, propuesta, o foto de papel)'
      : 'el PDF adjunto (propuesta o presupuesto)';
  return [
    'Sos un asistente que extrae datos de leads/propuestas/presupuestos de un catering argentino llamado La Fondiatta.',
    `Tu tarea: analizar ${sourceLabel} y devolver UN ÚNICO objeto JSON (no un array) con EXACTAMENTE estos keys:`,
    '',
    JSON.stringify(HEADERS, null, 2),
    '',
    'Reglas para cada campo:',
    '- "Contacto": persona contactada del lado del cliente (ej: "Daniela Castro", "Ricky Dordoni"). Es la persona, no la empresa.',
    '- "Cliente / Empresa": empresa que paga / cuenta corporativa, o "Particular" para eventos personales. SOLO si aparece explícitamente en el PDF (ej: "Danone", "Microsoft", "PedidosYa"). Si el PDF está dirigido a una persona y no aclara empresa, dejá "" — NO confundas con el lugar del evento. La locación NO es el cliente: "Teatro Coliseo" o "Hotel X" son la sede, no la cuenta.',
    '- "Cantidad Personas": número entero (solo el número, sin "pax"). Si dice rango "60-70", devolvé "60-70". Si son N fechas con el mismo pax por fecha (ej "4 miércoles × 40 pax"), poné el pax POR FECHA, no el agregado.',
    '- "Estado Presupuesto": dejá string vacío "" (lo completa el equipo después).',
    `- "Vendedor": uno de ${JSON.stringify(VENDEDORES)} si aparece firmado o mencionado, sino "".`,
    `- "Tipo de Evento": uno de ${JSON.stringify(TIPOS_EVENTO)}, el que mejor encaje. Si nada encaja, "".`,
    '- "Fecha Envío": "" (lo completa el sistema con la fecha de hoy).',
    '- "Fecha Evento": copiá tal cual aparece (ej: "jueves 4 de junio", "9/05/2026", "26 de junio"). No la conviertas. Tiene que ser una FECHA (día, mes y/o año). Si el documento NO menciona la fecha del evento, devolvé "" — NUNCA pongas un nombre de persona, empresa o lugar acá. Si son MÚLTIPLES fechas de la misma propuesta (ej "3, 10, 17 y 24 de junio" o "4 miércoles de junio: 3, 10, 17, 24"), poné el rango completo legible aquí (ej: "3, 10, 17 y 24 de junio 2026").',
    '- "Fecha Último Contacto": "".',
    '- "Mes": deducir del año/mes del evento en formato "mes 26" (ej: "junio 26", "mayo 26"). Si la fecha no es clara, "". Si son varias fechas del mismo mes, usá ese mes.',
    '- "ID Presupuesto": "" (lo asigna el sistema).',
    '- "Teléfono / WhatsApp": si aparece teléfono del contacto.',
    '- "Email": si aparece email del contacto.',
    '- "Nombre Evento": título o referencia del evento si aparece (ej: "Family Day ZS", "VIP Teatro Coliseo").',
    '- "Locación": dirección o referencia del lugar donde se hace el evento (ej: "Av. del Libertador 2601", "Teatro Coliseo", "Vivanco 1509, Tigre"). Es DISTINTA de "Cliente / Empresa".',
    '- "Detalle Cotizado": resumen CORTO en una línea de las opciones cotizadas (ej: "Tapeo A + Barra Clasica + Dulce + Sonido"). NO copies todo el menú línea por línea, hacé un resumen ejecutivo. Si son N fechas idénticas, agregá " · × N fechas" al final (ej: "Finger food 5 bocados + Salentein · × 4 fechas").',
    '- "Canal de adquisición": "Cliente Existente" si la empresa aparece varias veces en presupuestos LF (Danone, PedidosYa, Microsoft, Albaugh, ZS, etc.), "Vendedor" si fue por contacto del equipo, "Web" si pidieron por la web. Si no es claro, "".',
    '- "Resultado": "".',
    '- "Motivo de Pérdida": "".',
    '- "Se agendo en Calendar?": "".',
    '- "Observaciones": cualquier dato extra relevante que no entró en otro campo (forma de pago especial, requisitos, etc.). Mantenelo corto. Si hay multi-jornada con un total agregado distinto al por-fecha, incluí "$X/fecha · TOTAL $Y + IVA" acá.',
    '- "INTERNAL ID": "".',
    '',
    'CAMPO EXTRA — "fechas_individuales": array de strings (puede ser []).',
    '  - Si el PDF cubre la MISMA propuesta repetida en N fechas (ej "4 miércoles de junio: 3, 10, 17 y 24", "todos los viernes de mayo", "5, 6 y 7 de mayo"), llená este array con cada fecha por separado en formato D/M/YYYY (ej ["3/6/2026","10/6/2026","17/6/2026","24/6/2026"]).',
    '  - Si es una sola fecha, dejalo como [] (array vacío). NO repitas la única fecha.',
    '  - Si son N fechas con menús/pax DISTINTOS (ej civil un día + fiesta otro día), igual incluí las N fechas acá. El usuario decide después si agendar todas o no.',
    '  - Año: si el PDF dice "junio 2026" usá 2026. Si no aclara año, asumí el próximo en el calendario.',
    '',
    'Si algún dato NO está presente, devolvé string vacío "" para ese campo. NO inventes datos.',
    'Si el input es informal (ej: "30 pax / Jorge Lanza / propuesta general / 26 de junio"), igual extraé lo que puedas: nombre del contacto, pax, tipo de evento, fecha, etc. Lo que no esté → "".',
    'Devolvé SOLO el JSON, sin markdown, sin texto adicional.'
  ].join('\n');
}

// ==================================================================
// Sheet helpers
// ==================================================================

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getSheetId() === TAB_GID) return s;
  }
  return ss.getSheets()[0];
}

function nextIdPresupuesto() {
  const sheet = getSheet();
  const colIdx = HEADERS.indexOf('ID Presupuesto') + 1;
  const lastRow = findLastWrittenRow_(sheet);
  if (lastRow < 2) return 'P-2026-001';

  const values = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
  let max = 0;
  values.forEach(r => {
    const v = String(r[0] || '').trim();
    const m = v.match(/^P-2026-(\d+)$/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  });
  const next = max + 1;
  return 'P-2026-' + String(next).padStart(3, '0');
}

// Encuentra la última fila *escrita* mirando columnas de datos reales
// (Contacto, Cliente, ID Presupuesto, INTERNAL ID). Evita el bug de
// `appendRow()` / `getLastRow()` que se "estiran" cuando hay validación,
// formato condicional o formulas vacías muchas filas abajo.
function findLastWrittenRow_(sheet) {
  // K (ID Presupuesto) queda afuera a propósito: si un append crashea
  // entre setValue(K) y los setValues posteriores queda una fila
  // huérfana con solo K. Esa huérfana NO debe inflar lastRow y empujar
  // nuevos entries al final. A/B/V son ancla real.
  const primaryCols = [
    HEADERS.indexOf('Contacto') + 1,            // A
    HEADERS.indexOf('Cliente / Empresa') + 1,    // B
    HEADERS.indexOf('INTERNAL ID') + 1            // V
  ];
  const maxRow = sheet.getMaxRows();
  if (maxRow < 2) return 1;
  const scanTo = Math.min(maxRow, 10000);
  const minCol = Math.min.apply(null, primaryCols);
  const maxCol = Math.max.apply(null, primaryCols);
  const width = maxCol - minCol + 1;
  const values = sheet.getRange(2, minCol, scanTo - 1, width).getValues();
  const offsets = primaryCols.map(c => c - minCol);
  for (let i = values.length - 1; i >= 0; i--) {
    for (let k = 0; k < offsets.length; k++) {
      const v = values[i][offsets[k]];
      if (v !== '' && v !== null && v !== undefined) return i + 2;
    }
  }
  return 1;
}

function appendRow(data) {
  const sheet = getSheet();
  const row = HEADERS.map(h => data[h] != null ? data[h] : '');
  const target = findLastWrittenRow_(sheet) + 1;
  if (target > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }
  sheet.getRange(target, 1, 1, row.length).setValues([row]);
  return target;
}

// Actualiza la columna T ("Se agendo en Calendar?") de la fila que tenga
// `presupuestoId` (internal ID estable del cotizador, formato "p_xxx") en la
// columna V. La fila la escribió el bound v4.0; como el cotizador manda ambos
// requests en paralelo (bound con no-cors, no espera respuesta), puede haber
// una pequeña ventana en la que la fila todavía no existe → reintentamos.
//
// No sobreescribimos si la columna ya tiene un valor distinto: preservamos
// edición manual del Sheet.
function updateCalendarUrlForPresupuesto(presupuestoId, calendarEventUrl) {
  if (!presupuestoId || !calendarEventUrl) {
    return { ok: false, error: 'presupuestoId y calendarEventUrl son requeridos' };
  }
  const INTERNAL_ID_COL = 22; // V
  const CALENDAR_COL = 20;    // T — "Se agendo en Calendar?"
  const sheet = getSheet();

  // Hasta 6 reintentos × 1s = 6 seg de tolerancia para race condition con bound
  for (let attempt = 0; attempt < 6; attempt++) {
    const lastRow = findLastWrittenRow_(sheet);
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, INTERNAL_ID_COL, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0] || '').trim() === String(presupuestoId).trim()) {
          const targetRow = i + 2;
          const cell = sheet.getRange(targetRow, CALENDAR_COL);
          const existing = String(cell.getValue() || '').trim();
          if (existing && existing !== calendarEventUrl) {
            // Ya tenía un link distinto — no pisamos.
            return { ok: true, row: targetRow, mode: 'skipped_existing', existing: existing };
          }
          cell.setValue(calendarEventUrl);
          return { ok: true, row: targetRow, mode: 'written' };
        }
      }
    }
    Utilities.sleep(1000);
  }
  return { ok: false, error: 'No se encontró la fila con presupuestoId=' + presupuestoId + ' después de 6 segundos. Verificá que el bound v4.0 esté escribiendo.' };
}

function todayDDMMYYYY() {
  const d = new Date();
  return `${d.getDate()}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

// ==================================================================
// Google Calendar — match + crear POSIBLES
// ==================================================================
//
// Flujo:
//   1. Frontend manda { action: 'searchCalendar', fechaEvento, cliente }
//      -> Devuelve { date, sameDayEvents, nearbyEvents, posibleExisting }
//   2. Si el usuario decide linkear o crear POSIBLE, el frontend manda
//      { action: 'createPosibleEvent', fechaEvento, cliente, pax }
//      -> Devuelve { eventId, url, title, allDay, alreadyExisted }
//   3. Al confirmar el alta del presupuesto, el frontend pasa
//      calendarEventUrl y se guarda en columna "Se agendo en Calendar?"
//
// Diseño:
//   - El Calendar "LF Posibles" se crea automáticamente la primera vez
//     y su ID queda guardado en Script Properties.
//   - Antes de crear un POSIBLE, se busca por título+fecha para evitar
//     duplicados si el usuario reupload el mismo PDF o hace doble click.
//   - El parser de fechas acepta los formatos típicos de PDFs LF
//     ("9/05/2026", "9-5-2026", "4 de junio", "jueves 4 de junio 2026").
//     Si falla, devuelve null y el frontend muestra el botón deshabilitado.

function ensurePosiblesCalendar() {
  // El Calendar "LF Posibles" vive en la cuenta lafondiatta@gmail.com y se
  // comparte a la cuenta personal de Santi con permiso de escritura. El ID
  // se setea a mano en Script Properties como POSIBLES_CALENDAR_ID — NO se
  // autocrea desde acá para que nunca quede en la cuenta personal por error.
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('POSIBLES_CALENDAR_ID');
  if (!stored) {
    throw new Error(
      'Falta setear POSIBLES_CALENDAR_ID en Script Properties. ' +
      'Crear un Calendar "LF Posibles" en lafondiatta@gmail.com, ' +
      'compartirlo a santiagocambaceres@gmail.com con "Hacer cambios en los eventos", ' +
      'y pegar el Calendar ID en la property.'
    );
  }
  const cal = CalendarApp.getCalendarById(stored);
  if (!cal) {
    throw new Error(
      'POSIBLES_CALENDAR_ID está seteado (' + stored + ') pero no se accede al calendar. ' +
      'Verificar que lafondiatta@gmail.com compartió el calendar con permisos de escritura.'
    );
  }
  return cal;
}

function searchCalendarMatches(fechaEvento, cliente) {
  const date = parseFechaEvento(fechaEvento);
  if (!date) {
    return { ok: false, reason: 'fecha_no_parseada', fechaEvento: fechaEvento || '' };
  }

  const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
  const dayEnd = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
  const windowStart = new Date(dayStart.getTime() - CALENDAR_SEARCH_WINDOW_DAYS * 86400000);
  const windowEnd = new Date(dayEnd.getTime() + CALENDAR_SEARCH_WINDOW_DAYS * 86400000);

  const posiblesCal = ensurePosiblesCalendar();
  const posiblesId = posiblesCal.getId();

  const allCalendars = CalendarApp.getAllCalendars();
  const sameDay = [];
  const nearby = [];
  let posibleExisting = null;

  allCalendars.forEach(cal => {
    let events;
    try {
      events = cal.getEvents(windowStart, windowEnd);
    } catch (err) {
      // Algunos calendarios remotos pueden tirar permisos — los salteamos.
      return;
    }
    events.forEach(ev => {
      const start = ev.getStartTime();
      const evDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const sameDayMatch = (evDay.getTime() === dayStart.getTime());
      const item = {
        eventId: ev.getId(),
        title: ev.getTitle(),
        calendar: cal.getName(),
        calendarId: cal.getId(),
        isPosible: (cal.getId() === posiblesId),
        start: start.toISOString(),
        end: ev.getEndTime().toISOString(),
        allDay: ev.isAllDayEvent(),
        url: buildCalendarEventUrl(ev.getId(), cal.getId()),
        sameDay: sameDayMatch
      };
      if (sameDayMatch) {
        sameDay.push(item);
        // Detectar si ya existe un POSIBLE para este cliente+fecha (para de-dup)
        if (cal.getId() === posiblesId && cliente && titleMatchesCliente(ev.getTitle(), cliente)) {
          posibleExisting = item;
        }
      } else {
        nearby.push(item);
      }
    });
  });

  // Ordenar por fecha
  sameDay.sort((a, b) => a.start.localeCompare(b.start));
  nearby.sort((a, b) => a.start.localeCompare(b.start));

  return {
    ok: true,
    date: dayStart.toISOString(),
    dateLabel: formatDDMMYYYY(dayStart),
    sameDayEvents: sameDay,
    nearbyEvents: nearby,
    posibleExisting: posibleExisting,
    posiblesCalendarId: posiblesId
  };
}

function createPosibleEvent(fechaEvento, cliente, pax) {
  const date = parseFechaEvento(fechaEvento);
  if (!date) throw new Error('No se pudo parsear la fecha del evento: ' + fechaEvento);
  const empresa = (cliente || 'Sin cliente').trim();
  const paxStr = pax ? String(pax).trim() : '';
  const title = paxStr
    ? `Posible Evento ${empresa} (${paxStr}p)`
    : `Posible Evento ${empresa}`;

  const cal = ensurePosiblesCalendar();
  const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());

  // De-dup: si ya existe un POSIBLE en el mismo día con título equivalente,
  // lo reusamos en vez de crear un duplicado.
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existingSameDay = cal.getEventsForDay(dayStart);
    for (const ev of existingSameDay) {
      if (ev.getTitle().trim() === title || titleMatchesCliente(ev.getTitle(), empresa)) {
        return {
          eventId: ev.getId(),
          url: buildCalendarEventUrl(ev.getId(), cal.getId()),
          title: ev.getTitle(),
          calendar: cal.getName(),
          allDay: true,
          alreadyExisted: true
        };
      }
    }

    const event = cal.createAllDayEvent(title, dayStart, {
      description: `Presupuesto LF cargado automáticamente desde cotizador.\nCliente: ${empresa}${paxStr ? `\nPax: ${paxStr}` : ''}\nFecha original PDF: ${fechaEvento}`
    });
    event.setColor(POSIBLES_CALENDAR_COLOR);

    return {
      eventId: event.getId(),
      url: buildCalendarEventUrl(event.getId(), cal.getId()),
      title: event.getTitle(),
      calendar: cal.getName(),
      allDay: true,
      alreadyExisted: false
    };
  } finally {
    lock.releaseLock();
  }
}

function createPosibleEventsBulk(fechas, cliente, pax) {
  if (!Array.isArray(fechas) || fechas.length === 0) {
    throw new Error('createPosibleEventsBulk: fechas debe ser un array no vacío');
  }
  return fechas.map(f => {
    try {
      const ev = createPosibleEvent(f, cliente, pax);
      return Object.assign({ ok: true, fechaInput: f }, ev);
    } catch (err) {
      return { ok: false, fechaInput: f, error: err.message };
    }
  });
}

function buildCalendarEventUrl(eventId, calendarId) {
  // Apps Script devuelve el eventId con formato "abc123@google.com".
  // El base64 del eventId base + calendar va en la URL del visor:
  const base = eventId.split('@')[0];
  // El formato eid de Google Calendar es base64(eventId + " " + calendarId)
  const raw = base + ' ' + calendarId;
  const eid = Utilities.base64Encode(raw).replace(/=+$/, '');
  return 'https://calendar.google.com/calendar/event?eid=' + eid;
}

function titleMatchesCliente(title, cliente) {
  if (!title || !cliente) return false;
  const norm = s => s.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, ' ')
    .trim();
  const t = norm(title);
  const c = norm(cliente);
  if (!c) return false;
  // Match si el cliente está como palabra completa en el título
  return (' ' + t + ' ').indexOf(' ' + c + ' ') >= 0 || t.indexOf(c) >= 0;
}

function parseFechaEvento(raw) {
  if (!raw) return null;
  const s = String(raw).trim().toLowerCase();
  if (!s) return null;

  // Formato DD/MM/YYYY o D/M/YY (con / o - o .)
  let m = s.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);
  if (m) {
    const d = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const date = new Date(y, mo, d);
    if (!isNaN(date.getTime()) && date.getMonth() === mo) return date;
  }

  // Formato texto: "4 de junio", "jueves 4 de junio", "4 de junio de 2026"
  const meses = {
    enero: 0, ene: 0,
    febrero: 1, feb: 1,
    marzo: 2, mar: 2,
    abril: 3, abr: 3,
    mayo: 4, may: 4,
    junio: 5, jun: 5,
    julio: 6, jul: 6,
    agosto: 7, ago: 7,
    septiembre: 8, setiembre: 8, sept: 8, sep: 8,
    octubre: 9, oct: 9,
    noviembre: 10, nov: 10,
    diciembre: 11, dic: 11
  };
  m = s.match(/(\d{1,2})\s*(?:de\s+)?([a-záéíóúñ]+)(?:\s+(?:de\s+)?(\d{2,4}))?/i);
  if (m) {
    const d = parseInt(m[1], 10);
    const mesKey = m[2].normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const mo = meses[mesKey];
    if (mo != null) {
      let y;
      if (m[3]) {
        y = parseInt(m[3], 10);
        if (y < 100) y += 2000;
      } else {
        // Sin año explícito: usamos el año del evento más cercano en el futuro
        const today = new Date();
        y = today.getFullYear();
        const candidate = new Date(y, mo, d);
        if (candidate < today) y++;
      }
      const date = new Date(y, mo, d);
      if (!isNaN(date.getTime()) && date.getMonth() === mo) return date;
    }
  }

  return null;
}

function formatDDMMYYYY(d) {
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

// ==================================================================
// Tests manuales (correr desde el editor)
// ==================================================================

function test_nextId() {
  Logger.log(nextIdPresupuesto());
}

function test_geminiKey() {
  const k = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  Logger.log(k ? `OK key length=${k.length}` : 'FALTA GEMINI_API_KEY');
}

function test_idempotency() {
  const data = { 'Contacto': 'Test', 'Cliente / Empresa': 'Test SA' };
  const r1 = registrarConfirmado(data, 'test-key-' + Date.now());
  Logger.log(JSON.stringify(r1));
}
