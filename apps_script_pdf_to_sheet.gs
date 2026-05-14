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

    // Caso 1: registro confirmado con datos editados por el usuario
    //         (no se vuelve a llamar a Gemini, se escribe directo)
    if (body.confirmedData && !body.dryRun) {
      return jsonResponse(registrarConfirmado(body.confirmedData, body.idempotencyKey, body.calendarEventUrl));
    }

    // Caso 2: extracción inicial desde el PDF (preview)
    if (!body.pdfBase64) throw new Error('Falta pdfBase64');
    const extracted = extractFromPdf(body.pdfBase64, body.filename || 'presupuesto.pdf');
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

function extractFromPdf(pdfBase64, filename) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Falta GEMINI_API_KEY en Script Properties');

  const prompt = buildPrompt();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{
      parts: [
        { inline_data: { mime_type: 'application/pdf', data: pdfBase64 } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      response_mime_type: 'application/json',
      response_schema: buildResponseSchema()
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error(`Gemini ${code}: ${res.getContentText().slice(0, 500)}`);
  }

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
  out.INTERNAL_FILENAME = filename;
  return out;
}

function buildResponseSchema() {
  // Schema OpenAPI subset: garantiza que Gemini devuelva exactamente
  // los 22 campos como strings. No usamos `enum` porque Gemini rechaza
  // strings vacíos dentro de enum, y muchos de estos campos deben poder
  // quedar vacíos. El prompt ya instruye qué valores son válidos.
  const properties = {};
  HEADERS.forEach(h => {
    properties[h] = { type: 'string' };
  });
  return {
    type: 'object',
    properties: properties,
    required: HEADERS
  };
}

function buildPrompt() {
  return [
    'Sos un asistente que extrae datos de propuestas/presupuestos PDF de un catering argentino llamado La Fondiatta.',
    'Tu tarea: leer el PDF adjunto y devolver UN ÚNICO objeto JSON (no un array) con EXACTAMENTE estos keys:',
    '',
    JSON.stringify(HEADERS, null, 2),
    '',
    'Reglas para cada campo:',
    '- "Contacto": persona contactada del lado del cliente (ej: "Daniela Castro").',
    '- "Cliente / Empresa": empresa o "Particular" para eventos sociales (ej: "Danone", "PedidosYa", "Particular").',
    '- "Cantidad Personas": número entero (solo el número, sin "pax"). Si dice rango "60-70", devolvé "60-70".',
    '- "Estado Presupuesto": dejá string vacío "" (lo completa el equipo después).',
    `- "Vendedor": uno de ${JSON.stringify(VENDEDORES)} si aparece firmado o mencionado, sino "".`,
    `- "Tipo de Evento": uno de ${JSON.stringify(TIPOS_EVENTO)}, el que mejor encaje. Si nada encaja, "".`,
    '- "Fecha Envío": "" (lo completa el sistema con la fecha de hoy).',
    '- "Fecha Evento": copiá tal cual aparece en el PDF (ej: "jueves 4 de junio", "9/05/2026"). No la conviertas.',
    '- "Fecha Último Contacto": "".',
    '- "Mes": deducir del año/mes del evento en formato "mes 26" (ej: "junio 26", "mayo 26"). Si la fecha no es clara, "".',
    '- "ID Presupuesto": "" (lo asigna el sistema).',
    '- "Teléfono / WhatsApp": si aparece teléfono del contacto.',
    '- "Email": si aparece email del contacto.',
    '- "Nombre Evento": título o referencia del evento si aparece (ej: "Family Day ZS").',
    '- "Locación": dirección o referencia (ej: "Av. del Libertador 2601" o "Vivanco 1509, Tigre").',
    '- "Detalle Cotizado": resumen CORTO en una línea de las opciones cotizadas (ej: "Tapeo A + Barra Clasica + Dulce + Sonido"). NO copies todo el menú línea por línea, hacé un resumen ejecutivo.',
    '- "Canal de adquisición": "Cliente Existente" si la empresa aparece varias veces en presupuestos LF (Danone, PedidosYa, Microsoft, Albaugh, ZS, etc.), "Vendedor" si fue por contacto del equipo, "Web" si pidieron por la web. Si no es claro, "".',
    '- "Resultado": "".',
    '- "Motivo de Pérdida": "".',
    '- "Se agendo en Calendar?": "".',
    '- "Observaciones": cualquier dato extra relevante que no entró en otro campo (forma de pago especial, requisitos, etc.). Mantenelo corto.',
    '- "INTERNAL ID": "".',
    '',
    'Si algún dato NO está en el PDF, devolvé string vacío "" para ese campo. NO inventes datos.',
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
  const lastRow = sheet.getLastRow();
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

function appendRow(data) {
  const sheet = getSheet();
  const row = HEADERS.map(h => data[h] != null ? data[h] : '');
  sheet.appendRow(row);
  return sheet.getLastRow();
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
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('POSIBLES_CALENDAR_ID');
  if (stored) {
    const cal = CalendarApp.getCalendarById(stored);
    if (cal) return cal;
    // Si el ID guardado ya no existe (cal borrado), lo recreamos.
  }
  const cal = CalendarApp.createCalendar(POSIBLES_CALENDAR_NAME, {
    summary: 'Presupuestos cargados como evento POSIBLE desde el cotizador LF.'
  });
  props.setProperty('POSIBLES_CALENDAR_ID', cal.getId());
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
