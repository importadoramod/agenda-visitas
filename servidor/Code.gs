/**
 * Agenda de Visitas — servidor (Google Apps Script)
 *
 * Guarda las citas en una hoja de Google y pone cada cita en Google Calendar
 * con sus alarmas. Corre con la cuenta de Google de quien lo instala, así que
 * las alarmas llegan al celular de esa persona sin depender de nada más.
 *
 * Instalación: ver INSTALAR.md. Resumen:
 *   1. Ejecutar configurar() una vez (crea la hoja y la clave).
 *   2. Implementar → Nueva implementación → Aplicación web
 *      (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *   3. Ejecutar enlace() y abrir en el celular el enlace que aparece en el registro.
 */

const COLS = ['id', 'fecha', 'hora', 'cliente', 'ciudad', 'tipo', 'mercaderia', 'monto', 'nota',
  'estado', 'creado', 'actualizado', 'calId', 'eventId', 'opticaId'];
const VERSION_SERVIDOR = 2;
// Tablas que la app puede guardar sin cambiar este código (ópticas, abonos y las que vengan después)
const TABLA_OK = /^[a-z][a-z_]{1,29}$/;
const HOJA = 'Citas';
const PROPS = PropertiesService.getScriptProperties();

/* ===================== instalación ===================== */

// Ejecutar UNA vez desde el editor (botón ▶ Ejecutar).
function configurar() {
  let id = PROPS.getProperty('SHEET_ID');
  if (!id || !abrirLibro_(id)) {
    const libro = SpreadsheetApp.create('Agenda de Visitas — datos');
    id = libro.getId();
    PROPS.setProperty('SHEET_ID', id);
  }
  hoja_();
  if (!PROPS.getProperty('TOKEN')) {
    PROPS.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  }
  CalendarApp.getDefaultCalendar(); // pide el permiso de Calendar ahora y no después
  Logger.log('Listo. Hoja de datos: https://docs.google.com/spreadsheets/d/' + id);
  Logger.log('Ahora: Implementar → Nueva implementación → Aplicación web. Luego ejecute enlace().');
}

// Ejecutar DESPUÉS de implementar la aplicación web. Muestra el enlace para el celular.
// Cambie APP si publicó la app en otra dirección.
function enlace() {
  const APP = PROPS.getProperty('APP_URL') || 'https://importadoramod.github.io/agenda-visitas/';
  // Desde el editor, Google suele dar la dirección de prueba (/dev), que solo sirve con la sesión abierta.
  // Por eso se usa la dirección /exec guardada con guardarDireccionServidor().
  const guardada = PROPS.getProperty('SERVER_URL');
  const url = (guardada && guardada.indexOf('PEGUE') !== 0 ? guardada : '') || ScriptApp.getService().getUrl();
  if (!url || /\/dev$/.test(url)) {
    Logger.log('Falta la dirección real del servidor (termina en /exec).');
    Logger.log('Implementar → Gestionar implementaciones → URL → Copiar. Péguela en guardarDireccionServidor(), ejecútela y vuelva a ejecutar enlace().');
    return;
  }
  const link = APP + '#s=' + encodeURIComponent(url) + '&t=' + PROPS.getProperty('TOKEN');
  Logger.log('Abra este enlace en el celular y en la computadora (solo la primera vez):');
  Logger.log(link);
  Logger.log('Si prefiere escribirlo a mano → Servidor: ' + url + '  ·  Clave: ' + PROPS.getProperty('TOKEN'));
}

// Pegue aquí la URL de Implementar → Gestionar implementaciones (termina en /exec) y ejecute una vez.
function guardarDireccionServidor() {
  PROPS.setProperty('SERVER_URL', 'PEGUE-AQUI-LA-URL-QUE-TERMINA-EN-/exec');
}

// Opcional: guarda la dirección donde publicó la app, para que enlace() la use.
function guardarDireccionApp() {
  PROPS.setProperty('APP_URL', 'https://importadoramod.github.io/agenda-visitas/');
}

// Opcional: para que el dictado entienda frases libres con IA.
// Pegue su clave de la API de Anthropic y ejecute esta función una vez.
function guardarClaveIA() {
  PROPS.setProperty('ANTHROPIC_KEY', 'PEGUE-AQUI-SU-CLAVE');
}

/* ===================== entrada web ===================== */

function doGet(e) {
  if (e && e.parameter && e.parameter.payload) return responder_(manejar_(JSON.parse(e.parameter.payload)));
  return ContentService.createTextOutput('Servidor de la Agenda de Visitas funcionando.');
}

function doPost(e) {
  let pedido = {};
  try { pedido = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return responder_({ ok: false, error: 'pedido_invalido' }); }
  return responder_(manejar_(pedido));
}

function responder_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function manejar_(p) {
  if (!p || p.t !== PROPS.getProperty('TOKEN')) return { ok: false, error: 'clave' };
  try {
    switch (p.a) {
      case 'listar': return { ok: true, v: VERSION_SERVIDOR, citas: leerCitas_(), opticas: tLeer_('opticas'), abonos: tLeer_('abonos'), calendario: calElegido_(), cuenta: cuenta_() };
      case 'tlistar': return { ok: true, filas: tLeer_(p.tabla) };
      case 'tguardar': return conCandado_(() => ({ ok: true, fila: tGuardar_(p.tabla, p.fila) }));
      case 'tborrar': return conCandado_(() => { tBorrar_(p.tabla, p.id); return { ok: true }; });
      case 'guardar': return conCandado_(() => ({ ok: true, cita: guardar_(p.cita) }));
      case 'borrar': return conCandado_(() => { borrar_(p.id); return { ok: true }; });
      case 'calendarios': return { ok: true, calendarios: listarCalendarios_(), calendario: calElegido_(), cuenta: cuenta_() };
      case 'elegirCalendario': return elegirCalendario_(p.id || '');
      case 'interpretar': return interpretar_(p.texto || '', p.hoy || '', p.dia || '');
      default: return { ok: false, error: 'accion' };
    }
  } catch (err) {
    console.error(err);
    return { ok: false, error: 'servidor', detalle: String(err && err.message || err) };
  }
}

function conCandado_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ===================== hoja de datos ===================== */

function abrirLibro_(id) { try { return SpreadsheetApp.openById(id); } catch (e) { return null; } }

function hoja_() {
  const libro = abrirLibro_(PROPS.getProperty('SHEET_ID'));
  if (!libro) throw new Error('Falta ejecutar configurar()');
  let sh = libro.getSheetByName(HOJA);
  if (!sh) {
    sh = libro.getSheets()[0];
    sh.setName(HOJA);
  }
  if (sh.getLastRow() === 0 || sh.getLastColumn() < COLS.length) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, sh.getMaxRows(), COLS.length).setNumberFormat('@'); // todo como texto
  return sh;
}

function texto_(v, col) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(v, tz, col === 'hora' ? 'HH:mm' : 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

function leerCitas_() {
  const sh = hoja_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, COLS.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      COLS.forEach((k, i) => { o[k] = texto_(r[i], k); });
      o.monto = Number(o.monto || 0);
      return o;
    });
}

function filaDe_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function limpiar_(c) {
  const s = v => String(v === undefined || v === null ? '' : v).slice(0, 500);
  const tipo = c.tipo === 'cobro' ? 'cobro' : 'venta';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s(c.fecha))) throw new Error('fecha');
  const hora = /^\d{2}:\d{2}$/.test(s(c.hora)) ? s(c.hora) : '';
  return {
    id: s(c.id) || Utilities.getUuid(), fecha: s(c.fecha), hora: hora, cliente: s(c.cliente), ciudad: s(c.ciudad),
    tipo: tipo, mercaderia: tipo === 'venta' ? s(c.mercaderia) : '', monto: tipo === 'cobro' ? Math.round(Number(c.monto || 0) * 100) / 100 : 0,
    nota: s(c.nota), estado: c.estado === 'hecho' ? 'hecho' : 'pendiente', creado: s(c.creado) || new Date().toISOString(),
    actualizado: new Date().toISOString(), calId: '', eventId: '', opticaId: s(c.opticaId)
  };
}

function guardar_(entrada) {
  const sh = hoja_();
  const c = limpiar_(entrada || {});
  const fila = filaDe_(sh, c.id);
  if (fila) {
    const prev = {};
    sh.getRange(fila, 1, 1, COLS.length).getValues()[0].forEach((v, i) => { prev[COLS[i]] = texto_(v, COLS[i]); });
    c.creado = prev.creado || c.creado;
    c.calId = prev.calId; c.eventId = prev.eventId;
  }
  try { sincronizar_(c); } catch (err) { console.error('Calendar', err); }
  const valores = [COLS.map(k => c[k] === undefined ? '' : String(c[k]))];
  sh.getRange(fila || sh.getLastRow() + 1, 1, 1, COLS.length).setValues(valores);
  return c;
}

function borrar_(id) {
  const sh = hoja_(), fila = filaDe_(sh, id);
  if (!fila) return;
  const r = sh.getRange(fila, 1, 1, COLS.length).getValues()[0];
  const calId = texto_(r[COLS.indexOf('calId')]), eventId = texto_(r[COLS.indexOf('eventId')]);
  try { const ev = buscarEvento_(calId, eventId); if (ev) ev.deleteEvent(); } catch (err) { console.error(err); }
  sh.deleteRow(fila);
}

/* ===================== Google Calendar ===================== */

function cuenta_() { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } }

function calElegido_() {
  const id = PROPS.getProperty('CAL_ID') || '';
  const cal = id ? CalendarApp.getCalendarById(id) : null;
  if (cal) return { id: id, nombre: cal.getName() };
  return { id: '', nombre: 'Calendario principal' };
}

function calendario_() {
  const id = PROPS.getProperty('CAL_ID');
  return (id && CalendarApp.getCalendarById(id)) || CalendarApp.getDefaultCalendar();
}

function listarCalendarios_() {
  const principal = CalendarApp.getDefaultCalendar().getId();
  return CalendarApp.getAllCalendars()
    .filter(c => !/#holiday@|#contacts@|#weather@/.test(c.getId()))
    .map(c => ({
      id: c.getId() === principal ? '' : c.getId(),
      nombre: c.getId() === principal ? 'Calendario principal' : c.getName(),
      detalle: c.getId() === principal ? c.getId() : (c.isOwnedByMe() ? 'Calendario de esta cuenta' : 'De otra cuenta: ' + c.getId()),
      otra: !c.isOwnedByMe()
    }));
}

function elegirCalendario_(id) {
  if (!id) { PROPS.deleteProperty('CAL_ID'); return { ok: true, calendario: calElegido_() }; }
  let cal = CalendarApp.getCalendarById(id);
  if (!cal) { try { cal = CalendarApp.subscribeToCalendar(id); } catch (e) { cal = null; } }
  if (!cal) return { ok: false, error: 'sin_permiso' };
  // Comprobar que se puede escribir: crear y borrar un evento de prueba
  try {
    const prueba = cal.createAllDayEvent('Prueba de la Agenda', new Date(2000, 0, 1));
    prueba.deleteEvent();
  } catch (e) { return { ok: false, error: 'solo_lectura' }; }
  PROPS.setProperty('CAL_ID', id);
  return { ok: true, calendario: calElegido_() };
}

function buscarEvento_(calId, eventId) {
  if (!eventId) return null;
  const cal = calId ? CalendarApp.getCalendarById(calId) : CalendarApp.getDefaultCalendar();
  if (!cal) return null;
  try { return cal.getEventById(eventId); } catch (e) { return null; }
}

function dinero_(n) { return '$' + Number(n || 0).toFixed(2); }

function sincronizar_(c) {
  const cal = calendario_();
  const calId = PROPS.getProperty('CAL_ID') || '';
  const tz = Session.getScriptTimeZone();
  const titulo = (c.estado === 'hecho' ? '✓ ' : '') + (c.tipo === 'cobro' ? 'Cobrar ' + dinero_(c.monto) : 'Venta') + ' · ' + c.cliente;
  const detalle = (c.tipo === 'venta' ? 'Llevar: ' + c.mercaderia : 'Cobrar: ' + dinero_(c.monto)) +
    (c.nota ? '\nNota: ' + c.nota : '') + '\n\nRef agenda: ' + c.id;

  let ev = null;
  if (c.eventId) {
    if ((c.calId || '') === calId) ev = buscarEvento_(calId, c.eventId);
    else { const viejo = buscarEvento_(c.calId, c.eventId); if (viejo) { try { viejo.deleteEvent(); } catch (e) {} } }
  }
  // Si cambió entre "con hora" y "todo el día", se crea de nuevo
  if (ev && ev.isAllDayEvent() !== !c.hora) { try { ev.deleteEvent(); } catch (e) {} ev = null; }

  const dia = Utilities.parseDate(c.fecha + ' 00:00', tz, 'yyyy-MM-dd HH:mm');
  if (c.hora) {
    const ini = Utilities.parseDate(c.fecha + ' ' + c.hora, tz, 'yyyy-MM-dd HH:mm');
    const fin = new Date(ini.getTime() + 60 * 60 * 1000);
    if (ev) { ev.setTime(ini, fin); ev.setTitle(titulo); }
    else ev = cal.createEvent(titulo, ini, fin);
  } else {
    if (ev) { ev.setAllDayDate(dia); ev.setTitle(titulo); }
    else ev = cal.createAllDayEvent(titulo, dia);
  }
  ev.setDescription(detalle);
  ev.setLocation(c.ciudad || '');
  alarmas_(ev, c);
  c.calId = calId;
  c.eventId = ev.getId();
}

// Dos alarmas: 24 horas antes y 1 hora antes. Sin hora (todo el día): el día anterior a las 9:00 y a las 19:00.
function alarmas_(ev, c) {
  ev.removeAllReminders();
  if (c.estado === 'hecho') return;
  if (c.hora) { ev.addPopupReminder(24 * 60); ev.addPopupReminder(60); }
  else { ev.addPopupReminder(15 * 60); ev.addPopupReminder(5 * 60); }
}

// Ejecutar UNA vez desde el editor: pone las dos alarmas nuevas a las citas que ya estaban en el calendario.
function actualizarAlarmas() {
  const tz = Session.getScriptTimeZone();
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  let n = 0;
  leerCitas_().forEach(c => {
    if (!c.eventId || c.fecha < hoy || c.estado === 'hecho') return;
    const ev = buscarEvento_(c.calId || PROPS.getProperty('CAL_ID') || '', c.eventId);
    if (ev) { alarmas_(ev, c); n++; }
  });
  Logger.log('Listo: ' + n + ' citas con alarma 24 horas antes y 1 hora antes.');
}

/* ===================== dictado con IA (opcional) ===================== */

function interpretar_(texto, hoy, dia) {
  const clave = PROPS.getProperty('ANTHROPIC_KEY');
  if (!clave || clave.indexOf('PEGUE') === 0) return { ok: false, error: 'sin_ia' };
  const instr = 'Hoy es ' + dia + ' ' + hoy + ' (Ecuador). Una persona dictó una cita de visita de trabajo. ' +
    'Responde SOLO con un objeto JSON con estas claves: fecha (YYYY-MM-DD; convierte "mañana", "el martes", "el 15" a la fecha futura más cercana; si no dice, hoy), ' +
    'hora (HH:MM en 24 horas, o "" si no dice; "las tres" sin más contexto es 15:00), cliente, ciudad, tipo ("venta" o "cobro"), ' +
    'mercaderia (o ""), monto (número en dólares, o 0), nota (o ""). Texto dictado: """' + texto.slice(0, 1000) + '"""';
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': clave, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: 'claude-haiku-5-5', max_tokens: 400, messages: [{ role: 'user', content: instr }] })
  });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'ia', detalle: r.getContentText().slice(0, 300) };
  const txt = (JSON.parse(r.getContentText()).content || []).map(b => b.text || '').join('');
  const m = txt.match(/\{[\s\S]*\}/);
  if (!m) return { ok: false, error: 'ia' };
  return { ok: true, datos: JSON.parse(m[0]) };
}

/* ===================== tablas genéricas (ópticas, abonos…) ===================== */
// Cada tabla es una pestaña de la hoja. La primera fila son los nombres de los campos;
// si la app manda un campo nuevo, se agrega una columna sola. Así no hay que tocar este código.

function tHoja_(tabla) {
  if (!TABLA_OK.test(String(tabla || '')) || tabla === 'citas') throw new Error('tabla');
  const libro = abrirLibro_(PROPS.getProperty('SHEET_ID'));
  if (!libro) throw new Error('Falta ejecutar configurar()');
  const nombre = tabla.charAt(0).toUpperCase() + tabla.slice(1);
  let sh = libro.getSheetByName(nombre);
  if (!sh) {
    sh = libro.insertSheet(nombre);
    sh.getRange(1, 1, 1, 2).setValues([['id', 'actualizado']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function tEncabezado_(sh) {
  const n = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, n).getValues()[0].map(String).filter(Boolean);
}

function tLeer_(tabla) {
  const sh = tHoja_(tabla), n = sh.getLastRow();
  if (n < 2) return [];
  const cab = tEncabezado_(sh);
  return sh.getRange(2, 1, n - 1, cab.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; cab.forEach((k, i) => { o[k] = texto_(r[i], k); }); return o; });
}

function tGuardar_(tabla, fila) {
  if (!fila || typeof fila !== 'object') throw new Error('fila');
  const sh = tHoja_(tabla);
  const o = {};
  Object.keys(fila).slice(0, 60).forEach(k => {
    if (!/^[A-Za-z][A-Za-z0-9_]{0,40}$/.test(k)) return;
    const v = fila[k];
    o[k] = v === null || v === undefined ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v)).slice(0, 2000);
  });
  o.id = o.id || Utilities.getUuid();
  o.actualizado = new Date().toISOString();
  let cab = tEncabezado_(sh);
  const nuevas = Object.keys(o).filter(k => cab.indexOf(k) < 0);
  if (nuevas.length) {
    sh.getRange(1, cab.length + 1, 1, nuevas.length).setValues([nuevas]).setFontWeight('bold');
    cab = cab.concat(nuevas);
  }
  sh.getRange(1, 1, sh.getMaxRows(), cab.length).setNumberFormat('@');
  const r = filaDe_(sh, o.id) || sh.getLastRow() + 1;
  const previa = r <= sh.getLastRow() ? sh.getRange(r, 1, 1, cab.length).getValues()[0] : [];
  sh.getRange(r, 1, 1, cab.length).setValues([cab.map((k, i) => k in o ? o[k] : (previa[i] === undefined ? '' : previa[i]))]);
  return o;
}

function tBorrar_(tabla, id) {
  const sh = tHoja_(tabla), r = filaDe_(sh, id);
  if (r) sh.deleteRow(r);
}
