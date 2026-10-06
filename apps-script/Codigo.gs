const SPREADSHEET_ID = '1xUZ1Q7XDGQCw406Woy1imLppYszBJ9cEjk0__fmP3wU';
const SHEET_TAREAS = 'Tareas';
const SHEET_DESCARTES = 'Descartes_app';
const SHEET_AGENDA = 'Agenda'; // Ajusta este nombre si la pestaña real se llama de otra forma

// Alias aceptados para cada columna de la Agenda (coinciden con los que ya lee index.html).
// El script usa la primera cabecera que encuentre; si ninguna existe, crea una columna nueva
// con el primer nombre de la lista (o el nombre indicado en AGENDA_NEW_COLS para las técnicas).
const AGENDA_COL_ALIASES = {
  fecha: ['Fecha'],
  dia: ['Día', 'Dia'],
  hora: ['Hora'],
  tipo: ['Tipo'],
  organo: ['Órgano / actividad', 'Organo / actividad', 'Órgano', 'Organo'],
  lugar: ['Lugar'],
  detalle: ['Asunto / detalle', 'Asunto', 'Detalle'],
  expediente: ['Expediente relacionado', 'Expediente'],
  agenda_link: ['Agenda / orden del día', 'Agenda / orden del dia', 'Orden del día', 'Orden del dia'],
  expediente_link: ['Ficha expediente', 'Enlace expediente'],
  parlamentario: ['Parlamentario', 'Responsable', 'Asignado a'],
  area_label: ['Etiqueta / área', 'Etiqueta / area', 'Área', 'Area'],
  id_evento: ['ID_evento'],
  creado_por: ['Creado_por'],
  fecha_creacion: ['Fecha_creacion'],
  ultima_modificacion: ['Ultima_modificacion'],
  estado: ['Estado']
};

function doPost(e) {
  // Un único bloqueo para todas las escrituras: si dos personas guardan a la vez,
  // la segunda espera a que termine la primera en lugar de pisarla.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (error) {
    return jsonResponse({ ok: false, error: 'Servicio ocupado, inténtalo de nuevo' });
  }
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (data.action === 'crear_tarea') return crearTarea(data);
    if (data.action === 'actualizar_tarea') return actualizarTarea(data);
    if (data.action === 'eliminar_tarea') return eliminarTarea(data);
    if (data.action === 'descartar_item') return guardarDescarte(data);
    if (data.action === 'restaurar_descarte') return restaurarDescarte(data);
    if (data.action === 'crear_evento') return conSincronizacion(crearEvento(data), 'agenda');
    if (data.action === 'actualizar_evento') return conSincronizacion(actualizarEvento(data), 'agenda');
    if (data.action === 'cancelar_evento') return conSincronizacion(cancelarEvento(data), 'agenda');

    return jsonResponse({ ok: false, error: 'Acción no válida' });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// Tras guardar un evento, la copia de la agenda en Supabase se actualiza al momento
// (ver Sincronizacion.gs) para que la app lo vea sin esperar a la siguiente sincronización.
function conSincronizacion(respuesta, clave) {
  SpreadsheetApp.flush();
  sincronizarClave(clave);
  return respuesta;
}

/* ============ TAREAS ============ */

function crearTarea(data) {
  const sheet = getTareasSheet();
  const tarea = String(data.tarea || '').trim();
  const parlamentario = String(data.parlamentario || '').trim();

  if (!tarea) throw new Error('Falta la tarea');
  if (!parlamentario) throw new Error('Falta el usuario');

  const id = String(data.id || '').trim() || ('T-' + new Date().getTime());

  // Si la app reintenta el mismo envío, se actualiza la fila en lugar de duplicarla.
  const filaExistente = buscarFilaTarea(sheet, id, parlamentario);
  if (filaExistente) {
    return actualizarTarea(Object.assign({}, data, { id: id }));
  }

  sheet.appendRow([
    textoId(id),
    parlamentario,
    tarea,
    String(data.area || '').trim(),
    String(data.prioridad || 'Pronto').trim(),
    String(data.fecha || 'Sin fecha').trim(),
    String(data.estado || 'Pendiente').trim(),
    fechaActual()
  ]);

  return jsonResponse({ ok: true, id: id });
}

function actualizarTarea(data) {
  const sheet = getTareasSheet();
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Falta ID de tarea');

  const parlamentario = String(data.parlamentario || '').trim();
  if (!parlamentario) throw new Error('Falta el usuario');

  const fila = buscarFilaTarea(sheet, id, parlamentario);
  if (!fila) throw new Error('No se ha encontrado la tarea');

  sheet.getRange(fila, 1, 1, 8).setValues([[
    textoId(id),
    parlamentario,
    String(data.tarea || '').trim(),
    String(data.area || '').trim(),
    String(data.prioridad || 'Pronto').trim(),
    String(data.fecha || 'Sin fecha').trim(),
    String(data.estado || 'Pendiente').trim(),
    fechaActual()
  ]]);

  return jsonResponse({ ok: true, id: id });
}

function eliminarTarea(data) {
  const sheet = getTareasSheet();
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Falta ID de tarea');

  const parlamentario = String(data.parlamentario || '').trim();
  if (!parlamentario) throw new Error('Falta el usuario');

  const fila = buscarFilaTarea(sheet, id, parlamentario);
  // Si ya no está (por ejemplo, un reintento), el resultado es el mismo: la tarea no existe.
  if (fila) sheet.deleteRow(fila);

  return jsonResponse({ ok: true, id: id });
}

function listarTareas(usuario) {
  const nombre = String(usuario || '').trim();
  if (!nombre) return [];

  const sheet = getTareasSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, 8);
  const values = range.getValues();
  return range.getDisplayValues()
    .map(function(row, i) {
      return { row: row, id: leerId(values[i][0], row[0]) };
    })
    .filter(function(item) {
      return normalizarTexto(item.row[1]) === normalizarTexto(nombre);
    })
    .map(function(item) {
      const row = item.row;
      return {
        id: item.id,
        title: String(row[2] || '').trim(),
        area: String(row[3] || '').trim(),
        priority: String(row[4] || 'Pronto').trim(),
        when: String(row[5] || 'Sin fecha').trim(),
        done: normalizarTexto(row[6]) === 'hecho'
      };
    })
    .filter(function(task) { return task.id && task.title; });
}

/* ============ DESCARTES ============ */

function guardarDescarte(data) {
  const usuario = String(data.usuario || data.parlamentario || '').trim();
  const tipo = normalizarTipoDescarte(data.tipo);
  const itemId = String(data.item_id || data.id || '').trim();
  if (!usuario) throw new Error('Falta el usuario');
  if (!itemId) throw new Error('Falta el identificador del elemento');

  const sheet = getDescartesSheet();
  if (!buscarFilaDescarte(sheet, usuario, tipo, itemId)) {
    sheet.appendRow([usuario, tipo, itemId, fechaActual()]);
  }

  return jsonResponse({ ok: true, usuario: usuario, tipo: tipo, item_id: itemId });
}

function restaurarDescarte(data) {
  const usuario = String(data.usuario || data.parlamentario || '').trim();
  const tipo = normalizarTipoDescarte(data.tipo);
  const itemId = String(data.item_id || data.id || '').trim();
  if (!usuario) throw new Error('Falta el usuario');
  if (!itemId) throw new Error('Falta el identificador del elemento');

  const sheet = getDescartesSheet();
  const fila = buscarFilaDescarte(sheet, usuario, tipo, itemId);
  if (fila) sheet.deleteRow(fila);

  return jsonResponse({ ok: true, usuario: usuario, tipo: tipo, item_id: itemId });
}

function listarDescartes(usuario) {
  const nombre = String(usuario || '').trim();
  if (!nombre) return [];

  const sheet = getDescartesSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
    .filter(function(row) {
      return normalizarTexto(row[0]) === normalizarTexto(nombre);
    })
    .map(function(row) {
      return { tipo: normalizarTipoDescarte(row[1]), item_id: String(row[2] || '').trim() };
    })
    .filter(function(item) { return item.item_id; });
}

function normalizarTipoDescarte(tipo) {
  const value = String(tipo || '').trim().toLowerCase();
  if (value !== 'radar' && value !== 'bopv' && value !== 'buzon') {
    throw new Error('Tipo de descarte no válido');
  }
  return value;
}

function buscarFilaDescarte(sheet, usuario, tipo, itemId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const targetUser = normalizarTexto(usuario);

  for (let i = 0; i < rows.length; i++) {
    if (
      normalizarTexto(rows[i][0]) === targetUser &&
      String(rows[i][1] || '').trim().toLowerCase() === tipo &&
      String(rows[i][2] || '').trim() === itemId
    ) return i + 2;
  }
  return null;
}

/* ============ Utilidades ============ */

function normalizarTexto(value) {
  return String(value || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
}

// Los ID de tarea de la app son números largos (1790757701798). Sheets los convierte en
// número y, según el formato de la columna, los muestra como "1,79076E+12", con lo que ya
// no coinciden. Se guardan como texto y se leen a partir del valor real de la celda.
function textoId(id) {
  const value = String(id || '').trim();
  return /^\d+$/.test(value) ? "'" + value : value;
}

function leerId(rawValue, displayValue) {
  if (typeof rawValue === 'number' && isFinite(rawValue)) return String(Math.round(rawValue));
  return String(rawValue === '' || rawValue === null ? displayValue || '' : rawValue).trim();
}

function buscarFilaTarea(sheet, id, usuario) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const range = sheet.getRange(2, 1, lastRow - 1, 2);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const targetUser = normalizarTexto(usuario);
  const targetId = String(id).trim();

  for (let i = 0; i < values.length; i++) {
    if (
      leerId(values[i][0], display[i][0]) === targetId &&
      normalizarTexto(display[i][1]) === targetUser
    ) return i + 2;
  }
  return null;
}

function getTareasSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_TAREAS);
  if (!sheet) throw new Error('No existe la pestaña Tareas');
  return sheet;
}

function getDescartesSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_DESCARTES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DESCARTES);
    sheet.getRange(1, 1, 1, 4).setValues([['Usuario', 'Tipo', 'ID elemento', 'Fecha descarte']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function fechaActual() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/* ============ AGENDA: crear / editar / cancelar eventos ============ */

function getAgendaSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_AGENDA);
  if (!sheet) throw new Error('No existe la pestaña ' + SHEET_AGENDA + ' (ajusta SHEET_AGENDA en el script si se llama de otra forma)');
  return sheet;
}

// Devuelve { headers: [...], indexOf: {campo: colIndex0based} }, creando al vuelo las columnas
// técnicas (ID_evento, Creado_por, Fecha_creacion, Ultima_modificacion, Estado) si faltan.
function getAgendaHeaderMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
  const indexOf = {};

  Object.keys(AGENDA_COL_ALIASES).forEach(key => {
    const aliases = AGENDA_COL_ALIASES[key];
    let found = -1;
    for (const alias of aliases) {
      const i = headers.indexOf(alias);
      if (i >= 0) { found = i; break; }
    }
    indexOf[key] = found;
  });

  // Las columnas técnicas se crean automáticamente al final si no existen todavía.
  const tecnicas = ['id_evento', 'creado_por', 'fecha_creacion', 'ultima_modificacion', 'estado'];
  let changed = false;
  tecnicas.forEach(key => {
    if (indexOf[key] < 0) {
      headers.push(AGENDA_COL_ALIASES[key][0]);
      indexOf[key] = headers.length - 1;
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return { headers, indexOf, numCols: headers.length };
}

function setCell(rowValues, indexOf, key, value) {
  if (indexOf[key] >= 0) rowValues[indexOf[key]] = value;
}

function parseFechaSheet(fechaStr) {
  // Espera 'DD/MM/YYYY' (formato que ya usa la columna Fecha existente).
  const m = String(fechaStr || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) throw new Error('Fecha no válida, se esperaba DD/MM/YYYY');
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0);
}

function buscarFilaEvento(sheet, indexOf, id) {
  if (indexOf.id_evento < 0) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const col = indexOf.id_evento + 1;
  const ids = sheet.getRange(2, col, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(id).trim()) return i + 2;
  }
  return null;
}

function crearEvento(data) {
  const titulo = String(data.titulo || '').trim();
  const fecha = String(data.fecha || '').trim();
  const destinatarios = String(data.destinatarios || '').trim();
  const creadoPor = String(data.creado_por || '').trim();

  if (!titulo) throw new Error('Falta el título del evento');
  if (!fecha) throw new Error('Falta la fecha del evento');
  if (!destinatarios) throw new Error('Falta al menos un destinatario');
  if (!creadoPor) throw new Error('Falta quién crea el evento');

  const sheet = getAgendaSheet();
  const { indexOf, numCols } = getAgendaHeaderMap(sheet);

  const id = String(data.id || '').trim() || ('EVT-' + new Date().getTime());

  // Evita duplicados si el cliente reintenta el mismo id.
  if (buscarFilaEvento(sheet, indexOf, id)) {
    return actualizarEvento(Object.assign({}, data, { id: id }));
  }

  const row = new Array(numCols).fill('');
  const fechaDate = parseFechaSheet(fecha);
  setCell(row, indexOf, 'fecha', fechaDate);
  setCell(row, indexOf, 'hora', String(data.hora || '').trim());
  setCell(row, indexOf, 'tipo', String(data.tipo || 'Reunión').trim());
  setCell(row, indexOf, 'organo', titulo);
  setCell(row, indexOf, 'lugar', String(data.lugar || '').trim());
  setCell(row, indexOf, 'detalle', String(data.detalle || '').trim());
  setCell(row, indexOf, 'expediente_link', String(data.enlace || '').trim());
  setCell(row, indexOf, 'parlamentario', destinatarios);
  setCell(row, indexOf, 'id_evento', id);
  setCell(row, indexOf, 'creado_por', creadoPor);
  setCell(row, indexOf, 'fecha_creacion', fechaActual());
  setCell(row, indexOf, 'ultima_modificacion', fechaActual());
  setCell(row, indexOf, 'estado', 'Activo');

  sheet.appendRow(row);
  return jsonResponse({ ok: true, id: id });
}

function actualizarEvento(data) {
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Falta ID de evento');

  const sheet = getAgendaSheet();
  const { indexOf, numCols } = getAgendaHeaderMap(sheet);
  const fila = buscarFilaEvento(sheet, indexOf, id);
  if (!fila) throw new Error('No se ha encontrado el evento (¿fue creado desde la web?)');

  const current = sheet.getRange(fila, 1, 1, numCols).getValues()[0];

  if (data.fecha) current[indexOf.fecha] = parseFechaSheet(data.fecha);
  if (data.hora !== undefined && indexOf.hora >= 0) current[indexOf.hora] = String(data.hora).trim();
  if (data.tipo !== undefined && indexOf.tipo >= 0) current[indexOf.tipo] = String(data.tipo).trim();
  if (data.titulo !== undefined && indexOf.organo >= 0) current[indexOf.organo] = String(data.titulo).trim();
  if (data.lugar !== undefined && indexOf.lugar >= 0) current[indexOf.lugar] = String(data.lugar).trim();
  if (data.detalle !== undefined && indexOf.detalle >= 0) current[indexOf.detalle] = String(data.detalle).trim();
  if (data.enlace !== undefined && indexOf.expediente_link >= 0) current[indexOf.expediente_link] = String(data.enlace).trim();
  if (data.destinatarios !== undefined && indexOf.parlamentario >= 0) current[indexOf.parlamentario] = String(data.destinatarios).trim();
  setCell(current, indexOf, 'ultima_modificacion', fechaActual());

  sheet.getRange(fila, 1, 1, numCols).setValues([current]);
  return jsonResponse({ ok: true, id: id });
}

function cancelarEvento(data) {
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Falta ID de evento');

  const sheet = getAgendaSheet();
  const { indexOf } = getAgendaHeaderMap(sheet);
  const fila = buscarFilaEvento(sheet, indexOf, id);
  // Un reintento de algo ya cancelado o inexistente no es un error.
  if (!fila) return jsonResponse({ ok: true, id: id });

  if (indexOf.estado >= 0) sheet.getRange(fila, indexOf.estado + 1).setValue('Cancelado');
  if (indexOf.ultima_modificacion >= 0) sheet.getRange(fila, indexOf.ultima_modificacion + 1).setValue(fechaActual());

  return jsonResponse({ ok: true, id: id });
}

/* ============ ÍNDICE DEL REPOSITORIO (buscador del Panel general) ============ */

// Carpeta raíz del Repositorio GPV en Drive. La app busca por título en la pestaña
// Repositorio_indice, que se rellena recorriendo esta carpeta y todas sus subcarpetas.
const REPOSITORIO_FOLDER_ID = '1mmgAD6Qx-FnXC97IVSB6D9kOkMikt43j';
const SHEET_REPOSITORIO = 'Repositorio_indice';

// Ejecutar UNA VEZ a mano desde el editor: pide permiso de Drive, crea el índice
// y programa su actualización automática una vez al día (hacia las 6:00).
function instalarIndiceRepositorio() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarIndiceRepositorio')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarIndiceRepositorio').timeBased().everyDays(1).atHour(6).create();
  actualizarIndiceRepositorio();
}

function actualizarIndiceRepositorio() {
  const inicio = Date.now();
  const limiteMs = 5 * 60 * 1000; // Apps Script corta a los 6 minutos
  const tz = Session.getScriptTimeZone();
  const filas = [];
  const pendientes = [{ folder: DriveApp.getFolderById(REPOSITORIO_FOLDER_ID), ruta: '' }];
  const vistas = {};
  let completo = true;

  while (pendientes.length) {
    if (Date.now() - inicio > limiteMs) { completo = false; break; }
    const actual = pendientes.shift();
    const files = actual.folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      filas.push([
        f.getName(),
        f.getUrl(),
        actual.ruta || 'Repositorio',
        Utilities.formatDate(f.getLastUpdated(), tz, 'dd/MM/yyyy'),
        tipoDocumento(f.getMimeType()),
        Utilities.formatDate(f.getDateCreated(), tz, 'dd/MM/yyyy'),
        f.getDateCreated().getTime()
      ]);
    }
    const subs = actual.folder.getFolders();
    while (subs.hasNext()) {
      const s = subs.next();
      const id = s.getId();
      if (vistas[id] || s.isTrashed()) continue; // evita bucles con accesos directos
      vistas[id] = true;
      pendientes.push({ folder: s, ruta: (actual.ruta ? actual.ruta + ' / ' : '') + s.getName() });
    }
  }

  // Si no ha dado tiempo a recorrerlo todo, no se sustituye un índice completo por uno a medias.
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_REPOSITORIO);
  if (!completo && sheet && sheet.getLastRow() > filas.length + 1) return;
  if (!sheet) sheet = ss.insertSheet(SHEET_REPOSITORIO);

  // Los más recientes (por fecha de subida) primero: la app muestra los primeros como
  // "Últimos documentos del repositorio".
  filas.sort((a, b) => b[6] - a[6]);
  const datos = [['Título', 'Enlace', 'Carpeta', 'Modificado', 'Tipo', 'Creado']].concat(filas.map(r => r.slice(0, 6)));
  sheet.clearContents();
  sheet.getRange(1, 1, datos.length, 6).setNumberFormat('@').setValues(datos);
  sheet.setFrozenRows(1);
}

function tipoDocumento(mime) {
  const m = String(mime || '');
  if (m.indexOf('document') >= 0 || m.indexOf('word') >= 0) return 'Documento';
  if (m.indexOf('spreadsheet') >= 0 || m.indexOf('excel') >= 0) return 'Hoja de cálculo';
  if (m.indexOf('presentation') >= 0 || m.indexOf('powerpoint') >= 0) return 'Presentación';
  if (m.indexOf('pdf') >= 0) return 'PDF';
  if (m.indexOf('image') >= 0) return 'Imagen';
  return 'Archivo';
}

/* ============ doGet / helpers de respuesta ============ */

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.action === 'listar_descartes') {
      const result = { ok: true, descartes: listarDescartes(params.usuario) };
      return params.callback ? jsonpResponse(params.callback, result) : jsonResponse(result);
    }
    if (params.action === 'ics') return icsResponse(params.ics);
    if (params.action === 'listar_tareas') {
      const result = { ok: true, tareas: listarTareas(params.usuario) };
      return params.callback ? jsonpResponse(params.callback, result) : jsonResponse(result);
    }
    return jsonResponse({ ok: true, service: 'API tareas, descartes y agenda GPV' });
  } catch (error) {
    const params = (e && e.parameter) || {};
    const result = { ok: false, error: error.message, descartes: [], tareas: [] };
    return params.callback ? jsonpResponse(params.callback, result) : jsonResponse(result);
  }
}

// "Añadir a mi calendario" en iPhone: Safari solo abre el calendario si la cita llega desde una
// dirección web con tipo text/calendar, así que la app manda aquí la cita ya preparada.
function icsResponse(ics) {
  const texto = String(ics || '');
  if (texto.indexOf('BEGIN:VCALENDAR') !== 0 || texto.length > 6000) {
    return ContentService.createTextOutput('Cita no válida').setMimeType(ContentService.MimeType.TEXT);
  }
  return ContentService.createTextOutput(texto).setMimeType(ContentService.MimeType.ICAL);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(callback, obj) {
  const safeCallback = String(callback || '').replace(/[^a-zA-Z0-9_$.]/g, '');
  if (!safeCallback) return jsonResponse(obj);
  return ContentService.createTextOutput(safeCallback + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
