/* ============ SINCRONIZACIÓN EXCEL → SUPABASE ============
 * La app ya no descarga cada pestaña del Excel desde cada móvil: lee una copia en Supabase
 * con una sola consulta. Este script sube esa copia:
 *   - cada 5 minutos (y al momento cuando alguien edita el Excel a mano),
 *   - y justo después de que la app guarde un evento de agenda.
 * Solo se envían las pestañas que han cambiado desde la última vez.
 *
 * Instalación (una vez): pega el token de Supabase en SUPABASE_TOKEN, guarda y ejecuta
 * instalarSincronizacion().
 */

const SUPABASE_URL = 'https://dnukecumlmfoomqbnkuw.supabase.co';
const SUPABASE_KEY = 'PEGA_AQUI_LA_CLAVE_PUBLICA';
const SUPABASE_TOKEN = 'PEGA_AQUI_EL_TOKEN';

const ID_EXCEL_PRINCIPAL = '1xUZ1Q7XDGQCw406Woy1imLppYszBJ9cEjk0__fmP3wU';
const ID_EXCEL_BOPV = '1tmsuoYBO1fTLp7oQ-m5COxuKg9qVlX24JURRllospNE';
const ID_EXCEL_BANCO_DATOS = '1lm2ujBzBe0GSAjKLlmFB-48PATZ5QnYTVLp69Jldxwo';

// clave en Supabase → [Excel, pestaña]. Si la pestaña es null se usa la primera.
const PESTANAS_SINCRONIZADAS = {
  agenda: [ID_EXCEL_PRINCIPAL, 'Agenda'],
  plazos: [ID_EXCEL_PRINCIPAL, 'Plazos enmiendas'],
  iniciativas: [ID_EXCEL_PRINCIPAL, 'Todas iniciativas'],
  asignadas: [ID_EXCEL_PRINCIPAL, 'PNL y mociones asignadas'],
  radar: [ID_EXCEL_PRINCIPAL, 'Radar'],
  accesos: [ID_EXCEL_PRINCIPAL, 'Accesos directos'],
  pendientes: [ID_EXCEL_PRINCIPAL, 'Pendientes'],
  nuevos_documentos: [ID_EXCEL_PRINCIPAL, 'Nuevos documentos'],
  documentos_personales: [ID_EXCEL_PRINCIPAL, 'Documentos personales'],
  santi_cosas: [ID_EXCEL_PRINCIPAL, 'Santi_Cosas'],
  alba_cosas: [ID_EXCEL_PRINCIPAL, 'Alba_Cosas'],
  repositorio: [ID_EXCEL_PRINCIPAL, 'Repositorio_indice'],
  bopv: [ID_EXCEL_BOPV, 'Entradas BOPV'],
  banco_datos: [ID_EXCEL_BANCO_DATOS, null]
};

// Ejecutar UNA VEZ a mano: programa la sincronización y hace la primera subida completa.
function instalarSincronizacion() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['sincronizarTodo', 'sincronizarAlEditar'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarTodo').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('sincronizarAlEditar').forSpreadsheet(ID_EXCEL_PRINCIPAL).onChange().create();
  // Olvida las huellas anteriores para que la primera subida sea completa.
  const props = PropertiesService.getScriptProperties();
  Object.keys(PESTANAS_SINCRONIZADAS).forEach(clave => props.deleteProperty('huella_' + clave));
  sincronizarTodo();
}

function sincronizarAlEditar() {
  sincronizarTodo();
}

function sincronizarTodo() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // ya hay otra sincronización en marcha
  try {
    const abiertos = {};
    const errores = [];
    Object.keys(PESTANAS_SINCRONIZADAS).forEach(clave => {
      try {
        sincronizarClaveInterna(clave, abiertos, false);
      } catch (error) {
        errores.push(clave + ': ' + error.message);
      }
    });
    if (errores.length) console.warn('Sincronización con errores: ' + errores.join(' | '));
  } finally {
    lock.releaseLock();
  }
}

// Para llamar desde doPost justo después de guardar algo (sin bloquear: ya está dentro del lock).
function sincronizarClave(clave) {
  try {
    sincronizarClaveInterna(clave, {}, true);
  } catch (error) {
    console.warn('No se pudo sincronizar ' + clave + ': ' + error.message);
  }
}

function sincronizarClaveInterna(clave, abiertos, forzar) {
  const destino = PESTANAS_SINCRONIZADAS[clave];
  if (!destino) throw new Error('Clave desconocida');
  const libro = abiertos[destino[0]] || (abiertos[destino[0]] = SpreadsheetApp.openById(destino[0]));
  const hoja = destino[1] ? libro.getSheetByName(destino[1]) : libro.getSheets()[0];
  if (!hoja) throw new Error('No existe la pestaña ' + destino[1]);

  const filas = recortarFilas(hoja.getDataRange().getDisplayValues());
  const json = JSON.stringify(filas);
  const huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json, Utilities.Charset.UTF_8));
  const props = PropertiesService.getScriptProperties();
  if (!forzar && props.getProperty('huella_' + clave) === huella) return false;

  const respuesta = UrlFetchApp.fetch(SUPABASE_URL + '/rest/v1/rpc/sincronizar_contenido', {
    method: 'post',
    contentType: 'application/json',
    headers: { apikey: SUPABASE_KEY },
    payload: JSON.stringify({ p_token: SUPABASE_TOKEN, p_clave: clave, p_filas: filas }),
    muteHttpExceptions: true
  });
  if (respuesta.getResponseCode() >= 300) {
    throw new Error('Supabase ' + respuesta.getResponseCode() + ' ' + respuesta.getContentText().slice(0, 200));
  }
  props.setProperty('huella_' + clave, huella);
  return true;
}

// Quita filas y columnas vacías del final para no enviar celdas en blanco.
function recortarFilas(filas) {
  let ultimaFila = filas.length;
  while (ultimaFila > 0 && filas[ultimaFila - 1].every(v => String(v).trim() === '')) ultimaFila--;
  let ultimaCol = 0;
  for (let i = 0; i < ultimaFila; i++) {
    for (let j = filas[i].length; j > ultimaCol; j--) {
      if (String(filas[i][j - 1]).trim() !== '') { ultimaCol = j; break; }
    }
  }
  return filas.slice(0, ultimaFila).map(f => f.slice(0, ultimaCol));
}
