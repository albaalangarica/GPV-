/* ============ ENLACES ESCONDIDOS EN CELDAS ============
 * En el Excel hay celdas que muestran un texto («Abrir ficha», «Abrir convocatoria»…) con el
 * enlace real por debajo. Al copiar solo el texto visible, la app se quedaba sin el enlace.
 * Esta función devuelve las filas de una pestaña igual que getDisplayValues(), pero en las
 * columnas de enlaces sustituye ese texto por la dirección real.
 */

// Columnas cuyo contenido es un enlace (se reconocen por el nombre de la cabecera).
const CABECERAS_DE_ENLACE = /enlace|link|url|ficha|orden del d[ií]a|convocatoria/i;

function filasConEnlaces(hoja) {
  const rango = hoja.getDataRange();
  const filas = rango.getDisplayValues();
  if (!filas.length) return filas;
  const columnas = [];
  filas[0].forEach(function(cabecera, j) {
    if (CABECERAS_DE_ENLACE.test(String(cabecera || ''))) columnas.push(j);
  });
  if (!columnas.length) return filas;

  const ricos = rango.getRichTextValues();
  for (let i = 1; i < filas.length; i++) {
    columnas.forEach(function(j) {
      const texto = String(filas[i][j] || '').trim();
      if (!texto || /^https?:\/\//i.test(texto)) return;
      const url = enlaceDeCelda(ricos[i] && ricos[i][j]);
      if (url) filas[i][j] = url;
    });
  }
  return filas;
}

function enlaceDeCelda(rico) {
  if (!rico) return '';
  const directo = rico.getLinkUrl();
  if (directo) return directo;
  const trozos = rico.getRuns();
  for (let k = 0; k < trozos.length; k++) {
    const url = trozos[k].getLinkUrl();
    if (url) return url;
  }
  return '';
}
