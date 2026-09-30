# GPV · Dashboard parlamentario

Aplicación web (PWA) de una sola página: todo está en `index.html`. Los datos se leen de
hojas de Google (CSV) y los cambios (tareas, eventos, descartes) se guardan a través del
script de Google Apps Script.

## Archivos
- `index.html` — la app completa.
- `sw.js` — service worker: guarda solo los archivos de la app para que abra rápido y sin red.
- `manifest.webmanifest` y `gpv-icon*.png` — instalación en el móvil.
- `supabase.sql` — esquema para una posible migración futura (no se usa ahora).

## Cómo funciona la carga
- Al abrir, se muestra al instante la última copia de los datos guardada en el dispositivo
  y cada hoja se actualiza en segundo plano.
- Si una hoja no responde, se mantiene la copia anterior y aparece un aviso con "Reintentar".
- Cada cambio se comprueba releyendo el Excel; si no aparece, se reenvía y, si sigue
  fallando, se avisa en pantalla.

## Cómo publicar cambios
Edita `index.html` directamente (botón del lápiz en GitHub) en lugar de borrarlo y subir
otro archivo con otro nombre: así se ve qué ha cambiado y no se pierden arreglos anteriores.
Si cambias `sw.js`, sube también el número de `CACHE_NAME`.
