# GPV · Notas para continuar el trabajo

App web (PWA) del Grupo Parlamentario Popular Vasco. Publicada con GitHub Pages desde `main`
en https://albaalangarica.github.io/GPV-/ (el enlace y el dominio no deben cambiar).
Hablar con Alba en español, sin tecnicismos.

## Archivos
- `index.html` — la app entera (diseño nuevo, estilo NNGG Euskadi). El script es un IIFE:
  la parte final («ni…») redefine `renderView`/`renderLogin` y gana a las de arriba.
- `anterior/index.html` — diseño antiguo; sigue en /anterior/ pero ya no hay enlace en la app.
- `nueva/index.html` — solo redirige a la raíz (era la versión de prueba).
- `sw.js` — service worker. Si cambias archivos de la app, sube `CACHE_NAME`.
- `logo-ppvasco.png` (original), `logo-ppvasco-blanco.png`, `gpv-icon*.png`, `favicon.png`.
- `apps-script/Codigo.gs` y `apps-script/Sincronizacion.gs` — copias de referencia del
  Apps Script. Alba las pega a mano y vuelve a implementar (Gestionar implementaciones →
  lápiz → Nueva versión). En el repo el token va como `PEGA_AQUI_EL_TOKEN`; nunca subir el real.

## Datos
- Excel «Seguimiento parlamentario PP Vasco»: `1xUZ1Q7XDGQCw406Woy1imLppYszBJ9cEjk0__fmP3wU`.
- Excel → Apps Script (cada 5 min + al editar) → Supabase (proyecto GPV `dnukecumlmfoomqbnkuw`,
  tabla `contenido`) → función `datos` → la app. Si falla, la app lee el CSV de la hoja.
- Escrituras (tareas, descartes, eventos, importantes) van por POST al Apps Script
  (`TASKS_API_URL`) con cola de pendientes que se verifica releyendo.
- No poner claves de Supabase en la página ni en el repo.
- Enlaces: `apps-script/Enlaces.gs` (filasConEnlaces) sube la URL real de las celdas con texto
  enlazado («Abrir ficha») en columnas cuya cabecera es de enlace. Si la celda es solo texto,
  no hay enlace. En Asignadas solo se aceptan fichas de legebiltzarra.eus (nunca Gmail).

## Reglas que ha pedido Alba
- No cambiar ni mover la contraseña del Buzón (`BUZON_PASSWORD`).
- Eventos del buzón: solo entran en agenda si se añaden a mano y solo los ve su dueño.
- Agenda: cada uno ve lo común (Plenos y Comisiones, Grupo, Mesa y Junta) y lo suyo;
  lo personal de otros no aparece; lo propio de «Otros» solo al pulsar. Descartados ocultos.
  La etiqueta del Excel manda sobre la categoría deducida.
- Plenos destacados, con enlace al orden del día siempre visible (cabecera e Inicio).
- Tarjetas de agenda compactas: sin subtítulo, solo título + tipo/hora/lugar.
- Navegación: Inicio · Iniciativas · Asignadas · Panel (sin «Más»). Asignadas = PNL y mociones
  adjudicadas (pestaña «PNL y mociones asignadas» + datos de «Plazos enmiendas» por título). Agenda y Radar se abren desde Inicio;
  Panel reúne carpeta del Drive de cada uno, Buzón, Tareas, Radar y recursos. Cambiar de
  usuario: tocando las iniciales arriba a la derecha.
- Menos clics y menos líneas; filtros siempre a la vista (sin deslizar).
- Iniciativas importantes: hasta 3 por parlamentario, pestaña «Importantes» del Excel.

## Pendiente
- Alba debe actualizar `Codigo.gs` en Apps Script para que se guarden las importantes.
- Más adelante: acceso con nombre y contraseña (en lugar de elegir perfil).

## Cómo trabajar
- Rama de trabajo: partir siempre de `origin/main`, PR y merge con `gh api`
  (`repos/albaalangarica/GPV-/pulls` POST y `/pulls/N/merge` PUT).
- Probar con Playwright (Chromium en `/opt/pw-browsers`) sirviendo la carpeta en local y
  simulando Supabase/Apps Script; comprobar sintaxis del `<script>` antes de publicar.
