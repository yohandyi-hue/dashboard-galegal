# Dashboard de Facturación — GALEGAL SAS
Guía de instalación y conexión de datos

## Qué contiene esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Página principal del dashboard (ábrela en el navegador) |
| `styles.css` | Diseño visual (colores, tipografía, responsive) |
| `config.js` | **El único archivo que debes editar** para conectar tu Excel |
| `dashboard.js` | Lógica: lectura del Excel, normalización de datos, filtros y gráficos |
| `logo-galegal.png` | Logo de la empresa usado en el encabezado |

No necesitas instalar nada ni tener servidor propio: es un sitio 100% estático (HTML/CSS/JS) que corre en cualquier navegador.

---

## 1. Cómo funciona el cálculo de los datos

Para que sepas qué esperar y puedas validarlo contra tu Excel:

- **Estado de la factura**: se calcula a partir de la columna `N. Factura`.
  - Empieza con `GA` → **Pagada**
  - Empieza con `NC` o `NCV` → **Devuelta**
  - Cualquier otro valor (o vacío) se excluye del conteo.
- **Tipo de proceso** (columna `Clase de proceso`) se normaliza así:
  - `ACTUACIONES` → `ACTUACION`
  - `AJUSTES` → `AJUSTE`
  - `FISCAL` → `FISCALES`
  - `JUDICIAL` → `JUDICIALES`
  - `AUDIENCIA PREJUDICIAL`, `AUDIENCIA EXTRAJUDICIAL`, `CONCILIACION PREJUDICIAL`, `CONCILIACION EXTRAJUDICIAL` → `AUDIENCIAS PREJUDICIALES`
  - Cualquier otro valor se deja tal cual (en mayúsculas).
- **Valor facturado** = suma de la columna `Valor a pagar` de las filas visibles según los filtros activos.
- Filas sin `N. Factura`, sin fecha válida en `Fecha Factura`, o con un prefijo de factura no reconocido, se excluyen automáticamente (para no dañar los totales).

Si en el futuro tu Excel usa nuevos textos en "Clase de proceso" o nuevos prefijos de factura, dímelo y actualizamos el mapa en `dashboard.js` (sección `TIPO_PROCESO_MAP` y función `estadoFromNFactura`).

---

## 2. Publicar el código en GitHub

1. Crea una cuenta en [github.com](https://github.com) si no tienes una.
2. Crea un repositorio nuevo, por ejemplo `galegal-dashboard-facturacion` (puede ser privado).
3. Sube estos 5 archivos (arrastrar y soltar funciona en la interfaz web de GitHub, o usa Git):
   ```bash
   git init
   git add index.html styles.css config.js dashboard.js logo-galegal.png
   git commit -m "Dashboard de facturación GALEGAL SAS"
   git branch -M main
   git remote add origin https://github.com/TU-USUARIO/galegal-dashboard-facturacion.git
   git push -u origin main
   ```
4. (Opcional, para verlo en línea antes de meterlo a SharePoint) activa **GitHub Pages**:
   - Ve a **Settings → Pages**
   - En "Source" elige la rama `main` y carpeta `/ (root)`
   - Guarda. En unos minutos tu dashboard estará en `https://TU-USUARIO.github.io/galegal-dashboard-facturacion/`
   - **Nota:** en GitHub Pages usa el modo `manual` o `direct_url` en `config.js` — el modo `sharepoint_rest` solo funciona cuando el HTML se sirve desde el propio SharePoint.

Guardar el código en GitHub te sirve como respaldo y control de versiones; **no es obligatorio** para que el dashboard funcione en SharePoint (puedes subir los archivos directamente ahí, ver paso 3).

---

## 3. Incrustar el dashboard en SharePoint

Tienes dos caminos según el tipo de sitio de SharePoint:

### Opción A — SharePoint moderno (recomendada)

1. Sube los 5 archivos a una **Biblioteca de documentos** del sitio (por ejemplo, crea una carpeta `Dashboard-Facturacion` dentro de "Documentos compartidos").
2. Abre la página donde quieres mostrar el dashboard → **Editar página**.
3. Agrega una **Web Part de tipo "Insertar código" (Embed)**.
4. Pega el siguiente código, reemplazando la URL por la ruta real de tu `index.html` (clic derecho sobre el archivo en SharePoint → "Copiar vínculo", y usa la URL "descargar" o la ruta directa `.../index.html`):
   ```html
   <iframe src="https://TUDOMINIO.sharepoint.com/sites/TU-SITIO/Documentos%20compartidos/Dashboard-Facturacion/index.html"
           style="width:100%; height:900px; border:none;">
   </iframe>
   ```
5. Publica la página.

> Si SharePoint fuerza la descarga del `.html` en vez de mostrarlo (por seguridad, es el comportamiento por defecto), sube el sitio a través de una **biblioteca de "Páginas de sitio" habilitada para scripts**, o usa la Opción B (SPFx), o aloja el `index.html` en GitHub Pages y solo referencia esa URL en el iframe — así evitas el problema y el resto de instrucciones (auto-actualización) siguen aplicando igual si usas el modo `direct_url`.

### Opción B — Extensión SPFx (para IT / más control)

Si tu organización permite soluciones SPFx (SharePoint Framework), un desarrollador puede empaquetar estos mismos 4 archivos (`index.html`, `styles.css`, `config.js`, `dashboard.js`) como un **Web Part personalizado**. Esto da acceso más limpio a la API REST de SharePoint sin depender de iframes. Si tu equipo de IT quiere este camino, puedo ayudar a adaptar el código a un proyecto SPFx.

---

## 4. Conectar el dashboard para que se actualice solo

Todo se controla desde `config.js`. Edita la variable `DATA_SOURCE_MODE`:

### Modo recomendado: `"sharepoint_rest"` (auto-actualización real, sin exponer el archivo)

Úsalo cuando el dashboard **vive dentro del mismo sitio de SharePoint** (Opción A o B arriba). Aprovecha que el usuario ya inició sesión en SharePoint para leer el Excel con la API REST nativa — no necesitas contraseñas ni tokens.

1. En `config.js`:
   ```js
   DATA_SOURCE_MODE: "sharepoint_rest",
   SHAREPOINT_FILE_SERVER_RELATIVE_URL: "/sites/TU-SITIO/Documentos compartidos/CONTROL_FACTURAS.xlsx",
   ```
2. ¿Cómo obtienes la "ruta servidor-relativa"? Abre el Excel en SharePoint → panel de **Detalles** → copia la ruta que empieza en `/sites/...` (NO la URL completa con `https://`).
3. Cada vez que alguien abra el dashboard, éste leerá la versión más reciente del Excel directamente desde SharePoint. Además, `AUTO_REFRESH_MINUTES` (15 por defecto) hace que la página se refresque sola cada cierto tiempo sin que el usuario tenga que recargar.
4. El botón **"↻ Actualizar"** en la parte superior del dashboard permite forzar una recarga inmediata en cualquier momento.

**Requisito:** el usuario que abre el dashboard debe tener permiso de lectura sobre el archivo Excel en SharePoint (los mismos permisos que ya usa para ver la biblioteca).

### Modo alternativo: `"direct_url"` (si el dashboard NO vive en SharePoint, p. ej. GitHub Pages)

1. Genera un enlace para compartir el Excel ("Cualquiera con el vínculo puede ver"), y ajusta la URL para forzar descarga directa (agregando `?download=1` al final del enlace de OneDrive/SharePoint).
2. En `config.js`:
   ```js
   DATA_SOURCE_MODE: "direct_url",
   DIRECT_FILE_URL: "https://TUDOMINIO.sharepoint.com/.../CONTROL_FACTURAS.xlsx?download=1",
   ```
3. **Importante:** este modo hace el archivo accesible por enlace público a cualquiera que lo tenga — solo recomendado si los datos de facturación no son sensibles o si el enlace se comparte de forma controlada.

### Modo de respaldo: `"manual"` (por defecto en esta entrega)

El dashboard muestra un botón para **cargar el Excel manualmente** desde el computador del usuario. Es la forma más simple de probar el dashboard hoy mismo, sin configurar nada de SharePoint. No se actualiza solo: cada persona debe volver a cargar el archivo cuando quiera ver datos nuevos.

---

## 5. Probar el dashboard ahora mismo (sin SharePoint)

1. Abre `index.html` haciendo doble clic (se abrirá en tu navegador).
2. Como `config.js` viene en modo `"manual"`, verás un cuadro para **"Cargar archivo de facturación"**.
3. Arrastra tu `CONTROL_FACTURAS.xlsx` o haz clic para seleccionarlo.
4. Los KPIs, gráficos y filtros (Aseguradora, Año, Mes, Estado) se llenarán automáticamente.

Cuando quieras pasar a producción en SharePoint con auto-actualización, solo cambia `DATA_SOURCE_MODE` a `"sharepoint_rest"` como se explicó arriba.

---

## 6. Notas técnicas y librerías usadas

- **SheetJS (xlsx.js)** — lee el archivo Excel directamente en el navegador, sin backend.
- **Chart.js** — dibuja los gráficos (barras, dona) de forma responsive.
- Ambas se cargan desde CDN (`cdnjs.cloudflare.com`) en `index.html`. Si la red de tu organización bloquea CDNs externos, descarga ambos archivos `.js` y colócalos en la misma carpeta, reemplazando las etiquetas `<script src="https://cdnjs...">` por `<script src="xlsx.full.min.js">` y `<script src="chart.umd.min.js">` respectivamente.
- El dashboard es completamente responsive: se adapta a computador, tablet y celular (probado con breakpoints en 900px y 640px).

---

¿Dudas o cambios? Puedo ajustar colores, agregar más filtros, cambiar qué columna se suma como "Valor facturado", o adaptar el código a un Web Part SPFx si tu equipo de IT lo prefiere.
