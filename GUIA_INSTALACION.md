# Guía de instalación — Dashboard de Facturación GALEGAL S.A.S.

Este paquete contiene:
- `dashboard_facturacion.html` → el dashboard (código completo, HTML+CSS+JS).
- `galegal-logo.png` → el logo, debe ir **en la misma carpeta** que el HTML.

El dashboard funciona 100% en el navegador: cuando alguien carga el Excel, los
datos se procesan localmente y no se envían a ningún servidor.

---

## 1. Sobre la librería de lectura de Excel (SheetJS)

Para leer archivos `.xlsx` en el navegador, el dashboard usa la librería
gratuita **SheetJS**. Por su tamaño (cerca de 1 MB minificado) no es práctico
pegarla dentro del código del HTML a mano, así que el archivo la referencia
con una línea `<script src="...">`.

Tienes dos opciones:

**Opción A — Más simple (usa un CDN público):**
El HTML ya viene configurado así. Funciona de inmediato, pero requiere que el
equipo que abre el dashboard tenga salida a internet (a `cdnjs.cloudflare.com`).
Si tu SharePoint/intranet bloquea dominios externos, algunos usuarios verían
el dashboard sin gráficos.

**Opción B — 100% intranet, sin salir a internet (recomendada para SharePoint):**
1. Descarga una vez el archivo:
   `https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`
2. Súbelo a la **misma biblioteca de documentos de SharePoint** donde subirás
   `dashboard_facturacion.html`, con el nombre `xlsx.full.min.js`.
3. Abre `dashboard_facturacion.html` con un editor de texto y reemplaza la
   línea:
   ```html
   <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
   ```
   por:
   ```html
   <script src="xlsx.full.min.js"></script>
   ```
4. Guarda el archivo. Ahora el dashboard no depende de internet.

---

## 2. Subir el código a GitHub (opcional, si quieres versionarlo o servirlo por GitHub Pages)

1. Crea un repositorio nuevo en GitHub (puede ser privado si el contenido es
   sensible).
2. Sube los archivos:
   ```bash
   git init
   git add dashboard_facturacion.html galegal-logo.png
   git commit -m "Dashboard de facturación GALEGAL"
   git branch -M main
   git remote add origin https://github.com/<tu-usuario>/<tu-repo>.git
   git push -u origin main
   ```
3. Si quieres una URL pública para probarlo (por ejemplo, para incrustarlo
   luego en SharePoint como último recurso si el "Contenido incrustado" de
   SharePoint da problemas), activa **GitHub Pages** en
   `Settings → Pages → Deploy from branch → main`. Tu dashboard quedará en
   `https://<tu-usuario>.github.io/<tu-repo>/dashboard_facturacion.html`.

   ⚠️ Ten en cuenta que si el repositorio es público, cualquier persona con el
   link podría abrir el dashboard (aunque igual necesitaría subir el Excel
   manualmente para ver datos). Para uso interno, la ruta recomendada es subir
   los archivos directamente a SharePoint (paso 3), no usar GitHub Pages.

---

## 3. Incrustar el dashboard en SharePoint (ruta recomendada)

1. Ve a la biblioteca de documentos de tu sitio de SharePoint (o crea una
   nueva, por ejemplo "Dashboards").
2. Sube ahí `dashboard_facturacion.html`, `galegal-logo.png` y (si elegiste la
   Opción B) `xlsx.full.min.js`.
3. Copia el enlace directo al archivo `.html`:
   - Clic derecho sobre el archivo → **Copiar vínculo** → elige
     "Personas con acceso específico" o "Personas de tu organización" según
     tu política de permisos → copia la URL.
4. Ve a la página de SharePoint (o Team Site) donde quieres mostrar el
   dashboard → **Editar página** → agrega una web part de tipo
   **"Incrustar"** (Embed).
5. En el campo de la web part, pega el código:
   ```html
   <iframe src="URL_DEL_ARCHIVO_HTML" width="100%" height="900" frameborder="0"></iframe>
   ```
   Usa la URL que copiaste en el paso 3, pero cambiando el final `?web=1` (si
   lo tiene) por algo que permita visualización directa; si SharePoint insiste
   en descargar el archivo en vez de mostrarlo, usa en su lugar la web part
   **"Visor de archivos"** apuntando directamente al `.html`, o sube el
   archivo a una carpeta de **Site Assets** (que sí permite render directo de
   HTML en algunos tenants).
6. Guarda y publica la página.

> Nota: algunos tenants de SharePoint Online restringen la ejecución de HTML/
> JavaScript personalizado dentro de la biblioteca de documentos por
> seguridad. Si el iframe no carga el contenido, la alternativa es:
> - Publicarlo como página con un **App Part** / **Script Editor** (si tu
>   organización aún lo tiene habilitado), o
> - Hospedar el HTML en GitHub Pages (paso 2) o en un sitio estático interno
>   (Azure Static Web Apps, por ejemplo) y solo incrustar esa URL en
>   SharePoint con la web part "Incrustar".

---

## 4. Cómo se "conecta" con tu Excel de facturación

El dashboard **no lee automáticamente un archivo en OneDrive/SharePoint por
sí mismo** desde el navegador sin backend: por seguridad, los navegadores no
pueden abrir archivos del disco/SharePoint sin que el usuario los seleccione,
salvo que se conecte a la API de Microsoft Graph con permisos e inicio de
sesión (esto requeriría registrar una app en Azure AD, lo cual es un paso de
TI adicional, no un simple HTML).

Con el diseño actual (y siguiendo lo que pediste: "aviso de carga poco
invasivo"), el flujo es:

1. La persona autorizada abre el dashboard en SharePoint.
2. Hace clic en el botón discreto **"⭱ Cargar Excel"** junto al encabezado.
3. Selecciona el archivo `CONTROL_FACTURAS.xlsx` (puede descargarlo primero
   desde la biblioteca de SharePoint donde vive el Excel real).
4. El dashboard procesa el archivo al instante, en el navegador, y todos los
   gráficos y filtros se actualizan.

**Si más adelante quieres automatizar la carga (sin clic manual)**, las
opciones reales son:
- **Automatizar con Microsoft Graph API**: registrar una aplicación en Azure
  AD, dar permisos de lectura sobre el archivo en SharePoint, y que el
  dashboard pida un token (login con la cuenta corporativa) para descargar el
  Excel automáticamente al abrir la página. Esto sí es posible pero es un
  desarrollo adicional (autenticación OAuth) que no es un cambio trivial de
  una línea.
- **Power Automate**: crear un flujo que, cada vez que el Excel se actualice
  en SharePoint, lo convierta a un JSON y lo publique en un archivo estático
  que el dashboard lea por `fetch()` sin necesidad de clic. Es una alternativa
  más simple de mantener que Graph API directo, y con Power Automate no se
  requiere escribir código de autenticación.

Si quieres, puedo ayudarte a preparar cualquiera de esas dos rutas cuando
decidas cuál se ajusta mejor a los permisos de TI de tu organización.

---

## 5. Reglas de negocio ya aplicadas en el dashboard

- **Estado de factura**: se calcula a partir de la columna `N. Factura`.
  - Empieza con `GA` → **Pagada**
  - Empieza con `NC` o `NCV` → **Devuelta**
- **Normalización de "Tipo de Proceso"** (columna `Clase de proceso` en tu
  Excel):
  - `ACTUACIONES` → `ACTUACION`
  - `AJUSTES` → `AJUSTE`
  - `FISCAL` → `FISCALES`
  - `JUDICIAL` → `JUDICIALES`
  - `AUDIENCIA PREJUDICIAL`, `AUDIENCIA EXTRAJUDICIAL`,
    `CONCILIACION PREJUDICIAL`, `CONCILIACION EXTRAJUDICIAL` →
    `AUDIENCIAS PREJUDICIALES`
- **Filtros disponibles**: Año, Mes, Aseguradora y Estado de Factura.
- El dashboard lee las columnas por nombre de forma flexible (ignora espacios
  extra y mayúsculas/minúsculas), así que si cambian ligeramente los
  encabezados del Excel debería seguir funcionando.
