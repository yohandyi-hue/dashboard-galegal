/* ============================================================
   CONFIGURACIÓN DEL DASHBOARD — GALEGAL SAS
   Edita SOLO este archivo para conectar el dashboard a tu Excel.
   No es necesario tocar dashboard.js ni styles.css.
   ============================================================ */

const DASHBOARD_CONFIG = {

  // ----------------------------------------------------------
  // MODO DE CARGA DE DATOS. Elige uno:
  //
  //  "sharepoint_rest" → El dashboard vive DENTRO del mismo sitio
  //      de SharePoint (subido a una Biblioteca de sitio o vía
  //      SPFx / "Insertar código"). Usa la sesión del usuario que
  //      ya inició sesión en SharePoint para leer el Excel con la
  //      API REST nativa. ES LA OPCIÓN RECOMENDADA PARA
  //      ACTUALIZACIÓN AUTOMÁTICA porque no requiere credenciales
  //      adicionales ni exponer el archivo públicamente.
  //
  //  "direct_url" → Descarga el Excel desde una URL directa
  //      (por ejemplo un enlace "para descargar" de OneDrive/SharePoint
  //      compartido públicamente, o cualquier URL accesible por HTTPS).
  //      Más simple, pero el archivo debe ser accesible sin iniciar sesión.
  //
  //  "manual" → El usuario sube el archivo manualmente con el botón
  //      "Cargar archivo". No se actualiza solo, útil para pruebas
  //      locales o mientras configuras el modo automático.
  // ----------------------------------------------------------
  DATA_SOURCE_MODE: "manual",

  // ----------------------------------------------------------
  // Usado solo si DATA_SOURCE_MODE = "sharepoint_rest"
  // Ruta SERVIDOR-RELATIVA del archivo (NO la URL completa).
  // Ejemplo: "/sites/Facturacion/Documentos compartidos/CONTROL_FACTURAS.xlsx"
  // La encuentras abriendo el archivo en SharePoint → Detalles → Ruta.
  // ----------------------------------------------------------
  SHAREPOINT_FILE_SERVER_RELATIVE_URL: "/sites/TU-SITIO/Documentos compartidos/CONTROL_FACTURAS.xlsx",

  // El "web" de SharePoint donde corre la API REST. Normalmente
  // coincide con el origen de la página donde insertes el dashboard.
  // Déjalo vacío ("") para usar automáticamente el origen actual.
  SHAREPOINT_SITE_URL: "",

  // ----------------------------------------------------------
  // Usado solo si DATA_SOURCE_MODE = "direct_url"
  // URL directa y descargable del archivo .xlsx
  // ----------------------------------------------------------
  DIRECT_FILE_URL: "https://tudominio.sharepoint.com/.../CONTROL_FACTURAS.xlsx?download=1",

  // ----------------------------------------------------------
  // Nombre de la hoja de Excel que contiene los datos de facturas.
  // ----------------------------------------------------------
  SHEET_NAME: "Hoja1",

  // ----------------------------------------------------------
  // Minutos entre auto-recargas de datos cuando el modo NO es
  // "manual" (0 = no recargar automáticamente, solo al abrir la página).
  // ----------------------------------------------------------
  AUTO_REFRESH_MINUTES: 15,
};
