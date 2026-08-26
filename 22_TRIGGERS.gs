// ============================================================
// 22. TRIGGERS Y DISPARADORES CENTRALES
// ARCHIVO: 22_TRIGGERS.gs
// RESPONSABILIDAD: Centralizar todos los entrypoints y triggers globales del ERP
// ============================================================

/**
 * Evento global al abrir el libro de Google Sheets.
 * Delega la creación de menús a 01_CORE.gs de forma segura.
 */
function onOpen() {
  try {
    if (typeof CORE_onOpen === "function") {
      CORE_onOpen();
    }
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("onOpen", "TRIGGERS", error);
    } else {
      console.error("Error en onOpen: " + error.toString());
    }
  }
}

/**
 * Evento global de edición en Sheets.
 * Centraliza despachos de eventos para auditorías y validaciones operativas en caliente.
 */
function onEdit(e) {
  try {
    if (!e) return;
    // Central para futuras implementaciones de validación/auditoría reactiva
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("onEdit", "TRIGGERS", error);
    } else {
      console.error("Error en onEdit: " + error.toString());
    }
  }
}

/**
 * Punto de entrada GET de la Web App.
 * Delega de forma segura a 25_WEB.gs para renderizar la interfaz requerida.
 */
function doGet(e) {
  try {
    if (typeof WEB_doGet === "function") {
      return WEB_doGet(e);
    } else {
      return HtmlService.createHtmlOutput("<h2>Error Crítico</h2><p>El enrutador web WEB_doGet no está definido.</p>");
    }
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("doGet", "TRIGGERS", error);
    }
    return HtmlService.createHtmlOutput("<h2>Error de Servidor</h2><p>" + error.toString() + "</p>");
  }
}

/**
 * Punto de entrada POST de la Web App.
 * Delega peticiones asíncronas de servidor de forma segura.
 */
function doPost(e) {
  try {
    if (typeof WEB_doPost === "function") {
      return WEB_doPost(e);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ EXITO: false, MENSAJE: "WEB_doPost no está disponible." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("doPost", "TRIGGERS", error);
    }
    return ContentService.createTextOutput(JSON.stringify({ EXITO: false, ERROR: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
