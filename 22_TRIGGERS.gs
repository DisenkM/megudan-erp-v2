/**************************************************************
* 22_TRIGGERS.gs
* RESPONSABILIDAD:
* - Centralizar todos los entrypoints y triggers globales del ERP.
**************************************************************/

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

function onEdit(e) {
  try {
    if (!e) return;
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("onEdit", "TRIGGERS", error);
    } else {
      console.error("Error en onEdit: " + error.toString());
    }
  }
}

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
