/**************************************************************
* 18_INTEGRACIONES.gs
* RESPONSABILIDAD:
* - Despachar peticiones HTTP REST API hacia Siigo o Alegra mediante UrlFetchApp.
* - Controlar la cola de envíos y almacenar logs de respuestas API.
**************************************************************/

const INT_CONFIG = {
  ENDPOINT_SIIGO: "https://api.siigo.com/v1",
  ENDPOINT_ALEGRA: "https://api.alegra.com/webapi/v1",
  TIMEOUT_MS: 15000
};

function INT_ENVIAR_PETICION_API(servicio, endpoint, payload) {
  const urlBase = (servicio === "SIIGO") ? INT_CONFIG.ENDPOINT_SIIGO : INT_CONFIG.ENDPOINT_ALEGRA;
  const token = PropertiesService.getScriptProperties().getProperty(servicio + "_API_KEY");

  if (!token) {
    return { ok: false, mensaje: "API Key de integración no configurada." };
  }

  const opciones = {
    method: "POST",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const respuesta = UrlFetchApp.fetch(urlBase + endpoint, opciones);
    const codigo = respuesta.getResponseCode();
    const contenido = respuesta.getContentText();
    return { ok: (codigo >= 200 && codigo < 300), codigo: codigo, respuesta: contenido };
  } catch (error) {
    console.error("Fallo de API: " + error.toString());
    return { ok: false, mensaje: error.message };
  }
}

