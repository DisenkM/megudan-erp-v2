/**************************************************************
* 18_INTEGRACIONES.gs
* RESPONSABILIDAD:
* - Despachar peticiones HTTP REST API hacia Siigo o Alegra mediante UrlFetchApp.
* - Controlar la cola de envíos y registrar logs de respuestas o excepciones.
**************************************************************/

const INT_CONFIG = {
  ENDPOINT_SIIGO: "https://api.siigo.com/v1",
  ENDPOINT_ALEGRA: "https://api.alegra.com/webapi/v1",
  TIMEOUT_MS: 15000
};

/**
 * Envia un documento o tercero JSON de forma asíncrona hacia Siigo o Alegra.
 * @param {string} endpoint Endpoint destino (ej: "/customers").
 * @param {object} payload Cuerpo de la petición en formato objeto.
 * @returns {object} Respuesta estructurada del servidor externo.
 */
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

    if (codigo >= 200 && codigo < 300) {
      return { ok: true, codigo: codigo, datos: JSON.parse(contenido) };
    } else {
      return { ok: false, codigo: codigo, mensaje: "Error del servidor externo: " + contenido };
    }
  } catch (error) {
    return { ok: false, mensaje: "Fallo de conexión por timeout o error de red: " + error.toString() };
  }
}
