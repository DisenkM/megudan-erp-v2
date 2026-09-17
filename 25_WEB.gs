// (VERSIÓN 24.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 25_WEB.gs (VERSIÓN 24.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Enrutador maestro HTTP GET (WEB_doGet) para la Web App SPA.
* - Servir la compilación asíncrona de sub-vistas del iFrame/Div en memoria.
* - Redirección limpia compatible con el Sandbox de Google Chrome (sin navegación de marco no autorizada).
* - Prevenir errores HTTP 403 mediante normalización de URLs absolutas.
**************************************************************/

const WEB_CONFIG = {
  LOGIN: "F3_WEB_LOGIN",
  DASHBOARD: "F4_WEB_DASHBOARD",
  CLIENTES_FORM: "F1_CLI_FORM",
  COMPRAS_FORM: "F8_COM_VIEW",
  VENTAS_FORM: "F7_VEN_VIEW",
  SEGURIDAD_FORM: "F2_USR_GESTION",
  PRODUCTOS_FORM: "F5_PROD_VIEW",
  INVENTARIO_FORM: "F6_INV_VIEW",
  FINANZAS_FORM: "F9_FIN_VIEW",
  PLANEACION_FORM: "F11_PLA_VIEW",
  OBRAS_FORM: "F10_OBR_VIEW",
  NOMINA_FORM: "F12_NOM_VIEW",
  DOCUMENTOS_FORM: "F19_DOC_VIEW",
  TITULO_ERP: "MEGUDAN ERP"
};

function WEB_doGet(e) {
  try {
    const parametros = e && e.parameter ? e.parameter : {};
    const ruta = String(parametros.ruta || "login").trim().toLowerCase();
    const token = String(parametros.token || "").trim();
    
    if (ruta === "dashboard") {
      return WEB_MOSTRAR_DASHBOARD(parametros);
    }
    
    return WEB_MOSTRAR_LOGIN(parametros);
  } catch (error) {
    return WEB_MOSTRAR_ERROR("Error de procesamiento en enrutador: " + error.toString());
  }
}

function WEB_MOSTRAR_LOGIN(parametros) {
  try {
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.LOGIN);
    let webAppUrl = "";
    try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) { webAppUrl = ""; }
    plantilla.WEB_APP_URL = webAppUrl;
    
    return plantilla.evaluate()
      .setTitle("MEGUDAN ERP | Iniciar sesión")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("<h2>Error de Carga</h2><p>No se pudo cargar la vista de login.</p>");
  }
}

function WEB_MOSTRAR_DASHBOARD(parametros) {
  parametros = parametros || {};
  try {
    const tokenSesion = String(parametros.token || "").trim();
    if (!tokenSesion) return WEB_REDIRECCION_LOGIN("Debe iniciar sesión para acceder al panel.");
    
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || validacion.VALIDA !== true) {
      return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.DASHBOARD);
    plantilla.TOKEN_SESION = tokenSesion;
    plantilla.ID_USUARIO = validacion.SESION ? (validacion.SESION.ID_USUARIO || "") : "";
    plantilla.USUARIO = validacion.SESION ? (validacion.SESION.USUARIO || "") : "";
    
    let webAppUrl = "";
    try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) { webAppUrl = ""; }
    plantilla.WEB_APP_URL = webAppUrl;
    
    return plantilla.evaluate()
      .setTitle("MEGUDAN ERP | Panel Principal")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar panel: " + error.toString());
  }
}

function WEB_REDIRECCION_LOGIN(mensaje) {
  let webAppUrl = "";
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) { webAppUrl = ""; }
  const mensajeSeguro = String(mensaje || "Debe iniciar sesión.");
  const urlDestino = webAppUrl ? (webAppUrl + "?ruta=login&mensaje=" + encodeURIComponent(mensajeSeguro)) : ("?ruta=login&mensaje=" + encodeURIComponent(mensajeSeguro));
  
  const html = '<!DOCTYPE html><html><head><base target="_top"><meta charset="UTF-8"><style>body{margin:0;padding:0;font-family:sans-serif;background:#0f172a;color:#ffffff;display:flex;align-items:center;justify-content:center;min-height:100vh;}.card{background:#1e293b;border:1px solid rgba(255,255,255,0.1);padding:32px;border-radius:16px;text-align:center;max-width:400px;box-shadow:0 20px 40px rgba(0,0,0,0.5);}.btn{background:linear-gradient(135deg,#2563eb 0%,#1d4ed8 100%);color:#ffffff;padding:12px 24px;border-radius:10px;text-decoration:none;font-weight:800;display:inline-block;margin-top:16px;font-size:14px;}</style></head><body><div class="card"><h3 style="margin-top:0;color:#f87171;">⚠️ Control de Sesión</h3><p style="color:#94a3b8;font-size:13px;line-height:1.5;">' + mensajeSeguro + '</p><a href="' + urlDestino + '" target="_top" class="btn">🔑 Volver al Inicio de Sesión</a></div></body></html>';
  
  return HtmlService.createHtmlOutput(html).setTitle("MEGUDAN ERP | Redirigiendo").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function WEB_MOSTRAR_ERROR(mensaje) {
  const html = '<!DOCTYPE html><html><head><base target="_top"><title>Error de Sistema</title></head><body><div style="font-family:sans-serif; padding:40px; text-align:center;"><h2 style="color:#dc2626;">⚠️ Control de Acceso</h2><p>' + mensaje + '</p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle("MEGUDAN ERP | Error");
}

function WEB_OBTENER_COMPILACION_VISTA(ruta, tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!tokenSesion || !validacion || validacion.VALIDA !== true) {
      throw new Error("Sesión inválida o expirada. Por favor, reinicie la página.");
    }
    
    let archivoHtml = "";
    const rutaNormalizada = String(ruta).trim().toLowerCase();
    
    switch (rutaNormalizada) {
      case "ventas": archivoHtml = WEB_CONFIG.VENTAS_FORM; break;
      case "compras": archivoHtml = WEB_CONFIG.COMPRAS_FORM; break;
      case "clientes": archivoHtml = WEB_CONFIG.CLIENTES_FORM; break;
      case "finanzas": archivoHtml = WEB_CONFIG.FINANZAS_FORM; break;
      case "planeacion": archivoHtml = WEB_CONFIG.PLANEACION_FORM; break;
      case "obras": archivoHtml = WEB_CONFIG.OBRAS_FORM; break;
      case "nomina": archivoHtml = WEB_CONFIG.NOMINA_FORM; break;
      case "documentos": archivoHtml = WEB_CONFIG.DOCUMENTOS_FORM; break;
      case "seguridad":
        const acceso = SEG_VALIDAR_ACCESO(tokenSesion, "SEGURIDAD", "VER");
        if (!acceso || acceso.AUTORIZADO !== true) throw new Error("ACCESO DENEGADO: No cuenta con permisos.");
        archivoHtml = WEB_CONFIG.SEGURIDAD_FORM;
        break;
      case "productos": archivoHtml = WEB_CONFIG.PRODUCTOS_FORM; break;
      case "inventario": archivoHtml = WEB_CONFIG.INVENTARIO_FORM; break;
      default: throw new Error("El módulo solicitado '" + ruta + "' no existe.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(archivoHtml);
    plantilla.TOKEN_SESION = tokenSesion;
    plantilla.USUARIO_ACTUAL = validacion.SESION ? validacion.SESION.USUARIO : "ADMINISTRADOR";
    
    let htmlFinal = plantilla.evaluate().getContent();
    const shim = `<script>(function(){ if (typeof google === 'undefined' || !google.script || !google.script.run) { try { const parentWindow = window.parent; if (parentWindow && parentWindow.google && parentWindow.google.script && parentWindow.google.script.run) { window.google = window.google || {}; window.google.script = window.google.script || {}; window.google.script.run = parentWindow.google.script.run; } } catch(e){} } })();</script>`;
    htmlFinal = htmlFinal.replace("<head>", "<head>" + shim);
    return htmlFinal;
  } catch (error) {
    return `<!DOCTYPE html><html><body style="font-family:sans-serif; padding:40px; color:#b91c1c; text-align:center;"><h3>⚠️ Error de Compilación del Módulo</h3><p>${error.message}</p></body></html>`;
  }
}

function OBTENER_VISTA_HTML(nombreVista, tokenSesion) {
  return WEB_OBTENER_COMPILACION_VISTA(nombreVista, tokenSesion);
}

