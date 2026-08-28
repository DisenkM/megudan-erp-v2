// ============================================================
// 25. WEB - SISTEMA DE ENRUTAMIENTO Y RENDERIZADO WEB APP (V2)
// ARCHIVO: 25_WEB.gs
// RESPONSABILIDAD: Enrutar de forma segura peticiones HTTP delegadas desde 22_TRIGGERS.gs
// ============================================================

const WEB_CONFIG = {
  LOGIN: "F3_WEB_LOGIN",
  DASHBOARD: "F4_WEB_DASHBOARD",
  CLIENTES_FORM: "F1_CLI_FORM",
  SEGURIDAD_FORM: "F2_USR_GESTION",
  
  RUTA_LOGIN: "login",
  RUTA_DASHBOARD: "dashboard",
  RUTA_CLIENTES: "clientes_form",
  RUTA_SEGURIDAD: "seguridad_form",
  
  TITULO_ERP: "MEGUDAN ERP"
};

/**
 * Procesa la petición GET de la Web App delegada desde 22_TRIGGERS.gs.
 */
function WEB_doGet(e) {
  const parametros = e && e.parameter ? e.parameter : {};
  const ruta = String(parametros.ruta || WEB_CONFIG.RUTA_LOGIN).trim().toLowerCase();
  
  switch (ruta) {
    case WEB_CONFIG.RUTA_LOGIN:
      return WEB_MOSTRAR_LOGIN();
    case WEB_CONFIG.RUTA_DASHBOARD:
      return WEB_MOSTRAR_DASHBOARD(parametros);
    case WEB_CONFIG.RUTA_CLIENTES:
      return WEB_MOSTRAR_CLIENTES_FORM(parametros);
    case WEB_CONFIG.RUTA_SEGURIDAD:
      return WEB_MOSTRAR_SEGURIDAD_FORM(parametros);
    default:
      return WEB_MOSTRAR_ERROR("La página solicitada no existe.");
  }
}

/**
 * Carga la pantalla de inicio de sesión (F3_WEB_LOGIN).
 */
function WEB_MOSTRAR_LOGIN() {
  return HtmlService
    .createTemplateFromFile(WEB_CONFIG.LOGIN)
    .evaluate()
    .setTitle(WEB_CONFIG.TITULO_ERP + " | Iniciar sesión")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Carga el Dashboard Principal (F4_WEB_DASHBOARD) validando previamente el token.
 */
function WEB_MOSTRAR_DASHBOARD(parametros) {
  const tokenSesion = String(parametros.token || "").trim();
  if (!tokenSesion) {
    return WEB_REDIRECCION_LOGIN("Debe iniciar sesión.");
  }
  
  // Validación robusta cruzada de sesión contra 23_SEGURIDAD.gs
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion.VALIDA) {
    return WEB_REDIRECCION_LOGIN(validacion.MENSAJE);
  }
  
  const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.DASHBOARD);
  
  // Inyección segura de variables asíncronas de sesión
  plantilla.TOKEN_SESION = tokenSesion;
  plantilla.ID_USUARIO = validacion.SESION.ID_USUARIO;
  plantilla.USUARIO = validacion.SESION.USUARIO;
  
  return plantilla
    .evaluate()
    .setTitle(WEB_CONFIG.TITULO_ERP + " | Panel principal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Renderiza el formulario de Clientes (F1_CLI_FORM) dentro de la Web App.
 * Valida la existencia de la sesión activa a través de su token.
 */
function WEB_MOSTRAR_CLIENTES_FORM(parametros) {
  const token = String(parametros.token || "").trim();
  const validacion = SEG_VALIDAR_SESION(token);
  if (!validacion.VALIDA) {
    return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
  }
  
  const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.CLIENTES_FORM);
  plantilla.TOKEN_SESION = token;
  plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
  
  return plantilla.evaluate()
    .setTitle("Gestión de Terceros | ERP")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Renderiza la interfaz de Seguridad y Usuarios (F2_USR_GESTION) en la Web App.
 * Valida la sesión y verifica que el rol cuente con el permiso de VER en el módulo SEGURIDAD.
 */
function WEB_MOSTRAR_SEGURIDAD_FORM(parametros) {
  const token = String(parametros.token || "").trim();
  const validacion = SEG_VALIDAR_SESION(token);
  if (!validacion.VALIDA) {
    return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
  }
  
  const acceso = SEG_VALIDAR_ACCESO(token, "SEGURIDAD", "VER");
  if (!acceso.AUTORIZADO) {
    return WEB_MOSTRAR_ERROR("ACCESO DENEGADO: No cuenta con permisos para ver este módulo.");
  }
  
  const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.SEGURIDAD_FORM);
  plantilla.TOKEN_SESION = token;
  plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
  
  return plantilla.evaluate()
    .setTitle("Gestión de Seguridad | ERP")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Genera la pantalla intermedia que fuerza la redirección JavaScript al login.
 */
function WEB_REDIRECCION_LOGIN(mensaje) {
  const mensajeSeguro = String(mensaje || "Debe iniciar sesión.").replace(/"/g, "&quot;");
  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <script>
      window.top.location.href = 
        window.location.origin + 
        window.location.pathname + 
        "?ruta=login&mensaje=${mensajeSeguro}";
    </script>
  </head>
  <body>
    Redirigiendo al inicio de sesión...
  </body>
  </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(WEB_CONFIG.TITULO_ERP + " | Redirigiendo");
}

/**
 * Genera una página de error web uniforme.
 */
function WEB_MOSTRAR_ERROR(mensaje) {
  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <title>Error</title>
  </head>
  <body>
    <h2>Error</h2>
    <p>${mensaje}</p>
  </body>
  </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(WEB_CONFIG.TITULO_ERP + " | Error");
}
