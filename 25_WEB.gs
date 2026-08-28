// ============================================================
// 25. WEB - SISTEMA DE ENRUTAMIENTO Y RENDERIZADO WEB APP (V3)
// ARCHIVO: 25_WEB.gs
// RESPONSABILIDAD: Enrutar de forma segura peticiones HTTP delegadas desde 22_TRIGGERS.gs
//                  con validación estricta y protección contra fallos silenciosos.
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
 * Maneja las rutas GET del sistema y previene fallas silenciosas.
 */
function WEB_doGet(e) {
  try {
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
        return WEB_MOSTRAR_ERROR("La página solicitada no existe o la ruta ingresada es inválida.");
    }
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_doGet", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de procesamiento crítico en el enrutamiento: " + error.toString());
  }
}

/**
 * Carga la pantalla de inicio de sesión (F3_WEB_LOGIN).
 */
function WEB_MOSTRAR_LOGIN() {
  try {
    return HtmlService
      .createTemplateFromFile(WEB_CONFIG.LOGIN)
      .evaluate()
      .setTitle(WEB_CONFIG.TITULO_ERP + " | Iniciar sesión")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_LOGIN", "WEB", error);
    }
    return HtmlService.createHtmlOutput("<h2>Error de Carga</h2><p>No se pudo cargar la vista de login: " + error.toString() + "</p>");
  }
}

/**
 * Carga el Dashboard Principal (F4_WEB_DASHBOARD) validando previamente el token.
 */
function WEB_MOSTRAR_DASHBOARD(parametros) {
  try {
    const tokenSesion = String(parametros.token || "").trim();
    if (!tokenSesion) {
      return WEB_REDIRECCION_LOGIN("Debe iniciar sesión para acceder al panel.");
    }
    
    // Validación robusta cruzada de sesión contra 23_SEGURIDAD.gs
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || validacion.VALIDA !== true) {
      const msg = (validacion && validacion.MENSAJE) ? validacion.MENSAJE : "Sesión inválida o expirada. Por favor inicie sesión.";
      return WEB_REDIRECCION_LOGIN(msg);
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.DASHBOARD);
    
    // Inyección segura de variables asíncronas de sesión
    plantilla.TOKEN_SESION = tokenSesion;
    plantilla.ID_USUARIO = (validacion.SESION && validacion.SESION.ID_USUARIO) ? validacion.SESION.ID_USUARIO : "";
    plantilla.USUARIO = (validacion.SESION && validacion.SESION.USUARIO) ? validacion.SESION.USUARIO : "";
    
    return plantilla
      .evaluate()
      .setTitle(WEB_CONFIG.TITULO_ERP + " | Panel principal")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_DASHBOARD", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar panel: " + error.toString());
  }
}

/**
 * Renderiza el formulario de Clientes (F1_CLI_FORM) dentro de la Web App.
 * Valida la existencia de la sesión activa a través de su token y verifica los permisos.
 */
function WEB_MOSTRAR_CLIENTES_FORM(parametros) {
  try {
    const token = String(parametros.token || "").trim();
    if (!token) {
      return WEB_REDIRECCION_LOGIN("Debe iniciar sesión para acceder al módulo de clientes.");
    }
    
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      const msg = (validacion && validacion.MENSAJE) ? validacion.MENSAJE : "Sesión inválida o expirada. Por favor inicie sesión.";
      return WEB_REDIRECCION_LOGIN(msg);
    }
    
    // Validar autorización de acceso al módulo CLIENTES para ver
    const acceso = SEG_VALIDAR_ACCESO(token, "CLIENTES", "VER");
    if (!acceso || acceso.AUTORIZADO !== true) {
      const msgError = (acceso && acceso.MENSAJE) ? acceso.MENSAJE : "ACCESO DENEGADO: No cuenta con permisos para ver este módulo.";
      return WEB_MOSTRAR_ERROR(msgError);
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.CLIENTES_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = (validacion.SESION && validacion.SESION.USUARIO) ? validacion.SESION.USUARIO : "";
    
    return plantilla.evaluate()
      .setTitle("Gestión de Terceros | ERP")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_CLIENTES_FORM", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar formulario de terceros: " + error.toString());
  }
}

/**
 * Renderiza la interfaz de Seguridad y Usuarios (F2_USR_GESTION) en la Web App.
 * Valida la sesión y verifica que el rol cuente con el permiso de VER en el módulo SEGURIDAD.
 */
function WEB_MOSTRAR_SEGURIDAD_FORM(parametros) {
  try {
    const token = String(parametros.token || "").trim();
    if (!token) {
      return WEB_REDIRECCION_LOGIN("Debe iniciar sesión para acceder al módulo de seguridad.");
    }
    
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      const msg = (validacion && validacion.MENSAJE) ? validacion.MENSAJE : "Sesión inválida o expirada. Por favor inicie sesión.";
      return WEB_REDIRECCION_LOGIN(msg);
    }
    
    const acceso = SEG_VALIDAR_ACCESO(token, "SEGURIDAD", "VER");
    if (!acceso || acceso.AUTORIZADO !== true) {
      const msgError = (acceso && acceso.MENSAJE) ? acceso.MENSAJE : "ACCESO DENEGADO: No cuenta con permisos para ver este módulo.";
      return WEB_MOSTRAR_ERROR(msgError);
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.SEGURIDAD_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = (validacion.SESION && validacion.SESION.USUARIO) ? validacion.SESION.USUARIO : "";
    
    return plantilla.evaluate()
      .setTitle("Gestión de Seguridad | ERP")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_SEGURIDAD_FORM", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar seguridad: " + error.toString());
  }
}

/**
 * Genera la pantalla intermedia que fuerza la redirección JavaScript al login.
 * Utiliza window.location.replace() y encodeURIComponent() para mayor seguridad y evitar bucles.
 */
function WEB_REDIRECCION_LOGIN(mensaje) {
  const mensajeSeguro = String(mensaje || "Debe iniciar sesión.");
  const parametroMensaje = encodeURIComponent(mensajeSeguro);
  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <script>
      const urlDestino = window.location.origin + window.location.pathname + "?ruta=login&mensaje=" + "${parametroMensaje}";
      if (window.top) {
        window.top.location.replace(urlDestino);
      } else {
        window.location.replace(urlDestino);
      }
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
    <title>Error de Sistema</title>
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f9fafb; color: #111827; padding: 40px; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
      .error-card { background: white; border: 1.5px solid #fca5a5; border-radius: 8px; padding: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); max-width: 500px; text-align: center; }
      h2 { margin: 0 0 15px; color: #dc2626; font-size: 20px; }
      p { margin: 0 0 20px; font-size: 14px; color: #4b5563; line-height: 1.5; }
      .btn { background-color: #1f2937; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; text-decoration: none; display: inline-block; }
      .btn:hover { background-color: #111827; }
    </style>
  </head>
  <body>
    <div class="error-card">
      <h2>⚠️ Control de Acceso</h2>
      <p>${mensaje}</p>
      <a href="javascript:void(0)" onclick="if(window.top){window.top.location.replace(window.location.origin + window.location.pathname + '?ruta=login');}else{window.location.replace(window.location.origin + window.location.pathname + '?ruta=login');}" class="btn">Volver al Login</a>
    </div>
  </body>
  </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(WEB_CONFIG.TITULO_ERP + " | Error");
}
