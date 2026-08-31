/**************************************************************
* 25_WEB.gs (VERSIÓN 9.0 - ARQUITECTURA HÍBRIDA Standalone & RPC)
* RESPONSABILIDAD:
* - Enrutar de forma segura peticiones HTTP delegadas desde 22_TRIGGERS.gs.
* - Servir la compilación asíncrona de sub-vistas del iFrame en memoria.
* - Inyectar de forma transparente el Polyfill/Shim de comunicación RPC.
**************************************************************/

const WEB_CONFIG = {
  LOGIN: "F3_WEB_LOGIN",
  DASHBOARD: "F4_WEB_DASHBOARD",
  CLIENTES_FORM: "F1_CLI_FORM",
  SEGURIDAD_FORM: "F2_USR_GESTION",
  PRODUCTOS_FORM: "F5_PROD_VIEW",
  INVENTARIO_FORM: "F6_INV_VIEW",
  
  RUTA_LOGIN: "login",
  RUTA_DASHBOARD: "dashboard",
  RUTA_CLIENTES: "clientes_form",
  RUTA_SEGURIDAD: "seguridad_form",
  RUTA_PRODUCTOS: "productos_form",
  RUTA_INVENTARIO: "inventario_form",
  
  TITULO_ERP: "MEGUDAN ERP"
};

/**
 * Enrutador principal de peticiones HTTP GET (Web App Standalone)
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
      case WEB_CONFIG.RUTA_PRODUCTOS:
        return WEB_MOSTRAR_PRODUCTOS_FORM(parametros);
      case WEB_CONFIG.RUTA_INVENTARIO:
        return WEB_MOSTRAR_INVENTARIO_FORM(parametros);
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

function WEB_MOSTRAR_LOGIN() {
  try {
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.LOGIN);
    let webAppUrl = "";
    try {
      webAppUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      webAppUrl = "";
    }
    plantilla.WEB_APP_URL = webAppUrl; // ◄ Inyectamos la URL real de script.google.com para evitar redirecciones a googleusercontent
    
    return plantilla
      .evaluate()
      .setTitle(WEB_CONFIG.TITULO_ERP + " | Iniciar sesión")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_LOGIN", "WEB", error);
    }
    return HtmlService.createHtmlOutput("<h2>Error de Carga</h2><p>No se pudo cargar la vista de login.</p>");
  }
}

function WEB_MOSTRAR_DASHBOARD(parametros) {
  parametros = parametros || {};
  try {
    const tokenSesion = String(parametros.token || "").trim();
    if (!tokenSesion) {
      return WEB_REDIRECCION_LOGIN("Debe iniciar sesión para acceder al panel.");
    }
    
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || validacion.VALIDA !== true) {
      const msg = (validacion && validacion.MENSAJE) ? validacion.MENSAJE : "Sesión inválida o expirada. Por favor inicie sesión.";
      return WEB_REDIRECCION_LOGIN(msg);
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.DASHBOARD);
    plantilla.TOKEN_SESION = tokenSesion;
    plantilla.ID_USUARIO = (validacion.SESION && validacion.SESION.ID_USUARIO) ? validacion.SESION.ID_USUARIO : "";
    plantilla.USUARIO = (validacion.SESION && validacion.SESION.USUARIO) ? validacion.SESION.USUARIO : "";
    
    let webAppUrl = "";
    try {
      webAppUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      webAppUrl = "";
    }
    plantilla.WEB_APP_URL = webAppUrl;
    
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

function WEB_MOSTRAR_CLIENTES_FORM(parametros) {
  parametros = parametros || {};
  try {
    const token = String(parametros.token || "").trim();
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.CLIENTES_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
    
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

function WEB_MOSTRAR_SEGURIDAD_FORM(parametros) {
  parametros = parametros || {};
  try {
    const token = String(parametros.token || "").trim();
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
    }
    
    const acceso = SEG_VALIDAR_ACCESO(token, "SEGURIDAD", "VER");
    if (!acceso || acceso.AUTORIZADO !== true) {
      return WEB_MOSTRAR_ERROR("ACCESO DENEGADO: No cuenta con permisos para ver este módulo.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.SEGURIDAD_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
    
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

function WEB_MOSTRAR_PRODUCTOS_FORM(parametros) {
  parametros = parametros || {};
  try {
    const token = String(parametros.token || "").trim();
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.PRODUCTOS_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
    
    return plantilla.evaluate()
      .setTitle("Catálogo de Productos | ERP")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_PRODUCTOS_FORM", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar catálogo de productos: " + error.toString());
  }
}

function WEB_MOSTRAR_INVENTARIO_FORM(parametros) {
  parametros = parametros || {};
  try {
    const token = String(parametros.token || "").trim();
    const validacion = SEG_VALIDAR_SESION(token);
    if (!validacion || validacion.VALIDA !== true) {
      return WEB_REDIRECCION_LOGIN("Sesión inválida o expirada. Por favor inicie sesión.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(WEB_CONFIG.INVENTARIO_FORM);
    plantilla.TOKEN_SESION = token;
    plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
    
    return plantilla.evaluate()
      .setTitle("Existencias de Inventario | ERP")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_MOSTRAR_INVENTARIO_FORM", "WEB", error);
    }
    return WEB_MOSTRAR_ERROR("Error de sistema al cargar existencias de inventario: " + error.toString());
  }
}

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
      try {
        window.top.location.replace(urlDestino);
      } catch (e) {
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

function WEB_MOSTRAR_ERROR(mensaje) {
  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <title>Error de Sistema</title>
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f9fafb; color: #111827; padding: 40px; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
      .error-card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); max-width: 500px; text-align: center; border: 1px solid #e5e7eb; }
      h2 { color: #dc2626; margin-top: 0; }
      p { font-size: 14px; color: #4b5563; line-height: 1.6; margin-bottom: 25px; }
      .btn { background-color: #1f2937; color: white; padding: 10px 20px; text-decoration: none; border-radius: 6px; font-size: 13px; font-weight: bold; transition: background 0.2s; }
      .btn:hover { background-color: #111827; }
    </style>
  </head>
  <body>
    <div class="error-card">
      <h2>⚠️ Control de Acceso</h2>
      <p>${mensaje}</p>
      <a href="javascript:void(0)" onclick="try{window.top.location.replace(window.location.origin + window.location.pathname + '?ruta=login');}catch(e){window.location.replace(window.location.origin + window.location.pathname + '?ruta=login');}" class="btn">Volver al Login</a>
    </div>
  </body>
  </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(WEB_CONFIG.TITULO_ERP + " | Error");
}

/**
 * Compila las sub-vistas asíncronamente para inyección srcdoc
 */
function WEB_OBTENER_COMPILACION_VISTA(ruta, tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!tokenSesion || !validacion || validacion.VALIDA !== true) {
      throw new Error("Sesión inválida o expirada. Por favor, reinicie la página.");
    }
    
    let archivoHtml = "";
    const rutaNormalizada = String(ruta).trim().toLowerCase();
    
    switch (rutaNormalizada) {
      case "clientes":
        archivoHtml = WEB_CONFIG.CLIENTES_FORM;
        break;
      case "seguridad":
        const acceso = SEG_VALIDAR_ACCESO(tokenSesion, "SEGURIDAD", "VER");
        if (!acceso || acceso.AUTORIZADO !== true) {
          throw new Error("ACCESO DENEGADO: No cuenta con permisos para ver este módulo.");
        }
        archivoHtml = WEB_CONFIG.SEGURIDAD_FORM;
        break;
      case "productos":
        archivoHtml = WEB_CONFIG.PRODUCTOS_FORM;
        break;
      case "inventario":
        archivoHtml = WEB_CONFIG.INVENTARIO_FORM;
        break;
      default:
        throw new Error("El módulo solicitado '" + ruta + "' no existe.");
    }
    
    const plantilla = HtmlService.createTemplateFromFile(archivoHtml);
    plantilla.TOKEN_SESION = tokenSesion;
    plantilla.USUARIO_ACTUAL = validacion.SESION.USUARIO;
    
    let htmlFinal = plantilla.evaluate().getContent();
    const shim = `
    <script>
      (function() {
        if (typeof google === 'undefined' || !google.script || !google.script.run) {
          try {
            const parentWindow = window.parent;
            if (parentWindow && parentWindow.google && parentWindow.google.script && parentWindow.google.script.run) {
              window.google = window.google || {};
              window.google.script = window.google.script || {};
              window.google.script.run = parentWindow.google.script.run;
              console.log("🛡️ [MEGUDAN SHIM] Conexión RPC heredada con éxito.");
            }
          } catch (e) {
            console.error("❌ [MEGUDAN SHIM] Error al heredar RPC: ", e.message);
          }
        }
      })();
    </script>
    `;
    htmlFinal = htmlFinal.replace("<head>", "<head>" + shim);
    return htmlFinal;

  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("WEB_OBTENER_COMPILACION_VISTA", "WEB", error);
    }
    return `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 40px; background: #fff5f5; color: #b91c1c; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin:0; }
          .error-container { border: 1.5px solid #fca5a5; background: #fee2e2; padding: 25px; border-radius: 8px; max-width: 500px; text-align: center; }
        </style>
      </head>
      <body>
        <div class="error-container">
          <h3>⚠️ Error de Compilación del Módulo</h3>
          <p>\${error.message}</p>
        </div>
      </body>
      </html>
    `;
  }
}