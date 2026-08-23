// ============================================================
// 25. WEB
// Aplicación web del ERP.
// Gestiona el acceso mediante enlace, rutas y vistas HTML.
// ============================================================


  // ============================================================
  // 01. CONFIGURACIÓN DEL MÓDULO WEB
  // ============================================================

  const WEB_CONFIG = {

    // ----------------------------------------------------------
    // VISTAS HTML
    // ----------------------------------------------------------

    LOGIN: "F2_LOGIN",

    DASHBOARD: "F3_DASHBOARD",


    // ----------------------------------------------------------
    // RUTAS PRINCIPALES
    // ----------------------------------------------------------

    RUTA_LOGIN: "login",

    RUTA_DASHBOARD: "dashboard",


    // ----------------------------------------------------------
    // CONFIGURACIÓN DE INTERFAZ
    // ----------------------------------------------------------

    TITULO_ERP: "ERP",

    ANCHO_LOGIN: 1200,

    ALTO_LOGIN: 800

  };


  // ============================================================
  // 02. PUNTO DE ENTRADA DE LA APLICACIÓN WEB
  // Esta función se ejecuta cuando un usuario abre el enlace
  // de la aplicación web.
  // ============================================================

  function doGet(e) {

    // ----------------------------------------------------------
    // OBTENER PARÁMETROS DE LA URL
    // Ejemplo:
    //
    // ?ruta=login
    // ?ruta=dashboard
    // ----------------------------------------------------------

    const parametros =
      e && e.parameter
        ? e.parameter
        : {};


    const ruta =
      String(
        parametros.ruta ||
        WEB_CONFIG.RUTA_LOGIN
      )
        .trim()
        .toLowerCase();


    // ----------------------------------------------------------
    // ENRUTAMIENTO
    // ----------------------------------------------------------

    switch (ruta) {


      // ========================================================
      // LOGIN
      // ========================================================

      case WEB_CONFIG.RUTA_LOGIN:

        return WEB_MOSTRAR_LOGIN();


      // ========================================================
      // DASHBOARD
      // ========================================================

      case WEB_CONFIG.RUTA_DASHBOARD:

        return WEB_MOSTRAR_DASHBOARD(
          parametros
        );


      // ========================================================
      // RUTA NO ENCONTRADA
      // ========================================================

      default:

        return WEB_MOSTRAR_ERROR(
          "La página solicitada no existe."
        );

    }

  }


  // ============================================================
  // 03. MOSTRAR LOGIN
  // Carga la página de inicio de sesión.
  // ============================================================

  function WEB_MOSTRAR_LOGIN() {

    return HtmlService
      .createTemplateFromFile(
        WEB_CONFIG.LOGIN
      )
      .evaluate()
      .setTitle(
        WEB_CONFIG.TITULO_ERP +
        " | Iniciar sesión"
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }


  // ============================================================
  // 04. MOSTRAR DASHBOARD
  // Carga el panel principal del ERP.
  //
  // Por ahora valida únicamente que se haya recibido un token.
  // La validación completa de sesión se realizará utilizando
  // el módulo 23_SEGURIDAD.gs.
  // ============================================================

  function WEB_MOSTRAR_DASHBOARD(parametros) {

    // ----------------------------------------------------------
    // OBTENER TOKEN
    // ----------------------------------------------------------

    const tokenSesion =
      String(
        parametros.token ||
        ""
      ).trim();


    // ----------------------------------------------------------
    // VALIDAR TOKEN
    // ----------------------------------------------------------

    if (!tokenSesion) {

      return WEB_REDIRECCION_LOGIN(
        "Debe iniciar sesión."
      );

    }


    // ----------------------------------------------------------
    // VALIDAR SESIÓN
    // ----------------------------------------------------------

    const validacion =
      SEG_VALIDAR_SESION(
        tokenSesion
      );


    if (!validacion.VALIDA) {

      return WEB_REDIRECCION_LOGIN(
        validacion.MENSAJE
      );

    }


    // ----------------------------------------------------------
    // CARGAR DASHBOARD
    // ----------------------------------------------------------

    const plantilla =
      HtmlService.createTemplateFromFile(
        WEB_CONFIG.DASHBOARD
      );


    // ----------------------------------------------------------
    // ENVIAR DATOS A LA VISTA
    // ----------------------------------------------------------

    plantilla.TOKEN_SESION =
      tokenSesion;

    plantilla.ID_USUARIO =
      validacion.SESION.ID_USUARIO;

    plantilla.USUARIO =
      validacion.SESION.USUARIO;


    return plantilla
      .evaluate()
      .setTitle(
        WEB_CONFIG.TITULO_ERP +
        " | Panel principal"
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  }


  // ============================================================
  // 05. REDIRECCIÓN AL LOGIN
  // Muestra una página indicando que el acceso requiere
  // autenticación.
  //
  // Apps Script no permite una redirección HTTP tradicional
  // desde HtmlOutput, por lo que posteriormente la vista
  // realizará la navegación mediante JavaScript.
  // ============================================================

  function WEB_REDIRECCION_LOGIN(mensaje) {

    const mensajeSeguro =
      String(mensaje || "Debe iniciar sesión.")
        .replace(/"/g, "&quot;");


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


    return HtmlService
      .createHtmlOutput(html)
      .setTitle(
        WEB_CONFIG.TITULO_ERP +
        " | Redirigiendo"
      );

  }


  // ============================================================
  // 06. MOSTRAR ERROR WEB
  // Muestra una página básica cuando ocurre un error
  // de navegación.
  // ============================================================

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


    return HtmlService
      .createHtmlOutput(html)
      .setTitle(
        WEB_CONFIG.TITULO_ERP +
        " | Error"
      );

  }


  // ============================================================
  // 07. ABRIR LOGIN DESDE GOOGLE SHEETS
  // Permite abrir el formulario de login desde el menú
  // interno de Google Sheets.
  // ============================================================

  function WEB_ABRIR_LOGIN() {

    const html =
      HtmlService
        .createHtmlOutputFromFile(
          WEB_CONFIG.LOGIN
        )
        .setWidth(
          WEB_CONFIG.ANCHO_LOGIN
        )
        .setHeight(
          WEB_CONFIG.ALTO_LOGIN
        );


    SpreadsheetApp
      .getUi()
      .showModalDialog(
        html,
        WEB_CONFIG.TITULO_ERP +
        " | Iniciar sesión"
      );

  }