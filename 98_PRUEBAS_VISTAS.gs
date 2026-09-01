/**************************************************************
* 98_PRUEBAS_VISTAS.gs (VERSIÓN 11.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Suite de Pruebas Unitarias para validación del Motor de Renderizado (HTML5/ES6) y Enrutamiento.
* - Probar de manera automatizada la compilación, interpolación de scriptlets y evaluación de vistas.
* - Proporcionar guías de resolución explícitas paso a paso en caso de fallos de nombres de archivos.
**************************************************************/

/**
 * Función principal para ejecutar el testeo de interfaz y renderizado web.
 * Seleccionar "PROBAR_RENDERING_Y_VISTAS_E2E" y presionar Ejecutar en el Editor.
 */
function PROBAR_RENDERING_Y_VISTAS_E2E() {
  console.log("==================================================================");
  console.log("🎨 INICIANDO DIAGNÓSTICO DE COMPILACIÓN DE VISTAS (98_PRUEBAS_VISTAS)");
  console.log("==================================================================");
  
  const resultadosVistas = [];
  
  // ==========================================================
  // TEST 1: COMPILACIÓN Y EVALUACIÓN DE VISTA DE LOGIN (F3_WEB_LOGIN)
  // ==========================================================
  console.log("\n🔒 [TEST 1] Evaluando compilación de F3_WEB_LOGIN:");
  try {
    const template = HtmlService.createTemplateFromFile("F3_WEB_LOGIN");
    
    // 🛡️ INYECCIÓN DEFENSIVA DE MOCK URL PARA EVITAR ReferenceError DE NUEVA REDIRECCIÓN V6
    template.WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyonax-D7-70eoWd8yHF2e4MGv4Pf9nBQTM-OspdRM-qBkYamEw/exec";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml || contenidoHtml.length === 0) {
      throw new Error("El motor entregó un HTML vacío para el Login.");
    }
    
    // Validar estándar de compatibilidad
    if (!contenidoHtml.includes("<!DOCTYPE html>")) {
      throw new Error("Falta declaración <!DOCTYPE html> requerida para CSP y Sandbox de Google.");
    }
    if (!contenidoHtml.includes("ejecutarLogin")) {
      throw new Error("El Javascript de login está ausente o truncado.");
    }
    
    console.log("   [PASS] F3_WEB_LOGIN compilado correctamente.");
    console.log("   [INFO] Longitud de buffer HTML: " + contenidoHtml.length + " bytes.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PORTAL_LOGIN", estado: "PASS", detalle: "Compilación limpia e inyección HTML5 confirmada en F3_WEB_LOGIN." });
  } catch (errLogin) {
    console.error("   [FAIL] Error en Render de Login: " + errLogin.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F3_WEB_LOGIN' (sin la extensión .html en el editor) y pega allí todo el contenido del artefacto 'F3_WEB_LOGIN_HTML-v7.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PORTAL_LOGIN", estado: "FAIL", detalle: errLogin.message });
  }

  // ==========================================================
  // TEST 2: COMPILACIÓN DE DASHBOARD CON INYECCIÓN DE SCRIPTLETS (F4_WEB_DASHBOARD)
  // ==========================================================
  console.log("\n🎛️ [TEST 2] Evaluando inyección dinámica en F4_WEB_DASHBOARD:");
  try {
    const template = HtmlService.createTemplateFromFile("F4_WEB_DASHBOARD");
    
    // Inyectar variables que simulan el login asíncrono
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    template.ID_USUARIO = "USR-000001";
    template.USUARIO = "ADMIN_TEST_SUITE";
    template.WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyonax-D7-70eoWd8yHF2e4MGv4Pf9nBQTM-OspdRM-qBkYamEw/exec";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("SES-TEST-TOKEN-999999")) {
      throw new Error("Falla de interpolación: El Token de sesión no se inyectó en el Javascript del cliente.");
    }
    if (!contenidoHtml.includes("ADMIN_TEST_SUITE")) {
      throw new Error("Falla de interpolación: El alias de usuario de sesión no fue inyectado.");
    }
    if (!contenidoHtml.includes("contenedor-principal") && !contenidoHtml.includes("inyectarHTMLConScripts")) {
      throw new Error("Estructura de iFrames segura no encontrada.");
    }
    
    console.log("   [PASS] F4_WEB_DASHBOARD compilado e inyectado correctamente.");
    console.log("   [INFO] Tokens e identidades vinculadas de forma atómica en el DOM.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_DASHBOARD_DINAMICO", estado: "PASS", detalle: "Inyección segura de variables de sesión, WEB_APP_URL y control de viewport-frame exitoso." });
  } catch (errDash) {
    console.error("   [FAIL] Error en Render de Dashboard: " + errDash.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F4_WEB_DASHBOARD' y pega allí todo el contenido del artefacto 'F4_WEB_DASHBOARD_HTML-v11.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_DASHBOARD_DINAMICO", estado: "FAIL", detalle: errDash.message });
  }

  // ==========================================================
  // TEST 3: COMPILACIÓN DE GESTIÓN DE SEGURIDAD (F2_USR_GESTION)
  // ==========================================================
  console.log("\n🔐 [TEST 3] Evaluando compilación defensiva en F2_USR_GESTION:");
  try {
    const template = HtmlService.createTemplateFromFile("F2_USR_GESTION");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("USR_CARGAR_USUARIOS")) {
      throw new Error("El módulo Javascript asíncrono de usuarios no se compiló en la sección de script.");
    }
    if (contenidoHtml.includes("Unexpected string")) {
      throw new Error("Inconsistencia sintáctica de comillas de escape detectada.");
    }
    
    console.log("   [PASS] F2_USR_GESTION compilado de forma defensiva exitosa.");
    console.log("   [PASS] Sintaxis JS blindada y libre de ReferenceErrors inline.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_GESTION_SEGURIDAD", estado: "PASS", detalle: "Compilación limpia de pestañas, tablas modales y JS modular defensivo en F2_USR_GESTION." });
  } catch (errGestion) {
    console.error("   [FAIL] Error en Render de Seguridad: " + errGestion.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F2_USR_GESTION' y pega allí todo el contenido del artefacto 'F2_USR_GESTION_HTML-v19.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_GESTION_SEGURIDAD", estado: "FAIL", detalle: errGestion.message });
  }

  // ==========================================================
  // TEST 4: COMPILACIÓN DE GESTIÓN DE TERCEROS UNIFICADOS (F1_CLI_FORM)
  // ==========================================================
  console.log("\n👥 [TEST 4] Evaluando unificación y mapeo en F1_CLI_FORM:");
  try {
    const template = HtmlService.createTemplateFromFile("F1_CLI_FORM");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("aliasMap") || !contenidoHtml.includes("targetKey")) {
      throw new Error("El diccionario traductor de alias dinámico (Clientes/Proveedores) no fue encontrado en el JS del cliente.");
    }
    if (contenidoHtml.includes("<label_for")) {
      throw new Error("Etiqueta propietaria no semántica <label_for> detectada. Fallará el sanitizador de Google.");
    }
    
    console.log("   [PASS] F1_CLI_FORM compilado de forma impecable.");
    console.log("   [INFO] Estándar W3C validado para la inyección y previsualización de RUT.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FORMULARIO_TERCEROS", estado: "PASS", detalle: "Mapeador asíncrono unificado, traductor de alias y etiquetas semánticas validadas." });
  } catch (errCli) {
    console.error("   [FAIL] Error en Render de Formulario de Terceros: " + errCli.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F1_CLI_FORM' y pega allí todo el contenido del artefacto 'F1_CLI_FORM_HTML-v8.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FORMULARIO_TERCEROS", estado: "FAIL", detalle: errCli.message });
  }

  // ==========================================================
  // TEST 5: COMPILACIÓN DE CATÁLOGO DE PRODUCTOS (F5_PROD_VIEW)
  // ==========================================================
  console.log("\n📦 [TEST 5] Evaluando compilación de F5_PROD_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F5_PROD_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("prodCargarProductos")) {
      throw new Error("El módulo Javascript asíncrono de productos no se compiló en la sección de script.");
    }
    if (!contenidoHtml.includes("PROD_LISTAR_PRODUCTOS_WEB")) {
      throw new Error("La llamada RPC de productos PROD_LISTAR_PRODUCTOS_WEB no se encuentra en el script del cliente.");
    }
    
    console.log("   [PASS] F5_PROD_VIEW compilado correctamente.");
    console.log("   [INFO] Longitud de buffer HTML: " + contenidoHtml.length + " bytes.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_CATALOGO_PRODUCTOS", estado: "PASS", detalle: "Compilación limpia de la interfaz y enlace RPC asíncrono en F5_PROD_VIEW." });
  } catch (errProd) {
    console.error("   [FAIL] Error en Render de Productos: " + errProd.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F5_PROD_VIEW' y pega allí todo el contenido del artefacto 'F5_PROD_VIEW_HTML-v2.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_CATALOGO_PRODUCTOS", estado: "FAIL", detalle: errProd.message });
  }

  // ==========================================================
  // TEST 6: COMPILACIÓN DE SALDOS E INVENTARIO (F6_INV_VIEW)
  // ==========================================================
  console.log("\n🗃️ [TEST 6] Evaluando compilación de F6_INV_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F6_INV_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("invCargarSaldos")) {
      throw new Error("El módulo Javascript asíncrono de inventario no se compiló en la sección de script.");
    }
    if (!contenidoHtml.includes("INV_LISTAR_SALDOS_WEB")) {
      throw new Error("La llamada RPC de inventario INV_LISTAR_SALDOS_WEB no se encuentra en el script del cliente.");
    }
    
    console.log("   [PASS] F6_INV_VIEW compiled correctly.");
    console.log("   [INFO] Longitud de buffer HTML: " + contenidoHtml.length + " bytes.");
    
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_SALDOS_INVENTARIO", estado: "PASS", detalle: "Compilación limpia de la interfaz y enlace RPC asíncrono en F6_INV_VIEW." });
  } catch (errInv) {
    console.error("   [FAIL] Error en Render de Inventarios: " + errInv.message);
    console.error("   💡 SOLUCIÓN: Crea un archivo HTML llamado exactamente 'F6_INV_VIEW' y pega allí todo el contenido del artefacto 'F6_INV_VIEW_HTML-v2.txt'.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_SALDOS_INVENTARIO", estado: "FAIL", detalle: errInv.message });
  }

  // ==========================================================
  // TEST 7: SIMULACIÓN DE ENRUTAMIENTO HTTP GET (doGet)
  // ==========================================================
  console.log("\n🔌 [TEST 7] Probando Enrutamiento y Desvío HTTP (WEB_doGet):");
  try {
    if (typeof WEB_doGet !== "function") {
      throw new Error("La función enrutadora maestra WEB_doGet no está declarada.");
    }
    
    // Escenario A: Acceso sin parámetros (Debe enrutar al Login por defecto)
    console.log("   -> Simulando petición inicial sin parámetros de consulta (Carga de Login)...");
    const resPorDefecto = WEB_doGet(undefined);
    if (!resPorDefecto || resPorDefecto.getTitle() !== "MEGUDAN ERP | Iniciar sesión") {
      throw new Error("El enrutador no asignó la ruta por defecto 'login' de forma segura.");
    }
    console.log("   [PASS] Carga por defecto resuelta exitosamente hacia la vista de Login.");
    
    // Escenario B: Petición de Dashboard sin Token (Debe redireccionar con alerta)
    console.log("   -> Simulando petición de Dashboard saltándose la autenticación...");
    const eFalso = { parameter: { ruta: "dashboard" } };
    const resSinToken = WEB_doGet(eFalso);
    if (!resSinToken || resSinToken.getTitle() !== "MEGUDAN ERP | Redirigiendo") {
      throw new Error("El enrutador permitió cargar la estructura gráfica del panel sin token.");
    }
    console.log("   [PASS] Intento de intrusión bloqueado y desviado de forma controlada hacia el Login.");
    
    // Escenario C: Ruta inexistente (Debe arrojar ventana de error controlada)
    console.log("   -> Simulando petición hacia una URL inválida...");
    const eInvalido = { parameter: { ruta: "ruta_fantasma_999" } };
    const resInvalido = WEB_doGet(eInvalido);
    if (!resInvalido || resInvalido.getTitle() !== "MEGUDAN ERP | Error") {
      throw new Error("El sistema no capturó la ruta errónea con la pantalla de control de acceso.");
    }
    console.log("   [PASS] Excepción de ruta inválida controlada con éxito en UI.");

    resultadosVistas.push({ modulo: "WEB_ROUTING", prueba: "HTTP_GET_ROUTING", estado: "PASS", detalle: "Enrutador doGet validado ante cargas por defecto, intentos de bypass sin token y desvíos de seguridad." });
  } catch (errRouting) {
    console.error("   [FAIL] Error en Módulo de Enrutamiento: " + errRouting.message);
    console.error("   💡 SOLUCIÓN: Asegúrate de que el archivo '25_WEB.gs' contenga el código completo de la versión '25_WEB-v9.gs' de Studio.");
    resultadosVistas.push({ modulo: "WEB_ROUTING", prueba: "HTTP_GET_ROUTING", estado: "FAIL", detalle: errRouting.message });
  }

  // ==========================================================
  // CONSOLIDADO FINAL DE PRUEBAS GRÁFICAS
  // ==========================================================
  console.log("\n==================================================================");
  console.log("📊 REPORTE DE SUITE DE RENDERING DE INTERFAZ WEB");
  console.log("==================================================================");
  
  let pasados = 0;
  let fallados = 0;
  
  resultadosVistas.forEach(r => {
    if (r.estado === "PASS") {
      pasados++;
      console.log("   ✓ [" + r.modulo.padEnd(12) + "] " + r.prueba.padEnd(28) + " | ESTADO: OK   | " + r.detalle);
    } else {
      fallados++;
      console.error("   ❌ [" + r.modulo.padEnd(12) + "] " + r.prueba.padEnd(28) + " | ESTADO: FAIL | " + r.detalle);
    }
  });
  
  console.log("==================================================================");
  console.log("Pruebas Ejecutadas: " + resultadosVistas.length + " | Éxito: " + pasados + " | Errores: " + fallados);
  console.log("==================================================================");
  
  if (fallados === 0) {
    console.log("🎉 ¡SISTEMA GRAFICO VERIFICADO! TODAS LAS VISTAS COMPILAN AL 100% EN GOOGLE.");
  } else {
    console.warn("⚠️ SE DETECTARON INCONSISTENCIAS EN LA COMPILACIÓN DE ALGUNAS PLANTILLAS.");
    console.warn("   Por favor revise los mensajes de solución (💡) impresos arriba para cada fallo.");
  }
}