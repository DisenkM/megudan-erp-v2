// (VERSIÓN 35.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 98_PRUEBAS_VISTAS.gs (VERSIÓN 35.0 - V2 ERP - LIBRO 1)
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
  // TEST 1: LOGIN (F3_WEB_LOGIN)
  // ==========================================================
  console.log("\n🔒 [TEST 1] Evaluando compilación de F3_WEB_LOGIN:");
  try {
    const template = HtmlService.createTemplateFromFile("F3_WEB_LOGIN");
    template.WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyonax/exec";
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml || contenidoHtml.length === 0) {
      throw new Error("El motor entregó un HTML vacío para el Login.");
    }
    if (!contenidoHtml.includes("<!DOCTYPE html>")) {
      throw new Error("Falta declaración <!DOCTYPE html> requerida para CSP y Sandbox de Google.");
    }
    if (!contenidoHtml.includes("ejecutarLogin")) {
      throw new Error("El Javascript de login está ausente o truncado.");
    }
    
    console.log("   [PASS] F3_WEB_LOGIN compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PORTAL_LOGIN", estado: "PASS", detalle: "Compilación limpia e inyección HTML5 confirmada en F3_WEB_LOGIN." });
  } catch (errLogin) {
    console.error("   [FAIL] Error en Render de Login: " + errLogin.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PORTAL_LOGIN", estado: "FAIL", detalle: errLogin.message });
  }

  // ==========================================================
  // TEST 2: DASHBOARD (F4_WEB_DASHBOARD)
  // ==========================================================
  console.log("\n🎛️ [TEST 2] Evaluando inyección dinámica en F4_WEB_DASHBOARD:");
  try {
    const template = HtmlService.createTemplateFromFile("F4_WEB_DASHBOARD");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    template.ID_USUARIO = "USR-000001";
    template.USUARIO = "ADMIN_TEST_SUITE";
    template.WEB_APP_URL = "https://script.google.com/macros/s/AKfycbyonax/exec";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("SES-TEST-TOKEN-999999")) {
      throw new Error("Falla de interpolación: El Token de sesión no se inyectó.");
    }
    if (!contenidoHtml.includes("SEG_OBTENER_MENU_NIVEL")) {
      throw new Error("Llamada RPC de menú dinámico no encontrada.");
    }
    
    console.log("   [PASS] F4_WEB_DASHBOARD v18 compilado e inyectado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_DASHBOARD_DINAMICO", estado: "PASS", detalle: "Inyección de variables y menú v18 ok." });
  } catch (errDash) {
    console.error("   [FAIL] Error en Render de Dashboard: " + errDash.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_DASHBOARD_DINAMICO", estado: "FAIL", detalle: errDash.message });
  }

  // ==========================================================
  // TEST 3: SEGURIDAD (F2_USR_GESTION v22)
  // ==========================================================
  console.log("\n🔐 [TEST 3] Evaluando compilación de F2_USR_GESTION v22:");
  try {
    const template = HtmlService.createTemplateFromFile("F2_USR_GESTION");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("USR_CARGAR_USUARIOS")) {
      throw new Error("El módulo Javascript de usuarios no se compiló.");
    }
    if (!contenidoHtml.includes("SEG_GUARDAR_PERMISO_WEB")) {
      throw new Error("Llamada RPC SEG_GUARDAR_PERMISO_WEB no encontrada en permisos dinámicos.");
    }
    
    console.log("   [PASS] F2_USR_GESTION v22 compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_GESTION_SEGURIDAD", estado: "PASS", detalle: "Pestaña de permisos dinámicos v22 ok." });
  } catch (errGestion) {
    console.error("   [FAIL] Error en Render de Seguridad: " + errGestion.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_GESTION_SEGURIDAD", estado: "FAIL", detalle: errGestion.message });
  }

  // ==========================================================
  // TEST 4: TERCEROS (F1_CLI_FORM)
  // ==========================================================
  console.log("\n👥 [TEST 4] Evaluando F1_CLI_FORM:");
  try {
    const template = HtmlService.createTemplateFromFile("F1_CLI_FORM");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("aliasMap") || !contenidoHtml.includes("targetKey")) {
      throw new Error("Traductor de alias no encontrado.");
    }
    
    console.log("   [PASS] F1_CLI_FORM compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FORMULARIO_TERCEROS", estado: "PASS", detalle: "Terceros unificados ok." });
  } catch (errCli) {
    console.error("   [FAIL] Error en Render de Terceros: " + errCli.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FORMULARIO_TERCEROS", estado: "FAIL", detalle: errCli.message });
  }

  // ==========================================================
  // TEST 5: PRODUCTOS (F5_PROD_VIEW)
  // ==========================================================
  console.log("\n📦 [TEST 5] Evaluando F5_PROD_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F5_PROD_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("PROD_LISTAR_PRODUCTOS_WEB")) {
      throw new Error("Llamada RPC PROD_LISTAR_PRODUCTOS_WEB no encontrada.");
    }
    
    console.log("   [PASS] F5_PROD_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_CATALOGO_PRODUCTOS", estado: "PASS", detalle: "Catálogo de productos ok." });
  } catch (errProd) {
    console.error("   [FAIL] Error en Render de Productos: " + errProd.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_CATALOGO_PRODUCTOS", estado: "FAIL", detalle: errProd.message });
  }

  // ==========================================================
  // TEST 6: INVENTARIO (F6_INV_VIEW)
  // ==========================================================
  console.log("\n🗃️ [TEST 6] Evaluando F6_INV_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F6_INV_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("INV_LISTAR_SALDOS_WEB")) {
      throw new Error("Llamada RPC INV_LISTAR_SALDOS_WEB no encontrada.");
    }
    
    console.log("   [PASS] F6_INV_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_SALDOS_INVENTARIO", estado: "PASS", detalle: "Saldos de inventario ok." });
  } catch (errInv) {
    console.error("   [FAIL] Error en Render de Inventarios: " + errInv.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_SALDOS_INVENTARIO", estado: "FAIL", detalle: errInv.message });
  }

  // ==========================================================
  // TEST 7: VENTAS (F7_VEN_VIEW)
  // ==========================================================
  console.log("\n💰 [TEST 7] Evaluando F7_VEN_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F7_VEN_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("VEN_GUARDAR_VENTA_WEB")) {
      throw new Error("Llamada RPC VEN_GUARDAR_VENTA_WEB no encontrada.");
    }
    
    console.log("   [PASS] F7_VEN_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_EMISION_VENTAS", estado: "PASS", detalle: "Ventas AIU ok." });
  } catch (errVen) {
    console.error("   [FAIL] Error en Render de Ventas: " + errVen.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_EMISION_VENTAS", estado: "FAIL", detalle: errVen.message });
  }

  // ==========================================================
  // TEST 8: COMPRAS (F8_COM_VIEW)
  // ==========================================================
  console.log("\n🛒 [TEST 8] Evaluando F8_COM_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F8_COM_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("COM_GUARDAR_COMPRA_WEB")) {
      throw new Error("Llamada RPC COM_GUARDAR_COMPRA_WEB no encontrada.");
    }
    
    console.log("   [PASS] F8_COM_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_EMISION_COMPRAS", estado: "PASS", detalle: "Compras y CxP ok." });
  } catch (errCom) {
    console.error("   [FAIL] Error en Render de Compras: " + errCom.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_EMISION_COMPRAS", estado: "FAIL", detalle: errCom.message });
  }

  // ==========================================================
  // TEST 9: FINANZAS (F9_FIN_VIEW v4)
  // ==========================================================
  console.log("\n📊 [TEST 9] Evaluando F9_FIN_VIEW v4:");
  try {
    const template = HtmlService.createTemplateFromFile("F9_FIN_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("FIN_CALCULAR_ESTADOS_FINANCIEROS_WEB")) {
      throw new Error("Llamada RPC FIN_CALCULAR_ESTADOS_FINANCIEROS_WEB no encontrada.");
    }
    if (!contenidoHtml.includes("TAX_CALCULAR_RESUMEN_TRIBUTARIO_WEB")) {
      throw new Error("Llamada RPC TAX_CALCULAR_RESUMEN_TRIBUTARIO_WEB no encontrada.");
    }
    
    console.log("   [PASS] F9_FIN_VIEW v4 compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FINANZAS", estado: "PASS", detalle: "Finanzas e Impuestos v4 ok." });
  } catch (errFin) {
    console.error("   [FAIL] Error en Render de Finanzas: " + errFin.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_FINANZAS", estado: "FAIL", detalle: errFin.message });
  }

  // ==========================================================
  // TEST 10: PLANEACIÓN Y PRESUPUESTOS (F11_PLA_VIEW v4)
  // ==========================================================
  console.log("\n🗓️ [TEST 10] Evaluando F11_PLA_VIEW v4:");
  try {
    const template = HtmlService.createTemplateFromFile("F11_PLA_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("PLA_OBTENER_NOTICIAS_GLOBALES_WEB")) {
      throw new Error("Llamada RPC PLA_OBTENER_NOTICIAS_GLOBALES_WEB no encontrada.");
    }
    if (!contenidoHtml.includes("PLA_GUARDAR_PRESUPUESTO_WEB")) {
      throw new Error("Llamada RPC PLA_GUARDAR_PRESUPUESTO_WEB no encontrada.");
    }
    
    console.log("   [PASS] F11_PLA_VIEW v4 compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PLANEACION", estado: "PASS", detalle: "Planeación v4 con noticias HD y presupuestos ok." });
  } catch (errPla) {
    console.error("   [FAIL] Error en Render de Planeación: " + errPla.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_PLANEACION", estado: "FAIL", detalle: errPla.message });
  }

  // ==========================================================
  // TEST 11: OBRAS Y PROYECTOS (F10_OBR_VIEW)
  // ==========================================================
  console.log("\n🏗️ [TEST 11] Evaluando F10_OBR_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F10_OBR_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("OBR_LISTAR_OBRAS_WEB")) {
      throw new Error("Llamada RPC OBR_LISTAR_OBRAS_WEB no encontrada.");
    }
    
    console.log("   [PASS] F10_OBR_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_OBRAS", estado: "PASS", detalle: "Obras y proyectos ok." });
  } catch (errObr) {
    console.error("   [FAIL] Error en Render de Obras: " + errObr.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_OBRAS", estado: "FAIL", detalle: errObr.message });
  }

  // ==========================================================
  // TEST 12: NÓMINA OPERATIVA (F12_NOM_VIEW)
  // ==========================================================
  console.log("\n👥 [TEST 12] Evaluando F12_NOM_VIEW:");
  try {
    const template = HtmlService.createTemplateFromFile("F12_NOM_VIEW");
    template.TOKEN_SESION = "SES-TEST-TOKEN-999999";
    
    const htmlOutput = template.evaluate();
    const contenidoHtml = htmlOutput.getContent();
    
    if (!contenidoHtml.includes("NOM_LISTAR_EMPLEADOS_WEB")) {
      throw new Error("Llamada RPC NOM_LISTAR_EMPLEADOS_WEB no encontrada.");
    }
    
    console.log("   [PASS] F12_NOM_VIEW compilado correctamente.");
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_NOMINA", estado: "PASS", detalle: "Nómina operativa ok." });
  } catch (errNom) {
    console.error("   [FAIL] Error en Render de Nómina: " + errNom.message);
    resultadosVistas.push({ modulo: "VISTAS", prueba: "RENDER_NOMINA", estado: "FAIL", detalle: errNom.message });
  }

  // ==========================================================
  // TEST 13: ENRUTAMIENTO HTTP GET (doGet)
  // ==========================================================
  console.log("\n🔌 [TEST 13] Probando Enrutamiento y Desvío HTTP (WEB_doGet):");
  try {
    if (typeof WEB_doGet !== "function") {
      throw new Error("La función enrutadora maestra WEB_doGet no está declarada.");
    }
    
    const resPorDefecto = WEB_doGet(undefined);
    if (!resPorDefecto || resPorDefecto.getTitle() !== "MEGUDAN ERP | Iniciar sesión") {
      throw new Error("El enrutador no asignó la ruta por defecto 'login' de forma segura.");
    }
    
    const eFalso = { parameter: { ruta: "dashboard" } };
    const resSinToken = WEB_doGet(eFalso);
    if (!resSinToken || resSinToken.getTitle() !== "MEGUDAN ERP | Redirigiendo") {
      throw new Error("El enrutador permitió cargar la estructura gráfica del panel sin token.");
    }

    console.log("   [PASS] Enrutamiento doGet validado exitosamente.");
    resultadosVistas.push({ modulo: "WEB_ROUTING", prueba: "HTTP_GET_ROUTING", estado: "PASS", detalle: "Enrutador doGet validado ante cargas por defecto y bypass sin token." });
  } catch (errRouting) {
    console.error("   [FAIL] Error en Enrutamiento: " + errRouting.message);
    resultadosVistas.push({ modulo: "WEB_ROUTING", prueba: "HTTP_GET_ROUTING", estado: "FAIL", detalle: errRouting.message });
  }

  // ==========================================================
  // REPORTE CONSOLIDADO FINAL
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
    console.log("🎉 ¡SISTEMA GRÁFICO VERIFICADO! TODAS LAS VISTAS COMPILAN AL 100% EN GOOGLE.");
  } else {
    console.warn("⚠️ SE DETECTARON INCONSISTENCIAS EN LA COMPILACIÓN DE ALGUNAS PLANTILLAS.");
  }
}
