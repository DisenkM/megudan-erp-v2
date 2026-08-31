/**************************************************************
* 24_LOGS.gs (VERSIÓN ACTUALIZADA - MEGUDAN ERP V2)
* RESPONSABILIDAD:
* - Registro unificado de eventos de sistema, auditoría y control de errores.
* - Sistema de pruebas y diagnóstico en caliente del ERP.
**************************************************************/

/**
 * Registra un error de forma centralizada en el sistema.
 * @param {string} funcion Nombre de la función donde ocurrió el error.
 * @param {string} modulo Módulo del sistema afectado.
 * @param {Error} error Objeto de error capturado.
 */
function LOG_REGISTRAR_ERROR(funcion, modulo, error) {
  const ahora = new Date();
  const mensaje = error.message || error.toString();
  const stack = error.stack || "";
  const usuario = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "SISTEMA") : "SISTEMA";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAuditoria = ss.getSheetByName("USR_AUDITORIA");
    if (!hojaAuditoria) {
      console.error("La hoja USR_AUDITORIA no existe. Error original en [" + funcion + "]: " + mensaje);
      return;
    }
    
    let idAuditoria = "AUD-ERR-" + ahora.getTime();
    if (typeof SEG_GENERAR_ID === "function") {
      try {
        idAuditoria = SEG_GENERAR_ID("USR_AUDITORIA", "ID_AUDITORIA", "AUD");
      } catch (e) {
        // Fallback en caso de que falle la generación de ID
      }
    }
    
    const headers = hojaAuditoria.getRange(1, 1, 1, hojaAuditoria.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
    
    const logObj = {
      ID_AUDITORIA: idAuditoria,
      FECHA_HORA: ahora,
      USUARIO: usuario,
      MODULO: modulo,
      ACCION: "ERROR",
      DESCRIPCION: "ERROR en [" + funcion + "]: " + mensaje + " | Stack: " + stack.substring(0, 300),
      RESULTADO: "ERROR",
      MENSAJE_RESULTADO: mensaje,
      ORIGEN_ACCESO: "SISTEMA",
      FECHA_CREACION: ahora
    };
    
    const rowData = headers.map(h => logObj[h] !== undefined ? logObj[h] : "");
    hojaAuditoria.appendRow(rowData);
  } catch (errDb) {
    console.error("CRÍTICO: No se pudo escribir el log de error en Sheets: " + errDb.toString());
  }
}

/**
 * Registra una acción operativa o administrativa exitosa en la bitácora de auditoría.
 */
function LOG_REGISTRAR_ACCION(modulo, accion, idRegistro, descripcion, resultado, valorAnterior, valorNuevo) {
  try {
    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: modulo,
        ACCION: accion,
        ID_REGISTRO: idRegistro,
        DESCRIPCION: descripcion,
        RESULTADO: resultado || "EXITOSO",
        VALOR_ANTERIOR: valorAnterior || "",
        VALOR_NUEVO: valorNuevo || "",
        ORIGEN_ACCESO: "GOOGLE_SHEETS"
      });
    } else {
      console.log("[AUDIT] Modulo: " + modulo + " | Accion: " + accion + " | Desc: " + descripcion);
    }
  } catch (error) {
    console.error("Error al registrar acción de auditoría: " + error.toString());
  }
}

/**
 * Ejecuta un diagnóstico completo en caliente del ERP.
 * Analiza la integridad de base de datos, encriptación, validadores y permisos.
 * Imprime un panel visual educativo detallado del estado del sistema.
 */
function LOG_EJECUTAR_DIAGNOSTICO_COMPLETO() {
  console.log("==================================================================");
  console.log("🧪 INICIANDO SUITE DE PRUEBAS Y DIAGNÓSTICO DE MEGUDAN ERP V2");
  console.log("==================================================================");
  
  const resultados = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Verificación de Hojas Críticas (Base de Datos)
  const hojasCriticas = ["CFG_SISTEMA", "CFG_EMPRESA", "USR_USUARIOS", "USR_ROLES", "USR_PERMISOS", "USR_SESIONES", "USR_AUDITORIA", "CLI_MAESTRO"];
  console.log("\n1. Verificando Hojas Estructurales en Sheets:");
  hojasCriticas.forEach(function(nombre) {
    const hoja = ss.getSheetByName(nombre);
    if (hoja) {
      console.log("   [PASS] Hoja '" + nombre + "' detectada correctamente.");
      resultados.push({ prueba: "HOJA_" + nombre, estado: "OK", detalle: "Presente y accesible." });
    } else {
      console.error("   [FAIL] Hoja '" + nombre + "' NO encontrada. Falla del instalador inicial.");
      resultados.push({ prueba: "HOJA_" + nombre, estado: "FALLA", detalle: "Ausente de la base de datos." });
    }
  });
  
  // 2. Diagnóstico del Bug Crítico de Campos Obligatorios (ReferenceError de value)
  console.log("\n2. Diagnosticando Validador de Campos Obligatorios (Bug de la variable 'value'):");
  try {
    const datosPrueba = { TEST_CAMPO: "Valor Correcto" };
    if (typeof SEG_VALIDAR_OBLIGATORIOS === "function") {
      SEG_VALIDAR_OBLIGATORIOS(datosPrueba, ["TEST_CAMPO"]);
      try {
        SEG_VALIDAR_OBLIGATORIOS(datosPrueba, ["CAMPO_FALTANTE"]);
        console.error("   [FAIL] El validador no detectó el campo faltante.");
        resultados.push({ prueba: "VALIDADOR_BUG", estado: "FALLA", detalle: "El validador ignoró campos obligatorios vacíos." });
      } catch (valErr) {
        if (valErr.toString().includes("ReferenceError")) {
          console.error("   [BUG DETECTADO] El validador arrojó un ReferenceError. Bug de variable 'value' activo.");
          resultados.push({ prueba: "VALIDADOR_BUG", estado: "CRÍTICO", detalle: "ReferenceError detectado en el validador original." });
        } else {
          console.log("   [PASS] El validador detuvo la ejecución con un mensaje controlado de campo obligatorio.");
          resultados.push({ prueba: "VALIDADOR_BUG", estado: "OK", detalle: "Validador limpio y libre de ReferenceError." });
        }
      }
    } else {
      throw new Error("La función SEG_VALIDAR_OBLIGATORIOS no está disponible.");
    }
  } catch (errObl) {
    console.error("   [FAIL] Error al probar validador: " + errObl.toString());
    resultados.push({ prueba: "VALIDADOR_BUG", estado: "FALLA", detalle: errObl.toString() });
  }
  
  // 3. Prueba de Encriptación SHA-256 (Generador de Hash)
  console.log("\n3. Probando Encriptación SHA-256 (Generador de Hash):");
  try {
    if (typeof SEG_GENERAR_HASH_CONTRASENA === "function") {
      const hashOriginal = SEG_GENERAR_HASH_CONTRASENA("Admin123!");
      const hashEsperado = "3eb3fe66b31e3b4d10fa70b5cad49c7112294af6ae4e476a1c405155d45aa121"; // hash of Admin123!
      if (hashOriginal === hashEsperado) {
        console.log("   [PASS] Generador Hash SHA-256 produce resultados idénticos al estándar.");
        resultados.push({ prueba: "HASH_SHA256", estado: "OK", detalle: "Hash criptográfico seguro verificado." });
      } else {
        console.error("   [FAIL] Mismatch de Hash. Obtenido: " + hashOriginal);
        resultados.push({ prueba: "HASH_SHA256", estado: "FALLA", detalle: "Resultado hash inconsistente." });
      }
    } else {
      throw new Error("La función SEG_GENERAR_HASH_CONTRASENA no está disponible.");
    }
  } catch (errCrypt) {
    console.error("   [FAIL] Fallo al encriptar: " + errCrypt.toString());
    resultados.push({ prueba: "HASH_SHA256", estado: "FALLA", detalle: errCrypt.toString() });
  }
  
  // 4. Prueba del Algoritmo del Dígito de Verificación (DIAN)
  console.log("\n4. Probando cálculo matemático del Dígito de Verificación:");
  try {
    if (typeof CLI_CALCULAR_DV === "function") {
      const dvCalculado = CLI_CALCULAR_DV("901915723"); // NIT real de MEGUDAN
      if (dvCalculado === "2") {
        console.log("   [PASS] NIT 901915723 calculado exitosamente con Dv 2.");
        resultados.push({ prueba: "DIAN_DV", estado: "OK", detalle: "Algoritmo de la DIAN operando al 100% de precisión." });
      } else {
        console.error("   [FAIL] Dv calculado incorrecto: " + dvCalculado + " (Esperado: 2).");
        resultados.push({ prueba: "DIAN_DV", estado: "FALLA", detalle: "Mismatch en el Dv matemático de Colombia." });
      }
    } else {
      throw new Error("La función CLI_CALCULAR_DV no está disponible.");
    }
  } catch (errDv) {
    console.error("   [FAIL] Fallo al calcular Dv: " + errDv.toString());
    resultados.push({ prueba: "DIAN_DV", estado: "FALLA", detalle: errDv.toString() });
  }
  
  // 5. Prueba de Roles y Autocuración (Self-Healing)
  console.log("\n5. Probando Inicialización y Autocuración de Roles y Permisos:");
  try {
    if (typeof SEG_INICIALIZAR_ROLES_PREDEFINIDOS === "function" && typeof SEG_LISTAR_ROLES === "function") {
      SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
      const roles = SEG_LISTAR_ROLES();
      if (roles.length >= 7) {
        console.log("   [PASS] Base de datos de roles autocurada y cargada con " + roles.length + " roles.");
        resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "OK", detalle: "Roles del sistema inicializados." });
      } else {
        console.warn("   [WARN] Roles incompletos detectados.");
        resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "WARN", detalle: "Faltan roles por poblar." });
      }
    } else {
      throw new Error("Funciones de inicialización de roles no disponibles.");
    }
  } catch (errHealing) {
    console.error("   [FAIL] Error en autocuración: " + errHealing.toString());
    resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "FALLA", detalle: errHealing.toString() });
  }
  
  // 6. Prueba de Seguridad Dual (Sheets vs Web App)
  console.log("\n6. Probando Validación de Contexto de Seguridad Dual:");
  try {
    const realGetUi = SpreadsheetApp.getUi;
    
    // Escenario A: Simular Sheets Local Confiable (getUi funciona correctamente)
    SpreadsheetApp.getUi = function() { return { alert: function(){} }; };
    try {
      const accesoLocal = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(undefined, "SEGURIDAD", "CREAR");
      if (accesoLocal && accesoLocal.AUTORIZADO && accesoLocal.CODIGO === "CONTEXTO_SHEETS_TRUSTED") {
        console.log("   [PASS] Escenario Sheets Local: Autorizado automáticamente con rol de Administrador.");
        resultados.push({ prueba: "SEGURIDAD_DUAL_LOCAL", estado: "OK", detalle: "Local Sheets trusted bypass verificado." });
      } else {
        console.error("   [FAIL] Escenario Sheets Local: Falló el bypass confiable.");
        resultados.push({ prueba: "SEGURIDAD_DUAL_LOCAL", estado: "FALLA", detalle: "Fallo en bypass local de Sheets." });
      }
    } catch (e) {
      console.error("   [FAIL] Escenario Sheets Local arrojó error inesperado: " + e.toString());
      resultados.push({ prueba: "SEGURIDAD_DUAL_LOCAL", estado: "FALLA", detalle: e.toString() });
    } finally {
      SpreadsheetApp.getUi = realGetUi; // Restablecer UI original de inmediato
    }
    
    // Escenario B: Simular Contexto Web App (getUi arroja excepción de forma nativa)
    SpreadsheetApp.getUi = function() { throw new Error("No UI"); };
    try {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(null, "SEGURIDAD", "CREAR");
      console.error("   [FAIL] Escenario Web App: El sistema no bloqueó el acceso sin token!");
      resultados.push({ prueba: "SEGURIDAD_DUAL_WEB", estado: "FALLA", detalle: "Web App aceptó llamada sin Token." });
    } catch (webErr) {
      if (webErr.toString().includes("ACCESO DENEGADO")) {
        console.log("   [PASS] Escenario Web App: Bloqueo exitoso. Lanzó error de Token Requerido controlado.");
        resultados.push({ prueba: "SEGURIDAD_DUAL_WEB", estado: "OK", detalle: "Bloqueo correcto de accesos externos anónimos." });
      } else {
        console.error("   [FAIL] Escenario Web App: Error inconsistente: " + webErr.toString());
        resultados.push({ prueba: "SEGURIDAD_DUAL_WEB", estado: "FALLA", detalle: webErr.toString() });
      }
    } finally {
      SpreadsheetApp.getUi = realGetUi; // Restablecer UI original de inmediato
    }
  } catch (errDual) {
    console.error("   [FAIL] Fallo en la suite de seguridad dual: " + errDual.toString());
  }
  
  // Imprimir Resumen del Reporte
  console.log("\n==================================================================");
  console.log("📊 RESUMEN FINAL DEL DIAGNÓSTICO");
  console.log("==================================================================");
  
  let totalPasadas = 0;
  let totalCriticas = 0;
  
  resultados.forEach(function(r) {
    if (r.estado === "OK") {
      totalPasadas++;
      console.log("   ✓ " + r.prueba.padEnd(25) + " | ESTADO: PASÓ  | Detalle: " + r.detalle);
    } else {
      totalCriticas++;
      console.error("   ❌ " + r.prueba.padEnd(25) + " | ESTADO: FALLÓ | Detalle: " + r.detalle);
    }
  });
  
  console.log("==================================================================");
  console.log("Pruebas Ejecutadas: " + resultados.length + " | Pasadas: " + totalPasadas + " | Falladas: " + totalCriticas);
  console.log("==================================================================");
  
  // Escribir en USR_AUDITORIA como registro de test de sistema
  try {
    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: "LOGS",
        SUBMODULO: "DIAGNOSTICO",
        ACCION: "TEST_DIAGNOSTICO",
        TIPO_REGISTRO: "SISTEMA",
        DESCRIPCION: "Diagnóstico completo del ERP ejecutado. Pasó: " + totalPasadas + ", Falló: " + totalCriticas,
        RESULTADO: totalCriticas === 0 ? "EXITOSO" : "ERROR",
        MENSAJE_RESULTADO: "Ejecución de test suite terminada."
      });
    }
  } catch (e) {
    // Silenciar fallas de escritura
  }
  
  return {
    EXITO: totalCriticas === 0,
    PASADAS: totalPasadas,
    FALLADAS: totalCriticas,
    DETALLE: resultados
  };
}
