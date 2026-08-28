/**************************************************************
* 24_LOGS.gs
* RESPONSABILIDAD:
* - Registro unificado de eventos de sistema, auditoría y control de errores.
* - SISTEMA DE PRUEBAS Y DIAGNÓSTICO INTEGRAL (Automático para desarrolladores).
**************************************************************/

/**
 * Registra un error de forma centralizada en el sistema.
 */
function LOG_REGISTRAR_ERROR(funcion, modulo, error) {
  const ahora = new Date();
  const mensaje = error.message || error.toString();
  const stack = error.stack || "No disponible";
  const usuario = Session.getActiveUser().getEmail() || "SISTEMA";

  console.error(
    "=================================\n" +
    "❌ ERROR EN ERP MEGUDAN\n" +
    "=================================\n" +
    "Fecha: " + ahora.toISOString() + "\n" +
    "Módulo: " + modulo + "\n" +
    "Función: " + funcion + "\n" +
    "Mensaje: " + mensaje + "\n" +
    "Usuario: " + usuario + "\n" +
    "Stack Trace:\n" + stack + "\n" +
    "================================="
  );

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAuditoria = ss.getSheetByName("USR_AUDITORIA");
    if (!hojaAuditoria) return;

    let idAuditoria = "AUD-ERR-" + ahora.getTime();
    if (typeof SEG_GENERAR_ID === "function") {
      try {
        idAuditoria = SEG_GENERAR_ID("USR_AUDITORIA", "ID_AUDITORIA", "AUD");
      } catch (e) {}
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
      console.log(`[AUDIT] Modulo: ${modulo} | Accion: ${accion} | Desc: ${descripcion}`);
    }
  } catch (error) {
    console.error("Error al registrar acción de auditoría: " + error.toString());
  }
}

// ============================================================
// 🧪 SUITE DE DIAGNÓSTICO Y UNIT TESTS PARA DESARROLLADORES
// ============================================================

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
    // Llamada con campo presente
    SEG_VALIDAR_OBLIGATORIOS(datosPrueba, ["TEST_CAMPO"]);
    
    // Provocar fallo de validación de forma intencional para verificar si se dispara el ReferenceError original
    try {
      SEG_VALIDAR_OBLIGATORIOS(datosPrueba, ["CAMPO_FALTANTE"]);
      console.error("   [FAIL] El validador no detectó el campo faltante.");
      resultados.push({ prueba: "VALIDADOR_BUG", estado: "FALLA", detalle: "El validador ignoró campos obligatorios vacíos." });
    } catch (valErr) {
      if (valErr.toString().includes("ReferenceError")) {
        console.error("   [BUG DETECTADO] El validador arrojó un ReferenceError. Bug de variable 'value' inactivo.");
        resultados.push({ prueba: "VALIDADOR_BUG", estado: "CRÍTICO", detalle: "ReferenceError detectado en el validador original." });
      } else {
        console.log("   [PASS] El validador detuvo la ejecución con un mensaje controlado de campo obligatorio.");
        resultados.push({ prueba: "VALIDADOR_BUG", estado: "OK", detalle: "Validador limpio y libre de ReferenceError." });
      }
    }
  } catch (errGeneral) {
    console.error("   [FAIL] Fallo crítico al ejecutar validador: " + errGeneral.toString());
    resultados.push({ prueba: "VALIDADOR_BUG", estado: "FALLA", detalle: errGeneral.toString() });
  }

  // 3. Prueba de Algoritmo de Criptografía (SHA-256)
  console.log("\n3. Probando encriptación Hash SHA-256:");
  try {
    const hashOriginal = SEG_GENERAR_HASH_CONTRASENA("Admin123!");
    const hashEsperado = "316140f00d434a983d5fb6ead09b98a493c565978ea4fa79f9875a2461bd7df3";
    if (hashOriginal === hashEsperado) {
      console.log("   [PASS] Generador Hash SHA-256 produce resultados idénticos al estándar.");
      resultados.push({ prueba: "HASH_SHA256", estado: "OK", detalle: "Hash criptográfico seguro verificado." });
    } else {
      console.error("   [FAIL] Mismatch de Hash. Obtenido: " + hashOriginal);
      resultados.push({ prueba: "HASH_SHA256", estado: "FALLA", detalle: "Resultado hash inconsistente." });
    }
  } catch (errCrypt) {
    console.error("   [FAIL] Fallo al encriptar: " + errCrypt.toString());
    resultados.push({ prueba: "HASH_SHA256", estado: "FALLA", detalle: errCrypt.toString() });
  }

  // 4. Prueba del Algoritmo del Dígito de Verificación (DIAN)
  console.log("\n4. Probando cálculo matemático del Dígito de Verificación:");
  try {
    const dvCalculado = CLI_CALCULAR_DV("901915723"); // NIT real de MEGUDAN
    if (dvCalculado === "2") {
      console.log("   [PASS] NIT 901915723 calculado exitosamente con Dv 2.");
      resultados.push({ prueba: "DIAN_DV", estado: "OK", detalle: "Algoritmo de la DIAN operando al 100% de precisión." });
    } else {
      console.error("   [FAIL] Dv calculado incorrecto: " + dvCalculado + " (Esperado: 2).");
      resultados.push({ prueba: "DIAN_DV", estado: "FALLA", detalle: "Mismatch en el Dv matemático de Colombia." });
    }
  } catch (errDv) {
    console.error("   [FAIL] Fallo al calcular Dv: " + errDv.toString());
    resultados.push({ prueba: "DIAN_DV", estado: "FALLA", detalle: errDv.toString() });
  }

  // 5. Prueba de Roles y Autocuración (Self-Healing)
  console.log("\n5. Probando Inicialización y Autocuración de Roles y Permisos:");
  try {
    SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
    const roles = SEG_LISTAR_ROLES();
    if (roles.length >= 7) {
      console.log("   [PASS] Base de datos de roles autocurada y cargada con " + roles.length + " roles.");
      resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "OK", detalle: "Roles del sistema inicializados." });
    } else {
      console.warn("   [WARN] Roles incompletos detectados.");
      resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "WARN", detalle: "Faltan roles por poblar." });
    }
  } catch (errHealing) {
    console.error("   [FAIL] Error en autocuración: " + errHealing.toString());
    resultados.push({ prueba: "SELF_HEALING_ROLES", estado: "FALLA", detalle: errHealing.toString() });
  }

  // 6. Prueba de Seguridad Dual (Sheets vs Web App)
  console.log("\n6. Probando Validación de Contexto de Seguridad Dual:");
  try {
    // Escenario A: Contexto Sheets Local (Debería autorizar automáticamente)
    const accesoLocal = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(undefined, "SEGURIDAD", "CREAR");
    if (accesoLocal.AUTORIZADO && accesoLocal.CODIGO === "CONTEXTO_SHEETS_TRUSTED") {
      console.log("   [PASS] Escenario Sheets Local: Autorizado automáticamente con rol de Administrador.");
      resultados.push({ prueba: "SEGURIDAD_DUAL_LOCAL", estado: "OK", detalle: "Local Sheets trusted bypass verificado." });
    } else {
      console.error("   [FAIL] Escenario Sheets Local: Falló el bypass confiable.");
      resultados.push({ prueba: "SEGURIDAD_DUAL_LOCAL", estado: "FALLA", detalle: "Fallo en bypass local de Sheets." });
    }
    
    // Escenario B: Contexto Web sin Token (Debería denegar de forma segura)
    try {
      // Forzar contexto de simulación de Web App removiendo temporalmente SpreadsheetApp (envolviendo llamada para simular Web app)
      const mockWebCall = function() {
        // Al llamarse, simulamos que SpreadsheetApp.getUi lanza excepción como en la Web app
        const realGetUi = SpreadsheetApp.getUi;
        SpreadsheetApp.getUi = function() { throw new Error("No UI"); };
        try {
          SEG_VERIFICAR_CONTEXTO_Y_ACCESO(null, "SEGURIDAD", "CREAR");
        } finally {
          SpreadsheetApp.getUi = realGetUi; // Restablecer
        }
      };
      
      try {
        mockWebCall();
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
      }
    } catch (errContext) {
      console.error("   [FAIL] Error de contexto: " + errContext.toString());
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
    SEG_REGISTRAR_AUDITORIA({
      MODULO: "LOGS",
      SUBMODULO: "DIAGNOSTICO",
      ACCION: "TEST_DIAGNOSTICO",
      TIPO_REGISTRO: "SISTEMA",
      DESCRIPCION: "Diagnóstico completo del ERP ejecutado. Pasó: " + totalPasadas + ", Falló: " + totalCriticas,
      RESULTADO: totalCriticas === 0 ? "EXITOSO" : "ERROR",
      MENSAJE_RESULTADO: "Ejecución de test suite terminada."
    });
  } catch (e) {}

  return {
    EXITO: totalCriticas === 0,
    PASADAS: totalPasadas,
    FALLADAS: totalCriticas,
    DETALLE: resultados
  };
}
