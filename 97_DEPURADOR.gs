// (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 97_DEPURADOR.gs (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Capturar, registrar y auditar todos los errores de ejecución en Apps Script (Backend).
* - Asignar valores por defecto (fallback) para evitar registros con 'undefined'.
* - Proveer un interceptor global de excepciones para llamadas google.script.run (Frontend/Backend).
* - Proporcionar herramientas de autodiagnóstico en tiempo real para verificar funciones y tablas.
**************************************************************/

const DEP_CONFIG = {
  HOJA_AUDITORIA: "USR_AUDITORIA",
  ENTORNO: "PRODUCCION",
  LOG_MODO_ESTRICTO: true
};

/**
 * 🛠️ 1. CAPTURADOR GLOBAL DE ERRORES BACKEND (CON BLINDAJE ANTI-UNDEFINED)
 * Recibe cualquier excepción capturada en bloques try-catch y la registra en USR_AUDITORIA.
 * Si se ejecuta sin parámetros, asigna automáticamente los valores por defecto.
 */
function DEP_REGISTRAR_ERROR_SISTEMA(origenFuncion, modulo, error, contextoAdicional) {
  const ahora = new Date();
  
  // Normalizar parámetros para evitar 'undefined' en los logs
  const fnOrigen = (origenFuncion && String(origenFuncion).trim() !== "") ? String(origenFuncion).trim() : "EJECUCION_DIRECTA_O_ANONIMA";
  const modOrigen = (modulo && String(modulo).trim() !== "") ? String(modulo).trim().toUpperCase() : "SISTEMA_GENERAL";
  
  let mensajeError = "Error no especificado";
  let stackTrace = "Sin traza de pila";
  
  if (error) {
    if (typeof error === "object") {
      mensajeError = error.message || error.toString() || "Error de objeto sin mensaje";
      stackTrace = error.stack ? String(error.stack).substring(0, 500) : "Sin traza de pila";
    } else {
      mensajeError = String(error);
    }
  } else {
    mensajeError = "Invocación directa de prueba o excepción sin objeto 'error'";
  }

  const detalleContexto = (contextoAdicional && String(contextoAdicional).trim() !== "") ? String(contextoAdicional).trim() : "N/A";
  const usuario = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "SISTEMA") : "SISTEMA";

  console.error("❌ [ERP ERROR] En [" + fnOrigen + "] (" + modOrigen + "): " + mensajeError + " | Contexto: " + detalleContexto);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaAuditoria = ss.getSheetByName(DEP_CONFIG.HOJA_AUDITORIA);
    
    if (!hojaAuditoria) {
      hojaAuditoria = ss.insertSheet(DEP_CONFIG.HOJA_AUDITORIA);
      hojaAuditoria.appendRow([
        "ID_AUDITORIA", "FECHA_HORA", "ID_USUARIO", "USUARIO", "ID_SESION", "ID_ROL",
        "MODULO", "SUBMODULO", "ACCION", "TIPO_REGISTRO", "ID_REGISTRO", "DESCRIPCION",
        "VALOR_ANTERIOR", "VALOR_NUEVO", "RESULTADO", "MENSAJE_RESULTADO", "ORIGEN_ACCESO",
        "IP_ORIGEN", "FECHA_CREACION", "OBSERVACIONES"
      ]);
    }
    
    const idError = "ERR-" + ahora.getTime();
    const encabezados = hojaAuditoria.getRange(1, 1, 1, hojaAuditoria.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    
    const registroObj = {
      ID_AUDITORIA: idError,
      FECHA_HORA: ahora,
      ID_USUARIO: usuario,
      USUARIO: usuario,
      MODULO: modOrigen,
      SUBMODULO: "DEPURADOR",
      ACCION: "ERROR_CAPTURADO",
      TIPO_REGISTRO: "EXCEPCION",
      ID_REGISTRO: fnOrigen,
      DESCRIPCION: "FALLO EN [" + fnOrigen + "]: " + mensajeError + " | Stack: " + stackTrace + " | Detalle: " + detalleContexto,
      RESULTADO: "ERROR",
      MENSAJE_RESULTADO: mensajeError,
      ORIGEN_ACCESO: "APPS_SCRIPT",
      FECHA_CREACION: ahora
    };

    const filaNueva = encabezados.map(col => registroObj[col] !== undefined ? registroObj[col] : "");
    hojaAuditoria.appendRow(filaNueva);
  } catch (errWrite) {
    console.error("⚠️ No se pudo escribir el log de error en Sheets: " + errWrite.toString());
  }

  return {
    EXITO: false,
    CODIGO_ERROR: "ERR_BACKEND_RPC",
    FUNCION_ORIGEN: fnOrigen,
    MODULO: modOrigen,
    MENSAJE: "Error en servidor [" + fnOrigen + "]: " + mensajeError
  };
}

/**
 * 🛠️ 2. SUITE DE DIAGNÓSTICO RÁPIDO DE FUNCIONES Y TABLAS
 * Ejecuta un chequeo integral de la salud del ERP y retorna un reporte estructurado
 */
function DEP_EJECUTAR_DIAGNOSTICO_SALUD_ERP() {
  console.log("==================================================================");
  console.log("🔍 INICIANDO DIAGNÓSTICO TÉCNICO DE SALUD DEL ERP (97_DEPURADOR v2)");
  console.log("==================================================================");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reporte = [];

  // A. Verificación de Hojas Clave en Google Sheets
  const tablasRequeridas = [
    "CFG_EMPRESA", "CFG_SISTEMA", "USR_USUARIOS", "USR_ROLES", "USR_PERMISOS",
    "USR_SESIONES", "USR_AUDITORIA", "CLI_MAESTRO", "PROV_MAESTRO", "PROD_MAESTRO",
    "VEN_CABECERA", "COM_CABECERA", "INV_SALDOS", "TES_CUENTAS", "NOM_EMPLEADOS"
  ];

  tablasRequeridas.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (hoja) {
      const filas = hoja.getLastRow();
      reporte.push({ tipo: "TABLA", entidad: nombreHoja, estado: "OK", detalle: "Hoja detectada con " + filas + " filas." });
      console.log("  ✓ Tablas: Hoja '" + nombreHoja + "' activa (" + filas + " filas).");
    } else {
      reporte.push({ tipo: "TABLA", entidad: nombreHoja, estado: "FAIL", detalle: "¡Hoja no encontrada en Google Sheets!" });
      console.error("  ❌ Tablas: Hoja '" + nombreHoja + "' NO existe.");
    }
  });

  // B. Verificación de Funciones Backend Críticas
  const funcionesGlobales = [
    "SEG_VALIDAR_SESION", "SEG_AUTENTICAR_USUARIO", "CLI_LISTAR_CLIENTES",
    "PROD_LISTAR_PRODUCTOS_WEB", "VEN_GUARDAR_VENTA_WEB", "NOM_LISTAR_EMPLEADOS_WEB",
    "FIN_CALCULAR_ESTADOS_FINANCIEROS_WEB", "WEB_doGet"
  ];

  funcionesGlobales.forEach(nombreFn => {
    if (typeof this[nombreFn] === "function") {
      reporte.push({ tipo: "FUNCION_RPC", entidad: nombreFn, estado: "OK", detalle: "Función declarada e invocable globalmente." });
      console.log("  ✓ RPC: Función '" + nombreFn + "()' verificada.");
    } else {
      reporte.push({ tipo: "FUNCION_RPC", entidad: nombreFn, estado: "FAIL", detalle: "¡Función ReferenceError / No definida!" });
      console.error("  ❌ RPC: Función '" + nombreFn + "()' NO declarada.");
    }
  });

  console.log("==================================================================");
  return {
    EXITO: true,
    TOTAL_CHEQUEOS: reporte.length,
    DETALLE: reporte
  };
}

/**
 * 🛠️ 3. SHIM Y POLYFILL DE DEPURACIÓN EN FRONTEND (HTML / JS)
 */
function DEP_OBTENER_SHIM_FRONTEND_HTML() {
  return `
  <script>
    /* Interceptor de Depuración Frontend ERP */
    window.addEventListener('error', function(e) {
      console.error('❌ [FRONTEND SYNTAX/RUNTIME ERROR]:', e.message, 'en línea:', e.lineno, 'archivo:', e.filename);
      var alertBox = document.getElementById('VEN_ALERT') || document.getElementById('NOM_ALERT') || document.getElementById('FIN_ALERT') || document.getElementById('USR_MENSAJE_ALERT');
      if (alertBox) {
        alertBox.textContent = '⚠️ Error JavaScript Frontend: ' + e.message + ' (Línea ' + e.lineno + ')';
        alertBox.className = 'message message-error';
        alertBox.style.display = 'block';
      }
    });

    /* Handler universal de errores RPC para google.script.run */
    function DEP_MANEJAR_ERROR_RPC(err, funcionOrigen) {
      console.error('❌ [RPC ERROR] Fallo al invocar ' + (funcionOrigen || 'servidor') + ':', err);
      alert('⚠️ Fallo de Comunicación con Apps Script:\n\nFunción: ' + (funcionOrigen || 'RPC') + '\nDetalle: ' + (err.message || err.toString()) + '\n\nVerifica que hayas publicado una "Nueva Versión" en Implementar.');
    }
  </script>
  `;
}

