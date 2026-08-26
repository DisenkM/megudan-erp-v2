// ============================================================
// 24. LOGS Y AUDITORÍA CENTRALIZADA
// ARCHIVO: 24_LOGS.gs
// RESPONSABILIDAD: Registro unificado de eventos de sistema, auditoría y control de errores
// ============================================================

/**
 * Registra un error de forma centralizada en el sistema.
 * Guarda detalles ricos como fecha, función, módulo, mensaje de error y el stack trace.
 * Adicionalmente, escribe el evento en el panel de auditoría física USR_AUDITORIA de forma segura.
 * 
 * @param {string} funcion - Nombre de la función donde se capturó el error.
 * @param {string} modulo - Nombre del módulo o archivo (ej. "SEGURIDAD", "MENU", "VENTAS").
 * @param {Error|Object} error - El objeto de error de JavaScript capturado en el catch.
 */
function LOG_REGISTRAR_ERROR(funcion, modulo, error) {
  const ahora = new Date();
  const mensaje = error.message || error.toString();
  const stack = error.stack || "No disponible";
  const usuario = Session.getActiveUser().getEmail() || "SISTEMA";
  
  // 1. Registro oficial en la consola de Google Cloud (Stackdriver Logs)
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
  
  // 2. Registro físico tolerante a fallos en USR_AUDITORIA (Base de Datos)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAuditoria = ss.getSheetByName("USR_AUDITORIA");
    if (!hojaAuditoria) return; // Si no existe la hoja, evitamos colapsar el sistema
    
    // Obtenemos un secuencial de ID usando el generador si está disponible
    let idAuditoria = "AUD-ERR-" + ahora.getTime();
    if (typeof SEG_GENERAR_ID === "function") {
      try {
        idAuditoria = SEG_GENERAR_ID("USR_AUDITORIA", "ID_AUDITORIA", "AUD");
      } catch (e) {
        // Fallback a timestamp si falla el generador secuencial
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
    // Si la escritura física en Sheets falla, el error queda en Stackdriver para no causar loops de fallos
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
