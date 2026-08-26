/**************************************************************
* 24_LOGS.gs
* RESPONSABILIDAD:
* - Administrar la capa de trazabilidad, logs y auditoría técnica del ERP.
* - Escribir en USR_AUDITORIA cualquier cambio crítico de datos.
**************************************************************/

const LOGS_CONFIG = {
  HOJA_AUDITORIA: "USR_AUDITORIA",
  PREFIJO_AUDITORIA: "AUD",
  DIGITOS_ID: 6
};

/**
 * Registra una acción de auditoría en la hoja USR_AUDITORIA de forma segura (sin fallos).
 * @param {object} datos Datos del evento a auditar.
 */
function LOGS_REGISTRAR_AUDITORIA(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(LOGS_CONFIG.HOJA_AUDITORIA);
    if (!hoja) return;

    const ultimaFila = hoja.getLastRow();
    const ahora = new Date();
    
    // Generar ID único incremental
    let sigNumero = 1;
    if (ultimaFila >= 2) {
      const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
      const num = parseInt(ultimoID.replace(LOGS_CONFIG.PREFIJO_AUDITORIA + "-", ""), 10);
      if (!isNaN(num)) {
        sigNumero = num + 1;
      }
    }
    const idAuditoria = LOGS_CONFIG.PREFIJO_AUDITORIA + "-" + String(sigNumero).padStart(LOGS_CONFIG.DIGITOS_ID, "0");

    // Estructurar fila alineada con columnas de USR_AUDITORIA
    const fila = [
      idAuditoria,
      ahora,
      datos.ID_USUARIO || "",
      datos.USUARIO || "SISTEMA",
      datos.ID_SESION || "",
      datos.ID_ROL || "",
      datos.MODULO || "SISTEMA",
      datos.SUBMODULO || "",
      datos.ACCION || "CREAR",
      datos.TIPO_REGISTRO || "",
      datos.ID_REGISTRO || "",
      datos.DESCRIPCION || "",
      datos.VALOR_ANTERIOR || "",
      datos.VALOR_NUEVO || "",
      datos.RESULTADO || "EXITOSO",
      datos.MENSAJE_RESULTADO || "",
      datos.ORIGEN_ACCESO || "SISTEMA",
      datos.IP_ORIGEN || "",
      ahora,
      datos.OBSERVACIONES || ""
    ];

    hoja.appendRow(fila);
  } catch (error) {
    console.error("Fallo crítico en el auditor técnico: " + error.toString());
  }
}