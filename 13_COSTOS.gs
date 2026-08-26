/**************************************************************
* 13_COSTOS.gs
* RESPONSABILIDAD:
* - Controlar el registro de Costos Operativos asociados a Obras.
* - Generar alertas automáticas de desviación de presupuesto.
**************************************************************/

const COS_CONFIG = {
  HOJA_MOVIMIENTOS: "COS_MOVIMIENTOS",
  PREFIJO_ID: "COS",
  DIGITOS_ID: 6
};

function COS_REGISTRAR_COSTO(datos) {
  if (!datos) throw new Error("Datos de costo vacíos.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(COS_CONFIG.HOJA_MOVIMIENTOS);
  if (!hoja) throw new Error("Hoja COS_MOVIMIENTOS no encontrada.");

  const idCosto = COS_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  datos.ID_COSTO = idCosto;
  datos.FECHA = ahora;
  datos.ESTADO = "PROCESADO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  
  hoja.appendRow(fila);
  return { ok: true, idCosto: idCosto };
}

function COS_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(COS_CONFIG.HOJA_MOVIMIENTOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return COS_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(COS_CONFIG.PREFIJO_ID + "-", ""), 10);
  return COS_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(COS_CONFIG.DIGITOS_ID, "0");
}
