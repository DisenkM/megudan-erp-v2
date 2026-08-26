/**************************************************************
* 14_GASTOS.gs
* RESPONSABILIDAD:
* - Registrar y clasificar los Gastos Operativos y Administrativos.
* - Afectar de manera automatizada las cuentas de egreso de Tesorería.
**************************************************************/

const GAS_CONFIG = {
  HOJA_MOVIMIENTOS: "GAS_MOVIMIENTOS",
  PREFIJO_ID: "GAS",
  DIGITOS_ID: 6
};

function GAS_REGISTRAR_GASTO(datos) {
  if (!datos) throw new Error("Datos de gasto vacíos.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(GAS_CONFIG.HOJA_MOVIMIENTOS);
  if (!hoja) throw new Error("Hoja GAS_MOVIMIENTOS no existe.");

  const idGasto = GAS_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  datos.ID_GASTO = idGasto;
  datos.FECHA = ahora;
  datos.ESTADO = "PROCESADO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  
  hoja.appendRow(fila);

  // Afectar Tesorería (Egreso de dinero)
  try {
    TES_REGISTRAR_MOVIMIENTO({
      TIPO_MOVIMIENTO: "EGRESO",
      ID_CUENTA: datos.ID_CUENTA,
      INGRESO: 0,
      EGRESO: datos.VALOR,
      METODO_PAGO: datos.METODO_PAGO,
      DOCUMENTO_ORIGEN: "GASTO",
      ID_ORIGEN: idGasto,
      OBSERVACION: datos.OBSERVACION
    });
  } catch (err) {
    console.warn("Tesorería no afectada por gasto: " + idGasto);
  }

  return { ok: true, idGasto: idGasto };
}

function GAS_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(GAS_CONFIG.HOJA_MOVIMIENTOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return GAS_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(GAS_CONFIG.PREFIJO_ID + "-", ""), 10);
  return GAS_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(GAS_CONFIG.DIGITOS_ID, "0");
}