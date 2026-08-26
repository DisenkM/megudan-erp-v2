/**************************************************************
* 12_INGRESOS.gs
* RESPONSABILIDAD:
* - Capturar y clasificar los movimientos de ingresos del ERP.
* - Sincronizar el recaudo asíncronamente con Tesorería (TES_).
**************************************************************/

const ING_CONFIG = {
  HOJA_MOVIMIENTOS: "ING_MOVIMIENTOS",
  HOJA_RECAUDOS: "ING_RECAUDOS",
  PREFIJO_ID: "ING",
  DIGITOS_ID: 6
};

function ING_REGISTRAR_INGRESO(datos) {
  if (!datos) throw new Error("Datos de ingreso no válidos.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(ING_CONFIG.HOJA_MOVIMIENTOS);
  if (!hoja) throw new Error("Hoja de ingresos no instalada.");

  const idIngreso = ING_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  datos.ID_INGRESO = idIngreso;
  datos.FECHA = ahora;
  datos.ESTADO = "PROCESADO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  
  hoja.appendRow(fila);

  // Afectar Tesorería (Entrada de dinero)
  try {
    TES_REGISTRAR_MOVIMIENTO({
      TIPO_MOVIMIENTO: "INGRESO",
      ID_CUENTA: datos.ID_CUENTA_DESTINO,
      INGRESO: datos.VALOR,
      EGRESO: 0,
      METODO_PAGO: datos.METODO_PAGO,
      DOCUMENTO_ORIGEN: "INGRESO",
      ID_ORIGEN: idIngreso,
      OBSERVACION: datos.OBSERVACION
    });
  } catch (err) {
    console.warn("Tesorería no afectada por ingreso: " + idIngreso);
  }

  return { ok: true, idIngreso: idIngreso };
}

function ING_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(ING_CONFIG.HOJA_MOVIMIENTOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return ING_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(ING_CONFIG.PREFIJO_ID + "-", ""), 10);
  return ING_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(ING_CONFIG.DIGITOS_ID, "0");
}
