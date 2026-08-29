/**************************************************************
* 11_INVENTARIO.gs
* RESPONSABILIDAD:
* - Administrar las existencias de productos en bodegas y obras.
* - Calcular de forma automatizada el costo promedio ponderado.
* - Registrar movimientos de Kardex (entradas, salidas, ajustes).
**************************************************************/

const INV_CONFIG = {
  HOJA_MOVIMIENTOS: "INV_MOVIMIENTOS",
  HOJA_SALDOS: "INV_SALDOS",
  HOJA_KARDEX: "INV_KARDEX",
  PREFIJO_ID: "MOV",
  DIGITOS_ID: 6
};

function INV_REGISTRAR_MOVIMIENTO(mov) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMov = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
  const hojaSaldos = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
  if (!hojaMov || !hojaSaldos) throw new Error("Hojas de inventario no encontradas.");

  const idMovimiento = INV_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(INV_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();

  const cant = Number(mov.CANTIDAD || 0);
  const costo = Number(mov.COSTO_UNITARIO || 0);

  let filaSaldo = -1;
  let saldoActual = 0;
  let costoPromedio = costo;

  if (hojaSaldos.getLastRow() >= 2) {
    const saldos = hojaSaldos.getRange(2, 2, hojaSaldos.getLastRow() - 1, 7).getValues();
    for (let i = 0; i < saldos.length; i++) {
      if (String(saldos[i][0]) === String(mov.ID_PRODUCTO)) {
        filaSaldo = i + 2;
        saldoActual = Number(saldos[i][5] || 0);
        costoPromedio = Number(saldos[i][6] || 0);
        break;
      }
    }
  }

  let saldoFinal = saldoActual;
  if (mov.TIPO_MOVIMIENTO === "ENTRADA") {
    const costoAnteriorTotal = saldoActual * costoPromedio;
    const costoNuevoTotal = cant * costo;
    saldoFinal = saldoActual + cant;
    if (saldoFinal > 0) {
      costoPromedio = (costoAnteriorTotal + costoNuevoTotal) / saldoFinal;
    }
  } else if (mov.TIPO_MOVIMIENTO === "SALIDA") {
    saldoFinal = saldoActual - cant;
  }

  hojaMov.appendRow([
    idMovimiento, ahora, mov.TIPO_MOVIMIENTO, mov.ID_PRODUCTO, mov.ID_OBRA || "",
    mov.DOCUMENTO_ORIGEN || "", mov.ID_ORIGEN || "",
    mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : "",
    mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : "",
    costoPromedio, (cant * (mov.TIPO_MOVIMIENTO === "ENTRADA" ? costo : costoPromedio)),
    saldoFinal, Session.getActiveUser().getEmail() || "SISTEMA", ""
  ]);

  if (filaSaldo !== -1) {
    hojaSaldos.getRange(filaSaldo, 3, 1, 7).setValues([[
      ahora, saldoActual,
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]]);
  } else {
    hojaSaldos.appendRow([
      "SAL-" + idMovimiento, mov.ID_PRODUCTO, ahora, 0,
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]);
  }
}
