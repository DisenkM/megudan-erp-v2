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

/**
 * Registra un movimiento físico de stock y recalcula saldos de promedio ponderado.
 * @param {object} mov Objeto con los datos del movimiento.
 */
function INV_REGISTRAR_MOVIMIENTO(mov) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMov = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
  const hojaSaldos = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
  const hojaKardex = ss.getSheetByName(INV_CONFIG.HOJA_KARDEX);

  if (!hojaMov || !hojaSaldos) throw new Error("Módulo de inventarios no instalado.");

  const idMovimiento = INV_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  const cant = Number(mov.CANTIDAD || 0);
  const costo = Number(mov.COSTO_UNITARIO || 0);

  // Buscar saldo existente del producto
  let saldoInicial = 0;
  let costoPromedio = 0;
  let filaSaldo = -1;

  if (hojaSaldos.getLastRow() >= 2) {
    const saldos = hojaSaldos.getRange(2, 2, hojaSaldos.getLastRow() - 1, 7).getValues();
    for (let i = 0; i < saldos.length; i++) {
      if (String(saldos[i][0]) === String(mov.ID_PRODUCTO)) {
        saldoInicial = Number(saldos[i][5]); // SALDO_FINAL
        costoPromedio = Number(saldos[i][6]); // COSTO_PROMEDIO
        filaSaldo = i + 2;
        break;
      }
    }
  }

  // Calcular nuevo Saldo Final y Costo Promedio Ponderado
  let saldoFinal = saldoInicial;
  let nuevoCostoTotal = saldoInicial * costoPromedio;

  if (mov.TIPO_MOVIMIENTO === "ENTRADA") {
    saldoFinal += cant;
    nuevoCostoTotal += (cant * costo);
    if (saldoFinal > 0) {
      costoPromedio = nuevoCostoTotal / saldoFinal;
    }
  } else if (mov.TIPO_MOVIMIENTO === "SALIDA") {
    saldoFinal -= cant;
    // La salida se costea al promedio anterior (Ponderado)
    nuevoCostoTotal -= (cant * costoPromedio);
  }

  // Registrar en la hoja de Movimientos
  hojaMov.appendRow([
    idMovimiento, ahora, mov.TIPO_MOVIMIENTO, mov.ID_PRODUCTO, mov.ID_OBRA || "",
    mov.DOCUMENTO_ORIGEN || "", mov.ID_ORIGEN || "", 
    mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : "",
    mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : "",
    costoPromedio, (cant * (mov.TIPO_MOVIMIENTO === "ENTRADA" ? costo : costoPromedio)),
    saldoFinal, Session.getActiveUser().getEmail() || "SISTEMA", ""
  ]);

  // Actualizar o insertar fila de Saldos
  if (filaSaldo !== -1) {
    hojaSaldos.getRange(filaSaldo, 3, 1, 7).setValues([[
      ahora, saldoInicial, 
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]]);
  } else {
    hojaSaldos.appendRow([
      "SAL-" + Utilities.getUuid().substring(0,8), mov.ID_PRODUCTO, ahora, 
      0, mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]);
  }

  // Registrar línea en el Kardex Histórico
  if (hojaKardex) {
    hojaKardex.appendRow([
      ahora, mov.ID_PRODUCTO, mov.TIPO_MOVIMIENTO, 
      (mov.DOCUMENTO_ORIGEN || "") + " " + (mov.ID_ORIGEN || ""),
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : "",
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : "",
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]);
  }
}

function INV_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return INV_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(INV_CONFIG.PREFIJO_ID + "-", ""), 10);
  return INV_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(INV_CONFIG.DIGITOS_ID, "0");
}