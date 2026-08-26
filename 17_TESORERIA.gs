/**************************************************************
* 17_TESORERIA.gs
* RESPONSABILIDAD:
* - Administrar cuentas bancarias y caja física del ERP.
* - Registrar movimientos consolidados de entradas y salidas de fondos.
**************************************************************/

const TES_CONFIG = {
  HOJA_CUENTAS: "TES_CUENTAS",
  HOJA_MOVIMIENTOS: "TES_MOVIMIENTOS",
  PREFIJO_ID: "MOV",
  DIGITOS_ID: 6
};

function TES_REGISTRAR_MOVIMIENTO(mov) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCuentas = ss.getSheetByName(TES_CONFIG.HOJA_CUENTAS);
  const hojaMov = ss.getSheetByName(TES_CONFIG.HOJA_MOVIMIENTOS);

  if (!hojaCuentas || !hojaMov) throw new Error("Módulo de tesorería no instalado.");

  const idMovimiento = TES_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(TES_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();

  // Buscar y actualizar saldo de la cuenta de destino
  let saldoActual = 0;
  let filaCuenta = -1;

  if (hojaCuentas.getLastRow() >= 2) {
    const cuentas = hojaCuentas.getRange(2, 1, hojaCuentas.getLastRow() - 1, 7).getValues();
    for (let i = 0; i < cuentas.length; i++) {
      if (String(cuentas[i][0]) === String(mov.ID_CUENTA)) {
        saldoActual = Number(cuentas[i][6] || 0); // SALDO_ACTUAL
        filaCuenta = i + 2;
        break;
      }
    }
  }

  // Calcular nuevo saldo consolidado
  const ingreso = Number(mov.INGRESO || 0);
  const egreso = Number(mov.EGRESO || 0);
  const nuevoSaldo = saldoActual + ingreso - egreso;

  // Actualizar cuenta física
  if (filaCuenta !== -1) {
    hojaCuentas.getRange(filaCuenta, 7).setValue(nuevoSaldo);
  }

  // Escribir movimiento físico
  hojaMov.appendRow([
    idMovimiento, ahora, mov.TIPO_MOVIMIENTO, mov.ID_CUENTA, mov.DOCUMENTO_ORIGEN,
    mov.ID_ORIGEN, ingreso, egreso, nuevoSaldo, mov.METODO_PAGO, 
    Session.getActiveUser().getEmail() || "SISTEMA", mov.OBSERVACION || ""
  ]);
}