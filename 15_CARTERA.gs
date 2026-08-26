/**************************************************************
* 15_CARTERA.gs
* RESPONSABILIDAD:
* - Controlar el saldo y antigüedad de Cuentas por Cobrar de clientes.
* - Generar el recaudo contra las obligaciones vigentes de Cartera.
**************************************************************/

const CAR_CONFIG = {
  HOJA_CUENTAS: "CAR_CUENTAS",
  PREFIJO_ID: "CAR",
  DIGITOS_ID: 6
};

function CAR_CREAR_CUENTA_COBRAR(idVenta, idCliente, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CAR_CONFIG.HOJA_CUENTAS);
  if (!hoja) return;

  const idCartera = CAR_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CAR_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();
  const vencimiento = new Date(ahora.getTime() + (30 * 24 * 60 * 60 * 1000)); // 30 días plazo por defecto

  hoja.appendRow([
    idCartera, idCliente, idVenta, idVenta, ahora, vencimiento, total, 0, total, 0, "PENDIENTE"
  ]);
}