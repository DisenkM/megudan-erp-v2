/**************************************************************
* 16_CxP.gs
* RESPONSABILIDAD:
* - Administrar las Cuentas por Pagar asociadas a compras.
* - Controlar saldos y amortizar obligaciones con abonos.
**************************************************************/

const CXP_CONFIG = {
  HOJA_CUENTAS: "CXP_CUENTAS",
  PREFIJO_ID: "CXP",
  DIGITOS_ID: 6
};

function CXP_CREAR_CUENTA_PAGAR(idCompra, idProveedor, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CXP_CONFIG.HOJA_CUENTAS);
  if (!hoja) return;

  const idCxp = CXP_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CXP_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();
  const vencimiento = new Date(ahora.getTime() + (30 * 24 * 60 * 60 * 1000));

  hoja.appendRow([
    idCxp, idProveedor, idCompra, idCompra, ahora, vencimiento, total, 0, total, 0, "PENDIENTE"
  ]);
}
