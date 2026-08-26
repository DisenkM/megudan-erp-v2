/**************************************************************
* 20_CONSECUTIVOS.gs
* RESPONSABILIDAD:
* - Administrar y centralizar los rangos de numeración del ERP.
* - Garantizar el incremento secuencial atómico evitando saltos accidentales de consecutivo.
**************************************************************/

const CONS_CONFIG = {
  HOJA_CONSECUTIVOS: "CFG_DOCUMENTOS"
};

/**
 * Incrementa de forma atómica el consecutivo de un documento y lo reserva.
 * @param {string} tipoDocumento Tipo de documento (ej: "FACTURA_VENTA").
 * @returns {string} Consecutivo asignado con prefijo (ej: "FE-10002").
 */
function CONS_OBTENER_CONSECUTIVO_PROXIMO(tipoDocumento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONS_CONFIG.HOJA_CONSECUTIVOS);
  if (!hoja) throw new Error("Hoja de consecutivo no instalada.");

  const datos = hoja.getRange(2, 2, hoja.getLastRow() - 1, 4).getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0].toString().trim().toUpperCase() === tipoDocumento.toUpperCase()) {
      const prefijo = datos[i][1].toString();
      const actual = parseInt(datos[i][2], 10);
      const proximo = actual + 1;

      // Actualizar en caliente
      hoja.getRange(i + 2, 4).setValue(proximo);
      return prefijo + String(proximo);
    }
  }
  throw new Error("Rango de numeración no configurado para: " + tipoDocumento);
}
