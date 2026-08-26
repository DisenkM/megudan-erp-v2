/**************************************************************
* 02_CONFIGURACION.gs
* RESPONSABILIDAD:
* - Administrar las configuraciones locales de forma centralizada.
* - Leer y actualizar parámetros tributarios y de sistema desde Sheets.
**************************************************************/

const CONF_LOCAL = {
  HOJA_SISTEMA: "CFG_SISTEMA",
  HOJA_EMPRESA: "CFG_EMPRESA"
};

/**
 * Lee un parámetro de configuración general desde la hoja CFG_SISTEMA.
 * @param {string} parametro Nombre del parámetro a buscar.
 * @returns {string} Valor de la configuración encontrada.
 */
function CONF_LEER_PARAMETRO(parametro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONF_LOCAL.HOJA_SISTEMA);
  if (!hoja) return "";
  
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  const resultado = datos.find(fila => fila[0].toString().trim().toUpperCase() === parametro.toUpperCase());
  return resultado ? resultado[1] : "";
}

/**
 * Actualiza el valor de un parámetro de configuración en CFG_SISTEMA.
 * @param {string} parametro Nombre del parámetro.
 * @param {string} nuevoValor Nuevo valor a asignar.
 * @returns {boolean} Retorna verdadero si la operación fue exitosa.
 */
function CONF_ESCRIBIR_PARAMETRO(parametro, nuevoValor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONF_LOCAL.HOJA_SISTEMA);
  if (!hoja) return false;
  
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0].toString().trim().toUpperCase() === parametro.toUpperCase()) {
      hoja.getRange(i + 2, 2).setValue(nuevoValor);
      return true;
    }
  }
  return false;
}