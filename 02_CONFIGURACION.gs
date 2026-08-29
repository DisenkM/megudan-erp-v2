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

function CONF_LEER_PARAMETRO(parametro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONF_LOCAL.HOJA_SISTEMA);
    if (!hoja) return null;
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0].toString().trim().toUpperCase() === parametro.toUpperCase()) {
        return datos[i][1];
      }
    }
  } catch (e) {
    console.error("Error al leer parámetro: " + parametro);
  }
  return null;
}

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
