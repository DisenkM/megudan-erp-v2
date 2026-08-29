/**************************************************************
* 03_MENU.gs
* RESPONSABILIDAD:
* - Compilación atómica e hipervínculos de la hoja principal 🏠MENU.
**************************************************************/

const MENU_CONFIG = {
  NOMBRE_MENU: "🏠MENU",
  SIMBOLO_ENLACE: "↳"
};

function ACTUALIZAR_MENU() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menu = ss.getSheetByName(MENU_CONFIG.NOMBRE_MENU);
  if (!menu) {
    throw new Error("No existe la hoja principal unificada '" + MENU_CONFIG.NOMBRE_MENU + "'.");
  }
  const ultimaFila = menu.getLastRow();
  if (ultimaFila < 1) return;

  const range = menu.getRange(1, 1, ultimaFila, 1);
  const displayValues = range.getDisplayValues();
  const currentRichText = range.getRichTextValues();
  const newRichText = [];

  const hojasDisponibles = ss.getSheets().map(h => h.getName());
  const cacheGids = {};
  ss.getSheets().forEach(h => {
    cacheGids[h.getName()] = h.getSheetId();
  });

  let enlaces = 0;
  const noEncontradas = [];

  for (let i = 0; i < displayValues.length; i++) {
    const texto = displayValues[i][0];
    if (!texto || !texto.includes(MENU_CONFIG.SIMBOLO_ENLACE)) {
      newRichText.push([currentRichText[i][0]]);
      continue;
    }

    const nombreHoja = OBTENER_NOMBRE_HOJA_MENU(texto);
    if (!nombreHoja || !hojasDisponibles.includes(nombreHoja)) {
      noEncontradas.push(nombreHoja || texto);
      newRichText.push([currentRichText[i][0]]);
      continue;
    }

    const gid = cacheGids[nombreHoja];
    const url = ss.getUrl() + "#gid=" + gid;
    const textoVisible = MENU_CONFIG.SIMBOLO_ENLACE + " " + nombreHoja;

    const enlace = SpreadsheetApp.newRichTextValue()
      .setText(textoVisible)
      .setLinkUrl(url)
      .build();

    newRichText.push([enlace]);
    enlaces++;
  }

  range.setRichTextValues(newRichText);
  SpreadsheetApp.flush();
  REGISTRAR_RESULTADO_MENU(enlaces, noEncontradas);
}

function OBTENER_NOMBRE_HOJA_MENU(texto) {
  if (!texto) return "";
  let nombre = String(texto);
  nombre = nombre.replace(/↳/g, "");
  nombre = nombre.trim();
  nombre = nombre.replace(/\s+/g, " ");
  return nombre;
}

function REGISTRAR_RESULTADO_MENU(enlaces, noEncontradas) {
  console.log("=================================");
  console.log("🔗 ACTUALIZACIÓN DEL MENÚ ERP (🏠MENU)");
  console.log("=================================");
  console.log("Enlaces creados o actualizados: " + enlaces);
  console.log("Hojas no encontradas: " + noEncontradas.length);
  if (noEncontradas.length > 0) {
    console.warn("Hojas no encontradas en el libro: " + noEncontradas.join(", "));
  }
  console.log("=================================");
}
