/**************************************************************
* 03_MENU.gs
* RESPONSABILIDAD:
* - Administrar la navegación del ERP mediante enlaces dinámicos en Sheets.
* - Escanear la pestaña 🏠MENU para inyectar enlaces asíncronos en tiempo real.
**************************************************************/

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE MENÚ NAVEGABLE
// ============================================================
const MENU_CONFIG = {
  NOMBRE_MENU: "🏠MENU",
  SIMBOLO_ENLACE: "↳"
};

// ============================================================
// 02. ACTUALIZAR ENLACES DEL MENÚ EN GOOGLE SHEETS
// ============================================================
function ACTUALIZAR_MENU() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menu = ss.getSheetByName(MENU_CONFIG.NOMBRE_MENU);
  
  if (!menu) {
    throw new Error("No existe la hoja de Menú principal.");
  }

  const ultimaFila = menu.getLastRow();
  if (ultimaFila < 1) {
    throw new Error("El Menú de navegación se encuentra vacío.");
  }

  let enlacesCreados = 0;
  let noEncontradas = [];

  for (let fila = 1; fila <= ultimaFila; fila++) {
    const celda = menu.getRange(fila, 1);
    const texto = celda.getDisplayValue();

    if (!texto || !texto.includes(MENU_CONFIG.SIMBOLO_ENLACE)) {
      continue;
    }

    // Extraer y limpiar el nombre físico de la pestaña
    const nombreHoja = texto.replace(/↳/g, "").trim().replace(/\s+/g, " ");
    const hojaDestino = ss.getSheetByName(nombreHoja);

    if (!hojaDestino) {
      noEncontradas.push(nombreHoja);
      continue;
    }

    const url = ss.getUrl() + "#gid=" + hojaDestino.getSheetId();
    const textoVisible = MENU_CONFIG.SIMBOLO_ENLACE + " " + nombreHoja;

    // RichTextValue para máxima legibilidad e independencia de fórmulas locales
    const enlace = SpreadsheetApp.newRichTextValue()
      .setText(textoVisible)
      .setLinkUrl(url)
      .build();

    celda.setRichTextValue(enlace);
    enlacesCreados++;
  }

  SpreadsheetApp.flush();
  console.log("=== Sincronización de navegación finalizada ===");
  console.log("Enlaces creados exitosamente: " + enlacesCreados);
  if (noEncontradas.length > 0) {
    console.warn("Hojas no encontradas físicamente: " + noEncontradas.join(", "));
  }
}
