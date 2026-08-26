/**************************************************************
* 01_CORE.gs
* RESPONSABILIDAD:
* - Definir el punto de enlace y eventos globales de Google Sheets.
* - Generar el menú nativo superior para el lanzamiento de los formularios modal.
**************************************************************/

// ============================================================
// 01. EVENTOS GLOBALES DE GOOGLE SHEETS
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("FORMULARIOS")
    .addItem("👥 Abrir Formulario de Clientes", "ABRIR_CLIENTES")
    .addItem("👤 Abrir Gestión de Usuarios", "ABRIR_USUARIOS")
    .addToUi();
}

// ============================================================
// 02. LANZADORES DE FORMULARIOS MODALES (FRONTEND)
// ============================================================
function ABRIR_CLIENTES() {
  // Asegúrate de que el archivo HTML se llame exactamente "F1_CLI_FORM"
  const html = HtmlService
    .createHtmlOutputFromFile("F1_CLI_FORM")
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "👥 Gestión de Clientes - MEGUDAN");
}

function ABRIR_USUARIOS() {
  // Asegúrate de que el archivo HTML se llame exactamente "F2_USR_GESTION"
  const html = HtmlService
    .createHtmlOutputFromFile("F2_USR_GESTION")
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "🔐 Control de Seguridad y Usuarios");
}
