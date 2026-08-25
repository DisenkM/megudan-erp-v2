// ============================================================
// EVENTOS PRINCIPALES Y LANZADORES DEL ERP
// ARCHIVO: 01_CORE.gs
// ============================================================

/**
 * Evento oficial de Google Sheets que crea el menú superior al abrir el archivo.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("FORMULARIOS")
    .addItem("👥 Abrir Clientes", "ABRIR_CLIENTES")
    .addItem("👤 Abrir Gestión de Usuarios", "ABRIR_USUARIOS")
    .addToUi();
}

/**
 * Lanzador para el módulo de Gestión de Seguridad y Usuarios (F2_USR_GESTION).
 */
function ABRIR_USUARIOS() {
  // Asegúrate de que el archivo HTML en Apps Script se llame exactamente "F2_USR_GESTION"
  const html = HtmlService
    .createHtmlOutputFromFile("F2_USR_GESTION")
    .setWidth(1200)
    .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(html, "🔐 Gestión de Seguridad y Usuarios");
}