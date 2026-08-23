// ============================================================
// 01. EVENTOS PRINCIPALES DEL ERP
// ============================================================


// ============================================================
// 01.01 EVENTO AL ABRIR EL ERP
// Crea el menú principal de formularios.
// ============================================================

function onOpen() {

  const ui =
    SpreadsheetApp.getUi();


  ui.createMenu("FORMULARIOS")

    .addItem(
      "👥 Abrir Clientes",
      "ABRIR_CLIENTES"
    )

    .addItem(
      "👤 Abrir Gestión de Usuarios",
      "ABRIR_USUARIOS"
    )

    .addToUi();

}