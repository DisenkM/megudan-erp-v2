/// ============================================================
// 01. CORE - NAVEGACIÓN Y CONFIGURACIÓN DE INTERFAZ SPREADSHEET
// ARCHIVO: 01_CORE.gs
// RESPONSABILIDAD: Menús personalizados, diálogos modales y utilidades de UI seguras
// ============================================================

/**
 * Genera de forma segura el menú personalizado en Google Sheets.
 * Es invocado por el disparador central onOpen() en 22_TRIGGERS.gs.
 */
function CORE_onOpen() {
  const ui = CORE_OBTENER_UI();
  if (!ui) {
    console.log("onOpen ejecutado en contexto sin interfaz gráfica (Headless).");
    return;
  }
  try {
    ui.createMenu("FORMULARIOS")
      .addItem("👥 Abrir Clientes Web", "ABRIR_CLIENTES")
      .addItem("👤 Abrir Gestión de Usuarios Web", "ABRIR_USUARIOS")
      .addSeparator()
      .addSubMenu(ui.createMenu("Clientes en Hoja")
        .addItem("💾 Guardar Nuevo", "CLI_SHEET_GUARDAR")
        .addItem("🔍 Cargar por ID/NIT", "CLI_SHEET_CARGAR")
        .addItem("✏️ Actualizar Existente", "CLI_SHEET_ACTUALIZAR")
        .addItem("🚫 Inactivar Tercero", "CLI_SHEET_INACTIVAR"))
      .addSubMenu(ui.createMenu("Proveedores en Hoja")
        .addItem("💾 Guardar Nuevo", "PROV_SHEET_GUARDAR")
        .addItem("🔍 Cargar por ID/NIT", "PROV_SHEET_CARGAR")
        .addItem("✏️ Actualizar Existente", "PROV_SHEET_ACTUALIZAR")
        .addItem("🚫 Inactivar Tercero", "PROV_SHEET_INACTIVAR"))
      .addToUi();
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("CORE_onOpen", "CORE", error);
    }
  }
}

/**
 * Abre la ventana modal del Formulario de Clientes de forma segura.
 */
function ABRIR_CLIENTES() {
  const ui = CORE_OBTENER_UI();
  if (!ui) {
    console.warn("Spreadsheet UI no disponible.");
    return;
  }
  try {
    const template = HtmlService.createTemplateFromFile("F1_CLI_FORM");
    template.TOKEN_SESION = "SHEETS_CONTEXT";
    template.USUARIO_ACTUAL = Session.getActiveUser().getEmail() || "ADMIN_SHEETS";
    const html = template.evaluate()
      .setWidth(1200)
      .setHeight(800);
    ui.showModalDialog(html, "👥 Gestión de Clientes");
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("ABRIR_CLIENTES", "CORE", error);
    }
    ui.alert("Error", "No se pudo abrir el formulario de clientes: " + error.message, ui.ButtonSet.OK);
  }
}

function ABRIR_USUARIOS() {
  const ui = CORE_OBTENER_UI();
  if (!ui) {
    console.warn("Spreadsheet UI no disponible.");
    return;
  }
  try {
    const template = HtmlService.createTemplateFromFile("F2_USR_GESTION");
    template.TOKEN_SESION = "SHEETS_CONTEXT";
    template.USUARIO_ACTUAL = Session.getActiveUser().getEmail() || "ADMIN_SHEETS";
    const html = template.evaluate()
      .setWidth(1200)
      .setHeight(800);
    ui.showModalDialog(html, "🔐 Gestión de Seguridad y Control de Acceso");
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("ABRIR_USUARIOS", "CORE", error);
    }
    ui.alert("Error", "No se pudo abrir la gestión de seguridad: " + error.message, ui.ButtonSet.OK);
  }
}

function CORE_OBTENER_UI() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}
