/**************************************************************
* 22_TRIGGERS.gs
* RESPONSABILIDAD:
* - Administrar y enrutar de forma segura los disparadores asíncronos de Google Sheets.
* - Redireccionar el flujo onEdit() hacia los validadores específicos.
**************************************************************/

const TRIG_CONFIG = {
  REGISTRAR_ACTIVIDAD: true,
  AUDITAR_MODIFICACIONES: true
};

/**
 * Disparador onEdit nativo de Google Sheets.
 * Redirecciona cambios hacia validaciones asíncronas para optimizar velocidad.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const rango = e.range;
  const hoja = rango.getSheet();
  const nombreHoja = hoja.getName();
  const valorNuevo = e.value;
  const valorAnterior = e.oldValue;

  // Lógica de enrutamiento rápido según hoja editada
  if (nombreHoja === "🏠MENU" && valorNuevo === "RECOMPILAR") {
    ACTUALIZAR_MENU();
    rango.setValue("↳ NAVEGAR");
  }
}