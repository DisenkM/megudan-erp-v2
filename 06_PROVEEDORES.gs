/**************************************************************
* 06_PROVEEDORES.gs
* RESPONSABILIDAD:
* - Administrar el catálogo y operaciones (CRUD) de Proveedores.
* - Control de NITs colombianos y cálculo automático de DV.
**************************************************************/

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE PROVEEDORES
// ============================================================
const PROV_CONFIG = {
  HOJA_MAESTRO: "PROV_MAESTRO",
  HOJA_HISTORIAL: "PROV_HISTORIAL",
  PREFIJO_ID: "PROV",
  DIGITOS_ID: 6
};

// ============================================================
// 02. OPERACIONES DE BACKEND
// ============================================================
function PROV_GUARDAR_PROVEEDOR(datos) {
  if (!datos) throw new Error("Datos de proveedor no especificados.");
  
  // Calcular DV automático de forma segura si es NIT
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DV = CLI_CALCULAR_DV(datos.NIT_CC); // Reutiliza la función matemática de clientes
  } else {
    datos.DV = "";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROV_MAESTRO no existe.");

  const idProveedor = PROV_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  const usuario = Session.getActiveUser().getEmail() || "SISTEMA";

  datos.ID_PROVEEDOR = idProveedor;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_MODIFICACION = ahora;
  datos.ESTADO = "ACTIVO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  
  hoja.appendRow(fila);
  return { ok: true, idProveedor: idProveedor };
}

function PROV_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return PROV_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(PROV_CONFIG.PREFIJO_ID + "-", ""), 10);
  return PROV_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(PROV_CONFIG.DIGITOS_ID, "0");
}