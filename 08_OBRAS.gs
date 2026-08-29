/**************************************************************
* 08_OBRAS.gs
* RESPONSABILIDAD:
* - Administrar la base de datos de Obras y Proyectos.
* - Controlar asignación de recursos y ejecución de presupuestos de Guadua.
**************************************************************/

const OBR_CONFIG = {
  HOJA_MAESTRO: "OBR_MAESTRO",
  HOJA_PRESUPUESTO: "OBR_PRESUPUESTO",
  PREFIJO_ID: "OBR",
  DIGITOS_ID: 6
};

function OBR_GUARDAR_OBRA(datos, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "OBRAS", "CREAR");
  if (!datos) throw new Error("Datos de obra incompletos.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(OBR_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja OBR_MAESTRO no instalada.");

  const idObra = OBR_OBTENER_SIGUIENTE_ID();
  datos.ID_OBRA = idObra;
  datos.FECHA_CREACION = new Date();
  datos.ESTADO = "ACTIVA";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);

  return { ok: true, idObra: idObra, mensaje: "Obra guardada exitosamente." };
}

function OBR_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(OBR_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return OBR_CONFIG.PREFIJO_ID + "-000001";
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(OBR_CONFIG.PREFIJO_ID + "-", ""), 10);
  return OBR_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(OBR_CONFIG.DIGITOS_ID, "0");
}