/**************************************************************
* 07_PRODUCTOS.gs
* RESPONSABILIDAD:
* - Administrar el catálogo de Productos, Categorías y Precios.
* - Validar e integrar stocks de Guadua.
**************************************************************/

const PROD_CONFIG = {
  HOJA_MAESTRO: "PROD_MAESTRO",
  PREFIJO_ID: "PRD",
  DIGITOS_ID: 6
};

function PROD_GUARDAR_PRODUCTO(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "CREAR");
  if (!datos) throw new Error("Datos del producto no válidos.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROD_MAESTRO no encontrada.");

  const idProducto = PROD_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  datos.ID_PRODUCTO = idProducto;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_MODIFICACION = ahora;
  datos.ESTADO = "ACTIVO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);

  return { ok: true, idProducto: idProducto, mensaje: "Producto registrado exitosamente." };
}

function PROD_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return PROD_CONFIG.PREFIJO_ID + "-000001";
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(PROD_CONFIG.PREFIJO_ID + "-", ""), 10);
  return PROD_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(PROD_CONFIG.DIGITOS_ID, "0");
}

function PROD_BUSCAR_PRODUCTO(codigo, tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "VER");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();

  const idxCodigo = encabezados.indexOf("CODIGO");
  const fila = registros.find(f => String(f[idxCodigo]).toUpperCase() === String(codigo).trim().toUpperCase());
  if (!fila) return null;

  const prod = {};
  encabezados.forEach((col, idx) => {
    prod[col] = fila[idx];
  });
  return prod;
}
