/**************************************************************
* 07_PRODUCTOS.gs
* RESPONSABILIDAD:
* - Administrar el catálogo de Productos, Categorías y Precios.
* - Validar e integrar stocks de Guadua.
* - Exponer la consulta y edición de productos para la Web App.
**************************************************************/

const PROD_CONFIG = {
  HOJA_MAESTRO: "PROD_MAESTRO",
  PREFIJO_ID: "PRD",
  DIGITOS_ID: 6
};

function PROD_GUARDAR_PRODUCTO(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
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

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);

  return { ok: true, idProducto: idProducto, mensaje: "Producto registrado exitosamente en el catálogo." };
}

function PROD_ACTUALIZAR_PRODUCTO(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  if (!datos || !datos.ID_PRODUCTO) throw new Error("ID_PRODUCTO es requerido para actualizar.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROD_MAESTRO no encontrada.");

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const idxID = encabezados.indexOf("ID_PRODUCTO");
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();

  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxID]) === String(datos.ID_PRODUCTO)) {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar === -1) throw new Error("No se encontró el producto '" + datos.ID_PRODUCTO + "'.");

  // Cargar producto actual
  const valoresFilaActual = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const productoActualizado = {};
  encabezados.forEach((col, idx) => {
    productoActualizado[col] = valoresFilaActual[idx];
  });

  const protegidos = ["ID_PRODUCTO", "FECHA_CREACION", "USUARIO_CREACION"];
  Object.keys(datos).forEach(col => {
    const colUpper = col.toUpperCase();
    if (!protegidos.includes(colUpper) && encabezados.includes(colUpper)) {
      productoActualizado[colUpper] = datos[col];
    }
  });

  productoActualizado.FECHA_MODIFICACION = new Date();
  productoActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;

  const filaNueva = encabezados.map(col => productoActualizado[col] !== undefined ? productoActualizado[col] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaNueva]);

  return { ok: true, idProducto: datos.ID_PRODUCTO, mensaje: "Producto actualizado correctamente en el catálogo." };
}

function PROD_ELIMINAR_PRODUCTO(idProducto, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "ELIMINAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  if (!idProducto) throw new Error("El idProducto es requerido.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROD_MAESTRO no encontrada.");

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const idxID = encabezados.indexOf("ID_PRODUCTO");
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();

  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxID]) === String(idProducto)) {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar === -1) throw new Error("No se encontró el producto '" + idProducto + "'.");

  const idxEstado = encabezados.indexOf("ESTADO");
  hoja.getRange(filaModificar, idxEstado + 1).setValue("INACTIVO");

  return { ok: true, idProducto: idProducto, mensaje: "Producto marcado como INACTIVO de forma segura." };
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

function PROD_LISTAR_PRODUCTOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PRODUCTOS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(PROD_CONFIG.HOJA_MAESTRO);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "La hoja PROD_MAESTRO no existe en la base de datos." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay productos registrados." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Productos obtenidos de forma segura."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("PROD_LISTAR_PRODUCTOS_WEB", "PRODUCTOS", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los productos: " + error.message };
  }
}