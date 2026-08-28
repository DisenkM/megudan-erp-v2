/**************************************************************
* 06_PROVEEDORES.gs
* RESPONSABILIDAD:
* - Administrar el catálogo y operaciones (CRUD) de Proveedores.
* - Control de NITs colombianos y cálculo de DV reutilizando clientes.
* - Soportar la arquitectura de seguridad dual.
**************************************************************/

const PROV_CONFIG = {
  HOJA_MAESTRO: "PROV_MAESTRO",
  HOJA_HISTORIAL: "PROV_HISTORIAL",
  PREFIJO_ID: "PROV",
  DIGITOS_ID: 6
};

function PROV_GUARDAR_PROVEEDOR(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos) throw new Error("Datos de proveedor no especificados.");

  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DV = CLI_CALCULAR_DV(datos.NIT_CC);
  } else {
    datos.DV = "";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROV_MAESTRO no existe.");

  const idProveedor = PROV_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  datos.ID_PROVEEDOR = idProveedor;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_MODIFICACION = ahora;
  datos.ESTADO = "ACTIVO";

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);

  return { ok: true, idProveedor: idProveedor, mensaje: "Proveedor guardado exitosamente." };
}

function PROV_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return PROV_CONFIG.PREFIJO_ID + "-000001";
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(PROV_CONFIG.PREFIJO_ID + "-", ""), 10);
  return PROV_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(PROV_CONFIG.DIGITOS_ID, "0");
}

function PROV_ACTUALIZAR_PROVEEDOR(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos || !datos.ID_PROVEEDOR) throw new Error("ID_PROVEEDOR es requerido.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROV_MAESTRO no encontrada.");

  const idProveedor = datos.ID_PROVEEDOR;
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idxID = encabezados.indexOf("ID_PROVEEDOR");
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();

  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxID]) === String(idProveedor)) {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar === -1) throw new Error("No se encontró el proveedor '" + idProveedor + "'.");

  const valoresFilaActual = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const proveedorActualizado = {};
  encabezados.forEach((col, idx) => {
    proveedorActualizado[col] = valoresFilaActual[idx];
  });

  const protegidos = ["ID_PROVEEDOR", "FECHA_CREACION", "USUARIO_CREACION"];
  Object.keys(datos).forEach(col => {
    const colUpper = col.toUpperCase();
    if (!protegidos.includes(colUpper) && encabezados.includes(colUpper)) {
      proveedorActualizado[colUpper] = datos[col];
    }
  });

  proveedorActualizado.FECHA_MODIFICACION = new Date();
  const filaNueva = encabezados.map(col => proveedorActualizado[col] !== undefined ? proveedorActualizado[col] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaNueva]);

  return { ok: true, idProveedor: idProveedor, mensaje: "Proveedor actualizado con éxito." };
}

function PROV_INACTIVAR_PROVEEDOR(idProveedor, tokenSesion) {
  return PROV_ACTUALIZAR_PROVEEDOR({ ID_PROVEEDOR: idProveedor, ESTADO: "INACTIVO" }, tokenSesion);
}

function PROV_BUSCAR_PROVEEDOR(criterio, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "VER");

  if (!criterio) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();

  const idxId = encabezados.indexOf("ID_PROVEEDOR");
  const idxDoc = encabezados.indexOf("NIT_CC");
  const idxRazon = encabezados.indexOf("RAZON_SOCIAL");
  const criterioNormalizado = String(criterio).trim().toUpperCase();

  const filaEncontrada = registros.find(fila => {
    return String(fila[idxId]).toUpperCase() === criterioNormalizado ||
           String(fila[idxDoc]).toUpperCase() === criterioNormalizado ||
           String(fila[idxRazon]).toUpperCase().includes(criterioNormalizado);
  });

  if (!filaEncontrada) return null;
  const proveedor = {};
  encabezados.forEach((col, idx) => {
    proveedor[col] = filaEncontrada[idx];
  });
  return proveedor;
}

function PROV_GENERAR_ID(tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "CREAR");
  }
  return PROV_OBTENER_SIGUIENTE_ID();
}
