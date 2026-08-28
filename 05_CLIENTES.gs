/**************************************************************
* 05_CLIENTES.gs
* RESPONSABILIDAD:
* - Administrar el ciclo de vida (CRUD) de Clientes (CLI_MAESTRO).
* - Proteger accesos bajo la arquitectura de seguridad dual.
* - Calcular dígitos de verificación DIAN y registrar historiales.
**************************************************************/

const CLI_CONFIG = {
  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  FORMULARIO_HTML: "F1_CLI_FORM",
  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6
};

function CLI_CALCULAR_DV(nit) {
  nit = String(nit).trim().replace(/\D/g, "");
  if (!nit) return "";
  const vprimas = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  let suma = 0;
  const len = nit.length;
  for (let i = 0; i < len; i++) {
    suma += Number(nit.charAt(len - 1 - i)) * vprimas[i];
  }
  const residuo = suma % 11;
  return residuo > 1 ? String(11 - residuo) : String(residuo);
}

function CLI_VALIDAR_DUPLICADO(tipoDoc, numDoc, idExcluir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return false;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  return datos.some(fila => {
    if (idExcluir && String(fila[0]) === String(idExcluir)) return false;
    return String(fila[2]) === String(tipoDoc) && String(fila[3]) === String(numDoc);
  });
}

function CLI_GUARDAR_CLIENTE(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos) throw new Error("Datos no proporcionados.");
  
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }
  
  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    datos.RAZON_SOCIAL = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && n.trim() !== "").join(" ");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");
  
  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un cliente registrado.");
  }
  
  const idCliente = CLI_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  
  datos.ID_CLIENTE = idCliente;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_ACTUALIZACION = ahora;
  datos.USUARIO_CREACION = usuarioEjecutor;
  datos.USUARIO_ACTUALIZACION = usuarioEjecutor;
  
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);
  
  CLI_REGISTRAR_HISTORIAL(idCliente, "CREACION", "Cliente creado por " + usuarioEjecutor, usuarioEjecutor);
  return { ok: true, idCliente: idCliente, mensaje: "Cliente registrado correctamente." };
}

function CLI_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return CLI_CONFIG.PREFIJO_ID + "-000001";
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(CLI_CONFIG.PREFIJO_ID + "-", ""), 10);
  return CLI_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

function CLI_REGISTRAR_HISTORIAL(idCliente, accion, obs, usuarioEjecutor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
    if (!hoja) return;
    const fila = [
      "HIS-" + String(Math.max(1, hoja.getLastRow())).padStart(6, "0"),
      idCliente, new Date().toLocaleDateString(), new Date().toLocaleTimeString(),
      usuarioEjecutor || "SISTEMA", accion, "", "", "", "", "CLIENTES", idCliente, "", "ACTIVO", obs
    ];
    hoja.appendRow(fila);
  } catch (err) {
    console.error("No se pudo registrar log del cliente: " + idCliente);
  }
}

function CLI_ACTUALIZAR_CLIENTE(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos || !datos.ID_CLIENTE) throw new Error("ID_CLIENTE es obligatorio para actualizar.");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");
  
  const idCliente = datos.ID_CLIENTE;
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idxID = encabezados.indexOf("ID_CLIENTE");
  
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxID]) === String(idCliente)) {
      filaModificar = i + 2;
      break;
    }
  }
  
  if (filaModificar === -1) throw new Error("No se encontró el cliente '" + idCliente + "'.");
  
  const valoresFilaActual = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const clienteActualizado = {};
  encabezados.forEach((col, idx) => {
    clienteActualizado[col] = valoresFilaActual[idx];
  });
  
  const protegidos = ["ID_CLIENTE", "FECHA_CREACION", "USUARIO_CREACION"];
  Object.keys(datos).forEach(col => {
    const colUpper = col.toUpperCase();
    if (!protegidos.includes(colUpper) && encabezados.includes(colUpper)) {
      clienteActualizado[colUpper] = datos[col];
    }
  });
  
  clienteActualizado.FECHA_ACTUALIZACION = new Date();
  clienteActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;
  
  const filaNueva = encabezados.map(col => clienteActualizado[col] !== undefined ? clienteActualizado[col] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaNueva]);
  
  CLI_REGISTRAR_HISTORIAL(idCliente, "MODIFICACION", "Cliente actualizado por " + usuarioEjecutor, usuarioEjecutor);
  return { ok: true, idCliente: idCliente, mensaje: "Cliente actualizado correctamente." };
}

function CLI_INACTIVAR_CLIENTE(idCliente, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "ELIMINAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  
  return CLI_ACTUALIZAR_CLIENTE({ ID_CLIENTE: idCliente, ESTADO_CLIENTE: "INACTIVO" }, tokenSesion);
}

function CLI_BUSCAR_CLIENTE(criterio, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "VER");
  
  if (!criterio) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxDoc = encabezados.indexOf("NUMERO_DOCUMENTO");
  const idxRazon = encabezados.indexOf("RAZON_SOCIAL");
  
  const criterioNormalizado = String(criterio).trim().toUpperCase();
  const filaEncontrada = registros.find(fila => {
    return String(fila[idxId]).toUpperCase() === criterioNormalizado ||
           String(fila[idxDoc]).toUpperCase() === criterioNormalizado ||
           String(fila[idxRazon]).toUpperCase().includes(criterioNormalizado);
  });
  
  if (!filaEncontrada) return null;
  const cliente = {};
  encabezados.forEach((col, idx) => {
    cliente[col] = filaEncontrada[idx];
  });
  return cliente;
}

function CLI_SHEET_GUARDAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  const hojaMaestro = ss.getSheetByName("CLI_MAESTRO");
  if (!hojaForm || !hojaMaestro) throw new Error("Hojas CLI_FORM o CLI_MAESTRO no instaladas.");
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const valores = hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).getValues()[0];
  const datos = {};
  headers.forEach((h, idx) => { datos[h] = valores[idx]; });
  
  if (!datos.TIPO_PERSONA || !datos.TIPO_DOCUMENTO || !datos.NUMERO_DOCUMENTO || !datos.CORREO) {
    throw new Error("Por favor complete los campos obligatorios en la Fila 2.");
  }
  
  const res = CLI_GUARDAR_CLIENTE(datos, undefined); // Contexto Sheets (Token undefined)
  hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).clearContent();
  hojaForm.getRange(2, headers.indexOf("ID_CLIENTE") + 1).setValue(res.idCliente);
  hojaForm.getRange(2, headers.indexOf("ESTADO_CLIENTE") + 1).setValue("ACTIVO");
  
  const ui = SpreadsheetApp.getUi();
  if (ui) ui.alert("Éxito", "Cliente guardado correctamente con ID: " + res.idCliente, ui.ButtonSet.OK);
}

function CLI_SHEET_CARGAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  const hojaMaestro = ss.getSheetByName("CLI_MAESTRO");
  if (!hojaForm || !hojaMaestro) throw new Error("Hojas no instaladas.");
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const idxId = headers.indexOf("ID_CLIENTE");
  const idxDoc = headers.indexOf("NUMERO_DOCUMENTO");
  
  const idBusqueda = hojaForm.getRange(2, idxId + 1).getValue().toString().trim();
  const docBusqueda = hojaForm.getRange(2, idxDoc + 1).getValue().toString().trim();
  const criterio = idBusqueda || docBusqueda;
  if (!criterio) throw new Error("Debe ingresar un ID_CLIENTE o NUMERO_DOCUMENTO en la Fila 2.");
  
  const cliente = CLI_BUSCAR_CLIENTE(criterio, undefined);
  if (!cliente) throw new Error("No se encontró ningún cliente.");
  
  const filaValores = headers.map(h => cliente[h] !== undefined ? cliente[h] : "");
  hojaForm.getRange(2, 1, 1, headers.length).setValues([filaValores]);
  
  const ui = SpreadsheetApp.getUi();
  if (ui) ui.toast("Cliente cargado correctamente.", "ERP Operativo");
}

function CLI_SHEET_ACTUALIZAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  if (!hojaForm) throw new Error("No se encontró la hoja CLI_FORM.");
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const valores = hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).getValues()[0];
  const datos = {};
  headers.forEach((h, idx) => { datos[h] = valores[idx]; });
  
  if (!datos.ID_CLIENTE) throw new Error("ID_CLIENTE es requerido. Cargue un cliente primero.");
  
  CLI_ACTUALIZAR_CLIENTE(datos, undefined);
  const ui = SpreadsheetApp.getUi();
  if (ui) ui.alert("Éxito", "Cliente " + datos.ID_CLIENTE + " actualizado.", ui.ButtonSet.OK);
}

function CLI_SHEET_INACTIVAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  if (!hojaForm) throw new Error("No se encontró la hoja CLI_FORM.");
  
  const idCliente = hojaForm.getRange(2, 1).getValue().toString().trim();
  if (!idCliente) throw new Error("Debe cargar un cliente antes de inactivar.");
  
  CLI_INACTIVAR_CLIENTE(idCliente, undefined);
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  hojaForm.getRange(2, headers.indexOf("ESTADO_CLIENTE") + 1).setValue("INACTIVO");
  
  const ui = SpreadsheetApp.getUi();
  if (ui) ui.toast("Cliente inactivado.", "ERP Operativo");
}

function CLI_GENERAR_ID(tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "CREAR");
  }
  return CLI_OBTENER_SIGUIENTE_ID();
}
