/**************************************************************
* 05_CLIENTES.gs
* RESPONSABILIDAD:
* - Administrar el ciclo de vida (CRUD) del maestro de Clientes.
* - Validar datos tributarios de Colombia (NIT, DV automático, responsabilidades fiscales).
* - Registrar automáticamente historial de cambios y auditorías operativas.
**************************************************************/

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE CLIENTES
// ============================================================
const CLI_CONFIG = {
  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  FORMULARIO_HTML: "F1_CLI_FORM",
  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6
};

// ============================================================
// 02. ALGORITMO OFICIAL DE LA DIAN: CALCULAR DV
// ============================================================
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
  if (residuo > 1) {
    return String(11 - residuo);
  }
  return String(residuo);
}

// ============================================================
// 03. VALIDACIÓN DE DUPLICADOS EN BASE DE DATOS
// ============================================================
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

// ============================================================
// 04. OPERACIONES CRUD: GUARDAR Y ACTUALIZAR CLIENTES
// ============================================================
function CLI_GUARDAR_CLIENTE(datos) {
  if (!datos) throw new Error("Datos no proporcionados.");
  
  // Forzar cálculo asíncrono seguro del DV para NITs
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }
  
  // Dinamismo del Tipo de Persona
  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    datos.RAZON_SOCIAL = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && n.trim() !== "").join(" ");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");
  
  // Validar unicidad
  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un cliente registrado.");
  }
  
  // Generar ID e inyectar auditorías
  const idCliente = CLI_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  const usuario = Session.getActiveUser().getEmail() || "SISTEMA";
  
  datos.ID_CLIENTE = idCliente;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_ACTUALIZACION = ahora;
  datos.USUARIO_CREACION = usuario;
  datos.USUARIO_ACTUALIZACION = usuario;
  
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  hoja.appendRow(fila);
  
  // Registrar Historial
  CLI_REGISTRAR_HISTORIAL(idCliente, "CREACION", "Cliente creado de forma exitosa.");
  return { ok: true, mensaje: "Cliente creado de forma exitosa.", idCliente: idCliente };
}

function CLI_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  
  if (ultimaFila < 2) {
    return CLI_CONFIG.PREFIJO_ID + "-000001";
  }
  
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(CLI_CONFIG.PREFIJO_ID + "-", ""), 10);
  return CLI_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

function CLI_REGISTRAR_HISTORIAL(idCliente, accion, obs) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
    if (!hoja) return;
    
    const fila = [
      "HIS-" + String(Math.max(1, hoja.getLastRow())).padStart(6, "0"),
      idCliente,
      new Date().toLocaleDateString(),
      new Date().toLocaleTimeString(),
      Session.getActiveUser().getEmail() || "SISTEMA",
      accion,
      "", "", "", "", "CLIENTES", idCliente, "", "ACTIVO", obs
    ];
    hoja.appendRow(fila);
  } catch (err) {
    console.error("No se pudo registrar log del cliente: " + idCliente);
  }
}

/**
 * Actualiza los campos de un cliente en CLI_MAESTRO.
 */
function CLI_ACTUALIZAR_CLIENTE(datos) {
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
      filaModificar = i + 2; // +2 por índice base 1 y encabezados
      break;
    }
  }
  
  if (filaModificar === -1) {
    throw new Error("No se encontró el cliente '" + idCliente + "' para actualizar.");
  }
  
  // Conservar valores previos y mezclar nuevos
  const valoresFilaActual = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const clienteActualizado = {};
  encabezados.forEach((col, idx) => {
    clienteActualizado[col] = valoresFilaActual[idx];
  });
  
  // Actualizar con datos recibidos (excepto campos protegidos)
  const protegidos = ["ID_CLIENTE", "FECHA_CREACION", "USUARIO_CREACION"];
  Object.keys(datos).forEach(col => {
    const colUpper = col.toUpperCase();
    if (!protegidos.includes(colUpper) && encabezados.includes(colUpper)) {
      clienteActualizado[colUpper] = datos[col];
    }
  });
  
  clienteActualizado.FECHA_ACTUALIZACION = new Date();
  clienteActualizado.USUARIO_ACTUALIZACION = Session.getActiveUser().getEmail() || "SISTEMA";
  
  const filaNueva = encabezados.map(col => clienteActualizado[col] !== undefined ? clienteActualizado[col] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaNueva]);
  
  CLI_REGISTRAR_HISTORIAL(idCliente, "MODIFICACION", "Cliente actualizado de forma exitosa.");
  return { ok: true, idCliente: idCliente };
}

/**
 * Inactiva un cliente modificando su estado a INACTIVO.
 */
function CLI_INACTIVAR_CLIENTE(idCliente) {
  return CLI_ACTUALIZAR_CLIENTE({ ID_CLIENTE: idCliente, ESTADO_CLIENTE: "INACTIVO" });
}

/**
 * Busca un cliente por NIT, CC, Razón Social o ID del ERP.
 */
function CLI_BUSCAR_CLIENTE(criterio) {
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

// ============================================================
// 05. OPERACIONES EXCLUSIVAS DE HOJAS DE GOOGLE SHEETS (CRUD LOCAL)
// ============================================================

/**
 * Guarda un cliente desde la fila de captura (Fila 2) de la hoja CLI_FORM.
 */
function CLI_SHEET_GUARDAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  const hojaMaestro = ss.getSheetByName("CLI_MAESTRO");
  
  if (!hojaForm || !hojaMaestro) {
    throw new Error("No se encontraron las hojas físicas CLI_FORM o CLI_MAESTRO.");
  }
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const valores = hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).getValues()[0];
  
  const datos = {};
  headers.forEach((h, idx) => {
    datos[h] = valores[idx];
  });
  
  if (!datos.TIPO_PERSONA || !datos.TIPO_DOCUMENTO || !datos.NUMERO_DOCUMENTO || !datos.CORREO) {
    throw new Error("Por favor complete los campos obligatorios: TIPO_PERSONA, TIPO_DOCUMENTO, NUMERO_DOCUMENTO y CORREO en la Fila 2.");
  }
  
  const res = CLI_GUARDAR_CLIENTE(datos);
  
  hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).clearContent();
  hojaForm.getRange(2, headers.indexOf("ID_CLIENTE") + 1).setValue(res.idCliente);
  hojaForm.getRange(2, headers.indexOf("ESTADO_CLIENTE") + 1).setValue("ACTIVO");
  
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert("Éxito", "Cliente guardado correctamente con ID: " + res.idCliente, ui.ButtonSet.OK);
  }
}

/**
 * Carga la información de un cliente en la fila de captura (Fila 2) de CLI_FORM buscando por el ID_CLIENTE o NUMERO_DOCUMENTO ingresado.
 */
function CLI_SHEET_CARGAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  const hojaMaestro = ss.getSheetByName("CLI_MAESTRO");
  
  if (!hojaForm || !hojaMaestro) {
    throw new Error("No se encontraron las hojas físicas CLI_FORM o CLI_MAESTRO.");
  }
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const idxId = headers.indexOf("ID_CLIENTE");
  const idxDoc = headers.indexOf("NUMERO_DOCUMENTO");
  
  const idBusqueda = hojaForm.getRange(2, idxId + 1).getValue().toString().trim();
  const docBusqueda = hojaForm.getRange(2, idxDoc + 1).getValue().toString().trim();
  
  let criterio = idBusqueda || docBusqueda;
  if (!criterio) {
    throw new Error("Debe ingresar un ID_CLIENTE o NUMERO_DOCUMENTO en la Fila 2 para realizar la búsqueda.");
  }
  
  const cliente = CLI_BUSCAR_CLIENTE(criterio);
  if (!cliente) {
    throw new Error("No se encontró ningún cliente con el criterio provisto: " + criterio);
  }
  
  const filaValores = headers.map(h => cliente[h] !== undefined ? cliente[h] : "");
  hojaForm.getRange(2, 1, 1, headers.length).setValues([filaValores]);
  
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.toast("Cliente cargado correctamente.", "ERP Operativo");
  }
}

/**
 * Actualiza un cliente en CLI_MAESTRO con los valores modificados en la Fila 2 de CLI_FORM.
 */
function CLI_SHEET_ACTUALIZAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  
  if (!hojaForm) {
    throw new Error("No se encontró la hoja física CLI_FORM.");
  }
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  const valores = hojaForm.getRange(2, 1, 1, hojaForm.getLastColumn()).getValues()[0];
  
  const datos = {};
  headers.forEach((h, idx) => {
    datos[h] = valores[idx];
  });
  
  if (!datos.ID_CLIENTE) {
    throw new Error("No se puede actualizar. Primero debe cargar un cliente existente usando 'Cargar por ID/NIT'.");
  }
  
  CLI_ACTUALIZAR_CLIENTE(datos);
  
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert("Éxito", "Cliente " + datos.ID_CLIENTE + " actualizado correctamente.", ui.ButtonSet.OK);
  }
}

/**
 * Inactiva un cliente cargado en la Fila 2 de CLI_FORM.
 */
function CLI_SHEET_INACTIVAR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaForm = ss.getSheetByName("CLI_FORM");
  
  if (!hojaForm) {
    throw new Error("No se encontró la hoja física CLI_FORM.");
  }
  
  const idCliente = hojaForm.getRange(2, 1).getValue().toString().trim();
  if (!idCliente) {
    throw new Error("Debe cargar un cliente en el formulario antes de inactivarlo.");
  }
  
  CLI_INACTIVAR_CLIENTE(idCliente);
  
  const headers = hojaForm.getRange(1, 1, 1, hojaForm.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
  hojaForm.getRange(2, headers.indexOf("ESTADO_CLIENTE") + 1).setValue("INACTIVO");
  
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.toast("Cliente inactivado correctamente.", "ERP Operativo");
  }
}


/**
 * Wrapper de compatibilidad para el frontend del formulario de Clientes.
 */
function CLI_GENERAR_ID() {
  return CLI_OBTENER_SIGUIENTE_ID();
}
