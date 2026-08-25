// ============================================================
// CLIENTES.gs
// MÓDULO: CLIENTES
// ============================================================

const CLI_CONFIG = {
  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  FORMULARIO: "F1_CLI_FORM",
  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6
};

function ABRIR_CLIENTES() {
  const html = HtmlService
    .createHtmlOutputFromFile(CLI_CONFIG.FORMULARIO)
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "👥 Gestión de Clientes");
}

function CLI_GENERAR_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("No existe la hoja " + CLI_CONFIG.HOJA_MAESTRO);
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return CLI_CONFIG.PREFIJO_ID + "-000001";
  
  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues().flat();
  let mayorNumero = 0;
  ids.forEach(function(id) {
    if (!id) return;
    const partes = String(id).split("-");
    const numero = Number(partes[1]);
    if (!isNaN(numero) && numero > mayorNumero) {
      mayorNumero = numero;
    }
  });
  return CLI_CONFIG.PREFIJO_ID + "-" + String(mayorNumero + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

/**
 * Calcula el Dígito de Verificación (DV) oficial para Colombia (Algoritmo DIAN/Siigo).
 * Expuesta para llamadas asíncronas desde el Frontend HTML.
 */
function CLI_CALCULAR_DV(nit) {
  if (!nit || isNaN(nit)) return "";
  const factores = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  let suma = 0;
  const temp = nit.toString().trim();
  for (let i = 0; i < temp.length; i++) {
    suma += parseInt(temp.charAt(temp.length - 1 - i), 10) * factores[i];
  }
  const residuo = suma % 11;
  return residuo > 1 ? 11 - residuo : residuo;
}

function CLI_VALIDAR_DUPLICADO(tipoDocumento, numeroDocumento, idClienteActual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("No existe la hoja CLI_MAESTRO.");
  if (hoja.getLastRow() < 2) return false;

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxTipo = encabezados.indexOf("TIPO_DOCUMENTO");
  const idxNumero = encabezados.indexOf("NUMERO_DOCUMENTO");

  if (idxId === -1 || idxTipo === -1 || idxNumero === -1) {
    throw new Error("Faltan columnas requeridas en CLI_MAESTRO.");
  }

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  return datos.some(function(fila) {
    if (String(fila[idxId]) === String(idClienteActual || "")) return false;
    return (String(fila[idxTipo]) === String(tipoDocumento) && String(fila[idxNumero]) === String(numeroDocumento));
  });
}

function CLI_OBTENER_USUARIO_ACTUAL() {
  return Session.getActiveUser().getEmail() || "USUARIO_SISTEMA";
}

function CLI_OBTENER_FECHA_ACTUAL() {
  return new Date();
}

/**
 * Guarda un nuevo cliente aplicando validaciones estrictas de DIAN y Siigo
 */
function CLI_GUARDAR_CLIENTE(datos) {
  if (!datos) throw new Error("No se recibieron datos del cliente.");

  // Validaciones de obligatoriedad estructural
  const obligatorios = ["TIPO_PERSONA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO", "TIPO_CLIENTE", "PAIS"];
  obligatorios.forEach(function(campo) {
    if (!datos[campo] || String(datos[campo]).trim() === "") {
      throw new Error("El campo " + campo + " es obligatorio.");
    }
  });

  // Validaciones dinámicas por tipo de persona (Natural vs Jurídica)
  if (datos.TIPO_PERSONA === "PERSONA_JURIDICA") {
    if (!datos.RAZON_SOCIAL || datos.RAZON_SOCIAL.trim() === "") {
      throw new Error("La Razón Social es obligatoria para Personas Jurídicas.");
    }
    // Limpiar campos exclusivos de persona natural
    datos.PRIMER_NOMBRE = "";
    datos.SEGUNDO_NOMBRE = "";
    datos.PRIMER_APELLIDO = "";
    datos.SEGUNDO_APELLIDO = "";
  } else if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    if (!datos.PRIMER_NOMBRE || datos.PRIMER_NOMBRE.trim() === "" || !datos.PRIMER_APELLIDO || datos.PRIMER_APELLIDO.trim() === "") {
      throw new Error("El Primer Nombre y Primer Apellido son obligatorios para Personas Naturales.");
    }
    // Autocomponer Razón Social a partir de los nombres (Para mantener integridad comercial y fiscal)
    datos.RAZON_SOCIAL = (datos.PRIMER_NOMBRE + " " + (datos.SEGUNDO_NOMBRE || "") + " " + datos.PRIMER_APELLIDO + " " + (datos.SEGUNDO_APELLIDO || "")).replace(/\s+/g, ' ').trim();
  }

  // Asegurar el cálculo del Dígito de Verificación si el tipo de documento es NIT
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }

  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("Ya existe un cliente con este documento.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("No existe la hoja CLI_MAESTRO.");

  const idCliente = CLI_GENERAR_ID();
  const fecha = CLI_OBTENER_FECHA_ACTUAL();
  const usuario = CLI_OBTENER_USUARIO_ACTUAL();

  datos.ID_CLIENTE = idCliente;
  datos.FECHA_CREACION = fecha;
  datos.FECHA_ACTUALIZACION = fecha;
  datos.USUARIO_CREACION = usuario;
  datos.USUARIO_ACTUALIZACION = usuario;

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const nuevaFila = encabezados.map(function(campo) {
    return datos[campo] !== undefined ? datos[campo] : "";
  });

  hoja.appendRow(nuevaFila);

  CLI_REGISTRAR_HISTORIAL(
    idCliente, "CREACION", "CREAR", "", "", "", "CLIENTES", idCliente, usuario, "ACTIVO", "Cliente creado correctamente."
  );

  return {
    ok: true,
    mensaje: "Cliente registrado correctamente.",
    idCliente: idCliente
  };
}

function CLI_BUSCAR_CLIENTE(criterio) {
  criterio = String(criterio || "").trim();
  if (!criterio) throw new Error("Ingrese un ID de cliente, NIT, CC o razón social.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("No existe la hoja " + CLI_CONFIG.HOJA_MAESTRO + ".");

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) return null;

  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0].map(function(e) {
    return String(e).trim().toUpperCase();
  });

  const registros = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getDisplayValues();

  function normalizarTexto(valor) {
    return String(valor || "").trim().toUpperCase().replace(/\s+/g, " ");
  }

  function normalizarDocumento(valor) {
    return String(valor || "").trim().replace(/[.\s,\-]/g, "");
  }

  const criterioTexto = normalizarTexto(criterio);
  const criterioDocumento = normalizarDocumento(criterio);

  const idxIdCliente = encabezados.indexOf("ID_CLIENTE");
  const idxDocumento = encabezados.indexOf("NUMERO_DOCUMENTO");
  const idxRazonSocial = encabezados.indexOf("RAZON_SOCIAL");

  if (idxIdCliente === -1 || idxDocumento === -1 || idxRazonSocial === -1) {
    throw new Error("Faltan columnas requeridas en CLI_MAESTRO.");
  }

  const filaEncontrada = registros.find(function(fila) {
    const idCliente = normalizarTexto(fila[idxIdCliente]);
    const documento = normalizarDocumento(fila[idxDocumento]);
    const razonSocial = normalizarTexto(fila[idxRazonSocial]);

    if (idCliente === criterioTexto) return true;
    if (documento === criterioDocumento) return true;
    if (razonSocial.includes(criterioTexto)) return true;
    return false;
  });

  if (!filaEncontrada) return null;

  const cliente = {};
  encabezados.forEach(function(campo, indice) {
    cliente[campo] = filaEncontrada[indice];
  });

  return cliente;
}

function CLI_ACTUALIZAR_CLIENTE(datos) {
  if (!datos || !datos.ID_CLIENTE) throw new Error("No se indicó el ID del cliente.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) throw new Error("No existen clientes registrados.");

  // Validaciones dinámicas por tipo de persona antes de actualizar
  if (datos.TIPO_PERSONA === "PERSONA_JURIDICA") {
    if (!datos.RAZON_SOCIAL || datos.RAZON_SOCIAL.trim() === "") {
      throw new Error("La Razón Social es obligatoria para Personas Jurídicas.");
    }
    datos.PRIMER_NOMBRE = "";
    datos.SEGUNDO_NOMBRE = "";
    datos.PRIMER_APELLIDO = "";
    datos.SEGUNDO_APELLIDO = "";
  } else if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    if (!datos.PRIMER_NOMBRE || datos.PRIMER_NOMBRE.trim() === "" || !datos.PRIMER_APELLIDO || datos.PRIMER_APELLIDO.trim() === "") {
      throw new Error("El Primer Nombre y Primer Apellido son obligatorios para Personas Naturales.");
    }
    datos.RAZON_SOCIAL = (datos.PRIMER_NOMBRE + " " + (datos.SEGUNDO_NOMBRE || "") + " " + datos.PRIMER_APELLIDO + " " + (datos.SEGUNDO_APELLIDO || "")).replace(/\s+/g, ' ').trim();
  }

  // Actualizar el DV si el documento es NIT
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxRegistro = registros.findIndex(function(fila) {
    return String(fila[idxId]) === String(datos.ID_CLIENTE);
  });

  if (idxId === -1 || idxRegistro === -1) throw new Error("No se encontró el cliente.");

  const datosAnteriores = registros[idxRegistro];

  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, datos.ID_CLIENTE)) {
    throw new Error("Ya existe otro cliente con este documento.");
  }

  const usuario = CLI_OBTENER_USUARIO_ACTUAL();
  datos.FECHA_ACTUALIZACION = CLI_OBTENER_FECHA_ACTUAL();
  datos.USUARIO_ACTUALIZACION = usuario;

  const idxFechaCreacion = encabezados.indexOf("FECHA_CREACION");
  const idxUsuarioCreacion = encabezados.indexOf("USUARIO_CREACION");
  if (idxFechaCreacion !== -1) datos.FECHA_CREACION = datosAnteriores[idxFechaCreacion];
  if (idxUsuarioCreacion !== -1) datos.USUARIO_CREACION = datosAnteriores[idxUsuarioCreacion];

  const nuevaFila = encabezados.map(function(campo, indice) {
    return datos[campo] !== undefined ? datos[campo] : datosAnteriores[indice];
  });

  hoja.getRange(idxRegistro + 2, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  encabezados.forEach(function(campo, indice) {
    if (String(datosAnteriores[indice]) !== String(nuevaFila[indice]) && campo !== "FECHA_ACTUALIZACION" && campo !== "USUARIO_ACTUALIZACION") {
      CLI_REGISTRAR_HISTORIAL(
        datos.ID_CLIENTE, "MODIFICACION", "EDITAR", campo, datosAnteriores[indice], nuevaFila[indice], "CLIENTES", datos.ID_CLIENTE, usuario, "ACTIVO", "Campo modificado: " + campo
      );
    }
  });

  return {
    ok: true,
    mensaje: "Cliente actualizado correctamente.",
    idCliente: datos.ID_CLIENTE
  };
}

function CLI_INACTIVAR_CLIENTE(idCliente) {
  const cliente = CLI_BUSCAR_CLIENTE(idCliente);
  if (!cliente) throw new Error("Cliente no encontrado.");
  if (cliente.ESTADO_CLIENTE === "INACTIVO") throw new Error("El cliente ya está inactivo.");
  cliente.ESTADO_CLIENTE = "INACTIVO";
  CLI_ACTUALIZAR_CLIENTE(cliente);
  return { ok: true, mensaje: "Cliente inactivado correctamente." };
}

function CLI_REGISTRAR_HISTORIAL(idCliente, tipoEvento, accion, campoModificado, valorAnterior, valorNuevo, moduloOrigen, idRegistroOrigen, usuario, estadoEvento, observaciones) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
  if (!hoja) throw new Error("No existe la hoja CLI_HISTORIAL.");

  const consecutivo = Math.max(hoja.getLastRow() - 1, 0) + 1;
  const evento = {
    ID_HISTORIAL: "HIS-" + String(consecutivo).padStart(6, "0"),
    ID_CLIENTE: idCliente,
    TIPO_EVENTO: tipoEvento,
    FECHA_HORA: CLI_OBTENER_FECHA_ACTUAL(),
    USUARIO: usuario,
    ACCION: accion,
    CAMPO_MODIFICADO: campoModificado,
    VALOR_ANTERIOR: valorAnterior,
    VALOR_NUEVO: valorNuevo,
    MOTIVO_ORIGEN: "",
    MODULO_ORIGEN: moduloOrigen,
    ID_REGISTRO_ORIGEN: idRegistroOrigen,
    IP_USUARIO: "",
    ESTADO_EVENTO: estadoEvento,
    OBSERVACIONES: observaciones
  };

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(function(campo) {
    return evento[campo] !== undefined ? evento[campo] : "";
  });
  hoja.appendRow(fila);
}