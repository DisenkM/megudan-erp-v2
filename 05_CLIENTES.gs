/**************************************************************
* 05_CLIENTES.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el ciclo de vida (CRUD) de Clientes (CLI_MAESTRO).
* - Proteger accesos bajo la arquitectura de seguridad dual.
* - Calcular dígitos de verificación DIAN y registrar historiales.
* - Mapear identificadores de Siigo y Alegra desde el día uno.
* - Determinar automáticamente retenciones basadas en responsabilidades.
**************************************************************/

const CLI_CONFIG = {
  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6
};

/**
 * Obtiene la hoja física de clientes
 */
function CLI_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No existe la hoja física '" + nombreHoja + "' en la base de datos.");
  return hoja;
}

/**
 * Obtiene los encabezados de la hoja de clientes
 */
function CLI_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = CLI_OBTENER_HOJA(nombreHoja);
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaColumna === 0) throw new Error("La hoja '" + nombreHoja + "' no contiene encabezados.");
  return hoja.getRange(1, 1, 1, ultimaColumna).getDisplayValues()[0].map(h => String(h || "").trim().toUpperCase());
}

/**
 * Obtiene todos los registros de la hoja de clientes
 */
function CLI_OBTENER_REGISTROS(nombreHoja) {
  const hoja = CLI_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
}

/**
 * Calcula el Dígito de Verificación (DV) oficial de la DIAN (Módulo 11)
 */
function CLI_CALCULAR_DV(numDoc) {
  if (!numDoc) return "";
  const nit = String(numDoc).replace(/\D/g, "");
  if (!nit || isNaN(nit)) return "";
  
  const v = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  let suma = 0;
  const len = nit.length;
  for (let i = 0; i < len; i++) {
    suma += parseInt(nit.charAt(len - 1 - i), 10) * v[i];
  }
  const residuo = suma % 11;
  return (residuo > 1) ? String(11 - residuo) : String(residuo);
}

/**
 * Valida si un documento ya se encuentra registrado
 */
function CLI_VALIDAR_DUPLICADO(tipoDoc, numDoc, idExcluir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return false;
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxTipoDoc = encabezados.indexOf("TIPO_DOCUMENTO");
  const idxNumDoc = encabezados.indexOf("NUMERO_DOCUMENTO");
  
  if (idxId === -1 || idxTipoDoc === -1 || idxNumDoc === -1) return false;
  
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getDisplayValues();
  return datos.some(fila => {
    if (idExcluir && String(fila[idxId]) === String(idExcluir)) return false;
    return String(fila[idxTipoDoc]).trim().toUpperCase() === String(tipoDoc).trim().toUpperCase() && 
           String(fila[idxNumDoc]).trim().replace(/\D/g, "") === String(numDoc).trim().replace(/\D/g, "");
  });
}

/**
 * Obtiene el siguiente ID de cliente disponible
 */
function CLI_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) return CLI_CONFIG.PREFIJO_ID + "-" + String(1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return CLI_CONFIG.PREFIJO_ID + "-" + String(1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_CLIENTE");
  if (idxId === -1) throw new Error("No se encontró ID_CLIENTE en CLI_MAESTRO");
  
  const registros = hoja.getRange(2, idxId + 1, ultimaFila - 1, 1).getDisplayValues().flat();
  let numeroMayor = 0;
  registros.forEach(id => {
    const textoID = String(id || "").trim();
    if (textoID.startsWith(CLI_CONFIG.PREFIJO_ID + "-")) {
      const numero = parseInt(textoID.replace(CLI_CONFIG.PREFIJO_ID + "-", ""), 10);
      if (!isNaN(numero) && numero > numeroMayor) numeroMayor = numero;
    }
  });
  
  return CLI_CONFIG.PREFIJO_ID + "-" + String(numeroMayor + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

/**
 * Generador de ID seguro para el cliente
 */
function CLI_GENERAR_ID(tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "CREAR");
  }
  return CLI_OBTENER_SIGUIENTE_ID();
}

/**
 * Guarda un nuevo cliente en el sistema
 */
function CLI_GUARDAR_CLIENTE(datos, tokenSesion) {
  if (datos) {
    if (datos.NUMERO_DOCUMENTO === undefined && datos.NIT_CC !== undefined) datos.NUMERO_DOCUMENTO = datos.NIT_CC;
    if (datos.DIGITO_VERIFICACION === undefined && datos.DV !== undefined) datos.DIGITO_VERIFICACION = datos.DV;
    if (datos.ID_CLIENTE === undefined && datos.ID_PROVEEDOR !== undefined) datos.ID_CLIENTE = datos.ID_PROVEEDOR;
    if (datos.ESTADO_CLIENTE === undefined && datos.ESTADO !== undefined) datos.ESTADO_CLIENTE = datos.ESTADO;
  }
  
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  
  if (!datos) throw new Error("Datos no proporcionados.");
  
  // Validar NIT/CC duplicado
  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un cliente registrado.");
  }
  
  // Calcular DV si es NIT
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }
  
  // Asegurar que exista Responsabilidades Fiscales de la DIAN. Por defecto R-99-PN si no se especifica.
  if (!datos.RESPONSABILIDAD_FISCAL || String(datos.RESPONSABILIDAD_FISCAL).trim() === "") {
    datos.RESPONSABILIDAD_FISCAL = "R-99-PN";
  }
  
  // Mapear Siigo y Alegra por defecto
  datos.ID_SIIGO = datos.ID_SIIGO || "";
  datos.ID_ALEGRA = datos.ID_ALEGRA || "";
  
  // Si es persona natural, componer razón social
  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    datos.RAZON_SOCIAL = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && n.trim() !== "").join(" ");
  }
  
  const idCliente = CLI_OBTENER_SIGUIENTE_ID();
  datos.ID_CLIENTE = idCliente;
  
  const ahora = new Date();
  datos.FECHA_CREACION = ahora;
  datos.FECHA_ACTUALIZACION = ahora;
  datos.USUARIO_CREACION = usuarioEjecutor;
  datos.USUARIO_ACTUALIZACION = usuarioEjecutor;
  datos.ESTADO_CLIENTE = datos.ESTADO_CLIENTE || "ACTIVO";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  
  // Convertir objeto a fila ordenado
  const filaNueva = encabezados.map(campo => datos[campo] !== undefined ? datos[campo] : "");
  hoja.appendRow(filaNueva);
  
  // Registrar historial
  CLI_REGISTRAR_HISTORIAL(idCliente, "CREACION", "", "", "Cliente creado correctamente.", usuarioEjecutor);
  
  // Registrar Auditoría Global
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: datos.USUARIO_CREACION,
    USUARIO: usuarioEjecutor,
    MODULO: "CLIENTES",
    SUBMODULO: "MAESTRO",
    ACCION: "CREAR",
    TIPO_REGISTRO: "CLIENTE",
    ID_REGISTRO: idCliente,
    DESCRIPCION: "Cliente " + datos.RAZON_SOCIAL + " (" + idCliente + ") creado con éxito.",
    RESULTADO: "EXITOSO"
  });
  
  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_CLIENTE: idCliente, CLIENTE: datos });
}

/**
 * Actualiza un cliente existente
 */
function CLI_ACTUALIZAR_CLIENTE(datos, tokenSesion) {
  if (datos) {
    if (datos.NUMERO_DOCUMENTO === undefined && datos.NIT_CC !== undefined) datos.NUMERO_DOCUMENTO = datos.NIT_CC;
    if (datos.DIGITO_VERIFICACION === undefined && datos.DV !== undefined) datos.DIGITO_VERIFICACION = datos.DV;
    if (datos.ID_CLIENTE === undefined && datos.ID_PROVEEDOR !== undefined) datos.ID_CLIENTE = datos.ID_PROVEEDOR;
    if (datos.ESTADO_CLIENTE === undefined && datos.ESTADO !== undefined) datos.ESTADO_CLIENTE = datos.ESTADO;
  }
  
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  
  if (!datos || !datos.ID_CLIENTE) throw new Error("ID_CLIENTE es requerido para actualizar.");
  
  // Validar NIT/CC duplicado excluyendo este ID
  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, datos.ID_CLIENTE)) {
    throw new Error("El NIT/CC ingresado ya pertenece a otro cliente registrado.");
  }
  
  // Calcular DV si es NIT
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }
  
  // Asegurar responsabilidades DIAN
  if (!datos.RESPONSABILIDAD_FISCAL || String(datos.RESPONSABILIDAD_FISCAL).trim() === "") {
    datos.RESPONSABILIDAD_FISCAL = "R-99-PN";
  }
  
  // Si es persona natural, componer razón social
  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    datos.RAZON_SOCIAL = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && n.trim() !== "").join(" ");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  
  const idxId = encabezados.indexOf("ID_CLIENTE");
  let filaModificar = -1;
  
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxId]).trim() === String(datos.ID_CLIENTE).trim()) {
      filaModificar = i + 2;
      break;
    }
  }
  
  if (filaModificar === -1) throw new Error("Cliente " + datos.ID_CLIENTE + " no encontrado.");
  
  const clienteActualValores = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const clienteActual = {};
  encabezados.forEach((h, index) => {
    clienteActual[h] = clienteActualValores[index];
  });
  
  // Fusionar datos respetando campos que no se envíen
  const clienteActualizado = Object.assign({}, clienteActual, datos);
  clienteActualizado.FECHA_ACTUALIZACION = new Date();
  clienteActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;
  
  const filaActualizada = encabezados.map(campo => clienteActualizado[campo] !== undefined ? clienteActualizado[campo] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaActualizada]);
  
  // Buscar diferencias para registrar en historial
  encabezados.forEach(campo => {
    if (["FECHA_ACTUALIZACION", "USUARIO_ACTUALIZACION"].includes(campo)) return;
    const valAnt = String(clienteActual[campo] || "").trim();
    const valNue = String(clienteActualizado[campo] || "").trim();
    if (valAnt !== valNue) {
      CLI_REGISTRAR_HISTORIAL(datos.ID_CLIENTE, "MODIFICACION", campo, valAnt, "Campo modificado: " + campo, usuarioEjecutor, valNue);
    }
  });
  
  // Registrar Auditoría Global
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: usuarioEjecutor,
    USUARIO: usuarioEjecutor,
    MODULO: "CLIENTES",
    SUBMODULO: "MAESTRO",
    ACCION: "EDITAR",
    TIPO_REGISTRO: "CLIENTE",
    ID_REGISTRO: datos.ID_CLIENTE,
    DESCRIPCION: "Cliente " + clienteActualizado.RAZON_SOCIAL + " (" + datos.ID_CLIENTE + ") actualizado con éxito.",
    RESULTADO: "EXITOSO"
  });
  
  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_CLIENTE: datos.ID_CLIENTE, CLIENTE: clienteActualizado });
}

/**
 * Busca un tercero en Clientes por NIT, Razón Social o ID_CLIENTE
 */
function CLI_BUSCAR_CLIENTE(criterio) {
  if (!criterio) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxDoc = encabezados.indexOf("NUMERO_DOCUMENTO");
  const idxRazon = encabezados.indexOf("RAZON_SOCIAL");
  
  const criterioNormalizado = String(criterio).trim().toUpperCase();
  const criterioSoloNumeros = criterioNormalizado.replace(/\D/g, "");
  
  const filaEncontrada = registros.find(fila => {
    const valId = String(fila[idxId] || "").trim().toUpperCase();
    const valDoc = String(fila[idxDoc] || "").trim().toUpperCase();
    const valDocSoloNumeros = valDoc.replace(/\D/g, "");
    const valRazon = String(fila[idxRazon] || "").trim().toUpperCase();
    
    return valId === criterioNormalizado || 
           (criterioSoloNumeros !== "" && valDocSoloNumeros === criterioSoloNumeros) || 
           valRazon.includes(criterioNormalizado);
  });
  
  return filaEncontrada ? SEG_SANITIZAR_PARA_CLIENTE(CLI_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada)) : null;
}

function CLI_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const objeto = {};
  encabezados.forEach((campo, indice) => {
    objeto[campo] = fila[indice] !== undefined ? fila[indice] : "";
  });
  return objeto;
}

/**
 * Obtiene un cliente por su ID de forma directa
 */
function CLI_OBTENER_CLIENTE(idCliente, tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "VER");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  
  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  const idxId = encabezados.indexOf("ID_CLIENTE");
  
  const fila = registros.find(f => String(f[idxId]).trim() === String(idCliente).trim());
  return fila ? SEG_SANITIZAR_PARA_CLIENTE(CLI_CONVERTIR_FILA_OBJETO(encabezados, fila)) : null;
}

/**
 * Retorna todos los clientes registrados
 */
function CLI_LISTAR_CLIENTES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "VER");
    const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
    const registros = CLI_OBTENER_REGISTROS(CLI_CONFIG.HOJA_MAESTRO);
    const lista = registros.map(r => CLI_CONVERTIR_FILA_OBJETO(encabezados, r));
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de clientes obtenida con éxito."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar clientes: " + error.toString() };
  }
}

/**
 * Registra historial de cambios del cliente
 */
function CLI_REGISTRAR_HISTORIAL(idCliente, tipoEvento, campo, valAnt, obs, usuario, valNue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
    if (!hoja) return;
    
    const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_HISTORIAL);
    const idHistorial = "HIS-" + String(new Date().getTime()) + "-" + String(Math.floor(Math.random() * 1000));
    const ahora = new Date();
    
    const hist = {
      ID_HISTORIAL: idHistorial,
      ID_CLIENTE: idCliente,
      TIPO_EVENTO: tipoEvento,
      FECHA_HORA: ahora,
      USUARIO: usuario || "SISTEMA",
      ACCION: tipoEvento,
      CAMPO_MODIFICADO: campo || "",
      VALOR_ANTERIOR: valAnt || "",
      VALOR_NUEVO: valNue || "",
      MOTIVO_ORIGEN: "MAESTRO_CLIENTES",
      MODULO_ORIGEN: "CLIENTES",
      ID_REGISTRO_ORIGEN: idCliente,
      ESTADO_EVENTO: "ACTIVO",
      OBSERVACIONES: obs || ""
    };
    
    const filaNueva = encabezados.map(c => hist[c] !== undefined ? hist[c] : "");
    hoja.appendRow(filaNueva);
  } catch (e) {
    console.error("Error al escribir CLI_HISTORIAL: " + e.toString());
  }
}

/**
 * Determina dinámicamente si al cliente se le aplica retenciones en transacciones.
 * Analiza la columna RESPONSABILIDAD_FISCAL (ej: 'O-15' Autoretentor, 'O-23' Agente retención, etc.)
 */
function CLI_DEBE_APLICAR_RETENCION(idCliente) {
  const cliente = CLI_OBTENER_CLIENTE(idCliente);
  if (!cliente) return { RETENCION_FUENTE: true, RETENCION_IVA: false, MOTIVO: "Cliente no encontrado. Se aplican reglas generales." };
  
  const responsabilidades = String(cliente.RESPONSABILIDAD_FISCAL || "").toUpperCase();
  
  // Si el cliente es Gran Contribuyente (O-13) o Agente de Retención de IVA (O-23), nos aplicará retención de IVA
  const esGranContribuyente = responsabilidades.includes("O-13");
  const esAgenteRetencionIva = responsabilidades.includes("O-23");
  const esAutorretenedor = responsabilidades.includes("O-15");
  
  return {
    RETENCION_FUENTE: !esAutorretenedor, // Si es autorretenedor, no se le retiene fuente en la compra
    RETENCION_IVA: esGranContribuyente || esAgenteRetencionIva,
    MOTIVO: "Evaluado según responsabilidades fiscales DIAN: " + responsabilidades
  };
}
