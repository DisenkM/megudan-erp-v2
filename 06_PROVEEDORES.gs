/**************************************************************
* 06_PROVEEDORES.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el catálogo y operaciones (CRUD) de Proveedores (PROV_MAESTRO).
* - Control de NITs colombianos y cálculo de DV reutilizando clientes.
* - Soportar la arquitectura de seguridad dual.
* - Mapear identificadores de Siigo y Alegra desde el día uno.
* - Analizar responsabilidades tributarias para retenciones.
**************************************************************/

const PROV_CONFIG = {
  HOJA_MAESTRO: "PROV_MAESTRO",
  HOJA_HISTORIAL: "PROV_HISTORIAL",
  PREFIJO_ID: "PROV",
  DIGITOS_ID: 6
};

/**
 * Obtiene la hoja física de proveedores
 */
function PROV_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No existe la hoja física '" + nombreHoja + "' en la base de datos.");
  return hoja;
}

/**
 * Obtiene encabezados de proveedores
 */
function PROV_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = PROV_OBTENER_HOJA(nombreHoja);
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaColumna === 0) throw new Error("La hoja '" + nombreHoja + "' no contiene encabezados.");
  return hoja.getRange(1, 1, 1, ultimaColumna).getDisplayValues()[0].map(h => String(h || "").trim().toUpperCase());
}

/**
 * Obtiene registros de proveedores
 */
function PROV_OBTENER_REGISTROS(nombreHoja) {
  const hoja = PROV_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
}

/**
 * Valida duplicado en proveedores
 */
function PROV_VALIDAR_DUPLICADO(tipoDoc, numDoc, idExcluir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return false;
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  
  const idxId = encabezados.indexOf("ID_PROVEEDOR") !== -1 ? encabezados.indexOf("ID_PROVEEDOR") : encabezados.indexOf("ID_CLIENTE");
  const idxTipoDoc = encabezados.indexOf("TIPO_DOCUMENTO");
  const idxNumDoc = encabezados.indexOf("NIT_CC") !== -1 ? encabezados.indexOf("NIT_CC") : (encabezados.indexOf("NUMERO_DOCUMENTO") !== -1 ? encabezados.indexOf("NUMERO_DOCUMENTO") : encabezados.indexOf("NUMERO"));
  
  if (idxId === -1 || idxTipoDoc === -1 || idxNumDoc === -1) return false;
  
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getDisplayValues();
  return datos.some(fila => {
    if (idExcluir && String(fila[idxId]) === String(idExcluir)) return false;
    return String(fila[idxTipoDoc]).trim().toUpperCase() === String(tipoDoc).trim().toUpperCase() && 
           String(fila[idxNumDoc]).trim().replace(/\D/g, "") === String(numDoc).trim().replace(/\D/g, "");
  });
}

/**
 * Obtiene siguiente consecutivo para el ID de proveedor
 */
function PROV_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) return PROV_CONFIG.PREFIJO_ID + "-" + String(1).padStart(PROV_CONFIG.DIGITOS_ID, "0");
  
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return PROV_CONFIG.PREFIJO_ID + "-" + String(1).padStart(PROV_CONFIG.DIGITOS_ID, "0");
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_PROVEEDOR") !== -1 ? encabezados.indexOf("ID_PROVEEDOR") : encabezados.indexOf("ID_CLIENTE");
  if (idxId === -1) throw new Error("No se pudo identificar la columna ID_PROVEEDOR en PROV_MAESTRO");
  
  const registros = hoja.getRange(2, idxId + 1, ultimaFila - 1, 1).getDisplayValues().flat();
  let numeroMayor = 0;
  registros.forEach(id => {
    const textoID = String(id || "").trim();
    if (textoID.startsWith(PROV_CONFIG.PREFIJO_ID + "-")) {
      const numero = parseInt(textoID.replace(PROV_CONFIG.PREFIJO_ID + "-", ""), 10);
      if (!isNaN(numero) && numero > numeroMayor) numeroMayor = numero;
    }
  });
  
  return PROV_CONFIG.PREFIJO_ID + "-" + String(numeroMayor + 1).padStart(PROV_CONFIG.DIGITOS_ID, "0");
}

/**
 * Generador seguro de ID desde la vista
 */
function PROV_GENERAR_ID(tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "CREAR");
  }
  return PROV_OBTENER_SIGUIENTE_ID();
}

/**
 * Guarda un proveedor en el maestro
 */
function PROV_GUARDAR_PROVEEDOR(datos, tokenSesion) {
  if (datos) {
    if (datos.NIT_CC === undefined && datos.NUMERO_DOCUMENTO !== undefined) datos.NIT_CC = datos.NUMERO_DOCUMENTO;
    if (datos.DV === undefined && datos.DIGITO_VERIFICACION !== undefined) datos.DV = datos.DIGITO_VERIFICACION;
    if (datos.ID_PROVEEDOR === undefined && datos.ID_CLIENTE !== undefined) datos.ID_PROVEEDOR = datos.ID_CLIENTE;
    if (datos.ESTADO === undefined && datos.ESTADO_CLIENTE !== undefined) datos.ESTADO = datos.ESTADO_CLIENTE;
  }
  
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  
  if (!datos) throw new Error("Datos de proveedor no especificados.");
  
  if (PROV_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NIT_CC, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un proveedor registrado.");
  }
  
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DV = CLI_CALCULAR_DV(datos.NIT_CC);
  } else {
    datos.DV = "";
  }
  
  // Soportar campos requeridos y mapeos
  datos.ID_SIIGO = datos.ID_SIIGO || "";
  datos.ID_ALEGRA = datos.ID_ALEGRA || "";
  datos.NOMBRE_COMERCIAL = datos.NOMBRE_COMERCIAL || "";
  datos.NOMBRE_CONTACTO = datos.NOMBRE_CONTACTO || "";
  datos.RESPONSABILIDAD_FISCAL = datos.RESPONSABILIDAD_FISCAL || "R-99-PN";
  
  const idProveedor = PROV_OBTENER_SIGUIENTE_ID();
  datos.ID_PROVEEDOR = idProveedor;
  
  const ahora = new Date();
  datos.FECHA_CREACION = ahora;
  datos.FECHA_MODIFICACION = ahora;
  datos.ESTADO = datos.ESTADO || "ACTIVO";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROV_MAESTRO no encontrada.");
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  
  // Alinear con los campos del formulario si tienen alias
  const filaNueva = encabezados.map(campo => {
    // Al mapear TERCEROS UNIFICADOS el form puede mandar campos CLI_
    let valor = datos[campo];
    if (valor === undefined) {
      if (campo === "CONDICION_PAGO" && datos.FORMA_PAGO !== undefined) valor = datos.FORMA_PAGO;
      if (campo === "CUPO_CREDITO" && datos.LIMITE_CREDITO !== undefined) valor = datos.LIMITE_CREDITO;
      if (campo === "NOMBRE_CONTACTO" && datos.CONTACTO_PRINCIPAL !== undefined) valor = datos.CONTACTO_PRINCIPAL;
    }
    return valor !== undefined ? valor : "";
  });
  
  hoja.appendRow(filaNueva);
  
  // Registrar historial
  PROV_REGISTRAR_HISTORIAL(idProveedor, "CREACION", "", "", "Proveedor creado correctamente.", usuarioEjecutor);
  
  // Registrar Auditoría Global
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: usuarioEjecutor,
    USUARIO: usuarioEjecutor,
    MODULO: "PROVEEDORES",
    SUBMODULO: "MAESTRO",
    ACCION: "CREAR",
    TIPO_REGISTRO: "PROVEEDOR",
    ID_REGISTRO: idProveedor,
    DESCRIPCION: "Proveedor " + datos.RAZON_SOCIAL + " (" + idProveedor + ") creado con éxito.",
    RESULTADO: "EXITOSO"
  });
  
  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_PROVEEDOR: idProveedor, PROVEEDOR: datos });
}

/**
 * Actualiza un proveedor existente
 */
function PROV_ACTUALIZAR_PROVEEDOR(datos, tokenSesion) {
  if (datos) {
    if (datos.NIT_CC === undefined && datos.NUMERO_DOCUMENTO !== undefined) datos.NIT_CC = datos.NUMERO_DOCUMENTO;
    if (datos.DV === undefined && datos.DIGITO_VERIFICACION !== undefined) datos.DV = datos.DIGITO_VERIFICACION;
    if (datos.ID_PROVEEDOR === undefined && datos.ID_CLIENTE !== undefined) datos.ID_PROVEEDOR = datos.ID_CLIENTE;
    if (datos.ESTADO === undefined && datos.ESTADO_CLIENTE !== undefined) datos.ESTADO = datos.ESTADO_CLIENTE;
  }
  
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  
  if (!datos || !datos.ID_PROVEEDOR) throw new Error("ID_PROVEEDOR es requerido para actualizar.");
  
  if (PROV_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NIT_CC, datos.ID_PROVEEDOR)) {
    throw new Error("El NIT/CC ingresado ya pertenece a otro proveedor registrado.");
  }
  
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DV = CLI_CALCULAR_DV(datos.NIT_CC);
  } else {
    datos.DV = "";
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja PROV_MAESTRO no encontrada.");
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  
  const idxId = encabezados.indexOf("ID_PROVEEDOR") !== -1 ? encabezados.indexOf("ID_PROVEEDOR") : encabezados.indexOf("ID_CLIENTE");
  let filaModificar = -1;
  
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxId]).trim() === String(datos.ID_PROVEEDOR).trim()) {
      filaModificar = i + 2;
      break;
    }
  }
  
  if (filaModificar === -1) throw new Error("Proveedor " + datos.ID_PROVEEDOR + " no encontrado.");
  
  const proveedorActualValores = hoja.getRange(filaModificar, 1, 1, encabezados.length).getValues()[0];
  const proveedorActual = {};
  encabezados.forEach((h, index) => {
    proveedorActual[h] = proveedorActualValores[index];
  });
  
  // Mapeo dinámico del form unificado
  const datosMapeados = Object.assign({}, datos);
  if (datosMapeados.FORMA_PAGO !== undefined) datosMapeados.CONDICION_PAGO = datosMapeados.FORMA_PAGO;
  if (datosMapeados.LIMITE_CREDITO !== undefined) datosMapeados.CUPO_CREDITO = datosMapeados.LIMITE_CREDITO;
  if (datosMapeados.CONTACTO_PRINCIPAL !== undefined) datosMapeados.NOMBRE_CONTACTO = datosMapeados.CONTACTO_PRINCIPAL;
  
  const proveedorActualizado = Object.assign({}, proveedorActual, datosMapeados);
  proveedorActualizado.FECHA_MODIFICACION = new Date();
  
  const filaActualizada = encabezados.map(campo => proveedorActualizado[campo] !== undefined ? proveedorActualizado[campo] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaActualizada]);
  
  // Buscar diferencias para registrar en historial
  encabezados.forEach(campo => {
    if (["FECHA_MODIFICACION"].includes(campo)) return;
    const valAnt = String(proveedorActual[campo] || "").trim();
    const valNue = String(proveedorActualizado[campo] || "").trim();
    if (valAnt !== valNue) {
      PROV_REGISTRAR_HISTORIAL(datos.ID_PROVEEDOR, "MODIFICACION", campo, valAnt, "Campo modificado: " + campo, usuarioEjecutor, valNue);
    }
  });
  
  // Registrar Auditoría Global
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: usuarioEjecutor,
    USUARIO: usuarioEjecutor,
    MODULO: "PROVEEDORES",
    SUBMODULO: "MAESTRO",
    ACCION: "EDITAR",
    TIPO_REGISTRO: "PROVEEDOR",
    ID_REGISTRO: datos.ID_PROVEEDOR,
    DESCRIPCION: "Proveedor " + proveedorActualizado.RAZON_SOCIAL + " (" + datos.ID_PROVEEDOR + ") actualizado con éxito.",
    RESULTADO: "EXITOSO"
  });
  
  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_PROVEEDOR: datos.ID_PROVEEDOR, PROVEEDOR: proveedorActualizado });
}

/**
 * Busca un proveedor por nit, id o razon social
 */
function PROV_BUSCAR_PROVEEDOR(criterio) {
  if (!criterio) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  
  const idxId = encabezados.indexOf("ID_PROVEEDOR") !== -1 ? encabezados.indexOf("ID_PROVEEDOR") : encabezados.indexOf("ID_CLIENTE");
  const idxDoc = encabezados.indexOf("NIT_CC") !== -1 ? encabezados.indexOf("NIT_CC") : (encabezados.indexOf("NUMERO_DOCUMENTO") !== -1 ? encabezados.indexOf("NUMERO_DOCUMENTO") : encabezados.indexOf("NUMERO"));
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
           valRazon.includes(criterioNormalized);
  });
  
  return filaEncontrada ? SEG_SANITIZAR_PARA_CLIENTE(PROV_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada)) : null;
}

function PROV_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const objeto = {};
  encabezados.forEach((campo, indice) => {
    objeto[campo] = fila[indice] !== undefined ? fila[indice] : "";
  });
  return objeto;
}

/**
 * Obtiene proveedor por ID
 */
function PROV_OBTENER_PROVEEDOR(idProveedor, tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "VER");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;
  
  const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  const idxId = encabezados.indexOf("ID_PROVEEDOR") !== -1 ? encabezados.indexOf("ID_PROVEEDOR") : encabezados.indexOf("ID_CLIENTE");
  
  const fila = registros.find(f => String(f[idxId]).trim() === String(idProveedor).trim());
  return fila ? SEG_SANITIZAR_PARA_CLIENTE(PROV_CONVERTIR_FILA_OBJETO(encabezados, fila)) : null;
}

/**
 * Retorna todos los proveedores registrados
 */
function PROV_LISTAR_PROVEEDORES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "PROVEEDORES", "VER");
    const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_MAESTRO);
    const registros = PROV_OBTENER_REGISTROS(PROV_CONFIG.HOJA_MAESTRO);
    const lista = registros.map(r => PROV_CONVERTIR_FILA_OBJETO(encabezados, r));
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de proveedores obtenida con éxito."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar proveedores: " + error.toString() };
  }
}

/**
 * Historial de cambios de proveedores
 */
function PROV_REGISTRAR_HISTORIAL(idProveedor, tipoEvento, campo, valAnt, obs, usuario, valNue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(PROV_CONFIG.HOJA_HISTORIAL);
    if (!hoja) return;
    
    const encabezados = PROV_OBTENER_ENCABEZADOS(PROV_CONFIG.HOJA_HISTORIAL);
    const idHistorial = "HIS-" + String(new Date().getTime()) + "-" + String(Math.floor(Math.random() * 1000));
    const ahora = new Date();
    
    const hist = {
      ID_HISTORIAL: idHistorial,
      ID_PROVEEDOR: idProveedor,
      FECHA: ahora,
      HORA: Utilities.formatDate(ahora, Session.getScriptTimeZone() || "America/Bogota", "HH:mm:ss"),
      USUARIO: usuario || "SISTEMA",
      ACCION: tipoEvento,
      CAMPO: campo || "",
      VALOR_ANTERIOR: valAnt || "",
      VALOR_NUEVO: valNue || "",
      OBSERVACION: obs || ""
    };
    
    const filaNueva = encabezados.map(c => hist[c] !== undefined ? hist[c] : "");
    hoja.appendRow(filaNueva);
  } catch (e) {
    console.error("Error al escribir PROV_HISTORIAL: " + e.toString());
  }
}

/**
 * Determina dinámicamente si al proveedor se le aplica retenciones en compras.
 * Analiza la columna RESPONSABILIDAD_FISCAL (ej: 'O-15' Autoretentor, 'O-23' Agente retención, etc.)
 */
function PROV_DEBE_APLICAR_RETENCION(idProveedor) {
  const prov = PROV_OBTENER_PROVEEDOR(idProveedor);
  if (!prov) return { RETENCION_FUENTE: true, RETENCION_IVA: false, MOTIVO: "Proveedor no encontrado. Se aplican reglas generales." };
  
  const responsabilidades = String(prov.RESPONSABILIDAD_FISCAL || "").toUpperCase();
  const esAutorretenedor = responsabilidades.includes("O-15");
  const esSimplificado = responsabilidades.includes("O-47") || responsabilidades.includes("R-99-PN");
  
  return {
    RETENCION_FUENTE: !esAutorretenedor, // Si es autorretenedor, no le hacemos retención en la fuente (él se auto-retiene)
    RETENCION_IVA: esSimplificado, // Si es simplificado, podemos tener retención de IVA según normativas (Régimen Simple / No responsable)
    MOTIVO: "Evaluado según responsabilidades fiscales del proveedor: " + responsabilidades
  };
}
