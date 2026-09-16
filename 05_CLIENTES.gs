// (VERSIÓN 5.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 05_CLIENTES.gs (VERSIÓN 5.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el ciclo de vida (CRUD) de Clientes (CLI_MAESTRO).
* - Proteger accesos bajo la arquitectura de seguridad dual.
* - Calcular dígitos de verificación DIAN (DV) y registrar historiales.
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
 * Algoritmo oficial DIAN para el cálculo del Dígito de Verificación (DV)
 */
function CLI_CALCULAR_DV(nit) {
  if (!nit) return "0";
  const strNit = String(nit).replace(/\D/g, "");
  if (!strNit) return "0";
  const vpri = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  let z = strNit.length;
  let y = 0;
  for (let i = 0; i < z; i++) {
    y += parseInt(strNit.charAt(z - 1 - i), 10) * vpri[i];
  }
  let x = y % 11;
  return x > 1 ? String(11 - x) : String(x);
}

function CLI_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  return hoja;
}

function CLI_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = CLI_OBTENER_HOJA(nombreHoja);
  const uc = hoja.getLastColumn();
  if (uc === 0) return [];
  return hoja.getRange(1, 1, 1, uc).getValues()[0].map(h => String(h || "").trim().toUpperCase());
}

function CLI_OBTENER_REGISTROS(nombreHoja) {
  const hoja = CLI_OBTENER_HOJA(nombreHoja);
  const ulFila = hoja.getLastRow();
  if (ulFila < 2) return [];
  return hoja.getRange(2, 1, ulFila - 1, hoja.getLastColumn()).getValues();
}

function CLI_VALIDAR_DUPLICADO(tipoDoc, numDoc, idExcluir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return false;

  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxTipoDoc = encabezados.indexOf("TIPO_DOCUMENTO");
  const idxNumDoc = encabezados.indexOf("NUMERO_DOCUMENTO") !== -1 ? encabezados.indexOf("NUMERO_DOCUMENTO") : (encabezados.indexOf("NIT_CC") !== -1 ? encabezados.indexOf("NIT_CC") : (encabezados.indexOf("NUMERO") !== -1 ? encabezados.indexOf("NUMERO") : -1));

  if (idxId === -1 || idxTipoDoc === -1 || idxNumDoc === -1) return false;

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();

  return datos.some(fila => {
    if (idExcluir && String(fila[idxId]).trim() === String(idExcluir).trim()) return false;
    return String(fila[idxTipoDoc]).trim().toUpperCase() === String(tipoDoc).trim().toUpperCase() &&
           String(fila[idxNumDoc]).trim().replace(/\D/g, "") === String(numDoc).trim().replace(/\D/g, "");
  });
}

function CLI_OBTENER_SIGUIENTE_ID() {
  const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_MAESTRO);
  const ulFila = hoja.getLastRow();
  if (ulFila < 2) return CLI_CONFIG.PREFIJO_ID + "-" + String(1).padStart(CLI_CONFIG.DIGITOS_ID, "0");

  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_CLIENTE");
  if (idxId === -1) return CLI_CONFIG.PREFIJO_ID + "-" + String(ulFila).padStart(CLI_CONFIG.DIGITOS_ID, "0");

  const registros = hoja.getRange(2, idxId + 1, ulFila - 1, 1).getValues().flat();
  let maxNum = 0;
  registros.forEach(id => {
    const str = String(id || "").trim();
    if (str.startsWith(CLI_CONFIG.PREFIJO_ID + "-")) {
      const num = parseInt(str.replace(CLI_CONFIG.PREFIJO_ID + "-", ""), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  });

  return CLI_CONFIG.PREFIJO_ID + "-" + String(maxNum + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

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

  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un cliente registrado.");
  }

  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }

  if (!datos.RESPONSABILIDAD_FISCAL || String(datos.RESPONSABILIDAD_FISCAL).trim() === "") {
    datos.RESPONSABILIDAD_FISCAL = "R-99-PN";
  }

  datos.ID_SIIGO = datos.ID_SIIGO || "";
  datos.ID_ALEGRA = datos.ID_ALEGRA || "";

  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    const nomComp = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && String(n).trim() !== "").join(" ");
    if (nomComp) datos.RAZON_SOCIAL = nomComp;
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
  const filaNueva = encabezados.map(campo => datos[campo] !== undefined ? datos[campo] : "");
  hoja.appendRow(filaNueva);

  CLI_REGISTRAR_HISTORIAL(idCliente, "CREACION", "", "", "Cliente creado correctamente.", usuarioEjecutor);

  if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: usuarioEjecutor,
      USUARIO: usuarioEjecutor,
      MODULO: "CLIENTES",
      SUBMODULO: "MAESTRO",
      ACCION: "CREAR",
      TIPO_REGISTRO: "CLIENTE",
      ID_REGISTRO: idCliente,
      DESCRIPCION: "Cliente " + datos.RAZON_SOCIAL + " (" + idCliente + ") creado con éxito.",
      RESULTADO: "EXITOSO"
    });
  }

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_CLIENTE: idCliente, idCliente: idCliente, CLIENTE: datos, MENSAJE: "¡Cliente " + idCliente + " guardado exitosamente!" });
}

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

  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, datos.ID_CLIENTE)) {
    throw new Error("El NIT/CC ingresado ya pertenece a otro cliente registrado.");
  }

  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }

  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    const nomComp = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && String(n).trim() !== "").join(" ");
    if (nomComp) datos.RAZON_SOCIAL = nomComp;
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

  const clienteActualizado = Object.assign({}, clienteActual, datos);
  clienteActualizado.FECHA_ACTUALIZACION = new Date();
  clienteActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;

  const filaActualizada = encabezados.map(campo => clienteActualizado[campo] !== undefined ? clienteActualizado[campo] : "");
  hoja.getRange(filaModificar, 1, 1, encabezados.length).setValues([filaActualizada]);

  CLI_REGISTRAR_HISTORIAL(datos.ID_CLIENTE, "MODIFICACION", "", "", "Cliente actualizado de forma completa.", usuarioEjecutor);

  if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
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
  }

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_CLIENTE: datos.ID_CLIENTE, CLIENTE: clienteActualizado, MENSAJE: "¡Cliente " + datos.ID_CLIENTE + " actualizado con éxito!" });
}

function CLI_ELIMINAR_CLIENTE(idCliente, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "ELIMINAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!idCliente) throw new Error("ID_CLIENTE es requerido para inactivar.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");

  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxEst = encabezados.indexOf("ESTADO_CLIENTE") !== -1 ? encabezados.indexOf("ESTADO_CLIENTE") : encabezados.indexOf("ESTADO");

  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxId]).trim() === String(idCliente).trim()) {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar === -1) throw new Error("Cliente " + idCliente + " no encontrado.");

  if (idxEst !== -1) {
    hoja.getRange(filaModificar, idxEst + 1).setValue("INACTIVO");
  }

  CLI_REGISTRAR_HISTORIAL(idCliente, "INACTIVACION", "ESTADO_CLIENTE", "ACTIVO", "Cliente marcado como INACTIVO de forma lógica.", usuarioEjecutor, "INACTIVO");

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_CLIENTE: idCliente, MENSAJE: "Cliente " + idCliente + " inactivado correctamente." });
}

function CLI_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const objeto = {};
  encabezados.forEach((campo, indice) => {
    let val = fila[indice] !== undefined && fila[indice] !== null ? fila[indice] : "";
    if (["NUMERO_DOCUMENTO", "NIT_CC", "DIGITO_VERIFICACION", "DV", "TELEFONO", "CELULAR", "PLAZO_PAGO_DIAS"].includes(campo)) {
      val = String(val).trim();
    }
    objeto[campo] = val;
  });
  if (objeto.NUMERO_DOCUMENTO === undefined || objeto.NUMERO_DOCUMENTO === "") {
    if (objeto.NIT_CC) objeto.NUMERO_DOCUMENTO = String(objeto.NIT_CC).trim();
    else if (objeto.NUMERO) objeto.NUMERO_DOCUMENTO = String(objeto.NUMERO).trim();
  }
  if (objeto.NIT_CC === undefined || objeto.NIT_CC === "") {
    if (objeto.NUMERO_DOCUMENTO) objeto.NIT_CC = String(objeto.NUMERO_DOCUMENTO).trim();
  }
  if (objeto.DIGITO_VERIFICACION === undefined || objeto.DIGITO_VERIFICACION === "") {
    if (objeto.DV) objeto.DIGITO_VERIFICACION = String(objeto.DV).trim();
  }
  if (objeto.DV === undefined || objeto.DV === "") {
    if (objeto.DIGITO_VERIFICACION) objeto.DV = String(objeto.DIGITO_VERIFICACION).trim();
  }
  return objeto;
}

function CLI_BUSCAR_CLIENTE(criterio, tokenSesion) {
  if (!criterio) return null;
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "VER");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return null;

  const encabezados = CLI_OBTENER_ENCABEZADOS(CLI_CONFIG.HOJA_MAESTRO);
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();

  const idxId = encabezados.indexOf("ID_CLIENTE");
  const idxDoc = encabezados.indexOf("NUMERO_DOCUMENTO") !== -1 ? encabezados.indexOf("NUMERO_DOCUMENTO") : (encabezados.indexOf("NIT_CC") !== -1 ? encabezados.indexOf("NIT_CC") : (encabezados.indexOf("NUMERO") !== -1 ? encabezados.indexOf("NUMERO") : -1));
  const idxRazon = encabezados.indexOf("RAZON_SOCIAL");

  const criterioNormalizado = String(criterio).trim().toUpperCase();

  let filaEncontrada = registros.find(fila => String(fila[idxId] || "").trim().toUpperCase() === criterioNormalizado);

  if (!filaEncontrada) {
    const criterioSoloNumeros = criterioNormalizado.replace(/\D/g, "");
    filaEncontrada = registros.find(fila => {
      const valDoc = String(fila[idxDoc] || "").trim().toUpperCase();
      const valDocSoloNumeros = valDoc.replace(/\D/g, "");
      const valRazon = String(fila[idxRazon] || "").trim().toUpperCase();

      return (criterioSoloNumeros !== "" && valDocSoloNumeros === criterioSoloNumeros) ||
             (criterioNormalizado.length >= 3 && valRazon.includes(criterioNormalizado));
    });
  }

  return filaEncontrada ? SEG_SANITIZAR_PARA_CLIENTE(CLI_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada)) : null;
}

function CLI_OBTENER_CLIENTE(idCliente, tokenSesion) {
  return CLI_BUSCAR_CLIENTE(idCliente, tokenSesion);
}

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

function CLI_LISTAR_CLIENTES_WEB(tokenSesion) {
  return CLI_LISTAR_CLIENTES(tokenSesion);
}

function CLI_REGISTRAR_HISTORIAL(idCliente, tipoEvento, campo, valAnt, obs, usuario, valNue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
    if (!hoja) {
      hoja = ss.insertSheet(CLI_CONFIG.HOJA_HISTORIAL);
      hoja.appendRow(["ID_HISTORIAL", "ID_CLIENTE", "TIPO_EVENTO", "FECHA_HORA", "USUARIO", "ACCION", "CAMPO_MODIFICADO", "VALOR_ANTERIOR", "VALOR_NUEVO", "MOTIVO_ORIGEN", "MODULO_ORIGEN", "ID_REGISTRO_ORIGEN", "ESTADO_EVENTO", "OBSERVACIONES"]);
    }
    const idHist = "HIS-" + new Date().getTime();
    hoja.appendRow([idHist, idCliente, tipoEvento, new Date(), usuario || "SISTEMA", tipoEvento, campo || "", valAnt || "", valNue || "", "MAESTRO_CLIENTES", "CLIENTES", idCliente, "ACTIVO", obs || ""]);
  } catch (e) {}
}

function CLI_DEBE_APLICAR_RETENCION(idCliente) {
  const cliente = CLI_OBTENER_CLIENTE(idCliente);
  if (!cliente) return { RETENCION_FUENTE: true, RETENCION_IVA: false, MOTIVO: "Cliente general." };
  const responsabilidades = String(cliente.RESPONSABILIDAD_FISCAL || "").toUpperCase();
  return {
    RETENCION_FUENTE: !responsabilidades.includes("O-15"),
    RETENCION_IVA: responsabilidades.includes("O-13"),
    MOTIVO: "Evaluado según responsabilidades DIAN."
  };
}

