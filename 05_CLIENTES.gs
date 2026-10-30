// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 05_CLIENTES.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Módulo Core de Clientes, Proveedores, Terceros y Cartera / Cuentas por Cobrar.
* - Persistencia relacional en las hojas CLI_MAESTRO, PROV_MAESTRO y CAR_CUENTAS_COBRAR.
* - Métodos RPC asíncronos para la interfaz F1_CLI_FORM y los demás módulos (Ventas, Obras, Documentos).
**************************************************************/

const CLI_CONFIG = {
  HOJA_CLIENTES: "CLI_MAESTRO",
  HOJA_PROVEEDORES: "PROV_MAESTRO",
  HOJA_CARTERA: "CAR_CUENTAS_COBRAR",
  PREFIJO_CLIENTE: "CLI",
  PREFIJO_PROVEEDOR: "PROV",
  PREFIJO_CARTERA: "CXC",
  DIGITOS_ID: 6
};

function CLI_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
    if (nombreHoja === CLI_CONFIG.HOJA_CLIENTES) {
      hoja.appendRow([
        "ID_CLIENTE", "TIPO_TERCERO", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO",
        "DIGITO_VERIFICACION", "RAZON_SOCIAL", "PRIMER_NOMBRE", "SEGUNDO_NOMBRE", "PRIMER_APELLIDO",
        "SEGUNDO_APELLIDO", "NOMBRE_COMERCIAL", "DIRECCION", "CIUDAD", "DEPARTAMENTO",
        "PAIS", "CELULAR", "CORREO", "RESPONSABILIDAD_FISCAL", "FORMA_PAGO",
        "PLAZO_PAGO_DIAS", "LIMITE_CREDITO", "ESTADO_CLIENTE", "FECHA_CREACION", "USUARIO_CREACION"
      ]);
      hoja.getRange(1, 1, 1, 25).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    } else if (nombreHoja === CLI_CONFIG.HOJA_PROVEEDORES) {
      hoja.appendRow([
        "ID_PROVEEDOR", "TIPO_TERCERO", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO",
        "DIGITO_VERIFICACION", "RAZON_SOCIAL", "PRIMER_NOMBRE", "SEGUNDO_NOMBRE", "PRIMER_APELLIDO",
        "SEGUNDO_APELLIDO", "NOMBRE_COMERCIAL", "DIRECCION", "CIUDAD", "DEPARTAMENTO",
        "PAIS", "CELULAR", "CORREO", "RESPONSABILIDAD_FISCAL", "FORMA_PAGO",
        "PLAZO_PAGO_DIAS", "LIMITE_CREDITO", "ESTADO", "FECHA_CREACION", "USUARIO_CREACION"
      ]);
      hoja.getRange(1, 1, 1, 25).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    } else if (nombreHoja === CLI_CONFIG.HOJA_CARTERA) {
      hoja.appendRow([
        "ID_CARTERA", "ID_VENTA", "ID_CLIENTE", "FECHA_EMISION", "FECHA_VENCIMIENTO",
        "VALOR_DOCUMENTO", "ABONOS", "SALDO", "ESTADO"
      ]);
      hoja.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  }
  return hoja;
}

/**
 * RPC: Listar todos los clientes registrados
 */
function CLI_LISTAR_CLIENTES_WEB(tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "VER");
    }
    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_CLIENTES);
    const ulFila = hoja.getLastRow();
    if (ulFila < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay clientes registrados." };
    }
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ulFila - 1, enc.length).getValues();

    const lista = reg.map(fila => {
      const obj = {};
      enc.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
      if (!obj.TIPO_TERCERO) obj.TIPO_TERCERO = "CLIENTE";
      return obj;
    });

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(lista) : JSON.parse(JSON.stringify(lista)),
      MENSAJE: "Clientes listados correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar clientes: " + error.toString() };
  }
}

function CLI_LISTAR_CLIENTES(tokenSesion) {
  return CLI_LISTAR_CLIENTES_WEB(tokenSesion);
}

/**
 * RPC: Guardar nuevo cliente comercial
 */
function CLI_GUARDAR_CLIENTE(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "CREAR");
    }

    if (!datos || (!datos.NUMERO_DOCUMENTO && !datos.NIT_CC)) {
      throw new Error("El número de documento o NIT es obligatorio.");
    }

    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_CLIENTES);
    const idCli = datos.ID_CLIENTE || (CLI_CONFIG.PREFIJO_CLIENTE + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CLI_CONFIG.DIGITOS_ID, "0"));
    const ahora = new Date();
    const usuarioEjecutor = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "OPERADOR") : "OPERADOR";

    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const docNum = datos.NUMERO_DOCUMENTO || datos.NIT_CC || "";
    const dvVal = datos.DIGITO_VERIFICACION || datos.DV || "";

    const filaNueva = enc.map(col => {
      switch (col) {
        case "ID_CLIENTE": return idCli;
        case "TIPO_TERCERO": return datos.TIPO_TERCERO || "CLIENTE";
        case "TIPO_PERSONA": return datos.TIPO_PERSONA || "PERSONA_JURIDICA";
        case "TIPO_DOCUMENTO": return datos.TIPO_DOCUMENTO || "NIT";
        case "NUMERO_DOCUMENTO": return docNum;
        case "DIGITO_VERIFICACION": return dvVal;
        case "RAZON_SOCIAL": return String(datos.RAZON_SOCIAL || datos.NOMBRE_COMERCIAL || "CLIENTE SIN NOMBRE").trim().toUpperCase();
        case "PRIMER_NOMBRE": return datos.PRIMER_NOMBRE || "";
        case "SEGUNDO_NOMBRE": return datos.SEGUNDO_NOMBRE || "";
        case "PRIMER_APELLIDO": return datos.PRIMER_APELLIDO || "";
        case "SEGUNDO_APELLIDO": return datos.SEGUNDO_APELLIDO || "";
        case "NOMBRE_COMERCIAL": return datos.NOMBRE_COMERCIAL || datos.RAZON_SOCIAL || "";
        case "DIRECCION": return datos.DIRECCION || "";
        case "CIUDAD": return datos.CIUDAD || "Pitalito";
        case "DEPARTAMENTO": return datos.DEPARTAMENTO || "Huila";
        case "PAIS": return datos.PAIS || "COLOMBIA";
        case "CELULAR": return datos.CELULAR || datos.TELEFONO || "";
        case "CORREO": return datos.CORREO || datos.EMAIL || "";
        case "RESPONSABILIDAD_FISCAL": return datos.RESPONSABILIDAD_FISCAL || "48";
        case "FORMA_PAGO": return datos.FORMA_PAGO || "CONTADO";
        case "PLAZO_PAGO_DIAS": return Number(datos.PLAZO_PAGO_DIAS || 30);
        case "LIMITE_CREDITO": return Number(datos.LIMITE_CREDITO || datos.CUPO_CREDITO || 0);
        case "ESTADO_CLIENTE": return datos.ESTADO_CLIENTE || datos.ESTADO || "ACTIVO";
        case "FECHA_CREACION": return ahora;
        case "USUARIO_CREACION": return usuarioEjecutor;
        default: return "";
      }
    });

    hoja.appendRow(filaNueva);

    return {
      EXITO: true,
      ID_CLIENTE: idCli,
      MENSAJE: "¡Cliente '" + (datos.RAZON_SOCIAL || idCli) + "' guardado con éxito!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar cliente: " + error.toString() };
  }
}

/**
 * RPC: Actualizar cliente comercial existente
 */
function CLI_ACTUALIZAR_CLIENTE(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "CLIENTES", "EDITAR");
    }

    const idCli = datos.ID_CLIENTE || datos.ID_PROVEEDOR;
    if (!datos || !idCli) throw new Error("El ID_CLIENTE es requerido para actualizar.");

    if (String(idCli).indexOf("PROV-") === 0 || datos.TIPO_TERCERO === "PROVEEDOR") {
      return PROV_ACTUALIZAR_PROVEEDOR_WEB(datos, tokenSesion);
    }

    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_CLIENTES);
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const idxId = enc.indexOf("ID_CLIENTE");
    const reg = hoja.getRange(2, 1, Math.max(1, hoja.getLastRow() - 1), enc.length).getValues();

    let filaMod = -1;
    for (let i = 0; i < reg.length; i++) {
      if (String(reg[i][idxId]).trim() === String(idCli).trim()) {
        filaMod = i + 2;
        break;
      }
    }

    if (filaMod === -1) {
      return PROV_ACTUALIZAR_PROVEEDOR_WEB(datos, tokenSesion);
    }

    const docNum = datos.NUMERO_DOCUMENTO || datos.NIT_CC || "";
    const dvVal = datos.DIGITO_VERIFICACION || datos.DV || "";

    const filaActualizada = enc.map(col => {
      switch (col) {
        case "ID_CLIENTE": return datos.ID_CLIENTE;
        case "TIPO_TERCERO": return datos.TIPO_TERCERO || "CLIENTE";
        case "TIPO_PERSONA": return datos.TIPO_PERSONA || "PERSONA_JURIDICA";
        case "TIPO_DOCUMENTO": return datos.TIPO_DOCUMENTO || "NIT";
        case "NUMERO_DOCUMENTO": return docNum;
        case "DIGITO_VERIFICACION": return dvVal;
        case "RAZON_SOCIAL": return String(datos.RAZON_SOCIAL || datos.NOMBRE_COMERCIAL || "").trim().toUpperCase();
        case "PRIMER_NOMBRE": return datos.PRIMER_NOMBRE || "";
        case "SEGUNDO_NOMBRE": return datos.SEGUNDO_NOMBRE || "";
        case "PRIMER_APELLIDO": return datos.PRIMER_APELLIDO || "";
        case "SEGUNDO_APELLIDO": return datos.SEGUNDO_APELLIDO || "";
        case "NOMBRE_COMERCIAL": return datos.NOMBRE_COMERCIAL || "";
        case "DIRECCION": return datos.DIRECCION || "";
        case "CIUDAD": return datos.CIUDAD || "Pitalito";
        case "DEPARTAMENTO": return datos.DEPARTAMENTO || "Huila";
        case "PAIS": return datos.PAIS || "COLOMBIA";
        case "CELULAR": return datos.CELULAR || datos.TELEFONO || "";
        case "CORREO": return datos.CORREO || datos.EMAIL || "";
        case "RESPONSABILIDAD_FISCAL": return datos.RESPONSABILIDAD_FISCAL || "48";
        case "FORMA_PAGO": return datos.FORMA_PAGO || "CONTADO";
        case "PLAZO_PAGO_DIAS": return Number(datos.PLAZO_PAGO_DIAS || 30);
        case "LIMITE_CREDITO": return Number(datos.LIMITE_CREDITO || datos.CUPO_CREDITO || 0);
        case "ESTADO_CLIENTE": return datos.ESTADO_CLIENTE || datos.ESTADO || "ACTIVO";
        default: return hoja.getRange(filaMod, enc.indexOf(col) + 1).getValue();
      }
    });

    hoja.getRange(filaMod, 1, 1, enc.length).setValues([filaActualizada]);

    return {
      EXITO: true,
      ID_CLIENTE: datos.ID_CLIENTE,
      MENSAJE: "¡Cliente '" + datos.ID_CLIENTE + "' actualizado correctamente!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al actualizar cliente: " + error.toString() };
  }
}

/**
 * RPC: Buscar cliente por documento o nombre
 */
function CLI_BUSCAR_CLIENTE(q, tokenSesion) {
  try {
    const res = CLI_LISTAR_CLIENTES_WEB(tokenSesion);
    if (!res.EXITO || !res.DATOS) return res;
    const query = String(q || "").trim().toUpperCase();
    if (!query) return res;

    const filtrados = res.DATOS.filter(c => {
      return String(c.RAZON_SOCIAL || "").toUpperCase().indexOf(query) !== -1 ||
             String(c.NUMERO_DOCUMENTO || c.NIT_CC || "").toUpperCase().indexOf(query) !== -1 ||
             String(c.ID_CLIENTE || "").toUpperCase().indexOf(query) !== -1;
    });

    return { EXITO: true, DATOS: filtrados, MENSAJE: "Búsqueda finalizada." };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: error.toString() };
  }
}

/**
 * RPC: Listar Proveedores registrados
 */
function PROV_LISTAR_PROVEEDORES_WEB(tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "VER");
    }
    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_PROVEEDORES);
    const ulFila = hoja.getLastRow();
    if (ulFila < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay proveedores registrados." };
    }
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ulFila - 1, enc.length).getValues();

    const lista = reg.map(fila => {
      const obj = {};
      enc.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
      if (!obj.TIPO_TERCERO) obj.TIPO_TERCERO = "PROVEEDOR";
      if (!obj.ID_CLIENTE) obj.ID_CLIENTE = obj.ID_PROVEEDOR;
      return obj;
    });

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(lista) : JSON.parse(JSON.stringify(lista)),
      MENSAJE: "Proveedores listados correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar proveedores: " + error.toString() };
  }
}

/**
 * RPC: Guardar nuevo Proveedor
 */
function PROV_GUARDAR_PROVEEDOR_WEB(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "CREAR");
    }

    if (!datos || (!datos.NUMERO_DOCUMENTO && !datos.NIT_CC)) {
      throw new Error("El número de documento o NIT del proveedor es obligatorio.");
    }

    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_PROVEEDORES);
    const idProv = datos.ID_PROVEEDOR || datos.ID_CLIENTE || (CLI_CONFIG.PREFIJO_PROVEEDOR + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CLI_CONFIG.DIGITOS_ID, "0"));
    const ahora = new Date();
    const usuarioEjecutor = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "OPERADOR") : "OPERADOR";

    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const docNum = datos.NUMERO_DOCUMENTO || datos.NIT_CC || "";
    const dvVal = datos.DIGITO_VERIFICACION || datos.DV || "";

    const filaNueva = enc.map(col => {
      switch (col) {
        case "ID_PROVEEDOR": return idProv;
        case "TIPO_TERCERO": return datos.TIPO_TERCERO || "PROVEEDOR";
        case "TIPO_PERSONA": return datos.TIPO_PERSONA || "PERSONA_JURIDICA";
        case "TIPO_DOCUMENTO": return datos.TIPO_DOCUMENTO || "NIT";
        case "NUMERO_DOCUMENTO": return docNum;
        case "DIGITO_VERIFICACION": return dvVal;
        case "RAZON_SOCIAL": return String(datos.RAZON_SOCIAL || datos.NOMBRE_COMERCIAL || "PROVEEDOR SIN NOMBRE").trim().toUpperCase();
        case "PRIMER_NOMBRE": return datos.PRIMER_NOMBRE || "";
        case "SEGUNDO_NOMBRE": return datos.SEGUNDO_NOMBRE || "";
        case "PRIMER_APELLIDO": return datos.PRIMER_APELLIDO || "";
        case "SEGUNDO_APELLIDO": return datos.SEGUNDO_APELLIDO || "";
        case "NOMBRE_COMERCIAL": return datos.NOMBRE_COMERCIAL || datos.RAZON_SOCIAL || "";
        case "DIRECCION": return datos.DIRECCION || "";
        case "CIUDAD": return datos.CIUDAD || "Pitalito";
        case "DEPARTAMENTO": return datos.DEPARTAMENTO || "Huila";
        case "PAIS": return datos.PAIS || "COLOMBIA";
        case "CELULAR": return datos.CELULAR || datos.TELEFONO || "";
        case "CORREO": return datos.CORREO || datos.EMAIL || "";
        case "RESPONSABILIDAD_FISCAL": return datos.RESPONSABILIDAD_FISCAL || "48";
        case "FORMA_PAGO": return datos.FORMA_PAGO || "CONTADO";
        case "PLAZO_PAGO_DIAS": return Number(datos.PLAZO_PAGO_DIAS || 30);
        case "LIMITE_CREDITO": return Number(datos.LIMITE_CREDITO || datos.CUPO_CREDITO || 0);
        case "ESTADO": return datos.ESTADO || datos.ESTADO_CLIENTE || "ACTIVO";
        case "FECHA_CREACION": return ahora;
        case "USUARIO_CREACION": return usuarioEjecutor;
        default: return "";
      }
    });

hoja.appendRow(filaNueva);

    return {
      EXITO: true,
      ID_PROVEEDOR: idProv,
      MENSAJE: "¡Proveedor '" + (datos.RAZON_SOCIAL || idProv) + "' guardado con éxito!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar proveedor: " + error.toString() };
  }
}

/**
 * RPC: Actualizar Proveedor existente
 */
function PROV_ACTUALIZAR_PROVEEDOR_WEB(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "EDITAR");
    }

    const idProv = datos.ID_PROVEEDOR || datos.ID_CLIENTE;
    if (!idProv) throw new Error("El ID de proveedor es requerido para actualizar.");

    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_PROVEEDORES);
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const idxId = enc.indexOf("ID_PROVEEDOR");
    const reg = hoja.getRange(2, 1, Math.max(1, hoja.getLastRow() - 1), enc.length).getValues();

    let filaMod = -1;
    for (let i = 0; i < reg.length; i++) {
      if (String(reg[i][idxId]).trim() === String(idProv).trim()) {
        filaMod = i + 2;
        break;
      }
    }

    if (filaMod === -1) {
      if (String(idProv).indexOf("CLI-") === 0) {
        return CLI_ACTUALIZAR_CLIENTE(datos, tokenSesion);
      }
      throw new Error("Proveedor " + idProv + " no encontrado.");
    }

    const docNum = datos.NUMERO_DOCUMENTO || datos.NIT_CC || "";
    const dvVal = datos.DIGITO_VERIFICACION || datos.DV || "";

    const filaActualizada = enc.map(col => {
      switch (col) {
        case "ID_PROVEEDOR": return idProv;
        case "TIPO_TERCERO": return datos.TIPO_TERCERO || "PROVEEDOR";
        case "TIPO_PERSONA": return datos.TIPO_PERSONA || "PERSONA_JURIDICA";
        case "TIPO_DOCUMENTO": return datos.TIPO_DOCUMENTO || "NIT";
        case "NUMERO_DOCUMENTO": return docNum;
        case "DIGITO_VERIFICACION": return dvVal;
        case "RAZON_SOCIAL": return String(datos.RAZON_SOCIAL || datos.NOMBRE_COMERCIAL || "").trim().toUpperCase();
        case "PRIMER_NOMBRE": return datos.PRIMER_NOMBRE || "";
        case "SEGUNDO_NOMBRE": return datos.SEGUNDO_NOMBRE || "";
        case "PRIMER_APELLIDO": return datos.PRIMER_APELLIDO || "";
        case "SEGUNDO_APELLIDO": return datos.SEGUNDO_APELLIDO || "";
        case "NOMBRE_COMERCIAL": return datos.NOMBRE_COMERCIAL || "";
        case "DIRECCION": return datos.DIRECCION || "";
        case "CIUDAD": return datos.CIUDAD || "Pitalito";
        case "DEPARTAMENTO": return datos.DEPARTAMENTO || "Huila";
        case "PAIS": return datos.PAIS || "COLOMBIA";
        case "CELULAR": return datos.CELULAR || datos.TELEFONO || "";
        case "CORREO": return datos.CORREO || datos.EMAIL || "";
        case "RESPONSABILIDAD_FISCAL": return datos.RESPONSABILIDAD_FISCAL || "48";
        case "FORMA_PAGO": return datos.FORMA_PAGO || "CONTADO";
        case "PLAZO_PAGO_DIAS": return Number(datos.PLAZO_PAGO_DIAS || 30);
        case "LIMITE_CREDITO": return Number(datos.LIMITE_CREDITO || datos.CUPO_CREDITO || 0);
        case "ESTADO": return datos.ESTADO || datos.ESTADO_CLIENTE || "ACTIVO";
        default: return hoja.getRange(filaMod, enc.indexOf(col) + 1).getValue();
      }
    });

    hoja.getRange(filaMod, 1, 1, enc.length).setValues([filaActualizada]);

    return {
      EXITO: true,
      ID_PROVEEDOR: idProv,
      MENSAJE: "¡Proveedor '" + idProv + "' actualizado correctamente!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al actualizar proveedor: " + error.toString() };
  }
}

/**
 * RPC: Listar cuentas de Cartera (CxC)
 */
function CAR_LISTAR_CUENTAS_WEB(tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "FINANZAS", "VER");
    }

    const hoja = CLI_OBTENER_HOJA(CLI_CONFIG.HOJA_CARTERA);
    const ulFila = hoja.getLastRow();
    if (ulFila < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay registros de cartera pendientes." };
    }

    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ulFila - 1, enc.length).getValues();

    const lista = reg.map(fila => {
      const obj = {};
      enc.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
      return obj;
    });

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(lista) : JSON.parse(JSON.stringify(lista)),
      MENSAJE: "Cartera listada correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al consultar cartera: " + error.toString() };
  }
}


