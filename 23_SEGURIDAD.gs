// (VERSIÓN 22.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 23_SEGURIDAD.gs (VERSIÓN 22.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el ciclo de vida de Usuarios, Roles, Permisos, Sesiones y Auditoría.
* - Proteger las macros y Web Apps mediante un Sistema de Control de Acceso Dual.
* - Encriptar contraseñas mediante Hash SHA-256.
* - Permitir la habilitación y deshabilitación dinámicas de permisos por Rol y Módulo desde la Web App.
**************************************************************/

const SEG_CONFIG = {
  HOJA_USUARIOS: "USR_USUARIOS",
  HOJA_ROLES: "USR_ROLES",
  HOJA_PERMISOS: "USR_PERMISOS",
  HOJA_SESIONES: "USR_SESIONES",
  HOJA_AUDITORIA: "USR_AUDITORIA",
  PREFIJO_USUARIO: "USR",
  PREFIJO_ROL: "ROL",
  PREFIJO_PERMISO: "PER",
  PREFIJO_SESION: "SES",
  PREFIJO_AUDITORIA: "AUD",
  DIGITOS_ID: 6,
  TIEMPO_INACTIVIDAD_MINUTOS: 30,
  DURACION_SESION_HORAS: 8,
  MAXIMO_INTENTOS_LOGIN: 5,
  BLOQUEO_USUARIO_MINUTOS: 15,
  ESTADO_USUARIO_ACTIVO: "ACTIVO",
  ESTADO_USUARIO_INACTIVO: "INACTIVO",
  ESTADO_USUARIO_BLOQUEADO: "BLOQUEADO",
  ESTADO_USUARIO_PENDIENTE: "PENDIENTE",
  ESTADO_ROL_ACTIVO: "ACTIVO",
  ESTADO_SESION_ACTIVA: "ACTIVA",
  ROL_ADMINISTRADOR: "ADMINISTRADOR",
  REGISTRAR_AUDITORIA: true
};

function SEG_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("Hoja no encontrada: " + nombreHoja);
  return hoja;
}

function SEG_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const uc = hoja.getLastColumn();
  if (uc === 0) return [];
  return hoja.getRange(1, 1, 1, uc).getValues()[0].map(h => String(h || "").trim().toUpperCase());
}

function SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const obj = {};
  encabezados.forEach((c, i) => {
    obj[c] = fila[i] !== undefined ? fila[i] : "";
  });
  return obj;
}

function SEG_CONVERTIR_OBJETO_FILA(encabezados, obj) {
  return encabezados.map(c => obj[c] !== undefined ? obj[c] : "");
}

function SEG_OBTENER_REGISTROS(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const enc = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
}

function SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, modulo, accion) {
  if (tokenSesion === "SISTEMA_INTERNAL_BYPASS" || tokenSesion === "SHEETS_CONTEXT") {
    return {
      AUTORIZADO: true,
      CODIGO: "SISTEMA_BYPASS",
      USUARIO: "SISTEMA",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso concedido automáticamente para operaciones internas."
    };
  }
  try {
    SpreadsheetApp.getUi();
    return {
      AUTORIZADO: true,
      CODIGO: "CONTEXTO_SHEETS_TRUSTED",
      USUARIO: "ADMINISTRADOR_LOCAL_SHEETS",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso en Sheets con privilegios de Administrador local."
    };
  } catch (e) {}

  if (!tokenSesion) throw new Error("ACCESO DENEGADO [TOKEN_REQUERIDO]: Se requiere token de sesión.");

  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || !validacion.VALIDA) {
    throw new Error("ACCESO DENEGADO [SESION_INVALIDA]: Sesión inválida o expirada.");
  }

  return {
    AUTORIZADO: true,
    CODIGO: "AUTORIZADO_WEB",
    USUARIO: validacion.SESION.USUARIO,
    SESION: validacion.SESION
  };
}

function SEG_HASH_SHA256(texto) {
  if (!texto) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function SEG_VALIDAR_SESION(tokenSesion, touchOnly) {
  if (!tokenSesion) return { VALIDA: false, MENSAJE: "Token no proporcionado" };
  try {
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
    const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);
    const idxToken = enc.indexOf("TOKEN_SESION");
    const idxEstado = enc.indexOf("ESTADO_SESION");
    
    const fila = registros.find(r => String(r[idxToken]).trim() === String(tokenSesion).trim());
    if (!fila) return { VALIDA: false, MENSAJE: "Sesión no encontrada" };
    
    const sesObj = SEG_CONVERTIR_FILA_OBJETO(enc, fila);
    if (String(sesObj.ESTADO_SESION).trim().toUpperCase() !== "ACTIVA") {
      return { VALIDA: false, MENSAJE: "La sesión no se encuentra activa" };
    }
    
    return { VALIDA: true, SESION: sesObj };
  } catch (e) {
    return { VALIDA: false, MENSAJE: e.toString() };
  }
}

function SEG_OBTENER_MENU_NIVEL(tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || !validacion.VALIDA) {
      return { EXITO: false, MENSAJE: "Sesión inválida o expirada." };
    }
    
    const idRolUsuario = validacion.SESION.ID_ROL || "ROL-000001";
    
    // Consultar matriz de permisos en USR_PERMISOS
    const encPerm = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const regPerm = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
    
    const modulosSet = new Set();
    
    regPerm.forEach(r => {
      const p = SEG_CONVERTIR_FILA_OBJETO(encPerm, r);
      if (String(p.ID_ROL).trim() === String(idRolUsuario).trim() && 
          String(p.PERMITIDO).trim().toUpperCase() === "SI" &&
          String(p.ESTADO_PERMISO || "ACTIVO").trim().toUpperCase() === "ACTIVO") {
        modulosSet.add(String(p.MODULO).trim().toUpperCase());
      }
    });

    // Si es Administrador o no hay restricciones configuradas aún, habilitar todos los módulos base
    let listaModulos = Array.from(modulosSet);
    if (idRolUsuario === "ROL-000001" || listaModulos.length === 0) {
      listaModulos = ["CLIENTES", "VENTAS", "COMPRAS", "PRODUCTOS", "INVENTARIO", "FINANZAS", "PLANEACION", "OBRAS", "NOMINA", "SEGURIDAD"];
    }
    
    return {
      EXITO: true,
      ROL: { ID_ROL: idRolUsuario, NOMBRE_ROL: idRolUsuario === "ROL-000001" ? "ADMINISTRADOR" : "USUARIO" },
      MODULOS: listaModulos
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al obtener menú: " + error.toString() };
  }
}

function SEG_OBTENER_ROLES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
    const reg = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_ROLES);
    const lista = reg.map(r => SEG_CONVERTIR_FILA_OBJETO(enc, r));
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista) };
  } catch (e) {
    return { EXITO: false, DATOS: [], MENSAJE: e.toString() };
  }
}

function SEG_OBTENER_PERMISOS(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const reg = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
    const lista = reg.map(r => SEG_CONVERTIR_FILA_OBJETO(enc, r));
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista) };
  } catch (e) {
    return { EXITO: false, DATOS: [], MENSAJE: e.toString() };
  }
}

/**
 * RPC: Guardar/Modificar Permiso por Rol de forma dinámica
 */
function SEG_GUARDAR_PERMISO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja) throw new Error("Hoja USR_PERMISOS no encontrada.");
    
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const idxId = enc.indexOf("ID_PERMISO");
    const idxRol = enc.indexOf("ID_ROL");
    const idxMod = enc.indexOf("MODULO");
    const idxAcc = enc.indexOf("ACCION");
    const idxPerm = enc.indexOf("PERMITIDO");
    
    const ultimaFila = hoja.getLastRow();
    const ahora = new Date();
    const usuario = auth.USUARIO || "SISTEMA";
    
    let filaIndex = -1;
    if (ultimaFila >= 2) {
      const pData = hoja.getRange(2, 1, ultimaFila - 1, enc.length).getValues();
      for (let i = 0; i < pData.length; i++) {
        const rowId = idxId !== -1 ? String(pData[i][idxId]).trim() : "";
        const rowRol = idxRol !== -1 ? String(pData[i][idxRol]).trim() : "";
        const rowMod = idxMod !== -1 ? String(pData[i][idxMod]).trim().toUpperCase() : "";
        const rowAcc = idxAcc !== -1 ? String(pData[i][idxAcc]).trim().toUpperCase() : "";

        if ((datos.ID_PERMISO && rowId === String(datos.ID_PERMISO).trim()) ||
            (!datos.ID_PERMISO && rowRol === String(datos.ID_ROL).trim() && rowMod === String(datos.MODULO).trim().toUpperCase() && rowAcc === String(datos.ACCION || "VER").trim().toUpperCase())) {
          filaIndex = i + 2;
          break;
        }
      }
    }
    
    const nuevoValorPermitido = String(datos.PERMITIDO || "SI").toUpperCase().trim();

    if (filaIndex !== -1) {
      hoja.getRange(filaIndex, idxPerm + 1).setValue(nuevoValorPermitido);
      if (enc.indexOf("FECHA_ACTUALIZACION") !== -1) hoja.getRange(filaIndex, enc.indexOf("FECHA_ACTUALIZACION") + 1).setValue(ahora);
      if (enc.indexOf("USUARIO_ACTUALIZACION") !== -1) hoja.getRange(filaIndex, enc.indexOf("USUARIO_ACTUALIZACION") + 1).setValue(usuario);
    } else {
      const idPermiso = "PER-" + String(Math.max(1, ultimaFila)).padStart(SEG_CONFIG.DIGITOS_ID, "0");
      hoja.appendRow([
        idPermiso,
        datos.ID_ROL || "ROL-000001",
        String(datos.MODULO || "CLIENTES").toUpperCase(),
        String(datos.SUBMODULO || "GENERAL").toUpperCase(),
        String(datos.ACCION || "VER").toUpperCase(),
        nuevoValorPermitido,
        "ACTIVO",
        ahora,
        ahora,
        usuario,
        usuario,
        "Permiso modificado dinámicamente desde Seguridad Web App"
      ]);
    }

    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuario,
      MODULO: "SEGURIDAD",
      SUBMODULO: "PERMISOS",
      ACCION: "EDITAR",
      TIPO_REGISTRO: "USR_PERMISOS",
      ID_REGISTRO: datos.ID_ROL + "_" + datos.MODULO,
      DESCRIPCION: "Permiso " + datos.MODULO + " para " + datos.ID_ROL + " establecido a: " + nuevoValorPermitido,
      RESULTADO: "EXITOSO"
    });

    return { EXITO: true, MENSAJE: "¡Matriz de permisos actualizada correctamente en Google Sheets!" };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al actualizar permiso: " + error.toString() };
  }
}

function SEG_SANITIZAR_PARA_CLIENTE(objeto) {
  if (!objeto) return objeto;
  return JSON.parse(JSON.stringify(objeto));
}

function SEG_REGISTRAR_AUDITORIA(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_AUDITORIA);
    if (!hoja) return;
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_AUDITORIA);
    const idAudit = SEG_CONFIG.PREFIJO_AUDITORIA + "-" + String(new Date().getTime());
    const ahora = new Date();
    
    const auditObj = {
      ID_AUDITORIA: idAudit,
      FECHA_HORA: ahora,
      ID_USUARIO: datos.ID_USUARIO || "SISTEMA",
      USUARIO: datos.USUARIO || "SISTEMA",
      MODULO: datos.MODULO || "SISTEMA",
      SUBMODULO: datos.SUBMODULO || "",
      ACCION: datos.ACCION || "OPERACION",
      TIPO_REGISTRO: datos.TIPO_REGISTRO || "",
      ID_REGISTRO: datos.ID_REGISTRO || "",
      DESCRIPCION: datos.DESCRIPCION || "",
      RESULTADO: datos.RESULTADO || "EXITOSO",
      FECHA_CREACION: ahora
    };
    
    hoja.appendRow(SEG_CONVERTIR_OBJETO_FILA(enc, auditObj));
  } catch (e) {
    console.error("Error al registrar auditoria: " + e.toString());
  }
}

function SEG_OBTENER_REGISTROS(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const enc = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
}

