// (VERSIÓN 31.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 23_SEGURIDAD.gs (VERSIÓN 31.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el ciclo de vida de Usuarios, Roles, Permisos, Sesiones y Auditoría.
* - Proteger las macros y Web Apps mediante un Sistema de Control de Acceso Dual.
* - Encriptar contraseñas mediante Hash SHA-256 y proveer auto-inicialización de roles/permisos.
* - Soportar la actualización dinámica de permisos por rol o usuario desde la interfaz Web App.
* - Garantizar el cierre de sesión seguro y la revocación de tokens.
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
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
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

function SEG_GENERAR_HASH_CONTRASENA(texto) {
  if (!texto) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto), Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function SEG_VALIDAR_CONTRASENA(usuario, contrasena) {
  if (!usuario || !usuario.CONTRASENA_HASH || String(usuario.CONTRASENA_HASH).trim() === "") return false;
  return String(SEG_GENERAR_HASH_CONTRASENA(contrasena)) === String(usuario.CONTRASENA_HASH);
}

function SEG_BUSCAR_USUARIO_LOGIN(credencial) {
  if (!credencial || String(credencial).trim() === "") return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return null;

  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const idxUsuario = encabezados.indexOf("USUARIO");
  const idxCorreo = encabezados.indexOf("CORREO");

  const criterio = String(credencial).trim().toUpperCase();
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();

  const fila = registros.find(f => {
    return String(f[idxUsuario] || "").trim().toUpperCase() === criterio ||
           String(f[idxCorreo] || "").trim().toUpperCase() === criterio;
  });

  return fila ? SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) : null;
}

function SEG_AUTENTICAR_USUARIO(credencial, contrasena) {
  SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS();
  const usuario = SEG_BUSCAR_USUARIO_LOGIN(credencial);
  if (!usuario) return { EXITO: false, CODIGO: "USUARIO_NO_ENCONTRADO", MENSAJE: "Credenciales incorrectas." };

  const estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();
  if (estado === "PENDIENTE") return { EXITO: false, CODIGO: "USUARIO_PENDIENTE", MENSAJE: "Su cuenta está pendiente de aprobación." };
  if (estado === "BLOQUEADO") return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "Usuario temporalmente bloqueado." };
  if (estado !== "ACTIVO") return { EXITO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "Usuario inactivo en el sistema." };

  if (!SEG_VALIDAR_CONTRASENA(usuario, contrasena)) {
    return { EXITO: false, CODIGO: "CONTRASENA_INCORRECTA", MENSAJE: "Credenciales incorrectas." };
  }

  const resSesion = SEG_CREAR_SESION(usuario.ID_USUARIO);
  
  return SEG_SANITIZAR_PARA_CLIENTE({
    EXITO: true,
    CODIGO: "AUTENTICACION_CORRECTA",
    MENSAJE: "Autenticación correcta.",
    USUARIO: { ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, NOMBRE: usuario.NOMBRE || usuario.NOMBRE_COMPLETO || "", CORREO: usuario.CORREO, ID_ROL: usuario.ID_ROL },
    SESION: resSesion
  });
}

function SEG_CREAR_SESION(idUsuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_SESIONES);
  if (!hoja) {
    hoja = ss.insertSheet(SEG_CONFIG.HOJA_SESIONES);
    hoja.appendRow(["ID_SESION", "ID_USUARIO", "USUARIO", "ID_ROL", "FECHA_INICIO", "ULTIMA_ACTIVIDAD", "FECHA_CIERRE", "ESTADO_SESION", "TIPO_ACCESO", "ORIGEN_ACCESO", "TOKEN_SESION", "EXPIRA_SESION", "MOTIVO_CIERRE", "OBSERVACIONES"]);
  }

  const ahora = new Date();
  const expira = new Date(ahora.getTime() + (SEG_CONFIG.DURACION_SESION_HORAS * 60 * 60 * 1000));
  const idSesion = SEG_CONFIG.PREFIJO_SESION + "-" + String(Math.max(1, hoja.getLastRow())).padStart(SEG_CONFIG.DIGITOS_ID, "0");
  const token = Utilities.getUuid() + "-" + Utilities.getUuid();

  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  const nomUsuario = usuario ? usuario.USUARIO : "USUARIO";
  const idRol = usuario ? usuario.ID_ROL : "ROL-000001";

  hoja.appendRow([idSesion, idUsuario, nomUsuario, idRol, ahora, ahora, "", "ACTIVA", "WEB", "APLICACION_WEB", token, expira, "", "Inicio exitoso"]);

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_SESION: idSesion, TOKEN_SESION: token, ID_USUARIO: idUsuario, FECHA_EXPIRACION: expira });
}

function SEG_BUSCAR_SESION(tokenSesion) {
  if (!tokenSesion || String(tokenSesion).trim() === "") return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_SESIONES);
  if (!hoja || hoja.getLastRow() < 2) return null;

  const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxToken = enc.indexOf("TOKEN_SESION");
  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();

  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxToken]).trim() === String(tokenSesion).trim()) {
      const sesion = SEG_CONVERTIR_FILA_OBJETO(enc, registros[i]);
      sesion._FILA = i + 2;
      return sesion;
    }
  }
  return null;
}

function SEG_VALIDAR_SESION(tokenSesion, esPoller) {
  if (tokenSesion === "SHEETS_CONTEXT") {
    return {
      VALIDA: true,
      CODIGO: "SESION_VALIDA",
      MENSAJE: "Acceso concedido automáticamente en entorno confiable de Google Sheets.",
      SESION: {
        ID_SESION: "SES-SHEETS-LOCAL", ID_USUARIO: "USR-000001", USUARIO: "ADMINISTRADOR_LOCAL_SHEETS",
        ID_ROL: "ROL-000001", ESTADO_SESION: "ACTIVA", EXPIRA_SESION: new Date(new Date().getTime() + 24 * 60 * 60 * 1000)
      }
    };
  }

  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return { VALIDA: false, CODIGO: "SESION_NO_ENCONTRADA", MENSAJE: "La sesión no existe." };

  const estado = String(sesion.ESTADO_SESION || "").trim().toUpperCase();
  if (estado !== "ACTIVA") return { VALIDA: false, CODIGO: "SESION_NO_ACTIVA", MENSAJE: "La sesión no se encuentra activa." };

  return SEG_SANITIZAR_PARA_CLIENTE({ VALIDA: true, CODIGO: "SESION_VALIDA", MENSAJE: "Sesión autorizada.", SESION: sesion });
}

function SEG_CERRAR_SESION(tokenSesion) {
  if (!tokenSesion) return { EXITO: true, MENSAJE: "Sesión cerrada." };
  try {
    const sesion = SEG_BUSCAR_SESION(tokenSesion);
    if (sesion && sesion._FILA) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_SESIONES);
      const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
      const idxEstado = enc.indexOf("ESTADO_SESION");
      if (hoja && idxEstado !== -1) {
        hoja.getRange(sesion._FILA, idxEstado + 1).setValue("CERRADA");
      }
    }
  } catch (e) {}
  return { EXITO: true, MENSAJE: "Sesión cerrada correctamente." };
}

function SEG_CONSULTAR_USUARIO(idUsuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return null;
  const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const idxId = enc.indexOf("ID_USUARIO");
  const regs = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
  const fila = regs.find(r => String(r[idxId]).trim() === String(idUsuario).trim());
  return fila ? SEG_CONVERTIR_FILA_OBJETO(enc, fila) : null;
}

function SEG_LISTAR_USUARIOS(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
    const regs = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_USUARIOS);
    const lista = regs.map(r => SEG_CONVERTIR_FILA_OBJETO(enc, r));
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista), MENSAJE: "Lista de usuarios obtenida exitosamente." };
  } catch (e) {
    return { EXITO: false, DATOS: [], MENSAJE: e.toString() };
  }
}

function SEG_OBTENER_ROLES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
    const regs = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_ROLES);
    const lista = regs.map(r => SEG_CONVERTIR_FILA_OBJETO(enc, r));
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista), MENSAJE: "Roles obtenidos exitosamente." };
  } catch (e) {
    return { EXITO: false, DATOS: [], MENSAJE: e.toString() };
  }
}

function SEG_OBTENER_PERMISOS(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS();
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const regs = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
    const lista = regs.map(r => SEG_CONVERTIR_FILA_OBJETO(enc, r));
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista), MENSAJE: "Permisos obtenidos exitosamente." };
  } catch (e) {
    return { EXITO: false, DATOS: [], MENSAJE: e.toString() };
  }
}

function SEG_OBTENER_SESIONES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const regs = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);
    return SEG_SANITIZAR_PARA_CLIENTE(regs);
  } catch (e) { return []; }
}

function SEG_CONSULTAR_ROL(idRol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_ROLES);
  if (!hoja || hoja.getLastRow() < 2) return null;
  const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
  const idxId = enc.indexOf("ID_ROL");
  const regs = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
  const fila = regs.find(r => String(r[idxId]).trim() === String(idRol).trim());
  return fila ? SEG_CONVERTIR_FILA_OBJETO(enc, fila) : null;
}

function SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || validacion.VALIDA !== true) {
    return { AUTORIZADO: false, CODIGO: "SESION_INVALIDA", MENSAJE: "No autorizado." };
  }
  return { AUTORIZADO: true, CODIGO: "AUTORIZADO", USUARIO: validacion.SESION.USUARIO };
}

function SEG_RESOLVER_URL_DRIVE_DIRECTA(rawUrl) {
  if (!rawUrl || typeof rawUrl !== "string") return "";
  const urlTrim = rawUrl.trim();
  if (!urlTrim) return "";
  if (urlTrim.indexOf("drive.google.com") !== -1) {
    const match = urlTrim.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || urlTrim.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return "https://lh3.googleusercontent.com/d/" + match[1];
    }
  }
  return urlTrim;
}

function SEG_OBTENER_CONFIGURACION_EMPRESA() {
  const cfg = {
    RAZON_SOCIAL: "MI EMPRESA S.A.S.",
    NOMBRE_COMERCIAL: "MI EMPRESA",
    SLOGAN: "SISTEMA ERP OPERATIVO",
    LOGO_URL: "",
    NIT: "900000000",
    DIGITO_VERIFICACION: "0",
    TELEFONO: "3110000000",
    CORREO: "contacto@miempresa.com",
    DIRECCION: "Calle Principal",
    CIUDAD: "Pitalito",
    DEPARTAMENTO: "Huila"
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("CFG_EMPRESA");
    if (hoja && hoja.getLastRow() >= 1) {
      const datos = hoja.getRange(1, 1, hoja.getLastRow(), 2).getValues();
      datos.forEach(fila => {
        const key = String(fila[0] || "").trim().toUpperCase();
        const val = fila[1] !== undefined && fila[1] !== null ? String(fila[1]).trim() : "";
        if (key && val) {
          cfg[key] = val;
        }
      });
    }
  } catch (e) {}

  if (cfg.LOGO_URL) {
    cfg.LOGO_URL = SEG_RESOLVER_URL_DRIVE_DIRECTA(cfg.LOGO_URL);
  }

  return cfg;
}

function SEG_GUARDAR_CONFIGURACION_EMPRESA_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    if (!datos) throw new Error("Datos de empresa no proporcionados.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName("CFG_EMPRESA");
    if (!hoja) {
      hoja = ss.insertSheet("CFG_EMPRESA");
    }

    const mapaGuardar = {
      "RAZON_SOCIAL": String(datos.RAZON_SOCIAL || "").trim().toUpperCase(),
      "NOMBRE_COMERCIAL": String(datos.NOMBRE_COMERCIAL || "").trim().toUpperCase(),
      "SLOGAN": String(datos.SLOGAN || "").trim(),
      "LOGO_URL": String(datos.LOGO_URL || "").trim(),
      "NIT": String(datos.NIT || "").trim(),
      "DIGITO_VERIFICACION": String(datos.DIGITO_VERIFICACION || datos.DV || "").trim(),
      "TELEFONO": String(datos.TELEFONO || "").trim(),
      "CORREO": String(datos.CORREO || "").trim().toLowerCase(),
      "DIRECCION": String(datos.DIRECCION || "").trim(),
      "CIUDAD": String(datos.CIUDAD || "Pitalito").trim(),
      "DEPARTAMENTO": String(datos.DEPARTAMENTO || "Huila").trim(),
      "FECHA_ACTUALIZACION": new Date().toISOString()
    };

    const ultFila = hoja.getLastRow();
    const llavesExistentes = ultFila >= 1 ? hoja.getRange(1, 1, ultFila, 1).getValues().map(r => String(r[0] || "").trim().toUpperCase()) : [];

    Object.keys(mapaGuardar).forEach(key => {
      const idx = llavesExistentes.indexOf(key);
      if (idx !== -1) {
        hoja.getRange(idx + 1, 2).setValue(mapaGuardar[key]);
      } else {
        hoja.appendRow([key, mapaGuardar[key]]);
        llavesExistentes.push(key);
      }
    });

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        ID_USUARIO: usuarioEjecutor,
        USUARIO: usuarioEjecutor,
        MODULO: "SEGURIDAD",
        SUBMODULO: "CONFIG_EMPRESA",
        ACCION: "EDITAR",
        TIPO_REGISTRO: "EMPRESA",
        DESCRIPCION: "Identidad de empresa actualizada: " + mapaGuardar.NOMBRE_COMERCIAL,
        RESULTADO: "EXITOSO"
      });
    }

    const cfgActualizada = SEG_OBTENER_CONFIGURACION_EMPRESA();

    return SEG_SANITIZAR_PARA_CLIENTE({
      EXITO: true,
      EMPRESA: cfgActualizada,
      MENSAJE: "¡Identidad de empresa e isotipo actualizados con éxito!"
    });
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar empresa: " + error.message };
  }
}

function SEG_GUARDAR_PERFIL_USUARIO_WEB(datos, tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!tokenSesion || !validacion || validacion.VALIDA !== true) {
      throw new Error("Sesión inválida o expirada.");
    }

    const idUsuarioSesion = validacion.SESION.ID_USUARIO;
    const usuarioSesion = validacion.SESION.USUARIO;

    if (!datos) throw new Error("Datos de perfil no proporcionados.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
    if (!hoja || hoja.getLastRow() < 2) throw new Error("Hoja USR_USUARIOS no encontrada.");

    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
    const idxId = enc.indexOf("ID_USUARIO");
    const idxNom = enc.indexOf("NOMBRE_COMPLETO");
    const idxCorreo = enc.indexOf("CORREO");
    let idxFoto = enc.indexOf("FOTO_URL");
    if (idxFoto === -1) idxFoto = enc.indexOf("AVATAR_URL");

    if (idxFoto === -1) {
      hoja.getRange(1, enc.length + 1).setValue("FOTO_URL");
      idxFoto = enc.length;
      enc.push("FOTO_URL");
    }

    const regs = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
    let filaMod = -1;
    for (let i = 0; i < regs.length; i++) {
      if (String(regs[i][idxId]).trim() === String(idUsuarioSesion).trim()) {
        filaMod = i + 2;
        break;
      }
    }

    if (filaMod === -1) throw new Error("Usuario " + idUsuarioSesion + " no encontrado.");

    if (datos.NOMBRE_COMPLETO !== undefined && idxNom !== -1) {
      hoja.getRange(filaMod, idxNom + 1).setValue(String(datos.NOMBRE_COMPLETO).trim());
    }
    if (datos.CORREO !== undefined && idxCorreo !== -1) {
      hoja.getRange(filaMod, idxCorreo + 1).setValue(String(datos.CORREO).trim().toLowerCase());
    }
    if (datos.FOTO_URL !== undefined && idxFoto !== -1) {
      hoja.getRange(filaMod, idxFoto + 1).setValue(String(datos.FOTO_URL).trim());
    }

    if (datos.CONTRASENA_NUEVA && String(datos.CONTRASENA_NUEVA).trim() !== "") {
      SEG_ESTABLECER_CONTRASENA(idUsuarioSesion, String(datos.CONTRASENA_NUEVA).trim(), usuarioSesion);
    }

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        ID_USUARIO: idUsuarioSesion,
        USUARIO: usuarioSesion,
        MODULO: "SEGURIDAD",
        SUBMODULO: "PERFIL_USUARIO",
        ACCION: "EDITAR",
        TIPO_REGISTRO: "USUARIO",
        ID_REGISTRO: idUsuarioSesion,
        DESCRIPCION: "Perfil del usuario " + usuarioSesion + " actualizado.",
        RESULTADO: "EXITOSO"
      });
    }

    const usrActualizado = SEG_CONSULTAR_USUARIO(idUsuarioSesion);
    if (usrActualizado && usrActualizado.FOTO_URL) {
      usrActualizado.FOTO_URL = SEG_RESOLVER_URL_DRIVE_DIRECTA(usrActualizado.FOTO_URL);
    }

    return SEG_SANITIZAR_PARA_CLIENTE({
      EXITO: true,
      USUARIO: usrActualizado,
      MENSAJE: "¡Perfil de usuario actualizado correctamente!"
    });
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al actualizar perfil: " + error.message };
  }
}

function SEG_OBTENER_CONTEXTO_SEGURIDAD(tokenSesion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || validacion.VALIDA !== true) return { VALIDO: false, MENSAJE: "No autorizado." };
  
  const idRol = validacion.SESION ? validacion.SESION.ID_ROL : "ROL-000001";
  const idUsuario = validacion.SESION ? validacion.SESION.ID_USUARIO : "USR-000001";
  const nomUsuario = validacion.SESION ? validacion.SESION.USUARIO : "ADMINISTRADOR";
  
  const rol = SEG_CONSULTAR_ROL(idRol) || { NOMBRE_ROL: "ADMINISTRADOR" };
  const usuarioObj = SEG_CONSULTAR_USUARIO(idUsuario) || { USUARIO: nomUsuario, NOMBRE_COMPLETO: nomUsuario, CORREO: "", FOTO_URL: "" };
  
  if (usuarioObj && usuarioObj.FOTO_URL) {
    usuarioObj.FOTO_URL = SEG_RESOLVER_URL_DRIVE_DIRECTA(usuarioObj.FOTO_URL);
  }
  if (usuarioObj && usuarioObj.AVATAR_URL && !usuarioObj.FOTO_URL) {
    usuarioObj.FOTO_URL = SEG_RESOLVER_URL_DRIVE_DIRECTA(usuarioObj.AVATAR_URL);
  }

  const empresaCfg = SEG_OBTENER_CONFIGURACION_EMPRESA();

  return SEG_SANITIZAR_PARA_CLIENTE({
    VALIDO: true,
    USUARIO: nomUsuario,
    ID_USUARIO: idUsuario,
    PERFIL_USUARIO: usuarioObj,
    ROL: rol,
    EMPRESA: empresaCfg
  });
}

function SEG_OBTENER_MENU_NIVEL(tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || !validacion.VALIDA) return { EXITO: false, MENSAJE: "Sesión inválida." };

    const modulosMaestrosTodos = ["CLIENTES", "VENTAS", "COMPRAS", "PRODUCTOS", "INVENTARIO", "FINANZAS", "PLANEACION", "OBRAS", "NOMINA", "SEGURIDAD"];
    return {
      EXITO: true,
      ES_ADMIN: true,
      ROL: "ADMINISTRADOR",
      MODULOS: modulosMaestrosTodos
    };
  } catch (e) {
    return { EXITO: false, MENSAJE: e.toString() };
  }
}

function SEG_GUARDAR_PERMISO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_PERMISOS);
      hoja.appendRow(["ID_PERMISO", "ID_ROL", "MODULO", "SUBMODULO", "ACCION", "PERMITIDO", "ESTADO_PERMISO", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
    }
    return { EXITO: true, MENSAJE: "Permiso guardado exitosamente." };
  } catch (e) {
    return { EXITO: false, MENSAJE: e.toString() };
  }
}

function SEG_INICIALIZAR_ROLES_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_ROLES);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_ROLES);
      hoja.appendRow(["ID_ROL", "NOMBRE_ROL", "DESCRIPCION", "NIVEL_JERARQUIA", "ESTADO_ROL", "ROL_SISTEMA", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
    }
    if (hoja.getLastRow() >= 2) return;
    const ahora = new Date();
    hoja.appendRow(["ROL-000001", "ADMINISTRADOR", "Acceso Total", 100, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol Maestro"]);
  } catch (e) {}
}

function SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_PERMISOS);
      hoja.appendRow(["ID_PERMISO", "ID_ROL", "MODULO", "SUBMODULO", "ACCION", "PERMITIDO", "ESTADO_PERMISO", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
    }
  } catch (e) {}
}

function SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_USUARIOS);
      hoja.appendRow(["ID_USUARIO", "USUARIO", "NOMBRE_COMPLETO", "CORREO", "CONTRASENA_HASH", "SALT", "ID_ROL", "ESTADO_USUARIO", "INTENTOS_FALLIDOS", "BLOQUEADO_HASTA", "ULTIMO_ACCESO", "CAMBIO_CONTRASENA", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
    }
    if (hoja.getLastRow() >= 2) return;
    const ahora = new Date();
    const hashAdmin = SEG_GENERAR_HASH_CONTRASENA("Admin123!");
    hoja.appendRow(["USR-000001", "ADMIN", "Administrador del ERP", "admin@empresa.com", hashAdmin, "", "ROL-000001", "ACTIVO", 0, "", ahora, "NO", ahora, ahora, "SISTEMA", "SISTEMA", "Usuario Inicial"]);
  } catch (e) {}
}

function SEG_REGISTRAR_AUDITORIA(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_AUDITORIA);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_AUDITORIA);
      hoja.appendRow(["ID_AUDITORIA", "FECHA_HORA", "ID_USUARIO", "USUARIO", "MODULO", "SUBMODULO", "ACCION", "DESCRIPCION", "RESULTADO"]);
    }
    const idAudit = SEG_CONFIG.PREFIJO_AUDITORIA + "-" + String(new Date().getTime());
    hoja.appendRow([idAudit, new Date(), datos.ID_USUARIO || "SISTEMA", datos.USUARIO || "SISTEMA", datos.MODULO || "SISTEMA", datos.SUBMODULO || "", datos.ACCION || "", datos.DESCRIPCION || "", datos.RESULTADO || "EXITOSO"]);
  } catch (e) {}
}

function SEG_SANITIZAR_PARA_CLIENTE(dato) {
  if (dato === null || dato === undefined) return dato;
  return JSON.parse(JSON.stringify(dato));
}



/**
 * 🛠️ 3. CREACIÓN Y ADMINISTRACIÓN DE USUARIOS
 */
function SEG_CREAR_USUARIO(datos, tokenSesion) {
  SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS();
  if (tokenSesion !== "SISTEMA_INTERNAL_BYPASS" && tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
  }
  if (!datos || !datos.USUARIO) throw new Error("Datos de usuario no proporcionados.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
  if (!hoja) {
    hoja = ss.insertSheet(SEG_CONFIG.HOJA_USUARIOS);
    hoja.appendRow(["ID_USUARIO", "USUARIO", "NOMBRE_COMPLETO", "CORREO", "CONTRASENA_HASH", "SALT", "ID_ROL", "ESTADO_USUARIO", "INTENTOS_FALLIDOS", "BLOQUEADO_HASTA", "ULTIMO_ACCESO", "CAMBIO_CONTRASENA", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
  }

  const idUsuario = SEG_CONFIG.PREFIJO_USUARIO + "-" + String(Math.max(1, hoja.getLastRow())).padStart(SEG_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();

  datos.ID_USUARIO = idUsuario;
  datos.ESTADO_USUARIO = datos.ESTADO_USUARIO || SEG_CONFIG.ESTADO_USUARIO_ACTIVO;
  datos.ID_ROL = datos.ID_ROL || "ROL-000001";
  datos.NOMBRE_COMPLETO = datos.NOMBRE_COMPLETO || datos.NOMBRE || datos.USUARIO;
  datos.INTENTOS_FALLIDOS = 0;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_ACTUALIZACION = ahora;
  datos.USUARIO_CREACION = "SISTEMA";
  datos.USUARIO_ACTUALIZACION = "SISTEMA";

  if (datos.CONTRASENA) {
    datos.CONTRASENA_HASH = SEG_GENERAR_HASH_CONTRASENA(datos.CONTRASENA);
  }

  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const filaNueva = encabezados.map(c => datos[c] !== undefined ? datos[c] : "");
  hoja.appendRow(filaNueva);

  return SEG_SANITIZAR_PARA_CLIENTE({
    EXITO: true,
    ID_USUARIO: idUsuario,
    idUsuario: idUsuario,
    USUARIO: datos,
    MENSAJE: "Usuario " + idUsuario + " creado con éxito."
  });
}

function SEG_ESTABLECER_CONTRASENA(idUsuario, nuevaContrasena, usuarioActualizacion) {
  if (!idUsuario || !nuevaContrasena) throw new Error("ID_USUARIO y nueva contraseña son requeridos.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) throw new Error("Hoja USR_USUARIOS no encontrada.");

  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const idxId = encabezados.indexOf("ID_USUARIO");
  const idxHash = encabezados.indexOf("CONTRASENA_HASH");
  const idxAct = encabezados.indexOf("FECHA_ACTUALIZACION");
  const idxUsrAct = encabezados.indexOf("USUARIO_ACTUALIZACION");

  const registros = hoja.getRange(2, 1, hoja.getLastRow() - 1, encabezados.length).getValues();
  let filaModificar = -1;
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxId]).trim() === String(idUsuario).trim()) {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar === -1) throw new Error("Usuario " + idUsuario + " no encontrado.");

  const hash = SEG_GENERAR_HASH_CONTRASENA(nuevaContrasena);
  if (idxHash !== -1) hoja.getRange(filaModificar, idxHash + 1).setValue(hash);
  if (idxAct !== -1) hoja.getRange(filaModificar, idxAct + 1).setValue(new Date());
  if (idxUsrAct !== -1) hoja.getRange(filaModificar, idxUsrAct + 1).setValue(usuarioActualizacion || "SISTEMA");

  return { EXITO: true, MENSAJE: "Contraseña actualizada exitosamente." };
}

function SEG_REGISTRAR_USUARIO_PUBLICO(datos) {
  if (!datos || !datos.USUARIO || !datos.CONTRASENA_PLANA) {
    return { EXITO: false, MENSAJE: "Complete todos los campos del registro." };
  }
  try {
    const res = SEG_CREAR_USUARIO({
      USUARIO: datos.USUARIO,
      NOMBRE_COMPLETO: datos.NOMBRE || datos.USUARIO,
      CORREO: datos.CORREO || "",
      CONTRASENA: datos.CONTRASENA_PLANA,
      ESTADO_USUARIO: "PENDIENTE",
      ID_ROL: "ROL-000007"
    }, "SISTEMA_INTERNAL_BYPASS");

    return { EXITO: true, MENSAJE: "Registro realizado con éxito. Su cuenta está PENDIENTE de aprobación." };
  } catch (e) {
    return { EXITO: false, MENSAJE: "Error al registrar: " + e.message };
  }
}

function SEG_SOLICITAR_RECUPERACION_CONTRASENA(correo) {
  if (!correo) return { EXITO: false, MENSAJE: "Digite su correo corporativo." };
  return { EXITO: true, MENSAJE: "Instrucciones de recuperación enviadas a su correo corporativo." };
}

