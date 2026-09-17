/**************************************************************
 * 23_SEGURIDAD.gs (VERSIÓN 28.0 - V2 ERP - LIBRO 1 - PARCHE DE SEGURIDAD)
 * RESPONSABILIDAD:
 * - Administrar el ciclo de vida de Usuarios, Roles, Permisos, Sesiones y Auditoría.
 * - Proteger las macros y Web Apps mediante un Sistema de Control de Acceso REAL
 *   (valida rol + módulo + acción contra USR_PERMISOS, no un bypass ciego).
 * - Encriptar contraseñas mediante Hash SHA-256 + SALT único por usuario.
 * - Bloquear usuarios tras exceder el número máximo de intentos fallidos.
 * - Soportar la actualización dinámica de permisos por rol o usuario desde la Web App.
 * - Garantizar el cierre de sesión seguro y la revocación de tokens.
 *
 * CAMBIOS RESPECTO A LA VERSIÓN ANTERIOR (27.0):
 * 1) SEG_VERIFICAR_CONTEXTO_Y_ACCESO ya NO otorga ADMINISTRADOR automático a
 *    cualquiera que abra el Sheet. Ahora resuelve la identidad real por el
 *    correo de Google (Session.getActiveUser) y valida su rol y permisos
 *    reales contra USR_PERMISOS. Si el usuario no está registrado, sólo
 *    obtiene acceso de LECTURA (VER).
 * 2) Nueva función SEG_TIENE_PERMISO(idRol, modulo, accion) que consulta
 *    USR_PERMISOS de verdad (antes SEG_VALIDAR_ACCESO ignoraba modulo/accion).
 * 3) Contraseñas: SEG_GENERAR_HASH_CONTRASENA ahora usa un SALT único por
 *    usuario (columna SALT, que antes existía en la hoja pero nunca se usaba).
 * 4) SEG_AUTENTICAR_USUARIO ahora incrementa INTENTOS_FALLIDOS y bloquea al
 *    usuario tras MAXIMO_INTENTOS_LOGIN, respetando BLOQUEO_USUARIO_MINUTOS.
 * 5) El usuario ADMIN inicial ya no se crea con password fija "Admin123!":
 *    se genera una contraseña temporal aleatoria que se muestra UNA sola vez
 *    y se obliga a cambiarla en el primer login (CAMBIO_CONTRASENA = "SI").
 * 6) Nueva función SEG_SEMBRAR_MATRIZ_PERMISOS_DEFECTO() para poblar
 *    USR_PERMISOS con una matriz de permisos razonable por rol, de forma
 *    que el sistema no quede "todo denegado" tras aplicar este parche.
 *    EJECUTAR ESTA FUNCIÓN UNA SOLA VEZ DESPUÉS DE INSTALAR EL PARCHE.
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
  ROL_CONSULTA_DEFECTO: "ROL-000007", // rol de solo lectura para identidades no registradas
  REGISTRAR_AUDITORIA: true
};

/* ============================================================
 * UTILIDADES DE HOJA (sin cambios respecto a la versión original)
 * ============================================================ */

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

/* ============================================================
 * CONTROL DE ACCESO REAL (NÚCLEO DEL PARCHE)
 * ============================================================ */

/**
 * Punto único de entrada usado por TODOS los módulos (INVENTARIO, VENTAS,
 * COMPRAS, TESORERIA, NOMINA, etc.) antes de ejecutar cualquier operación.
 *
 * Resuelve la identidad real del usuario (por token de sesión Web, o por
 * el correo de Google si se ejecuta dentro de Sheets) y valida su permiso
 * real contra USR_PERMISOS. YA NO otorga ADMINISTRADOR automático.
 */
function SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, modulo, accion) {
  let usuario = null;
  let esContextoSheetsSinRegistro = false;

  if (tokenSesion === "SISTEMA_INTERNAL_BYPASS") {
    // Reservado exclusivamente para procesos 100% internos del servidor
    // (triggers programados, instalador). No debe usarse desde la Web App.
    return {
      AUTORIZADO: true,
      CODIGO: "SISTEMA_BYPASS",
      USUARIO: "SISTEMA",
      ID_USUARIO: "SISTEMA",
      ROL: SEG_CONFIG.ROL_ADMINISTRADOR
    };
  }

  if (tokenSesion === "SHEETS_CONTEXT") {
    // Ejecutado desde dentro del Sheet (menú/diálogo). Ya NO se asume admin:
    // se resuelve la identidad real por el correo de Google del editor.
    let email = "";
    try { email = Session.getActiveUser().getEmail(); } catch (e) { email = ""; }

    usuario = email ? SEG_BUSCAR_USUARIO_LOGIN(email) : null;

    if (!usuario) {
      esContextoSheetsSinRegistro = true;
    }
  } else {
    if (!tokenSesion) throw new Error("ACCESO DENEGADO [TOKEN_REQUERIDO]: Se requiere token de sesión.");
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || !validacion.VALIDA) {
      throw new Error("ACCESO DENEGADO [SESION_INVALIDA]: " + (validacion ? validacion.MENSAJE : "Sesión inválida o expirada."));
    }
    usuario = SEG_CONSULTAR_USUARIO(validacion.SESION.ID_USUARIO);
    if (!usuario) throw new Error("ACCESO DENEGADO [USUARIO_NO_ENCONTRADO]: El usuario de la sesión ya no existe.");
  }

  // Identidad de Sheets no registrada en USR_USUARIOS: acceso de solo lectura.
  if (esContextoSheetsSinRegistro) {
    if (String(accion || "").toUpperCase() !== "VER") {
      SEG_REGISTRAR_AUDITORIA({
        USUARIO: "DESCONOCIDO_SHEETS", MODULO: modulo, ACCION: accion,
        DESCRIPCION: "Intento de escritura desde Sheets por identidad no registrada", RESULTADO: "DENEGADO"
      });
      throw new Error("ACCESO DENEGADO [USUARIO_NO_REGISTRADO]: Tu cuenta de Google no está registrada en USR_USUARIOS. Solo tienes acceso de lectura hasta que un administrador te cree un usuario.");
    }
    return {
      AUTORIZADO: true,
      CODIGO: "SHEETS_CONTEXT_SOLO_LECTURA",
      USUARIO: "DESCONOCIDO_SHEETS",
      ID_USUARIO: "",
      ROL: SEG_CONFIG.ROL_CONSULTA_DEFECTO
    };
  }

  // Validaciones de estado de la cuenta
  const estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();
  if (estado === SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO) {
    throw new Error("ACCESO DENEGADO [USUARIO_BLOQUEADO]: El usuario se encuentra bloqueado.");
  }
  if (estado !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO && tokenSesion !== "SHEETS_CONTEXT") {
    throw new Error("ACCESO DENEGADO [USUARIO_INACTIVO]: El usuario no está activo.");
  }

  // Validación real de permiso por rol/módulo/acción
  const tienePermiso = SEG_TIENE_PERMISO(usuario.ID_ROL, modulo, accion);
  if (!tienePermiso) {
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, MODULO: modulo, ACCION: accion,
      DESCRIPCION: "Intento de acceso denegado por falta de permiso", RESULTADO: "DENEGADO"
    });
    throw new Error("ACCESO DENEGADO [SIN_PERMISO]: Tu rol no tiene permiso para '" + accion + "' en el módulo '" + modulo + "'.");
  }

  return {
    AUTORIZADO: true,
    CODIGO: "AUTORIZADO",
    USUARIO: usuario.USUARIO,
    ID_USUARIO: usuario.ID_USUARIO,
    ROL: usuario.ID_ROL
  };
}

/**
 * Consulta real de permisos: ADMINISTRADOR siempre pasa; el resto de roles
 * necesita una fila explícita en USR_PERMISOS con PERMITIDO = SI/SÍ/TRUE.
 * Deniega por defecto si no hay ninguna fila que coincida (fail-safe).
 */
function SEG_TIENE_PERMISO(idRol, modulo, accion) {
  if (!idRol) return false;

  const rol = SEG_CONSULTAR_ROL(idRol);
  if (rol && String(rol.NOMBRE_ROL || "").trim().toUpperCase() === SEG_CONFIG.ROL_ADMINISTRADOR) {
    return true;
  }
  if (!modulo || !accion) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
  if (!hoja || hoja.getLastRow() < 2) return false;

  const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
  const idxRol = enc.indexOf("ID_ROL");
  const idxMod = enc.indexOf("MODULO");
  const idxAcc = enc.indexOf("ACCION");
  const idxPermitido = enc.indexOf("PERMITIDO");
  if (idxRol === -1 || idxMod === -1 || idxAcc === -1 || idxPermitido === -1) return false;

  const regs = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
  const modUpper = String(modulo).trim().toUpperCase();
  const accUpper = String(accion).trim().toUpperCase();

  for (let i = 0; i < regs.length; i++) {
    if (String(regs[i][idxRol]).trim() === String(idRol).trim() &&
        String(regs[i][idxMod]).trim().toUpperCase() === modUpper &&
        String(regs[i][idxAcc]).trim().toUpperCase() === accUpper) {
      const val = String(regs[i][idxPermitido]).trim().toUpperCase();
      return (val === "SI" || val === "SÍ" || val === "TRUE" || val === "1");
    }
  }
  return false; // deny by default: si no hay regla explícita, no se autoriza
}

/* ============================================================
 * CONTRASEÑAS CON SALT
 * ============================================================ */

function SEG_GENERAR_SALT() {
  return Utilities.getUuid().replace(/-/g, "");
}

/**
 * Genera el hash SHA-256 de (salt + contraseña). Si no se pasa salt,
 * se mantiene compatibilidad hacia atrás (hash sin salt), pero TODO
 * usuario nuevo o reseteado a partir de este parche SIEMPRE debe tener
 * salt asignado.
 */
function SEG_GENERAR_HASH_CONTRASENA(texto, salt) {
  if (!texto) return "";
  const combinado = String(salt || "") + String(texto);
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combinado, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function SEG_VALIDAR_CONTRASENA(usuario, contrasena) {
  if (!usuario || !usuario.CONTRASENA_HASH || String(usuario.CONTRASENA_HASH).trim() === "") return false;
  const hashCalculado = SEG_GENERAR_HASH_CONTRASENA(contrasena, usuario.SALT);
  return String(hashCalculado) === String(usuario.CONTRASENA_HASH);
}

/* ============================================================
 * LOGIN / SESIONES
 * ============================================================ */

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

/**
 * Autentica al usuario. A diferencia de la versión anterior, ahora:
 * - Incrementa INTENTOS_FALLIDOS en cada intento incorrecto.
 * - Bloquea al usuario (ESTADO_USUARIO = BLOQUEADO, BLOQUEADO_HASTA)
 *   tras alcanzar SEG_CONFIG.MAXIMO_INTENTOS_LOGIN.
 * - Libera el bloqueo automáticamente al vencer BLOQUEO_USUARIO_MINUTOS.
 * - Resetea INTENTOS_FALLIDOS a 0 tras un login exitoso.
 */
function SEG_AUTENTICAR_USUARIO(credencial, contrasena) {
  SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS();

  const usuario = SEG_BUSCAR_USUARIO_LOGIN(credencial);
  if (!usuario) return { EXITO: false, CODIGO: "USUARIO_NO_ENCONTRADO", MENSAJE: "Credenciales incorrectas." };

  let estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();

  // Si estaba bloqueado, revisar si el bloqueo ya venció
  if (estado === SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO) {
    const bloqueadoHasta = usuario.BLOQUEADO_HASTA ? new Date(usuario.BLOQUEADO_HASTA) : null;
    if (bloqueadoHasta && bloqueadoHasta.getTime() > Date.now()) {
      const minutosRestantes = Math.ceil((bloqueadoHasta.getTime() - Date.now()) / 60000);
      return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "Usuario bloqueado. Intenta de nuevo en " + minutosRestantes + " minuto(s)." };
    } else {
      // Vencido el bloqueo: se libera automáticamente
      SEG_ACTUALIZAR_CAMPOS_USUARIO(usuario.ID_USUARIO, { ESTADO_USUARIO: SEG_CONFIG.ESTADO_USUARIO_ACTIVO, INTENTOS_FALLIDOS: 0, BLOQUEADO_HASTA: "" });
      estado = SEG_CONFIG.ESTADO_USUARIO_ACTIVO;
    }
  }

  if (estado === "PENDIENTE") return { EXITO: false, CODIGO: "USUARIO_PENDIENTE", MENSAJE: "Su cuenta está pendiente de aprobación." };
  if (estado !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) return { EXITO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "Usuario inactivo en el sistema." };

  if (!SEG_VALIDAR_CONTRASENA(usuario, contrasena)) {
    const intentosPrevios = Number(usuario.INTENTOS_FALLIDOS || 0);
    const nuevosIntentos = intentosPrevios + 1;

    if (nuevosIntentos >= SEG_CONFIG.MAXIMO_INTENTOS_LOGIN) {
      const bloqueadoHasta = new Date(Date.now() + SEG_CONFIG.BLOQUEO_USUARIO_MINUTOS * 60000);
      SEG_ACTUALIZAR_CAMPOS_USUARIO(usuario.ID_USUARIO, {
        INTENTOS_FALLIDOS: nuevosIntentos,
        ESTADO_USUARIO: SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO,
        BLOQUEADO_HASTA: bloqueadoHasta
      });
      SEG_REGISTRAR_AUDITORIA({ ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, MODULO: "SEGURIDAD", ACCION: "LOGIN", DESCRIPCION: "Usuario bloqueado por exceso de intentos fallidos", RESULTADO: "BLOQUEADO" });
      return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "Demasiados intentos fallidos. Usuario bloqueado por " + SEG_CONFIG.BLOQUEO_USUARIO_MINUTOS + " minutos." };
    }

    SEG_ACTUALIZAR_CAMPOS_USUARIO(usuario.ID_USUARIO, { INTENTOS_FALLIDOS: nuevosIntentos });
    return { EXITO: false, CODIGO: "CONTRASENA_INCORRECTA", MENSAJE: "Credenciales incorrectas (" + (SEG_CONFIG.MAXIMO_INTENTOS_LOGIN - nuevosIntentos) + " intento(s) restante(s))." };
  }

  // Login correcto: resetear contador de intentos y actualizar último acceso
  SEG_ACTUALIZAR_CAMPOS_USUARIO(usuario.ID_USUARIO, { INTENTOS_FALLIDOS: 0, ULTIMO_ACCESO: new Date() });

  const resSesion = SEG_CREAR_SESION(usuario.ID_USUARIO);
  SEG_REGISTRAR_AUDITORIA({ ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, MODULO: "SEGURIDAD", ACCION: "LOGIN", DESCRIPCION: "Autenticación exitosa", RESULTADO: "EXITOSO" });

  return SEG_SANITIZAR_PARA_CLIENTE({
    EXITO: true,
    CODIGO: "AUTENTICACION_CORRECTA",
    MENSAJE: "Autenticación correcta.",
    REQUIERE_CAMBIO_CONTRASENA: String(usuario.CAMBIO_CONTRASENA || "").toUpperCase() === "SI",
    USUARIO: { ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, NOMBRE: usuario.NOMBRE || usuario.NOMBRE_COMPLETO || "", CORREO: usuario.CORREO, ID_ROL: usuario.ID_ROL },
    SESION: resSesion
  });
}

/**
 * Actualiza campos puntuales de una fila de USR_USUARIOS sin reescribir
 * toda la fila (evita pisar otros campos concurrentemente).
 */
function SEG_ACTUALIZAR_CAMPOS_USUARIO(idUsuario, campos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
  if (!hoja || hoja.getLastRow() < 2) return false;
  const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const idxId = enc.indexOf("ID_USUARIO");
  const ultimaFila = hoja.getLastRow();
  const regs = hoja.getRange(2, 1, ultimaFila - 1, enc.length).getValues();
  for (let i = 0; i < regs.length; i++) {
    if (String(regs[i][idxId]).trim() === String(idUsuario).trim()) {
      const filaReal = i + 2;
      Object.keys(campos).forEach(campo => {
        const idxCampo = enc.indexOf(campo);
        if (idxCampo !== -1) {
          hoja.getRange(filaReal, idxCampo + 1).setValue(campos[campo]);
        }
      });
      return true;
    }
  }
  return false;
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
  const idRol = usuario ? usuario.ID_ROL : SEG_CONFIG.ROL_CONSULTA_DEFECTO;
  hoja.appendRow([idSesion, idUsuario, nomUsuario, idRol, ahora, ahora, "", "ACTIVA", "WEB", "APLICACION_WEB", token, expira, "", "Inicio exitoso"]);
  return { EXITO: true, ID_SESION: idSesion, TOKEN_SESION: token, ID_USUARIO: idUsuario, FECHA_EXPIRACION: expira };
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

/**
 * Valida una sesión. El contexto SHEETS_CONTEXT ya NO se marca como
 * ADMINISTRADOR aquí; la resolución de rol real ocurre en
 * SEG_VERIFICAR_CONTEXTO_Y_ACCESO. Esta función se mantiene para consultas
 * rápidas de validez de sesión Web.
 */
function SEG_VALIDAR_SESION(tokenSesion, esPoller) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return { VALIDA: false, CODIGO: "SESION_NO_ENCONTRADA", MENSAJE: "La sesión no existe." };
  const estado = String(sesion.ESTADO_SESION || "").trim().toUpperCase();
  if (estado !== SEG_CONFIG.ESTADO_SESION_ACTIVA) return { VALIDA: false, CODIGO: "SESION_NO_ACTIVA", MENSAJE: "La sesión no se encuentra activa." };
  const expira = sesion.EXPIRA_SESION ? new Date(sesion.EXPIRA_SESION) : null;
  if (expira && expira.getTime() < Date.now()) {
    return { VALIDA: false, CODIGO: "SESION_EXPIRADA", MENSAJE: "La sesión ha expirado." };
  }
  return { VALIDA: true, CODIGO: "SESION_VALIDA", MENSAJE: "Sesión autorizada.", SESION: sesion };
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
      const idxCierre = enc.indexOf("FECHA_CIERRE");
      if (hoja && idxEstado !== -1) {
        hoja.getRange(sesion._FILA, idxEstado + 1).setValue("CERRADA");
      }
      if (hoja && idxCierre !== -1) {
        hoja.getRange(sesion._FILA, idxCierre + 1).setValue(new Date());
      }
    }
  } catch (e) {
    if (typeof LOG_REGISTRAR_ERROR === "function") LOG_REGISTRAR_ERROR("SEG_CERRAR_SESION", "SEGURIDAD", e);
  }
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

/* ============================================================
 * CONSULTAS WEB (protegidas con el control de acceso real)
 * ============================================================ */

function SEG_LISTAR_USUARIOS(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
    const regs = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_USUARIOS);
    const lista = regs.map(r => {
      const obj = SEG_CONVERTIR_FILA_OBJETO(enc, r);
      delete obj.CONTRASENA_HASH; // nunca exponer el hash al cliente
      delete obj.SALT;
      return obj;
    });
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

/**
 * Mantenido por compatibilidad con quien ya llame a SEG_VALIDAR_ACCESO
 * directamente. Ahora sí valida módulo/acción de verdad, delegando en
 * SEG_TIENE_PERMISO.
 */
function SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || validacion.VALIDA !== true) {
    return { AUTORIZADO: false, CODIGO: "SESION_INVALIDA", MENSAJE: "No autorizado." };
  }
  const idRol = validacion.SESION ? validacion.SESION.ID_ROL : SEG_CONFIG.ROL_CONSULTA_DEFECTO;
  const permitido = SEG_TIENE_PERMISO(idRol, modulo, accion);
  if (!permitido) {
    return { AUTORIZADO: false, CODIGO: "SIN_PERMISO", MENSAJE: "Tu rol no tiene permiso para esta acción." };
  }
  return { AUTORIZADO: true, CODIGO: "AUTORIZADO", USUARIO: validacion.SESION.USUARIO };
}

function SEG_OBTENER_CONTEXTO_SEGURIDAD(tokenSesion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || validacion.VALIDA !== true) return { VALIDO: false, MENSAJE: "No autorizado." };
  const idRol = validacion.SESION ? validacion.SESION.ID_ROL : SEG_CONFIG.ROL_CONSULTA_DEFECTO;
  const rol = SEG_CONSULTAR_ROL(idRol) || { NOMBRE_ROL: "CONSULTA" };
  return { VALIDO: true, USUARIO: validacion.SESION.USUARIO, ROL: rol };
}

/**
 * Devuelve solo los módulos donde el rol real del usuario tiene al menos
 * permiso de VER. ADMINISTRADOR sigue viendo todo el menú.
 */
function SEG_OBTENER_MENU_NIVEL(tokenSesion) {
  try {
    const validacion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacion || !validacion.VALIDA) return { EXITO: false, MENSAJE: "Sesión inválida." };

    const idRol = validacion.SESION.ID_ROL;
    const rol = SEG_CONSULTAR_ROL(idRol) || {};
    const esAdmin = String(rol.NOMBRE_ROL || "").trim().toUpperCase() === SEG_CONFIG.ROL_ADMINISTRADOR;

    const modulosMaestrosTodos = ["CLIENTES", "VENTAS", "COMPRAS", "PRODUCTOS", "INVENTARIO", "FINANZAS", "PLANEACION", "OBRAS", "NOMINA", "SEGURIDAD"];

    if (esAdmin) {
      return { EXITO: true, ES_ADMIN: true, ROL: rol.NOMBRE_ROL, MODULOS: modulosMaestrosTodos };
    }

    const modulosPermitidos = modulosMaestrosTodos.filter(m => SEG_TIENE_PERMISO(idRol, m, "VER"));
    return { EXITO: true, ES_ADMIN: false, ROL: rol.NOMBRE_ROL || "SIN_ROL", MODULOS: modulosPermitidos };
  } catch (e) {
    return { EXITO: false, MENSAJE: e.toString() };
  }
}

function SEG_GUARDAR_PERMISO_WEB(datos, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja) {
      hoja = ss.insertSheet(SEG_CONFIG.HOJA_PERMISOS);
      hoja.appendRow(["ID_PERMISO", "ID_ROL", "MODULO", "SUBMODULO", "ACCION", "PERMITIDO", "ESTADO_PERMISO", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"]);
    }
    const ahora = new Date();
    const idPermiso = SEG_CONFIG.PREFIJO_PERMISO + "-" + String(Math.max(1, hoja.getLastRow())).padStart(SEG_CONFIG.DIGITOS_ID, "0");
    hoja.appendRow([
      idPermiso, datos.ID_ROL || "", datos.MODULO || "", datos.SUBMODULO || "",
      datos.ACCION || "", datos.PERMITIDO ? "SI" : "NO", "ACTIVO",
      ahora, ahora, "WEB", "WEB", datos.OBSERVACIONES || ""
    ]);
    return { EXITO: true, MENSAJE: "Permiso guardado exitosamente." };
  } catch (e) {
    return { EXITO: false, MENSAJE: e.toString() };
  }
}

/* ============================================================
 * INICIALIZACIÓN / SEMBRADO DE DATOS POR DEFECTO
 * ============================================================ */

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
    const rolesDefecto = [
      ["ROL-000001", "ADMINISTRADOR", "Administración general y acceso total al ERP", 100, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol maestro protegido"],
      ["ROL-000002", "CONTADOR", "Gestión y supervisión de procesos contables y financieros", 80, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a reportes contables"],
      ["ROL-000003", "AUXILIAR_CONTABLE", "Apoyo en registros y procesos contables autorizados", 60, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Permiso de registro operativo"],
      ["ROL-000004", "TESORERO", "Gestión de tesorería, pagos, recaudos y conciliaciones autorizadas", 70, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a bancos y cajas"],
      ["ROL-000005", "COMERCIAL", "Gestión de clientes y operaciones comerciales autorizadas", 50, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso limitado a ventas"],
      ["ROL-000006", "OPERATIVO", "Registro de operaciones asignadas", 40, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol operativo general"],
      ["ROL-000007", "CONSULTA", "Acceso exclusivamente de lectura general", 10, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol de sólo lectura"]
    ];
    rolesDefecto.forEach(rol => hoja.appendRow(rol));
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

/**
 * *** EJECUTAR MANUALMENTE UNA SOLA VEZ DESPUÉS DE INSTALAR ESTE PARCHE ***
 * (Editor de Apps Script -> seleccionar esta función -> Ejecutar)
 *
 * Sin esta matriz, como SEG_TIENE_PERMISO ahora deniega por defecto,
 * TODOS los roles distintos de ADMINISTRADOR quedarían sin acceso a nada
 * hasta que existan filas explícitas en USR_PERMISOS. Esta función siembra
 * una matriz de permisos razonable de partida; ajústala luego desde la
 * pantalla de Seguridad según las necesidades reales de tu operación.
 */
function SEG_SEMBRAR_MATRIZ_PERMISOS_DEFECTO() {
  SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
  SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
  if (hoja.getLastRow() >= 2) {
    console.log("USR_PERMISOS ya tiene datos, no se sobreescribe. Bórralos manualmente si quieres resembrar.");
    return;
  }

  // matriz: [ID_ROL, MODULO, ACCION, PERMITIDO]
  const M = {
    CONTADOR: {
      FINANZAS: ["VER", "CREAR", "EDITAR"], VENTAS: ["VER", "CREAR", "EDITAR"], COMPRAS: ["VER", "CREAR", "EDITAR"],
      INVENTARIO: ["VER", "CREAR", "EDITAR"], NOMINA: ["VER", "CREAR", "EDITAR"], PLANEACION: ["VER", "CREAR", "EDITAR"],
      CLIENTES: ["VER"], PRODUCTOS: ["VER"], OBRAS: ["VER"], SEGURIDAD: ["VER"]
    },
    AUXILIAR_CONTABLE: {
      VENTAS: ["VER", "CREAR"], COMPRAS: ["VER", "CREAR"], INVENTARIO: ["VER", "CREAR"], FINANZAS: ["VER", "CREAR"],
      CLIENTES: ["VER"], PRODUCTOS: ["VER"], OBRAS: ["VER"]
    },
    TESORERO: {
      FINANZAS: ["VER", "CREAR", "EDITAR"], VENTAS: ["VER"], COMPRAS: ["VER"], CLIENTES: ["VER"]
    },
    COMERCIAL: {
      CLIENTES: ["VER", "CREAR", "EDITAR"], VENTAS: ["VER", "CREAR", "EDITAR"], PRODUCTOS: ["VER"], INVENTARIO: ["VER"]
    },
    OPERATIVO: {
      OBRAS: ["VER", "CREAR"], INVENTARIO: ["VER", "CREAR"], PRODUCTOS: ["VER"], CLIENTES: ["VER"]
    },
    CONSULTA: {
      CLIENTES: ["VER"], VENTAS: ["VER"], COMPRAS: ["VER"], PRODUCTOS: ["VER"], INVENTARIO: ["VER"],
      FINANZAS: ["VER"], PLANEACION: ["VER"], OBRAS: ["VER"], NOMINA: ["VER"], SEGURIDAD: ["VER"]
    }
  };

  const nombreARol = {};
  SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_ROLES).forEach(fila => {
    const enc = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
    const obj = SEG_CONVERTIR_FILA_OBJETO(enc, fila);
    nombreARol[obj.NOMBRE_ROL] = obj.ID_ROL;
  });

  const ahora = new Date();
  let contador = 1;
  Object.keys(M).forEach(nombreRol => {
    const idRol = nombreARol[nombreRol];
    if (!idRol) return;
    Object.keys(M[nombreRol]).forEach(modulo => {
      M[nombreRol][modulo].forEach(accion => {
        const idPermiso = SEG_CONFIG.PREFIJO_PERMISO + "-" + String(contador++).padStart(SEG_CONFIG.DIGITOS_ID, "0");
        hoja.appendRow([idPermiso, idRol, modulo, "", accion, "SI", "ACTIVO", ahora, ahora, "SISTEMA", "SISTEMA", "Sembrado por defecto"]);
      });
    });
  });

  console.log("✓ Matriz de permisos por defecto sembrada en USR_PERMISOS. Revísala y ajústala en la pantalla de Seguridad.");
}

/**
 * Crea el usuario ADMIN inicial con SALT real y una contraseña temporal
 * ALEATORIA (ya no la fija "Admin123!"). La contraseña se muestra una sola
 * vez en el log/alerta de instalación y se obliga a cambiarla en el primer
 * login (columna CAMBIO_CONTRASENA = "SI", que la pantalla de login debe
 * respetar redirigiendo a un formulario de cambio de contraseña).
 */
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
    const salt = SEG_GENERAR_SALT();
    const passwordTemporal = SEG_GENERAR_PASSWORD_TEMPORAL();
    const hashAdmin = SEG_GENERAR_HASH_CONTRASENA(passwordTemporal, salt);

    hoja.appendRow(["USR-000001", "ADMIN", "Administrador del ERP", "admin@megudan.com", hashAdmin, salt, "ROL-000001", "ACTIVO", 0, "", ahora, "SI", ahora, ahora, "SISTEMA", "SISTEMA", "Usuario Inicial - password temporal, debe cambiarse"]);

    const mensaje = "USUARIO ADMIN CREADO.\nUsuario: ADMIN\nContraseña temporal: " + passwordTemporal + "\n\n" +
      "⚠️ Guarda esta contraseña ahora: no volverá a mostrarse. Deberás cambiarla en el primer inicio de sesión.";
    console.log(mensaje);
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Usuario administrador creado", mensaje, ui.ButtonSet.OK);
    } catch (e) { /* contexto headless, ya quedó en el log */ }
  } catch (e) {
    if (typeof LOG_REGISTRAR_ERROR === "function") LOG_REGISTRAR_ERROR("SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS", "SEGURIDAD", e);
  }
}

function SEG_GENERAR_PASSWORD_TEMPORAL() {
  const caracteres = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%";
  let pass = "";
  for (let i = 0; i < 14; i++) {
    pass += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }
  return pass;
}

/**
 * RPC para cambio de contraseña (requerido tras el primer login, o a
 * petición del usuario). Genera un nuevo salt en cada cambio.
 */
function SEG_CAMBIAR_CONTRASENA(idUsuario, contrasenaActual, contrasenaNueva, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER"); // solo exige sesión válida, no permiso especial
    const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
    if (!usuario) throw new Error("Usuario no encontrado.");
    if (!SEG_VALIDAR_CONTRASENA(usuario, contrasenaActual)) throw new Error("La contraseña actual no es correcta.");
    if (!contrasenaNueva || String(contrasenaNueva).length < 8) throw new Error("La nueva contraseña debe tener al menos 8 caracteres.");

    const nuevoSalt = SEG_GENERAR_SALT();
    const nuevoHash = SEG_GENERAR_HASH_CONTRASENA(contrasenaNueva, nuevoSalt);
    SEG_ACTUALIZAR_CAMPOS_USUARIO(idUsuario, {
      CONTRASENA_HASH: nuevoHash, SALT: nuevoSalt, CAMBIO_CONTRASENA: "NO", FECHA_ACTUALIZACION: new Date()
    });
    SEG_REGISTRAR_AUDITORIA({ ID_USUARIO: idUsuario, USUARIO: usuario.USUARIO, MODULO: "SEGURIDAD", ACCION: "CAMBIO_CONTRASENA", DESCRIPCION: "Contraseña actualizada por el usuario", RESULTADO: "EXITOSO" });
    return { EXITO: true, MENSAJE: "Contraseña actualizada correctamente." };
  } catch (e) {
    return { EXITO: false, MENSAJE: e.toString() };
  }
}

/* ============================================================
 * AUDITORÍA
 * ============================================================ */

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
  } catch (e) {
    // Si la auditoría misma falla, al menos déjalo en el log de ejecución
    console.error("Fallo registrando auditoría: " + e);
  }
}

function SEG_SANITIZAR_PARA_CLIENTE(dato) {
  if (dato === null || dato === undefined) return dato;
  return JSON.parse(JSON.stringify(dato));
}

