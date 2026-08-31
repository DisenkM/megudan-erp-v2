/**************************************************************
* 23_SEGURIDAD.gs (VERSIÓN 10.0 - MEGUDAN ERP V2)
* RESPONSABILIDAD:
* - Administrar el ciclo de vida de Usuarios, Roles, Permisos, Sesiones y Auditoría.
* - Proteger las macros y Web Apps mediante un Sistema de Control de Acceso Dual.
* - Encriptar contraseñas mediante Hash SHA-256 y proveer auto-inicialización de roles/permisos.
* - Soportar recuperación de claves por correo y subida segura de logotipos.
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

function SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, modulo, accion) {
  if (tokenSesion === "SISTEMA_INTERNAL_BYPASS" || tokenSesion === "SHEETS_CONTEXT") {
    return {
      AUTORIZADO: true,
      CODIGO: "SISTEMA_BYPASS",
      USUARIO: "SISTEMA",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso concedido automáticamente para operaciones internas o locales del sistema."
    };
  }
  try {
    SpreadsheetApp.getUi();
    return {
      AUTORIZADO: true,
      CODIGO: "CONTEXTO_SHEETS_TRUSTED",
      USUARIO: "ADMINISTRADOR_LOCAL_SHEETS",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso concedido automáticamente en entorno confiable de Google Sheets."
    };
  } catch (uiError) {
    if (!tokenSesion || String(tokenSesion).trim() === "") {
      throw new Error("ACCESO DENEGADO [TOKEN_REQUERIDO]: Se requiere un token de sesión activo para operar desde la Web App.");
    }
    const validacion = SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion);
    if (!validacion || validacion.AUTORIZADO !== true) {
      throw new Error("ACCESO DENEGADO [" + (validacion ? validacion.CODIGO : "ERROR") + "]: " + (validacion ? validacion.MENSAJE : "No autorizado."));
    }
    return validacion;
  }
}

function SEG_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No existe la hoja física '" + nombreHoja + "' en la base de datos.");
  return hoja;
}

function SEG_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaColumna === 0) throw new Error("La hoja '" + nombreHoja + "' no contiene encabezados.");
  return hoja.getRange(1, 1, 1, ultimaColumna).getDisplayValues().map(h => String(h || "").trim().toUpperCase());
}

function SEG_OBTENER_REGISTROS(nombreHoja) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) return [];
  return hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
}

function SEG_NORMALIZAR_TEXT(valor) {
  return String(valor || "").trim().toUpperCase().replace(/\s+/g, " ");
}

function SEG_GENERAR_ID(nombreHoja, campoID, prefijo) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return prefijo + "-" + String(1).padStart(SEG_CONFIG.DIGITOS_ID, "0");

  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const indiceID = encabezados.indexOf(campoID.toUpperCase());
  if (indiceID === -1) throw new Error("No se encontró la columna '" + campoID + "' en " + nombreHoja);

  const registros = hoja.getRange(2, indiceID + 1, ultimaFila - 1, 1).getDisplayValues().flat();
  let numeroMayor = 0;
  registros.forEach(id => {
    const textoID = String(id || "").trim();
    if (textoID.startsWith(prefijo + "-")) {
      const numero = parseInt(textoID.replace(prefijo + "-", ""), 10);
      if (!isNaN(numero) && numero > numeroMayor) numeroMayor = numero;
    }
  });

  return prefijo + "-" + String(numeroMayor + 1).padStart(SEG_CONFIG.DIGITOS_ID, "0");
}

function SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const objeto = {};
  encabezados.forEach((campo, indice) => {
    objeto[campo] = fila[indice] !== undefined ? fila[indice] : "";
  });
  return objeto;
}

function SEG_CONVERTIR_OBJETO_FILA(encabezados, objeto) {
  return encabezados.map(campo => objeto[campo] !== undefined ? objeto[campo] : "");
}

function SEG_BUSCAR_REGISTRO(nombreHoja, campoBusqueda, valorBusqueda) {
  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const registros = SEG_OBTENER_REGISTROS(nombreHoja);
  const idx = encabezados.indexOf(campoBusqueda.toUpperCase());
  if (idx === -1) throw new Error("Campo " + campoBusqueda + " no encontrado en " + nombreHoja);

  const normalizado = SEG_NORMALIZAR_TEXT(valorBusqueda);
  const fila = registros.find(f => SEG_NORMALIZAR_TEXT(f[idx]) === normalizado);
  return fila ? SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) : null;
}

function SEG_BUSCAR_FILA_REGISTRO(nombreHoja, campoBusqueda, valorBusqueda) {
  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const registros = SEG_OBTENER_REGISTROS(nombreHoja);
  const idx = encabezados.indexOf(campoBusqueda.toUpperCase());
  if (idx === -1) throw new Error("Campo " + campoBusqueda + " no encontrado en " + nombreHoja);

  const normalizado = SEG_NORMALIZAR_TEXT(valorBusqueda);
  for (let i = 0; i < registros.length; i++) {
    if (SEG_NORMALIZAR_TEXT(registros[i][idx]) === normalizado) return i + 2;
  }
  return null;
}

function SEG_VALIDAR_OBLIGATORIOS(datos, obligatorios) {
  obligatorios.forEach(campo => {
    const valor = datos[campo];
    if (valor === undefined || valor === null || String(valor).trim() === "") {
      throw new Error("El campo " + campo + " es obligatorio.");
    }
  });
}

function SEG_AHORA() { return new Date(); }

function SEG_CREAR_USUARIO(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos || typeof datos !== "object") throw new Error("Debe proporcionar la información del usuario.");
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);

  let campoNombre = encabezados.indexOf("NOMBRE") !== -1 ? "NOMBRE" : "NOMBRE_COMPLETO";
  if (campoNombre === "NOMBRE_COMPLETO" && datos.NOMBRE_COMPLETO === undefined && datos.NOMBRE !== undefined) {
    datos.NOMBRE_COMPLETO = datos.NOMBRE;
  } else if (campoNombre === "NOMBRE" && datos.NOMBRE === undefined && datos.NOMBRE_COMPLETO !== undefined) {
    datos.NOMBRE = datos.NOMBRE_COMPLETO;
  }

  datos.USUARIO = String(datos.USUARIO || "").trim().toUpperCase().replace(/\s/g, "");
  datos.CORREO = String(datos.CORREO || "").trim().toLowerCase();

  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "USUARIO", datos.USUARIO)) {
    throw new Error("El nombre de usuario '" + datos.USUARIO + "' ya existe.");
  }
  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", datos.CORREO)) {
    throw new Error("El correo '" + datos.CORREO + "' ya está registrado.");
  }

  const ahora = SEG_AHORA();
  const idUsuario = SEG_GENERAR_ID(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", SEG_CONFIG.PREFIJO_USUARIO);

  const usuario = {};
  encabezados.forEach(c => { usuario[c] = datos[c] !== undefined ? datos[c] : ""; });

  usuario.ID_USUARIO = idUsuario;
  usuario.USUARIO = datos.USUARIO;
  usuario[campoNombre] = datos[campoNombre] || datos.NOMBRE || datos.NOMBRE_COMPLETO || "";
  usuario.CORREO = datos.CORREO;
  usuario.ESTADO_USUARIO = datos.ESTADO_USUARIO || SEG_CONFIG.ESTADO_USUARIO_PENDIENTE;
  usuario.INTENTOS_FALLIDOS = 0;
  usuario.FECHA_CREACION = ahora;
  usuario.FECHA_ACTUALIZACION = ahora;
  usuario.USUARIO_CREACION = usuarioEjecutor;
  usuario.USUARIO_ACTUALIZACION = usuarioEjecutor;

  if (datos.CONTRASENA_PLANA) {
    const seguridadValida = SEG_VALIDAR_SEGURIDAD_CONTRASENA(datos.CONTRASENA_PLANA);
    if (!seguridadValida.VALIDA) throw new Error(seguridadValida.MENSAJES.join(" "));
    usuario.CONTRASENA_HASH = SEG_GENERAR_HASH_CONTRASENA(datos.CONTRASENA_PLANA);
    usuario.DEBE_CAMBIAR_CONTRASENA = datos.DEBE_CAMBIAR_CONTRASENA || "NO";
  }

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, usuario);
  hoja.appendRow(nuevaFila);

  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: idUsuario,
    USUARIO: usuario.USUARIO,
    MODULO: "SEGURIDAD",
    SUBMODULO: "USUARIOS",
    ACCION: "CREAR",
    TIPO_REGISTRO: "USUARIOS",
    ID_REGISTRO: idUsuario,
    DESCRIPCION: "Usuario " + usuario.USUARIO + " creado con éxito.",
    RESULTADO: "EXITOSO"
  });

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_USUARIO: idUsuario, USUARIO: usuario });
}

function SEG_CONSULTAR_USUARIO(idUsuario, tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  }
  return SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
}

function SEG_LISTAR_USUARIOS(tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
    const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_USUARIOS);
    const usuarios = registros.filter(r => r && String(r).trim() !== "").map(r => SEG_CONVERTIR_FILA_OBJETO(encabezados, r));
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(usuarios),
      MENSAJE: "Lista de usuarios obtenida exitosamente."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_LISTAR_USUARIOS", "SEGURIDAD", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los usuarios: " + error.toString() };
  }
}

function SEG_ACTUALIZAR_USUARIO(idUsuario, datos, tokenSesion) {
  if (typeof idUsuario === "object" && datos === undefined) {
    datos = idUsuario;
    idUsuario = datos.ID_USUARIO;
  }
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!idUsuario) throw new Error("Debe indicar el ID_USUARIO.");
  const usuarioActual = SEG_CONSULTAR_USUARIO(idUsuario, undefined);
  if (!usuarioActual) throw new Error("No se encontró el usuario '" + idUsuario + "'.");

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const filaUsuario = SEG_BUSCAR_FILA_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);

  const camposProtegidos = ["ID_USUARIO", "FECHA_CREACION", "USUARIO_CREACION"];
  const usuarioActualizado = Object.assign({}, usuarioActual);

  Object.keys(datos).forEach(campo => {
    const colNormalizada = String(campo).trim().toUpperCase();
    if (!camposProtegidos.includes(colNormalizada) && encabezados.indexOf(colNormalizada) !== -1) {
      usuarioActualizado[colNormalizada] = datos[campo];
    }
  });

  if (datos.USUARIO !== undefined) {
    usuarioActualizado.USUARIO = String(datos.USUARIO || "").trim().toUpperCase().replace(/\s/g, "");
    const duplicado = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "USUARIO", usuarioActualizado.USUARIO);
    if (duplicado && duplicado.ID_USUARIO !== idUsuario) throw new Error("Nombre de usuario ocupado.");
  }

  if (datos.CORREO !== undefined) {
    usuarioActualizado.CORREO = String(datos.CORREO || "").trim().toLowerCase();
    const duplicado = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", usuarioActualizado.CORREO);
    if (duplicado && duplicado.ID_USUARIO !== idUsuario) throw new Error("Correo electrónico ocupado.");
  }

  usuarioActualizado.FECHA_ACTUALIZACION = SEG_AHORA();
  usuarioActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;

  const filaActualizada = SEG_CONVERTIR_OBJETO_FILA(encabezados, usuarioActualizado);
  hoja.getRange(filaUsuario, 1, 1, encabezados.length).setValues([filaActualizada]);

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_USUARIO: idUsuario, USUARIO: usuarioActualizado });
}

function SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, nuevoEstado, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  return SEG_ACTUALIZAR_USUARIO(idUsuario, { ESTADO_USUARIO: nuevoEstado.toUpperCase(), USUARIO_ACTUALIZACION: usuarioEjecutor }, tokenSesion);
}

function SEG_ELIMINAR_USUARIO(idUsuario, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ELIMINAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  const res = SEG_ACTUALIZAR_USUARIO(idUsuario, { ESTADO_USUARIO: "ELIMINADO", USUARIO_ACTUALIZACION: usuarioEjecutor }, tokenSesion);
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: idUsuario,
    USUARIO: res.USUARIO.USUARIO,
    MODULO: "SEGURIDAD",
    SUBMODULO: "USUARIOS",
    ACCION: "ELIMINAR",
    TIPO_REGISTRO: "USUARIOS",
    ID_REGISTRO: idUsuario,
    DESCRIPCION: "Usuario " + res.USUARIO.USUARIO + " marcado como ELIMINADO de forma lógica por " + usuarioEjecutor,
    RESULTADO: "EXITOSO"
  });
  return { EXITO: true, MENSAJE: "Usuario marcado como eliminado lógicamente." };
}

function SEG_BLOQUEAR_USUARIO(idUsuario, tokenSesion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO, tokenSesion);
}

function SEG_DESBLOQUEAR_USUARIO(idUsuario, tokenSesion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_ACTIVO, tokenSesion);
}

function SEG_CREAR_ROL(datos, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
  if (!datos || typeof datos !== "object") throw new Error("Debe proporcionar los datos del rol.");
  if (!datos.NOMBRE_ROL) throw new Error("El NOMBRE_ROL es mandatorio.");

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_ROLES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);

  const nombreRol = String(datos.NOMBRE_ROL).trim().toUpperCase();
  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_ROLES, "NOMBRE_ROL", nombreRol)) {
    throw new Error("El rol '" + nombreRol + "' ya existe.");
  }

  const ahora = SEG_AHORA();
  const idRol = SEG_GENERAR_ID(SEG_CONFIG.HOJA_ROLES, "ID_ROL", SEG_CONFIG.PREFIJO_ROL);

  const rol = {};
  encabezados.forEach(c => { rol[c] = datos[c] !== undefined ? datos[c] : ""; });

  rol.ID_ROL = idRol;
  rol.NOMBRE_ROL = nombreRol;
  rol.ESTADO_ROL = datos.ESTADO_ROL || "ACTIVO";
  rol.FECHA_CREACION = ahora;
  rol.FECHA_ACTUALIZACION = ahora;
  rol.USUARIO_CREACION = Session.getActiveUser().getEmail() || "SISTEMA";
  rol.USUARIO_ACTUALIZACION = rol.USUARIO_CREACION;

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, rol);
  hoja.appendRow(nuevaFila);

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_ROL: idRol, ROL: rol });
}

function SEG_CONSULTAR_ROL(idRol) {
  return SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_ROLES, "ID_ROL", idRol);
}

function SEG_LISTAR_ROLES(tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  }
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_ROLES);
  return registros.filter(r => r && String(r).trim() !== "").map(r => SEG_CONVERTIR_FILA_OBJETO(encabezados, r));
}

function SEG_ACTUALIZAR_ROL(idRol, datos, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  if (!idRol || String(idRol).trim() === "") throw new Error("ID_ROL mandatorio.");

  const rolActual = SEG_CONSULTAR_ROL(idRol);
  if (!rolActual) throw new Error("No existe el rol '" + idRol + "'.");

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_ROLES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
  const filaRol = SEG_BUSCAR_FILA_REGISTRO(SEG_CONFIG.HOJA_ROLES, "ID_ROL", idRol);

  const camposProtegidos = ["ID_ROL", "FECHA_CREACION", "USUARIO_CREACION"];
  const rolActualizado = Object.assign({}, rolActual);

  Object.keys(datos).forEach(campo => {
    const colNormalizada = String(campo).trim().toUpperCase();
    if (!camposProtegidos.includes(colNormalizada) && encabezados.indexOf(colNormalizada) !== -1) {
      rolActualizado[colNormalizada] = datos[campo];
    }
  });

  rolActualizado.FECHA_ACTUALIZACION = SEG_AHORA();
  rolActualizado.USUARIO_ACTUALIZACION = Session.getActiveUser().getEmail() || "SISTEMA";

  const filaActualizada = SEG_CONVERTIR_OBJETO_FILA(encabezados, rolActualizado);
  hoja.getRange(filaRol, 1, 1, encabezados.length).setValues([filaActualizada]);

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_ROL: idRol, ROL: rolActualizado });
}

function SEG_CAMBIAR_ESTADO_ROL(idRol, nuevoEstado, tokenSesion) {
  return SEG_ACTUALIZAR_ROL(idRol, { ESTADO_ROL: nuevoEstado }, tokenSesion);
}

function SEG_OBTENER_PERMISOS_ROL(idRol) {
  if (!idRol) throw new Error("Debe indicar el ID_ROL.");
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
  const idxIdRol = encabezados.indexOf("ID_ROL");

  const filtrados = registros.filter(fila => String(fila[idxIdRol] || "").trim().toUpperCase() === String(idRol).trim().toUpperCase());
  return filtrados.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
}

function SEG_BUSCAR_PERMISO(idRol, modulo, accion) {
  const permisos = SEG_OBTENER_PERMISOS_ROL(idRol);
  const moduloNormalizado = SEG_NORMALIZAR_TEXT(modulo);
  const accionNormalizada = SEG_NORMALIZAR_TEXT(accion);
  return permisos.find(p => SEG_NORMALIZAR_TEXT(p.MODULO) === moduloNormalizado && SEG_NORMALIZAR_TEXT(p.ACCION) === accionNormalizada) || null;
}

function SEG_VALIDAR_PERMISO_ROL(idRol, modulo, accion) {
  const rol = SEG_CONSULTAR_ROL(idRol);
  if (!rol || String(rol.ESTADO_ROL || "").toUpperCase() !== "ACTIVO") return false;

  const permiso = SEG_BUSCAR_PERMISO(idRol, modulo, accion);
  if (!permiso || String(permiso.ESTADO_PERMISO || "").toUpperCase() !== "ACTIVO") return false;

  const permitido = String(permiso.PERMITIDO || "").trim().toUpperCase();
  return ["SI", "SÍ", "TRUE", "1", "VERDADERO"].includes(permitido);
}

function SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario, undefined);
  if (!usuario || String(usuario.ESTADO_USUARIO || "").toUpperCase() !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) return false;
  if (!usuario.ID_ROL) return false;
  return SEG_VALIDAR_PERMISO_ROL(usuario.ID_ROL, modulo, accion);
}

function SEG_VERIFICAR_ACCESO(idUsuario, modulo, accion) {
  const authorized = SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion);
  if (!authorized) {
    throw new Error("ACCESO DENEGADO. No tiene autorización para '" + accion + "' en el módulo '" + modulo + "'.");
  }
  return true;
}

function SEG_GENERAR_HASH_CONTRASENA(contrasena) {
  if (contrasena === undefined || contrasena === null || String(contrasena) === "") {
    throw new Error("Debe proporcionar una contraseña para procesar.");
  }
  const bytes = Utilities.newBlob(String(contrasena)).getBytes();
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return hashBytes.map(byte => {
    const valor = byte < 0 ? byte + 256 : byte;
    return ("0" + valor.toString(16)).slice(-2);
  }).join("");
}

function SEG_VALIDAR_SEGURIDAD_CONTRASENA(contrasena) {
  const password = String(contrasena || "");
  const resultado = { VALIDA: false, MENSAJES: [] };
  if (password.length < 8) resultado.MENSAES.push("Mínimo 8 caracteres.");
  if (!/[A-Z]/.test(password)) resultado.MENSAJES.push("Requiere una mayúscula.");
  if (!/[a-z]/.test(password)) resultado.MENSAJES.push("Requiere una minúscula.");
  if (!/[1-9]/.test(password)) resultado.MENSAJES.push("Requiere un número.");
  resultado.VALIDA = resultado.MENSAJES.length === 0;
  return resultado;
}

function SEG_BUSCAR_USUARIO_LOGIN(credencial) {
  if (!credencial || String(credencial).trim() === "") return null;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return null;

  const idxUsuario = encabezados.indexOf("USUARIO");
  const idxCorreo = encabezados.indexOf("CORREO");

  const criterio = String(credencial).trim().toUpperCase();
  const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();

  const fila = registros.find(f => {
    return String(f[idxUsuario] || "").trim().toUpperCase() === criterio ||
           String(f[idxCorreo] || "").trim().toUpperCase() === criterio;
  });

  return fila ? SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) : null;
}

function SEG_ESTABLECER_CONTRASENA(idUsuario, nuevaContrasena, usuarioActualizacion, tokenSesion) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
  if (!usuario) throw new Error("Usuario no encontrado.");

  const validacionSeguridad = SEG_VALIDAR_SEGURIDAD_CONTRASENA(nuevaContrasena);
  if (!validacionSeguridad.VALIDA) throw new Error(validacionSeguridad.MENSAJES.join(" "));

  const hash = SEG_GENERAR_HASH_CONTRASENA(nuevaContrasena);
  const token = tokenSesion || (usuarioActualizacion === "SISTEMA_TEST" || usuarioActualizacion === "SISTEMA" ? "SISTEMA_INTERNAL_BYPASS" : undefined);
  return SEG_ACTUALIZAR_USUARIO(idUsuario, {
    CONTRASENA_HASH: hash,
    FECHA_CAMBIO_CONTRASENA: SEG_AHORA(),
    DEBE_CAMBIAR_CONTRASENA: "NO",
    USUARIO_ACTUALIZACION: usuarioActualizacion || "SISTEMA"
  }, token);
}

function SEG_VALIDAR_CONTRASENA(usuario, contrasena) {
  if (!usuario || !usuario.CONTRASENA_HASH || String(usuario.CONTRASENA_HASH).trim() === "") return false;
  return String(SEG_GENERAR_HASH_CONTRASENA(contrasena)) === String(usuario.CONTRASENA_HASH);
}

function SEG_AUTENTICAR_USUARIO(credencial, contrasena) {
  const usuario = SEG_BUSCAR_USUARIO_LOGIN(credencial);
  if (!usuario) return { EXITO: false, CODIGO: "USUARIO_NO_ENCONTRADO", MENSAJE: "Credenciales incorrectas." };

  const estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();
  if (estado === "PENDIENTE") return { EXITO: false, CODIGO: "USUARIO_PENDIENTE", MENSAJE: "Su cuenta está pendiente de aprobación." };
  if (estado === "BLOQUEADO") return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "Usuario temporalmente bloqueado." };
  if (estado !== "ACTIVO") return { EXITO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "Usuario inactivo en el sistema." };

  if (!SEG_VALIDAR_CONTRASENA(usuario, contrasena)) {
    let fallidos = parseInt(usuario.INTENTOS_FALLIDOS || 0, 10) + 1;
    const updates = { INTENTOS_FALLIDOS: fallidos };
    if (fallidos >= SEG_CONFIG.MAXIMO_INTENTOS_LOGIN) updates.ESTADO_USUARIO = "BLOQUEADO";
    SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, updates, "SISTEMA_INTERNAL_BYPASS");
    return { EXITO: false, CODIGO: "CONTRASENA_INCORRECTA", MENSAJE: "Credenciales incorrectas." };
  }

  SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, { INTENTOS_FALLIDOS: 0, ULTIMO_ACCESO: SEG_AHORA() }, "SISTEMA_INTERNAL_BYPASS");
  return {
    EXITO: true,
    CODIGO: "AUTENTICACION_CORRECTA",
    MENSAJE: "Autenticación correcta.",
    USUARIO: { ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, NOMBRE: usuario.NOMBRE || usuario.NOMBRE_COMPLETO || "", CORREO: usuario.CORREO, ID_ROL: usuario.ID_ROL }
  };
}

function SEG_GENERAR_TOKEN_SESION() { return Utilities.getUuid() + "-" + Utilities.getUuid(); }

function SEG_CREAR_SESION(idUsuario) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
  if (!usuario) throw new Error("Usuario no encontrado.");

  const ahora = SEG_AHORA();
  const fechaExp = new Date(ahora.getTime() + (SEG_CONFIG.DURACION_SESION_HORAS * 60 * 60 * 1000));
  const idSesion = SEG_GENERAR_ID(SEG_CONFIG.HOJA_SESIONES, "ID_SESION", SEG_CONFIG.PREFIJO_SESION);
  const token = SEG_GENERAR_TOKEN_SESION();

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);

  const sesion = {
    ID_SESION: idSesion, ID_USUARIO: usuario.ID_USUARIO, USUARIO: usuario.USUARIO, ID_ROL: usuario.ID_ROL,
    TOKEN_SESION: token, ESTADO_SESION: "ACTIVA", FECHA_INICIO: ahora, EXPIRA_SESION: fechaExp,
    ULTIMA_ACTIVIDAD: ahora, TIPO_ACCESO: "WEB", ORIGEN_ACCESO: "APLICACION_WEB"
  };

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, sesion);
  hoja.appendRow(nuevaFila);

  return SEG_SANITIZAR_PARA_CLIENTE({ EXITO: true, ID_SESION: idSesion, TOKEN_SESION: token, ID_USUARIO: usuario.ID_USUARIO, FECHA_EXPIRACION: fechaExp });
}

function SEG_BUSCAR_SESION(tokenSesion) {
  if (!tokenSesion || String(tokenSesion).trim() === "") return null;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);

  const idxToken = encabezados.indexOf("TOKEN_SESION");
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxToken]) === String(tokenSesion)) {
      const sesion = SEG_CONVERTIR_FILA_OBJETO(encabezados, registros[i]);
      sesion._FILA = i + 2;
      return sesion;
    }
  }
  return null;
}

function SEG_VALIDAR_SESION(tokenSesion) {
  if (tokenSesion === "SHEETS_CONTEXT") {
    return SEG_SANITIZAR_PARA_CLIENTE({
      VALIDA: true,
      CODIGO: "SESION_VALIDA",
      MENSAJE: "Acceso concedido automáticamente en entorno confiable de Google Sheets.",
      SESION: {
        ID_SESION: "SES-SHEETS-LOCAL",
        ID_USUARIO: "USR-000001",
        USUARIO: "ADMINISTRADOR_LOCAL_SHEETS",
        ID_ROL: "ROL-000001",
        ESTADO_SESION: "ACTIVA",
        EXPIRA_SESION: new Date(new Date().getTime() + 24 * 60 * 60 * 1000),
        ULTIMA_ACTIVIDAD: new Date()
      }
    });
  }
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return { VALIDA: false, CODIGO: "SESION_NO_ENCONTRADA", MENSAJE: "La sesión no existe." };

  const estado = String(sesion.ESTADO_SESION || "").trim().toUpperCase();
  if (estado !== "ACTIVA") return { VALIDA: false, CODIGO: "SESION_NO_ACTIVA", MENSAJE: "La sesión no se encuentra activa." };

  const ahora = SEG_AHORA();
  const fechaExp = new Date(sesion.EXPIRA_SESION || sesion.FECHA_EXPIRACION);
  if (ahora.getTime() >= fechaExp.getTime()) {
    SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
    return { VALIDA: false, CODIGO: "SESION_EXPIRADA", MENSAJE: "La sesión ha expirado por límite de tiempo." };
  }

  const ultimaActividad = new Date(sesion.ULTIMA_ACTIVIDAD);
  const diferenciaMinutos = (ahora.getTime() - ultimaActividad.getTime()) / (1000 * 60);
  if (diferenciaMinutos > SEG_CONFIG.TIEMPO_INACTIVIDAD_MINUTOS) {
    SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
    return { VALIDA: false, CODIGO: "SESION_EXPIRADA_INACTIVIDAD", MENSAJE: "La sesión ha expirado por inactividad." };
  }

  SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion);
  return SEG_SANITIZAR_PARA_CLIENTE({ VALIDA: true, CODIGO: "SESION_VALIDA", MENSAJE: "Sesión autorizada.", SESION: sesion });
}

function SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion) {
  if (tokenSesion === "SHEETS_CONTEXT") return true;
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return false;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxActividad = encabezados.indexOf("ULTIMA_ACTIVIDAD");
  hoja.getRange(sesion._FILA, idxActividad + 1).setValue(SEG_AHORA());
  return true;
}

function SEG_CAMBIAR_ESTADO_SESION(tokenSesion, nuevoEstado) {
  if (tokenSesion === "SHEETS_CONTEXT") return true;
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return false;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxEstado = encabezados.indexOf("ESTADO_SESION");
  hoja.getRange(sesion._FILA, idxEstado + 1).setValue(nuevoEstado.toUpperCase());
  return true;
}

function SEG_CERRAR_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return { EXITO: false, MENSAJE: "No se encontró la sesión." };
  SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "CERRADA");
  return { EXITO: true, MENSAJE: "Sesión cerrada correctamente." };
}

function SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion) {
  const validacionSesion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacionSesion || validacionSesion.VALIDA !== true) {
    return { AUTORIZADO: false, CODIGO: "SESION_INVALIDA", MENSAJE: validacionSesion ? validacionSesion.MENSAJE : "No autorizado." };
  }

  const sesion = validacionSesion.SESION;
  const rol = SEG_CONSULTAR_ROL(sesion.ID_ROL);
  if (!rol || String(rol.ESTADO_ROL || "").trim().toUpperCase() !== "ACTIVO") {
    return { AUTORIZADO: false, CODIGO: "ROL_INACTIVO", MENSAJE: "El rol de la sesión se encuentra inactivo." };
  }

  if (rol.NOMBRE_ROL === SEG_CONFIG.ROL_ADMINISTRADOR) {
    return SEG_SANITIZAR_PARA_CLIENTE({ AUTORIZADO: true, CODIGO: "ACCESO_ADMINISTRADOR", USUARIO: sesion.USUARIO, ROL: rol.NOMBRE_ROL, SESION: sesion });
  }

  const autorizado = SEG_VALIDAR_PERMISO_ROL(rol.ID_ROL, modulo, accion);
  if (!authorized) {
    return { AUTORIZADO: false, CODIGO: "ACCESO_DENEGADO", MENSAJE: "Su rol no tiene permisos de " + accion + " en " + modulo + "." };
  }

  return SEG_SANITIZAR_PARA_CLIENTE({ AUTORIZADO: true, CODIGO: "ACCESO_CONCEDIDO", USUARIO: sesion.USUARIO, ROL: rol.NOMBRE_ROL, SESION: sesion });
}

function SEG_OBTENER_CONTEXTO_SEGURIDAD(tokenSesion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion || validacion.VALIDA !== true) return { VALIDO: false, MENSAJE: "No autorizado." };
  const rol = SEG_CONSULTAR_ROL(validacion.SESION.ID_ROL);
  return SEG_SANITIZAR_PARA_CLIENTE({ VALIDO: true, USUARIO: validacion.SESION.USUARIO, ROL: rol });
}

function SEG_REGISTRAR_AUDITORIA(datos) {
  if (!SEG_CONFIG.REGISTRAR_AUDITORIA) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_AUDITORIA);
    if (!hoja) return;
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_AUDITORIA);
    const ahora = SEG_AHORA();
    const idAuditoria = SEG_GENERAR_ID(SEG_CONFIG.HOJA_AUDITORIA, "ID_AUDITORIA", SEG_CONFIG.PREFIJO_AUDITORIA);

    const registro = { ID_AUDITORIA: idAuditoria, FECHA_HORA: ahora, FECHA_CREACION: ahora };
    encabezados.forEach(campo => {
      if (datos[campo] !== undefined) registro[campo] = datos[campo];
      else if (registro[campo] === undefined) registro[campo] = "";
    });

    hoja.appendRow(SEG_CONVERTIR_OBJETO_FILA(encabezados, registro));
  } catch (err) {
    console.error("No se pudo registrar log de auditoría: " + err.toString());
  }
}

function SEG_OBTENER_ROLES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
    const roles = SEG_LISTAR_ROLES();
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(roles),
      MENSAJE: "Roles del sistema obtenidos exitosamente."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_OBTENER_ROLES", "SEGURIDAD", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los roles: " + error.toString() };
  }
}

function SEG_OBTENER_PERMISOS(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS();
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Permisos obtenidos exitosamente."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_OBTENER_PERMISOS", "SEGURIDAD", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los permisos: " + error.toString() };
  }
}

function SEG_OBTENER_SESIONES(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
    const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Sesiones activas obtenidas exitosamente."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_OBTENER_SESIONES", "SEGURIDAD", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar las sesiones: " + error.toString() };
  }
}

function SEG_APROBAR_USUARIO(idUsuario, idRol, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "APROBAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  const rol = SEG_CONSULTAR_ROL(idRol);
  if (!rol) throw new Error("El rol seleccionado '" + idRol + "' no existe.");

  SEG_ACTUALIZAR_USUARIO(idUsuario, { ID_ROL: idRol, ESTADO_USUARIO: "ACTIVO", USUARIO_ACTUALIZACION: usuarioEjecutor }, tokenSesion);

  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: idUsuario,
    USUARIO: usuarioEjecutor,
    MODULO: "SEGURIDAD",
    SUBMODULO: "USUARIOS",
    ACCION: "APROBAR",
    TIPO_REGISTRO: "USUARIOS",
    ID_REGISTRO: idUsuario,
    DESCRIPCION: "Usuario aprobado con rol: " + rol.NOMBRE_ROL + " por " + usuarioEjecutor,
    RESULTADO: "EXITOSO"
  });

  return { EXITO: true, MENSAJE: "Usuario aprobado con éxito." };
}

function SEG_GUARDAR_LOGO(base64Data, nombreArchivo) {
  try {
    SpreadsheetApp.getUi();
  } catch (e) {
    throw new Error("Esta operación solo está permitida para administradores dentro de Google Sheets.");
  }
  const splitData = base64Data.split(",");
  const contentType = splitData.match(/:(.*?);/)[1];
  const bytes = Utilities.base64Decode(splitData[1]);
  const blob = Utilities.newBlob(bytes, contentType, nombreArchivo);

  let carpeta;
  const carpetas = DriveApp.getFoldersByName("MEGUDAN_RECURSOS");
  if (carpetas.hasNext()) {
    carpeta = carpetas.next();
  } else {
    carpeta = DriveApp.createFolder("MEGUDAN_RECURSOS");
  }

  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const urlLogo = archivo.getUrl();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEmpresa = ss.getSheetByName("CFG_EMPRESA");
  if (!hojaEmpresa) throw new Error("No se encontró la hoja CFG_EMPRESA.");

  const datos = hojaEmpresa.getRange(2, 1, hojaEmpresa.getLastRow() - 1, 2).getValues();
  let filaModificar = -1;
  for (let i = 0; i < datos.length; i++) {
    if (datos[i].toString().trim().toUpperCase() === "LOGO_URL" || datos[i].toString().trim().toUpperCase() === "LOGO") {
      filaModificar = i + 2;
      break;
    }
  }

  if (filaModificar !== -1) {
    hojaEmpresa.getRange(filaModificar, 2).setValue(urlLogo);
  } else {
    hojaEmpresa.appendRow(["LOGO_URL", urlLogo, "URL", "NO", "Ubicación del logo de la empresa"]);
  }

  return { EXITO: true, MENSAJE: "¡Logo corporativo subido y guardado exitosamente!", URL: urlLogo };
}

function SEG_SOLICITAR_RECUPERACION_CONTRASENA(correo) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", correo);
  if (!usuario) return { EXITO: false, MENSAJE: "El correo no se encuentra registrado en el ERP." };

  const caracteres = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#%";
  let contrasenaTemp = "M9!";
  for (let i = 0; i < 7; i++) {
    contrasenaTemp += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }

  const hash = SEG_GENERAR_HASH_CONTRASENA(contrasenaTemp);
  SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, {
    CONTRASENA_HASH: hash,
    DEBE_CAMBIAR_CONTRASENA: "SI",
    FECHA_CAMBIO_CONTRASENA: SEG_AHORA()
  }, "SISTEMA_INTERNAL_BYPASS");

  const asunto = "Recuperación de Acceso - MEGUDAN ERP";
  const cuerpoHTML = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 500px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #374151; color: white; padding: 20px; text-align: center;">
        <h2 style="margin: 0;">MEGUDAN ERP</h2>
      </div>
      <div style="padding: 25px; color: #1f2937;">
        <p>Hola <strong>` + (usuario.NOMBRE || usuario.NOMBRE_COMPLETO || "Usuario") + `</strong>,</p>
        <p>Tu contraseña temporal de acceso es:</p>
        <p style="text-align: center;"><code style="font-size: 18px; font-weight: bold; background: #e5e7eb; padding: 4px 10px;">` + contrasenaTemp + `</code></p>
        <p style="color: #dc2626;">⚠️ Debes cambiar esta clave inmediatamente al iniciar sesión.</p>
      </div>
    </div>`;

  MailApp.sendEmail({ to: correo, subject: asunto, htmlBody: cuerpoHTML });
  return { EXITO: true, MENSAJE: "¡Se ha enviado una contraseña temporal a su correo con éxito!" };
}

function SEG_INICIALIZAR_ROLES_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_ROLES);
    if (!hoja || hoja.getLastRow() >= 2) return;

    console.log("Inicializando roles predefinidos...");
    const ahora = new Date();
    const roles = [
      ["ROL-000001", "ADMINISTRADOR", "Administración general y acceso total al ERP", 100, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol maestro protegido"],
      ["ROL-000002", "CONTADOR", "Gestión y supervisión de procesos contables y financieros", 80, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a reportes contables"],
      ["ROL-000003", "AUXILIAR_CONTABLE", "Apoyo en registros y procesos contables autorizados", 60, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Permiso de registro operativo"],
      ["ROL-000004", "TESORERO", "Gestión de tesorería, pagos, recaudos y conciliaciones", 70, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a bancos y cajas"],
      ["ROL-000005", "COMERCIAL", "Gestión de clientes y operaciones comerciales", 50, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso limitado a ventas"],
      ["ROL-000006", "OPERATIVO", "Registro de operaciones asignadas e inventarios", 40, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol operativo general"],
      ["ROL-000007", "CONSULTA", "Acceso exclusivamente de lectura general", 10, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol de sólo lectura"]
    ];

    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
    roles.forEach(fila => {
      hoja.appendRow(SEG_CONVERTIR_OBJETO_FILA(encabezados, SEG_CONVERTIR_FILA_OBJETO(["ID_ROL", "NOMBRE_ROL", "DESCRIPCION", "NIVEL_JERARQUIA", "ESTADO_ROL", "ROL_SISTEMA", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"], fila)));
    });
  } catch (error) {
    console.error("Error al inicializar roles: " + error.toString());
  }
}

function SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja || hoja.getLastRow() >= 2) return;

    console.log("Inicializando permisos...");
    const ahora = new Date();
    const modulos = ["SEGURIDAD", "CLIENTES", "PROVEEDORES", "PRODUCTOS", "OBRAS", "VENTAS", "COMPRAS", "INVENTARIO", "TESORERIA"];
    const accionesAdmin = ["VER", "CREAR", "EDITAR", "ELIMINAR", "ANULAR", "APROBAR", "ADMINISTRAR"];
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);

    let consecutivo = 1;
    modulos.forEach(mod => {
      accionesAdmin.forEach(acc => {
        const idPermiso = "PER-" + String(consecutivo).padStart(SEG_CONFIG.DIGITOS_ID, "0");
        const obj = {
          ID_PERMISO: idPermiso, ID_ROL: "ROL-000001", MODULO: mod, SUBMODULO: "GENERAL", ACCION: acc,
          PERMITIDO: "SI", ESTADO_PERMISO: "ACTIVO", FECHA_CREACION: ahora, FECHA_ACTUALIZACION: ahora,
          USUARIO_CREACION: "SISTEMA", USUARIO_ACTUALIZACION: "SISTEMA", OBSERVACIONES: "Permiso administrativo maestro"
        };
        hoja.appendRow(SEG_CONVERTIR_OBJETO_FILA(encabezados, obj));
        consecutivo++;
      });
    });
  } catch (error) {
    console.error("Error al inicializar permisos: " + error.toString());
  }
}

function SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
    if (!hoja || hoja.getLastRow() >= 2) return;

    SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
    const adminUser = {
      USUARIO: "ADMIN",
      NOMBRE: "Administrador del ERP",
      CORREO: "admin@megudan.com",
      ID_ROL: "ROL-000001",
      ESTADO_USUARIO: "ACTIVO",
      CONTRASENA_PLANA: "Admin123!",
      DEBE_CAMBIAR_CONTRASENA: "NO"
    };

    SEG_CREAR_USUARIO(adminUser, "SISTEMA_INTERNAL_BYPASS");
    console.log("Usuario administrador creado.");
  } catch (error) {
    console.error("Error al inicializar usuarios: " + error.toString());
  }
}

function SEG_SANITIZAR_PARA_CLIENTE(dato) {
  if (dato === null || dato === undefined) return dato;
  
  if (dato instanceof Date || (dato && typeof dato === "object" && typeof dato.getMonth === "function")) {
    try {
      return Utilities.formatDate(new Date(dato), Session.getScriptTimeZone() || "America/Bogota", "yyyy-MM-dd HH:mm:ss");
    } catch (e) {
      try {
        return new Date(dato).toISOString().replace("T", " ").substring(0, 19);
      } catch (err) {
        return String(dato);
      }
    }
  }
  
  if (Array.isArray(dato)) {
    return dato.map(SEG_SANITIZAR_PARA_CLIENTE);
  }
  
  if (typeof dato === "object") {
    const nuevoObj = {};
    for (const key in dato) {
      if (dato.hasOwnProperty(key)) {
        nuevoObj[key] = SEG_SANITIZAR_PARA_CLIENTE(dato[key]);
      }
    }
    return nuevoObj;
  }
  
  return dato;
}

function SEG_REGISTRAR_USUARIO_PUBLICO(datos) {
  try {
    if (!datos) return { EXITO: false, MENSAJE: "No se proporcionaron datos." };
    if (!datos.USUARIO || String(datos.USUARIO).trim() === "") {
      return { EXITO: false, MENSAJE: "El nombre de usuario es obligatorio." };
    }
    if (!datos.NOMBRE || String(datos.NOMBRE).trim() === "") {
      return { EXITO: false, MENSAJE: "El nombre completo es obligatorio." };
    }
    if (!datos.CORREO || String(datos.CORREO).trim() === "") {
      return { EXITO: false, MENSAJE: "El correo electrónico es obligatorio." };
    }
    if (!datos.CONTRASENA_PLANA || String(datos.CONTRASENA_PLANA).trim() === "") {
      return { EXITO: false, MENSAJE: "La contraseña es obligatoria." };
    }

    const payload = {
      USUARIO: datos.USUARIO,
      NOMBRE: datos.NOMBRE,
      CORREO: datos.CORREO,
      CONTRASENA_PLANA: datos.CONTRASENA_PLANA,
      ID_ROL: "ROL-000006", 
      ESTADO_USUARIO: "PENDIENTE",
      DEBE_CAMBIAR_CONTRASENA: "NO"
    };

    const res = SEG_CREAR_USUARIO(payload, "SISTEMA_INTERNAL_BYPASS");
    if (res && res.EXITO) {
      SEG_ESTABLECER_CONTRASENA(res.ID_USUARIO, datos.CONTRASENA_PLANA, "AUTO_REGISTRO", "SISTEMA_INTERNAL_BYPASS");
      return {
        EXITO: true,
        MENSAJE: "¡Registro exitoso! Tu usuario '" + datos.USUARIO + "' ha sido creado con estado PENDIENTE. Un administrador debe aprobar tu cuenta para permitir el ingreso."
      };
    } else {
      return { EXITO: false, MENSAJE: "Ocurrió un error inesperado al registrar el usuario." };
    }
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_REGISTRAR_USUARIO_PUBLICO", "SEGURIDAD", error);
    }
    return { EXITO: false, MENSAJE: error.message || error.toString() };
  }
}