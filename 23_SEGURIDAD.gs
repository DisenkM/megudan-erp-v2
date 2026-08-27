// ============================================================
// 23. SEGURIDAD Y CONTROL DE ACCESO
// ARCHIVO: 23_SEGURIDAD.gs
// ============================================================

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE SEGURIDAD
// ============================================================
const SEG_CONFIG = {
  // Hojas de base de datos
  HOJA_USUARIOS: "USR_USUARIOS",
  HOJA_ROLES: "USR_ROLES",
  HOJA_PERMISOS: "USR_PERMISOS",
  HOJA_SESIONES: "USR_SESIONES",
  HOJA_AUDITORIA: "USR_AUDITORIA",

  // Prefijos de identificadores únicos
  PREFIJO_USUARIO: "USR",
  PREFIJO_ROL: "ROL",
  PREFIJO_PERMISO: "PER",
  PREFIJO_SESION: "SES",
  PREFIJO_AUDITORIA: "AUD",

  // Formato de IDs (dígitos para autoincremento)
  DIGITOS_ID: 6,

  // Estados del sistema
  ESTADO_USUARIO_ACTIVO: "ACTIVO",
  ESTADO_USUARIO_INACTIVO: "INACTIVO",
  ESTADO_USUARIO_BLOQUEADO: "BLOQUEADO",
  ESTADO_USUARIO_PENDIENTE: "PENDIENTE",

  ESTADO_ROL_ACTIVO: "ACTIVO",
  ESTADO_ROL_INACTIVO: "INACTIVO",

  ESTADO_PERMISO_ACTIVO: "ACTIVO",
  ESTADO_PERMISO_INACTIVO: "INACTIVO",

  ESTADO_SESION_ACTIVA: "ACTIVA",
  ESTADO_SESION_CERRADA: "CERRADA",
  ESTADO_SESION_EXPIRADA: "EXPIRADA",
  ESTADO_SESION_BLOQUEADA: "BLOQUEADA",
  ESTADO_SESION_REVOCADA: "REVOCADA",

  // Control de tiempos y duración
  DURACION_SESION_HORAS: 8,
  TIEMPO_INACTIVIDAD_MINUTOS: 30,

  // Seguridad de fuerza bruta
  MAXIMO_INTENTOS_LOGIN: 5,
  BLOQUEO_USUARIO_MINUTOS: 30,

  // Orígenes de acceso
  ORIGEN_GOOGLE_SHEETS: "GOOGLE_SHEETS",
  ORIGEN_APLICACION_WEB: "APLICACION_WEB",
  ORIGEN_SISTEMA: "SISTEMA",

  // Tipos de acceso
  TIPO_ACCESO_WEB: "WEB",
  TIPO_ACCESO_GOOGLE_SHEETS: "GOOGLE_SHEETS",

  // Resultados de auditoría
  RESULTADO_EXITOSO: "EXITOSO",
  RESULTADO_DENEGADO: "DENEGADO",
  RESULTADO_ERROR: "ERROR",

  // Catálogo de acciones de auditoría
  ACCION_ACCEDER: "ACCEDER",
  ACCION_VER: "VER",
  ACCION_CREAR: "CREAR",
  ACCION_EDITAR: "EDITAR",
  ACCION_ELIMINAR: "ELIMINAR",
  ACCION_ANULAR: "ANULAR",
  ACCION_APROBAR: "APROBAR",
  ACCION_EXPORTAR: "EXPORTAR",
  ACCION_IMPORTAR: "IMPORTAR",
  ACCION_CONFIGURAR: "CONFIGURAR",
  ACCION_ADMINISTRAR: "ADMINISTRAR",
  ACCION_INICIAR_SESION: "INICIAR_SESION",
  ACCION_CERRAR_SESION: "CERRAR_SESION",

  // Roles protegidos del ERP
  ROL_ADMINISTRADOR: "ADMINISTRADOR",

  // Parámetros de auditoría
  REGISTRAR_AUDITORIA: true,
  REGISTRAR_LOGIN: true,
  REGISTRAR_ERRORES_SEGURIDAD: true
};

// ============================================================
// 02. AUXILIARES Y ACCESO A BASE DE DATOS
// ============================================================

/**
 * Obtiene un objeto de tipo Sheet de Google Sheets por su nombre.
 */
function SEG_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    throw new Error("No existe la hoja física '" + nombreHoja + "' en la base de datos.");
  }
  return hoja;
}

/**
 * Obtiene y normaliza los encabezados de una hoja de base de datos.
 */
function SEG_OBTENER_ENCABEZADOS(nombreHoja) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaColumna === 0) {
    throw new Error("La hoja '" + nombreHoja + "' no contiene una fila de encabezados.");
  }
  return hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getDisplayValues()[0]
    .map(function(encabezado) {
      return String(encabezado || "").trim().toUpperCase();
    });
}

/**
 * Obtiene todos los registros físicos de una hoja (omitiendo la fila de encabezados).
 */
function SEG_OBTENER_REGISTROS(nombreHoja) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) {
    return [];
  }
  return hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
}

/**
 * Normaliza un texto para comparaciones robustas (elimina espacios múltiples y mayúsculas).
 */
function SEG_NORMALIZAR_TEXTO(valor) {
  return String(valor || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

/**
 * Genera el siguiente ID de tipo correlativo con prefijo para cualquier maestro del módulo.
 * Ejemplo: USR-000001, ROL-000005, SES-000124, AUD-000003
 */
function SEG_GENERAR_ID(nombreHoja, campoID, prefijo) {
  const hoja = SEG_OBTENER_HOJA(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return prefijo + "-" + String(1).padStart(SEG_CONFIG.DIGITOS_ID, "0");
  }

  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const indiceID = encabezados.indexOf(campoID.toUpperCase());
  if (indiceID === -1) {
    throw new Error("No se encontró la columna '" + campoID + "' en la hoja '" + nombreHoja + "'.");
  }

  const registros = hoja.getRange(2, indiceID + 1, ultimaFila - 1, 1).getDisplayValues().flat();
  let numeroMayor = 0;
  
  registros.forEach(function(id) {
    const textoID = String(id || "").trim();
    if (textoID.startsWith(prefijo + "-")) {
      const numero = parseInt(textoID.replace(prefijo + "-", ""), 10);
      if (!isNaN(numero) && numero > numeroMayor) {
        numeroMayor = numero;
      }
    }
  });

  return prefijo + "-" + String(numeroMayor + 1).padStart(SEG_CONFIG.DIGITOS_ID, "0");
}

/**
 * Convierte un arreglo plano de base de datos (fila) a un objeto clave-valor utilizando los encabezados.
 */
function SEG_CONVERTIR_FILA_OBJETO(encabezados, fila) {
  const objeto = {};
  encabezados.forEach(function(campo, indice) {
    objeto[campo] = fila[indice] !== undefined ? fila[indice] : "";
  });
  return objeto;
}

/**
 * Convierte un objeto clave-valor en un arreglo plano alineado con el orden de los encabezados.
 */
function SEG_CONVERTIR_OBJETO_FILA(encabezados, objeto) {
  return encabezados.map(function(campo) {
    return objeto[campo] !== undefined ? objeto[campo] : "";
  });
}

/**
 * Busca un registro exacto en un campo de una hoja y lo devuelve en formato Objeto.
 */
function SEG_BUSCAR_REGISTRO(nombreHoja, campoBusqueda, valorBusqueda) {
  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const registros = SEG_OBTENER_REGISTROS(nombreHoja);
  const indiceCampo = encabezados.indexOf(campoBusqueda.toUpperCase());
  if (indiceCampo === -1) {
    throw new Error("No se encontró el campo '" + campoBusqueda + "' en la hoja '" + nombreHoja + "'.");
  }

  const valorNormalizado = SEG_NORMALIZAR_TEXTO(valorBusqueda);
  const filaEncontrada = registros.find(function(fila) {
    return SEG_NORMALIZAR_TEXTO(fila[indiceCampo]) === valorNormalizado;
  });

  if (!filaEncontrada) {
    return null;
  }

  return SEG_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada);
}

/**
 * Busca el número de fila física (1-based index de Sheets) para un registro específico.
 */
function SEG_BUSCAR_FILA_REGISTRO(nombreHoja, campoBusqueda, valorBusqueda) {
  const encabezados = SEG_OBTENER_ENCABEZADOS(nombreHoja);
  const registros = SEG_OBTENER_REGISTROS(nombreHoja);
  const indiceCampo = encabezados.indexOf(campoBusqueda.toUpperCase());
  if (indiceCampo === -1) {
    throw new Error("No se encontró el campo '" + campoBusqueda + "' en la hoja '" + nombreHoja + "'.");
  }

  const valorNormalizado = SEG_NORMALIZAR_TEXTO(valorBusqueda);
  for (let i = 0; i < registros.length; i++) {
    if (SEG_NORMALIZAR_TEXTO(registros[i][indiceCampo]) === valorNormalizado) {
      return i + 2; // +2 por fila de encabezado y conversión a 1-based index
    }
  }
  return null;
}

/**
 * Valida de forma estricta que campos indicados como obligatorios tengan valor.
 */
function SEG_VALIDAR_OBLIGATORIOS(datos, camposObligatorios) {
  camposObligatorios.forEach(function(campo) {
    const valor = datos[campo];
    if (value === undefined || value === null || String(valor).trim() === "") {
      throw new Error("El campo " + campo + " es obligatorio.");
    }
  });
}

/**
 * Centraliza la obtención de la fecha y hora oficial del ERP.
 */
function SEG_AHORA() {
  return new Date();
}

// ============================================================
// 03. CRUD DE USUARIOS
// ============================================================

/**
 * Crea un usuario en USR_USUARIOS con su validación de estructura dinámica.
 */
function SEG_CREAR_USUARIO(datos) {
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar la información del usuario.");
  }

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);

  // Soporta dinámicamente si el campo de nombre es NOMBRE o NOMBRE_COMPLETO
  let campoNombre = encabezados.indexOf("NOMBRE") !== -1 ? "NOMBRE" : "NOMBRE_COMPLETO";
  
  // Normalizar llaves para compatibilidad con datos del frontend
  if (campoNombre === "NOMBRE_COMPLETO" && datos.NOMBRE_COMPLETO === undefined && datos.NOMBRE !== undefined) {
    datos.NOMBRE_COMPLETO = datos.NOMBRE;
  } else if (campoNombre === "NOMBRE" && datos.NOMBRE === undefined && datos.NOMBRE_COMPLETO !== undefined) {
    datos.NOMBRE = datos.NOMBRE_COMPLETO;
  }

  // Normalizar llaves lógicas
  datos.USUARIO = String(datos.USUARIO || "").trim().toUpperCase().replace(/\s+/g, "");
  datos.CORREO = String(datos.CORREO || "").trim().toLowerCase();

  // Validar duplicados de usuario o correo
  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "USUARIO", datos.USUARIO)) {
    throw new Error("El nombre de usuario '" + datos.USUARIO + "' ya se encuentra registrado.");
  }
  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", datos.CORREO)) {
    throw new Error("El correo electrónico '" + datos.CORREO + "' ya se encuentra registrado.");
  }

  const ahora = SEG_AHORA();
  const idUsuario = SEG_GENERAR_ID(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", SEG_CONFIG.PREFIJO_USUARIO);

  const usuario = {};
  encabezados.forEach(function(campo) {
    usuario[campo] = datos[campo] !== undefined ? datos[campo] : "";
  });

  // Poblar valores de sistema
  usuario.ID_USUARIO = idUsuario;
  usuario.USUARIO = datos.USUARIO;
  usuario[campoNombre] = datos[campoNombre] || datos.NOMBRE || datos.NOMBRE_COMPLETO || "";
  usuario.CORREO = datos.CORREO;
  usuario.ESTADO_USUARIO = datos.ESTADO_USUARIO || SEG_CONFIG.ESTADO_USUARIO_ACTIVO || "ACTIVO";
  usuario.INTENTOS_FALLIDOS = 0;
  usuario.FECHA_CREACION = ahora;
  usuario.FECHA_ACTUALIZACION = ahora;
  usuario.USUARIO_CREACION = datos.USUARIO_CREACION || Session.getActiveUser().getEmail() || "SISTEMA";
  usuario.USUARIO_ACTUALIZACION = datos.USUARIO_ACTUALIZACION || usuario.USUARIO_CREACION;

  // Si se incluye una contraseña plana para crear
  if (datos.CONTRASENA_PLANA) {
    const seguridadValida = SEG_VALIDAR_SEGURIDAD_CONTRASENA(datos.CONTRASENA_PLANA);
    if (!seguridadValida.VALIDA) {
      throw new Error(seguridadValida.MENSAJES.join(" "));
    }
    usuario.CONTRASENA_HASH = SEG_GENERAR_HASH_CONTRASENA(datos.CONTRASENA_PLANA);
    usuario.DEBE_CAMBIAR_CONTRASENA = datos.DEBE_CAMBIAR_CONTRASENA || "NO";
  }

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, usuario);
  hoja.appendRow(nuevaFila);

  // Registrar en la auditoría
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: idUsuario,
    USUARIO: usuario.USUARIO,
    MODULO: "SEGURIDAD",
    SUBMODULO: "USUARIOS",
    ACCION: SEG_CONFIG.ACCION_CREAR,
    TIPO_REGISTRO: "USUARIOS",
    ID_REGISTRO: idUsuario,
    DESCRIPCION: "Usuario " + usuario.USUARIO + " creado con éxito.",
    RESULTADO: SEG_CONFIG.RESULTADO_EXITOSO
  });

  return {
    EXITO: true,
    ID_USUARIO: idUsuario,
    USUARIO: usuario
  };
}

/**
 * Consulta la ficha completa de un usuario por su ID.
 */
function SEG_CONSULTAR_USUARIO(idUsuario) {
  return SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
}

/**
 * Recupera un listado completo con todos los usuarios registrados.
 */
function SEG_LISTAR_USUARIOS() {
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_USUARIOS);
  return registros
    .filter(function(fila) {
      return fila[0] && String(fila[0]).trim() !== "";
    })
    .map(function(fila) {
      return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
    });
}

/**
 * Actualiza los campos permitidos de un usuario en base a su ID.
 */
function SEG_ACTUALIZAR_USUARIO(idUsuario, datos) {
  if (!idUsuario || String(idUsuario).trim() === "") {
    throw new Error("Debe indicar el ID_USUARIO.");
  }
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar los datos a actualizar.");
  }

  const usuarioActual = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuarioActual) {
    throw new Error("No se encontró el usuario '" + idUsuario + "'.");
  }

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const filaUsuario = SEG_BUSCAR_FILA_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);

  const camposProtegidos = ["ID_USUARIO", "FECHA_CREACION", "USUARIO_CREACION"];
  const usuarioActualizado = Object.assign({}, usuarioActual);

  Object.keys(datos).forEach(function(campo) {
    const campoNormalizado = String(campo).trim().toUpperCase();
    if (!camposProtegidos.includes(campoNormalizado) && encabezados.indexOf(campoNormalizado) !== -1) {
      usuarioActualizado[campoNormalizado] = datos[campo];
    }
  });

  // Validaciones adicionales si se actualiza campos sensibles
  if (datos.USUARIO !== undefined) {
    usuarioActualizado.USUARIO = String(datos.USUARIO || "").trim().toUpperCase().replace(/\s+/g, "");
    const duplicado = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "USUARIO", usuarioActualizado.USUARIO);
    if (duplicado && duplicado.ID_USUARIO !== idUsuario) {
      throw new Error("El nombre de usuario '" + usuarioActualizado.USUARIO + "' ya se encuentra registrado por otro usuario.");
    }
  }

  if (datos.CORREO !== undefined) {
    usuarioActualizado.CORREO = String(datos.CORREO || "").trim().toLowerCase();
    const duplicado = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", usuarioActualizado.CORREO);
    if (duplicado && duplicado.ID_USUARIO !== idUsuario) {
      throw new Error("El correo '" + usuarioActualizado.CORREO + "' ya se encuentra registrado por otro usuario.");
    }
  }

  usuarioActualizado.FECHA_ACTUALIZACION = SEG_AHORA();
  usuarioActualizado.USUARIO_ACTUALIZACION = datos.USUARIO_ACTUALIZACION || Session.getActiveUser().getEmail() || "SISTEMA";

  const filaActualizada = SEG_CONVERTIR_OBJETO_FILA(encabezados, usuarioActualizado);
  hoja.getRange(filaUsuario, 1, 1, encabezados.length).setValues([filaActualizada]);

  return {
    EXITO: true,
    ID_USUARIO: idUsuario,
    USUARIO: usuarioActualizado
  };
}

/**
 * Modifica el estado operacional de un usuario (ACTIVO, INACTIVO, BLOQUEADO).
 */
function SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, nuevoEstado, usuarioActualizacion) {
  const estado = String(nuevoEstado || "").trim().toUpperCase();
  const estadosPermitidos = [
    SEG_CONFIG.ESTADO_USUARIO_ACTIVO,
    SEG_CONFIG.ESTADO_USUARIO_INACTIVO,
    SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO,
    SEG_CONFIG.ESTADO_USUARIO_PENDIENTE
  ];

  if (!estadosPermitidos.includes(estado)) {
    throw new Error("Estado de usuario no válido: " + nuevoEstado);
  }

  return SEG_ACTUALIZAR_USUARIO(idUsuario, {
    ESTADO_USUARIO: estado,
    USUARIO_ACTUALIZACION: usuarioActualizacion || "SISTEMA"
  });
}

function SEG_BLOQUEAR_USUARIO(idUsuario, usuarioActualizacion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO, usuarioActualizacion);
}

function SEG_DESBLOQUEAR_USUARIO(idUsuario, usuarioActualizacion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_ACTIVO, usuarioActualizacion);
}

// ============================================================
// 04. CRUD DE ROLES
// ============================================================

/**
 * Crea un rol en la hoja USR_ROLES.
 */
function SEG_CREAR_ROL(datos) {
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar la información del rol.");
  }
  if (!datos.NOMBRE_ROL || String(datos.NOMBRE_ROL).trim() === "") {
    throw new Error("El campo NOMBRE_ROL es obligatorio.");
  }

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_ROLES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);

  const nombreRol = String(datos.NOMBRE_ROL).trim().toUpperCase();
  if (SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_ROLES, "NOMBRE_ROL", nombreRol)) {
    throw new Error("El rol '" + nombreRol + "' ya existe en la base de datos.");
  }

  const ahora = SEG_AHORA();
  const idRol = SEG_GENERAR_ID(SEG_CONFIG.HOJA_ROLES, "ID_ROL", SEG_CONFIG.PREFIJO_ROL);

  const rol = {};
  encabezados.forEach(function(campo) {
    rol[campo] = datos[campo] !== undefined ? datos[campo] : "";
  });

  rol.ID_ROL = idRol;
  rol.NOMBRE_ROL = nombreRol;
  rol.ESTADO_ROL = datos.ESTADO_ROL || SEG_CONFIG.ESTADO_ROL_ACTIVO || "ACTIVO";
  rol.FECHA_CREACION = ahora;
  rol.FECHA_ACTUALIZACION = ahora;
  rol.USUARIO_CREACION = datos.USUARIO_CREACION || Session.getActiveUser().getEmail() || "SISTEMA";
  rol.USUARIO_ACTUALIZACION = datos.USUARIO_ACTUALIZACION || rol.USUARIO_CREACION;

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, rol);
  hoja.appendRow(nuevaFila);

  return {
    EXITO: true,
    ID_ROL: idRol,
    ROL: rol
  };
}

/**
 * Consulta un rol específico de la base de datos por su ID.
 */
function SEG_CONSULTAR_ROL(idRol) {
  return SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_ROLES, "ID_ROL", idRol);
}

/**
 * Obtiene el listado completo de todos los roles.
 */
function SEG_LISTAR_ROLES() {
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_ROLES);
  return registros
    .filter(function(fila) {
      return fila[0] && String(fila[0]).trim() !== "";
    })
    .map(function(fila) {
      return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
    });
}

/**
 * Actualiza los campos de un rol según su ID.
 */
function SEG_ACTUALIZAR_ROL(idRol, datos) {
  if (!idRol || String(idRol).trim() === "") {
    throw new Error("Debe indicar el ID_ROL.");
  }
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar los datos a actualizar.");
  }

  const rolActual = SEG_CONSULTAR_ROL(idRol);
  if (!rolActual) {
    throw new Error("No se encontró el rol '" + idRol + "'.");
  }

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_ROLES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
  const filaRol = SEG_BUSCAR_FILA_REGISTRO(SEG_CONFIG.HOJA_ROLES, "ID_ROL", idRol);

  const camposProtegidos = ["ID_ROL", "FECHA_CREACION", "USUARIO_CREACION"];
  const rolActualizado = Object.assign({}, rolActual);

  Object.keys(datos).forEach(function(campo) {
    const campoNormalizado = String(campo).trim().toUpperCase();
    if (!camposProtegidos.includes(campoNormalizado) && encabezados.indexOf(campoNormalizado) !== -1) {
      rolActualizado[campoNormalizado] = datos[campo];
    }
  });

  if (datos.NOMBRE_ROL !== undefined) {
    rolActualizado.NOMBRE_ROL = String(datos.NOMBRE_ROL).trim().toUpperCase();
    const duplicado = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_ROLES, "NOMBRE_ROL", rolActualizado.NOMBRE_ROL);
    if (duplicado && duplicado.ID_ROL !== idRol) {
      throw new Error("El nombre de rol '" + rolActualizado.NOMBRE_ROL + "' ya está registrado por otro rol.");
    }
  }

  rolActualizado.FECHA_ACTUALIZACION = SEG_AHORA();
  rolActualizado.USUARIO_ACTUALIZACION = datos.USUARIO_ACTUALIZACION || Session.getActiveUser().getEmail() || "SISTEMA";

  const filaActualizada = SEG_CONVERTIR_OBJETO_FILA(encabezados, rolActualizado);
  hoja.getRange(filaRol, 1, 1, encabezados.length).setValues([filaActualizada]);

  return {
    EXITO: true,
    ID_ROL: idRol,
    ROL: rolActualizado
  };
}

/**
 * Modifica el estado lógico de un rol (ACTIVO, INACTIVO).
 */
function SEG_CAMBIAR_ESTADO_ROL(idRol, nuevoEstado, usuarioActualizacion) {
  const estado = String(nuevoEstado || "").trim().toUpperCase();
  const estadosPermitidos = ["ACTIVO", "INACTIVO"];
  if (!estadosPermitidos.includes(estado)) {
    throw new Error("Estado de rol no válido: " + nuevoEstado);
  }
  return SEG_ACTUALIZAR_ROL(idRol, {
    ESTADO_ROL: estado,
    USUARIO_ACTUALIZACION: usuarioActualizacion || "SISTEMA"
  });
}

function SEG_ACTIVAR_ROL(idRol, usuarioActualizacion) {
  return SEG_CAMBIAR_ESTADO_ROL(idRol, "ACTIVO", usuarioActualizacion);
}

function SEG_INACTIVAR_ROL(idRol, usuarioActualizacion) {
  return SEG_CAMBIAR_ESTADO_ROL(idRol, "INACTIVO", usuarioActualizacion);
}

// ============================================================
// 05. GESTIÓN Y VALIDACIÓN DE PERMISOS
// ============================================================

/**
 * Obtiene todos los registros de permisos asignados a un ID_ROL.
 */
function SEG_OBTENER_PERMISOS_ROL(idRol) {
  if (!idRol || String(idRol).trim() === "") {
    throw new Error("Debe indicar el ID_ROL.");
  }
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
  const idxIdRol = encabezados.indexOf("ID_ROL");
  if (idxIdRol === -1) {
    throw new Error("No se encontró la columna ID_ROL en la hoja de permisos.");
  }
  
  const filtrados = registros.filter(function(fila) {
    return String(fila[idxIdRol] || "").trim().toUpperCase() === String(idRol).trim().toUpperCase();
  });

  return filtrados.map(function(fila) {
    return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
  });
}

/**
 * Busca una regla de permiso específica para un rol dentro de un módulo y acción.
 */
function SEG_BUSCAR_PERMISO(idRol, modulo, accion) {
  const permisos = SEG_OBTENER_PERMISOS_ROL(idRol);
  const moduloNormalizado = SEG_NORMALIZAR_TEXTO(modulo);
  const accionNormalizada = SEG_NORMALIZAR_TEXTO(accion);

  return permisos.find(function(permiso) {
    const moduloRegistro = SEG_NORMALIZAR_TEXTO(permiso.MODULO);
    const accionRegistro = SEG_NORMALIZAR_TEXTO(permiso.ACCION);
    return moduloRegistro === moduloNormalizado && accionRegistro === accionNormalizada;
  }) || null;
}

/**
 * Determina si un Rol específico está habilitado para realizar una Acción en un Módulo.
 */
function SEG_VALIDAR_PERMISO_ROL(idRol, modulo, accion) {
  const rol = SEG_CONSULTAR_ROL(idRol);
  if (!rol || String(rol.ESTADO_ROL || "").trim().toUpperCase() !== "ACTIVO") {
    return false;
  }

  const permiso = SEG_BUSCAR_PERMISO(idRol, modulo, accion);
  if (!permiso || String(permiso.ESTADO_PERMISO || "").trim().toUpperCase() !== "ACTIVO") {
    return false;
  }

  const permitido = String(permiso.PERMITIDO || "").trim().toUpperCase();
  const valoresVerdaderos = ["SI", "SÍ", "TRUE", "1", "VERDADERO"];
  return valoresVerdaderos.includes(permitido);
}

/**
 * Determina si el usuario activo tiene permiso para ejecutar una Acción en un Módulo.
 */
function SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuario || String(usuario.ESTADO_USUARIO || "").trim().toUpperCase() !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) {
    return false;
  }
  if (!usuario.ID_ROL || String(usuario.ID_ROL).trim() === "") {
    return false;
  }
  return SEG_VALIDAR_PERMISO_ROL(usuario.ID_ROL, modulo, accion);
}

/**
 * Verifica acceso a nivel de usuario, deteniendo la ejecución con error en caso de denegación.
 */
function SEG_VERIFICAR_ACCESO(idUsuario, modulo, accion) {
  const autorizado = SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion);
  if (!autorizado) {
    throw new Error("ACCESO DENEGADO. El usuario no cuenta con autorización para '" + accion + "' en el módulo '" + modulo + "'.");
  }
  return true;
}

/**
 * Obtiene el listado de reglas de permisos cargados para el usuario actual.
 */
function SEG_LISTAR_PERMISOS_USUARIO(idUsuario) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuario || !usuario.ID_ROL) {
    return [];
  }
  return SEG_OBTENER_PERMISOS_ROL(usuario.ID_ROL);
}

// ============================================================
// 06. AUTENTICACIÓN Y CONTRASEÑAS
// ============================================================

/**
 * Convierte una contraseña en texto plano en un Hash hexadecimal robusto bajo SHA-256.
 */
function SEG_GENERAR_HASH_CONTRASENA(contrasena) {
  if (contrasena === undefined || contrasena === null || String(contrasena) === "") {
    throw new Error("Debe proporcionar una contraseña para procesar.");
  }
  const bytes = Utilities.newBlob(String(contrasena)).getBytes();
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return hashBytes.map(function(byte) {
    const valor = byte < 0 ? byte + 256 : byte;
    return ("0" + valor.toString(16)).slice(-2);
  }).join("");
}

/**
 * Valida los requisitos de seguridad corporativa para las contraseñas del ERP.
 */
function SEG_VALIDAR_SEGURIDAD_CONTRASENA(contrasena) {
  const password = String(contrasena || "");
  const resultado = {
    VALIDA: false,
    MENSAJES: []
  };

  if (password.length < 8) {
    resultado.MENSAJES.push("La contraseña debe tener al menos 8 caracteres.");
  }
  if (!/[A-Z]/.test(password)) {
    resultado.MENSAJES.push("La contraseña debe incluir al menos una letra mayúscula.");
  }
  if (!/[a-z]/.test(password)) {
    resultado.MENSAJES.push("La contraseña debe incluir al menos una letra minúscula.");
  }
  if (!/[0-9]/.test(password)) {
    resultado.MENSAJES.push("La contraseña debe incluir al menos un número.");
  }
  if (!/[!@#$%^&*()_+\-=\[\]{};':"\|,.<>\/?]/.test(password)) {
    resultado.MENSAJES.push("La contraseña debe incluir al menos un carácter especial.");
  }

  resultado.VALIDA = resultado.MENSAJES.length === 0;
  return resultado;
}

/**
 * Busca a un usuario elegible para el login comparando contra su campo USUARIO o CORREO.
 */
function SEG_BUSCAR_USUARIO_LOGIN(credencial) {
  if (!credencial || String(credencial).trim() === "") {
    return null;
  }
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) {
    return null;
  }

  const idxUsuario = encabezados.indexOf("USUARIO");
  const idxCorreo = encabezados.indexOf("CORREO");
  if (idxUsuario === -1 || idxCorreo === -1) {
    throw new Error("No existen las columnas necesarias (USUARIO, CORREO) en USR_USUARIOS.");
  }

  const criterio = String(credencial).trim().toUpperCase();
  const registros = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();

  const filaEncontrada = registros.find(function(fila) {
    const usuarioReg = String(fila[idxUsuario] || "").trim().toUpperCase();
    const correoReg = String(fila[idxCorreo] || "").trim().toUpperCase();
    return usuarioReg === criterio || correoReg === criterio;
  });

  if (!filaEncontrada) {
    return null;
  }

  return SEG_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada);
}

/**
 * Encripta y establece una contraseña para un usuario determinado de la base de datos.
 */
function SEG_ESTABLECER_CONTRASENA(idUsuario, nuevaContrasena, usuarioActualizacion) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuario) {
    throw new Error("No se encontró el usuario indicado para cambiar contraseña.");
  }

  const validacionSeguridad = SEG_VALIDAR_SEGURIDAD_CONTRASENA(nuevaContrasena);
  if (!validacionSeguridad.VALIDA) {
    throw new Error(validacionSeguridad.MENSAJES.join(" "));
  }

  const hash = SEG_GENERAR_HASH_CONTRASENA(nuevaContrasena);
  return SEG_ACTUALIZAR_USUARIO(idUsuario, {
    CONTRASENA_HASH: hash,
    FECHA_CAMBIO_CONTRASENA: SEG_AHORA(),
    DEBE_CAMBIAR_CONTRASENA: "NO",
    USUARIO_ACTUALIZACION: usuarioActualizacion || "SISTEMA"
  });
}

/**
 * Compara y valida una contraseña enviada en texto plano contra el hash almacenado del usuario.
 */
function SEG_VALIDAR_CONTRASENA(usuario, contrasena) {
  if (!usuario || !usuario.CONTRASENA_HASH || String(usuario.CONTRASENA_HASH).trim() === "") {
    return false;
  }
  const hashIngresado = SEG_GENERAR_HASH_CONTRASENA(contrasena);
  return String(hashIngresado) === String(usuario.CONTRASENA_HASH);
}

/**
 * Autentica credenciales de ingreso (usuario/correo + contraseña) y retorna su contexto.
 */
function SEG_AUTENTICAR_USUARIO(credencial, contrasena) {
  const usuario = SEG_BUSCAR_USUARIO_LOGIN(credencial);
  if (!usuario) {
    return { EXITO: false, CODIGO: "USUARIO_NO_ENCONTRADO", MENSAJE: "Credenciales incorrectas de acceso." };
  }

  const estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();
  if (estado === SEG_CONFIG.ESTADO_USUARIO_PENDIENTE) {
    return { EXITO: false, CODIGO: "USUARIO_PENDIENTE", MENSAJE: "Su cuenta está pendiente de aprobación por un administrador." };
  }
  if (estado === SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO) {
    return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "El usuario se encuentra temporalmente bloqueado." };
  }
  if (estado !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) {
    return { EXITO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "El usuario no está activo en el ERP." };
  }

  const contrasenaValida = SEG_VALIDAR_CONTRASENA(usuario, contrasena);
  if (!contrasenaValida) {
    // Control de intentos fallidos
    let fallidos = parseInt(usuario.INTENTOS_FALLIDOS || 0, 10) + 1;
    const updates = { INTENTOS_FALLIDOS: fallidos };
    if (fallidos >= SEG_CONFIG.MAXIMO_INTENTOS_LOGIN) {
      updates.ESTADO_USUARIO = SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO;
      updates.BLOQUEADO_HASTA = new Date(SEG_AHORA().getTime() + (SEG_CONFIG.BLOQUEO_USUARIO_MINUTOS * 60 * 1000));
    }
    SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, updates);
    
    return { 
      EXITO: false, 
      CODIGO: "CONTRASENA_INCORRECTA", 
      MENSAJE: fallidos >= SEG_CONFIG.MAXIMO_INTENTOS_LOGIN 
        ? "Contraseña incorrecta. El usuario ha sido bloqueado por superar el límite de intentos."
        : "Credenciales incorrectas de acceso." 
    };
  }

  // Login Exitoso: Resetear intentos fallidos
  SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, {
    INTENTOS_FALLIDOS: 0,
    BLOQUEADO_HASTA: "",
    ULTIMO_ACCESO: SEG_AHORA()
  });

  return {
    EXITO: true,
    CODIGO: "AUTENTICACION_CORRECTA",
    MENSAJE: "Autenticación correcta.",
    USUARIO: {
      ID_USUARIO: usuario.ID_USUARIO,
      USUARIO: usuario.USUARIO,
      NOMBRE: usuario.NOMBRE || usuario.NOMBRE_COMPLETO || "",
      CORREO: usuario.CORREO,
      ID_ROL: usuario.ID_ROL
    },
    DEBE_CAMBIAR_CONTRASENA: String(usuario.DEBE_CAMBIAR_CONTRASENA || "").trim().toUpperCase() === "SI" || String(usuario.DEBE_CAMBIAR_CONTRASENA || "").trim().toUpperCase() === "SÍ"
  };
}

/**
 * Modifica la contraseña validando previamente que la contraseña actual sea la correcta.
 */
function SEG_CAMBIAR_CONTRASENA(idUsuario, contrasenaActual, nuevaContrasena, usuarioActualizacion) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuario) {
    throw new Error("No se encontró el usuario indicado.");
  }

  if (!SEG_VALIDAR_CONTRASENA(usuario, contrasenaActual)) {
    throw new Error("La contraseña actual es incorrecta.");
  }

  if (String(contrasenaActual) === String(nuevaContrasena)) {
    throw new Error("La nueva contraseña debe ser diferente de la contraseña actual.");
  }

  return SEG_ESTABLECER_CONTRASENA(idUsuario, nuevaContrasena, usuarioActualizacion || usuario.USUARIO);
}

// ============================================================
// 07. GESTIÓN DE SESIONES
// ============================================================

/**
 * Genera un Token robusto y único de sesión utilizando un identificador UUID doble.
 */
function SEG_GENERAR_TOKEN_SESION() {
  return Utilities.getUuid() + "-" + Utilities.getUuid();
}

/**
 * Crea una sesión de base de datos activa para un usuario autenticado correctamente.
 */
function SEG_CREAR_SESION(idUsuario) {
  const usuario = SEG_CONSULTAR_USUARIO(idUsuario);
  if (!usuario) {
    throw new Error("No se encontró el usuario para crear la sesión.");
  }
  if (String(usuario.ESTADO_USUARIO || "").trim().toUpperCase() !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) {
    throw new Error("No se puede crear sesión para un usuario no activo.");
  }

  const ahora = SEG_AHORA();
  const fechaExpiracion = new Date(ahora.getTime() + (SEG_CONFIG.DURACION_SESION_HORAS * 60 * 60 * 1000));
  const idSesion = SEG_GENERAR_ID(SEG_CONFIG.HOJA_SESIONES, "ID_SESION", SEG_CONFIG.PREFIJO_SESION);
  const tokenSesion = SEG_GENERAR_TOKEN_SESION();

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);

  const sesion = {};
  encabezados.forEach(function(campo) {
    sesion[campo] = "";
  });

  sesion.ID_SESION = idSesion;
  sesion.ID_USUARIO = usuario.ID_USUARIO;
  sesion.USUARIO = usuario.USUARIO;
  sesion.ID_ROL = usuario.ID_ROL;
  sesion.TOKEN_SESION = tokenSesion;
  sesion.ESTADO_SESION = SEG_CONFIG.ESTADO_SESION_ACTIVA || "ACTIVA";
  sesion.FECHA_INICIO = ahora;
  sesion.EXPIRA_SESION = fechaExpiracion;
  sesion.ULTIMA_ACTIVIDAD = ahora;
  sesion.TIPO_ACCESO = SEG_CONFIG.TIPO_ACCESO_WEB || "WEB";
  sesion.ORIGEN_ACCESO = SEG_CONFIG.ORIGEN_APLICACION_WEB || "APLICACION_WEB";

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, sesion);
  hoja.appendRow(nuevaFila);

  return {
    EXITO: true,
    ID_SESION: idSesion,
    TOKEN_SESION: tokenSesion,
    ID_USUARIO: usuario.ID_USUARIO,
    FECHA_EXPIRACION: fechaExpiracion
  };
}

/**
 * Busca y retorna un objeto de sesión recuperado mediante su token único.
 */
function SEG_BUSCAR_SESION(tokenSesion) {
  if (!tokenSesion || String(tokenSesion).trim() === "") {
    return null;
  }
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) {
    return null;
  }

  const idxToken = encabezados.indexOf("TOKEN_SESION");
  if (idxToken === -1) {
    throw new Error("No existe la columna TOKEN_SESION en USR_SESIONES.");
  }

  const registros = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();
  for (let i = 0; i < registros.length; i++) {
    if (String(registros[i][idxToken]) === String(tokenSesion)) {
      const sesion = SEG_CONVERTIR_FILA_OBJETO(encabezados, registros[i]);
      sesion._FILA = i + 2;
      return sesion;
    }
  }
  return null;
}

/**
 * Valida minuciosamente si una sesión activa está vigente y libre de inactividad o timeouts.
 */
function SEG_VALIDAR_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) {
    return { VALIDA: false, CODIGO: "SESION_NO_ENCONTRADA", MENSAJE: "La sesión no existe o no es válida." };
  }

  const estado = String(sesion.ESTADO_SESION || "").trim().toUpperCase();
  if (estado !== "ACTIVA") {
    return { VALIDA: false, CODIGO: "SESION_NO_ACTIVA", MENSAJE: "La sesión no se encuentra activa." };
  }

  const ahora = SEG_AHORA();
  
  // Validar expiración absoluta de duración máxima
  const fechaExpiracion = new Date(sesion.EXPIRA_SESION);
  if (ahora.getTime() >= fechaExpiracion.getTime()) {
    SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
    return { VALIDA: false, CODIGO: "SESION_EXPIRADA", MENSAJE: "La sesión ha expirado por duración máxima." };
  }

  // Validar timeout de inactividad
  const ultimaActividad = new Date(sesion.ULTIMA_ACTIVIDAD);
  const diferenciaMinutos = (ahora.getTime() - ultimaActividad.getTime()) / (1000 * 60);
  if (diferenciaMinutos > SEG_CONFIG.TIEMPO_INACTIVIDAD_MINUTOS) {
    SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
    return { VALIDA: false, CODIGO: "SESION_EXPIRADA_INACTIVIDAD", MENSAJE: "La sesión ha expirado por inactividad." };
  }

  // Sesión válida: Actualizar la última actividad
  SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion);
  return {
    VALIDA: true,
    CODIGO: "SESION_VALIDA",
    MENSAJE: "Sesión válida.",
    SESION: sesion
  };
}

/**
 * Actualiza la última hora de actividad registrada para una sesión.
 */
function SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return false;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxActividad = encabezados.indexOf("ULTIMA_ACTIVIDAD");
  if (idxActividad === -1) return false;
  
  hoja.getRange(sesion._FILA, idxActividad + 1).setValue(SEG_AHORA());
  return true;
}

/**
 * Cambia el estado de una sesión de base de datos.
 */
function SEG_CAMBIAR_ESTADO_SESION(tokenSesion, nuevoEstado) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return false;
  
  const estado = String(nuevoEstado || "").trim().toUpperCase();
  const estadosPermitidos = ["ACTIVA", "CERRADA", "EXPIRADA", "BLOQUEADA", "REVOCADA"];
  if (!estadosPermitidos.includes(estado)) {
    throw new Error("Estado de sesión no válido: " + nuevoEstado);
  }

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxEstado = encabezados.indexOf("ESTADO_SESION");
  if (idxEstado === -1) {
    throw new Error("No existe la columna ESTADO_SESION en USR_SESIONES.");
  }

  hoja.getRange(sesion._FILA, idxEstado + 1).setValue(estado);
  return true;
}

/**
 * Cierra voluntariamente una sesión de base de datos.
 */
function SEG_CERRAR_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) {
    return { EXITO: false, MENSAJE: "No se encontró la sesión activa." };
  }

  SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "CERRADA");
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  
  const idxFechaCierre = encabezados.indexOf("FECHA_CIERRE");
  const idxMotivo = encabezados.indexOf("MOTIVO_CIERRE");
  if (idxFechaCierre !== -1) {
    hoja.getRange(sesion._FILA, idxFechaCierre + 1).setValue(SEG_AHORA());
  }
  if (idxMotivo !== -1) {
    hoja.getRange(sesion._FILA, idxMotivo + 1).setValue("CIERRE_VOLUNTARIO");
  }

  return { EXITO: true, MENSAJE: "Sesión cerrada correctamente." };
}

// ============================================================
// 08. VALIDACIÓN DE ACCESO GENERAL
// ============================================================

/**
 * Obtiene el objeto completo de usuario de la base de datos asociado a una sesión.
 */
function SEG_OBTENER_USUARIO_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return null;
  return SEG_CONSULTAR_USUARIO(sesion.ID_USUARIO);
}

/**
 * Verifica y valida que el usuario de la sesión se mantenga Activo.
 */
function SEG_VALIDAR_USUARIO_SESION(tokenSesion) {
  const usuario = SEG_OBTENER_USUARIO_SESION(tokenSesion);
  if (!usuario) {
    return { VALIDO: false, CODIGO: "USUARIO_SESION_NO_ENCONTRADO", MENSAJE: "No se encontró un usuario válido para la sesión." };
  }
  if (String(usuario.ESTADO_USUARIO || "").trim().toUpperCase() !== SEG_CONFIG.ESTADO_USUARIO_ACTIVO) {
    return { VALIDO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "El usuario asignado a la sesión se encuentra inactivo." };
  }
  return { VALIDO: true, USUARIO: usuario };
}

/**
 * Valida la existencia y estado del Rol asociado a la sesión del usuario.
 */
function SEG_VALIDAR_ROL_SESION(tokenSesion) {
  const validacionUsuario = SEG_VALIDAR_USUARIO_SESION(tokenSesion);
  if (!validacionUsuario.VALIDO) {
    return { VALIDO: false, CODIGO: validacionUsuario.CODIGO, MENSAJE: validacionUsuario.MENSAJE };
  }

  const usuario = validacionUsuario.USUARIO;
  if (!usuario.ID_ROL || String(usuario.ID_ROL).trim() === "") {
    return { VALIDO: false, CODIGO: "USUARIO_SIN_ROL", MENSAJE: "El usuario de la sesión no tiene rol asignado." };
  }

  const rol = SEG_CONSULTAR_ROL(usuario.ID_ROL);
  if (!rol) {
    return { VALIDO: false, CODIGO: "ROL_NO_ENCONTRADO", MENSAJE: "No se encontró el rol asignado al usuario." };
  }

  if (String(rol.ESTADO_ROL || "").trim().toUpperCase() !== "ACTIVO") {
    return { VALIDO: false, CODIGO: "ROL_INACTIVO", MENSAJE: "El rol asignado al usuario de la sesión no está activo." };
  }

  return { VALIDO: true, USUARIO: usuario, ROL: rol };
}

/**
 * Valida el acceso total de un token a una acción de un módulo específico.
 */
function SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion) {
  if (!tokenSesion || String(tokenSesion).trim() === "") {
    return { AUTORIZADO: false, CODIGO: "TOKEN_NO_PROPORCIONADO", MENSAJE: "Debe proporcionar un token de sesión." };
  }
  if (!modulo || String(modulo).trim() === "") {
    return { AUTORIZADO: false, CODIGO: "MODULO_NO_PROPORCIONADO", MENSAJE: "Debe indicar el módulo a validar." };
  }
  if (!accion || String(accion).trim() === "") {
    return { AUTORIZADO: false, CODIGO: "ACCION_NO_PROPORCIONADA", MENSAJE: "Debe indicar la acción a validar." };
  }

  const validacionSession = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacionSession.VALIDA) {
    return { AUTORIZADO: false, CODIGO: validacionSession.CODIGO, MENSAJE: validacionSession.MENSAJE };
  }

  const validacionRol = SEG_VALIDAR_ROL_SESION(tokenSesion);
  if (!validacionRol.VALIDO) {
    return { AUTORIZADO: false, CODIGO: validacionRol.CODIGO, MENSAJE: validacionRol.MENSAJE };
  }

  const usuario = validacionRol.USUARIO;
  const rol = validacionRol.ROL;

  // REGLA SUPREMA: El administrador de sistema tiene pase libre para todo
  if (rol.NOMBRE_ROL === SEG_CONFIG.ROL_ADMINISTRADOR) {
    return {
      AUTORIZADO: true,
      CODIGO: "ACCESO_AUTORIZADO",
      MENSAJE: "Acceso supremo autorizado (Administrador).",
      ID_USUARIO: usuario.ID_USUARIO,
      USUARIO: usuario.USUARIO,
      ID_ROL: rol.ID_ROL,
      ROL: rol.NOMBRE_ROL,
      MODULO: modulo,
      ACCION: accion,
      SESION: validacionSession.SESION
    };
  }

  // Validación de permisos detallados
  const autorizado = SEG_VALIDAR_PERMISO_ROL(rol.ID_ROL, modulo, accion);
  if (!autorizado) {
    return {
      AUTORIZADO: false,
      CODIGO: "PERMISO_DENEGADO",
      MENSAJE: "Acceso denegado. El rol no cuenta con la regla de permiso activa para esta acción.",
      ID_USUARIO: usuario.ID_USUARIO,
      ID_ROL: rol.ID_ROL,
      MODULO: modulo,
      ACCION: accion
    };
  }

  return {
    AUTORIZADO: true,
    CODIGO: "ACCESO_AUTORIZADO",
    MENSAJE: "Acceso autorizado con éxito.",
    ID_USUARIO: usuario.ID_USUARIO,
    USUARIO: usuario.USUARIO,
    ID_ROL: rol.ID_ROL,
    ROL: rol.NOMBRE_ROL,
    MODULO: modulo,
    ACCION: accion,
    SESION: validacionSession.SESION
  };
}

/**
 * Valida acceso obligatorio, rompiendo ejecución con throw Error en caso de desautorización.
 */
function SEG_VERIFICAR_ACCESO_SESION(tokenSesion, modulo, accion) {
  const resultado = SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion);
  if (!resultado.AUTORIZADO) {
    throw new Error("ACCESO DENEGADO [" + resultado.CODIGO + "]: " + resultado.MENSAJE);
  }
  return resultado;
}

/**
 * Devuelve todo el contexto lógico de seguridad para un token de sesión válido.
 */
function SEG_OBTENER_CONTEXTO_SEGURIDAD(tokenSesion) {
  const validacionSesion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacionSesion.VALIDA) {
    return { VALIDO: false, CODIGO: validacionSesion.CODIGO, MENSAJE: validacionSesion.MENSAJE };
  }
  const validacionRol = SEG_VALIDAR_ROL_SESION(tokenSesion);
  if (!validacionRol.VALIDO) {
    return { VALIDO: false, CODIGO: validacionRol.CODIGO, MENSAJE: validacionRol.MENSAJE };
  }
  return {
    VALIDO: true,
    CODIGO: "CONTEXTO_VALIDO",
    MENSAJE: "Contexto de seguridad recuperado con éxito.",
    SESION: validacionSesion.SESION,
    USUARIO: validacionRol.USUARIO,
    ROL: validacionRol.ROL
  };
}

/**
 * Valida múltiples permisos en un solo payload masivo.
 */
function SEG_VALIDAR_MULTIPLES_ACCESOS(tokenSesion, accesos) {
  if (!Array.isArray(accesos)) {
    throw new Error("La lista de accesos a evaluar debe ser un arreglo de objetos.");
  }
  const resultados = [];
  accesos.forEach(function(acceso) {
    const modulo = acceso.MODULO || acceso.modulo;
    const accion = acceso.ACCION || acceso.accion;
    const resultado = SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion);
    resultados.push({
      MODULO: modulo,
      ACCION: accion,
      AUTORIZADO: resultado.AUTORIZADO,
      CODIGO: resultado.CODIGO,
      MENSAJE: resultado.MENSAJE
    });
  });
  return resultados;
}

// ============================================================
// 09. REGISTRO DE AUDITORÍA (USR_AUDITORIA)
// ============================================================

/**
 * Registra un evento de auditoría en la hoja USR_AUDITORIA de manera automática y tolerante a fallos.
 */
function SEG_REGISTRAR_AUDITORIA(datos) {
  if (!SEG_CONFIG.REGISTRAR_AUDITORIA) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_AUDITORIA);
    if (!hoja) return; // Si no existe físicamente, omitimos para evitar romper procesos clave
    
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_AUDITORIA);
    const ahora = SEG_AHORA();
    const idAuditoria = SEG_GENERAR_ID(SEG_CONFIG.HOJA_AUDITORIA, "ID_AUDITORIA", SEG_CONFIG.PREFIJO_AUDITORIA);
    
    const registro = {};
    encabezados.forEach(function(campo) {
      registro[campo] = datos[campo] !== undefined ? datos[campo] : "";
    });
    
    // Forzar campos de sistema
    registro.ID_AUDITORIA = idAuditoria;
    registro.FECHA_HORA = ahora;
    registro.FECHA_CREACION = ahora;
    
    const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, registro);
    hoja.appendRow(nuevaFila);
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_REGISTRAR_AUDITORIA", "SEGURIDAD", error);
    } else {
      console.error(error);
    }
  }
}

/**
 * Recupera un listado completo con todas las sesiones registradas.
 */
function SEG_OBTENER_SESIONES() {
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);
  return registros
    .filter(function(fila) {
      return fila[0] && String(fila[0]).trim() !== "";
    })
    .map(function(fila) {
      return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
    });
}

/**
 * Registra de forma autónoma un usuario desde la pantalla web pública.
 * Por defecto se guarda con estado "PENDIENTE" y con el rol CONSULTA (ROL-000007).
 */
function SEG_REGISTRAR_AUTONOMO(datos) {
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar la información de registro.");
  }
  
  // Validaciones obligatorias para registro autónomo
  const camposObligatorios = ["USUARIO", "NOMBRE", "CORREO", "CONTRASENA_PLANA"];
  camposObligatorios.forEach(function(campo) {
    if (datos[campo] === undefined || datos[campo] === null || String(datos[campo]).trim() === "") {
      throw new Error("El campo " + campo + " es obligatorio para registrarse.");
    }
  });

  // Forzar estado PENDIENTE de aprobación y rol básico CONSULTA (ROL-000007) por defecto
  datos.ESTADO_USUARIO = SEG_CONFIG.ESTADO_USUARIO_PENDIENTE || "PENDIENTE";
  datos.ID_ROL = "ROL-000007"; // Rol de consulta básico predefinido en USR_ROLES
  datos.USUARIO_CREACION = "AUTOREGISTRO";
  
  // Usar la función de creación nativa para validar duplicados, cifrar contraseña y guardar
  const resultado = SEG_CREAR_USUARIO(datos);
  
  if (resultado.EXITO) {
    return {
      EXITO: true,
      MENSAJE: "Su solicitud de registro ha sido enviada. Por favor, espere a que un administrador apruebe su cuenta."
    };
  }
  
  return resultado;
}

// ============================================================
// FUNCIONES ADICIONALES DE INTEGRACIÓN CON EL FRONTEND
// ============================================================

/**
 * Obtiene el listado completo de todos los roles (Puente para el frontend).
 */
function SEG_OBTENER_ROLES() {
  try {
    SEG_INICIALIZAR_ROLES_PREDEFINIDOS(); // Autopoblar si está vacía
    return SEG_LISTAR_ROLES();
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_OBTENER_ROLES", "SEGURIDAD", error);
    } else {
      console.error(error);
    }
    throw new Error("No se pudieron cargar los roles de la base de datos.");
  }
}

/**
 * Obtiene el listado completo de todos los permisos parametrizados en USR_PERMISOS (Puente para el frontend).
 */
function SEG_OBTENER_PERMISOS() {
  try {
    SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS(); // Autopoblar si está vacía
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
    return registros.map(function(fila) {
      return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
    });
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_OBTENER_PERMISOS", "SEGURIDAD", error);
    } else {
      console.error(error);
    }
    throw new Error("No se pudieron cargar los permisos de la base de datos.");
  }
}

// ============================================================
// SELF-HEALING: INICIALIZACIÓN DE ROLES Y PERMISOS PREDEFINIDOS
// ============================================================

/**
 * Auto-inicializa roles del ERP si la tabla USR_ROLES está vacía.
 */
function SEG_INICIALIZAR_ROLES_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_ROLES);
    if (!hoja) return;
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila >= 2) return; // Ya existen roles, no es necesario inicializar
    
    console.log("Inicializando roles predefinidos in USR_ROLES...");
    const ahora = new Date();
    const rolesPredefinidos = [
      ["ROL-000001", "ADMINISTRADOR", "Administración general y acceso total al ERP", 100, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol maestro protegido"],
      ["ROL-000002", "CONTADOR", "Gestión y supervisión de procesos contables y financieros", 80, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a reportes contables"],
      ["ROL-000003", "AUXILIAR_CONTABLE", "Apoyo en registros y procesos contables autorizados", 60, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Permiso de registro operativo"],
      ["ROL-000004", "TESORERO", "Gestión de tesorería, pagos, recaudos y conciliaciones", 70, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a bancos y cajas"],
      ["ROL-000005", "COMERCIAL", "Gestión de clientes y operaciones comerciales", 50, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso limitado a ventas"],
      ["ROL-000006", "OPERATIVO", "Registro de operaciones asignadas e inventarios", 40, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol operativo general"],
      ["ROL-000007", "CONSULTA", "Acceso exclusivamente de lectura general", 10, "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol de sólo lectura"]
    ];
    
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_ROLES);
    rolesPredefinidos.forEach(function(filaRol) {
      const objetoRol = {
        ID_ROL: filaRol[0],
        NOMBRE_ROL: filaRol[1],
        DESCRIPCION: filaRol[2],
        NIVEL_JERARQUIA: filaRol[3],
        ESTADO_ROL: filaRol[4],
        ROL_SISTEMA: filaRol[5],
        FECHA_CREACION: filaRol[6],
        FECHA_ACTUALIZACION: filaRol[7],
        USUARIO_CREACION: filaRol[8],
        USUARIO_ACTUALIZACION: filaRol[9],
        OBSERVACIONES: filaRol[10]
      };
      const filaInsertar = SEG_CONVERTIR_OBJETO_FILA(encabezados, objetoRol);
      hoja.appendRow(filaInsertar);
    });
    console.log("Roles predefinidos creados correctamente.");
  } catch (error) {
    console.error("Error en SEG_INICIALIZAR_ROLES_PREDEFINIDOS: " + error.toString());
  }
}

/**
 * Auto-inicializa permisos básicos del ERP si la tabla USR_PERMISOS está vacía.
 */
function SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_PERMISOS);
    if (!hoja) return;
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila >= 2) return; // Ya existen permisos
    
    console.log("Inicializando permisos predefinidos in USR_PERMISOS...");
    const ahora = new Date();
    const modulos = ["SEGURIDAD", "CLIENTES", "PROVEEDORES", "PRODUCTOS", "OBRAS", "VENTAS", "COMPRAS", "INVENTARIO", "TESORERIA"];
    const accionesAdmin = ["VER", "CREAR", "EDITAR", "ELIMINAR", "ANULAR", "APROBAR", "ADMINISTRAR"];
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
    let consecutivo = 1;
    
    // 1. ADMINISTRADOR gets all permisos
    modulos.forEach(function(mod) {
      accionesAdmin.forEach(function(acc) {
        const idPermiso = "PER-" + String(consecutivo).padStart(SEG_CONFIG.DIGITOS_ID, "0");
        const objetoPermiso = {
          ID_PERMISO: idPermiso,
          ID_ROL: "ROL-000001", // ADMINISTRADOR
          MODULO: mod,
          SUBMODULO: "GENERAL",
          ACCION: acc,
          PERMITIDO: "SI",
          ESTADO_PERMISO: "ACTIVO",
          FECHA_CREACION: ahora,
          FECHA_ACTUALIZACION: ahora,
          USUARIO_CREACION: "SISTEMA",
          USUARIO_ACTUALIZACION: "SISTEMA",
          OBSERVACIONES: "Permiso administrativo maestro"
        };
        const filaInsertar = SEG_CONVERTIR_OBJETO_FILA(encabezados, objetoPermiso);
        hoja.appendRow(filaInsertar);
        consecutivo++;
      });
    });
    
    // 2. CONSULTA gets VER only
    modulos.forEach(function(mod) {
      const idPermiso = "PER-" + String(consecutivo).padStart(SEG_CONFIG.DIGITOS_ID, "0");
      const objetoPermiso = {
        ID_PERMISO: idPermiso,
        ID_ROL: "ROL-000007", // CONSULTA
        MODULO: mod,
        SUBMODULO: "GENERAL",
        ACCION: "VER",
        PERMITIDO: "SI",
        ESTADO_PERMISO: "ACTIVO",
        FECHA_CREACION: ahora,
        FECHA_ACTUALIZACION: ahora,
        USUARIO_CREACION: "SISTEMA",
        USUARIO_ACTUALIZACION: "SISTEMA",
        OBSERVACIONES: "Permiso de sólo lectura"
      };
      const filaInsertar = SEG_CONVERTIR_OBJETO_FILA(encabezados, objetoPermiso);
      hoja.appendRow(filaInsertar);
      consecutivo++;
    });
    console.log("Permisos predefinidos creados correctamente.");
  } catch (error) {
    console.error("Error en SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS: " + error.toString());
  }
}

/**
 * Aprueba un usuario pendiente asignándole su rol definitivo y activándolo en un solo paso atómico.
 */
function SEG_APROBAR_USUARIO(idUsuario, idRol, usuarioActualizador) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rol = SEG_CONSULTAR_ROL(idRol);
    if (!rol) {
      throw new Error("El rol seleccionado '" + idRol + "' no es válido o no existe en USR_ROLES.");
    }
    
    const resultadoRol = SEG_ACTUALIZAR_USUARIO(idUsuario, {
      ID_ROL: idRol,
      USUARIO_ACTUALIZACION: usuarioActualizador || "SISTEMA"
    });
    if (!resultadoRol.EXITO) {
      throw new Error("No se pudo actualizar el rol del usuario.");
    }
    
    const resultadoEstado = SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, "ACTIVO", usuarioActualizador || "SISTEMA");
    if (!resultadoEstado.EXITO) {
      throw new Error("No se pudo activar el estado del usuario.");
    }
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: idUsuario,
      USUARIO: resultadoRol.USUARIO.USUARIO,
      MODULO: "SEGURIDAD",
      SUBMODULO: "USUARIOS",
      ACCION: "APROBAR",
      TIPO_REGISTRO: "USUARIOS",
      ID_REGISTRO: idUsuario,
      DESCRIPCION: "Usuario '" + resultadoRol.USUARIO.USUARIO + "' aprobado con éxito. Rol asignado: " + rol.NOMBRE_ROL + ".",
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      MENSAJE: "Usuario '" + resultadoRol.USUARIO.USUARIO + "' aprobado con éxito. Rol asignado: " + rol.NOMBRE_ROL + "."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_APROBAR_USUARIO", "SEGURIDAD", error);
    } else {
      console.error(error);
    }
    throw new Error(error.message || error.toString());
  }
}

// ============================================================
// 10. NUEVOS MÉTODOS: LOGO Y RECUPERACIÓN DE CONTRASEÑA
// ============================================================

/**
 * Guarda el logotipo de la empresa en Drive y actualiza la URL en la hoja de configuración CFG_EMPRESA.
 * RESTRICCIÓN: Solo ejecutable si se llama desde la interfaz de Sheets por un administrador.
 */
function SEG_GUARDAR_LOGO(base64Data, nombreArchivo) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (!ui) {
      throw new Error("Esta operación solo está permitida para administradores dentro de Google Sheets.");
    }
    
    const splitData = base64Data.split(",");
    const contentType = splitData[0].match(/:(.*?);/)[1];
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
    if (!hojaEmpresa) {
      throw new Error("No se encontró la hoja de configuración CFG_EMPRESA.");
    }
    
    const datos = hojaEmpresa.getRange(2, 1, hojaEmpresa.getLastRow() - 1, 2).getValues();
    let filaModificar = -1;
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0].toString().trim().toUpperCase() === "LOGO_URL" || datos[i][0].toString().trim().toUpperCase() === "LOGO") {
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
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_GUARDAR_LOGO", "SEGURIDAD", error);
    }
    throw new Error(error.message || error.toString());
  }
}

/**
 * Busca un usuario por correo, genera una contraseña temporal de un solo uso,
 * actualiza el Sheets y envía un correo asíncrono con las instrucciones de acceso.
 */
function SEG_SOLICITAR_RECUPERACION_CONTRASENA(correo) {
  try {
    if (!correo || String(correo).trim() === "") {
      throw new Error("Debe proporcionar un correo electrónico válido.");
    }
    const correoNormalizado = String(correo).trim().toLowerCase();
    
    const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", correoNormalizado);
    if (!usuario) {
      return { EXITO: false, MENSAJE: "El correo electrónico no se encuentra registrado en el ERP de MEGUDAN." };
    }
    
    const caracteres = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#%";
    let contrasenaTemporal = "";
    contrasenaTemporal += "M";
    contrasenaTemporal += "9";
    contrasenaTemporal += "!";
    for (let i = 0; i < 7; i++) {
      const randIdx = Math.floor(Math.random() * caracteres.length);
      contrasenaTemporal += caracteres.charAt(randIdx);
    }
    
    SEG_ESTABLECER_CONTRASENA(usuario.ID_USUARIO, contrasenaTemporal, "SISTEMA_RECUPERACION");
    SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, { DEBE_CAMBIAR_CONTRASENA: "SI" });
    
    const asunto = "Recuperación de Acceso - MEGUDAN ERP";
    const cuerpoHTML = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 500px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        <div style="background-color: #374151; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0; font-size: 20px;">MEGUDAN ERP</h2>
          <p style="margin: 5px 0 0; font-size: 12px; opacity: 0.8;">Sistema de Gestión Operativa</p>
        </div>
        <div style="padding: 25px; color: #1f2937; line-height: 1.6;">
          <p style="margin: 0 0 15px;">Hola <strong>\${usuario.NOMBRE || usuario.NOMBRE_COMPLETO || usuario.USUARIO}</strong>,</p>
          <p style="margin: 0 0 15px;">Hemos recibido una solicitud para restablecer la contraseña de tu cuenta en el ERP de MEGUDAN.</p>
          <div style="background-color: #f3f4f6; border-left: 4px solid #374151; padding: 15px; margin: 20px 0; text-align: center; border-radius: 4px;">
            <p style="margin: 0 0 5px; font-size: 11px; color: #6b7280; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px;">Tu contraseña temporal de acceso:</p>
            <code style="font-size: 18px; font-weight: bold; color: #111827; background: #e5e7eb; padding: 4px 10px; border-radius: 4px; display: inline-block; letter-spacing: 1px;">\${contrasenaTemporal}</code>
          </div>
          <p style="margin: 0 0 15px; font-size: 13px; color: #dc2626; font-weight: bold;">⚠️ Por seguridad, el sistema te exigirá cambiar esta contraseña temporal inmediatamente al iniciar sesión.</p>
          <p style="margin: 0 0 15px;">Si tú no solicitaste este cambio, por favor ponte en contacto de inmediato con el administrador de seguridad del sistema.</p>
          <p style="margin: 20px 0 0; font-size: 12px; color: #9ca3af; border-top: 1px dashed #e5e7eb; padding-top: 15px; text-align: center;">MEGUDAN CONSTRUCCIONES SOSTENIBLES SAS &copy; 2026</p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: correoNormalizado,
      subject: asunto,
      htmlBody: cuerpoHTML
    });
    
    return { EXITO: true, MENSAJE: "¡Se ha enviado una contraseña temporal a su correo de registro con éxito!" };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_SOLICITAR_RECUPERACION_CONTRASENA", "SEGURIDAD", error);
    }
    return { EXITO: false, MENSAJE: "Ocurrió un error al procesar la solicitud: " + error.message };
  }
}
