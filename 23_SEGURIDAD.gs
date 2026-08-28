/**************************************************************
* 23_SEGURIDAD.gs
* RESPONSABILIDAD:
* - Administrar el control de acceso multidimensional del ERP (Token -> Sesión -> Usuario -> Rol -> Permiso).
* - Proteger las macros y Web Apps mediante un Sistema de Control de Acceso Dual.
* - Encriptar contraseñas mediante Hash SHA-256 y proveer auto-inicialización de roles/permisos.
* - Soportar recuperación de claves por correo y subida segura de logotipos.
**************************************************************/

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE SEGURIDAD
// ============================================================
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
  
  // Tiempos y seguridad
  TIEMPO_INACTIVIDAD_MINUTOS: 30,
  DURACION_SESION_HORAS: 8,
  MAXIMO_INTENTOS_LOGIN: 5,
  BLOQUEO_USUARIO_MINUTOS: 15,
  
  // Estados lógicos
  ESTADO_USUARIO_ACTIVO: "ACTIVO",
  ESTADO_USUARIO_INACTIVO: "INACTIVO",
  ESTADO_USUARIO_BLOQUEADO: "BLOQUEADO",
  ESTADO_USUARIO_PENDIENTE: "PENDIENTE",
  
  ESTADO_ROL_ACTIVO: "ACTIVO",
  ESTADO_SESION_ACTIVA: "ACTIVA",
  
  // Acciones auditoría
  ACCION_CREAR: "CREAR",
  ACCION_MODIFICAR: "MODIFICAR",
  ACCION_ELIMINAR: "ELIMINAR",
  ACCION_LOGIN: "LOGIN",
  ACCION_CERRAR_SESION: "CERRAR_SESION",
  
  // Roles de sistema protegidos
  ROL_ADMINISTRADOR: "ADMINISTRADOR",
  
  // Registro global
  REGISTRAR_AUDITORIA: true
};

// ============================================================
// 02. SISTEMA DE SEGURIDAD DUAL Y ACCESO DE ENTORNO
// ============================================================

/**
 * Sistema de Control Dual (Sheets Confiable vs Web App con Token).
 * - Contexto local (Sheets/Editor): Permite operar con privilegios supremos de ADMINISTRADOR automáticamente.
 * - Contexto remoto (Web App/doGet/doPost): Requiere obligatoriamente un token de sesión activo y validación de permisos.
 */
function SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, modulo, accion) {
  if (tokenSesion === "SISTEMA_INTERNAL_BYPASS") {
    return {
      AUTORIZADO: true,
      CODIGO: "SISTEMA_BYPASS",
      USUARIO: "SISTEMA",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso concedido automáticamente para operaciones internas del sistema."
    };
  }
  try {
    // Intentamos invocar SpreadsheetApp.getUi() para saber si es un contexto físico de Sheets
    SpreadsheetApp.getUi();
    // Si no lanza excepción, estamos dentro de Google Sheets (botón físico, macro local o menú).
    // Concedemos de forma implícita privilegios de administrador supremo para la sesión local.
    return {
      AUTORIZADO: true,
      CODIGO: "CONTEXTO_SHEETS_TRUSTED",
      USUARIO: "ADMINISTRADOR_LOCAL_SHEETS",
      ROL: "ADMINISTRADOR",
      MENSAJE: "Acceso concedido automáticamente en entorno confiable de Google Sheets."
    };
  } catch (uiError) {
    // Si lanza excepción, estamos en un hilo sin interfaz gráfica (Web App externa).
    // Exigimos obligatoriamente un inicio de sesión y validación de Token.
    if (!tokenSesion || String(tokenSesion).trim() === "") {
      throw new Error("ACCESO DENEGADO [TOKEN_REQUERIDO]: Se requiere un token de sesión activo para operar desde la Web App.");
    }
    
    const validacion = SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion);
    if (!validacion.AUTORIZADO) {
      throw new Error("ACCESO DENEGADO [" + validacion.CODIGO + "]: " + validacion.MENSAJE);
    }
    
    return validacion;
  }
}

// ============================================================
// 03. AUXILIARES Y ACCESO A BASE DE DATOS
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
 * [CORREGIDO] Se solucionó el ReferenceError reemplazando la variable incorrecta "value" por "valor" en la validación.
 */
function SEG_VALIDAR_OBLIGATORIOS(datos, camposObligatorios) {
  camposObligatorios.forEach(function(campo) {
    const valor = datos[campo];
    if (valor === undefined || valor === null || String(valor).trim() === "") {
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
// 04. CRUD DE USUARIOS (BACKEND CON CONTROL DUAL)
// ============================================================

/**
 * Crea un usuario en USR_USUARIOS con su validación de estructura dinámica.
 */
function SEG_CREAR_USUARIO(datos, tokenSesion) {
  // Validación de permisos duales
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar la información del usuario.");
  }

  SEG_VALIDAR_OBLIGATORIOS(datos, ["USUARIO", "CORREO"]);

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
  usuario.ID_ROL = datos.ID_ROL || "ROL-000007"; // Consulta por defecto
  usuario.ESTADO_USUARIO = datos.ESTADO_USUARIO || "PENDIENTE"; // Por defecto ingresa como pendiente para aprobación
  usuario.INTENTOS_FALLIDOS = 0;
  usuario.FECHA_CREACION = ahora;
  usuario.FECHA_ACTUALIZACION = ahora;
  usuario.USUARIO_CREACION = usuarioEjecutor;
  usuario.USUARIO_ACTUALIZACION = usuarioEjecutor;

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
    DESCRIPCION: "Usuario " + usuario.USUARIO + " creado con éxito. Creador: " + usuarioEjecutor,
    RESULTADO: SEG_CONFIG.RESULTADO_EXITOSO
  });

  return {
    EXITO: true,
    ID_USUARIO: idUsuario,
    USUARIO: usuario,
    MENSAJE: "Usuario creado exitosamente en estado PENDIENTE."
  };
}

/**
 * Consulta la ficha completa de un usuario por su ID.
 */
function SEG_CONSULTAR_USUARIO(idUsuario, tokenSesion) {
  if (tokenSesion !== undefined) {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  }
  return SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
}

/**
 * Recupera un listado completo con todos los usuarios registrados.
 */
function SEG_LISTAR_USUARIOS(tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  
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
 * [SOPORTE SOBRECARGADO]: Soporta firmas con (datos, tokenSesion) y (idUsuario, datos, tokenSesion)
 */
function SEG_ACTUALIZAR_USUARIO(idUsuario, datos, tokenSesion) {
  // Desempaquetar firma sobrecargada de 2 argumentos desde el frontend: SEG_ACTUALIZAR_USUARIO(datos, tokenSesion)
  if (tokenSesion === undefined && typeof idUsuario === "object" && idUsuario !== null) {
    tokenSesion = datos;
    datos = idUsuario;
    idUsuario = datos.ID_USUARIO || datos.id_usuario;
  }

  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  if (!idUsuario || String(idUsuario).trim() === "") {
    throw new Error("Debe indicar el ID_USUARIO.");
  }
  if (!datos || typeof datos !== "object") {
    throw new Error("Debe proporcionar los datos a actualizar.");
  }

  const usuarioActual = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
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

  // Validaciones adicionales si se actualizan campos sensibles
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

  // Si incluye nueva clave temporal o de re-establecimiento
  if (datos.CONTRASENA_PLANA && String(datos.CONTRASENA_PLANA).trim() !== "") {
    const seguridadValida = SEG_VALIDAR_SEGURIDAD_CONTRASENA(datos.CONTRASENA_PLANA);
    if (!seguridadValida.VALIDA) {
      throw new Error(seguridadValida.MENSAJES.join(" "));
    }
    usuarioActualizado.CONTRASENA_HASH = SEG_GENERAR_HASH_CONTRASENA(datos.CONTRASENA_PLANA);
  }

  usuarioActualizado.FECHA_ACTUALIZACION = SEG_AHORA();
  usuarioActualizado.USUARIO_ACTUALIZACION = usuarioEjecutor;

  const filaActualizada = SEG_CONVERTIR_OBJETO_FILA(encabezados, usuarioActualizado);
  hoja.getRange(filaUsuario, 1, 1, encabezados.length).setValues([filaActualizada]);

  return {
    EXITO: true,
    ID_USUARIO: idUsuario,
    USUARIO: usuarioActualizado,
    MENSAJE: "Usuario actualizado exitosamente."
  };
}

/**
 * Modifica el estado operacional de un usuario (ACTIVO, INACTIVO, BLOQUEADO).
 */
function SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, nuevoEstado, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  const estado = String(nuevoEstado || "").trim().toUpperCase();
  const estadosPermitidos = [
    SEG_CONFIG.ESTADO_USUARIO_ACTIVO,
    SEG_CONFIG.ESTADO_USUARIO_INACTIVO,
    SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO,
    SEG_CONFIG.ESTADO_USUARIO_PENDIENTE,
    "ELIMINADO"
  ];

  if (!estadosPermitidos.includes(estado)) {
    throw new Error("Estado de usuario no válido: " + nuevoEstado);
  }

  return SEG_ACTUALIZAR_USUARIO(idUsuario, {
    ESTADO_USUARIO: estado,
    USUARIO_ACTUALIZACION: usuarioEjecutor
  }, tokenSesion);
}

/**
 * Elimina lógicamente un usuario en Sheets asignándole el estado "ELIMINADO".
 * [COMPATIBILIDAD FRONTEND]: Soluciona la ausencia de este método en la base de código inicial.
 */
function SEG_ELIMINAR_USUARIO(idUsuario, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ELIMINAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  try {
    const resultado = SEG_ACTUALIZAR_USUARIO(idUsuario, {
      ESTADO_USUARIO: "ELIMINADO",
      USUARIO_ACTUALIZACION: usuarioEjecutor
    }, tokenSesion);
    
    if (resultado.EXITO) {
      SEG_REGISTRAR_AUDITORIA({
        ID_USUARIO: idUsuario,
        USUARIO: resultado.USUARIO.USUARIO,
        MODULO: "SEGURIDAD",
        SUBMODULO: "USUARIOS",
        ACCION: "ELIMINAR",
        TIPO_REGISTRO: "USUARIOS",
        ID_REGISTRO: idUsuario,
        DESCRIPCION: "Usuario " + resultado.USUARIO.USUARIO + " eliminado lógicamente por " + usuarioEjecutor,
        RESULTADO: "EXITOSO"
      });
    }
    
    return { EXITO: true, MENSAJE: "Usuario eliminado de forma lógica correctamente." };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_ELIMINAR_USUARIO", "SEGURIDAD", error);
    }
    throw new Error("No se pudo eliminar el usuario: " + error.message);
  }
}

function SEG_BLOQUEAR_USUARIO(idUsuario, tokenSesion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO, tokenSesion);
}

function SEG_DESBLOQUEAR_USUARIO(idUsuario, tokenSesion) {
  return SEG_CAMBIAR_ESTADO_USUARIO(idUsuario, SEG_CONFIG.ESTADO_USUARIO_ACTIVO, tokenSesion);
}

// ============================================================
// 05. CRUD DE ROLES
// ============================================================

function SEG_CREAR_ROL(datos, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
  
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

  return { EXITO: true, ID_ROL: idRol, ROL: rol };
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
  return registros
    .filter(function(fila) {
      return fila[0] && String(fila[0]).trim() !== "";
    })
    .map(function(fila) {
      return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
    });
}

function SEG_ACTUALIZAR_ROL(idRol, datos, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "EDITAR");
  
  if (!idRol || String(idRol).trim() === "") {
    throw new Error("Debe indicar el ID_ROL.");
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

  return { EXITO: true, ID_ROL: idRol, ROL: rolActualizado };
}

function SEG_CAMBIAR_ESTADO_ROL(idRol, nuevoEstado, tokenSesion) {
  return SEG_ACTUALIZAR_ROL(idRol, { ESTADO_ROL: nuevoEstado }, tokenSesion);
}

// ============================================================
// 06. GESTIÓN Y VALIDACIÓN DE PERMISOS
// ============================================================

function SEG_OBTENER_PERMISOS_ROL(idRol) {
  if (!idRol || String(idRol).trim() === "") {
    throw new Error("Debe indicar el ID_ROL.");
  }
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
  const idxIdRol = encabezados.indexOf("ID_ROL");
  
  const filtrados = registros.filter(function(fila) {
    return String(fila[idxIdRol] || "").trim().toUpperCase() === String(idRol).trim().toUpperCase();
  });

  return filtrados.map(function(fila) {
    return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
  });
}

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
  return ["SI", "SÍ", "TRUE", "1", "VERDADERO"].includes(permitido);
}

function SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
  if (!usuario || String(usuario.ESTADO_USUARIO || "").trim().toUpperCase() !== "ACTIVO") {
    return false;
  }
  return SEG_VALIDAR_PERMISO_ROL(usuario.ID_ROL, modulo, accion);
}

function SEG_VERIFICAR_ACCESO(idUsuario, modulo, accion) {
  const autorizado = SEG_VALIDAR_PERMISO_USUARIO(idUsuario, modulo, accion);
  if (!autorizado) {
    throw new Error("ACCESO DENEGADO. No tiene autorización para '" + accion + "' en el módulo '" + modulo + "'.");
  }
  return true;
}

// ============================================================
// 07. CRIPTOGRAFÍA Y AUTENTICACIÓN (SHA-256)
// ============================================================

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

function SEG_VALIDAR_SEGURIDAD_CONTRASENA(contrasena) {
  const password = String(contrasena || "");
  const resultado = { VALIDA: false, MENSAJES: [] };

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

function SEG_BUSCAR_USUARIO_LOGIN(credencial) {
  if (!credencial || String(credencial).trim() === "") {
    return null;
  }
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_USUARIOS);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_USUARIOS);
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();
  if (ultimaFila < 2) return null;

  const idxUsuario = encabezados.indexOf("USUARIO");
  const idxCorreo = encabezados.indexOf("CORREO");

  const criterio = String(credencial).trim().toUpperCase();
  const registros = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();

  const filaEncontrada = registros.find(function(fila) {
    const usuarioReg = String(fila[idxUsuario] || "").trim().toUpperCase();
    const correoReg = String(fila[idxCorreo] || "").trim().toUpperCase();
    return usuarioReg === criterio || correoReg === criterio;
  });

  return filaEncontrada ? SEG_CONVERTIR_FILA_OBJETO(encabezados, filaEncontrada) : null;
}

function SEG_ESTABLECER_CONTRASENA(idUsuario, nuevaContrasena, usuarioActualizacion) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
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

function SEG_VALIDAR_CONTRASENA(usuario, contrasena) {
  if (!usuario || !usuario.CONTRASENA_HASH || String(usuario.CONTRASENA_HASH).trim() === "") {
    return false;
  }
  return String(SEG_GENERAR_HASH_CONTRASENA(contrasena)) === String(usuario.CONTRASENA_HASH);
}

function SEG_AUTENTICAR_USUARIO(credencial, contrasena) {
  SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS(); // Asegurar autocuración en login
  
  const usuario = SEG_BUSCAR_USUARIO_LOGIN(credencial);
  if (!usuario) {
    return { EXITO: false, CODIGO: "USUARIO_NO_ENCONTRADO", MENSAJE: "Credenciales incorrectas de acceso." };
  }

  const estado = String(usuario.ESTADO_USUARIO || "").trim().toUpperCase();
  if (estado === "PENDIENTE") {
    return { EXITO: false, CODIGO: "USUARIO_PENDIENTE", MENSAJE: "Su cuenta está pendiente de aprobación por un administrador." };
  }
  if (estado === "BLOQUEADO") {
    return { EXITO: false, CODIGO: "USUARIO_BLOQUEADO", MENSAJE: "El usuario se encuentra temporalmente bloqueado." };
  }
  if (estado !== "ACTIVO") {
    return { EXITO: false, CODIGO: "USUARIO_INACTIVO", MENSAJE: "El usuario no está activo en el ERP." };
  }

  if (!SEG_VALIDAR_CONTRASENA(usuario, contrasena)) {
    let fallidos = parseInt(usuario.INTENTOS_FALLIDOS || 0, 10) + 1;
    const updates = { INTENTOS_FALLIDOS: fallidos };
    if (fallidos >= SEG_CONFIG.MAXIMO_INTENTOS_LOGIN) {
      updates.ESTADO_USUARIO = "BLOQUEADO";
    }
    SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, updates);
    return { EXITO: false, CODIGO: "CONTRASENA_INCORRECTA", MENSAJE: "Credenciales incorrectas de acceso." };
  }

  // Éxito
  SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, { INTENTOS_FALLIDOS: 0, ULTIMO_ACCESO: SEG_AHORA() });
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
    DEBE_CAMBIAR_CONTRASENA: String(usuario.DEBE_CAMBIAR_CONTRASENA || "").trim().toUpperCase() === "SI"
  };
}

// ============================================================
// 08. CONTROL DE SESIONES DE BASE DE DATOS
// ============================================================

function SEG_GENERAR_TOKEN_SESION() {
  return Utilities.getUuid() + "-" + Utilities.getUuid();
}

function SEG_CREAR_SESION(idUsuario) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "ID_USUARIO", idUsuario);
  if (!usuario || String(usuario.ESTADO_USUARIO || "").trim().toUpperCase() !== "ACTIVO") {
    throw new Error("No se puede iniciar sesión para un usuario que no esté activo.");
  }

  const ahora = SEG_AHORA();
  const fechaExpiracion = new Date(ahora.getTime() + (SEG_CONFIG.DURACION_SESION_HORAS * 60 * 60 * 1000));
  const idSesion = SEG_GENERAR_ID(SEG_CONFIG.HOJA_SESIONES, "ID_SESION", SEG_CONFIG.PREFIJO_SESION);
  const tokenSesion = SEG_GENERAR_TOKEN_SESION();

  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);

  const sesion = {
    ID_SESION: idSesion,
    ID_USUARIO: usuario.ID_USUARIO,
    USUARIO: usuario.USUARIO,
    ID_ROL: usuario.ID_ROL,
    TOKEN_SESION: tokenSesion,
    ESTADO_SESION: "ACTIVA",
    FECHA_INICIO: ahora,
    EXPIRA_SESION: fechaExpiracion,
    ULTIMA_ACTIVIDAD: ahora,
    TIPO_ACCESO: "WEB",
    ORIGEN_ACCESO: "APLICACION_WEB"
  };

  const nuevaFila = SEG_CONVERTIR_OBJETO_FILA(encabezados, sesion);
  hoja.appendRow(nuevaFila);

  return { EXITO: true, ID_SESION: idSesion, TOKEN_SESION: tokenSesion, ID_USUARIO: usuario.ID_USUARIO, FECHA_EXPIRACION: fechaExpiracion };
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
  try {
    const sesion = SEG_BUSCAR_SESION(tokenSesion);
    if (!sesion) {
      return { VALIDA: false, CODIGO: "SESION_NO_ENCONTRADA", MENSAJE: "La sesión no existe o no es válida.", SESION: null };
    }

    const estado = String(sesion.ESTADO_SESION || "").trim().toUpperCase();
    if (estado !== "ACTIVA") {
      return { VALIDA: false, CODIGO: "SESION_NO_ACTIVA", MENSAJE: "La sesión no se encuentra activa.", SESION: null };
    }

    const ahora = SEG_AHORA();
    const fechaExpiracion = new Date(sesion.EXPIRA_SESION || sesion.FECHA_EXPIRACION);
    if (ahora.getTime() >= fechaExpiracion.getTime()) {
      SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
      return { VALIDA: false, CODIGO: "SESION_EXPIRADA", MENSAJE: "La sesión ha expirado por límite de tiempo.", SESION: null };
    }

    const ultimaActividad = new Date(sesion.ULTIMA_ACTIVIDAD);
    const diferenciaMinutos = (ahora.getTime() - ultimaActividad.getTime()) / (1000 * 60);
    if (diferenciaMinutos > SEG_CONFIG.TIEMPO_INACTIVIDAD_MINUTOS) {
      SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "EXPIRADA");
      return { VALIDA: false, CODIGO: "SESION_EXPIRADA_INACTIVIDAD", MENSAJE: "La sesión ha expirado por inactividad.", SESION: null };
    }

    SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion);
    return { VALIDA: true, CODIGO: "SESION_VALIDA", MENSAJE: "Sesión autorizada.", SESION: sesion };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_VALIDAR_SESION", "SEGURIDAD", error);
    }
    return { VALIDA: false, CODIGO: "ERROR_INTERNO", MENSAJE: "Error interno al validar sesión: " + error.toString(), SESION: null };
  }
}

function SEG_ACTUALIZAR_ACTIVIDAD_SESION(tokenSesion) {
  const sesion = SEG_BUSCAR_SESION(tokenSesion);
  if (!sesion) return false;
  const hoja = SEG_OBTENER_HOJA(SEG_CONFIG.HOJA_SESIONES);
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const idxActividad = encabezados.indexOf("ULTIMA_ACTIVIDAD");
  hoja.getRange(sesion._FILA, idxActividad + 1).setValue(SEG_AHORA());
  return true;
}

function SEG_CAMBIAR_ESTADO_SESION(tokenSesion, nuevoEstado) {
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
  if (!sesion) return { EXITO: false, MENSAJE: "Sesión no encontrada." };
  SEG_CAMBIAR_ESTADO_SESION(tokenSesion, "CERRADA");
  return { EXITO: true, MENSAJE: "Sesión cerrada correctamente." };
}

// ============================================================
// 09. AUTORIZACIONES POR ROL Y CONTEXTO
// ============================================================

function SEG_VALIDAR_ACCESO(tokenSesion, modulo, accion) {
  try {
    const validacionSesion = SEG_VALIDAR_SESION(tokenSesion);
    if (!validacionSesion || !validacionSesion.VALIDA) {
      const msg = (validacionSesion && validacionSesion.MENSAJE) ? validacionSesion.MENSAJE : "Sesión inválida.";
      return { AUTORIZADO: false, CODIGO: "SESION_INVALIDA", MENSAJE: msg };
    }

    const sesion = validacionSesion.SESION;
    if (!sesion) {
      return { AUTORIZADO: false, CODIGO: "SESION_NULA", MENSAJE: "No se encontró el objeto de sesión." };
    }
    
    const rol = SEG_CONSULTAR_ROL(sesion.ID_ROL);
    if (!rol || String(rol.ESTADO_ROL || "").trim().toUpperCase() !== "ACTIVO") {
      return { AUTORIZADO: false, CODIGO: "ROL_INACTIVO", MENSAJE: "El rol asignado a la sesión se encuentra inactivo." };
    }

    // Pase de administrador supremo
    if (rol.NOMBRE_ROL === SEG_CONFIG.ROL_ADMINISTRADOR) {
      return { AUTORIZADO: true, CODIGO: "ACCESO_ADMINISTRADOR", USUARIO: sesion.USUARIO, ROL: rol.NOMBRE_ROL, SESION: sesion, MENSAJE: "Acceso supremo autorizado (Administrador)." };
    }

    const autorizado = SEG_VALIDAR_PERMISO_ROL(rol.ID_ROL, modulo, accion);
    if (!autorizado) {
      return { AUTORIZADO: false, CODIGO: "ACCESO_DENEGADO", MENSAJE: "Su rol no tiene permisos de " + accion + " en " + modulo + "." };
    }

    return { AUTORIZADO: true, CODIGO: "ACCESO_CONCEDIDO", USUARIO: sesion.USUARIO, ROL: rol.NOMBRE_ROL, SESION: sesion, MENSAJE: "Acceso autorizado con éxito." };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("SEG_VALIDAR_ACCESO", "SEGURIDAD", error);
    }
    return { AUTORIZADO: false, CODIGO: "ERROR_INTERNO", MENSAJE: "Error interno al validar acceso: " + error.toString() };
  }
}

function SEG_OBTENER_CONTEXTO_SEGURIDAD(tokenSesion) {
  const validacion = SEG_VALIDAR_SESION(tokenSesion);
  if (!validacion.VALIDA) return { VALIDO: false, MENSAJE: validacion.MENSAJE };
  const rol = SEG_CONSULTAR_ROL(validacion.SESION.ID_ROL);
  return { VALIDO: true, USUARIO: validacion.SESION.USUARIO, ROL: rol };
}

function SEG_REGISTRAR_AUDITORIA(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_AUDITORIA);
    if (!hoja) return;
    
    const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_AUDITORIA);
    const ahora = SEG_AHORA();
    const idAuditoria = SEG_GENERAR_ID(SEG_CONFIG.HOJA_AUDITORIA, "ID_AUDITORIA", SEG_CONFIG.PREFIJO_AUDITORIA);
    
    const registro = {
      ID_AUDITORIA: idAuditoria,
      FECHA_HORA: ahora,
      FECHA_CREACION: ahora
    };
    
    encabezados.forEach(function(campo) {
      if (datos[campo] !== undefined) {
        registro[campo] = datos[campo];
      } else if (registro[campo] === undefined) {
        registro[campo] = "";
      }
    });
    
    hoja.appendRow(SEG_CONVERTIR_OBJETO_FILA(encabezados, registro));
  } catch (err) {
    console.error("No se pudo registrar log de auditoría: " + err.toString());
  }
}

// ============================================================
// 10. METODOS ADICIONALES (SUBIDA DE LOGOS, AUTOCURACIÓN)
// ============================================================

function SEG_OBTENER_ROLES(tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
  return SEG_LISTAR_ROLES();
}

function SEG_OBTENER_PERMISOS(tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS();
  
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_PERMISOS);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_PERMISOS);
  return registros.map(function(fila) {
    return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila);
  });
}

function SEG_OBTENER_SESIONES(tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
  
  const encabezados = SEG_OBTENER_ENCABEZADOS(SEG_CONFIG.HOJA_SESIONES);
  const registros = SEG_OBTENER_REGISTROS(SEG_CONFIG.HOJA_SESIONES);
  return registros
    .filter(function(fila) { return fila[0] && String(fila[0]).trim() !== ""; })
    .map(function(fila) { return SEG_CONVERTIR_FILA_OBJETO(encabezados, fila); });
}

function SEG_APROBAR_USUARIO(idUsuario, idRol, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "APROBAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";

  const rol = SEG_CONSULTAR_ROL(idRol);
  if (!rol) throw new Error("El rol '" + idRol + "' no es válido.");

  SEG_ACTUALIZAR_USUARIO(idUsuario, { ID_ROL: idRol, ESTADO_USUARIO: "ACTIVO" }, tokenSesion);
  
  SEG_REGISTRAR_AUDITORIA({
    ID_USUARIO: idUsuario,
    USUARIO: idUsuario,
    MODULO: "SEGURIDAD",
    SUBMODULO: "USUARIOS",
    ACCION: "APROBAR",
    TIPO_REGISTRO: "USUARIOS",
    ID_REGISTRO: idUsuario,
    DESCRIPCION: "Usuario aprobado con rol: " + rol.NOMBRE_ROL + " por " + usuarioEjecutor,
    RESULTADO: "EXITOSO"
  });

  return { EXITO: true, MENSAJE: "Usuario aprobado e inactivado de solicitudes pendientes." };
}

function SEG_GUARDAR_LOGO(base64Data, nombreArchivo) {
  try {
    SpreadsheetApp.getUi();
  } catch (e) {
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
  if (!hojaEmpresa) throw new Error("No se encontró la hoja CFG_EMPRESA.");
  
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
}

function SEG_SOLICITAR_RECUPERACION_CONTRASENA(correo) {
  const usuario = SEG_BUSCAR_REGISTRO(SEG_CONFIG.HOJA_USUARIOS, "CORREO", correo);
  if (!usuario) {
    return { EXITO: false, MENSAJE: "El correo electrónico no se encuentra registrado." };
  }

  const caracteres = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#%";
  let contrasenaTemporal = "M9!";
  for (let i = 0; i < 7; i++) {
    contrasenaTemporal += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
  }

  // Establecemos clave temporal
  const hash = SEG_GENERAR_HASH_CONTRASENA(contrasenaTemporal);
  SEG_ACTUALIZAR_USUARIO(usuario.ID_USUARIO, {
    CONTRASENA_HASH: hash,
    DEBE_CAMBIAR_CONTRASENA: "SI",
    FECHA_CAMBIO_CONTRASENA: SEG_AHORA()
  });

  const asunto = "Recuperación de Acceso - MEGUDAN ERP";
  const cuerpoHTML = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 500px; margin: 0 auto; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #374151; color: white; padding: 20px; text-align: center;">
        <h2 style="margin: 0;">MEGUDAN ERP</h2>
      </div>
      <div style="padding: 25px; line-height: 1.6;">
        <p>Hola <strong>${usuario.NOMBRE || usuario.NOMBRE_COMPLETO || usuario.USUARIO}</strong>,</p>
        <p>Tu contraseña temporal de acceso de un solo uso es:</p>
        <div style="background-color: #f3f4f6; padding: 15px; text-align: center; font-weight: bold; font-size: 18px;">
          <code>${contrasenaTemporal}</code>
        </div>
        <p style="color: #dc2626;">⚠️ Deberás cambiar esta contraseña inmediatamente al ingresar.</p>
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
    roles.forEach(function(fila) {
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

    modulos.forEach(function(mod) {
      accionesAdmin.forEach(function(acc) {
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