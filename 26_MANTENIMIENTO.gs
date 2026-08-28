/**************************************************************
* 26_MANTENIMIENTO.gs
* RESPONSABILIDAD:
* - Motor centralizado de limpieza, depuración y mantenimiento.
* - Truncar de forma segura tablas de historial (Auditoría, Sesiones).
* - Administrar y purgar el CacheService de Google Apps Script.
**************************************************************/

const MNT_CONFIG = {
  HOJA_AUDITORIA: "USR_AUDITORIA",
  HOJA_SESIONES: "USR_SESIONES",
  HOJA_CLI_HISTORIAL: "CLI_HISTORIAL",
  HOJA_PROV_HISTORIAL: "PROV_HISTORIAL"
};

function MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion) {
  if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function") {
    return SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ADMINISTRAR");
  }
  try {
    SpreadsheetApp.getUi();
    return { AUTORIZADO: true, USUARIO: "ADMINISTRADOR_LOCAL_SHEETS" };
  } catch (e) {
    throw new Error("ACCESO DENEGADO: No se pudo verificar el contexto de seguridad.");
  }
}

function MNT_PURGAR_CACHE_SISTEMA(tokenSesion) {
  try {
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    const scriptCache = CacheService.getScriptCache();
    const documentCache = CacheService.getDocumentCache();
    const userCache = CacheService.getUserCache();

    if (scriptCache) scriptCache.removeAll(["rss-feed-contents", "session_cache"]);
    if (documentCache) documentCache.removeAll(["document_state"]);
    if (userCache) userCache.removeAll(["user_context"]);

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: "SEGURIDAD",
        SUBMODULO: "MANTENIMIENTO",
        ACCION: "LIMPIAR_CACHE",
        TIPO_REGISTRO: "SISTEMA",
        DESCRIPCION: "Caché general del ERP purgada con éxito por " + usuarioEjecutor,
        RESULTADO: "EXITOSO"
      });
    }

    return { EXITO: true, MENSAJE: "¡Caché del sistema purgada correctamente!" };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_PURGAR_CACHE_SISTEMA", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: "Fallo al purgar caché: " + error.message };
  }
}

function MNT_PURGAR_TABLA_HISTORIAL(nombreHoja, campoFecha, diasRetencion, tokenSesion) {
  try {
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) throw new Error("La hoja '" + nombreHoja + "' no existe.");

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { EXITO: true, MENSAJE: "La tabla ya está limpia.", eliminadas: 0 };
    }

    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
    const idxFecha = encabezados.indexOf(String(campoFecha).toUpperCase());
    if (idxFecha === -1) throw new Error("No se encontró columna: " + campoFecha);

    const rangoDatos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn());
    const registros = rangoDatos.getValues();
    
    const ahora = new Date();
    const milisegundosRetencion = Number(diasRetencion || 0) * 24 * 60 * 60 * 1000;
    const limiteFecha = new Date(ahora.getTime() - milisegundosRetencion);

    const registrosNuevos = [];
    let filasEliminadas = 0;

    registros.forEach(fila => {
      let conservar = true;
      if (diasRetencion === 0) {
        conservar = false;
      } else {
        const fechaRegistroVal = fila[idxFecha];
        if (fechaRegistroVal) {
          const fechaReg = new Date(fechaRegistroVal);
          if (!isNaN(fechaReg.getTime()) && fechaReg.getTime() < limiteFecha.getTime()) {
            conservar = false;
          }
        }
      }
      if (conservar) registrosNuevos.push(fila);
      else filasEliminadas++;
    });

    rangoDatos.clearContent();
    if (registrosNuevos.length > 0) {
      hoja.getRange(2, 1, registrosNuevos.length, encabezados.length).setValues(registrosNuevos);
    }

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: "SEGURIDAD",
        SUBMODULO: "MANTENIMIENTO",
        ACCION: "DEPURAR_TABLA",
        TIPO_REGISTRO: "SISTEMA",
        DESCRIPCION: "Purga realizada sobre la tabla " + nombreHoja + ". Filas purgadas: " + filasEliminadas + " por " + usuarioEjecutor,
        RESULTADO: "EXITOSO"
      });
    }

    try {
      if (hoja.getFilter()) hoja.getFilter().remove();
      hoja.getRange(1, 1, Math.max(2, hoja.getLastRow()), encabezados.length).createFilter();
    } catch (e) {}

    return { EXITO: true, MENSAJE: "¡Tabla " + nombreHoja + " depurada! " + filasEliminadas + " filas eliminadas.", eliminadas: filasEliminadas };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_PURGAR_TABLA_HISTORIAL", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: error.message };
  }
}

function MNT_PURGAR_AUDITORIA_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_AUDITORIA, "FECHA_HORA", diasRetencion, tokenSesion);
}

function MNT_PURGAR_SESIONES_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_SESIONES, "FECHA_INICIO", diasRetencion, tokenSesion);
}

function MNT_RESET_USUARIOS_SISTEMA(tokenSesion) {
  try {
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(SEG_CONFIG.HOJA_USUARIOS);
    if (!hoja) throw new Error("La hoja '" + SEG_CONFIG.HOJA_USUARIOS + "' no existe.");

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila >= 2) {
      hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).clearContent();
    }

    if (typeof SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS === "function") {
      SEG_INICIALIZAR_USUARIOS_PREDEFINIDOS();
    }

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: "SEGURIDAD",
        SUBMODULO: "MANTENIMIENTO",
        ACCION: "RESET_USUARIOS",
        TIPO_REGISTRO: "SISTEMA",
        DESCRIPCION: "Base de datos de usuarios reiniciada e inicializada por " + usuarioEjecutor,
        RESULTADO: "EXITOSO"
      });
    }

    return { EXITO: true, MENSAJE: "¡Base de datos de usuarios reiniciada exitosamente! Se ha re-creado la cuenta ADMIN por defecto (Admin123!)." };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_RESET_USUARIOS_SISTEMA", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: error.message };
  }
}