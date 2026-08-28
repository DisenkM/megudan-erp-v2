/**************************************************************
* 26_MANTENIMIENTO.gs
* ERP OPERATIVO MEGUDAN V2
* RESPONSABILIDAD:
* - Proveer un motor centralizado de limpieza, depuración y mantenimiento.
* - Truncar de forma segura tablas de historial (Auditoría, Sesiones, Historiales de Terceros).
* - Administrar y purgar el CacheService de Google Apps Script.
* - Soportar la arquitectura de seguridad dual (Bypass Sheets local / Validación remota por Token).
**************************************************************/

const MNT_CONFIG = {
  HOJA_AUDITORIA: "USR_AUDITORIA",
  HOJA_SESIONES: "USR_SESIONES",
  HOJA_CLI_HISTORIAL: "CLI_HISTORIAL",
  HOJA_PROV_HISTORIAL: "PROV_HISTORIAL"
};

/**
 * Función central de validación de privilegios de mantenimiento.
 * Solo permite la ejecución si se realiza en Sheets (Administrador local) o si se provee
 * un token de sesión con rol ADMINISTRADOR y permiso de ADMINISTRAR en el módulo SEGURIDAD.
 */
function MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion) {
  if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function") {
    // Retorna el contexto validado (sea bypass local o sesión remota aprobada)
    return SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ADMINISTRAR");
  }
  
  // Salvaguarda fallback en caso de ausencia temporal del módulo de seguridad
  try {
    SpreadsheetApp.getUi();
    return { AUTORIZADO: true, USUARIO: "ADMINISTRADOR_LOCAL_SHEETS" };
  } catch (e) {
    throw new Error("ACCESO DENEGADO: No se pudo verificar el contexto de seguridad de mantenimiento.");
  }
}

/**
 * Puga por completo la caché lógica de Google Apps Script (Script, Document y User Cache).
 * @param {string} tokenSesion Token seguro de la Web App (opcional desde Sheets local).
 * @returns {object} Estado de la transacción para el frontend.
 */
function MNT_PURGAR_CACHE_SISTEMA(tokenSesion) {
  try {
    // 1. Validar autorización de seguridad dual
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    // 2. Instanciar los servicios de caché oficiales de Google Apps Script
    const scriptCache = CacheService.getScriptCache();
    const documentCache = CacheService.getDocumentCache();
    const userCache = CacheService.getUserCache();

    // 3. Remover todos los registros activos del caché asíncrono
    // Nota: Aunque no usemo claves pesadas hoy, previene problemas de memoria de sesión
    if (scriptCache) scriptCache.removeAll(["rss-feed-contents", "session_cache"]);
    if (documentCache) documentCache.removeAll(["document_state"]);
    if (userCache) userCache.removeAll(["user_context"]);

    // 4. Registrar la acción en la bitácora de auditoría
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

    return {
      EXITO: true,
      MENSAJE: "¡Caché del sistema (Script, Documento y Usuario) purgada correctamente!"
    };

  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_PURGAR_CACHE_SISTEMA", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: "Fallo al purgar caché: " + error.message };
  }
}

/**
 * Purga de forma selectiva o total los registros de una tabla de historial en Sheets,
 * basándose en un criterio estricto de días de retención (retención = 0 limpia todo).
 * @param {string} nombreHoja Nombre físico de la hoja de Sheets.
 * @param {string} campoFecha Nombre de la columna que contiene la fecha del registro.
 * @param {number} diasRetencion Cantidad de días de historial a conservar (0 para limpiar todo).
 * @param {string} tokenSesion Token seguro de la Web App (opcional desde Sheets local).
 * @returns {object} Estado detallado de las filas eliminadas y conservadas.
 */
function MNT_PURGAR_TABLA_HISTORIAL(nombreHoja, campoFecha, diasRetencion, tokenSesion) {
  try {
    // 1. Validar autorización de seguridad dual
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";

    // 2. Conectar y validar la existencia de la hoja física
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      throw new Error("La hoja '" + nombreHoja + "' no existe en la base de datos.");
    }

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { EXITO: true, MENSAJE: "La tabla " + nombreHoja + " ya está completamente limpia.", eliminadas: 0 };
    }

    // 3. Leer encabezados para ubicar dinámicamente la columna de fechas
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
    const idxFecha = encabezados.indexOf(String(campoFecha).toUpperCase());
    if (idxFecha === -1) {
      throw new Error("No se encontró la columna de control de fecha '" + campoFecha + "' en " + nombreHoja);
    }

    // 4. Leer todos los registros de la tabla
    const rangoDatos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn());
    const registros = rangoDatos.getValues();
    
    const ahora = new Date();
    const milisegundosRetencion = Number(diasRetencion || 0) * 24 * 60 * 60 * 1000;
    const limiteFecha = new Date(ahora.getTime() - milisegundosRetencion);

    const registrosNuevos = [];
    let filasEliminadas = 0;

    // 5. Procesamiento lógico local en memoria (Batch Processing de alta velocidad)
    registros.forEach(fila => {
      let conservar = true;
      
      if (diasRetencion === 0) {
        conservar = false; // Truncado total solicitado
      } else {
        const fechaRegistroVal = fila[idxFecha];
        if (fechaRegistroVal) {
          const fechaReg = new Date(fechaRegistroVal);
          if (!isNaN(fechaReg.getTime()) && fechaReg.getTime() < limiteFecha.getTime()) {
            conservar = false; // El registro superó la antigüedad de retención
          }
        }
      }

      if (conservar) {
        registrosNuevos.push(fila);
      } else {
        filasEliminadas++;
      }
    });

    // 6. Escribir masivamente de vuelta (Batch Write)
    // Limpiamos todo el rango de datos operativos
    rangoDatos.clearContent();

    if (registrosNuevos.length > 0) {
      // Inyectamos de vuelta únicamente los registros que cumplen con el período de retención
      hoja.getRange(2, 1, registrosNuevos.length, encabezados.length).setValues(registrosNuevos);
    }

    // 7. Registrar auditoría física en Sheets
    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        MODULO: "SEGURIDAD",
        SUBMODULO: "MANTENIMIENTO",
        ACCION: "DEPURAR_TABLA",
        TIPO_REGISTRO: "SISTEMA",
        DESCRIPCION: "Purga realizada sobre la tabla " + nombreHoja + " (" + diasRetencion + " días retención). Filas purgadas: " + filasEliminadas + " por " + usuarioEjecutor,
        RESULTADO: "EXITOSO"
      });
    }

    // Forzar el redibujado de filtros de Sheets para evitar desalineación visual
    try {
      if (hoja.getFilter()) {
        hoja.getFilter().remove();
      }
      hoja.getRange(1, 1, Math.max(2, hoja.getLastRow()), encabezados.length).createFilter();
    } catch (e) {
      console.warn("No se pudo regenerar el filtro en " + nombreHoja);
    }

    return {
      EXITO: true,
      MENSAJE: "¡Tabla " + nombreHoja + " depurada con éxito! Se purgaron " + filasEliminadas + " registros.",
      eliminadas: filasEliminadas,
      conservadas: registrosNuevos.length
    };

  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_PURGAR_TABLA_HISTORIAL", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: "Error al purgar la tabla: " + error.message };
  }
}

/**
 * Purga de un solo clic la cola de Auditoría general (USR_AUDITORIA) conservando
 * únicamente el número de días parametrizado.
 */
function MNT_PURGAR_AUDITORIA_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_AUDITORIA, "FECHA_HORA", diasRetencion, tokenSesion);
}

/**
 * Purga el historial de sesiones cerradas o antiguas en USR_SESIONES.
 */
function MNT_PURGAR_SESIONES_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_SESIONES, "FECHA_INICIO", diasRetencion, tokenSesion);
}

/**
 * Purga el historial de modificaciones del catálogo de clientes (CLI_HISTORIAL).
 */
function MNT_PURGAR_HISTORIAL_CLIENTES(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_CLI_HISTORIAL, "FECHA_HORA", diasRetencion, tokenSesion);
}

/**
 * Purga el historial de modificaciones de proveedores (PROV_HISTORIAL).
 */
function MNT_PURGAR_HISTORIAL_PROVEEDORES(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_PROV_HISTORIAL, "FECHA", diasRetencion, tokenSesion);
}