/**************************************************************
* 26_MANTENIMIENTO.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Motor centralizado de limpieza, depuración y mantenimiento de base de datos.
* - Truncar de forma segura tablas de historial (Auditoría, Sesiones) preservando retenciones.
* - Administrar y purgar el CacheService de Google Apps Script.
* - Eliminar de forma atómica y recursiva los registros de prueba generados en las pruebas E2E.
**************************************************************/

const MNT_CONFIG = {
  HOJA_AUDITORIA: "USR_AUDITORIA",
  HOJA_SESIONES: "USR_SESIONES",
  HOJA_CLI_HISTORIAL: "CLI_HISTORIAL",
  HOJA_PROV_HISTORIAL: "PROV_HISTORIAL"
};

/**
 * Valida el acceso del usuario para tareas de mantenimiento administrativo
 */
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

/**
 * RPC: Limpia de forma segura y completa todos los registros de pruebas generados en la base de datos de Sheets.
 * Identifica y purga clientes, proveedores, compras, ventas, movimientos de inventario, etc., creados durante el testeo.
 * Borra de abajo hacia arriba para evitar el desplazamiento de índices en Google Sheets.
 */
function MNT_LIMPIAR_DATOS_PRUEBA_WEB(tokenSesion) {
  try {
    const auth = MNT_VERIFICAR_ACCESO_MANTENIMIENTO(tokenSesion);
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let totalEliminados = 0;
    const logDetallado = [];
    
    // Definición de hojas y sus criterios de limpieza específicos de pruebas
    const reglasLimpieza = [
      {
        hoja: "USR_USUARIOS",
        columnaId: "USUARIO",
        evaluar: function(valor) { return String(valor).toLowerCase().startsWith("test_"); }
      },
      {
        hoja: "USR_AUDITORIA",
        columnaId: "USUARIO",
        evaluar: function(valor, fila) {
          return String(valor).toLowerCase().startsWith("test_") || 
                 String(fila["DESCRIPCION"]).toLowerCase().includes("suite de pruebas") ||
                 String(fila["DESCRIPCION"]).toLowerCase().includes("diagnóstico completo");
        }
      },
      {
        hoja: "CLI_MAESTRO",
        columnaId: "RAZON_SOCIAL",
        evaluar: function(valor) { return String(valor).toUpperCase() === "CLIENTE TEST INTEGRACION SAS"; }
      },
      {
        hoja: "CLI_HISTORIAL",
        columnaId: "OBSERVACIONES",
        evaluar: function(valor) { 
          return String(valor).toLowerCase().includes("test_") || 
                 String(valor).toLowerCase().includes("suite de pruebas") || 
                 String(valor).toLowerCase().includes("cliente creado por test_") ||
                 String(valor).toLowerCase().includes("cliente actualizado por test_"); 
        }
      },
      {
        hoja: "PROV_MAESTRO",
        columnaId: "RAZON_SOCIAL",
        evaluar: function(valor) { return String(valor).toUpperCase() === "PROVEEDOR TEST CONSTRUCCION SAS"; }
      },
      {
        hoja: "PROV_HISTORIAL",
        columnaId: "OBSERVACION",
        evaluar: function(valor) { 
          return String(valor).toLowerCase().includes("test_") || 
                 String(valor).toLowerCase().includes("proveedor creado por test_"); 
        }
      },
      {
        hoja: "PROD_MAESTRO",
        columnaId: "DESCRIPCION",
        evaluar: function(valor) { return String(valor).toUpperCase() === "GUADUA DE PRUEBA INMUNIZADA 6M"; }
      },
      {
        hoja: "COM_CABECERA",
        columnaId: "NUM_DOCUMENTO",
        evaluar: function(valor) { return String(valor).toUpperCase() === "FAC-COMPRA-789"; }
      },
      {
        hoja: "COM_DETALLE",
        columnaId: "DESCRIPCION",
        evaluar: function(valor) { return String(valor).toUpperCase() === "GUADUA DE PRUEBA INMUNIZADA 6M"; }
      },
      {
        hoja: "VEN_CABECERA",
        columnaId: "NUM_DOCUMENTO",
        evaluar: function(valor) { return String(valor).toUpperCase() === "FAC-VENTA-101"; }
      },
      {
        hoja: "VEN_DETALLE",
        columnaId: "DESCRIPCION",
        evaluar: function(valor) { return String(valor).toUpperCase() === "GUADUA DE PRUEBA INMUNIZADA 6M"; }
      },
      {
        hoja: "INV_MOVIMIENTOS",
        columnaId: "OBSERVACION",
        evaluar: function(valor) {
          return String(valor).toLowerCase().includes("salida automatizada por facturación de venta") ||
                 String(valor).toLowerCase().includes("entrada automática por facturación de compra") ||
                 String(valor).toLowerCase().includes("guadua de prueba");
        }
      },
      {
        hoja: "INV_SALDOS",
        columnaId: "ID_PRODUCTO",
        evaluar: function(valor) {
          return String(valor).toUpperCase().startsWith("PRD-") && !["PRD-000010"].includes(String(valor).toUpperCase()); 
        }
      },
      {
        hoja: "CAR_CUENTAS",
        columnaId: "DOCUMENTO",
        evaluar: function(valor) { return String(valor).toUpperCase().startsWith("VEN-"); } 
      },
      {
        hoja: "CAR_RECAUDOS",
        columnaId: "OBSERVACION",
        evaluar: function(valor) { return String(valor).toLowerCase().includes("recaudo sobre cartera"); }
      },
      {
        hoja: "CXP_CUENTAS",
        columnaId: "DOCUMENTO",
        evaluar: function(valor) { return String(valor).toUpperCase().startsWith("COM-"); }
      },
      {
        hoja: "CXP_PAGOS",
        columnaId: "OBSERVACION",
        evaluar: function(valor) { return String(valor).toLowerCase().includes("abono a obligación"); }
      },
      {
        hoja: "GAS_MOVIMIENTOS",
        columnaId: "OBSERVACION",
        evaluar: function(valor) { return String(valor).toLowerCase().includes("pago de energía eléctrica"); }
      },
      {
        hoja: "TES_MOVIMIENTOS",
        columnaId: "OBSERVACION",
        evaluar: function(valor) {
          return String(valor).toLowerCase().includes("gasto de administración gas-") ||
                 String(valor).toLowerCase().includes("recaudo de cartera") ||
                 String(valor).toLowerCase().includes("pago de cxp");
        }
      }
    ];
    
    // Procesar cada regla de limpieza de forma iterativa y segura
    reglasLimpieza.forEach(function(regla) {
      const sheet = ss.getSheetByName(regla.hoja);
      if (!sheet) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      const colHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase());
      const colIdx = colHeaders.indexOf(regla.columnaId.toUpperCase());
      if (colIdx === -1) return;
      
      const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      let eliminadosEnHoja = 0;
      
      // Recorrer de abajo hacia arriba para borrar de forma segura
      for (let i = values.length - 1; i >= 0; i--) {
        const filaFisica = i + 2;
        const filaObj = {};
        colHeaders.forEach((h, index) => {
          filaObj[h] = values[i][index];
        });
        
        const valorEvaluado = values[i][colIdx];
        if (regla.evaluar(valorEvaluado, filaObj)) {
          sheet.deleteRow(filaFisica);
          eliminadosEnHoja++;
        }
      }
      
      if (eliminadosEnHoja > 0) {
        totalEliminados += eliminadosEnHoja;
        logDetallado.push(`Pestaña ${regla.hoja}: ${eliminadosEnHoja} registros eliminados.`);
        
        // Re-inicializar filtros para evitar roturas visuales
        try {
          if (sheet.getFilter()) sheet.getFilter().remove();
          sheet.getRange(1, 1, Math.max(2, sheet.getLastRow()), colHeaders.length).createFilter();
        } catch (e) {
          // Silenciar errores de filtro
        }
      }
    });
    
    // Registrar auditoría de mantenimiento
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "MANTENIMIENTO",
      ACCION: "DEPURAR_PRUEBAS",
      TIPO_REGISTRO: "SISTEMA",
      DESCRIPCION: `Limpiador de pruebas ejecutado con éxito por ${usuarioEjecutor}. Registros eliminados: ${totalEliminados}. Detalles: ${logDetallado.join(" | ")}`,
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      ELIMINADOS: totalEliminados,
      DETALLE: logDetallado,
      MENSAJE: `Se limpiaron con éxito ${totalEliminados} registros de pruebas en la base de datos de Sheets.`
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_LIMPIAR_DATOS_PRUEBA_WEB", "MANTENIMIENTO", error);
    }
    return { EXITO: false, ELIMINADOS: 0, MENSAJE: "Fallo crítico en el limpiador de pruebas: " + error.message };
  }
}

/**
 * Wrapper de compatibilidad para la suite de pruebas y ejecuciones locales
 */
function MNT_LIMPIAR_DATOS_PRUEBA(tokenSesion) {
  return MNT_LIMPIAR_DATOS_PRUEBA_WEB(tokenSesion);
}

/**
 * RPC: Purga la caché de sesión global en memoria del ERP
 */
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
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "MANTENIMIENTO",
      ACCION: "LIMPIAR_CACHE",
      TIPO_REGISTRO: "SISTEMA",
      DESCRIPCION: "Caché general del ERP purgada con éxito por " + usuarioEjecutor,
      RESULTADO: "EXITOSO"
    });
    
    return { EXITO: true, MENSAJE: "¡Caché del sistema purgada correctamente!" };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_PURGAR_CACHE_SISTEMA", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: "Fallo al purgar caché: " + error.message };
  }
}

/**
 * RPC: Depuración genérica de tablas de historial y registros
 */
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
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "MANTENIMIENTO",
      ACCION: "DEPURAR_TABLA",
      TIPO_REGISTRO: "SISTEMA",
      DESCRIPCION: "Purga realizada sobre la tabla " + nombreHoja + ". Filas purgadas: " + filasEliminadas + " por " + usuarioEjecutor,
      RESULTADO: "EXITOSO"
    });
    
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

/**
 * RPC: Depurar registros antiguos en USR_AUDITORIA
 */
function MNT_PURGAR_AUDITORIA_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_AUDITORIA, "FECHA_HORA", diasRetencion, tokenSesion);
}

/**
 * RPC: Depurar registros antiguos en USR_SESIONES
 */
function MNT_PURGAR_SESIONES_SISTEMA(diasRetencion, tokenSesion) {
  return MNT_PURGAR_TABLA_HISTORIAL(MNT_CONFIG.HOJA_SESIONES, "FECHA_INICIO", diasRetencion, tokenSesion);
}

/**
 * RPC: Re-inicializa la tabla USR_USUARIOS preservando únicamente el ADMIN maestro
 */
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
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "MANTENIMIENTO",
      ACCION: "RESET_USUARIOS",
      TIPO_REGISTRO: "SISTEMA",
      DESCRIPCION: "Base de datos de usuarios reiniciada e inicializada por " + usuarioEjecutor,
      RESULTADO: "EXITOSO"
    });
    
    return { EXITO: true, MENSAJE: "¡Base de datos de usuarios reiniciada exitosamente! Se ha re-creado la cuenta ADMIN por defecto (Admin123!)." };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("MNT_RESET_USUARIOS_SISTEMA", "MANTENIMIENTO", error);
    }
    return { EXITO: false, MENSAJE: error.message };
  }
}