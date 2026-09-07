// (VERSIÓN 1.0 - V2 ERP - LIBRO 2)
/**************************************************************
* 18_ESTADOS_FINANCIEROS.gs
* RESPONSABILIDAD:
* - Calcular de forma modular y dinámica Balance General y PyG (Colombia) a partir de los datos operacionales de Sheets.
* - Administrar notas y observaciones personalizables de cierre financiero.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const FIN_CONFIG_CORE = {
  HOJA_NOTAS: "FIN_ESTADOS_NOTAS",
  PREFIJO_ID: "NOT",
  DIGITOS_ID: 6
};

/**
 * RPC: Obtener las notas y observaciones guardadas para un periodo específico
 */
function FIN_OBTENER_NOTAS_ESTADOS_WEB(periodo, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(FIN_CONFIG_CORE.HOJA_NOTAS);
    if (!hoja) {
      // Auto-inicializar la hoja de notas si no existe para evitar crash
      hoja = ss.insertSheet(FIN_CONFIG_CORE.HOJA_NOTAS);
      hoja.setTabColor("#6b7280");
      hoja.appendRow(["ID_NOTA", "SECCION", "PERIODO", "CONTENIDO", "OBSERVACIONES", "FECHA_ACTUALIZACION", "USUARIO"]);
    }
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay notas guardadas." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros
      .map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila))
      .filter(f => String(f.PERIODO).trim() === String(periodo).trim());
      
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Notas recuperadas con éxito para el período " + periodo
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al obtener notas financieras: " + error.toString() };
  }
}

/**
 * RPC: Guardar o actualizar una nota/observación contable para el periodo
 */
function FIN_GUARDAR_NOTA_ESTADO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.SECCION || !datos.PERIODO || !datos.CONTENIDO) {
      throw new Error("Datos de notas incompletos.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(FIN_CONFIG_CORE.HOJA_NOTAS);
    if (!hoja) {
      hoja = ss.insertSheet(FIN_CONFIG_CORE.HOJA_NOTAS);
      hoja.setTabColor("#6b7280");
      hoja.appendRow(["ID_NOTA", "SECCION", "PERIODO", "CONTENIDO", "OBSERVACIONES", "FECHA_ACTUALIZACION", "USUARIO"]);
    }
    
    const ultimaFila = hoja.getLastRow();
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    
    let filaIndex = -1;
    let registros = [];
    if (ultimaFila >= 2) {
      registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
      for (let i = 0; i < registros.length; i++) {
        if (String(registros[i][1]).trim().toUpperCase() === String(datos.SECCION).trim().toUpperCase() &&
            String(registros[i][2]).trim() === String(datos.PERIODO).trim()) {
          filaIndex = i + 2;
          break;
        }
      }
    }
    
    const ahora = new Date();
    
    if (filaIndex !== -1) {
      // Actualizar nota existente
      const colContenido = encabezados.indexOf("CONTENIDO") + 1;
      const colObs = encabezados.indexOf("OBSERVACIONES") + 1;
      const colFecha = encabezados.indexOf("FECHA_ACTUALIZACION") + 1;
      const colUsuario = encabezados.indexOf("USUARIO") + 1;
      
      hoja.getRange(filaIndex, colContenido).setValue(datos.CONTENIDO);
      hoja.getRange(filaIndex, colObs).setValue(datos.OBSERVACIONES || "");
      hoja.getRange(filaIndex, colFecha).setValue(ahora);
      hoja.getRange(filaIndex, colUsuario).setValue(usuarioEjecutor);
    } else {
      // Crear nueva nota
      const idNota = FIN_CONFIG_CORE.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(FIN_CONFIG_CORE.DIGITOS_ID, "0");
      hoja.appendRow([
        idNota,
        datos.SECCION.toUpperCase(),
        datos.PERIODO,
        datos.CONTENIDO,
        datos.OBSERVACIONES || "",
        ahora,
        usuarioEjecutor
      ]);
    }
    
    return {
      EXITO: true,
      MENSAJE: "¡Nota financiera y observaciones actualizadas exitosamente en Sheets!"
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("FIN_GUARDAR_NOTA_ESTADO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al guardar notas financieras: " + error.toString() };
  }
}

/**
 * RPC: Calcular dinámicamente Balance General y PyG a partir de datos operacionales
 */
function FIN_CALCULAR_ESTADOS_FINANCIEROS_WEB(periodo, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Inicializar totales de cálculo real
    let ingresosVentas = 0;
    let otrosIngresos = 0;
    let costosObra = 0;
    let gastosOperativos = 0;
    let cajaBancos = 0;
    let carteraClientes = 0;
    let cxpProveedores = 0;
    
    // 1. Calcular Ingresos de Ventas reales (VEN_CABECERA)
    const hojaVentas = ss.getSheetByName("VEN_CABECERA");
    if (hojaVentas && hojaVentas.getLastRow() > 1) {
      const vData = hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 15).getValues();
      vData.forEach(v => {
        if (v[12] !== "ANULADA") { // No anulada
          ingresosVentas += Number(v[8] || 0); // SUBTOTAL
        }
      });
    }
    
    // 2. Otros ingresos (ING_MOVIMIENTOS)
    const hojaIngresos = ss.getSheetByName("ING_MOVIMIENTOS");
    if (hojaIngresos && hojaIngresos.getLastRow() > 1) {
      const iData = hojaIngresos.getRange(2, 1, hojaIngresos.getLastRow() - 1, 12).getValues();
      iData.forEach(i => {
        if (i[10] === "PROCESADO" || i[10] === "PAGADO") {
          otrosIngresos += Number(i[6] || 0); // VALOR
        }
      });
    }
    
    // 3. Costos directos de Obras (COS_MOVIMIENTOS)
    const hojaCostos = ss.getSheetByName("COS_MOVIMIENTOS");
    if (hojaCostos && hojaCostos.getLastRow() > 1) {
      const cData = hojaCostos.getRange(2, 1, hojaCostos.getLastRow() - 1, 15).getValues();
      cData.forEach(c => {
        if (c[13] !== "ANULADO") {
          costosObra += Number(c[9] || 0); // VALOR
        }
      });
    }
    
    // 4. Gastos operativos y administrativos (GAS_MOVIMIENTOS)
    const hojaGastos = ss.getSheetByName("GAS_MOVIMIENTOS");
    if (hojaGastos && hojaGastos.getLastRow() > 1) {
      const gData = hojaGastos.getRange(2, 1, hojaGastos.getLastRow() - 1, 16).getValues();
      gData.forEach(g => {
        if (g[14] !== "ANULADO") {
          gastosOperativos += Number(g[7] || 0); // VALOR
        }
      });
    }
    
    // 5. Caja y Bancos (Saldos en TES_CUENTAS)
    const hojaTesoreria = ss.getSheetByName("TES_CUENTAS");
    if (hojaTesoreria && hojaTesoreria.getLastRow() > 1) {
      const tData = hojaTesoreria.getRange(2, 1, hojaTesoreria.getLastRow() - 1, 9).getValues();
      tData.forEach(t => {
        if (t[8] === "ACTIVO") {
          cajaBancos += Number(t[6] || 0); // SALDO_INICIAL / SALDO
        }
      });
    }
    
    // 6. Cartera / CxC (Saldos en CAR_CUENTAS)
    const hojaCartera = ss.getSheetByName("CAR_CUENTAS");
    if (hojaCartera && hojaCartera.getLastRow() > 1) {
      const carData = hojaCartera.getRange(2, 1, hojaCartera.getLastRow() - 1, 11).getValues();
      carData.forEach(car => {
        if (car[10] !== "ANULADA") {
          carteraClientes += Number(car[8] || 0); // SALDO
        }
      });
    }
    
    // 7. Cuentas por Pagar / CxP (Saldos en CXP_CUENTAS)
    const hojaCxp = ss.getSheetByName("CXP_CUENTAS");
    if (hojaCxp && hojaCxp.getLastRow() > 1) {
      const cxpData = hojaCxp.getRange(2, 1, hojaCxp.getLastRow() - 1, 11).getValues();
      cxpData.forEach(cxp => {
        if (cxp[10] !== "ANULADA") {
          cxpProveedores += Number(cxp[8] || 0); // SALDO
        }
      });
    }
    
    return {
      EXITO: true,
      DATOS: {
        INGRESOS_VENTAS: ingresosVentas,
        OTROS_INGRESOS: otrosIngresos,
        COSTOS_OBRA: costosObra,
        GASTOS_OPERATIVOS: gastosOperativos,
        CAJA_BANCOS: cajaBancos,
        CARTERA_CLIENTES: carteraClientes,
        CXP_PROVEEDORES: cxpProveedores,
        TOTAL_ACTIVOS: cajaBancos + carteraClientes,
        TOTAL_PASIVOS: cxpProveedores,
        PATRIMONIO_ESTIMADO: (cajaBancos + carteraClientes) - cxpProveedores
      },
      MENSAJE: "Cálculo en caliente de estados financieros completado."
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: "Error al calcular estados: " + error.toString() };
  }
}