// (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 14_GASTOS.gs
* RESPONSABILIDAD:
* - Registrar y clasificar los Gastos Operativos y Administrativos del ERP.
* - Afectar de manera automatizada las cuentas de egreso de Tesorería.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const GAS_CONFIG = {
  HOJA_MOVIMIENTOS: "GAS_MOVIMIENTOS",
  HOJA_CATEGORIAS: "GAS_CATEGORIAS",
  PREFIJO_ID: "GAS",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar todos los gastos registrados, uniendo con PROV_MAESTRO para el nombre comercial
 */
function GAS_LISTAR_GASTOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(GAS_CONFIG.HOJA_MOVIMIENTOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de gastos no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran gastos en la base de datos." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con PROV_MAESTRO
    const hojaProv = ss.getSheetByName("PROV_MAESTRO");
    if (hojaProv) {
      const provEnc = hojaProv.getRange(1, 1, 1, hojaProv.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const provRegs = hojaProv.getLastRow() > 1 ? hojaProv.getRange(2, 1, hojaProv.getLastRow() - 1, hojaProv.getLastColumn()).getValues() : [];
      const mapaProv = {};
      
      provRegs.forEach(row => {
        const p = SEG_CONVERTIR_FILA_OBJETO(provEnc, row);
        mapaProv[p.ID_PROVEEDOR] = p.RAZON_SOCIAL || p.NOMBRE_COMERCIAL || p.NOMBRE_CONTACTO;
      });
      
      lista.forEach(item => {
        item.RAZON_SOCIAL_PROVEEDOR = mapaProv[item.ID_PROVEEDOR] || "Tercero / Empleado / Sin Proveedor (" + item.ID_PROVEEDOR + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de egresos/gastos obtenida de forma exitosa."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar gastos: " + error.toString() };
  }
}

/**
 * RPC: Registrar un nuevo egreso de gastos administrativos o comerciales
 */
function GAS_REGISTRAR_GASTO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.VALOR || datos.VALOR <= 0 || !datos.ID_CUENTA) {
      throw new Error("Datos de gasto incompletos o cuenta de origen ausente.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = ss.getSheetByName(GAS_CONFIG.HOJA_MOVIMIENTOS);
    if (!hojaMov) throw new Error("Hoja GAS_MOVIMIENTOS no encontrada.");
    
    const idGasto = GAS_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(GAS_CONFIG.DIGITOS_ID, "0");
    const ahora = new Date();
    
    // Columnas: ID_GASTO, FECHA, TIPO_GASTO, CATEGORIA, ID_PROVEEDOR, DOCUMENTO_ORIGEN, ID_ORIGEN, VALOR, IVA, CENTRO_COSTO, CUENTA_CONTABLE, METODO_PAGO, ID_CUENTA, RESPONSABLE, ESTADO, OBSERVACION
    hojaMov.appendRow([
      idGasto,
      ahora,
      datos.TIPO_GASTO || "ADMINISTRATIVO",
      datos.CATEGORIA || "Otros Gastos",
      datos.ID_PROVEEDOR || "PROV-GENERAL",
      datos.DOCUMENTO_ORIGEN || "MANUAL",
      datos.ID_ORIGEN || idGasto,
      Number(datos.VALOR),
      Number(datos.IVA || 0),
      datos.CENTRO_COSTO || "ADMINISTRATIVO",
      datos.CUENTA_CONTABLE || "519505",
      datos.METODO_PAGO || "TRANSFERENCIA",
      datos.ID_CUENTA,
      usuarioEjecutor,
      "PROCESADO",
      datos.OBSERVACION || "Gasto general registrado manualmente"
    ]);
    
    // Afectar la cuenta en Tesorería
    try {
      const hojaCuentas = ss.getSheetByName("TES_CUENTAS");
      const hojaTesMov = ss.getSheetByName("TES_MOVIMIENTOS");
      if (hojaCuentas && hojaTesMov) {
        let nuevoSaldo = 0;
        const cuentasData = hojaCuentas.getRange(2, 1, Math.max(1, hojaCuentas.getLastRow() - 1), 10).getValues();
        const cIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA).trim());
        if (cIndex !== -1) {
          const saldoAnterior = Number(cuentasData[cIndex][6] || 0);
          nuevoSaldo = saldoAnterior - Number(datos.VALOR);
          hojaCuentas.getRange(cIndex + 2, 7).setValue(nuevoSaldo); // Actualizar saldo de cuenta
        }
        
        const idTes = "MOV-" + String(Math.max(1, hojaTesMov.getLastRow())).padStart(6, "0");
        hojaTesMov.appendRow([
          idTes,
          ahora,
          "EGRESO",
          datos.ID_CUENTA,
          "GASTOS",
          idGasto,
          0,
          Number(datos.VALOR),
          nuevoSaldo,
          datos.METODO_PAGO || "TRANSFERENCIA",
          usuarioEjecutor,
          datos.OBSERVACION || "Gasto de administración " + idGasto
        ]);
      }
    } catch (errTes) {
      console.warn("No se pudo afectar el libro de tesorería: " + errTes.toString());
    }
    
    // Registrar Auditoría
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "TESORERIA",
      SUBMODULO: "GASTOS",
      ACCION: "CREAR",
      TIPO_REGISTRO: "GAS_MOVIMIENTOS",
      ID_REGISTRO: idGasto,
      DESCRIPCION: "Gasto operativo registrado por valor de: " + datos.VALOR + " COP cargado a cuenta: " + datos.ID_CUENTA,
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      ID_GASTO: idGasto,
      MENSAJE: "¡Gasto general registrado exitosamente! Consecutivo: " + idGasto
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("GAS_REGISTRAR_GASTO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al registrar gasto: " + error.toString() };
  }
}


/**
 * Función de compatibilidad y enrutamiento directo de Gastos para Suite de Pruebas E2E
 * (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
 */
function GAS_REGISTRAR_GASTO(datos, tokenSesion) {
  return GAS_REGISTRAR_GASTO_WEB(datos, tokenSesion);
}
