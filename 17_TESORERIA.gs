// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 17_TESORERIA.gs
* RESPONSABILIDAD:
* - Administrar cuentas bancarias, caja física y conciliación.
* - Registrar movimientos de entrada y salida actualizando saldos de bancos.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const TES_CONFIG_CORE = {
  HOJA_CUENTAS: "TES_CUENTAS",
  HOJA_MOVIMIENTOS: "TES_MOVIMIENTOS",
  HOJA_CONCILIACION: "TES_CONCILIACION",
  PREFIJO_ID: "MOV",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar las cuentas de caja y bancos registradas
 */
function TES_LISTAR_CUENTAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(TES_CONFIG_CORE.HOJA_CUENTAS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de cuentas de tesorería no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran cuentas de bancos o cajas." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de cuentas financieras obtenida exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar cuentas: " + error.toString() };
  }
}

/**
 * RPC: Listar todos los movimientos de caja y bancos (Kardex financiero unificado)
 */
function TES_LISTAR_MOVIMIENTOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(TES_CONFIG_CORE.HOJA_MOVIMIENTOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de movimientos de tesorería no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran movimientos bancarios." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con TES_CUENTAS para jalar el nombre legible del banco
    const hojaCta = ss.getSheetByName(TES_CONFIG_CORE.HOJA_CUENTAS);
    if (hojaCta) {
      const ctaEnc = hojaCta.getRange(1, 1, 1, hojaCta.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const ctaRegs = hojaCta.getLastRow() > 1 ? hojaCta.getRange(2, 1, hojaCta.getLastRow() - 1, hojaCta.getLastColumn()).getValues() : [];
      const mapaCta = {};
      
      ctaRegs.forEach(row => {
        const c = SEG_CONVERTIR_FILA_OBJETO(ctaEnc, row);
        mapaCta[c.ID_CUENTA] = c.NOMBRE_CUENTA + " (" + c.ENTIDAD + ")";
      });
      
      lista.forEach(item => {
        item.NOMBRE_CUENTA_VINCULADA = mapaCta[item.ID_CUENTA] || "Cuenta Desconocida (" + item.ID_CUENTA + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Kardex de tesorería obtenido exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar movimientos de tesorería: " + error.toString() };
  }
}

/**
 * RPC: Registrar transferencias o traslados internos de efectivo entre tus propias cuentas del ERP
 */
function TES_REGISTRAR_MOVIMIENTO_INTERNO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.ID_CUENTA_ORIGEN || !datos.ID_CUENTA_DESTINO || !datos.VALOR || datos.VALOR <= 0) {
      throw new Error("Datos de traslado interno de fondos incompletos.");
    }
    
    if (String(datos.ID_CUENTA_ORIGEN).trim() === String(datos.ID_CUENTA_DESTINO).trim()) {
      throw new Error("La cuenta de origen y destino no pueden ser iguales.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCuentas = ss.getSheetByName(TES_CONFIG_CORE.HOJA_CUENTAS);
    const hojaMov = ss.getSheetByName(TES_CONFIG_CORE.HOJA_MOVIMIENTOS);
    if (!hojaCuentas || !hojaMov) throw new Error("Hojas de tesorería no encontradas.");
    
    const cuentasData = hojaCuentas.getRange(2, 1, Math.max(1, hojaCuentas.getLastRow() - 1), 10).getValues();
    const oIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA_ORIGEN).trim());
    const dIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA_DESTINO).trim());
    
    if (oIndex === -1) throw new Error("Cuenta de origen no encontrada.");
    if (dIndex === -1) throw new Error("Cuenta de destino no encontrada.");
    
    const saldoOrigen = Number(cuentasData[oIndex][6] || 0);
    const saldoDestino = Number(cuentasData[dIndex][6] || 0);
    const valorTraslado = Number(datos.VALOR);
    
    if (valorTraslado > saldoOrigen) {
      throw new Error("Saldo insuficiente en cuenta origen. Disponible: $" + saldoOrigen.toLocaleString());
    }
    
    const nuevoSaldoOrigen = saldoOrigen - valorTraslado;
    const nuevoSaldoDestino = saldoDestino + valorTraslado;
    
    // Escribir nuevos saldos de vuelta en Sheets
    hojaCuentas.getRange(oIndex + 2, 7).setValue(nuevoSaldoOrigen);
    hojaCuentas.getRange(dIndex + 2, 7).setValue(nuevoSaldoDestino);
    
    const ahora = new Date();
    
    // Registrar salida de origen
    const idMovO = "MOV-TRO-" + ahora.getTime();
    hojaMov.appendRow([
      idMovO, ahora, "EGRESO", datos.ID_CUENTA_ORIGEN, "TRASLADO", idMovO, 0, valorTraslado, nuevoSaldoOrigen, "TRANSFERENCIA", usuarioEjecutor, "Traslado de fondos hacia cuenta: " + datos.ID_CUENTA_DESTINO
    ]);
    
    // Registrar entrada de destino
    const idMovD = "MOV-TRD-" + ahora.getTime();
    hojaMov.appendRow([
      idMovD, ahora, "INGRESO", datos.ID_CUENTA_DESTINO, "TRASLADO", idMovD, valorTraslado, 0, nuevoSaldoDestino, "TRANSFERENCIA", usuarioEjecutor, "Recepción de fondos desde cuenta: " + datos.ID_CUENTA_ORIGEN
    ]);
    
    // Registrar auditoría
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "TESORERIA",
      SUBMODULO: "BANCARIO",
      ACCION: "CREAR",
      TIPO_REGISTRO: "TES_MOVIMIENTOS",
      ID_REGISTRO: idMovO,
      DESCRIPCION: "Traslado interno de fondos. Valor: " + valorTraslado + " COP. Origen: " + datos.ID_CUENTA_ORIGEN + " -> Destino: " + datos.ID_CUENTA_DESTINO,
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      MENSAJE: "¡Traslado de fondos unificado registrado correctamente!"
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("TES_REGISTRAR_MOVIMIENTO_INTERNO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al registrar traslado: " + error.toString() };
  }
}
