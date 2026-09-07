// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 12_INGRESOS.gs
* RESPONSABILIDAD:
* - Registrar y clasificar los movimientos de Ingresos del ERP.
* - Sincronizar el recaudo asíncronamente con Tesorería (TES_).
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const ING_CONFIG = {
  HOJA_MOVIMIENTOS: "ING_MOVIMIENTOS",
  HOJA_RECAUDOS: "ING_RECAUDOS",
  PREFIJO_ID: "ING",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar todos los movimientos de ingresos, uniendo con CLI_MAESTRO para la razón social
 */
function ING_LISTAR_INGRESOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(ING_CONFIG.HOJA_MOVIMIENTOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de ingresos no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran movimientos de ingresos." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con CLI_MAESTRO
    const hojaClientes = ss.getSheetByName("CLI_MAESTRO");
    if (hojaClientes) {
      const cliEnc = hojaClientes.getRange(1, 1, 1, hojaClientes.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const cliRegs = hojaClientes.getLastRow() > 1 ? hojaClientes.getRange(2, 1, hojaClientes.getLastRow() - 1, hojaClientes.getLastColumn()).getValues() : [];
      const mapaClientes = {};
      
      cliRegs.forEach(row => {
        const c = SEG_CONVERTIR_FILA_OBJETO(cliEnc, row);
        mapaClientes[c.ID_CLIENTE] = c.RAZON_SOCIAL || c.NOMBRE_COMERCIAL || (c.PRIMER_NOMBRE + " " + c.PRIMER_APELLIDO);
      });
      
      lista.forEach(item => {
        item.RAZON_SOCIAL_CLIENTE = mapaClientes[item.ID_CLIENTE] || "Otro / Sin Cliente (" + item.ID_CLIENTE + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de ingresos obtenida exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar ingresos: " + error.toString() };
  }
}

/**
 * RPC: Registrar un nuevo movimiento de ingreso manual (No asociado a ventas a crédito)
 */
function ING_REGISTRAR_INGRESO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.VALOR || datos.VALOR <= 0 || !datos.ID_CUENTA_DESTINO) {
      throw new Error("Datos de ingreso incompletos.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = ss.getSheetByName(ING_CONFIG.HOJA_MOVIMIENTOS);
    if (!hojaMov) throw new Error("Hoja ING_MOVIMIENTOS no encontrada.");
    
    const idIngreso = ING_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(ING_CONFIG.DIGITOS_ID, "0");
    const ahora = new Date();
    
    // Guardar fila de ingresos: ID_INGRESO, FECHA, TIPO_INGRESO, ORIGEN, ID_ORIGEN, ID_CLIENTE, VALOR, METODO_PAGO, ID_CUENTA_DESTINO, RESPONSABLE, ESTADO, OBSERVACION
    hojaMov.appendRow([
      idIngreso,
      ahora,
      datos.TIPO_INGRESO || "OTROS_INGRESOS",
      datos.ORIGEN || "MANUAL",
      datos.ID_ORIGEN || idIngreso,
      datos.ID_CLIENTE || "PUBLICO_GENERAL",
      Number(datos.VALOR),
      datos.METODO_PAGO || "TRANSFERENCIA",
      datos.ID_CUENTA_DESTINO,
      usuarioEjecutor,
      "PROCESADO",
      datos.OBSERVACION || "Ingreso registrado manualmente"
    ]);
    
    // Afectar la cuenta en Tesorería
    try {
      const hojaCuentas = ss.getSheetByName("TES_CUENTAS");
      const hojaTesMov = ss.getSheetByName("TES_MOVIMIENTOS");
      if (hojaCuentas && hojaTesMov) {
        let nuevoSaldo = 0;
        const cuentasData = hojaCuentas.getRange(2, 1, Math.max(1, hojaCuentas.getLastRow() - 1), 10).getValues();
        const cIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA_DESTINO).trim());
        if (cIndex !== -1) {
          const saldoAnterior = Number(cuentasData[cIndex][6] || 0);
          nuevoSaldo = saldoAnterior + Number(datos.VALOR);
          hojaCuentas.getRange(cIndex + 2, 7).setValue(nuevoSaldo); // Actualizar saldo de cuenta
        }
        
        const idTes = "MOV-" + String(Math.max(1, hojaTesMov.getLastRow())).padStart(6, "0");
        hojaTesMov.appendRow([
          idTes,
          ahora,
          "INGRESO",
          datos.ID_CUENTA_DESTINO,
          "INGRESOS",
          idIngreso,
          Number(datos.VALOR),
          0,
          nuevoSaldo,
          datos.METODO_PAGO || "TRANSFERENCIA",
          usuarioEjecutor,
          datos.OBSERVACION || "Ingreso manual " + idIngreso
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
      SUBMODULO: "INGRESOS",
      ACCION: "CREAR",
      TIPO_REGISTRO: "ING_MOVIMIENTOS",
      ID_REGISTRO: idIngreso,
      DESCRIPCION: "Ingreso manual registrado por valor de: " + datos.VALOR + " COP en cuenta: " + datos.ID_CUENTA_DESTINO,
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      ID_INGRESO: idIngreso,
      MENSAJE: "¡Ingreso registrado exitosamente! Consecutivo: " + idIngreso
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("ING_REGISTRAR_INGRESO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al registrar ingreso: " + error.toString() };
  }
}