// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 15_CARTERA.gs
* RESPONSABILIDAD:
* - Administrar las Cuentas por Cobrar (CxC / Cartera) asociadas a ventas a crédito.
* - Controlar saldos, abonos y recabar obligaciones de clientes en Sheets.
* - Afectar de forma automatizada las cuentas de tesorería ante recaudos.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const CAR_CONFIG = {
  HOJA_CUENTAS: "CAR_CUENTAS",
  HOJA_RECAUDOS: "CAR_RECAUDOS",
  PREFIJO_ID: "CAR",
  PREFIJO_RECAUDO: "REC",
  DIGITOS_ID: 6
};

/**
 * Crea una nueva cuenta por cobrar (CxC) a partir de una venta a crédito
 */
function CAR_CREAR_CUENTA_COBRAR(idVenta, idCliente, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CAR_CONFIG.HOJA_CUENTAS);
  if (!hoja) return;
  
  const idCartera = CAR_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CAR_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();
  
  // Plazo general de 30 días de crédito para el cobro
  const vencimiento = new Date(ahora.getTime() + (30 * 24 * 60 * 60 * 1000));
  
  // Columnas: ID_CARTERA, ID_CLIENTE, ID_VENTA, DOCUMENTO, FECHA_EMISION, FECHA_VENCIMIENTO, VALOR_DOCUMENTO, ABONOS, SALDO, DIAS_VENCIDOS, ESTADO
  hoja.appendRow([
    idCartera, 
    idCliente, 
    idVenta, 
    idVenta, // Documento de soporte (Factura o consecutivo)
    ahora, 
    vencimiento, 
    total, 
    0, // Abonos iniciales
    total, // Saldo inicial
    0, // Días vencidos
    "PENDIENTE"
  ]);
}

/**
 * RPC: Listar todas las cuentas de cartera activas, unificando con CLI_MAESTRO para la razón social
 */
function CAR_LISTAR_CUENTAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER"); // Permiso de finanzas / tesorería
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CAR_CONFIG.HOJA_CUENTAS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de cartera no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay cuentas por cobrar registradas." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con CLI_MAESTRO para el nombre real o razón social del cliente
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
        item.RAZON_SOCIAL_CLIENTE = mapaClientes[item.ID_CLIENTE] || "Cliente Desconocido (" + item.ID_CLIENTE + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Cuentas de cartera obtenidas de forma exitosa."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar cartera: " + error.toString() };
  }
}

/**
 * RPC: Registrar un abono o recaudo recibido de un cliente para amortizar su deuda (CxC)
 */
function CAR_REGISTRAR_RECAUDO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.ID_CARTERA || !datos.VALOR || datos.VALOR <= 0) {
      throw new Error("Datos de recaudo o valor no válidos.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCuentas = ss.getSheetByName(CAR_CONFIG.HOJA_CUENTAS);
    if (!hojaCuentas) throw new Error("Hoja CAR_CUENTAS no encontrada.");
    
    const ultimaFila = hojaCuentas.getLastRow();
    if (ultimaFila < 2) throw new Error("No hay cuentas de cartera registradas.");
    
    const encabezados = hojaCuentas.getRange(1, 1, 1, hojaCuentas.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registrosRange = hojaCuentas.getRange(2, 1, ultimaFila - 1, hojaCuentas.getLastColumn());
    const registros = registrosRange.getValues();
    
    let filaIndex = -1;
    let cuentaObj = null;
    
    for (let i = 0; i < registros.length; i++) {
      if (String(registros[i][0]).trim() === String(datos.ID_CARTERA).trim()) {
        filaIndex = i + 2; // +2 por encabezado y base 1
        cuentaObj = SEG_CONVERTIR_FILA_OBJETO(encabezados, registros[i]);
        break;
      }
    }
    
    if (!cuentaObj) throw new Error("La cuenta de cartera especificada no fue encontrada.");
    
    const valorRecaudo = Number(datos.VALOR);
    const saldoActual = Number(cuentaObj.SALDO);
    if (valorRecaudo > saldoActual) {
      throw new Error('El valor del recaudo ($' + valorRecaudo.toLocaleString() + ') supera el saldo pendiente ($' + saldoActual.toLocaleString() + ').');
    }
    
    // 1. Actualizar saldos en CAR_CUENTAS
    const nuevosAbonos = Number(cuentaObj.ABONOS) + valorRecaudo;
    const nuevoSaldo = saldoActual - valorRecaudo;
    let nuevoEstado = "PARCIAL";
    if (nuevoSaldo <= 0) {
      nuevoEstado = "PAGADO";
    }
    
    // Escribir de vuelta a la fila de la cuenta
    const colAbonos = encabezados.indexOf("ABONOS") + 1;
    const colSaldo = encabezados.indexOf("SALDO") + 1;
    const colEstado = encabezados.indexOf("ESTADO") + 1;
    
    hojaCuentas.getRange(filaIndex, colAbonos).setValue(nuevosAbonos);
    hojaCuentas.getRange(filaIndex, colSaldo).setValue(nuevoSaldo);
    hojaCuentas.getRange(filaIndex, colEstado).setValue(nuevoEstado);
    
    // 2. Registrar en la tabla CAR_RECAUDOS
    // Columnas: ID_RECAUDO, ID_CARTERA, ID_CLIENTE, ID_VENTA, FECHA, VALOR, METODO_PAGO, ID_CUENTA_DESTINO, RESPONSABLE, OBSERVACION
    const hojaRecaudos = ss.getSheetByName(CAR_CONFIG.HOJA_RECAUDOS);
    if (!hojaRecaudos) throw new Error("Hoja CAR_RECAUDOS no encontrada.");
    
    const idRecaudo = CAR_CONFIG.PREFIJO_RECAUDO + '-' + String(Math.max(1, hojaRecaudos.getLastRow())).padStart(CAR_CONFIG.DIGITOS_ID, '0');
    const ahora = new Date();
    
    hojaRecaudos.appendRow([
      idRecaudo,
      datos.ID_CARTERA,
      cuentaObj.ID_CLIENTE,
      cuentaObj.ID_VENTA,
      ahora,
      valorRecaudo,
      datos.METODO_PAGO || "TRANSFERENCIA",
      datos.ID_CUENTA_DESTINO || "CTA-000001",
      usuarioEjecutor,
      datos.OBSERVACION || "Recaudo sobre cartera de la venta " + cuentaObj.ID_VENTA
    ]);
    
    // 3. Registrar un movimiento de ingreso en TES_MOVIMIENTOS (Tesorería)
    try {
      const hojaTes = ss.getSheetByName("TES_MOVIMIENTOS");
      if (hojaTes) {
        const idTes = 'MOV-' + String(Math.max(1, hojaTes.getLastRow())).padStart(6, '0');
        // Columnas: ID_MOVIMIENTO, FECHA, TIPO_MOVIMIENTO, ID_CUENTA, ORIGEN, ID_ORIGEN, INGRESO, EGRESO, SALDO, METODO_PAGO, RESPONSABLE, OBSERVACION
        const hojaCuentasTes = ss.getSheetByName("TES_CUENTAS");
        let nuevoSaldoBanco = 0;
        if (hojaCuentasTes) {
          const cuentasData = hojaCuentasTes.getRange(2, 1, Math.max(1, hojaCuentasTes.getLastRow() - 1), 10).getValues();
          const cuentaIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA_DESTINO).trim());
          if (cuentaIndex !== -1) {
            const saldoBancario = Number(cuentasData[cuentaIndex][6] || 0); // columna SALDO_INICIAL / SALDO
            nuevoSaldoBanco = saldoBancario + valorRecaudo;
            // Actualizar saldo de la cuenta de tesorería
            hojaCuentasTes.getRange(cuentaIndex + 2, 7).setValue(nuevoSaldoBanco);
          }
        }
        
        hojaTes.appendRow([
          idTes,
          ahora,
          "INGRESO",
          datos.ID_CUENTA_DESTINO || "CTA-000001",
          "CARTERA",
          idRecaudo,
          valorRecaudo, // Ingreso
          0,            // Egreso
          nuevoSaldoBanco, // Nuevo Saldo
          datos.METODO_PAGO || "TRANSFERENCIA",
          usuarioEjecutor,
          datos.OBSERVACION || "Recaudo de cartera de la venta " + cuentaObj.ID_VENTA + " (" + cuentaObj.ID_CLIENTE + ")"
        ]);
      }
    } catch (errTes) {
      console.warn("No se pudo afectar el libro de tesorería: " + errTes.toString());
    }
    
    // Registrar auditoría del pago
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "TESORERIA",
      SUBMODULO: "CARTERA",
      ACCION: "APROBAR",
      TIPO_REGISTRO: "CAR_RECAUDOS",
      ID_REGISTRO: idRecaudo,
      DESCRIPCION: "Recaudo ingresado de la Cartera " + datos.ID_CARTERA + " por un total de: " + valorRecaudo + " COP. Saldo restante: " + nuevoSaldo + " COP.",
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      ID_RECAUDO: idRecaudo,
      NUEVO_SALDO: nuevoSaldo,
      MENSAJE: "Abono de cartera registrado correctamente. ID de recaudo: " + idRecaudo
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("CAR_REGISTRAR_RECAUDO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al registrar recaudo: " + error.toString() };
  }
}

/**
 * RPC: Obtener el listado de recaudos asociados a una cuenta de cartera
 */
function CAR_OBTENER_HISTORIAL_RECAUDOS_WEB(idCartera, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CAR_CONFIG.HOJA_RECAUDOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de recaudos no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran recaudos para esta cartera." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    
    const filtrados = registros
      .filter(fila => String(fila[1]).trim() === String(idCartera).trim())
      .map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
      
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(filtrados),
      MENSAJE: "Historial de recaudos recuperado exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al obtener historial de recaudos: " + error.toString() };
  }
}