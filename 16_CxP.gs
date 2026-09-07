// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 16_CxP.gs
* RESPONSABILIDAD:
* - Administrar las Cuentas por Pagar asociadas a compras.
* - Controlar saldos y amortizar obligaciones con abonos/pagos.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const CXP_CONFIG = {
  HOJA_CUENTAS: "CXP_CUENTAS",
  HOJA_PAGOS: "CXP_PAGOS",
  PREFIJO_ID: "CXP",
  PREFIJO_PAGO: "PAG",
  DIGITOS_ID: 6
};

/**
 * Crea una nueva obligación por pagar (CxP) a partir de una compra a crédito
 */
function CXP_CREAR_CUENTA_PAGAR(idCompra, idProveedor, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CXP_CONFIG.HOJA_CUENTAS);
  if (!hoja) return;
  
  const idCxp = CXP_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(CXP_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();
  
  // Plazo general de 30 días de crédito
  const vencimiento = new Date(ahora.getTime() + (30 * 24 * 60 * 60 * 1000));
  
  // Columnas: ID_CXP, ID_PROVEEDOR, ID_COMPRA, DOCUMENTO, FECHA_EMISION, FECHA_VENCIMIENTO, VALOR_DOCUMENTO, ABONOS, SALDO, DIAS_VENCIDOS, ESTADO
  hoja.appendRow([
    idCxp, 
    idProveedor, 
    idCompra, 
    idCompra, // Documento de soporte
    ahora, 
    vencimiento, 
    total, 
    0, // Abonos iniciales
    total, // Saldo inicial
    0, // Días vencidos iniciales
    "PENDIENTE"
  ]);
}

/**
 * RPC: Listar todas las cuentas por pagar, unificando con PROV_MAESTRO para el nombre comercial
 */
function CXP_LISTAR_CUENTAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER"); // Permiso financiero
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CXP_CONFIG.HOJA_CUENTAS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de cuentas por pagar no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay obligaciones pendientes." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con PROV_MAESTRO para el nombre real del proveedor
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
        item.RAZON_SOCIAL_PROVEEDOR = mapaProv[item.ID_PROVEEDOR] || "Proveedor Desconocido (" + item.ID_PROVEEDOR + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de obligaciones financieras obtenida de forma exitosa."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar CxP: " + error.toString() };
  }
}

/**
 * RPC: Registrar un abono o pago a una obligación (CxP) específica
 */
function CXP_REGISTRAR_PAGO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.ID_CXP || !datos.VALOR || datos.VALOR <= 0) {
      throw new Error("Datos de pago o valor no válidos.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCuentas = ss.getSheetByName(CXP_CONFIG.HOJA_CUENTAS);
    if (!hojaCuentas) throw new Error("Hoja CXP_CUENTAS no encontrada.");
    
    const ultimaFila = hojaCuentas.getLastRow();
    if (ultimaFila < 2) throw new Error("No hay cuentas por pagar registradas.");
    
    const encabezados = hojaCuentas.getRange(1, 1, 1, hojaCuentas.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registrosRange = hojaCuentas.getRange(2, 1, ultimaFila - 1, hojaCuentas.getLastColumn());
    const registros = registrosRange.getValues();
    
    let filaIndex = -1;
    let cuentaObj = null;
    
    for (let i = 0; i < registros.length; i++) {
      if (String(registros[i][0]).trim() === String(datos.ID_CXP).trim()) {
        filaIndex = i + 2; // +2 por encabezado y base 1
        cuentaObj = SEG_CONVERTIR_FILA_OBJETO(encabezados, registros[i]);
        break;
      }
    }
    
    if (!cuentaObj) throw new Error("La obligación CxP especificada no fue encontrada.");
    
    const valorPago = Number(datos.VALOR);
    const saldoActual = Number(cuentaObj.SALDO);
    if (valorPago > saldoActual) {
      throw new Error("El valor del pago ($" + valorPago.toLocaleString() + ") supera el saldo pendiente ($" + saldoActual.toLocaleString() + ").");
    }
    
    // 1. Actualizar saldos en CXP_CUENTAS
    const nuevosAbonos = Number(cuentaObj.ABONOS) + valorPago;
    const nuevoSaldo = saldoActual - valorPago;
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
    
    // 2. Registrar en la tabla CXP_PAGOS
    // Columnas: ID_PAGO, ID_CXP, ID_PROVEEDOR, ID_COMPRA, FECHA, VALOR, METODO_PAGO, ID_CUENTA_ORIGEN, RESPONSABLE, OBSERVACION
    const hojaPagos = ss.getSheetByName(CXP_CONFIG.HOJA_PAGOS);
    if (!hojaPagos) throw new Error("Hoja CXP_PAGOS no encontrada.");
    
    const idPago = CXP_CONFIG.PREFIJO_PAGO + "-" + String(Math.max(1, hojaPagos.getLastRow())).padStart(CXP_CONFIG.DIGITOS_ID, "0");
    const ahora = new Date();
    
    hojaPagos.appendRow([
      idPago,
      datos.ID_CXP,
      cuentaObj.ID_PROVEEDOR,
      cuentaObj.ID_COMPRA,
      ahora,
      valorPago,
      datos.METODO_PAGO || "TRANSFERENCIA",
      datos.ID_CUENTA_ORIGEN || "CTA-000001",
      usuarioEjecutor,
      datos.OBSERVACION || "Abono a obligación " + datos.ID_CXP
    ]);
    
    // 3. Registrar un movimiento de egreso en TES_MOVIMIENTOS (Tesoreria)
    try {
      const hojaTes = ss.getSheetByName("TES_MOVIMIENTOS");
      if (hojaTes) {
        const idTes = "MOV-" + String(Math.max(1, hojaTes.getLastRow())).padStart(6, "0");
        // Columnas: ID_MOVIMIENTO, FECHA, TIPO_MOVIMIENTO, ID_CUENTA, ORIGEN, ID_ORIGEN, INGRESO, EGRESO, SALDO, METODO_PAGO, RESPONSABLE, OBSERVACION
        // Obtenemos el saldo actual de la cuenta para recalcular el disponible
        const hojaCuentasTes = ss.getSheetByName("TES_CUENTAS");
        let nuevoSaldoBanco = 0;
        if (hojaCuentasTes) {
          const cuentasData = hojaCuentasTes.getRange(2, 1, Math.max(1, hojaCuentasTes.getLastRow() - 1), 10).getValues();
          const cuentaIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA_ORIGEN).trim());
          if (cuentaIndex !== -1) {
            const saldoBancario = Number(cuentasData[cuentaIndex][6] || 0); // columna SALDO_INICIAL / SALDO
            nuevoSaldoBanco = saldoBancario - valorPago;
            // Actualizar saldo de la cuenta de tesorería
            hojaCuentasTes.getRange(cuentaIndex + 2, 7).setValue(nuevoSaldoBanco);
          }
        }
        
        hojaTes.appendRow([
          idTes,
          ahora,
          "EGRESO",
          datos.ID_CUENTA_ORIGEN || "CTA-000001",
          "CXP",
          idPago,
          0,          // Ingreso
          valorPago,  // Egreso
          nuevoSaldoBanco, // Nuevo Saldo
          datos.METODO_PAGO || "TRANSFERENCIA",
          usuarioEjecutor,
          datos.OBSERVACION || "Pago de CxP " + datos.ID_CXP + " (" + cuentaObj.ID_PROVEEDOR + ")"
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
      SUBMODULO: "CXP",
      ACCION: "APROBAR",
      TIPO_REGISTRO: "CXP_PAGOS",
      ID_REGISTRO: idPago,
      DESCRIPCION: "Abono realizado a la CxP " + datos.ID_CXP + " por un total de: " + valorPago + " COP. Saldo restante: " + nuevoSaldo + " COP.",
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      DATOS: { ID_PAGO: idPago, SALDO_RESTANTE: nuevoSaldo, ESTADO: nuevoEstado },
      MENSAJE: "Abono registrado con éxito. ID de comprobante: " + idPago
    };
    
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("CXP_REGISTRAR_PAGO_WEB", "TESORERIA", error);
    }
    return { EXITO: false, MENSAJE: "Error al aplicar el pago: " + error.toString() };
  }
}

/**
 * RPC: Consultar el historial de pagos asociados a una obligación CxP específica
 */
function CXP_OBTENER_PAGOS_WEB(idCxp, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CXP_CONFIG.HOJA_PAGOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de pagos no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [] };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    
    const filtrados = registros
      .filter(fila => String(fila[1]).trim() === String(idCxp).trim())
      .map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
      
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(filtrados),
      MENSAJE: "Historial de pagos cargado correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al consultar pagos: " + error.toString() };
  }
}