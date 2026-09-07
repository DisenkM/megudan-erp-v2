/**************************************************************
* 09_VENTAS.gs (VERSIÓN 3.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Capturar, procesar y guardar las transacciones de ventas y pedidos de clientes.
* - Calcular de forma exacta bases gravables, IVA tradicional e IVA sobre AIU (Colombia).
* - Afectar asíncronamente existencias de inventario (INV_MOVIMIENTOS) y generar cartera (CAR_CUENTAS).
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const VEN_CONFIG = {
  HOJA_CABECERA: "VEN_CABECERA",
  HOJA_DETALLE: "VEN_DETALLE",
  PREFIJO_ID: "VEN",
  DIGITOS_ID: 6
};

/**
 * Endpoint RPC seguro para guardar una venta desde el Frontend SPA
 */
function VEN_GUARDAR_VENTA_WEB(cabecera, detalles, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    cabecera.USUARIO = usuarioEjecutor;
    const resultado = VEN_GUARDAR_VENTA(cabecera, detalles, tokenSesion);
    
    // Loguear acción de auditoría
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "VENTAS",
      SUBMODULO: "FACTURACION",
      ACCION: "CREAR",
      TIPO_REGISTRO: "VENTAS",
      ID_REGISTRO: resultado.idVenta,
      DESCRIPCION: "Venta registrada exitosamente. Documento N°: " + (cabecera.NUM_DOCUMENTO || "") + " por un total de: " + resultado.total + " COP.",
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(resultado),
      MENSAJE: "¡Venta guardada y procesada exitosamente en Sheets! ID Interno: " + resultado.idVenta
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("VEN_GUARDAR_VENTA_WEB", "VENTAS", error);
    }
    return {
      EXITO: false,
      MENSAJE: "No se pudo registrar la venta: " + error.toString()
    };
  }
}

/**
 * Endpoint RPC seguro para listar las ventas registradas
 */
function VEN_LISTAR_VENTAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de ventas no configurada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay registros de ventas." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con CLI_MAESTRO para jalar la Razón Social real
    const hojaClientes = ss.getSheetByName("CLI_MAESTRO");
    if (hojaClientes) {
      const cliEnc = hojaClientes.getRange(1, 1, 1, hojaClientes.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const cliRegs = hojaClientes.getLastRow() > 1 ? hojaClientes.getRange(2, 1, hojaClientes.getLastRow() - 1, hojaClientes.getLastColumn()).getValues() : [];
      const mapaClientes = {};
      
      cliRegs.forEach(row => {
        const c = SEG_CONVERTIR_FILA_OBJETO(cliEnc, row);
        mapaClientes[c.ID_CLIENTE] = c.RAZON_SOCIAL || c.NOMBRE_COMERCIAL || (c.PRIMER_NOMBRE + " " + c.PRIMER_APELLIDO);
      });
      
      lista.forEach(v => {
        v.RAZON_SOCIAL_CLIENTE = mapaClientes[v.ID_CLIENTE] || "Tercero Desconocido (" + v.ID_CLIENTE + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de transacciones de venta obtenida de forma exitosa."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar ventas: " + error.toString() };
  }
}

/**
 * Endpoint RPC seguro para consultar el detalle de una venta
 */
function VEN_OBTENER_DETALLE_VENTA_WEB(idVenta, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Obtener información de la cabecera de la venta
    const hojaCab = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
    if (!hojaCab) throw new Error("La hoja de cabecera de ventas no existe.");
    
    const ultCab = hojaCab.getLastRow();
    let cabeceraObj = {};
    if (ultCab >= 2) {
      const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const regCab = hojaCab.getRange(2, 1, ultCab - 1, hojaCab.getLastColumn()).getValues();
      const filaCab = regCab.find(f => String(f[0]) === String(idVenta));
      if (filaCab) {
        cabeceraObj = SEG_CONVERTIR_FILA_OBJETO(encCab, filaCab);
        
        // Unir con CLI_MAESTRO para jalar la Razón Social real
        const hojaClientes = ss.getSheetByName("CLI_MAESTRO");
        if (hojaClientes) {
          const cliEnc = hojaClientes.getRange(1, 1, 1, hojaClientes.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
          const cliRegs = hojaClientes.getLastRow() > 1 ? hojaClientes.getRange(2, 1, hojaClientes.getLastRow() - 1, hojaClientes.getLastColumn()).getValues() : [];
          const clienteFila = cliRegs.find(row => {
            const c = SEG_CONVERTIR_FILA_OBJETO(cliEnc, row);
            return String(c.ID_CLIENTE) === String(cabeceraObj.ID_CLIENTE);
          });
          if (clienteFila) {
            const cObj = SEG_CONVERTIR_FILA_OBJETO(cliEnc, clienteFila);
            cabeceraObj.RAZON_SOCIAL_CLIENTE = cObj.RAZON_SOCIAL || cObj.NOMBRE_COMERCIAL || (cObj.PRIMER_NOMBRE + " " + cObj.PRIMER_APELLIDO);
          } else {
            cabeceraObj.RAZON_SOCIAL_CLIENTE = "Tercero Desconocido (" + cabeceraObj.ID_CLIENTE + ")";
          }
        }
      }
    }
    
    // 2. Obtener detalles de la venta (artículos vendidos)
    const hojaDet = ss.getSheetByName(VEN_CONFIG.HOJA_DETALLE);
    if (!hojaDet) throw new Error("La hoja de detalle de ventas no existe.");
    
    const ultDet = hojaDet.getLastRow();
    let filtrados = [];
    if (ultDet >= 2) {
      const encDet = hojaDet.getRange(1, 1, 1, hojaDet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const regDet = hojaDet.getRange(2, 1, ultDet - 1, hojaDet.getLastColumn()).getValues();
      filtrados = regDet
        .filter(fila => String(fila[1]) === String(idVenta))
        .map(fila => SEG_CONVERTIR_FILA_OBJETO(encDet, fila));
    }
      
    return {
      EXITO: true,
      DATOS: {
        CABECERA: SEG_SANITIZAR_PARA_CLIENTE(cabeceraObj),
        DETALLES: SEG_SANITIZAR_PARA_CLIENTE(filtrados)
      },
      MENSAJE: "Detalle de venta obtenido correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: "Error al obtener detalles: " + error.toString() };
  }
}

/**
 * Motor central de guardado y afectación de ventas
 */
function VEN_GUARDAR_VENTA(cabecera, detalles, tokenSesion) {
  if (!cabecera || !detalles || detalles.length === 0) {
    throw new Error("La transacción de venta se encuentra incompleta o vacía.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCab = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
  const hojaDet = ss.getSheetByName(VEN_CONFIG.HOJA_DETALLE);
  if (!hojaCab || !hojaDet) throw new Error("Hojas físicas de ventas (VEN_CABECERA o VEN_DETALLE) no encontradas.");
  
  const idVenta = VEN_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  
  let subtotal = 0;
  let descuento = 0;
  let iva = 0;
  
  // Procesamiento de cálculos bajo normativa DIAN Colombia
  const esVentaAIU = (cabecera.TIPO_DOCUMENTO === "FACTURA_AIU" || cabecera.TIPO_VENTA === "AIU");
  
  detalles.forEach(function(det) {
    const cant = Number(det.CANTIDAD || 0);
    const precio = Number(det.PRECIO_UNITARIO || 0);
    const desc = Number(det.DESCUENTO || 0);
    let pctIva = Number(String(det.IVA || "0").replace("%", "")) / 100;
    
    let lineaSub = cant * precio;
    let lineaDesc = desc;
    let lineaIva = 0;
    
    if (esVentaAIU) {
      // Regla de AIU Colombia: El IVA se calcula EXCLUSIVAMENTE sobre la Utilidad (U)
      const a_pct = Number(cabecera.AIU_A || 15) / 100;
      const i_pct = Number(cabecera.AIU_I || 5) / 100;
      const u_pct = Number(cabecera.AIU_U || 10) / 100;
      
      const subtotalAIU = lineaSub - lineaDesc;
      const baseUtilidad = subtotalAIU * u_pct;
      
      lineaIva = baseUtilidad * 0.19; // El IVA del AIU es del 19% sobre la utilidad pactada
    } else {
      // IVA tradicional colombiano
      lineaIva = (lineaSub - lineaDesc) * pctIva;
    }
    
    subtotal += lineaSub;
    descuento += lineaDesc;
    iva += lineaIva;
    
    const idDetalle = "DET-" + Utilities.getUuid().substring(0, 8);
    
    // Escribir fila de VEN_DETALLE
    hojaDet.appendRow([
      idDetalle, 
      idVenta, 
      det.ID_PRODUCTO, 
      det.DESCRIPCION || "", 
      cant, 
      det.ID_UNIDAD || "UND", 
      precio, 
      lineaDesc, 
      lineaIva, 
      (lineaSub - lineaDesc + lineaIva)
    ]);
    
    // Descarga de Inventario (Kardex físico) de forma automatizada
    try {
      INV_REGISTRAR_MOVIMIENTO({
        TIPO_MOVIMIENTO: "SALIDA",
        ID_PRODUCTO: det.ID_PRODUCTO,
        CANTIDAD: cant,
        ID_OBRA: cabecera.ID_OBRA || "",
        DOCUMENTO_ORIGEN: "VENTA",
        ID_ORIGEN: idVenta,
        COSTO_UNITARIO: 0, // El Kardex de salida calcula costo promedio ponderado automáticamente
        OBSERVACION: "Salida automatizada por Facturación de Venta " + idVenta
      }, tokenSesion);
    } catch (err) {
      console.warn("Inventario omitido u opcional para venta: " + idVenta + ". Detalle: " + err.toString());
    }
  });
  
  const total = subtotal - descuento + iva;
  
  cabecera.ID_VENTA = idVenta;
  cabecera.FECHA = ahora;
  cabecera.SUBTOTAL = subtotal;
  cabecera.DESCUENTO = descuento;
  cabecera.IVA = iva;
  cabecera.TOTAL = total;
  cabecera.ESTADO = "EMITIDA";
  cabecera.FECHA_CREACION = ahora;
  
  // Calcular vencimiento de cartera según plazo comercial
  const plazoDias = Number(cabecera.PLAZO_PAGO_DIAS || 0);
  cabecera.FECHA_VENCIMIENTO = new Date(ahora.getTime() + (plazoDias * 24 * 60 * 60 * 1000));
  
  const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const filaCab = encCab.map(col => cabecera[col] !== undefined ? cabecera[col] : "");
  hojaCab.appendRow(filaCab);
  
  // Disparar Cartera (Cuentas por Cobrar) en caliente si es a CRÉDITO
  try {
    if (cabecera.FORMA_PAGO === "CREDITO") {
      CAR_CREAR_CUENTA_COBRAR(idVenta, cabecera.ID_CLIENTE, total);
    }
  } catch (err) {
    console.warn("Cartera de cliente omitida u opcional para venta: " + idVenta + ". Detalle: " + err.toString());
  }
  
  return { ok: true, idVenta: idVenta, total: total };
}

/**
 * Obtener consecutivo siguiente indexado de la cabecera
 */
function VEN_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return VEN_CONFIG.PREFIJO_ID + "-000001";
  
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(VEN_CONFIG.PREFIJO_ID + "-", ""), 10);
  if (isNaN(numero)) {
    return VEN_CONFIG.PREFIJO_ID + "-000001";
  }
  return VEN_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(VEN_CONFIG.DIGITOS_ID, "0");
}