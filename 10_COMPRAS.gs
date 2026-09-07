// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 10_COMPRAS.gs
* RESPONSABILIDAD:
* - Registrar las adquisiciones e insumos de proveedores.
* - Sincronizar costo promedio de inventario y registrar CxP.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const COM_CONFIG = {
  HOJA_CABECERA: "COM_CABECERA",
  HOJA_DETALLE: "COM_DETALLE",
  PREFIJO_ID: "COM",
  DIGITOS_ID: 6
};

/**
 * Endpoint RPC seguro para guardar una compra desde el Frontend SPA
 */
function COM_GUARDAR_COMPRA_WEB(cabecera, detalles, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    cabecera.USUARIO = usuarioEjecutor;
    const resultado = COM_GUARDAR_COMPRA(cabecera, detalles, tokenSesion);
    
    // Loguear acción de auditoría
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "COMPRAS",
      SUBMODULO: "OPERACION",
      ACCION: "CREAR",
      TIPO_REGISTRO: "COMPRAS",
      ID_REGISTRO: resultado.idCompra,
      DESCRIPCION: "Compra registrada exitosamente. Documento N°: " + (cabecera.NUM_DOCUMENTO || "") + " por un total de: " + resultado.total + " COP.",
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(resultado),
      MENSAJE: "¡Compra guardada y procesada exitosamente en Sheets! ID Interno: " + resultado.idCompra
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("COM_GUARDAR_COMPRA_WEB", "COMPRAS", error);
    }
    return {
      EXITO: false,
      MENSAJE: "No se pudo registrar la compra: " + error.toString()
    };
  }
}

/**
 * Endpoint RPC seguro para listar las compras registradas
 */
function COM_LISTAR_COMPRAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de compras no configurada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay registros de compras." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con PROV_MAESTRO para jalar la Razón Social real
    const hojaProv = ss.getSheetByName("PROV_MAESTRO");
    if (hojaProv) {
      const provEnc = hojaProv.getRange(1, 1, 1, hojaProv.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const provRegs = hojaProv.getLastRow() > 1 ? hojaProv.getRange(2, 1, hojaProv.getLastRow() - 1, hojaProv.getLastColumn()).getValues() : [];
      const mapaProv = {};
      
      provRegs.forEach(row => {
        const p = SEG_CONVERTIR_FILA_OBJETO(provEnc, row);
        mapaProv[p.ID_PROVEEDOR] = p.RAZON_SOCIAL || p.NOMBRE_COMERCIAL || p.NOMBRE_CONTACTO;
      });
      
      lista.forEach(v => {
        v.RAZON_SOCIAL_PROVEEDOR = mapaProv[v.ID_PROVEEDOR] || "Proveedor Desconocido (" + v.ID_PROVEEDOR + ")";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de transacciones de compra obtenida de forma exitosa."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar compras: " + error.toString() };
  }
}

/**
 * Endpoint RPC seguro para consultar el detalle de una compra
 */
function COM_OBTENER_DETALLE_COMPRA_WEB(idCompra, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "COMPRAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Obtener información de la cabecera de la compra
    const hojaCab = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
    if (!hojaCab) throw new Error("La hoja de cabecera de compras no existe.");
    
    const ultCab = hojaCab.getLastRow();
    let cabeceraObj = {};
    if (ultCab >= 2) {
      const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const regCab = hojaCab.getRange(2, 1, ultCab - 1, hojaCab.getLastColumn()).getValues();
      const filaCab = regCab.find(f => String(f[0]) === String(idCompra));
      if (filaCab) {
        cabeceraObj = SEG_CONVERTIR_FILA_OBJETO(encCab, filaCab);
        
        // Unir con PROV_MAESTRO para jalar la Razón Social real
        const hojaProv = ss.getSheetByName("PROV_MAESTRO");
        if (hojaProv) {
          const provEnc = hojaProv.getRange(1, 1, 1, hojaProv.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
          const provRegs = hojaProv.getLastRow() > 1 ? hojaProv.getRange(2, 1, hojaProv.getLastRow() - 1, hojaProv.getLastColumn()).getValues() : [];
          const provFila = provRegs.find(row => {
            const p = SEG_CONVERTIR_FILA_OBJETO(provEnc, row);
            return String(p.ID_PROVEEDOR) === String(cabeceraObj.ID_PROVEEDOR);
          });
          if (provFila) {
            const pObj = SEG_CONVERTIR_FILA_OBJETO(provEnc, provFila);
            cabeceraObj.RAZON_SOCIAL_PROVEEDOR = pObj.RAZON_SOCIAL || pObj.NOMBRE_COMERCIAL || pObj.NOMBRE_CONTACTO;
          } else {
            cabeceraObj.RAZON_SOCIAL_PROVEEDOR = "Proveedor Desconocido (" + cabeceraObj.ID_PROVEEDOR + ")";
          }
        }
      }
    }
    
    // 2. Obtener detalles de la compra (artículos comprados)
    const hojaDet = ss.getSheetByName(COM_CONFIG.HOJA_DETALLE);
    if (!hojaDet) throw new Error("La hoja de detalle de compras no existe.");
    
    const ultDet = hojaDet.getLastRow();
    let filtrados = [];
    if (ultDet >= 2) {
      const encDet = hojaDet.getRange(1, 1, 1, hojaDet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const regDet = hojaDet.getRange(2, 1, ultDet - 1, hojaDet.getLastColumn()).getValues();
      filtrados = regDet
        .filter(fila => String(fila[1]) === String(idCompra))
        .map(fila => SEG_CONVERTIR_FILA_OBJETO(encDet, fila));
    }
      
    return {
      EXITO: true,
      DATOS: {
        CABECERA: SEG_SANITIZAR_PARA_CLIENTE(cabeceraObj),
        DETALLES: SEG_SANITIZAR_PARA_CLIENTE(filtrados)
      },
      MENSAJE: "Detalle de compra obtenido correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: "Error al obtener detalles: " + error.toString() };
  }
}

/**
 * Motor central de guardado y afectación de compras
 */
function COM_GUARDAR_COMPRA(cabecera, detalles, tokenSesion) {
  if (!cabecera || !detalles || detalles.length === 0) {
    throw new Error("La transacción de compra se encuentra incompleta o vacía.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCab = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
  const hojaDet = ss.getSheetByName(COM_CONFIG.HOJA_DETALLE);
  if (!hojaCab || !hojaDet) throw new Error("Hojas físicas de compras (COM_CABECERA o COM_DETALLE) no encontradas.");
  
  const idCompra = COM_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  
  let subtotal = 0;
  let descuento = 0;
  let iva = 0;
  
  detalles.forEach(function(det) {
    const cant = Number(det.CANTIDAD || 0);
    const costo = Number(det.COSTO_UNITARIO || 0);
    const desc = Number(det.DESCUENTO || 0);
    let pctIva = Number(String(det.PORCENTAJE_IVA || "0").replace("%", "")) / 100;
    
    let lineaSub = cant * costo;
    let lineaDesc = desc;
    let lineaIva = (lineaSub - lineaDesc) * pctIva;
    
    subtotal += lineaSub;
    descuento += lineaDesc;
    iva += lineaIva;

    const idDetalle = "DET-" + Utilities.getUuid().substring(0, 8);
    
    // Escribir fila de COM_DETALLE
    hojaDet.appendRow([
      idDetalle, 
      idCompra, 
      det.ID_PRODUCTO, 
      det.DESCRIPCION || "", 
      cant, 
      det.ID_UNIDAD || "UND", 
      costo, 
      lineaDesc, 
      lineaIva, 
      (lineaSub - lineaDesc + lineaIva)
    ]);
    
    // Entrada de Inventario (Kardex físico) de forma automatizada
    try {
      INV_REGISTRAR_MOVIMIENTO({
        TIPO_MOVIMIENTO: "ENTRADA",
        ID_PRODUCTO: det.ID_PRODUCTO,
        CANTIDAD: cant,
        COSTO_UNITARIO: costo,
        ID_OBRA: cabecera.ID_OBRA || "",
        DOCUMENTO_ORIGEN: "COMPRA",
        ID_ORIGEN: idCompra,
        OBSERVACION: "Entrada automática por Facturación de Compra " + idCompra
      }, tokenSesion);
    } catch (err) {
      console.warn("Inventario omitido u opcional para compra: " + idCompra + ". Detalle: " + err.toString());
    }
  });
  
  const total = subtotal - descuento + iva;
  
  cabecera.ID_COMPRA = idCompra;
  cabecera.FECHA = ahora;
  cabecera.SUBTOTAL = subtotal;
  cabecera.DESCUENTO = descuento;
  cabecera.IVA = iva;
  cabecera.TOTAL = total;
  cabecera.ESTADO = "RECIBIDA";
  cabecera.FECHA_CREACION = ahora;
  
  // Calcular vencimiento de cartera según plazo comercial
  const plazoDias = Number(cabecera.PLAZO_PAGO_DIAS || 0);
  cabecera.FECHA_VENCIMIENTO = new Date(ahora.getTime() + (plazoDias * 24 * 60 * 60 * 1000));
  
  const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const filaCab = encCab.map(col => cabecera[col] !== undefined ? cabecera[col] : "");
  hojaCab.appendRow(filaCab);
  
  // Disparar Cuentas por Pagar (CXP) en caliente si es a CRÉDITO
  try {
    if (cabecera.FORMA_PAGO === "CREDITO") {
      CXP_CREAR_CUENTA_PAGAR(idCompra, cabecera.ID_PROVEEDOR, total);
    }
  } catch (err) {
    console.warn("CxP omitida u opcional para compra: " + idCompra + ". Detalle: " + err.toString());
  }
  
  return { ok: true, idCompra: idCompra, total: total };
}

/**
 * Obtener consecutivo siguiente indexado de la cabecera
 */
function COM_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return COM_CONFIG.PREFIJO_ID + "-000001";
  
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(COM_CONFIG.PREFIJO_ID + "-", ""), 10);
  if (isNaN(numero)) {
    return COM_CONFIG.PREFIJO_ID + "-000001";
  }
  return COM_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(COM_CONFIG.DIGITOS_ID, "0");
}
