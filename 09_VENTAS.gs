// (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 09_VENTAS.gs (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Módulo Core de Ventas, Facturación Comercial, Cotizaciones y Remisiones.
* - Registro en Sheets de VEN_CABECERA y VEN_DETALLE con soporte de AIU y Retenciones.
* - Afectación automática de Cartera (CAR_CUENTAS_COBRAR) y Descuento de Stock en Kardex (INV_SALDOS).
* - Integración RPC asíncrona unificada para la interfaz Web App F7_VEN_VIEW.
**************************************************************/

const VEN_CONFIG = {
  HOJA_CABECERA: "VEN_CABECERA",
  HOJA_DETALLE: "VEN_DETALLE",
  HOJA_CARTERA: "CAR_CUENTAS_COBRAR",
  PREFIJO_VENTA: "VEN",
  PREFIJO_FACTURA: "FAC",
  PREFIJO_REMISION: "REM",
  PREFIJO_COTIZACION: "COT",
  DIGITOS_ID: 6
};

function VEN_OBTENER_HOJA(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = ss.insertSheet(nombreHoja);
  }
  if (hoja.getLastColumn() === 0 || hoja.getLastRow() === 0) {
    if (nombreHoja === VEN_CONFIG.HOJA_CABECERA) {
      hoja.appendRow([
        "ID_VENTA", "FECHA", "ID_CLIENTE", "RAZON_SOCIAL_CLIENTE", "TIPO_DOCUMENTO", 
        "NUM_DOCUMENTO", "ID_OBRA", "FORMA_PAGO", "PLAZO_DIAS", "SUBTOTAL", 
        "VALOR_AIU", "TOTAL_IVA", "TOTAL_RETENCIONES", "TOTAL", "ESTADO", 
        "FECHA_CREACION", "USUARIO_CREACION", "OBSERVACIONES"
      ]);
      hoja.getRange(1, 1, 1, 18).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    } else if (nombreHoja === VEN_CONFIG.HOJA_DETALLE) {
      hoja.appendRow([
        "ID_DETALLE", "ID_VENTA", "ID_PRODUCTO", "CODIGO", "DESCRIPCION", 
        "CANTIDAD", "PRECIO_UNITARIO", "TARIFA_IVA", "VALOR_IVA", "SUBTOTAL"
      ]);
      hoja.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    } else if (nombreHoja === VEN_CONFIG.HOJA_CARTERA) {
      hoja.appendRow([
        "ID_CARTERA", "ID_VENTA", "ID_CLIENTE", "FECHA_EMISION", "FECHA_VENCIMIENTO",
        "VALOR_DOCUMENTO", "ABONOS", "SALDO", "ESTADO"
      ]);
      hoja.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  }
  return hoja;
}

/**
 * RPC: Cargar catálogos combinados (Clientes, Proveedores, Productos, Obras) para la vista de Ventas
 */
function VEN_OBTENER_CATALOGOS_COMBINADOS_WEB(tokenSesion) {
  try {
    const result = {
      CLIENTES: [],
      PRODUCTOS: [],
      OBRAS: []
    };

    // 1. Obtener Clientes y Proveedores
    try {
      if (typeof CLI_LISTAR_CLIENTES_WEB === "function") {
        const resCli = CLI_LISTAR_CLIENTES_WEB(tokenSesion);
        if (resCli && resCli.DATOS) result.CLIENTES = result.CLIENTES.concat(resCli.DATOS);
      }
    } catch (e1) {}

    try {
      if (typeof PROV_LISTAR_PROVEEDORES_WEB === "function") {
        const resProv = PROV_LISTAR_PROVEEDORES_WEB(tokenSesion);
        if (resProv && resProv.DATOS) {
          resProv.DATOS.forEach(p => {
            p.ID_CLIENTE = p.ID_CLIENTE || p.ID_PROVEEDOR;
            result.CLIENTES.push(p);
          });
        }
      }
    } catch (e2) {}

    // Deduplicar Clientes por ID/Documento
    const mapCli = {};
    result.CLIENTES.forEach(c => {
      const k = String(c.ID_CLIENTE || c.ID_PROVEEDOR || c.NUMERO_DOCUMENTO || "").trim();
      if (k && !mapCli[k]) mapCli[k] = c;
    });
    result.CLIENTES = Object.values(mapCli);

    // 2. Obtener Productos
    try {
      if (typeof PROD_LISTAR_PRODUCTOS_WEB === "function") {
        const resProd = PROD_LISTAR_PRODUCTOS_WEB(tokenSesion);
        if (resProd && resProd.DATOS) result.PRODUCTOS = resProd.DATOS;
      }
    } catch (e3) {}

    if (result.PRODUCTOS.length === 0) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const hojaProd = ss.getSheetByName("PROD_MAESTRO");
        if (hojaProd && hojaProd.getLastRow() >= 2) {
          const encP = hojaProd.getRange(1, 1, 1, hojaProd.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
          const regP = hojaProd.getRange(2, 1, hojaProd.getLastRow() - 1, encP.length).getValues();
          result.PRODUCTOS = regP.map(fila => {
            const obj = {};
            encP.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
            return obj;
          });
        }
      } catch (e4) {}
    }

    // 3. Obtener Obras
    try {
      if (typeof OBR_LISTAR_OBRAS_WEB === "function") {
        const resObras = OBR_LISTAR_OBRAS_WEB(tokenSesion);
        if (resObras && resObras.DATOS) result.OBRAS = resObras.DATOS;
      }
    } catch (e5) {}

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(result) : JSON.parse(JSON.stringify(result)),
      MENSAJE: "Catálogos de ventas obtenidos exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: { CLIENTES: [], PRODUCTOS: [], OBRAS: [] }, MENSAJE: error.toString() };
  }
}

/**
 * RPC: Listar todas las ventas/facturas emitidas para la interfaz Web App
 */
function VEN_LISTAR_VENTAS_WEB(tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "VER");
    }

    const hoja = VEN_OBTENER_HOJA(VEN_CONFIG.HOJA_CABECERA);
    const ulFila = hoja.getLastRow();
    if (ulFila < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay transacciones de venta registradas." };
    }

    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ulFila - 1, enc.length).getValues();

    const lista = reg.map(fila => {
      const obj = {};
      enc.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
      return obj;
    }).reverse();

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(lista) : JSON.parse(JSON.stringify(lista)),
      MENSAJE: "Ventas listadas correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al consultar ventas: " + error.toString() };
  }
}

/**
 * RPC: Registrar una nueva Venta / Factura / Remisión / Cotización en Sheets
 */
function VEN_GUARDAR_VENTA_WEB(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "CREAR");
    }

    if (!datos || !datos.ID_CLIENTE || !datos.ITEMS || datos.ITEMS.length === 0) {
      throw new Error("Debe seleccionar un cliente y agregar al menos un ítem al detalle de la venta.");
    }

    const hojaCab = VEN_OBTENER_HOJA(VEN_CONFIG.HOJA_CABECERA);
    const hojaDet = VEN_OBTENER_HOJA(VEN_CONFIG.HOJA_DETALLE);

    const idVenta = datos.ID_VENTA || (VEN_CONFIG.PREFIJO_VENTA + "-" + String(Math.max(1, hojaCab.getLastRow())).padStart(VEN_CONFIG.DIGITOS_ID, "0"));
    const ahora = new Date();
    const usuarioEjecutor = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "OPERADOR") : "OPERADOR";

    const tipoDoc = String(datos.TIPO_DOCUMENTO || "FACTURA").toUpperCase();
    const formaPago = String(datos.FORMA_PAGO || "CONTADO").toUpperCase();
    const estadoInicial = datos.ESTADO || (tipoDoc === "COTIZACION" ? "COTIZADO" : (formaPago === "CREDITO" ? "PENDIENTE_CARTERA" : "PAGADA"));

    // 1. Guardar Cabecera de Venta
    const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registroCab = encCab.map(col => {
      switch (col) {
        case "ID_VENTA": return idVenta;
        case "FECHA": return datos.FECHA || ahora;
        case "ID_CLIENTE": return datos.ID_CLIENTE;
        case "RAZON_SOCIAL_CLIENTE": return datos.RAZON_SOCIAL_CLIENTE || datos.NOMBRE_CLIENTE || "CLIENTE GENERAL";
        case "TIPO_DOCUMENTO": return tipoDoc;
        case "NUM_DOCUMENTO": return datos.NUM_DOCUMENTO || idVenta;
        case "ID_OBRA": return datos.ID_OBRA || "";
        case "FORMA_PAGO": return formaPago;
        case "PLAZO_DIAS": return Number(datos.PLAZO_DIAS || 30);
        case "SUBTOTAL": return Number(datos.SUBTOTAL || 0);
        case "VALOR_AIU": return Number(datos.VALOR_AIU || 0);
        case "TOTAL_IVA": return Number(datos.TOTAL_IVA || 0);
        case "TOTAL_RETENCIONES": return Number(datos.TOTAL_RETENCIONES || 0);
        case "TOTAL": return Number(datos.TOTAL || 0);
        case "ESTADO": return estadoInicial;
        case "FECHA_CREACION": return ahora;
        case "USUARIO_CREACION": return usuarioEjecutor;
        case "OBSERVACIONES": return datos.OBSERVACIONES || "Registrado desde Módulo de Ventas V2";
        default: return "";
      }
    });

    hojaCab.appendRow(registroCab);

    // 2. Guardar Detalle de Ítems
    const encDet = hojaDet.getRange(1, 1, 1, hojaDet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    datos.ITEMS.forEach((item, index) => {
      const idDetalle = idVenta + "-D" + String(index + 1).padStart(3, "0");
      const filaDet = encDet.map(col => {
        switch (col) {
          case "ID_DETALLE": return idDetalle;
          case "ID_VENTA": return idVenta;
          case "ID_PRODUCTO": return item.ID_PRODUCTO || "PROD-GEN";
          case "CODIGO": return item.CODIGO || "";
          case "DESCRIPCION": return item.DESCRIPCION || "Item de venta";
          case "CANTIDAD": return Number(item.CANTIDAD || 1);
          case "PRECIO_UNITARIO": return Number(item.PRECIO_UNITARIO || 0);
          case "TARIFA_IVA": return String(item.TARIFA_IVA || "19%");
          case "VALOR_IVA": return Number(item.VALOR_IVA || 0);
          case "SUBTOTAL": return Number(item.SUBTOTAL || 0);
          default: return "";
        }
      });
      hojaDet.appendRow(filaDet);

      // Descuento automático de Inventario Kardex si aplica
      if (tipoDoc !== "COTIZACION" && typeof INV_REGISTRAR_MOVIMIENTO_KARDEX === "function" && item.ID_PRODUCTO) {
        try {
          INV_REGISTRAR_MOVIMIENTO_KARDEX({
            ID_PRODUCTO: item.ID_PRODUCTO,
            TIPO_MOVIMIENTO: "SALIDA_VENTA",
            CANTIDAD: Number(item.CANTIDAD || 1),
            DOCUMENTO_SOPORTE: idVenta,
            OBSERVACION: "Venta a " + (datos.RAZON_SOCIAL_CLIENTE || datos.ID_CLIENTE)
          });
        } catch (errInv) { console.error("Aviso Kardex:", errInv.message); }
      }
    });

    // 3. Registrar Cartera (CxC) si la venta es a Crédito
    if (formaPago === "CREDITO" && tipoDoc !== "COTIZACION") {
      VEN_REGISTRAR_CARTERA_CXC(idVenta, datos);
    }

    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        ID_USUARIO: usuarioEjecutor,
        USUARIO: usuarioEjecutor,
        MODULO: "VENTAS",
        ACCION: "CREAR",
        TIPO_REGISTRO: tipoDoc,
        ID_REGISTRO: idVenta,
        DESCRIPCION: "Venta " + idVenta + " registrada por $" + datos.TOTAL + " a " + (datos.RAZON_SOCIAL_CLIENTE || datos.ID_CLIENTE),
        RESULTADO: "EXITOSO"
      });
    }

    return {
      EXITO: true,
      ID_VENTA: idVenta,
      MENSAJE: "¡" + tipoDoc + " '" + idVenta + "' emitida y registrada en Sheets con éxito!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al procesar venta: " + error.toString() };
  }
}

/**
 * Helper interno para registrar la obligación en Cartera (CAR_CUENTAS_COBRAR)
 */
function VEN_REGISTRAR_CARTERA_CXC(idVenta, datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaCar = ss.getSheetByName(VEN_CONFIG.HOJA_CARTERA);
    if (!hojaCar) {
      hojaCar = ss.insertSheet(VEN_CONFIG.HOJA_CARTERA);
      hojaCar.appendRow(["ID_CARTERA", "ID_VENTA", "ID_CLIENTE", "FECHA_EMISION", "FECHA_VENCIMIENTO", "VALOR_DOCUMENTO", "ABONOS", "SALDO", "ESTADO"]);
    }

    const idCartera = "CXC-" + String(Math.max(1, hojaCar.getLastRow())).padStart(6, "0");
    const plazoDias = Number(datos.PLAZO_DIAS || 30);
    const fechaEmision = new Date();
    const fechaVenc = new Date(fechaEmision.getTime() + (plazoDias * 24 * 60 * 60 * 1000));
    const totalMonto = Number(datos.TOTAL || 0);

    hojaCar.appendRow([
      idCartera, idVenta, datos.ID_CLIENTE, fechaEmision, fechaVenc, totalMonto, 0, totalMonto, "PENDIENTE"
    ]);
  } catch (e) {
    console.error("Aviso Cartera:", e.message);
  }
}

/**
 * RPC: Obtener el detalle completo (Cabecera + Ítems) de una venta específica
 */
function VEN_OBTENER_DETALLE_VENTA_WEB(idVenta, tokenSesion) {
  try {
    if (!idVenta) throw new Error("ID_VENTA es requerido.");
    const hojaCab = VEN_OBTENER_HOJA(VEN_CONFIG.HOJA_CABECERA);
    const hojaDet = VEN_OBTENER_HOJA(VEN_CONFIG.HOJA_DETALLE);

    const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const regCab = hojaCab.getRange(2, 1, Math.max(1, hojaCab.getLastRow() - 1), encCab.length).getValues();
    const idxIdCab = encCab.indexOf("ID_VENTA");

    const filaCab = regCab.find(r => String(r[idxIdCab]).trim() === String(idVenta).trim());
    if (!filaCab) throw new Error("Venta '" + idVenta + "' no encontrada.");

    const cabecera = {};
    encCab.forEach((col, idx) => { cabecera[col] = filaCab[idx] !== undefined ? filaCab[idx] : ""; });

    const encDet = hojaDet.getRange(1, 1, 1, hojaDet.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const regDet = hojaDet.getRange(2, 1, Math.max(1, hojaDet.getLastRow() - 1), encDet.length).getValues();
    const idxIdDet = encDet.indexOf("ID_VENTA");

    const items = regDet.filter(r => String(r[idxIdDet]).trim() === String(idVenta).trim()).map(r => {
      const item = {};
      encDet.forEach((col, idx) => { item[col] = r[idx] !== undefined ? r[idx] : ""; });
      return item;
    });

    const resObj = { CABECERA: cabecera, ITEMS: items };

    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(resObj) : JSON.parse(JSON.stringify(resObj)),
      MENSAJE: "Detalle de venta obtenido exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: error.toString() };
  }
}

