/**************************************************************
* 10_COMPRAS.gs
* RESPONSABILIDAD:
* - Registrar las adquisiciones e insumos de proveedores.
* - Sincronizar el costo promedio de inventario y registrar la cuenta por pagar (CXP_).
**************************************************************/

const COM_CONFIG = {
  HOJA_CABECERA: "COM_CABECERA",
  HOJA_DETALLE: "COM_DETALLE",
  PREFIJO_ID: "COM",
  DIGITOS_ID: 6
};

function COM_GUARDAR_COMPRA(cabecera, detalles) {
  if (!cabecera || !detalles || detalles.length === 0) {
    throw new Error("Transacción incompleta.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCab = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
  const hojaDet = ss.getSheetByName(COM_CONFIG.HOJA_DETALLE);

  const idCompra = COM_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  let subtotal = 0;
  let descuento = 0;
  let iva = 0;

  detalles.forEach(function(det) {
    const cant = Number(det.CANTIDAD || 0);
    const costo = Number(det.COSTO_UNITARIO || 0);
    const desc = Number(det.DESCUENTO || 0);
    const pctIva = Number(det.PORCENTAJE_IVA || 0) / 100;

    const lineaSub = cant * costo;
    const lineaDesc = desc;
    const lineaIva = (lineaSub - lineaDesc) * pctIva;

    subtotal += lineaSub;
    descuento += lineaDesc;
    iva += lineaIva;

    const idDetalle = "DET-" + Utilities.getUuid().substring(0, 8);
    hojaDet.appendRow([
      idDetalle, idCompra, det.ID_PRODUCTO, det.DESCRIPCION, cant, 
      det.ID_UNIDAD, costo, desc, lineaIva, (lineaSub - lineaDesc + lineaIva)
    ]);

    // Afectar Inventario (Kardex de Entrada / Costo Promedio)
    try {
      INV_REGISTRAR_MOVIMIENTO({
        TIPO_MOVIMIENTO: "ENTRADA",
        ID_PRODUCTO: det.ID_PRODUCTO,
        CANTIDAD: cant,
        COSTO_UNITARIO: costo,
        DOCUMENTO_ORIGEN: "COMPRA",
        ID_ORIGEN: idCompra
      });
    } catch (err) {
      console.warn("Inventario omitido para compra: " + idCompra);
    }
  });

  cabecera.ID_COMPRA = idCompra;
  cabecera.FECHA = ahora;
  cabecera.SUBTOTAL = subtotal;
  cabecera.DESCUENTO = descuento;
  cabecera.IVA = iva;
  cabecera.TOTAL = subtotal - descuento + iva;
  cabecera.ESTADO = "RECIBIDO";
  cabecera.FECHA_CREACION = ahora;

  const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0];
  const filaCab = encCab.map(col => cabecera[col] !== undefined ? cabecera[col] : "");
  hojaCab.appendRow(filaCab);

  // Afectar Cuentas por Pagar si es a crédito
  if (cabecera.FORMA_PAGO === "CREDITO") {
    try {
      CXP_CREAR_CUENTA_PAGAR(idCompra, cabecera.ID_PROVEEDOR, cabecera.TOTAL);
    } catch (err) {
      console.warn("No se pudo generar cuenta por pagar.");
    }
  }

  return { ok: true, idCompra: idCompra, total: cabecera.TOTAL };
}

function COM_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(COM_CONFIG.HOJA_CABECERA);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return COM_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(COM_CONFIG.PREFIJO_ID + "-", ""), 10);
  return COM_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(COM_CONFIG.DIGITOS_ID, "0");
}
