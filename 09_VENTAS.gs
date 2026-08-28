/**************************************************************
* 09_VENTAS.gs
* RESPONSABILIDAD:
* - Capturar las transacciones de ventas y pedidos de clientes.
* - Calcular de forma modular: Subtotal, Descuento, IVA y Total.
* - Afectar asíncronamente stock (INV_) y generar cartera (CAR_).
**************************************************************/

const VEN_CONFIG = {
  HOJA_CABECERA: "VEN_CABECERA",
  HOJA_DETALLE: "VEN_DETALLE",
  PREFIJO_ID: "VEN",
  DIGITOS_ID: 6
};

function VEN_GUARDAR_VENTA(cabecera, detalles, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "CREAR");
  if (!cabecera || !detalles || detalles.length === 0) {
    throw new Error("Transacción incompleta de venta.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCab = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
  const hojaDet = ss.getSheetByName(VEN_CONFIG.HOJA_DETALLE);

  const idVenta = VEN_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();

  let subtotal = 0;
  let descuento = 0;
  let iva = 0;

  detalles.forEach(function(det) {
    const cant = Number(det.CANTIDAD || 0);
    const precio = Number(det.PRECIO_UNITARIO || 0);
    const desc = Number(det.DESCUENTO || 0);
    const pctIva = Number(det.PORCENTAJE_IVA || 0) / 100;

    const lineaSub = cant * precio;
    const lineaDesc = desc;
    const lineaIva = (lineaSub - lineaDesc) * pctIva;

    subtotal += lineaSub;
    descuento += lineaDesc;
    iva += lineaIva;

    const idDetalle = "DET-" + Utilities.getUuid().substring(0, 8);
    hojaDet.appendRow([
      idDetalle, idVenta, det.ID_PRODUCTO, det.DESCRIPCION || "",
      cant, det.ID_UNIDAD || "", precio, desc, lineaIva, (lineaSub - lineaDesc + lineaIva)
    ]);

    try {
      INV_REGISTRAR_MOVIMIENTO({
        TIPO_MOVIMIENTO: "SALIDA",
        ID_PRODUCTO: det.ID_PRODUCTO,
        CANTIDAD: cant,
        ID_OBRA: cabecera.ID_OBRA,
        DOCUMENTO_ORIGEN: "VENTA",
        ID_ORIGEN: idVenta
      });
    } catch (err) {
      console.warn("Inventario omitido para venta: " + idVenta);
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

  const encCab = hojaCab.getRange(1, 1, 1, hojaCab.getLastColumn()).getValues()[0];
  const filaCab = encCab.map(col => cabecera[col] !== undefined ? cabecera[col] : "");
  hojaCab.appendRow(filaCab);

  try {
    if (cabecera.FORMA_PAGO === "CREDITO") {
      CAR_CREAR_CUENTA_COBRAR(idVenta, cabecera.ID_CLIENTE, total);
    }
  } catch (err) {
    console.warn("Cartera no configurada para venta: " + idVenta);
  }

  return { ok: true, idVenta: idVenta, total: total };
}

function VEN_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(VEN_CONFIG.HOJA_CABECERA);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return VEN_CONFIG.PREFIJO_ID + "-000001";
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(VEN_CONFIG.PREFIJO_ID + "-", ""), 10);
  return VEN_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(VEN_CONFIG.DIGITOS_ID, "0");
}
