/**************************************************************
* 11_INVENTARIO.gs
* RESPONSABILIDAD:
* - Administrar las existencias de productos en bodegas y obras.
* - Calcular de forma automatizada el costo promedio ponderado.
* - Registrar movimientos de Kardex (entradas, salidas, ajustes).
* - Exponer la consulta de inventarios de forma segura para la Web App.
**************************************************************/

const INV_CONFIG = {
  HOJA_MOVIMIENTOS: "INV_MOVIMIENTOS",
  HOJA_SALDOS: "INV_SALDOS",
  HOJA_KARDEX: "INV_KARDEX",
  PREFIJO_ID: "MOV",
  DIGITOS_ID: 6
};

function INV_REGISTRAR_MOVIMIENTO(mov, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "CREAR");
  if (!mov || !mov.ID_PRODUCTO) throw new Error("Movimiento de inventario no válido.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMov = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
  const hojaSaldos = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
  if (!hojaMov || !hojaSaldos) throw new Error("Hojas de inventario no encontradas.");

  const idMovimiento = INV_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(INV_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();

  // Calcular saldos y costo promedio
  let saldoActual = 0;
  let costoPromedio = Number(mov.COSTO_UNITARIO || 0);
  
  const saldosData = hojaSaldos.getLastRow() > 1 ? hojaSaldos.getRange(2, 1, hojaSaldos.getLastRow() - 1, 9).getValues() : [];
  let filaSaldo = -1;
  
  for (let i = 0; i < saldosData.length; i++) {
    if (String(saldosData[i][1]) === String(mov.ID_PRODUCTO)) {
      saldoActual = Number(saldosData[i][6] || 0);
      costoPromedio = Number(saldosData[i][7] || costoPromedio);
      filaSaldo = i + 2;
      break;
    }
  }

  const cant = Number(mov.CANTIDAD || 0);
  const costo = Number(mov.COSTO_UNITARIO || 0);
  let saldoFinal = saldoActual;

  if (mov.TIPO_MOVIMIENTO === "ENTRADA") {
    saldoFinal = saldoActual + cant;
    if (saldoFinal > 0) {
      costoPromedio = ((saldoActual * costoPromedio) + (cant * costo)) / saldoFinal;
    }
  } else if (mov.TIPO_MOVIMIENTO === "SALIDA") {
    saldoFinal = saldoActual - cant;
  }

  hojaMov.appendRow([
    idMovimiento, ahora, mov.TIPO_MOVIMIENTO, mov.ID_PRODUCTO, mov.ID_OBRA || "",
    mov.DOCUMENTO_ORIGEN || "", mov.ID_ORIGEN || "",
    mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : "",
    mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : "",
    costoPromedio, (cant * (mov.TIPO_MOVIMIENTO === "ENTRADA" ? costo : costoPromedio)),
    saldoFinal, Session.getActiveUser().getEmail() || "SISTEMA", ""
  ]);

  if (filaSaldo !== -1) {
    hojaSaldos.getRange(filaSaldo, 3, 1, 7).setValues([[
      ahora, saldoActual,
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]]);
  } else {
    hojaSaldos.appendRow([
      "SAL-" + idMovimiento, mov.ID_PRODUCTO, ahora, 0,
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal, costoPromedio, (saldoFinal * costoPromedio)
    ]);
  }

  return { EXITO: true, idMovimiento: idMovimiento, saldoFinal: saldoFinal, costoPromedio: costoPromedio };
}

function INV_LISTAR_SALDOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "La hoja INV_SALDOS no existe en la base de datos." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay saldos de inventario registrados." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Para conveniencia de la vista, buscaremos la descripción del producto uniendo con PROD_MAESTRO
    const hojaProd = ss.getSheetByName("PROD_MAESTRO");
    if (hojaProd) {
      const prodEnc = hojaProd.getRange(1, 1, 1, hojaProd.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const prodRegs = hojaProd.getLastRow() > 1 ? hojaProd.getRange(2, 1, hojaProd.getLastRow() - 1, hojaProd.getLastColumn()).getValues() : [];
      const mapProd = {};
      const idxId = prodEnc.indexOf("ID_PRODUCTO");
      const idxDesc = prodEnc.indexOf("DESCRIPCION");
      const idxCod = prodEnc.indexOf("CODIGO");
      prodRegs.forEach(r => {
        mapProd[String(r[idxId])] = { desc: r[idxDesc], cod: r[idxCod] };
      });
      
      lista.forEach(item => {
        const p = mapProd[String(item.ID_PRODUCTO)];
        item.DESCRIPCION_PRODUCTO = p ? p.desc : "Producto Desconocido";
        item.CODIGO_PRODUCTO = p ? p.cod : "N/A";
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Saldos de inventario obtenidos de forma segura."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("INV_LISTAR_SALDOS_WEB", "INVENTARIO", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los saldos de inventario: " + error.message };
  }
}


function INV_GUARDAR_AJUSTE(datos, tokenSesion) {
  const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "CREAR");
  const usuarioEjecutor = auth.USUARIO || "SISTEMA";
  if (!datos || !datos.ID_PRODUCTO) throw new Error("Datos de ajuste no válidos.");

  const mov = {
    TIPO_MOVIMIENTO: datos.TIPO_MOVIMIENTO,
    ID_PRODUCTO: datos.ID_PRODUCTO,
    CANTIDAD: Number(datos.CANTIDAD || 0),
    COSTO_UNITARIO: Number(datos.COSTO_UNITARIO || 0),
    DOCUMENTO_ORIGEN: "AJUSTE",
    ID_ORIGEN: "AJU-" + new Date().getTime(),
    OBSERVACION: datos.OBSERVACION || "Ajuste manual desde la vista del ERP"
  };

  const res = INV_REGISTRAR_MOVIMIENTO(mov, tokenSesion);
  return { EXITO: true, mensaje: "Ajuste de inventario registrado correctamente.", idMovimiento: res.idMovimiento };
}
