/**************************************************************
 * 11_INVENTARIO.gs (VERSIÓN 3.0 - V2 ERP - LIBRO 1 - PARCHE)
 * RESPONSABILIDAD:
 * - Administrar las existencias de productos en bodegas y obras.
 * - Calcular de forma automatizada el costo promedio ponderado.
 * - Registrar movimientos de Kardex (entradas, salidas, ajustes).
 * - Soportar la asignación y actualización de ubicaciones físicas de productos.
 * - Exponer un reporte Kardex PEPS (FIFO) de carácter INFORMATIVO.
 * - Exponer la consulta de inventarios de forma segura para la Web App.
 *
 * CAMBIOS RESPECTO A LA VERSIÓN ANTERIOR (2.0):
 * 1) MÉTODO DE COSTEO ÚNICO Y OFICIAL: el sistema valora inventario y costo
 *    de venta con PROMEDIO PONDERADO (el que alimenta INV_SALDOS y por lo
 *    tanto costo de venta / estados financieros). El reporte PEPS/FIFO se
 *    deja explícitamente como INFORMATIVO/GERENCIAL, y así se documenta y
 *    se marca en la respuesta (ADVERTENCIA), para que nadie lo tome por
 *    error como la valoración contable oficial. Si tu empresa decide
 *    adoptar PEPS como método oficial, hay que migrar también el costeo
 *    de ventas (módulo 09_VENTAS/13_COSTOS) para que use las capas PEPS en
 *    vez del promedio — ese cambio NO está incluido aquí porque no se
 *    revisó el código de esos módulos.
 * 2) CONCURRENCIA: INV_REGISTRAR_MOVIMIENTO ahora usa LockService para
 *    evitar que dos movimientos simultáneos del mismo producto lean el
 *    mismo saldo inicial y uno "pise" al otro (condición de carrera).
 * 3) VALIDACIÓN DE STOCK NEGATIVO: por defecto ya no se permite una SALIDA
 *    que deje el saldo en negativo (configurable con
 *    INV_CONFIG.PERMITIR_STOCK_NEGATIVO).
 **************************************************************/

const INV_CONFIG = {
  HOJA_MOVIMIENTOS: "INV_MOVIMIENTOS",
  HOJA_SALDOS: "INV_SALDOS",
  HOJA_KARDEX: "INV_KARDEX",
  PREFIJO_ID: "MOV",
  DIGITOS_ID: 6,
  METODO_COSTEO_OFICIAL: "PROMEDIO_PONDERADO", // método que alimenta INV_SALDOS / costo de venta
  PERMITIR_STOCK_NEGATIVO: false,              // cambia a true solo si tu operación lo requiere explícitamente
  NOMBRE_LOCK_INVENTARIO: "LOCK_INV_REGISTRAR_MOVIMIENTO",
  TIMEOUT_LOCK_MS: 30000
};

/**
 * Registra un movimiento de entrada o salida, afectando el Kardex y
 * recalculando Saldos con PROMEDIO PONDERADO (método oficial del sistema).
 *
 * Protegido con LockService: mientras un movimiento del mismo script se
 * está procesando, cualquier otro movimiento concurrente espera su turno
 * en vez de leer un saldo desactualizado.
 */
function INV_REGISTRAR_MOVIMIENTO(mov, tokenSesion) {
  SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "CREAR");
  if (!mov || !mov.ID_PRODUCTO) throw new Error("Movimiento de inventario no válido.");
  if (mov.TIPO_MOVIMIENTO !== "ENTRADA" && mov.TIPO_MOVIMIENTO !== "SALIDA") {
    throw new Error("TIPO_MOVIMIENTO debe ser 'ENTRADA' o 'SALIDA'.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(INV_CONFIG.TIMEOUT_LOCK_MS);
  } catch (e) {
    throw new Error("El sistema está ocupado procesando otro movimiento de inventario. Intenta de nuevo en unos segundos.");
  }

  try {
    return INV_REGISTRAR_MOVIMIENTO_INTERNO_(mov);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lógica real de registro, SIEMPRE debe ejecutarse ya dentro del lock
 * adquirido por INV_REGISTRAR_MOVIMIENTO. No llamar directamente desde
 * fuera de este archivo.
 */
function INV_REGISTRAR_MOVIMIENTO_INTERNO_(mov) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMov = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
  const hojaSaldos = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
  if (!hojaMov || !hojaSaldos) throw new Error("Hojas de inventario no encontradas.");

  const idMovimiento = INV_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(INV_CONFIG.DIGITOS_ID, "0");
  const ahora = new Date();

  let saldoActual = 0;
  let costoPromedio = Number(mov.COSTO_UNITARIO || 0);

  const saldosData = hojaSaldos.getLastRow() > 1 ? hojaSaldos.getRange(2, 1, hojaSaldos.getLastRow() - 1, 9).getValues() : [];
  let filaSaldo = -1;
  for (let i = 0; i < saldosData.length; i++) {
    if (String(saldosData[i][1]).trim() === String(mov.ID_PRODUCTO).trim()) {
      saldoActual = Number(saldosData[i][6] || 0);
      costoPromedio = Number(saldosData[i][7] || costoPromedio);
      filaSaldo = i + 2;
      break;
    }
  }

  const cant = Number(mov.CANTIDAD || 0);
  if (cant <= 0) throw new Error("La cantidad del movimiento debe ser mayor a cero.");
  const costoInput = Number(mov.COSTO_UNITARIO || 0);
  let saldoFinal = saldoActual;

  if (mov.TIPO_MOVIMIENTO === "ENTRADA") {
    saldoFinal = saldoActual + cant;
    // Recalcular Costo Promedio Ponderado (método oficial del sistema)
    if (saldoFinal > 0) {
      costoPromedio = ((saldoActual * costoPromedio) + (cant * costoInput)) / saldoFinal;
    } else {
      costoPromedio = costoInput;
    }
  } else if (mov.TIPO_MOVIMIENTO === "SALIDA") {
    saldoFinal = saldoActual - cant;
    if (saldoFinal < 0 && !INV_CONFIG.PERMITIR_STOCK_NEGATIVO) {
      throw new Error("Stock insuficiente para el producto " + mov.ID_PRODUCTO + ". Disponible: " + saldoActual + ", solicitado: " + cant + ".");
    }
    // El costo de salida es el costo promedio actual (método oficial)
  }

  // Guardar en INV_MOVIMIENTOS
  // Columnas: ID_MOVIMIENTO, FECHA, TIPO_MOVIMIENTO, ID_PRODUCTO, ID_OBRA, DOCUMENTO_ORIGEN, ID_ORIGEN, ENTRADA, SALIDA, COSTO_UNITARIO, COSTO_TOTAL, SALDO, RESPONSABLE, OBSERVACION
  hojaMov.appendRow([
    idMovimiento,
    ahora,
    mov.TIPO_MOVIMIENTO,
    mov.ID_PRODUCTO,
    mov.ID_OBRA || "",
    mov.DOCUMENTO_ORIGEN || "MANUAL",
    mov.ID_ORIGEN || "N/A",
    mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : "",
    mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : "",
    mov.TIPO_MOVIMIENTO === "ENTRADA" ? costoInput : costoPromedio,
    (cant * (mov.TIPO_MOVIMIENTO === "ENTRADA" ? costoInput : costoPromedio)),
    saldoFinal,
    Session.getActiveUser().getEmail() || "SISTEMA",
    mov.OBSERVACION || ""
  ]);

  // Actualizar o crear fila en INV_SALDOS
  // Columnas: ID_SALDO, ID_PRODUCTO, FECHA, SALDO_INICIAL, ENTRADAS, SALIDAS, SALDO_FINAL, COSTO_PROMEDIO, VALOR_INVENTARIO
  if (filaSaldo !== -1) {
    hojaSaldos.getRange(filaSaldo, 3, 1, 7).setValues([[
      ahora,
      saldoActual,
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal,
      costoPromedio,
      (saldoFinal * costoPromedio)
    ]]);
  } else {
    hojaSaldos.appendRow([
      "SAL-" + idMovimiento,
      mov.ID_PRODUCTO,
      ahora,
      0, // Saldo Inicial
      mov.TIPO_MOVIMIENTO === "ENTRADA" ? cant : 0,
      mov.TIPO_MOVIMIENTO === "SALIDA" ? cant : 0,
      saldoFinal,
      costoPromedio,
      (saldoFinal * costoPromedio)
    ]);
  }

  return { EXITO: true, idMovimiento: idMovimiento, saldoFinal: saldoFinal, costoPromedio: costoPromedio, metodoCosteo: INV_CONFIG.METODO_COSTEO_OFICIAL };
}

/**
 * RPC: Listar todos los saldos de bodega uniendo con el maestro de productos
 * para obtener datos adicionales. Sin cambios funcionales respecto a la
 * versión anterior.
 */
function INV_LISTAR_SALDOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaSaldos = ss.getSheetByName(INV_CONFIG.HOJA_SALDOS);
    if (!hojaSaldos) return { EXITO: false, DATOS: [], MENSAJE: "La hoja de saldos no existe." };

    const ultimaFila = hojaSaldos.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay existencias registradas." };

    const encabezados = hojaSaldos.getRange(1, 1, 1, hojaSaldos.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hojaSaldos.getRange(2, 1, ultimaFila - 1, hojaSaldos.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));

    const hojaProd = ss.getSheetByName("PROD_MAESTRO");
    if (hojaProd) {
      const prodEnc = hojaProd.getRange(1, 1, 1, hojaProd.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const prodRegs = hojaProd.getLastRow() > 1 ? hojaProd.getRange(2, 1, hojaProd.getLastRow() - 1, hojaProd.getLastColumn()).getValues() : [];
      const mapProd = {};
      const idxId = prodEnc.indexOf("ID_PRODUCTO");
      const idxDesc = prodEnc.indexOf("DESCRIPCION");
      const idxCod = prodEnc.indexOf("CODIGO");
      const idxUbi = prodEnc.indexOf("UBICACION");
      const idxImg = prodEnc.indexOf("IMAGEN_URL");
      const idxObs = prodEnc.indexOf("OBSERVACIONES");
      prodRegs.forEach(r => {
        mapProd[String(r[idxId]).trim()] = {
          desc: r[idxDesc],
          cod: r[idxCod],
          ubi: idxUbi !== -1 ? r[idxUbi] : "",
          img: idxImg !== -1 ? r[idxImg] : "",
          obs: idxObs !== -1 ? r[idxObs] : ""
        };
      });
      lista.forEach(item => {
        const p = mapProd[String(item.ID_PRODUCTO).trim()];
        item.DESCRIPCION_PRODUCTO = p ? p.desc : "Producto Desconocido";
        item.CODIGO_PRODUCTO = p ? p.cod : "N/A";
        item.UBICACION = p ? p.ubi : "BODEGA_PRINCIPAL";
        item.IMAGEN_URL = p ? p.img : "";
        item.OBSERVACIONES_PRODUCTO = p ? p.obs : "";
      });
    }

    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      METODO_COSTEO: INV_CONFIG.METODO_COSTEO_OFICIAL,
      MENSAJE: "Saldos de inventario obtenidos de forma segura."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("INV_LISTAR_SALDOS_WEB", "INVENTARIO", error);
    }
    return { EXITO: false, DATOS: [], MENSAJE: "No se pudieron cargar los saldos: " + error.message };
  }
}

/**
 * RPC: Guardar un ajuste manual de inventario. Los ajustes de tipo SALIDA
 * también respetan la validación de stock negativo (a menos que se
 * habilite explícitamente lo contrario en INV_CONFIG).
 */
function INV_GUARDAR_AJUSTE(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "CREAR");
    if (!datos || !datos.ID_PRODUCTO) throw new Error("Datos de ajuste no válidos.");

    const mov = {
      TIPO_MOVIMIENTO: datos.TIPO_MOVIMIENTO,
      ID_PRODUCTO: datos.ID_PRODUCTO,
      CANTIDAD: Number(datos.CANTIDAD || 0),
      COSTO_UNITARIO: Number(datos.COSTO_UNITARIO || 0),
      DOCUMENTO_ORIGEN: "AJUSTE",
      ID_ORIGEN: "AJU-" + new Date().getTime(),
      OBSERVACION: datos.OBSERVACION || "Ajuste manual de inventario"
    };
    const res = INV_REGISTRAR_MOVIMIENTO(mov, tokenSesion);
    return { EXITO: true, mensaje: "Ajuste de inventario registrado correctamente.", idMovimiento: res.idMovimiento };
  } catch (error) {
    return { EXITO: false, mensaje: "Error al registrar ajuste: " + error.toString() };
  }
}

/**
 * RPC: Actualizar la ubicación física de un producto en el maestro.
 * Sin cambios funcionales respecto a la versión anterior.
 */
function INV_ACTUALIZAR_UBICACION_WEB(idProducto, ubicacion, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "EDITAR");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaProd = ss.getSheetByName("PROD_MAESTRO");
    if (!hojaProd) throw new Error("Hoja PROD_MAESTRO no encontrada.");

    const encabezados = hojaProd.getRange(1, 1, 1, hojaProd.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    let idxUbi = encabezados.indexOf("UBICACION");
    if (idxUbi === -1) {
      hojaProd.getRange(1, encabezados.length + 1).setValue("UBICACION");
      idxUbi = encabezados.length;
    }
    const idxID = encabezados.indexOf("ID_PRODUCTO");
    const registros = hojaProd.getRange(2, 1, Math.max(1, hojaProd.getLastRow() - 1), hojaProd.getLastColumn()).getValues();
    let filaModificar = -1;
    for (let i = 0; i < registros.length; i++) {
      if (String(registros[i][idxID]).trim() === String(idProducto).trim()) {
        filaModificar = i + 2;
        break;
      }
    }
    if (filaModificar === -1) throw new Error("Producto no encontrado en el maestro.");
    hojaProd.getRange(filaModificar, idxUbi + 1).setValue(String(ubicacion).trim());
    return { EXITO: true, MENSAJE: "Ubicación del producto actualizada correctamente a: " + ubicacion };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al actualizar ubicación: " + error.toString() };
  }
}

/**
 * RPC: Obtener el Kardex Cronológico y calcular las Capas de Costo PEPS
 * (FIFO) para un producto.
 *
 * *** IMPORTANTE (contable) ***
 * Este reporte es INFORMATIVO/GERENCIAL. El costeo OFICIAL del sistema
 * (el que valora INV_SALDOS y de ahí el costo de venta) es PROMEDIO
 * PONDERADO. Los valores de "costo por capa" que ves aquí NO deben usarse
 * para libros contables mientras el método oficial siga siendo promedio
 * ponderado, porque darán una utilidad bruta distinta a la reportada.
 */
function INV_OBTENER_KARDEX_PEPS_WEB(idProducto, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = ss.getSheetByName(INV_CONFIG.HOJA_MOVIMIENTOS);
    if (!hojaMov) return { EXITO: false, KARDEX: [], CAPAS_PEPS: [], MENSAJE: "Hoja de movimientos no encontrada." };

    const ultimaFila = hojaMov.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, KARDEX: [], CAPAS_PEPS: [], MENSAJE: "No hay movimientos registrados." };

    const encabezados = hojaMov.getRange(1, 1, 1, hojaMov.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hojaMov.getRange(2, 1, ultimaFila - 1, hojaMov.getLastColumn()).getValues();

    const movs = registros
      .map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila))
      .filter(m => String(m.ID_PRODUCTO).trim() === String(idProducto).trim());

    const capas = [];
    movs.forEach(m => {
      const cant = Number(m.ENTRADA || m.SALIDA || 0);
      const costo = Number(m.COSTO_UNITARIO || 0);
      if (m.TIPO_MOVIMIENTO === "ENTRADA") {
        capas.push({
          idMovimiento: m.ID_MOVIMIENTO,
          idOrigen: m.ID_ORIGEN || "N/A",
          origen: m.DOCUMENTO_ORIGEN || "AJUSTE",
          fecha: m.FECHA,
          cantidadOriginal: cant,
          cantidadRestante: cant,
          costoUnitario: costo,
          observacion: m.OBSERVACION || ""
        });
      } else if (m.TIPO_MOVIMIENTO === "SALIDA") {
        let pendienteConsumo = cant;
        for (let i = 0; i < capas.length; i++) {
          if (capas[i].cantidadRestante > 0) {
            if (pendienteConsumo >= capas[i].cantidadRestante) {
              pendienteConsumo -= capas[i].cantidadRestante;
              capas[i].cantidadRestante = 0;
            } else {
              capas[i].cantidadRestante -= pendienteConsumo;
              pendienteConsumo = 0;
              break;
            }
          }
        }
      }
    });

    const capasActivas = capas.filter(c => c.cantidadRestante > 0);

    return {
      EXITO: true,
      KARDEX: SEG_SANITIZAR_PARA_CLIENTE(movs),
      CAPAS_PEPS: SEG_SANITIZAR_PARA_CLIENTE(capasActivas),
      METODO_COSTEO_OFICIAL: INV_CONFIG.METODO_COSTEO_OFICIAL,
      ADVERTENCIA: "Este reporte PEPS es informativo. El costeo oficial del sistema es " + INV_CONFIG.METODO_COSTEO_OFICIAL + ".",
      MENSAJE: "Kardex y capas PEPS calculados correctamente."
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("INV_OBTENER_KARDEX_PEPS_WEB", "INVENTARIO", error);
    }
    return { EXITO: false, KARDEX: [], CAPAS_PEPS: [], MENSAJE: "Error al calcular Kardex PEPS: " + error.toString() };
  }
}
