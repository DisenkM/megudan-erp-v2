// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 08_OBRAS.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el maestro de Obras y Proyectos (OBR_MAESTRO).
* - Presupuestación de recursos por obra (OBR_PRESUPUESTO).
* - Control de avance físico/financiero (OBR_AVANCE).
* - Asignación de materiales, mano de obra y equipos (OBR_RECURSOS).
**************************************************************/

const OBR_CONFIG_CORE = {
  HOJA_MAESTRO: "OBR_MAESTRO",
  HOJA_PRESUPUESTO: "OBR_PRESUPUESTO",
  HOJA_AVANCE: "OBR_AVANCE",
  HOJA_RECURSOS: "OBR_RECURSOS",
  PREFIJO_ID: "OBR",
  PREFIJO_PRESUPUESTO: "PRE-OBR",
  PREFIJO_AVANCE: "AVN",
  PREFIJO_RECURSO: "REC-OBR",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar todas las obras registradas
 */
function OBR_LISTAR_OBRAS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "OBRAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(OBR_CONFIG_CORE.HOJA_MAESTRO);
    if (!hoja) {
      hoja = ss.insertSheet(OBR_CONFIG_CORE.HOJA_MAESTRO);
      hoja.setTabColor("#92400e");
      hoja.appendRow(["ID_OBRA", "CODIGO_OBRA", "NOMBRE_OBRA", "ID_CLIENTE", "DESCRIPCION", "UBICACION", "CIUDAD", "FECHA_INICIO", "FECHA_FIN_ESTIMADA", "FECHA_FIN_REAL", "PRESUPUESTO", "RESPONSABLE", "ESTADO", "FECHA_CREACION"]);
    }
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran obras o proyectos en la base de datos." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unificar Razón Social del Cliente
    const hojaClientes = ss.getSheetByName("CLI_MAESTRO");
    if (hojaClientes && hojaClientes.getLastRow() > 1) {
      const cliEnc = hojaClientes.getRange(1, 1, 1, hojaClientes.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const cliRegs = hojaClientes.getRange(2, 1, hojaClientes.getLastRow() - 1, hojaClientes.getLastColumn()).getValues();
      const mapaClientes = {};
      cliRegs.forEach(r => {
        const c = SEG_CONVERTIR_FILA_OBJETO(cliEnc, r);
        mapaClientes[c.ID_CLIENTE] = c.RAZON_SOCIAL || (c.PRIMER_NOMBRE + " " + c.PRIMER_APELLIDO);
      });
      lista.forEach(o => {
        o.RAZON_SOCIAL_CLIENTE = mapaClientes[o.ID_CLIENTE] || "Cliente " + o.ID_CLIENTE;
      });
    }
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Obras y proyectos recuperados exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar obras: " + error.toString() };
  }
}

/**
 * RPC: Guardar o actualizar una obra/proyecto
 */
function OBR_GUARDAR_OBRA_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "OBRAS", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.NOMBRE_OBRA || !datos.ID_CLIENTE) {
      throw new Error("El nombre de la obra y el cliente son campos obligatorios.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(OBR_CONFIG_CORE.HOJA_MAESTRO);
    if (!hoja) {
      hoja = ss.insertSheet(OBR_CONFIG_CORE.HOJA_MAESTRO);
      hoja.setTabColor("#92400e");
      hoja.appendRow(["ID_OBRA", "CODIGO_OBRA", "NOMBRE_OBRA", "ID_CLIENTE", "DESCRIPCION", "UBICACION", "CIUDAD", "FECHA_INICIO", "FECHA_FIN_ESTIMADA", "FECHA_FIN_REAL", "PRESUPUESTO", "RESPONSABLE", "ESTADO", "FECHA_CREACION"]);
    }
    
    const idObra = datos.ID_OBRA || (OBR_CONFIG_CORE.PREFIJO_ID + "-" + String(Math.max(1, hoja.getLastRow())).padStart(OBR_CONFIG_CORE.DIGITOS_ID, "0"));
    const codigoObra = datos.CODIGO_OBRA || ("OBR-" + datos.NOMBRE_OBRA.substring(0, 3).toUpperCase() + "-" + String(Math.floor(Math.random() * 900) + 100));
    const ahora = new Date();
    
    hoja.appendRow([
      idObra,
      codigoObra,
      datos.NOMBRE_OBRA.trim(),
      datos.ID_CLIENTE.trim(),
      datos.DESCRIPCION || "",
      datos.UBICACION || "Pitalito, Huila",
      datos.CIUDAD || "Pitalito",
      datos.FECHA_INICIO || ahora,
      datos.FECHA_FIN_ESTIMADA || ahora,
      "",
      Number(datos.PRESUPUESTO || 0),
      datos.RESPONSABLE || usuarioEjecutor,
      datos.ESTADO || "ACTIVA",
      ahora
    ]);
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "OBRAS",
      SUBMODULO: "MAESTRO",
      ACCION: "CREAR",
      TIPO_REGISTRO: "OBR_MAESTRO",
      ID_REGISTRO: idObra,
      DESCRIPCION: "Obra '" + datos.NOMBRE_OBRA + "' creada por $" + datos.PRESUPUESTO,
      RESULTADO: "EXITOSO"
    });
    
    return { EXITO: true, ID_OBRA: idObra, MENSAJE: "¡Obra / Proyecto '" + datos.NOMBRE_OBRA + "' creado exitosamente!" };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar obra: " + error.toString() };
  }
}

/**
 * RPC: Obtener el detalle transaccional de costos, recursos e ingresos de una obra
 */
function OBR_OBTENER_DETALLE_OBRA_WEB(idObra, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "OBRAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let totalIngresos = 0;
    let totalCostosDirectos = 0;
    let totalGastosObra = 0;
    const listaCostos = [];
    
    // 1. Ingresos por ventas asociadas a la obra
    const hojaVentas = ss.getSheetByName("VEN_CABECERA");
    if (hojaVentas && hojaVentas.getLastRow() > 1) {
      const vData = hojaVentas.getRange(2, 1, hojaVentas.getLastRow() - 1, 15).getValues();
      vData.forEach(v => {
        if (String(v[5]).trim() === String(idObra).trim() && v[12] !== "ANULADA") {
          totalIngresos += Number(v[11] || 0);
        }
      });
    }
    
    // 2. Costos directos de la obra
    const hojaCostos = ss.getSheetByName("COS_MOVIMIENTOS");
    if (hojaCostos && hojaCostos.getLastRow() > 1) {
      const cEnc = hojaCostos.getRange(1, 1, 1, hojaCostos.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const cData = hojaCostos.getRange(2, 1, hojaCostos.getLastRow() - 1, hojaCostos.getLastColumn()).getValues();
      cData.forEach(c => {
        const itemObj = SEG_CONVERTIR_FILA_OBJETO(cEnc, c);
        if (String(itemObj.ID_OBRA).trim() === String(idObra).trim() && itemObj.ESTADO !== "ANULADO") {
          totalCostosDirectos += Number(itemObj.VALOR || 0);
          listaCostos.push(itemObj);
        }
      });
    }
    
    // 3. Gastos asignados a la obra
    const hojaGastos = ss.getSheetByName("GAS_MOVIMIENTOS");
    if (hojaGastos && hojaGastos.getLastRow() > 1) {
      const gData = hojaGastos.getRange(2, 1, hojaGastos.getLastRow() - 1, 16).getValues();
      gData.forEach(g => {
        if (String(g[5]).trim() === String(idObra).trim() && g[14] !== "ANULADO") {
          totalGastosObra += Number(g[7] || 0);
        }
      });
    }
    
    const costoTotalEjecutado = totalCostosDirectos + totalGastosObra;
    const utilidadEstimada = totalIngresos - costoTotalEjecutado;
    const margenPorcentaje = totalIngresos > 0 ? ((utilidadEstimada / totalIngresos) * 100).toFixed(2) : 0;
    
    return {
      EXITO: true,
      DATOS: {
        ID_OBRA: idObra,
        TOTAL_INGRESOS: totalIngresos,
        TOTAL_COSTOS_DIRECTOS: totalCostosDirectos,
        TOTAL_GASTOS: totalGastosObra,
        COSTO_TOTAL: costoTotalEjecutado,
        UTILIDAD_ESTIMADA: utilidadEstimada,
        MARGEN_PORCENTAJE: margenPorcentaje,
        HISTORIAL_COSTOS: SEG_SANITIZAR_PARA_CLIENTE(listaCostos)
      },
      MENSAJE: "Detalle de obra y liquidación de rentabilidad calculado exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: "Error al consolidar detalle de obra: " + error.toString() };
  }
}

