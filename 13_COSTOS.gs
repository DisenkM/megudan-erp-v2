// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 13_COSTOS.gs
* RESPONSABILIDAD:
* - Registrar y clasificar los movimientos de Costos Operativos asociados a Obras.
* - Generar alertas automáticas de desviación de presupuesto.
* - Exponer métodos seguros e interactivos para la Web App (SPA) mediante RPC.
**************************************************************/

const COS_CONFIG = {
  HOJA_MOVIMIENTOS: "COS_MOVIMIENTOS",
  HOJA_CATEGORIAS: "COS_CATEGORIAS",
  PREFIJO_ID: "COS",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar costos directos e indirectos, uniendo con PROV_MAESTRO y OBR_MAESTRO
 */
function COS_LISTAR_COSTOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(COS_CONFIG.HOJA_MOVIMIENTOS);
    if (!hoja) return { EXITO: false, DATOS: [], MENSAJE: "Hoja de costos no encontrada." };
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No se registran movimientos de costos." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    // Unir con PROV_MAESTRO
    const hojaProv = ss.getSheetByName("PROV_MAESTRO");
    const mapaProv = {};
    if (hojaProv) {
      const provEnc = hojaProv.getRange(1, 1, 1, hojaProv.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const provRegs = hojaProv.getLastRow() > 1 ? hojaProv.getRange(2, 1, hojaProv.getLastRow() - 1, hojaProv.getLastColumn()).getValues() : [];
      provRegs.forEach(row => {
        const p = SEG_CONVERTIR_FILA_OBJETO(provEnc, row);
        mapaProv[p.ID_PROVEEDOR] = p.RAZON_SOCIAL || p.NOMBRE_COMERCIAL || p.NOMBRE_CONTACTO;
      });
    }
    
    // Unir con OBR_MAESTRO
    const hojaObr = ss.getSheetByName("OBR_MAESTRO");
    const mapaObras = {};
    if (hojaObr) {
      const obrEnc = hojaObr.getRange(1, 1, 1, hojaObr.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const obrRegs = hojaObr.getLastRow() > 1 ? hojaObr.getRange(2, 1, hojaObr.getLastRow() - 1, hojaObr.getLastColumn()).getValues() : [];
      obrRegs.forEach(row => {
        const o = SEG_CONVERTIR_FILA_OBJETO(obrEnc, row);
        mapaObras[o.ID_OBRA] = o.NOMBRE_OBRA || o.CODIGO_OBRA;
      });
    }
    
    lista.forEach(item => {
      item.RAZON_SOCIAL_PROVEEDOR = mapaProv[item.ID_PROVEEDOR] || "Otro / Sin Proveedor (" + item.ID_PROVEEDOR + ")";
      item.NOMBRE_OBRA = mapaObras[item.ID_OBRA] || "Gasto General / Sin Obra (" + item.ID_OBRA + ")";
    });
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de costos obtenida exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar costos: " + error.toString() };
  }
}

/**
 * RPC: Registrar un nuevo movimiento de costo directo asociado a un proyecto / obra
 */
function COS_REGISTRAR_COSTO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "INVENTARIO", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.VALOR || datos.VALOR <= 0 || !datos.ID_CUENTA) {
      throw new Error("Datos de costo incompletos.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMov = ss.getSheetByName(COS_CONFIG.HOJA_MOVIMIENTOS);
    if (!hojaMov) throw new Error("Hoja COS_MOVIMIENTOS no encontrada.");
    
    const idCosto = COS_CONFIG.PREFIJO_ID + "-" + String(Math.max(1, hojaMov.getLastRow())).padStart(COS_CONFIG.DIGITOS_ID, "0");
    const ahora = new Date();
    
    // Columnas: ID_COSTO, FECHA, TIPO_COSTO, CATEGORIA, ID_OBRA, ID_PRODUCTO, ID_PROVEEDOR, DOCUMENTO_ORIGEN, ID_ORIGEN, VALOR, CENTRO_COSTO, CUENTA_CONTABLE, RESPONSABLE, ESTADO, OBSERVACION
    hojaMov.appendRow([
      idCosto,
      ahora,
      datos.TIPO_COSTO || "DIRECTO",
      datos.CATEGORIA || "Materiales",
      datos.ID_OBRA || "OBRA-GENERAL",
      datos.ID_PRODUCTO || "PRD-GENERAL",
      datos.ID_PROVEEDOR || "PROV-GENERAL",
      datos.DOCUMENTO_ORIGEN || "MANUAL",
      datos.ID_ORIGEN || idCosto,
      Number(datos.VALOR),
      datos.CENTRO_COSTO || "OBRAS",
      datos.CUENTA_CONTABLE || "519505",
      usuarioEjecutor,
      "PROCESADO",
      datos.OBSERVACION || "Costo operativo de obra registrado manualmente"
    ]);
    
    // Afectar la cuenta en Tesorería (EGRESO)
    try {
      const hojaCuentas = ss.getSheetByName("TES_CUENTAS");
      const hojaTesMov = ss.getSheetByName("TES_MOVIMIENTOS");
      if (hojaCuentas && hojaTesMov) {
        let nuevoSaldo = 0;
        const cuentasData = hojaCuentas.getRange(2, 1, Math.max(1, hojaCuentas.getLastRow() - 1), 10).getValues();
        const cIndex = cuentasData.findIndex(f => String(f[0]).trim() === String(datos.ID_CUENTA).trim());
        if (cIndex !== -1) {
          const saldoAnterior = Number(cuentasData[cIndex][6] || 0);
          nuevoSaldo = saldoAnterior - Number(datos.VALOR);
          hojaCuentas.getRange(cIndex + 2, 7).setValue(nuevoSaldo); // Actualizar saldo de cuenta
        }
        
        const idTes = "MOV-" + String(Math.max(1, hojaTesMov.getLastRow())).padStart(6, "0");
        hojaTesMov.appendRow([
          idTes,
          ahora,
          "EGRESO",
          datos.ID_CUENTA,
          "COSTOS",
          idCosto,
          0,
          Number(datos.VALOR),
          nuevoSaldo,
          "TRANSFERENCIA",
          usuarioEjecutor,
          datos.OBSERVACION || "Costo de obra " + idCosto
        ]);
      }
    } catch (errTes) {
      console.warn("No se pudo afectar el libro de tesorería: " + errTes.toString());
    }
    
    // Registrar Auditoría
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "INVENTARIO",
      SUBMODULO: "COSTOS",
      ACCION: "CREAR",
      TIPO_REGISTRO: "COS_MOVIMIENTOS",
      ID_REGISTRO: idCosto,
      DESCRIPCION: "Costo directo registrado por valor de: " + datos.VALOR + " COP asociado a la obra: " + datos.ID_OBRA,
      RESULTADO: "EXITOSO"
    });
    
    return {
      EXITO: true,
      ID_COSTO: idCosto,
      MENSAJE: "¡Costo operativo registrado exitosamente! Consecutivo: " + idCosto
    };
  } catch (error) {
    if (typeof LOG_REGISTRAR_ERROR === "function") {
      LOG_REGISTRAR_ERROR("COS_REGISTRAR_COSTO_WEB", "INVENTARIO", error);
    }
    return { EXITO: false, MENSAJE: "Error al registrar costo: " + error.toString() };
  }
}