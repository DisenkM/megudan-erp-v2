// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 21_NOMINA.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el personal operativo y empleados de planta (NOM_EMPLEADOS).
* - Procesar la liquidación de nómina quincenal/mensual (NOM_LIQUIDACION).
* - Sincronizar egresos con el módulo de Gastos y Tesorería.
**************************************************************/

const NOM_CONFIG_CORE = {
  HOJA_EMPLEADOS: "NOM_EMPLEADOS",
  HOJA_LIQUIDACION: "NOM_LIQUIDACION",
  PREFIJO_EMPLEADO: "EMP",
  PREFIJO_LIQUIDACION: "NOM",
  DIGITOS_ID: 6
};

/**
 * RPC: Listar todos los empleados del personal operativo
 */
function NOM_LISTAR_EMPLEADOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
    if (!hoja) {
      hoja = ss.insertSheet(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
      hoja.setTabColor("#10b981");
      hoja.appendRow(["ID_EMPLEADO", "DOCUMENTO", "NOMBRES_APELLIDOS", "CARGO", "SALARIO_BASE", "TIPO_CONTRATO", "FECHA_INGRESO", "ESTADO", "FECHA_CREACION"]);
    }
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay empleados registrados en nómina." };
    
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const registros = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
    const lista = registros.map(fila => SEG_CONVERTIR_FILA_OBJETO(encabezados, fila));
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista),
      MENSAJE: "Lista de empleados recuperada exitosamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar empleados: " + error.toString() };
  }
}

/**
 * RPC: Guardar/Crear nuevo empleado
 */
function NOM_GUARDAR_EMPLEADO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.DOCUMENTO || !datos.NOMBRES_APELLIDOS) {
      throw new Error("El documento y el nombre completo son campos obligatorios.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
    if (!hoja) {
      hoja = ss.insertSheet(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
      hoja.setTabColor("#10b981");
      hoja.appendRow(["ID_EMPLEADO", "DOCUMENTO", "NOMBRES_APELLIDOS", "CARGO", "SALARIO_BASE", "TIPO_CONTRATO", "FECHA_INGRESO", "ESTADO", "FECHA_CREACION"]);
    }
    
    const idEmp = NOM_CONFIG_CORE.PREFIJO_EMPLEADO + "-" + String(Math.max(1, hoja.getLastRow())).padStart(NOM_CONFIG_CORE.DIGITOS_ID, "0");
    const ahora = new Date();
    
    hoja.appendRow([
      idEmp,
      datos.DOCUMENTO.trim(),
      datos.NOMBRES_APELLIDOS.trim(),
      datos.CARGO || "OPERARIO",
      Number(datos.SALARIO_BASE || 1300000),
      datos.TIPO_CONTRATO || "INDEFINIDO",
      datos.FECHA_INGRESO || ahora,
      "ACTIVO",
      ahora
    ]);
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "NOMINA",
      ACCION: "CREAR",
      TIPO_REGISTRO: "NOM_EMPLEADOS",
      ID_REGISTRO: idEmp,
      DESCRIPCION: "Empleado " + datos.NOMBRES_APELLIDOS + " registrado con éxito.",
      RESULTADO: "EXITOSO"
    });
    
    return { EXITO: true, ID_EMPLEADO: idEmp, MENSAJE: "¡Empleado " + datos.NOMBRES_APELLIDOS + " registrado en nómina!" };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar empleado: " + error.toString() };
  }
}

/**
 * RPC: Liquidar y procesar nómina del período
 */
function NOM_LIQUIDAR_NOMINA_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.ID_EMPLEADO || datos.VALOR_NETO <= 0) {
      throw new Error("Debe seleccionar un empleado y especificar el valor neto devengado.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_LIQUIDACION);
    if (!hoja) {
      hoja = ss.insertSheet(NOM_CONFIG_CORE.HOJA_LIQUIDACION);
      hoja.setTabColor("#10b981");
      hoja.appendRow(["ID_NOMINA", "PERIODO", "ID_EMPLEADO", "DEVENGADO_BASE", "AUX_TRANSPORTE", "DEDUCCIONES", "NETO_PAGADO", "FECHA_PAGO", "USUARIO"]);
    }
    
    const idNomina = NOM_CONFIG_CORE.PREFIJO_LIQUIDACION + "-" + String(Math.max(1, hoja.getLastRow())).padStart(NOM_CONFIG_CORE.DIGITOS_ID, "0");
    const ahora = new Date();
    
    hoja.appendRow([
      idNomina,
      datos.PERIODO || "2026-09",
      datos.ID_EMPLEADO,
      Number(datos.DEVENGADO_BASE || 0),
      Number(datos.AUX_TRANSPORTE || 0),
      Number(datos.DEDUCCIONES || 0),
      Number(datos.VALOR_NETO || 0),
      ahora,
      usuarioEjecutor
    ]);
    
    // Registrar gasto de nómina en GAS_MOVIMIENTOS de forma automatizada
    try {
      GAS_REGISTRAR_GASTO_WEB({
        TIPO_GASTO: "ADMINISTRATIVO",
        CATEGORIA: "Nómina y Salarios",
        VALOR: Number(datos.VALOR_NETO || 0),
        ID_CUENTA: datos.ID_CUENTA || "CTA-000001",
        OBSERVACION: "Pago de nómina (" + datos.PERIODO + ") para empleado " + datos.ID_EMPLEADO
      }, tokenSesion);
    } catch (eGas) {
      console.warn("Nómina liquidada. Alerta en registro automático de gasto: " + eGas.toString());
    }
    
    return { EXITO: true, ID_NOMINA: idNomina, MENSAJE: "¡Nómina " + idNomina + " liquidada y procesada correctamente!" };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al liquidar nómina: " + error.toString() };
  }
}
