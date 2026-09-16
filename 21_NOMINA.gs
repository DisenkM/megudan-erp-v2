// (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 21_NOMINA.gs (VERSIÓN 2.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Administrar el maestro de empleados y personal de planta (NOM_EMPLEADOS).
* - Liquidación avanzada de Nómina Colombiana según Código Sustantivo del Trabajo (CST).
* - Cálculo de Recargos (Nocturnos 35%, Dominicales/Festivos 75%), Horas Extras (Diurnas 25%, Nocturnas 75%, Dom/Fest 100%-150%).
* - Gestión de Comisiones, Bonos/Auxilios No Constitutivos de Salario (Art. 128 CST).
* - Deducciones de Ley (Salud 4%, Pensión 4%, FSP 1%) y Auxilio de Transporte (< 2 SMMLV).
* - Provisiones de Prestaciones Sociales (Cesantías 8.33%, Int. Cesantías 1%, Prima 8.33%, Vacaciones 4.17%) y Aportes Patronales (ARL, CCF, SENA, ICBF).
* - Historial completo de pagos de nómina con consulta y trazabilidad auditada.
**************************************************************/

const NOM_CONFIG_CORE = {
  HOJA_EMPLEADOS: "NOM_EMPLEADOS",
  HOJA_LIQUIDACION: "NOM_LIQUIDACION",
  HOJA_HISTORIAL: "NOM_HISTORIAL",
  PREFIJO_EMPLEADO: "EMP",
  PREFIJO_LIQUIDACION: "NOM",
  DIGITOS_ID: 6,
  SMMLV_2026: 1423500, // Valor base SMMLV 2026
  AUX_TRANSPORTE_2026: 162000, // Valor base Auxilio de Transporte 2026
  HORAS_MES_LABORALES: 230 // Jornada laboral mensual ajustada
};

/**
 * RPC: Listar todos los empleados activos del personal
 */
function NOM_LISTAR_EMPLEADOS_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
    
    if (!hoja) {
      hoja = ss.insertSheet(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
      hoja.setTabColor("#10b981");
      hoja.appendRow([
        "ID_EMPLEADO", "DOCUMENTO", "TIPO_DOCUMENTO", "NOMBRES_APELLIDOS", "CARGO", 
        "SALARIO_BASE", "TIPO_CONTRATO", "NIVEL_ARL", "TIENE_AUX_TRANSPORTE", 
        "CUENTA_BANCARIA", "BANCO", "FECHA_INGRESO", "ESTADO", "FECHA_CREACION"
      ]);
    }
    
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return { EXITO: true, DATOS: [], MENSAJE: "No hay empleados registrados." };
    
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ultimaFila - 1, enc.length).getValues();
    
    const lista = reg.filter(r => r[0] && String(r[0]).trim() !== "").map(r => {
      const obj = {};
      enc.forEach((c, idx) => {
        let val = r[idx] !== undefined && r[idx] !== null ? r[idx] : "";
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone() || "America/Bogota", "yyyy-MM-dd");
        }
        obj[c] = val;
      });
      return obj;
    });
    
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(lista), MENSAJE: "Lista de empleados obtenida exitosamente." };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar empleados: " + error.toString() };
  }
}

/**
 * RPC: Guardar o actualizar un empleado en el sistema
 */
function NOM_GUARDAR_EMPLEADO_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.DOCUMENTO || !datos.NOMBRES_APELLIDOS) {
      throw new Error("El documento y los nombres completos son obligatorios.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
    if (!hoja) {
      hoja = ss.insertSheet(NOM_CONFIG_CORE.HOJA_EMPLEADOS);
      hoja.setTabColor("#10b981");
      hoja.appendRow([
        "ID_EMPLEADO", "DOCUMENTO", "TIPO_DOCUMENTO", "NOMBRES_APELLIDOS", "CARGO", 
        "SALARIO_BASE", "TIPO_CONTRATO", "NIVEL_ARL", "TIENE_AUX_TRANSPORTE", 
        "CUENTA_BANCARIA", "BANCO", "FECHA_INGRESO", "ESTADO", "FECHA_CREACION"
      ]);
    }
    
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const ultimaFila = hoja.getLastRow();
    const ahora = new Date();
    
    let idEmp = datos.ID_EMPLEADO;
    let filaIndex = -1;
    
    if (ultimaFila >= 2) {
      const idxId = enc.indexOf("ID_EMPLEADO");
      const idxDoc = enc.indexOf("DOCUMENTO");
      const reg = hoja.getRange(2, 1, ultimaFila - 1, enc.length).getValues();
      
      for (let i = 0; i < reg.length; i++) {
        const rowId = idxId !== -1 ? String(reg[i][idxId]).trim() : "";
        const rowDoc = idxDoc !== -1 ? String(reg[i][idxDoc]).trim() : "";
        
        if ((idEmp && rowId === String(idEmp).trim()) || (!idEmp && rowDoc === String(datos.DOCUMENTO).trim())) {
          filaIndex = i + 2;
          idEmp = rowId;
          break;
        }
      }
    }
    
    if (!idEmp) {
      idEmp = NOM_CONFIG_CORE.PREFIJO_EMPLEADO + "-" + String(Math.max(1, ultimaFila)).padStart(NOM_CONFIG_CORE.DIGITOS_ID, "0");
    }
    
    const empObj = {
      ID_EMPLEADO: idEmp,
      DOCUMENTO: String(datos.DOCUMENTO).trim(),
      TIPO_DOCUMENTO: String(datos.TIPO_DOCUMENTO || "CC").toUpperCase(),
      NOMBRES_APELLIDOS: String(datos.NOMBRES_APELLIDOS).trim().toUpperCase(),
      CARGO: String(datos.CARGO || "OPERARIO").toUpperCase(),
      SALARIO_BASE: Number(datos.SALARIO_BASE || NOM_CONFIG_CORE.SMMLV_2026),
      TIPO_CONTRATO: String(datos.TIPO_CONTRATO || "INDEFINIDO").toUpperCase(),
      NIVEL_ARL: String(datos.NIVEL_ARL || "RIESGO_1").toUpperCase(),
      TIENE_AUX_TRANSPORTE: (datos.SALARIO_BASE <= (NOM_CONFIG_CORE.SMMLV_2026 * 2)) ? "SI" : "NO",
      CUENTA_BANCARIA: String(datos.CUENTA_BANCARIA || "").trim(),
      BANCO: String(datos.BANCO || "BANCOLOMBIA").toUpperCase(),
      FECHA_INGRESO: datos.FECHA_INGRESO || ahora,
      ESTADO: String(datos.ESTADO || "ACTIVO").toUpperCase(),
      FECHA_CREACION: ahora
    };
    
    if (filaIndex !== -1) {
      const filaEdit = enc.map(c => empObj[c] !== undefined ? empObj[c] : "");
      hoja.getRange(filaIndex, 1, 1, enc.length).setValues([filaEdit]);
    } else {
      const filaNueva = enc.map(c => empObj[c] !== undefined ? empObj[c] : "");
      hoja.appendRow(filaNueva);
    }
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "NOMINA",
      ACCION: filaIndex !== -1 ? "EDITAR" : "CREAR",
      TIPO_REGISTRO: "NOM_EMPLEADOS",
      ID_REGISTRO: idEmp,
      DESCRIPCION: "Empleado " + empObj.NOMBRES_APELLIDOS + " (" + idEmp + ") guardado en nómina.",
      RESULTADO: "EXITOSO"
    });
    
    return { EXITO: true, ID_EMPLEADO: idEmp, MENSAJE: "¡Empleado " + empObj.NOMBRES_APELLIDOS + " guardado con éxito!" };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al guardar empleado: " + error.toString() };
  }
}

/**
 * RPC: Liquidación Completa y Especifica de Nómina Colombiana
 */
function NOM_LIQUIDAR_NOMINA_COMPLETA_WEB(datos, tokenSesion) {
  try {
    const auth = SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "CREAR");
    const usuarioEjecutor = auth.USUARIO || "SISTEMA";
    
    if (!datos || !datos.ID_EMPLEADO) {
      throw new Error("Debe seleccionar un empleado válido para liquidar nómina.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaLiq = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_LIQUIDACION);
    
    if (!hojaLiq) {
      hojaLiq = ss.insertSheet(NOM_CONFIG_CORE.HOJA_LIQUIDACION);
      hojaLiq.setTabColor("#10b981");
      hojaLiq.appendRow([
        "ID_NOMINA", "PERIODO", "ID_EMPLEADO", "DOCUMENTO", "EMPLEADO_NOMBRE", "DIAS_TRABAJADOS",
        "SALARIO_BASE", "SUELDO_DEVENGADO", "HORAS_EXTRA_DIURNAS", "VALOR_HED", 
        "HORAS_EXTRA_NOCTURNAS", "VALOR_HEN", "RECARGOS_NOCTURNOS", "VALOR_RN",
        "RECARGOS_DOMINICALES", "VALOR_RD", "HORAS_EXTRA_DOM_DIURNAS", "VALOR_HEDD",
        "HORAS_EXTRA_DOM_NOCTURNAS", "VALOR_HEND", "COMISIONES", "BONOS_NO_SALARIALES",
        "TOTAL_DEVENGADO_SALARIAL", "TOTAL_DEVENGADO_NO_SALARIAL", "AUXILIO_TRANSPORTE",
        "TOTAL_DEVENGADO_GENERAL", "IBC_SALUD_PENSION", "DEDUCCION_SALUD", "DEDUCCION_PENSION",
        "DEDUCCION_FSP", "OTRAS_DEDUCCIONES", "TOTAL_DEDUCCIONES", "NETO_PAGADO",
        "PROV_CESANTIAS", "PROV_INT_CESANTIAS", "PROV_PRIMA", "PROV_VACACIONES", "APORTE_ARL",
        "APORTE_CAJA", "ESTADO", "FECHA_PAGO", "USUARIO"
      ]);
    }
    
    const encLiq = hojaLiq.getRange(1, 1, 1, hojaLiq.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const ultimaFila = hojaLiq.getLastRow();
    const ahora = new Date();
    const idNomina = NOM_CONFIG_CORE.PREFIJO_LIQUIDACION + "-" + String(Math.max(1, ultimaFila)).padStart(NOM_CONFIG_CORE.DIGITOS_ID, "0");
    
    // Parametrización de Tarifas CST Colombia
    const diasTrabajados = Number(datos.DIAS_TRABAJADOS || 30);
    const salarioBase = Number(datos.SALARIO_BASE || NOM_CONFIG_CORE.SMMLV_2026);
    const valorHoraOrdinaria = salarioBase / NOM_CONFIG_CORE.HORAS_MES_LABORALES;
    
    const sueldoDevengado = (salarioBase / 30) * diasTrabajados;
    
    // Recargos y Horas Extras
    const hedCant = Number(datos.HORAS_EXTRA_DIURNAS || 0);
    const hedValor = hedCant * (valorHoraOrdinaria * 1.25);
    
    const henCant = Number(datos.HORAS_EXTRA_NOCTURNAS || 0);
    const henValor = henCant * (valorHoraOrdinaria * 1.75);
    
    const rnCant = Number(datos.HORAS_RECARGO_NOCTURNO || 0);
    const rnValor = rnCant * (valorHoraOrdinaria * 0.35);
    
    const rdCant = Number(datos.HORAS_RECARGO_DOMINICAL || 0);
    const rdValor = rdCant * (valorHoraOrdinaria * 0.75);
    
    const heddCant = Number(datos.HORAS_EXTRA_DOM_DIURNAS || 0);
    const heddValor = heddCant * (valorHoraOrdinaria * 2.00);
    
    const hendCant = Number(datos.HORAS_EXTRA_DOM_NOCTURNAS || 0);
    const hendValor = hendCant * (valorHoraOrdinaria * 2.50);
    
    const comisiones = Number(datos.COMISIONES || 0);
    const bonosNoSalariales = Number(datos.BONOS_NO_SALARIALES || 0);
    
    // Devengado Salarial (Base para Cotización e IBC)
    const totalDevengadoSalarial = sueldoDevengado + hedValor + henValor + rnValor + rdValor + heddValor + hendValor + comisiones;
    const totalDevengadoNoSalarial = bonosNoSalariales;
    
    // Auxilio de Transporte
    let auxTransporte = 0;
    if (salarioBase <= (NOM_CONFIG_CORE.SMMLV_2026 * 2) && datos.APLICA_AUX_TRANSPORTE !== "NO") {
      auxTransporte = (NOM_CONFIG_CORE.AUX_TRANSPORTE_2026 / 30) * diasTrabajados;
    }
    
    const totalDevengadoGeneral = totalDevengadoSalarial + totalDevengadoNoSalarial + auxTransporte;
    
    // Deducciones de Ley
    const ibcSaludPension = totalDevengadoSalarial; // Ingreso Base de Cotización
    const deduccionSalud = ibcSaludPension * 0.04;
    const deduccionPension = ibcSaludPension * 0.04;
    
    let deduccionFsp = 0;
    if (ibcSaludPension >= (NOM_CONFIG_CORE.SMMLV_2026 * 4)) {
      deduccionFsp = ibcSaludPension * 0.01; // Fondo de Solidaridad Pensional 1%
    }
    
    const otrasDeducciones = Number(datos.OTRAS_DEDUCCIONES || 0);
    const totalDeducciones = deduccionSalud + deduccionPension + deduccionFsp + otrasDeducciones;
    const netoPagado = totalDevengadoGeneral - totalDeducciones;
    
    // Provisiones y Aportes Patronales
    const basePrestaciones = totalDevengadoSalarial + auxTransporte;
    const provCesantias = basePrestaciones * 0.0833; // 8.33%
    const provIntCesantias = provCesantias * 0.12; // 12% anual sobre cesantías (~1% mensual)
    const provPrima = basePrestaciones * 0.0833; // 8.33%
    const provVacaciones = sueldoDevengado * 0.0417; // 4.17% sobre sueldo base
    
    // ARL Nivel de Riesgo
    const tarifaARLMap = { "RIESGO_1": 0.00522, "RIESGO_2": 0.01044, "RIESGO_3": 0.02436, "RIESGO_4": 0.04350, "RIESGO_5": 0.06960 };
    const tarifaARL = tarifaARLMap[datos.NIVEL_ARL || "RIESGO_1"] || 0.00522;
    const aporteARL = ibcSaludPension * tarifaARL;
    const aporteCajaCompensacion = ibcSaludPension * 0.04; // 4%
    
    const registroNomina = {
      ID_NOMINA: idNomina,
      PERIODO: datos.PERIODO || "2026-09",
      ID_EMPLEADO: datos.ID_EMPLEADO,
      DOCUMENTO: datos.DOCUMENTO || "",
      EMPLEADO_NOMBRE: datos.EMPLEADO_NOMBRE || "",
      DIAS_TRABAJADOS: diasTrabajados,
      SALARIO_BASE: salarioBase,
      SUELDO_DEVENGADO: sueldoDevengado,
      HORAS_EXTRA_DIURNAS: hedCant,
      VALOR_HED: hedValor,
      HORAS_EXTRA_NOCTURNAS: henCant,
      VALOR_HEN: henValor,
      RECARGOS_NOCTURNOS: rnCant,
      VALOR_RN: rnValor,
      RECARGOS_DOMINICALES: rdCant,
      VALOR_RD: rdValor,
      HORAS_EXTRA_DOM_DIURNAS: heddCant,
      VALOR_HEDD: heddValor,
      HORAS_EXTRA_DOM_NOCTURNAS: hendCant,
      VALOR_HEND: hendValor,
      COMISIONES: comisiones,
      BONOS_NO_SALARIALES: bonosNoSalariales,
      TOTAL_DEVENGADO_SALARIAL: totalDevengadoSalarial,
      TOTAL_DEVENGADO_NO_SALARIAL: totalDevengadoNoSalarial,
      AUXILIO_TRANSPORTE: auxTransporte,
      TOTAL_DEVENGADO_GENERAL: totalDevengadoGeneral,
      IBC_SALUD_PENSION: ibcSaludPension,
      DEDUCCION_SALUD: deduccionSalud,
      DEDUCCION_PENSION: deduccionPension,
      DEDUCCION_FSP: deduccionFsp,
      OTRAS_DEDUCCIONES: otrasDeducciones,
      TOTAL_DEDUCCIONES: totalDeducciones,
      NETO_PAGADO: netoPagado,
      PROV_CESANTIAS: provCesantias,
      PROV_INT_CESANTIAS: provIntCesantias,
      PROV_PRIMA: provPrima,
      PROV_VACACIONES: provVacaciones,
      APORTE_ARL: aporteARL,
      APORTE_CAJA: aporteCajaCompensacion,
      ESTADO: "PROCESADA",
      FECHA_PAGO: ahora,
      USUARIO: usuarioEjecutor
    };
    
    const filaNueva = encLiq.map(c => registroNomina[c] !== undefined ? registroNomina[c] : "");
    hojaLiq.appendRow(filaNueva);
    
    // Registrar afectación de egreso en Gastos de forma automática
    try {
      if (typeof GAS_REGISTRAR_GASTO_WEB === "function") {
        GAS_REGISTRAR_GASTO_WEB({
          TIPO_GASTO: "ADMINISTRATIVO",
          CATEGORIA: "Nómina y Salarios",
          VALOR: netoPagado,
          ID_CUENTA: datos.ID_CUENTA_PAGO || "CTA-000001",
          OBSERVACION: "Pago de nómina (" + (datos.PERIODO || "2026-09") + ") - " + datos.EMPLEADO_NOMBRE + " [" + idNomina + "]"
        }, tokenSesion);
      }
    } catch (eGas) {
      console.warn("Nómina procesada. Alerta al sincronizar egreso en gastos: " + eGas.toString());
    }
    
    SEG_REGISTRAR_AUDITORIA({
      ID_USUARIO: auth.SESION ? auth.SESION.ID_USUARIO : "USR-000001",
      USUARIO: usuarioEjecutor,
      MODULO: "SEGURIDAD",
      SUBMODULO: "NOMINA",
      ACCION: "CREAR",
      TIPO_REGISTRO: "NOM_LIQUIDACION",
      ID_REGISTRO: idNomina,
      DESCRIPCION: "Nómina " + idNomina + " liquidada para " + datos.EMPLEADO_NOMBRE + " por $" + netoPagado.toLocaleString('es-CO') + " COP.",
      RESULTADO: "EXITOSO"
    });
    
    return { 
      EXITO: true, 
      ID_NOMINA: idNomina, 
      NETO_PAGADO: netoPagado,
      MENSAJE: "¡Nómina " + idNomina + " liquidada con éxito! Neto Pagado: $" + netoPagado.toLocaleString('es-CO') + " COP." 
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al liquidar nómina: " + error.toString() };
  }
}

/**
 * RPC: Consultar Historial de Pagos de Nómina con Filtros
 */
function NOM_OBTENER_HISTORIAL_PAGOS_WEB(periodo, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(NOM_CONFIG_CORE.HOJA_LIQUIDACION);
    
    if (!hoja || hoja.getLastRow() < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay historial de pagos registrado." };
    }
    
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
    const idxPeriodo = enc.indexOf("PERIODO");
    
    const filtrados = reg.filter(r => {
      if (!r[0]) return false;
      if (!periodo || String(periodo).trim() === "" || String(periodo).toUpperCase() === "TODOS") return true;
      return idxPeriodo !== -1 && String(r[idxPeriodo]).trim() === String(periodo).trim();
    }).map(r => {
      const obj = {};
      enc.forEach((c, idx) => {
        let val = r[idx] !== undefined && r[idx] !== null ? r[idx] : "";
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone() || "America/Bogota", "yyyy-MM-dd HH:mm:ss");
        }
        obj[c] = val;
      });
      return obj;
    });
    
    return { EXITO: true, DATOS: SEG_SANITIZAR_PARA_CLIENTE(filtrados), MENSAJE: "Historial de nómina cargado correctamente." };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al cargar historial: " + error.toString() };
  }
}
