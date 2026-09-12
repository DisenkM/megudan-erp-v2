/**************************************************************
* 19_TRIBUTARIO.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 2)
* RESPONSABILIDAD:
* - Administrar las configuraciones de impuestos y retenciones de MEGUDAN.
* - Calcular en caliente bases gravables, IVA bajo cláusula de AIU e IVA tradicional.
* - Estructurar el calendario tributario nacional y distrital para 2026 (NIT 901915723-2).
* - Exponer métodos seguros de auditoría fiscal para la Web App (SPA) mediante RPC.
**************************************************************/

const TAX_CONFIG = {
  HOJA_IMPUESTOS: "CFG_IMPUESTOS",
  NIT_EMPRESA: "901915723",
  DV_EMPRESA: "2",
  TARIFA_ICA_PITALITO: 0.005, // 5 x mil (0.5% construcción/servicios)
  SOBRETASA_BOMBERIL: 0.04,   // 4% sobre impuesto de ICA
  AVISOS_Y_TABLEROS: 0.15     // 15% sobre impuesto de ICA
};

/**
 * RPC: Obtener el Calendario Tributario personalizado de MEGUDAN para el año 2026
 * Parametrizado bajo NIT terminado en 23 (último dígito 3) y domicilio en Pitalito, Huila.
 */
function TAX_OBTENER_CALENDARIO_TRIBUTARIO_WEB(tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    
    const calendario = [
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Enero 2026",
        vencimiento: "2026-02-12",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas a título de Renta, IVA y Timbre."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Febrero 2026",
        vencimiento: "2026-03-11",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas a título de Renta, IVA y Timbre."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Marzo 2026",
        vencimiento: "2026-04-14",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas a título de Renta, IVA y Timbre."
      },
      {
        impuesto: "Impuesto al Patrimonio Extraordinario",
        formulario: "Formulario 420",
        periodo: "Declaración / Cuota 1",
        vencimiento: "2026-04-22",
        entidad: "DIAN",
        descripcion: "Impuesto extraordinario por emergencia climática (Decreto 0173 de 2026). Tarifa del 0.5%."
      },
      {
        impuesto: "Industria y Comercio (ICA)",
        formulario: "Declaración Anual",
        periodo: "Año Gravable 2025",
        vencimiento: "2026-04-30",
        entidad: "Alcaldía de Pitalito",
        descripcion: "Declaración y pago anual del impuesto municipal de Industria y Comercio en Pitalito."
      },
      {
        impuesto: "Impuesto sobre la Renta (S.A.S.)",
        formulario: "Formulario 110",
        periodo: "Año Gravable 2025 (Declaración / Cuota 1)",
        vencimiento: "2026-05-14",
        entidad: "DIAN",
        descripcion: "Declaración anual de renta de personas jurídicas. Pago de la primera cuota (50%). Tarifa 35%."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Abril 2026",
        vencimiento: "2026-05-12",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas a título de Renta, IVA y Timbre."
      },
      {
        impuesto: "Información Exógena",
        formulario: "Medios Magnéticos",
        periodo: "Año Gravable 2025",
        vencimiento: "2026-05-21",
        entidad: "DIAN",
        descripcion: "Presentación de reportes de terceros, clientes, proveedores, costos y retenciones anuales."
      },
      {
        impuesto: "Impuesto al Patrimonio Extraordinario",
        formulario: "Formulario 420",
        periodo: "Pago Cuota 2",
        vencimiento: "2026-05-22",
        entidad: "DIAN",
        descripcion: "Pago de la segunda cuota del Impuesto extraordinario por emergencia (Decreto 0173)."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Mayo 2026",
        vencimiento: "2026-06-11",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas a título de Renta, IVA y Timbre."
      },
      {
        impuesto: "Impuesto sobre la Renta (S.A.S.)",
        formulario: "Formulario 110",
        periodo: "Año Gravable 2025 (Cuota 2)",
        vencimiento: "2026-07-10",
        entidad: "DIAN",
        descripcion: "Pago de la segunda cuota de renta (saldo restante del 50%)."
      },
      {
        impuesto: "Impuesto sobre las Ventas (I.V.A.)",
        formulario: "Formulario 300",
        periodo: "Enero-Abril 2026 (Cuatrimestre 1)",
        vencimiento: "2026-05-13",
        entidad: "DIAN",
        descripcion: "Declaración cuatrimestral de IVA cobrado en ventas (sujeto a deducciones bajo AIU de Obras)."
      },
      {
        impuesto: "Impuesto sobre las Ventas (I.V.A.)",
        formulario: "Formulario 300",
        periodo: "Mayo-Agosto 2026 (Cuatrimestre 2)",
        vencimiento: "2026-09-11",
        entidad: "DIAN",
        descripcion: "Declaración cuatrimestral de IVA cobrado en ventas (sujeto a deducciones bajo AIU de Obras)."
      },
      {
        impuesto: "Impuesto sobre las Ventas (I.V.A.)",
        formulario: "Formulario 300",
        periodo: "Septiembre-Diciembre 2026 (Cuatrimestre 3)",
        vencimiento: "2027-01-13",
        entidad: "DIAN",
        descripcion: "Declaración cuatrimestral de IVA cobrado en ventas (sujeto a deducciones bajo AIU de Obras)."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Junio 2026",
        vencimiento: "2026-07-14",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Julio 2026",
        vencimiento: "2026-08-12",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Agosto 2026",
        vencimiento: "2026-09-11",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Septiembre 2026",
        vencimiento: "2026-10-13",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Octubre 2026",
        vencimiento: "2026-11-12",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Noviembre 2026",
        vencimiento: "2026-12-11",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      },
      {
        impuesto: "Retención en la Fuente",
        formulario: "Formulario 350",
        periodo: "Diciembre 2026",
        vencimiento: "2027-01-13",
        entidad: "DIAN",
        descripcion: "Declaración mensual de retenciones practicadas."
      }
    ];
    
    return { EXITO: true, DATOS: calendario, MENSAJE: "Calendario tributario unificado obtenido para el NIT: " + TAX_CONFIG.NIT_EMPRESA };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al cargar calendario: " + error.toString() };
  }
}

/**
 * RPC: Calcular de forma consolidada e interactiva la liquidación de impuestos de un periodo específico
 * Cruza las hojas físicas de compras, ventas y egresos para pre-calcular bases gravables e IVA.
 */
function TAX_CALCULAR_RESUMEN_TRIBUTARIO_WEB(periodo, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "TESORERIA", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let totalVentasBrutas = 0;
    let totalIvaGenerado = 0;
    let totalDescuentosOtorgados = 0;
    let retencionesA_Favor = 0;
    
    let totalComprasBrutas = 0;
    let totalIvaDescontableCompras = 0;
    let retencionesPracticadasCompras = 0;
    
    let totalGastosBrutos = 0;
    let totalIvaDescontableGastos = 0;
    let retencionesPracticadasGastos = 0;
    
    const periodoNormalizado = String(periodo || "").trim(); // Esperado "YYYY-MM"
    
    // 1. Leer Ventas (VEN_CABECERA)
    const hojaVen = ss.getSheetByName("VEN_CABECERA");
    if (hojaVen && hojaVen.getLastRow() >= 2) {
      const enc = hojaVen.getRange(1, 1, 1, hojaVen.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const idxFecha = enc.indexOf("FECHA");
      const idxSub = enc.indexOf("SUBTOTAL");
      const idxDesc = enc.indexOf("DESCUENTO");
      const idxIva = enc.indexOf("IVA");
      const idxTotal = enc.indexOf("TOTAL");
      const idxEstado = enc.indexOf("ESTADO");
      
      const vRegs = hojaVen.getRange(2, 1, hojaVen.getLastRow() - 1, enc.length).getValues();
      vRegs.forEach(row => {
        if (row[idxEstado] === "ANULADA") return;
        const fechaStr = row[idxFecha] instanceof Date ? Utilities.formatDate(row[idxFecha], "America/Bogota", "yyyy-MM") : String(row[idxFecha]);
        if (fechaStr.substring(0, 7) === periodoNormalizado) {
          totalVentasBrutas += Number(row[idxSub] || 0);
          totalIvaGenerado += Number(row[idxIva] || 0);
          totalDescuentosOtorgados += Number(row[idxDesc] || 0);
          
          // Estimar retenciones recibidas si es crédito (tarifa promedio de 2% en construcción)
          if (row[enc.indexOf("FORMA_PAGO")] === "CREDITO") {
            retencionesA_Favor += Number(row[idxSub] || 0) * 0.02;
          }
        }
      });
    }
    
    // 2. Leer Compras (COM_CABECERA)
    const hojaCom = ss.getSheetByName("COM_CABECERA");
    if (hojaCom && hojaCom.getLastRow() >= 2) {
      const enc = hojaCom.getRange(1, 1, 1, hojaCom.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const idxFecha = enc.indexOf("FECHA");
      const idxSub = enc.indexOf("SUBTOTAL");
      const idxIva = enc.indexOf("IVA");
      const idxDesc = enc.indexOf("DESCUENTO");
      const idxEstado = enc.indexOf("ESTADO");
      
      const cRegs = hojaCom.getRange(2, 1, hojaCom.getLastRow() - 1, enc.length).getValues();
      cRegs.forEach(row => {
        if (row[idxEstado] === "ANULADA") return;
        const fechaStr = row[idxFecha] instanceof Date ? Utilities.formatDate(row[idxFecha], "America/Bogota", "yyyy-MM") : String(row[idxFecha]);
        if (fechaStr.substring(0, 7) === periodoNormalizado) {
          totalComprasBrutas += Number(row[idxSub] || 0);
          totalIvaDescontableCompras += Number(row[idxIva] || 0);
          
          // Estimar retenciones en la fuente aplicadas (compras promedio 2.5% o servicios 4%)
          retencionesPracticadasCompras += Number(row[idxSub] || 0) * 0.025;
        }
      });
    }
    
    // 3. Leer Gastos (GAS_MOVIMIENTOS)
    const hojaGas = ss.getSheetByName("GAS_MOVIMIENTOS");
    if (hojaGas && hojaGas.getLastRow() >= 2) {
      const enc = hojaGas.getRange(1, 1, 1, hojaGas.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
      const idxFecha = enc.indexOf("FECHA");
      const idxVal = enc.indexOf("VALOR");
      const idxIva = enc.indexOf("IVA");
      const idxEstado = enc.indexOf("ESTADO");
      
      const gRegs = hojaGas.getRange(2, 1, hojaGas.getLastRow() - 1, enc.length).getValues();
      gRegs.forEach(row => {
        if (row[idxEstado] === "ANULADO") return;
        const fechaStr = row[idxFecha] instanceof Date ? Utilities.formatDate(row[idxFecha], "America/Bogota", "yyyy-MM") : String(row[idxFecha]);
        if (fechaStr.substring(0, 7) === periodoNormalizado) {
          totalGastosBrutos += Number(row[idxVal] || 0);
          totalIvaDescontableGastos += Number(row[idxIva] || 0);
          
          // Estimar retenciones sobre gastos fijos/servicios
          retencionesPracticadasGastos += Number(row[idxVal] || 0) * 0.04;
        }
      });
    }
    
    // 4. Cálculos Consolidados
    const baseIvaNeto = totalIvaGenerado - (totalIvaDescontableCompras + totalIvaDescontableGastos);
    const retencionesPorPagar = retencionesPracticadasCompras + retencionesPracticadasGastos;
    
    // Calcular ICA de Pitalito: Ingresos Brutos de Ventas * Tarifa Base
    const impuestoIcaBase = totalVentasBrutas * TAX_CONFIG.TARIFA_ICA_PITALITO;
    const bomberos = impuestoIcaBase * TAX_CONFIG.SOBRETASA_BOMBERIL;
    const avisos = impuestoIcaBase * TAX_CONFIG.AVISOS_Y_TABLEROS;
    const icaConsolidadoPitalito = impuestoIcaBase + bomberos + avisos;
    
    // Estimar Renta (35%) sobre utilidad operativa aproximada
    const utilidadEstimada = totalVentasBrutas - totalDescuentosOtorgados - totalComprasBrutas - totalGastosBrutos;
    const provisionRenta = utilidadEstimada > 0 ? (utilidadEstimada * 0.35) : 0;
    
    const datosConsolidados = {
      ventas: totalVentasBrutas,
      ventasDescuentos: totalDescuentosOtorgados,
      ivaGenerado: totalIvaGenerado,
      ivaDescontable: totalIvaDescontableCompras + totalIvaDescontableGastos,
      ivaNetoPagar: baseIvaNeto,
      retencionesAFavor: retencionesA_Favor,
      retencionesPracticadas: retencionesPorPagar,
      icaBase: impuestoIcaBase,
      icaAvisos: avisos,
      icaBomberos: bomberos,
      icaTotalPagar: icaConsolidadoPitalito,
      utilidadOperativa: utilidadEstimada,
      rentaProvision: provisionRenta
    };
    
    return {
      EXITO: true,
      DATOS: SEG_SANITIZAR_PARA_CLIENTE(datosConsolidados),
      MENSAJE: "Resumen tributario precalculado con éxito para el periodo " + periodoNormalizado
    };
  } catch (error) {
    return { EXITO: false, DATOS: null, MENSAJE: "Error en cálculo tributario: " + error.toString() };
  }
}