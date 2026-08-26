/**************************************************************
* 21_CALCULOS.gs
* RESPONSABILIDAD:
* - Realizar cálculos tributarios y financieros de Colombia (bases, IVA, retenciones, ICA).
* - Asegurar consistencia de decimales y redondeos del sistema.
**************************************************************/

const CALC_CONFIG = {
  IVA_ESTANDAR: 19.0,
  RETEN_ESTANDAR: 2.5,
  DECIMALES: 2
};

/**
 * Realiza el cálculo matemático de impuestos locales (IVA, Retención).
 * @param {number} base Base gravable sobre la cual calcular.
 * @param {number} tarifaIva Tarifa del IVA (ej: 19.0).
 * @param {number} tarifaReten Tarifa de Retención en la fuente (ej: 2.5).
 * @returns {object} Resultados detallados con redondeo estructurado.
 */
function CALC_CALCULAR_IMPUESTOS_COLOMBIA(base, tarifaIva, tarifaReten) {
  const baseNum = Number(base || 0);
  const ivaPct = Number(tarifaIva !== undefined ? tarifaIva : CALC_CONFIG.IVA_ESTANDAR) / 100;
  const retPct = Number(tarifaReten !== undefined ? tarifaReten : CALC_CONFIG.RETEN_ESTANDAR) / 100;

  const calculoIva = baseNum * ivaPct;
  const calculoRet = baseNum * retPct;
  const total = baseNum + calculoIva - calculoRet;

  const d = CALC_CONFIG.DECIMALES;
  return {
    base: Number(baseNum.toFixed(d)),
    iva: Number(calculoIva.toFixed(d)),
    retencion: Number(calculoRet.toFixed(d)),
    total: Number(total.toFixed(d))
  };
}