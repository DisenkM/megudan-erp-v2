/**************************************************************
* 21_CALCULOS.gs
* RESPONSABILIDAD:
* - Realizar cálculos tributarios de Colombia (bases, IVA, retenciones).
* - Asegurar consistencia de decimales y redondeos del sistema.
**************************************************************/

const CALC_CONFIG = {
  IVA_ESTANDAR: 19.0,
  RETEN_ESTANDAR: 2.5,
  DECIMALES: 2
};

function CALC_CALCULAR_IMPUESTOS_COLOMBIA(base, tarifaIva, tarifaReten) {
  const d = CALC_CONFIG.DECIMALES;
  const tIva = (tarifaIva !== undefined) ? tarifaIva : CALC_CONFIG.IVA_ESTANDAR;
  const tRet = (tarifaReten !== undefined) ? tarifaReten : CALC_CONFIG.RETEN_ESTANDAR;

  const iva = base * (tIva / 100);
  const retencion = base * (tRet / 100);
  const total = base + iva - retencion;

  return {
    base: Number(base.toFixed(d)),
    iva: Number(iva.toFixed(d)),
    retencion: Number(retencion.toFixed(d)),
    total: Number(total.toFixed(d))
  };
}
