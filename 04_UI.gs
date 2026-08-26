/**************************************************************
* 04_UI.gs
* RESPONSABILIDAD:
* - Administrar la identidad visual de Google Sheets.
* - Colorear las pestañas físicas del libro ERP de forma automática según módulo.
**************************************************************/

// ============================================================
// 01. PALETA DE COLORES POR GRUPO MODULAR (ÚNICA FUENTE)
// ============================================================
const COLORES_ERP = {
  CONFIGURACION: "#6B7280",
  CLIENTES: "#2563EB",
  PROVEEDORES: "#7C3AED",
  PRODUCTOS: "#16A34A",
  OBRAS: "#92400E",
  VENTAS: "#1D4ED8",
  COMPRAS: "#EA580C",
  INVENTARIO: "#0891B2",
  INGRESOS: "#0F766E",
  COSTOS: "#CA8A04",
  GASTOS: "#DC2626",
  CARTERA: "#0284C7",
  CUENTAS_POR_PAGAR: "#C026D3",
  TESORERIA: "#4F46E5",
  SEGURIDAD: "#9333EA",
  MENU: "#111827"
};

// ============================================================
// 02. ASOCIACIÓN DE HOJAS FÍSICAS A CADA GRUPO
// ============================================================
const HOJAS_ERP = {
  CONFIGURACION: [
    "CFG_EMPRESA", "CFG_SISTEMA", "CFG_DOCUMENTOS", "CFG_IMPUESTOS", 
    "CFG_CONTABILIDAD", "CFG_INVENTARIO", "CFG_COSTOS"
  ],
  CLIENTES: [
    "CLI_MAESTRO", "CLI_HISTORIAL"
  ],
  PROVEEDORES: [
    "PROV_MAESTRO", "PROV_FORM", "PROV_HISTORIAL"
  ],
  PRODUCTOS: [
    "PROD_MAESTRO", "PROD_CATEGORIAS", "PROD_UNIDADES", "PROD_TIPOS", "PROD_PRECIOS"
  ],
  OBRAS: [
    "OBR_MAESTRO", "OBR_PRESUPUESTO", "OBR_AVANCE", "OBR_RECURSOS"
  ],
  VENTAS: [
    "VEN_CABECERA", "VEN_DETALLE", "VEN_DOCUMENTOS", "VEN_HISTORIAL"
  ],
  COMPRAS: [
    "COM_CABECERA", "COM_DETALLE", "COM_DOCUMENTOS", "COM_HISTORIAL"
  ],
  INVENTARIO: [
    "INV_MOVIMIENTOS", "INV_SALDOS", "INV_KARDEX", "INV_AJUSTES", "INV_TRASLADOS"
  ],
  INGRESOS: [
    "ING_MOVIMIENTOS", "ING_RECAUDOS", "ING_HISTORIAL"
  ],
  COSTOS: [
    "COS_MOVIMIENTOS", "COS_CATEGORIAS", "COS_HISTORIAL"
  ],
  GASTOS: [
    "GAS_MOVIMIENTOS", "GAS_CATEGORIAS", "GAS_HISTORIAL"
  ],
  CARTERA: [
    "CAR_CUENTAS", "CAR_RECAUDOS", "CAR_VENCIMIENTOS", "CAR_HISTORIAL"
  ],
  CUENTAS_POR_PAGAR: [
    "CXP_CUENTAS", "CXP_PAGOS", "CXP_VENCIMIENTOS", "CXP_HISTORIAL"
  ],
  TESORERIA: [
    "TES_CUENTAS", "TES_MOVIMIENTOS", "TES_PAGOS", "TES_RECAUDOS", "TES_CONCILIACION"
  ],
  SEGURIDAD: [
    "USR_USUARIOS", "USR_ROLES", "USR_PERMISOS", "USR_SESIONES", "USR_AUDITORIA"
  ]
};

// ============================================================
// 03. APLICAR COLORES DE PESTAÑAS AUTOMÁTICAMENTE
// ============================================================
function COLOREAR_TODAS_LAS_PESTANAS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let coloreadas = 0;
  let noEncontradas = [];

  Object.entries(HOJAS_ERP).forEach(function([modulo, hojas]) {
    const color = COLORES_ERP[modulo];
    hojas.forEach(function(nombreHoja) {
      const hoja = ss.getSheetByName(nombreHoja);
      if (hoja) {
        hoja.setTabColor(color);
        coloreadas++;
      } else {
        noEncontradas.push(nombreHoja);
      }
    });
  });

  const menu = ss.getSheetByName("🏠MENU");
  if (menu) {
    menu.setTabColor(COLORES_ERP.MENU);
    coloreadas++;
  }

  SpreadsheetApp.flush();
  console.log("Pestañas coloreadas de manera exitosa: " + coloreadas);
  if (noEncontradas.length > 0) {
    console.info("Hojas que no se han creado en el libro: " + noEncontradas.join(", "));
  }
}