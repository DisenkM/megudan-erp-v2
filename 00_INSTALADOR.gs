/**
 * ============================================================
 * ERP V2 — LIBRO 1: ERP OPERATIVO
 * INSTALADOR INICIAL
 * ============================================================
 *
 * CREA:
 * - MENU
 * - CONFIGURACIÓN
 * - CLIENTES
 * - PROVEEDORES
 * - PRODUCTOS
 * - OBRAS Y PROYECTOS
 * - VENTAS
 * - COMPRAS
 * - INVENTARIO
 * - INGRESOS
 * - COSTOS
 * - GASTOS
 * - CARTERA
 * - CUENTAS POR PAGAR
 * - TESORERÍA
 *
 * NO ELIMINA HOJAS EXISTENTES.
 * SI UNA HOJA YA EXISTE, LA CONSERVA.
 * ============================================================
 */

function INSTALAR_ERP_OPERATIVO() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==========================================================
  // CONFIGURACIÓN GENERAL
  // ==========================================================

  const modulos = [

    {
      nombre: "⚙️ CONFIGURACIÓN",
      color: "#6B7280",
      hojas: {

        "CFG_EMPRESA": [
          "ID_EMPRESA",
          "NIT",
          "DV",
          "RAZON_SOCIAL",
          "NOMBRE_COMERCIAL",
          "TIPO_PERSONA",
          "DIRECCION",
          "CIUDAD",
          "DEPARTAMENTO",
          "TELEFONO",
          "EMAIL",
          "REPRESENTANTE_LEGAL",
          "LOGO",
          "ESTADO"
        ],

        "CFG_SISTEMA": [
          "PARAMETRO",
          "VALOR",
          "TIPO_DATO",
          "DESCRIPCION",
          "ESTADO"
        ],

        "CFG_DOCUMENTOS": [
          "ID_DOCUMENTO",
          "TIPO_DOCUMENTO",
          "PREFIJO",
          "CONSECUTIVO_ACTUAL",
          "CONSECUTIVO_INICIAL",
          "CONSECUTIVO_FINAL",
          "RESOLUCION",
          "FECHA_RESOLUCION",
          "FECHA_VENCIMIENTO",
          "ESTADO"
        ],

        "CFG_IMPUESTOS": [
          "ID_IMPUESTO",
          "CODIGO",
          "NOMBRE",
          "TIPO",
          "TARIFA",
          "BASE_MINIMA",
          "CUENTA_GENERADO",
          "CUENTA_DESCONTABLE",
          "ESTADO"
        ],

        "CFG_CONTABILIDAD": [
          "ID_PARAMETRO",
          "CONCEPTO",
          "CUENTA_CONTABLE",
          "CENTRO_COSTO",
          "DESCRIPCION",
          "ESTADO"
        ],

        "CFG_INVENTARIO": [
          "PARAMETRO",
          "VALOR",
          "DESCRIPCION",
          "ESTADO"
        ],

        "CFG_COSTOS": [
          "ID_PARAMETRO",
          "TIPO_COSTO",
          "METODO",
          "CUENTA_CONTABLE",
          "CENTRO_COSTO",
          "ESTADO"
        ],

        "CFG_USUARIOS": [
          "ID_USUARIO",
          "NOMBRE",
          "EMAIL",
          "ROL",
          "ESTADO",
          "FECHA_CREACION",
          "ULTIMO_ACCESO"
        ],

        "CFG_PERMISOS": [
          "ID_PERMISO",
          "ROL",
          "MODULO",
          "ACCION",
          "PERMITIDO",
          "ESTADO"
        ]
      }
    },

    {
      nombre: "👥 CLIENTES",
      color: "#2563EB",
      hojas: {

        "CLI_MAESTRO": [
          "ID_CLIENTE",
          "TIPO_PERSONA",
          "TIPO_DOCUMENTO",
          "NIT_CC",
          "DV",
          "RAZON_SOCIAL",
          "NOMBRE_COMERCIAL",
          "NOMBRE_CONTACTO",
          "TELEFONO",
          "EMAIL",
          "DIRECCION",
          "CIUDAD",
          "DEPARTAMENTO",
          "CONDICION_PAGO",
          "CUPO_CREDITO",
          "ESTADO",
          "ID_SIIGO",
          "ID_ALEGRA",
          "FECHA_CREACION",
          "FECHA_MODIFICACION"
        ],

        "CLI_FORM": [
          "ID_CLIENTE",
          "TIPO_PERSONA",
          "TIPO_DOCUMENTO",
          "NIT_CC",
          "DV",
          "RAZON_SOCIAL",
          "NOMBRE_COMERCIAL",
          "NOMBRE_CONTACTO",
          "TELEFONO",
          "EMAIL",
          "DIRECCION",
          "CIUDAD",
          "DEPARTAMENTO",
          "CONDICION_PAGO",
          "CUPO_CREDITO",
          "OBSERVACION"
        ],

        "CLI_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_CLIENTE",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "CAMPO",
          "VALOR_ANTERIOR",
          "VALOR_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "🏢 PROVEEDORES",
      color: "#7C3AED",
      hojas: {

        "PROV_MAESTRO": [
          "ID_PROVEEDOR",
          "TIPO_PERSONA",
          "TIPO_DOCUMENTO",
          "NIT_CC",
          "DV",
          "RAZON_SOCIAL",
          "NOMBRE_COMERCIAL",
          "NOMBRE_CONTACTO",
          "TELEFONO",
          "EMAIL",
          "DIRECCION",
          "CIUDAD",
          "DEPARTAMENTO",
          "CONDICION_PAGO",
          "CUPO_CREDITO",
          "ESTADO",
          "ID_SIIGO",
          "ID_ALEGRA",
          "FECHA_CREACION",
          "FECHA_MODIFICACION"
        ],

        "PROV_FORM": [
          "ID_PROVEEDOR",
          "TIPO_PERSONA",
          "TIPO_DOCUMENTO",
          "NIT_CC",
          "DV",
          "RAZON_SOCIAL",
          "NOMBRE_COMERCIAL",
          "NOMBRE_CONTACTO",
          "TELEFONO",
          "EMAIL",
          "DIRECCION",
          "CIUDAD",
          "DEPARTAMENTO",
          "CONDICION_PAGO",
          "CUPO_CREDITO",
          "OBSERVACION"
        ],

        "PROV_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_PROVEEDOR",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "CAMPO",
          "VALOR_ANTERIOR",
          "VALOR_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "📦 PRODUCTOS",
      color: "#16A34A",
      hojas: {

        "PROD_MAESTRO": [
          "ID_PRODUCTO",
          "CODIGO",
          "DESCRIPCION",
          "TIPO_PRODUCTO",
          "ID_CATEGORIA",
          "ID_UNIDAD",
          "PRECIO_VENTA",
          "COSTO_REFERENCIA",
          "IVA",
          "CONTROL_INVENTARIO",
          "STOCK_MINIMO",
          "STOCK_MAXIMO",
          "ESTADO",
          "ID_SIIGO",
          "ID_ALEGRA",
          "FECHA_CREACION",
          "FECHA_MODIFICACION"
        ],

        "PROD_CATEGORIAS": [
          "ID_CATEGORIA",
          "CODIGO_CATEGORIA",
          "NOMBRE_CATEGORIA",
          "DESCRIPCION",
          "ESTADO"
        ],

        "PROD_UNIDADES": [
          "ID_UNIDAD",
          "CODIGO",
          "NOMBRE",
          "ABREVIATURA",
          "TIPO",
          "ESTADO"
        ],

        "PROD_TIPOS": [
          "ID_TIPO",
          "CODIGO",
          "TIPO_PRODUCTO",
          "DESCRIPCION",
          "AFECTA_INVENTARIO",
          "ESTADO"
        ],

        "PROD_PRECIOS": [
          "ID_PRECIO",
          "ID_PRODUCTO",
          "TIPO_PRECIO",
          "PRECIO",
          "FECHA_INICIO",
          "FECHA_FIN",
          "ESTADO"
        ]
      }
    },

    {
      nombre: "🏗️ OBRAS Y PROYECTOS",
      color: "#92400E",
      hojas: {

        "OBR_MAESTRO": [
          "ID_OBRA",
          "CODIGO_OBRA",
          "NOMBRE_OBRA",
          "ID_CLIENTE",
          "DESCRIPCION",
          "UBICACION",
          "CIUDAD",
          "FECHA_INICIO",
          "FECHA_FIN_ESTIMADA",
          "FECHA_FIN_REAL",
          "PRESUPUESTO",
          "RESPONSABLE",
          "ESTADO",
          "FECHA_CREACION"
        ],

        "OBR_PRESUPUESTO": [
          "ID_PRESUPUESTO",
          "ID_OBRA",
          "CATEGORIA",
          "TIPO_RECURSO",
          "DESCRIPCION",
          "CANTIDAD",
          "UNIDAD",
          "COSTO_UNITARIO",
          "COSTO_TOTAL",
          "FECHA",
          "REGISTRO_USUARIO"
        ],

        "OBR_AVANCE": [
          "ID_AVANCE",
          "ID_OBRA",
          "FECHA",
          "PORCENTAJE_AVANCE",
          "DESCRIPCION",
          "VALOR_EJECUTADO",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "OBR_RECURSOS": [
          "ID_RECURSO",
          "ID_OBRA",
          "FECHA",
          "TIPO_RECURSO",
          "ID_PRODUCTO",
          "ID_PROVEEDOR",
          "DESCRIPCION",
          "CANTIDAD",
          "VALOR",
          "COSTO_TOTAL",
          "RESPONSABLE"
        ]
      }
    },

    {
      nombre: "💰 VENTAS",
      color: "#1D4ED8",
      hojas: {

        "VEN_CABECERA": [
          "ID_VENTA",
          "FECHA",
          "ID_CLIENTE",
          "TIPO_DOCUMENTO",
          "NUM_DOCUMENTO",
          "FORMA_PAGO",
          "CONDICION_PAGO",
          "FECHA_VENCIMIENTO",
          "SUBTOTAL",
          "DESCUENTO",
          "IVA",
          "TOTAL",
          "ESTADO",
          "USUARIO",
          "FECHA_CREACION"
        ],

        "VEN_DETALLE": [
          "ID_DETALLE",
          "ID_VENTA",
          "ID_PRODUCTO",
          "DESCRIPCION",
          "CANTIDAD",
          "ID_UNIDAD",
          "PRECIO_UNITARIO",
          "DESCUENTO",
          "IVA",
          "TOTAL"
        ],

        "VEN_DOCUMENTOS": [
          "ID_DOCUMENTO",
          "ID_VENTA",
          "TIPO_DOCUMENTO",
          "NUM_DOCUMENTO",
          "PREFIJO",
          "FECHA",
          "ESTADO_DOCUMENTO",
          "RUTA_DOCUMENTO",
          "OBSERVACION"
        ],

        "VEN_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_VENTA",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "🛒 COMPRAS",
      color: "#EA580C",
      hojas: {

        "COM_CABECERA": [
          "ID_COMPRA",
          "FECHA",
          "ID_PROVEEDOR",
          "TIPO_DOCUMENTO",
          "NUM_DOCUMENTO",
          "FORMA_PAGO",
          "CONDICION_PAGO",
          "FECHA_VENCIMIENTO",
          "SUBTOTAL",
          "DESCUENTO",
          "IVA",
          "TOTAL",
          "ESTADO",
          "USUARIO",
          "FECHA_CREACION"
        ],

        "COM_DETALLE": [
          "ID_DETALLE",
          "ID_COMPRA",
          "ID_PRODUCTO",
          "DESCRIPCION",
          "CANTIDAD",
          "ID_UNIDAD",
          "COSTO_UNITARIO",
          "DESCUENTO",
          "IVA",
          "TOTAL"
        ],

        "COM_DOCUMENTOS": [
          "ID_DOCUMENTO",
          "ID_COMPRA",
          "TIPO_DOCUMENTO",
          "NUM_DOCUMENTO",
          "FECHA",
          "ESTADO_DOCUMENTO",
          "RUTA_DOCUMENTO",
          "OBSERVACION"
        ],

        "COM_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_COMPRA",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "🗃️ INVENTARIO",
      color: "#0891B2",
      hojas: {

        "INV_MOVIMIENTOS": [
          "ID_MOVIMIENTO",
          "FECHA",
          "TIPO_MOVIMIENTO",
          "ID_PRODUCTO",
          "ID_OBRA",
          "DOCUMENTO_ORIGEN",
          "ID_ORIGEN",
          "ENTRADA",
          "SALIDA",
          "COSTO_UNITARIO",
          "COSTO_TOTAL",
          "SALDO",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "INV_SALDOS": [
          "ID_SALDO",
          "ID_PRODUCTO",
          "FECHA",
          "SALDO_INICIAL",
          "ENTRADAS",
          "SALIDAS",
          "SALDO_FINAL",
          "COSTO_PROMEDIO",
          "VALOR_INVENTARIO"
        ],

        "INV_KARDEX": [
          "FECHA",
          "ID_PRODUCTO",
          "TIPO_MOVIMIENTO",
          "DOCUMENTO",
          "ENTRADA",
          "SALIDA",
          "SALDO",
          "COSTO_UNITARIO",
          "COSTO_TOTAL"
        ],

        "INV_AJUSTES": [
          "ID_AJUSTE",
          "FECHA",
          "ID_PRODUCTO",
          "TIPO_AJUSTE",
          "CANTIDAD",
          "COSTO_UNITARIO",
          "VALOR",
          "MOTIVO",
          "RESPONSABLE",
          "ESTADO"
        ],

        "INV_TRASLADOS": [
          "ID_TRASLADO",
          "FECHA",
          "ID_PRODUCTO",
          "ORIGEN",
          "DESTINO",
          "CANTIDAD",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "💵 INGRESOS",
      color: "#0F766E",
      hojas: {

        "ING_MOVIMIENTOS": [
          "ID_INGRESO",
          "FECHA",
          "TIPO_INGRESO",
          "ORIGEN",
          "ID_ORIGEN",
          "ID_CLIENTE",
          "VALOR",
          "METODO_PAGO",
          "ID_CUENTA_DESTINO",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ],

        "ING_RECAUDOS": [
          "ID_RECAUDO",
          "FECHA",
          "ID_CLIENTE",
          "ID_CARTERA",
          "ID_VENTA",
          "VALOR",
          "METODO_PAGO",
          "ID_CUENTA_DESTINO",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "ING_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_INGRESO",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "VALOR_ANTERIOR",
          "VALOR_NUEVO",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "🧾 COSTOS",
      color: "#CA8A04",
      hojas: {

        "COS_MOVIMIENTOS": [
          "ID_COSTO",
          "FECHA",
          "TIPO_COSTO",
          "CATEGORIA",
          "ID_OBRA",
          "ID_PRODUCTO",
          "ID_PROVEEDOR",
          "DOCUMENTO_ORIGEN",
          "ID_ORIGEN",
          "VALOR",
          "CENTRO_COSTO",
          "CUENTA_CONTABLE",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ],

        "COS_CATEGORIAS": [
          "ID_CATEGORIA",
          "CODIGO",
          "NOMBRE_CATEGORIA",
          "TIPO_COSTO",
          "DESCRIPCION",
          "ESTADO"
        ],

        "COS_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_COSTO",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "VALOR_ANTERIOR",
          "VALOR_NUEVO",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "💸 GASTOS",
      color: "#DC2626",
      hojas: {

        "GAS_MOVIMIENTOS": [
          "ID_GASTO",
          "FECHA",
          "TIPO_GASTO",
          "CATEGORIA",
          "ID_PROVEEDOR",
          "DOCUMENTO_ORIGEN",
          "ID_ORIGEN",
          "VALOR",
          "IVA",
          "CENTRO_COSTO",
          "CUENTA_CONTABLE",
          "METODO_PAGO",
          "ID_CUENTA",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ],

        "GAS_CATEGORIAS": [
          "ID_CATEGORIA",
          "CODIGO",
          "NOMBRE_CATEGORIA",
          "DESCRIPCION",
          "CUENTA_CONTABLE",
          "CENTRO_COSTO",
          "ESTADO"
        ],

        "GAS_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_GASTO",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "VALOR_ANTERIOR",
          "VALOR_NUEVO",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "📋 CARTERA",
      color: "#0284C7",
      hojas: {

        "CAR_CUENTAS": [
          "ID_CARTERA",
          "ID_CLIENTE",
          "ID_VENTA",
          "DOCUMENTO",
          "FECHA_EMISION",
          "FECHA_VENCIMIENTO",
          "VALOR_DOCUMENTO",
          "ABONOS",
          "SALDO",
          "DIAS_VENCIDOS",
          "ESTADO"
        ],

        "CAR_RECAUDOS": [
          "ID_RECAUDO",
          "ID_CARTERA",
          "ID_CLIENTE",
          "ID_VENTA",
          "FECHA",
          "VALOR",
          "METODO_PAGO",
          "ID_CUENTA_DESTINO",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "CAR_VENCIMIENTOS": [
          "ID_CARTERA",
          "ID_CLIENTE",
          "DOCUMENTO",
          "FECHA_VENCIMIENTO",
          "VALOR_DOCUMENTO",
          "ABONOS",
          "SALDO",
          "DIAS_VENCIDOS",
          "RANGO_MORA",
          "ESTADO"
        ],

        "CAR_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_CARTERA",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "SALDO_ANTERIOR",
          "SALDO_NUEVO",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "📑 CUENTAS POR PAGAR",
      color: "#C026D3",
      hojas: {

        "CXP_CUENTAS": [
          "ID_CXP",
          "ID_PROVEEDOR",
          "ID_COMPRA",
          "DOCUMENTO",
          "FECHA_EMISION",
          "FECHA_VENCIMIENTO",
          "VALOR_DOCUMENTO",
          "ABONOS",
          "SALDO",
          "DIAS_VENCIDOS",
          "ESTADO"
        ],

        "CXP_PAGOS": [
          "ID_PAGO",
          "ID_CXP",
          "ID_PROVEEDOR",
          "ID_COMPRA",
          "FECHA",
          "VALOR",
          "METODO_PAGO",
          "ID_CUENTA_ORIGEN",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "CXP_VENCIMIENTOS": [
          "ID_CXP",
          "ID_PROVEEDOR",
          "DOCUMENTO",
          "FECHA_VENCIMIENTO",
          "VALOR_DOCUMENTO",
          "ABONOS",
          "SALDO",
          "DIAS_VENCIDOS",
          "RANGO_MORA",
          "ESTADO"
        ],

        "CXP_HISTORIAL": [
          "ID_HISTORIAL",
          "ID_CXP",
          "FECHA",
          "HORA",
          "USUARIO",
          "ACCION",
          "SALDO_ANTERIOR",
          "SALDO_NUEVO",
          "ESTADO_ANTERIOR",
          "ESTADO_NUEVO",
          "OBSERVACION"
        ]
      }
    },

    {
      nombre: "🏦 TESORERÍA",
      color: "#4F46E5",
      hojas: {

        "TES_CUENTAS": [
          "ID_CUENTA",
          "TIPO_CUENTA",
          "NOMBRE_CUENTA",
          "ENTIDAD",
          "NUMERO_CUENTA",
          "MONEDA",
          "SALDO_INICIAL",
          "FECHA_SALDO_INICIAL",
          "ESTADO"
        ],

        "TES_MOVIMIENTOS": [
          "ID_MOVIMIENTO",
          "FECHA",
          "TIPO_MOVIMIENTO",
          "ID_CUENTA",
          "ORIGEN",
          "ID_ORIGEN",
          "INGRESO",
          "EGRESO",
          "SALDO",
          "METODO_PAGO",
          "RESPONSABLE",
          "OBSERVACION"
        ],

        "TES_PAGOS": [
          "ID_PAGO",
          "FECHA",
          "TIPO_PAGO",
          "ORIGEN",
          "ID_ORIGEN",
          "ID_CUENTA",
          "VALOR",
          "METODO_PAGO",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ],

        "TES_RECAUDOS": [
          "ID_RECAUDO",
          "FECHA",
          "ORIGEN",
          "ID_ORIGEN",
          "ID_CUENTA",
          "VALOR",
          "METODO_PAGO",
          "RESPONSABLE",
          "ESTADO",
          "OBSERVACION"
        ],

        "TES_CONCILIACION": [
          "ID_CONCILIACION",
          "ID_CUENTA",
          "PERIODO",
          "FECHA_EXTRACTO",
          "SALDO_EXTRACTO",
          "SALDO_SISTEMA",
          "DIFERENCIA",
          "ESTADO",
          "USUARIO",
          "OBSERVACION"
        ]
      }
    }
  ];

  // ==========================================================
  // CREAR MENU
  // ==========================================================

  let menu = ss.getSheetByName("MENU");

  if (!menu) {
    menu = ss.insertSheet("MENU");
  }

  menu.clear();

  menu.getRange("A1")
    .setValue("📘 ERP V2 — ERP OPERATIVO")
    .setFontSize(18)
    .setFontWeight("bold");

  menu.getRange("A3")
    .setValue("MÓDULOS DEL SISTEMA")
    .setFontWeight("bold");

  let filaMenu = 4;

  // ==========================================================
  // CREAR MÓDULOS Y HOJAS
  // ==========================================================

  modulos.forEach(modulo => {

    menu.getRange(filaMenu, 1)
      .setValue(modulo.nombre)
      .setFontWeight("bold")
      .setBackground(modulo.color)
      .setFontColor("#FFFFFF");

    filaMenu++;

    Object.entries(modulo.hojas).forEach(([nombreHoja, encabezados]) => {

      let hoja = ss.getSheetByName(nombreHoja);

      // Si no existe, crearla
      if (!hoja) {
        hoja = ss.insertSheet(nombreHoja);
      }

      // Limpiar contenido existente
      hoja.clear();

      // Encabezados
      hoja.getRange(1, 1, 1, encabezados.length)
        .setValues([encabezados]);

      // Formato encabezado
      hoja.getRange(1, 1, 1, encabezados.length)
        .setFontWeight("bold")
        .setBackground(modulo.color)
        .setFontColor("#FFFFFF")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      // Congelar encabezado
      hoja.setFrozenRows(1);

      // Filtro
      if (hoja.getFilter()) {
        hoja.getFilter().remove();
      }

      hoja.getRange(
        1,
        1,
        Math.max(2, hoja.getMaxRows()),
        encabezados.length
      ).createFilter();

      // Ajustar columnas
      hoja.autoResizeColumns(1, encabezados.length);

      // Altura encabezado
      hoja.setRowHeight(1, 30);

      // Agregar hoja al menú
      menu.getRange(filaMenu, 1)
        .setValue("   ↳ " + nombreHoja)
        .setFontColor(modulo.color);

      filaMenu++;
    });

    filaMenu++;
  });

  // ==========================================================
  // FORMATO DEL MENU
  // ==========================================================

  menu.setColumnWidth(1, 300);

  menu.getRange("A1:A" + filaMenu)
    .setVerticalAlignment("middle");

  menu.setFrozenRows(3);

  // ==========================================================
  // MOVER MENU AL PRINCIPIO
  // ==========================================================

  ss.setActiveSheet(menu);
  ss.moveActiveSheet(1);

  // ==========================================================
  // MENSAJE FINAL
  // ==========================================================

  SpreadsheetApp.getUi().alert(
    "ERP OPERATIVO INSTALADO",
    "Se creó la estructura inicial del 📘 ERP OPERATIVO.\n\n" +
    "Las hojas fueron organizadas con sus encabezados, colores y filtros.\n\n" +
    "Siguiente etapa: revisar y validar la estructura antes de programar las automatizaciones.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
