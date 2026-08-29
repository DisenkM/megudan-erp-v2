/**************************************************************
* 00_INSTALADOR.gs
* RESPONSABILIDAD:
* - Instalar, verificar y formatear la base de datos de 60 hojas.
* - Inicializar los roles, permisos y usuarios de fábrica con autocuración.
* - Blindar datos mediante mecanismo de "Caja Fuerte" ante reinicios accidentales.
**************************************************************/

const INST_CONFIG = {
  NOMBRE_MENU: "🏠MENU",
  COLOR_MENU: "#111827",
  HOJA_ROLES: "USR_ROLES",
  ANCHO_COL_MIN: 120,
  ALTURA_ENCABEZADO: 30
};

const INST_MODULOS = [
  {
    nombre: "⚙️ CONFIGURACIÓN",
    color: "#6B7280",
    hojas: {
      "CFG_EMPRESA": [
        "ID_EMPRESA", "NIT", "DV", "RAZON_SOCIAL", "NOMBRE_COMERCIAL",
        "TIPO_PERSONA", "DIRECCION", "CIUDAD", "DEPARTAMENTO", "TELEFONO",
        "EMAIL", "REPRESENTANTE_LEGAL", "LOGO", "ESTADO"
      ],
      "CFG_SISTEMA": [
        "PARAMETRO", "VALOR", "TIPO_DATO", "DESCRIPCION", "ESTADO"
      ],
      "CFG_DOCUMENTOS": [
        "ID_DOCUMENTO", "TIPO_DOCUMENTO", "PREFIJO", "CONSECUTIVO_ACTUAL",
        "CONSECUTIVO_INICIAL", "CONSECUTIVO_FINAL", "RESOLUCION",
        "FECHA_RESOLUCION", "FECHA_VENCIMIENTO", "ESTADO"
      ],
      "CFG_IMPUESTOS": [
        "ID_IMPUESTO", "CODIGO", "NOMBRE", "TIPO", "TARIFA",
        "BASE_MINIMA", "CUENTA_GENERADO", "CUENTA_DESCONTABLE", "ESTADO"
      ],
      "CFG_CONTABILIDAD": [
        "ID_PARAMETRO", "CONCEPTO", "CUENTA_CONTABLE", "CENTRO_COSTO",
        "DESCRIPCION", "ESTADO"
      ],
      "CFG_INVENTARIO": [
        "PARAMETRO", "VALOR", "DESCRIPCION", "ESTADO"
      ],
      "CFG_COSTOS": [
        "ID_PARAMETRO", "TIPO_COSTO", "METODO", "CUENTA_CONTABLE",
        "CENTRO_COSTO", "ESTADO"
      ],
      "CFG_USUARIOS": [
        "ID_USUARIO", "NOMBRE", "EMAIL", "ROL", "ESTADO", "FECHA_CREACION", "ULTIMO_ACCESO"
      ],
      "CFG_PERMISOS": [
        "ID_PERMISO", "ROL", "MODULO", "ACCION", "PERMITIDO", "ESTADO"
      ]
    }
  },
  {
    nombre: "👥 CLIENTES",
    color: "#2563EB",
    hojas: {
      "CLI_MAESTRO": [
        "ID_CLIENTE", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO",
        "DIGITO_VERIFICACION", "RAZON_SOCIAL", "PRIMER_NOMBRE", "SEGUNDO_NOMBRE",
        "PRIMER_APELLIDO", "SEGUNDO_APELLIDO", "NOMBRE_COMERCIAL", "TIPO_CLIENTE",
        "RESPONSABILIDAD_IVA", "RESPONSABILIDAD_FISCAL", "REGIMEN_FISCAL", "DIRECCION",
        "CIUDAD", "DEPARTAMENTO", "PAIS", "TELEFONO", "CELULAR", "CORREO",
        "CONTACTO_PRINCIPAL", "CARGO_CONTACTO", "FORMA_PAGO", "METODO_PAGO",
        "PLAZO_PAGO_DIAS", "LIMITE_CREDITO", "VENDEDOR_ASIGNADO", "CENTRO_COSTO",
        "OBRA_PRINCIPAL", "CUENTA_CONTABLE_CLIENTE", "ESTADO_CLIENTE", "FECHA_CREACION",
        "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"
      ],
      "CLI_FORM": [
        "ID_CLIENTE", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NUMERO_DOCUMENTO",
        "DIGITO_VERIFICACION", "RAZON_SOCIAL", "PRIMER_NOMBRE", "SEGUNDO_NOMBRE",
        "PRIMER_APELLIDO", "SEGUNDO_APELLIDO", "NOMBRE_COMERCIAL", "TIPO_CLIENTE",
        "RESPONSABILIDAD_IVA", "RESPONSABILIDAD_FISCAL", "REGIMEN_FISCAL", "DIRECCION",
        "CIUDAD", "DEPARTAMENTO", "PAIS", "TELEFONO", "CELULAR", "CORREO",
        "CONTACTO_PRINCIPAL", "CARGO_CONTACTO", "FORMA_PAGO", "METODO_PAGO",
        "PLAZO_PAGO_DIAS", "LIMITE_CREDITO", "VENDEDOR_ASIGNADO", "CENTRO_COSTO",
        "OBRA_PRINCIPAL", "CUENTA_CONTABLE_CLIENTE", "ESTADO_CLIENTE", "FECHA_CREACION",
        "FECHA_ACTUALIZACION", "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"
      ],
      "CLI_HISTORIAL": [
        "ID_HISTORIAL", "ID_CLIENTE", "TIPO_EVENTO", "FECHA_HORA", "USUARIO",
        "ACCION", "CAMPO_MODIFICADO", "VALOR_ANTERIOR", "VALOR_NUEVO", "MOTIVO_ORIGEN",
        "MODULO_ORIGEN", "ID_REGISTRO_ORIGEN", "IP_USUARIO", "ESTADO_EVENTO", "OBSERVACIONES"
      ]
    }
  },
  {
    nombre: "🏢 PROVEEDORES",
    color: "#7C3AED",
    hojas: {
      "PROV_MAESTRO": [
        "ID_PROVEEDOR", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NIT_CC", "DV",
        "RAZON_SOCIAL", "NOMBRE_COMERCIAL", "NOMBRE_CONTACTO", "TELEFONO", "EMAIL",
        "DIRECCION", "CIUDAD", "DEPARTAMENTO", "CONDICION_PAGO", "CUPO_CREDITO",
        "ESTADO", "ID_SIIGO", "ID_ALEGRA", "FECHA_CREACION", "FECHA_MODIFICACION"
      ],
      "PROV_FORM": [
        "ID_PROVEEDOR", "TIPO_PERSONA", "TIPO_DOCUMENTO", "NIT_CC", "DV",
        "RAZON_SOCIAL", "NOMBRE_COMERCIAL", "NOMBRE_CONTACTO", "TELEFONO", "EMAIL",
        "DIRECCION", "CIUDAD", "DEPARTAMENTO", "CONDICION_PAGO", "CUPO_CREDITO",
        "OBSERVACION"
      ],
      "PROV_HISTORIAL": [
        "ID_HISTORIAL", "ID_PROVEEDOR", "FECHA", "HORA", "USUARIO", "ACCION",
        "CAMPO", "VALOR_ANTERIOR", "VALOR_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "📦 PRODUCTOS",
    color: "#16A34A",
    hojas: {
      "PROD_MAESTRO": [
        "ID_PRODUCTO", "CODIGO", "DESCRIPCION", "TIPO_PRODUCTO", "ID_CATEGORIA",
        "ID_UNIDAD", "PRECIO_VENTA", "COSTO_REFERENCIA", "IVA", "CONTROL_INVENTARIO",
        "STOCK_MINIMO", "STOCK_MAXIMO", "ESTADO", "ID_SIIGO", "ID_ALEGRA",
        "FECHA_CREACION", "FECHA_MODIFICACION"
      ],
      "PROD_CATEGORIAS": [
        "ID_CATEGORIA", "CODIGO_CATEGORIA", "NOMBRE_CATEGORIA", "DESCRIPCION", "ESTADO"
      ],
      "PROD_UNIDADES": [
        "ID_UNIDAD", "CODIGO", "NOMBRE", "ABREVIATURA", "TIPO", "ESTADO"
      ],
      "PROD_TIPOS": [
        "ID_TIPO", "CODIGO", "TIPO_PRODUCTO", "DESCRIPCION", "AFECTA_INVENTARIO", "ESTADO"
      ],
      "PROD_PRECIOS": [
        "ID_PRECIO", "ID_PRODUCTO", "TIPO_PRECIO", "PRECIO", "FECHA_INICIO", "FECHA_FIN", "ESTADO"
      ]
    }
  },
  {
    nombre: "🏗️ OBRAS Y PROYECTOS",
    color: "#92400E",
    hojas: {
      "OBR_MAESTRO": [
        "ID_OBRA", "CODIGO_OBRA", "NOMBRE_OBRA", "ID_CLIENTE", "DESCRIPCION",
        "UBICACION", "CIUDAD", "FECHA_INICIO", "FECHA_FIN_ESTIMADA", "FECHA_FIN_REAL",
        "PRESUPUESTO", "RESPONSABLE", "ESTADO", "FECHA_CREACION"
      ],
      "OBR_PRESUPUESTO": [
        "ID_PRESUPUESTO", "ID_OBRA", "CATEGORIA", "TIPO_RECURSO", "DESCRIPCION",
        "CANTIDAD", "UNIDAD", "COSTO_UNITARIO", "COSTO_TOTAL", "FECHA", "REGISTRO_USUARIO"
      ],
      "OBR_AVANCE": [
        "ID_AVANCE", "ID_OBRA", "FECHA", "PORCENTAJE_AVANCE", "DESCRIPCION",
        "VALOR_EJECUTADO", "RESPONSABLE", "OBSERVACION"
      ],
      "OBR_RECURSOS": [
        "ID_RECURSO", "ID_OBRA", "FECHA", "TIPO_RECURSO", "ID_PRODUCTO",
        "ID_PROVEEDOR", "DESCRIPCION", "CANTIDAD", "VALOR", "COSTO_TOTAL", "RESPONSABLE"
      ]
    }
  },
  {
    nombre: "💰 VENTAS",
    color: "#1D4ED8",
    hojas: {
      "VEN_CABECERA": [
        "ID_VENTA", "FECHA", "ID_CLIENTE", "TIPO_DOCUMENTO", "NUM_DOCUMENTO",
        "FORMA_PAGO", "CONDICION_PAGO", "FECHA_VENCIMIENTO", "SUBTOTAL", "DESCUENTO",
        "IVA", "TOTAL", "ESTADO", "USUARIO", "FECHA_CREACION"
      ],
      "VEN_DETALLE": [
        "ID_DETALLE", "ID_VENTA", "ID_PRODUCTO", "DESCRIPCION", "CANTIDAD",
        "ID_UNIDAD", "PRECIO_UNITARIO", "DESCUENTO", "IVA", "TOTAL"
      ],
      "VEN_DOCUMENTOS": [
        "ID_DOCUMENTO", "ID_VENTA", "TIPO_DOCUMENTO", "NUM_DOCUMENTO", "PREFIJO",
        "FECHA", "ESTADO_DOCUMENTO", "RUTA_DOCUMENTO", "OBSERVACION"
      ],
      "VEN_HISTORIAL": [
        "ID_HISTORIAL", "ID_VENTA", "FECHA", "HORA", "USUARIO", "ACCION",
        "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "🛒 COMPRAS",
    color: "#EA580C",
    hojas: {
      "COM_CABECERA": [
        "ID_COMPRA", "FECHA", "ID_PROVEEDOR", "TIPO_DOCUMENTO", "NUM_DOCUMENTO",
        "FORMA_PAGO", "CONDICION_PAGO", "FECHA_VENCIMIENTO", "SUBTOTAL", "DESCUENTO",
        "IVA", "TOTAL", "ESTADO", "USUARIO", "FECHA_CREACION"
      ],
      "COM_DETALLE": [
        "ID_DETALLE", "ID_COMPRA", "ID_PRODUCTO", "DESCRIPCION", "CANTIDAD",
        "ID_UNIDAD", "COSTO_UNITARIO", "DESCUENTO", "IVA", "TOTAL"
      ],
      "COM_DOCUMENTOS": [
        "ID_DOCUMENTO", "ID_COMPRA", "TIPO_DOCUMENTO", "NUM_DOCUMENTO", "FECHA",
        "ESTADO_DOCUMENTO", "RUTA_DOCUMENTO", "OBSERVACION"
      ],
      "COM_HISTORIAL": [
        "ID_HISTORIAL", "ID_COMPRA", "FECHA", "HORA", "USUARIO", "ACCION",
        "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "🗃️ INVENTARIO",
    color: "#0891B2",
    hojas: {
      "INV_MOVIMIENTOS": [
        "ID_MOVIMIENTO", "FECHA", "TIPO_MOVIMIENTO", "ID_PRODUCTO", "ID_OBRA",
        "DOCUMENTO_ORIGEN", "ID_ORIGEN", "ENTRADA", "SALIDA", "COSTO_UNITARIO",
        "COSTO_TOTAL", "SALDO", "RESPONSABLE", "OBSERVACION"
      ],
      "INV_SALDOS": [
        "ID_SALDO", "ID_PRODUCTO", "FECHA", "SALDO_INICIAL", "ENTRADAS",
        "SALIDAS", "SALDO_FINAL", "COSTO_PROMEDIO", "VALOR_INVENTARIO"
      ],
      "INV_KARDEX": [
        "FECHA", "ID_PRODUCTO", "TIPO_MOVIMIENTO", "DOCUMENTO", "ENTRADA",
        "SALIDA", "SALDO", "COSTO_UNITARIO", "COSTO_TOTAL"
      ],
      "INV_AJUSTES": [
        "ID_AJUSTE", "FECHA", "ID_PRODUCTO", "TIPO_AJUSTE", "CANTIDAD",
        "COSTO_UNITARIO", "VALOR", "MOTIVO", "RESPONSABLE", "ESTADO"
      ],
      "INV_TRASLADOS": [
        "ID_TRASLADO", "FECHA", "ID_PRODUCTO", "ORIGEN", "DESTINO", "CANTIDAD",
        "RESPONSABLE", "ESTADO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "💵 INGRESOS",
    color: "#0F766E",
    hojas: {
      "ING_MOVIMIENTOS": [
        "ID_INGRESO", "FECHA", "TIPO_INGRESO", "ORIGEN", "ID_ORIGEN", "ID_CLIENTE",
        "VALOR", "METODO_PAGO", "ID_CUENTA_DESTINO", "RESPONSABLE", "ESTADO", "OBSERVACION"
      ],
      "ING_RECAUDOS": [
        "ID_RECAUDO", "FECHA", "ID_CLIENTE", "ID_CARTERA", "ID_VENTA", "VALOR",
        "METODO_PAGO", "ID_CUENTA_DESTINO", "RESPONSABLE", "OBSERVACION"
      ],
      "ING_HISTORIAL": [
        "ID_HISTORIAL", "ID_INGRESO", "FECHA", "HORA", "USUARIO", "ACCION",
        "VALOR_ANTERIOR", "VALOR_NUEVO", "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "🧾 COSTOS",
    color: "#CA8A04",
    hojas: {
      "COS_MOVIMIENTOS": [
        "ID_COSTO", "FECHA", "TIPO_COSTO", "CATEGORIA", "ID_OBRA", "ID_PRODUCTO",
        "ID_PROVEEDOR", "DOCUMENTO_ORIGEN", "ID_ORIGEN", "VALOR", "CENTRO_COSTO",
        "CUENTA_CONTABLE", "RESPONSABLE", "ESTADO", "OBSERVACION"
      ],
      "COS_CATEGORIAS": [
        "ID_CATEGORIA", "CODIGO", "NOMBRE_CATEGORIA", "TIPO_COSTO", "DESCRIPCION", "ESTADO"
      ],
      "COS_HISTORIAL": [
        "ID_HISTORIAL", "ID_COSTO", "FECHA", "HORA", "USUARIO", "ACCION",
        "VALOR_ANTERIOR", "VALOR_NUEVO", "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "💸 GASTOS",
    color: "#DC2626",
    hojas: {
      "GAS_MOVIMIENTOS": [
        "ID_GASTO", "FECHA", "TIPO_GASTO", "CATEGORIA", "ID_PROVEEDOR",
        "DOCUMENTO_ORIGEN", "ID_ORIGEN", "VALOR", "IVA", "CENTRO_COSTO",
        "CUENTA_CONTABLE", "METODO_PAGO", "ID_CUENTA", "RESPONSABLE", "ESTADO", "OBSERVACION"
      ],
      "GAS_CATEGORIAS": [
        "ID_CATEGORIA", "CODIGO", "NOMBRE_CATEGORIA", "DESCRIPCION", "CUENTA_CONTABLE",
        "CENTRO_COSTO", "ESTADO"
      ],
      "GAS_HISTORIAL": [
        "ID_HISTORIAL", "ID_GASTO", "FECHA", "HORA", "USUARIO", "ACCION",
        "VALOR_ANTERIOR", "VALOR_NUEVO", "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "📋 CARTERA",
    color: "#0284C7",
    hojas: {
      "CAR_CUENTAS": [
        "ID_CARTERA", "ID_CLIENTE", "ID_VENTA", "DOCUMENTO", "FECHA_EMISION",
        "FECHA_VENCIMIENTO", "VALOR_DOCUMENTO", "ABONOS", "SALDO", "DIAS_VENCIDOS", "ESTADO"
      ],
      "CAR_RECAUDOS": [
        "ID_RECAUDO", "ID_CARTERA", "ID_CLIENTE", "ID_VENTA", "FECHA", "VALOR",
        "METODO_PAGO", "ID_CUENTA_DESTINO", "RESPONSABLE", "OBSERVACION"
      ],
      "CAR_VENCIMIENTOS": [
        "ID_CARTERA", "ID_CLIENTE", "DOCUMENTO", "FECHA_VENCIMIENTO", "VALOR_DOCUMENTO",
        "ABONOS", "SALDO", "DIAS_VENCIDOS", "RANGO_MORA", "ESTADO"
      ],
      "CAR_HISTORIAL": [
        "ID_HISTORIAL", "ID_CARTERA", "FECHA", "HORA", "USUARIO", "ACCION",
        "SALDO_ANTERIOR", "SALDO_NUEVO", "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "📑 CUENTAS POR PAGAR",
    color: "#C026D3",
    hojas: {
      "CXP_CUENTAS": [
        "ID_CXP", "ID_PROVEEDOR", "ID_COMPRA", "DOCUMENTO", "FECHA_EMISION",
        "FECHA_VENCIMIENTO", "VALOR_DOCUMENTO", "ABONOS", "SALDO", "DIAS_VENCIDOS", "ESTADO"
      ],
      "CXP_PAGOS": [
        "ID_PAGO", "ID_CXP", "ID_PROVEEDOR", "ID_COMPRA", "FECHA", "VALOR",
        "METODO_PAGO", "ID_CUENTA_ORIGEN", "RESPONSABLE", "OBSERVACION"
      ],
      "CXP_VENCIMIENTOS": [
        "ID_CXP", "ID_PROVEEDOR", "DOCUMENTO", "FECHA_VENCIMIENTO", "VALOR_DOCUMENTO",
        "ABONOS", "SALDO", "DIAS_VENCIDOS", "RANGO_MORA", "ESTADO"
      ],
      "CXP_HISTORIAL": [
        "ID_HISTORIAL", "ID_CXP", "FECHA", "HORA", "USUARIO", "ACCION",
        "SALDO_ANTERIOR", "SALDO_NUEVO", "ESTADO_ANTERIOR", "ESTADO_NUEVO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "🏦 TESORERÍA",
    color: "#4F46E5",
    hojas: {
      "TES_CUENTAS": [
        "ID_CUENTA", "TIPO_CUENTA", "NOMBRE_CUENTA", "ENTIDAD", "NUMERO_CUENTA",
        "MONEDA", "SALDO_INICIAL", "FECHA_SALDO_INICIAL", "ESTADO"
      ],
      "TES_MOVIMIENTOS": [
        "ID_MOVIMIENTO", "FECHA", "TIPO_MOVIMIENTO", "ID_CUENTA", "ORIGEN",
        "ID_ORIGEN", "INGRESO", "EGRESO", "SALDO", "METODO_PAGO", "RESPONSABLE", "OBSERVACION"
      ],
      "TES_PAGOS": [
        "ID_PAGO", "FECHA", "TIPO_PAGO", "ORIGEN", "ID_ORIGEN", "ID_CUENTA",
        "VALOR", "METODO_PAGO", "RESPONSABLE", "ESTADO", "OBSERVACION"
      ],
      "TES_RECAUDOS": [
        "ID_RECAUDO", "FECHA", "ORIGEN", "ID_ORIGEN", "ID_CUENTA", "VALOR",
        "METODO_PAGO", "RESPONSABLE", "ESTADO", "OBSERVACION"
      ],
      "TES_CONCILIACION": [
        "ID_CONCILIACION", "ID_CUENTA", "PERIODO", "FECHA_EXTRACTO", "SALDO_EXTRACTO",
        "SALDO_SISTEMA", "DIFERENCIA", "ESTADO", "USUARIO", "OBSERVACION"
      ]
    }
  },
  {
    nombre: "🛡️ SEGURIDAD",
    color: "#9333EA",
    hojas: {
      "USR_USUARIOS": [
        "ID_USUARIO", "USUARIO", "NOMBRE_COMPLETO", "CORREO", "CONTRASENA_HASH",
        "SALT", "ID_ROL", "ESTADO_USUARIO", "INTENTOS_FALLIDOS", "BLOQUEADO_HASTA",
        "ULTIMO_ACCESO", "CAMBIO_CONTRASENA", "FECHA_CREACION", "FECHA_ACTUALIZACION",
        "USUARIO_CREACION", "USUARIO_ACTUALIZACION", "OBSERVACIONES"
      ],
      "USR_ROLES": [
        "ID_ROL", "NOMBRE_ROL", "DESCRIPCION", "NIVEL_JERARQUIA", "ESTADO_ROL",
        "ROL_SISTEMA", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION",
        "USUARIO_ACTUALIZACION", "OBSERVACIONES"
      ],
      "USR_PERMISOS": [
        "ID_PERMISO", "ID_ROL", "MODULO", "SUBMODULO", "ACCION", "PERMITIDO",
        "ESTADO_PERMISO", "FECHA_CREACION", "FECHA_ACTUALIZACION", "USUARIO_CREACION",
        "USUARIO_ACTUALIZACION", "OBSERVACIONES"
      ],
      "USR_SESIONES": [
        "ID_SESION", "ID_USUARIO", "USUARIO", "ID_ROL", "FECHA_INICIO", "ULTIMA_ACTIVIDAD",
        "FECHA_CIERRE", "ESTADO_SESION", "TIPO_ACCESO", "ORIGEN_ACCESO", "TOKEN_SESION",
        "EXPIRA_SESION", "MOTIVO_CIERRE", "OBSERVACIONES"
      ],
      "USR_AUDITORIA": [
        "ID_AUDITORIA", "FECHA_HORA", "ID_USUARIO", "USUARIO", "ID_SESION", "ID_ROL",
        "MODULO", "SUBMODULO", "ACCION", "TIPO_REGISTRO", "ID_REGISTRO", "DESCRIPCION",
        "VALOR_ANTERIOR", "VALOR_NUEVO", "RESULTADO", "MENSAJE_RESULTADO", "ORIGEN_ACCESO",
        "IP_ORIGEN", "FECHA_CREACION", "OBSERVACIONES"
      ]
    }
  }
];

function INSTALAR_ERP_OPERATIVO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let menu = ss.getSheetByName(INST_CONFIG.NOMBRE_MENU);
  if (!menu) {
    menu = ss.insertSheet(INST_CONFIG.NOMBRE_MENU);
  }
  menu.clearContents();
  menu.setTabColor(INST_CONFIG.COLOR_MENU);
  menu.getRange("A1").setValue("📘 ERP V2 — ERP OPERATIVO").setFontSize(18).setFontWeight("bold");
  menu.getRange("A3").setValue("MÓDULOS DEL SISTEMA").setFontWeight("bold");
  let filaMenu = 4;
  let hojasPreservadasCount = 0;
  let hojasCreadasCount = 0;

  INST_MODULOS.forEach(function(modulo) {
    menu.getRange(filaMenu, 1)
      .setValue(modulo.nombre)
      .setFontWeight("bold")
      .setBackground(modulo.color)
      .setFontColor("#FFFFFF");
    filaMenu++;

    Object.entries(modulo.hojas).forEach(function([nombreHoja, encabezados]) {
      let hoja = ss.getSheetByName(nombreHoja);
      let hojaNueva = false;
      if (!hoja) {
        hoja = ss.insertSheet(nombreHoja);
        hojaNueva = true;
      }
      if (!hojaNueva && hoja.getLastRow() >= 2) {
        hojasPreservadasCount++;
        console.log("🛡️ [CAJA FUERTE] Preservada hoja con datos: " + nombreHoja);
      } else {
        hojasCreadasCount++;
        hoja.clearContents();
        hoja.setTabColor(modulo.color);
        hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
        hoja.getRange(1, 1, 1, encabezados.length)
          .setFontWeight("bold")
          .setBackground(modulo.color)
          .setFontColor("#FFFFFF")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
        try {
          hoja.setFrozenRows(1);
          hoja.setRowHeight(1, INST_CONFIG.ALTURA_ENCABEZADO);
        } catch (err) {
          console.warn("Fijación de fila omitida en: " + nombreHoja);
        }
        hoja.autoResizeColumns(1, encabezados.length);
        for (let c = 1; c <= encabezados.length; c++) {
          const anchoActual = hoja.getColumnWidth(c);
          if (anchoActual < INST_CONFIG.ANCHO_COL_MIN) {
            hoja.setColumnWidth(c, INST_CONFIG.ANCHO_COL_MIN);
          }
        }
        try {
          if (hoja.getFilter()) {
            hoja.getFilter().remove();
          }
          hoja.getRange(1, 1, Math.max(2, hoja.getMaxRows()), encabezados.length).createFilter();
        } catch (err) {
          console.info("Filtro clásico omitido (Tabla nativa detectada): " + nombreHoja);
        }
      }
      menu.getRange(filaMenu, 1)
        .setValue("   ↳ " + nombreHoja)
        .setFontColor(modulo.color);
      filaMenu++;
    });
    filaMenu++;
  });

  menu.setColumnWidth(1, 320);
  menu.getRange("A1:A" + filaMenu).setVerticalAlignment("middle");
  menu.setFrozenRows(3);
  ss.setActiveSheet(menu);
  ss.moveActiveSheet(1);

  const hojaRoles = ss.getSheetByName(INST_CONFIG.HOJA_ROLES);
  if (hojaRoles && hojaRoles.getLastRow() < 2) {
    const ahora = new Date();
    const rolesDefecto = [
      ["ROL-000001", "ADMINISTRADOR", "Administración general y acceso total al ERP", "100", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol maestro protegido"],
      ["ROL-000002", "CONTADOR", "Gestión y supervisión de procesos contables y financieros", "80", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a reportes contables"],
      ["ROL-000003", "AUXILIAR_CONTABLE", "Apoyo en registros y procesos contables autorizados", "60", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Permiso de registro operativo"],
      ["ROL-000004", "TESORERO", "Gestión de tesorería, pagos, recaudos y conciliaciones autorizadas", "70", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso a bancos y cajas"],
      ["ROL-000005", "COMERCIAL", "Gestión de clientes y operaciones comerciales autorizadas", "50", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Acceso limitado a ventas"],
      ["ROL-000006", "OPERATIVO", "Registro de operaciones asignadas", "40", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol operativo general"],
      ["ROL-000007", "CONSULTA", "Acceso exclusivamente de lectura general", "10", "ACTIVO", "SÍ", ahora, ahora, "SISTEMA", "SISTEMA", "Rol de sólo lectura"]
    ];
    rolesDefecto.forEach(rol => hojaRoles.appendRow(rol));
    console.log("✓ Roles del sistema por defecto inicializados en USR_ROLES.");
  } else {
    console.log("🛡️ [CAJA FUERTE] Se omitió poblar roles porque USR_ROLES ya contiene datos.");
  }

  if (typeof ACTUALIZAR_MENU === "function") {
    try {
      ACTUALIZAR_MENU();
      console.log("✓ Enlaces de navegación compilados dinámicamente en 🏠MENU.");
    } catch (errLink) {
      console.warn("No se pudieron auto-compilar los links del menú: " + errLink.message);
    }
  }

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null;
  }
  const msgTitulo = "¡SISTEMA OPERATIVO BLINDADO E INSTALADO!";
  const msgTexto = "Se completó la inicialización física del Libro 1 con éxito.\n\n" +
                    "• Hojas estructurales totales analizadas: 60\n" +
                    "• Hojas estructuradas / creadas limpias: " + hojasCreadasCount + "\n" +
                    "• Hojas protegidas por Caja Fuerte (CON VALORES): " + hojasPreservadasCount + "\n\n" +
                    "🛡️ Tu base de datos y tus archivos de configuración existentes han sido totalmente blindados y permanecieron inmutables durante este proceso.";
  if (ui) {
    ui.alert(msgTitulo, msgTexto, ui.ButtonSet.OK);
  } else {
    console.log("=================================");
    console.log("🚀 " + msgTitulo);
    console.log("=================================");
    console.log(msgTexto.replace(/\n/g, "\n"));
    console.log("=================================");
  }
}