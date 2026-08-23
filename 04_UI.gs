/**************************************************************
 * 04_UI.gs
 * ERP OPERATIVO
 *
 * RESPONSABILIDAD:
 * - Administrar la apariencia visual de las pestañas.
 * - Aplicar colores por módulo.
 * - Identificar las hojas físicas de cada módulo.
 * - Colorear automáticamente las pestañas existentes.
 *
 * ESTE ARCHIVO NO:
 * - Crea hojas.
 * - Elimina hojas.
 * - Crea hipervínculos.
 * - Abre formularios HTML.
 * - Modifica datos operativos.
 * - Gestiona usuarios, roles o permisos.
 **************************************************************/


  // ============================================================
  // 01. CONFIGURACIÓN DE COLORES
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
  // 02. HOJAS POR MÓDULO
  // Solo se incluyen hojas físicas de Google Sheets.
  // NO se incluyen archivos .gs ni formularios .html.
  // ============================================================

    const HOJAS_ERP = {

      // ----------------------------------------------------------
      // CONFIGURACIÓN
      // ----------------------------------------------------------

      CONFIGURACION: [

        "CFG_EMPRESA",
        "CFG_SISTEMA",
        "CFG_DOCUMENTOS",
        "CFG_IMPUESTOS",
        "CFG_CONTABILIDAD",
        "CFG_INVENTARIO",
        "CFG_COSTOS"

      ],


      // ----------------------------------------------------------
      // CLIENTES
      // ----------------------------------------------------------

      CLIENTES: [

        "CLI_MAESTRO",
        "CLI_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // PROVEEDORES
      // ----------------------------------------------------------

      PROVEEDORES: [

        "PROV_MAESTRO",
        "PROV_FORM",
        "PROV_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // PRODUCTOS
      // ----------------------------------------------------------

      PRODUCTOS: [

        "PROD_MAESTRO",
        "PROD_CATEGORIAS",
        "PROD_UNIDADES",
        "PROD_TIPOS",
        "PROD_PRECIOS"

      ],


      // ----------------------------------------------------------
      // OBRAS
      // ----------------------------------------------------------

      OBRAS: [

        "OBR_MAESTRO",
        "OBR_PRESUPUESTO",
        "OBR_AVANCE",
        "OBR_RECURSOS"

      ],


      // ----------------------------------------------------------
      // VENTAS
      // ----------------------------------------------------------

      VENTAS: [

        "VEN_CABECERA",
        "VEN_DETALLE",
        "VEN_DOCUMENTOS",
        "VEN_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // COMPRAS
      // ----------------------------------------------------------

      COMPRAS: [

        "COM_CABECERA",
        "COM_DETALLE",
        "COM_DOCUMENTOS",
        "COM_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // INVENTARIO
      // ----------------------------------------------------------

      INVENTARIO: [

        "INV_MOVIMIENTOS",
        "INV_SALDOS",
        "INV_KARDEX",
        "INV_AJUSTES",
        "INV_TRASLADOS"

      ],


      // ----------------------------------------------------------
      // INGRESOS
      // ----------------------------------------------------------

      INGRESOS: [

        "ING_MOVIMIENTOS",
        "ING_RECAUDOS",
        "ING_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // COSTOS
      // ----------------------------------------------------------

      COSTOS: [

        "COS_MOVIMIENTOS",
        "COS_CATEGORIAS",
        "COS_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // GASTOS
      // ----------------------------------------------------------

      GASTOS: [

        "GAS_MOVIMIENTOS",
        "GAS_CATEGORIAS",
        "GAS_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // CARTERA
      // ----------------------------------------------------------

      CARTERA: [

        "CAR_CUENTAS",
        "CAR_RECAUDOS",
        "CAR_VENCIMIENTOS",
        "CAR_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // CUENTAS POR PAGAR
      // ----------------------------------------------------------

      CUENTAS_POR_PAGAR: [

        "CXP_CUENTAS",
        "CXP_PAGOS",
        "CXP_VENCIMIENTOS",
        "CXP_HISTORIAL"

      ],


      // ----------------------------------------------------------
      // TESORERÍA
      // ----------------------------------------------------------

      TESORERIA: [

        "TES_CUENTAS",
        "TES_MOVIMIENTOS",
        "TES_PAGOS",
        "TES_RECAUDOS",
        "TES_CONCILIACION"

      ],


      // ----------------------------------------------------------
      // SEGURIDAD
      // ----------------------------------------------------------

      SEGURIDAD: [

        "USR_USUARIOS",
        "USR_ROLES",
        "USR_PERMISOS"

      ]

    };



  // ============================================================
  // 03. COLOREAR TODAS LAS PESTAÑAS
  // Aplica a cada hoja existente el color correspondiente
  // a su módulo.
  // ============================================================

    function COLOREAR_TODAS_LAS_PESTANAS() {

      // ----------------------------------------------------------
      // 03.01 CONECTAR CON EL LIBRO
      // ----------------------------------------------------------

        const ss =
          SpreadsheetApp.getActiveSpreadsheet();


      // ----------------------------------------------------------
      // 03.02 VARIABLES DE CONTROL
      // ----------------------------------------------------------

        let coloreadas = 0;

        const noEncontradas = [];


      // ----------------------------------------------------------
      // 03.03 RECORRER MÓDULOS
      // ----------------------------------------------------------

        Object.keys(HOJAS_ERP).forEach(
          function(modulo) {

            const color =
              COLORES_ERP[modulo];

            const hojas =
              HOJAS_ERP[modulo];


            // ------------------------------------------------------
            // 03.03.01 RECORRER HOJAS DEL MÓDULO
            // ------------------------------------------------------

              hojas.forEach(
                function(nombreHoja) {

                  const hoja =
                    ss.getSheetByName(
                      nombreHoja
                    );


                  // --------------------------------------------------
                  // HOJA ENCONTRADA
                  // --------------------------------------------------

                  if (hoja) {

                    hoja.setTabColor(
                      color
                    );

                    coloreadas++;

                  }


                  // --------------------------------------------------
                  // HOJA NO ENCONTRADA
                  // --------------------------------------------------

                  else {

                    noEncontradas.push(
                      nombreHoja
                    );

                  }

                }
              );

            }
          );


      // ----------------------------------------------------------
      // 03.04 COLOREAR HOJA PRINCIPAL
      // ----------------------------------------------------------

        const menu =
          ss.getSheetByName(
            "🏠MENU"
          );


        if (menu) {

          menu.setTabColor(
            COLORES_ERP.MENU
          );

          coloreadas++;

        }

        else {

          noEncontradas.push(
            "🏠MENU"
          );

        }


      // ----------------------------------------------------------
      // 03.05 APLICAR CAMBIOS
      // ----------------------------------------------------------

        SpreadsheetApp.flush();


      // ----------------------------------------------------------
      // 03.06 REGISTRAR RESULTADO
      // ----------------------------------------------------------

        REGISTRAR_RESULTADO_COLORES(
          coloreadas,
          noEncontradas
        );

      }



  // ============================================================
  // 04. REGISTRAR RESULTADO
  // Registra en el historial de ejecución el resultado del proceso.
  // ============================================================

    function REGISTRAR_RESULTADO_COLORES(
      coloreadas,
      noEncontradas
    ) {

      console.log(
        "================================="
      );

      console.log(
        "🎨 COLORES DEL ERP"
      );

      console.log(
        "================================="
      );

      console.log(
        "Pestañas coloreadas: " +
        coloreadas
      );

      console.log(
        "Hojas no encontradas: " +
        noEncontradas.length
      );


    // ----------------------------------------------------------
    // 04.01 MOSTRAR HOJAS NO ENCONTRADAS
    // ----------------------------------------------------------

      if (
        noEncontradas.length > 0
      ) {

        console.log(
          "Hojas no encontradas:"
        );

        noEncontradas.forEach(
          function(nombreHoja) {

            console.log(
              "- " +
              nombreHoja
            );

          }
        );

      }


      console.log(
        "================================="
      );

  }