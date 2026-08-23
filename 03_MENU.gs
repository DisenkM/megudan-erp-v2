/**************************************************************
 * 03_MENU.gs
 * ERP OPERATIVO
 *
 * RESPONSABILIDAD:
 * - Administrar la navegación interna del ERP.
 * - Crear y actualizar hipervínculos del 🏠MENU.
 * - Detectar las hojas existentes.
 *
 * ESTE ARCHIVO NO:
 * - Crea hojas.
 * - Elimina hojas.
 * - Modifica columnas.
 * - Colorea pestañas.
 * - Abre formularios HTML.
 **************************************************************/


  // ============================================================
  // 01. CONFIGURACIÓN
  // ============================================================

  const MENU_CONFIG = {

    // ----------------------------------------------------------
    // 01.01 NOMBRE DE LA HOJA PRINCIPAL
    // ----------------------------------------------------------

    NOMBRE_MENU: "🏠MENU",


    // ----------------------------------------------------------
    // 01.02 SÍMBOLO DE OPCIÓN NAVEGABLE
    // ----------------------------------------------------------

    SIMBOLO_ENLACE: "↳"

  };



  // ============================================================
  // 02. ACTUALIZAR MENÚ
  // Recorre la columna A del 🏠MENU y crea los hipervínculos
  // hacia las hojas existentes del ERP.
  // ============================================================

  function ACTUALIZAR_MENU() {

    // ----------------------------------------------------------
    // 02.01 CONECTAR CON EL LIBRO
    // ----------------------------------------------------------

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    // ----------------------------------------------------------
    // 02.02 OBTENER HOJA DEL MENÚ
    // ----------------------------------------------------------

    const menu =
      ss.getSheetByName(
        MENU_CONFIG.NOMBRE_MENU
      );


    // ----------------------------------------------------------
    // 02.03 VALIDAR HOJA DEL MENÚ
    // ----------------------------------------------------------

    if (!menu) {

      throw new Error(
        "No existe la hoja " +
        MENU_CONFIG.NOMBRE_MENU + "."
      );

    }


    // ----------------------------------------------------------
    // 02.04 OBTENER ÚLTIMA FILA
    // ----------------------------------------------------------

    const ultimaFila =
      menu.getLastRow();


    if (ultimaFila < 1) {

      throw new Error(
        "La hoja " +
        MENU_CONFIG.NOMBRE_MENU +
        " está vacía."
      );

    }


    // ----------------------------------------------------------
    // 02.05 VARIABLES DE CONTROL
    // ----------------------------------------------------------

    let enlaces = 0;

    const noEncontradas = [];


    // ----------------------------------------------------------
    // 02.06 RECORRER EL MENÚ
    // ----------------------------------------------------------

    for (
      let fila = 1;
      fila <= ultimaFila;
      fila++
    ) {

      const celda =
        menu.getRange(
          fila,
          1
        );


      // --------------------------------------------------------
      // 02.06.01 OBTENER TEXTO DE LA CELDA
      // --------------------------------------------------------

      const texto =
        celda.getDisplayValue();


      // --------------------------------------------------------
      // 02.06.02 IGNORAR CELDAS VACÍAS
      // --------------------------------------------------------

      if (!texto) {

        continue;

      }


      // --------------------------------------------------------
      // 02.06.03 IDENTIFICAR OPCIONES NAVEGABLES
      // Solo procesa las filas que contienen ↳
      // --------------------------------------------------------

      if (
        !texto.includes(
          MENU_CONFIG.SIMBOLO_ENLACE
        )
      ) {

        continue;

      }


      // --------------------------------------------------------
      // 02.06.04 OBTENER NOMBRE DE LA HOJA DESTINO
      // --------------------------------------------------------

      const nombreHoja =
        OBTENER_NOMBRE_HOJA_MENU(
          texto
        );


      if (!nombreHoja) {

        continue;

      }


      // --------------------------------------------------------
      // 02.06.05 BUSCAR HOJA DESTINO
      // --------------------------------------------------------

      const hojaDestino =
        ss.getSheetByName(
          nombreHoja
        );


      // --------------------------------------------------------
      // 02.06.06 VALIDAR EXISTENCIA DE LA HOJA
      // --------------------------------------------------------

      if (!hojaDestino) {

        noEncontradas.push(
          nombreHoja
        );

        continue;

      }


      // --------------------------------------------------------
      // 02.06.07 OBTENER ID DE LA HOJA
      // --------------------------------------------------------

      const gid =
        hojaDestino.getSheetId();


      // --------------------------------------------------------
      // 02.06.08 CONSTRUIR URL DE NAVEGACIÓN
      // --------------------------------------------------------

      const url =
        ss.getUrl() +
        "#gid=" +
        gid;


      // --------------------------------------------------------
      // 02.06.09 CREAR ENLACE EN LA CELDA
      // Se utiliza RichTextValue para evitar problemas con
      // configuración regional de fórmulas HYPERLINK.
      // --------------------------------------------------------

      const textoVisible =
        MENU_CONFIG.SIMBOLO_ENLACE +
        " " +
        nombreHoja;


      const enlace =
        SpreadsheetApp
          .newRichTextValue()
          .setText(
            textoVisible
          )
          .setLinkUrl(
            url
          )
          .build();


      // --------------------------------------------------------
      // 02.06.10 ESCRIBIR HIPERVÍNCULO
      // --------------------------------------------------------

      celda.setRichTextValue(
        enlace
      );


      enlaces++;

    }


    // ----------------------------------------------------------
    // 02.07 APLICAR CAMBIOS
    // ----------------------------------------------------------

    SpreadsheetApp.flush();


    // ----------------------------------------------------------
    // 02.08 REGISTRAR RESULTADO
    // ----------------------------------------------------------

    REGISTRAR_RESULTADO_MENU(
      enlaces,
      noEncontradas
    );

  }



  // ============================================================
  // 03. OBTENER NOMBRE DE HOJA DEL MENÚ
  // Elimina el símbolo ↳ y devuelve el nombre exacto de la hoja.
  // ============================================================

  function OBTENER_NOMBRE_HOJA_MENU(
    texto
  ) {

    // ----------------------------------------------------------
    // 03.01 VALIDAR TEXTO
    // ----------------------------------------------------------

    if (!texto) {

      return "";

    }


    // ----------------------------------------------------------
    // 03.02 CONVERTIR A TEXTO
    // ----------------------------------------------------------

    let nombre =
      String(texto);


    // ----------------------------------------------------------
    // 03.03 ELIMINAR SÍMBOLO DE NAVEGACIÓN
    // ----------------------------------------------------------

    nombre =
      nombre.replace(
        /↳/g,
        ""
      );


    // ----------------------------------------------------------
    // 03.04 ELIMINAR ESPACIOS EXTERNOS
    // ----------------------------------------------------------

    nombre =
      nombre.trim();


    // ----------------------------------------------------------
    // 03.05 NORMALIZAR ESPACIOS
    // ----------------------------------------------------------

    nombre =
      nombre.replace(
        /\s+/g,
        " "
      );


    return nombre;

  }



  // ============================================================
  // 04. REGISTRAR RESULTADO
  // Muestra en el registro el resultado de la actualización.
  // ============================================================

  function REGISTRAR_RESULTADO_MENU(
    enlaces,
    noEncontradas
  ) {

    console.log(
      "================================="
    );

    console.log(
      "🔗 ACTUALIZACIÓN DEL MENÚ ERP"
    );

    console.log(
      "================================="
    );

    console.log(
      "Enlaces creados o actualizados: " +
      enlaces
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

      console.log(
        noEncontradas.join(
          ", "
        )
      );

    }


    console.log(
      "================================="
    );

  }