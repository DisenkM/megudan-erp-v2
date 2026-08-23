// ============================================================
// CLIENTES.gs
// MÓDULO: CLIENTES
// ============================================================
//
// FUNCIONES:
// 01. Configuración del módulo
// 02. Abrir formulario
// 03. Generar ID
// 04. Validar duplicados
// 05. Obtener usuario y fecha
// 06. Guardar cliente
// 07. Buscar cliente
// 08. Actualizar cliente
// 09. Inactivar cliente
// 10. Registrar historial
//
// HOJAS RELACIONADAS:
// - CLI_MAESTRO
// - CLI_HISTORIAL
//
// INTERFAZ:
// - CLI_FORM.html
// ============================================================


// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO
// ============================================================

const CLI_CONFIG = {

  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  FORMULARIO: "F1_CLI_FORM",

  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6

};


// ============================================================
// 02. ABRIR FORMULARIO DE CLIENTES
// ============================================================

function ABRIR_CLIENTES() {

  const html = HtmlService
    .createHtmlOutputFromFile(CLI_CONFIG.FORMULARIO)
    .setWidth(1200)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    "👥 Gestión de Clientes"
  );

}


// ============================================================
// 03. GENERAR NUEVO ID_CLIENTE
// ============================================================

function CLI_GENERAR_ID() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hoja = ss.getSheetByName(
    CLI_CONFIG.HOJA_MAESTRO
  );

  if (!hoja) {
    throw new Error(
      "No existe la hoja " + CLI_CONFIG.HOJA_MAESTRO
    );
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return CLI_CONFIG.PREFIJO_ID + "-000001";
  }

  const ids = hoja
    .getRange(2, 1, ultimaFila - 1, 1)
    .getValues()
    .flat();

  let mayorNumero = 0;

  ids.forEach(function(id) {

    if (!id) return;

    const partes = String(id).split("-");
    const numero = Number(partes[1]);

    if (!isNaN(numero) && numero > mayorNumero) {
      mayorNumero = numero;
    }

  });

  const siguienteNumero = mayorNumero + 1;

  return (
    CLI_CONFIG.PREFIJO_ID +
    "-" +
    String(siguienteNumero).padStart(
      CLI_CONFIG.DIGITOS_ID,
      "0"
    )
  );

}


// ============================================================
// 04. VALIDAR CLIENTE DUPLICADO
// ============================================================

function CLI_VALIDAR_DUPLICADO(
  tipoDocumento,
  numeroDocumento,
  idClienteActual
) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hoja = ss.getSheetByName(
    CLI_CONFIG.HOJA_MAESTRO
  );

  if (!hoja) {
    throw new Error("No existe la hoja CLI_MAESTRO.");
  }

  if (hoja.getLastRow() < 2) {
    return false;
  }

  const encabezados = hoja
    .getRange(1, 1, 1, hoja.getLastColumn())
    .getValues()[0];

  const indiceId =
    encabezados.indexOf("ID_CLIENTE");

  const indiceTipo =
    encabezados.indexOf("TIPO_DOCUMENTO");

  const indiceNumero =
    encabezados.indexOf("NUMERO_DOCUMENTO");

  if (
    indiceId === -1 ||
    indiceTipo === -1 ||
    indiceNumero === -1
  ) {
    throw new Error(
      "Faltan columnas requeridas en CLI_MAESTRO."
    );
  }

  const datos = hoja
    .getRange(
      2,
      1,
      hoja.getLastRow() - 1,
      hoja.getLastColumn()
    )
    .getValues();

  return datos.some(function(fila) {

    const esMismoCliente =
      String(fila[indiceId]) ===
      String(idClienteActual || "");

    if (esMismoCliente) {
      return false;
    }

    return (
      String(fila[indiceTipo]) ===
      String(tipoDocumento) &&
      String(fila[indiceNumero]) ===
      String(numeroDocumento)
    );

  });

}


// ============================================================
// 05. UTILIDADES DEL MÓDULO
// ============================================================

function CLI_OBTENER_USUARIO_ACTUAL() {

  return (
    Session.getActiveUser().getEmail() ||
    "USUARIO_SISTEMA"
  );

}


function CLI_OBTENER_FECHA_ACTUAL() {
  return new Date();
}


// ============================================================
// 06. GUARDAR NUEVO CLIENTE
// ============================================================

function CLI_GUARDAR_CLIENTE(datos) {

  if (!datos) {
    throw new Error(
      "No se recibieron datos del cliente."
    );
  }

  const obligatorios = [
    "TIPO_PERSONA",
    "TIPO_DOCUMENTO",
    "NUMERO_DOCUMENTO",
    "TIPO_CLIENTE",
    "PAIS"
  ];

  obligatorios.forEach(function(campo) {

    if (
      !datos[campo] ||
      String(datos[campo]).trim() === ""
    ) {
      throw new Error(
        "El campo " + campo + " es obligatorio."
      );
    }

  });

  const duplicado = CLI_VALIDAR_DUPLICADO(
    datos.TIPO_DOCUMENTO,
    datos.NUMERO_DOCUMENTO,
    null
  );

  if (duplicado) {
    throw new Error(
      "Ya existe un cliente con este documento."
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hoja = ss.getSheetByName(
    CLI_CONFIG.HOJA_MAESTRO
  );

  if (!hoja) {
    throw new Error("No existe la hoja CLI_MAESTRO.");
  }

  const idCliente = CLI_GENERAR_ID();
  const fecha = CLI_OBTENER_FECHA_ACTUAL();
  const usuario = CLI_OBTENER_USUARIO_ACTUAL();

  datos.ID_CLIENTE = idCliente;
  datos.FECHA_CREACION = fecha;
  datos.FECHA_ACTUALIZACION = fecha;
  datos.USUARIO_CREACION = usuario;
  datos.USUARIO_ACTUALIZACION = usuario;

  const encabezados = hoja
    .getRange(1, 1, 1, hoja.getLastColumn())
    .getValues()[0];

  const nuevaFila = encabezados.map(function(campo) {
    return datos[campo] !== undefined
      ? datos[campo]
      : "";
  });

  hoja.appendRow(nuevaFila);

  CLI_REGISTRAR_HISTORIAL(
    idCliente,
    "CREACION",
    "CREAR",
    "",
    "",
    "",
    "CLIENTES",
    idCliente,
    usuario,
    "ACTIVO",
    "Cliente creado correctamente."
  );

  return {
    ok: true,
    mensaje: "Cliente registrado correctamente.",
    idCliente: idCliente
  };

}

// =============================================================
// 07. BUSCAR CLIENTE
// =============================================================
// FUNCIÓN:
// Busca un cliente registrado en CLI_MAESTRO utilizando:
//
// - ID_CLIENTE
// - NIT
// - CC
// - Cualquier número de documento registrado
// - RAZÓN SOCIAL
//
// CRITERIOS DE BÚSQUEDA:
//
// ID_CLIENTE          → Coincidencia exacta
// NUMERO_DOCUMENTO    → Coincidencia exacta
// RAZON_SOCIAL        → Coincidencia parcial
// ==============================================================

  function CLI_BUSCAR_CLIENTE(criterio) {


    // ==========================================================
    // 07.01 VALIDAR CRITERIO DE BÚSQUEDA
    // ==========================================================

    criterio = String(criterio || "").trim();

    if (!criterio) {

      throw new Error(
        "Ingrese un ID de cliente, NIT, CC o razón social."
      );

    }


    // ==========================================================
    // 07.02 OBTENER HOJA MAESTRA DE CLIENTES
    // ==========================================================

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hoja = ss.getSheetByName(
      CLI_CONFIG.HOJA_MAESTRO
    );

    if (!hoja) {

      throw new Error(
        "No existe la hoja " + CLI_CONFIG.HOJA_MAESTRO + "."
      );

    }


    // ==========================================================
    // 07.03 VALIDAR QUE EXISTAN CLIENTES REGISTRADOS
    // ==========================================================

    const ultimaFila = hoja.getLastRow();
    const ultimaColumna = hoja.getLastColumn();

    if (ultimaFila < 2) {

      return null;

    }


    // ==========================================================
    // 07.04 OBTENER ENCABEZADOS DE CLI_MAESTRO
    // ==========================================================
    // Se convierten a mayúsculas para evitar problemas de
    // comparación entre los nombres de las columnas.
    // ==========================================================

    const encabezados = hoja
      .getRange(1, 1, 1, ultimaColumna)
      .getDisplayValues()[0]
      .map(function(encabezado) {

        return String(encabezado)
          .trim()
          .toUpperCase();

      });


    // ==========================================================
    // 07.05 OBTENER REGISTROS DE CLIENTES
    // ==========================================================
    // Se utiliza getDisplayValues() para conservar los valores
    // tal como se muestran en Google Sheets.
    // ==========================================================

    const registros = hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        ultimaColumna
      )
      .getDisplayValues();


    // ==========================================================
    // 07.06 NORMALIZAR TEXTO
    // ==========================================================
    // Convierte a mayúsculas y elimina espacios innecesarios.
    // ==========================================================

    function normalizarTexto(valor) {

      return String(valor || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, " ");

    }


    // ==========================================================
    // 07.07 NORMALIZAR DOCUMENTO
    // ==========================================================
    // Permite encontrar el mismo documento aunque esté escrito:
    //
    // 123456789
    // 123.456.789
    // 123 456 789
    // 123-456-789
    // ==========================================================

    function normalizarDocumento(valor) {

      return String(valor || "")
        .trim()
        .replace(/[.\s,\-]/g, "");

    }


    // ==========================================================
    // 07.08 PREPARAR CRITERIOS DE BÚSQUEDA
    // ==========================================================

    const criterioTexto = normalizarTexto(criterio);

    const criterioDocumento = normalizarDocumento(criterio);


    // ==========================================================
    // 07.09 IDENTIFICAR COLUMNAS NECESARIAS
    // ==========================================================

    const idxIdCliente =
      encabezados.indexOf("ID_CLIENTE");

    const idxDocumento =
      encabezados.indexOf("NUMERO_DOCUMENTO");

    const idxRazonSocial =
      encabezados.indexOf("RAZON_SOCIAL");


    // ==========================================================
    // 07.10 VALIDAR ESTRUCTURA DE CLI_MAESTRO
    // ==========================================================

    if (idxIdCliente === -1) {

      throw new Error(
        "No se encontró la columna ID_CLIENTE en CLI_MAESTRO."
      );

    }

    if (idxDocumento === -1) {

      throw new Error(
        "No se encontró la columna NUMERO_DOCUMENTO en CLI_MAESTRO."
      );

    }

    if (idxRazonSocial === -1) {

      throw new Error(
        "No se encontró la columna RAZON_SOCIAL en CLI_MAESTRO."
      );

    }


    // ==========================================================
    // 07.11 BUSCAR CLIENTE
    // ==========================================================
    // PRIORIDAD DE BÚSQUEDA:
    //
    // 1. ID_CLIENTE       → Coincidencia exacta
    // 2. NIT / CC         → Coincidencia exacta
    // 3. RAZÓN SOCIAL     → Coincidencia parcial
    // ==========================================================

    const filaEncontrada = registros.find(function(fila) {


      // ========================================================
      // 07.11.01 OBTENER VALORES DE LA FILA
      // ========================================================

      const idCliente =
        normalizarTexto(fila[idxIdCliente]);

      const documento =
        normalizarDocumento(fila[idxDocumento]);

      const razonSocial =
        normalizarTexto(fila[idxRazonSocial]);


      // ========================================================
      // 07.11.02 BUSCAR POR ID_CLIENTE
      // ========================================================
      // Ejemplo:
      //
      // CLI-000001
      // ========================================================

      if (idCliente === criterioTexto) {

        return true;

      }


      // ========================================================
      // 07.11.03 BUSCAR POR NIT, CC O DOCUMENTO
      // ========================================================
      // NIT y CC se encuentran en NUMERO_DOCUMENTO.
      //
      // La coincidencia es exacta para evitar encontrar
      // documentos incorrectos.
      // ========================================================

      if (documento === criterioDocumento) {

        return true;

      }


      // ========================================================
      // 07.11.04 BUSCAR POR RAZÓN SOCIAL
      // ========================================================
      // Permite búsquedas parciales.
      //
      // Ejemplo:
      //
      // CONSTRUCTORA ABC
      //
      // Encuentra:
      //
      // CONSTRUCTORA ABC S.A.S.
      // ========================================================

      if (razonSocial.includes(criterioTexto)) {

        return true;

      }


      // ========================================================
      // 07.11.05 NO EXISTE COINCIDENCIA
      // ========================================================

      return false;

    });


    // ==========================================================
    // 07.12 CLIENTE NO ENCONTRADO
    // ==========================================================

    if (!filaEncontrada) {

      return null;

    }


    // ==========================================================
    // 07.13 CONVERTIR FILA EN OBJETO
    // ==========================================================
    // Convierte la fila encontrada en un objeto utilizando los
    // encabezados de CLI_MAESTRO como propiedades.
    // ==========================================================

    const cliente = {};

    encabezados.forEach(function(campo, indice) {

      cliente[campo] =
        filaEncontrada[indice];

    });


    // ==========================================================
    // 07.14 DEVOLVER CLIENTE AL HTML
    // ==========================================================
    // El resultado es enviado a F1_CLI_FORM mediante:
    //
    // google.script.run
    //   .CLI_BUSCAR_CLIENTE(criterio)
    // ==========================================================

    return cliente;

  }


// ============================================================
// 08. ACTUALIZAR CLIENTE
// ============================================================

  function CLI_ACTUALIZAR_CLIENTE(datos) {

    if (!datos || !datos.ID_CLIENTE) {
      throw new Error(
        "No se indicó el ID del cliente."
      );
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hoja = ss.getSheetByName(
      CLI_CONFIG.HOJA_MAESTRO
    );

    if (!hoja || hoja.getLastRow() < 2) {
      throw new Error(
        "No existen clientes registrados."
      );
    }

    const encabezados = hoja
      .getRange(1, 1, 1, hoja.getLastColumn())
      .getValues()[0];

    const registros = hoja
      .getRange(
        2,
        1,
        hoja.getLastRow() - 1,
        hoja.getLastColumn()
      )
      .getValues();

    const indiceId = encabezados.indexOf("ID_CLIENTE");

    const indiceRegistro = registros.findIndex(function(fila) {

      return String(fila[indiceId]) ===
        String(datos.ID_CLIENTE);

    });

    if (indiceRegistro === -1) {
      throw new Error(
        "No se encontró el cliente."
      );
    }

    const datosAnteriores =
      registros[indiceRegistro];

    const duplicado = CLI_VALIDAR_DUPLICADO(
      datos.TIPO_DOCUMENTO,
      datos.NUMERO_DOCUMENTO,
      datos.ID_CLIENTE
    );

    if (duplicado) {
      throw new Error(
        "Ya existe otro cliente con este documento."
      );
    }

    const usuario =
      CLI_OBTENER_USUARIO_ACTUAL();

    datos.FECHA_ACTUALIZACION =
      CLI_OBTENER_FECHA_ACTUAL();

    datos.USUARIO_ACTUALIZACION =
      usuario;

    const indiceFechaCreacion =
      encabezados.indexOf("FECHA_CREACION");

    const indiceUsuarioCreacion =
      encabezados.indexOf("USUARIO_CREACION");

    if (indiceFechaCreacion !== -1) {
      datos.FECHA_CREACION =
        datosAnteriores[indiceFechaCreacion];
    }

    if (indiceUsuarioCreacion !== -1) {
      datos.USUARIO_CREACION =
        datosAnteriores[indiceUsuarioCreacion];
    }

    const nuevaFila = encabezados.map(
      function(campo, indice) {

        return datos[campo] !== undefined
          ? datos[campo]
          : datosAnteriores[indice];

      }
    );

    hoja
      .getRange(
        indiceRegistro + 2,
        1,
        1,
        nuevaFila.length
      )
      .setValues([nuevaFila]);

    encabezados.forEach(function(campo, indice) {

      if (
        String(datosAnteriores[indice]) !==
        String(nuevaFila[indice]) &&
        campo !== "FECHA_ACTUALIZACION" &&
        campo !== "USUARIO_ACTUALIZACION"
      ) {

        CLI_REGISTRAR_HISTORIAL(
          datos.ID_CLIENTE,
          "MODIFICACION",
          "EDITAR",
          campo,
          datosAnteriores[indice],
          nuevaFila[indice],
          "CLIENTES",
          datos.ID_CLIENTE,
          usuario,
          "ACTIVO",
          "Campo modificado: " + campo
        );

      }

    });

    return {
      ok: true,
      mensaje: "Cliente actualizado correctamente.",
      idCliente: datos.ID_CLIENTE
    };

  }


// ============================================================
// 09. INACTIVAR CLIENTE
// ============================================================

function CLI_INACTIVAR_CLIENTE(idCliente) {

  const cliente = CLI_BUSCAR_CLIENTE(idCliente);

  if (!cliente) {
    throw new Error("Cliente no encontrado.");
  }

  if (cliente.ESTADO_CLIENTE === "INACTIVO") {
    throw new Error("El cliente ya está inactivo.");
  }

  cliente.ESTADO_CLIENTE = "INACTIVO";

  CLI_ACTUALIZAR_CLIENTE(cliente);

  return {
    ok: true,
    mensaje: "Cliente inactivado correctamente."
  };

}


// ============================================================
// 10. REGISTRAR HISTORIAL Y AUDITORÍA
// ============================================================

function CLI_REGISTRAR_HISTORIAL(
  idCliente,
  tipoEvento,
  accion,
  campoModificado,
  valorAnterior,
  valorNuevo,
  moduloOrigen,
  idRegistroOrigen,
  usuario,
  estadoEvento,
  observaciones
) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hoja = ss.getSheetByName(
    CLI_CONFIG.HOJA_HISTORIAL
  );

  if (!hoja) {
    throw new Error(
      "No existe la hoja CLI_HISTORIAL."
    );
  }

  const consecutivo =
    Math.max(hoja.getLastRow() - 1, 0) + 1;

  const evento = {

    ID_HISTORIAL:
      "HIS-" + String(consecutivo).padStart(6, "0"),

    ID_CLIENTE:
      idCliente,

    TIPO_EVENTO:
      tipoEvento,

    FECHA_HORA:
      CLI_OBTENER_FECHA_ACTUAL(),

    USUARIO:
      usuario,

    ACCION:
      accion,

    CAMPO_MODIFICADO:
      campoModificado,

    VALOR_ANTERIOR:
      valorAnterior,

    VALOR_NUEVO:
      valorNuevo,

    MOTIVO_ORIGEN:
      "",

    MODULO_ORIGEN:
      moduloOrigen,

    ID_REGISTRO_ORIGEN:
      idRegistroOrigen,

    IP_USUARIO:
      "",

    ESTADO_EVENTO:
      estadoEvento,

    OBSERVACIONES:
      observaciones

  };

  const encabezados = hoja
    .getRange(1, 1, 1, hoja.getLastColumn())
    .getValues()[0];

  const fila = encabezados.map(function(campo) {
    return evento[campo] !== undefined
      ? evento[campo]
      : "";
  });

  hoja.appendRow(fila);

}


// ============================================================
// 11. FIN DEL MÓDULO CLIENTES
// ============================================================