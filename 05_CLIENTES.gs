/**************************************************************
* 05_CLIENTES.gs
* RESPONSABILIDAD:
* - Administrar el ciclo de vida (CRUD) del maestro de Clientes.
* - Validar datos tributarios de Colombia (NIT, DV automático, responsabilidades fiscales).
* - Registrar automáticamente historial de cambios y auditorías operativas.
**************************************************************/

// ============================================================
// 01. CONFIGURACIÓN DEL MÓDULO DE CLIENTES
// ============================================================
const CLI_CONFIG = {
  HOJA_MAESTRO: "CLI_MAESTRO",
  HOJA_HISTORIAL: "CLI_HISTORIAL",
  FORMULARIO_HTML: "F1_CLI_FORM",
  PREFIJO_ID: "CLI",
  DIGITOS_ID: 6
};

// ============================================================
// 02. ALGORITMO OFICIAL DE LA DIAN: CALCULAR DV
// ============================================================
function CLI_CALCULAR_DV(nit) {
  if (!nit) return "";
  const tempNit = String(nit).replace(/[^0-9]/g, "");
  if (!tempNit || isNaN(tempNit)) return "";

  const pesos = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  let suma = 0;
  for (let i = 0; i < tempNit.length; i++) {
    const digito = parseInt(tempNit.charAt(tempNit.length - 1 - i), 10);
    suma += digito * pesos[i];
  }
  const residuo = suma % 11;
  if (residuo > 1) {
    return String(11 - residuo);
  }
  return String(residuo);
}

// ============================================================
// 03. VALIDACIÓN DE DUPLICADOS EN BASE DE DATOS
// ============================================================
function CLI_VALIDAR_DUPLICADO(tipoDoc, numDoc, idExcluir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja || hoja.getLastRow() < 2) return false;

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
  return datos.some(fila => {
    if (idExcluir && String(fila[0]) === String(idExcluir)) return false;
    return String(fila[2]) === String(tipoDoc) && String(fila[3]) === String(numDoc);
  });
}

// ============================================================
// 04. OPERACIONES CRUD: GUARDAR Y ACTUALIZAR CLIENTES
// ============================================================
function CLI_GUARDAR_CLIENTE(datos) {
  if (!datos) throw new Error("Datos no proporcionados.");

  // Forzar cálculo asíncrono seguro del DV para NITs
  if (datos.TIPO_DOCUMENTO === "NIT") {
    datos.DIGITO_VERIFICACION = CLI_CALCULAR_DV(datos.NUMERO_DOCUMENTO);
  } else {
    datos.DIGITO_VERIFICACION = "";
  }

  // Dinamismo del Tipo de Persona
  if (datos.TIPO_PERSONA === "PERSONA_NATURAL") {
    datos.RAZON_SOCIAL = [datos.PRIMER_NOMBRE, datos.SEGUNDO_NOMBRE, datos.PRIMER_APELLIDO, datos.SEGUNDO_APELLIDO]
      .filter(n => n && n.trim() !== "").join(" ");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  if (!hoja) throw new Error("Hoja CLI_MAESTRO no encontrada.");

  // Validar unicidad
  if (CLI_VALIDAR_DUPLICADO(datos.TIPO_DOCUMENTO, datos.NUMERO_DOCUMENTO, null)) {
    throw new Error("El NIT/CC ingresado ya pertenece a un cliente registrado.");
  }

  // Generar ID e inyectar auditorías
  const idCliente = CLI_OBTENER_SIGUIENTE_ID();
  const ahora = new Date();
  const usuario = Session.getActiveUser().getEmail() || "SISTEMA";

  datos.ID_CLIENTE = idCliente;
  datos.FECHA_CREACION = ahora;
  datos.FECHA_ACTUALIZACION = ahora;
  datos.USUARIO_CREACION = usuario;
  datos.USUARIO_ACTUALIZACION = usuario;

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const fila = encabezados.map(col => datos[col] !== undefined ? datos[col] : "");
  
  hoja.appendRow(fila);

  // Registrar Historial
  CLI_REGISTRAR_HISTORIAL(idCliente, "CREACION", "Cliente creado de forma exitosa.");

  return { ok: true, mensaje: "Cliente creado de forma exitosa.", idCliente: idCliente };
}

function CLI_OBTENER_SIGUIENTE_ID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_MAESTRO);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return CLI_CONFIG.PREFIJO_ID + "-000001";
  }
  const ultimoID = hoja.getRange(ultimaFila, 1).getValue().toString();
  const numero = parseInt(ultimoID.replace(CLI_CONFIG.PREFIJO_ID + "-", ""), 10);
  return CLI_CONFIG.PREFIJO_ID + "-" + String(numero + 1).padStart(CLI_CONFIG.DIGITOS_ID, "0");
}

function CLI_REGISTRAR_HISTORIAL(idCliente, accion, obs) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CLI_CONFIG.HOJA_HISTORIAL);
    if (!hoja) return;

    const fila = [
      "HIS-" + String(Math.max(1, hoja.getLastRow())).padStart(6, "0"),
      idCliente,
      new Date().toLocaleDateString(),
      new Date().toLocaleTimeString(),
      Session.getActiveUser().getEmail() || "SISTEMA",
      accion,
      "", "", "", "", "CLIENTES", idCliente, "", "ACTIVO", obs
    ];
    hoja.appendRow(fila);
  } catch (err) {
    console.error("No se pudo registrar log del cliente: " + idCliente);
  }
}
