// (VERSIÓN 2.0 - V2 ERP - PROVISIONADOR MULTIEMPRESA)
/**************************************************************
* 30_PROVISIONADOR_MULTIEMPRESA.gs
* RESPONSABILIDAD:
* - Módulo SuperAdministrador Multi-Tenant para el despliegue de ERPs independientes por empresa.
* - Clonación automatizada de la plantilla maestra de Google Sheets (DriveApp.makeCopy).
* - Auto-configuración de la hoja CFG_EMPRESA y siembra de usuario Administrador de fábrica.
* - Registro y auditoría centralizada en la tabla de control maestro EMP_EMPRESAS_MAESTRO.
**************************************************************/

const PROV_MULTI_CONFIG = {
  HOJA_CONTROL_EMPRESAS: "EMP_EMPRESAS_MAESTRO",
  NOMBRE_CARPETA_TENANTS: "ERP_EMPRESAS_CLIENTES",
  PREFIJO_EMPRESA: "EMP",
  DIGITOS_ID: 6
};

/**
 * Registra o recupera la carpeta raíz en Google Drive donde se alojan los ERPs independientes
 */
function PROV_OBTENER_CARPETA_MATRIZ_DRIVE() {
  const carpetas = DriveApp.getFoldersByName(PROV_MULTI_CONFIG.NOMBRE_CARPETA_TENANTS);
  if (carpetas.hasNext()) {
    return carpetas.next();
  }
  return DriveApp.createFolder(PROV_MULTI_CONFIG.NOMBRE_CARPETA_TENANTS);
}

/**
 * Inicializa la tabla central de control de Tenants/Empresas en la hoja actual
 */
function PROV_INICIALIZAR_TABLA_MAESTRA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(PROV_MULTI_CONFIG.HOJA_CONTROL_EMPRESAS);
  if (!hoja) {
    hoja = ss.insertSheet(PROV_MULTI_CONFIG.HOJA_CONTROL_EMPRESAS);
    hoja.setTabColor("#0f172a");
    hoja.appendRow([
      "ID_EMPRESA", "NIT", "DV", "RAZON_SOCIAL", "NOMBRE_COMERCIAL", 
      "CORREO_ADMINISTRADOR", "TELEFONO", "CIUDAD", "SPREADSHEET_ID", 
      "SPREADSHEET_URL", "ESTADO_TENANT", "FECHA_PROVISION", "USUARIO_CREACION", "OBSERVACIONES"
    ]);
    hoja.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
  }
  return hoja;
}

/**
 * RPC: Listar todas las empresas registradas en el Panel de SuperAdmin
 */
function PROV_LISTAR_EMPRESAS_REGISTRADAS_WEB(tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ADMINISTRAR");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = PROV_INICIALIZAR_TABLA_MAESTRA();
    const ulFila = hoja.getLastRow();
    
    if (ulFila < 2) {
      return { EXITO: true, DATOS: [], MENSAJE: "No hay empresas registradas en el provisionador." };
    }
    
    const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    const reg = hoja.getRange(2, 1, ulFila - 1, enc.length).getValues();
    
    const lista = reg.map(fila => {
      const obj = {};
      enc.forEach((col, idx) => { obj[col] = fila[idx] !== undefined ? fila[idx] : ""; });
      return obj;
    });
    
    return {
      EXITO: true,
      DATOS: typeof SEG_SANITIZAR_PARA_CLIENTE === "function" ? SEG_SANITIZAR_PARA_CLIENTE(lista) : lista,
      MENSAJE: "Empresas listadas correctamente."
    };
  } catch (error) {
    return { EXITO: false, DATOS: [], MENSAJE: "Error al listar empresas: " + error.toString() };
  }
}

/**
 * RPC: Desplegar y provisionar una nueva empresa independiente con su propio Google Sheet ERP
 */
function PROV_CREAR_NUEVA_EMPRESA_ERP_WEB(datos, tokenSesion) {
  try {
    if (typeof SEG_VERIFICAR_CONTEXTO_Y_ACCESO === "function" && tokenSesion && tokenSesion !== "SHEETS_CONTEXT") {
      SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "SEGURIDAD", "ADMINISTRAR");
    }
    
    if (!datos || !datos.NIT || !datos.RAZON_SOCIAL || !datos.CORREO_ADMINISTRADOR) {
      throw new Error("El NIT, la Razón Social y el Correo del Administrador son obligatorios.");
    }
    
    const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
    const idTemplateMatriz = ssOrigen.getId(); // Toma la hoja actual como plantilla base
    
    const hojaMaestra = PROV_INICIALIZAR_TABLA_MAESTRA();
    const encMaestra = hojaMaestra.getRange(1, 1, 1, hojaMaestra.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
    
    // 1. Validar si el NIT ya tiene una empresa provisionada
    const idxNit = encMaestra.indexOf("NIT");
    if (idxNit !== -1 && hojaMaestra.getLastRow() >= 2) {
      const nitsExistentes = hojaMaestra.getRange(2, idxNit + 1, hojaMaestra.getLastRow() - 1, 1).getValues().flat();
      const nitLimpio = String(datos.NIT).replace(/\D/g, "");
      const existe = nitsExistentes.some(n => String(n).replace(/\D/g, "") === nitLimpio);
      if (existe) {
        throw new Error("Ya existe una empresa registrada en la plataforma con el NIT " + datos.NIT);
      }
    }
    
    // 2. Generar ID_EMPRESA único
    const idEmpresa = PROV_MULTI_CONFIG.PREFIJO_EMPRESA + "-" + String(Math.max(1, hojaMaestra.getLastRow())).padStart(PROV_MULTI_CONFIG.DIGITOS_ID, "0");
    
    // 3. Clonar la plantilla del ERP en la carpeta de empresas en Google Drive
    const carpetaDrive = PROV_OBTENER_CARPETA_MATRIZ_DRIVE();
    const archivoMatriz = DriveApp.getFileById(idTemplateMatriz);
    const nombreNuevoERP = "ERP Empresarial - " + String(datos.RAZON_SOCIAL).trim().toUpperCase() + " (" + datos.NIT + ")";
    
    const copiaArchivo = archivoMatriz.makeCopy(nombreNuevoERP, carpetaDrive);
    copiaArchivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const idNuevoSheet = copiaArchivo.getId();
    const urlNuevoSheet = copiaArchivo.getUrl();
    const ssNuevo = SpreadsheetApp.openById(idNuevoSheet);
    
    // 4. Configurar la hoja CFG_EMPRESA en el nuevo ERP independiente
    PROV_CONFIGURAR_HOJA_CFG_EMPRESA_NUEVA(ssNuevo, idEmpresa, datos);
    
    // 5. Configurar el usuario Administrador inicial en USR_USUARIOS del nuevo ERP
    PROV_CONFIGURAR_ADMIN_INICIAL_NUEVO_ERP(ssNuevo, datos);
    
    // 6. Registrar el nuevo Tenant en la tabla maestra central de control
    const ahora = new Date();
    const usuarioCreador = (typeof Session !== "undefined" && Session.getActiveUser()) ? (Session.getActiveUser().getEmail() || "SUPERADMIN") : "SUPERADMIN";
    
    const registroFila = encMaestra.map(col => {
      switch (col) {
        case "ID_EMPRESA": return idEmpresa;
        case "NIT": return datos.NIT;
        case "DV": return datos.DV || "";
        case "RAZON_SOCIAL": return String(datos.RAZON_SOCIAL).trim().toUpperCase();
        case "NOMBRE_COMERCIAL": return datos.NOMBRE_COMERCIAL || "";
        case "CORREO_ADMINISTRADOR": return String(datos.CORREO_ADMINISTRADOR).trim().toLowerCase();
        case "TELEFONO": return datos.TELEFONO || "";
        case "CIUDAD": return datos.CIUDAD || "Pitalito";
        case "SPREADSHEET_ID": return idNuevoSheet;
        case "SPREADSHEET_URL": return urlNuevoSheet;
        case "ESTADO_TENANT": return "ACTIVA";
        case "FECHA_PROVISION": return ahora;
        case "USUARIO_CREACION": return usuarioCreador;
        case "OBSERVACIONES": return datos.OBSERVACIONES || "Provisionamiento automático SaaS Multi-Empresa";
        default: return "";
      }
    });
    
    hojaMaestra.appendRow(registroFila);
    
    if (typeof SEG_REGISTRAR_AUDITORIA === "function") {
      SEG_REGISTRAR_AUDITORIA({
        ID_USUARIO: usuarioCreador,
        USUARIO: usuarioCreador,
        MODULO: "SEGURIDAD",
        SUBMODULO: "MULTI_EMPRESA",
        ACCION: "CREAR",
        TIPO_REGISTRO: "TENANT",
        ID_REGISTRO: idEmpresa,
        DESCRIPCION: "Empresa '" + datos.RAZON_SOCIAL + "' provisionada con éxito. Sheet ID: " + idNuevoSheet,
        RESULTADO: "EXITOSO"
      });
    }
    
    return {
      EXITO: true,
      ID_EMPRESA: idEmpresa,
      SPREADSHEET_ID: idNuevoSheet,
      SPREADSHEET_URL: urlNuevoSheet,
      MENSAJE: "¡ERP Independiente para '" + datos.RAZON_SOCIAL + "' provisionado y configurado exitosamente!"
    };
  } catch (error) {
    return { EXITO: false, MENSAJE: "Error al provisionar empresa: " + error.toString() };
  }
}

/**
 * Escribe la configuración propia de la nueva empresa en la hoja CFG_EMPRESA del Sheet clonado
 */
function PROV_CONFIGURAR_HOJA_CFG_EMPRESA_NUEVA(ssNuevo, idEmpresa, datos) {
  const hojaCfg = ssNuevo.getSheetByName("CFG_EMPRESA");
  if (!hojaCfg) return;
  
  const valoresCfg = hojaCfg.getRange(1, 1, Math.max(1, hojaCfg.getLastRow()), 2).getValues();
  const mapaValores = {
    "ID_EMPRESA": idEmpresa,
    "RAZON_SOCIAL": String(datos.RAZON_SOCIAL || "").trim().toUpperCase(),
    "NOMBRE_COMERCIAL": String(datos.NOMBRE_COMERCIAL || datos.RAZON_SOCIAL || "").trim().toUpperCase(),
    "NIT": String(datos.NIT || "").trim(),
    "DIGITO_VERIFICACION": String(datos.DV || ""),
    "CORREO": String(datos.CORREO_ADMINISTRADOR || "").trim().toLowerCase(),
    "CORREO_ADMINISTRATIVO": String(datos.CORREO_ADMINISTRADOR || "").trim().toLowerCase(),
    "CORREO_FACTURACION": String(datos.CORREO_ADMINISTRADOR || "").trim().toLowerCase(),
    "TELEFONO": String(datos.TELEFONO || "").trim(),
    "CIUDAD": String(datos.CIUDAD || "Pitalito").trim(),
    "DEPARTAMENTO": String(datos.DEPARTAMENTO || "Huila").trim(),
    "DIRECCION": String(datos.DIRECCION || "").trim(),
    "ESTADO": "ACTIVA",
    "FECHA_ACTUALIZACION": new Date()
  };
  
  for (let i = 0; i < valoresCfg.length; i++) {
    const campoKey = String(valoresCfg[i][0] || "").trim().toUpperCase();
    if (mapaValores[campoKey] !== undefined) {
      hojaCfg.getRange(i + 1, 2).setValue(mapaValores[campoKey]);
    }
  }
}

/**
 * Registra al administrador titular en la hoja USR_USUARIOS del nuevo Sheet clonado
 */
function PROV_CONFIGURAR_ADMIN_INICIAL_NUEVO_ERP(ssNuevo, datos) {
  const hojaUsr = ssNuevo.getSheetByName("USR_USUARIOS");
  if (!hojaUsr) return;
  
  const enc = hojaUsr.getRange(1, 1, 1, hojaUsr.getLastColumn()).getValues()[0].map(h => String(h || "").trim().toUpperCase());
  const idxCorreo = enc.indexOf("CORREO");
  const idxUsr = enc.indexOf("USUARIO");
  
  const correoAdmin = String(datos.CORREO_ADMINISTRADOR || "").trim().toLowerCase();
  const nomUsuario = correoAdmin.split("@")[0].replace(/[^a-zA-Z0-9_]/g, "").toUpperCase();
  
  // Limpiar usuarios de prueba si existieran en la plantilla
  if (hojaUsr.getLastRow() >= 2) {
    hojaUsr.getRange(2, 1, hojaUsr.getLastRow() - 1, enc.length).clearContent();
  }

  let hashPass = "";
  if (typeof SEG_GENERAR_HASH_CONTRASENA === "function") {
    hashPass = SEG_GENERAR_HASH_CONTRASENA("Admin123!");
  } else {
    hashPass = "3eb3fe66b31e3b4d10fa70b5cad49c7112294af6ae4e476a1c405155d45aa121";
  }

  const ahora = new Date();
  
  const registroAdmin = enc.map(col => {
    switch (col) {
      case "ID_USUARIO": return "USR-000001";
      case "USUARIO": return nomUsuario || "ADMIN";
      case "NOMBRE_COMPLETO": return String(datos.NOMBRE_ADMINISTRADOR || "Administrador " + datos.RAZON_SOCIAL).trim();
      case "CORREO": return correoAdmin;
      case "CONTRASENA_HASH": return hashPass;
      case "ID_ROL": return "ROL-000001";
      case "ESTADO_USUARIO": return "ACTIVO";
      case "INTENTOS_FALLIDOS": return 0;
      case "FECHA_CREACION": return ahora;
      case "FECHA_ACTUALIZACION": return ahora;
      case "USUARIO_CREACION": return "PROVISIONADOR_SaaS";
      default: return "";
    }
  });
  
  hojaUsr.appendRow(registroAdmin);
}
