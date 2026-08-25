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
  // Nombre de la hoja principal de navegación
  NOMBRE_MENU: "🏠MENU",
  
  // Símbolo que identifica una opción navegable
  SIMBOLO_ENLACE: "↳"
};

// ============================================================
// 02. ACTUALIZAR MENÚ
// Recorre la columna A del 🏠MENU y crea los hipervínculos
// hacia las hojas existentes del ERP.
// ============================================================
function ACTUALIZAR_MENU() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener la hoja del menú
  const menu = ss.getSheetByName(MENU_CONFIG.NOMBRE_MENU);
  if (!menu) {
    throw new Error("No existe la hoja " + MENU_CONFIG.NOMBRE_MENU + ".");
  }
  
  const ultimaFila = menu.getLastRow();
  if (ultimaFila < 1) {
    throw new Error("La hoja " + MENU_CONFIG.NOMBRE_MENU + " está vacía.");
  }
  
  let enlaces = 0;
  const noEncontradas = [];
  
  // Recorrer la columna A del menú
  for (let fila = 1; fila <= ultimaFila; fila++) {
    const celda = menu.getRange(fila, 1);
    const texto = celda.getDisplayValue();
    
    // Ignorar celdas vacías
    if (!texto) {
      continue;
    }
    
    // Identificar opciones navegables
    if (!texto.includes(MENU_CONFIG.SIMBOLO_ENLACE)) {
      continue;
    }
    
    // Obtener el nombre real de la hoja destino
    const nombreHoja = OBTENER_NOMBRE_HOJA_MENU(texto);
    if (!nombreHoja) {
      continue;
    }
    
    // Buscar la hoja de destino en el libro activo
    const hojaDestino = ss.getSheetByName(nombreHoja);
    if (!hojaDestino) {
      noEncontradas.push(nombreHoja);
      continue;
    }
    
    // Obtener el ID único de la pestaña (gid)
    const gid = hojaDestino.getSheetId();
    
    // Construir la URL interna de navegación
    const url = ss.getUrl() + "#gid=" + gid;
    
    // Crear el hipervínculo utilizando RichTextValue para máxima compatibilidad regional
    const textoVisible = MENU_CONFIG.SIMBOLO_ENLACE + " " + nombreHoja;
    const enlace = SpreadsheetApp.newRichTextValue()
      .setText(textoVisible)
      .setLinkUrl(url)
      .build();
    
    celda.setRichTextValue(enlace);
    enlaces++;
  }
  
  SpreadsheetApp.flush();
  REGISTRAR_RESULTADO_MENU(enlaces, noEncontradas);
}

// ============================================================
// 03. OBTENER NOMBRE DE HOJA DEL MENÚ
// Elimina el símbolo ↳ y los espacios extras para extraer el nombre
// ============================================================
function OBTENER_NOMBRE_HOJA_MENU(texto) {
  if (!texto) {
    return "";
  }
  
  let nombre = String(texto);
  nombre = nombre.replace(/↳/g, ""); // Quitar símbolo de navegación
  nombre = nombre.trim();             // Eliminar espacios extremos
  nombre = nombre.replace(/\s+/g, " "); // Normalizar espacios internos
  
  return nombre;
}

// ============================================================
// 04. REGISTRAR RESULTADO DE LA NAVEGACIÓN
// Muestra en el registro el resultado detallado de la sincronización.
// ============================================================
function REGISTRAR_RESULTADO_MENU(enlaces, noEncontradas) {
  console.log("=================================");
  console.log("🔗 ACTUALIZACIÓN DEL MENÚ ERP");
  console.log("=================================");
  console.log("Enlaces creados o actualizados: " + enlaces);
  console.log("Hojas no encontradas: " + noEncontradas.length);
  
  if (noEncontradas.length > 0) {
    console.log("Hojas no encontradas:");
    console.log(noEncontradas.join(", "));
  }
  console.log("=================================");
}