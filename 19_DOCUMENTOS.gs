/**************************************************************
* 19_DOCUMENTOS.gs
* RESPONSABILIDAD:
* - Generar representaciones visuales e imprimibles (PDFs) de documentos del ERP.
* - Exportar e indexar PDFs de facturas, cotizaciones y órdenes de compra en Google Drive.
**************************************************************/

const DOC_CONFIG = {
  CARPETA_PDFS: "MEGUDAN_PDFS_AUDIT",
  FORMATO_DEFECTO: "PDF"
};

/**
 * Crea un PDF de forma automatizada y lo almacena en Google Drive.
 * @param {string} idDocumento ID operativo del documento.
 * @param {string} htmlTemplate Código HTML estructurado con la plantilla visual del documento.
 * @returns {string} URL de visualización del archivo generado.
 */
function DOC_GENERAR_PDF_OPERATIVO(idDocumento, htmlTemplate) {
  try {
    const blob = Utilities.newBlob(htmlTemplate, "text/html", "plantilla.html");
    const pdf = blob.getAs("application/pdf").setName(idDocumento + ".pdf");
    
    // Buscar o crear carpeta de destino
    let carpeta;
    const carpetas = DriveApp.getFoldersByName(DOC_CONFIG.CARPETA_PDFS);
    if (carpetas.hasNext()) {
      carpeta = carpetas.next();
    } else {
      carpeta = DriveApp.createFolder(DOC_CONFIG.CARPETA_PDFS);
    }

    const archivo = carpeta.createFile(pdf);
    return archivo.getUrl();
  } catch (err) {
    console.error("No se pudo compilar representación imprimible: " + err.toString());
    return "";
  }
}