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

function DOC_GENERAR_PDF_OPERATIVO(idDocumento, htmlTemplate) {
  try {
    const blob = Utilities.newBlob(htmlTemplate, "text/html", idDocumento + ".html");
    let carpeta;
    const carpetas = DriveApp.getFoldersByName(DOC_CONFIG.CARPETA_PDFS);
    if (carpetas.hasNext()) {
      carpeta = carpetas.next();
    } else {
      carpeta = DriveApp.createFolder(DOC_CONFIG.CARPETA_PDFS);
    }
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (err) {
    console.error("No se pudo compilar representación imprimible: " + err.toString());
    return "";
  }
}
