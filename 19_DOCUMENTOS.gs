// (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
/**************************************************************
* 19_DOCUMENTOS.gs (VERSIÓN 1.0 - V2 ERP - LIBRO 1)
* RESPONSABILIDAD:
* - Generar representaciones visuales de impresos (PDFs) de Facturas, Cotizaciones y Pedidos.
* - Guardar los archivos PDF generados de forma automatizada en Google Drive (DriveApp).
* - Indexar la URL_PDF pública en las cabeceras transaccionales de Sheets.
**************************************************************/

const DOC_CONFIG_CORE = {
  NOMBRE_CARPETA_DRIVE: "MEGUDAN_PDFS_AUDIT"
};

/**
 * Obtiene o crea la carpeta de almacenamiento de PDFs en Google Drive
 */
function DOC_OBTENER_CARPETA_DRIVE() {
  const carpetas = DriveApp.getFoldersByName(DOC_CONFIG_CORE.NOMBRE_CARPETA_DRIVE);
  if (carpetas.hasNext()) {
    return carpetas.next();
  }
  return DriveApp.createFolder(DOC_CONFIG_CORE.NOMBRE_CARPETA_DRIVE);
}

/**
 * RPC: Generar e Imprimir PDF oficial de Factura de Venta / Cotización
 */
function DOC_GENERAR_PDF_VENTA_WEB(idVenta, tokenSesion) {
  try {
    SEG_VERIFICAR_CONTEXTO_Y_ACCESO(tokenSesion, "VENTAS", "VER");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Obtener datos de la venta
    const resDetalle = VEN_OBTENER_DETALLE_VENTA_WEB(idVenta, tokenSesion);
    if (!resDetalle || !resDetalle.EXITO || !resDetalle.DATOS) {
      throw new Error("No se pudieron recuperar los datos de la venta " + idVenta);
    }
    
    const cab = resDetalle.DATOS.CABECERA || {};
    const dets = resDetalle.DATOS.DETALLES || [];
    
    // 2. Construir plantilla HTML para el PDF
    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; color: #0f172a; }
          .header-box { display: flex; justify-content: space-between; border-bottom: 2px solid #1d4ed8; padding-bottom: 15px; margin-bottom: 20px; }
          .title { font-size: 20px; font-weight: 800; color: #1d4ed8; }
          .meta { font-size: 12px; color: #475569; margin-top: 5px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th { background: #f1f5f9; padding: 10px; font-size: 11px; text-transform: uppercase; border-bottom: 2px solid #cbd5e1; }
          td { padding: 10px; font-size: 12px; border-bottom: 1px solid #e2e8f0; }
          .totals { margin-top: 20px; width: 300px; margin-left: auto; }
          .total-row { display: flex; justify-content: space-between; padding: 6px 0; font-size: 13px; }
          .total-grand { font-size: 16px; font-weight: 800; color: #1d4ed8; border-top: 2px solid #1d4ed8; padding-top: 8px; }
        </style>
      </head>
      <body>
        <div class="header-box">
          <div>
            <div class="title">MEGUDAN CONSTRUCCIONES SOSTENIBLES SAS</div>
            <div class="meta">NIT: 901.915.723-2 · Pitalito, Huila</div>
            <div class="meta">Correo: megudancs@gmail.com · Tel: 3118172601</div>
          </div>
          <div style="text-align: right;">
            <div style="font-size: 16px; font-weight: 800; color: #1e3a8a;">${cab.TIPO_DOCUMENTO || 'FACTURA DE VENTA'}</div>
            <div style="font-size: 14px; font-weight: 700;">N° ${cab.NUM_DOCUMENTO || idVenta}</div>
            <div class="meta">Fecha: ${cab.FECHA || 'N/A'}</div>
          </div>
        </div>
        
        <div style="background: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; font-size: 12px; margin-bottom: 20px;">
          <strong>CLIENTE / TERCERO:</strong> ${cab.RAZON_SOCIAL_CLIENTE || cab.ID_CLIENTE}<br>
          <strong>ID CLIENTE:</strong> ${cab.ID_CLIENTE} | <strong>FORMA DE PAGO:</strong> ${cab.FORMA_PAGO} (${cab.CONDICION_PAGO || 'Contado'})
        </div>
        
        <table>
          <thead>
            <tr>
              <th style="text-align: left;">Producto / Servicio</th>
              <th style="text-align: center;">Cantidad</th>
              <th style="text-align: right;">Precio Unit.</th>
              <th style="text-align: right;">IVA</th>
              <th style="text-align: right;">Total Línea</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    dets.forEach(d => {
      htmlContent += `
        <tr>
          <td><strong>${d.ID_PRODUCTO}</strong> - ${d.DESCRIPCION || ''}</td>
          <td style="text-align: center;">${d.CANTIDAD} ${d.ID_UNIDAD || 'UND'}</td>
          <td style="text-align: right;">$${Number(d.PRECIO_UNITARIO || 0).toLocaleString('es-CO')}</td>
          <td style="text-align: right;">$${Number(d.IVA || 0).toLocaleString('es-CO')}</td>
          <td style="text-align: right; font-weight: 700;">$${Number(d.TOTAL || 0).toLocaleString('es-CO')} COP</td>
        </tr>
      `;
    });
    
    htmlContent += `
          </tbody>
        </table>
        
        <div class="totals">
          <div class="total-row"><span>Subtotal:</span><span>$${Number(cab.SUBTOTAL || 0).toLocaleString('es-CO')} COP</span></div>
          <div class="total-row"><span>Descuento:</span><span>-$${Number(cab.DESCUENTO || 0).toLocaleString('es-CO')} COP</span></div>
          <div class="total-row"><span>I.V.A.:</span><span>$${Number(cab.IVA || 0).toLocaleString('es-CO')} COP</span></div>
          <div class="total-row total-grand"><span>TOTAL A PAGAR:</span><span>$${Number(cab.TOTAL || 0).toLocaleString('es-CO')} COP</span></div>
        </div>
        
        <div style="margin-top: 40px; font-size: 10px; color: #64748b; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 10px;">
          Documento generado por MEGUDAN ERP V2 · Representación Impresa Oficial
        </div>
      </body>
      </html>
    `;
    
    // 3. Crear Blob en PDF y guardar en Google Drive
    const blobHtml = HtmlService.createHtmlOutput(htmlContent).getBlob().getAs("application/pdf");
    blobHtml.setName("MEGUDAN_DOC_" + idVenta + "_" + String(new Date().getTime()).substring(5) + ".pdf");
    
    const carpetaDrive = DOC_OBTENER_CARPETA_DRIVE();
    const archivoPdf = carpetaDrive.createFile(blobHtml);
    archivoPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const urlPdf = archivoPdf.getUrl();
    
    return {
      EXITO: true,
      URL_PDF: urlPdf,
      ID_VENTA: idVenta,
      MENSAJE: "¡Documento PDF impreso y exportado con éxito a Google Drive!"
    };
  } catch (error) {
    return { EXITO: false, URL_PDF: "", MENSAJE: "Error al generar PDF: " + error.toString() };
  }
}

