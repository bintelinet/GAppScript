/**
 * Fundamentos 15 - Crea pdfs personalizados utilizando una plantilla html
 * https://www.youtube.com/watch?v=KAQBTWIDUKI&list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV&index=15
 * Al ejecutar la función crearPdfs revisar en Mi Unida la carpeta PdfsPersonalizados con los PDFs creados
 */


// VARIABLES GLOBALES
const AREA = 0;
const RESPONSABLE = 1;
const PUESTO = 2;
const ASUNTO = 3;

function crearPdfs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos');
  var data = sheet.getDataRange().getValues();
  // Borramos la primer fila de los encabezados
  data.shift()
  // Mostramos la información de la hoja "Datos" console.log(data);

  // Carpeta para almacenar los PDFs
  var folderName = 'PDFs personalizados';
  var folder = DriveApp.getFoldersByName(folderName);

  //Verificamos los folders encontrados
  if (folder.hasNext()) {
    folder = folder.next(); // Si encuentra un folder con el nombre buscado
  } else {
    // Si no se encuentra el folder buscado, lo creamos con el nombre indicado
    folder = DriveApp.createFolder(folderName);
  }

  data.forEach(row => {
    // No utilizamos el prefijo var porque la variable será global
    responsable = row[RESPONSABLE];
    puesto = row[PUESTO];
    area = row[AREA];
    asunto = row[ASUNTO];

    var html = HtmlService.createTemplateFromFile('Plantilla').evaluate();
    var pdf = html.getAs('application/pdf').setName('circular - ' + area + '.pdf');
    folder.createFile(pdf);
  });
}
