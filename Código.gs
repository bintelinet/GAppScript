/**
 * Fundamentos GAS Script 39 - onChange events para checkboxes & selects (quicktip 03)
 * https://youtu.be/EllKT72CBls?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV
 */

function doGet() {
  const html = HtmlService.createTemplateFromFile('FormFacturas');
  const output = html.evaluate();
  return output;
}

function getDataMedicamentos(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMedicamentos = ss.getSheetByName('BD_Medicamentos');
  const dataMedicamentos = sheetMedicamentos.getDataRange().getDisplayValues();
  dataMedicamentos.shift();
  //console.log(dataMedicamentos);
  return dataMedicamentos;
}