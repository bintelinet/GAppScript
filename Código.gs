/**
 * Fundamentos GAS 40 - validaciones del lado del cliente (quicktip 04)
 * https://youtu.be/iZnIJJEs3xs?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV
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