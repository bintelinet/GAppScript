/**
 * Fundamentos GAS 37 - Datalist para tus formularios web (💨 quicktip 01)
 * https://youtu.be/RDwCPP6yyg8?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV
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