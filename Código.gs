/**
 * Fundamentos GAS Script 38 - onChange events para datalist ( quicktip 02)
 * https://youtu.be/_UeGY2AAu8s?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV
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