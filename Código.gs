/**
 * Fundamentos GAS 41 - validaciones del lado del servidor (💨 quicktip 05)
 * https://youtu.be/o7sCY4iNejI
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

function verificarEmailUsuario(userEmail = 'bintelinet@gmail.com'){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUsuarios = ss.getSheetByName('Users');
  const result = sheetUsuarios.createTextFinder(userEmail).findAll().map((range) => range.getA1Notation());
  console.log(result);

  // Verificamos si encontró la cadena en una celda
  if(result.length > 0){
    let message = `El usuario ${userEmail} ya está siendo utilizado. Por vavor elige otro correo.`;
    console.log(message);
  }else{
    return 'El registro se ha realizado con éxito';
  }

}