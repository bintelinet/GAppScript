/**
 * Apps Script Fundamentos 14 - Creando una WebApp de 0 a 100
 * https://youtu.be/nke1zk82RfI?list=PLFVYPW43NcuzRignaoqLX1BBoNmN-cVQV
 */

// Global variables
const ss = SpreadsheetApp.openById('1jX7fiFJk4oGxzGloQa8-uMuSwJtUwIvNfYUUiZQV2tA')
const sheetRegistros = ss.getSheetByName('Registros');
const sheetBD = ss.getSheetByName('BD');

// Se ejecuta cuando se carga la página
function doGet() {
  // Obtener los datos de la hoja
  var data = sheetBD.getDataRange().getValues();
  // Borramos los encabezados del rango de datos
  data.shift(); 
  // Creamos un arreglo para guardar las opciones que se cargan en la página
  var talleres = [];
  // Iteramos cada una de las filas para llenar el arreglo
  data.forEach( row =>{
    // Si en la columna 2 que contiene los lugares disponibles es mayor a 0 hay lugares disponibles
    if( row[2] > 0){
      // Agregamos el taller usando el método push del dato que está en columna A Talleres
      talleres.push(row[0]);
    }
  });

  // LLamamos la plantilla Form
  var template = HtmlService.createTemplateFromFile('Form');
  // Creamos la propiedad talleres dentro del objeto template para llenar dinámicamente las opciones
  template.talleres = talleres;
  
  // Verificamos si hay algún error en el código de la plantilla
  var output = template.evaluate();
  // Devolvemos la plantilla
  return output;
}

/* https://developers.google.com/apps-script/add-ons/guides/css */

// Función para incluir archivos html en el Form
function include(fileName){
  // Obtenemos el contenido del archivo html
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}


// Se ejecuta si hay algún error del lado del servidor

// Se ejecuta del lado del servidor cuando se presionea el botón enviar
function addRecord(form){
  // Mostramos la propiedad talleres para verificar que recibimos el objeto adecuadamente
  console.log(form.talleres);

  // Validar el cupo al momento del envío, cabe la posibilidad que al momento de cargar la página todavía haya lugares, 
  // pero al momento del envío ya se hayan ocupado
  var cupoDisponible = verCupoTaller(form.talleres);
  
  // Agregar registro Fecha, email y taller elegido
  sheetRegistros.appendRow([new Date, Session.getActiveUser().getEmail(), form.talleres]);
  
  return "Registro recibido"; //Enviamos el mensaje del lado del cliente
}

function verCupoTaller(tallerElegido){
  // Almacenamos los valores en la hoja BD
  var talleres = sheetBD.getDataRange().getValues();
  // Quitamos el encabezado
  talleres.shift();
  // Iteramos a través de los talleres
  talleres.map( taller => { console.log(taller)
    if( taller[0] == tallerElegido){
      if( taller[2] > 0 ){
        // Existen lugares disponibles
        return true;
      }
      else{
        // Si no hay cupo arrojamos una excepción y ejecutará la función CallBackFunction(showError)
        throw "Este taller ya no tiene cupo, recarga la página"  // Se muestra del lado del cliente y ya no se mostrará nada de lado del servidos
      }
    }
  });
}


















