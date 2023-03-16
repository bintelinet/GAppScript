// ----------- VARIABLES GLOBALES ----------- //

// Url de las carpetas
const urlCrm = 'https://docs.google.com/spreadsheets/d/1A-TGb019R_DpojO6cwz9lZqTMDgIucB2csduD4P0gdM/edit#gid=0';
const urlFolderCtes = 'https://drive.google.com/drive/folders/1CTTf9wVvVjR016JCS6x4HDZP71PvxbRo';
const urlContenedorCFDIs = 'https://drive.google.com/drive/folders/1CJLlomZm6ul-2JEWzyiqkeD_pS1jDgBP';
const urlFolderAnexoRMF = 'https://drive.google.com/drive/folders/1R4v8QAMD58SrHhsGA8uKTYkQWYiIOwFP';

// Portafolio clientes  urlCrm.match(/[-\w]{25,}/)[0];
const idCrm = '1A-TGb019R_DpojO6cwz9lZqTMDgIucB2csduD4P0gdM';
//Repositorios de clientes: Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
const idFolderCtes = '1CTTf9wVvVjR016JCS6x4HDZP71PvxbRo';
//Mi unidad\Compuprintt\02 Administración\02 Contabilidad\Contenedor CFDIS
const idContenedorCfdis = '1CJLlomZm6ul-2JEWzyiqkeD_pS1jDgBP';
const idFolderEmitidas = '1FCGzs41RLegf-yaFA9ndRpeyrOcgS5e1';
const idFolderRecibidas = '1FAtwTDvCeF4we0ByHcywiyBTX3Dt5IUi';
// Carpeta de Anexo de la Resolución de la Miscelánea Fiscal
const idFolderAnexoRMF = '1R4v8QAMD58SrHhsGA8uKTYkQWYiIOwFP';

//Datos del negocio
const nameBusiness = '';
const ulrBusiness = 'https://bintelinet.com.mx/';
const urlImgBusiness = 'https://drive.google.com/uc?export=download&id=1Am9W5rs5KwWv2kpg2R_vrlXeGoI-PHO5';

//Datos del negocio
// const nameBusiness = 'Compurprintt';
// const ulrBusiness = 'https://compuprintt.mx/';
// const urlImgBusiness = 'https://drive.google.com/uc?export=download&id=1TkZWKxJcDAQ0fK-LDVNkxQhNo6OYzKox';

// ----------- FIN VARIABLES GLOBALES ----------- //


// ----------- DOGET DOPOST ----------- //

// Verificar del lado del contribuyente
function doGet() {
  const html = HtmlService.createTemplateFromFile('dataCFDIs');
  const output = html.evaluate();
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();

}

// ----------- FIN DOGET DOPOST ----------- //


// ----------- CREACIÓN DE MENUS Y SIDE BARS ----------- //
/**
 * Crea los menús en la hoja de cálculo
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet();

  //* Crea Menús
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("Bintelinet")
    .addSubMenu(ui.createMenu("Crear carpetas")
      .addItem("de contribuyente", 'uiCreateRepoContribuyente')
      .addItem("de facturas", 'uiCreateRepoFacturas'))

    .addSubMenu(ui.createMenu("Gestionar CFDIs")
      .addItem('Paso 1: Distribuir CFDIs', 'uiDistribuirCFDI')
      .addItem('Paso 2: Listar CFDIs', 'uiListCFDI')
      .addItem('Paso 3: Extraer datos', 'uiExtractDataCfdi'))

    .addSubMenu(ui.createMenu('Contribuyente')
      .addItem('Calcular Declaración', 'uiCalcularDeclaraSAT')
      .addItem('Presentar Declaración', 'uiPresentarDeclaraSAT'))

    .addToUi();
}


/**
 * Muestra el sidebar del Repositorio de Contribuyente
 */
function uiCreateRepoContribuyente() {

  var html = HtmlService.createTemplateFromFile('repoContribuyente')
    .evaluate()
    .setTitle("Creación de Carpetas del Contribuyente")

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar del Repositorio de Facturas
 */
function uiCreateRepoFacturas() {

  var html = HtmlService.createTemplateFromFile('repoFacturas')
    .evaluate()
    .setTitle("Creación de Carpetas de Facturas")

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar del Distribución de CFDIs
 */
function uiDistribuirCFDI() {
  // Interfaz UI

  var html = HtmlService.createTemplateFromFile('distribuirCFDI')
    .evaluate()
    .setTitle('Distribuición de CFDIs del Contribuyente');

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar del Listado de CFDIs
 */
function uiListCFDI() {
  // Interfaz UI

  var html = HtmlService.createTemplateFromFile('listarCFDIs')
    .evaluate()
    .setTitle('Listado de CFDIs del Contribuyente');

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar de la Extracción de Datos de los CFDIs
 */
function uiExtractDataCfdi() {

  var html = HtmlService.createTemplateFromFile('dataCFDIs')
    .evaluate()
    .setTitle('Extracción de Datos de los CFDIs');

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar para el cálculo de Declaraciones del Contribuyente
 */
function uiCalcularDeclaraSAT() {

  var html = HtmlService.createTemplateFromFile('calcDeclaraSAT')
    .evaluate()
    .setTitle('Cálculo de Declaraciones')

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

/**
 * Muestra el sidebar para el generar Informes de Declaraciones del Contribuyente
 */
function uiPresentarDeclaraSAT() {

  var html = HtmlService.createTemplateFromFile('presentaDeclaraSAT')
    .evaluate()
    .setTitle('Presentación de Declaraciones')

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

// ----------- FIN CREACIÓN DE MENUS Y SIDE BARS ----------- //



// ----------- FUNCIONES GENERALES ----------- //

/**
 * Muestra el aviso que no se ha seleccionado un elemento
 */
  function alerta(message, seg){
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Aviso', seg);
  }

/**
 * Activa la hoja de trabajo seleccionada de GSheet 
 */
function activeSheet(nameSheet = 'calculosCFDI') {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheet).activate();
}

/**
 * Obtiene la fecha actual
 */
function getToday() {
  var actual = new Date();
  var anio = actual.getFullYear();
  var mes = actual.getMonth() + 1;
  var dia = actual.getUTCDate();
  var hora = actual.getHours();
  var minutos = actual.getMinutes();
  var segundos = actual.getSeconds();


  today = {
    date: actual,
    year: anio,
    month: mes,
    day: dia,
    hour: hora,
    min: minutos,
    sec: segundos
  }
  console.log(today.year + "/" + today.month + "/" + today.day + "-" + today.hour + ":" + today.min + ":" + today.sec + "\n");
  return today;
}

/**
 * Obtiene los años y meses de cada RFC para el listado y/o extracción de datos de los CFDIs
 */
function getAniosMesesRepoRFC(rfc, cveRegimen, cvePeriodo) {
  // rfc = 'SAMR7510317E2', cveRegimen = '605', cvePeriodo = '06' 

  // c_Periodicidad	Descripción
  // 01	Diario
  // 02	Semanal
  // 03	Quincenal
  // 04	Mensual
  // 05	Bimestral
  // 06 Trimestral
  // 07 Cuatrimestral
  // 08 Anual

  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  if (rfc !== null && cvePeriodo !== null) {
    // Identifica la carpeta RFC del cliente
    var findFolder = getCarpeta(idFolderCtes, rfc);
    //Logger.log('Carpeta encontrada: ' + findFolder.name + ', ID: ' + findFolder.id)

    // Verifica si existe la carpeta RFC
    if (findFolder.name != null) {
      // Identifica la carpeta de Facturas del RFC encontrado
      var findFacturas = getCarpeta(findFolder.id, 'Facturas');
      var listaAnios = [];
      // Obten los las carpetas "años" de las facturas
      var aniosFacturacion = getCarpetas(findFacturas.id);
      aniosFacturacion.forEach(anio => {
        listaAnios.push(anio.name);
      });
      // Ordenamos los años de forma descendente
      listaAnios.sort((a, b) => b - a)
      //console.log(listaAnios);
    }

    // CFDI v4  https://docs.google.com/spreadsheets/d/1ECamUNkmD3B1c4kFbpBflOUEBg_j5U3mjLrwuyZ0W9o/edit#gid=1306767523
    // c_Periodicidad

    var listaMeses = [];


    switch (cvePeriodo) {

      case '04': // Periodicidad Mensual
        console.log('Periodicidad Mensual');
        var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
        ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre'], ['19-anual', 'Todos los Meses']];
        break;

      case '05': // Periodicidad Bimestral
        console.log('Periodicidad Bimestral');
        var listaMeses = [['13-ene-feb', 'Enero-Febrero'], ['14-mar-abr', 'Marzo-Abril'], ['15-may-jun', 'Mayo-Junio'], ['16-jul-ago', 'Julio-Agosto'],
        ['17-sep-oct', 'Septiembre-Octubre'], ['18-nov-dic', 'Noviembre-Diciembre'], ['19-anual', 'Todos los Meses']];
        break;

      case '06': // Periodicidad Trimestral
        var listaMeses = [['20-ene-feb-mar', 'Enero-Febrero-Marzo'], ['21-abr-may-jun', 'Abril-Mayo-Junio'], ['22-jul-ago-sep', 'Julio-Agosto-Septiembre'],
        ['23-oct-nov-dic', 'Octubre-Noviembre-Diciembre'], ['19-anual', 'Todos los Meses']];
        break;

      case '08': // Periodicidad Anual
        console.log('Periodicidad Anual');
        var listaMeses = [['19-anual', 'Todos los Meses']];
        break;

      default:
        console.log('Elija un periodo declaración correcta');
        var listaMeses = [];
    }
    console.log(listaMeses);

    var listaAniosMeses = {
      anios: listaAnios,
      meses: listaMeses
    }
    //console.log(listaAniosMeses);
    return listaAniosMeses;
  }
}

/**
 * Obtiene los años para cada RFC y los meses de declaración de acuerdo al regimen fiscal del contribuyente
 */
function getAniosMesesDeclaraSat(rfc, cveRegimen, cvePeriodo) {
  // rfc = 'SAMR7510317E2', cveRegimen = '605', cvePeriodo = '06' 

  // c_Periodicidad	Descripción
  // 01	Diario
  // 02	Semanal
  // 03	Quincenal
  // 04	Mensual
  // 05	Bimestral
  // 06 Trimestral
  // 07 Cuatrimestral
  // 08 Anual

  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  if (rfc !== null && cvePeriodo !== null) {
    // Identifica la carpeta RFC del cliente
    var findFolder = getCarpeta(idFolderCtes, rfc);
    //Logger.log('Carpeta encontrada: ' + findFolder.name + ', ID: ' + findFolder.id)

    // Verifica si existe la carpeta RFC
    if (findFolder.name != null) {
      // Identifica la carpeta de Facturas del RFC encontrado
      var findFacturas = getCarpeta(findFolder.id, 'Facturas');
      var listaAnios = [];
      // Obten los las carpetas "años" de las facturas
      var aniosFacturacion = getCarpetas(findFacturas.id);
      aniosFacturacion.forEach(anio => {
        listaAnios.push(anio.name);
      });
      // Ordenamos los años de forma descendente
      listaAnios.sort((a, b) => b - a)
      //console.log(listaAnios);
    }

    // CFDI v4  https://docs.google.com/spreadsheets/d/1ECamUNkmD3B1c4kFbpBflOUEBg_j5U3mjLrwuyZ0W9o/edit#gid=1306767523
    // c_Periodicidad

    var listaMeses = [];


    switch (cvePeriodo) {

      case '04': // Periodicidad Mensual
        console.log('Periodicidad Mensual');

        switch (cveRegimen) {
          // No incluye declaración anual
          case '626':	//Régimen Simplificado de Confianza
            var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
            ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];
            break;
          // Incluye declaración anual
          case '601':	// General de Ley Personas Morales
          case '606':	// Arrendamiento
          case '607':	// Régimen de Enajenación o Adquisición de Bienes
          case '610': // Residentes en el Extranjero sin Establecimiento Permanente en México
          case '612': // Personas Físicas con Actividades Empresariales y Profesionales
          case '622':	// Actividades Agrícolas, Ganaderas, Silvícolas y Pesqueras
          case '623':	// Opcional para Grupos de Sociedades
          case '624':	// Coordinados
          case '625':	// Régimen de las Actividades Empresariales con ingresos a través de Plataformas Tecnológicas
          case '626':	// Régimen Simplificado de Confianza
            var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
            ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre'], ['19-anual', 'Anual']];
            break;

          default:
            console.log('El regimen no forma parte de la periodicidad mensual')
            var listaMeses = [];
        }
        break;

      case '05': // Periodicidad Bimestral
        console.log('Periodicidad Bimestral');
        switch (cveRegimen) {
          case '621':	// Incorporación Fiscal
            var listaMeses = [['13-ene-feb', 'Enero-Febrero'], ['14-mar-abr', 'Marzo-Abril'], ['15-may-jun', 'Mayo-Junio'], ['16-jul-ago', 'Julio-Agosto'],
            ['17-sep-oct', 'Septiembre-Octubre'], ['18-nov-dic', 'Noviembre-Diciembre']];
            break;
          default:
            console.log('El regimen no forma parte de la periodicidad bimestral')
            var listaMeses = [];
        }
        break;

      case '06': // Periodicidad Trimestral
        console.log('Periodicidad Trimestral');
        switch (cveRegimen) {
          case '606':	// Arrendamiento
            var listaMeses = [['20-ene-feb-mar', 'Enero-Febrero-Marzo'], ['21-abr-may-jun', 'Abril-Mayo-Junio'], ['22-jul-ago-sep', 'Julio-Agosto-Septiembre'], ['23-oct-nov-dic', 'Octubre-Noviembre-Diciembre']];
            break;
          default:
            console.log('El regimen no forma parte de la periodicidad trimestral')
            var listaMeses = [];
        }
        break;

      case '08': // Periodicidad Anual
        console.log('Periodicidad Anual');
        switch (cveRegimen) {
          case '603':	// Personas Morales con Fines no Lucrativos
          case '605':	// Sueldos y Salarios e Ingresos Asimilados a Salarios
          case '608':	// Demás ingresos
          case '611':	// Ingresos por Dividendos (socios y accionistas)
          case '614':	// Ingresos por intereses
          case '615':	// Régimen de los ingresos por obtención de premios
          case '616':	// Sin obligaciones fiscales
          case '620':	// Sociedades Cooperativas de Producción que optan por diferir sus ingresos
            var listaMeses = [['19-anual', 'Anual']];
            break;

          default:
            console.log('El regimen no forma parte de la periodicidad anual')
            var listaMeses = [];
        }
        break;

      default:
        console.log('Elija un periodo declaración correcta');
        var listaMeses = [];
    }
    console.log(listaMeses);

    var listaAniosMeses = {
      anios: listaAnios,
      meses: listaMeses
    }
    //console.log(listaAniosMeses);
    return listaAniosMeses;
  }
}

/**
 * Modal para mostrar los mensajes de las funciones
 */
function showModal(title = 'Título', subtitle = 'Subtítulo', message = 'Contenido') {
  //const userInterface = HtmlService.createHtmlOutputFromFile('modal');
  /*   var printBar = '';
    if (progressBar == true) {
      printBar = '<hr>' +
        '<div class="progress">' +
        '<div class="progress-bar" role="progressbar" style="width: 25%;" aria-valuenow="25" aria-valuemin="0" aria-valuemax="100">25%' +
        '</div></div>'
    } */
  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<base target="_top">' +
    '<title>' + title + '</title>' +
    '</head>' +
    '<body>' +
    '<div class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">' +
    '<div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">' +
    '<div class="modal-content">' +
    '<div class="modal-header">' +
    '<h4 class="modal-title">' + subtitle + '</h4>' +
    '</div>' +
    '<div class="modal-body">' +
    message +
    '</div>' + //printBar +
    '</div>' +
    '</div>' +
    '</div>' +
    '</body>' +
    '</html>';

  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(userInterface, title);
}

// ----------- FIN FUNCIONES GENERALES ----------- //





