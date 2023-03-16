// ----------- CREAR REPOSITORIOS ----------- //
/**
 * Crea el Repositorio de Contribuyente dado su RFC
 * Carpetas:
 *  - _RFC_
 *    - Auditum
 *    - Documentos
 *    - Facturas
 *    - FIEL
 * 
 *    - Documentos
 *      - Acuses
 *        - _AAAA_ // Año vigente
 *      - Declaraciones
 *        - _AAAA_ // Año vigente
 *      - Fiscales
 *      - Registros
 *        - _AAAA_ // Año vigente
 * 
 *    - Facturas
 *      - _AAAA_ // Año vigente
 *        - 01-ene
 *        - 02-feb
 *        - 03-mar
 *        - 04-abr
 *        - 05-may
 *        - 06-jun
 *        - 07-jul
 *        - 08-ago
 *        - 09-sep
 *        - 10-oct
 *        - 11-nov
 *        - 12-dic
 * 
 *        // Dentro de cada mes las siguientes carpetas
 *          - emitidas
 *          - recibidas
 */
function createRepoRFC(rfc = 'FIMC820127H28') {
  // rfc = 'SAMR7510317E2'
  // rfc = 'FIMC820127H28'
  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  var carpetasClientes = DriveApp.getFolderById(idFolderCtes);

  //Imprime la carpeta contenedora de los clientes: 
  //Logger.log(carpetasClientes.getName());

  //Obtenemos el año vigente
  var date = new Date();
  var year = date.getFullYear();
  //Logger.log(year);

  var findFolder = getCarpeta(idFolderCtes, rfc);
  Logger.log(rfc);

  // Verifica si ya existe la carpeta
  if (findFolder.name == null) {

    SpreadsheetApp.getActiveSpreadsheet().toast('Creando repositorio de contribuyente', 'Aviso', 5);

    var carpetaRFC = carpetasClientes.createFolder(rfc);
    // Carpetas de Nivel 1
    var auditum = carpetaRFC.createFolder('Auditum');
    var fiel = carpetaRFC.createFolder('FIEL');
    var documentos = carpetaRFC.createFolder('Documentos');
    var aniofacturacion = carpetaRFC.createFolder('Facturas');
    documentos.createFolder('Acuses');
    documentos.createFolder('Declaraciones');
    documentos.createFolder('Fiscales');
    documentos.createFolder('Registros');

    var rfcUrl = getCarpeta(idFolderCtes, rfc);
    var title = 'Carpetas del Contribuyente';
    var subtitle = 'RFC: ' + '<a href="' + rfcUrl.url + '" target="a_blank">' + rfc + '</a>';
    var message = '<p>Se crearon exitosamente las carpetas: </p>' +
      'Auditum<br/>' +
      'FIEL' + '<br/>' +
      'Documentos' + '<br/>' +
      'Facturas' + '<br/>' +
      'Acuses' + '<br/>' +
      'Declaraciones' + '<br/>' +
      'Fiscales' + '<br/>' +
      'Registros';
    //Logger.log(title + ' ' + subtitle);
    showModal(title, subtitle, message);

    // Carpetas de Nivel 2
    createRepoFacturas(rfc, year);StatusLimpiarDatos
  } else {
    // LLamando al modal
    var title = 'Carpetas del Contribuyente';
    var subtitle = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a>';
    var message = 'Ya existe el repositorio de contribuyente';
    showModal(title, subtitle, message);
    console.log(message);

  }
} // Fin crearRepoRfc

/**
 * Crea el Reposito de Facturas dado su RFC
 * Carpetas:
 *  - _RFC_
 *    - Documentos
 *      - Acuses
 *        - _AAAA_ // Año vigente
 *      - Declaraciones
 *        - _AAAA_ // Año vigente
 *      - Registros
 *        - _AAAA_ // Año vigente
 *    - Facturas
 *      - _AAAA_ // Año vigente
 *        - 01-ene
 *        - 02-feb
 *        - 03-mar
 *        - 04-abr
 *        - 05-may
 *        - 06-jun
 *        - 07-jul
 *        - 08-ago
 *        - 09-sep
 *        - 10-oct
 *        - 11-nov
 *        - 12-dic
 * 
 *        // Dentro de cada mes las siguientes carpetas
 *          - emitidas
 *          - recibidas
 */
function createRepoFacturas(rfc, year) {
  //rfc = 'SAMR7510317E2', year = '2021'

  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  SpreadsheetApp.getActiveSpreadsheet().toast('Preparando repositorio...', 'Aviso', 5);
  // Identifica la carpeta RFC del cliente
  var findFolder = getCarpeta(idFolderCtes, rfc);
  //Logger.log('Carpeta encontrada: ' + findFolder.name + ', ID: ' + findFolder.id)


  // Verifica si existe la carpeta RFC
  if (findFolder.name != null) {
    // Identifica la carpeta de Documentos del RFC encontrado
    var findDocumentos = getCarpeta(findFolder.id, 'Documentos');
    // Identifica la carpeta Acuses del RFC encontrado
    var findYearAcuses = getCarpeta(findDocumentos.id, 'Acuses');
    // Identifica la carpeta Declaraciones del RFC encontrado
    var findYearDeclaraciones = getCarpeta(findDocumentos.id, 'Declaraciones');
    // Identifica la carpeta Registros del RFC encontrado
    var findYearRegistros = getCarpeta(findDocumentos.id, 'Registros');

    // Identifica la carpeta de Facturas del RFC encontrado
    var findFacturas = getCarpeta(findFolder.id, 'Facturas');
    // Identifica la carpeta año 'AAAA' que se desea crear dentro de la carpeta Facturas
    var findYearFacturas = getCarpeta(findFacturas.id, year);

    // Crea la carpeta del año fiscal si no existe
    if (findYearFacturas.name == null) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Creando carpetas de facturas', 'Aviso', 5);
      var anioFacturacion = DriveApp.getFolderById(findFacturas.id).createFolder(year);

      var ene = '01-ene';
      var feb = '02-feb';
      var mar = '03-mar';
      var abr = '04-abr';
      var may = '05-may';
      var jun = '06-jun';
      var jul = '07-jul';
      var ago = '08-ago';
      var sep = '09-sep';
      var oct = '10-oct';
      var nov = '11-nov';
      var dic = '12-dic';

      // Crear las carpetas de los meses
      anioFacturacion.createFolder(ene);
      anioFacturacion.createFolder(feb);
      anioFacturacion.createFolder(mar);
      anioFacturacion.createFolder(abr);
      anioFacturacion.createFolder(may);
      anioFacturacion.createFolder(jun);
      anioFacturacion.createFolder(jul);
      anioFacturacion.createFolder(ago);
      anioFacturacion.createFolder(sep);
      anioFacturacion.createFolder(oct);
      anioFacturacion.createFolder(nov);
      anioFacturacion.createFolder(dic);

      // LLamando al modal
      var title = 'Carpetas de Facturas';
      var subtitle = 'Espere un momento por favor, no interrumpa el proceso...';
      var message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/>' +
        '<p>Se crearon exitosamente las carpetas: </p>' +
        '<a href="' + findYearAcuses.url + '" target="a_blank">Acuses</a>/' + year + '<br/>' +
        '<a href="' + findYearDeclaraciones.url + '" target="a_blank">Declaraciones</a>/' + year + '<br/>' +
        '<a href="' + findYearRegistros.url + '" target="a_blank">Registros</a>/' + year + '<br/>' +
        '<a href="' + findFacturas.url + '" target="a_blank">Facturas</a>/' + year + '<br/>' +
        ene + ', ' + feb + '<br/>' +
        mar + ', ' + abr + '<br/>' +
        may + ', ' + jun + '<br/>' +
        jul + ', ' + ago + '<br/>' +
        sep + ', ' + oct + '<br/>' +
        nov + ', ' + dic;
      Logger.log(message);
      showModal(title, subtitle, message);

      var carpetasMeses = getCarpetas(anioFacturacion.getId());

      carpetasMeses.forEach(meses => {
        var mes = DriveApp.getFolderById(meses.id);
        mes.createFolder('emitidas');
        mes.createFolder('recibidas');
      });

      DriveApp.getFolderById(findYearAcuses.id).createFolder(year);
      DriveApp.getFolderById(findYearDeclaraciones.id).createFolder(year);
      DriveApp.getFolderById(findYearRegistros.id).createFolder(year);

    } else {
      // LLamando al modal
      var title = 'Carpetas de Facturas';
      var subtitle = 'Cierre esta venta para continuar...';
      var message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br>'
        + 'Ya existe el repositorio de ' + '<a href="' + findFacturas.url + '" target="a_blank">Facturas</a> para el año fiscal: '
        + '<a href="' + findYearFacturas.url + '" target="a_blank">' + year + '</a>';
      showModal(title, subtitle, message);
      console.log(message);
    }
  }

}


/**
 * Redistribuye autómaticamente los CFDIS (xml y pdf) correspodiente para cada RFC, Año y Mes
 * Paso 1: Distribuir XMLs
 */
function distributeCfdis(destinoCfdis = 'emitidas') {
  console.log('Preparando archivos CFDIs...');

  var fecha = getToday();
  var anio = fecha.year;
  var mes = fecha.month;
  var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
  ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];
  //console.log(listaMeses[mes -1]);
  var nombreHoja = listaMeses[mes - 1][0] + '-' + destinoCfdis;

  // Sheet AAAA_Bitácora_CFDIs.gsheet
  var bitacora = getArchivo(idContenedorCfdis, anio + '-Bitácora-CFDIs');

  var selectBitacora = SpreadsheetApp.openById(bitacora.id);
  var ssBitacoraCfdis = selectBitacora.getSheetByName(nombreHoja);
  //ssBitacoraCfdis.getDataRange().clearContent();  

  // Identifica la carpeta emitidas o recibidas
  var tipoCfdis = '';
  if (destinoCfdis == 'emitidas') {
    tipoCfdis = idFolderEmitidas;
    var headerRfc = 'RFC Emisor';
  } else {
    tipoCfdis = idFolderRecibidas;
    var headerRfc = 'RFC Receptor';
  }

  var numCfdis = getNumCFDIs(tipoCfdis);
  var StatusDistributeCfdis = {
    title: 'Paso 1: Distribuir XMLs',
    data: [destinoCfdis, numCfdis],
    status: false   // Indica si ya concluyó la ejecución de la función actual
  };

  // Obteniendo datos para el modal
  var title = 'Distribuición de CFDIs del Contribuyente';


  if (numCfdis == 0) {
    var subtitle = 'Cierre esta venta para continuar...';
    var message = 'No se encontraron CFDIs en la carpeta: ' + destinoCfdis + '<br> Diríjase a la carpeta de Contenedor CFDIs';
  } else {
    var subtitle = 'Espere un momento por favor, no interrumpa el proceso...';
    var message = 'Distribuyendo ' + numCfdis + ' CFDIs ' + destinoCfdis;
  }
  //console.log(message);
  showModal(title, subtitle, message);

  var folder = DriveApp.getFolderById(tipoCfdis);
  var getXmlCfdiTxt = folder.searchFiles('mimeType = "text/xml" ');
  var getXmlCfdiApp = folder.searchFiles('mimeType = "application/xml" ');

  var headerCfdis = ['Time Stamp', 'Nombre CFDI', 'Tipo de Comprobante', 'ID Carpeta Origen', 'ID XML', 'ID PDF', 'Versión', headerRfc, 'Año Emisión', 'Mes Emisión',
    'ID Carpeta Destino', 'Reubicado', 'Error'];
  //ssBitacoraCfdis.appendRow(headerCfdis);

  var title = 'CFDIs ' + nombreHoja;
  //console.log(title);

  while (getXmlCfdiTxt.hasNext()) {

    var idXmlFile = getXmlCfdiTxt.next().getId();
    //console.log(idXmlFile)
    var xmlCfdi = getCfdi(idXmlFile);
    //console.log(xmlCfdi.xmlName)

    if (destinoCfdis == 'emitidas') {
      var valueRfc = xmlCfdi.xmlRfc;
    } else {
      var valueRfc = xmlCfdi.xmlRfcReceptor;
    }

    var fechaEmision = xmlCfdi.xmlDate;
    var anioEmision = fechaEmision.split('T')[0].split('-')[0];
    var numMes = fechaEmision.split('T')[0].split('-')[1];
    var mesEmision = listaMeses[numMes - 1][0];
    var carpetaDestino = getRepoCFDI(valueRfc, anioEmision, mesEmision, destinoCfdis);
    var idCarpetaDestino = carpetaDestino.id;
    var error = carpetaDestino.msg;
    var reubicado = false;

    if (idCarpetaDestino != null) {
      //console.log('CFDI: ' + xmlCfdi.xmlName)
      //console.log('Reubicando XML');
      moveFiles(xmlCfdi.xmlId, idCarpetaDestino);
      if (xmlCfdi.pdfId != null) {
        //console.log('Reubicando PDF');
        moveFiles(xmlCfdi.pdfId, idCarpetaDestino);
      }
      reubicado = true;
    }

    ssBitacoraCfdis.appendRow([new Date(), xmlCfdi.xmlName, xmlCfdi.xmlParentFolder, xmlCfdi.xmlIdParentFolder, xmlCfdi.xmlId, xmlCfdi.pdfId, xmlCfdi.xmlVersion, valueRfc, anioEmision, mesEmision, idCarpetaDestino, reubicado, error]);

    SpreadsheetApp.flush();//opcional
    //console.log(new Date() + ' | ' + xmlCfdi.xmlName + ' | ' + xmlCfdi.xmlParentFolder + ' | ' + xmlCfdi.xmlIdParentFolder + ' | ' + xmlCfdi.xmlId + ' | ' + xmlCfdi.pdfId + ' | ' + xmlCfdi.xmlVersion + ' | ' + valueRfc + ' | ' + anioEmision + ' | ' + mesEmision + ' | ' + idCarpetaDestino);

  }

  while (getXmlCfdiApp.hasNext()) {

    var idXmlFile = getXmlCfdiApp.next().getId();
    //console.log(idXmlFile)
    var xmlCfdi = getCfdi(idXmlFile);
    //console.log(xmlCfdi.xmlName)

    if (destinoCfdis == 'emitidas') {
      var valueRfc = xmlCfdi.xmlRfcEmisor;
    } else {
      var valueRfc = xmlCfdi.xmlRfcReceptor;
    }

    var fechaEmision = xmlCfdi.xmlDate;
    var anioEmision = fechaEmision.split('T')[0].split('-')[0];
    var numMes = fechaEmision.split('T')[0].split('-')[1];
    var mesEmision = listaMeses[numMes - 1][0];
    var carpetaDestino = getRepoCFDI(valueRfc, anioEmision, mesEmision, destinoCfdis);
    var idCarpetaDestino = carpetaDestino.id;
    var error = carpetaDestino.msg;
    var reubicado = false;

    if (idCarpetaDestino != null) {
      //console.log('CFDI: ' + xmlCfdi.xmlName)
      //console.log('Reubicando XML');
      moveFiles(xmlCfdi.xmlId, idCarpetaDestino);
      if (xmlCfdi.pdfId != null) {
        //console.log('Reubicando PDF');
        moveFiles(xmlCfdi.pdfId, idCarpetaDestino);
      }
      reubicado = true;
    }

    ssBitacoraCfdis.appendRow([new Date(), xmlCfdi.xmlName, xmlCfdi.xmlParentFolder, xmlCfdi.xmlIdParentFolder, xmlCfdi.xmlId, xmlCfdi.pdfId, xmlCfdi.xmlVersion, valueRfc, anioEmision, mesEmision, idCarpetaDestino, reubicado, error]);
    SpreadsheetApp.flush();//opcional
    //console.log(new Date() + ' | ' + xmlCfdi.xmlName + ' | ' + xmlCfdi.xmlParentFolder + ' | ' + xmlCfdi.xmlIdParentFolder + ' | ' + xmlCfdi.pdfId + ' | ' + xmlCfdi.xmlId + ' | ' + xmlCfdi.xmlVersion + ' | ' + valueRfc + ' | ' + anioEmision + ' | ' + mesEmision + ' | ' + idCarpetaDestino);
  }

  StatusDistributeCfdis.status = true;
  //console.log(StatusDistributeCfdis);
  return StatusDistributeCfdis;
}

// ----------- LISTAR CFDIs ----------- //
/**
 * Lista los CFDIs (xml y pdf) de cada RFC, año y mes para preparar su extracción de datos
 * Paso 2: Listar XMLs
 */
function listFilesCFDIs(rfc = 'GORD730303AN5', year = '2021', month = '01-ene') {
  //rfc, year, month
  //rfc = 'FIMC820127H28', year = '2018', month = '01-ene'
  //rfc = 'SAMR7510317E2', year = '2021', month = '10-oct'
  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  var title = 'Listar CFDIs del contribuyente';
  // Identifica la carpeta RFC del cliente
  var findFolder = getCarpeta(idFolderCtes, rfc);
  //Logger.log('Carpeta encontrada: ' + findFolder.name + ', ID: ' + findFolder.id)

  // Verifica si existe la carpeta RFC
  if (findFolder.name != null) {
    // Identifica la carpeta de Documentos del RFC encontrado
    var findDocumentos = getCarpeta(findFolder.id, 'Documentos');
    // Identifica la carpeta año 'AAAA' que se desea crear dentro de la carpeta Acuses
    //var findYearAcuses = getCarpeta(findDocumentos.id, 'Acuses');
    // Identifica la carpeta año 'AAAA' que se desea crear dentro de la carpeta Declaraciones
    //var findYearDeclaraciones = getCarpeta(findDocumentos.id, 'Declaraciones');

    // Identifica la carpeta de Facturas del RFC encontrado
    var findFacturas = getCarpeta(findFolder.id, 'Facturas');
    // Identifica la carpeta año 'AAAA' que se desea crear dentro de la carpeta Facturas
    var findYearFacturas = getCarpeta(findFacturas.id, year);

    // Verifica si exite la carpeta "AAAA" del año fiscal a importar 
    if (findYearFacturas.name != null) {
      var anioFacturacion = DriveApp.getFolderById(findFacturas.id);
      var findMonthFacturas = getCarpeta(findYearFacturas.id, month);
      //Logger.log(findMonthFacturas.id);

      // Verifica si existe la carpeta del mes a importar
      if (findMonthFacturas.name != null) {

        var tipoCfdiFolder = getCarpetas(findMonthFacturas.id);
        var ssRepoCfdiFiles = SpreadsheetApp.getActive().getSheetByName(month + '-reposCFDI');

        // Limpiando el contenido de la hoja
        ssRepoCfdiFiles.getDataRange().clearContent();
        var headerCfdi = ['Folio Fiscal', 'Nombre CFDI', 'Tipo CFDI', 'URL Carpeta', 'URL XML', 'URL PDF',
          'ID Carpeta', 'ID XML', 'ID PDF', 'Versión', 'Fecha Emisión', 'RFC Emisor', 'Razón Social Emisor', 'RFC Receptor', 'Razón Social Receptor', 'Retenciones', 'Traslados', 'Deducible'];
        ssRepoCfdiFiles.appendRow(headerCfdi);
        var idTipoCfdi = ''; var nameTipoCfdi = '';

        var meses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'], ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];

        if (month != '19-anual') {

          for (i = 0; i < meses.length; i++) {
            if (meses[i][0] == month) {
              var selMonth = meses[i][1];
              //console.log(month)
              break;
            }
          }
        }
        // Posicionarse al principio de la hoja
        ssRepoCfdiFiles.getRange("A2").activateAsCurrentCell();

        // -------  EMITIDAS -------//
        idTipoCfdi = tipoCfdiFolder[1].id;
        nameTipoCfdi = tipoCfdiFolder[1].name;

        // Obtiene el total de CFDIs dentro de la carpeta emitidas del mes correspondiente
        var totalCfdisEmitidos = getNumCFDIs(idTipoCfdi);

        // Obteniendo datos para el modal
        var title = 'Listado de CFDIs del Contribuyente';
        var subtitle = 'Espere un momento por favor, no interrumpa el proceso...';
        var message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/><br/>'
          + 'Listado de CFDIs del mes de ' + '<a href="' + findMonthFacturas.url + '" target="a_blank">' + selMonth + '</a>' + ' de '
          + '<a href="' + findYearFacturas.url + '" target="a_blank">' + year + '</a><br/>'
          + 'Listando ' + totalCfdisEmitidos + ' CFDIs ' + '<a href="' + findFolder.url + '" target="a_blank">emitidos</a><br/>';
        //console.log(message);
        showModal(title, subtitle, message);

        // Listando los CFDIs emitidos
        getCfdiFiles(idTipoCfdi, month);


        // -------  RECIBIDAS -------//
        idTipoCfdi = tipoCfdiFolder[0].id;
        nameTipoCfdi = tipoCfdiFolder[0].name;

        // Obtiene los archivos de la carpeta emitidas
        var totalCfdisRecibidos = getNumCFDIs(idTipoCfdi);

        var totalCfdis = totalCfdisEmitidos + totalCfdisRecibidos;

        message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/><br/>'
          + 'Listado de CFDIs del mes de ' + '<a href="' + findMonthFacturas.url + '" target="a_blank">' + selMonth + '</a>'
          + ' de ' + '<a href="' + findYearFacturas.url + '" target="a_blank">' + year + '</a><br/>'
          + 'Se listaron ' + totalCfdisEmitidos + ' CFDIs ' + '<a href="' + tipoCfdiFolder[1].url + '" target="a_blank">emitidos</a><br/>'
          + 'Listando ' + totalCfdisRecibidos + ' CFDIs ' + '<a href="' + tipoCfdiFolder[0].url + '" target="a_blank">recibidos</a><br/>'
          + 'Total de CFDIs encontrados ' + totalCfdis;
        showModal(title, subtitle, message);
        // Listando los CFDIs recibidos
        getCfdiFiles(idTipoCfdi, month);

        var StatusListCFDIs = {
          title: 'Paso 2: Listar CFDIs',
          data: [selMonth, totalCfdis, totalCfdisEmitidos, totalCfdisRecibidos],
          status: false
        };


        // Identificando la posición de las columnas        
        const tipoComprobante = 3;
        const fechaEmision = 11;
        const lastCol = 14;  // Col: Deducible
        var lastRow = ssRepoCfdiFiles.getLastRow();
        if (lastRow > 1) {
          // Reordenando los datos
          // https://www.youtube.com/watch?v=Wko1-ywLyVI
          SpreadsheetApp.getActiveSpreadsheet().toast('Reordenando datos del mes de ' + selMonth + ' del ' + year, 'Aviso');
          var dataSource = ssRepoCfdiFiles.getRange(2, 1, lastRow - 1, lastCol);
          //console.log(dataSource);
          // Sorts ascending by column fechaEmision, then ascending by column tipoComprobante
          dataSource.sort([{ column: fechaEmision, ascending: true }]);
          dataSource.sort([{ column: tipoComprobante, ascending: true }]);
        }
        ssRepoCfdiFiles.getRange(totalCfdis + 2, 1).activateAsCurrentCell();
        // Se concluyó el listado de CFDIs
      } else {
        // No se encontró la carpeta del mes seleccionado
        var title = 'Listado de CFDIs del Contribuyente';
        var subtitle = 'Verifique el Repositorio del Contribuyente...';
        var message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/><br/>'
          + '<a href="' + findYearFacturas.url + '" target="a_blank">' + 'Carpeta del Año Fiscal ' + year + '</a><br/>'
          + 'No existe la carpeta "' + month + '" para año fiscal ' + year;
        //console.log(message);
        showModal(title, subtitle, message);
      }
      // FIN if (findYearFacturas.name != null)
    } else {
      // No se encontró la carpeta del año seleccionado
      var title = 'Listado de CFDIs del Contribuyente';
      var subtitle = 'Verifique el Repositorio del Contribuyente...';
      var message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/><br/>'
        + 'No existe la carpeta del año fiscal "' + year + '"';
      //console.log(message);
      showModal(title, subtitle, message);
    }
  } else {
    // No se encontró la carpeta del RFC seleccionado  
    var title = 'Listado de CFDIs del Contribuyente';
    var subtitle = 'Verifique el Repositorio de Clientes...';
    var message = '<a href="' + urlCarpetaCtes + '" target="a_blank">' + 'Carpetal del RFC' + '</a><br/>'
      + 'No existe el RFC: ' + rfc + ' seleccionado, cree primero su Carpeta de Contribuyente';
    //console.log(message);
    showModal(title, subtitle, message);
  }
  subtitle = 'Cierre esta venta para continuar...';
  message = 'RFC: ' + '<a href="' + findFolder.url + '" target="a_blank">' + rfc + '</a><br/><br/>'
    + 'Listado de CFDIs del mes de ' + '<a href="' + findMonthFacturas.url + '" target="a_blank">' + selMonth + '</a>'
    + ' de ' + '<a href="' + findYearFacturas.url + '" target="a_blank">' + year + '</a><br/>'
    + 'Se listaron ' + totalCfdisEmitidos + ' CFDIs ' + '<a href="' + tipoCfdiFolder[1].url + '" target="a_blank">emitidos</a><br/>'
    + 'Se listaron ' + totalCfdisRecibidos + ' CFDIs ' + '<a href="' + tipoCfdiFolder[0].url + '" target="a_blank">recibidos</a><br/>'
    + 'Se listaron un total de ' + totalCfdis + ' CFDIs';
  //console.log(message);
  showModal(title, subtitle, message);
  StatusListCFDIs.status = true;
  Logger.log('Fin del listado de CFDIs')
  return StatusListCFDIs;
}

// ----------- FIN LISTAR CFDIs ----------- //

/**
 * Extrae los datos de los CFDIs según su RFC, año fiscal y mes
 * Paso 3: Extraer datos
 */
function getDataCFDI(rfc = 'GORD730303AN5', year = '2021', month = '11-nov') {

  // rfc, year, month
  // rfc = 'FIMC820127H28', year = '2018', month = '01-ene'
  // rfc = 'SAMR7510317E2', year = '2021', month = '10-oct'

  var ssRepoCfdiFiles = SpreadsheetApp.getActive().getSheetByName(month + '-reposCFDI');
  var ssDataCfdiFiles = SpreadsheetApp.getActive().getSheetByName(month + '-dataCFDI');
  var reposCfdi = ssRepoCfdiFiles.getDataRange().getValues();

  const RFC_EMISOR = reposCfdi[0].indexOf('RFC Emisor');
  const RFC_RECEPTOR = reposCfdi[0].indexOf('RFC Receptor');
  const TIPO_CFDI = reposCfdi[0].indexOf('Tipo CFDI');
  const ANIO = reposCfdi[0].indexOf('Fecha Emisión');

  const idXML = reposCfdi[0].indexOf('ID XML');
  const tipoCfdi = reposCfdi[0].indexOf('Tipo CFDI');

  var headerCfdi = ['Folio Fiscal', 'Serie', 'Folio', 'Año', 'Mes', 'Día', 'Hora', 'Tipo CFDI', 'Tipo de Comprobante',
    'RFC Emisor', 'Razón Social Emisor', 'Regimen Fiscal', 'RFC Receptor', 'Razón Social Receptor', 'Uso de CFDI', 'SubTotal', 'Descuento', 'Total',
    'ISR Retenido', 'IVA Retenido', 'IEPS Retenido', 'Total Imp. Retenidos',
    'Tipo Factor IVA', 'Tipo Factor IEPS', 'Tasa IVA Trasladado', 'Tasa IEPS Trasladado', 'IVA Trasladado', 'IEPS Trasladado', 'Total Imp. Trasladados',
    'Moneda', 'Tipo de Cambio', 'Forma de pago', 'Método de Pago', 'Condiciones de Pago',
    'Código Postal', 'No. Serie del CSD', 'No. de Certificado SAT', 'Sello CFDI', 'Sello SAT'];

  reposCfdi.shift();

  var meses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
  ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];

  if (month != '19-anual') {

    for (i = 0; i < meses.length; i++) {
      if (meses[i][0] == month) {
        var selMonth = meses[i][1];
        //console.log(month)
        break;
      }
    }
  }

  // Verifica si los datos de xx-mes-dataCFDI pertenecen al RFC seleccionado
  var datosValidos = true;
  var numTipoCfdis = [0, 0];
  var errorFila = 2;  //errorFila [Emisor,Receptor]
  var errorCol = RFC_EMISOR; // Error de la columna

  /*  Convertir un string a un objeto una fecha
  var fechaEmision = reposCfdi[fila][ANIO];
  var ms = Date.parse(fechaEmision);
  var anio = new Date(ms); */

  for (var fila in reposCfdi) {    
    // La columna fecha emisión debe tener el tipo de de dato: Texto sin formato
    var anio = reposCfdi[fila][ANIO].toLocaleString().split(' ')[0].split('-')[0];
    //console.log(anio);

    switch (reposCfdi[fila][TIPO_CFDI]) {
      case 'emitidas':
        if (reposCfdi[fila][RFC_EMISOR] == rfc && anio == year) {
          numTipoCfdis[0]++;
        } else {
          errorCol = RFC_EMISOR;
          datosValidos = false;
        }
        break;
      case 'recibidas':
        if (reposCfdi[fila][RFC_RECEPTOR] == rfc && anio == year) {
          numTipoCfdis[1]++;
        } else {
          errorCol = RFC_RECEPTOR;
          datosValidos = false;
        }
        break;
    }

  }

  var totalCfdis = numTipoCfdis[0] + numTipoCfdis[1];

  var StatusGetDataCFDI = {
    title: 'Paso 3: Extraer datos',
    data: [selMonth, totalCfdis, numTipoCfdis[0], numTipoCfdis[1]],
    status: false
  };

  var title = 'Extracción de Datos de los CFDIs'; var subtitle = ''; var message = '';
  if (datosValidos == true) {
    // Limpia la hoja por completo
    ssDataCfdiFiles.getDataRange().clearContent();
    // Escribe el encabezado
    ssDataCfdiFiles.appendRow(headerCfdi);
    //Posicionarse al princio de la hoja
    ssDataCfdiFiles.getRange("A2").activateAsCurrentCell();
    //SpreadsheetApp.getActiveSpreadsheet().toast('Importando datos de las facturas ', 'Aviso', 3);
    subtitle = 'Espere un momento por favor, no interrumpa el proceso...';
    message = 'RFC: ' + rfc + '<br/><br/>'
      + 'Extrayendo los datos de cada CFDI del mes de ' + selMonth + ' de ' + year + '<br/>'
      + 'Se encontraron ' + numTipoCfdis[0] + ' CDFIs emitidos <br/>'
      + 'Se encontraron ' + numTipoCfdis[1] + ' CDFIs recibidos <br/>'
      + 'Analizando ' + totalCfdis + ' CFDIs registrados en la hoja: ' + month + '-reposCFDI';
    //console.log(message);
    showModal(title, subtitle, message);

    // Extrae los datos de cada CFDI
    reposCfdi.forEach((fileXml, index) => {
      //Logger.log('(' + index + ')' + fileXml[idXML])
      // Obten el ID del XML
      var idXmlFile = fileXml[idXML];
      // Obten los datos de cada CFDI
      var dataCfdi = getDataCFDIfromGDrive(idXmlFile);

      var destinoCfdi = '';
      if (fileXml[tipoCfdi] == 'emitidas') {
        destinoCfdi = 'emitidas'
      } else {
        destinoCfdi = 'recibidas';
      }

      var uuid = dataCfdi.dataComplementoTimbre[1][2];
      var serieCSD = dataCfdi.dataComprobante[1][15];
      var fechaEmision = dataCfdi.dataComprobante[1][2];            //console.log(fechaEmision);
      var anioEmision = fechaEmision.split('T')[0].split('-')[0];   //console.log(anioEmision);
      var mesEmision = fechaEmision.split('T')[0].split('-')[1];    //console.log(mesEmision);
      var diaEmision = fechaEmision.split('T')[0].split('-')[2];    //console.log(diaEmision);
      var horaEmision = fechaEmision.split('T')[1].split('-')[0];   //console.log(horaEmision);
      var folio = dataCfdi.dataComprobante[1][14];
      var emisorRFC = dataCfdi.dataEmisor[1][0];
      var emisorRSocial = dataCfdi.dataEmisor[1][1];
      var emisorRegimen = dataCfdi.dataEmisor[1][2];
      var receptorRFC = dataCfdi.dataReceptor[1][0];
      var receptorRSocial = dataCfdi.dataReceptor[1][1];
      var receptorUsoCfdi = dataCfdi.dataReceptor[1][2];
      var subTotal = dataCfdi.dataComprobante[1][6];
      var descuento = dataCfdi.dataComprobante[1][5];
      var total = dataCfdi.dataComprobante[1][7];
      var totalImpuestos = dataCfdi.dataTotalImpuestos;

      var isrRetenido = null;
      var ivaRetenido = null;
      var iepsRetenido = null;
      var totalImpRetenidos = null;
      var ivaTipoFactor = null;
      var iepsTipoFactor = null;
      var tasaIvaTrasladado = null;
      var ivaTrasladado = null;
      var tasaIepsTrasladado = null;
      var iepsTrasladado = null;
      var totalImpTrasladados = null;

      if (totalImpuestos.length > 0) {
        // EL CFDIS incluye impuestos
        // https://docs.google.com/spreadsheets/d/1ECamUNkmD3B1c4kFbpBflOUEBg_j5U3mjLrwuyZ0W9o/edit#gid=1598084234
        // Catálogo de tasas o cuotas de impuestos: c_TasaOCuota
        // Catálogo de impuestos: c_Impuesto

        // Sección impuestos trasladados o retenidos
        var totalImpRetenidos = totalImpuestos[1][0];
        var totalImpTrasladados = totalImpuestos[1][1];

        if (totalImpuestos[1][0] != null) {
          // TotalImpuestosRetenidos
          //console.log(dataCfdi.dataImpRentenciones);          
          dataCfdi.dataImpRentenciones.forEach(retencion => {
            //Logger.log(retencion);
            if (retencion[0] == '001') {
              // ISR Retenido
              isrRetenido = retencion[1];
              //Logger.log('ISR Retenido: ' + isrRetenido);
            } else if (retencion[0] == '002') {
              // IVA Retenido
              ivaRetenido = retencion[1];
              //Logger.log('IVA Retenido: ' + ivaRetenido);
            } else if (retencion[0] == '003') {
              // IEPS Retenido
              iepsRetenido = retencion[1];
              //Logger.log('IEPS Retenido: ' + iepsTrasladado);
            }
          });
        }

        if (totalImpuestos[1][1] != null) {
          // TotalImpuestosTrasladados
          //console.log(dataCfdi.dataImpTraslados);
          dataCfdi.dataImpTraslados.forEach(traslado => {
            //Logger.log(traslado);
            if (traslado[0] == '002') {
              // IVA Trasladado
              ivaTipoFactor = traslado[1];
              tasaIvaTrasladado = traslado[2];
              ivaTrasladado = traslado[3];
              //Logger.log('IVA Trasladado: ' + ivaTrasladado);
            } else if (traslado[0] == '003') {
              // IEPS Trasladado
              iepsTipoFactor = traslado[1];
              tasaIepsTrasladado = traslado[2];
              iepsTrasladado = traslado[3];
              //Logger.log('IEPS trasladado: ' + iepsTrasladado);
            }
          });
        }

      }

      var moneda = dataCfdi.dataComprobante[1][3];
      var tipoCambio = dataCfdi.dataComprobante[1][4];
      var tipoComprobante = dataCfdi.dataComprobante[1][10];
      var formaPago = dataCfdi.dataComprobante[1][8];
      var metodoPago = dataCfdi.dataComprobante[1][11];
      var condicionesDePago = dataCfdi.dataComprobante[1][9];
      var codigoPostal = dataCfdi.dataComprobante[1][12];
      var serie = dataCfdi.dataComprobante[1][13];
      var noCertificadoSAT = dataCfdi.dataComplementoTimbre[1][6];
      var selloCFD = dataCfdi.dataComplementoTimbre[1][5];
      var selloSAT = dataCfdi.dataComplementoTimbre[1][7];

      if (receptorRFC == 'XAXX010101000') receptorRSocial = 'PÚBLICO GENERAL';
      var registro = [uuid, serie, folio, anioEmision, mesEmision, diaEmision, horaEmision, destinoCfdi, tipoComprobante, emisorRFC, emisorRSocial, emisorRegimen,
        receptorRFC, receptorRSocial, receptorUsoCfdi, subTotal, descuento, total, isrRetenido, ivaRetenido, iepsRetenido, totalImpRetenidos,
        ivaTipoFactor, iepsTipoFactor, tasaIvaTrasladado, tasaIepsTrasladado, ivaTrasladado, iepsTrasladado, totalImpTrasladados,
        moneda, tipoCambio, formaPago, metodoPago, condicionesDePago, codigoPostal, serieCSD, noCertificadoSAT, selloCFD, selloSAT];
      //console.log(registro);
      // Escribe los valores en la última fila de hoja   
      ssDataCfdiFiles.appendRow(registro);

    });
    //console.log('Total de CFDIs ' + totalCfdis)
    // Muestra el mensaje que ha finalizado el proceso
    subtitle = 'Cierre esta venta para continuar...';
    if (totalCfdis > 0) {
      message = 'RFC: ' + rfc + '<br/><br/>'
        + numTipoCfdis[0] + ' CDFIs emitidos <br/>'
        + numTipoCfdis[1] + ' CDFIs recibidos <br/>'
        + 'Se extrajeron con éxito los datos de ' + totalCfdis + ' CFDIs del mes ' + selMonth + ' de ' + year + '<br/>';
      ssDataCfdiFiles.getRange(totalCfdis + 2, 1).activateAsCurrentCell();
    }

  }
  else {
    subtitle = 'Cierre esta venta para continuar...';
    message = 'Los CFDIs listados en la hoja "' + month + '-reposCFDI' + '" no corresponden a los datos seleccionados:' + '<br/><br/>'
      + 'RFC ' + rfc + ' de ' + selMonth + ' de ' + year + '<br/><br/>'
      + 'Lista nuevamente los CFDIs adecuados del contribuyente seleccionado';

    console.log(numTipoCfdis[0] + ' | ' + numTipoCfdis[1])

    if (numTipoCfdis[0] == 0 && numTipoCfdis[1] > 0) {
      console.log('Contiene solo CFDIs recibidas');
      errorFila += numTipoCfdis[1];
      console.log('Error en fila: ' + errorFila);
      ssRepoCfdiFiles.getRange(errorFila, RFC_RECEPTOR + 1).activateAsCurrentCell();
    }
    if (numTipoCfdis[1] == 0 && numTipoCfdis[0] > 0) {
      console.log('Contiene solo CFDIs emitidas');
      errorFila += numTipoCfdis[0];
      console.log('Error en fila: ' + errorFila);
      ssRepoCfdiFiles.getRange(errorFila, RFC_EMISOR + 1).activateAsCurrentCell();
    }
    if (numTipoCfdis[0] > 0 && numTipoCfdis[1] > 0) {
      errorFila += numTipoCfdis[0] + numTipoCfdis[1];
      console.log('Error en fila: ' + errorFila);
      ssRepoCfdiFiles.getRange(errorFila, errorCol + 1).activateAsCurrentCell();
    }

  }

  //console.log(message);
  showModal(title, subtitle, message);

  StatusGetDataCFDI.status = true;
  //console.log(StatusGetDataCFDI);
  return StatusGetDataCFDI;
}
