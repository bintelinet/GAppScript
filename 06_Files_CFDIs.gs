// -----------OBTENCIÓN DE DATOS DE LOS CFDIs ----------- //

/**
 * Obtiene todos los datos contenidos dentro de un XML dado su id
 */
function getDataCFDIfromGDrive(idXmlFile = '1F-vQcXKwgOQauXKmvP7_LTuJhFe0rsDf') {
  // Obtenemos el ID del documento   
  // 1F-vQcXKwgOQauXKmvP7_LTuJhFe0rsDf    Sin conceptos
  // excentos 1bRSK1KhkreBEjRb2zIoKTdHdvylq-FFW
  // barcel 1mO9-kViLFJu7K0UILygnSX1P5j8oRWbv
  // 15CM5DZMXDeDFWYxYRBjHkwiTi9WQipED
  var nameCFDI = DriveApp.getFileById(idXmlFile);
  // Muestra el nombre del archivo  :   
  Logger.log('Extrayendo datos del CFDI: ' + nameCFDI);
  // Muestra la URL:    
  Logger.log('ID: ' + idXmlFile + ' | URL: ' + nameCFDI.getUrl());

  // https://stackoverflow.com/questions/62915684/parse-utf-8-bom-xml-files-using-xmlservice-google-app-script
  var bytes = DriveApp.getFileById(idXmlFile).getBlob().getBytes();
  //Logger.log(bytes);
  // Remueve los bits del Boom
  var boom = bytes.splice(0, 3); 
  //console.log(boom);
  //Guarda el archivo XML UTF8 en content
  if (boom[0] < 0) {
    //Logger.log('Con Boom');
    var xmlFile = bytes;
    //Logger.log(xmlFile);
  } else {
    var xmlFile = DriveApp.getFileById(idXmlFile).getBlob().getBytes();
  }
  var content = Utilities.newBlob(xmlFile).getDataAsString();
  // Muestra el contenido del XML: 
  //Logger.log(content);
  // Analiza el la gramática del XML
  var document = XmlService.parse(content);
  //Logger.log(document);
  var root = document.getRootElement();
  //Logger.log('Root element: '+root);
  var uriCFDI = root.getNamespace().getURI();
  //Logger.log('URI: ' + uriCFDI);
  var namePrefix = root.getNamespace().getPrefix();
  //Logger.log(namePrefix + ':' + uriCFDI)

  var itemCfdi = root.getChildren();
  var comprobante = [];
  var emisor = [];
  var receptor = [];
  var conceptos = [];
  var impuestoConceptos = [];
  var totalImpuestos = [];
  var retenciones = [];
  var traslados = [];
  var complementoPago = [];
  var complementoTimbre = [];
  var temp1 = [];
  var temp2 = [];

  // --- Sección: Comprobante --- //
  //Logger.log('*-- ' + root.getName() + ' --*');
  var headerComprobante = ['schemaLocation', 'Version', 'Fecha', 'Moneda', 'TipoCambio', 'Descuento', 'SubTotal', 'Total', 'FormaPago',
    'CondicionesDePago', 'TipoDeComprobante', 'MetodoPago', 'LugarExpedicion', 'Serie', 'Folio', 'NoCertificado', 'Certificado', 'Sello'];
  comprobante.push(headerComprobante);
  //var atribRoot = root.getAttributes();

  headerComprobante.forEach((elem, index) => {
    // console.log('Atributo: ' + elem);
    var attribComprobante = root.getAttribute(elem);
    if (attribComprobante != null) {
      //Logger.log(index + ') ' + attribComprobante.getName() + ': ' + attribComprobante.getValue());
      temp1[index] = attribComprobante.getValue();
    } else {
      temp1[index] = null;
    }
  });
  comprobante.push(temp1);
  temp1 = [];
  //Logger.log(comprobante);

  //Elementos del CFDI:  ['Comprobante', 'Emisor', 'Receptor', 'Conceptos', 'Impuestos', 'Complemento']
  var indexItem = [];
  itemCfdi.forEach(item => {
    indexItem.push(item.getName());
  })
  //Logger.log(indexItem);

  // --- Sección: Emisor --- //
  var idxEmisor = indexItem.indexOf('Emisor');
  //Logger.log('*-- ' + itemCfdi[idxEmisor].getName() + ' --*');

  var headerEmisor = ['Rfc', 'Nombre', 'RegimenFiscal'];
  emisor.push(headerEmisor);
  var attribEmisor = itemCfdi[idxEmisor].getAttributes();
  attribEmisor.map(elem => {
    var index = headerEmisor.indexOf(elem.getName());
    //Logger.log(elem.getName() + ': ' + elem.getValue())
    temp1[index] = elem.getValue();
    //Logger.log(temp1);
  });
  emisor.push(temp1);
  temp1 = [];
  //Logger.log(emisor);


  // --- Sección: Receptor --- //
  var idxReceptor = indexItem.indexOf('Receptor');
  //Logger.log('*-- ' + itemCfdi[idxReceptor].getName() + ' --*');

  var headerReceptor = ['Rfc', 'Nombre', 'UsoCFDI'];
  receptor.push(headerReceptor);
  var attribReceptor = itemCfdi[idxReceptor].getAttributes();
  attribReceptor.map(elem => {
    var index = headerReceptor.indexOf(elem.getName());
    //Logger.log(elem.getName() + ': ' + elem.getValue())
    temp1[index] = elem.getValue();
    //Logger.log(temp1);
  });
  receptor.push(temp1);
  temp1 = [];
  //Logger.log(receptor);



  // --- Sección: Conceptos --- //
  var idxConceptos = indexItem.indexOf('Conceptos')

  if (idxConceptos >= 0) {
    //Logger.log('*-- ' + itemCfdi[idxConcept].getName() + ' --*');
    var headerConceptos = ['ClaveProdServ', 'NoIdentificacion', 'Cantidad', 'ClaveUnidad', 'Unidad', 'Descripcion', 'ValorUnitario', 'Importe', 'Descuento', 'NoPedimento', 'NoCuentaPredial'];
    conceptos.push(headerConceptos);
    // Impuestos trasladados 
    var headerImpuestoConceptos = ['Base', 'Impuesto', 'TipoFactor', 'TasaOCuota', 'Importe'];
    impuestoConceptos.push(headerImpuestoConceptos);

    var itemConceptos = itemCfdi[idxConceptos].getChildren();
    itemConceptos.map((mapConceptos, idxItemConceptos) => {

      // --- Descripción Conceptos --- //
      // Logger.log(' -> ' + mapConceptos.getName() + ' ' + numConcepto + ' <-');
      headerConceptos.forEach((elem, index) => {
        //console.log('Atributo: ' + elem);
        var attribConcepto = mapConceptos.getAttribute(elem);
        if (attribConcepto != null) {
          //Logger.log('(' + idxItemConceptos + ',' + index + ') Desc Concepto ' + attribConcepto.getName() + ': ' + attribConcepto.getValue());
          temp1[index] = attribConcepto.getValue();
        } else {
          temp1[index] = null;
        }
      });
      conceptos.push(temp1);
      temp1 = [];
      //Logger.log(conceptos);

      // --- Impuestos de Conceptos --- //        
      if (mapConceptos.getChildren().length == 0) {
        //Logger.log('No contiene impuestos');
        //Logger.log('(' + idxItemConceptos + ') Sin Imp Concepto ');
        impuestoConceptos.push([]);
        //Logger.log(impuestoConceptos);
      } else if (mapConceptos.getChildren().length > 0) {
        var itemImpConcepto = mapConceptos.getChildren()[0].getChildren();
        //console.log(itemImpConcepto);

        itemImpConcepto.map((impuesto, idxItemImpuestos) => {
          var attribImpuesto = impuesto.getChildren()[0].getAttributes();
          //Logger.log('(' + idxItemConceptos + ',' + idxItemImpuestos + ') Imp Concepto ' + attribImpuesto);
          attribImpuesto.map(elem => {
            var index = headerImpuestoConceptos.indexOf(elem.getName());
            temp2[index] = elem.getValue();
          });
          impuestoConceptos.push(temp2);

        });
      }
    }); // Fin itemConceptos
  }

  // --- Fin Sección: Conceptos --- //


  // --- Sección: Impuestos --- //

  var idxImp = indexItem.indexOf('Impuestos')

  if (idxImp >= 0) {

    //Logger.log('*-- ' + itemCfdi[idxImp].getName() + ' --*');
    var headerTotalImpuestos = ['TotalImpuestosRetenidos', 'TotalImpuestosTrasladados'];
    totalImpuestos.push(headerTotalImpuestos);
    var atribTotalImpuestos = itemCfdi[idxImp].getAttributes();
    atribTotalImpuestos.map(elem => {
      index = headerTotalImpuestos.indexOf(elem.getName());
      //Logger.log(elem.getName() + ': ' + elem.getValue());
      temp1[index] = elem.getValue();
    });
    totalImpuestos.push(temp1);
    temp1 = [];
    //Logger.log(totalImpuestos);

    var itemImpuestos = itemCfdi[idxImp].getChildren();
    itemImpuestos.map(impuesto => {

      // --- Descripción Impuestos --- //

      numImpuestos = impuesto.getChildren().length;

      if (impuesto.getName() == 'Traslados') {
        //Logger.log(numImpuestos + '. ' + impuesto.getName());
        var headerTraslados = ['Impuesto', 'TipoFactor', 'TasaOCuota', 'Importe'];
        traslados.push(headerTraslados);
        temp2 = [];
        for (i = 0; i < impuesto.getChildren().length; i++) {
          var nombImpuesto = impuesto.getChildren()[i].getAttributes();
          nombImpuesto.map(elem => {
            var index = headerTraslados.indexOf(elem.getName());
            temp2[index] = elem.getValue();
          });
          traslados.push(temp2);
        }
        //Logger.log(traslados);
      } else {
        //Logger.log(numImpuestos + '. ' + impuesto.getName());
        var headerRetenciones = ['Impuesto', 'Importe'];
        retenciones.push(headerRetenciones);
        temp2 = [];
        for (i = 0; i < impuesto.getChildren().length; i++) {
          var nombImpuesto = impuesto.getChildren()[i].getAttributes();
          nombImpuesto.map(elem => {
            var index = headerRetenciones.indexOf(elem.getName());
            temp2[index] = elem.getValue()
          });
          retenciones.push(temp2);
        }
        //Logger.log(retenciones);
      }
    });
  }

  // --- Fin Sección: Impuestos --- //


  // --- Sección: Complemento --- //

  var idxComp = indexItem.indexOf('Complemento')

  if (idxComp >= 0) {
    //Logger.log('*-- ' + itemCfdi[idxComp].getName() + ' --*');
    var complementos = itemCfdi[idxComp].getChildren();

    var tipoComplementos = [];
    var tipoPagos = [];
    complementos.forEach(complemento => {
      tipoComplementos.push(complemento.getName());
      temp1.push(complemento.getName());

      switch (complemento.getName()) {
        case 'TimbreFiscalDigital':
          //Logger.log(complemento.getName());
          var headerTimbreFiscal = ['schemaLocation', 'Version', 'UUID', 'FechaTimbrado', 'RfcProvCertif', 'SelloCFD', 'NoCertificadoSAT', 'SelloSAT'];
          complementoTimbre.push(headerTimbreFiscal);
          var trimbreFiscal = complemento.getAttributes();
          trimbreFiscal.map(complemento => {
            index = headerTimbreFiscal.indexOf(complemento.getName());
            //Logger.log(index +') '+complemento.getName());
            //Logger.log(complemento.getName() + ': ' +complemento.getValue() + ' | Col: ' + index);
            temp1[index] = complemento.getValue();
          });
          complementoTimbre.push(temp1);
          temp1 = [];
          break;
        case 'Pagos':
          //Logger.log(complemento.getName());
          var headerPago = ['FechaPago', 'FormaDePagoP', 'MonedaP', 'Monto', 'NumOperacion', 'RfcEmisorCtaOrd', 'NomBancoOrdExt', 'CtaOrdenante',
            'RfcEmisorCtaBen', 'CtaBeneficiario'];
          var headerDocRelacionado = ['IdDocumento', 'MonedaDR', 'Folio', 'MetodoDePagoDR', 'NumParcialidad', 'ImpSaldoAnt', 'ImpSaldoInsoluto', 'ImpPagado'];

          var pagos = complemento.getDescendants();
          //console.log(pagos.map(elem => elem.getValue()))
          pagos.forEach((pago, index) => {
            if (pago.asElement() !== null) {
              //console.log(pago.getValue())
              //Logger.log(pagos[index].getAttributes());
              temp1[index] = pagos[index].getAttributes();
            }
          });
          //complementoPago.push(temp1);
          temp1 = [];
          //Logger.log(pagos[3].getAttributes());
          break;
      }

    });
    //console.log(tipoComplementos);
  }

  // --- Fin Sección: Complemento --- //

  //Logger.log(comprobante);
  //Logger.log(emisor);
  //Logger.log(receptor);
  //Logger.log(conceptos);
  //Logger.log(impuestoConceptos);
  //Logger.log(totalImpuestos);
  //Logger.log(retenciones);
  //Logger.log(traslados);
  //Logger.log(tipoComplementos)
  //Logger.log(complementoTimbre);

  // Contruyendo el objeto con los datos del CFDI
  var dataCFDI = {
    title: 'Esquema CFDI',
    dataComprobante: comprobante,
    dataEmisor: emisor,
    dataReceptor: receptor,
    dataConceptos: conceptos,
    dataImpuestoConceptos: impuestoConceptos,
    dataTotalImpuestos: totalImpuestos,
    dataImpRentenciones: retenciones,
    dataImpTraslados: traslados,
    headerTipoComplementos: tipoComplementos,
    headerTipoPagos: tipoPagos,
    dataComplementoPago: complementoPago,
    dataComplementoTimbre: complementoTimbre
  }
  //Logger.log(dataCFDI.dataImpTraslados)
  return dataCFDI;
}


// -----------FIN OBTENCIÓN DE DATOS DE LOS CFDIs ----------- //


// ----------- OBTENCIÓN DE ARCHIVOS XML Y PDF ----------- //

/**
 * Escribe en la hoja del mes seleccionado los datos de los CFDIs obtenidos
 * de la carpeta emitidas y recibidas para el mes seleccionado
 */
function getCfdiFiles(idSearchFolder, month) {
  // idSearchFolder = '1NjZVLQZ-s25Qqu9YNGqVhhJ9jiZBbxjZ', month = '03-mar'
  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_\Facturas\##-mes-reposCFDI
  var ssRepoCfdiFiles = SpreadsheetApp.getActive().getSheetByName(month + '-reposCFDI');

  // Identifica la carpeta RFC del cliente
  var folder = DriveApp.getFolderById(idSearchFolder);
  var getXmlCfdiTxt = folder.searchFiles('mimeType = "text/xml" ');
  var getXmlCfdiApp = folder.searchFiles('mimeType = "application/xml" ');

  while (getXmlCfdiTxt.hasNext()) {
    // headear: ['Folio Fiscal', 'Nombre CFDI', 'Tipo CFDI', 'URL Carpeta', 'URL XML', 'URL PDF', 'ID Carpeta', 'ID XML', 'ID PDF', 'Versión', 'Fecha Emisión', 'RFC Emisor', 
    var idXmlFile = getXmlCfdiTxt.next().getId();
    var xmlCfdi = getCfdi(idXmlFile);
    console.log(idSearchFolder);
    //Logger.log(idXmlFile + ' | '  + xmlCfdi.xmlName);
    ssRepoCfdiFiles.appendRow([xmlCfdi.xmlUuid, xmlCfdi.xmlName, xmlCfdi.xmlParentFolder, xmlCfdi.xmlUrlParentFolder, xmlCfdi.xmlUrl, xmlCfdi.pdfUrl,
    xmlCfdi.xmlIdParentFolder, xmlCfdi.xmlId, xmlCfdi.pdfId, xmlCfdi.xmlVersion, xmlCfdi.xmlDate, xmlCfdi.xmlRfcEmisor, xmlCfdi.xmlRfcNameEmisor, xmlCfdi.xmlRfcReceptor, xmlCfdi.xmlRfcNameReceptor, null, null, null]);

    SpreadsheetApp.flush();//opcional
    //Logger.log(xmlCfdi.xmlUuid + ' | ' + xmlCfdi.xmlName + ' | ' + xmlCfdi.xmlVersion + ' | ' + xmlCfdi.xmlSerie + ' | ' + xmlCfdi.xmlFolio + ' | ' + xmlCfdi.xmlDate);
  }



  while (getXmlCfdiApp.hasNext()) {
    // headear: ['Folio Fiscal', 'Nombre CFDI', 'Tipo CFDI', 'URL Carpeta', 'URL XML', 'URL PDF', 'ID Carpeta', 'ID XML', 'ID PDF', 'Versión', 'Fecha Emisión''RFC', 'Razón Social', 'Retenciones', 'Traslados', 'Deducible'];

    var idXmlFile = getXmlCfdiApp.next().getId();
    var xmlCfdi = getCfdi(idXmlFile);
    //Logger.log(idXmlFile + ' | '  + xmlCfdi.xmlName);
    ssRepoCfdiFiles.appendRow([xmlCfdi.xmlUuid, xmlCfdi.xmlName, xmlCfdi.xmlParentFolder, xmlCfdi.xmlUrlParentFolder, xmlCfdi.xmlUrl, xmlCfdi.pdfUrl,
    xmlCfdi.xmlIdParentFolder, xmlCfdi.xmlId, xmlCfdi.pdfId, xmlCfdi.xmlVersion, xmlCfdi.xmlDate, xmlCfdi.xmlRfcEmisor, xmlCfdi.xmlRfcNameEmisor, xmlCfdi.xmlRfcReceptor, xmlCfdi.xmlRfcNameReceptor, null, null, null]);

    SpreadsheetApp.flush();//opcional
    //Logger.log(xmlCfdi.xmlUuid + ' | ' + xmlCfdi.xmlName + ' | ' + xmlCfdi.xmlVersion + ' | ' + xmlCfdi.xmlSerie + ' | ' + xmlCfdi.xmlFolio + ' | ' + xmlCfdi.xmlDate);
  }

}


function moveFiles(sourceFileId, targetFolderId) {
  var file = DriveApp.getFileById(sourceFileId);
  var folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}

// ----------- FIN OBTENCIÓN DE ARCHIVOS XML Y PDF ----------- //


// ----------- OBTENCIÓN DE CARPETAS Y ARCHIVOS ----------- //

/**
 * Busca la carpeta padre del contribuyente y su carpeta de facturación
 */
function getRepoCFDI(rfc, year, month, cfdis) {
  //rfc = 'SAMR7510317E2', year = '2021'

  //ID carpeta "04 Clientes"
  //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes
  //SpreadsheetApp.getActiveSpreadsheet().toast('Encontrando repositorio del RFC: ' + rfc + '...', 'Aviso', 5);

  // Identifica la carpeta RFC del cliente
  var findRFC = getCarpeta(idFolderCtes, rfc);
  var idCarpetaDestino = {};
  var message = null;

  // Verifica si existe la carpeta RFC
  if (findRFC.name != null) {
    //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_
    //Logger.log('Carpeta encontrada: ' + findRFC.name + ', ID: ' + findRFC.id);
    // Identifica la carpeta de Facturas del RFC encontrado
    var findFacturas = getCarpeta(findRFC.id, 'Facturas');
    //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_\Facturas
    //Logger.log('Carpeta encontrada: ' + findFacturas.name + ', ID: ' + findFacturas.id);
    // Identifica la carpeta año 'AAAA' que se desea encuentra dentro de la carpeta Facturas
    var findYearFacturas = getCarpeta(findFacturas.id, year);
    //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_\Facturas\_AAAA_

    // Encuentra el años fiscal de la factura
    if (findYearFacturas.name != null) {
      //Logger.log('Carpeta encontrada: ' + findYearFacturas.name + ', ID: ' + findYearFacturas.id);
      var mesFacturacion = getCarpeta(findYearFacturas.id, month);
      //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_\Facturas\_MMMM_
      //Logger.log('Carpeta encontrada: ' + mesFacturacion.name + ', ID: ' + mesFacturacion.id);
      idCarpetaDestino = getCarpeta(mesFacturacion.id, cfdis);
      //Mi unidad\Compuprintt\01 Dirección\01 Empresa\04 Clientes\_RFC_\Facturas\_MMMM_\_TipoComprobante_
      //Logger.log('Carpeta encontrada: ' + carpetaDestino.name + ', ID: ' + carpetaDestino.id);
    } else {
      message = 'No se encontró la carpeta del año fiscal: ' + year;
      console.log(message);
    }
  } else {
    message = 'No se encontró la carpeta del contribuyente RFC: ' + rfc;
    console.log(message);
  }
  var carpetaDestino = {
    id: idCarpetaDestino.id,
    msg: message
  }
  return carpetaDestino;
}

/**
 * Obtiene los meta datos de un XML dado su id de archivo
 */
function getCfdi(idCfdiFile = '19l3m1KtuHa10lPnISmm9tTRy5BaMT5Vn') {
  //function getCfdi(idCfdiFile) {
  var xmlCfdi = DriveApp.getFileById(idCfdiFile);
  var pdfCfdi = {};
  var fileCfdi = {};

  // Verifica si es un archivo XML
  if (xmlCfdi.getMimeType() === 'text/xml' || xmlCfdi.getMimeType() === 'application/xml') {
    var splitFileName = xmlCfdi.getName().split(".");
    //Logger.log(xmlCfdi.getMimeType());
    // Obtenen el PDF del XML
    if (splitFileName[1] == 'xml') {
      var pdf = splitFileName[0] + '.pdf';
    } else {
      var pdf = splitFileName[0] + '.PDF';
    }

    // https://stackoverflow.com/questions/62915684/parse-utf-8-bom-xml-files-using-xmlservice-google-app-script
    var bytes = DriveApp.getFileById(idCfdiFile).getBlob().getBytes();
    // Remueve los bits del Boom
    var boom = bytes.splice(0, 3)
    //Guarda el archivo XML UTF8 en content
    if (boom[0] < 0) {
      //Logger.log('Con Boom: ' + boom);
      var xmlFile = bytes;
    } else {
      //Logger.log('Sin Boom');
      var xmlFile = DriveApp.getFileById(idCfdiFile).getBlob().getBytes();
    }
    var content = Utilities.newBlob(xmlFile).getDataAsString();
    // Muestra el contenido del XML:     Logger.log(content);
    // Analiza el la gramática del XML
    var document = XmlService.parse(content);
    //Logger.log(document);
    var root = document.getRootElement();
    //Logger.log('Root element: '+root);
    // Verifica si es un CFDI
    var uriCFDI = root.getNamespace().getURI();
    var itemCfdi = root.getChildren();
    var temp1 = [];
    var temp2 = [];

    //Elementos del CFDI:  ['Comprobante', 'Emisor', 'Receptor', 'Conceptos', 'Impuestos', 'Complemento']
    var indexItem = [];
    itemCfdi.forEach(item => {
      indexItem.push(item.getName());
    })
    //Logger.log(indexItem);

    // --- Sección: Comprobante --- //
    //Logger.log('*-- ' + root.getName() + ' --*');
    var comprobante = [];
    var headerComprobante = ['schemaLocation', 'Version', 'Fecha', 'Moneda', 'TipoCambio', 'Descuento', 'SubTotal', 'Total', 'FormaPago',
      'CondicionesDePago', 'TipoDeComprobante', 'MetodoPago', 'LugarExpedicion', 'Serie', 'Folio', 'NoCertificado', 'Certificado', 'Sello'];
    comprobante.push(headerComprobante);


    headerComprobante.forEach((elem, index) => {
      // console.log('Atributo: ' + elem);
      var attribComprobante = root.getAttribute(elem);
      if (attribComprobante != null) {
        //Logger.log(index + ') ' + attribComprobante.getName() + ': ' + attribComprobante.getValue());
        temp1[index] = attribComprobante.getValue();
      } else {
        temp1[index] = null;
      }
    });
    comprobante.push(temp1);
    temp1 = [];
    //Logger.log(comprobante);

    // ---Fin Sección: Comprobante --- //

    // --- Sección: Emisor --- //
    var idxEmisor = indexItem.indexOf('Emisor');
    //Logger.log('*-- ' + itemCfdi[idxEmisor].getName() + ' --*');
    var emisor = [];
    var headerEmisor = ['Rfc', 'Nombre', 'RegimenFiscal'];
    emisor.push(headerEmisor);
    var atribEmisor = itemCfdi[idxEmisor].getAttributes();
    atribEmisor.map(elem => {
      index = headerEmisor.indexOf(elem.getName());
      //Logger.log(elem.getName() + ': ' + elem.getValue())
      temp1[index] = elem.getValue();
      //Logger.log(temp1);
    });
    emisor.push(temp1);
    temp1 = [];
    //Logger.log(emisor);

    // --- Sección: Receptor --- //
    var idxReceptor = indexItem.indexOf('Receptor');
    //Logger.log('*-- ' + itemCfdi[idxReceptor].getName() + ' --*');
    var receptor = [];
    var headerReceptor = ['Rfc', 'Nombre', 'UsoCFDI'];
    receptor.push(headerReceptor);
    var atribReceptor = itemCfdi[idxReceptor].getAttributes();
    atribReceptor.map(elem => {
      index = headerReceptor.indexOf(elem.getName());
      //Logger.log(elem.getName() + ': ' +elem.getValue())
      temp1[index] = elem.getValue();
      //Logger.log(temp1);
    });
    receptor.push(temp1);
    temp1 = [];
    //Logger.log(receptor);


    // --- Sección: Complemento --- //

    var idxComp = indexItem.indexOf('Complemento')
    var complementoTimbre = [];
    if (idxComp >= 0) {

      //Logger.log('*-- ' + itemCfdi[idxComp].getName() + ' --*');
      var complemento = [];
      var complementos = itemCfdi[idxComp].getChildren();

      complementos.forEach(complemento => {
        if (complemento.getName() == 'TimbreFiscalDigital') {
          //Logger.log(complemento.getName());
          var headerTimbreFiscal = ['schemaLocation', 'Version', 'UUID', 'FechaTimbrado', 'RfcProvCertif', 'SelloCFD', 'NoCertificadoSAT', 'SelloSAT'];
          complementoTimbre.push(headerTimbreFiscal);
          var trimbreFiscal = complemento.getAttributes();
          trimbreFiscal.map(complemento => {
            index = headerTimbreFiscal.indexOf(complemento.getName());
            //Logger.log(index +') '+complemento.getName());

            temp1[index] = complemento.getValue();
          });
          complementoTimbre.push(temp1);
          temp1 = [];
        }
      });

    }

    // --- Fin Sección: Complemento --- //

    // Verifica si existe el archivo PDF del XML
    if (DriveApp.getFilesByName(pdf).hasNext()) {
      pdfCfdi = DriveApp.getFilesByName(pdf).next();

      fileCfdi = {
        xmlUuid: complementoTimbre[1][2],
        xmlVersion: comprobante[1][1],
        xmlDate: comprobante[1][2],
        xmlId: xmlCfdi.getId(),
        xmlName: xmlCfdi.getName(),
        xmlType: xmlCfdi.getMimeType(),
        xmlUrl: xmlCfdi.getUrl(),
        xmlIdParentFolder: xmlCfdi.getParents().next().getId(),
        xmlParentFolder: xmlCfdi.getParents().next().getName(),
        xmlUrlParentFolder: xmlCfdi.getParents().next().getUrl(),
        pdfId: pdfCfdi.getId(),
        pdfName: pdfCfdi.getName(),
        pdfType: pdfCfdi.getMimeType(),
        pdfUrl: pdfCfdi.getUrl(),
        pdfIdParentFolder: pdfCfdi.getParents().next().getId(),
        pdfParentFolder: pdfCfdi.getParents().next().getName(),
        pdfUrlParentFolder: pdfCfdi.getParents().next().getUrl(),
        xmlRfcEmisor: emisor[1][0],
        xmlRfcNameEmisor: emisor[1][1],
        xmlRfcReceptor: receptor[1][0],
        xmlRfcNameReceptor: receptor[1][1]
      }
    } else {
      // Si solo existe el XML 
      fileCfdi = {
        xmlUuid: complementoTimbre[1][2],
        xmlVersion: comprobante[1][1],
        xmlDate: comprobante[1][2],
        xmlId: xmlCfdi.getId(),
        xmlName: xmlCfdi.getName(),
        xmlType: xmlCfdi.getMimeType(),
        xmlUrl: xmlCfdi.getUrl(),
        xmlIdParentFolder: xmlCfdi.getParents().next().getId(),
        xmlParentFolder: xmlCfdi.getParents().next().getName(),
        xmlUrlParentFolder: xmlCfdi.getParents().next().getUrl(),
        pdfId: null,
        pdfName: null,
        pdfType: null,
        pdfUrl: null,
        pdfIdParentFolder: null,
        pdfParentFolder: null,
        pdfUrlParentFolder: null,
        xmlRfcEmisor: emisor[1][0],
        xmlRfcNameEmisor: emisor[1][1],
        xmlRfcReceptor: receptor[1][0],
        xmlRfcNameReceptor: receptor[1][1]
      }
    }
  }
  //Logger.log(fileCfdi.xmlRfc + ' | ' + fileCfdi.xmlUuid + ' | ' + fileCfdi.xmlName + ' | ' + fileCfdi.pdfName + ' | ' + fileCfdi.xmlId + ' | ' + fileCfdi.pdfId)
  //console.log(fileCfdi);
  return fileCfdi;
}

/**
 * Limpiar Datos de hoja
 */
function limpiarDatos(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastRow = ss.getLastRow(); var lastCol = ss.getLastColumn();
  var rows = lastRow - 1;

  var StatusLimpiarDatos = {
    title: 'Limpiando datos de la hoja: ' + sheetName,
    data: [sheetName],
    status: false
  };

  // Limpia los datos de la hoja
  if (rows > 0) {
    ss.getRange(2, 1, rows, lastCol).clearContent();
    ss.getRange("A2").activateAsCurrentCell();
  }

  StatusLimpiarDatos.status = true;
  return StatusLimpiarDatos;
}

/**
 * Limpiar DeclaraSAT
 */
function limpiarDeclaraSAT(sheetName = 'calculoDeclaraSAT', month = '19-anual') {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssLimpiarDeclaraSat = ss.getSheetByName(sheetName);

  // Encabezado: Datos del Contribuyente RFC
  var headerCalDeclaraSat = [['RFC', ''],
  ['Razón Social', ''],
  ['Regimen Fiscal', ''],
  ['Correo', ''],
  ['Teléfono', '']];
  ssLimpiarDeclaraSat.getRange(3, 3, headerCalDeclaraSat.length, 2).setValues(headerCalDeclaraSat);
  //Fecha y hora de cálculo
  ssLimpiarDeclaraSat.getRange("H3").setValue('');
  ssLimpiarDeclaraSat.getRange("H4").setValue('');

  // Obteniendo información del contribuyente
  var listaMeses = ['01-ene', '02-feb', '03-mar', '04-abr', '05-may', '06-jun', '07-jul', '08-ago', '09-sep', '10-oct', '11-nov', '12-dic'];
  var startRow = 9;
  var lastRow = ssLimpiarDeclaraSat.getRange('A:A').getLastRow();

  if (month == '19-anual') {
    console.log('Limpiando Todos los Meses');
    var lastColumn = 13;
    console.log(lastColumn)
    ssLimpiarDeclaraSat.getRange(startRow, 2, lastRow - startRow, lastColumn - 1).clearContent();
  } else {
    console.log('Limpiando el mes: ' + month);
    var numMes = listaMeses.indexOf(month);
    var col = numMes + 2;
    // Limpiando la columna del mes correspondiente
    ssLimpiarDeclaraSat.getRange(startRow, col, lastRow - startRow, 1).clearContent();
  }
}


/**
 * Obtiene los meta datos de una subcarpeta dado el id de su carpeta padre
 */
function getCarpeta(idSearchFolder, folder) {
  // Busca dentro de una carpeta por su id
  var carpetas = DriveApp.getFolderById(idSearchFolder).getFolders();
  var dataCarpeta = {};

  while (carpetas.hasNext()) {
    var carpeta = carpetas.next();
    if (carpeta.getName() === folder) {
      dataCarpeta = {
        id: carpeta.getId(),
        name: carpeta.getName(),
        size: carpeta.getSize(),
        url: carpeta.getUrl(),
        idParentFolder: carpeta.getParents().next().getId(),
        parentFolder: carpeta.getParents().next().getName(),
        urlParentFolder: carpeta.getParents().next().getUrl()
      }
    }
  }
  //Logger.log(dataCarpeta.name);
  return dataCarpeta;
}


/**
 * Obtiene los meta datos de subcarpetas dentro de una carpeta
 */
function getCarpetas(idSearchFolder) {
  // idSearchFolder = '11pw8STEqlIX01A4oODaD_Lf1m-wRTbls'
  // Busca dentro de una carpeta por su id
  var carpetas = DriveApp.getFolderById(idSearchFolder).getFolders();
  //Logger.log(carpetas)
  var listCarpetas = [];

  var index = 0;
  while (carpetas.hasNext()) {
    var carpeta = carpetas.next();

    dataCarpeta = {
      id: carpeta.getId(),
      name: carpeta.getName(),
      size: carpeta.getSize(),
      url: carpeta.getUrl(),
      idParentFolder: carpeta.getParents().next().getId(),
      parentFolder: carpeta.getParents().next().getName(),
      urlParentFolder: carpeta.getParents().next().getUrl()
    }
    listCarpetas[index] = dataCarpeta;
    index++;
  }
  //Logger.log(listCarpetas.length + ' carpeta(s) encontrada(s)')
  //Logger.log(listCarpetas[0]);
  return listCarpetas;
}


/**
 * Obtiene el número de archivos dentro de una carpeta
 */
function getNumCFDIs(idSearchFolder = '13IZ6HijhKGzwf4Y5buxS_ik9WOGGHBE3') {
  // idSearchFolder = '11pw8STEqlIX01A4oODaD_Lf1m-wRTbls'

  // Identifica la carpeta RFC del cliente
  var folder = DriveApp.getFolderById(idSearchFolder);
  var getXmlCfdiTxt = folder.searchFiles('mimeType = "text/xml" ');
  var getXmlCfdiApp = folder.searchFiles('mimeType = "application/xml" ');

  var index = 0;
  // Cuenta los XMLs tipo txt
  while (getXmlCfdiTxt.hasNext()) {
    getXmlCfdiTxt.next();
    //console.log(index +') XML txt: ');
    index++;
  }
  // Cuenta los XMLs tipo txt
  while (getXmlCfdiApp.hasNext()) {
    getXmlCfdiApp.next();
    //console.log(index +') XML app: ');
    index++;
  }
  //console.log('Total XML de CFDIs: ' + index);
  return index;
}

/**
 * Obtiene el número de archivos dentro de una carpeta
 */
function getNumArchivos(idSearchFolder = '13IZ6HijhKGzwf4Y5buxS_ik9WOGGHBE3') {
  // idSearchFolder = '11pw8STEqlIX01A4oODaD_Lf1m-wRTbls'
  // Busca dentro de una carpeta por su id
  var archivos = DriveApp.getFolderById(idSearchFolder).getFiles();
  var index = 0;
  while (archivos.hasNext()) {
    archivos.next();
    //console.log(archivo);
    index++;
  }
  //console.log('Total de archivos: ' + index);
  return index;
}

function getArchivo(idSearchFolder, fileName) {
  // Identifica la carpeta de contenedor de CFDIs
  var folder = DriveApp.getFolderById(idSearchFolder);
  var archivos = folder.getFilesByName(fileName);
  if (archivos.hasNext()) {
    var archivo = archivos.next();

    var dataArchivo = {
      id: archivo.getId(),
      name: archivo.getName(),
      size: archivo.getSize(),
      extension: archivo.getMimeType(),
      url: archivo.getUrl(),
      idParentFolder: archivo.getParents().next().getId(),
      parentFolder: archivo.getParents().next().getName(),
      urlParentFolder: archivo.getParents().next().getUrl()
    }
  }
  //console.log(dataArchivo);
  return dataArchivo;

}

/**
 * Obtiene los meta datos de archivos dentro de una carpeta
 */
function getArchivos(idSearchFolder = '13IZ6HijhKGzwf4Y5buxS_ik9WOGGHBE3') {
  // idSearchFolder = '11pw8STEqlIX01A4oODaD_Lf1m-wRTbls'
  // Busca dentro de una carpeta por su id
  var archivos = DriveApp.getFolderById(idSearchFolder).getFiles();
  //Logger.log(archivos)
  var listaArchivos = [];

  var index = 0;
  while (archivos.hasNext()) {
    var archivo = archivos.next();

    dataArchivo = {
      id: archivo.getId(),
      name: archivo.getName(),
      size: archivo.getSize(),
      extension: archivo.getMimeType(),
      url: archivo.getUrl(),
      idParentFolder: archivo.getParents().next().getId(),
      parentFolder: archivo.getParents().next().getName(),
      urlParentFolder: archivo.getParents().next().getUrl()
    }
    listaArchivos[index] = dataArchivo;
    index++;
  }
  //Logger.log(listaArchivos.length + ' carpeta(s) encontrada(s)')
  //Logger.log(listaArchivos[0]);
  return listaArchivos;
}

// ----------- FIN OBTENCIÓN DE CARPETAS ----------- //


