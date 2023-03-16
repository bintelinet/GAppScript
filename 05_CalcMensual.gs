/**
 * Suma de Ingresos/Egresos, Imp. Retenidos e Imp.Trasladados
 */
function calculoMensual(rfc, year, month, tarifaSat) {

  var ssCalDeclaraSat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('calculoDeclaraSAT');
  var ssRepoCfdis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(month + '-reposCFDI').getDataRange().getValues();
  var ssDataCfdis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(month + '-dataCFDI').getDataRange().getValues();
  var ssCalculosCfdis = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('calculosCFDI');

  var ANIO = ssDataCfdis[0].indexOf('Año');
  var TIPO_CFDI = ssDataCfdis[0].indexOf('Tipo CFDI');
  var RFC_EMISOR = ssDataCfdis[0].indexOf('RFC Emisor');
  var RFC_RECEPTOR = ssDataCfdis[0].indexOf('RFC Receptor');
  var TIPOCOMPROBANTE = ssDataCfdis[0].indexOf('Tipo de Comprobante');
  var METODOPAGO = ssDataCfdis[0].indexOf('Método de Pago');
  var SUBTOTAL = ssDataCfdis[0].indexOf('SubTotal');
  var DESCUENTO = ssDataCfdis[0].indexOf('Descuento');
  var TOTAL = ssDataCfdis[0].indexOf('Total');
  var TIPOCFDI = ssDataCfdis[0].indexOf('Tipo CFDI');
  var ISR_RETENIDO = ssDataCfdis[0].indexOf('ISR Retenido');
  var IVA_RETENIDO = ssDataCfdis[0].indexOf('IVA Retenido');
  var IEPS_RETENIDO = ssDataCfdis[0].indexOf('IEPS Retenido');
  var TOTAL_IMP_RETENIDOS = ssDataCfdis[0].indexOf('Total Imp. Retenidos');
  var TIPOFACTOR_IVA = ssDataCfdis[0].indexOf('Tipo Factor IVA');
  var TIPOFACTOR_IEPS = ssDataCfdis[0].indexOf('Tipo Factor IEPS');
  var TASA_IVA = ssDataCfdis[0].indexOf('Tasa IVA Trasladado');
  var TASA_IEPS = ssDataCfdis[0].indexOf('Tasa IEPS Trasladado');
  var IVA_TRASLADADO = ssDataCfdis[0].indexOf('IVA Trasladado');
  var IEPS_TRASLADADO = ssDataCfdis[0].indexOf('IEPS Trasladado');
  var TOTAL_IMP_TRASLADADOS = ssDataCfdis[0].indexOf('Total Imp. Trasladados');
  ssDataCfdis.shift();

  // Obteniendo información del contribuyente
  var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'],
  ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];
  var nombMes = ''; var numMes = 0; var periodoDeclara = '';

  for (i = 0; i < listaMeses.length; i++) {
    if (listaMeses[i][0] == month) {
      nombMes = listaMeses[i][1];
      numMes = i;
      periodoDeclara = 'CÁLCULO DE LA DECLARACIÓN DE ' + nombMes.toUpperCase() + ' DE ' + year;
      var col = numMes + 2;
      // Limpiando la columna del mes correspondiente
      ssCalDeclaraSat.getRange(9, col, 35, 1).clearContent();
      break;
    } else {
      periodoDeclara = 'La periodicidad de declaración es incorrecta ';
    }
  }
  console.log(periodoDeclara);

  // Calculando el ISR
  var baseGravableISR = 0; var baseGravableIVA = 0; var tasaIVA = '0.16';
  var actividadesGravadas = 0; var limiteInferior = 0; var baseImpuesto = 0; var tasa = 0; var impuestoMarginal = 0;
  var cuotaFija = 0; var importeIsr = 0; var isrPagar = 0; var ivaPagar = 0; ivaFavor = 0;
  var isrRetenido = 0; var ivaRetenido = 0; var iepsRetenido = 0; var ivaAcreditable = 0; var ivaCobrado = 0;
  var ivaTrasladado = 0; var iepsTrasladado = 0; var baseCalculoISR = 0; var deducciones = 0;

  var totalIEcfdis = [['CFDIs', 'Emitidas', 'Recibidas'],
  ['Ingresos', 0, 0],
  ['Egresos', 0, 0],
  ['Total', 0, 0]];

  // var tipoCfdis = [['CFDI', 'SubTotal', 'Descuento', 'Total'],
  // ['Ingresos PUE', [1][1], [1][2], [1][3],
  // ['Ingresos PPD', [2][1], [2][2], [2][3],
  // ['Total Emitidas', [3][1], [3][2], [3][3],    // Emitidas
  // ['Egresos PUE', [4][1], [4][2], [4][3],
  // ['Egresos PPD', [5][1], [5][2], [4][3],
  // ['Total Recibidas', [6][1], [6][2], [6][3],   // Recibidas

  var tipoCfdisI = [['CFDI Tipo I', 'SubTotal', 'Descuento', 'Total'],
  ['Ingresos PUE', 0, 0, 0],
  ['Ingresos PPD', 0, 0, 0],
  ['Total Ingresos', 0, 0, 0],    // Emitidas
  ['Egresos PUE', 0, 0, 0],
  ['Egresos PPD', 0, 0, 0],
  ['Total Deducciones', 0, 0, 0]]; // Recibidas

  var tipoCfdisE = [['CFDI Tipo E', 'SubTotal', 'Descuento', 'Total'],
  ['Ingresos PUE', 0, 0, 0],
  ['Ingresos PPD', 0, 0, 0],
  ['Total Ingresos', 0, 0, 0],    // Emitidas
  ['Egresos PUE', 0, 0, 0],
  ['Egresos PPD', 0, 0, 0],
  ['Total Deducciones', 0, 0, 0]]; // Recibidas

  // var impRetenidosCfdis = [
  //   ['Pago', 'ISR Retenido', 'IVA Retenido', 'IEPS Retenido', 'Total Imp. Retenidos'],
  //   ['ImpRet PUE', [1][1], [1][2], [1][3], [1][4]],                // Emitidas
  //   ['ImpRet PPD', [2][1], [2][2], [2][3], [2][4]],                // Emitidas
  //   ['Total ImpRet Emitidos', [3][1], [3][2], [3][3], [3][4]],     // Emitidas
  //   ['ImpRet PUE', [4][1], [4][2], [4][3], [4][4]],                // Recibidas
  //   ['ImpRet PPD', [5][1], [5][2], [5][3], [5][4]],                // Recibidas
  //   ['Total ImpRet Recibidos', [6][1], [6][2], [6][3], [6][4]]];   // Recibidas  

  var impRetenidosCfdisI = [
    ['Pago', 'ISR Retenido', 'IVA Retenido', 'IEPS Retenido', 'Total Imp. Ret.'],
    ['ImpRet PUE', 0, 0, 0, 0],                // Emitidas
    ['ImpRet PPD', 0, 0, 0, 0],                // Emitidas
    ['Total ImpRet Emitidos', 0, 0, 0, 0],     // Emitidas
    ['ImpRet PUE', 0, 0, 0, 0],                // Recibidas
    ['ImpRet PPD', 0, 0, 0, 0],                // Recibidas
    ['Total ImpRet Recibidos', 0, 0, 0, 0]];   // Recibidas

  var impRetenidosCfdisE = [
    ['Pago', 'ISR Retenido', 'IVA Retenido', 'IEPS Retenido', 'Total Imp. Ret.'],
    ['ImpRet PUE', 0, 0, 0, 0],                // Emitidas
    ['ImpRet PPD', 0, 0, 0, 0],                // Emitidas
    ['Total ImpRet Emitidos', 0, 0, 0, 0],     // Emitidas
    ['ImpRet PUE', 0, 0, 0, 0],                // Recibidas
    ['ImpRet PPD', 0, 0, 0, 0],                // Recibidas
    ['Total ImpRet Recibidos', 0, 0, 0, 0]];   // Recibidas 

  // var impTrasladadosCfdisI = [
  //   ['Pago','FactorIVA','FactorEPS','IVA16%','IVA0%','IVAExcento','IEPS%','IEPSExcento', 'Total Imp. Trasladados'],
  //   ['ImpTras PUE', [1][1], [1][2], [1][3], [1][4], [1][5], [1][6], [1][7], [1][8]],                     // Emitidas
  //   ['ImpTras PPD', [2][1], [2][2], [2][3], [2][4], [2][5], [2][6], [2][7]],[2][8],                      // Emitidas
  //   ['Total ImpTras Emitidos', [3][1], [3][2], [3][3], [3][4], [3][5], [3][6], [3][7],[3][8]],           // Emitidas
  //   ['ImpTras PUE', [4][1], [4][2], [4][3], [4][4], [4][5], [4][6], [4][7],[4][8]],                      // Recibidas
  //   ['ImpTras PPD', [5][1], [5][2], [5][3], [5][4], [5][5], [5][6], [5][7],[5][8]],	                    // Recibidas
  //   ['Total ImpTras Recibidos', [6][1], [6][2], [6][3], [6][4], [6][5], [6][6], [6][7], [6][8]]];        // Recibidas

  var impTrasladadosCfdisI = [
    ['Pago', 'Tasa IVA', 'Tasa IEPS', 'IVA 16%', 'IVA 0%', 'IVA Excento', 'IEPS %', 'IEPS Excento', 'Total Imp. Tras.'],
    ['ImpTras PUE', '', '', 0, 0, 0, 0, 0, 0],               // Emitidas
    ['ImpTras PPD', '', '', 0, 0, 0, 0, 0, 0],               // Emitidas
    ['Total ImpTras Emitidos', '', '', 0, 0, 0, 0, 0, 0],    // Emitidas
    ['ImpTras PUE', '', '', 0, 0, 0, 0, 0, 0],               // Recibidas
    ['ImpTras PPD', '', '', 0, 0, 0, 0, 0, 0],               // Recibidas
    ['Total ImpTras Recibidos', '', '', 0, 0, 0, 0, 0, 0]];  // Recibidas

  var impTrasladadosCfdisE = [
    ['Pago', 'Tasa IVA', 'Tasa IEPS', 'IVA 16%', 'IVA 0%', 'IVA Excento', 'IEPS %', 'IEPS Excento', 'Total Imp. Tras.'],
    ['ImpTras PUE', '', '', 0, 0, 0, 0, 0, 0],               // Emitidas
    ['ImpTras PPD', '', '', 0, 0, 0, 0, 0, 0],               // Emitidas
    ['Total ImpTras Emitidos', '', '', 0, 0, 0, 0, 0, 0],    // Emitidas
    ['ImpTras PUE', '', '', 0, 0, 0, 0, 0, 0],               // Recibidas
    ['ImpTras PPD', '', '', 0, 0, 0, 0, 0, 0],               // Recibidas
    ['Total ImpTras Recibidos', '', '', 0, 0, 0, 0, 0, 0]];  // Recibidas

  var actGravadasxIngresosCfdisI = [['Método Pago', 'Tasa', 'Tasa 0', 'Excentas'],
  ['PUE', 0, 0, 0],
  ['PPD', 0, 0, 0],
  ['Total', 0, 0, 0]];

  var actGravadasxEgresosCfdisI = [['Método Pago', 'Tasa', 'Tasa 0', 'Excentas'],
  ['PUE', 0, 0, 0],
  ['PPD', 0, 0, 0],
  ['Total', 0, 0, 0]];

  // var dataCfdis = [['Pago', 'Emitidas', 'Recibidas'],
  // ['PUE', [1][1], [1][2]],
  // ['PPD', [2][1], [2][2]],
  // ['Total', [3][1], [3][2]]];

  var dataCfdisI = [['Pago', 'Emitidas', 'Recibidas'],
  ['PUE', 0, 0],
  ['PPD', 0, 0],
  ['Total', 0, 0]];

  var dataCfdisE = [['Pago', 'Emitidas', 'Recibidas'],
  ['PUE', 0, 0],
  ['PPD', 0, 0],
  ['Total', 0, 0]];


  // ------ Deducciones Autorizadas -------- //
  // IVA Acreditable
  //      IVA por Acreditar (Por pagar)   |  IVA Acreditable (Pagado)
  // IVA Trasladado
  //      IVA por Trasladar (Por pagar)   |  IVA Trasladado (Cobrado)

  var DEDUCIBLE = ssRepoCfdis[0].indexOf('Deducible');
  ssRepoCfdis.shift();

  //console.log(ssDataCfdis[index][TIPO_CFDI] + ' ' + ssDataCfdis[TIPOCOMPROBANTE])
  // Verifica si los datos de xx-mes-dataCFDI pertenecen al RFC seleccionado
  var datosValidos = true;
  for (index in ssDataCfdis) {

    switch (ssDataCfdis[index][TIPO_CFDI]) {
      case 'emitidas':
        if (ssDataCfdis[index][RFC_EMISOR] == rfc && ssDataCfdis[index][ANIO] == year && datosValidos == true) {
          //console.log(index + ') RFC Contribuyente Válido');
          switch (ssDataCfdis[index][TIPOCOMPROBANTE]) {
            case 'I': // Ingreso 
              // Total Ingresos CFDIs emitidas
              totalIEcfdis[1][1] += Number(ssDataCfdis[index][SUBTOTAL]);
              break;
            case 'E': // Egreso
              // Total Egresos CFDIs emitidas
              totalIEcfdis[2][1] += Number(ssDataCfdis[index][SUBTOTAL]);
              break;
          }
        } else {
          console.log(index + ') RFC Contribuyente Inválido');
          totalIEcfdis = [['Total', 'Emitidas', 'Recibidas'],
          ['Ingresos', 0, 0],
          ['Egresos', 0, 0],
          ['Total', 0, 0]];
          datosValidos = false;
          break;
        }
        break;
      case 'recibidas':
        if (ssDataCfdis[index][RFC_RECEPTOR] == rfc && ssDataCfdis[index][ANIO] == year && datosValidos == true) {
          //console.log(index + ') RFC Contribuyente Válido');
          switch (ssDataCfdis[index][TIPOCOMPROBANTE]) {
            case 'I': // Ingreso 
              // Total Ingresos CFDIs emitidas
              totalIEcfdis[1][2] += Number(ssDataCfdis[index][SUBTOTAL]);
              break;
            case 'E': // Egreso
              // Total Egresos CFDIs emitidas
              totalIEcfdis[2][2] += Number(ssDataCfdis[index][SUBTOTAL]);
              break;
          }
        } else {
          //console.log(index + ') RFC Contribuyente Inválido');
          totalIEcfdis = [['Total', 'Emitidas', 'Recibidas'],
          ['Ingresos', 0, 0],
          ['Egresos', 0, 0],
          ['Total', 0, 0]];
          datosValidos = false;
          break;
        }
        break;
    }
  }
  // Total Ingresos Emitidas
  totalIEcfdis[3][1] = totalIEcfdis[1][1] - totalIEcfdis[2][1];
  // Total Egresos Recibidas
  totalIEcfdis[3][2] = totalIEcfdis[1][2] - totalIEcfdis[2][2];

  // Si el datos extraídos de los CFDIs pertenence al RFC seleccionado
  // realiza ejecuta los cálculos
  if (datosValidos == true) {
    // Total Ingresos Emitidas

    ssDataCfdis.forEach((cfdi, index) => {
      // Verifica si es deducible el CFDI
      if (ssRepoCfdis[index][DEDUCIBLE] === 'Si') {

        switch (cfdi[TIPOCOMPROBANTE]) {
          case 'I': // Ingreso   <<<<---------------------------------------------------------------------------

            switch (cfdi[TIPOCFDI]) {

              case 'emitidas':
                dataCfdisI[3][1]++;
                //console.log(dataCfdisI[3][1] + ') CFDIs emitidos');

                switch (cfdi[METODOPAGO]) {
                  case 'PUE': //console.log(cfdi[METODOPAGO])
                    // Contabiliza los CFDIs PUE totalmente cobrados
                    dataCfdisI[1][1]++;
                    // SubTotal
                    tipoCfdisI[1][1] = Number(cfdi[SUBTOTAL]); //console.log('Subtotal PUE: ' + tipoCfdisI[1][1]);
                    // Descuento
                    tipoCfdisI[1][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisI[1][2]);
                    // Total
                    tipoCfdisI[1][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisI[1][3]);

                    // -------- EMITIDAS: IMPUESTOS RETENIDOS PUE -------- //

                    // ISR Retenido
                    impRetenidosCfdisI[1][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PUE: ' + impRetenidosCfdisI[1][1]);
                    // IVA Retenido
                    impRetenidosCfdisI[1][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PUE: ' + impRetenidosCfdisI[1][2]);
                    // IEPS Retenido
                    impRetenidosCfdisI[1][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PUE: ' + impRetenidosCfdisI[1][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisI[1][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PUE: ' + impRetenidosCfdisI[1][4]);

                    // -------- FIN EMITIDAS: IMPUESTOS RETENIDOS PUE -------- //


                    // -------- EMITIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    impTrasladadosCfdisI[1][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {

                      case 'Tasa': //Tasa IVA                      
                        //console.log('---> ' + impTrasladadosCfdisI[1][1]);
                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisI[1][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisI[1][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[1][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisI[1][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisI[1][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisI[1][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisI[1][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[1][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');                    
                        impTrasladadosCfdisI[1][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {
                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisI[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[1][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisI[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[1][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisI[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[1][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');
                    }

                    // -------- FIN EMITIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    break; // FIN case 'PUE'

                  case 'PPD': //console.log(cfdi[METODOPAGO]);
                    // Contabiliza los CFDIs PUE totalmente cobrados
                    dataCfdisI[2][1]++;
                    // SubTotal
                    tipoCfdisI[2][1] += Number(cfdi[SUBTOTAL]); //console.log('Subtotal PPD: ' + tipoCfdisI[2][1]);
                    // Descuento
                    tipoCfdisI[2][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PPD: ' + tipoCfdisI[2][2]);
                    // Total
                    tipoCfdisI[2][3] += Number(cfdi[TOTAL]); //console.log('Total PPD: ' + tipoCfdisI[2][3]);

                    // -------- EMITIDAS: IMPUESTOS RETENIDOS PPD -------- //

                    // ISR Retenido
                    impRetenidosCfdisI[2][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PPD: ' + impRetenidosCfdisI[2][1]);
                    // IVA Retenido
                    impRetenidosCfdisI[2][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PPD: ' + impRetenidosCfdisI[2][2]);
                    // IEPS Retenido
                    impRetenidosCfdisI[2][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PPD: ' + impRetenidosCfdisI[2][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisI[2][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PPD: ' + impRetenidosCfdisI[2][4]);                  


                    // -------- FIN EMITIDAS: IMPUESTOS RETENIDOS PPD -------- //


                    // -------- EMITIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    impTrasladadosCfdisI[2][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);


                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA

                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisI[2][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisI[2][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[2][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisI[2][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisI[2][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisI[2][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisI[2][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[2][3]);
                        }

                        break;
                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        impTrasladadosCfdisI[2][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisI[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[2][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisI[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[2][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisI[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[2][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN EMITIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    break; // FIN case 'PPD'

                  default:
                    console.log('Indique un método de pago')
                } // Fin switch (cfdi[METODOPAGO])
                break;

              case 'recibidas':
                dataCfdisI[3][2]++;
                //console.log(totalCfdisRecibidas + ') CFDIs recibidos');

                switch (cfdi[METODOPAGO]) {
                  case 'PUE':
                    // Contabiliza los CFDIs totalmente pagados
                    dataCfdisI[1][2]++;
                    // SubTotal
                    tipoCfdisI[4][1] += Number(cfdi[SUBTOTAL]); //console.log('SubTotal PUE: ' + tipoCfdisI[4][1]);
                    // Descuento
                    tipoCfdisI[4][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisI[4][2]);
                    // Total
                    tipoCfdisI[4][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisI[4][3]);


                    // -------- RECIBIDAS: IMPUESTOS RETENIDOS PUE -------- //
                    impTrasladadosCfdisI[4][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // ISR Retenido
                    impRetenidosCfdisI[4][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PUE: ' + impRetenidosCfdisI[4][1]);
                    // IVA Retenido
                    impRetenidosCfdisI[4][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PUE: ' + impRetenidosCfdisI[4][2]);
                    // IEPS Retenido
                    impRetenidosCfdisI[4][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PUE: ' + impRetenidosCfdisI[4][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisI[4][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PUE: ' + impRetenidosCfdisI[4][4]);

                    // -------- FIN RECIBIDAS: IMPUESTOS RETENIDOS PUE -------- //           


                    // -------- RECIBIDAS: IMPUESTOS TRASLADADOS PUE -------- //

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA
                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisI[4][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisI[4][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[4][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisI[4][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisI[4][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisI[4][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisI[4][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[4][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IVA');                      
                      //impTrasladadosCfdisI[4][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisI[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[4][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisI[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[4][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisI[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[4][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN RECIBIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    break;

                  case 'PPD':
                    //console.log(cfdi[METODOPAGO]);
                    dataCfdisI[2][2]++;
                    // SubTotal
                    tipoCfdisI[5][1] += Number(cfdi[SUBTOTAL]); //console.log('SubTotal PUE: ' + tipoCfdisI[5][1]);
                    // Descuento
                    tipoCfdisI[5][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisI[5][2]);
                    // Total
                    tipoCfdisI[5][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisI[5][3]);

                    // -------- RECIBIDAS: IMPUESTOS RETENIDOS PPD -------- //
                    impTrasladadosCfdisI[5][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // ISR Retenido
                    impRetenidosCfdisI[5][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PPD: ' + impRetenidosCfdisI[5][1]);
                    // IVA Retenido
                    impRetenidosCfdisI[5][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PPD: ' + impRetenidosCfdisI[5][2]);
                    // IEPS Retenido
                    impRetenidosCfdisI[5][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PPD: ' + impRetenidosCfdisI[5][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisI[5][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PPD: ' + impRetenidosCfdisI[5][4]);  

                    // -------- FIN RECIBIDAS: IMPUESTOS RETENIDOS PPD -------- //


                    // -------- RECIBIDAS: IMPUESTOS TRASLADADOS PPD -------- //

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA

                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisI[5][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisI[5][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[5][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisI[5][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisI[5][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisI[5][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisI[5][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisI[5][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        impTrasladadosCfdisI[5][5] += Number(cfdi[IVA_TRASLADADO]);
                    }

                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisI[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[5][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisI[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[5][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisI[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisI[5][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN RECIBIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    break;
                  default:
                    console.log('Indique un método de pago')
                } // Fin switch (cfdi[METODOPAGO])
                break;
              default:
                console.log('Indique un tipo de CFDI');
            } // Fin switch (cfdi[TIPOCFDI])
            break;

          case 'E':	// Egreso  <<<<---------------------------------------------------------------------------
            switch (cfdi[TIPOCFDI]) {

              case 'emitidas':
                dataCfdisE[3][1]++;
                //console.log(dataCfdisE[3][1] + ') CFDIs emitidos');

                switch (cfdi[METODOPAGO]) {
                  case 'PUE': //console.log(cfdi[METODOPAGO])
                    // Contabiliza los CFDIs PUE totalmente cobrados
                    dataCfdisE[1][1]++;
                    // SubTotal
                    tipoCfdisE[1][1] = Number(cfdi[SUBTOTAL]); //console.log('Subtotal PUE: ' + tipoCfdisE[1][1]);
                    // Descuento
                    tipoCfdisE[1][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisE[1][2]);
                    // Total
                    tipoCfdisE[1][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisE[1][3]);

                    // -------- EMITIDAS: IMPUESTOS RETENIDOS PUE -------- //

                    // ISR Retenido
                    impRetenidosCfdisE[1][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PUE: ' + impRetenidosCfdisE[1][1]);
                    // IVA Retenido
                    impRetenidosCfdisE[1][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PUE: ' + impRetenidosCfdisE[1][2]);
                    // IEPS Retenido
                    impRetenidosCfdisE[1][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PUE: ' + impRetenidosCfdisE[1][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisE[1][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PUE: ' + impRetenidosCfdisE[1][4]);

                    // -------- FIN EMITIDAS: IMPUESTOS RETENIDOS PUE -------- //


                    // -------- EMITIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    impTrasladadosCfdisE[1][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {

                      case 'Tasa': //Tasa IVA                      
                        //console.log('---> ' + impTrasladadosCfdisE[1][1]);
                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisE[1][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisE[1][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[1][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisE[1][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisE[1][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisE[1][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisE[1][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[1][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');                    
                        impTrasladadosCfdisE[1][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {
                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisE[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[1][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisE[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[1][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisE[1][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[1][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');
                    }

                    // -------- FIN EMITIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    break; // FIN case 'PUE'

                  case 'PPD': //console.log(cfdi[METODOPAGO]);
                    // Contabiliza los CFDIs PUE totalmente cobrados
                    dataCfdisE[2][1]++;
                    // SubTotal
                    tipoCfdisE[2][1] += Number(cfdi[SUBTOTAL]); //console.log('Subtotal PPD: ' + tipoCfdisE[2][1]);
                    // Descuento
                    tipoCfdisE[2][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PPD: ' + tipoCfdisE[2][2]);
                    // Total
                    tipoCfdisE[2][3] += Number(cfdi[TOTAL]); //console.log('Total PPD: ' + tipoCfdisE[2][3]);

                    // -------- EMITIDAS: IMPUESTOS RETENIDOS PPD -------- //

                    // ISR Retenido
                    impRetenidosCfdisE[2][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PPD: ' + impRetenidosCfdisE[2][1]);
                    // IVA Retenido
                    impRetenidosCfdisE[2][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PPD: ' + impRetenidosCfdisE[2][2]);
                    // IEPS Retenido
                    impRetenidosCfdisE[2][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PPD: ' + impRetenidosCfdisE[2][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisE[2][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PPD: ' + impRetenidosCfdisE[2][4]);                  


                    // -------- FIN EMITIDAS: IMPUESTOS RETENIDOS PPD -------- //


                    // -------- EMITIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    impTrasladadosCfdisE[2][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);


                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA

                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisE[2][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisE[2][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[2][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisE[2][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisE[2][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisE[2][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisE[2][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[2][3]);
                        }

                        break;
                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        impTrasladadosCfdisE[2][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisE[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[2][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisE[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[2][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisE[2][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[2][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN EMITIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    break; // FIN case 'PPD'

                  default:
                    console.log('Indique un método de pago')
                } // Fin switch (cfdi[METODOPAGO])
                break;

              case 'recibidas':
                dataCfdisE[3][2]++;
                //console.log(totalCfdisRecibidas + ') CFDIs recibidos');

                switch (cfdi[METODOPAGO]) {
                  case 'PUE':
                    // Contabiliza los CFDIs totalmente pagados
                    dataCfdisE[1][2]++;
                    // SubTotal
                    tipoCfdisE[4][1] += Number(cfdi[SUBTOTAL]); //console.log('SubTotal PUE: ' + tipoCfdisE[4][1]);
                    // Descuento
                    tipoCfdisE[4][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisE[4][2]);
                    // Total
                    tipoCfdisE[4][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisE[4][3]);


                    // -------- RECIBIDAS: IMPUESTOS RETENIDOS PUE -------- //
                    impTrasladadosCfdisE[4][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // ISR Retenido
                    impRetenidosCfdisE[4][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PUE: ' + impRetenidosCfdisE[4][1]);
                    // IVA Retenido
                    impRetenidosCfdisE[4][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PUE: ' + impRetenidosCfdisE[4][2]);
                    // IEPS Retenido
                    impRetenidosCfdisE[4][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PUE: ' + impRetenidosCfdisE[4][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisE[4][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PUE: ' + impRetenidosCfdisE[4][4]);

                    // -------- FIN RECIBIDAS: IMPUESTOS RETENIDOS PUE -------- //           


                    // -------- RECIBIDAS: IMPUESTOS TRASLADADOS PUE -------- //

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA
                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisE[4][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisE[4][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[4][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisE[4][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisE[4][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisE[4][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisE[4][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[4][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IVA');                      
                      //impTrasladadosCfdisE[4][5] += Number(cfdi[IVA_TRASLADADO]);
                    }


                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisE[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[4][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisE[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[4][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisE[4][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[4][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN RECIBIDAS: IMPUESTOS TRASLADADOS PUE -------- //
                    break;

                  case 'PPD':
                    //console.log(cfdi[METODOPAGO]);
                    dataCfdisE[2][2]++;
                    // SubTotal
                    tipoCfdisE[5][1] += Number(cfdi[SUBTOTAL]); //console.log('SubTotal PUE: ' + tipoCfdisE[5][1]);
                    // Descuento
                    tipoCfdisE[5][2] += Number(cfdi[DESCUENTO]); //console.log('Descuento PUE: ' + tipoCfdisE[5][2]);
                    // Total
                    tipoCfdisE[5][3] += Number(cfdi[TOTAL]); //console.log('Total PUE: ' + tipoCfdisE[5][3]);

                    // -------- RECIBIDAS: IMPUESTOS RETENIDOS PPD -------- //
                    impTrasladadosCfdisE[5][8] += Number(cfdi[TOTAL_IMP_TRASLADADOS]);

                    // ISR Retenido
                    impRetenidosCfdisE[5][1] += Number(cfdi[ISR_RETENIDO]); //console.log('ISR RETENIDO PPD: ' + impRetenidosCfdisE[5][1]);
                    // IVA Retenido
                    impRetenidosCfdisE[5][2] += Number(cfdi[IVA_RETENIDO]); //console.log('IVA RETENIDO PPD: ' + impRetenidosCfdisE[5][2]);
                    // IEPS Retenido
                    impRetenidosCfdisE[5][3] += Number(cfdi[IEPS_RETENIDO]); //console.log('IEPS RETENIDO PPD: ' + impRetenidosCfdisE[5][3]);
                    // Total Imp. Retenidos
                    impRetenidosCfdisE[5][4] += Number(cfdi[TOTAL_IMP_RETENIDOS]); //console.log('IMPUESTOS RETENIDOS PPD: ' + impRetenidosCfdisE[5][4]);  

                    // -------- FIN RECIBIDAS: IMPUESTOS RETENIDOS PPD -------- //


                    // -------- RECIBIDAS: IMPUESTOS TRASLADADOS PPD -------- //

                    // IVA Trasladado
                    switch (cfdi[TIPOFACTOR_IVA]) {
                      case 'Tasa': //Tasa IVA

                        if (Number(cfdi[TASA_IVA]) == tasaIVA) {
                          // Tasa IVA 16%
                          impTrasladadosCfdisE[5][1] = cfdi[TASA_IVA];
                          // IVA 16%                    
                          impTrasladadosCfdisE[5][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[5][3]);
                        } else if (Number(cfdi[TASA_IVA]) == 0) {
                          // Tasa IVA 0%
                          impTrasladadosCfdisE[5][1] = cfdi[TASA_IVA];
                          // IVA 0% 
                          impTrasladadosCfdisE[5][4] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //  + ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA]);                        
                        } else if (Number(cfdi[TASA_IVA]) < tasaIVA) {
                          // Tasa IVA %
                          impTrasladadosCfdisE[5][1] = cfdi[TASA_IVA];
                          // IVA %                    
                          impTrasladadosCfdisE[5][3] += Number(cfdi[IVA_TRASLADADO]);
                          //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                          //+ ' | Actividad Gravada Tasa IVA: ' + cfdi[TASA_IVA] +' | '+ impTrasladadosCfdisE[5][3]);
                        }
                        break;

                      default: //Tasa IVA excento
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        impTrasladadosCfdisE[5][5] += Number(cfdi[IVA_TRASLADADO]);
                    }

                    // IEPS Trasladado
                    switch (cfdi[TIPOFACTOR_IEPS]) {

                      case 'Tasa': //Tasa IEPS                      
                        // Tasa IEPS %
                        impTrasladadosCfdisE[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[5][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //  + ' | Actividad Exenta de IVA');
                        break;
                      case 'Cuota': //Cuota IEPS                      
                        // Cuota IEPS %
                        impTrasladadosCfdisE[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[5][6] += Number(cfdi[IEPS_TRASLADADO]);
                        //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                        //+ ' | Actividad Exenta de IVA');
                        break;
                      default: //Tasa IEPS excento
                        impTrasladadosCfdisE[5][2] = cfdi[TASA_IEPS];
                        impTrasladadosCfdisE[5][7] += Number(cfdi[IEPS_TRASLADADO]);
                      //console.log('(' + index + ') ' + 'Tipo CFDI ' + cfdi[TIPOCFDI] + ': ' + cfdi[TIPOCOMPROBANTE] + ' | Pago: ' + cfdi[METODOPAGO]
                      //+ ' | Actividad Exenta de IEPS');

                    }

                    // -------- FIN RECIBIDAS: IMPUESTOS TRASLADADOS PPD -------- //
                    break;
                  default:
                    console.log('Indique un método de pago')
                } // Fin switch (cfdi[METODOPAGO])
                break;
              default:
                console.log('Indique un tipo de CFDI');
            } // Fin switch (cfdi[TIPOCFDI])
            break;

          case 'T':	// Traslado
            break;

          case 'N':	// Nómina
            break;

          case 'P': // Pago
            break;

          default:
            console.log('Indique un tipo de comprobante')

        }
      } // FIN: if (ssRepoCfdis[0][DEDUCIBLE] == 'Si')

    });

    // -----------------------------  TIPO CFDI I: INGRESO ----------------------------- //
    // Total Ingresos = Ingresos PUE + Ingresos PPD

    // ------- CFDIs: SUBTOTAL ------- //
    // Emitidas: SubTotal = SubTotal PUE + SubTotal PPD
    tipoCfdisI[3][1] = tipoCfdisI[1][1] + tipoCfdisI[2][1];
    // Recibidas: SubTotal = SubTotal PUE + SubTotal PPD
    tipoCfdisI[6][1] = tipoCfdisI[4][1] + tipoCfdisI[5][1];

    // ------- CFDIs: DESCUENTO ------- //
    // Emitidas: Descuento = Descuento PUE + Descuento PPD
    tipoCfdisI[3][2] = tipoCfdisI[1][2] + tipoCfdisI[2][2];
    // Recibidas: Descuento = SubTotal PUE + Descuento PPD
    tipoCfdisI[6][2] = tipoCfdisI[4][2] + tipoCfdisI[5][2];

    // ------- CFDIs: TOTAL ------- //
    // Emitidas: Total = Total PUE + Total PPD
    tipoCfdisI[3][3] = tipoCfdisI[1][3] + tipoCfdisI[2][3];
    // Recibidas: Total = Total PUE + Total PPD
    tipoCfdisI[6][3] = tipoCfdisI[4][3] + tipoCfdisI[5][3];

    // ------- IMPUESTOS RETENIDOS ------- //

    // Total Imp Retenidos - CFDIs I - EMITIDAS
    // Total ISR Retenido = ISR Retenido PUE + ISR Retenido PPD 
    impRetenidosCfdisI[3][1] = impRetenidosCfdisI[1][1] + impRetenidosCfdisI[2][1];
    // Total IVA Retenido = IVA Retenido PUE + IVA Retenido PPD 
    impRetenidosCfdisI[3][2] = impRetenidosCfdisI[1][2] + impRetenidosCfdisI[2][2];
    // Total IEPS Retenido = IEPS Retenido PUE + IEPS Retenido PPD
    impRetenidosCfdisI[3][3] = impRetenidosCfdisI[1][3] + impRetenidosCfdisI[2][3];
    // Total Imp Retenidos = Total Imp Retenidos PUE + Total Imp Retenidos PPD
    impRetenidosCfdisI[3][4] = impRetenidosCfdisI[1][4] + impRetenidosCfdisI[2][4];

    // Total Imp Retenidos - CFDIs I - RECIBIDAS 
    // Total ISR Retenido = ISR Retenido PUE + ISR Retenido PPD 
    impRetenidosCfdisI[6][1] = impRetenidosCfdisI[4][1] + impRetenidosCfdisI[5][1];
    // Total IVA Retenido = IVA Retenido PUE + IVA Retenido PPD 
    impRetenidosCfdisI[6][2] = impRetenidosCfdisI[4][2] + impRetenidosCfdisI[5][2];
    // Total IEPS Retenido = IEPS Retenido PUE + IEPS Retenido PPD
    impRetenidosCfdisI[6][3] = impRetenidosCfdisI[4][3] + impRetenidosCfdisI[5][3];
    // Total Imp Retenidos = Total Imp Retenidos PUE + Total Imp Retenidos PPD
    impRetenidosCfdisI[6][4] = impRetenidosCfdisI[4][4] + impRetenidosCfdisI[5][4];

    // Total Imp. Retenidos = ISR Retenido + IVA Retenido + IEPS Retenido
    // totalImpRetCfdisIemitidas = impRetenidosCfdisI[3][4];
    // totalImpRetCfdisIrecibidas = impRetenidosCfdisI[6][4];

    // ------- FIN IMPUESTOS RETENIDOS ------- //

    // ------- IMPUESTOS TRASLADADOS ------- //

    // Total Imp Trasladados - CFDIs I - EMITIDAS  
    // Total IVA16 Traslado = IVA16 Traslado PUE + IVA16 Traslado PPD 
    impTrasladadosCfdisI[3][3] = impTrasladadosCfdisI[1][3] + impTrasladadosCfdisI[2][3];
    // Total IVA0 Traslado = IVA0 Traslado PUE + IVA0 Traslado PPD 
    impTrasladadosCfdisI[3][4] = impTrasladadosCfdisI[1][4] + impTrasladadosCfdisI[2][4];
    // Total IVAExcento Traslado = IVAExcento Traslado PUE + IVAExcento Traslado PPD 
    impTrasladadosCfdisI[3][5] = impTrasladadosCfdisI[1][5] + impTrasladadosCfdisI[2][5];
    // Total IEPS Traslado = IEPS Traslado PUE + IEPS Traslado PPD
    impTrasladadosCfdisI[3][6] = impTrasladadosCfdisI[1][6] + impTrasladadosCfdisI[2][6];
    // Total IEPSExcento Traslado = IEPSExcento Traslado PUE + IEPSExcento Traslado PPD
    impTrasladadosCfdisI[3][7] = impTrasladadosCfdisI[1][7] + impTrasladadosCfdisI[2][7];
    // Total Imp. Traslados = Total Imp. Traslado PUE + Total Imp. Traslado PPD
    impTrasladadosCfdisI[3][8] = impTrasladadosCfdisI[1][8] + impTrasladadosCfdisI[2][8];

    // -------------- 

    // console.log('Total IVA16 Traslado (' + impTrasladadosCfdisI[3][3] + ') = IVA16 Traslado PUE (' + impTrasladadosCfdisI[1][3] + ') + IVA16 Traslado PPD (' + impTrasladadosCfdisI[2][3] + ')');
    // console.log('Total IVA0 Traslado (' + impTrasladadosCfdisI[3][4] + ') = IVA0 Traslado PUE (' + impTrasladadosCfdisI[1][4] + ') + IVA0 Traslado PPD (' + impTrasladadosCfdisI[2][4] + ')');
    // console.log('Total IVAExcento Traslado (' + impTrasladadosCfdisI[3][5] + ') = IVAExcento Traslado PUE (' + impTrasladadosCfdisI[1][5] + ') + IVAExcento Traslado PPD (' + impTrasladadosCfdisI[2][5] + ')');
    // console.log('Total IEPS Traslado (' + impTrasladadosCfdisI[3][6] + ') = IEPS Traslado PUE (' + impTrasladadosCfdisI[1][6] + ') + IEPS Traslado PPD (' + impTrasladadosCfdisI[2][6] + ')');
    // console.log('Total IEPSExcento Traslado (' + impTrasladadosCfdisI[3][7] + ') = IEPSExcento Traslado PUE (' + impTrasladadosCfdisI[1][7] + ') + IEPSExcento Traslado PPD (' + impTrasladadosCfdisI[2][7] + ')');


    // Total Imp Trasladados - CFDIs I - RECIBIDAS 
    // Total IVA16 Traslado = IVA16 Traslado PUE + IVA16 Traslado PPD 
    impTrasladadosCfdisI[6][3] = impTrasladadosCfdisI[4][3] + impTrasladadosCfdisI[5][3];
    // Total IVA0 Traslado = IVA0 Traslado PUE + IVA0 Traslado PPD 
    impTrasladadosCfdisI[6][4] = impTrasladadosCfdisI[4][4] + impTrasladadosCfdisI[5][4];
    // Total IVAExcento Traslado = IVAExcento Traslado PUE + IVAExcento Traslado PPD 
    impTrasladadosCfdisI[6][5] = impTrasladadosCfdisI[4][5] + impTrasladadosCfdisI[5][5];
    // Total IEPS Traslado = IEPS Traslado PUE + IEPS Traslado PPD
    impTrasladadosCfdisI[6][6] = impTrasladadosCfdisI[4][6] + impTrasladadosCfdisI[5][6];
    // Total IEPSExcento Traslado = IEPSExcento Traslado PUE + IEPSExcento Traslado PPD
    impTrasladadosCfdisI[6][7] = impTrasladadosCfdisI[4][7] + impTrasladadosCfdisI[5][7];
    // Total Imp. Traslados = Total Imp. Traslado PUE + Total Imp. Traslado PPD
    impTrasladadosCfdisI[6][8] = impTrasladadosCfdisI[4][8] + impTrasladadosCfdisI[5][8];

    // --------------

    // console.log('Total IVA16 Traslado (' + impTrasladadosCfdisI[6][3] + ') = IVA16 Traslado PUE (' + impTrasladadosCfdisI[4][3] + ') + IVA16 Traslado PPD (' + impTrasladadosCfdisI[5][3] + ')');
    // console.log('Total IVA0 Traslado (' + impTrasladadosCfdisI[6][4] + ') = IVA0 Traslado PUE (' + impTrasladadosCfdisI[4][4] + ') + IVA0 Traslado PPD (' + impTrasladadosCfdisI[5][4] + ')');
    // console.log('Total IVAExcento Traslado (' + impTrasladadosCfdisI[6][5] + ') = IVAExcento Traslado PUE (' + impTrasladadosCfdisI[4][5] + ') + IVAExcento Traslado PPD (' + impTrasladadosCfdisI[5][5] + ')');
    // console.log('Total IEPS Traslado (' + impTrasladadosCfdisI[6][6] + ') = IEPS Traslado PUE (' + impTrasladadosCfdisI[4][6] + ') + IEPS Traslado PPD (' + impTrasladadosCfdisI[5][6] + ')');
    // console.log('Total IEPSExcento Traslado (' + impTrasladadosCfdisI[6][7] + ') = IEPSExcento Traslado PUE (' + impTrasladadosCfdisI[4][7] + ') + IEPSExcento Traslado PPD (' + impTrasladadosCfdisI[5][7] + ')');

    // ------- FIN IMPUESTOS TRASLADADOS ------- //

    // --------------------------- FIN TIPO CFDI I: INGRESO ----------------------------- //


    // -----------------------------  TIPO CFDI I: EGRESO ----------------------------- //
    // Total Ingresos = Ingresos PUE + Ingresos PPD

    // ------- CFDIs: SUBTOTAL ------- //
    // Emitidas: SubTotal = SubTotal PUE + SubTotal PPD
    tipoCfdisE[3][1] = tipoCfdisE[1][1] + tipoCfdisE[2][1];
    // Recibidas: SubTotal = SubTotal PUE + SubTotal PPD
    tipoCfdisE[6][1] = tipoCfdisE[4][1] + tipoCfdisE[5][1];

    // ------- CFDIs: DESCUENTO ------- //
    // Emitidas: Descuento = Descuento PUE + Descuento PPD
    tipoCfdisE[3][2] = tipoCfdisE[1][2] + tipoCfdisE[2][2];
    // Recibidas: Descuento = SubTotal PUE + Descuento PPD
    tipoCfdisE[6][2] = tipoCfdisE[4][2] + tipoCfdisE[5][2];

    // ------- CFDIs: TOTAL ------- //
    // Emitidas: Total = Total PUE + Total PPD
    tipoCfdisE[3][3] = tipoCfdisE[1][3] + tipoCfdisE[2][3];
    // Recibidas: Total = Total PUE + Total PPD
    tipoCfdisE[6][3] = tipoCfdisE[4][3] + tipoCfdisE[5][3];

    // ------- IMPUESTOS RETENIDOS ------- //

    // Total Imp Retenidos - CFDIs I - EMITIDAS
    // Total ISR Retenido = ISR Retenido PUE + ISR Retenido PPD 
    impRetenidosCfdisE[3][1] = impRetenidosCfdisE[1][1] + impRetenidosCfdisE[2][1];
    // Total IVA Retenido = IVA Retenido PUE + IVA Retenido PPD 
    impRetenidosCfdisE[3][2] = impRetenidosCfdisE[1][2] + impRetenidosCfdisE[2][2];
    // Total IEPS Retenido = IEPS Retenido PUE + IEPS Retenido PPD
    impRetenidosCfdisE[3][3] = impRetenidosCfdisE[1][3] + impRetenidosCfdisE[2][3];
    // Total Imp Retenidos = Total Imp Retenidos PUE + Total Imp Retenidos PPD
    impRetenidosCfdisE[3][4] = impRetenidosCfdisE[1][4] + impRetenidosCfdisE[2][4];

    // Total Imp Retenidos - CFDIs I - RECIBIDAS 
    // Total ISR Retenido = ISR Retenido PUE + ISR Retenido PPD 
    impRetenidosCfdisE[6][1] = impRetenidosCfdisE[4][1] + impRetenidosCfdisE[5][1];
    // Total IVA Retenido = IVA Retenido PUE + IVA Retenido PPD 
    impRetenidosCfdisE[6][2] = impRetenidosCfdisE[4][2] + impRetenidosCfdisE[5][2];
    // Total IEPS Retenido = IEPS Retenido PUE + IEPS Retenido PPD
    impRetenidosCfdisE[6][3] = impRetenidosCfdisE[4][3] + impRetenidosCfdisE[5][3];
    // Total Imp Retenidos = Total Imp Retenidos PUE + Total Imp Retenidos PPD
    impRetenidosCfdisE[6][4] = impRetenidosCfdisE[4][4] + impRetenidosCfdisE[5][4];

    // Total Imp. Retenidos = ISR Retenido + IVA Retenido + IEPS Retenido
    // totalImpRetCfdisIemitidas = impRetenidosCfdisE[3][4];
    // totalImpRetCfdisIrecibidas = impRetenidosCfdisE[6][4];

    // ------- FIN IMPUESTOS RETENIDOS ------- //

    // ------- IMPUESTOS TRASLADADOS ------- //

    // Total Imp Trasladados - CFDIs I - EMITIDAS  
    // Total IVA16 Traslado = IVA16 Traslado PUE + IVA16 Traslado PPD 
    impTrasladadosCfdisE[3][3] = impTrasladadosCfdisE[1][3] + impTrasladadosCfdisE[2][3];
    // Total IVA0 Traslado = IVA0 Traslado PUE + IVA0 Traslado PPD 
    impTrasladadosCfdisE[3][4] = impTrasladadosCfdisE[1][4] + impTrasladadosCfdisE[2][4];
    // Total IVAExcento Traslado = IVAExcento Traslado PUE + IVAExcento Traslado PPD 
    impTrasladadosCfdisE[3][5] = impTrasladadosCfdisE[1][5] + impTrasladadosCfdisE[2][5];
    // Total IEPS Traslado = IEPS Traslado PUE + IEPS Traslado PPD
    impTrasladadosCfdisE[3][6] = impTrasladadosCfdisE[1][6] + impTrasladadosCfdisE[2][6];
    // Total IEPSExcento Traslado = IEPSExcento Traslado PUE + IEPSExcento Traslado PPD
    impTrasladadosCfdisE[3][7] = impTrasladadosCfdisE[1][7] + impTrasladadosCfdisE[2][7];
    // Total Imp. Traslados = Total Imp. Traslado PUE + Total Imp. Traslado PPD
    impTrasladadosCfdisE[3][8] = impTrasladadosCfdisE[1][8] + impTrasladadosCfdisE[2][8];

    // -------------- 

    // console.log('Total IVA16 Traslado (' + impTrasladadosCfdisE[3][3] + ') = IVA16 Traslado PUE (' + impTrasladadosCfdisE[1][3] + ') + IVA16 Traslado PPD (' + impTrasladadosCfdisE[2][3] + ')');
    // console.log('Total IVA0 Traslado (' + impTrasladadosCfdisE[3][4] + ') = IVA0 Traslado PUE (' + impTrasladadosCfdisE[1][4] + ') + IVA0 Traslado PPD (' + impTrasladadosCfdisE[2][4] + ')');
    // console.log('Total IVAExcento Traslado (' + impTrasladadosCfdisE[3][5] + ') = IVAExcento Traslado PUE (' + impTrasladadosCfdisE[1][5] + ') + IVAExcento Traslado PPD (' + impTrasladadosCfdisE[2][5] + ')');
    // console.log('Total IEPS Traslado (' + impTrasladadosCfdisE[3][6] + ') = IEPS Traslado PUE (' + impTrasladadosCfdisE[1][6] + ') + IEPS Traslado PPD (' + impTrasladadosCfdisE[2][6] + ')');
    // console.log('Total IEPSExcento Traslado (' + impTrasladadosCfdisE[3][7] + ') = IEPSExcento Traslado PUE (' + impTrasladadosCfdisE[1][7] + ') + IEPSExcento Traslado PPD (' + impTrasladadosCfdisE[2][7] + ')');


    // Total Imp Trasladados - CFDIs I - RECIBIDAS 
    // Total IVA16 Traslado = IVA16 Traslado PUE + IVA16 Traslado PPD 
    impTrasladadosCfdisE[6][3] = impTrasladadosCfdisE[4][3] + impTrasladadosCfdisE[5][3];
    // Total IVA0 Traslado = IVA0 Traslado PUE + IVA0 Traslado PPD 
    impTrasladadosCfdisE[6][4] = impTrasladadosCfdisE[4][4] + impTrasladadosCfdisE[5][4];
    // Total IVAExcento Traslado = IVAExcento Traslado PUE + IVAExcento Traslado PPD 
    impTrasladadosCfdisE[6][5] = impTrasladadosCfdisE[4][5] + impTrasladadosCfdisE[5][5];
    // Total IEPS Traslado = IEPS Traslado PUE + IEPS Traslado PPD
    impTrasladadosCfdisE[6][6] = impTrasladadosCfdisE[4][6] + impTrasladadosCfdisE[5][6];
    // Total IEPSExcento Traslado = IEPSExcento Traslado PUE + IEPSExcento Traslado PPD
    impTrasladadosCfdisE[6][7] = impTrasladadosCfdisE[4][7] + impTrasladadosCfdisE[5][7];
    // Total Imp. Traslados = Total Imp. Traslado PUE + Total Imp. Traslado PPD
    impTrasladadosCfdisE[6][8] = impTrasladadosCfdisE[4][8] + impTrasladadosCfdisE[5][8];

    // --------------

    // console.log('Total IVA16 Traslado (' + impTrasladadosCfdisE[6][3] + ') = IVA16 Traslado PUE (' + impTrasladadosCfdisE[4][3] + ') + IVA16 Traslado PPD (' + impTrasladadosCfdisE[5][3] + ')');
    // console.log('Total IVA0 Traslado (' + impTrasladadosCfdisE[6][4] + ') = IVA0 Traslado PUE (' + impTrasladadosCfdisE[4][4] + ') + IVA0 Traslado PPD (' + impTrasladadosCfdisE[5][4] + ')');
    // console.log('Total IVAExcento Traslado (' + impTrasladadosCfdisE[6][5] + ') = IVAExcento Traslado PUE (' + impTrasladadosCfdisE[4][5] + ') + IVAExcento Traslado PPD (' + impTrasladadosCfdisE[5][5] + ')');
    // console.log('Total IEPS Traslado (' + impTrasladadosCfdisE[6][6] + ') = IEPS Traslado PUE (' + impTrasladadosCfdisE[4][6] + ') + IEPS Traslado PPD (' + impTrasladadosCfdisE[5][6] + ')');
    // console.log('Total IEPSExcento Traslado (' + impTrasladadosCfdisE[6][7] + ') = IEPSExcento Traslado PUE (' + impTrasladadosCfdisE[4][7] + ') + IEPSExcento Traslado PPD (' + impTrasladadosCfdisE[5][7] + ')');

    // ------- FIN IMPUESTOS TRASLADADOS ------- //

    // --------------------------- FIN TIPO CFDI I: EGRESO ----------------------------- //

  } // FIN if (datosValidos == true)


  // -----------  Cáculo de la Declaración Mensual del ISR ----- //

  // Base Gravable ISR = Total Ingresos
  baseGravableISR = tipoCfdisI[3][1]
  // Deducciones = Total Egresos autorizados
  deducciones = tipoCfdisI[6][1];
  // Base Cálculo ISR = Total Ingresos - Total Egresos
  baseCalculoISR = baseGravableISR - deducciones;

  var filaIniReporte = 3; var filaFinReporte = filaIniReporte + 6;
  ssCalculosCfdis.getRange('A1').setValue(nombMes);
  ssCalculosCfdis.getRange('A' + filaIniReporte + ':' + 'D' + filaFinReporte).setValues(tipoCfdisI);
  ssCalculosCfdis.getRange('F' + filaIniReporte + ':' + 'J' + filaFinReporte).setValues(impRetenidosCfdisI);
  ssCalculosCfdis.getRange('L' + filaIniReporte + ':' + 'T' + filaFinReporte).setValues(impTrasladadosCfdisI);

  filaIniReporte += 9; filaFinReporte = filaIniReporte + 6;
  ssCalculosCfdis.getRange('A1').setValue(nombMes);
  ssCalculosCfdis.getRange('A' + filaIniReporte + ':' + 'D' + filaFinReporte).setValues(tipoCfdisE);
  ssCalculosCfdis.getRange('F' + filaIniReporte + ':' + 'J' + filaFinReporte).setValues(impRetenidosCfdisE);
  ssCalculosCfdis.getRange('L' + filaIniReporte + ':' + 'T' + filaFinReporte).setValues(impTrasladadosCfdisE);

  Logger.log(tipoCfdisI);
  Logger.log(impRetenidosCfdisI);
  Logger.log(impTrasladadosCfdisI);

  tarifaSat.shift();
  //Logger.log(tarifaSat);
  if (baseCalculoISR >= 0.1) {

    for (fila in tarifaSat) {

      if (baseCalculoISR < tarifaSat[fila][1]) {
        //Logger.log(fila + ') ' +baseCalculoISR + '<' + tarifaSat[fila][1])
        limiteInferior = tarifaSat[fila][0];
        tasa = tarifaSat[fila][3];
        cuotaFija = tarifaSat[fila][2];
        break;
      } else if (baseCalculoISR > tarifaSat[fila][0]) {
        //Logger.log(fila + ') ' +baseCalculoISR + '>' + tarifaSat[fila][1])
        limiteInferior = tarifaSat[fila][0];
        tasa = tarifaSat[fila][3];
        cuotaFija = tarifaSat[fila][2];
      }
    }

  }

  // -----------  Cáculo de la Declaración Mensual del ISR ----- //
  baseImpuesto = baseCalculoISR - limiteInferior;
  impuestoMarginal = baseImpuesto * tasa;
  importeIsr = impuestoMarginal + cuotaFija;
  isrPagar = importeIsr - isrRetenido;
  if (isrPagar < 0) isrPagar = 0;

  // -----------  Escribiendo el Cáculo de la Declaración Mensual del IVA ---------- //


  // Actividades gravadas = Actividades a tasa 16% + Actividades a tasa 0%
  actividadesGravadas = impTrasladadosCfdisI[3][3] + impTrasladadosCfdisI[3][4];
  baseGravableIVA = actividadesGravadas;

  // Razón de Acreditamiento
  var acreditamiento = 0;
  if (tipoCfdisI[3][1] == 0) {  // Total Ingresos - CFDIs emitidas
    console.log('No se puede dividir entre cero')
    acreditamiento = 0;
  } else {
    // acreditamiento = actividadesGravadas / totalIngresos
    acreditamiento = actividadesGravadas / tipoCfdisI[3][1];
  }

  // IVA Acreditable = IVA Traslado 16% + IVA Traslado 0%
  // ivaPagar = ivaTrasladado - ivaAcreditable - ivaRetenido
  //ivaTrasladado = impTrasladadosCfdisI

  ivaPagar = ivaTrasladado.toFixed(0) - ivaAcreditable.toFixed(0) - ivaRetenido.toFixed(0);
  if (ivaPagar < 0) {
    ivaFavor = -ivaPagar;
    ivaPagar = 0;
  }

  var labelCfdisTipoI = [['Relación CFDIs Tipo I'],
  ['(=) Total Ingresos'],
  ['(=) Total Egresos']];

  var valuesCfdisTipoI = [[nombMes],
  [Math.round10(totalIEcfdis[3][1], -2)],
  [Math.round10(totalIEcfdis[3][2], -2)]];
  //console.log(valuesCfdisTipoI);

  //console.log(labelCfdisTipoI.length);
  //console.log(transpose(labelCfdisTipoI));
  var labelCalculoISR = [['Cálculo de ISR'],
  ['Base gravable ISR'],
  ['(-) Deducciones Autorizadas'],
  ['(=) Base del cálculo'],
  ['(-) Límite Inferior'],
  ['(=) Base de Impuesto'],
  ['(x) Tasa'],
  ['(=) Impuesto Marginal'],
  ['(+) Cuota Fija'],
  ['(=) Importe ISR'],
  ['(-) ISR Retenido'],
  ['(=) ISR Neto a Pagar']];

  var valuesCalculoISR = [[nombMes],
  [Math.round10(baseGravableISR, -2)],
  [Math.round10(deducciones, -2)],
  [Math.round10(baseCalculoISR, -2)],
  [Math.round10(limiteInferior, -2)],
  [Math.round10(baseImpuesto, -2)],
  [Math.round10(tasa, -4)],
  [Math.round10(impuestoMarginal, -2)],
  [Math.round10(cuotaFija, -2)],
  [Math.round10(importeIsr, -2)],
  [Math.round10(isrRetenido, -2)],
  [Math.round10(isrPagar, -2)]];
  //console.log(valuesCalculoISR);

  var labelCalculoIVA = [['Cálculo de IVA'],
  ['(=) Actividades Exentas'],
  ['(+) Actividades a tasa 16%'],
  ['(+) Actividades a tasa 0%'],
  ['(=) Actividades gravadas'],
  ['(/) Razón de Acreditamiento'],
  ['Base gravable IVA'],
  ['(x) IVA Cobrado al 16%'],
  ['(-) IVA Acreditable'],
  ['(-) IVA Retenido 10.6667%'],
  ['(=) IVA Neto a Pagar'],
  ['(=) IVA Neto a Favor']];

  var valuesCalculoIVA = [[nombMes],
  [Math.round10(impTrasladadosCfdisI[3][5], -2)], // Total IVAExcento Traslado (emitidas)
  [Math.round10(impTrasladadosCfdisI[3][3], -2)], // Total IVA16 Traslado (emitidas)
  [Math.round10(impTrasladadosCfdisI[3][4], -2)], // Total IVA0 Traslado (emitidas)
  [Math.round10(actividadesGravadas, -2)], // Actividades gravadas
  [Math.round10(acreditamiento, -2)], // Razón de Acreditamiento 
  [Math.round10(impTrasladadosCfdisI[3][3] + impTrasladadosCfdisI[3][4], -2)], // Base gravable IVA = Actividades
  ['(x) IVA Cobrado al 16%'],
  ['(-) IVA Acreditable'],
  ['(-) IVA Retenido 10.6667%'],
  [ivaPagar],
  [ivaFavor]];
  //console.log(valuesCalculoIVA);

  var labelTipoCfdis = [[nombMes],
  ['Emitidas PUE'],
  ['Emitidas PPD'],
  ['Total Emitidas'],
  ['Recibidas PUE'],
  ['Recibidas PPD'],
  ['Total Recibidas']];

  var valuesTipoCfdis = [[nombMes],
  [dataCfdisI[1][1]],
  [dataCfdisI[2][1]],
  [dataCfdisI[3][1]],
  [dataCfdisI[1][2]],
  [dataCfdisI[2][2]],
  [dataCfdisI[3][2]],
  [rfc]];

  // Escribiendo los valores en la hója del cálculo
  var labels = labelCfdisTipoI.concat([['']]).concat(labelCalculoISR).concat([['']]).concat(labelCalculoIVA).concat([['']]).concat(labelTipoCfdis).concat([['']]);
  //console.log labels)
  var values = valuesCfdisTipoI.concat([['']]).concat(valuesCalculoISR).concat([['']]).concat(valuesCalculoIVA).concat([['']]).concat(valuesTipoCfdis);
  //console.log(values);

  var fila = 9;
  // Título del cálculo mensual
  ssCalDeclaraSat.getRange("A1").setValue(periodoDeclara);
  // Escribe los encabezados
  ssCalDeclaraSat.getRange(fila, 1, labels.length, 1).setValues(labels);
  // Escribe los values del mes seleccionado
  ssCalDeclaraSat.getRange(fila, 2 + numMes, values.length, 1).setValues(values);
  // Imprime el rango de values: console.log(ssCalDeclaraSat.getRange(fila, 1, labels.length, 1).getA1Notation());

}

/**
 * Calcula la presentación de declaraciones  mensuales del SAT
 */

function declaraMensual() {

}