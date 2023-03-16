// ----------- CÁLCULO DE DECLARACIONES ----------- //


function calcDeclaraSAT(rfc = 'GORD730303AN5', year = '2021', month = '02-feb', cveRegimen = '612', cvePeriodo = '04', reCalc = true) {

  var ssCrm = SpreadsheetApp.openById(idCrm).getSheetByName('T_PersonaFiscal').getDataRange().getValues();
  var ssDomicilio = SpreadsheetApp.openById(idCrm).getSheetByName('T_Domicilio').getDataRange().getValues();
  var ssCalDeclaraSat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('calculoDeclaraSAT');

  // Anexo 8 RMF - Tablas IVA e ISR para PFAE, RIF, Arrendamiento, Actividades Ganadera
  var anexo8Rmf = getArchivo(idFolderAnexoRMF, year + ' Anexo 8 RMF');
  var ssAnexo8RMF = SpreadsheetApp.openById(anexo8Rmf.id);

  var RFC = ssCrm[0].indexOf('RFC');
  var RSOCIAL = ssCrm[0].indexOf('Razón Social');
  var CVEREGFIS = ssCrm[0].indexOf('Cve Regimen Fiscal');
  var REGFISCAL = ssCrm[0].indexOf('Regimen Fiscal');
  var TELEFONO = ssCrm[0].indexOf('Teléfono Principal');
  var CORREO = ssCrm[0].indexOf('Correo Empresarial');
  var DOMICILIO = ssDomicilio[0].indexOf('Domicilio Fiscal');


  var lastCol = ssCalDeclaraSat.getRange('A9').getLastColumn();
  // Borra 4 columnas a partir de la columna 9: ssCalDeclaraSat.deleteColumns(9, 4)

  var razonSocial = ''; var regimenFiscal = '';
  var telefono = ''; var correo = '';

  for (var fila in ssCrm) {
    if (ssCrm[fila][RFC] == rfc) {
      razonSocial = ssCrm[fila][RSOCIAL];
      regimenFiscal = ssCrm[fila][REGFISCAL];
      telefono = ssCrm[fila][TELEFONO];
      correo = ssCrm[fila][CORREO];
      break;
    }
  }
  var fechaHora = (new Date()).toISOString().split('T');
  var fecha = fechaHora[0];
  var hora = fechaHora[1].substring(0, 8);

  console.log('RFC: ' + rfc + ' | Razón Social: ' + razonSocial + ' | Regimen Fiscal: ' + regimenFiscal);

  // Encabezado: Datos del Contribuyente RFC
  var headerCalDeclaraSat = [['RFC', rfc],
  ['Razón Social', razonSocial],
  ['Regimen Fiscal', regimenFiscal],
  ['Correo', correo],
  ['Teléfono', telefono]];
  ssCalDeclaraSat.getRange(3, 3, headerCalDeclaraSat.length, 2).setValues(headerCalDeclaraSat);

  //Fecha y hora de cálculo
  ssCalDeclaraSat.getRange("H3").setValue(fecha);
  ssCalDeclaraSat.getRange("H4").setValue(hora);

  // Pie de Página: Datos del Contador Público
  var rfcContador = ssCrm[1][RFC];
  razonSocial = ssCrm[1][RSOCIAL];
  cveRegFiscal = ssCrm[1][CVEREGFIS];
  var domicilio = ssDomicilio[1][DOMICILIO];
  telefono = ssCrm[1][TELEFONO];
  correo = ssCrm[1][CORREO];
  if (cveRegFiscal == 612) razonSocial = 'C.P. ' + razonSocial

  // Escribiendo el pie de página
  var filaPiePag = 47;
  ssCalDeclaraSat.getRange(filaPiePag, 1).setValue(rfcContador);
  ssCalDeclaraSat.getRange(filaPiePag, 2).setValue(razonSocial);
  ssCalDeclaraSat.getRange(filaPiePag, 5).setValue(domicilio);
  ssCalDeclaraSat.getRange(filaPiePag, 11).setValue('Teléfono ' + telefono);
  ssCalDeclaraSat.getRange(filaPiePag, 12).setValue('Correo Electrónico ' + correo);

  //---------------- FIN FORMATO DE DOCUMENTO --------//

  var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'], ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];

  if (month != '19-anual') {

    for (i = 0; i < listaMeses.length; i++) {
      if (listaMeses[i][0] == month) {
        var selMonth = listaMeses[i][1];
        //console.log(month)
        break;
      }
    }
  }

  var StatusCalcDeclaraSat = {
    title: 'Calcular Declaración',
    data: [selMonth],
    status: false
  };


  if (month != '19-anual') {
    // Selecciona el tipo de declaración
    switch (cvePeriodo) {
      case '04': // Periodicidad Mensual
        console.log('Declaración Mensual');
        declaracionMensual(rfc, year, month, cveRegimen);
        break;
      case '05': // Periodicidad Bimestral
        console.log('Declaración Bimestral');
        declaracionBimestral(rfc, year, month, cveRegimen);
        break;
      case '06': // Periodicidad Trimestral
        console.log('Declaración Trimestral');
        declaracionTrimestral(rfc, year, month, cveRegimen);
        break;
      default:
        console.log('Elija un periodo declaración correcta');
    }
    StatusCalcDeclaraSat.status = true;

  } else {
    // Declaración Anual   
    switch (cvePeriodo) {

      case '04':  // Periodicidad Mensual
        var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'],
        ['06-jun', 'Junio'], ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'],
        ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];

        if (reCalc == true) {
          console.log('Recálculando cada mes');
          // Borrando valores en el informe
          var filaDatos = 9;
          ssCalDeclaraSat.getRange(filaDatos, 2, filaPiePag - filaDatos - 2, 12).clearContent();
          // Recalcula todos los meses
          listaMeses.forEach(mes => {
            var month = mes[0];
            var nombre = mes[1];
            declaracionMensual(rfc, year, month, cveRegimen);
          });
        }

        // Calcula la declaración anual
        var anual = ssAnexo8RMF.getRange('Anual').getValues();
        break;

      case '05':  // Periodicidad Bimestral
        var listaMeses = [['13-ene-feb', 'Enero-Febrero'], ['14-mar-abr', 'Marzo-Abril'], ['15-may-jun', 'Mayo-Junio'],
        ['16-jul-ago', 'Julio-Agosto'], ['17-sep-oct', 'Septiembre-Octubre'], ['18-nov-dic', 'Noviembre-Diciembre']];

        if (reCalc == true) {
          console.log('Recálculando cada bimestre');
          // Borrando valores en el informe
          ssCalDeclaraSat.getRange(9, 2, filaPiePag - 2, 12).clearContent();
          listaMeses.forEach(bimestre => {
            var month = bimestre[0];
            var nombre = bimestre[1];
            //declaracionBimestral(rfc, year, month, cveRegimen);
          });
        }

        // Calcula la declaración anual
        var anual = ssAnexo8RMF.getRange('Anual').getValues();
        break;

      case '06': // Periodicidad Trimestral
        var listaTrimestre = [['20-ene-feb-mar', 'Enero-Febrero-Marzo'], ['21-abr-may-jun', 'Abril-Mayo-Junio'],
        ['22-jul-ago-sep', 'Julio-Agosto-Septiembre'], ['23-oct-nov-dic', 'Octubre-Noviembre-Diciembre']];

        if (reCalc == true) {
          console.log('Recálculando cada trimestre');
          // Borrando valores en el informe
          ssCalDeclaraSat.getRange(9, 2, filaPiePag - 2, 12).clearContent();
          // Recalcula todos los meses
          listaTrimestre.forEach(trimestre => {
            var month = trimestre[0];
            var nombre = trimestre[1];
            //declaracionTrimestral(rfc, year, month, cveRegimen)
          });
        }
        break;

      default:
        console.log('Elija un periodo declaración correcta');
    }

    var periodoDeclara = 'CÁLCULO DE LA DECLARACIÓN ANUAL DE ' + year;
    ssCalDeclaraSat.getRange("A1").setValue(periodoDeclara);
    //Logger.log(anual);
  }
  return StatusCalcDeclaraSat;
}

function declaracionMensual(rfc = 'RIAJ5811167Y0', year = '2021', month = '04-abr', cveRegimen = '612') {

  // Anexo 8 RMF - Tablas IVA e ISR para PFAE, RIF, Arrendamiento, Actividades Ganadera
  var anexo8Rmf = getArchivo(idFolderAnexoRMF, year + ' Anexo 8 RMF');
  var ssAnexo8RMF = SpreadsheetApp.openById(anexo8Rmf.id);
  var nombreRango = '';

  switch (cveRegimen) {
    case '601':
      console.log('General de Ley Personas Morales')
      //declaraRegimen601();
      break;
    case '606':
      console.log('Regimen de Arrendamiento');
      nombreRango = 'Arrendamiento_Mensual';
      //var ssArrendamiento = ssAnexo8RMF.getSheetByName('Arrendamiento');
      var arrendTarifaMes = ssAnexo8RMF.getRange(nombreRango).getValues();
      //Logger.log(arrendTarifaMes);
      break;
    case '607':
      console.log('Régimen de Enajenación o Adquisición de Bienes');
      //declaraRegimen607()
      break;
    case '610':
      console.log('Residentes en el Extranjero sin Establecimiento Permanente en México');
      break;
    case '612':
      console.log('Regimen de Personas Físicas con Actividades Empresariales y Profesionales');
      var rangosPfae = ['PFAE_01_ene', 'PFAE_02_feb', 'PFAE_03_mar', 'PFAE_04_abr', 'PFAE_05_may', 'PFAE_06_jun',
        'PFAE_07_jul', 'PFAE_08_ago', 'PFAE_09_sep', 'PFAE_10_oct', 'PFAE_11_nov', 'PFAE_12_dic'];
      nombreRango = 'PFAE_' + month.replace('-', '_',);
      var pfaeTarifaMes = ssAnexo8RMF.getRange(nombreRango).getValues();
      calculoMensual(rfc, year, month, pfaeTarifaMes);
      break;
    case '622':
      console.log('Actividades Agrícolas, Ganaderas, Silvícolas y Pesqueras');
      break;
    case '623':
      console.log('Opcional para Grupos de Sociedades');
      break;
    case '624':
      console.log('Coordinados');
      break;
    case '625':
      console.log('Régimen de las Actividades Empresariales con ingresos a través de Plataformas Tecnológicas');
      break;
    case '626':
      console.log('Régimen Simplificado de Confianza');
      break;
    default:
      console.log('Elija un Regimen Fiscal correcto');
  }
}

function declaracionBimestral(rfc = 'RIAJ5811167Y0', year = '2021', month = '16-jul-ago', cveRegimen = '621') {

  // Anexo 8 RMF - Tablas IVA e ISR para PFAE, RIF, Arrendamiento, Actividades Ganadera
  var anexo8Rmf = getArchivo(idFolderAnexoRMF, year + ' Anexo 8 RMF');
  var ssAnexo8RMF = SpreadsheetApp.openById(anexo8Rmf.id);
  var nombreRango = '';

  var listaBimestre = [['13-ene-feb', 'Enero-Febrero'], ['14-mar-abr', 'Marzo-Abril'], ['15-may-jun', 'Mayo-Junio'],
  ['16-jul-ago', 'Julio-Agosto'], ['17-sep-oct', 'Septiembre-Octubre'], ['18-nov-dic', 'Noviembre-Diciembre'], ['19-anual', 'Anual']];

  for (i = 0; i < listaBimestre.length; i++) {
    if (listaBimestre[i][0] == month) {
      var selMonth = listaBimestre[i][1];
      var periodoDeclara = 'Cálculo de la declaración de ' + selMonth + ' de ' + year;
      break;
    } else {
      var periodoDeclara = 'La periodicidad de declaración es incorrecta ';
    }
  }

  switch (cveRegimen) {
    case '621':
      console.log('Regimen de Incorporación Fiscal');
      console.log(periodoDeclara);
      var rangosRif = ['RIF_13_ene_feb', 'RIF_14_mar_abr', 'RIF_15_may_jun', 'RIF_16_jul_ago', 'RIF_17_sep_oct', 'RIF_18_nov_dic'];
      nombreRango = 'RIF_' + month.replace('-', '_',).replace('-', '_');
      var rifTarifaBimes = ssAnexo8RMF.getRange(nombreRango).getValues();
      Logger.log(rifTarifaBimes);
      break;
    default:
      console.log('Elija un Regimen Fiscal correcto');
  }
}

function declaracionTrimesal(rfc = 'RIAJ5811167Y0', year = '2021', month = '20-ene-feb-ma', cveRegimen = '606') {

  // Anexo 8 RMF - Tablas IVA e ISR para PFAE, RIF, Arrendamiento, Actividades Ganadera
  var anexo8Rmf = getArchivo(idFolderAnexoRMF, year + ' Anexo 8 RMF');
  var ssAnexo8RMF = SpreadsheetApp.openById(anexo8Rmf.id);
  var nombreRango = '';

  var listaTrimestre = [['20-ene-feb-mar', 'Enero-Febrero-Marzo'], ['21-abr-may-jun', 'Abril-Mayo-Junio'], ['22-jul-ago-sep', 'Julio-Agosto-Septiembre'], ['23-oct-nov-dic', 'Octubre-Noviembre-Diciembre'], ['19-anual', 'Anual']];

  for (i = 0; i < listaTrimestre.length; i++) {
    if (listaTrimestre[i][0] == month) {
      var selMonth = listaTrimestre[i][1];
      var periodoDeclara = 'Cálculo de la declaración de ' + selMonth + ' de ' + year;
      break;
    } else {
      var periodoDeclara = 'La periodicidad de declaración es incorrecta ';
    }
  }

  switch (cveRegimen) {
    case '606':
      console.log('Regimen de Arrendamiento');
      console.log(periodoDeclara);
      nombreRango = 'Arrendamiento_Trimestral';
      //var ssArrendamiento = ssAnexo8RMF.getSheetByName('Arrendamiento');
      var arrendTarifaTrimes = ssAnexo8RMF.getRange(nombreRango).getValues();
      Logger.log(arrendTarifaTrimes);
      break;
    default:
      console.log('Elija un Regimen Fiscal correcto');
  }
}


function presentaDeclaracionSat(rfc = 'GORD730303AN5', year = '2021', month = '02-feb', cveRegimen = '612', cvePeriodo = '04') {
  var ssCrm = SpreadsheetApp.openById(idCrm).getSheetByName('T_PersonaFiscal').getDataRange().getValues();
  var ssDomicilio = SpreadsheetApp.openById(idCrm).getSheetByName('T_Domicilio').getDataRange().getValues();
  var ssPresentaDeclaraSat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('presentaDeclaraSAT');
  var ssCalDeclaraSat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('calculoDeclaraSAT');

  var RFC = ssCrm[0].indexOf('RFC');
  var RSOCIAL = ssCrm[0].indexOf('Razón Social');
  var CVEREGFIS = ssCrm[0].indexOf('Cve Regimen Fiscal');
  var REGFISCAL = ssCrm[0].indexOf('Regimen Fiscal');
  var TELEFONO = ssCrm[0].indexOf('Teléfono Principal');
  var CORREO = ssCrm[0].indexOf('Correo Empresarial');
  var DOMICILIO = ssDomicilio[0].indexOf('Domicilio Fiscal');

  var razonSocial = ''; var regimenFiscal = '';
  var telefono = ''; var correo = '';

  for (var fila in ssCrm) {
    if (ssCrm[fila][RFC] == rfc) {
      razonSocial = ssCrm[fila][RSOCIAL];
      regimenFiscal = ssCrm[fila][REGFISCAL];
      telefono = ssCrm[fila][TELEFONO];
      correo = ssCrm[fila][CORREO];
      break;
    }
  }
  var fechaHora = (new Date()).toISOString().split('T');
  var fecha = fechaHora[0];
  var hora = fechaHora[1].substring(0, 8);

  console.log('RFC: ' + rfc + ' | Razón Social: ' + razonSocial + ' | Regimen Fiscal: ' + regimenFiscal);

  // Encabezado: Datos del Contribuyente RFC
  var headerCalDeclaraSat = [['RFC', rfc],
  ['Razón Social', razonSocial],
  ['Regimen Fiscal', regimenFiscal],
  ['Correo', correo],
  ['Teléfono', telefono]];
  ssPresentaDeclaraSat.getRange(3, 3, headerCalDeclaraSat.length, 2).setValues(headerCalDeclaraSat);

  //Fecha y hora de cálculo
  ssPresentaDeclaraSat.getRange("H3").setValue(fecha);
  ssPresentaDeclaraSat.getRange("H4").setValue(hora);

  // Pie de Página: Datos del Contador Público
  var rfcContador = ssCrm[1][RFC];
  razonSocial = ssCrm[1][RSOCIAL];
  cveRegFiscal = ssCrm[1][CVEREGFIS];
  var domicilio = ssDomicilio[1][DOMICILIO];
  telefono = ssCrm[1][TELEFONO];
  correo = ssCrm[1][CORREO];
  if (cveRegFiscal == 612) razonSocial = 'C.P. ' + razonSocial

  // Escribiendo el pie de página
  var filaPiePag = 55;
  ssPresentaDeclaraSat.getRange(filaPiePag, 1).setValue(rfcContador);
  ssPresentaDeclaraSat.getRange(filaPiePag, 2).setValue(razonSocial);
  ssPresentaDeclaraSat.getRange(filaPiePag, 5).setValue(domicilio);
  ssPresentaDeclaraSat.getRange(filaPiePag, 11).setValue('Teléfono ' + telefono);
  ssPresentaDeclaraSat.getRange(filaPiePag, 12).setValue('Correo Electrónico ' + correo);

  //---------------- FIN FORMATO DE DOCUMENTO --------//

  var listaMeses = [['01-ene', 'Enero'], ['02-feb', 'Febrero'], ['03-mar', 'Marzo'], ['04-abr', 'Abril'], ['05-may', 'Mayo'], ['06-jun', 'Junio'], ['07-jul', 'Julio'], ['08-ago', 'Agosto'], ['09-sep', 'Septiembre'], ['10-oct', 'Octubre'], ['11-nov', 'Noviembre'], ['12-dic', 'Diciembre']];
  // Cargamos los meses calculados para el RFC seleccionado
  var checkRfc = ssCalDeclaraSat.getRange(ssCalDeclaraSat.getLastRow() - 1, 2, 1, ssCalDeclaraSat.getLastColumn() - 1).getValues();
  //console.log(checkRfc);

  var indexMes = 0; 
  
  for (i = 0; i < listaMeses.length; i++) {
  indexMes = listaMeses[i][0].indexOf(month);
    if (indexMes == 0) {
      indexMes = i + 1;
      console.log(indexMes);
      break;
    }
  }
  
  var StatusPresentaDeclaraSat = {
    title: 'Calcular Declaración',
    data: [],
    status: false
  };
  var title = 'Presentación de Declaraciones';
  var subtitle = ''; var message = '';

  for (i = 0; i < indexMes; i++) {

    if (rfc != checkRfc[0][i] && year == '2021') {
      var selMonth = listaMeses[i][1];
      subtitle = 'Cierre esta venta para continuar...';
      message = 'RFC: ' + rfc + '<br/><br/>'
        + 'No se encuentran los cálculos correspondientes de ' + selMonth + ' del ' + year;
      //console.log(message);
      showModal(title, subtitle, message);
      break;
    } else {
      var selMonth = listaMeses[i][1];
      var numMes = i;
      periodoDeclara = 'PRESENTACIÓN DE LA DECLARACIÓN DE ' + selMonth.toUpperCase() + ' DE ' + year;
      console.log(periodoDeclara);
      var col = numMes + 2;
      // Limpiando la columna del mes correspondiente
      ssPresentaDeclaraSat.getRange(9, col, 35, 1).clearContent();
      ssPresentaDeclaraSat.getRange(9, col).setValue(selMonth);
      StatusPresentaDeclaraSat.data = [selMonth];
      StatusPresentaDeclaraSat.status = true;
    }
  }


  return StatusPresentaDeclaraSat;
}
// ----------- FIN CÁLCULO DE DECLARACIONES ----------- //


function col2row(column) {
  return [column.map(function (row) { return row[0]; })];
}

function row2col(row) {
  return row[0].map(function (elem) { return [elem]; });
}

function transpose(matrix) {
  return matrix[0].map((col, i) => matrix.map(row => row[i]));
}

// Conclusión
(function () {
  /**
   * Ajuste decimal de un número.
   *
   * @param {String}  tipo  El tipo de ajuste.
   * @param {Number}  valor El numero.
   * @param {Integer} exp   El exponente (el logaritmo 10 del ajuste base).
   * @returns {Number} El valor ajustado.
   */
  function decimalAdjust(type, value, exp) {
    // Si el exp no está definido o es cero...
    if (typeof exp === 'undefined' || +exp === 0) {
      return Math[type](value);
    }
    value = +value;
    exp = +exp;
    // Si el valor no es un número o el exp no es un entero...
    if (isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0)) {
      return NaN;
    }
    // Shift
    value = value.toString().split('e');
    value = Math[type](+(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp)));
    // Shift back
    value = value.toString().split('e');
    return +(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp));
  }

  // Decimal round
  if (!Math.round10) {
    Math.round10 = function (value, exp) {
      return decimalAdjust('round', value, exp);
    };
  }
  // Decimal floor
  if (!Math.floor10) {
    Math.floor10 = function (value, exp) {
      return decimalAdjust('floor', value, exp);
    };
  }
  // Decimal ceil
  if (!Math.ceil10) {
    Math.ceil10 = function (value, exp) {
      return decimalAdjust('ceil', value, exp);
    };
  }
})();

/* // Round
Math.round10(55.55, -1);   // 55.6
Math.round10(55.549, -1);  // 55.5
Math.round10(55, 1);       // 60
Math.round10(54.9, 1);     // 50
Math.round10(-55.55, -1);  // -55.5
Math.round10(-55.551, -1); // -55.6
Math.round10(-55, 1);      // -50
Math.round10(-55.1, 1);    // -60
Math.round10(1.005, -2);   // 1.01 -- compare this with Math.round(1.005*100)/100 above
// Floor
Math.floor10(55.59, -1);   // 55.5
Math.floor10(59, 1);       // 50
Math.floor10(-55.51, -1);  // -55.6
Math.floor10(-51, 1);      // -60
// Ceil
Math.ceil10(55.51, -1);    // 55.6
Math.ceil10(51, 1);        // 60
Math.ceil10(-55.59, -1);   // -55.5
Math.ceil10(-59, 1);       // -50 */
























