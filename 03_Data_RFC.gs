// ----------- OBTENCIÓN DE DATOS DE CONTRIBUYENTE ----------- //

/**
 * Obtiene los datos de las Personas Fiscales registradas en el Portafolio de Clientes
 */
function getPersonaFiscal() {
  // Contribuyente | Portafolio clientes
  var ssCrm = SpreadsheetApp.openById(idCrm);
  var sheetContactos = ssCrm.getSheetByName('T_Contactos');
  var sheetPersonaFiscal = ssCrm.getSheetByName('T_PersonaFiscal');
  var dataPersonaFiscal = sheetPersonaFiscal.getDataRange().getDisplayValues();

  const RFC = dataPersonaFiscal[0].indexOf('RFC');
  const RAZON_SOCIAL = dataPersonaFiscal[0].indexOf('Razón Social');
  const REGIMEN_FISCAL = dataPersonaFiscal[0].indexOf('Regimen Fiscal');
  const REG_IDTRIBUTARIO = dataPersonaFiscal[0].indexOf('RegIdTrib');
  const CORREO_FISCAL = dataPersonaFiscal[0].indexOf('Correo Empresarial');
  const TELEFONO_1 = dataPersonaFiscal[0].indexOf('Teléfono Principal');
  const CVE_SERVICIO = dataPersonaFiscal[0].indexOf('Cve Servicio');
  const SERVICIO = dataPersonaFiscal[0].indexOf('Servicio');
  const CVE_PERIODO = dataPersonaFiscal[0].indexOf('Cve Periodicidad');
  const PERIODO_DECLARASAT = dataPersonaFiscal[0].indexOf('Periodicidad');

  //Logger.log(RFC);
  // Elimina encabezado
  dataPersonaFiscal.shift();


  for (var fila in dataPersonaFiscal) {
    // Verificar el RFC de form.rfc

    if (dataPersonaFiscal[fila][RFC] == 'HETF930303FA9') {
      //if (dataPersonaFiscal[fila][RFC] == form.rfc) {
      var registroPersonaFiscal = [];

      var rfc = dataPersonaFiscal[fila][RFC];
      var razonSocial = dataPersonaFiscal[fila][RAZON_SOCIAL];
      var regimenFiscal = dataPersonaFiscal[fila][REGIMEN_FISCAL];
      var regIdTrib = dataPersonaFiscal[fila][REG_IDTRIBUTARIO];
      var correoFiscal = dataPersonaFiscal[fila][CORREO_FISCAL];
      var telefonoPF = dataPersonaFiscal[fila][TELEFONO_1];
      var cveServicio = dataPersonaFiscal[fila][CVE_SERVICIO];
      //var servicio = dataPersonaFiscal[fila][SERVICIO];
      var cve_periodo = dataPersonaFiscal[fila][CVE_PERIODO];
      registroPersonaFiscal.push(rfc, razonSocial, regimenFiscal, regIdTrib, correoFiscal, telefonoPF, cveServicio, cve_periodo);

      var personaFiscal = {
        title: 'DATOS FISCALES',
        headers: ['RFC', 'Razón Social', 'Regimen Fiscal', 'RegIDTributario', 'Correo Empresarial', 'Teléfono 1', 'Cve Servicio', 'Cve Periodicidad'],
        data: registroPersonaFiscal
      }
      Logger.log(personaFiscal);

      return personaFiscal;
    }
  }
  throw ('Error al obtener los datos');
}

/**
 * Obtiene los datos de los Domicilios Fiscales registradas en el Portafolio de Clientes
 */
function getDomicilioFiscal(form) {

  // Contribuyente | Portafolio clientes
  var ssCrm = SpreadsheetApp.openById(idCrm);
  var sheetPersonaFiscal = ssCrm.getSheetByName('T_PersonaFiscal');
  var sheetDomicilio = ssCrm.getSheetByName('T_Domicilio');
  var sheetDomicilioFiscal = ssCrm.getSheetByName('T_DomicilioFiscal');

  var dataPersonaFiscal = sheetPersonaFiscal.getDataRange().getDisplayValues();
  var dataDomicilio = sheetDomicilio.getDataRange().getDisplayValues();
  var dataDomicilioFiscal = sheetDomicilioFiscal.getDataRange().getDisplayValues(); //console.log(dataDomicilioFiscal)

  const PAIS = dataDomicilio[0].indexOf('País');
  const ENTIDAD = dataDomicilio[0].indexOf('Entidad');
  const MUNICIPIO = dataDomicilio[0].indexOf('Municipio');
  const CODIGO_POSTAL = dataDomicilio[0].indexOf('Código Postal');
  const COLONIA = dataDomicilio[0].indexOf('Colonia');
  const LOCALIDAD = dataDomicilio[0].indexOf('Localidad');
  const CALLE = dataDomicilio[0].indexOf('Calle');
  const NUMEXT = dataDomicilio[0].indexOf('NumExt');
  const NUMINT = dataDomicilio[0].indexOf('NumInt');
  const REFERENCIA = dataDomicilio[0].indexOf('Referencia');

  //Eliminamos los encabezados
  dataPersonaFiscal.shift();
  dataDomicilio.shift();
  //Logger.log(dataDomicilio.length)
  var domicilioFiscal = []

  dataPersonaFiscal.forEach(personaFiscal => {

    if (personaFiscal[1] == 'HETF930303FA0') {
      //if (personaFiscal[1] == form.rfc) {
      //  personaFiscal[1] == 'HETF930303FA0'
      var idPersonaFisica = personaFiscal[0]
      //Logger.log("ID Persona Física: " + idPersonaFisica)
      var registroDomicilio = [];
      dataDomicilioFiscal.forEach(domicilioFiscal => {
        if (domicilioFiscal[0] == idPersonaFisica) {
          //Logger.log('ID Domicilio Fiscal: ' + domicilioFiscal[1]);
          // Creamos un arreglo para cada domicilio

          var fila = domicilioFiscal[1] - 1;
          var vPais = dataDomicilio[fila][PAIS];
          var vEntidad = dataDomicilio[fila][ENTIDAD];
          var vMunicipio = dataDomicilio[fila][MUNICIPIO];
          var vCodigoPostal = dataDomicilio[fila][CODIGO_POSTAL];
          var vColonia = dataDomicilio[fila][COLONIA];
          var vLocalidad = dataDomicilio[fila][LOCALIDAD];
          var vNumExt = dataDomicilio[fila][NUMEXT];
          var vNumInt = dataDomicilio[fila][NUMINT];
          var vCalle = dataDomicilio[fila][CALLE];
          var vReferencia = dataDomicilio[fila][REFERENCIA];

          if (vNumExt != '') {
            vCalle += ' Num. Ext. ' + vNumExt;
            if (vNumInt != '') vCalle += ' Num. Int. ' + vNumInt;
          } else {
            if (vNumInt != '') vCalle += ' Num. Ext. SN Num. Int. ' + vNumInt;
          }

          // Recopilamos los domicilios fiscales para el RFC
          registroDomicilio.push([vPais, vEntidad, vMunicipio, vCodigoPostal, vColonia, vCalle, vLocalidad, vReferencia]);
        }

      });

      domicilioFiscal = {
        title: 'DOMICILIO FISCAL',
        headers: ['País', 'Entidad', 'Municipio', 'C.P.', 'Colonia', 'Calle ', 'Localidad', 'Referencia'],
        data: registroDomicilio
      }
      //Logger.log(domicilioFiscal.data[0]);
    }

  });
  //Logger.log(domicilioFiscal.data);
  return domicilioFiscal;
}



// ----------- FIN OBTENCIÓN DE DATOS DE CONTRIBUYENTE ----------- //

/**
 * Obtiene los datos fiscales de cada contribuyente incluyendo su domicilio fiscal principal
 */

function getPersonasFiscales(rfc) {
  // Contribuyente | Portafolio clientes
  var ssCrm = SpreadsheetApp.openById(idCrm);
  var sheetPersonaFiscal = ssCrm.getSheetByName('T_PersonaFiscal');
  var sheetDomicilio = ssCrm.getSheetByName('T_Domicilio');
  var dataPersonaFiscal = sheetPersonaFiscal.getDataRange().getValues();
  var dataDomicilio = sheetDomicilio.getDataRange().getDisplayValues();

  const RFC = dataPersonaFiscal[0].indexOf('RFC');
  const RAZON_SOCIAL = dataPersonaFiscal[0].indexOf('Razón Social');
  const CVE_REGFISCAL = dataPersonaFiscal[0].indexOf('Cve Regimen Fiscal');
  const REGIMEN_FISCAL = dataPersonaFiscal[0].indexOf('Regimen Fiscal');
  const REG_IDTRIBUTARIO = dataPersonaFiscal[0].indexOf('RegIdTrib');
  const CORREO_FISCAL = dataPersonaFiscal[0].indexOf('Correo Empresarial');
  const TELEFONO_1 = dataPersonaFiscal[0].indexOf('Teléfono Principal');
  const DOMICILIO = dataDomicilio[0].indexOf('Domicilio Fiscal');
  const USUARIO = dataPersonaFiscal[0].indexOf('Usuario');
  const CVE_SERVICIO = dataPersonaFiscal[0].indexOf('Cve Servicio');
  const REPOSITORIO = dataPersonaFiscal[0].indexOf('Repositorio');
  const CVE_PERIODICIDAD = dataPersonaFiscal[0].indexOf('Cve Periodicidad');
  // Elimina encabezado
  dataPersonaFiscal.shift();
  dataDomicilio.shift();
  //Logger.log(dataDomicilio[0][DOMICILIO]);
  var dataPersonasFiscales = [];

  //Obten los datos fiscales de todos los constribuyentes registrados
  for (var fila in dataPersonaFiscal) {

    if (dataPersonaFiscal[fila][REPOSITORIO] == 'VERDADERO') {

      var rfc_persona = dataPersonaFiscal[fila][RFC];
      var razonSocial = dataPersonaFiscal[fila][RAZON_SOCIAL];
      var cveRegFiscal = dataPersonaFiscal[fila][CVE_REGFISCAL];
      var regimenFiscal = dataPersonaFiscal[fila][REGIMEN_FISCAL];
      var regIdTrib = dataPersonaFiscal[fila][REG_IDTRIBUTARIO];
      var correoFiscal = dataPersonaFiscal[fila][CORREO_FISCAL];
      var telefono_1 = dataPersonaFiscal[fila][TELEFONO_1];
      var domicilio = dataDomicilio[fila][DOMICILIO];
      var cveServicio = dataPersonaFiscal[fila][CVE_SERVICIO];
      var cve_periodo = dataPersonaFiscal[fila][CVE_PERIODICIDAD];

      if (rfc == dataPersonaFiscal[fila][RFC]) {
        dataPersonasFiscales.push([rfc_persona, razonSocial, cveRegFiscal, regimenFiscal, regIdTrib, correoFiscal, telefono_1, domicilio, cveServicio, cve_periodo]);
        break;
      } else if (rfc == null) {
        dataPersonasFiscales.push([rfc_persona, razonSocial, cveRegFiscal, regimenFiscal, regIdTrib, correoFiscal, telefono_1, domicilio, cveServicio, cve_periodo]);
      }
    }
  }
  Logger.log(dataPersonasFiscales);
  dataPersonasFiscales.sort(sortColum);
  
  return dataPersonasFiscales;
}


/**
 * Ordena los números por la columna 3 del arreglo
 */
function sortNumbers(a, b) {
  // Ordenar la matrix
  // https://youtu.be/hPCIOohF0Fg
  return a[3] - b[3];
}

/**
 * Ordena un arreglo por su columna 1
 */
function sortColum(r1, r2) {
  a = r1[1].toString().toLowerCase();
  b = r2[1].toString().toLowerCase();
  if (a > b) {
    return 1;
  } else if (a < b) {
    return -1
  }
  return 0;
}