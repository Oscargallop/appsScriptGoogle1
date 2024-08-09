function obtenerUltimasFilas() {
  var sourceSpreadsheetId = '1TMRxbgx-kyYvnj6IXZA9AUSF1Ci_mVCBxnoGCeC0vZ4';
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var destinationSpreadsheetId = '17lJhD7hmKxrRST-bl3bj52ECOSwBmjsQESqBtuA9ooc';
  var destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
  var hojas = ['E1', 'E5', 'E6', 'E8', 'E9', 'E3'];
  var resultados = {};
  
  hojas.forEach(function(nombreHoja) {
    var hoja = sourceSpreadsheet.getSheetByName(nombreHoja);
    if (hoja) {
      var datos = hoja.getRange('B:K').getValues();
      var ultimasFilas = [];
      for (var i = datos.length - 1; i >= 0 && ultimasFilas.length < 9; i--) {
        if (datos[i][0]) {
          ultimasFilas.unshift(datos[i]);
        }
      }
      resultados[nombreHoja] = ultimasFilas;
    }
  });

  var hojaTemporal = destinationSpreadsheet.insertSheet('HojaTemporal');
  var sheets = destinationSpreadsheet.getSheets();
  for (var i = sheets.length - 1; i >= 0; i--) {
    var sheet = sheets[i];
    if (sheet.getName().startsWith('Resumen ')) {
      destinationSpreadsheet.deleteSheet(sheet);
    }
  }

  var fecha = new Date();
  var fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var resumenNombre = 'Resumen ' + fechaStr;
  var resumenHoja = destinationSpreadsheet.insertSheet(resumenNombre);
  
  var titulos = ["Maquina", "Referencia", "Item", "Color", "Orden de Produccion", "Cantidad", "Estado", "Estado Estampado", "Observaciones"];
  resumenHoja.getRange(1, 1, 1, titulos.length).setValues([titulos]);
  resumenHoja.getRange(1, 1, 1, titulos.length).setFontWeight("bold");
  
  var fila = 2;
  for (var nombreHoja in resultados) {
    resumenHoja.getRange(fila, 1).setValue("Hoja: " + nombreHoja);
    fila++;
    var datos = resultados[nombreHoja];
    for (var j = 0; j < datos.length; j++) {
      resumenHoja.getRange(fila, 1, 1, datos[j].length).setValues([datos[j]]);
      fila++;
    }
    fila++;
  }
  
  resumenHoja.autoResizeColumns(1, titulos.length);
  destinationSpreadsheet.deleteSheet(hojaTemporal);
  return fechaStr;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function abrirHoja() {
  return 'https://docs.google.com/spreadsheets/d/17lJhD7hmKxrRST-bl3bj52ECOSwBmjsQESqBtuA9ooc/edit?gid=342091769#gid=342091769';
  //window.location.href = "https://docs.google.com/spreadsheets/d/17lJhD7hmKxrRST-bl3bj52ECOSwBmjsQESqBtuA9ooc/edit?gid=342091769#gid=342091769";
}

function actualizarResumen() {
  var fechaStr = obtenerUltimasFilas();
  return 'Resumen actualizado: ' + fechaStr;
}
