// Nicolas Pinzon Aparicio 1151634
function myFunction() {
var ui=SpreadsheetApp.getUi();
  ui.createMenu('Funciones');
  ui.addItem('Enviar Correo', 'envioCorreo').addToUi();
  ui.addItem('Generar Archivo', 'generaArchivo').addToUi();
  ui.addItem('Generar Archivo', 'generaPdf').addToUi();
}

function envioCorreo(){
var sheet = SpreadsheetApp.getActiveSheet();
var startRow = 2; // Primera Fila A Procesar
var numRows = 10; // Filas A Procesar
var dataRange = sheet.getRange(startRow, 1, numRows, 10);
var data = dataRange.getValues();
for (var i in data) {
var row = data[i];
var emailAddress = row[2]; // Primera columna
var message = "Estimado (a) Estudiante "+row[0]+" "+row[1]+" "+"Código " +row[3]+"\n"+
"Nos permitimos informarle su calificación definitiva en el curso Sistemas de Información Gerencial:"+"\n"+
"Calificación: "+row[4]+"\n"+
"Observaciones: " +row[5]+"\n"+
"Felicitaciones y éxitos en su vida académica y profesional."
var subject = 'Anuncio Cloud';
MailApp.sendEmail(emailAddress, subject, message);
}
}

function generaArchivo(){
 limpiarDocumento();
 var archivo = SpreadsheetApp.getActive();
 var hojas = archivo.getSheets();
 var datos = hojas[0].getDataRange().getValues();
 const ultimaFila = hojas[0].getLastRow()-1;
 var documento = DocumentApp.openById('1CKLGwjbCZSeu7X8lYLICyF60f-GeK-pmYRu96RgPebE');
 var columnas = [
     ['Nombre','Apellido','Correo','Código','calificación','Observaciones']
   ];
  for(var i = 1; i<ultimaFila; i++){
     const fila = datos[i];
    
     const nombre = fila[0];
     const apellido = fila[1];
     const correo = fila[2];
     const codigo = fila[3];
     const nota = fila[4];
     const observaciones = fila[5];
    
    var alumno = [nombre,apellido,correo,codigo,nota,observaciones]
    columnas.push(alumno);
  }
  documento.getBody().appendTable(columnas);
}
function limpiarDocumento(){
   var documento = DocumentApp.openById('1CKLGwjbCZSeu7X8lYLICyF60f-GeK-pmYRu96RgPebE');
   documento.getBody().clear();
}
function generarPdf(){
  var documento = DocumentApp.openById('1CKLGwjbCZSeu7X8lYLICyF60f-GeK-pmYRu96RgPebE');
  DriveApp.createFile(documento.getAs('application/pdf'));
}