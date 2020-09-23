function onOpen(e){
    SpreadsheetApp.getUi().createMenu('Script-WYMH')
    .addItem('Combinar correspondencia', 'combinarCorrespondencia')
   .addItem('Generar document', 'generarArchivo')
  .addItem('Generar PDF', 'generarPdf')
    .addToUi();
  }

function combinarCorrespondencia(){
let sheet = SpreadsheetApp.getActiveSheet();
let startRow = 2;
let rows = sheet.getLastRow();
let dataRange = sheet.getRange(startRow, 1, rows, rows);
let data = dataRange.getValues();
  
for (let i in data) {
  let row = data[i];
  let emailAddress = row[2]; 
  let subject = 'Combinacion Correspondencia';
  let message = "Estimado (a) Estudiante "+row[0]+" "+row[1]+" "+"Código " +row[3]+"\n"+
    "Nos permitimos informarle su calificación definitiva en el curso Computacion en la nube:"+"\n"+
      "Calificación: "+row[4]+"\n"+
        "Observaciones: " +row[5]+"\n"+
          "Felicitaciones y éxitos en su vida académica y profesional."
        
  GmailApp.sendEmail(emailAddress, subject, message);
}
}


function generaArchivo(){
  
 let sheet = SpreadsheetApp.getActiveSheet();
 let data = sheet.getDataRange().getValues();
 const maxRow = hojas[0].getLastRow()-1;
 let document = DocumentApp.openById('1L6Ds1zXK0JupeA8HpT6PNUv31ZaEbrCuCc8r3MTUeao');
 let columns = [
     ['Nombre','Apellido','Correo','Código','calificación','Observaciones']
   ];
  for(let i = 1; i<maxRow; i++){
     let fila = data[i];
    
     let nombre = fila[0];
     let apellido = fila[1];
     let correo = fila[2];
     let codigo = fila[3];
     let nota = fila[4];
     let observaciones = fila[5];
    
    columns.push([nombre,apellido,correo,codigo,nota,observaciones]);
    
  }
  document.getBody().appendTable(columns);
}

function generarPdf(){
  let document = DocumentApp.openById('1L6Ds1zXK0JupeA8HpT6PNUv31ZaEbrCuCc8r3MTUeao');
  DriveApp.createFile(document.getAs('application/pdf'));
}
 