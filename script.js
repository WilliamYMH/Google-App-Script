function onOpen(e){
    SpreadsheetApp.getUi().createMenu('Script-WYMH')
    .addItem('Combinar correspondencia', 'combinarCorrespondencia')
    .addToUi();
  }
  
  function combinarCorrespondencia() {
    
    let tabla = getTabla();
    //Logger.log(tabla);
    if(tabla){
      for(let i=2; i<tabla.length; i++) {
        //Logger.log(tabla[i][7]);
        GmailApp.sendEmail((tabla[i][7]), 
                           "Combinacion de correspondecia",
                           "Estimado (a) estudiante: "+tabla[i][3]+" "+tabla[i][4]+" \nNos permite informarle su reporte en el curso Computacion en la nube."+
                           "\nCalificacion: "+tabla[i][5]+
                           ".\nObservaciones: "+tabla[i][6]+
                           ".\nPuesto en el grupo: "+tabla[i][1]+
                           ".\nFelicitaciones y exitos en su vida academica y profesional.");
      }
      return;
    }
  }

 function getTabla() {
    
    var sheet = SpreadsheetApp.getActiveSheet();
    if(sheet){
      Logger.log('sheeet');
      let colum = 8;
      let rows = sheet.getLastRow();
      let table = [];
      //Logger.log(rows);
      for(let i=1; i<rows; i++) {
        table[i] = new Array(colum);
        for(let j=1; j<colum; j++) {
          table[i][j] = sheet.getRange(i, j).getValue();
        }
      }
      return table;
    }
    return null;
  }
 