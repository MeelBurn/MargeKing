
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Dialog')
      .addItem('Open MargeKing', 'openDialog')
      .addToUi();
  
  
}




function newPage(page) {
  return HtmlService.createHtmlOutputFromFile(page).getContent()
}



function openDialog() {
  
  var htmlProduct = HtmlService.createHtmlOutputFromFile('mainPage').setWidth(800).setHeight(500);
  SpreadsheetApp.getUi() // 
      .showModalDialog(htmlProduct, 'MargeKing');
  
  
}
function openDialog2(){
  
   var htmlProduct = HtmlService.createHtmlOutputFromFile('Index').setWidth(700).setHeight(500);
  SpreadsheetApp.getUi() // 
      .showModalDialog(htmlProduct, 'MargeKing');
  
}

function sendMailing(txt01,topic){
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
 
  
  for(var i=0; i < data.length; i++){
    
    if(data[i][0].includes("@")){
      
      var txt02 = txt01;
      var txttop = topic;
      for(var j=0; j<sheet.getLastColumn(); j++){
        var rx = new RegExp("{{"+data[0][j]+"}}",'g');
        txttop=txttop.replace(rx, data[i][j]);
        txt02=txt02.replace(rx, data[i][j]);
        
      }            
      
      
      /*GmailApp.sendEmail(data[i][0],
                           txttop,
                           "Open cos tam",
                         {htmlBody: txt02});
      
      */
   
      Logger.log(txttop);
      Logger.log(txt02);
      
      
    }
                            
  }
  
 
}

