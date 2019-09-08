function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
    // Or DocumentApp or FormApp.
  ui.createMenu('Siparişler')
      .addItem("Siparişleri Aktar", "startOrders")
      .addItem('Faturaları Düzenle', 'startInvoicing')
      .addItem('Faturaları E-Postala', 'startSendingInvoices')
      .addSeparator()
      .addSubMenu(ui.createMenu('Entegrasyon Ayarları')
        .addItem("Platform Ayarları", "startConfiguration")
        .addItem('Paraşüte Bağlan', 'showParasutSidebar')
        .addSubMenu(ui.createMenu('Ürünleri Eşle')
          .addItem("Paraşüt", "getProducts")
          .addItem("Woocommerce", "getProductVariations")
          .addItem("N11", "mapN11Products")
          .addItem("Trendyol", "mapTyProducts")
          .addItem("Hepsiburada", "mapHbProducts")))
      //.addMenu("Ürünleri Eşle", menuEntries)
      
      .addToUi();
      
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
      var parasutStatus = sheet.getRange("ParasutConfStatus").getValue();
      var wooStatus = sheet.getRange("WC_ConfStatus").getValue();
      var n11Status = sheet.getRange("N11_ConfStatus").getValue();
      var tyStatus = sheet.getRange("TY_ConfStatus").getValue();
      var hbStatus = sheet.getRange("HB_ConfStatus").getValue();
      
      var startDate = new Date();
      startDate = startDate.getFullYear() + "-" + ("0"+(startDate.getMonth()+1)).slice(-2) + "-" + ("0" + startDate.getDate()).slice(-2);
      
      if (!wooStatus) {
        sheet.getRange("WC_StartDate").setValue(startDate);
      }
      
      if (!n11Status) {
        sheet.getRange("N11_StartDate").setValue(startDate);
      }
      
      if (!tyStatus) {
        sheet.getRange("TY_StartDate").setValue(startDate);
      }
      
      if(!parasutStatus && !wooStatus && !n11Status && !tyStatus && !hbStatus) {
        //Show Welcome Screen
         var response = ui.alert("getOrders'a Hoşgeldiniz!" +
                          "\n getOrders ile farklı e-ticaret platformlarındaki tüm siparişlerinizi " +
                          "tek bir yere toplayabilir, e-faturalarınızı otomatik düzenleyebilir ve email ile yollayabilirsiniz." + 
                          "\n\n Başlamadan önce platformlara bağlanabilmek için bazı ayarları yapmanız gerekiyor." +
                          "\n\n Entegrasyon Ayarlarına Menüden Siparişler -> Entegrasyon Ayarları -> Platform Ayarları seçerek ulaşabilirsiniz.");
      }
        
}

function startOrders() {
  var htmlOutput = HtmlService
       .createHtmlOutputFromFile('orderUI')
       .setWidth(500)
       .setHeight(510);
   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Siparişleri Aktar');
}

function startInvoicing() {
  var htmlOutput = HtmlService
       .createHtmlOutputFromFile('invoiceUI')
       .setWidth(500)
       .setHeight(510);
   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Faturaları Düzenle');
}

function startSendingInvoices() {
  var htmlOutput = HtmlService
       .createHtmlOutputFromFile('invoiceMailUI')
       .setWidth(500)
       .setHeight(510);
   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'E-postaları Gönder');
}

function startConfiguration(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("Configuration")
  .setWidth(500)
  .setHeight(510);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Entegrasyon Ayarları");

}

function ParasutConfig(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("ParasutConfig")
  .setWidth(600)
  .setHeight(700);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Parasut Configuration");

}

function woocommerceConfig(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("WooConfig")
  .setWidth(500)
  .setHeight(510);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Woocommerce Configuration");

}

function N11Config(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("N11Config")
  .setWidth(500)
  .setHeight(510);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "N11 Configuration");

}

function TrendyolConfig(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("TrendyolConfig")
  .setWidth(500)
  .setHeight(510);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Trendyol Configuration");

}

function HepsiburadaConfig(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("HepsiburadaConfig")
  .setWidth(500)
  .setHeight(510);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Hepsiburada Configuration");

}

function connectParasut(){
  var  htmlOutput = HtmlService
  .createHtmlOutputFromFile("ParasutConnect")
  .setWidth(500)
  .setHeight(300);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Parasut Connection");

}

/**
 * Removes duplicate rows from order details sheet.
 */
function removeDuplicates(sheet) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    if (row[29] == "#N/A" || row[34] == "#N/A") {
      row[39] = "HATA: Ürün Bulunamadı";
    }
    var duplicate = false;
    for (var j in newData) {
    if(row[0] == newData[j][0] && row[24] == newData[j][24]){   //Compare order ID (column 0) and item no (column 24)
        duplicate = true;
        if (row.join() !== newData[j].join() && row[1] > newData[j][1]) {   //If order/item is modified, copy newer values to existing record, keep invoice id
          if (newData[j][35] != "") {
            row[35] = newData[j][35];    //Keep invoice id
            row[36] = newData[j][36];    //Keep e-invoice id
            row[37] = newData[j][37];    //Keep mail send status
            row[38] = newData[j][38];    //Keep item sort order
          }
          newData[j] = row;   // TO DO: Also consider when an order item is deleted
        }
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function validateTcNo(tcNo) {

  tcNo = String(tcNo);
  var isValid = true;
  
  if (tcNo.length != 11 || tcNo[0] == 0 || !Number(tcNo)) {
    isValid = false;
    return isValid;
  }
  if (tcNo == "11111111111" || tcNo == "99999999999") {      //Special case to handle pseudo tc numbers
    return isValid;
  }
  var charTen = (Number(tcNo[0]) + Number(tcNo[2]) + Number(tcNo[4]) + Number(tcNo[6]) + Number(tcNo[8])) * 7
                - Number(tcNo[1]) - Number(tcNo[3]) - Number(tcNo[5]) - Number(tcNo[7]);
  var charTen = charTen % 10;
  var charEleven = (Number(tcNo[0]) + Number(tcNo[1]) + Number(tcNo[2]) + Number(tcNo[3]) + Number(tcNo[4]) + 
                        Number(tcNo[5]) + Number(tcNo[6]) + Number(tcNo[7]) + Number(tcNo[8]) + charTen) % 10;
  
  if (tcNo[9] != charTen || tcNo[10] != charEleven) {
    isValid = false;
    return isValid;
  }
  return isValid;

}

function validateTaxNo(taxNo){
  
  if(taxNo < 1000000000){
    taxNo = "0" + String(taxNo);
  
  }else{
    taxNo=String(taxNo);
  }
 
  //Logger.log(taxNo);
  //Logger.log(Number(taxNo[1]));
  
  if(taxNo.length!=10){
    return false;
  }
  
   var sum=0;
   var p;
   var q;
  
  for(var i = 0 ; i < 9 ; i++){
    p=(Number(taxNo[i]) + 10 - (i + 1)) % 10;
    if (p == 9){
      q = 9;
    } else {
      q = (p * Math.pow(2, 10 - i - 1)) % 9; 
    }
    sum += q;
  }
  
  return (10 - (sum % 10)) % 10 == Number(taxNo[9]);
  /*
  Logger.log((10-(sum % 10)) % 10);
  Logger.log(Number(taxNo[9]));
  Logger.log((10 - (sum % 10)) % 10 == Number(taxNo[9]));
  */
}


/*
Gets all orders from the spreadsheet for each platform provided in the platforms parameter.
Returns an array containing the caught errors and also logs them to the stackdriver.
*/
function getAllOrders(platforms){
  var errors=[];
  var platform;
  for(var i = 0 ; i < platforms.length ; i++){
    platform=platforms[i];
    try{
      errors.push(platform + ": Başarılı")
      if(platform == "Woocommerce"){
        //qweqwe
        getOrders();
      }else if(platform == "N11"){
        getN11Orders();
      }else if(platform == "Trendyol"){
        getTyOrders();
      }else if(platform == "Hepsiburada"){
        getHbOrders();
      }
     
    }catch(error){
      errors[i]=platform + ": Hata->"+error.message;
      console.error("getAllOrders yielded an error: " + error.message);
      continue;
    }
  }
  return errors;
}


/*
Creates all the invoices for each platform provided in the platforms parameter.
Returns an array containing the caught errors and also logs them to the stackdriver.
*/
function createAllInvoices(platforms){
    
    var errors=[];
    var platform;
    for(var i = 0 ; i < platforms.length ; i++){
      platform=platforms[i];
      try{
        errors.push(platform + ": Başarılı")
        createInvoice(platforms[i]);
        
        }catch(error){
          errors[i]=platform + ": Hata->"+error.message;
          console.error("createAllInvoices yielded an error: " + error.message);
          continue;
        }
    }
    return errors;
}


/* 
Sends all invoices for the provided platforms in the platforms parameter. 
Returns an array containing the caught errors and logs them to the stackdriver.
Utilizes the sendEmail() method below
*/
function sendEinvoices(platforms) {

  //var platforms = ["Woocommerce", "N11", "Trendyol"];
  var errors = [];
  for (var p = 0; p < platforms.length; p++) {
    
    if (platforms[p] == "Woocommerce") {
      var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order Details");
    } else if (platforms[p] == "Hepsiburada") {
      var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB Order Details");
    } else if (platforms[p] == "N11") {
      var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("N11 Order Details");
    } else if (platforms[p] == "Trendyol") {
      var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TY Order Details");
    }
    
    var orderData = orderSheet.getDataRange().getValues();
    var prevInvoiceNo;
    
    var errorCount=0;
    
      for (var i = 1; i < orderData.length; i++) {
        
        // Check and send e-invoice by email for only unsent orders
        if (orderData[i][35] != "" && orderData[i][37] == "") {
          
          var customerEmail = orderData[i][16];
          //var customerEmail = ""      //For testing
          var invoiceNo = orderData[i][35];
            if (prevInvoiceNo != invoiceNo) {
              try {
                var eInvoiceNo = sendEmail(invoiceNo, customerEmail);
                Utilities.sleep(2000);
                prevInvoiceNo = invoiceNo;
                var cell = orderSheet.getRange("A1");
                cell.offset(i, 36).setValue(eInvoiceNo);
                cell.offset(i, 37).setValue("Gönderildi");
                cell.offset(i, 39).setValue("");
                console.log("E-invoice no: " + eInvoiceNo + " successfully delivered to " + customerEmail);
              } catch(error) {
                errorCount++;
                console.error(error);                     // TO DO: Pass error to UI
                prevInvoiceNo = invoiceNo;
                var cell = orderSheet.getRange("A1");
                cell.offset(i, 39).setValue(error);
              }
            } else {
                var cell = orderSheet.getRange("A1");
                cell.offset(i, 36).setValue(eInvoiceNo);
                cell.offset(i, 37).setValue("Gönderildi");
            }
        }
      }
    if(errorCount<=0){
       errors.push(platforms[p]+ ": Başarılı");
    }else{
       errors.push(platforms[p]+ ": Toplamda "+ errorCount + " hata oluştu.");
    }
   
  
  }
  return errors;
}


/* 
Sends an email to the given email adress containing the pdf of the invoice with the given id.
*/

function sendEmail(invoiceId, customerEmail){
  
  var parasutService=getParasutService();
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = confSheet.getRange("ParasutCompanyId").getValue();
  
  var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/sales_invoices/'+ invoiceId +"?include=active_e_document", {
                                   headers: {
                                   Authorization: 'Bearer ' + parasutService.getAccessToken()
  }
  });
  

  if(response.getResponseCode()==200){
    Logger.log("Invoice ID: "+ eInvoiceId + " Invoice Type: "+ eInvoiceType);
    
    var json = response.getContentText();
    var data = JSON.parse(json);
    
    if(data.data.relationships.active_e_document.data == null){
      throw new Error("E-arsiv veya E-fatura bulunamadı.");
    
    }
    
    var eInvoiceId = data.data.relationships.active_e_document.data.id;
    var eInvoiceType = data.data.relationships.active_e_document.data.type;
    
    if(eInvoiceType=="e_archives"){
      response=UrlFetchApp.fetch("https://api.parasut.com/v4/"+ companyId + "/e_archives/"+ eInvoiceId +"/pdf",{
                                 headers: {
                                 Authorization: 'Bearer ' + parasutService.getAccessToken()
    }
  });
      json=response.getContentText();
      data=JSON.parse(json);
      var pdfLink=data.data.attributes.url;
      //Logger.log("URL: "+ pdfLink);
      response=UrlFetchApp.fetch(pdfLink,{
        "method" : "get",
        
                               
  });

      var pdfBlob=response.getAs("application/pdf");
      //Logger.log(pdfBlob.getName());
      //var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
      var mailSubject = confSheet.getRange("InvMailSubject").getValue();
      var mailBody = confSheet.getRange("InvMailBody").getValue();
      var aliases=GmailApp.getAliases();
      var from;
      for(var i=0;i<aliases.length;i++){
        if(aliases[i]==confSheet.getRange("InvMailFrom").getValue()){
          from=aliases[i];
          break;
        }
      }

      GmailApp.sendEmail(customerEmail, mailSubject + " - "+pdfBlob.getName(), mailBody+ "\n\n", {attachments : [pdfBlob],
        from : from});

     
      
    } else if(eInvoiceType=="e_invoices"){
      response=UrlFetchApp.fetch("https://api.parasut.com/v4/"+ companyId + "/e_invoices/"+ eInvoiceId +"/pdf", {
                                 headers: {
                                 Authorization: 'Bearer ' + parasutService.getAccessToken()
    }
  });
      json=response.getContentText();
      data=JSON.parse(json);
      var pdfLink=data.data.attributes.url;
      Logger.log("URL: "+ pdfLink);
      response=UrlFetchApp.fetch(pdfLink,{
        "method" : "get",
        
                               
      });
      

      var pdfBlob=response.getAs("application/pdf");
      //Logger.log(pdfBlob.getName());
      var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
      var mailSubject = confSheet.getRange("InvMailSubject").getValue();
      var mailBody = confSheet.getRange("InvMailBody").getValue();
      var aliases=GmailApp.getAliases();
      var from;
      for(var i=0;i<aliases.length;i++){
        if(aliases[i]==confSheet.getRange("InvMailFrom").getValue()){
          from=aliases[i];
          break;
        }
      }

      GmailApp.sendEmail(customerEmail, mailSubject + " - "+pdfBlob.getName(), mailBody+ "\n\n", {attachments : [pdfBlob],
        from : from});

     
    }
    
    return pdfBlob.getName();
    
}


}
