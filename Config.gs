function getInputsN11(){
  
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  var fields=[];
  
  fields.push(confSheet.getRange("N11_APIUser").getValue());
  fields.push(confSheet.getRange("N11_APIPassword").getValue());
  
  return fields;

}


function writeInputsN11(inputStrings){

var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
confSheet.getRange("N11_APIUser").setValue(inputStrings[0]);
confSheet.getRange("N11_APIPassword").setValue(inputStrings[1]);
  
  if(confSheet.getRange("N11_APIUser").getValue()==""){
    confSheet.getRange("N11_ConfStatus").setValue("False");
  }else if(confSheet.getRange("N11_APIPassword").getValue()==""){
    confSheet.getRange("N11_ConfStatus").setValue("False");
  }else{
    confSheet.getRange("N11_ConfStatus").setValue("True");
  }

}


function getInputsHb(){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  var fields=[];
  
  fields.push(confSheet.getRange("HB_MerchantID").getValue());
  fields.push(confSheet.getRange("HB_APIUser").getValue());
  fields.push(confSheet.getRange("HB_APIPassword").getValue());
  fields.push(confSheet.getRange("HB_InvoiceStatus").getValue());
  
  return fields;

}


function writeInputsHb(inputStrings){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  confSheet.getRange("HB_MerchantID").setValue(inputStrings[0]);
  confSheet.getRange("HB_APIUser").setValue(inputStrings[1]);
  confSheet.getRange("HB_APIPassword").setValue(inputStrings[2]);
  confSheet.getRange("HB_InvoiceStatus").setValue(inputStrings[3]);  
  
  if(confSheet.getRange("HB_MerchantID").getValue()==""){
    confSheet.getRange("HB_ConfStatus").setValue("False");
  }else if(confSheet.getRange("HB_APIUser").getValue()==""){
    confSheet.getRange("HB_ConfStatus").setValue("False");
  }else if(confSheet.getRange("HB_APIPassword").getValue()==""){
    confSheet.getRange("HB_ConfStatus").setValue("False");
  }else if(confSheet.getRange("HB_InvoiceStatus").getValue()==""){
    confSheet.getRange("HB_ConfStatus").setValue("False");
  }else{
    confSheet.getRange("HB_ConfStatus").setValue("True");
  }

}



function getInputsParasut(){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  var fields=[];
  
  fields.push(confSheet.getRange("ParasutCompanyId").getValue());
  fields.push(confSheet.getRange("ParasutClientId").getValue());
  fields.push(confSheet.getRange("ParasutClientSecret").getValue());
  fields.push(confSheet.getRange("InvMailSubject").getValue());
  fields.push(confSheet.getRange("InvMailBody").getValue());
  
  fields.push(getRedirectUri());
  
  fields.push(confSheet.getRange("ParasutConfStatus").getValue());
  
  return fields;

}


function writeInputsParasut(inputStrings){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  confSheet.getRange("ParasutCompanyId").setValue(inputStrings[0]);
  confSheet.getRange("ParasutClientId").setValue(inputStrings[1]);
  confSheet.getRange("ParasutClientSecret").setValue(inputStrings[2]);
  confSheet.getRange("InvMailSubject").setValue(inputStrings[3]);
  confSheet.getRange("InvMailBody").setValue(inputStrings[4]);
  
  if(confSheet.getRange("ParasutCompanyId").getValue()==""){
    confSheet.getRange("ParasutConfStatus").setValue("False");
  }else if(confSheet.getRange("ParasutClientId").getValue()==""){
    confSheet.getRange("ParasutConfStatus").setValue("False");
  }else if(confSheet.getRange("ParasutClientSecret").getValue()==""){
    confSheet.getRange("ParasutConfStatus").setValue("False");
  }else if(confSheet.getRange("InvMailSubject").getValue()==""){
    confSheet.getRange("ParasutConfStatus").setValue("False");
  }else if( confSheet.getRange("InvMailBody").getValue()==""){
    confSheet.getRange("ParasutConfStatus").setValue("False");    
  }else{
    confSheet.getRange("ParasutConfStatus").setValue("True");
  }
  
}




function getInputsTy(){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  var fields=[];
  
  fields.push(confSheet.getRange("TY_SupplierID").getValue());
  fields.push(confSheet.getRange("TY_APIUser").getValue());
  fields.push(confSheet.getRange("TY_APIPassword").getValue());
  fields.push(confSheet.getRange("TY_InvoiceStatus").getValue());
  
  return fields;
  
}



function writeInputsTy(inputStrings){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  confSheet.getRange("TY_SupplierID").setValue(inputStrings[0]);
  confSheet.getRange("TY_APIUser").setValue(inputStrings[1]);
  confSheet.getRange("TY_APIPassword").setValue(inputStrings[2]);
  confSheet.getRange("TY_InvoiceStatus").setValue(inputStrings[3]);  
  
  if(confSheet.getRange("TY_SupplierID").getValue()==""){
    confSheet.getRange("TY_ConfStatus").setValue("False");
  }else if(confSheet.getRange("TY_APIUser").getValue()==""){
    confSheet.getRange("TY_ConfStatus").setValue("False");
  }else if(confSheet.getRange("TY_APIPassword").getValue()==""){
    confSheet.getRange("TY_ConfStatus").setValue("False");
  }else if(confSheet.getRange("TY_InvoiceStatus").getValue()==""){
    confSheet.getRange("TY_ConfStatus").setValue("False");
  }else{
    confSheet.getRange("TY_ConfStatus").setValue("True");
  }
  
}



function getInputsWC(){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  var fields=[];
  
  fields.push(confSheet.getRange("WC_StoreUrl").getValue());
  fields.push(confSheet.getRange("WC_CustomerKey").getValue());
  fields.push(confSheet.getRange("WC_CustomerSecret").getValue());
  fields.push(confSheet.getRange("WC_InvoiceStatus").getValue());
  
  return fields;
  
}


function writeInputsWC(inputStrings){
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration Test");
  confSheet.getRange("WC_StoreUrl").setValue(inputStrings[0]);
  confSheet.getRange("WC_CustomerKey").setValue(inputStrings[1]);
  confSheet.getRange("WC_CustomerSecret").setValue(inputStrings[2]);
  confSheet.getRange("WC_InvoiceStatus").setValue(inputStrings[3]);
  
   if(confSheet.getRange("WC_StoreUrl").getValue()==""){
    confSheet.getRange("WC_ConfStatus").setValue("False");
  }else if(confSheet.getRange("WC_CustomerKey").getValue()==""){
    confSheet.getRange("WC_ConfStatus").setValue("False");
  }else if(confSheet.getRange("WC_CustomerSecret").getValue()==""){
    confSheet.getRange("WC_ConfStatus").setValue("False");
  }else if(confSheet.getRange("WC_InvoiceStatus").getValue()==""){
    confSheet.getRange("WC_ConfStatus").setValue("False");
  }else{
    confSheet.getRange("WC_ConfStatus").setValue("True");
  }
  
}
