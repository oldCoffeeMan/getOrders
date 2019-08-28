function getHbOrders() {

  console.time("HepsiBuradaOrders");

// TO DO: Select orders created after the last syncronization dynamically
  
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var baseUrl = confSheet.getRange("HB_BaseURL").getValue();
  var merchantId = confSheet.getRange("HB_MerchantID").getValue();
  var userName = confSheet.getRange("HB_APIUser").getValue();
  var passWord = confSheet.getRange("HB_APIPassword").getValue();
  var encoded = Utilities.base64Encode(userName + ":" + passWord);
  var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB Order Details");
  var gotErrors = 0;
  var totalOrders = 0;
  
    
  // Get & Append New Orders
  var page = 0;    //Start pagination loop
  
  do {
    
    var url = baseUrl + "orders/merchantid/" + merchantId + "?offset=" + page + "&limit=100";
    var options =

        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,
        };
        
    var result = UrlFetchApp.fetch(url, options);
    
    if (result.getResponseCode() == 200) {

      var params = JSON.parse(result.getContentText());
      var totalPages = params.pageCount;
      
      for (var j = 0; j < params.items.length; j++) {
        var container = [];
        container.push("'" + params.items[j].orderNumber);
        for (var q = 1; q < 23; q++) {
            container.push("");
          }
        container.push(params.items[j].orderDate);
        container.push(params.items[j].id);
        temp.appendRow(container);
        temp.getRange(temp.getLastRow(), 1).setNumberFormat('@STRING@');
        temp.getRange(temp.getLastRow(), 6).setNumberFormat('@STRING@');
        temp.getRange(temp.getLastRow(), 9).setNumberFormat('@STRING@');
      }

    } else {
      //Logger.log("HepsiBurada new orders fetch failed for page no: " + page + ". Response: " + result.getResponseCode());
      console.error("HepsiBurada new orders fetch failed for page no: " + page + ". Response: " + result.getResponseCode());
      gotErrors++;
    }
    
    page++;
  } while (page < totalPages);
  
  
  // Get & Append Packaged Orders
  var url = baseUrl + "packages/merchantid/" + merchantId + "?timespan=72";
  var options =

        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,
        };
        
  var result = UrlFetchApp.fetch(url, options);
  
  if (result.getResponseCode() == 200) {

    var params = JSON.parse(result.getContentText());
    
    for (var i = 0; i < params.length; i++) {
    
      for (var j = 0; j < params[i].items.length; j++) {
        var container = [];
        container.push("'" + params[i].items[j].orderNumber);
        for (var q = 1; q < 23; q++) {
            container.push("");
          }
        container.push(params[i].items[j].orderDate);
        container.push(params[i].items[j].lineItemId);
        temp.appendRow(container);
        temp.getRange(temp.getLastRow(), 1).setNumberFormat('@STRING@');
        temp.getRange(temp.getLastRow(), 6).setNumberFormat('@STRING@');
        temp.getRange(temp.getLastRow(), 9).setNumberFormat('@STRING@');
      }
      
    }
  
  } else {
      //Logger.log("HepsiBurada packaged orders fetch failed. Response: " + result.getResponseCode());
      console.error("HepsiBurada packaged orders fetch failed. Response: " + result.getResponseCode());
      gotErrors++;
  }
  
  
  // Update order status and last status update dates + add missing info // TO DO: Select orders created after the last syncronization dynamically
  
  var range = temp.getDataRange().offset(1, 0, temp.getDataRange().getNumRows() - 1).getValues();
  var newData = [];
  for (var i in range) {
    var orderNo = range[i][0];
    var duplicate = false;
      for (var j in newData) {
        if(orderNo == newData[j]){   //Compare order ID (column 0)
          duplicate = true;
        }
      }
    if (!duplicate) {
      var url = baseUrl + "orders/merchantid/" + merchantId + "/ordernumber/" + orderNo;
      var options =
      
        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,
        
        };
        
      var result = UrlFetchApp.fetch(url, options);
      
      if (result.getResponseCode() == 200) {
        var params = JSON.parse(result.getContentText());
        appendOrders(params);
        totalOrders++;
      } else {
          //Logger.log("HepsiBurada order details fetch failed for order no: " + orderNo + ". Response: " + result.getResponseCode());
          console.error("HepsiBurada order details fetch failed for order no: " + orderNo + ". Response: " + result.getResponseCode());
          gotErrors++;
      }
      
      newData.push(orderNo);
    }
  }
  
  removeDuplicates(temp);
  
  var range = temp.getDataRange().offset(1, 0, temp.getDataRange().getNumRows()-1);   // Offset a row to omit header row from range
  range.sort([{column: 2, ascending: false}, {column: 1, ascending: false}, 25]);
  
  if (gotErrors) {
  
    console.error("Hepsiburada order fetch completed with " + gotErrors + " error(s).");
    console.info("HepsiBurada - no of orders that are processed: " + totalOrders);
    console.timeEnd("HepsiBuradaOrders");
    //throw "Hepsiburada siparişleri bazı hatalarla tamamlandı. Hata sayısı: " + gotErrors;
    
  } else {
    console.info("HepsiBurada - no of orders successfully processed: " + totalOrders);
    console.timeEnd("HepsiBuradaOrders");
    return "HepsiBurada siparişleri başarıyla tamamlandı. İşlenen sipariş sayısı: " + totalOrders;
  }
    
}

function appendOrders(params) {

  // This function can be used with HepsiBurada Get List of Orders & Get List of Order Details endpoints
  // This function requires params which must include order items based on above API schemas

    var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB Order Details");
    
    var rNo = temp.getDataRange().getNumRows() + 1;                        // Start row for formulas, next row after current data
    
    var orders = params.items;
    for (var i = 0; i < orders.length; i++) {
      
      var container = [];
      var companyName, billingTaxOffice, billingTaxNo, merchantSku;
      
      container.push("'" + orders[i].orderNumber);
      container.push(orders[i].lastStatusUpdateDate);
      container.push(orders[i].status);
      container.push(orders[i].customerName);
      container.push("");                            // TO DO: parse customerName and split into First & Last name. No field for Last name in API!
      container.push("'" + orders[i].invoice.turkishIdentityNumber);
      
      billingTaxOffice = orders[i].invoice.taxOffice;
      billingTaxNo = orders[i].invoice.taxNumber;
      if (billingTaxOffice && billingTaxNo) {
        companyName = orders[i].invoice.address.name;
      } else {
        companyName = "";
      }
      
      container.push(companyName);
      container.push(billingTaxOffice);
      container.push("'" + billingTaxNo);
      
      container.push(orders[i].invoice.address.address);
      container.push(orders[i].invoice.address.town);
      container.push(orders[i].invoice.address.city);
      
      if (params.deliveryAddress == null) {
      
        container.push(""); 
        container.push("");
        container.push("");
      
      } else {
      
        container.push(params.deliveryAddress.name + " " + params.deliveryAddress.address); 
        container.push(params.deliveryAddress.town);
        container.push(params.deliveryAddress.city);
      
      }
      
      container.push(orders[i].invoice.address.phoneNumber);
      container.push(orders[i].invoice.address.email);
      container.push("");                             //No customer note field on order
      container.push("Online Ödeme");                 //Fixed payment type
      container.push(orders[i].totalPrice.amount);    //Total Order Price      
      container.push(0);                              //No refund is calculated for now
      container.push(orders[i].totalPrice.amount);    //Total Order Price
      container.push("");                             //No refund reason for now
      container.push(orders[i].orderDate);            //Order Date
      
      //Add item data
      container.push(orders[i].id);
      container.push("=VLOOKUP(\"" + orders[i].sku + "\";HB_ProductList;2;FALSE)");            //No merchant sku is supplied in end points. Read sku from HB product table
      container.push("=VLOOKUP(\"" + orders[i].sku + "\";HB_ProductList;4;FALSE)");                         //No product name is supplied in API. Read name from products table
      container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].sku + "\";HB_ProductList;3;FALSE);Products_Parasut;3;FALSE)");                   //Get Parasut id
      
      container.push(orders[i].quantity);
      var itemPrice = Number(orders[i].unitPrice.amount)/(1 + Number(orders[i].vatRate / 100));
      var itemTotal = Number(orders[i].totalPrice.amount - orders[i].vat);
      var taxTotal = Number(orders[i].vat);
      container.push(Number(Math.round(itemPrice+'e2')+'e-2'));
      container.push(Number(Math.round(itemTotal+'e2')+'e-2'));
      container.push(Number(Math.round(taxTotal+'e2')+'e-2'));
      container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].sku + "\";HB_ProductList;3;FALSE);Products_Parasut;5;FALSE)");
      
      temp.appendRow(container);
      temp.getRange(temp.getLastRow(), 1).setNumberFormat('@STRING@');
      temp.getRange(temp.getLastRow(), 6).setNumberFormat('@STRING@');
      temp.getRange(temp.getLastRow(), 9).setNumberFormat('@STRING@');

    }
    
    return orders.length;         // Return no of order items processed

}

function mapHbProducts() {

  // Check HepsiBurada products are copied properly
  var productsHb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products HB");
  if (productsHb != null) {
    
    var rowCount = productsHb.getDataRange().getNumRows();
    if (rowCount > 1 && 
       (productsHb.getRange(1, 1).getValue() == "HepsiburadaSku" || 
        productsHb.getRange(1, 3).getValue() == "HepsiburadaSku") &&
        productsHb.getRange(1, 2).getValue() == "MerchantSku") {
        
        // Set HB codes and product mapping formulas
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert("HepsiBurada ürünleri stok/faturalama ürünleriniz ile eşleştirilecek." +
                                  "\nDİKKAT, bu işlem HepsiBurada için daha önce yapılmış eşleştirmeleri siler!" + 
                                  "\n\n Devam etmek istiyor musunuz?", ui.ButtonSet.YES_NO);
        if (response == ui.Button.YES) {
          if (productsHb.getRange(1, 1).getValue() != "HepsiburadaSku" &&
              productsHb.getRange(1, 3).getValue() == "HepsiburadaSku") {
              
              productsHb.getRange(1, 3, rowCount).copyTo(productsHb.getRange(1, 1));
              productsHb.getRange(1, 3, rowCount).clearContent();
              productsHb.getRange(1, 1).setValue("HepsiburadaSku");
          }
          
          var range = [];
          for (var i = 2; i <= rowCount; i++) {
            range.push("=VLOOKUP($B" + i + ";Products_Parasut;1;FALSE)");
            //range.push("=VLOOKUP($B" + i + ";WC_ProductMap;7;FALSE)");                 //TO DO get mapping values from Products sheet          
          }
          productsHb.getRange(2, 3, rowCount-1).setValue(range);
          productsHb.getRange(1, 3).setValue("InvoiceSku");
          productsHb.getRange(1, 1, 1, 20).setFontWeight("bold");
          
          productsHb.getDataRange().removeDuplicates([1]);
          var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
          for (var n = 0; namedRanges.length; n++) {
            if (namedRanges[n].getName() == "HB_ProductList") {
              namedRanges[n].setRange(productsHb.getRange("A:T"));
              break;
            }
          }
        }
        
    } else {
      SpreadsheetApp.getUi().alert("HepsiBurada ürün listesi doğru kopyalanmamış veya tabloda değişiklikler yapılmış! \nLütfen yeniden kopyalayıp deneyiniz.");
      //Logger.log("HepsiBurada ürün tablosu doğru kopyalanmamış veya tabloda değişiklikler yapılmış! Lütfen yeniden kopyalayıp deneyiniz.");
    }
  
  } else {
    SpreadsheetApp.getUi().alert("HepsiBurada ürün tablosu bulunamadı! \nLütfen HepsiBurada'dan indirdiğiniz ürün listesini 'Products HB' adını vereceğiniz bir tabloya kopyalayınız.");
    //Logger.log("HepsiBurada ürün tablosu bulunamadı! Lütfen ürün tablosunu 'Products HB' adını vereceğiniz bir tabloya kopyalayınız.");
  }

}


function getHbPackagedOrders() {

  // Get list of packaged orders
  
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var baseUrl = confSheet.getRange("HB_BaseURL").getValue();
  var merchantId = confSheet.getRange("HB_MerchantID").getValue();
  var userName = confSheet.getRange("HB_APIUser").getValue();
  var passWord = confSheet.getRange("HB_APIPassword").getValue();
  var encoded = Utilities.base64Encode(userName + ":" + passWord);
  
  var url = baseUrl + "packages/merchantid/" + merchantId + "?timespan=72";
  
  var options =

        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,
        };
        
  var result = UrlFetchApp.fetch(url, options);
    
  if (result.getResponseCode() == 200) {

    var params = JSON.parse(result.getContentText());

  }   // Catch & throw error
    
    var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB Order Details");
    var ordersArray = [];
    var orderNoList = [];
    
    for (var p = 0; p < params.length; p++) {
    
      for (var i = 0; i < params[p].items.length; i++) {
      
        var orderNo = params[p].items[i].orderNumber;
        var isDuplicate = false;
        
        for (var j = 0; j < orderNoList.length; j++) {
        
          if (orderNo == orderNoList[j]) {
            isDuplicate = true;
            
            break;
          }
        }
        if (!isDuplicate) {
          orderNoList.push(orderNo);
          ordersArray.push(getHBOrderDetail(orderNo));
          
        }
      
      }
    
    }
    
    return ordersArray;

}


function getHBOrderDetail(orderNo) {

  orderNo = orderNo || "0302802785";
  //var userName = "coffeetropic_dev";      //Test System
  //var passWord = "Ct12345!";              //Test System
  //var encoded = Utilities.base64Encode(userName + ":" + passWord);
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var baseUrl = confSheet.getRange("HB_BaseURL").getValue();
  var merchantId = confSheet.getRange("HB_MerchantID").getValue();
  var userName = confSheet.getRange("HB_APIUser").getValue();
  var passWord = confSheet.getRange("HB_APIPassword").getValue();
  var encoded = Utilities.base64Encode(userName + ":" + passWord);
  
  //var url = "https://oms-external-sit.hepsiburada.com/orders/merchantid/e37a4aa2-3326-45c6-aaec-e67d374d0375?offset=0&limit=3";
  var url = baseUrl + "orders/merchantid/" + merchantId + "/ordernumber/" + orderNo;
  var options =

        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,

        };
  var result = UrlFetchApp.fetch(url, options);
  var params = JSON.parse(result.getContentText());
  
  return params;
  
}
