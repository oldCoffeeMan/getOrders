function getTyOrders() {

// TO DO: Select orders created after the last syncronization dynamically
    console.time("TrendyolOrders");
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
    var supplierId = sheet.getRange("TY_SupplierID").getValue();
    var userName = sheet.getRange("TY_APIUser").getValue();
    var passWord = sheet.getRange("TY_APIPassword").getValue();
    var encoded = Utilities.base64Encode(userName + ":" + passWord);
    var startDate = sheet.getRange("TY_StartDate").getValue();
    
    var page = 0;    //Start pagination loop
    
    do {
    
    var url = "https://api.trendyol.com/sapigw/suppliers/" + supplierId + "/orders?startDate=" + startDate.getTime()
                + "&page=" + page + "&orderByField=LastModifiedDate&orderByDirection=DESC";
  
    var options =

        {
            "method": "GET",
            "contentType": "application/json",
            "headers" : {"Authorization" : "Basic " + encoded},
            "muteHttpExceptions": true,
        };
        
    var result = UrlFetchApp.fetch(url, options);;
    
    if (result.getResponseCode() == 200) {

      var params = JSON.parse(result.getContentText());

      var totalPages = params.totalPages;
      
      var orders = params.content;

    } else {
      console.error(result);
      throw "Trendyol Erişim Hatası: " + result.getResponseCode();
    }
    
    var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TY Order Details");

    var arrayLength = orders.length;
    var rNo = temp.getDataRange().getNumRows() + 1;                        // Start row for formulas, next row after current data
  
    for (var i = 0; i < arrayLength; i++) {
      
      var container = [];
      var packageModifs = orders[i].packageHistories.length;
      if (packageModifs > 0) {
        var modifDate = orders[i].packageHistories[0].createdDate;
        var status = orders[i].packageHistories[0].status;
      }
      for (var m = 0; m < packageModifs; m++) {
        if (modifDate < orders[i].packageHistories[m].createdDate) {
          modifDate = orders[i].packageHistories[m].createdDate;
          status = orders[i].packageHistories[m].status;
        }
      }
      modifDate = new Date(modifDate);
      container.push(orders[i].orderNumber);
      container.push(modifDate);
      container.push(status);
      container.push(orders[i].invoiceAddress.firstName);
      container.push(orders[i].invoiceAddress.lastName);
      container.push(orders[i].tcIdentityNumber);
      container.push(orders[i].invoiceAddress.company);
      container.push("");    //No field for tax office in API??
      container.push(orders[i].taxNumber);
      
      container.push(orders[i].invoiceAddress.address1 + " " + orders[i].invoiceAddress.address2);
      container.push(orders[i].invoiceAddress.district);
      container.push(orders[i].invoiceAddress.city);
      container.push(orders[i].shipmentAddress.firstName + " " + orders[i].shipmentAddress.lastName + " " + orders[i].shipmentAddress.address1 + " " + orders[i].shipmentAddress.address2); 
      container.push(orders[i].shipmentAddress.district);
      container.push(orders[i].shipmentAddress.city);
      container.push("");    //No field for phone in API!
      container.push(orders[i].customerEmail);
      container.push(orders[i].customerId);    //Customer Note field is used to store Customer Id
      container.push("Online Ödeme");    //Fixed payment type
      container.push(orders[i].totalPrice);          //Total Order Price
        
      // Discount   TO DO: Implement discounts
            
      container.push(0);    //No refund is calculated for now
      container.push(orders[i].totalPrice);          //Total Order Price
      container.push("");    //No refund reason for now
      
      container.push(new Date(orders[i].orderDate));          //Order Date
      container.push(orders[i].lines[0].id);
      container.push(orders[i].lines[0].productName);
      container.push(orders[i].lines[0].merchantSku);
      container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].lines[0].merchantSku + "\";TY_ProductList;9;FALSE);Products_Parasut;3;FALSE)");
      //container.push("=VLOOKUP(\"" + orders[i].lines[0].merchantSku + "\";WC_ProductMap;8;FALSE)");
      container.push(orders[i].lines[0].quantity);
      
      var itemPrice = Number(orders[i].lines[0].price)/(1 + Number(orders[i].lines[0].vatBaseAmount / 100));
      var itemTotal = Number(orders[i].lines[0].quantity) * itemPrice;
      var taxTotal = (Number(orders[i].lines[0].price) - itemPrice) * orders[i].lines[0].quantity;
      container.push(Number(Math.round(itemPrice+'e2')+'e-2'));
      container.push(Number(Math.round(itemTotal+'e2')+'e-2'));
      container.push(Number(Math.round(taxTotal+'e2')+'e-2'));
      container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].lines[0].merchantSku + "\";TY_ProductList;9;FALSE);Products_Parasut;5;FALSE)");
      
      temp.appendRow(container);
      container = [];
      
      c = orders[i].lines.length;
        
      if (c > 1) {

        var items = "";
        var total_line_items_quantity = 0;
        for (var k = 1; k < c; k++) {
          var item, item_f, qty, meta, variation;
          
          container.push(orders[i].orderNumber);
          container.push(modifDate);
          container.push(status);
          container.push(orders[i].invoiceAddress.firstName);
          container.push(orders[i].invoiceAddress.lastName);
          
          for (var q = 1; q < 20; q++) {
            container.push("");
          }
          
          container.push(orders[i].lines[k].id);
          container.push(orders[i].lines[k].productName);
          container.push(orders[i].lines[k].merchantSku);
          container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].lines[k].merchantSku + "\";TY_ProductList;9;FALSE);Products_Parasut;3;FALSE)");
          //container.push("=VLOOKUP(\"" + orders[i].lines[k].merchantSku + "\";WC_ProductMap;8;FALSE)");
          container.push(orders[i].lines[k].quantity);
          
          //container.push(Number(orders[i].lines[k].subtotal));
      
          var itemPrice = Number(orders[i].lines[k].price)/(1 + Number(orders[i].lines[k].vatBaseAmount / 100));
          var itemTotal = Number(orders[i].lines[k].quantity) * itemPrice;
          var taxTotal = (Number(orders[i].lines[k].price) - itemPrice) * orders[i].lines[k].quantity;
          container.push(Number(Math.round(itemPrice+'e2')+'e-2'));
          container.push(Number(Math.round(itemTotal+'e2')+'e-2'));
          container.push(Number(Math.round(taxTotal+'e2')+'e-2'));
          container.push("=VLOOKUP(VLOOKUP(\"" + orders[i].lines[k].merchantSku + "\";TY_ProductList;9;FALSE);Products_Parasut;5;FALSE)");
          //container.push("=VLOOKUP(AB" + (rNo++) + ";ParasutProductIdList;7;FALSE)");
          
          temp.appendRow(container);
          container = [];
        }
        
      }
      
    }
    page++;
    } while (page < totalPages);
        
    removeDuplicates(temp);
    
    var range = temp.getDataRange().offset(1, 0, temp.getDataRange().getNumRows()-1);   // Offset a row to omit header row from range    
    range.sort([{column: 1, ascending: false}, 25]);
        
    console.info("Trendyol - no of orders successfully processed: " + params.totalElements);
    
    var updateDuration = sheet.getRange("TY_UpdateDuration").getValue();
    var updateDate = new Date((new Date()).getTime() - (updateDuration * 60 * 60000));      //Update duration must be in hours
    if (updateDate > new Date(startDate)) {
      sheet.getRange("TY_StartDate").setValue(updateDate.toISOString().slice(0,10));
    }
    
    //createInvoice();
    console.timeEnd("TrendyolOrders");
    
}

function mapTyProducts() {

  // Check HepsiBurada products are copied properly
  var productsTy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products TY");
  if (productsTy != null) {
    
    var rowCount = productsTy.getDataRange().getNumRows();
    if (rowCount > 1 && 
       (productsTy.getRange(1, 1).getValue() == "Tedarikçi Stok Kodu" || 
        productsTy.getRange(1, 9).getValue() == "Tedarikçi Stok Kodu")) {
        
        // Set TY codes and product mapping formulas
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert("Trendyol ürünleri stok/faturalama ürünleriniz ile eşleştirilecek." +
                                  "\nDİKKAT, bu işlem Trendyol için daha önce yapılmış eşleştirmeleri siler!" + 
                                  "\n\n Devam etmek istiyor musunuz?", ui.ButtonSet.YES_NO);
        if (response == ui.Button.YES) {
          if (productsTy.getRange(1, 1).getValue() != "Tedarikçi Stok Kodu" &&
              productsTy.getRange(1, 9).getValue() == "Tedarikçi Stok Kodu") {
              
              productsTy.getRange(1, 9, rowCount).copyTo(productsTy.getRange(1, 1));
              productsTy.getRange(1, 9, rowCount).clearContent();
              productsTy.getRange(1, 1).setValue("Tedarikçi Stok Kodu");
          }
          
          var range = [];
          for (var i = 2; i <= rowCount; i++) {
            range.push("=VLOOKUP($A" + i + ";Products_Parasut;1;FALSE)");
            //range.push("=VLOOKUP($A" + i + ";WC_ProductMap;7;FALSE)");                 //TO DO get mapping values from Products sheet          
          }
          productsTy.getRange(2, 9, rowCount-1).setValue(range);
          productsTy.getRange(1, 9).setValue("Faturalama Stok Kodu");
          productsTy.getRange(1, 1, 1, 20).setFontWeight("bold");
          
          productsTy.getDataRange().removeDuplicates([1]);
          var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
          for (var n = 0; namedRanges.length; n++) {
            if (namedRanges[n].getName() == "TY_ProductList") {
              namedRanges[n].setRange(productsTy.getRange("A:Q"));
              break;
            }
          }
        }
        
    } else {
      SpreadsheetApp.getUi().alert("Trendyol ürün listesi doğru kopyalanmamış veya tabloda değişiklikler yapılmış! \nLütfen yeniden kopyalayıp deneyiniz.");
      Logger.log("Trendyol ürün tablosu doğru kopyalanmamış veya tabloda değişiklikler yapılmış! Lütfen yeniden kopyalayıp deneyiniz.");
    }
  
  } else {
    SpreadsheetApp.getUi().alert("Trendyol ürün tablosu bulunamadı! \nLütfen Trendyol'dan indirdiğiniz ürün listesini 'Products TY' adını vereceğiniz bir tabloya kopyalayınız.");
    Logger.log("Trendyol ürün tablosu bulunamadı! Lütfen ürün tablosunu 'Products TY' adını vereceğiniz bir tabloya kopyalayınız.");
  }

}
