function getOrders() {

// TO DO: Select orders created after the last syncronization dynamically
    console.time("WoocommerceOrders");
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
    var storeUrl = sheet.getRange("WC_StoreUrl").getValue();
    var clientKey = sheet.getRange("WC_CustomerKey").getValue();
    var clientSecret = sheet.getRange("WC_CustomerSecret").getValue();
    var startDate = sheet.getRange("WC_StartDate").getValue();
    
    
    var tzoffset = (new Date()).getTimezoneOffset() * 60000; //offset in milliseconds
    startDate = (new Date(startDate - tzoffset)).toISOString().slice(0, -1);   //eliminate time zone difference
    
    var page = 1;    //Start pagination loop
    do {
    
    var url = storeUrl + "/wp-json/wc/v3/orders?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret + "&after=" + startDate + "&page=" + page + "&per_page=50"; 

    var options =

        {
            "method": "GET",
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "muteHttpExceptions": true,

        };

    var result = UrlFetchApp.fetch(url, options);
    
    if (result.getResponseCode() == 200) {

      var params = JSON.parse(result.getContentText());

      var totalPages = result.getHeaders()["x-wp-totalpages"];
      
      if (result.getHeaders()["x-cache-hit"] == "HIT" || result.getHeaders()["x-cache"] == "cached") {
        
        console.warn("Woocommerce getOrders hit cache. Total no of records: " + result.getHeaders()["x-wp-total"]);
      }
    
    } else {
      console.error(result);
      throw "Woocommerce Erişim Hatası: " + result.getResponseCode();
    }

    //var doc = SpreadsheetApp.getActiveSpreadsheet();

    var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order Details");

    var consumption = {};

    var arrayLength = params.length;
    var rNo = temp.getDataRange().getNumRows() + 1;             // Start row for formulas, next row after current data
  
    for (var i = 0; i < arrayLength; i++) {
      var a, c, d, meta;
      var container = [];
      var billingTcKimlik, billingVergiDaire, billingVergiNo;
      
      container.push(params[i]["id"]);
      container.push(params[i]["date_modified"]);
      container.push(params[i]["status"]);
      container.push(params[i]["billing"]["first_name"]);
      container.push(params[i]["billing"]["last_name"]);
      
      meta = params[i]["meta_data"].length;
      
      for (var j = 0; j < meta; j++) {
        if (params[i]["meta_data"][j]["key"] == "_billing_TC_Kimlik_No") {
          
          billingTcKimlik = params[i]["meta_data"][j]["value"];
          
        } else if (params[i]["meta_data"][j]["key"] == "_billing_Vergi_Dairesi") {
          
          billingVergiDaire = params[i]["meta_data"][j]["value"];
          
        } else if (params[i]["meta_data"][j]["key"] == "_billing_Vergi_No") {
          
          billingVergiNo = params[i]["meta_data"][j]["value"];
        }
        
      }
      
      container.push(billingTcKimlik);
      container.push(params[i]["billing"]["company"]);
      container.push(billingVergiDaire);
      container.push(billingVergiNo);
      container.push(params[i]["billing"]["address_1"]+ " "+ params[i]["billing"]["address_2"]+ " "+ params[i]["billing"]["postcode"]);
      container.push(params[i]["billing"]["city"]);
      container.push("=VLOOKUP(\"" + params[i]["billing"]["state"] + "\";WC_States;2;FALSE)");
      container.push(params[i]["shipping"]["first_name"] + " "+ params[i]["shipping"]["last_name"]+" "+ params[i]["shipping"]["address_1"] +" "+ params[i]["shipping"]["address_2"]+" "+params[i]["shipping"]["postcode"]); 
      container.push(params[i]["shipping"]["city"]);
      container.push("=VLOOKUP(\"" + params[i]["shipping"]["state"] + "\";WC_States;2;FALSE)");
      container.push(params[i]["billing"]["phone"]);
      container.push(params[i]["billing"]["email"]);
      container.push(params[i]["customer_note"]);
      container.push(params[i]["payment_method_title"]);
      container.push(Number(params[i]["total"]));          //Price
        
      //a = container.push(Number(params[i]["discount_total"])); // Discount   TO DO: Implement discounts
        
      d = params[i]["refunds"].length;
      var refundItems = "";
      var refundValue = 0;
      
      for (var r = 0; r < d; r++) {
        var item, item_f, value;
        item = params[i]["refunds"][r]["reason"];
        value = params[i]["refunds"][r]["total"];
        refundValue += Number(value);
        item_f = value +" - "+ item;
        refundItems += item_f + ", / ";
      }
      
      container.push(refundValue); //Refunded value from order
      container.push(parseFloat(container[19]) + refundValue); // Total minus refund
      container.push(refundItems); //Refunded items from order
      container.push(params[i]["date_created"]);
      // container.push(params[i]["date_created_gmt"]);
      // container.push(params[i]["order_key"]);
      container.push(params[i]["line_items"][0]["id"]);
      container.push(params[i]["line_items"][0]["name"]);
      container.push(params[i]["line_items"][0]["sku"]);
      container.push("=VLOOKUP(VLOOKUP(\"" + params[i]["line_items"][0]["sku"] + "\";WC_ProductList;4;FALSE);Products_Parasut;3;FALSE)");
      //container.push("=VLOOKUP(\"" + params[i]["line_items"][0]["sku"] + "\";WC_ProductMap;8;FALSE)");
      container.push(params[i]["line_items"][0]["quantity"]);
      container.push(Number(params[i]["line_items"][0]["price"]));
      // container.push(Number(params[i]["line_items"][0]["subtotal"]));
      container.push(Number(params[i]["line_items"][0]["total"]));
      container.push(Number(params[i]["line_items"][0]["total_tax"]));
      container.push("=VLOOKUP(VLOOKUP(\"" + params[i]["line_items"][0]["sku"] + "\";WC_ProductList;4;FALSE);Products_Parasut;5;FALSE)");
      
      temp.appendRow(container);
      container = [];
      
      c = params[i]["line_items"].length;
        
      if (c > 1) {

        var items = "";
        var total_line_items_quantity = 0;
        for (var k = 1; k < c; k++) {
          var item, item_f, qty, meta, variation;
                    
          container.push(params[i]["id"]);
          container.push(params[i]["date_modified"]);
          container.push(params[i]["status"]);
          container.push(params[i]["billing"]["first_name"]);
          container.push(params[i]["billing"]["last_name"]);
          
          for (var q = 1; q < 20; q++) {
            container.push("");
          }
          container.push(params[i]["line_items"][k]["id"]);
          container.push(params[i]["line_items"][k]["name"]);
          container.push(params[i]["line_items"][k]["sku"]);
          container.push("=VLOOKUP(VLOOKUP(\"" + params[i]["line_items"][k]["sku"] + "\";WC_ProductList;4;FALSE);Products_Parasut;3;FALSE)");
          //container.push("=VLOOKUP(\"" + params[i]["line_items"][k]["sku"] + "\";WC_ProductMap;8;FALSE)");
          container.push(params[i]["line_items"][k]["quantity"]);
          container.push(Number(params[i]["line_items"][k]["price"]));
          //container.push(Number(params[i]["line_items"][k]["subtotal"]));
          container.push(Number(params[i]["line_items"][k]["total"]));
          container.push(Number(params[i]["line_items"][k]["total_tax"]));
          container.push("=VLOOKUP(VLOOKUP(\"" + params[i]["line_items"][k]["sku"] + "\";WC_ProductList;4;FALSE);Products_Parasut;5;FALSE)");
          
          temp.appendRow(container);
          container = [];
        }
        
      }
     
        //Logger.log("Order No: " + params[i]["id"] + "  Status: " + params[i]["status"] + " Create Date: " + params[i]["date_created"]);

    }
    page++;
    } while (page <= totalPages);
        
    removeDuplicates(temp);
    
    var range = temp.getDataRange().offset(1, 0, temp.getDataRange().getNumRows()-1);   // Offset a row to omit header row from range
        
    range.sort([{column: 1, ascending: false}, 25]);
        
    console.info("Woocommerce no of orders successfully processed: " + result.getHeaders()["x-wp-total"]);
    
    var updateDuration = sheet.getRange("WC_UpdateDuration").getValue();
    var updateDate = new Date((new Date()).getTime() - (updateDuration * 60 * 60000));      //Update duration must be in hours
    
    if (updateDate > new Date(startDate)) {
      sheet.getRange("WC_StartDate").setValue(updateDate.toISOString().slice(0,10));
    }
    
    console.timeEnd("WoocommerceOrders");
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
    if (row[27] == "#N/A" || row[32] == "#N/A") {
      row[37] = "HATA: Ürün Bulunamadı";
    }
    var duplicate = false;
    for (var j in newData) {
    if(row[0] == newData[j][0] && row[24] == newData[j][24]){   //Compare order ID (column 0) and item no (column 24)
        duplicate = true;
        if (row.join() !== newData[j].join() && row[1] > newData[j][1]) {   //If order/item is modified, copy newer values to existing record, keep invoice id
          if (newData[j][33] != "") {
            row[33] = newData[j][33];    //Keep invoice id
            row[34] = newData[j][34];    //Keep e-invoice id
            row[35] = newData[j][35];    //Keep mail send status
            row[36] = newData[j][36];    //Keep item sort order
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


function getProductAttributes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var storeUrl = sheet.getRange("B3").getValue();
  var clientKey = sheet.getRange("B4").getValue();
  var clientSecret = sheet.getRange("B5").getValue();
  
  // Populate Woocommerce product codes and product mapping formulas
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Woocommerce ürünleri stok/faturalama ürünleriniz ile eşleştirilecek." +
                          "\nDİKKAT, bu işlem Woocommerce için daha önce yapılmış eşleştirmeleri siler!" + 
                          "\n\n Devam etmek istiyor musunuz?", ui.ButtonSet.YES_NO);
                          
  if (response == ui.Button.YES) {
    var productList = [];
    var rNo = 2;              // Start row for formulas, next row after current data
    var page = 1;             // Start pagination loop
    do {
    
      var sUrl = storeUrl + "/wp-json/wc/v3/products?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret + "&page=" + page + "&per_page=100";
      var options =
            {
                "method": "GET",
                "Content-Type": "application/json",
                "muteHttpExceptions": true,
            };
      var result = UrlFetchApp.fetch(sUrl, options);
      
      if (result.getResponseCode() == 200) {
      
        var data = JSON.parse(result.getContentText());
        var totalPages = result.getHeaders()["x-wp-totalpages"];
      
      } else {
      
        console.error(result);
        SpreadsheetApp.getUi().alert("Woocommerce ürünlerine erişilemiyor! \nLütfen daha sonra tekrar deneyiniz.");
        throw "Woocommerce Ürün Listesi Erişim Hatası: " + result.getResponseCode();
      
      }
      
      var products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products WC");
      
      for (var x in data) {
        
        var container = [];
        if (data[x].status == "publish") {
          if (Array.isArray(data[x].variations) && data[x].variations.length) {
          
            sUrl = storeUrl + "/wp-json/wc/v3/products/" + data[x].id + "/variations?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret + "&per_page=100";
            var options =
              {
                  "method": "GET",
                  "Content-Type": "application/json",
                  "muteHttpExceptions": true,
              };
            
            result = UrlFetchApp.fetch(sUrl, options);
            
            if (result.getResponseCode() == 200) {
              
              var variationData = JSON.parse(result.getContentText());
              for (var y in variationData) {
                if (variationData[y].status == "publish") {
                  var attrData = "";
                  for (var z in variationData[y].attributes) {
                    attrData += variationData[y].attributes[z].option + ", ";
                  }
                  container.push(variationData[y].sku);
                  container.push(variationData[y].id);
                  container.push(data[x].name);
                  //container.push("=VLOOKUP($A" + rNo++ + ";WC_ProductMap;7;FALSE)");
                  container.push("=VLOOKUP($A" + rNo++ + ";Products_Parasut;1;FALSE)");
                  container.push(Number(variationData[y].price));
                  container.push(Number(variationData[y].stock_quantity));
                  container.push(variationData[y].tax_class);
                  container.push(attrData);
                  productList.push(container);
                  container = [];
                  //doc.appendRow([variationData[y].sku,
                  //              data[x].name,
                  //              variationData[y].id, 
                  //              Number(variationData[y].price),
                  //              variationData[y].tax_class,
                  //              Number(variationData[y].stock_quantity),
                  //              ""].concat(attrData));
                  Logger.log("Variation ID: " + variationData[y].id + "   Variation Options: " + attrData);
                }
              }
            }
          } else {
          
              container.push(data[x].sku);
              container.push(data[x].id);
              container.push(data[x].name);
              container.push("=VLOOKUP($A" + rNo++ + ";Products_Parasut;1;FALSE)");
              container.push(Number(data[x].price));
              container.push(Number(data[x].stock_quantity));
              container.push(data[x].tax_class);
              container.push("");
              productList.push(container);
              container = [];
              //doc.appendRow([data[x].sku, data[x].name, data[x].id, Number(data[x].price), data[x].tax_class, Number(data[x].stock_quantity)]);
              Logger.log("Item: " + x + "   Item Name: " + data[x].name);
          }
        }
      }
    page++;
    } while (page <= totalPages);
    
    if (products.getDataRange().getNumRows() > 1) {
      products.getDataRange().offset(1, 0, products.getDataRange().getNumRows()-1).clearContent();
    }
    Logger.log("No of products: " + productList.length);
    products.getRange(2, 1, productList.length, 8).setValues(productList);
    products.getDataRange().removeDuplicates([1]);
    var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
    for (var n = 0; namedRanges.length; n++) {
      if (namedRanges[n].getName() == "WC_ProductList") {
        namedRanges[n].setRange(products.getRange("A:H"));
        break;
      }
    }
  }
}


function getVariationAttributes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var storeUrl = sheet.getRange("B3").getValue();
  var clientKey = sheet.getRange("B4").getValue();
  var clientSecret = sheet.getRange("B5").getValue();
  
  var sUrl = storeUrl + "/wp-json/wc/v3/products?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret;
  var options =
        {
            "method": "GET",
            "Content-Type": "application/json",
            "muteHttpExceptions": true,
        };

    var result = UrlFetchApp.fetch(sUrl, options);
    var data = JSON.parse(result.getContentText());
    for (var x in data) {
      var j = 0;
      var rowArray = [];
      rowArray[j] = data[x].id;
      j++;
      sUrl = storeUrl + "/wp-json/wc/v3/products/" + data[x].id + "/variations?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret + "&per_page=100";
      var options =
        {
            "method": "GET",
            "Content-Type": "application/json",
            "muteHttpExceptions": true,
        };

        result = UrlFetchApp.fetch(sUrl, options);
        var variationData = JSON.parse(result.getContentText());
        for (var y in variationData) {
          rowArray[j] = variationData[y].id;
          j++;
          rowArray[j] = variationData[y].sku;
          j++;
          rowArray[j] = variationData[y].price;
          j++;
          rowArray[j] = variationData[y].stock_quantity;
          j++;
        }
      var doc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
      doc.appendRow(rowArray);
    }
    Logger.log(rowArray);
}


function getCountryStates() {
  //Get all states of a country
  //Currently only lists states (which is used as cities) in Turkey
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var storeUrl = sheet.getRange("B3").getValue();
  var clientKey = sheet.getRange("B4").getValue();
  var clientSecret = sheet.getRange("B5").getValue();
  
  var sUrl = storeUrl + "/wp-json/wc/v3/data/countries/tr?consumer_key=" + clientKey + "&consumer_secret=" + clientSecret;
  var options =
        {
            "method": "GET",
            "Content-Type": "application/json",
            "muteHttpExceptions": true,
        };

  var result = UrlFetchApp.fetch(sUrl, options);
  var data = JSON.parse(result.getContentText());
  var states = data.states;
  for (var x in states) {
    sheet.appendRow([states[x].code, states[x].name])
  }
  
}
