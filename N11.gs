function getN11Orders() {

  console.time("N11Orders");
    
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var apiKey = sheet.getRange("N11_APIUser").getValue();
  var apiSecret = sheet.getRange("N11_APIPassword").getValue();
  var variants = sheet.getRange("N11_Variants").getValue();
  var startDate = sheet.getRange("N11_StartDate").getValue();
  startDate = ("0" + startDate.getDate()).slice(-2) + "/" + ("0"+(startDate.getMonth()+1)).slice(-2) + "/" +
              startDate.getFullYear();
  
  var page = 0;    //Start pagination loop
  do {
  
    var payload = '<?xml version="1.0" encoding="UTF-8"?> ' +
      '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://www.n11.com/ws/schemas"> ' +
        '<SOAP-ENV:Body> ' +
          '<ns1:OrderListRequest> ' +
            '<auth> ' +
              '<appKey>' + apiKey + '</appKey> ' +
              '<appSecret>' + apiSecret + '</appSecret> ' +
            '</auth> ' +
            '<searchData> ' +
              '<productId></productId> ' +
              '<status></status> ' +
              '<buyerName></buyerName> ' +
              '<orderNumber></orderNumber> ' +
              '<productSellerCode></productSellerCode> ' +
              '<recipient></recipient> ' +
              '<sameDayDelivery></sameDayDelivery> ' +
              '<period> ' +
                '<startDate>' + startDate + '</startDate> ' +
                '<endDate></endDate> ' +
              '</period> ' +
              '<sortForUpdateDate>true</sortForUpdateDate> ' +
            '</searchData> ' +
            '<pagingData> ' +
              '<currentPage>' + page + '</currentPage> ' +
              '<pageSize>50</pageSize> ' +
            '</pagingData> ' +
          '</ns1:OrderListRequest> ' +
        '</SOAP-ENV:Body> ' +
      '</SOAP-ENV:Envelope> ';
      
    var url = "https://api.n11.com/ws/OrderService.wsdl";
    
    var options =
    
          {
              "method": "get",
              "contentType": "text/xml; charset=utf-8",
              "payload": payload,
  
          };
    
    var result = UrlFetchApp.fetch(url, options);
    
    if (result.getResponseCode() == 200) {
      
      var xml = result.getContentText();
      var document = XmlService.parse(xml);
      var mainNs = XmlService.getNamespace("http://schemas.xmlsoap.org/soap/envelope/");
      var orderNs = XmlService.getNamespace("http://www.n11.com/ws/schemas");
      
      var totalPages = document.getRootElement().getChild("Body", mainNs).getChild("OrderListResponse", orderNs).getChild("pagingData").getChild("pageCount").getValue();
      Logger.log("Total Pages: " + totalPages);
      
      var orders = document.getRootElement().getChild("Body", mainNs).getChild("OrderListResponse", orderNs).getChild("orderList").getChildren("order");
      Logger.log("Orders: " + orders);
    
    } else {
      console.error(result);
      throw "N11 Sipariş Listesi Erişim Hatası: " + result.getResponseCode();
    }
    
    var temp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("N11 Order Details");
    var rNo = temp.getDataRange().getNumRows() + 1;                        // Start row for formulas, next row after current data
    
    for (var j = 0; j < orders.length; j++) {
        Logger.log("Order No: " + orders[j].getChild("orderNumber").getValue() + " Order date: " + orders[j].getChild("createDate").getValue());
        var orderId = orders[j].getChild("id").getValue();
        
        payload = '<?xml version="1.0" encoding="UTF-8"?> ' +
        '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://www.n11.com/ws/schemas"> ' +
          '<SOAP-ENV:Body> ' +
            '<ns1:OrderDetailRequest> ' +
              '<auth> ' +
                '<appKey>' + apiKey + '</appKey> ' +
                '<appSecret>' + apiSecret + '</appSecret> ' +
              '</auth> ' +
              '<orderRequest> ' +
                '<id>' + orderId + '</id> ' +
              '</orderRequest> ' +
            '</ns1:OrderDetailRequest> ' +
          '</SOAP-ENV:Body> ' +
        '</SOAP-ENV:Envelope> ';
        
        var options =
          
          {
              "method": "get",
              "contentType": "text/xml; charset=utf-8",
              "payload": payload,
        
          };
        
        var orderResult = UrlFetchApp.fetch(url, options);
        
        if (result.getResponseCode() == 200) {
          
          var orderXml = orderResult.getContentText();
          var orderDocument = XmlService.parse(orderXml);
          var orderMainNs = XmlService.getNamespace("http://schemas.xmlsoap.org/soap/envelope/");
          var orderDetailNs = XmlService.getNamespace("http://www.n11.com/ws/schemas");
          var orderDetail = orderDocument.getRootElement().getChild("Body", orderMainNs).getChild("OrderDetailResponse", orderDetailNs).getChild("orderDetail");
          //Logger.log("Order Detail XML: " + orderDocument);
          //Logger.log("Order Detail: " + orderDetail);
          //Logger.log("Order Detail Child Nodes: " + orderDetail.getChildren());
          //Logger.log("Order Billing Child Nodes: " + orderDetail.getChild("billingAddress").getChildren());
          //Logger.log("Order Item Descendants: " + orderDetail.getChild("itemList").getDescendants());
          
          //var orderModifs = orderDetail.getChild("itemList").getChildren("item").length;   TO DO: check all items for the latest modification date
          
          var modifDate = orderDetail.getChild("itemList").getChildren("item")[0].getChild("updatedDate").getValue();
          var status = orderDetail.getChild("itemList").getChildren("item")[0].getChild("status").getValue();
          
          var container = [];
          container.push(orderDetail.getChild("orderNumber").getValue());
          container.push(modifDate);
          container.push(status);
          container.push(orderDetail.getChild("billingAddress").getChild("fullName").getValue());
          container.push("");           //TO DO: Split first and last name from full name
          
          if (orderDetail.getChild("invoiceType").getValue() == "1") {    //If the customer is individual customer
          
            container.push(orderDetail.getChild("billingAddress").getChild("tcId").getValue());
            container.push("");           //Empty Company Name
            container.push("");           //Empty Tax Office
            container.push("");           //Empty Tax No
            
          } else {
            
            container.push("");           //Empty citizenship no
            container.push(orderDetail.getChild("buyer").getChild("fullName").getValue());   //Company Name
            container.push(orderDetail.getChild("billingAddress").getChild("taxHouse").getValue());
            container.push(orderDetail.getChild("billingAddress").getChild("taxId").getValue());
          
          }
          
          container.push(orderDetail.getChild("billingAddress").getChild("address").getValue() + " " + orderDetail.getChild("billingAddress").getChild("postalCode").getValue());
          container.push(orderDetail.getChild("billingAddress").getChild("district").getValue());
          container.push(orderDetail.getChild("billingAddress").getChild("city").getValue());
          container.push(orderDetail.getChild("shippingAddress").getChild("fullName").getValue() + " " + orderDetail.getChild("shippingAddress").getChild("address").getValue());
          container.push(orderDetail.getChild("shippingAddress").getChild("district").getValue());
          container.push(orderDetail.getChild("shippingAddress").getChild("city").getValue());
          container.push(orderDetail.getChild("billingAddress").getChild("gsm").getValue());
          container.push(orderDetail.getChild("buyer").getChild("email").getValue());
          container.push("");           //Empty customer note
          container.push(orderDetail.getChild("paymentType").getValue());
          container.push(Number(orderDetail.getChild("billingTemplate").getChild("sellerInvoiceAmount").getValue()));
          container.push(0);            //No refund is calculated for now
          container.push(Number(orderDetail.getChild("billingTemplate").getChild("sellerInvoiceAmount").getValue()));
          container.push("");           //No refund reason
          container.push(orderDetail.getChild("createDate").getValue());
          
          var orderItems = orderDetail.getChild("itemList").getChildren("item");
          container.push(orderItems[0].getChild("id").getValue());
          container.push(orderItems[0].getChild("productName").getValue());
          container.push(orderItems[0].getChild("productSellerCode").getValue());
          container.push("=VLOOKUP(VLOOKUP(\"" + orderItems[0].getChild("productSellerCode").getValue() + "\";N11_ProductList;2;FALSE);Products_Parasut;3;FALSE)");
          container.push(orderItems[0].getChild("quantity").getValue());
          
          var itemPrice = (Number(orderItems[0].getChild("price").getValue()) - Number(orderItems[0].getChild("sellerDiscount").getValue()))/(1 + (8 / 100));       //TO DO: read VAT rate from Product Table
          var itemTotal = Number(orderItems[0].getChild("quantity").getValue()) * itemPrice;
          var taxTotal = (Number(orderItems[0].getChild("price").getValue()) - Number(orderItems[0].getChild("sellerDiscount").getValue()) - itemPrice) * orderItems[0].getChild("quantity").getValue();
          container.push(Number(Math.round(itemPrice+'e2')+'e-2'));
          container.push(Number(Math.round(itemTotal+'e2')+'e-2'));
          container.push(Number(Math.round(taxTotal+'e2')+'e-2'));
          container.push("=VLOOKUP(VLOOKUP(\"" + orderItems[0].getChild("productSellerCode").getValue() + "\";N11_ProductList;2;FALSE);Products_Parasut;5;FALSE)");
          container.push("");            //TO DO: change item sort order column placement
          container.push("");            //TO DO: change item sort order column placement
          container.push("");            //TO DO: change item sort order column placement
          container.push(0);             //Item sort order
          
          temp.appendRow(container);
          
          container = [];
          c = orderItems.length;
            
          if (c > 1) {
            
            for (var k = 1; k < c; k++) {
            
              container.push(orderDetail.getChild("orderNumber").getValue());
              container.push(modifDate);
              container.push(status);
              container.push(orderDetail.getChild("billingAddress").getChild("fullName").getValue());
              container.push("");           //TO DO: Split first and last name from full name
              
              for (var q = 1; q < 20; q++) {
                container.push("");
              }
              
              container.push(orderItems[k].getChild("id").getValue());
              container.push(orderItems[k].getChild("productName").getValue());
              container.push(orderItems[k].getChild("productSellerCode").getValue());
              container.push("=VLOOKUP(VLOOKUP(\"" + orderItems[k].getChild("productSellerCode").getValue() + "\";N11_ProductList;2;FALSE);Products_Parasut;3;FALSE)");
              container.push(orderItems[k].getChild("quantity").getValue());
              
              var itemPrice = (Number(orderItems[k].getChild("price").getValue()) - Number(orderItems[k].getChild("sellerDiscount").getValue()))/(1 + (8 / 100));       //TO DO: read VAT rate from Product Table
              var itemTotal = Number(orderItems[k].getChild("quantity").getValue()) * itemPrice;
              var taxTotal = (Number(orderItems[k].getChild("price").getValue()) - Number(orderItems[k].getChild("sellerDiscount").getValue()) - itemPrice) * orderItems[k].getChild("quantity").getValue();
              container.push(Number(Math.round(itemPrice+'e2')+'e-2'));
              container.push(Number(Math.round(itemTotal+'e2')+'e-2'));
              container.push(Number(Math.round(taxTotal+'e2')+'e-2'));
              container.push("=VLOOKUP(VLOOKUP(\"" + orderItems[k].getChild("productSellerCode").getValue() + "\";N11_ProductList;2;FALSE);Products_Parasut;5;FALSE)");
              container.push("");            //TO DO: change item sort order column placement
              container.push("");            //TO DO: change item sort order column placement
              container.push("");            //TO DO: change item sort order column placement
              container.push(k);             //Item sort order
              
              temp.appendRow(container);
              container = [];
          
            }
          
          }
        } else {
        
          console.error("Error occured while getting details of N11 order no: " + orders[j].getChild("orderNumber").getValue());
        }
        
    }
      
    page++;
  } while (page < totalPages);
    
  removeDuplicates(temp);
    
  var range = temp.getDataRange().offset(1, 0, temp.getDataRange().getNumRows()-1);   // Offset a row to omit header row from range
  range.sort([{column: 2, ascending: false}, 1, 37]);
  
  console.info("N11 - no of orders successfully processed: " + orders.length);
    
  var updateDuration = sheet.getRange("N11_UpdateDuration").getValue();
  var updateDate = new Date((new Date()).getTime() - (updateDuration * 60 * 60000));      //Update duration must be in hours
  if (updateDate > new Date(startDate)) {
    sheet.getRange("N11_StartDate").setValue(updateDate.toISOString().slice(0,10));
  }
    
  //createInvoice();
  console.timeEnd("N11Orders");

}

function mapN11Products() {

  var productsN11 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products N11");
  if (productsN11 != null) {
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
    var apiKey = sheet.getRange("N11_APIUser").getValue();
    var apiSecret = sheet.getRange("N11_APIPassword").getValue();
    var variants = sheet.getRange("N11_Variants").getValue();
    
    // Populate N11 codes and product mapping formulas
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("N11 ürünleri stok/faturalama ürünleriniz ile eşleştirilecek." +
                            "\nDİKKAT, bu işlem N11 için daha önce yapılmış eşleştirmeleri siler!" + 
                            "\n\n Devam etmek istiyor musunuz?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      
      var rNo = 2;              // Start row for formulas, next row after current data
      var page = 0;    //Start pagination loop
      do {
      
        var payload = '<?xml version="1.0" encoding="UTF-8"?> ' +
          '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns1="http://www.n11.com/ws/schemas"> ' +
            '<SOAP-ENV:Body> ' +
              '<ns1:GetProductListRequest> ' +
                '<auth> ' +
                  '<appKey>' + apiKey + '</appKey> ' +
                  '<appSecret>' + apiSecret + '</appSecret> ' +
                '</auth> ' +
                '<pagingData> ' +
                  '<currentPage>' + page + '</currentPage> ' +
                  '<pageSize>100</pageSize> ' +
                '</pagingData> ' +
              '</ns1:GetProductListRequest> ' +
            '</SOAP-ENV:Body> ' +
          '</SOAP-ENV:Envelope> ';
          
        var url = "https://api.n11.com/ws/ProductService.wsdl";
        
        var options =
        
              {
                  "method": "get",
                  "contentType": "text/xml; charset=utf-8",
                  "payload": payload,
      
              };
        
        var result = UrlFetchApp.fetch(url, options);
        
        if (result.getResponseCode() == 200) {
          
          var xml = result.getContentText();
          var document = XmlService.parse(xml);
          var mainNs = XmlService.getNamespace("http://schemas.xmlsoap.org/soap/envelope/");
          var productNs = XmlService.getNamespace("http://www.n11.com/ws/schemas");
          
          var totalPages = document.getRootElement().getChild("Body", mainNs).getChild("GetProductListResponse", productNs).getChild("pagingData").getChild("pageCount").getValue();
          Logger.log("Total Pages: " + totalPages);
          
          var products = document.getRootElement().getChild("Body", mainNs).getChild("GetProductListResponse", productNs).getChild("products").getChildren("product");
          
        } else {
          console.error(result);
          SpreadsheetApp.getUi().alert("N11 ürünlerine erişilemiyor! \nLütfen daha sonra tekrar deneyiniz.");
          throw "N11 Ürün Listesi Erişim Hatası: " + result.getResponseCode();
        }
        
        var productList = [];
        for (var i = 0; i < products.length; i++) {
          Logger.log("Product Id: " + products[i].getChild("id").getValue() + " Seller code: " + products[i].getChild("productSellerCode").getValue());
          
          var container = [];
          var productSku = products[i].getChild("productSellerCode").getValue();
          var productId = products[i].getChild("id").getValue();
          var title = products[i].getChild("title").getValue();
          var price = products[i].getChild("price").getValue();
          var sale = products[i].getChild("saleStatus").getValue();
          var approval = products[i].getChild("approvalStatus").getValue();
          
          var stockItems = products[i].getChild("stockItems").getChildren("stockItem");
          Logger.log("No of stock items: " + stockItems.length);
          for (var j = 0; j < stockItems.length; j++) {
            
            var stockCode = stockItems[j].getChild("sellerStockCode") == null ? productSku : stockItems[j].getChild("sellerStockCode").getValue();
            var stockId = stockItems[j].getChild("id").getValue();
            var optionPrice = stockItems[j].getChild("optionPrice").getValue();
            var attributes = stockItems[j].getChild("attributes").getChildren("attribute");
            var options = "";
            if(Array.isArray(attributes) && attributes.length) {
              for (a = 0; a < attributes.length; a++) {
                options += attributes[a].getChild("name").getValue() + ": " + attributes[a].getChild("value").getValue() + ", ";
              }
            }
            var quantity = stockItems[j].getChild("quantity").getValue();
            container.push(stockCode);
            container.push(stockId);
            container.push(productSku);
            container.push("=VLOOKUP($" + (variants ? "A" : "C") + rNo++ + ";WC_ProductMap;7;FALSE)");
            container.push(productId);
            container.push(title);
            container.push(price);
            container.push(optionPrice);
            container.push(options);
            container.push(quantity);
            container.push(sale);
            container.push(approval);
            //productsN11.appendRow(container);
            productList.push(container);
            
            container = [];
          }
        }
        
      page++;
      } while (page < totalPages);
      
      if (productsN11.getDataRange().getNumRows() > 1) {
        productsN11.getDataRange().offset(1, 0, productsN11.getDataRange().getNumRows()-1).clearContent();
      }
      Logger.log("No of products: " + productList.length);
      productsN11.getRange(2, 1, productList.length, 12).setValues(productList);
      
      if (variants) {
        productsN11.getDataRange().removeDuplicates([1]);
        var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
        for (var n = 0; namedRanges.length; n++) {
          if (namedRanges[n].getName() == "N11_ProductList") {
            namedRanges[n].setRange(productsN11.getRange("A:L"));
            productsN11.unhideColumn(productsN11.getRange(1,1,1,2));
            productsN11.unhideColumn(productsN11.getRange(1,8,1,2));
            break;
          }
        }
      } else {
        productsN11.getDataRange().removeDuplicates([3]);
        var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
        for (var n = 0; namedRanges.length; n++) {
          if (namedRanges[n].getName() == "N11_ProductList") {
            namedRanges[n].setRange(productsN11.getRange("C:L"));
            productsN11.hideColumn(productsN11.getRange(1,1,1,2));
            productsN11.hideColumn(productsN11.getRange(1,8,1,2));
            break;
          }
        }
      }
      
    }
  } else {
      SpreadsheetApp.getUi().alert("N11 ürün tablosu bulunamadı! \nLütfen 'Products N11' adını vereceğiniz bir tablo oluşturup tekrar deneyiniz.");
      //Logger.log("N11 ürün tablosu bulunamadı! Lütfen 'Products N11' adını vereceğiniz bir tablo oluşturup tekrar deneyiniz.");
  }
      
}
