function createInvoice(platform) {
  
  //var platform = "Woocommerce";       //Use for platform tests
  console.time(platform+"Invoices");      //Start timer
  
  var invoiceConf = getInvoiceConfPrst(platform);
  
  // Create a new invoice in Paraşüt using data in order details sheet based on platform selection. TO DO: merge all order details into single sheet!
  if (platform == "Woocommerce") {
    var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Order Details");
  } else if (platform == "Hepsiburada") {
    var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HB Order Details");
  } else if (platform == "N11") {
    var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("N11 Order Details");
  } else if (platform == "Trendyol") {
    var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TY Order Details");
  }
  var orderData = orderSheet.getDataRange().getValues();
  
  /*
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = confSheet.getRange("ParasutCompanyId").getValue();
  
  var orderStatus = confSheet.getRange("WC_InvoiceStatus").getValue();
  var storeUrl = confSheet.getRange("WC_StoreUrl").getValue();
  var invoiceCategory = confSheet.getRange("WC_InvoiceCat").getValue();
  var cargo = confSheet.getRange("WC_CargoCompany").getValue();
  var cargoTaxNo = confSheet.getRange("WC_CargoCompanyTaxNo").getValue();
  var paymentType = "KREDIKARTI/BANKAKARTI";   //TO DO: Get payment Type from order/WC
  var paymentPlatform = confSheet.getRange("WC_PaymentPlatform").getValue();
  */
  
  var d = new Date();
  var invoiceDate = [
    d.getFullYear(),
    ('0' + (d.getMonth() + 1)).slice(-2),    //Month index starts at 0, increase it by 1 to get month no
    ('0' + d.getDate()).slice(-2)
    ].join('-');
  var shipmentDate = invoiceDate;
  var prevOrderId, orderRow, invoiceId, invoiceAmount;
  
  for (var i = 1; i < orderData.length;) {
  
  orderRow = i;
  
  // Check and create invoice for only orders with proper status
  if (orderData[i][2] == invoiceConf.orderStatus && orderData[i][33] == "") {
    
    var orderId = orderData[i][0];
    var orderDate = orderData[i][23];
    
    var customerId, customerName, isCompany, customerTaxOffice, customerTaxNumber, customerBalance = 0;
    var customerEmail = orderData[i][16];
    var customerPhone = orderData[i][15];
    var customerCity = orderData[i][11];
    var customerDistrict = orderData[i][10];
    var customerAddress = orderData[i][9];
    
    if (orderData[i][6] != "" && orderData[i][7] != "" && orderData[i][8] != "") {
      //console.log("Company tax fields: " + orderData[i][6] + " " + orderData[i][7] + " " + orderData[i][8]);
      isCompany = true;
      customerName = orderData[i][6];
      customerTaxOffice = orderData[i][7];
      customerTaxNumber = orderData[i][8];
      if (!validateTaxNo(customerTaxNumber)) {
        console.warn("Company tax number is not valid: " + customerTaxNumber + " changed to 11111111111");
        isCompany = false;
        customerTaxNumber ="11111111111";                 //Assign a pseudo tc number if company tax no is wrong
      }
        
    } else {
      //console.log("Person tax fields: " + orderData[i][5]);
      isCompany = false;
      customerName = orderData[i][3] + " " + orderData[i][4];
      customerTaxNumber = orderData[i][5];
      if (!validateTcNo(customerTaxNumber)) {
        console.warn("Person tax number is not valid: " + customerTaxNumber + " changed to 11111111111");
        customerTaxNumber ="11111111111";                 //Assign a pseudo tc number for wrong tc numbers
      }
    }
    //console.log("Order No: " + orderId + " / Customer email: " + customerEmail + " / Customer Tax Number: " + customerTaxNumber);
    
  
    // Step 1: Check if the customer exists using email (for individuals) and tax number (for companies)
    
    var parasutService = getParasutService();
    var response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/contacts?filter[tax_number]=' + customerTaxNumber, {
      headers: {
        Authorization: 'Bearer ' + parasutService.getAccessToken()
      }
    });
    
    if (response.getResponseCode() == 200) {
    
      // Get JSON from response and parse the data and then parse list of customers
      var json = response.getContentText();
      var data = JSON.parse(json);
      var customerFound = false;

    } else {
      console.error(response.getResponseCode() + response);
      throw new Error("Paraşüt müşteri bilgileri aramada hata: " + response.getResponseCode() + response);
    }
    
    if(Array.isArray(data.data) && data.data.length){
      // Customer array exists and is not empty
      for (var c = 0; c < data.data.length; c++) {
      
        if (data.data[c].attributes.contact_type == "company" && isCompany){
          //console.log('Tax number found!  Customer Id ' + data.data[c].id + ' Company? ' + data.data[c].attributes.contact_type + ' Name: ' + data.data[c].attributes.name);
          customerFound = true;
          customerId = data.data[c].id;
          customerBalance = data.data[c].attributes.trl_balance;
        } else if (data.data[c].attributes.contact_type == "person" && (data.data[c].attributes.email == customerEmail || data.data[c].attributes.name == customerName)) {
          //console.log('Tax number found!  Customer Id ' + data.data[c].id + ' Company? ' + data.data[c].attributes.contact_type + ' Name: ' + data.data[c].attributes.name);
          customerFound = true;
          customerId = data.data[c].id;
          customerBalance = data.data[c].attributes.trl_balance;
        }
      }
    }
    else {
      var response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/contacts?filter[email]=' + customerEmail, {
        headers: {
          Authorization: 'Bearer ' + parasutService.getAccessToken()
        }
      });
      
      if (response.getResponseCode() == 200) {
      
        json = response.getContentText();
        data = JSON.parse(json);

      } else {
        console.error(response.getResponseCode() + response);
        throw new Error("Paraşüt müşteri bilgileri aramada hata: " + response.getResponseCode() + response);
      }
      
      
      if(Array.isArray(data.data) && data.data.length){
        // Customer array exists and is not empty
        for (var c = 0; c < data.data.length; c++) {
        
          if (data.data[c].attributes.contact_type == "company" && isCompany){
            //console.log('Tax number found!  Customer Id ' + data.data[c].id + ' Company? ' + data.data[c].attributes.contact_type + ' Name: ' + data.data[c].attributes.name);
            //TO DO: Update customer to fill tax number
            customerFound = true;
            customerId = data.data[c].id;
            customerBalance = data.data[c].attributes.trl_balance;
          }
            else if (!isCompany && (data.data[c].attributes.contact_type == "company" && orderData[i][8] == data.data[c].attributes.tax_number)) {
              //console.log('Tax number found!  Customer Id ' + data.data[c].id + ' Company? ' + data.data[c].attributes.contact_type + ' Name: ' + data.data[c].attributes.name);
              customerFound = true;
              customerId = data.data[c].id;
              customerBalance = data.data[c].attributes.trl_balance;
          }
            else if (!isCompany && data.data[c].attributes.contact_type == "person") {
              customerFound = true;
              customerId = data.data[c].id;
              customerBalance = data.data[c].attributes.trl_balance;
          }
        }
      }
    }
    if (!customerFound) {
      //console.log('No such customer, create one!');
      customerId = createCustomer(customerEmail, customerName, isCompany, customerTaxOffice, customerTaxNumber, customerPhone, customerCity, customerDistrict, customerAddress);
      //console.log('New Customer ID: ' + customerId);
    }
    
    // Add product lines to invoice items array
    var invoiceItems = [];
    var productRow = 0;
    var productError = false;
    
    do {  // Loop through product lines
      
      // Use Paraşüt product code to create invoice item
      if (!productError && orderData[i][37] == "") {
        invoiceItems.push({
            type: 'sales_invoice_details',
            attributes: {
              quantity: orderData[i][28],
              unit_price: orderData[i][29],
              vat_rate: orderData[i][32]
            },
            relationships: {
              product: {
                data: {
                  id: orderData[i][27],
                  type: 'products'
                }
              }
            }}
        );
      } else {
        productError = true;
      }
      //console.log("i = " + i + " Order Id: " + orderId + ",Product id: " + orderData[i][27]);
      prevOrderId = orderId;
      i++;
      if (i < orderData.length) {
        orderId = orderData[i][0];
        if (prevOrderId == orderId) {productRow++}
      }
    } while (prevOrderId == orderId && i < orderData.length);
    
    if (!productError) {   //Start of invoicing if no error found in product lines
    
    // Step 2: Create an invoice at Parasut using order, product and customer data
    
    var invoiceData = {
      data: {
      type: 'sales_invoices',
      attributes: {
        item_type: 'invoice',
        description: platform + ' #' + prevOrderId,
        issue_date: Date(),
        due_date: Date(),
        currency: 'TRL',
        order_no: prevOrderId,
        order_date: orderDate
      },
      relationships: {
        details: {
          data : invoiceItems
        },
        contact: {
          data : {
            id: customerId,
            type: 'contacts'
          }
        },
        category: {
          data: {
            id: invoiceConf.invoiceCategory,
            type: 'item_categories'
          }
        }
      }}
    };
    
    var options = {
      'method' : 'post',
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string.
      'payload' : JSON.stringify(invoiceData),
      'headers': {
        'Authorization': 'Bearer ' + parasutService.getAccessToken()
      }
    };
    response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/sales_invoices', options);
    //console.log("Invoice Create API result: " + response.getResponseCode());
    
    if (response.getResponseCode() == 201) {     //Invoice created successfully
      
      json = response.getContentText();
      data = JSON.parse(json);
      invoiceId = data.data.id;
      invoiceAmount = data.data.attributes.net_total;
      
      console.info("Invoice created successfully of order no " + prevOrderId +
                    " for customer " + customerName + " & id " + customerId +
                    " with " + (1 + productRow) + " line items! Invoice id: " + invoiceId);
      
      var cell = orderSheet.getRange("A1");
      for (var r = 0; r <= productRow; r++) {
        cell.offset(orderRow + r, 33).setValue(invoiceId);
      }
    
    } else {
      
      orderSheet.getRange("A1").offset(orderRow, 37).setValue(response.getResponseCode() + response);
      console.error("Could not create invoice for order no " + prevOrderId + " Response: " + response);
      throw new Error(prevOrderId + " nolu sipariş için Paraşüt faturası düzenlenemedi. Hata: " + response.getResponseCode());
    }
    
    
    
    // Step 3: Make Payment for the invoice
    
    //var paymentAccount = confSheet.getRange("WC_PaymentAccount").getValue();
    if (invoiceConf.paymentAccount != "" && customerBalance >= 0) {
      paymentId = makePayment(invoiceId, invoiceConf.paymentAccount, invoiceAmount, Date());   // TO DO: Get payment date from order and make payment with correct date
      console.log("Payment Account: " + invoiceConf.paymentAccount + "  PaymentId: " + paymentId);
    }
    
    
    // Step 4: Create e-invoice
    
    var response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/e_invoice_inboxes?filter[vkn]=' + customerTaxNumber, {
      headers: {
        Authorization: 'Bearer ' + parasutService.getAccessToken()
      }
    });

    if (response.getResponseCode() == 200) {
      
      // Get JSON from response and parse the data
      json = response.getContentText();
      data = JSON.parse(json);

    } else {
      orderSheet.getRange("A1").offset(orderRow, 37).setValue(response.getResponseCode() + response);
      console.error("Could not determine e-invoice eligibility for " + customerId + " Response: " + response);
      throw new Error(customerId + " nolu müşteri için e-fatura düzenlenmesi sağlanamıyor. Hata: " + response.getResponseCode());
    }

    
    if(Array.isArray(data.data) && data.data.length){
        // Customer is an e-invoice customer
        var eInvoiceInbox = data.data[0].attributes.e_invoice_address;
        //console.log('Tax Number found!  Tax Number: ' + data.data[0].attributes.vkn + ' Address: ' + eInvoiceInbox + ' Name: ' + data.data[0].attributes.name);
        //Create an e-invoice
        
        var eInvoiceData = {
          data: {
            type: 'e_invoices',
            attributes: {
              vat_withholding_code: '',
              vat_exemption_reason_code: '',
              vat_exemption_reason: '',
              note: '',
              excise_duty_codes: [],
              scenario: 'commercial',
              to: eInvoiceInbox,
            },
            relationships: {
              invoice: {
                data: {
                  id: invoiceId,
                  type: "sales_invoices"
                }
              }
            }
          }
        }
        
        var options = {
          'method' : 'post',
          'contentType': 'application/json',
          // Convert the JavaScript object to a JSON string.
          'payload' : JSON.stringify(eInvoiceData),
          'headers': {
            'Authorization': 'Bearer ' + parasutService.getAccessToken()
          }
        };
        
        response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/e_invoices', options);
        
        if (response.getResponseCode() == 201 || response.getResponseCode() == 202) {
          
          console.log(response.getResponseCode());
          // Get JSON from response and parse the data
          json = response.getContentText();
          data = JSON.parse(json);
        
        } else {
          orderSheet.getRange("A1").offset(orderRow, 37).setValue(response.getResponseCode() + response);
          console.error("Could not create e-invoice for " + invoiceId + " Response: " + response);
          throw new Error(invoiceId + " nolu fatura için Paraşüt e-fatura düzenlenemedi. Hata: " + response.getResponseCode());
        }
        
        trackingId = data.data.id;
        eInvoiceStatus = data.data.attributes.status;
        console.log('E-invoice is being created. Tracking Id: ' + trackingId + ' Status: ' + eInvoiceStatus);
        
        
      }
      else {
        
        //Create an e-archive invoice
        //console.log('Store URL: ' + invoiceConf.storeUrl + '  Payment Type: ' + invoiceConf.paymentType + '  Payment Platform: ' + invoiceConf.paymentPlatform + '  Date: ' + invoiceDate)
        var eArchiveData = {
          data: {
            type: 'e_archives',
            attributes: {
              vat_withholding_code: '',
              vat_exemption_reason_code: '',
              vat_exemption_reason: '',
              note: '',
              excise_duty_codes: [],
              internet_sale: {
                url: invoiceConf.storeUrl,
                payment_type: invoiceConf.paymentType,
                payment_platform: invoiceConf.paymentPlatform,
                payment_date: invoiceDate
              },
              shipment: {
                title: invoiceConf.cargo,
                vkn: invoiceConf.cargoTaxNo,
                name: '',
                tckn: '',
                date: shipmentDate
              }
            },
            relationships: {
              sales_invoice: {
                data: {
                  id: invoiceId,
                  type: "sales_invoices"
                }
              }
            }
          }
        }
        
        var options = {
          'method' : 'post',
          'contentType': 'application/json',
          // Convert the JavaScript object to a JSON string.
          'payload' : JSON.stringify(eArchiveData),
          'headers': {
            'Authorization': 'Bearer ' + parasutService.getAccessToken()
          }
        };
        
        response = UrlFetchApp.fetch(invoiceConf.baseUrl + invoiceConf.companyId + '/e_archives', options);
        
        if (response.getResponseCode() == 201 || response.getResponseCode() == 202) {
          
          // Get JSON from response and parse the data
          json = response.getContentText();
          data = JSON.parse(json);
        
        } else {
          orderSheet.getRange("A1").offset(orderRow, 37).setValue(response.getResponseCode() + response);
          console.error("Could not create e-archive invoice for " + invoiceId + " Response: " + response);
          throw new Error(invoiceId + " nolu fatura için Paraşüt e-arşiv faturası düzenlenemedi. Hata: " + response.getResponseCode());
        }
        
        trackingId = data.data.id;
        eInvoiceStatus = data.data.attributes.status;
        console.log('E-ArchiveInvoice is being created. Tracking Id: ' + trackingId + ' Status: ' + eInvoiceStatus);
        
      }
      
    Utilities.sleep(6000);      //Wait for 6 seconds
    
    //console.log("Veri Alanı: " + cell.offset(orderRow, 33).getA1Notation() + " GetRowIndex: " + cell.offset(orderRow, 33).getRowIndex());
    
    } else {
      console.error("Invoice can not be created for order no " + prevOrderId + " due to errors in " + (1 + productRow) + " product line(s).");
      orderSheet.getRange("A1").offset(orderRow, 37).setValue("Error: Invoice not created!");
    }
  }
  else {
    i++;
  }
  }      // End of main order data loop
  console.timeEnd(platform+"Invoices");      // Stops the timer, logs execution duration.
}



function getInvoiceConfPrst(platform) {
  
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  
  if (platform == "Woocommerce") {
  
    var confPrstObj = {
      confSheet: confSheet,
      companyId: confSheet.getRange("ParasutCompanyId").getValue(),
      baseUrl: confSheet.getRange("ParasutBaseUrl").getValue(),
      orderStatus: confSheet.getRange("WC_InvoiceStatus").getValue(),
      storeUrl: confSheet.getRange("WC_StoreUrl").getValue(),
      invoiceCategory: confSheet.getRange("WC_InvoiceCat").getValue(),
      cargo: confSheet.getRange("WC_CargoCompany").getValue(),
      cargoTaxNo: confSheet.getRange("WC_CargoCompanyTaxNo").getValue(),
      paymentType: "KREDIKARTI/BANKAKARTI",        //TO DO: Get payment Type from order/WC
      paymentAccount: confSheet.getRange("WC_PaymentAccount").getValue(),
      paymentPlatform: confSheet.getRange("WC_PaymentPlatform").getValue()
    };
  
  } else if (platform == "Hepsiburada") {
  
    var confPrstObj = {
      confSheet: confSheet,
      companyId: confSheet.getRange("ParasutCompanyId").getValue(),
      baseUrl: confSheet.getRange("ParasutBaseUrl").getValue(),
      orderStatus: confSheet.getRange("HB_InvoiceStatus").getValue(),
      storeUrl: confSheet.getRange("HB_StoreUrl").getValue(),
      invoiceCategory: confSheet.getRange("HB_InvoiceCat").getValue(),
      cargo: confSheet.getRange("HB_CargoCompany").getValue(),
      cargoTaxNo: confSheet.getRange("HB_CargoCompanyTaxNo").getValue(),
      paymentType: "KREDIKARTI/BANKAKARTI",        //TO DO: Get payment Type from order/HB
      paymentAccount: confSheet.getRange("HB_PaymentAccount").getValue(),
      paymentPlatform: confSheet.getRange("HB_PaymentPlatform").getValue()
    };
  
  } else if (platform == "N11") {
  
    var confPrstObj = {
      confSheet: confSheet,
      companyId: confSheet.getRange("ParasutCompanyId").getValue(),
      baseUrl: confSheet.getRange("ParasutBaseUrl").getValue(),
      orderStatus: confSheet.getRange("N11_InvoiceStatus").getValue(),
      storeUrl: confSheet.getRange("N11_StoreUrl").getValue(),
      invoiceCategory: confSheet.getRange("N11_InvoiceCat").getValue(),
      cargo: confSheet.getRange("N11_CargoCompany").getValue(),
      cargoTaxNo: confSheet.getRange("N11_CargoCompanyTaxNo").getValue(),
      paymentType: "KREDIKARTI/BANKAKARTI",        //TO DO: Get payment Type from order/N11
      paymentAccount: confSheet.getRange("N11_PaymentAccount").getValue(),
      paymentPlatform: confSheet.getRange("N11_PaymentPlatform").getValue()
    };
  
  } else if (platform == "Trendyol") {
  
    var confPrstObj = {
      confSheet: confSheet,
      companyId: confSheet.getRange("ParasutCompanyId").getValue(),
      baseUrl: confSheet.getRange("ParasutBaseUrl").getValue(),
      orderStatus: confSheet.getRange("TY_InvoiceStatus").getValue(),
      storeUrl: confSheet.getRange("TY_StoreUrl").getValue(),
      invoiceCategory: confSheet.getRange("TY_InvoiceCat").getValue(),
      cargo: confSheet.getRange("TY_CargoCompany").getValue(),
      cargoTaxNo: confSheet.getRange("TY_CargoCompanyTaxNo").getValue(),
      paymentType: "KREDIKARTI/BANKAKARTI",        //TO DO: Get payment Type from order/TY
      paymentAccount: confSheet.getRange("TY_PaymentAccount").getValue(),
      paymentPlatform: confSheet.getRange("TY_PaymentPlatform").getValue()
    };
  
  }
  return confPrstObj;
}


function createCustomer(email, name, isCompany, taxOffice, taxNumber, phone, city, district, address) {
  
  // Make a POST request to Paraşüt with a JSON payload for creating a new Customer.
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = confSheet.getRange("ParasutCompanyId").getValue();

  if (isCompany) {
    var contactType = "company";
    var customerCategory = confSheet.getRange("ParasutCompanyCat").getValue();
  } else {
    var contactType = "person";
    var customerCategory = confSheet.getRange("ParasutPersonCat").getValue();
  }
  
  var customerData = {
    data: {
      type: 'contacts',
      attributes: {
        account_type: 'customer',
        email: email,
        name: name,
        contact_type: contactType,
        tax_number: taxNumber,
        tax_office: taxOffice,
        phone: phone,
        city: city,
        district: district,
        address: address
      },
      relationships: {
        category: {
          data: {
            id: customerCategory,
            type: 'item_categories'
          }
        }
      }
    }
  };
  var parasutService = getParasutService();
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(customerData),
    'headers': {
        'Authorization': 'Bearer ' + parasutService.getAccessToken()
      }
  };
  var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/contacts', options);
  Logger.log(response);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data.data.id;
  
}

function makePayment (invoiceId, paymentAccount, invoiceAmount, paymentDate) {

  var configuration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = configuration.getRange("ParasutCompanyId").getValue();

  var paymentData = {
    data: {
      type: 'payments',
      attributes: {
        description: 'QuickOrders Otomatik Ödeme',
        account_id: paymentAccount,
        date: paymentDate,
        amount: invoiceAmount,
        exchange_rate: '1.0'
      }
    }
  }
  
  var parasutService = getParasutService();
  var options = {
      'method' : 'post',
      'contentType': 'application/json',
      // Convert the JavaScript object to a JSON string.
      'payload' : JSON.stringify(paymentData),
      'headers': {
        'Authorization': 'Bearer ' + parasutService.getAccessToken()
      }
    };
    
    var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/sales_invoices/' + invoiceId + '/payments', options);
    var json = response.getContentText();
    var data = JSON.parse(json);
    return data.data.id;
}

function getProducts() {
  
  var configuration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = configuration.getRange("ParasutCompanyId").getValue();
  var pUrl = configuration.getRange("ParasutBaseUrl").getValue();
  var parasutService = getParasutService();
  
  // Get Product list & codes from Parasut and append into Products sheet
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Paraşüt'te faturalama için kullanılacak ürün listesi güncellenecek." +
                          "\nDİKKAT, bu işlem Paraşüt'ten daha önce indirilmiş ürün listelerini siler" +
                          " ve yaptığınız ürün eşleştirmelerini yeniden yapmanız gerekebilir!" + 
                          "\n\n Devam etmek istiyor musunuz?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
  
    var productList = [];
    
    var page = 1;
    do {
    
      // Get list of products from Parasut using OAuth2
      
      var result = UrlFetchApp.fetch(pUrl + companyId + '/products?page[number]=' + page , {
          headers: {
          Authorization: 'Bearer ' + parasutService.getAccessToken()
          }
      });
      
      if (result.getResponseCode() == 200) {
        // Get JSON from response and parse the data and then parse list of products
        var json = result.getContentText();
        var data = JSON.parse(json);
        //var productKeys = Object.keys(data.data[0].attributes);
      } else {
        console.error(result);
        SpreadsheetApp.getUi().alert("Paraşüt ürün listesine erişilemiyor! \nLütfen daha sonra tekrar deneyiniz.");
        throw "Paraşüt Ürün Listesi Erişim Hatası: " + result.getResponseCode();
      }
      
      Logger.log('Current Page: ' + data.meta.current_page + '  No of pages: ' + data.meta.total_pages);
      
      var products = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
      //products.appendRow(productKeys);
      
      for (var x in data.data) {
        var container = [];
        if (data.data[x].attributes.archived == false) {
          container.push(data.data[x].attributes.code);
          container.push(data.data[x].attributes.name);
          container.push(data.data[x].id);
          container.push(data.data[x].attributes.unit);
          container.push(Number(data.data[x].attributes.vat_rate));
          container.push(data.data[x].attributes.inventory_tracking)
          container.push(Number(data.data[x].attributes.stock_count));
          container.push("");                             //TO DO: add category info
          container.push(Number(data.data[x].attributes.list_price));
          container.push(data.data[x].attributes.currency);
          container.push(Number(data.data[x].attributes.list_price_in_trl));
          //rowArray.push(data.data[x].attributes.archived);
        }
        
        //products.appendRow(rowArray);
        productList.push(container);
        
        Logger.log('Row Array No ' + x + ' = ' + container);
      }
    
    Utilities.sleep(800);      //Wait for 1/2 second
    page++;
    } while (data.meta.current_page < data.meta.total_pages);
    
    if (products.getDataRange().getNumRows() > 1) {
      products.getDataRange().offset(1, 0, products.getDataRange().getNumRows()-1).clearContent();
    }
    Logger.log("No of products: " + productList.length);
    products.getRange(2, 1, productList.length, 11).setValues(productList);
    products.getDataRange().removeDuplicates([1]);
    var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
    for (var n = 0; namedRanges.length; n++) {
      if (namedRanges[n].getName() == "Products_Parasut") {
        namedRanges[n].setRange(products.getRange("A:K"));
        break;
      }
    }
  }
}

function callPayment () {

    var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
    var parasutService = getParasutService();
    var companyId = confSheet.getRange("ParasutCompanyId").getValue();
    var customerTaxNumber = 4650687329;
    var invoiceId = 19948810;
    var invoiceAmount = "35.90";
    
    var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/contacts?filter[tax_number]=' + customerTaxNumber, {
      headers: {
        Authorization: 'Bearer ' + parasutService.getAccessToken()
      }
    });
    
    // Get JSON from response and parse the data and then parse list of customers
    var json = response.getContentText();
    var data = JSON.parse(json);
    var customerId = data.data[0].id;
    var customerBalance = data.data[0].attributes.trl_balance;
    
    Logger.log("Customer Id: " + customerId + "  Customer Balance: " + customerBalance + "  Invoice Amount: " + invoiceAmount);

    // TEST: Make Payment for the invoice
    
    var paymentAccount = confSheet.getRange("WC_PaymentAccount").getValue();
    if (paymentAccount != "" && customerBalance >= 0) {
      paymentId = makePayment(invoiceId, paymentAccount, invoiceAmount, Date());   // TO DO: Get payment date from order and make payment with correct date
      //Logger.log("Payment Account: " + paymentAccount + "  PaymentId: " + paymentId);
    }

}

function getCategories() {
  
  // Get list of categories from Parasut using OAuth2
  //var configuration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  //var companyId = configuration.getRange("ParasutCompanyId").getValue();
  var parasutService = getParasutService();
  var response = UrlFetchApp.fetch('https://api.parasut.com/v4/217404/item_categories?filter[category_type]=Contact', {
    headers: {
      Authorization: 'Bearer ' + parasutService.getAccessToken()
    }
  });
  
  // Get JSON from response and parse the data and then parse list of categories
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  var j = 0;
  var rowArray = [];
  //var rangeArray = [];
  var sheet = SpreadsheetApp.getActiveSheet();
  
  for (var x in data.data) {
    //var invoiceAttributes = invoiceData[x].attributes
    rowArray[j] = data.data[x].id;
    j++;
    for (var y in data.data[x].attributes) {
      if (y == "name") {
        rowArray[j] = data.data[x].attributes.name;
        j++;
      } else if (y == "category_type") {
        rowArray[j] = data.data[x].attributes.category_type;
        j++;
      }
    }
    sheet.appendRow(rowArray);
    //rangeArray[x] = rowArray.slice();
    Logger.log('Row Array No ' + x + ' = ' + rowArray);
    j = 0;
  }
  
}



function getAccounts() {

  var configuration = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var companyId = configuration.getRange("ParasutCompanyId").getValue();
  var parasutService = getParasutService();
  var page = 1;
  do {
  
  // Get list of cash & bank accounts from Parasut using OAuth2
  
  var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/accounts?page[number]=' + page , {
    headers: {
      Authorization: 'Bearer ' + parasutService.getAccessToken()
    }
  });
  
  // Get JSON from response and parse the data and then parse list of accounts
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  var j = 0;
  Logger.log('Current Page: ' + data.meta.current_page + '  No of pages: ' + data.meta.total_pages);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  
  for (var x in data.data) {
    var rowArray = [];
    if (data.data[x].attributes.archived == false) {
      rowArray.push(data.data[x].id);
      rowArray.push(data.data[x].attributes.name);
      rowArray.push(data.data[x].attributes.currency);
      rowArray.push(data.data[x].attributes.account_type);
      //rowArray.push(data.data[x].attributes.archived);
    }
    
    sheet.appendRow(rowArray);
    
    Logger.log('Row Array No ' + x + ' = ' + rowArray);
  }
  page++;
  } while (data.meta.current_page < data.meta.total_pages);

}

/**
 * Logs the redirect URI to register.
 */
function getRedirectUri() {
  var service = getParasutService();
  Logger.log(service.getRedirectUri());
  return service.getRedirectUri();
}


function getParasutService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  
  var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
  var clientId = confSheet.getRange("ParasutClientId").getValue();
  var clientSecret = confSheet.getRange("ParasutClientSecret").getValue();
  
  return OAuth2.createService('parasut')

      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://api.parasut.com/oauth/authorize')
      .setTokenUrl('https://api.parasut.com/oauth/token')

      // Set the client ID and secret, by requesting from Parasut support (destek@parasut.com).
      .setClientId(clientId)
      .setClientSecret(clientSecret)
      //.setClientId('2fa1db42bf2fc882e51d09bba7e03f30f99e72a9f67bfbd4c128fa06c54cf7e0')
      //.setClientSecret('b96a33ab3ca84a1d1e2642e36716409f73efb783121e743061a0203922eeebb9')

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authParasutCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
  
}

function authParasutCallback(request) {
  var parasutService = getParasutService();
  var isAuthorized = parasutService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function showParasutSidebar() {
  var parasutService = getParasutService();
  if (!parasutService.hasAccess()) {
    var authorizationUrl = parasutService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
  Logger.log('You already have an access to Parasut');
  }
}

function logout() {
  var service = getParasutService()
  service.reset();
}

function testParasutAccess() {

    var confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuration");
    var parasutService = getParasutService();
    var companyId = confSheet.getRange("ParasutCompanyId").getValue();

    var response = UrlFetchApp.fetch('https://api.parasut.com/v4/' + companyId + '/contacts?filter[tax_number]=4780508560' , {
      headers: {
        Authorization: 'Bearer ' + parasutService.getAccessToken()
      }
    });
    Logger.log(parasutService.getAccessToken());
    // Get JSON from response and parse the data and then parse list of customers
    var json = response.getContentText();
    var data = JSON.parse(json);
    Logger.log(response.getResponseCode());
}
