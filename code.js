// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/generate-pdfs

/*
Copyright 2022 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

// TODO: To test this solution, set EMAIL_OVERRIDE to true and set EMAIL_ADDRESS_OVERRIDE to your email address.
const EMAIL_OVERRIDE = true;
const EMAIL_ADDRESS_OVERRIDE = 'dispatch@xxx.com';
// const EMAIL_ADDRESS_OVERRIDE = '.com';

// Application constants
const APP_TITLE = 'Dispatch PDFs';
const OUTPUT_FOLDER_NAME = "Dispatch PDFs";
const DUE_DATE_NUM_DAYS = 15

// Sheet name constants. Update if you change the names of the sheets.
const CUSTOMERS_SHEET_NAME = 'Customers';
const PRODUCTS_SHEET_NAME = 'Products';
const TRANSACTIONS_SHEET_NAME = 'Transactions';
const INVOICES_SHEET_NAME = 'Dispatches';
const TRUCK_SHEET_NAME = 'Freights';
const WAREHOUSE_SHEET_NAME = 'Warehouse';
const INVOICE_TEMPLATE_SHEET_NAME = 'Dispatch Template';

// Email constants
const EMAIL_SUBJECT = 'Order Preparation Form ';
const EMAIL_BODY = 'Hi Pauline:\r  \r        Please review this order preparation and confirm. \r \r \r Best Regards,   \r \r Eason';

const EMAIL_TRUCK_SUBJECT = 'Dispatch Instruction ';
const EMAIL_TRUCK_BODY = '\rPlease review this order preparation and confirm.';


const EMAIL_WAREHOUSE_SUBJECT = 'Order Preparation';
// const EMAIL_WAREHOUSE_BODY = 'Hi '
/**
 * Iterates through the worksheet data populating the template sheet with 
 * customer data, then saves each instance as a PDF document.
 * 
 * Called by user via custom menu item.
 */
function processDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoiceTemplateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);

  // Gets data from the storage sheets as objects.
  const customers = dataRangeToObject(customersSheet);
  const products = dataRangeToObject(productsSheet);
  const transactions = dataRangeToObject(transactionsSheet);
  const dispatches = dataRangeToObject(invoicesSheet);

  ss.toast('Creating Dispatch ', APP_TITLE, 1);
  const invoices = [];

  let transactions_by_dispatch = {};


  // Iterates for each customer calling createInvoiceForCustomer routine.
  customers.forEach(function (customer) {
    ss.toast(`Creating Dispatch for ${customer.customer_name}`, APP_TITLE, 1);
    transactions.forEach(function (transaction) {
      if (transaction.customer_name != customer.customer_name) {
      }
      else {
        const invoiceNumber = customer.customer_id + transaction.dispatch_num;
        let dispatchOrders = dispatches.filter(function (dispatch_ord) {
          return dispatch_ord.invoice_no_ == invoiceNumber;
        })
        if (dispatchOrders.length > 0) {
          console.log("Invoice number ", invoiceNumber, " already created");
        }
        else {
          if (transactions_by_dispatch.hasOwnProperty(invoiceNumber)) {
            transactions_by_dispatch[invoiceNumber].push(transaction);
          }
          else {
            transactions_by_dispatch[invoiceNumber] = [transaction];
          }
        }
      }
    });

    for (const [key, transac_list] of Object.entries(transactions_by_dispatch)) {
      let invoice = createInvoiceForCustomer(
        customer, products, transac_list, invoiceTemplateSheet, ss.getId(), dispatches, true);
      if (invoice != false) {
        invoices.push(invoice);
      }
    }



  });
  //if(invoices.length > 0){
  // Writes invoices data to the sheet.
  // invoicesSheet.getRange(2, 1, invoices.length, invoices[0].length).setValues(invoices);
  invoices.forEach(function (invoice) {
    invoicesSheet.appendRow(invoice);
  })
  //}

}

/**
 * Iterates through the worksheet data populating the template sheet with 
 * customer data, then saves each instance as a PDF document.
 * 
 * Called by user via custom menu item.
 */
function processDocuments2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  const productsSheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  const transactionsSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoiceTemplateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);

  // Gets data from the storage sheets as objects.
  const customers = dataRangeToObject(customersSheet);
  const products = dataRangeToObject(productsSheet);
  const transactions = dataRangeToObject(transactionsSheet);
  const dispatches = dataRangeToObject(invoicesSheet);

  ss.toast('Creating Dispatch ', APP_TITLE, 1);
  const invoices = [];

  let transactions_by_dispatch = {};


  // Iterates for each customer calling createInvoiceForCustomer routine.
  customers.forEach(function (customer) {
    ss.toast(`Creating Dispatch for ${customer.customer_name}`, APP_TITLE, 1);
    transactions.forEach(function (transaction) {
      if (transaction.customer_name != customer.customer_name) {
      }
      else {
        const invoiceNumber = customer.customer_id + transaction.dispatch_num;
        let dispatchOrders = dispatches.filter(function (dispatch_ord) {
          return dispatch_ord.invoice_no_ == invoiceNumber;
        })
        if (dispatchOrders.length > 0) {
          console.log("Invoice number ", invoiceNumber, " already created");
        }
        else {
          if (transactions_by_dispatch.hasOwnProperty(invoiceNumber)) {
            transactions_by_dispatch[invoiceNumber].push(transaction);
          }
          else {
            transactions_by_dispatch[invoiceNumber] = [transaction];
          }
        }
      }
    });

    for (const [key, transac_list] of Object.entries(transactions_by_dispatch)) {
      let invoice = createInvoiceForCustomer(
        customer, products, transac_list, invoiceTemplateSheet, ss.getId(), dispatches, false);
      if (invoice != false) {
        invoices.push(invoice);
      }
    }



  });
  //if(invoices.length > 0){
  // Writes invoices data to the sheet.
  // invoicesSheet.getRange(2, 1, invoices.length, invoices[0].length).setValues(invoices);
  invoices.forEach(function (invoice) {
    invoicesSheet.appendRow(invoice);
  })
  //}

}

/**
 * Processes each customer instance with passed in data parameters.
 * 
 * @param {object} customer - Object for the customer
 * @param {object} products - Object for all the products
 * @param {object} transactions - Object for all the transactions
 * @param {object} invoiceTemplateSheet - Object for the invoice template sheet
 * @param {string} ssId - Google Sheet ID     
 * Return {array} of instance customer invoice data
 */
function createInvoiceForCustomer(customer, products, transactions, templateSheet, ssId, dispatches, usePdf) {
  let customerTransactions = transactions.filter(function (transaction) {
    return transaction.customer_name == customer.customer_name;
  });

  // Clears existing data from the template.
  clearTemplateSheet();

  let lineItems = [];
  let totalAmount = 0;
  let totalPallets = 0;
  let dispatch_num = 0;
  let delivery_address = '';
  let delivery_instruction = '';
  let po_num = '';
  let notes = '';
  let freight_company = '';
  let freight_cost = 0;

  let warehouse = '';
  let seal_tag = customer.seal_tag;
  let delivery_date = '';

  customerTransactions.forEach(function (lineItem) {
    let lineItemProduct = products.filter(function (product) {
      return product.sku_id == lineItem.sku;
    })[0];
    const qty = parseInt(lineItem.pallets);
    const batchNumber = lineItem.batch;
    const weight = parseFloat(lineItemProduct.weight).toFixed(2);
    const amount = parseFloat(qty * weight).toFixed(2);
    dispatch_num = lineItem.dispatch_num;
    delivery_address = lineItem.delivery_address;
    delivery_instruction = lineItem.delivery_instruction;
    po_num = lineItem.po;

      // if the customer name is Bill Bar --> add BBCOL number
// =================================================================================================================================
    if (customer.name == "Bill Bar"){
      notes += "tesing";
      notes += "BBCOL# is ";
    }

    
// =================================================================================================================================


    // if the warehouse is ITL add some to the notes

    if(warehouse.warehouse_name == "ITL"){

      if (warehouse.delivery_method == "FOB"){
        notes += "COLLECT";
      }

      if (warehouse.delivery_method == "Delivery"){
        notes += "3rd PARTY";
      }
    }

    if (seal_tag == "Yes") {
      // notes += lineItem.notes ? "" + lineItem.notes : "";
      notes = "Please put seal tag on the truck \n" ;
      // notes += "Please put seal tag on the truck \n"
    }
    else {
      // notes = "" + lineItem.notes;
      notes += lineItem.notes ? "" + lineItem.notes : "";
      notes += '\n';
    }
    notes += lineItem.notes ? "" + lineItem.notes : "";

    freight_company = lineItem.freight_company;
    freight_cost = lineItem.freight_quote
    warehouse = lineItem.warehouse;
    delivery_date = lineItem.date;

    lineItems.push([lineItemProduct.sku_name, lineItemProduct.sku_description, batchNumber, lineItemProduct.package, qty, weight, amount]);
    totalAmount += parseFloat(amount);
    totalPallets += parseFloat(qty);
  });

  if (lineItems.length == 0) {
    return false;
  }
  // Generates a random invoice number. You can replace with your own document ID method.
  const invoiceNumber = customer.customer_id + dispatch_num;
  let dispatchOrders = dispatches.filter(function (dispatch_ord) {
    return dispatch_ord.invoice_no_ == invoiceNumber;
  })
  if (dispatchOrders.length > 0) {
    console.log("Invoice number ", invoiceNumber, " already created");
    return false;
  }

  // Calulates dates.
  const todaysDate = new Date().toDateString();

  // Sets values in the template.
  templateSheet.getRange('B10').setValue(customer.customer_name)
  templateSheet.getRange('B11').setValue(delivery_address)
  templateSheet.getRange('G10').setValue(invoiceNumber)
  templateSheet.getRange('G12').setValue(todaysDate)
  templateSheet.getRange('G14').setValue(po_num)
  templateSheet.getRange('B25').setValue(notes)
  //templateSheet.getRange('F14').setValue(dueDate)
  templateSheet.getRange(18, 2, lineItems.length, 7).setValues(lineItems);


  // Cleans up and creates PDF.
  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
  if (usePdf) {
    const pdf = createPDF(ssId, templateSheet, `Dispatch#${invoiceNumber}`);
    return [invoiceNumber, po_num, delivery_date, customer.customer_name, customer.email, warehouse, totalAmount, totalPallets, delivery_address, pdf.getUrl(), freight_company, freight_cost, 'No', 'No', delivery_instruction];
  }
  const xlsx = createXLSX(ssId, templateSheet, `Dispatch#${invoiceNumber}`);
  return [invoiceNumber, po_num, delivery_date, customer.customer_name, customer.email, warehouse, totalAmount, totalPallets, delivery_address, xlsx.getUrl(), freight_company, freight_cost, 'No', 'No', delivery_instruction];
}

/**
* Resets the template sheet by clearing out customer data.
* You use this to prepare for the next iteration or to view blank
* the template for design.
* 
* Called by createInvoiceForCustomer() or by the user via custom menu item.
*/
function clearTemplateSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
  // Clears existing data from the template.
  const rngClear = templateSheet.getRangeList(['B10:B11', 'G10', 'G12', 'G14', 'B18:H22', 'B25:B26']).getRanges()
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
  // This sample only accounts for six rows of data 'B18:G24'. You can extend or make dynamic as necessary.
  templateSheet.getRange(18, 2, 7, 6).clearContent();
}

/**
 * Creates a PDF for the customer given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created
 * @return {file object} PDF file as a blob
 */
function createXLSX(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=xlsx&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.xlsx');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const xlsxFile = folder.createFile(blob);
  return xlsxFile;
}

/**
 * Creates a PDF for the customer given sheet.
 * @param {string} ssId - Id of the Google Spreadsheet
 * @param {object} sheet - Sheet to be converted as PDF
 * @param {string} pdfName - File name of the PDF being created
 * @return {file object} PDF file as a blob
 */
function createPDF(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const pdfFile = folder.createFile(blob);
  return pdfFile;
}


/**
 * Sends emails with PDF as an attachment.
 * Checks/Sets 'Email Sent' column to 'Yes' to avoid resending.
 * 
 * Called by user via custom menu item.
 */
function sendEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();

  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));

  ss.toast('Emailing Invoices', APP_TITLE, 1);
  invoices.forEach(function (invoice, index) {

    if (invoice.approval_email_sent != 'Yes') {
      ss.toast(`Emailing Invoice for ${invoice.customer}`, APP_TITLE, 1);

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
      const attachment = DriveApp.getFileById(fileId);

      let recipient = invoice.email;
      if (EMAIL_OVERRIDE) {
        recipient = EMAIL_ADDRESS_OVERRIDE
      }

      if (attachment.getMimeType().includes("pdf")) {
        GmailApp.sendEmail(recipient, EMAIL_SUBJECT + ` ${invoice.invoice_no_}`, EMAIL_BODY, {
          attachments: [attachment.getAs(MimeType.PDF)],
          name: APP_TITLE + ` ${invoice.invoice_no_}`
        });
      }
      else {
        console.log("Sending xlsx");
        GmailApp.sendEmail(recipient, EMAIL_SUBJECT + ` ${invoice.invoice_no_}`, EMAIL_BODY, {
          attachments: [attachment],
          name: APP_TITLE + ` ${invoice.invoice_no_}`
        });
      }


      invoicesSheet.getRange(index + 2, 16).setValue('Yes');
    }
  });
}

function sendEmailToFreight() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();
  const truckSheet = ss.getSheetByName(TRUCK_SHEET_NAME);
  const truckData = truckSheet.getRange(1, 1, truckSheet.getLastRow(), truckSheet.getLastColumn()).getValues();
  const warehouseSheet = ss.getSheetByName(WAREHOUSE_SHEET_NAME);
  const warehouseData = warehouseSheet.getRange(1, 1, warehouseSheet.getLastRow(), warehouseSheet.getLastColumn()).getValues();

  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));

  const keysT = truckData.splice(0, 1)[0];
  const trucks = getObjects(truckData, createObjectKeys(keysT));

  const keysW = warehouseData.splice(0, 1)[0];
  const warehouses = getObjects(warehouseData, createObjectKeys(keysW));

  let freight_company = "";
  let warehouse = "";

  console.log('Emailing Freight Company ' + EMAIL_TRUCK_SUBJECT);
  invoices.forEach(function (invoice, index) {

    if (invoice.freight_email_sent != 'Yes') {
      ss.toast(`Emailing Freight Company for ${invoice.customer}`, APP_TITLE, 1);
      trucks.forEach(function (each_truck) {
        if (invoice.truck == each_truck.freight_company) {
          freight_company = each_truck;
        }
      })

      warehouses.forEach(function (each_warehouse) {
        if (invoice.warehouse == each_warehouse.warehouse_name) {
          warehouse = each_warehouse;
        }
      })

      // var delivery_instruction = warehouse.delivery_instruction;
      // delivery_instruction = "\r\r" + delivery_instruction.replace("{{PO_NUM}}", invoice.po_num).replace("{{DISPATCH_NUM}}", invoice.invoice_no_).replace("{{PALLETS_NUM}}", invoice.total_pallets).replace("{{TOTAL_WEIGHT}}", invoice.total_weight).replace("{{DELIVERY_DATE}}", invoice.delivery_date).replace("{{DELIVERY_ADDRESS}}", invoice.delivery_address).replace("{{FREIGHT_COST}}", invoice.truck_cost).replace("{{DELIVERY_INSTRUCTION}}", invoice.instruction);



//=================================================================================================================================


var delivery_instruction = warehouse.delivery_instruction;


// delivery_instruction = "<br><br>" + delivery_instruction
//     .replace("{{PO_NUM}}", `<span style="color:red;">${invoice.po_num}</span>`)
//     .replace("{{DISPATCH_NUM}}", `<span style="color:blue;"><strong>${invoice.invoice_no_}</strong></span><br>`)
//     .replace("{{PALLETS_NUM}}", `<span style="color:green;"><strong>${invoice.total_pallets}</span></strong> pallets <br>`  )
//     .replace("{{TOTAL_WEIGHT}}", `<span style="color:orange;">${invoice.total_weight}</span> KG<br> `)
//     .replace("{{DELIVERY_DATE}}", `<span style="color:purple;">${invoice.delivery_date}</span><br>`)
//     .replace("{{DELIVERY_ADDRESS}}", `<span style="color:teal;">${invoice.delivery_address}</span><br>`)
//     .replace("{{FREIGHT_COST}}", `<span style="color:brown;">${invoice.truck_cost}</span><br>`)
//     .replace("{{DELIVERY_INSTRUCTION}}", `<span style="color:navy;">${invoice.instruction}</span><br>`);

delivery_instruction = `
<div style="font-family: Arial, sans-serif; font-size: 12px;">
  <pre>
  <br>
  <span style="color:black;"><strong>Regarding:</strong></span><br>
    <span style="color:black;"><strong>CUSTOMER PO #: ${invoice.po_num}</strong></span><br>
    <span style="color:black;"><strong>DISPATCH #: ${invoice.invoice_no_}</strong></span><br>
    <span style="color:red;">TRUCK FREIGHT COSTS: $${invoice.truck_cost}</span><br>
    <span style="color:green;"><strong>TOTAL # OF PALLETS: ${invoice.total_pallets} pallets</strong></span><br>
    <span style="color:orange;">TOTAL METRIC TON: ${invoice.total_weight} KG</span><br>
    <span style="color:purple;">1) This order is ready for pick up at Partners Warehouse The warehouse address is: 21051 Walter Strawn Dr, Elwood, IL 60421.
   Please use the DISPATCH# when communicating with Partners warehouse.
   The warehouse email address is:  beansnutrition.csr@partnerswarehouse.com
   And please cc’ us at this address:   dispatch@.com

   Please make an appointment with them at least 24 hours in advance. Appointment time for pick up is from 7:00am - 2:30pm. Their break time is from 9:00am-9:30am. Their lunch break is from 12:00pm-12:30pm. 
   If the driver misses the confirmed appointment time by more than 1.5 hours, a new appointment will need to be made.

    Once the warehouse confirms your appointment, you are all set. If your 
    driver is lost, please call (815) 423-9100.</span><br>
    <span style="color:teal;">2) Please deliver ON ${invoice.delivery_date} to this customer's address: ${invoice.delivery_address}</span><br>
    <span style="color:navy;">3) ${invoice.instruction}</span><br>
    THE TRUCK MUST BE 53’ DRY VAN AND IN GOOD CONDITION<br>
    The shipping instruction will have the net weight of cargo only (excluding pallets and packaging). THE DRIVER MUST WEIGH THEIR EMPTY TRUCKS BEFORE arriving at the WAREHOUSE FOR LOADING. The driver shall determine prior to departing if the cargo will be overweight based on the net weight provided by the warehouse. <br>
    
    4) Fees of $25.00 will incur from the warehouse if the following happens:<br>
      a) failure to make an appointment<br>
      b) same day appointments<br>
      c) 1 day earlier or later than the appointment made will be considered as no appointment<br>
      d) late pick up (missing the specified deadline date)<br>
    Therefore, please follow the instructions in order to avoid these fees.<br>
    5) Fees of $150.00/day will be incurred from the customer if you miss the customer delivery date.<br>
    Thank you.<br>
    Most Sincerely,Eason<br>
  </pre>
</div>
`;
//=================================================================================================================================
      // delivery_instruction = `
      //   <br><br>
      //   ${delivery_instruction}
      //   <br>
      //   <span style="color:red;">${invoice.po_num}</span><br>
      //   <span style="color:blue;">${invoice.invoice_no_}</span><br>
      //   <span style="color:yellow;">${invoice.total_pallets}</span><br>
      //   <span style="color:green;">${invoice.total_weight}</span><br>
      //   <span style="color:pink;">${invoice.delivery_date}</span><br>
      //   <span style="color:purple;">${invoice.delivery_address}</span><br>
      //   <span style="color:orange;">${invoice.truck_cost}</span><br>
      //   <span style="color:black;">${invoice.instruction}</span><br>
      // `;


      let email_body = "Dear " + `${freight_company.freight_company}` + delivery_instruction;


      // let recipient = freight_company.contact_email;
       
      // let cc = 'dispatch@.com';

      let recipient = 'eason@test.com';
      let cc = 'eason@test.com';
      GmailApp.sendEmail(recipient, EMAIL_TRUCK_SUBJECT + ` ${invoice.invoice_no_}`, email_body, {
        //attachments: [attachment.getAs(MimeType.PDF)],
        cc: cc,
        name: EMAIL_TRUCK_SUBJECT + ` ${invoice.invoice_no_}`,
        htmlBody: email_body
      });

      invoicesSheet.getRange(index + 2, 14).setValue('Yes');
    }
  });
}


function sendEmailToWareHouse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();
  const truckSheet = ss.getSheetByName(TRUCK_SHEET_NAME);
  const truckData = truckSheet.getRange(1, 1, truckSheet.getLastRow(), truckSheet.getLastColumn()).getValues();
  const warehouseSheet = ss.getSheetByName(WAREHOUSE_SHEET_NAME);
  const warehouseData = warehouseSheet.getRange(1, 1, warehouseSheet.getLastRow(), warehouseSheet.getLastColumn()).getValues();

  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));

  const keysT = truckData.splice(0, 1)[0];
  const trucks = getObjects(truckData, createObjectKeys(keysT));

  const keysW = warehouseData.splice(0, 1)[0];
  const warehouses = getObjects(warehouseData, createObjectKeys(keysW));

  let freight_company = "";
  let warehouse = "";

  console.log('Emailing WareHouse ' + EMAIL_WAREHOUSE_SUBJECT);
  invoices.forEach(function (invoice, index) {

    if (invoice.approval_email_sent != 'Yes') {
      ss.toast(`Emailing WareHouse for ${invoice.customer}`, APP_TITLE, 1);

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
      const attachment = DriveApp.getFileById(fileId);


      trucks.forEach(function (each_truck) {
        if (invoice.truck == each_truck.freight_company) {
          freight_company = each_truck;
        }
      })

      warehouses.forEach(function (each_warehouse) {
        if (invoice.warehouse == each_warehouse.warehouse_name) {
          warehouse = each_warehouse;
 

        }
      })



 
      let email_body = "Hi " + `${warehouse.warehouse_name}` + ':\n' + '\n' +
        'Please see the attachment and prepare for us.\n' +
        'Thank you.\n' +

        '\n' +
        'Best Regards, \n' +
        '\n' +
        'Eason,\n' + 
        'xxx LLC';

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      // when send to Wollenweber need to cc more people 
      // chad@wollenweb.com
      // chris@wollenweb.com
      
      if (warehouse.warehouse_name == "Wollenweber"){
        cc = ["chad@wollenweb.com",'chris@wollenweb.com','dispatch@xxx.com']; 
      }


      let recipient = warehouse.email;
      let cc = 'dispatch@xxx.com';


      // let recipient = 'pauline@xxx.com';
      // let cc = 'eason@xxx.com';
 

      if (attachment.getMimeType().includes("pdf")) {
        GmailApp.sendEmail(recipient, EMAIL_WAREHOUSE_SUBJECT + ` ${invoice.invoice_no_}`, email_body, {
          attachments: [attachment.getAs(MimeType.PDF)],
          cc: cc,
          name: EMAIL_TRUCK_SUBJECT + ` ${invoice.invoice_no_}`
        });
      }
      else {
        console.log("Sending xlsx");
        GmailApp.sendEmail(recipient, EMAIL_WAREHOUSE_SUBJECT + ` ${invoice.invoice_no_}`, email_body, {
          attachments: [attachment.getAs(MimeType.PDF)],
          cc: cc,
          name: EMAIL_TRUCK_SUBJECT + ` ${invoice.invoice_no_}`
        });
      }
 

      invoicesSheet.getRange(index + 2, 13).setValue('Yes');
    }
  });
}


/**
 * Helper function that turns sheet data range into an object. 
 * 
 * @param {SpreadsheetApp.Sheet} sheet - Sheet to process
 * Return {object} of a sheet's datarange as an object 
 */
function dataRangeToObject(sheet) {
  const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const keys = dataRange.splice(0, 1)[0];
  return getObjects(dataRange, createObjectKeys(keys));
}

/**
 * Utility function for mapping sheet data to objects.
 */
function getObjects(data, keys) {
  let objects = [];
  for (let i = 0; i < data.length; ++i) {
    let object = {};
    let hasData = false;
    for (let j = 0; j < data[i].length; ++j) {
      let cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}
// Creates object keys for column headers.
function createObjectKeys(keys) {
  return keys.map(function (key) {
    return key.replace(/\W+/g, '_').toLowerCase();
  });
}
// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}