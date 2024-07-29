Streamline and Optimizing Post-Sale Workflow


 
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Pdf');
  menu.addItem('Create New Pdf', 'createOrderForm');
  menu.addToUi();
}

function createOrderForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  // IDs of the template document and target folder
  var templateId = '1jMgpUal8quv-tsny5zRFNgz2yUMvPYKAOQilE6JMeAw'; // Replace with your template document ID
  var folderId = '1c7u7YJnszOfB-9ijTaq82ob6-9nfXjkd'; // Replace with your target folder ID

  // Get the header row
  var headers = data[0];
  
  // Find the index of the relevant columns
  var timeIndex = headers.indexOf('Timestamp');
  var nameIndex = headers.indexOf('Your Name');
  var emailIndex = headers.indexOf('Email address');
  var shipIndex = headers.indexOf('Shipping Address');
  var quantityIndex = headers.indexOf('Quantity');
  var colorIndex = headers.indexOf('Color (White/Black)');
  var sizeIndex = headers.indexOf('Size (S, M, L, XL, 2XL, 3XL) ');
  var linkIndex = headers.indexOf('Document Link');
  
  // Add "Document Link" column if it does not exist
  if (linkIndex === -1) {
    linkIndex = headers.length;
    sheet.getRange(1, linkIndex + 1).setValue('Document Link');
  }
  
  // Iterate through rows to create order documents
  for (var i = 1; i < data.length; i++) {
    var yourName = data[i][nameIndex];
    var timeStamp = data[i][timeIndex];
    var emailAddress = data[i][emailIndex];
    var shipAddress = data[i][shipIndex];
    var quantityShirt = data[i][quantityIndex];
    var colorShirt = data[i][colorIndex];
    var sizeShirt = data[i][sizeIndex];

    if (!yourName || !timeStamp || !emailAddress || !shipAddress || !quantityShirt || !colorShirt || !sizeShirt) {
      continue; // Skip rows with missing data
    }
    
    // Create a new order document and get the PDF URL
    var pdfUrl = createOrderDocument(templateId, folderId, yourName, timeStamp, emailAddress, shipAddress, quantityShirt, colorShirt, sizeShirt);
    
    // Update the Google Sheet with the document link
    sheet.getRange(i + 1, linkIndex + 1).setValue(pdfUrl);
    
    // Send an email with the PDF attachment
    var pdfFile = DriveApp.getFileById(pdfUrl.split('/d/')[1].split('/')[0]); // Extract the file ID from the URL
    sendEmailWithPdf(emailAddress, yourName, pdfFile);
  }
}

function createOrderDocument(templateId, folderId, yourName, timeStamp, emailAddress, shipAddress, quantityShirt, colorShirt, sizeShirt) {
  // Get the template document and create a copy
  var templateDoc = DriveApp.getFileById(templateId);
  var newDoc = templateDoc.makeCopy(yourName + ' Order Form (Responses)', DriveApp.getFolderById(folderId));
  var docId = newDoc.getId();
  
  // Open the new document
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  // Replace placeholders with actual content
  body.replaceText('{{Your Name}}', yourName);
  body.replaceText('{{email}}', emailAddress);
  body.replaceText('{{address}}', shipAddress);
  body.replaceText('{{Timestamp}}', timeStamp);
  body.replaceText('{{quantity}}', quantityShirt);
  body.replaceText('{{color}}', colorShirt);
  body.replaceText('{{size}}', sizeShirt);

  // Save and close the document
  doc.saveAndClose();
  
  // Convert the Google Document to a PDF blob
  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs('application/pdf');
  
  // Retrieve the folder where the PDF will be saved
  var folder = DriveApp.getFolderById(folderId);
  
  // Create a new PDF file in the specified folder
  var pdfFile = folder.createFile(pdfBlob);
  
  // Return the URL of the new PDF file
  return pdfFile.getUrl();
}

function sendEmailWithPdf(emailAddress, fullName, pdfFile) {
  var subject = 'Your Order Confirmation';
  var body = 'Dear ' + fullName + ',\n\nThank you for your purchase. IF you find any problem, contact us.\n\nBest regards,\nDHYS Corporation';
  
  // Send the email with the PDF attachment
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs('application/pdf')]
  });
}

function testCreateOrderForm() {
  createOrderForm();
}



