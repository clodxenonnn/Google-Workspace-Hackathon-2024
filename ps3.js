// Standardizing policy documentation, Problem Statement 3 Code


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Pdf');
  menu.addItem('Create New Pdf', 'createPolicyDocuments');
  menu.addToUi();
}

function createPolicyDocuments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  // IDs of the template document and target folder
  var templateId = '1uby1HArzPrnLjLi0HdnfyubZxc1KBtIvvItRaAu88w4'; // Replace with your template document ID
  var folderId = '131UFqR_d4tnm2bP2g49DDftBiWxyIlqU'; // Replace with your target folder ID

  // Get the header row
  var headers = data[0];
  
  // Find the index of the relevant columns
  var nameIndex = headers.indexOf('Full Name');
  var emailIndex = headers.indexOf('Email Address');
  var linkIndex = headers.indexOf('Document Link');
  
  // Iterate through rows to create policy documents
  for (var i = 1; i < data.length; i++) {
    var fullName = data[i][nameIndex];
    var emailAddress = data[i][emailIndex];
    
    if (!fullName || !emailAddress) {
      continue; // Skip rows with missing data
    }
    
    // Create a new policy document and get its PDF file
    var { docUrl, pdfFile } = createPolicyDocument(templateId, folderId, fullName, emailAddress);
    
    // Send the PDF as an email attachment
    sendEmailWithPdf(emailAddress, fullName, pdfFile);
    
    // Update the Google Sheet with the document link
    sheet.getRange(i + 1, linkIndex + 1).setValue(docUrl);
  }
}

function createPolicyDocument(templateId, folderId, fullName, emailAddress) {
  // Get the template document and create a copy
  var templateDoc = DriveApp.getFileById(templateId);
  var newDoc = templateDoc.makeCopy(fullName + ' Policy Document', DriveApp.getFolderById(folderId));
  var docId = newDoc.getId();
  
  // Open the new document
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  // Replace placeholders with actual content
  body.replaceText('{{Full Name}}', fullName);
  body.replaceText('{{Email Address}}', emailAddress);
  body.replaceText('{{Today Date}}', new Date().toLocaleDateString());
  
  // Save and close the document
  doc.saveAndClose();

  // Convert the Google Document to a PDF blob
  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs('application/pdf');
  
  // Retrieve the folder where the PDF will be saved
  var folder = DriveApp.getFolderById(folderId);
  
  // Create a new PDF file in the specified folder
  var pdfFile = folder.createFile(pdfBlob);
  
  // Return the URL of the new PDF file and the PDF file itself
  return {
    docUrl: pdfFile.getUrl(),
    pdfFile: pdfFile
  };
}

function sendEmailWithPdf(emailAddress, fullName, pdfFile) {
  var subject = 'Your Policy Document';
  var body = 'Dear ' + fullName + ',\n\nPlease find attached your policy document.\n\nBest regards,\nYour Company';
  
  // Send the email with the PDF attachment
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs('application/pdf')]
  });
}
