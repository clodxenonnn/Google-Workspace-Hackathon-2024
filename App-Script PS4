//Compliance with company policies


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Pdf');
  menu.addItem('Create New Pdf', 'createPolicyForm');
  menu.addToUi();
}

function createPolicyForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  // IDs of the template document and target folder
  var templateId = '16BPmDW4iOZv1Q97DYi6ltiTm3ZFwOEX824TDB2dP2RU'; // Replace with your template document ID
  var folderId = '1Bk7D0b0AMPAFlhtAKNDyMGGdXVXpCMyu'; // Replace with your target folder ID

// Get the header row
  var headers = data[0];
  
  // Find the index of the relevant columns
  var numberIndex = headers.indexOf('Employee ID');
  var nameIndex = headers.indexOf('Employee Name');
  var departmentIndex = headers.indexOf('departmentPosition');
  var linkIndex = headers.indexOf('Document Link');
  
  // Iterate through rows to create policy documents
  for (var i = 1; i < data.length; i++) {
    var numberEmployee = data[i][numberIndex];
    var employeeName = data[i][nameIndex];
    var departmentPosition = data[i][departmentIndex];
    
    if ( !numberEmployee || !employeeName || !departmentPosition ) {
      continue; // Skip rows with missing data
    }
    
    // Create a new policy document
    var docUrl = createPolicyDocument(templateId, folderId, numberEmployee, employeeName, departmentPosition);
    
    // Update the Google Sheet with the document link
    sheet.getRange(i + 1, linkIndex + 1).setValue(docUrl);
  }
}

function createPolicyDocument(templateId, folderId, numberEmployee, employeeName, departmentPosition) {
  // Get the template document and create a copy
  var templateDoc = DriveApp.getFileById(templateId);
  var newDoc = templateDoc.makeCopy(employeeName + ' Compliance Policy Form', DriveApp.getFolderById(folderId));
  var docId = newDoc.getId();
  
  // Open the new document
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  // Replace placeholders with actual content
  body.replaceText('{{Employee ID}}',numberEmployee);
  body.replaceText('{{Employee Name}}',employeeName);
  body.replaceText('{{departmentPosition}}', departmentPosition);
  body.replaceText('{{Date}}', new Date().toLocaleDateString());
  
  // Save and close the document
  doc.saveAndClose();
  
  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs('application/pdf');
  
  // Retrieve the folder where the PDF will be saved
  var folder = DriveApp.getFolderById(folderId);
  
  // Create a new PDF file in the specified folder
  var pdfFile = folder.createFile(pdfBlob);
  
  // Return the URL of the new PDF file
  return pdfFile.getUrl();
}

function testCreatePolicyForm() {
  createPolicyForm();
}



