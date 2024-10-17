
# **Lunch and Learn: Automating Google Sheets with Apps Script**
**Generate and Email PDF Reports**

---

## **Objective**
In this session, we will learn how to use Google Apps Script to automate tasks within Google Sheets, focusing on:
- Accessing Google Sheets data programmatically.
- Sending emails using custom subject and body.
- Generating a PDF from a Google Sheet.
- Sending that PDF as an email attachment.

By the end of the session, you’ll be able to apply these concepts to automate reporting, invoicing, or any task involving Google Sheets and email notifications.

---

## **Outline / Table of Contents**

1. **Step 1**: Setting Up Variables and Printing to Console
2. **Step 2**: Writing a Function to Send a Simple Email
3. **Step 3**: Generating a PDF from a Google Sheet and Saving it to Google Drive
4. **Step 4**: Sending the Generated PDF as an Email Attachment
5. **Wrap-Up**: Summary and Discussion

---

## **Step 1: Setting Up Variables and Printing to Console**

**Objective**: Set up custom variables (such as sheet ID, email, etc.) and use `Logger.log()` to print these values to the Apps Script console.

### **Instructions**
1. Open the Google Apps Script editor by selecting **Extensions > Apps Script** in your Google Sheet.
2. Replace the default code with the following skeleton code.
3. Set your **sheet ID** and **sheet name** to match your Google Sheet.
4. Use the **Logger.log()** method to print the values to the console and confirm they are correct.

### **Code for Step 1**

```javascript
function main() {
  var sheetId = 'YOUR_SHEET_ID';  // <-- Add your Google Sheet ID here
  var sheetName = 'Sheet1';        // <-- Specify the sheet name to export as PDF
  var recipientEmail = 'recipient@example.com';  // <-- Enter the recipient's email address
  var emailSubject = 'Your Report';              // <-- Customize the email subject
  var emailBody = 'Attached is your PDF report.'; // <-- Customize the email body

  Logger.log('Sheet ID: ' + sheetId);
  Logger.log('Sheet Name: ' + sheetName);
  Logger.log('Recipient Email: ' + recipientEmail);
  Logger.log('Email Subject: ' + emailSubject);
  Logger.log('Email Body: ' + emailBody);
}
```

---

## **Step 2: Writing a Function to Send a Simple Email**

**Objective**: Write a function that sends a simple email using the variables set in Step 1 (recipient, subject, body).

### **Instructions**
1. Use the email variables from Step 1 to send a basic email with the subject and body.
2. The function `sendSimpleEmail()` will use `MailApp.sendEmail()` to send the email to the recipient.
3. Run the script to send the email to yourself.

### **Code for Step 2**

```javascript
function main() {
  var recipientEmail = 'recipient@example.com';  // <-- Add your email address here
  var emailSubject = 'Your Report';              // <-- Customize the email subject
  var emailBody = 'Here’s your report, but no attachment yet!';  // <-- Customize the body

  sendSimpleEmail(recipientEmail, emailSubject, emailBody);
}

function sendSimpleEmail(recipient, subject, body) {
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body
  });

  Logger.log('Email sent to ' + recipient);
}
```

---

## **Step 3: Generating a PDF from a Google Sheet and Saving It to Google Drive**

**Objective**: Learn how to generate a PDF from the specified Google Sheet and save it to Google Drive.

### **Instructions**
1. Replace your folder ID in the code to save the generated PDF to a specific Google Drive folder.
2. The `generatePdf()` function will generate the PDF using the sheet you specify in the variables.
3. Run the script and verify that the PDF is saved in your Google Drive.

### **Code for Step 3**

```javascript
function main() {
  var sheetId = 'YOUR_SHEET_ID';  
  var sheetName = 'Sheet1';        

  var pdfBlob = generatePdf(sheetId, sheetName);

  var folder = DriveApp.getFolderById('YOUR_FOLDER_ID');  // <-- Add your folder ID
  folder.createFile(pdfBlob);

  Logger.log('PDF saved to Google Drive');
}

function generatePdf(sheetId, sheetName) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet with name ' + sheetName + ' not found!');
  }

  var url = 'https://docs.google.com/spreadsheets/d/' + sheet.getParent().getId() +
            '/export?exportFormat=pdf&format=pdf' +
            '&gid=' + sheet.getSheetId() +  
            '&size=A4' +  
            '&portrait=true' +  
            '&fitw=true' +  
            '&sheetnames=false&printtitle=false&pagenumbers=false' +
            '&gridlines=false&fzr=false';  

  var pdfBlob = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }).getBlob().setName(sheet.getName() + '.pdf');

  return pdfBlob;
}
```

---

## **Step 4: Sending the Generated PDF as an Email Attachment**

**Objective**: Combine the previous steps to generate a PDF and attach it to an email before sending it.

### **Instructions**
1. Combine the `generatePdf()` function with the `sendEmailWithAttachment()` function.
2. After generating the PDF, attach it to the email and send it to the specified recipient.

### **Code for Step 4**

```javascript
function main() {
  var sheetId = 'YOUR_SHEET_ID';  
  var sheetName = 'Sheet1';        
  var recipientEmail = 'recipient@example.com';  
  var emailSubject = 'Your Report';              
  var emailBody = 'Attached is your PDF report.'; 

  var pdfBlob = generatePdf(sheetId, sheetName);

  sendEmailWithAttachment(pdfBlob, recipientEmail, emailSubject, emailBody);

  Logger.log('Email sent with PDF attachment');
}

function generatePdf(sheetId, sheetName) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet with name ' + sheetName + ' not found!');
  }

  var url = 'https://docs.google.com/spreadsheets/d/' + sheet.getParent().getId() +
            '/export?exportFormat=pdf&format=pdf' +
            '&gid=' + sheet.getSheetId() +  
            '&size=A4' +  
            '&portrait=true' +  
            '&fitw=true' +  
            '&sheetnames=false&printtitle=false&pagenumbers=false' +
            '&gridlines=false&fzr=false';  

  var pdfBlob = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }).getBlob().setName(sheet.getName() + '.pdf');

  return pdfBlob;
}

function sendEmailWithAttachment(pdfBlob, recipient, subject, body) {
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    attachments: [pdfBlob]
  });
}
```

---

## **Wrap-Up: Summary and Discussion**

- **What We Learned**: 
  - How to set up and use custom variables.
  - How to send a simple email using Google Apps Script.
  - How to generate a PDF from a Google Sheet.
  - How to send that PDF as an email attachment.
  
- **Real-World Applications**:
  - Automating weekly or monthly reports.
  - Sending invoices directly from a Google Sheet.
  - Generating PDFs for task tracking, project management, or performance reports.

---

## **Explore More Google Apps Script Use Cases**

Here are a few other Google Apps Script use cases to explore:
1. **Automate Task Reminders**: Use Apps Script to send reminder emails based on deadlines in a Google Sheet.
2. **Track Form Submissions**: Automate the processing of Google Form responses and send custom responses.
3. **Create a Dashboard**: Build an automated Google Sheet dashboard that pulls data from various sources and updates in real-time.
4. **Custom Invoice Generation**: Automatically generate and send invoices based on entries in Google Sheets.
5. **Daily/Weekly Data Exports**: Export data from Google Sheets to external services or create daily/weekly reports.

---

Feel free to experiment with these ideas and take your Google Apps Script knowledge further!
