
# **Lunch and Learn: Automating Google Sheets with Apps Script**
**Generate and Email PDF Reports**

---

## **Objective**
In this session, we will learn how to use Google Apps Script to automate tasks within Google Sheets, focusing on:
- Accessing Google Sheets data programmatically.
- Sending emails using custom subject and body.
- Generating a PDF link from a Google Sheet and sending it via email.

By the end of the session, you’ll be able to apply these concepts to automate reporting, invoicing, or any task involving Google Sheets and email notifications.

---

## **Outline / Table of Contents**

1. **Step 1**: Setting Up Variables and Logging Data from Google Sheets
2. **Step 2**: Writing a Function to Send a Simple Email
3. **Step 3**: Generating a PDF Link from a Google Sheet and Including It in an Email
4. **Wrap-Up**: Summary and Discussion

---

## **Step 1: Setting Up Variables and Logging Data from Google Sheets**

**Objective**: Set up custom variables (such as sheet ID, email, etc.), get a sheet object, and use `Logger.log()` to print these values and data from the sheet to the Apps Script console.

### **Instructions**
1. Open the Google Apps Script editor by selecting **Extensions > Apps Script** in your Google Sheet.
2. Replace the default code with the following skeleton code.
3. Set your **sheet ID** and **sheet name** to match your Google Sheet.
4. Use the **Logger.log()** method to log the sheet name and a value from a specific cell (A1).

### **Code for Step 1**

```javascript
function main() {
  // Custom variables
  var sheetId = 'YOUR_SHEET_ID';  // <-- Add your Google Sheet ID here
  var sheetName = 'Sheet1';        // <-- Specify the sheet name to export as PDF
  var recipientEmail = 'recipient@example.com';  // <-- Enter the recipient's email address
  var emailSubject = 'Your Report';              // <-- Customize the email subject
  var emailBody = 'Attached is your PDF report.'; // <-- Customize the email body

  // Get the sheet object
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  // Log some details from the sheet
  Logger.log('Sheet Name: ' + sheet.getName()); // Log the sheet's name
  Logger.log('First Cell Value: ' + sheet.getRange(1, 1).getValue()); // Log the value in cell A1

  // Print other variables to the Apps Script console
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

## **Step 3: Generating a PDF Link from a Google Sheet and Including It in an Email**

**Objective**: Generate a PDF link from the specified Google Sheet and include that link in the email body.

### **Instructions**
1. Generate the PDF link from the Google Sheet.
2. Include the link in the email body and send it to the recipient.
3. This method avoids needing to fetch a `Blob` object, working around network restrictions.

### **Code for Step 3**

```javascript
function main() {
  // Custom variables
  var sheetId = 'YOUR_SHEET_ID';  
  var sheetName = 'Sheet1';        
  var recipientEmail = 'recipient@example.com';  
  var emailSubject = 'Your Report';              
  var emailBody = 'Click the link below to download your PDF report.

'; 

  // Generate the PDF link from the Google Sheet
  var pdfLink = generatePdfLink(sheetId, sheetName);

  // Append the link to the email body
  emailBody += pdfLink;

  // Send the email with the PDF link instead of the attachment
  sendEmailWithLink(recipientEmail, emailSubject, emailBody);

  Logger.log('Email sent with PDF link');
}

// Function to generate the PDF export link
function generatePdfLink(sheetId, sheetName) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Sheet with name ' + sheetName + ' not found!');
  }

  // Build the URL to export the Google Sheet as a PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + sheet.getParent().getId() +
            '/export?exportFormat=pdf&format=pdf' +
            '&gid=' + sheet.getSheetId() +  
            '&size=A4' +  
            '&portrait=true' +  
            '&fitw=true' +  
            '&sheetnames=false&printtitle=false&pagenumbers=false' +
            '&gridlines=false&fzr=false';  

  return url;  // Return the link to the PDF
}

// Function to send an email with a link
function sendEmailWithLink(recipient, subject, body) {
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body
  });
}
```

---

## **Wrap-Up: Summary and Discussion**

- **What We Learned**: 
  - How to set up and use custom variables.
  - How to log data from Google Sheets using Apps Script.
  - How to send a simple email using Google Apps Script.
  - How to generate a PDF link from a Google Sheet and include it in an email.

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
