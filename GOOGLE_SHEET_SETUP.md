# Google Sheet Integration Setup Guide

Follow these quick steps (takes ~2 minutes) to receive all client orders from your portfolio directly in your Google Sheet in real time.

---

### Step 1: Create a Google Sheet
1. Open [Google Sheets](https://sheets.new) and create a new blank spreadsheet.
2. Name it: **"Zakaria Dou - Portfolio Orders"**.
3. In the first row (headers), you can add:
   - **A1**: `Timestamp`
   - **B1**: `Client Name`
   - **C1**: `Email`
   - **D1**: `Phone / WhatsApp`
   - **E1**: `Timeline`
   - **F1**: `Services`
   - **G1**: `Budget`
   - **H1**: `Project Details`

---

### Step 2: Open Google Apps Script
1. In your Google Sheet menu bar, click on **Extensions** > **Apps Script**.
2. Delete any existing code in the editor and paste the following Google Apps Script:

```javascript
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Automatically name/find sheet tab "Client Orders"
    var sheet = doc.getSheetByName("Client Orders");
    if (!sheet) {
      sheet = doc.getActiveSheet();
      sheet.setName("Client Orders");
    }
    
    // Parse incoming data
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }
    
    var timestamp = data.timestamp || new Date().toLocaleString();
    var name = data.name || "Anonymous Client";
    var email = data.email || "No Email Provided";
    var phone = data.phone || "N/A";
    var timeline = data.timeline || "Flexible";
    var services = Array.isArray(data.services) ? data.services.join(", ") : (data.services || "General Inquiry");
    var budget = data.budget || "Flexible / To discuss";
    var details = data.brief || data.details || "No details provided";
    
    // Add header row automatically if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Client Name",
        "Email",
        "Phone / WhatsApp",
        "Timeline",
        "Services",
        "Budget",
        "Project Details"
      ]);
      
      var headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#15181a");
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
    
    // 2. Append order row into the Google Sheet
    sheet.appendRow([
      timestamp,
      name,
      email,
      phone,
      timeline,
      services,
      budget,
      details
    ]);
    
    // 3. Send instant email notification to both your email addresses
    var targetEmail = "douzakaria98@gmail.com, douzakaria1998@gmail.com";
    var subject = "🚀 New Client Order: " + name + " (" + services + ")";
    
    var cleanPhone = phone.replace(/[^0-9+]/g, '');
    var waLink = cleanPhone ? "https://wa.me/" + cleanPhone.replace(/^\+/, '') : "";
    
    var htmlBody = `
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 620px; margin: 0 auto; background: #fafaf7; border: 1px solid #d5d9d1; border-radius: 12px; overflow: hidden; color: #15181a;">
        <div style="background: #15181a; color: #ffffff; padding: 22px 26px;">
          <h2 style="margin: 0; font-size: 20px; font-weight: 600;">Zakaria Dou · Portfolio</h2>
          <p style="margin: 6px 0 0; font-size: 14px; opacity: 0.85;">You received a new client project request!</p>
        </div>
        
        <div style="padding: 24px 26px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 15px;">
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; width: 140px; font-weight: 500;">Client Name:</td>
              <td style="padding: 10px 0; font-weight: 600; font-size: 16px;">${name}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; font-weight: 500;">Email:</td>
              <td style="padding: 10px 0;"><a href="mailto:${email}" style="color: #2a38f0; font-weight: 500; text-decoration: none;">${email}</a></td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; font-weight: 500;">Phone / WhatsApp:</td>
              <td style="padding: 10px 0;">
                <span style="font-weight: 500;">${phone}</span>
                ${waLink ? ` &nbsp;·&nbsp; <a href="${waLink}" style="color: #10b981; font-weight: 600; text-decoration: none;" target="_blank">🟢 Open WhatsApp</a>` : ''}
              </td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; font-weight: 500;">Budget:</td>
              <td style="padding: 10px 0; font-weight: 600; color: #2a38f0;">${budget}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; font-weight: 500;">Timeline:</td>
              <td style="padding: 10px 0; font-weight: 500;">${timeline}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
              <td style="padding: 10px 0; color: #596060; font-weight: 500;">Services:</td>
              <td style="padding: 10px 0; font-weight: 500;">${services}</td>
            </tr>
          </table>
          
          <div style="margin-top: 18px;">
            <p style="margin: 0 0 6px; color: #596060; font-weight: 500; font-size: 14px;">Project Details:</p>
            <div style="background: #f1f2ee; border: 1px solid #d5d9d1; border-radius: 8px; padding: 14px; font-size: 14px; line-height: 1.6; white-space: pre-wrap;">${details}</div>
          </div>
          
          <div style="margin-top: 22px; padding-top: 18px; border-top: 1px solid #e5e7eb;">
            <a href="mailto:${email}" style="display: inline-block; background: #15181a; color: #ffffff; padding: 9px 16px; border-radius: 6px; text-decoration: none; font-size: 13.5px; font-weight: 600;">✉️ Reply to ${name}</a>
            ${waLink ? `<a href="${waLink}" style="display: inline-block; background: #10b981; color: #ffffff; padding: 9px 16px; border-radius: 6px; text-decoration: none; font-size: 13.5px; font-weight: 600; margin-left: 8px;">🟢 Chat on WhatsApp</a>` : ''}
          </div>
        </div>
        
        <div style="background: #eeeee9; padding: 10px 24px; font-size: 12px; color: #596060; text-align: center;">
          Received on ${timestamp} via your Portfolio Website
        </div>
      </div>
    `;
    
    try {
      MailApp.sendEmail({
        to: targetEmail,
        replyTo: email,
        subject: subject,
        htmlBody: htmlBody
      });
    } catch (mailErr) {
      console.warn("Mail dispatch notice: " + mailErr.toString());
    }
    
    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "active", service: "Zakaria Dou Orders Webhook" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this once in Apps Script to authorize email sending:
function testEmail() {
  MailApp.sendEmail("douzakaria98@gmail.com", "🎉 Portfolio Notifications Active", "Your portfolio email notification system is now authorized and fully active!");
}
```

---

### Step 3: Deploy as Web App
1. In the top right of the Apps Script window, click the blue **Deploy** button > **New deployment**.
2. Beside "Select type", click the gear icon ⚙️ and choose **Web app**.
3. Fill in:
   - **Description**: `Portfolio Order Webhook`
   - **Execute as**: **Me (`douzakaria1998@gmail.com`)**
   - **Who has access**: **Anyone** *(Important: must be "Anyone" so visitors to your website can submit without logging into Google)*
4. Click **Deploy**.
5. Google may prompt you for permission:
   - Click **Authorize access**.
   - Choose your Google account.
   - Click **Advanced** > **Go to Untitled project (unsafe)** (Google shows this warning for personal scripts).
   - Click **Allow**.
6. Copy the **Web App URL** provided (it looks like `https://script.google.com/macros/s/AKfycb.../exec`).

---

### Step 4: Paste into your `index.html`
Open `index.html` and look for line ~690:
```javascript
const GOOGLE_SHEET_URL = 'https://script.google.com/macros/s/AKfycbz_YOUR_WEB_APP_ID/exec';
```
Replace `'https://script.google.com/macros/s/AKfycbz_YOUR_WEB_APP_ID/exec'` with your actual Web App URL, save, and push to Git!

---

### Fallback & Safety Guarantee
- If a client submits when the script URL isn't configured or if there is a network interruption, their order details are **always safely backed up in their browser's local storage** (`zakaria_orders`).
- All submissions are verified on the front-end before sending.
