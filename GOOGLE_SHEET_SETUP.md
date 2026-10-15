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
  lock.tryLock(10000); // Prevent concurrent write collisions
  
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();
    
    // Parse incoming data (handles both JSON and form payloads)
    var data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter || {};
    }
    
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
    
    // Append the new client order row
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString(),
      data.name || "",
      data.email || "",
      data.phone || "",
      data.timeline || "",
      data.services || "",
      data.budget || "",
      data.details || ""
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
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
