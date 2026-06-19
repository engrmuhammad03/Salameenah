// ==================== SALAMEENAH CONCEPT - MODERN ORDER SYSTEM ====================
// 
// This script handles:
// 1. Saving orders to Google Sheets
// 2. Sending beautiful branded email to customer
// 3. Sending professional notification to store owner
//
// ================================================

// ===== CONFIGURATION - UPDATE THESE VALUES =====
const SHEET_ID = "1DfFkcMfhxJrnw5-W3YKnXmDRJr9BcLGojjjTcpq0TS0";           // Your Google Sheet ID
const STORE_EMAIL = "ummusalmamuhammad32@gmail.com";      // Your store email
const STORE_NAME = "Salameenah Concept";                 // Your store name
const COMPANY_LOGO_URL = "";                             // OPTIONAL: Add your logo URL (leave empty for text logo)
const COMPANY_WHATSAPP = "+2348065512725";                // Your WhatsApp number
const COMPANY_PHONE = "+2348065512725";                     // Your phone number
const BANK_NAME = "Monie Point";
const BANK_ACCOUNT_NAME = "Salameenah Concept";
const BANK_ACCOUNT_NUMBER = "5721810092";

// ===== DO NOT EDIT BELOW THIS LINE =====

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.requestType === "order") {
      return handleOrderSubmission(data);
    } else {
      return createResponse(false, null, "Unknown request type");
    }
  } catch(error) {
    return createResponse(false, null, error.toString());
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ 
    success: true, 
    message: "Salameenah Concept Order System is running",
    version: "2.0"
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleOrderSubmission(data) {
  const sheet = getOrCreateSheet();
  const orderId = generateOrderId();
  const timestamp = new Date();
  
  sheet.appendRow([
    timestamp,
    orderId,
    data.customerName,
    data.customerPhone,
    data.customerEmail,
    data.deliveryAddress,
    data.paymentMethod,
    data.items,
    data.totalAmount,
    data.orderNotes || "",
    "Pending",
    data.pageSource || ""
  ]);
  
  sendCustomerOrderEmail(data, orderId);
  sendAdminNotificationEmail(data, orderId);
  
  return createResponse(true, { orderId: orderId }, null);
}

function getOrCreateSheet() {
  let sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Orders");
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SHEET_ID).insertSheet("Orders");
    sheet.appendRow([
      "Timestamp", "Order ID", "Customer Name", "Phone", "Email", 
      "Delivery Address", "Payment Method", "Items", "Total Amount", 
      "Order Notes", "Status", "Page Source"
    ]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground("#c8a45c").setFontColor("#1a1a1a");
    sheet.setColumnWidths(1, 12, 150);
  }
  return sheet;
}

function generateOrderId() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const random = Math.floor(Math.random() * 9999).toString().padStart(4, '0');
  return `SAL-${year}${month}${day}-${random}`;
}

function createResponse(success, data, error) {
  return ContentService.createTextOutput(JSON.stringify({ 
    success: success, 
    data: data, 
    error: error 
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==================== CUSTOMER ORDER EMAIL ====================

function sendCustomerOrderEmail(data, orderId) {
  const itemsHtml = formatItemsForEmail(data.items);
  const orderDate = new Date().toLocaleDateString('en-NG', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
  
  const logoHtml = COMPANY_LOGO_URL ? 
    `<img src="${COMPANY_LOGO_URL}" alt="${STORE_NAME}" style="height: 60px; width: auto; margin-bottom: 10px;">` : 
    `<div style="font-size: 28px; font-weight: 800; letter-spacing: 2px;">S<span style="color: #c8a45c;">C</span></div>`;
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Order Confirmation - ${STORE_NAME}</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
          font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif; 
          line-height: 1.5; 
          background: #f0ebe3;
          padding: 20px;
        }
        .email-container { 
          max-width: 600px; 
          margin: 0 auto; 
          background: #ffffff; 
          border-radius: 24px; 
          overflow: hidden;
          box-shadow: 0 20px 40px rgba(0,0,0,0.08);
        }
        .email-header { 
          background: #1a1a1a;
          padding: 30px 25px 25px; 
          text-align: center; 
          border-bottom: 3px solid #c8a45c;
        }
        .logo-area { margin-bottom: 12px; }
        .logo-text { 
          font-size: 28px; 
          font-weight: 800; 
          color: #c8a45c; 
          letter-spacing: 2px;
        }
        .logo-sub { 
          font-size: 10px; 
          color: #888; 
          letter-spacing: 3px;
          margin-top: 4px;
        }
        .order-badge {
          display: inline-block;
          background: rgba(200, 164, 92, 0.15);
          padding: 6px 18px;
          border-radius: 30px;
          margin-top: 12px;
        }
        .order-badge span {
          color: #c8a45c;
          font-weight: 600;
          font-size: 12px;
          letter-spacing: 1px;
        }
        .email-body { padding: 30px 25px; }
        .greeting { 
          font-size: 22px; 
          font-weight: 700; 
          color: #1a1a1a;
          margin-bottom: 6px;
        }
        .greeting-sub {
          color: #666;
          font-size: 14px;
          margin-bottom: 20px;
          padding-bottom: 15px;
          border-bottom: 2px solid #f0ebe3;
        }
        
        .order-id-card {
          background: #f8f6f3;
          border-radius: 16px;
          padding: 15px 20px;
          margin: 20px 0;
          text-align: center;
          border: 1px solid #e8e0d6;
        }
        .order-id-label {
          font-size: 10px;
          text-transform: uppercase;
          letter-spacing: 2px;
          color: #999;
        }
        .order-id-value {
          font-size: 18px;
          font-weight: 700;
          color: #c8a45c;
          font-family: monospace;
          letter-spacing: 1px;
          margin-top: 5px;
        }
        .order-date {
          font-size: 11px;
          color: #aaa;
          margin-top: 6px;
        }
        
        .section-title {
          font-size: 16px;
          font-weight: 700;
          color: #1a1a1a;
          margin: 25px 0 12px 0;
          padding-left: 10px;
          border-left: 3px solid #c8a45c;
        }
        
        /* INFO TABLE - REDUCED WIDTH WITH OVERFLOW PROTECTION */
        .info-table {
          width: 100%;
          border-collapse: collapse;
          background: #faf7f2;
          border-radius: 16px;
          overflow: hidden;
          margin: 15px 0;
          table-layout: fixed;
        }
        .info-table td {
          padding: 10px 12px;
          border-bottom: 1px solid #e8e0d6;
          vertical-align: top;
          word-wrap: break-word;
          word-break: break-word;
          overflow-wrap: break-word;
        }
        .info-table tr:last-child td {
          border-bottom: none;
        }
        .info-label {
          width: 85px;
          font-weight: 700;
          color: #a0784c;
          font-size: 12px;
        }
        .info-value {
          color: #1a1a1a;
          font-size: 13px;
          word-wrap: break-word;
          word-break: break-word;
          overflow-wrap: break-word;
        }
        
        /* ITEMS TABLE */
        .items-table {
          width: 100%;
          border-collapse: collapse;
          background: #faf7f2;
          border-radius: 16px;
          overflow: hidden;
          margin: 15px 0;
          table-layout: fixed;
        }
        .items-table th {
          background: #e8e0d6;
          padding: 10px 12px;
          text-align: left;
          font-size: 11px;
          font-weight: 800;
          color: #5a4a2a;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .items-table td {
          padding: 10px 12px;
          border-bottom: 1px solid #e8e0d6;
          font-size: 13px;
          vertical-align: top;
          word-wrap: break-word;
          word-break: break-word;
          overflow-wrap: break-word;
        }
        .items-table tr:last-child td {
          border-bottom: none;
        }
        .items-table .item-name {
          font-weight: 600;
          color: #1a1a1a;
        }
        .items-table .item-qty {
          text-align: center;
          width: 60px;
        }
        .items-table .item-price {
          text-align: right;
          width: 100px;
        }
        .total-row {
          background: #fff8e7;
        }
        .total-row td {
          font-weight: 800;
          font-size: 14px;
          color: #a0784c;
          border-top: 2px solid #c8a45c;
        }
        
        /* PAYMENT BOX */
        .payment-box {
          background: #fff8e7;
          border: 2px solid #c8a45c;
          border-radius: 16px;
          padding: 20px;
          margin: 20px 0;
        }
        .payment-box h4 {
          color: #a0784c;
          margin-bottom: 15px;
          font-size: 15px;
          font-weight: 700;
        }
        .bank-table {
          width: 100%;
          border-collapse: collapse;
          background: white;
          border-radius: 12px;
          overflow: hidden;
          margin: 10px 0;
          table-layout: fixed;
        }
        .bank-table td {
          padding: 10px 12px;
          border-bottom: 1px solid #e8e0d6;
          font-size: 13px;
          word-wrap: break-word;
          word-break: break-word;
        }
        .bank-table tr:last-child td {
          border-bottom: none;
        }
        .bank-label {
          font-weight: 700;
          color: #555;
          width: 110px;
        }
        .bank-value {
          color: #1a1a1a;
          font-family: monospace;
          font-weight: 600;
        }
        .copy-btn {
          background: #c8a45c;
          color: #1a1a1a;
          border: none;
          padding: 4px 12px;
          border-radius: 20px;
          font-size: 11px;
          font-weight: 600;
          cursor: pointer;
          margin-left: 10px;
        }
        .whatsapp-button {
          display: inline-block;
          background: #25D366;
          color: white;
          padding: 12px 24px;
          border-radius: 40px;
          text-decoration: none;
          font-weight: 700;
          font-size: 14px;
          margin-top: 15px;
        }
        .help-box {
          background: #f0ebe3;
          border-radius: 16px;
          padding: 15px;
          text-align: center;
          margin-top: 25px;
        }
        .help-box p {
          margin: 0;
          font-size: 12px;
          color: #666;
        }
        .help-box a {
          color: #25D366;
          font-weight: 600;
          text-decoration: none;
        }
        
        /* FOOTER */
        .footer {
          background: #1a1a1a;
          color: #888;
          padding: 25px;
          text-align: center;
          font-size: 11px;
        }
        .footer a {
          color: #c8a45c;
          text-decoration: none;
        }
        .social-icons {
          margin-bottom: 12px;
        }
        .social-icons a {
          display: inline-block;
          margin: 0 6px;
          color: #c8a45c;
          font-size: 16px;
          text-decoration: none;
        }
        
        /* MOBILE RESPONSIVE */
        @media (max-width: 500px) {
          .email-body { padding: 25px 20px; }
          .info-label { width: 75px; }
          .bank-label { width: 90px; }
          .bank-table td { padding: 8px 10px; }
          .items-table th, .items-table td { padding: 8px 10px; }
          .info-table td { padding: 8px 10px; }
          .copy-btn { margin-left: 5px; padding: 3px 8px; }
        }
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="email-header">
          <div class="logo-area">
            ${logoHtml}
          </div>
          <div class="logo-sub">PREMIUM MODEST WEAR</div>
          <div class="order-badge">
            <span>ORDER CONFIRMED</span>
          </div>
        </div>
        
        <div class="email-body">
          <div class="greeting">Thank you, ${data.customerName.split(' ')[0]}!</div>
          <div class="greeting-sub">We are delighted to confirm your order</div>
          
          <div class="order-id-card">
            <div class="order-id-label">YOUR ORDER NUMBER</div>
            <div class="order-id-value">${orderId}</div>
            <div class="order-date">Placed on ${orderDate}</div>
          </div>
          
          <div class="section-title">Order Summary</div>
          <table class="items-table" cellpadding="0" cellspacing="0" width="100%">
            <thead>
              <tr>
                <th align="left">Item</th>
                <th align="center" width="60">Qty</th>
                <th align="right" width="100">Total</th>
              </tr>
            </thead>
            <tbody>
              ${itemsHtml}
              <tr class="total-row">
                <td align="right" colspan="2" style="font-weight: 700;">GRAND TOTAL</td>
                <td align="right" style="font-weight: 800; color: #a0784c;">${data.totalAmount}</td>
              </tr>
            </tbody>
          </table>
          
          <div class="section-title">Delivery Details</div>
          <table class="info-table" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td class="info-label">Full Name</td>
              <td class="info-value">${data.customerName}</td>
            </tr>
            <tr>
              <td class="info-label">Phone Number</td>
              <td class="info-value">${data.customerPhone}</td>
            </tr>
            <tr>
              <td class="info-label">Email Address</td>
              <td class="info-value">${data.customerEmail}</td>
            </tr>
            <tr>
              <td class="info-label">Delivery Address</td>
              <td class="info-value">${data.deliveryAddress}</td>
            </tr>
            <tr>
              <td class="info-label">Payment Method</td>
              <td class="info-value">${data.paymentMethod === 'Bank Transfer' ? 'Bank Transfer' : 'Cash on Delivery'}</td>
            </tr>
          </table>
          
          ${data.paymentMethod === 'Bank Transfer' ? `
            <div class="payment-box">
              <h4>Bank Transfer Information</h4>
              <table class="bank-table" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td class="bank-label">Bank Name:</td>
                  <td class="bank-value">${BANK_NAME}</td>
                </tr>
                <tr>
                  <td class="bank-label">Account Name:</td>
                  <td class="bank-value">${BANK_ACCOUNT_NAME}</td>
                </tr>
                <tr>
                  <td class="bank-label">Account Number:</td>
                  <td class="bank-value">
                    ${BANK_ACCOUNT_NUMBER}
                    <button class="copy-btn" onclick="copyToClipboard('${BANK_ACCOUNT_NUMBER}')">Copy</button>
                  </td>
                </tr>
                <tr>
                  <td class="bank-label">Amount to Pay:</td>
                  <td class="bank-value" style="font-weight: 800; color: #a0784c;">${data.totalAmount}</td>
                </tr>
              </table>
              <p style="font-size: 12px; margin-top: 12px;"><strong>Next Step:</strong> After payment, click the button below to send your proof of payment.</p>
              <div style="text-align: center;">
                <a href="https://wa.me/${COMPANY_WHATSAPP}?text=I%20have%20made%20payment%20for%20Order%20${orderId}" class="whatsapp-button">
                  Send Payment Proof on WhatsApp
                </a>
              </div>
            </div>
          ` : `
            <div class="payment-box">
              <h4>Cash on Delivery</h4>
              <table class="bank-table" cellpadding="0" cellspacing="0" width="100%">
                <tr>
                  <td class="bank-label">Amount to Pay:</td>
                  <td class="bank-value" style="font-weight: 800; color: #a0784c;">${data.totalAmount}</td>
                </tr>
              </table>
              <p style="font-size: 12px; margin-top: 10px;">You will pay when your order is delivered to your address.</p>
              <p style="font-size: 12px;">Our delivery team will contact you within few minutes to schedule delivery.</p>
            </div>
          `}
          
          ${data.orderNotes ? `
            <div class="section-title">Additional Notes</div>
            <div style="background: #faf7f2; border-radius: 12px; padding: 12px 15px; margin: 10px 0;">
              <p style="margin: 0; color: #555; font-style: italic; font-size: 13px;">"${data.orderNotes}"</p>
            </div>
          ` : ''}
          
          <div class="help-box">
            <p>Need help? Contact us on 
              <a href="https://wa.me/${COMPANY_WHATSAPP}">WhatsApp</a> 
              or call <strong>${COMPANY_PHONE}</strong>
            </p>
          </div>
        </div>
        
        <div class="footer">
          <div class="social-icons">
            <a href="https://www.instagram.com/salameenah.concept/" target="_blank">Instagram</a>
            <a href="https://wa.me/${COMPANY_WHATSAPP}">WhatsApp</a>
          </div>
          <p style="margin-bottom: 6px;">${STORE_NAME} - Elegance Redefined</p>
          <p>${STORE_EMAIL} | ${COMPANY_PHONE}</p>
          <p style="margin-top: 10px;">&copy; ${new Date().getFullYear()} ${STORE_NAME}. All rights reserved.</p>
        </div>
      </div>
      
      <script>
        function copyToClipboard(text) {
          var input = document.createElement('textarea');
          input.value = text;
          document.body.appendChild(input);
          input.select();
          document.execCommand('copy');
          document.body.removeChild(input);
          alert('Account number copied!');
        }
      </script>
    </body>
    </html>
  `;
  
  const plainText = `${STORE_NAME} - Order Confirmation #${orderId}\n\nThank you ${data.customerName}!\nTotal: ${data.totalAmount}\nPayment: ${data.paymentMethod}\n\nWe will process your order shortly.`;
  
  try {
    GmailApp.sendEmail(data.customerEmail, `Your Order Confirmation #${orderId} - ${STORE_NAME}`, plainText, { htmlBody: htmlBody });
  } catch(e) {
    console.error("Customer email failed: " + e);
  }
}

// ==================== ADMIN NOTIFICATION EMAIL (FIXED) ====================

function sendAdminNotificationEmail(data, orderId) {
  const itemsHtml = formatItemsForAdmin(data.items);
  const orderDate = new Date().toLocaleString('en-NG');
  
  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>New Order - ${STORE_NAME}</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
          font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif; 
          background: #f0ebe3;
          padding: 20px;
        }
        .email-container { 
          max-width: 650px; 
          margin: 0 auto; 
          background: #ffffff; 
          border-radius: 24px; 
          overflow: hidden;
          box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        .email-header { 
          background: #c8a45c;
          padding: 25px 25px;
          text-align: center;
        }
        .header-title { 
          font-size: 24px; 
          font-weight: 800; 
          color: #1a1a1a;
          letter-spacing: 1px;
        }
        .header-sub { 
          font-size: 12px; 
          color: #1a1a1a; 
          opacity: 0.8;
          margin-top: 4px;
        }
        .email-body { padding: 30px 25px; }
        
        .alert-box {
          background: #fff8e7;
          border-left: 4px solid #c8a45c;
          border-radius: 12px;
          padding: 15px 20px;
          margin-bottom: 25px;
        }
        .alert-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
        }
        .alert-label {
          font-weight: 700;
          color: #a0784c;
          font-size: 13px;
        }
        .order-id {
          font-size: 14px;
          font-weight: 800;
          color: #c8a45c;
          font-family: monospace;
          letter-spacing: 1px;
          background: #1a1a1a;
          padding: 2px 10px;
          border-radius: 30px;
        }
        .order-time {
          font-size: 11px;
          color: #999;
          margin-top: 8px;
        }
        
        .section-title {
          font-size: 15px;
          font-weight: 800;
          color: #1a1a1a;
          margin: 20px 0 12px 0;
          padding-left: 10px;
          border-left: 3px solid #c8a45c;
        }
        
        /* INFO TABLE - REDUCED WIDTH WITH OVERFLOW PROTECTION */
        .info-table {
          width: 100%;
          border-collapse: collapse;
          background: #f8f6f3;
          border-radius: 12px;
          overflow: hidden;
          table-layout: fixed;
        }
        .info-table td {
          padding: 10px 12px;
          border-bottom: 1px solid #e8e0d6;
          vertical-align: top;
          word-wrap: break-word;
          word-break: break-word;
          overflow-wrap: break-word;
        }
        .info-table tr:last-child td {
          border-bottom: none;
        }
        .info-label {
          width: 85px;
          font-weight: 700;
          color: #a0784c;
          font-size: 12px;
        }
        .info-value {
          color: #1a1a1a;
          font-size: 13px;
          word-wrap: break-word;
          word-break: break-word;
          overflow-wrap: break-word;
        }
        .info-value a {
          color: #c8a45c;
          text-decoration: none;
        }
        
        /* ITEMS TABLE */
        .items-table {
          width: 100%;
          border-collapse: collapse;
          background: #f8f6f3;
          border-radius: 12px;
          overflow: hidden;
          table-layout: fixed;
        }
        .items-table th {
          background: #e8e0d6;
          padding: 10px 12px;
          text-align: left;
          font-size: 11px;
          font-weight: 800;
          color: #5a4a2a;
        }
        .items-table td {
          padding: 10px 12px;
          border-bottom: 1px solid #e8e0d6;
          font-size: 13px;
          word-wrap: break-word;
          word-break: break-word;
        }
        .items-table tr:last-child td {
          border-bottom: none;
        }
        .total-row td {
          font-weight: 800;
          color: #a0784c;
          border-top: 2px solid #c8a45c;
          background: #fff8e7;
        }
        
        /* BUTTONS */
        .button-container {
          display: flex;
          gap: 10px;
          margin: 25px 0;
          flex-wrap: wrap;
          justify-content: center;
        }
        .btn-sheets {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          background: #1a1a1a;
          color: #c8a45c;
          padding: 8px 18px;
          border-radius: 50px;
          text-decoration: none;
          font-weight: 700;
          font-size: 11px;
          transition: all 0.2s ease;
          border: 1px solid #c8a45c;
        }
        .btn-sheets:hover {
          background: #c8a45c;
          color: #1a1a1a;
          transform: translateY(-2px);
        }
        .btn-whatsapp {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          background: #25D366;
          color: white;
          padding: 8px 18px;
          border-radius: 50px;
          text-decoration: none;
          font-weight: 700;
          font-size: 11px;
          transition: all 0.2s ease;
        }
        .btn-whatsapp:hover {
          background: #128C7E;
          transform: translateY(-2px);
        }
        
        .footer-note {
          background: #f0ebe3;
          border-radius: 12px;
          padding: 12px 15px;
          text-align: center;
          font-size: 11px;
          color: #666;
          margin-top: 20px;
        }
        
        /* MOBILE RESPONSIVE */
        @media (max-width: 500px) {
          .email-body { padding: 25px 20px; }
          .alert-row { flex-direction: column; gap: 8px; align-items: flex-start; }
          .button-container { flex-direction: column; }
          .btn-sheets, .btn-whatsapp { width: 100%; text-align: center; }
          .info-label { width: 75px; }
          .items-table th, .items-table td { padding: 8px 10px; }
          .info-table td { padding: 8px 10px; }
        }
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="email-header">
          <div class="header-title">NEW ORDER RECEIVED</div>
          <div class="header-sub">${STORE_NAME}</div>
        </div>
        
        <div class="email-body">
          <div class="alert-box">
            <div class="alert-row">
              <span class="alert-label">Action Required</span>
              <span class="order-id">${orderId}</span>
            </div>
            <div class="order-time">Order placed on ${orderDate}</div>
          </div>
          
          <div class="section-title">Customer Information</div>
          <table class="info-table" cellpadding="0" cellspacing="0" width="100%">
            <tr>
              <td class="info-label">Full Name</td>
              <td class="info-value">${data.customerName}</td>
            </tr>
            <tr>
              <td class="info-label">Phone Number</td>
              <td class="info-value"><a href="tel:${data.customerPhone}">${data.customerPhone}</a></td>
            </tr>
            <tr>
              <td class="info-label">Email</td>
              <td class="info-value"><a href="mailto:${data.customerEmail}">${data.customerEmail}</a></td>
            </tr>
            <tr>
              <td class="info-label">Delivery Address</td>
              <td class="info-value">${data.deliveryAddress}</td>
            </tr>
            <tr>
              <td class="info-label">Payment Method</td>
              <td class="info-value">${data.paymentMethod === 'Bank Transfer' ? 'Bank Transfer' : 'Cash on Delivery'}</td>
            </tr>
          </table>
          
          <div class="section-title">Items Ordered</div>
          <table class="items-table" cellpadding="0" cellspacing="0" width="100%">
            <thead>
              <tr>
                <th>Item</th>
                <th align="center" width="60">Qty</th>
                <th align="right" width="100">Total</th>
              </tr>
            </thead>
            <tbody>
              ${itemsHtml}
              <tr class="total-row">
                <td colspan="2" align="right" style="font-weight: 700;">GRAND TOTAL</td>
                <td align="right" style="font-weight: 800;">${data.totalAmount}</td>
              </tr>
            </tbody>
          </table>
          
          ${data.orderNotes ? `
            <div class="section-title">Customer Notes</div>
            <div style="background: #f8f6f3; border-radius: 12px; padding: 12px 15px;">
              <p style="margin: 0; color: #555; font-style: italic; font-size: 13px;">${data.orderNotes}</p>
            </div>
          ` : ''}
          
          <div class="button-container">
            <a href="https://docs.google.com/spreadsheets/d/${SHEET_ID}" target="_blank" class="btn-sheets">
              View in Google Sheets
            </a>
            <a href="https://wa.me/${data.customerPhone}?text=Hello%20${encodeURIComponent(data.customerName)}%2C%20your%20order%20${orderId}%20has%20been%20received%20and%20is%20being%20processed." target="_blank" class="btn-whatsapp">
              WhatsApp Customer
            </a>
          </div>
          
          <div class="footer-note">
            Update the "Status" column in Google Sheets to "Processing" or "Shipped" when you begin working on this order.
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
  
  const plainText = `NEW ORDER #${orderId}\n\nCustomer: ${data.customerName}\nPhone: ${data.customerPhone}\nTotal: ${data.totalAmount}\nPayment: ${data.paymentMethod}\n\nView in Sheets: https://docs.google.com/spreadsheets/d/${SHEET_ID}`;
  
  try {
    GmailApp.sendEmail(STORE_EMAIL, `New Order #${orderId} - ${STORE_NAME}`, plainText, { htmlBody: htmlBody });
  } catch(e) {
    console.error("Admin email failed: " + e);
  }
}

function formatItemsForEmail(itemsText) {
  const lines = itemsText.split('\n');
  let html = '';
  for (const line of lines) {
    if (line.trim()) {
      const match = line.match(/(.+?) \(Qty:(\d+)\) - ₦([\d,]+)/);
      if (match) {
        html += `
          <tr>
            <td class="item-name">${match[1]}</td>
            <td align="center" class="item-qty">${match[2]}</td>
            <td align="right" class="item-price">₦${match[3]}</td>
          </tr>
        `;
      } else {
        html += `<tr><td colspan="3">${line}</td></tr>`;
      }
    }
  }
  return html;
}

function formatItemsForAdmin(itemsText) {
  const lines = itemsText.split('\n');
  let html = '';
  for (const line of lines) {
    if (line.trim()) {
      const match = line.match(/(.+?) \(Qty:(\d+)\) - ₦([\d,]+)/);
      if (match) {
        html += `
          <tr>
            <td><strong>${match[1]}</strong></td>
            <td align="center">${match[2]}</td>
            <td align="right">₦${match[3]}</td>
          </tr>
        `;
      } else {
        html += `<tr><td colspan="3">${line}</td></tr>`;
      }
    }
  }
  return html;
}