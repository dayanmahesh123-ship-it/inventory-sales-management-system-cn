// ============================================================
// MAHESH WEB APP SOLUTION — Management System Backend
// Demo Business: Liyanage Electronics
// Version: 6.0 — Complete Polished Release with Add/Edit/Delete
// ============================================================

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const ADMIN_EMAIL = 'admin@liyanageelectronics.lk'; // ⬅️ මේක ඔයාගේ email එකට change කරන්න

// ============================================================
// EMAIL ALERT HELPER — අඩු තොග email ඇඟවීම්
// ============================================================
function sendLowStockAlert(productId, productName, currentQty, minStock) {
  try {
    const subject = '⚠️ Low Stock Alert — ' + productName + ' [' + productId + ']';
    const htmlBody = 
      '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">' +
        '<div style="background:linear-gradient(135deg,#f59e0b,#ef4444);padding:20px;text-align:center">' +
          '<h2 style="color:#fff;margin:0;font-size:18px">⚠️ Low Stock Alert</h2>' +
          '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:12px">Liyanage Electronics — Inventory System</p>' +
        '</div>' +
        '<div style="padding:24px">' +
          '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
            '<tr><td style="padding:8px 0;color:#64748b">Product ID</td><td style="padding:8px 0;font-weight:700;text-align:right">' + productId + '</td></tr>' +
            '<tr><td style="padding:8px 0;color:#64748b">Product Name</td><td style="padding:8px 0;font-weight:700;text-align:right">' + productName + '</td></tr>' +
            '<tr style="border-top:1px solid #f1f5f9"><td style="padding:8px 0;color:#64748b">Current Stock</td><td style="padding:8px 0;font-weight:700;color:#ef4444;text-align:right">' + currentQty + ' units</td></tr>' +
            '<tr><td style="padding:8px 0;color:#64748b">Min Stock Level</td><td style="padding:8px 0;font-weight:700;text-align:right">' + minStock + ' units</td></tr>' +
          '</table>' +
          '<div style="margin-top:20px;padding:12px;background:#fef2f2;border-radius:8px;border-left:4px solid #ef4444">' +
            '<p style="margin:0;font-size:13px;color:#991b1b"><strong>Action Required:</strong> Please reorder this item immediately to avoid stockouts.</p>' +
          '</div>' +
        '</div>' +
        '<div style="padding:12px 24px;background:#f8fafc;text-align:center;font-size:11px;color:#94a3b8">' +
          'Powered by Mahesh Web App Solution | Automated Alert' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Liyanage Electronics System'
    });

    Logger.log('Low stock alert sent for: ' + productName + ' (Qty: ' + currentQty + ')');
    return true;
  } catch (err) {
    Logger.log('Email alert failed: ' + err.toString());
    return false;
  }
}

function checkAndAlertLowStock(productId, productName, newQty, minStock) {
  if (newQty <= minStock && newQty >= 0) {
    sendLowStockAlert(productId, productName, newQty, minStock);
  }
}

// ============================================================
// SHEET SETUP — Sheets හදන්න
// ============================================================
function setupSheets() {
  const sheetsConfig = {
    Products: {
      headers: ['ProductID','ProductName','Category','Brand','Model','UnitPrice','CostPrice','StockQty','MinStockLevel','WarrantyMonths','Supplier','DateAdded','Status'],
      data: [
        ['MOB001','Samsung Galaxy S24 Ultra','Mobile Phones','Samsung','S24 Ultra',189900,152000,15,5,12,'Samsung Sri Lanka',new Date('2024-12-01'),'Active'],
        ['MOB002','iPhone 15 Pro Max','Mobile Phones','Apple','15 Pro Max',249900,210000,8,3,12,'Apple Authorized',new Date('2024-12-05'),'Active'],
        ['MOB003','Samsung Galaxy A15','Mobile Phones','Samsung','A15',42900,33000,25,10,12,'Samsung Sri Lanka',new Date('2024-12-10'),'Active'],
        ['MOB004','Xiaomi Redmi Note 13','Mobile Phones','Xiaomi','Redmi Note 13',52900,40000,20,8,12,'Xiaomi Distributors',new Date('2025-01-02'),'Active'],
        ['APP001','LG 55" OLED Smart TV','Appliances','LG','OLED55C3',329900,275000,4,2,24,'LG Electronics Lanka',new Date('2024-11-15'),'Active'],
        ['APP002','Samsung 65" Crystal UHD TV','Appliances','Samsung','CU7000',219900,178000,6,2,24,'Samsung Sri Lanka',new Date('2024-11-20'),'Active'],
        ['APP003','LG 10kg Front Load Washer','Appliances','LG','FV1410S5W',159900,128000,5,2,36,'LG Electronics Lanka',new Date('2025-01-05'),'Active'],
        ['LAP001','Dell Inspiron 15','Laptops','Dell','Inspiron 3520',174900,142000,10,3,12,'Dell Technologies',new Date('2024-12-20'),'Active'],
        ['LAP002','HP Pavilion x360','Laptops','HP','Pavilion x360 14',189900,155000,7,3,12,'HP Lanka',new Date('2025-01-10'),'Active'],
        ['LAP003','Lenovo IdeaPad Slim 3','Laptops','Lenovo','IdeaPad Slim 3',134900,108000,12,4,12,'Lenovo Distributors',new Date('2025-01-12'),'Active'],
        ['ACC001','Samsung Galaxy Buds FE','Accessories','Samsung','Buds FE',18900,14000,30,10,6,'Samsung Sri Lanka',new Date('2025-01-15'),'Active'],
        ['ACC002','Anker PowerCore 20000mAh','Accessories','Anker','PowerCore 20K',8900,5800,40,15,18,'Anker Distributors',new Date('2025-01-18'),'Active'],
        ['ACC003','Logitech MX Master 3S','Accessories','Logitech','MX Master 3S',21900,16500,3,5,24,'Logitech Lanka',new Date('2025-01-20'),'Low Stock'],
        ['APP004','Philips Air Fryer XXL','Appliances','Philips','HD9270',49900,38000,2,5,24,'Philips Lanka',new Date('2025-02-01'),'Low Stock']
      ]
    },
    Sales: {
      headers: ['SaleID','Date','ProductID','ProductName','Category','Qty','UnitPrice','DiscountPct','TotalAmount','CustomerName','CustomerPhone','PaymentMethod','SoldBy','ReturnStatus'],
      data: [
        ['S0001',new Date('2025-05-20 09:30'),'MOB001','Samsung Galaxy S24 Ultra','Mobile Phones',1,189900,0,189900,'Kamal Perera','0771234567','Card','Nimal',''],
        ['S0001',new Date('2025-05-20 09:30'),'ACC002','Anker PowerCore 20000mAh','Accessories',2,8900,0,17800,'Kamal Perera','0771234567','Card','Nimal',''],
        ['S0002',new Date('2025-05-20 11:15'),'APP001','LG 55" OLED Smart TV','Appliances',1,329900,5,313405,'Saman Silva','0712345678','Card','Sunil',''],
        ['S0003',new Date('2025-05-21 10:00'),'MOB003','Samsung Galaxy A15','Mobile Phones',2,42900,0,85800,'Ruwan Fernando','0761122334','Cash','Nimal','Partial Return'],
        ['S0004',new Date('2025-05-21 14:30'),'LAP001','Dell Inspiron 15','Laptops',1,174900,0,174900,'Dilshan Jayawardena','0779988776','Card','Sunil',''],
        ['S0005',new Date('2025-05-22 09:00'),'MOB002','iPhone 15 Pro Max','Mobile Phones',1,249900,3,242403,'Nadeesha Kumari','0714455667','Card','Nimal',''],
        ['S0005',new Date('2025-05-22 09:00'),'ACC001','Samsung Galaxy Buds FE','Accessories',1,18900,3,18333,'Nadeesha Kumari','0714455667','Card','Nimal',''],
        ['S0006',new Date('2025-05-22 15:45'),'APP002','Samsung 65" Crystal UHD TV','Appliances',1,219900,0,219900,'Chaminda Bandara','0723344556','Cash','Sunil',''],
        ['S0007',new Date('2025-05-23 10:20'),'LAP003','Lenovo IdeaPad Slim 3','Laptops',1,134900,0,134900,'Amaya Ratnayake','0775566778','Cash','Nimal',''],
        ['S0008',new Date('2025-05-23 13:00'),'MOB004','Xiaomi Redmi Note 13','Mobile Phones',1,52900,0,52900,'Pradeep Wijesinghe','0769900112','Cash','Sunil','Returned'],
        ['S0009',new Date('2025-05-24 11:30'),'LAP002','HP Pavilion x360','Laptops',1,189900,0,189900,'Sachini De Silva','0711223344','Card','Nimal',''],
        ['S0010',new Date('2025-05-25 12:30'),'ACC002','Anker PowerCore 20000mAh','Accessories',3,8900,0,26700,'Thilina Gamage','0776677889','Cash','Sunil','']
      ]
    },
    Returns: {
      headers: ['ReturnID','Date','SaleID','ProductID','ProductName','Qty','Reason','RefundAmount','ProcessedBy','Status'],
      data: [
        ['R0001',new Date('2025-05-22'),'S0003','MOB003','Samsung Galaxy A15',1,'Defective unit — screen flickering',42900,'Nimal','Completed'],
        ['R0002',new Date('2025-05-24'),'S0008','MOB004','Xiaomi Redmi Note 13',1,'Customer changed mind within 7 days',52900,'Sunil','Completed']
      ]
    },
    RestockLog: {
      headers: ['RestockID','Date','ProductID','ProductName','Qty','Supplier','UnitCost','TotalCost','ReceivedBy','Notes'],
      data: [
        ['RS001',new Date('2025-05-15'),'MOB001','Samsung Galaxy S24 Ultra',10,'Samsung Sri Lanka',152000,1520000,'Nimal','Monthly restock'],
        ['RS002',new Date('2025-05-15'),'ACC002','Anker PowerCore 20000mAh',20,'Anker Distributors',5800,116000,'Sunil','High demand item'],
        ['RS003',new Date('2025-05-18'),'LAP001','Dell Inspiron 15',5,'Dell Technologies',142000,710000,'Nimal','New batch arrival']
      ]
    }
  };

  for (const [sheetName, config] of Object.entries(sheetsConfig)) {
    let sheet = SPREADSHEET.getSheetByName(sheetName);
    if (sheet) { sheet.clear(); } else { sheet = SPREADSHEET.insertSheet(sheetName); }
    sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    if (config.data.length > 0) {
      sheet.getRange(2, 1, config.data.length, config.headers.length).setValues(config.data);
    }
    sheet.getRange(1, 1, 1, config.headers.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, config.headers.length);
  }

  SpreadsheetApp.getUi().alert(
    '✅ සියලුම Sheets 4ම සාර්ථකව හැදුවා!\n\n' +
    '• Products: ' + sheetsConfig.Products.data.length + ' items\n' +
    '• Sales: ' + sheetsConfig.Sales.data.length + ' rows\n' +
    '• Returns: ' + sheetsConfig.Returns.data.length + ' records\n' +
    '• RestockLog: ' + sheetsConfig.RestockLog.data.length + ' records\n\n' +
    '📧 Low stock alerts යනවා: ' + ADMIN_EMAIL
  );
}

// ============================================================
// DATA RETRIEVAL — Data ගන්න
// ============================================================
function getSheetData(sheetName) {
  const sheet = SPREADSHEET.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return rows.map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      obj[h] = (row[i] instanceof Date) ? row[i].toISOString() : row[i];
    });
    return obj;
  });
}

function getAllData() {
  return {
    products: getSheetData('Products'),
    sales: getSheetData('Sales'),
    returns: getSheetData('Returns'),
    restockLog: getSheetData('RestockLog'),
    _timestamp: new Date().toISOString()
  };
}

// ============================================================
// GET ENDPOINT — Data request කරන්න
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  var result;
  
  switch (action) {
    case 'getProducts':
      result = { success: true, data: getSheetData('Products') };
      break;
    case 'getSales':
      result = { success: true, data: getSheetData('Sales') };
      break;
    case 'getReturns':
      result = { success: true, data: getSheetData('Returns') };
      break;
    case 'getRestockLog':
      result = { success: true, data: getSheetData('RestockLog') };
      break;
    default:
      result = { success: true, data: getAllData() };
      break;
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// POST ENDPOINT — Data write කරන්න
// ============================================================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // ╔═══════════════════════════════════════╗
    // ║        ADD PRODUCT — භාණ්ඩ එකතු කරන්න     ║
    // ╚═══════════════════════════════════════╝
    if (action === 'addProduct') {
      const p = payload.data;
      const sheet = SPREADSHEET.getSheetByName('Products');
      
      // Required fields check
      if (!p.productId || !p.productName || !p.category || !p.brand) {
        return _jr({ success: false, error: 'අවශ්‍ය fields හිස්ව ඇත. Product ID, Name, Category, Brand අවශ්‍යයි.' });
      }
      
      // Check duplicate Product ID
      const existing = getSheetData('Products');
      if (existing.find(function(x) { return x.ProductID === p.productId; })) {
        return _jr({ success: false, error: 'Product ID "' + p.productId + '" දැනටමත් තිබේ. වෙනත් ID එකක් යොදන්න.' });
      }
      
      // Check duplicate Product Name
      var duplicateName = existing.find(function(x) { 
        return x.ProductName.toLowerCase().trim() === String(p.productName).toLowerCase().trim(); 
      });
      if (duplicateName) {
        return _jr({ success: false, error: 'මෙම නමින් භාණ්ඩයක් දැනටමත් තිබේ: "' + duplicateName.ProductName + '" (' + duplicateName.ProductID + ')' });
      }
      
      var uP = Number(p.unitPrice) || 0;
      var cP = Number(p.costPrice) || 0;
      var sQ = Number(p.stockQty) || 0;
      var mS = Number(p.minStockLevel) || 5;
      var wM = Number(p.warrantyMonths) || 0;
      var status = sQ <= mS ? 'Low Stock' : 'Active';
      
      // Validate prices
      if (uP <= 0) {
        return _jr({ success: false, error: 'විකුණුම් මිල Rs. 0 ට වඩා වැඩි විය යුතුයි.' });
      }
      if (cP <= 0) {
        return _jr({ success: false, error: 'මිලදී ගත් මිල Rs. 0 ට වඩා වැඩි විය යුතුයි.' });
      }
      if (uP < cP) {
        // Warning but still allow — just log
        Logger.log('WARNING: Selling price (' + uP + ') is less than cost price (' + cP + ') for ' + p.productName);
      }

      sheet.appendRow([
        p.productId,
        p.productName,
        p.category,
        p.brand,
        p.model || '',
        uP,
        cP,
        sQ,
        mS,
        wM,
        p.supplier || '',
        new Date(),
        status
      ]);
      sheet.autoResizeColumns(1, sheet.getLastColumn());

      // Low stock alert check
      if (status === 'Low Stock') {
        checkAndAlertLowStock(p.productId, p.productName, sQ, mS);
      }

      return _jr({
        success: true,
        message: '✅ භාණ්ඩය සාර්ථකව එකතු කරන ලදී! Product: ' + p.productName,
        productId: p.productId,
        status: status
      });
    }

    // ╔═══════════════════════════════════════╗
    // ║     UPDATE PRODUCT — භාණ්ඩ edit කරන්න    ║
    // ╚═══════════════════════════════════════╝
    if (action === 'updateProduct') {
      const lock = LockService.getScriptLock();
      try { lock.waitLock(15000); } catch (lk) { 
        return _jr({ success: false, error: 'System busy. නැවත උත්සාහ කරන්න.' }); 
      }
      
      try {
        const p = payload.data;
        
        if (!p.productId) {
          return _jr({ success: false, error: 'Product ID අවශ්‍යයි.' });
        }
        
        const sheet = SPREADSHEET.getSheetByName('Products');
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        
        // Find the product row
        var rowIndex = -1;
        for (var i = 0; i < data.length; i++) {
          if (String(data[i][0]).trim() === String(p.productId).trim()) {
            rowIndex = i;
            break;
          }
        }
        
        if (rowIndex === -1) {
          return _jr({ success: false, error: 'Product "' + p.productId + '" හමු නොවීය.' });
        }
        
        var actualRow = rowIndex + 2; // +1 for header, +1 for 0-index
        var updatedFields = [];
        
        // Update ProductName
        if (p.productName !== undefined && p.productName !== '') {
          var nameIdx = headers.indexOf('ProductName');
          if (nameIdx !== -1) {
            // Check duplicate name (exclude current product)
            var existingProducts = getSheetData('Products');
            var dupName = existingProducts.find(function(x) { 
              return x.ProductID !== p.productId && x.ProductName.toLowerCase().trim() === String(p.productName).toLowerCase().trim(); 
            });
            if (dupName) {
              return _jr({ success: false, error: 'මෙම නමින් වෙනත් භාණ්ඩයක් තිබේ: "' + dupName.ProductName + '"' });
            }
            sheet.getRange(actualRow, nameIdx + 1).setValue(p.productName);
            updatedFields.push('Name');
          }
        }
        
        // Update Category
        if (p.category !== undefined && p.category !== '') {
          var catIdx = headers.indexOf('Category');
          if (catIdx !== -1) {
            sheet.getRange(actualRow, catIdx + 1).setValue(p.category);
            updatedFields.push('Category');
          }
        }
        
        // Update Brand
        if (p.brand !== undefined && p.brand !== '') {
          var brandIdx = headers.indexOf('Brand');
          if (brandIdx !== -1) {
            sheet.getRange(actualRow, brandIdx + 1).setValue(p.brand);
            updatedFields.push('Brand');
          }
        }
        
        // Update Model
        if (p.model !== undefined) {
          var modelIdx = headers.indexOf('Model');
          if (modelIdx !== -1) {
            sheet.getRange(actualRow, modelIdx + 1).setValue(p.model);
            updatedFields.push('Model');
          }
        }
        
        // Update UnitPrice
        if (p.unitPrice !== undefined) {
          var upIdx = headers.indexOf('UnitPrice');
          if (upIdx !== -1) {
            var newUP = Number(p.unitPrice) || 0;
            if (newUP <= 0) return _jr({ success: false, error: 'විකුණුම් මිල Rs. 0 ට වඩා වැඩි විය යුතුයි.' });
            sheet.getRange(actualRow, upIdx + 1).setValue(newUP);
            updatedFields.push('UnitPrice');
          }
        }
        
        // Update CostPrice
        if (p.costPrice !== undefined) {
          var cpIdx = headers.indexOf('CostPrice');
          if (cpIdx !== -1) {
            var newCP = Number(p.costPrice) || 0;
            if (newCP <= 0) return _jr({ success: false, error: 'මිලදී ගත් මිල Rs. 0 ට වඩා වැඩි විය යුතුයි.' });
            sheet.getRange(actualRow, cpIdx + 1).setValue(newCP);
            updatedFields.push('CostPrice');
          }
        }
        
        // Update StockQty
        if (p.stockQty !== undefined) {
          var sqIdx = headers.indexOf('StockQty');
          var msIdx = headers.indexOf('MinStockLevel');
          var stIdx = headers.indexOf('Status');
          if (sqIdx !== -1) {
            var newSQ = Number(p.stockQty) || 0;
            if (newSQ < 0) newSQ = 0;
            sheet.getRange(actualRow, sqIdx + 1).setValue(newSQ);
            updatedFields.push('StockQty');
            
            // Auto-update status
            var minLevel = msIdx !== -1 ? Number(data[rowIndex][msIdx]) || 5 : 5;
            if (p.minStockLevel !== undefined) minLevel = Number(p.minStockLevel) || 5;
            
            if (stIdx !== -1) {
              var newStatus = newSQ <= minLevel ? 'Low Stock' : 'Active';
              sheet.getRange(actualRow, stIdx + 1).setValue(newStatus);
            }
            
            // Low stock alert
            if (newSQ <= minLevel) {
              var prodName = data[rowIndex][headers.indexOf('ProductName')];
              checkAndAlertLowStock(p.productId, prodName, newSQ, minLevel);
            }
          }
        }
        
        // Update MinStockLevel
        if (p.minStockLevel !== undefined) {
          var minIdx = headers.indexOf('MinStockLevel');
          if (minIdx !== -1) {
            var newMin = Number(p.minStockLevel) || 5;
            sheet.getRange(actualRow, minIdx + 1).setValue(newMin);
            updatedFields.push('MinStockLevel');
            
            // Re-check status with new min level
            var sqIdx2 = headers.indexOf('StockQty');
            var stIdx2 = headers.indexOf('Status');
            var currentQty = p.stockQty !== undefined ? Number(p.stockQty) : Number(data[rowIndex][sqIdx2]) || 0;
            if (stIdx2 !== -1) {
              sheet.getRange(actualRow, stIdx2 + 1).setValue(currentQty <= newMin ? 'Low Stock' : 'Active');
            }
          }
        }
        
        // Update WarrantyMonths
        if (p.warrantyMonths !== undefined) {
          var wmIdx = headers.indexOf('WarrantyMonths');
          if (wmIdx !== -1) {
            sheet.getRange(actualRow, wmIdx + 1).setValue(Number(p.warrantyMonths) || 0);
            updatedFields.push('WarrantyMonths');
          }
        }
        
        // Update Supplier
        if (p.supplier !== undefined) {
          var supIdx = headers.indexOf('Supplier');
          if (supIdx !== -1) {
            sheet.getRange(actualRow, supIdx + 1).setValue(p.supplier);
            updatedFields.push('Supplier');
          }
        }
        
        if (updatedFields.length === 0) {
          return _jr({ success: false, error: 'Update කරන්න fields නැත.' });
        }
        
        return _jr({
          success: true,
          message: '✅ ' + p.productId + ' සාර්ථකව update කරන ලදී. Updated: ' + updatedFields.join(', '),
          productId: p.productId,
          updatedFields: updatedFields
        });
        
      } finally { lock.releaseLock(); }
    }

    // ╔═══════════════════════════════════════╗
    // ║    DELETE PRODUCT — භාණ්ඩ මකන්න       ║
    // ╚═══════════════════════════════════════╝
    if (action === 'deleteProduct') {
      const lock = LockService.getScriptLock();
      try { lock.waitLock(15000); } catch (lk) { 
        return _jr({ success: false, error: 'System busy. නැවත උත්සාහ කරන්න.' }); 
      }
      
      try {
        var productId = String(payload.productId || payload.data && payload.data.productId || '').trim();
        
        if (!productId) {
          return _jr({ success: false, error: 'Product ID අවශ්‍යයි.' });
        }
        
        // Check if product has sales records
        var salesData = getSheetData('Sales');
        var hasSales = salesData.find(function(s) { return s.ProductID === productId; });
        if (hasSales) {
          return _jr({ 
            success: false, 
            error: 'මෙම භාණ්ඩයට විකුණුම් records තිබේ. මකන්න බැහැ. Product: ' + productId + 
                   '\n\nSale ID: ' + hasSales.SaleID + ' — ' + hasSales.ProductName +
                   '\n\nඉඟිය: මෙම භාණ්ඩය මකනවා වෙනුවට Status "Discontinued" කරන්න.'
          });
        }
        
        var sheet = SPREADSHEET.getSheetByName('Products');
        var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        
        var rowIndex = -1;
        var productName = '';
        for (var i = 0; i < data.length; i++) {
          if (String(data[i][0]).trim() === productId) {
            rowIndex = i;
            productName = data[i][1];
            break;
          }
        }
        
        if (rowIndex === -1) {
          return _jr({ success: false, error: 'Product "' + productId + '" හමු නොවීය.' });
        }
        
        // Delete the row
        sheet.deleteRow(rowIndex + 2);
        
        return _jr({
          success: true,
          message: '🗑️ "' + productName + '" (' + productId + ') සාර්ථකව මකා දැමීය.',
          productId: productId,
          productName: productName
        });
        
      } finally { lock.releaseLock(); }
    }

    // ╔═══════════════════════════════════════╗
    // ║        ADD SALE — අලෙවිය එකතු කරන්න      ║
    // ╚═══════════════════════════════════════╝
    if (action === 'addSale') {
      const lock = LockService.getScriptLock();
      try { lock.waitLock(15000); } catch (lk) { 
        return _jr({ success: false, error: 'System busy. නැවත උත්සාහ කරන්න.' }); 
      }

      try {
        const order = payload.data;
        const items = order.items;
        const discPct = Number(order.discountPct) || 0;
        const custName = order.customerName || 'Walk-in';
        const custPhone = order.customerPhone || '';
        const payMethod = order.paymentMethod || 'Cash';
        const soldBy = order.soldBy || 'System';

        if (!items || items.length === 0) {
          return _jr({ success: false, error: 'භාණ්ඩ නැත. Cart එකට items එකතු කරන්න.' });
        }
        
        // Validate discount
        if (discPct < 0 || discPct > 100) {
          return _jr({ success: false, error: 'Discount 0% - 100% අතර විය යුතුයි.' });
        }

        var salesSheet = SPREADSHEET.getSheetByName('Sales');
        var maxNum = 0;
        if (salesSheet.getLastRow() >= 2) {
          var saleIDs = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 1).getValues();
          saleIDs.forEach(function(r) {
            var n = parseInt(String(r[0]).replace(/\D/g, '')) || 0;
            if (n > maxNum) maxNum = n;
          });
        }
        var saleId = 'S' + String(maxNum + 1).padStart(4, '0');

        var prodSheet = SPREADSHEET.getSheetByName('Products');
        var prodHeaders = prodSheet.getRange(1, 1, 1, prodSheet.getLastColumn()).getValues()[0];
        var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();

        // Validate all items first
        var validated = [];
        for (var idx = 0; idx < items.length; idx++) {
          var item = items[idx];
          var pi = -1;
          for (var j = 0; j < prodData.length; j++) {
            if (String(prodData[j][0]).trim() === String(item.productId).trim()) {
              pi = j;
              break;
            }
          }
          
          if (pi === -1) {
            return _jr({ success: false, error: 'භාණ්ඩය "' + item.productId + '" හමු නොවීය.' });
          }
          if (prodData[pi][7] < item.qty) {
            return _jr({ success: false, error: '"' + prodData[pi][1] + '" — තොගයේ ඇත්තේ ' + prodData[pi][7] + ' ක් පමණි. ඔබට ' + item.qty + ' ක් අවශ්‍යයි.' });
          }
          
          var uPrice = Number(prodData[pi][5]) || 0;
          var lineTotal = Math.round(uPrice * item.qty * (1 - discPct / 100));
          
          validated.push({
            prodIdx: pi,
            productId: prodData[pi][0],
            productName: prodData[pi][1],
            category: prodData[pi][2],
            qty: item.qty,
            unitPrice: uPrice,
            lineTotal: lineTotal
          });
        }

        var now = new Date();
        var saleRows = validated.map(function(v) {
          return [
            saleId, now, v.productId, v.productName, v.category,
            v.qty, v.unitPrice, discPct, v.lineTotal,
            custName, custPhone, payMethod, soldBy, ''
          ];
        });
        
        salesSheet.getRange(salesSheet.getLastRow() + 1, 1, saleRows.length, saleRows[0].length).setValues(saleRows);

        // Update stock and check low stock
        var lowStockAlerts = [];
        for (var vi = 0; vi < validated.length; vi++) {
          var v = validated[vi];
          var row = v.prodIdx + 2;
          var newQty = Number(prodData[v.prodIdx][7]) - v.qty;
          var minStock = Number(prodData[v.prodIdx][8]) || 5;
          
          prodSheet.getRange(row, 8).setValue(newQty);
          
          if (newQty <= minStock) {
            prodSheet.getRange(row, 13).setValue('Low Stock');
            lowStockAlerts.push({ id: v.productId, name: v.productName, qty: newQty, min: minStock });
          }
          
          prodData[v.prodIdx][7] = newQty; // Update local copy too
        }

        // Send email alerts for low stock items
        lowStockAlerts.forEach(function(alert) {
          checkAndAlertLowStock(alert.id, alert.name, alert.qty, alert.min);
        });

        var gt = validated.reduce(function(s, v) { return s + v.lineTotal; }, 0);
        var st = validated.reduce(function(s, v) { return s + (v.unitPrice * v.qty); }, 0);

        return _jr({
          success: true,
          message: '✅ Sale ' + saleId + ' සාර්ථකව සිදු කරන ලදී! මුළු මුදල: Rs. ' + gt.toLocaleString(),
          saleId: saleId,
          date: now.toISOString(),
          items: validated.map(function(v) {
            return {
              productId: v.productId,
              productName: v.productName,
              qty: v.qty,
              unitPrice: v.unitPrice,
              lineTotal: v.lineTotal
            };
          }),
          subtotal: st,
          discountPct: discPct,
          discountAmount: st - gt,
          grandTotal: gt,
          customerName: custName,
          customerPhone: custPhone,
          paymentMethod: payMethod,
          soldBy: soldBy,
          lowStockAlerts: lowStockAlerts.length > 0 ? lowStockAlerts : undefined
        });
      } finally { lock.releaseLock(); }
    }

    // ╔═══════════════════════════════════════╗
    // ║   PROCESS RETURN — ආපසු භාරගන්න       ║
    // ╚═══════════════════════════════════════╝
    if (action === 'processReturn') {
      const lock = LockService.getScriptLock();
      try { lock.waitLock(15000); } catch (lk) { 
        return _jr({ success: false, error: 'System busy. නැවත උත්සාහ කරන්න.' }); 
      }

      try {
        const ret = payload.data;
        var saleId = String(ret.saleId || '').trim().toUpperCase();
        var productId = String(ret.productId || '').trim();
        var returnQty = Number(ret.qty) || 0;
        var reason = ret.reason || 'No reason provided';
        var processedBy = ret.processedBy || 'Admin';

        if (!saleId || !productId || returnQty <= 0) {
          return _jr({ success: false, error: 'Sale ID, Product ID, සහ වලංගු ප්‍රමාණයක් අවශ්‍යයි.' });
        }

        var salesSheet = SPREADSHEET.getSheetByName('Sales');
        var salesHeaders = salesSheet.getRange(1, 1, 1, salesSheet.getLastColumn()).getValues()[0];
        var salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();

        var saleRowIndex = -1;
        var saleLine = null;
        for (var i = 0; i < salesData.length; i++) {
          if (String(salesData[i][0]).trim().toUpperCase() === saleId && String(salesData[i][2]).trim() === productId) {
            saleRowIndex = i;
            saleLine = {};
            salesHeaders.forEach(function(h, ci) {
              saleLine[h] = salesData[i][ci] instanceof Date ? salesData[i][ci].toISOString() : salesData[i][ci];
            });
            break;
          }
        }

        if (!saleLine) {
          return _jr({ success: false, error: 'Sale ID "' + saleId + '" සමග Product "' + productId + '" හමු නොවීය.' });
        }

        // Check already returned quantity
        var returnsData = getSheetData('Returns');
        var alreadyReturned = returnsData
          .filter(function(r) { return String(r.SaleID).toUpperCase() === saleId && r.ProductID === productId && r.Status === 'Completed'; })
          .reduce(function(sum, r) { return sum + (Number(r.Qty) || 0); }, 0);

        var soldQty = Number(saleLine.Qty) || 0;
        var maxReturnable = soldQty - alreadyReturned;

        if (maxReturnable <= 0) {
          return _jr({ success: false, error: 'මෙම භාණ්ඩයේ සියලුම units ආපසු ලබාදී ඇත.' });
        }
        if (returnQty > maxReturnable) {
          return _jr({ success: false, error: 'උපරිම ආපසු ගැනීමේ ප්‍රමාණය: ' + maxReturnable + ' ක්. ඔබ ඉල්ලා ඇත: ' + returnQty });
        }

        // Calculate refund
        var unitPrice = Number(saleLine.UnitPrice) || 0;
        var discPct = Number(saleLine.DiscountPct) || 0;
        var refundPerUnit = Math.round(unitPrice * (1 - discPct / 100));
        var refundAmount = refundPerUnit * returnQty;

        // Generate Return ID
        var retSheet = SPREADSHEET.getSheetByName('Returns');
        var maxRetNum = 0;
        if (retSheet.getLastRow() >= 2) {
          retSheet.getRange(2, 1, retSheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
            var n = parseInt(String(r[0]).replace(/\D/g, '')) || 0;
            if (n > maxRetNum) maxRetNum = n;
          });
        }
        var returnId = 'R' + String(maxRetNum + 1).padStart(4, '0');

        // Write return record
        retSheet.appendRow([
          returnId, new Date(), saleId, productId, saleLine.ProductName || '',
          returnQty, reason, refundAmount, processedBy, 'Completed'
        ]);
        retSheet.autoResizeColumns(1, retSheet.getLastColumn());

        // Update sale return status
        var rsColIdx = salesHeaders.indexOf('ReturnStatus');
        if (rsColIdx !== -1) {
          var newTotal = alreadyReturned + returnQty;
          salesSheet.getRange(saleRowIndex + 2, rsColIdx + 1).setValue(newTotal >= soldQty ? 'Returned' : 'Partial Return');
        }

        // Update product stock (add back)
        var prodSheet = SPREADSHEET.getSheetByName('Products');
        var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();
        for (var pi = 0; pi < prodData.length; pi++) {
          if (String(prodData[pi][0]).trim() === productId) {
            var newQty = Number(prodData[pi][7]) + returnQty;
            var minStock = Number(prodData[pi][8]) || 5;
            prodSheet.getRange(pi + 2, 8).setValue(newQty);
            // Update status both directions
            if (newQty > minStock && String(prodData[pi][12]) === 'Low Stock') {
              prodSheet.getRange(pi + 2, 13).setValue('Active');
            } else if (newQty <= minStock) {
              prodSheet.getRange(pi + 2, 13).setValue('Low Stock');
            }
            break;
          }
        }

        return _jr({
          success: true,
          returnId: returnId,
          refundAmount: refundAmount,
          productName: saleLine.ProductName,
          productId: productId,
          qty: returnQty,
          customerName: saleLine.CustomerName || '',
          saleId: saleId,
          message: '✅ Return ' + returnId + ' සාර්ථකව සකසන ලදී. ආපසු මුදල: Rs. ' + refundAmount.toLocaleString()
        });
      } finally { lock.releaseLock(); }
    }

    // ╔═══════════════════════════════════════╗
    // ║    ADD RESTOCK — තොග එකතු කරන්න        ║
    // ╚═══════════════════════════════════════╝
    if (action === 'addRestock') {
      const lock = LockService.getScriptLock();
      try { lock.waitLock(15000); } catch (lk) { 
        return _jr({ success: false, error: 'System busy. නැවත උත්සාහ කරන්න.' }); 
      }
      
      try {
        var r = payload.data;
        
        if (!r.productId || !r.qty || Number(r.qty) <= 0) {
          return _jr({ success: false, error: 'Product ID සහ වලංගු ප්‍රමාණයක් අවශ්‍යයි.' });
        }
        
        var productId = String(r.productId).trim();
        var qty = Number(r.qty);
        var supplier = r.supplier || '';
        var unitCost = Number(r.unitCost) || 0;
        var receivedBy = r.receivedBy || 'Admin';
        var notes = r.notes || '';
        
        // Find product
        var prodSheet = SPREADSHEET.getSheetByName('Products');
        var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();
        
        var prodIdx = -1;
        var productName = '';
        for (var i = 0; i < prodData.length; i++) {
          if (String(prodData[i][0]).trim() === productId) {
            prodIdx = i;
            productName = prodData[i][1];
            if (!supplier) supplier = prodData[i][10] || '';
            if (unitCost <= 0) unitCost = Number(prodData[i][6]) || 0;
            break;
          }
        }
        
        if (prodIdx === -1) {
          return _jr({ success: false, error: 'Product "' + productId + '" හමු නොවීය.' });
        }
        
        // Update stock
        var newQty = Number(prodData[prodIdx][7]) + qty;
        var minStock = Number(prodData[prodIdx][8]) || 5;
        prodSheet.getRange(prodIdx + 2, 8).setValue(newQty);
        
        // Update status
        if (newQty > minStock) {
          prodSheet.getRange(prodIdx + 2, 13).setValue('Active');
        }
        
        // Generate Restock ID
        var rsSheet = SPREADSHEET.getSheetByName('RestockLog');
        var maxRsNum = 0;
        if (rsSheet.getLastRow() >= 2) {
          rsSheet.getRange(2, 1, rsSheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
            var n = parseInt(String(row[0]).replace(/\D/g, '')) || 0;
            if (n > maxRsNum) maxRsNum = n;
          });
        }
        var restockId = 'RS' + String(maxRsNum + 1).padStart(3, '0');
        
        // Write restock record
        rsSheet.appendRow([
          restockId, new Date(), productId, productName,
          qty, supplier, unitCost, unitCost * qty,
          receivedBy, notes
        ]);
        rsSheet.autoResizeColumns(1, rsSheet.getLastColumn());
        
        return _jr({
          success: true,
          message: '✅ ' + productName + ' — ' + qty + ' units එකතු කරන ලදී. නව තොගය: ' + newQty,
          restockId: restockId,
          productId: productId,
          productName: productName,
          addedQty: qty,
          newStockQty: newQty,
          totalCost: unitCost * qty
        });
        
      } finally { lock.releaseLock(); }
    }

    // ╔═══════════════════════════════════════╗
    // ║     GET REPORT — වාර්තා ලබාගන්න         ║
    // ╚═══════════════════════════════════════╝
    if (action === 'getReport') {
      var reportType = payload.reportType || 'summary';
      var dateFrom = payload.dateFrom ? new Date(payload.dateFrom) : null;
      var dateTo = payload.dateTo ? new Date(payload.dateTo) : null;
      
      var salesData = getSheetData('Sales');
      var productsData = getSheetData('Products');
      var returnsData = getSheetData('Returns');
      
      // Filter by date range if provided
      if (dateFrom || dateTo) {
        salesData = salesData.filter(function(s) {
          var sd = new Date(s.Date);
          if (dateFrom && sd < dateFrom) return false;
          if (dateTo) {
            var endDate = new Date(dateTo);
            endDate.setHours(23, 59, 59, 999);
            if (sd > endDate) return false;
          }
          return true;
        });
      }
      
      // Total revenue
      var totalRevenue = salesData.reduce(function(s, x) { return s + (Number(x.TotalAmount) || 0); }, 0);
      
      // Unique sales count
      var uniqueSaleIds = {};
      salesData.forEach(function(s) { uniqueSaleIds[s.SaleID] = true; });
      var totalSales = Object.keys(uniqueSaleIds).length;
      
      // Total units sold
      var totalUnits = salesData.reduce(function(s, x) { return s + (Number(x.Qty) || 0); }, 0);
      
      // Category breakdown
      var catBreakdown = {};
      salesData.forEach(function(s) {
        if (!catBreakdown[s.Category]) catBreakdown[s.Category] = { revenue: 0, qty: 0, count: 0 };
        catBreakdown[s.Category].revenue += Number(s.TotalAmount) || 0;
        catBreakdown[s.Category].qty += Number(s.Qty) || 0;
        catBreakdown[s.Category].count++;
      });
      
      // Payment method breakdown
      var payBreakdown = {};
      salesData.forEach(function(s) {
        if (!payBreakdown[s.PaymentMethod]) payBreakdown[s.PaymentMethod] = { revenue: 0, count: 0 };
        payBreakdown[s.PaymentMethod].revenue += Number(s.TotalAmount) || 0;
        payBreakdown[s.PaymentMethod].count++;
      });
      
      // Top products
      var prodSales = {};
      salesData.forEach(function(s) {
        if (!prodSales[s.ProductID]) prodSales[s.ProductID] = { name: s.ProductName, revenue: 0, qty: 0 };
        prodSales[s.ProductID].revenue += Number(s.TotalAmount) || 0;
        prodSales[s.ProductID].qty += Number(s.Qty) || 0;
      });
      var topProducts = Object.keys(prodSales).map(function(k) {
        return { productId: k, productName: prodSales[k].name, revenue: prodSales[k].revenue, qty: prodSales[k].qty };
      }).sort(function(a, b) { return b.revenue - a.revenue; }).slice(0, 10);
      
      // Low stock items
      var lowStock = productsData.filter(function(p) { 
        return p.Status === 'Low Stock' || Number(p.StockQty) <= Number(p.MinStockLevel); 
      });
      
      // Return stats
      var totalRefunds = returnsData.reduce(function(s, r) { return s + (Number(r.RefundAmount) || 0); }, 0);
      
      // Estimated profit (revenue - cost for sold items)
      var estimatedProfit = 0;
      salesData.forEach(function(s) {
        var prod = productsData.find(function(p) { return p.ProductID === s.ProductID; });
        if (prod) {
          var costTotal = (Number(prod.CostPrice) || 0) * (Number(s.Qty) || 0);
          estimatedProfit += (Number(s.TotalAmount) || 0) - costTotal;
        }
      });
      
      return _jr({
        success: true,
        report: {
          type: reportType,
          dateRange: { from: dateFrom ? dateFrom.toISOString() : null, to: dateTo ? dateTo.toISOString() : null },
          totalRevenue: totalRevenue,
          totalSales: totalSales,
          totalUnits: totalUnits,
          averageOrderValue: totalSales > 0 ? Math.round(totalRevenue / totalSales) : 0,
          estimatedProfit: estimatedProfit,
          totalRefunds: totalRefunds,
          netRevenue: totalRevenue - totalRefunds,
          categoryBreakdown: catBreakdown,
          paymentBreakdown: payBreakdown,
          topProducts: topProducts,
          lowStockItems: lowStock,
          totalProducts: productsData.length,
          lowStockCount: lowStock.length
        }
      });
    }

    // Unknown action
    return _jr({ success: false, error: 'නොදන්නා action එකකි: "' + action + '"' });
    
  } catch (err) {
    Logger.log('doPost Error: ' + err.toString());
    return _jr({ success: false, error: 'Server error: ' + err.toString() });
  }
}

// ============================================================
// JSON RESPONSE HELPER
// ============================================================
function _jr(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// MENU — Google Sheets menu
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Mahesh App')
    .addItem('📋 Initialize All Sheets (සියලුම sheets හදන්න)', 'setupSheets')
    .addItem('📧 Test Low Stock Email', 'testLowStockEmail')
    .addItem('📊 Generate Summary Report', 'generateSummaryReport')
    .addSeparator()
    .addItem('🌐 Deploy Instructions', 'showDeployInfo')
    .addToUi();
}

function testLowStockEmail() {
  sendLowStockAlert('TEST001', 'Test Product', 2, 5);
  SpreadsheetApp.getUi().alert('📧 Test email sent to: ' + ADMIN_EMAIL + '\n\nEmail inbox එක check කරන්න.');
}

function generateSummaryReport() {
  var products = getSheetData('Products');
  var sales = getSheetData('Sales');
  var returns = getSheetData('Returns');
  
  var totalRevenue = sales.reduce(function(s, x) { return s + (Number(x.TotalAmount) || 0); }, 0);
  var uniqueSales = {};
  sales.forEach(function(s) { uniqueSales[s.SaleID] = true; });
  var totalOrders = Object.keys(uniqueSales).length;
  var totalRefunds = returns.reduce(function(s, r) { return s + (Number(r.RefundAmount) || 0); }, 0);
  var lowStock = products.filter(function(p) { return p.Status === 'Low Stock' || Number(p.StockQty) <= Number(p.MinStockLevel); });
  
  var inventoryValue = products.reduce(function(s, p) { return s + (Number(p.UnitPrice) || 0) * (Number(p.StockQty) || 0); }, 0);
  var costValue = products.reduce(function(s, p) { return s + (Number(p.CostPrice) || 0) * (Number(p.StockQty) || 0); }, 0);
  
  SpreadsheetApp.getUi().alert(
    '📊 SUMMARY REPORT — සාරාංශ වාර්තාව\n' +
    '═══════════════════════════\n\n' +
    '💰 මුළු ආදායම: Rs. ' + totalRevenue.toLocaleString() + '\n' +
    '🛒 මුළු orders: ' + totalOrders + '\n' +
    '📦 මුළු භාණ්ඩ: ' + products.length + '\n' +
    '⚠️ අඩු තොග: ' + lowStock.length + ' items\n' +
    '🔄 ආපසු මුදල: Rs. ' + totalRefunds.toLocaleString() + '\n' +
    '✅ ශුද්ධ ආදායම: Rs. ' + (totalRevenue - totalRefunds).toLocaleString() + '\n\n' +
    '📈 තොග වටිනාකම (විකුණුම්): Rs. ' + inventoryValue.toLocaleString() + '\n' +
    '📈 තොග වටිනාකම (මිලදී ගත්): Rs. ' + costValue.toLocaleString() + '\n\n' +
    (lowStock.length > 0 ? '⚠️ අඩු තොග භාණ්ඩ:\n' + lowStock.map(function(p) { return '  • ' + p.ProductName + ' (' + p.StockQty + '/' + p.MinStockLevel + ')'; }).join('\n') : '✅ සියලුම භාණ්ඩ හොඳ තොග මට්ටමක!')
  );
}

function showDeployInfo() {
  SpreadsheetApp.getUi().alert(
    '🌐 Deploy කරන ආකාරය\n' +
    '═══════════════════════════\n\n' +
    '1️⃣ Deploy > New Deployment click කරන්න\n' +
    '2️⃣ Type: "Web app" select කරන්න\n' +
    '3️⃣ Execute as: "Me" select කරන්න\n' +
    '4️⃣ Who has access: "Anyone" select කරන්න\n' +
    '5️⃣ Deploy click කරන්න\n' +
    '6️⃣ URL එක copy කරන්න\n' +
    '7️⃣ index.html එකේ API_URL එකට paste කරන්න\n\n' +
    '⚠️ වැදගත්:\n' +
    '• Code එක වෙනස් කළ සෑම විටම NEW deployment එකක් කරන්න\n' +
    '• "New Deployment" use කරන්න (Manage Deployments නොවේ)\n' +
    '• URL එක "https://script.google.com/macros/s/..." ආකාරයේ විය යුතුයි'
  );
}
