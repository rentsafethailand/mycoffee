// =====================================================
// ☕ ระบบ POS ร้านกาแฟ Multi-Tenant
// 📅 สร้างเมื่อ: 2025-01-13
// 👨‍💻 สำหรับ: Google Apps Script
// =====================================================

// =====================================================
// 🔧 CONFIGURATION - ตั้งค่าพื้นฐาน
// =====================================================

// ⚠️ สำคัญ: ให้แก้ไข ID เหล่านี้หลังจากรัน createMasterSheet()
var MASTER_SHEET_ID = 'YOUR_MASTER_SHEET_ID_HERE'; // ใส่ Sheet ID ของ Master Sheet
var MASTER_FOLDER_ID = 'YOUR_MASTER_FOLDER_ID_HERE'; // ใส่ Folder ID ของ Master Folder
var MASTER_PASSWORD = 'superadmin123'; // รหัสผ่าน Super Admin

// ราคาไลเซ่นส์
var LICENSE_PRICE = {
  monthly: 199,
  yearly: 1399
};

// =====================================================
// 📱 WEB APP ENTRY POINT
// =====================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('☕ ระบบ POS ร้านกาแฟ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================
// 🔨 SETUP FUNCTIONS - ฟังก์ชันสร้างระบบ
// =====================================================

/**
 * สร้าง Master Sheet (รันอัตโนมัติถ้ายังไม่มี)
 */
function createMasterSheet() {
  try {
    // สร้าง Master Sheet
    var ss = SpreadsheetApp.create('☕ POS Master Database');
    var masterSheetId = ss.getId();

    // สร้าง Master Folder
    var folder = DriveApp.createFolder('☕ POS Coffee Shops');
    var folderId = folder.getId();

    // ย้าย Master Sheet เข้า folder
    var file = DriveApp.getFileById(masterSheetId);
    file.moveTo(folder);

    Logger.log('===================================');
    Logger.log('✅ สร้าง Master Sheet สำเร็จ!');
    Logger.log('📋 Sheet ID: ' + masterSheetId);
    Logger.log('📁 Folder ID: ' + folderId);
    Logger.log('🔗 URL: ' + ss.getUrl());
    Logger.log('===================================');

    // สร้างชีตหลัก
    var sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
    if (sheetNames.indexOf('Sheet1') > -1) {
      ss.getSheetByName('Sheet1').setName('ร้านทั้งหมด');
    }

    var mainSheet = ss.getSheetByName('ร้านทั้งหมด');

    // สร้างหัวตาราง
    var headers = [
      'รหัสร้าน',
      'ชื่อร้าน',
      'Email',
      'รหัสผ่าน',
      'Sheet ID',
      'Folder ID',
      'วันที่สร้าง',
      'แพ็คเกจ',
      'วันหมดอายุ',
      'สถานะ',
      'หมายเหตุ'
    ];

    mainSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // ปรับความกว้างคอลัมน์
    mainSheet.setColumnWidth(1, 100); // รหัสร้าน
    mainSheet.setColumnWidth(2, 200); // ชื่อร้าน
    mainSheet.setColumnWidth(3, 250); // Email
    mainSheet.setColumnWidth(4, 120); // รหัสผ่าน
    mainSheet.setColumnWidth(5, 300); // Sheet ID
    mainSheet.setColumnWidth(6, 300); // Folder ID
    mainSheet.setColumnWidth(7, 120); // วันที่สร้าง
    mainSheet.setColumnWidth(8, 100); // แพ็คเกจ
    mainSheet.setColumnWidth(9, 120); // วันหมดอายุ
    mainSheet.setColumnWidth(10, 100); // สถานะ
    mainSheet.setColumnWidth(11, 200); // หมายเหตุ

    // Freeze header
    mainSheet.setFrozenRows(1);

    return {
      success: true,
      sheetId: masterSheetId,
      folderId: folderId,
      url: ss.getUrl()
    };

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * สร้างร้านใหม่
 */
function createNewShop(shopData) {
  try {
    // ตรวจสอบข้อมูล
    if (!shopData.shopName || !shopData.email || !shopData.password || !shopData.packageType) {
      throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน');
    }

    // เช็ค Email ซ้ำ
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();
    var emailCol = findColumnIndex(data[0], 'Email');

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol] === shopData.email) {
        throw new Error('Email นี้มีในระบบแล้ว');
      }
    }

    // สร้าง Google Sheet ใหม่
    var newSheet = SpreadsheetApp.create('☕ ' + shopData.shopName);
    var sheetId = newSheet.getId();

    // สร้างโฟลเดอร์ย่อย
    var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
    var shopFolder = masterFolder.createFolder(shopData.shopName + ' - ' + shopData.email);
    var folderId = shopFolder.getId();

    // ย้าย Sheet เข้าโฟลเดอร์
    var file = DriveApp.getFileById(sheetId);
    file.moveTo(shopFolder);

    // สร้างโครงสร้างชีต
    setupShopSheets(sheetId, shopData);

    // คำนวณวันหมดอายุ
    var today = new Date();
    var expiryDate = new Date(today);
    if (shopData.packageType === 'monthly') {
      expiryDate.setMonth(expiryDate.getMonth() + 1);
    } else {
      expiryDate.setFullYear(expiryDate.getFullYear() + 1);
    }

    // สร้างรหัสร้าน
    var shopId = 'SHOP' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    // บันทึกลง Master Sheet
    masterSheet.appendRow([
      shopId,
      shopData.shopName,
      shopData.email,
      shopData.password,
      sheetId,
      folderId,
      Utilities.formatDate(today, 'GMT+7', 'dd/MM/yyyy HH:mm:ss'),
      shopData.packageType === 'monthly' ? 'รายเดือน' : 'รายปี',
      Utilities.formatDate(expiryDate, 'GMT+7', 'dd/MM/yyyy'),
      'ใช้งาน',
      'สร้างโดย Super Admin'
    ]);

    Logger.log('✅ สร้างร้าน ' + shopData.shopName + ' สำเร็จ!');

    return {
      success: true,
      shopId: shopId,
      sheetId: sheetId,
      folderId: folderId,
      message: 'สร้างร้านสำเร็จ'
    };

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * สร้างโครงสร้างชีตสำหรับร้าน
 */
function setupShopSheets(sheetId, shopData) {
  var ss = SpreadsheetApp.openById(sheetId);

  // ลบชีตเริ่มต้น
  var sheets = ss.getSheets();
  if (sheets.length > 0) {
    ss.deleteSheet(sheets[0]);
  }

  // 1. ชีต: ตั้งค่า
  var settingsSheet = ss.insertSheet('ตั้งค่า');
  settingsSheet.getRange('A1:B1').setValues([['คีย์', 'ค่า']])
    .setBackground('#2196F3').setFontColor('#FFFFFF').setFontWeight('bold');
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 400);

  var settingsData = [
    ['ชื่อร้าน', shopData.shopName],
    ['Email', shopData.email],
    ['ที่อยู่', ''],
    ['เบอร์โทร', ''],
    ['หมายเลขพร้อมเพย์', ''],
    ['ช่องทางการสั่งซื้อ', 'Dine-in,Take away,Delivery'],
    ['รูปแบบหมายเลขออเดอร์', 'auto'],
    ['รูปแบบ OrderID', 'ORD-{DATE}-{NUMBER}'],
    ['เลขออเดอร์ล่าสุด', '0'],
    ['lastQueueUpdate', new Date().toISOString()],
    ['lastOrderUpdate', new Date().toISOString()]
  ];
  settingsSheet.getRange(2, 1, settingsData.length, 2).setValues(settingsData);

  // 2. ชีต: หมวดหมู่สินค้า
  var categorySheet = ss.insertSheet('หมวดหมู่สินค้า');
  categorySheet.getRange('A1:C1').setValues([['รหัสหมวดหมู่', 'ชื่อหมวดหมู่', 'ลำดับ']])
    .setBackground('#FF9800').setFontColor('#FFFFFF').setFontWeight('bold');
  categorySheet.setColumnWidths(1, 3, 150);

  // 3. ชีต: ตัวเลือกสินค้า
  var optionsSheet = ss.insertSheet('ตัวเลือกสินค้า');
  var optionHeaders = [
    'รหัสตัวเลือก',
    'ชื่อกลุ่ม',
    'ประเภทกลุ่ม',
    'ชื่อตัวเลือก',
    'ราคาเพิ่ม',
    'ตัวคูณวัตถุดิบ',
    'วัตถุดิบพิเศษ',
    'ลำดับ'
  ];
  optionsSheet.getRange(1, 1, 1, optionHeaders.length).setValues([optionHeaders])
    .setBackground('#9C27B0').setFontColor('#FFFFFF').setFontWeight('bold');
  optionsSheet.setColumnWidths(1, optionHeaders.length, 120);

  // 4. ชีต: สินค้า
  var productsSheet = ss.insertSheet('สินค้า');
  var productHeaders = [
    'รหัสสินค้า',
    'ชื่อสินค้า',
    'หมวดหมู่',
    'ราคาขาย',
    'ราคาทุน',
    'เวลาทำ(นาที)',
    'วัตถุดิบพื้นฐาน',
    'กลุ่มตัวเลือก',
    'ตัวเลือกเริ่มต้น',
    'สถานะ'
  ];
  productsSheet.getRange(1, 1, 1, productHeaders.length).setValues([productHeaders])
    .setBackground('#4CAF50').setFontColor('#FFFFFF').setFontWeight('bold');
  productsSheet.setColumnWidths(1, productHeaders.length, 120);

  // 5. ชีต: วัตถุดิบ
  var materialsSheet = ss.insertSheet('วัตถุดิบ');
  var materialHeaders = [
    'รหัสวัตถุดิบ',
    'ชื่อวัตถุดิบ',
    'ราคาต่อหน่วย',
    'สต๊อก',
    'หน่วย',
    'สถานะ'
  ];
  materialsSheet.getRange(1, 1, 1, materialHeaders.length).setValues([materialHeaders])
    .setBackground('#795548').setFontColor('#FFFFFF').setFontWeight('bold');
  materialsSheet.setColumnWidths(1, materialHeaders.length, 150);

  // 6. ชีต: การขาย
  var salesSheet = ss.insertSheet('การขาย');
  var salesHeaders = [
    'เลขที่ออเดอร์',
    'วันที่',
    'เวลาสั่ง',
    'รหัสสินค้า',
    'ชื่อสินค้า',
    'ตัวเลือกที่เลือก',
    'จำนวน',
    'ราคาต่อหน่วย',
    'ราคารวม',
    'สถานะ',
    'เวลาโดยประมาณ(นาที)',
    'เวลาเริ่มทำ',
    'เวลาเสร็จ',
    'ช่องทางสั่งซื้อ',
    'โต๊ะ/หมายเหตุ',
    'ช่องทางชำระเงิน',
    'เงินสด',
    'โอน',
    'QR Code',
    'URL สลิป',
    'พนักงาน'
  ];
  salesSheet.getRange(1, 1, 1, salesHeaders.length).setValues([salesHeaders])
    .setBackground('#F44336').setFontColor('#FFFFFF').setFontWeight('bold');
  salesSheet.setColumnWidths(1, salesHeaders.length, 120);

  // 7. ชีต: ต้นทุน
  var costsSheet = ss.insertSheet('ต้นทุน');
  var costHeaders = [
    'วันที่',
    'ประเภท',
    'รายการ',
    'จำนวนเงิน',
    'หมายเหตุ'
  ];
  costsSheet.getRange(1, 1, 1, costHeaders.length).setValues([costHeaders])
    .setBackground('#607D8B').setFontColor('#FFFFFF').setFontWeight('bold');
  costsSheet.setColumnWidths(1, costHeaders.length, 150);

  Logger.log('✅ สร้างโครงสร้างชีตสำเร็จ');
}

/**
 * สร้างร้านตัวอย่าง (พร้อมข้อมูลตัวอย่าง)
 */
function createSampleShop() {
  try {
    var sampleData = {
      shopName: 'ร้านกาแฟตัวอย่าง',
      email: 'sample@coffee.com',
      password: 'sample123',
      packageType: 'monthly'
    };

    var result = createNewShop(sampleData);

    if (result.success) {
      // เพิ่มข้อมูลตัวอย่าง
      addSampleData(result.sheetId);

      Logger.log('✅ สร้างร้านตัวอย่างสำเร็จ!');
      Logger.log('📧 Email: sample@coffee.com');
      Logger.log('🔑 Password: sample123');

      return {
        success: true,
        message: 'สร้างร้านตัวอย่างสำเร็จ',
        email: 'sample@coffee.com',
        password: 'sample123',
        sheetId: result.sheetId
      };
    }

    return result;

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มข้อมูลตัวอย่าง
 */
function addSampleData(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);

  // เพิ่มหมวดหมู่ตัวอย่าง
  var categorySheet = ss.getSheetByName('หมวดหมู่สินค้า');
  categorySheet.appendRow(['CAT001', 'กาแฟ', 1]);
  categorySheet.appendRow(['CAT002', 'ชา', 2]);
  categorySheet.appendRow(['CAT003', 'เครื่องดื่มปั่น', 3]);

  // เพิ่มตัวเลือกตัวอย่าง
  var optionsSheet = ss.getSheetByName('ตัวเลือกสินค้า');
  optionsSheet.appendRow(['OP001', 'ชนิด', 'single', 'ร้อน', 0, 1.0, '', 1]);
  optionsSheet.appendRow(['OP002', 'ชนิด', 'single', 'เย็น', 0, 1.0, '', 2]);
  optionsSheet.appendRow(['OP003', 'ชนิด', 'single', 'ปั่น', 5, 1.2, '', 3]);
  optionsSheet.appendRow(['OP004', 'ขนาด', 'single', 'Small', -10, 0.7, '', 1]);
  optionsSheet.appendRow(['OP005', 'ขนาด', 'single', 'Medium', 0, 1.0, '', 2]);
  optionsSheet.appendRow(['OP006', 'ขนาด', 'single', 'Large', 10, 1.3, '', 3]);
  optionsSheet.appendRow(['OP007', 'ความหวาน', 'single', 'หวานน้อย', 0, 0.5, '', 1]);
  optionsSheet.appendRow(['OP008', 'ความหวาน', 'single', 'หวานปกติ', 0, 1.0, '', 2]);
  optionsSheet.appendRow(['OP009', 'ความหวาน', 'single', 'หวานมาก', 0, 1.5, '', 3]);
  optionsSheet.appendRow(['OP010', 'ท็อปปิ้ง', 'multiple', 'วิปครีม', 15, 0, '', 1]);
  optionsSheet.appendRow(['OP011', 'ท็อปปิ้ง', 'multiple', 'ช็อตเพิ่ม', 20, 0, '', 2]);

  // เพิ่มวัตถุดิบตัวอย่าง
  var materialsSheet = ss.getSheetByName('วัตถุดิบ');
  materialsSheet.appendRow(['M001', 'เมล็ดกาแฟ', 0.5, 5000, 'g', 'ใช้งาน']);
  materialsSheet.appendRow(['M002', 'นมสด', 0.2, 10000, 'ml', 'ใช้งาน']);
  materialsSheet.appendRow(['M003', 'น้ำเชื่อม', 0.1, 3000, 'ml', 'ใช้งาน']);
  materialsSheet.appendRow(['M004', 'วิปครีม', 0.5, 500, 'g', 'ใช้งาน']);

  // เพิ่มสินค้าตัวอย่าง
  var productsSheet = ss.getSheetByName('สินค้า');
  productsSheet.appendRow([
    'P001',
    'Americano',
    'กาแฟ',
    40,
    15,
    3,
    'M001:15g,M002:50ml',
    'ชนิด,ขนาด,ความหวาน',
    'OP002,OP005,OP008',
    'ใช้งาน'
  ]);
  productsSheet.appendRow([
    'P002',
    'Latte',
    'กาแฟ',
    50,
    20,
    5,
    'M001:15g,M002:200ml,M003:15ml',
    'ชนิด,ขนาด,ความหวาน,ท็อปปิ้ง',
    'OP002,OP005,OP008',
    'ใช้งาน'
  ]);
  productsSheet.appendRow([
    'P003',
    'Cappuccino',
    'กาแฟ',
    55,
    22,
    5,
    'M001:15g,M002:150ml,M003:10ml',
    'ชนิด,ขนาด,ความหวาน,ท็อปปิ้ง',
    'OP002,OP005,OP008',
    'ใช้งาน'
  ]);

  Logger.log('✅ เพิ่มข้อมูลตัวอย่างสำเร็จ');
}

// =====================================================
// 🔐 AUTHENTICATION - ระบบ Login/Logout
// =====================================================

/**
 * Login ด้วย Email + Password
 */
function login(email, password) {
  try {
    if (!email || !password) {
      throw new Error('กรุณากรอก Email และรหัสผ่าน');
    }

    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    // หาคอลัมน์
    var emailCol = findColumnIndex(data[0], 'Email');
    var passwordCol = findColumnIndex(data[0], 'รหัสผ่าน');
    var sheetIdCol = findColumnIndex(data[0], 'Sheet ID');
    var shopNameCol = findColumnIndex(data[0], 'ชื่อร้าน');
    var statusCol = findColumnIndex(data[0], 'สถานะ');
    var expiryCol = findColumnIndex(data[0], 'วันหมดอายุ');
    var packageCol = findColumnIndex(data[0], 'แพ็คเกจ');

    // หาร้านที่ตรงกับ Email
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        // เช็ครหัสผ่าน
        if (data[i][passwordCol] !== password) {
          throw new Error('รหัสผ่านไม่ถูกต้อง');
        }

        // เช็คสถานะร้าน
        if (data[i][statusCol] === 'ระงับ') {
          throw new Error('ร้านของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ');
        }

        // เช็คไลเซ่นส์
        var expiryDate = new Date(data[i][expiryCol]);
        var today = new Date();

        if (expiryDate < today) {
          throw new Error('ไลเซ่นส์ของคุณหมดอายุแล้ว กรุณาต่ออายุ');
        }

        // คำนวณวันคงเหลือ
        var daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

        // Login สำเร็จ
        return {
          success: true,
          sheetId: data[i][sheetIdCol],
          shopName: data[i][shopNameCol],
          email: email,
          expiryDate: Utilities.formatDate(expiryDate, 'GMT+7', 'dd/MM/yyyy'),
          daysLeft: daysLeft,
          packageType: data[i][packageCol],
          warning: daysLeft <= 30 ? 'ไลเซ่นส์ของคุณเหลืออีก ' + daysLeft + ' วัน' : null
        };
      }
    }

    // ไม่พบ Email
    throw new Error('ไม่พบ Email นี้ในระบบ');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เช็คไลเซ่นส์
 */
function checkLicense(sheetId) {
  try {
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    var sheetIdCol = findColumnIndex(data[0], 'Sheet ID');
    var expiryCol = findColumnIndex(data[0], 'วันหมดอายุ');
    var packageCol = findColumnIndex(data[0], 'แพ็คเกจ');
    var shopNameCol = findColumnIndex(data[0], 'ชื่อร้าน');

    for (var i = 1; i < data.length; i++) {
      if (data[i][sheetIdCol] === sheetId) {
        var expiryDate = new Date(data[i][expiryCol]);
        var today = new Date();
        var daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

        return {
          success: true,
          shopName: data[i][shopNameCol],
          packageType: data[i][packageCol],
          expiryDate: Utilities.formatDate(expiryDate, 'GMT+7', 'dd/MM/yyyy'),
          daysLeft: daysLeft,
          isExpired: expiryDate < today
        };
      }
    }

    throw new Error('ไม่พบข้อมูลร้าน');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เปลี่ยนรหัสผ่าน
 */
function changePassword(email, oldPassword, newPassword) {
  try {
    if (!newPassword || newPassword.length < 6) {
      throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร');
    }

    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    var emailCol = findColumnIndex(data[0], 'Email');
    var passwordCol = findColumnIndex(data[0], 'รหัสผ่าน');

    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        if (data[i][passwordCol] !== oldPassword) {
          throw new Error('รหัสผ่านเก่าไม่ถูกต้อง');
        }

        masterSheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);

        return {
          success: true,
          message: 'เปลี่ยนรหัสผ่านสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบ Email นี้ในระบบ');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 📦 PRODUCT & OPTIONS MANAGEMENT
// =====================================================

/**
 * ดึงข้อมูลหมวดหมู่สินค้า
 */
function getCategories(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('หมวดหมู่สินค้า');
    var data = sheet.getDataRange().getValues();

    var categories = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // มีรหัสหมวดหมู่
        categories.push({
          id: data[i][0],
          name: data[i][1],
          order: data[i][2]
        });
      }
    }

    // เรียงตามลำดับ
    categories.sort(function(a, b) { return a.order - b.order; });

    return {
      success: true,
      data: categories
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มหมวดหมู่
 */
function addCategory(sheetId, categoryData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('หมวดหมู่สินค้า');

    var categoryId = 'CAT' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    sheet.appendRow([
      categoryId,
      categoryData.name,
      categoryData.order || 999
    ]);

    return {
      success: true,
      categoryId: categoryId,
      message: 'เพิ่มหมวดหมู่สำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทหมวดหมู่
 */
function updateCategory(sheetId, categoryId, categoryData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('หมวดหมู่สินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === categoryId) {
        sheet.getRange(i + 1, 2).setValue(categoryData.name);
        sheet.getRange(i + 1, 3).setValue(categoryData.order);

        return {
          success: true,
          message: 'อัพเดทหมวดหมู่สำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบหมวดหมู่นี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบหมวดหมู่
 */
function deleteCategory(sheetId, categoryId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('หมวดหมู่สินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === categoryId) {
        sheet.deleteRow(i + 1);

        return {
          success: true,
          message: 'ลบหมวดหมู่สำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบหมวดหมู่นี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงข้อมูลตัวเลือกสินค้าทั้งหมด
 */
function getProductOptions(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
    var data = sheet.getDataRange().getValues();

    var options = [];
    var groups = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { // มีรหัสตัวเลือก
        var groupName = data[i][1];

        if (!groups[groupName]) {
          groups[groupName] = {
            groupName: groupName,
            groupType: data[i][2],
            options: []
          };
        }

        groups[groupName].options.push({
          id: data[i][0],
          name: data[i][3],
          priceModifier: parseFloat(data[i][4]) || 0,
          materialMultiplier: parseFloat(data[i][5]) || 0,
          specialMaterials: data[i][6] || '',
          order: data[i][7] || 999
        });
      }
    }

    // แปลง object เป็น array
    for (var key in groups) {
      groups[key].options.sort(function(a, b) { return a.order - b.order; });
      options.push(groups[key]);
    }

    return {
      success: true,
      data: options
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มตัวเลือกสินค้า
 */
function addProductOption(sheetId, optionData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');

    var optionId = 'OP' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    sheet.appendRow([
      optionId,
      optionData.groupName,
      optionData.groupType || 'single',
      optionData.optionName,
      optionData.priceModifier || 0,
      optionData.materialMultiplier || 0,
      optionData.specialMaterials || '',
      optionData.order || 999
    ]);

    return {
      success: true,
      optionId: optionId,
      message: 'เพิ่มตัวเลือกสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทตัวเลือกสินค้า
 */
function updateProductOption(sheetId, optionId, optionData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === optionId) {
        sheet.getRange(i + 1, 2).setValue(optionData.groupName);
        sheet.getRange(i + 1, 3).setValue(optionData.groupType);
        sheet.getRange(i + 1, 4).setValue(optionData.optionName);
        sheet.getRange(i + 1, 5).setValue(optionData.priceModifier);
        sheet.getRange(i + 1, 6).setValue(optionData.materialMultiplier);
        sheet.getRange(i + 1, 7).setValue(optionData.specialMaterials || '');
        sheet.getRange(i + 1, 8).setValue(optionData.order);

        return {
          success: true,
          message: 'อัพเดทตัวเลือกสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบตัวเลือกนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * สร้างตัวเลือกสินค้าใหม่
 */
function createProductOption(sheetId, optionData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');

    // Generate new option ID
    var data = sheet.getDataRange().getValues();
    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var idNum = parseInt(data[i][0].replace(/[^\d]/g, ''));
        if (idNum > maxId) maxId = idNum;
      }
    }
    var newId = 'OP' + String(maxId + 1).padStart(3, '0');

    // Add new row
    sheet.appendRow([
      newId,
      optionData.groupName || '',
      optionData.groupType || 'single',
      optionData.optionName || '',
      parseFloat(optionData.priceModifier) || 0,
      parseFloat(optionData.materialMultiplier) || 1.0,
      optionData.specialMaterials || '',
      parseInt(optionData.order) || 1,
      optionData.status || 'เปิดใช้งาน'
    ]);

    return {
      success: true,
      message: 'เพิ่มตัวเลือกสำเร็จ',
      optionId: newId
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * บันทึกตัวเลือกสินค้า (Create or Update)
 */
function saveProductOption(sheetId, optionData) {
  if (optionData.id) {
    // Update existing
    return updateProductOption(sheetId, optionData.id, optionData);
  } else {
    // Create new
    return createProductOption(sheetId, optionData);
  }
}

/**
 * ลบตัวเลือกสินค้า
 */
function deleteProductOption(sheetId, optionId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === optionId) {
        sheet.deleteRow(i + 1);

        return {
          success: true,
          message: 'ลบตัวเลือกสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบตัวเลือกนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงรายการกลุ่มตัวเลือกทั้งหมด (ไม่ซ้ำ)
 */
function getOptionGroups(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
    var data = sheet.getDataRange().getValues();

    var groups = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) { // มีชื่อกลุ่ม
        var groupName = data[i][1];
        if (!groups[groupName]) {
          groups[groupName] = {
            name: groupName,
            type: data[i][2] || 'single',
            options: []
          };
        }

        groups[groupName].options.push({
          id: data[i][0],
          name: data[i][3],
          priceModifier: parseFloat(data[i][4]) || 0,
          materialMultiplier: parseFloat(data[i][5]) || 1.0,
          order: parseInt(data[i][7]) || 1,
          status: data[i][8] || 'เปิดใช้งาน'
        });
      }
    }

    // Convert to array and sort options
    var groupArray = [];
    for (var key in groups) {
      groups[key].options.sort(function(a, b) { return a.order - b.order; });
      groupArray.push(groups[key]);
    }

    return {
      success: true,
      data: groupArray
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงข้อมูลสินค้าทั้งหมด
 */
function getProducts(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('สินค้า');
    var data = sheet.getDataRange().getValues();

    var products = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][9] === 'ใช้งาน') { // มีรหัสสินค้าและสถานะใช้งาน
        products.push({
          id: data[i][0],
          name: data[i][1],
          category: data[i][2],
          price: parseFloat(data[i][3]) || 0,
          cost: parseFloat(data[i][4]) || 0,
          prepTime: parseInt(data[i][5]) || 0,
          baseMaterials: data[i][6] || '',
          optionGroups: data[i][7] || '',
          defaultOptions: data[i][8] || '',
          status: data[i][9]
        });
      }
    }

    return {
      success: true,
      data: products
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มสินค้า
 */
function addProduct(sheetId, productData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('สินค้า');

    var productId = 'P' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    sheet.appendRow([
      productId,
      productData.name,
      productData.category,
      productData.price || 0,
      productData.cost || 0,
      productData.prepTime || 0,
      productData.baseMaterials || '',
      productData.optionGroups || '',
      productData.defaultOptions || '',
      'ใช้งาน'
    ]);

    return {
      success: true,
      productId: productId,
      message: 'เพิ่มสินค้าสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทสินค้า
 */
function updateProduct(sheetId, productId, productData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('สินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        sheet.getRange(i + 1, 2).setValue(productData.name);
        sheet.getRange(i + 1, 3).setValue(productData.category);
        sheet.getRange(i + 1, 4).setValue(productData.price);
        sheet.getRange(i + 1, 5).setValue(productData.cost);
        sheet.getRange(i + 1, 6).setValue(productData.prepTime);
        sheet.getRange(i + 1, 7).setValue(productData.baseMaterials);
        sheet.getRange(i + 1, 8).setValue(productData.optionGroups);
        sheet.getRange(i + 1, 9).setValue(productData.defaultOptions);

        return {
          success: true,
          message: 'อัพเดทสินค้าสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบสินค้านี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบสินค้า (เปลี่ยนสถานะ)
 */
function deleteProduct(sheetId, productId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('สินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        sheet.getRange(i + 1, 10).setValue('ลบแล้ว');

        return {
          success: true,
          message: 'ลบสินค้าสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบสินค้านี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * คำนวณราคาสินค้ารวมตัวเลือก
 */
function calculateProductPrice(sheetId, productId, selectedOptions) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);

    // ดึงราคาพื้นฐาน
    var productSheet = ss.getSheetByName('สินค้า');
    var productData = productSheet.getDataRange().getValues();
    var basePrice = 0;

    for (var i = 1; i < productData.length; i++) {
      if (productData[i][0] === productId) {
        basePrice = parseFloat(productData[i][3]) || 0;
        break;
      }
    }

    // ดึงราคาเพิ่มจากตัวเลือก
    var optionsSheet = ss.getSheetByName('ตัวเลือกสินค้า');
    var optionsData = optionsSheet.getDataRange().getValues();
    var totalModifier = 0;

    var selectedIds = selectedOptions.split(',');
    for (var j = 0; j < selectedIds.length; j++) {
      for (var k = 1; k < optionsData.length; k++) {
        if (optionsData[k][0] === selectedIds[j].trim()) {
          totalModifier += parseFloat(optionsData[k][4]) || 0;
          break;
        }
      }
    }

    return {
      success: true,
      basePrice: basePrice,
      modifier: totalModifier,
      totalPrice: basePrice + totalModifier
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 🧪 MATERIALS MANAGEMENT
// =====================================================

/**
 * ดึงข้อมูลวัตถุดิบทั้งหมด
 */
function getMaterials(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');
    var data = sheet.getDataRange().getValues();

    var materials = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][5] === 'ใช้งาน') {
        materials.push({
          id: data[i][0],
          name: data[i][1],
          pricePerUnit: parseFloat(data[i][2]) || 0,
          stock: parseFloat(data[i][3]) || 0,
          unit: data[i][4],
          status: data[i][5]
        });
      }
    }

    return {
      success: true,
      data: materials
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มวัตถุดิบ
 */
function addMaterial(sheetId, materialData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');

    var materialId = 'M' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');

    sheet.appendRow([
      materialId,
      materialData.name,
      materialData.pricePerUnit || 0,
      materialData.stock || 0,
      materialData.unit || '',
      'ใช้งาน'
    ]);

    return {
      success: true,
      materialId: materialId,
      message: 'เพิ่มวัตถุดิบสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทวัตถุดิบ
 */
function updateMaterial(sheetId, materialId, materialData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === materialId) {
        sheet.getRange(i + 1, 2).setValue(materialData.name);
        sheet.getRange(i + 1, 3).setValue(materialData.pricePerUnit);
        sheet.getRange(i + 1, 4).setValue(materialData.stock);
        sheet.getRange(i + 1, 5).setValue(materialData.unit);

        return {
          success: true,
          message: 'อัพเดทวัตถุดิบสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบวัตถุดิบนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบวัตถุดิบ (เปลี่ยนสถานะ)
 */
function deleteMaterial(sheetId, materialId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === materialId) {
        sheet.getRange(i + 1, 6).setValue('ลบแล้ว');

        return {
          success: true,
          message: 'ลบวัตถุดิบสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบวัตถุดิบนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มสต๊อกวัตถุดิบ (สั่งซื้อ)
 */
function addMaterialStock(sheetId, materialId, quantity) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === materialId) {
        var currentStock = parseFloat(data[i][3]) || 0;
        var newStock = currentStock + parseFloat(quantity);
        sheet.getRange(i + 1, 4).setValue(newStock);

        return {
          success: true,
          message: 'เพิ่มสต๊อกสำเร็จ',
          newStock: newStock
        };
      }
    }

    throw new Error('ไม่พบวัตถุดิบนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลดสต๊อกวัตถุดิบ (ของเสีย)
 */
function reduceMaterialStock(sheetId, materialId, quantity, reason) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('วัตถุดิบ');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === materialId) {
        var currentStock = parseFloat(data[i][3]) || 0;
        var newStock = currentStock - parseFloat(quantity);

        if (newStock < 0) {
          throw new Error('สต๊อกไม่เพียงพอ');
        }

        sheet.getRange(i + 1, 4).setValue(newStock);

        return {
          success: true,
          message: 'ลดสต๊อกสำเร็จ',
          newStock: newStock
        };
      }
    }

    throw new Error('ไม่พบวัตถุดิบนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 💰 SALES & POS FUNCTIONS
// =====================================================

/**
 * สร้างหมายเลขออเดอร์
 */
function generateOrderNumber(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var settingsSheet = ss.getSheetByName('ตั้งค่า');
    var data = settingsSheet.getDataRange().getValues();

    // หาค่าต่างๆ
    var orderMode = '';
    var orderFormat = '';
    var lastNumber = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'รูปแบบหมายเลขออเดอร์') {
        orderMode = data[i][1];
      } else if (data[i][0] === 'รูปแบบ OrderID') {
        orderFormat = data[i][1];
      } else if (data[i][0] === 'เลขออเดอร์ล่าสุด') {
        lastNumber = parseInt(data[i][1]) || 0;
      }
    }

    if (orderMode === 'manual') {
      return {
        success: true,
        mode: 'manual',
        orderNumber: null
      };
    }

    // Auto mode
    var newNumber = lastNumber + 1;

    // สร้างเลขออเดอร์ตามรูปแบบ
    var orderNumber = orderFormat;
    orderNumber = orderNumber.replace('{DATE}', Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd'));
    orderNumber = orderNumber.replace('{NUMBER}', ('000' + newNumber).slice(-3));

    // อัพเดทเลขล่าสุด
    for (var j = 1; j < data.length; j++) {
      if (data[j][0] === 'เลขออเดอร์ล่าสุด') {
        settingsSheet.getRange(j + 1, 2).setValue(newNumber);
        break;
      }
    }

    return {
      success: true,
      mode: 'auto',
      orderNumber: orderNumber
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * บันทึกการขาย
 */
function saveSale(sheetId, saleData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var salesSheet = ss.getSheetByName('การขาย');

    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy');
    var timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');

    // บันทึกแต่ละรายการ
    for (var i = 0; i < saleData.items.length; i++) {
      var item = saleData.items[i];

      salesSheet.appendRow([
        saleData.orderNumber,
        dateStr,
        timeStr,
        item.productId,
        item.productName,
        item.selectedOptions || '',
        item.quantity,
        item.pricePerUnit,
        item.totalPrice,
        'รอทำ',
        item.prepTime * item.quantity,
        '',
        '',
        saleData.channel || 'Dine-in',
        saleData.tableNote || '',
        saleData.paymentMethods.join(','),
        saleData.cashAmount || 0,
        saleData.transferAmount || 0,
        saleData.qrAmount || 0,
        saleData.slipUrl || '',
        saleData.staff || ''
      ]);
    }

    // อัพเดท Timestamp
    updateQueueTimestamp(sheetId);

    return {
      success: true,
      message: 'บันทึกการขายสำเร็จ',
      orderNumber: saleData.orderNumber
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงยอดขายวันนี้
 */
function getTodaySales(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
    var orders = {};
    var totalSales = 0;

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === today) {
        var orderNum = data[i][0];
        var itemTotal = parseFloat(data[i][8]) || 0;
        totalSales += itemTotal;

        if (!orders[orderNum]) {
          orders[orderNum] = {
            orderNumber: orderNum,
            date: data[i][1],
            time: data[i][2],
            status: data[i][9],
            channel: data[i][13],
            items: [],
            total: 0
          };
        }

        orders[orderNum].items.push({
          productName: data[i][4],
          options: data[i][5],
          quantity: data[i][6],
          price: itemTotal
        });

        orders[orderNum].total += itemTotal;
      }
    }

    var orderArray = [];
    for (var key in orders) {
      orderArray.push(orders[key]);
    }

    return {
      success: true,
      totalSales: totalSales,
      orders: orderArray
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงประวัติการขายตามช่วงวันที่
 */
function getSalesByDateRange(sheetId, startDate, endDate) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var orders = {};
    var totalSales = 0;

    for (var i = 1; i < data.length; i++) {
      var orderDate = data[i][1];

      // เช็คว่าอยู่ในช่วงหรือไม่ (แบบง่าย)
      if (isDateInRange(orderDate, startDate, endDate)) {
        var orderNum = data[i][0];
        var itemTotal = parseFloat(data[i][8]) || 0;
        totalSales += itemTotal;

        if (!orders[orderNum]) {
          orders[orderNum] = {
            orderNumber: orderNum,
            date: data[i][1],
            time: data[i][2],
            status: data[i][9],
            channel: data[i][13],
            items: [],
            total: 0
          };
        }

        orders[orderNum].items.push({
          productName: data[i][4],
          options: data[i][5],
          quantity: data[i][6],
          price: itemTotal
        });

        orders[orderNum].total += itemTotal;
      }
    }

    var orderArray = [];
    for (var key in orders) {
      orderArray.push(orders[key]);
    }

    return {
      success: true,
      totalSales: totalSales,
      orderCount: orderArray.length,
      orders: orderArray
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบออเดอร์
 */
function deleteOrder(sheetId, orderNumber) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var deletedRows = 0;

    // ลบจากล่างขึ้นบน
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === orderNumber) {
        sheet.deleteRow(i + 1);
        deletedRows++;
      }
    }

    if (deletedRows > 0) {
      updateQueueTimestamp(sheetId);

      return {
        success: true,
        message: 'ลบออเดอร์สำเร็จ'
      };
    }

    throw new Error('ไม่พบออเดอร์นี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 👨‍🍳 KITCHEN QUEUE SYSTEM
// =====================================================

/**
 * ดึงคิวทั้งหมด
 */
function getKitchenQueue(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
    var orders = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === today) {
        var status = data[i][9];

        // เฉพาะคิวที่ยังไม่เสร็จ
        if (status === 'รอทำ' || status === 'กำลังทำ' || status === 'พร้อมเสิร์ฟ') {
          var orderNum = data[i][0];

          if (!orders[orderNum]) {
            orders[orderNum] = {
              orderNumber: orderNum,
              date: data[i][1],
              time: data[i][2],
              status: status,
              channel: data[i][13],
              tableNote: data[i][14],
              items: [],
              totalTime: 0,
              startTime: data[i][11] || null,
              finishTime: data[i][12] || null
            };
          }

          orders[orderNum].items.push({
            productName: data[i][4],
            options: data[i][5],
            quantity: data[i][6],
            prepTime: parseInt(data[i][10]) || 0
          });

          orders[orderNum].totalTime += parseInt(data[i][10]) || 0;

          // อัพเดทสถานะล่าสุด
          if (status === 'กำลังทำ') {
            orders[orderNum].status = 'กำลังทำ';
            if (data[i][11]) orders[orderNum].startTime = data[i][11];
          }
          if (status === 'พร้อมเสิร์ฟ') {
            orders[orderNum].status = 'พร้อมเสิร์ฟ';
            if (data[i][12]) orders[orderNum].finishTime = data[i][12];
          }
        }
      }
    }

    var orderArray = [];
    for (var key in orders) {
      orderArray.push(orders[key]);
    }

    // เรียงตามเวลา
    orderArray.sort(function(a, b) {
      return a.time > b.time ? 1 : -1;
    });

    return {
      success: true,
      data: orderArray
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงสถิติคิว
 */
function getQueueStats(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
    var stats = {
      pending: 0,
      inProgress: 0,
      ready: 0,
      totalTime: 0
    };

    var orders = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === today) {
        var orderNum = data[i][0];
        var status = data[i][9];

        if (!orders[orderNum]) {
          orders[orderNum] = status;
        } else {
          if (status === 'กำลังทำ') orders[orderNum] = 'กำลังทำ';
          if (status === 'พร้อมเสิร์ฟ') orders[orderNum] = 'พร้อมเสิร์ฟ';
        }

        if (status === 'รอทำ' || status === 'กำลังทำ') {
          stats.totalTime += parseInt(data[i][10]) || 0;
        }
      }
    }

    for (var key in orders) {
      if (orders[key] === 'รอทำ') stats.pending++;
      else if (orders[key] === 'กำลังทำ') stats.inProgress++;
      else if (orders[key] === 'พร้อมเสิร์ฟ') stats.ready++;
    }

    return {
      success: true,
      data: stats
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทสถานะออเดอร์
 */
function updateOrderStatus(sheetId, orderNumber, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var updated = false;
    var now = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === orderNumber) {
        // อัพเดทสถานะ
        sheet.getRange(i + 1, 10).setValue(newStatus);

        // บันทึก Timestamp
        if (newStatus === 'กำลังทำ') {
          sheet.getRange(i + 1, 12).setValue(now); // เวลาเริ่มทำ
        } else if (newStatus === 'พร้อมเสิร์ฟ') {
          sheet.getRange(i + 1, 13).setValue(now); // เวลาเสร็จ
        }

        updated = true;
      }
    }

    if (updated) {
      updateQueueTimestamp(sheetId);

      return {
        success: true,
        message: 'อัพเดทสถานะสำเร็จ'
      };
    }

    throw new Error('ไม่พบออเดอร์นี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * แก้ไขออเดอร์
 */
function updateOrder(sheetId, orderNumber, updatedData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    // ลบรายการเดิม
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === orderNumber) {
        sheet.deleteRow(i + 1);
      }
    }

    // เพิ่มรายการใหม่
    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy');
    var timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');

    for (var j = 0; j < updatedData.items.length; j++) {
      var item = updatedData.items[j];

      sheet.appendRow([
        orderNumber,
        dateStr,
        timeStr,
        item.productId,
        item.productName,
        item.selectedOptions || '',
        item.quantity,
        item.pricePerUnit,
        item.totalPrice,
        updatedData.status || 'รอทำ',
        item.prepTime * item.quantity,
        '',
        '',
        updatedData.channel || 'Dine-in',
        updatedData.tableNote || '',
        '',
        0,
        0,
        0,
        '',
        ''
      ]);
    }

    updateQueueTimestamp(sheetId);

    return {
      success: true,
      message: 'แก้ไขออเดอร์สำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงข้อมูลออเดอร์สำหรับแก้ไข
 */
function getOrderForEdit(sheetId, orderNumber) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var order = {
      orderNumber: orderNumber,
      items: [],
      channel: '',
      tableNote: '',
      status: ''
    };

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === orderNumber) {
        order.channel = data[i][13];
        order.tableNote = data[i][14];
        order.status = data[i][9];

        order.items.push({
          productId: data[i][3],
          productName: data[i][4],
          selectedOptions: data[i][5],
          quantity: data[i][6],
          pricePerUnit: data[i][7],
          totalPrice: data[i][8],
          prepTime: data[i][10]
        });
      }
    }

    if (order.items.length === 0) {
      throw new Error('ไม่พบออเดอร์นี้');
    }

    return {
      success: true,
      data: order
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 📊 REPORTS & DASHBOARD
// =====================================================

/**
 * ดึงข้อมูล Dashboard
 */
function getDashboardData(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var salesSheet = ss.getSheetByName('การขาย');
    var costsSheet = ss.getSheetByName('ต้นทุน');

    var today = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');

    // ยอดขายวันนี้
    var salesData = salesSheet.getDataRange().getValues();
    var todaySales = 0;
    var todayCost = 0;
    var productCount = {};

    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][1] === today) {
        todaySales += parseFloat(salesData[i][8]) || 0;

        var productName = salesData[i][4];
        var quantity = parseInt(salesData[i][6]) || 0;

        if (!productCount[productName]) {
          productCount[productName] = 0;
        }
        productCount[productName] += quantity;

        // ต้นทุนจากสินค้า (ประมาณการ)
        var productId = salesData[i][3];
        var cost = getProductCost(sheetId, productId);
        todayCost += cost * quantity;
      }
    }

    // ต้นทุนอื่นๆ วันนี้
    var costsData = costsSheet.getDataRange().getValues();
    for (var j = 1; j < costsData.length; j++) {
      if (costsData[j][0] === today && costsData[j][1] === 'รายวัน') {
        todayCost += parseFloat(costsData[j][3]) || 0;
      }
    }

    // Top 5 สินค้าขายดี
    var topProducts = [];
    for (var key in productCount) {
      topProducts.push({ name: key, count: productCount[key] });
    }
    topProducts.sort(function(a, b) { return b.count - a.count; });
    topProducts = topProducts.slice(0, 5);

    // ยอดขาย 7 วันล่าสุด
    var last7Days = getLast7DaysSales(salesData);

    return {
      success: true,
      todaySales: todaySales,
      todayCost: todayCost,
      todayProfit: todaySales - todayCost,
      topProducts: topProducts,
      last7Days: last7Days
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงต้นทุนสินค้า
 */
function getProductCost(sheetId, productId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('สินค้า');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === productId) {
        return parseFloat(data[i][4]) || 0;
      }
    }

    return 0;

  } catch (error) {
    return 0;
  }
}

/**
 * ดึงยอดขาย 7 วันล่าสุด
 */
function getLast7DaysSales(salesData) {
  var result = [];

  for (var i = 6; i >= 0; i--) {
    var date = new Date();
    date.setDate(date.getDate() - i);
    var dateStr = Utilities.formatDate(date, 'GMT+7', 'dd/MM/yyyy');

    var dayTotal = 0;

    for (var j = 1; j < salesData.length; j++) {
      if (salesData[j][1] === dateStr) {
        dayTotal += parseFloat(salesData[j][8]) || 0;
      }
    }

    result.push({
      date: dateStr,
      total: dayTotal
    });
  }

  return result;
}

/**
 * ดึงข้อมูลต้นทุน
 */
function getCosts(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');
    var data = sheet.getDataRange().getValues();

    var costs = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        costs.push({
          date: data[i][0],
          type: data[i][1],
          item: data[i][2],
          amount: parseFloat(data[i][3]) || 0,
          note: data[i][4]
        });
      }
    }

    return {
      success: true,
      data: costs
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เพิ่มต้นทุน
 */
function addCost(sheetId, costData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');

    sheet.appendRow([
      costData.date,
      costData.type,
      costData.item,
      costData.amount || 0,
      costData.note || ''
    ]);

    return {
      success: true,
      message: 'เพิ่มต้นทุนสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบต้นทุน
 */
function deleteCost(sheetId, rowIndex) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');

    sheet.deleteRow(rowIndex + 1);

    return {
      success: true,
      message: 'ลบต้นทุนสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * รายงาน P&L
 */
function getProfitLossReport(sheetId, month, year) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var salesSheet = ss.getSheetByName('การขาย');
    var costsSheet = ss.getSheetByName('ต้นทุน');

    var salesData = salesSheet.getDataRange().getValues();
    var costsData = costsSheet.getDataRange().getValues();

    var totalRevenue = 0;
    var totalCost = 0;

    // รายได้
    for (var i = 1; i < salesData.length; i++) {
      var dateStr = salesData[i][1];
      if (isInMonth(dateStr, month, year)) {
        totalRevenue += parseFloat(salesData[i][8]) || 0;

        var productId = salesData[i][3];
        var quantity = parseInt(salesData[i][6]) || 0;
        var cost = getProductCost(sheetId, productId);
        totalCost += cost * quantity;
      }
    }

    // ต้นทุนอื่นๆ
    for (var j = 1; j < costsData.length; j++) {
      var costDate = costsData[j][0];
      if (isInMonth(costDate, month, year)) {
        totalCost += parseFloat(costsData[j][3]) || 0;
      }
    }

    var profit = totalRevenue - totalCost;
    var profitMargin = totalRevenue > 0 ? (profit / totalRevenue * 100).toFixed(2) : 0;

    return {
      success: true,
      revenue: totalRevenue,
      cost: totalCost,
      profit: profit,
      profitMargin: profitMargin
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 👑 SUPER ADMIN & LICENSE MANAGEMENT
// =====================================================

/**
 * Login Super Admin
 */
function superAdminLogin(password) {
  try {
    if (password === MASTER_PASSWORD) {
      return {
        success: true,
        message: 'เข้าสู่ระบบ Super Admin สำเร็จ'
      };
    }

    throw new Error('รหัสผ่านไม่ถูกต้อง');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงข้อมูลร้านทั้งหมด
 */
function getAllShops() {
  try {
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    var shops = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var expiryDate = new Date(data[i][8]);
        var today = new Date();
        var daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

        shops.push({
          shopId: data[i][0],
          shopName: data[i][1],
          email: data[i][2],
          password: data[i][3],
          sheetId: data[i][4],
          folderId: data[i][5],
          createdDate: data[i][6],
          packageType: data[i][7],
          expiryDate: Utilities.formatDate(expiryDate, 'GMT+7', 'dd/MM/yyyy'),
          status: data[i][9],
          note: data[i][10],
          daysLeft: daysLeft,
          isExpired: expiryDate < today
        });
      }
    }

    return {
      success: true,
      data: shops
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทข้อมูลร้าน
 */
function updateShop(shopId, shopData) {
  try {
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === shopId) {
        masterSheet.getRange(i + 1, 2).setValue(shopData.shopName);
        masterSheet.getRange(i + 1, 3).setValue(shopData.email);
        if (shopData.password) {
          masterSheet.getRange(i + 1, 4).setValue(shopData.password);
        }
        masterSheet.getRange(i + 1, 10).setValue(shopData.status);
        masterSheet.getRange(i + 1, 11).setValue(shopData.note);

        return {
          success: true,
          message: 'อัพเดทข้อมูลร้านสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบร้านนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ต่ออายุไลเซ่นส์
 */
function renewLicense(shopId, packageType) {
  try {
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === shopId) {
        var currentExpiry = new Date(data[i][8]);
        var today = new Date();
        var startDate = currentExpiry > today ? currentExpiry : today;

        var newExpiry = new Date(startDate);
        if (packageType === 'monthly') {
          newExpiry.setMonth(newExpiry.getMonth() + 1);
        } else {
          newExpiry.setFullYear(newExpiry.getFullYear() + 1);
        }

        masterSheet.getRange(i + 1, 8).setValue(packageType === 'monthly' ? 'รายเดือน' : 'รายปี');
        masterSheet.getRange(i + 1, 9).setValue(Utilities.formatDate(newExpiry, 'GMT+7', 'dd/MM/yyyy'));
        masterSheet.getRange(i + 1, 10).setValue('ใช้งาน');

        return {
          success: true,
          message: 'ต่ออายุไลเซ่นส์สำเร็จ',
          newExpiryDate: Utilities.formatDate(newExpiry, 'GMT+7', 'dd/MM/yyyy')
        };
      }
    }

    throw new Error('ไม่พบร้านนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบร้าน
 */
function deleteShop(shopId) {
  try {
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var data = masterSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === shopId) {
        masterSheet.deleteRow(i + 1);

        return {
          success: true,
          message: 'ลบร้านสำเร็จ'
        };
      }
    }

    throw new Error('ไม่พบร้านนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// 🔧 HELPER FUNCTIONS
// =====================================================

/**
 * หาตำแหน่งคอลัมน์จากชื่อ
 */
function findColumnIndex(headerRow, columnName) {
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === columnName) {
      return i;
    }
  }
  return -1;
}

/**
 * อัพเดท Timestamp สำหรับคิว
 */
function updateQueueTimestamp(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var settingsSheet = ss.getSheetByName('ตั้งค่า');
    var data = settingsSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'lastQueueUpdate') {
        settingsSheet.getRange(i + 1, 2).setValue(new Date().toISOString());
        return;
      }
    }

  } catch (error) {
    Logger.log('Error updating timestamp: ' + error.toString());
  }
}

/**
 * ดึง Timestamp
 */
function getSystemTimestamp(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var settingsSheet = ss.getSheetByName('ตั้งค่า');
    var data = settingsSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'lastQueueUpdate') {
        return {
          success: true,
          lastQueueUpdate: data[i][1]
        };
      }
    }

    return {
      success: true,
      lastQueueUpdate: new Date().toISOString()
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงการตั้งค่า
 */
function getSettings(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตั้งค่า');
    var data = sheet.getDataRange().getValues();

    var settings = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings[data[i][0]] = data[i][1];
      }
    }

    return {
      success: true,
      data: settings
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * อัพเดทการตั้งค่า
 */
function updateSettings(sheetId, settingsData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ตั้งค่า');
    var data = sheet.getDataRange().getValues();

    for (var key in settingsData) {
      var found = false;

      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settingsData[key]);
          found = true;
          break;
        }
      }

      if (!found) {
        sheet.appendRow([key, settingsData[key]]);
      }
    }

    return {
      success: true,
      message: 'อัพเดทการตั้งค่าสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Upload ไฟล์ไป Drive
 */
function uploadFileToDrive(sheetId, base64Data, filename, mimeType) {
  try {
    // หา Folder ID
    var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('ร้านทั้งหมด');
    var masterData = masterSheet.getDataRange().getValues();
    var folderId = '';

    for (var i = 1; i < masterData.length; i++) {
      if (masterData[i][4] === sheetId) {
        folderId = masterData[i][5];
        break;
      }
    }

    if (!folderId) {
      throw new Error('ไม่พบโฟลเดอร์');
    }

    // Decode base64
    var bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    var blob = Utilities.newBlob(bytes, mimeType, filename);

    // Upload
    var folder = DriveApp.getFolderById(folderId);
    var file = folder.createFile(blob);

    return {
      success: true,
      fileId: file.getId(),
      url: file.getUrl()
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * เช็คว่าวันที่อยู่ในช่วงหรือไม่
 */
function isDateInRange(dateStr, startDate, endDate) {
  try {
    var parts = dateStr.split('/');
    var date = new Date(parts[2], parts[1] - 1, parts[0]);

    var startParts = startDate.split('/');
    var start = new Date(startParts[2], startParts[1] - 1, startParts[0]);

    var endParts = endDate.split('/');
    var end = new Date(endParts[2], endParts[1] - 1, endParts[0]);

    return date >= start && date <= end;

  } catch (error) {
    return false;
  }
}

/**
 * เช็คว่าวันที่อยู่ในเดือนหรือไม่
 */
function isInMonth(dateStr, month, year) {
  try {
    var parts = dateStr.split('/');
    var dateMonth = parseInt(parts[1]);
    var dateYear = parseInt(parts[2]);

    return dateMonth === month && dateYear === year;

  } catch (error) {
    return false;
  }
}

// ===================================
// REPORTS FUNCTIONS
// ===================================

/**
 * รายงานยอดขาย
 */
function getSalesReport(sheetId, startDate, endDate) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var start = new Date(startDate);
    var end = new Date(endDate);
    end.setHours(23, 59, 59, 999);

    var dailySales = {};
    var totalSales = 0;
    var totalOrders = 0;
    var paymentMethods = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var orderDate = new Date(data[i][1]);

        if (orderDate >= start && orderDate <= end) {
          var dateKey = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          var amount = parseFloat(data[i][5]) || 0;
          var payment = data[i][6] || 'เงินสด';

          // Daily sales
          if (!dailySales[dateKey]) {
            dailySales[dateKey] = 0;
          }
          dailySales[dateKey] += amount;

          // Total
          totalSales += amount;
          totalOrders++;

          // Payment methods
          if (!paymentMethods[payment]) {
            paymentMethods[payment] = 0;
          }
          paymentMethods[payment] += amount;
        }
      }
    }

    // Convert to arrays
    var dailyArray = [];
    for (var key in dailySales) {
      dailyArray.push({
        date: key,
        sales: dailySales[key]
      });
    }
    dailyArray.sort(function(a, b) {
      return a.date > b.date ? 1 : -1;
    });

    var paymentArray = [];
    for (var key in paymentMethods) {
      paymentArray.push({
        method: key,
        amount: paymentMethods[key],
        percentage: (paymentMethods[key] / totalSales * 100).toFixed(2)
      });
    }

    return {
      success: true,
      data: {
        totalSales: totalSales,
        totalOrders: totalOrders,
        averageOrder: totalOrders > 0 ? (totalSales / totalOrders) : 0,
        dailySales: dailyArray,
        paymentMethods: paymentArray
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * รายงานกำไร-ขาดทุน
 */
function getProfitLossReport(sheetId, month, year) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var salesSheet = ss.getSheetByName('การขาย');
    var costsSheet = ss.getSheetByName('ต้นทุน');

    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0, 23, 59, 59, 999);

    // Calculate sales
    var salesData = salesSheet.getDataRange().getValues();
    var totalSales = 0;
    var totalCost = 0;

    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][0]) {
        var orderDate = new Date(salesData[i][1]);
        if (orderDate >= startDate && orderDate <= endDate) {
          totalSales += parseFloat(salesData[i][5]) || 0;
          totalCost += parseFloat(salesData[i][10]) || 0; // ต้นทุนสินค้า
        }
      }
    }

    // Calculate other costs
    var costsData = costsSheet.getDataRange().getValues();
    var operatingCosts = 0;
    var costsByCategory = {};

    for (var i = 1; i < costsData.length; i++) {
      if (costsData[i][0]) {
        var costDate = new Date(costsData[i][0]);
        if (costDate >= startDate && costDate <= endDate) {
          var amount = parseFloat(costsData[i][2]) || 0;
          var category = costsData[i][1] || 'อื่นๆ';

          operatingCosts += amount;

          if (!costsByCategory[category]) {
            costsByCategory[category] = 0;
          }
          costsByCategory[category] += amount;
        }
      }
    }

    var totalCosts = totalCost + operatingCosts;
    var grossProfit = totalSales - totalCost;
    var netProfit = totalSales - totalCosts;
    var profitMargin = totalSales > 0 ? (netProfit / totalSales * 100) : 0;

    // Convert costs to array
    var costsArray = [];
    for (var key in costsByCategory) {
      costsArray.push({
        category: key,
        amount: costsByCategory[key]
      });
    }

    return {
      success: true,
      data: {
        totalSales: totalSales,
        productCost: totalCost,
        operatingCosts: operatingCosts,
        totalCosts: totalCosts,
        grossProfit: grossProfit,
        netProfit: netProfit,
        profitMargin: profitMargin,
        costsByCategory: costsArray
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * รายงานสินค้าขายดี
 */
function getBestSellersReport(sheetId, startDate, endDate, limit) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('การขาย');
    var data = sheet.getDataRange().getValues();

    var start = new Date(startDate);
    var end = new Date(endDate);
    end.setHours(23, 59, 59, 999);

    var productSales = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var orderDate = new Date(data[i][1]);

        if (orderDate >= start && orderDate <= end) {
          var items = data[i][4] || '';
          var amount = parseFloat(data[i][5]) || 0;

          // Parse items (format: "สินค้า1 x2, สินค้า2 x1")
          var itemList = items.split(',');

          for (var j = 0; j < itemList.length; j++) {
            var item = itemList[j].trim();
            var match = item.match(/(.+?)\s+x(\d+)/);

            if (match) {
              var productName = match[1].trim();
              var quantity = parseInt(match[2]) || 0;

              if (!productSales[productName]) {
                productSales[productName] = {
                  name: productName,
                  quantity: 0,
                  revenue: 0
                };
              }

              productSales[productName].quantity += quantity;
              // Estimate revenue (split total evenly)
              productSales[productName].revenue += (amount / itemList.length);
            }
          }
        }
      }
    }

    // Convert to array and sort
    var productArray = [];
    for (var key in productSales) {
      productArray.push(productSales[key]);
    }

    productArray.sort(function(a, b) {
      return b.quantity - a.quantity;
    });

    // Limit results
    if (limit && limit > 0) {
      productArray = productArray.slice(0, limit);
    }

    return {
      success: true,
      data: productArray
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ===================================
// COSTS MANAGEMENT FUNCTIONS
// ===================================

/**
 * ดึงรายจ่ายรายวัน
 */
function getDailyCosts(sheetId, date) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');
    var data = sheet.getDataRange().getValues();

    var targetDate = new Date(date);
    var costs = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var costDate = new Date(data[i][0]);

        if (Utilities.formatDate(costDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') ===
            Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')) {

          costs.push({
            id: 'COST' + i,
            date: Utilities.formatDate(costDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            category: data[i][1] || '',
            description: data[i][3] || '',
            amount: parseFloat(data[i][2]) || 0,
            note: data[i][4] || ''
          });
        }
      }
    }

    return {
      success: true,
      data: costs
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * บันทึกรายจ่ายรายวัน
 */
function saveDailyCost(sheetId, costData) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');

    if (costData.id && costData.id.startsWith('COST')) {
      // Update existing
      var rowIndex = parseInt(costData.id.replace('COST', ''));
      sheet.getRange(rowIndex, 1).setValue(new Date(costData.date));
      sheet.getRange(rowIndex, 2).setValue(costData.category);
      sheet.getRange(rowIndex, 3).setValue(parseFloat(costData.amount) || 0);
      sheet.getRange(rowIndex, 4).setValue(costData.description);
      sheet.getRange(rowIndex, 5).setValue(costData.note);
    } else {
      // Create new
      sheet.appendRow([
        new Date(costData.date),
        costData.category || '',
        parseFloat(costData.amount) || 0,
        costData.description || '',
        costData.note || ''
      ]);
    }

    return {
      success: true,
      message: 'บันทึกรายจ่ายสำเร็จ'
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ดึงข้อมูลรายจ่ายรายวันตาม ID
 */
function getDailyCost(sheetId, costId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');

    if (costId && costId.startsWith('COST')) {
      var rowIndex = parseInt(costId.replace('COST', ''));
      var row = sheet.getRange(rowIndex, 1, 1, 5).getValues()[0];

      var cost = {
        'รหัส': costId,
        'วันที่': row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        'ประเภท': row[1] || '',
        'จำนวนเงิน': row[2] || 0,
        'รายละเอียด': row[3] || '',
        'หมายเหตุ': row[4] || ''
      };

      return {
        success: true,
        cost: cost
      };
    }

    throw new Error('ไม่พบรายจ่ายนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ลบรายจ่ายรายวัน
 */
function deleteDailyCost(sheetId, costId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');

    if (costId && costId.startsWith('COST')) {
      var rowIndex = parseInt(costId.replace('COST', ''));
      sheet.deleteRow(rowIndex);

      return {
        success: true,
        message: 'ลบรายจ่ายสำเร็จ'
      };
    }

    throw new Error('ไม่พบรายจ่ายนี้');

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * สรุปต้นทุนรายเดือน
 */
function getMonthlyCostsSummary(sheetId, month, year) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName('ต้นทุน');
    var salesSheet = ss.getSheetByName('การขาย');

    var data = sheet.getDataRange().getValues();
    var salesData = salesSheet.getDataRange().getValues();

    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0, 23, 59, 59, 999);

    var dailyCosts = {};
    var categoryTotals = {};
    var totalCost = 0;
    var totalSales = 0;

    // Calculate costs
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var costDate = new Date(data[i][0]);

        if (costDate >= startDate && costDate <= endDate) {
          var dateKey = Utilities.formatDate(costDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          var category = data[i][1] || 'อื่นๆ';
          var amount = parseFloat(data[i][2]) || 0;

          // Daily costs
          if (!dailyCosts[dateKey]) {
            dailyCosts[dateKey] = 0;
          }
          dailyCosts[dateKey] += amount;

          // Category totals
          if (!categoryTotals[category]) {
            categoryTotals[category] = 0;
          }
          categoryTotals[category] += amount;

          totalCost += amount;
        }
      }
    }

    // Calculate sales
    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i][0]) {
        var salesDate = new Date(salesData[i][1]);
        if (salesDate >= startDate && salesDate <= endDate) {
          totalSales += parseFloat(salesData[i][5]) || 0;
        }
      }
    }

    // Convert to arrays
    var dailyArray = [];
    for (var key in dailyCosts) {
      dailyArray.push({
        date: key,
        amount: dailyCosts[key]
      });
    }
    dailyArray.sort(function(a, b) {
      return a.date > b.date ? 1 : -1;
    });

    var categoryArray = [];
    for (var key in categoryTotals) {
      categoryArray.push({
        category: key,
        amount: categoryTotals[key],
        percentage: totalCost > 0 ? (categoryTotals[key] / totalCost * 100).toFixed(2) : 0
      });
    }

    return {
      success: true,
      data: {
        totalCost: totalCost,
        totalSales: totalSales,
        profit: totalSales - totalCost,
        dailyCosts: dailyArray,
        categoryTotals: categoryArray
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// =====================================================
// ✅ จบไฟล์ Code.gs - อัพเดทครบถ้วน 100%
// =====================================================
