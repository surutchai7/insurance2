// ==================== Web App Entry Point ====================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Insurance Sale Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== Configuration ====================
// ใส่ Spreadsheet ID ของคุณที่นี่ (ต้องเปลี่ยนให้ตรงกับ Spreadsheet จริง)
const SPREADSHEET_ID = '1prUeqh_6pB84eqqqCbPFVxfOBpa7tFCv4RGA7GFLzks';

// Email settings
const ADMIN_EMAIL = 'your-email@gmail.com'; // เปลี่ยนเป็นอีเมลของคุณ
const EMAIL_ENABLED = false; // เปลี่ยนเป็น true เพื่อเปิดใช้งานการส่งอีเมล

// Service Type to Sheet Name mapping
const SERVICE_SHEET_MAPPING = {
  'ประกันภัย': 'ประกันภัย',
  'พรบ': 'ประกันภัย', // พรบ จะบันทึกในชีทเดียวกับประกันภัย
  'ภาษีรถ': 'ภาษี',
  'ค่าตรวจสภาพรถ': 'ค่าบริการ',
  'ค่าบริการ': 'ค่าบริการ',
  'อื่นๆ': 'อื่นๆ'
};

// Payment statuses
const PAYMENT_STATUS = {
  PENDING: 'รอชำระเงิน',
  PARTIAL: 'ชำระบางส่วน',
  COMPLETED: 'ชำระเต็มจำนวน',
  OVERDUE: 'เกินกำหนดชำระ',
  CANCELLED: 'ยกเลิก',
  INSTALLMENT: 'อยู่ระหว่างผ่อนชำระ' // เพิ่มสถานะใหม่
};

// Renewal statuses
const RENEWAL_STATUS = {
  PENDING: 'รอดำเนินการ',
  NOTIFIED: 'แจ้งผู้เอาประกันแล้ว',
  ORDERED: 'สั่งต่ออายุแล้ว',
  EXPIRED: 'กรมธรรม์ขาดต่อ'
};

// ==================== Error Handling Wrapper ====================
function safeExecute(functionName, fn, defaultReturn = null) {
  try {
    console.log(`Executing: ${functionName}`);
    const result = fn();
    console.log(`${functionName} completed successfully`);
    return result;
  } catch (error) {
    console.error(`Error in ${functionName}:`, error);
    console.error('Stack trace:', error.stack);
    return {
      success: false,
      error: error.toString(),
      message: `เกิดข้อผิดพลาดใน ${functionName}: ${error.toString()}`
    };
  }
}

// ==================== Sheet Management ====================
function getOrCreateSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log('Creating new sheet:', sheetName);
      sheet = ss.insertSheet(sheetName);
      const headers = getSheetHeaders(sheetName);
      
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
        
        // Auto-resize columns
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
      }
    }
    
    return sheet;
  } catch (error) {
    console.error('Error in getOrCreateSheet:', error);
    throw new Error(`ไม่สามารถเข้าถึงหรือสร้างชีท ${sheetName}: ${error.toString()}`);
  }
}

function getSheetHeaders(sheetName) {
  const headerMap = {
    'ลูกค้า': ['รหัสลูกค้า', 'ชื่อ-นามสกุล', 'เลขบัตรประชาชน', 'เบอร์โทร', 'อีเมล', 'ที่อยู่จัดส่ง', 'ที่อยู่ตามเอกสาร', 'วันที่สร้าง', 'สถานะ', 'หมายเหตุ'],
    'รถ': ['รหัสรถ', 'รหัสลูกค้า', 'ทะเบียนรถ', 'ยี่ห้อ', 'รุ่น', 'สี', 'ปีรถ', 'เลขตัวถัง', 'วันที่สร้าง', 'สถานะ'],
    'ทรัพย์สิน': ['รหัสทรัพย์สิน', 'รหัสลูกค้า', 'ประเภท', 'ชื่อทรัพย์สิน', 'ที่อยู่', 'มูลค่า', 'รายละเอียด', 'วันที่สร้าง', 'สถานะ', 'ประเภทรถ', 'รหัสรถ'],
    'ใบงาน': ['เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทประกัน', 'รหัสทรัพย์สิน', 'รายละเอียด', 'ทุนประกัน', 'เบี้ยประกัน', 'วันเริ่ม', 'วันสิ้นสุด', 'เลขกรมธรรม์', 'สถานะ', 'วันที่สร้าง', 'บริษัทประกัน', 'ผู้สร้าง', 'วันที่แก้ไข', 'สถานะการชำระเงิน', 'ประเภทการชำระ', 'จำนวนงวด', 'ยอดชำระแล้ว', 'ยอดค้างชำระ', 'ส่วนลด', 'ยอดชำระจริง'],
    'รายการใบงาน': ['รหัสรายการ', 'เลขที่ใบงาน', 'ประเภทบริการ', 'ประเภทประกัน', 'รายละเอียด', 'ทุนประกัน', 'เบี้ยประกัน', 'วันเริ่ม', 'วันสิ้นสุด', 'เลขกรมธรรม์', 'บริษัทประกัน', 'สถานะ', 'หมายเหตุ'],
    'บริษัทประกัน': ['รหัสบริษัท', 'ชื่อบริษัท', 'ที่อยู่', 'เบอร์โทร', 'ผู้ติดต่อ', 'สถานะ', 'อีเมล'],
    'ประวัติ': ['รหัสประวัติ', 'วันที่', 'ผู้ใช้', 'การกระทำ', 'รายละเอียด', 'ตารางที่เกี่ยวข้อง', 'รหัสอ้างอิง'],
    
    // Payment related sheets - อัพเดตเพิ่มฟิลด์ใหม่
    'การชำระเงิน': ['รหัสการชำระ', 'เลขที่ใบงาน', 'วันที่ชำระ', 'ประเภทการชำระ', 'จำนวนเงิน', 'หมายเลขอ้างอิง', 'ธนาคาร', 'หมายเหตุ', 'ผู้บันทึก', 'สถานะ', 'งวดที่'],
    'งวดการชำระ': ['รหัสงวด', 'เลขที่ใบงาน', 'งวดที่', 'จำนวนงวดทั้งหมด', 'จำนวนเงิน', 'วันครบกำหนด', 'วันที่ชำระ', 'สถานะ', 'หมายเลขอ้างอิง', 'หมายเหตุ', 'จำนวนเงินดาวน์', 'วันครบกำหนดเดิม'],
    
    // Service type specific sheets
    'ประกันภัย': ['รหัสรายการ', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทบริการ', 'ประเภทประกัน', 'รายละเอียด', 'ทุนประกัน', 'เบี้ยประกัน', 'วันเริ่ม', 'วันสิ้นสุด', 'เลขกรมธรรม์', 'บริษัทประกัน', 'สถานะ', 'วันที่สร้าง', 'หมายเหตุ', 'เลขตัวถัง'],
    'ภาษี': ['รหัสรายการ', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทบริการ', 'รายละเอียด', 'จำนวนเงิน', 'วันเริ่ม', 'วันสิ้นสุด', 'เลขที่อ้างอิง', 'สถานะ', 'วันที่สร้าง', 'หมายเหตุ', 'เลขตัวถัง'],
    'ค่าบริการ': ['รหัสรายการ', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทบริการ', 'รายละเอียด', 'จำนวนเงิน', 'วันที่ให้บริการ', 'สถานะ', 'วันที่สร้าง', 'หมายเหตุ', 'เลขตัวถัง'],
    'อื่นๆ': ['รหัสรายการ', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทบริการ', 'รายละเอียด', 'จำนวนเงิน', 'วันที่', 'สถานะ', 'วันที่สร้าง', 'หมายเหตุ', 'เลขตัวถัง'],
    
    // New sheets for tracking
    'ติดตามการผ่อนชำระ': ['รหัส', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'งวดที่', 'จำนวนเงิน', 'วันครบกำหนด', 'สถานะ', 'วันที่ชำระ', 'หมายเหตุ', 'ผู้ติดตาม'],
    'ติดตามการต่ออายุ': ['รหัส', 'เลขที่ใบงาน', 'รหัสลูกค้า', 'ชื่อลูกค้า', 'ประเภทประกัน', 'วันหมดอายุ', 'สถานะการต่ออายุ', 'วันที่แจ้ง', 'วันที่สั่งต่อ', 'เลขกรมธรรม์ใหม่', 'หมายเหตุ', 'ผู้ติดตาม']
  };
  
  return headerMap[sheetName] || [];
}

// ==================== Initialize System ====================
function initializeSystem() {
  return safeExecute('initializeSystem', function() {
    console.log('Initializing system...');
    
    // Test connection first
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('Connected to spreadsheet:', ss.getName());
    
    // Base sheets
    const baseSheets = ['ลูกค้า', 'รถ', 'ทรัพย์สิน', 'ใบงาน', 'รายการใบงาน', 'บริษัทประกัน', 'ประวัติ', 'การชำระเงิน', 'งวดการชำระ'];
    
    // Service type specific sheets
    const serviceSheets = ['ประกันภัย', 'ภาษี', 'ค่าบริการ', 'อื่นๆ'];
    
    // Tracking sheets
    const trackingSheets = ['ติดตามการผ่อนชำระ', 'ติดตามการต่ออายุ'];
    
    // Create all sheets
    [...baseSheets, ...serviceSheets, ...trackingSheets].forEach(sheetName => {
      console.log('Creating/checking sheet:', sheetName);
      getOrCreateSheet(sheetName);
    });
    
    // Add default insurance companies if needed
    initializeDefaultCompanies();
    
    // Migrate old work orders to items if needed
    const migrationResult = migrateWorkOrdersToItems();
    console.log('Migration result:', migrationResult);
    
    console.log('System initialization complete');
    return { 
      success: true, 
      message: 'ระบบพร้อมใช้งาน รองรับการแยกบันทึกตามประเภทบริการ การชำระเงิน และการติดตาม',
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      userEmail: Session.getActiveUser().getEmail(),
      migrationResult: migrationResult
    };
  });
}

function initializeDefaultCompanies() {
  try {
    const companySheet = getOrCreateSheet('บริษัทประกัน');
    if (companySheet.getLastRow() <= 1) {
      console.log('Adding default companies...');
      const defaultCompanies = [
        ['COM001', 'บริษัท ทิพยประกันภัย จำกัด (มหาชน)', 'กรุงเทพฯ', '02-123-4567', 'คุณสมชาย', 'Active', 'contact@dhipaya.co.th'],
        ['COM002', 'บริษัท วิริยะประกันภัย จำกัด (มหาชน)', 'กรุงเทพฯ', '02-234-5678', 'คุณสมหญิง', 'Active', 'contact@viriyah.co.th'],
        ['COM003', 'บริษัท กรุงเทพประกันภัย จำกัด (มหาชน)', 'กรุงเทพฯ', '02-345-6789', 'คุณสมศักดิ์', 'Active', 'contact@bangkokinsurance.com']
      ];
      companySheet.getRange(2, 1, defaultCompanies.length, defaultCompanies[0].length).setValues(defaultCompanies);
    }
  } catch (error) {
    console.error('Error initializing default companies:', error);
  }
}

// ==================== Payment Functions ====================
function generatePaymentId() {
  try {
    const sheet = getOrCreateSheet('การชำระเงิน');
    const lastRow = sheet.getLastRow();
    const today = new Date();
    const year = Utilities.formatDate(today, 'GMT+7', 'yyyy');
    const month = Utilities.formatDate(today, 'GMT+7', 'MM');
    return 'PAY' + year + month + String(lastRow).padStart(4, '0');
  } catch (error) {
    console.error('Error generating payment ID:', error);
    return 'PAY' + String(Date.now()).substr(-8);
  }
}

function generateInstallmentId() {
  try {
    const sheet = getOrCreateSheet('งวดการชำระ');
    const lastRow = sheet.getLastRow();
    return 'INST' + String(lastRow).padStart(6, '0');
  } catch (error) {
    console.error('Error generating installment ID:', error);
    return 'INST' + String(Date.now()).substr(-6);
  }
}

// Enhanced savePayment function
function savePayment(data) {
  return safeExecute('savePayment', function() {
    if (!data || !data.workOrderId || !data.amount || !data.paymentType) {
      throw new Error('ข้อมูลการชำระเงินไม่ครบถ้วน');
    }
    
    const paymentSheet = getOrCreateSheet('การชำระเงิน');
    const workOrderSheet = getOrCreateSheet('ใบงาน');
    const paymentId = generatePaymentId();
    const user = Session.getActiveUser().getEmail();
    const now = new Date();
    
    // Save payment record with installment number if applicable
    paymentSheet.appendRow([
      paymentId,
      data.workOrderId,
      now,
      data.paymentType,
      parseFloat(data.amount),
      data.referenceNumber || '',
      data.bank || '',
      data.notes || '',
      user,
      'สำเร็จ',
      data.installmentNumber || '' // เพิ่มเลขงวด
    ]);
    
    // Update work order payment status
    updateWorkOrderPaymentStatus(data.workOrderId);
    
    // Add to history
    addHistory('บันทึกการชำระเงิน', `ชำระเงิน ${data.paymentType} จำนวน ${data.amount} บาท`, 'การชำระเงิน', paymentId);
    
    return { 
      success: true, 
      message: 'บันทึกการชำระเงินสำเร็จ', 
      paymentId: paymentId 
    };
  });
}

// Enhanced createInstallmentPlan with custom down payment and dates
function createInstallmentPlan(data) {
  return safeExecute('createInstallmentPlan', function() {
    if (!data || !data.workOrderId || !data.totalAmount || !data.numberOfInstallments) {
      throw new Error('ข้อมูลการผ่อนชำระไม่ครบถ้วน');
    }
    
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const numberOfInstallments = parseInt(data.numberOfInstallments);
    
    if (numberOfInstallments < 2 || numberOfInstallments > 6) {
      throw new Error('จำนวนงวดต้องอยู่ระหว่าง 2-6 งวด');
    }
    
    const totalAmount = parseFloat(data.totalAmount);
    const downPayment = parseFloat(data.downPayment || 0);
    const remainingAmount = totalAmount - downPayment;
    const amountPerInstallment = Math.ceil(remainingAmount / (numberOfInstallments - 1));
    const startDate = new Date(data.startDate || new Date());
    const customDates = data.customDates || [];
    
    for (let i = 0; i < numberOfInstallments; i++) {
      const installmentId = generateInstallmentId();
      let dueDate;
      let amount;
      
      if (i === 0) {
        // งวดแรก (เงินดาวน์)
        amount = downPayment > 0 ? downPayment : amountPerInstallment;
        dueDate = startDate;
      } else {
        amount = amountPerInstallment;
        // ใช้วันที่กำหนดเองถ้ามี
        if (customDates[i]) {
          dueDate = new Date(customDates[i]);
        } else {
          dueDate = new Date(startDate);
          dueDate.setMonth(dueDate.getMonth() + i);
        }
      }
      
      installmentSheet.appendRow([
        installmentId,
        data.workOrderId,
        i + 1,
        numberOfInstallments,
        amount,
        dueDate,
        '', // วันที่ชำระ
        'รอชำระ',
        '', // หมายเลขอ้างอิง
        '', // หมายเหตุ
        i === 0 ? downPayment : '', // จำนวนเงินดาวน์
        dueDate // วันครบกำหนดเดิม (สำหรับการแก้ไข)
      ]);
    }
    
    // Update work order payment type
    updateWorkOrderPaymentType(data.workOrderId, 'ผ่อนชำระ', numberOfInstallments);
    
    // Add to tracking
    addInstallmentTracking(data.workOrderId);
    
    // Add to history
    addHistory('สร้างแผนผ่อนชำระ', `สร้างแผนผ่อนชำระ ${numberOfInstallments} งวด`, 'งวดการชำระ', data.workOrderId);
    
    return { 
      success: true, 
      message: `สร้างแผนผ่อนชำระ ${numberOfInstallments} งวด สำเร็จ` 
    };
  });
}

// Enhanced payInstallment function
function payInstallment(data) {
  return safeExecute('payInstallment', function() {
    if (!data || !data.installmentId || !data.amount) {
      throw new Error('ข้อมูลการชำระงวดไม่ครบถ้วน');
    }
    
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const dataRange = installmentSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.installmentId) {
        // Update installment record
        installmentSheet.getRange(i + 1, 7).setValue(new Date()); // วันที่ชำระ
        installmentSheet.getRange(i + 1, 8).setValue('ชำระแล้ว'); // สถานะ
        installmentSheet.getRange(i + 1, 9).setValue(data.referenceNumber || ''); // หมายเลขอ้างอิง
        installmentSheet.getRange(i + 1, 10).setValue(data.notes || ''); // หมายเหตุ
        
        // Record payment with installment number
        const paymentData = {
          workOrderId: values[i][1],
          amount: data.amount,
          paymentType: 'ผ่อนชำระ',
          referenceNumber: data.referenceNumber,
          bank: data.bank,
          notes: `ชำระงวดที่ ${values[i][2]}/${values[i][3]}`,
          installmentNumber: values[i][2]
        };
        
        savePayment(paymentData);
        
        // Update tracking
        updateInstallmentTracking(values[i][1], values[i][2]);
        
        return { 
          success: true, 
          message: `ชำระงวดที่ ${values[i][2]} สำเร็จ` 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'ไม่พบข้อมูลงวดการชำระ' 
    };
  });
}

// New function to update installment due date
function updateInstallmentDueDate(data) {
  return safeExecute('updateInstallmentDueDate', function() {
    if (!data || !data.installmentId || !data.newDueDate) {
      throw new Error('ข้อมูลไม่ครบถ้วน');
    }
    
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const dataRange = installmentSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.installmentId) {
        const oldDate = values[i][5];
        installmentSheet.getRange(i + 1, 6).setValue(new Date(data.newDueDate)); // วันครบกำหนดใหม่
        installmentSheet.getRange(i + 1, 12).setValue(oldDate); // เก็บวันเดิมไว้
        
        // Add to history
        addHistory('แก้ไขวันครบกำหนด', `เปลี่ยนวันครบกำหนดงวดที่ ${values[i][2]}`, 'งวดการชำระ', values[i][1]);
        
        return { 
          success: true, 
          message: 'อัพเดตวันครบกำหนดสำเร็จ' 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'ไม่พบข้อมูลงวดการชำระ' 
    };
  });
}

// New function to update installment amount
function updateInstallmentAmount(data) {
  return safeExecute('updateInstallmentAmount', function() {
    if (!data || !data.installmentId || !data.newAmount) {
      throw new Error('ข้อมูลไม่ครบถ้วน');
    }
    
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const dataRange = installmentSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.installmentId) {
        const oldAmount = values[i][4];
        installmentSheet.getRange(i + 1, 5).setValue(parseFloat(data.newAmount)); // จำนวนเงินใหม่
        
        // Recalculate remaining installments if first installment
        if (values[i][2] === 1 && data.recalculateRemaining) {
          const workOrderId = values[i][1];
          const totalInstallments = values[i][3];
          
          // Get work order total
          const workOrder = getWorkOrderById(workOrderId);
          if (workOrder) {
            const totalAmount = workOrder.actualAmount || workOrder.premium;
            const downPayment = parseFloat(data.newAmount);
            const remainingAmount = totalAmount - downPayment;
            const remainingInstallments = totalInstallments - 1;
            const newAmountPerInstallment = Math.ceil(remainingAmount / remainingInstallments);
            
            // Update remaining installments
            for (let j = i + 1; j < values.length && values[j][1] === workOrderId; j++) {
              installmentSheet.getRange(j + 1, 5).setValue(newAmountPerInstallment);
            }
          }
        }
        
        // Add to history
        addHistory('แก้ไขจำนวนเงิน', `เปลี่ยนจำนวนเงินงวดที่ ${values[i][2]} จาก ${oldAmount} เป็น ${data.newAmount}`, 'งวดการชำระ', values[i][1]);
        
        return { 
          success: true, 
          message: 'อัพเดตจำนวนเงินสำเร็จ' 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'ไม่พบข้อมูลงวดการชำระ' 
    };
  });
}

function getPaymentHistory(workOrderId) {
  return safeExecute('getPaymentHistory', function() {
    if (!workOrderId) {
      throw new Error('ไม่พบเลขที่ใบงาน');
    }
    
    const paymentSheet = getOrCreateSheet('การชำระเงิน');
    const lastRow = paymentSheet.getLastRow();
    
    console.log('Getting payment history for:', workOrderId, 'Last row:', lastRow);
    
    if (lastRow <= 1) return [];
    
    const values = paymentSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    
    const payments = values
      .filter(row => {
        const rowWorkOrderId = String(row[1] || '');
        const matches = rowWorkOrderId === String(workOrderId);
        if (matches) {
          console.log('Found payment:', row[0], 'Amount:', row[4]);
        }
        return matches;
      })
      .map(row => ({
        paymentId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        paymentDate: row[2] ? (row[2] instanceof Date ? row[2].toISOString().split('T')[0] : row[2]) : '',
        paymentType: String(row[3] || ''),
        amount: parseFloat(row[4]) || 0,
        referenceNumber: String(row[5] || ''),
        bank: String(row[6] || ''),
        notes: String(row[7] || ''),
        recordedBy: String(row[8] || ''),
        status: String(row[9] || ''),
        installmentNumber: row[10] || ''
      }));
    
    console.log('Found', payments.length, 'payments for work order', workOrderId);
    return payments;
  }, []);
}

function getInstallmentPlan(workOrderId) {
  return safeExecute('getInstallmentPlan', function() {
    if (!workOrderId) {
      throw new Error('ไม่พบเลขที่ใบงาน');
    }
    
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const lastRow = installmentSheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const values = installmentSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    const installments = values
      .filter(row => row[1] === workOrderId)
      .map(row => ({
        installmentId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        installmentNumber: parseInt(row[2]) || 0,
        totalInstallments: parseInt(row[3]) || 0,
        amount: parseFloat(row[4]) || 0,
        dueDate: row[5] || '',
        paymentDate: row[6] || '',
        status: String(row[7] || ''),
        referenceNumber: String(row[8] || ''),
        notes: String(row[9] || ''),
        downPaymentAmount: row[10] || '',
        originalDueDate: row[11] || ''
      }));
    
    return installments;
  });
}

// Enhanced updateWorkOrderPaymentStatus
function updateWorkOrderPaymentStatus(workOrderId) {
  try {
    const workOrderSheet = getOrCreateSheet('ใบงาน');
    const paymentSheet = getOrCreateSheet('การชำระเงิน');
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    
    // Get work order data
    const workOrderData = getWorkOrderById(workOrderId);
    if (!workOrderData) return;
    
    // Get all payments for this work order
    const payments = getPaymentHistory(workOrderId);
    
    // Calculate total paid
    const totalPaid = payments.reduce((sum, payment) => sum + payment.amount, 0);
    const totalAmount = workOrderData.actualAmount || workOrderData.premium || 0;
    const remainingAmount = totalAmount - totalPaid;
    
    // Determine payment status
    let paymentStatus = PAYMENT_STATUS.PENDING;
    
    if (workOrderData.paymentType === 'ผ่อนชำระ') {
      // Check installment status
      const installments = getInstallmentPlan(workOrderId);
      const allPaid = installments.every(inst => inst.status === 'ชำระแล้ว');
      const somePaid = installments.some(inst => inst.status === 'ชำระแล้ว');
      
      if (allPaid) {
        paymentStatus = PAYMENT_STATUS.COMPLETED;
      } else if (somePaid) {
        paymentStatus = PAYMENT_STATUS.INSTALLMENT;
      } else {
        paymentStatus = PAYMENT_STATUS.PENDING;
      }
    } else {
      // ชำระเต็มจำนวน
      if (totalPaid >= totalAmount) {
        paymentStatus = PAYMENT_STATUS.COMPLETED;
      } else if (totalPaid > 0) {
        paymentStatus = PAYMENT_STATUS.PARTIAL;
      }
    }
    
    // Update work order
    const dataRange = workOrderSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === workOrderId) {
        workOrderSheet.getRange(i + 1, 17).setValue(paymentStatus); // สถานะการชำระเงิน
        workOrderSheet.getRange(i + 1, 20).setValue(totalPaid); // ยอดชำระแล้ว
        workOrderSheet.getRange(i + 1, 21).setValue(remainingAmount); // ยอดค้างชำระ
        break;
      }
    }
  } catch (error) {
    console.error('Error updating work order payment status:', error);
  }
}

function updateWorkOrderPaymentType(workOrderId, paymentType, numberOfInstallments) {
  try {
    const workOrderSheet = getOrCreateSheet('ใบงาน');
    const dataRange = workOrderSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === workOrderId) {
        workOrderSheet.getRange(i + 1, 18).setValue(paymentType); // ประเภทการชำระ
        workOrderSheet.getRange(i + 1, 19).setValue(numberOfInstallments || ''); // จำนวนงวด
        break;
      }
    }
  } catch (error) {
    console.error('Error updating work order payment type:', error);
  }
}

// Enhanced getWorkOrderById with discount and actual amount
function getWorkOrderById(workOrderId) {
  try {
    const sheet = getOrCreateSheet('ใบงาน');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === workOrderId) {
        return {
          id: String(values[i][0] || ''),
          customerId: String(values[i][1] || ''),
          customerName: String(values[i][2] || ''),
          insuranceType: String(values[i][3] || ''),
          propertyId: String(values[i][4] || ''),
          details: String(values[i][5] || ''),
          sumInsured: parseFloat(values[i][6]) || 0,
          premium: parseFloat(values[i][7]) || 0,
          startDate: values[i][8] || '',
          endDate: values[i][9] || '',
          policyNumber: String(values[i][10] || ''),
          status: String(values[i][11] || ''),
          createdDate: values[i][12] || '',
          insuranceCompany: String(values[i][13] || ''),
          createdBy: String(values[i][14] || ''),
          modifiedDate: values[i][15] || '',
          paymentStatus: String(values[i][16] || PAYMENT_STATUS.PENDING),
          paymentType: String(values[i][17] || ''),
          numberOfInstallments: values[i][18] || '',
          totalPaid: parseFloat(values[i][19]) || 0,
          remainingAmount: parseFloat(values[i][20]) || 0,
          discount: parseFloat(values[i][21]) || 0,
          actualAmount: parseFloat(values[i][22]) || parseFloat(values[i][7]) || 0
        };
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting work order by ID:', error);
    return null;
  }
}

function getPaymentSummary(workOrderId) {
  return safeExecute('getPaymentSummary', function() {
    const workOrder = getWorkOrderById(workOrderId);
    if (!workOrder) {
      throw new Error('ไม่พบข้อมูลใบงาน');
    }
    
    const payments = getPaymentHistory(workOrderId);
    const installments = getInstallmentPlan(workOrderId);
    
    const totalAmount = workOrder.actualAmount || workOrder.premium || 0;
    const totalPaid = payments.reduce((sum, payment) => sum + payment.amount, 0);
    const remainingAmount = totalAmount - totalPaid;
    
    // Check overdue installments
    const today = new Date();
    const overdueInstallments = installments.filter(inst => {
      if (inst.status === 'ชำระแล้ว') return false;
      const dueDate = new Date(inst.dueDate);
      return dueDate < today;
    });
    
    return {
      workOrderId: workOrderId,
      totalAmount: totalAmount,
      discount: workOrder.discount || 0,
      actualAmount: workOrder.actualAmount || totalAmount,
      totalPaid: totalPaid,
      remainingAmount: remainingAmount,
      paymentStatus: workOrder.paymentStatus,
      paymentType: workOrder.paymentType,
      numberOfInstallments: workOrder.numberOfInstallments,
      payments: payments,
      installments: installments,
      overdueInstallments: overdueInstallments,
      isFullyPaid: totalPaid >= (workOrder.actualAmount || totalAmount)
    };
  });
}

// ==================== Tracking Functions ====================
// Add installment tracking
function addInstallmentTracking(workOrderId) {
  try {
    const trackingSheet = getOrCreateSheet('ติดตามการผ่อนชำระ');
    const installments = getInstallmentPlan(workOrderId);
    const workOrder = getWorkOrderById(workOrderId);
    
    if (!workOrder) return;
    
    installments.forEach(inst => {
      if (inst.status !== 'ชำระแล้ว') {
        const trackingId = 'TRK-INS' + String(trackingSheet.getLastRow()).padStart(5, '0');
        
        trackingSheet.appendRow([
          trackingId,
          workOrderId,
          workOrder.customerId,
          workOrder.customerName,
          inst.installmentNumber,
          inst.amount,
          inst.dueDate,
          'รอชำระ',
          '', // วันที่ชำระ
          '', // หมายเหตุ
          Session.getActiveUser().getEmail()
        ]);
      }
    });
  } catch (error) {
    console.error('Error adding installment tracking:', error);
  }
}

// Update installment tracking
function updateInstallmentTracking(workOrderId, installmentNumber) {
  try {
    const trackingSheet = getOrCreateSheet('ติดตามการผ่อนชำระ');
    const dataRange = trackingSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === workOrderId && values[i][4] === installmentNumber) {
        trackingSheet.getRange(i + 1, 8).setValue('ชำระแล้ว');
        trackingSheet.getRange(i + 1, 9).setValue(new Date());
        break;
      }
    }
  } catch (error) {
    console.error('Error updating installment tracking:', error);
  }
}

// Get installments due this month
function getInstallmentsDueThisMonth() {
  return safeExecute('getInstallmentsDueThisMonth', function() {
    const trackingSheet = getOrCreateSheet('ติดตามการผ่อนชำระ');
    const lastRow = trackingSheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const values = trackingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    const dueInstallments = values
      .filter(row => {
        if (row[7] === 'ชำระแล้ว') return false;
        const dueDate = new Date(row[6]);
        return dueDate.getMonth() === currentMonth && dueDate.getFullYear() === currentYear;
      })
      .map(row => ({
        trackingId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        installmentNumber: row[4] || 0,
        amount: parseFloat(row[5]) || 0,
        dueDate: row[6] || '',
        status: String(row[7] || ''),
        paymentDate: row[8] || '',
        notes: String(row[9] || ''),
        trackedBy: String(row[10] || '')
      }));
    
    return dueInstallments;
  });
}

// Add renewal tracking
function addRenewalTracking(workOrderId) {
  try {
    const trackingSheet = getOrCreateSheet('ติดตามการต่ออายุ');
    const workOrder = getWorkOrderById(workOrderId);
    
    if (!workOrder || !workOrder.endDate) return;
    
    const trackingId = 'TRK-REN' + String(trackingSheet.getLastRow()).padStart(5, '0');
    
    trackingSheet.appendRow([
      trackingId,
      workOrderId,
      workOrder.customerId,
      workOrder.customerName,
      workOrder.insuranceType,
      workOrder.endDate,
      RENEWAL_STATUS.PENDING,
      '', // วันที่แจ้ง
      '', // วันที่สั่งต่อ
      '', // เลขกรมธรรม์ใหม่
      '', // หมายเหตุ
      Session.getActiveUser().getEmail()
    ]);
  } catch (error) {
    console.error('Error adding renewal tracking:', error);
  }
}

// Update renewal status
function updateRenewalStatus(data) {
  return safeExecute('updateRenewalStatus', function() {
    if (!data || !data.trackingId || !data.status) {
      throw new Error('ข้อมูลไม่ครบถ้วน');
    }
    
    const trackingSheet = getOrCreateSheet('ติดตามการต่ออายุ');
    const dataRange = trackingSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.trackingId) {
        trackingSheet.getRange(i + 1, 7).setValue(data.status);
        
        // Update dates based on status
        if (data.status === RENEWAL_STATUS.NOTIFIED) {
          trackingSheet.getRange(i + 1, 8).setValue(new Date());
        } else if (data.status === RENEWAL_STATUS.ORDERED) {
          trackingSheet.getRange(i + 1, 9).setValue(new Date());
          if (data.newPolicyNumber) {
            trackingSheet.getRange(i + 1, 10).setValue(data.newPolicyNumber);
          }
        }
        
        if (data.notes) {
          trackingSheet.getRange(i + 1, 11).setValue(data.notes);
        }
        
        // Add to history
        addHistory('อัพเดตสถานะการต่ออายุ', `เปลี่ยนสถานะเป็น ${data.status}`, 'ติดตามการต่ออายุ', values[i][1]);
        
        return { 
          success: true, 
          message: 'อัพเดตสถานะสำเร็จ' 
        };
      }
    }
    
    return { 
      success: false, 
      message: 'ไม่พบข้อมูลการติดตาม' 
    };
  });
}

// Get renewals due this month
function getRenewalsDueThisMonth() {
  return safeExecute('getRenewalsDueThisMonth', function() {
    const trackingSheet = getOrCreateSheet('ติดตามการต่ออายุ');
    const lastRow = trackingSheet.getLastRow();
    
    if (lastRow <= 1) {
      // If no tracking records, check work orders directly
      const workOrders = getWorkOrderList();
      const today = new Date();
      const currentMonth = today.getMonth();
      const currentYear = today.getFullYear();
      const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      
      const dueRenewals = workOrders
        .filter(wo => {
          if (!wo.endDate || wo.status === 'ยกเลิก') return false;
          const endDate = new Date(wo.endDate);
          return endDate >= today && endDate <= nextMonth;
        })
        .map(wo => {
          // Auto-add to tracking
          addRenewalTracking(wo.id);
          
          return {
            workOrderId: wo.id,
            customerId: wo.customerId,
            customerName: wo.customerName,
            insuranceType: wo.insuranceType,
            expiryDate: wo.endDate,
            status: RENEWAL_STATUS.PENDING,
            policyNumber: wo.policyNumber
          };
        });
      
      return dueRenewals;
    }
    
    // Get from tracking sheet
    const values = trackingSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    
    const dueRenewals = values
      .filter(row => {
        const expiryDate = new Date(row[5]);
        return expiryDate >= today && expiryDate <= nextMonth;
      })
      .map(row => ({
        trackingId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        insuranceType: String(row[4] || ''),
        expiryDate: row[5] || '',
        status: String(row[6] || ''),
        notificationDate: row[7] || '',
        renewalOrderDate: row[8] || '',
        newPolicyNumber: String(row[9] || ''),
        notes: String(row[10] || ''),
        trackedBy: String(row[11] || '')
      }));
    
    return dueRenewals;
  });
}

// Get installments with month/year filter
function getInstallmentsByMonthYear(month, year) {
  return safeExecute('getInstallmentsByMonthYear', function() {
    const trackingSheet = getOrCreateSheet('ติดตามการผ่อนชำระ');
    const lastRow = trackingSheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const values = trackingSheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const targetMonth = parseInt(month) - 1; // JavaScript months are 0-based
    const targetYear = parseInt(year);
    
    const dueInstallments = values
      .filter(row => {
        const dueDate = new Date(row[6]);
        return dueDate.getMonth() === targetMonth && dueDate.getFullYear() === targetYear;
      })
      .map(row => ({
        trackingId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        installmentNumber: row[4] || 0,
        amount: parseFloat(row[5]) || 0,
        dueDate: row[6] || '',
        status: String(row[7] || ''),
        paymentDate: row[8] || '',
        notes: String(row[9] || ''),
        trackedBy: String(row[10] || ''),
        totalInstallments: 0 // Will be filled from installment plan
      }));
    
    // Get total installments from installment plan
    const installmentSheet = getOrCreateSheet('งวดการชำระ');
    const installmentData = installmentSheet.getDataRange().getValues();
    
    dueInstallments.forEach(inst => {
      for (let i = 1; i < installmentData.length; i++) {
        if (installmentData[i][1] === inst.workOrderId && installmentData[i][2] === inst.installmentNumber) {
          inst.totalInstallments = installmentData[i][3] || 0;
          break;
        }
      }
    });
    
    return dueInstallments;
  });
}

// Get renewals with month/year filter
function getRenewalsByMonthYear(month, year) {
  return safeExecute('getRenewalsByMonthYear', function() {
    const trackingSheet = getOrCreateSheet('ติดตามการต่ออายุ');
    const lastRow = trackingSheet.getLastRow();
    
    if (lastRow <= 1) {
      // Check work orders directly
      const workOrders = getWorkOrderList();
      const targetMonth = parseInt(month) - 1;
      const targetYear = parseInt(year);
      
      const dueRenewals = workOrders
        .filter(wo => {
          if (!wo.endDate || wo.status === 'ยกเลิก') return false;
          const endDate = new Date(wo.endDate);
          return endDate.getMonth() === targetMonth && endDate.getFullYear() === targetYear;
        })
        .map(wo => ({
          workOrderId: wo.id,
          customerId: wo.customerId,
          customerName: wo.customerName,
          insuranceType: wo.insuranceType,
          expiryDate: wo.endDate,
          status: RENEWAL_STATUS.PENDING,
          policyNumber: wo.policyNumber
        }));
      
      return dueRenewals;
    }
    
    // Get from tracking sheet
    const values = trackingSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const targetMonth = parseInt(month) - 1;
    const targetYear = parseInt(year);
    
    const dueRenewals = values
      .filter(row => {
        const expiryDate = new Date(row[5]);
        return expiryDate.getMonth() === targetMonth && expiryDate.getFullYear() === targetYear;
      })
      .map(row => ({
        trackingId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        insuranceType: String(row[4] || ''),
        expiryDate: row[5] || '',
        status: String(row[6] || ''),
        notificationDate: row[7] || '',
        renewalOrderDate: row[8] || '',
        newPolicyNumber: String(row[9] || ''),
        notes: String(row[10] || ''),
        trackedBy: String(row[11] || '')
      }));
    
    return dueRenewals;
  });
}

// ==================== Utility Functions ====================
function calculateEndDate(startDate) {
  if (!startDate) return '';
  
  try {
    const start = new Date(startDate);
    const end = new Date(start);
    end.setFullYear(end.getFullYear() + 1); // เพิ่ม 1 ปี
    
    return end.toISOString().split('T')[0]; // Return YYYY-MM-DD format
  } catch (error) {
    console.error('Error calculating end date:', error);
    return '';
  }
}

function formatDateThai(date) {
  if (!date) return '-';
  try {
    const d = new Date(date);
    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear() + 543}`;
  } catch (error) {
    console.error('Error formatting date:', error);
    return '-';
  }
}

// ==================== Enhanced Search Functions ====================
function searchWorkOrders(searchText) {
  return safeExecute('searchWorkOrders', function() {
    console.log('Searching work orders with text:', searchText);
    
    if (!searchText || searchText.trim() === '') {
      console.log('Empty search text, returning empty results');
      return { success: true, data: [], count: 0 };
    }
    
    const sheet = getOrCreateSheet('ใบงาน');
    const lastRow = sheet.getLastRow();
    
    console.log('Work orders sheet last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('No data in work orders sheet');
      return { success: true, data: [], count: 0 };
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
    const searchLower = searchText.toLowerCase();
    
    console.log('Total rows to search:', values.length);
    
    const results = values.filter(row => {
      // Convert all fields to string for safe searching
      const workOrderId = String(row[0] || '').toLowerCase();
      const customerName = String(row[2] || '').toLowerCase();
      const policyNumber = String(row[10] || '').toLowerCase();
      const insuranceType = String(row[3] || '').toLowerCase();
      const details = String(row[5] || '').toLowerCase();
      const status = String(row[11] || '').toLowerCase();
      const paymentStatus = String(row[16] || '').toLowerCase();
      
      const matches = workOrderId.includes(searchLower) || 
                     customerName.includes(searchLower) || 
                     policyNumber.includes(searchLower) ||
                     insuranceType.includes(searchLower) ||
                     details.includes(searchLower) ||
                     status.includes(searchLower) ||
                     paymentStatus.includes(searchLower);
      
      return matches;
    }).map(row => ({
      id: String(row[0] || ''),
      customerId: String(row[1] || ''),
      customerName: String(row[2] || ''),
      insuranceType: String(row[3] || ''),
      propertyId: String(row[4] || ''),
      details: String(row[5] || ''),
      sumInsured: parseFloat(row[6]) || 0,
      premium: parseFloat(row[7]) || 0,
      startDate: row[8] ? (row[8] instanceof Date ? row[8].toISOString().split('T')[0] : row[8]) : '',
      endDate: row[9] ? (row[9] instanceof Date ? row[9].toISOString().split('T')[0] : row[9]) : '',
      policyNumber: String(row[10] || ''),
      status: String(row[11] || ''),
      createdDate: row[12] ? (row[12] instanceof Date ? row[12].toISOString().split('T')[0] : row[12]) : '',
      insuranceCompany: String(row[13] || ''),
      createdBy: String(row[14] || ''),
      modifiedDate: row[15] ? (row[15] instanceof Date ? row[15].toISOString().split('T')[0] : row[15]) : '',
      paymentStatus: String(row[16] || PAYMENT_STATUS.PENDING),
      paymentType: String(row[17] || ''),
      numberOfInstallments: row[18] || '',
      totalPaid: parseFloat(row[19]) || 0,
      remainingAmount: parseFloat(row[20]) || 0,
      discount: parseFloat(row[21]) || 0,
      actualAmount: parseFloat(row[22]) || parseFloat(row[7]) || 0
    }));
    
    console.log('Search found', results.length, 'matches');
    
    return {
      success: true,
      data: results,
      count: results.length
    };
  }, { success: false, data: [], count: 0, error: 'Search failed' });
}

function searchCustomers(searchText) {
  return safeExecute('searchCustomers', function() {
    if (!searchText || searchText.trim() === '') {
      return { success: true, data: [], count: 0 };
    }
    
    const sheet = getOrCreateSheet('ลูกค้า');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: true, data: [], count: 0 };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const searchLower = searchText.toLowerCase();
    
    const results = values.filter(row => {
      const name = String(row[1] || '').toLowerCase();
      const idCard = String(row[2] || '').toLowerCase();
      const phone = String(row[3] || '').toLowerCase();
      const email = String(row[4] || '').toLowerCase();
      
      return name.includes(searchLower) || 
             idCard.includes(searchLower) || 
             phone.includes(searchLower) || 
             email.includes(searchLower);
    }).map(row => ({
      id: String(row[0] || ''),
      name: String(row[1] || ''),
      idCard: String(row[2] || ''),
      phone: String(row[3] || ''),
      email: String(row[4] || ''),
      shippingAddress: String(row[5] || ''),
      documentAddress: String(row[6] || ''),
      createdDate: row[7] ? new Date(row[7]).toISOString() : '',
      status: String(row[8] || 'Active'),
      notes: String(row[9] || '')
    }));
    
    return {
      success: true,
      data: results,
      count: results.length
    };
  }, { success: false, data: [], count: 0, error: 'Search failed' });
}

// ==================== Customer Functions ====================
function generateCustomerId() {
  try {
    const sheet = getOrCreateSheet('ลูกค้า');
    const lastRow = sheet.getLastRow();
    return 'CUS' + String(lastRow).padStart(5, '0');
  } catch (error) {
    console.error('Error generating customer ID:', error);
    return 'CUS' + String(Date.now()).substr(-5);
  }
}

function saveCustomer(data) {
  return safeExecute('saveCustomer', function() {
    console.log('Saving customer:', data);
    
    if (!data || !data.name || data.name.trim() === '') {
      throw new Error('ชื่อลูกค้าเป็นข้อมูลที่จำเป็น');
    }
    
    const sheet = getOrCreateSheet('ลูกค้า');
    const customerId = data.id || generateCustomerId();
    const now = new Date();
    
    if (data.id) {
      // Update existing customer
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === data.id) {
          sheet.getRange(i + 1, 1, 1, 10).setValues([[
            customerId,
            data.name || '',
            data.idCard || '',
            data.phone || '',
            data.email || '',
            data.shippingAddress || '',
            data.documentAddress || '',
            values[i][7], // Keep original creation date
            data.status || 'Active',
            data.notes || ''
          ]]);
          
          // Add to history
          addHistory('แก้ไขข้อมูลลูกค้า', `แก้ไขข้อมูล ${data.name}`, 'ลูกค้า', customerId);
          
          console.log('Customer updated successfully');
          return { success: true, message: 'อัพเดตข้อมูลลูกค้าสำเร็จ', id: customerId };
        }
      }
    }
    
    // Add new customer
    const newRow = [
      customerId,
      data.name || '',
      data.idCard || '',
      data.phone || '',
      data.email || '',
      data.shippingAddress || '',
      data.documentAddress || '',
      now,
      'Active',
      data.notes || ''
    ];
    
    sheet.appendRow(newRow);
    
    // Add to history
    addHistory('เพิ่มลูกค้าใหม่', `เพิ่ม ${data.name}`, 'ลูกค้า', customerId);
    
    console.log('Customer added successfully:', customerId);
    
    return { success: true, message: 'บันทึกข้อมูลลูกค้าสำเร็จ', id: customerId };
  });
}

function deleteCustomer(customerId) {
  return safeExecute('deleteCustomer', function() {
    if (!customerId) {
      throw new Error('ไม่พบรหัสลูกค้า');
    }
    
    const sheet = getOrCreateSheet('ลูกค้า');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === customerId) {
        // Soft delete - just update status
        sheet.getRange(i + 1, 9).setValue('Deleted');
        
        // Add to history
        addHistory('ลบลูกค้า', `ลบ ${values[i][1]}`, 'ลูกค้า', customerId);
        
        return { success: true, message: 'ลบข้อมูลลูกค้าสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบข้อมูลลูกค้า' };
  });
}

function getCustomerListWrapped() {
  return safeExecute('getCustomerListWrapped', function() {
    console.log('Getting customer list...');
    const sheet = getOrCreateSheet('ลูกค้า');
    const lastRow = sheet.getLastRow();
    console.log('Customer sheet last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('No customer data found');
      return {
        success: true,
        data: [],
        count: 0,
        message: 'No customers found'
      };
    }
    
    const numRows = lastRow - 1;
    const numCols = 10;
    const dataRange = sheet.getRange(2, 1, numRows, numCols);
    const values = dataRange.getValues();
    console.log('Retrieved', values.length, 'customer rows');
    
    const customers = values
      .filter(row => row[8] !== 'Deleted') // Filter out deleted customers
      .map((row, index) => {
        return {
          id: String(row[0] || ''),
          name: String(row[1] || ''),
          idCard: String(row[2] || ''),
          phone: String(row[3] || ''),
          email: String(row[4] || ''),
          shippingAddress: String(row[5] || ''),
          documentAddress: String(row[6] || ''),
          createdDate: row[7] ? new Date(row[7]).toISOString() : '',
          status: String(row[8] || 'Active'),
          notes: String(row[9] || '')
        };
      });
    
    console.log('Processed', customers.length, 'customers');
    
    return {
      success: true,
      data: customers,
      count: customers.length,
      message: 'Customers loaded successfully'
    };
  }, {
    success: false,
    data: [],
    count: 0,
    error: 'Failed to load customers',
    message: 'Failed to load customers'
  });
}

function searchCustomersForAutocomplete(searchText) {
  return safeExecute('searchCustomersForAutocomplete', function() {
    if (!searchText || searchText.length < 2) {
      return { success: true, data: [] };
    }
    
    const sheet = getOrCreateSheet('ลูกค้า');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: true, data: [] };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const searchLower = searchText.toLowerCase();
    
    const results = values
      .filter(row => {
        const name = String(row[1] || '').toLowerCase();
        const phone = String(row[3] || '').toLowerCase();
        const idCard = String(row[2] || '').toLowerCase();
        
        return row[8] !== 'Deleted' && (
          name.includes(searchLower) || 
          phone.includes(searchLower) ||
          idCard.includes(searchLower)
        );
      })
      .slice(0, 10) // Limit to 10 results
      .map(row => ({
        id: String(row[0] || ''),
        name: String(row[1] || ''),
        phone: String(row[3] || ''),
        idCard: String(row[2] || ''),
        label: `${row[1]} - ${row[3]}` // Display format
      }));
    
    return {
      success: true,
      data: results
    };
  }, { success: false, data: [], error: 'Autocomplete failed' });
}

// ==================== Vehicle Functions ====================
function generateVehicleId() {
  try {
    const sheet = getOrCreateSheet('รถ');
    const lastRow = sheet.getLastRow();
    return 'VEH' + String(lastRow).padStart(5, '0');
  } catch (error) {
    console.error('Error generating vehicle ID:', error);
    return 'VEH' + String(Date.now()).substr(-5);
  }
}

function saveVehicle(data) {
  return safeExecute('saveVehicle', function() {
    if (!data || !data.customerId || !data.licensePlate) {
      throw new Error('ข้อมูลเจ้าของและทะเบียนรถเป็นข้อมูลที่จำเป็น');
    }
    
    const sheet = getOrCreateSheet('รถ');
    const vehicleId = data.id || generateVehicleId();
    const now = new Date();
    
    if (data.id) {
      // Update existing vehicle
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === data.id) {
          sheet.getRange(i + 1, 1, 1, 10).setValues([[
            vehicleId,
            data.customerId,
            data.licensePlate,
            data.brand || '',
            data.model || '',
            data.color || '',
            data.year || '',
            data.chassisNumber || '',
            values[i][8], // Keep original creation date
            'Active'
          ]]);
          
          // Add to history
          addHistory('แก้ไขข้อมูลรถ', `แก้ไขทะเบียน ${data.licensePlate}`, 'รถ', vehicleId);
          
          return { success: true, message: 'อัพเดตข้อมูลรถสำเร็จ', id: vehicleId };
        }
      }
    }
    
    // Add new vehicle
    sheet.appendRow([
      vehicleId,
      data.customerId,
      data.licensePlate,
      data.brand || '',
      data.model || '',
      data.color || '',
      data.year || '',
      data.chassisNumber || '',
      now,
      'Active'
    ]);
    
    // Add to history
    addHistory('เพิ่มรถใหม่', `เพิ่มทะเบียน ${data.licensePlate}`, 'รถ', vehicleId);
    
    return { success: true, message: 'บันทึกข้อมูลรถสำเร็จ', id: vehicleId };
  });
}

function deleteVehicle(vehicleId) {
  return safeExecute('deleteVehicle', function() {
    if (!vehicleId) {
      throw new Error('ไม่พบรหัสรถ');
    }
    
    const sheet = getOrCreateSheet('รถ');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === vehicleId) {
        // Soft delete
        sheet.getRange(i + 1, 10).setValue('Deleted');
        
        // Add to history
        addHistory('ลบรถ', `ลบทะเบียน ${values[i][2]}`, 'รถ', vehicleId);
        
        return { success: true, message: 'ลบข้อมูลรถสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบข้อมูลรถ' };
  });
}

function getVehicleListWrapped() {
  return safeExecute('getVehicleListWrapped', function() {
    const sheet = getOrCreateSheet('รถ');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: [],
        count: 0
      };
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const vehicles = values
      .filter(row => row[9] !== 'Deleted')
      .map(row => ({
        id: String(row[0] || ''),
        customerId: String(row[1] || ''),
        licensePlate: String(row[2] || ''),
        brand: String(row[3] || ''),
        model: String(row[4] || ''),
        color: String(row[5] || ''),
        year: String(row[6] || ''),
        chassisNumber: String(row[7] || ''),
        createdDate: row[8] ? new Date(row[8]).toISOString() : '',
        status: String(row[9] || 'Active')
      }));
    
    return {
      success: true,
      data: vehicles,
      count: vehicles.length
    };
  }, {
    success: false,
    data: [],
    count: 0,
    error: 'Failed to load vehicles'
  });
}

// ==================== Property Functions ====================
function generatePropertyId() {
  try {
    const sheet = getOrCreateSheet('ทรัพย์สิน');
    const lastRow = sheet.getLastRow();
    return 'PROP' + String(lastRow).padStart(5, '0');
  } catch (error) {
    console.error('Error generating property ID:', error);
    return 'PROP' + String(Date.now()).substr(-5);
  }
}

function saveProperty(data) {
  return safeExecute('saveProperty', function() {
    if (!data || !data.customerId || !data.type || !data.name) {
      throw new Error('ข้อมูลเจ้าของ ประเภท และชื่อทรัพย์สินเป็นข้อมูลที่จำเป็น');
    }
    
    const sheet = getOrCreateSheet('ทรัพย์สิน');
    const propertyId = data.id || generatePropertyId();
    const now = new Date();
    
    if (data.id) {
      // Update existing property
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === data.id) {
          sheet.getRange(i + 1, 1, 1, 11).setValues([[
            propertyId,
            data.customerId,
            data.type,
            data.name,
            data.address || '',
            data.value || 0,
            data.description || '',
            values[i][7], // Keep original creation date
            'Active',
            data.vehicleType || '',
            data.vehicleCode || ''
          ]]);
          
          // Add to history
          addHistory('แก้ไขข้อมูลทรัพย์สิน', `แก้ไข ${data.name}`, 'ทรัพย์สิน', propertyId);
          
          return { success: true, message: 'อัพเดตข้อมูลทรัพย์สินสำเร็จ', id: propertyId };
        }
      }
    }
    
    // Add new property
    sheet.appendRow([
      propertyId,
      data.customerId,
      data.type,
      data.name,
      data.address || '',
      data.value || 0,
      data.description || '',
      now,
      'Active',
      data.vehicleType || '',
      data.vehicleCode || ''
    ]);
    
    // Add to history
    addHistory('เพิ่มทรัพย์สินใหม่', `เพิ่ม ${data.name}`, 'ทรัพย์สิน', propertyId);
    
    return { success: true, message: 'บันทึกข้อมูลทรัพย์สินสำเร็จ', id: propertyId };
  });
}

function deleteProperty(propertyId) {
  return safeExecute('deleteProperty', function() {
    if (!propertyId) {
      throw new Error('ไม่พบรหัสทรัพย์สิน');
    }
    
    const sheet = getOrCreateSheet('ทรัพย์สิน');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === propertyId) {
        // Soft delete
        sheet.getRange(i + 1, 9).setValue('Deleted');
        
        // Add to history
        addHistory('ลบทรัพย์สิน', `ลบ ${values[i][3]}`, 'ทรัพย์สิน', propertyId);
        
        return { success: true, message: 'ลบข้อมูลทรัพย์สินสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบข้อมูลทรัพย์สิน' };
  });
}

function getPropertyListWrapped() {
  return safeExecute('getPropertyListWrapped', function() {
    const sheet = getOrCreateSheet('ทรัพย์สิน');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: [],
        count: 0
      };
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const properties = values
      .filter(row => row[8] !== 'Deleted')
      .map(row => ({
        id: String(row[0] || ''),
        customerId: String(row[1] || ''),
        type: String(row[2] || ''),
        name: String(row[3] || ''),
        address: String(row[4] || ''),
        value: parseFloat(row[5]) || 0,
        description: String(row[6] || ''),
        createdDate: row[7] ? new Date(row[7]).toISOString() : '',
        status: String(row[8] || 'Active'),
        vehicleType: String(row[9] || ''),
        vehicleCode: String(row[10] || '')
      }));
    
    return {
      success: true,
      data: properties,
      count: properties.length
    };
  }, {
    success: false,
    data: [],
    count: 0,
    error: 'Failed to load properties'
  });
}

// ==================== Work Order Functions ====================
function generateWorkOrderId() {
  try {
    const sheet = getOrCreateSheet('ใบงาน');
    const today = new Date();
    const year = Utilities.formatDate(today, 'GMT+7', 'yyyy');
    const month = Utilities.formatDate(today, 'GMT+7', 'MM');
    
    // Get all work orders for this month
    const lastRow = sheet.getLastRow();
    let maxSequence = 0;
    
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      const prefix = 'WO' + year + month;
      
      // Find the highest sequence number for this month
      values.forEach(row => {
        const woId = String(row[0]);
        if (woId.startsWith(prefix)) {
          const sequence = parseInt(woId.substring(prefix.length)) || 0;
          maxSequence = Math.max(maxSequence, sequence);
        }
      });
    }
    
    // Increment sequence
    const newSequence = String(maxSequence + 1).padStart(4, '0');
    return 'WO' + year + month + newSequence;
  } catch (error) {
    console.error('Error generating work order ID:', error);
    return 'WO' + String(Date.now()).substr(-6);
  }
}

function getNextWorkOrderId() {
  return safeExecute('getNextWorkOrderId', function() {
    return generateWorkOrderId();
  });
}

function getOrCreateWorkOrderItemsSheet() {
  const sheetName = 'รายการใบงาน';
  const sheet = getOrCreateSheet(sheetName);
  
  // Check if headers exist and are correct
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() < 13) {
    console.log('Setting up work order items sheet headers...');
    const headers = [
      'รหัสรายการ', 'เลขที่ใบงาน', 'ประเภทบริการ', 'ประเภทประกัน', 'รายละเอียด', 
      'ทุนประกัน', 'เบี้ยประกัน', 'วันเริ่ม', 'วันสิ้นสุด', 
      'เลขกรมธรรม์', 'บริษัทประกัน', 'สถานะ', 'หมายเหตุ'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  
  return sheet;
}

// Enhanced saveWorkOrder function with discount support
function saveWorkOrder(data) {
  return safeExecute('saveWorkOrder', function() {
    if (!data || !data.customerId || !data.items || data.items.length === 0) {
      throw new Error('ข้อมูลลูกค้าและรายการเป็นข้อมูลที่จำเป็น');
    }
    
    const mainSheet = getOrCreateSheet('ใบงาน');
    const itemsSheet = getOrCreateWorkOrderItemsSheet();
    const user = Session.getActiveUser().getEmail();
    const now = new Date();
    
    // Use provided work order ID or generate new one
    let workOrderId;
    if (data.id) {
      workOrderId = data.id;
    } else if (data.workOrderNumber) {
      workOrderId = data.workOrderNumber;
    } else {
      workOrderId = generateWorkOrderId();
    }
    
    // คำนวณยอดรวมจาก items
    let totalPremium = 0;
    let totalSumInsured = 0;
    
    if (data.items && data.items.length > 0) {
      totalPremium = data.items.reduce((sum, item) => sum + (parseFloat(item.premium) || 0), 0);
      totalSumInsured = data.items.reduce((sum, item) => sum + (parseFloat(item.sumInsured) || 0), 0);
    }
    
    // คำนวณส่วนลดและยอดชำระจริง
    const discount = parseFloat(data.discount) || 0;
    const actualAmount = totalPremium - discount;
    
    // Save main work order
    if (data.id) {
      // Update existing work order
      updateMainWorkOrder(mainSheet, data, workOrderId, totalSumInsured, totalPremium, now, user, discount, actualAmount);
      deleteWorkOrderItems(workOrderId); // Delete old items from all sheets
    } else {
      // Add new work order
      createMainWorkOrder(mainSheet, data, workOrderId, totalSumInsured, totalPremium, now, user, discount, actualAmount);
    }
    
    // Save work order items to appropriate sheets based on service type
    if (data.items && data.items.length > 0) {
      saveWorkOrderItemsByServiceType(workOrderId, data.items, data);
    }
    
    // Create payment/installment plan if specified
    if (data.paymentType && !data.id) { // Only for new work orders
      if (data.paymentType === 'ผ่อนชำระ' && data.numberOfInstallments) {
        const installmentData = {
          workOrderId: workOrderId,
          totalAmount: actualAmount, // ใช้ยอดหลังหักส่วนลด
          numberOfInstallments: data.numberOfInstallments,
          startDate: data.paymentStartDate || new Date(),
          downPayment: data.installmentPlan && data.installmentPlan[0] ? data.installmentPlan[0].amount : 0,
          customDates: data.installmentPlan ? data.installmentPlan.map(p => p.dueDate) : []
        };
        createInstallmentPlan(installmentData);
      } else {
        // ชำระเต็มจำนวน - บันทึกการชำระเลย
        const paymentData = {
          workOrderId: workOrderId,
          amount: actualAmount,
          paymentType: data.paymentType,
          referenceNumber: data.paymentReference || '',
          bank: data.bank || '',
          notes: data.paymentNotes || ''
        };
        savePayment(paymentData);
      }
      
      // Update payment type in work order
      updateWorkOrderPaymentType(workOrderId, data.paymentType, data.numberOfInstallments);
    }
    
    // Add renewal tracking for insurance items
    const insuranceItems = data.items.filter(item => 
      item.serviceType === 'ประกันภัย' || item.serviceType === 'พรบ'
    );
    if (insuranceItems.length > 0 && !data.id) {
      addRenewalTracking(workOrderId);
    }
    
    // Add to history
    const action = data.id ? 'แก้ไขใบงาน' : 'สร้างใบงานใหม่';
    addHistory(action, `${action} ${workOrderId}`, 'ใบงาน', workOrderId);
    
    return { success: true, message: 'บันทึกใบงานสำเร็จ', workOrderId: workOrderId };
  });
}

// Function to save work order items to appropriate sheets based on service type
function saveWorkOrderItemsByServiceType(workOrderId, items, workOrderData) {
  const itemsSheet = getOrCreateWorkOrderItemsSheet(); // Keep original items sheet for compatibility
  
  // Get chassis number from vehicle data
  let chassisNumber = '';
  if (workOrderData.vehicleId) {
    try {
      const vehicleSheet = getOrCreateSheet('รถ');
      const vehicleData = vehicleSheet.getDataRange().getValues();
      for (let i = 1; i < vehicleData.length; i++) {
        if (vehicleData[i][0] === workOrderData.vehicleId) {
          chassisNumber = String(vehicleData[i][7] || ''); // Column 7 is chassis number
          break;
        }
      }
    } catch (error) {
      console.error('Error getting chassis number:', error);
    }
  }
  
  items.forEach((item, index) => {
    const itemId = workOrderId + '-' + String(index + 1).padStart(2, '0');
    const serviceType = item.serviceType || '';
    const targetSheetName = SERVICE_SHEET_MAPPING[serviceType] || 'อื่นๆ';
    
    // คำนวณวันหมดอายุอัตโนมัติถ้าไม่มี
    let endDate = item.endDate;
    if (!endDate && item.startDate) {
      endDate = calculateEndDate(item.startDate);
    }
    
    // Save to original items sheet (for compatibility)
    itemsSheet.appendRow([
      itemId,
      workOrderId,
      serviceType,
      item.insuranceType || '',
      item.details || '',
      item.sumInsured || 0,
      item.premium || 0,
      item.startDate || workOrderData.startDate || '',
      endDate || workOrderData.endDate || '',
      item.policyNumber || '',
      item.insuranceCompany || '',
      'ใช้งาน',
      item.notes || ''
    ]);
    
    // Save to service type specific sheet with chassis number
    saveToServiceTypeSheet(targetSheetName, {
      itemId: itemId,
      workOrderId: workOrderId,
      customerId: workOrderData.customerId,
      customerName: workOrderData.customerName,
      serviceType: serviceType,
      item: item,
      workOrderData: workOrderData,
      endDate: endDate,
      chassisNumber: chassisNumber
    });
  });
}

// Function to save to specific service type sheet
function saveToServiceTypeSheet(sheetName, data) {
  try {
    const sheet = getOrCreateSheet(sheetName);
    const { itemId, workOrderId, customerId, customerName, serviceType, item, workOrderData, endDate, chassisNumber } = data;
    
    switch (sheetName) {
      case 'ประกันภัย':
        sheet.appendRow([
          itemId,                           // รหัสรายการ
          workOrderId,                      // เลขที่ใบงาน
          customerId,                       // รหัสลูกค้า
          customerName,                     // ชื่อลูกค้า
          serviceType,                      // ประเภทบริการ
          item.insuranceType || '',         // ประเภทประกัน
          item.details || '',               // รายละเอียด
          item.sumInsured || 0,             // ทุนประกัน
          item.premium || 0,                // เบี้ยประกัน
          item.startDate || workOrderData.startDate || '', // วันเริ่ม
          endDate || workOrderData.endDate || '',          // วันสิ้นสุด
          item.policyNumber || '',          // เลขกรมธรรม์
          item.insuranceCompany || '',      // บริษัทประกัน
          'ใช้งาน',                          // สถานะ
          new Date(),                       // วันที่สร้าง
          item.notes || '',                 // หมายเหตุ
          chassisNumber || ''               // เลขตัวถัง
        ]);
        break;
        
      case 'ภาษี':
        sheet.appendRow([
          itemId,                           // รหัสรายการ
          workOrderId,                      // เลขที่ใบงาน
          customerId,                       // รหัสลูกค้า
          customerName,                     // ชื่อลูกค้า
          serviceType,                      // ประเภทบริการ
          item.details || '',               // รายละเอียด
          item.premium || 0,                // จำนวนเงิน
          item.startDate || workOrderData.startDate || '', // วันเริ่ม
          endDate || workOrderData.endDate || '',          // วันสิ้นสุด
          item.policyNumber || '',          // เลขที่อ้างอิง
          'ใช้งาน',                          // สถานะ
          new Date(),                       // วันที่สร้าง
          item.notes || '',                 // หมายเหตุ
          chassisNumber || ''               // เลขตัวถัง
        ]);
        break;
        
      case 'ค่าบริการ':
        sheet.appendRow([
          itemId,                           // รหัสรายการ
          workOrderId,                      // เลขที่ใบงาน
          customerId,                       // รหัสลูกค้า
          customerName,                     // ชื่อลูกค้า
          serviceType,                      // ประเภทบริการ
          item.details || '',               // รายละเอียด
          item.premium || 0,                // จำนวนเงิน
          item.startDate || new Date(),     // วันที่ให้บริการ
          'ใช้งาน',                          // สถานะ
          new Date(),                       // วันที่สร้าง
          item.notes || '',                 // หมายเหตุ
          chassisNumber || ''               // เลขตัวถัง
        ]);
        break;
        
      case 'อื่นๆ':
        sheet.appendRow([
          itemId,                           // รหัสรายการ
          workOrderId,                      // เลขที่ใบงาน
          customerId,                       // รหัสลูกค้า
          customerName,                     // ชื่อลูกค้า
          serviceType,                      // ประเภทบริการ
          item.details || '',               // รายละเอียด
          item.premium || 0,                // จำนวนเงิน
          item.startDate || new Date(),     // วันที่
          'ใช้งาน',                          // สถานะ
          new Date(),                       // วันที่สร้าง
          item.notes || '',                 // หมายเหตุ
          chassisNumber || ''               // เลขตัวถัง
        ]);
        break;
    }
    
    console.log(`Saved item ${itemId} to ${sheetName} sheet with chassis number: ${chassisNumber}`);
  } catch (error) {
    console.error(`Error saving to ${sheetName} sheet:`, error);
  }
}

// Enhanced helper functions for main work order operations
function updateMainWorkOrder(sheet, data, workOrderId, totalSumInsured, totalPremium, now, user, discount, actualAmount) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      // Keep existing payment information
      const existingPaymentStatus = values[i][16] || PAYMENT_STATUS.PENDING;
      const existingPaymentType = values[i][17] || '';
      const existingInstallments = values[i][18] || '';
      const existingTotalPaid = values[i][19] || 0;
      const existingRemaining = actualAmount - existingTotalPaid;
      
      sheet.getRange(i + 1, 1, 1, 23).setValues([[
        workOrderId,
        data.customerId,
        data.customerName,
        'รวม',
        data.vehicleId || data.propertyId || '',
        `${data.items.length} รายการ`,
        totalSumInsured,
        totalPremium,
        data.startDate || '',
        data.endDate || '',
        '',
        data.status || values[i][11],
        values[i][12],
        '',
        values[i][14],
        now,
        existingPaymentStatus,
        existingPaymentType,
        existingInstallments,
        existingTotalPaid,
        existingRemaining,
        discount,
        actualAmount
      ]]);
      break;
    }
  }
}

function createMainWorkOrder(sheet, data, workOrderId, totalSumInsured, totalPremium, now, user, discount, actualAmount) {
  sheet.appendRow([
    workOrderId,
    data.customerId,
    data.customerName,
    'รวม',
    data.vehicleId || data.propertyId || '',
    `${data.items.length} รายการ`,
    totalSumInsured,
    totalPremium,
    data.startDate || '',
    data.endDate || '',
    '',
    data.status || 'แจ้งงานแล้ว',
    now,
    '',
    user,
    now,
    PAYMENT_STATUS.PENDING, // สถานะการชำระเงิน
    '', // ประเภทการชำระ
    '', // จำนวนงวด
    0, // ยอดชำระแล้ว
    actualAmount, // ยอดค้างชำระ
    discount, // ส่วนลด
    actualAmount // ยอดชำระจริง
  ]);
}

// Enhanced deleteWorkOrderItems to delete from all sheets
function deleteWorkOrderItems(workOrderId) {
  try {
    // Delete from original items sheet
    const itemsSheet = getOrCreateWorkOrderItemsSheet();
    deleteItemsFromSheet(itemsSheet, workOrderId);
    
    // Delete from service type specific sheets
    const uniqueSheetNames = [...new Set(Object.values(SERVICE_SHEET_MAPPING))];
    uniqueSheetNames.forEach(sheetName => {
      try {
        const sheet = getOrCreateSheet(sheetName);
        deleteItemsFromSheet(sheet, workOrderId);
      } catch (error) {
        console.warn(`Could not delete from ${sheetName}:`, error);
      }
    });
  } catch (error) {
    console.error('Error deleting work order items:', error);
  }
}

// Helper function to delete items from a specific sheet
function deleteItemsFromSheet(sheet, workOrderId) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // Delete from bottom to top to avoid index issues
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][1] === workOrderId) { // Column 1 is work order ID
      sheet.deleteRow(i + 1);
    }
  }
}

function getWorkOrderItems(workOrderId) {
  return safeExecute('getWorkOrderItems', function() {
    console.log('Getting work order items for:', workOrderId);
    
    if (!workOrderId) {
      console.log('No workOrderId provided');
      return [];
    }
    
    const sheet = getOrCreateWorkOrderItemsSheet();
    const lastRow = sheet.getLastRow();
    
    console.log('Work order items sheet last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('No items in sheet');
      return [];
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    console.log('Total rows in items sheet:', values.length);
    
    const items = values
      .filter(row => {
        const rowWorkOrderId = String(row[1] || '');
        const matches = rowWorkOrderId === workOrderId;
        if (matches) {
          console.log('Found matching item:', row[0]);
        }
        return matches;
      })
      .map(row => ({
        id: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        serviceType: String(row[2] || ''),
        insuranceType: String(row[3] || ''),
        details: String(row[4] || ''),
        sumInsured: parseFloat(row[5]) || 0,
        premium: parseFloat(row[6]) || 0,
        startDate: row[7] ? (row[7] instanceof Date ? row[7].toISOString().split('T')[0] : row[7]) : '',
        endDate: row[8] ? (row[8] instanceof Date ? row[8].toISOString().split('T')[0] : row[8]) : '',
        policyNumber: String(row[9] || ''),
        insuranceCompany: String(row[10] || ''),
        status: String(row[11] || ''),
        notes: String(row[12] || '')
      }));
    
    console.log('Found', items.length, 'items for work order', workOrderId);
    return items;
  }, []);
}

function getWorkOrderList() {
  return safeExecute('getWorkOrderList', function() {
    const sheet = getOrCreateSheet('ใบงาน');
    const lastRow = sheet.getLastRow();
    
    console.log('getWorkOrderList - lastRow:', lastRow);
    
    if (lastRow <= 1) {
      console.log('No work orders found');
      return [];
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
    console.log('Retrieved', values.length, 'work order rows');
    
    const workOrders = values.map(row => ({
      id: String(row[0] || ''),
      customerId: String(row[1] || ''),
      customerName: String(row[2] || ''),
      insuranceType: String(row[3] || ''),
      propertyId: String(row[4] || ''),
      details: String(row[5] || ''),
      sumInsured: parseFloat(row[6]) || 0,
      premium: parseFloat(row[7]) || 0,
      startDate: row[8] ? (row[8] instanceof Date ? row[8].toISOString().split('T')[0] : row[8]) : '',
      endDate: row[9] ? (row[9] instanceof Date ? row[9].toISOString().split('T')[0] : row[9]) : '',
      policyNumber: String(row[10] || ''),
      status: String(row[11] || ''),
      createdDate: row[12] ? (row[12] instanceof Date ? row[12].toISOString().split('T')[0] : row[12]) : '',
      insuranceCompany: String(row[13] || ''),
      createdBy: String(row[14] || ''),
      modifiedDate: row[15] ? (row[15] instanceof Date ? row[15].toISOString().split('T')[0] : row[15]) : '',
      paymentStatus: String(row[16] || PAYMENT_STATUS.PENDING),
      paymentType: String(row[17] || ''),
      numberOfInstallments: row[18] || '',
      totalPaid: parseFloat(row[19]) || 0,
      remainingAmount: parseFloat(row[20]) || parseFloat(row[7]) || 0,
      discount: parseFloat(row[21]) || 0,
      actualAmount: parseFloat(row[22]) || parseFloat(row[7]) || 0
    }));
    
    console.log('Processed', workOrders.length, 'work orders');
    return workOrders;
  }, []);
}

function cancelWorkOrder(workOrderId) {
  return safeExecute('cancelWorkOrder', function() {
    if (!workOrderId) {
      throw new Error('ไม่พบรหัสใบงาน');
    }
    
    const sheet = getOrCreateSheet('ใบงาน');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === workOrderId) {
        // Update status to cancelled
        sheet.getRange(i + 1, 12).setValue('ยกเลิก');
        sheet.getRange(i + 1, 16).setValue(new Date());
        sheet.getRange(i + 1, 17).setValue(PAYMENT_STATUS.CANCELLED);
        
        // Cancel all pending installments
        const installmentSheet = getOrCreateSheet('งวดการชำระ');
        const installmentData = installmentSheet.getDataRange().getValues();
        
        for (let j = 1; j < installmentData.length; j++) {
          if (installmentData[j][1] === workOrderId && installmentData[j][7] === 'รอชำระ') {
            installmentSheet.getRange(j + 1, 8).setValue('ยกเลิก');
          }
        }
        
        // Add to history
        addHistory('ยกเลิกใบงาน', `ยกเลิกใบงาน ${workOrderId}`, 'ใบงาน', workOrderId);
        
        return { success: true, message: 'ยกเลิกใบงานสำเร็จ' };
      }
    }
    
    return { success: false, message: 'ไม่พบใบงาน' };
  });
}

function renewPolicy(workOrderId) {
  return safeExecute('renewPolicy', function() {
    if (!workOrderId) {
      throw new Error('ไม่พบรหัสใบงาน');
    }
    
    const workOrders = getWorkOrderList();
    const workOrder = workOrders.find(wo => wo.id === workOrderId);
    
    if (!workOrder) {
      return { success: false, message: 'ไม่พบใบงาน' };
    }
    
    // Get work order items
    const items = getWorkOrderItems(workOrderId);
    
    if (items.length === 0) {
      return { success: false, message: 'ไม่พบรายการในใบงาน' };
    }
    
    // Create new items with dates extended by 1 year
    const renewedItems = items.map(item => {
      const newStartDate = item.endDate ? new Date(item.endDate) : new Date();
      newStartDate.setDate(newStartDate.getDate() + 1); // Start next day after old end date
      
      const newEndDate = calculateEndDate(newStartDate.toISOString().split('T')[0]);
      
      return {
        ...item,
        startDate: newStartDate.toISOString().split('T')[0],
        endDate: newEndDate,
        policyNumber: '', // New policy number will be assigned
        status: 'รอดำเนินการ'
      };
    });
    
    // Create new work order
    const newWorkOrderData = {
      customerId: workOrder.customerId,
      customerName: workOrder.customerName,
      vehicleId: workOrder.propertyId,
      items: renewedItems,
      status: 'แจ้งงานแล้ว',
      startDate: renewedItems[0]?.startDate,
      endDate: renewedItems[0]?.endDate
    };
    
    const result = saveWorkOrder(newWorkOrderData);
    
    if (result.success) {
      // Add to history
      addHistory('ต่ออายุกรมธรรม์', `ต่ออายุจาก ${workOrderId} เป็น ${result.workOrderId}`, 'ใบงาน', result.workOrderId);
    }
    
    return result;
  });
}

// Clone Work Order Function
function cloneWorkOrder(workOrderId) {
  return safeExecute('cloneWorkOrder', function() {
    if (!workOrderId) {
      throw new Error('ไม่พบรหัสใบงาน');
    }
    
    const workOrders = getWorkOrderList();
    const workOrder = workOrders.find(wo => wo.id === workOrderId);
    
    if (!workOrder) {
      return { success: false, message: 'ไม่พบใบงาน' };
    }
    
    // Get work order items
    const items = getWorkOrderItems(workOrderId);
    
    if (items.length === 0) {
      return { success: false, message: 'ไม่พบรายการในใบงาน' };
    }
    
    // Clone items with cleared policy numbers and reset status
    const clonedItems = items.map(item => ({
      ...item,
      policyNumber: '', // Clear policy number for new work order
      status: 'รอดำเนินการ',
      startDate: '', // Clear dates to be set manually
      endDate: ''
    }));
    
    // Create cloned work order
    const clonedWorkOrderData = {
      customerId: workOrder.customerId,
      customerName: workOrder.customerName,
      vehicleId: workOrder.propertyId,
      items: clonedItems,
      status: 'รอดำเนินการ'
    };
    
    const result = saveWorkOrder(clonedWorkOrderData);
    
    if (result.success) {
      // Add to history
      addHistory('โคลนใบงาน', `โคลนจาก ${workOrderId} เป็น ${result.workOrderId}`, 'ใบงาน', result.workOrderId);
    }
    
    return result;
  });
}

// ==================== Service Type Functions ====================
// Function to get items by service type
function getItemsByServiceType(serviceType, limit = 100) {
  return safeExecute('getItemsByServiceType', function() {
    const targetSheetName = SERVICE_SHEET_MAPPING[serviceType] || 'อื่นๆ';
    const sheet = getOrCreateSheet(targetSheetName);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
    
    // Map data based on sheet structure
    if (targetSheetName === 'ประกันภัย') {
      return values.map(row => ({
        itemId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        serviceType: String(row[4] || ''),
        insuranceType: String(row[5] || ''),
        details: String(row[6] || ''),
        sumInsured: parseFloat(row[7]) || 0,
        amount: parseFloat(row[8]) || 0,
        startDate: row[9] || '',
        endDate: row[10] || '',
        policyNumber: String(row[11] || ''),
        insuranceCompany: String(row[12] || ''),
        status: String(row[13] || ''),
        createdDate: row[14] || '',
        notes: String(row[15] || ''),
        chassisNumber: String(row[16] || '')
      }));
    } else if (targetSheetName === 'ภาษี') {
      return values.map(row => ({
        itemId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        serviceType: String(row[4] || ''),
        details: String(row[5] || ''),
        amount: parseFloat(row[6]) || 0,
        startDate: row[7] || '',
        endDate: row[8] || '',
        referenceNumber: String(row[9] || ''),
        status: String(row[10] || ''),
        createdDate: row[11] || '',
        notes: String(row[12] || ''),
        chassisNumber: String(row[13] || '')
      }));
    } else if (targetSheetName === 'ค่าบริการ') {
      return values.map(row => ({
        itemId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        serviceType: String(row[4] || ''),
        details: String(row[5] || ''),
        amount: parseFloat(row[6]) || 0,
        date: row[7] || '',
        status: String(row[8] || ''),
        createdDate: row[9] || '',
        notes: String(row[10] || ''),
        chassisNumber: String(row[11] || '')
      }));
    } else { // อื่นๆ
      return values.map(row => ({
        itemId: String(row[0] || ''),
        workOrderId: String(row[1] || ''),
        customerId: String(row[2] || ''),
        customerName: String(row[3] || ''),
        serviceType: String(row[4] || ''),
        details: String(row[5] || ''),
        amount: parseFloat(row[6]) || 0,
        date: row[7] || '',
        status: String(row[8] || ''),
        createdDate: row[9] || '',
        notes: String(row[10] || ''),
        chassisNumber: String(row[11] || '')
      }));
    }
  }, []);
}

// Function to get service type statistics
function getServiceTypeStats() {
  return safeExecute('getServiceTypeStats', function() {
    const stats = {};
    
    const uniqueSheetNames = [...new Set(Object.values(SERVICE_SHEET_MAPPING))];
    uniqueSheetNames.forEach(sheetName => {
      try {
        const sheet = getOrCreateSheet(sheetName);
        const lastRow = sheet.getLastRow();
        const count = Math.max(0, lastRow - 1); // Subtract header row
        
        if (count > 0) {
          let amountColumnIndex;
          if (sheetName === 'ประกันภัย') {
            amountColumnIndex = 8; // เบี้ยประกัน column
          } else {
            amountColumnIndex = 6; // จำนวนเงิน column
          }
          
          const values = sheet.getRange(2, amountColumnIndex + 1, count, 1).getValues();
          const totalAmount = values.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);
          
          stats[sheetName] = {
            count: count,
            totalAmount: totalAmount
          };
        } else {
          stats[sheetName] = {
            count: 0,
            totalAmount: 0
          };
        }
      } catch (error) {
        console.warn(`Could not get stats for ${sheetName}:`, error);
        stats[sheetName] = {
          count: 0,
          totalAmount: 0
        };
      }
    });
    
    return stats;
  }, {});
}

// ==================== Company Functions ====================
function getCompanyListWrapped() {
  return safeExecute('getCompanyListWrapped', function() {
    const sheet = getOrCreateSheet('บริษัทประกัน');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: [],
        count: 0
      };
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    const companies = values.map(row => ({
      id: String(row[0] || ''),
      name: String(row[1] || ''),
      address: String(row[2] || ''),
      phone: String(row[3] || ''),
      contact: String(row[4] || ''),
      status: String(row[5] || ''),
      email: String(row[6] || '')
    }));
    
    return {
      success: true,
      data: companies,
      count: companies.length
    };
  }, {
    success: false,
    data: [],
    count: 0,
    error: 'Failed to load companies'
  });
}

// ==================== History Functions ====================
function generateHistoryId() {
  try {
    const sheet = getOrCreateSheet('ประวัติ');
    const lastRow = sheet.getLastRow();
    return 'HIST' + String(lastRow).padStart(5, '0');
  } catch (error) {
    console.error('Error generating history ID:', error);
    return 'HIST' + String(Date.now()).substr(-5);
  }
}

function addHistory(action, details, relatedTable, referenceId) {
  return safeExecute('addHistory', function() {
    const sheet = getOrCreateSheet('ประวัติ');
    const historyId = generateHistoryId();
    const user = Session.getActiveUser().getEmail();
    const now = new Date();
    
    sheet.appendRow([
      historyId,
      now,
      user,
      action,
      details,
      relatedTable,
      referenceId
    ]);
    
    return { success: true, id: historyId };
  }, { success: false, error: 'Failed to add history' });
}

function getHistory(limit = 50) {
  return safeExecute('getHistory', function() {
    const sheet = getOrCreateSheet('ประวัติ');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    
    const values = sheet.getRange(startRow, 1, numRows, 7).getValues();
    
    return values.reverse().map(row => ({
      id: String(row[0] || ''),
      date: row[1] || '',
      user: String(row[2] || ''),
      action: String(row[3] || ''),
      details: String(row[4] || ''),
      relatedTable: String(row[5] || ''),
      referenceId: String(row[6] || '')
    }));
  }, []);
}

// ==================== Dashboard Functions ====================
function getDashboardData() {
  return safeExecute('getDashboardData', function() {
    const customersResult = getCustomerListWrapped();
    const workOrders = getWorkOrderList();
    
    const today = new Date();
    const thirtyDaysLater = new Date(today.getTime() + (30 * 24 * 60 * 60 * 1000));
    
    let expiringCount = 0;
    let totalRevenue = 0;
    let pendingPayments = 0;
    let overduePayments = 0;
    
    workOrders.forEach(wo => {
      if (wo.endDate && wo.status !== 'ยกเลิก') {
        const endDate = new Date(wo.endDate);
        if (endDate >= today && endDate <= thirtyDaysLater) {
          expiringCount++;
        }
      }
      
      if (wo.status !== 'ยกเลิก') {
        totalRevenue += wo.actualAmount || wo.premium || 0;
        
        // Count payment statuses
        if (wo.paymentStatus === PAYMENT_STATUS.PENDING) {
          pendingPayments++;
        } else if (wo.paymentStatus === PAYMENT_STATUS.INSTALLMENT) {
          // Check for overdue installments
          const installments = getInstallmentPlan(wo.id);
          const hasOverdue = installments.some(inst => {
            if (inst.status !== 'ชำระแล้ว') {
              const dueDate = new Date(inst.dueDate);
              return dueDate < today;
            }
            return false;
          });
          if (hasOverdue) {
            overduePayments++;
          }
        }
      }
    });
    
    return {
      totalCustomers: customersResult.count || 0,
      totalWorkOrders: workOrders.filter(wo => wo.status !== 'ยกเลิก').length,
      expiringPolicies: expiringCount,
      totalRevenue: totalRevenue,
      pendingPayments: pendingPayments,
      overduePayments: overduePayments
    };
  }, {
    totalCustomers: 0,
    totalWorkOrders: 0,
    expiringPolicies: 0,
    totalRevenue: 0,
    pendingPayments: 0,
    overduePayments: 0
  });
}

// ==================== Backup Functions ====================
function createBackup() {
  return safeExecute('createBackup', function() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const backupName = `Backup_Insurance_${Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss')}`;
    
    const backup = ss.copy(backupName);
    
    // Try to create folder, but don't fail if it already exists
    let folder;
    try {
      folder = DriveApp.getFoldersByName('Insurance_Backups').next();
    } catch (e) {
      folder = DriveApp.createFolder('Insurance_Backups');
    }
    
    DriveApp.getFileById(backup.getId()).moveTo(folder);
    
    return {
      success: true,
      message: 'สร้างสำเนาสำรองสำเร็จ',
      backupId: backup.getId(),
      backupUrl: backup.getUrl()
    };
  });
}

// ==================== Test Functions ====================
function testConnection() {
  return safeExecute('testConnection', function() {
    console.log('Testing connection with ID:', SPREADSHEET_ID);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const result = {
      success: true,
      spreadsheetName: ss.getName(),
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: ss.getUrl(),
      userEmail: Session.getActiveUser().getEmail(),
      sheets: sheets.map(s => ({
        name: s.getName(),
        lastRow: s.getLastRow(),
        lastColumn: s.getLastColumn()
      }))
    };
    console.log('Test connection result:', result);
    return result;
  });
}

function testConnectionDetailed() {
  return testConnection(); // Use the same function for now
}

function addTestData() {
  return safeExecute('addTestData', function() {
    // Add test customer
    const customerResult = saveCustomer({
      name: 'คุณทดสอบ ระบบ',
      idCard: '1234567890123',
      phone: '0812345678',
      email: 'test@example.com',
      shippingAddress: '123 ถนนทดสอบ กรุงเทพฯ 10100',
      documentAddress: '456 ถนนจริง กรุงเทพฯ 10200',
      notes: 'ลูกค้าทดสอบระบบ'
    });
    console.log('Customer added:', customerResult);
    
    if (customerResult.success && customerResult.id) {
      // Add test vehicle
      const vehicleResult = saveVehicle({
        customerId: customerResult.id,
        licensePlate: 'กก 1234',
        brand: 'Toyota',
        model: 'Vios',
        color: 'ขาว',
        year: '2023',
        chassisNumber: 'ABC123456789'
      });
      console.log('Vehicle added:', vehicleResult);
      
      // Add test property
      const propertyResult = saveProperty({
        customerId: customerResult.id,
        type: 'บ้าน',
        name: 'บ้านทดสอบ',
        address: '789 ซอยทดสอบ',
        value: 1000000,
        description: 'บ้านเดี่ยว 2 ชั้น',
        vehicleType: 'รถเก๋ง',
        vehicleCode: 'CAR001'
      });
      console.log('Property added:', propertyResult);
      
      return {
        success: true,
        customer: customerResult,
        vehicle: vehicleResult,
        property: propertyResult
      };
    } else {
      return {
        success: false,
        error: 'Failed to create customer'
      };
    }
  });
}

// ==================== Additional Helper Functions ====================
function loadCustomerList() {
  // This is an alias for getCustomerListWrapped for consistency with frontend
  return getCustomerListWrapped();
}

function loadVehicleList() {
  // This is an alias for getVehicleListWrapped for consistency with frontend
  return getVehicleListWrapped();
}

function loadPropertyList() {
  // This is an alias for getPropertyListWrapped for consistency with frontend
  return getPropertyListWrapped();
}

function loadCompanies() {
  // This is an alias for getCompanyListWrapped for consistency with frontend
  return getCompanyListWrapped();
}

// ==================== Data Migration Functions ====================
function migrateWorkOrdersToItems() {
  // ฟังก์ชันสำหรับย้ายข้อมูลจาก sheet ใบงานไปยัง sheet รายการใบงาน
  // สำหรับใบงานที่สร้างก่อนมีระบบ items
  
  console.log('Starting work order migration...');
  
  const workOrderSheet = getOrCreateSheet('ใบงาน');
  const itemsSheet = getOrCreateWorkOrderItemsSheet();
  
  const lastRow = workOrderSheet.getLastRow();
  if (lastRow <= 1) {
    console.log('No work orders to migrate');
    return { success: true, message: 'No work orders to migrate' };
  }
  
  const workOrders = workOrderSheet.getRange(2, 1, lastRow - 1, 16).getValues();
  let migratedCount = 0;
  
  workOrders.forEach((row, index) => {
    const workOrderId = row[0];
    
    // Check if items already exist for this work order
    const existingItems = itemsSheet.getDataRange().getValues()
      .filter((itemRow, i) => i > 0 && itemRow[1] === workOrderId);
    
    if (existingItems.length === 0) {
      // No items exist, create one from the main work order data
      const itemId = workOrderId + '-01';
      
      // Determine service type based on insurance type
      let serviceType = 'ประกันภัย';
      if (row[3] === 'พรบ' || row[3] === 'พ.ร.บ.') {
        serviceType = 'พรบ';
      }
      
      itemsSheet.appendRow([
        itemId,                    // รหัสรายการ
        workOrderId,               // เลขที่ใบงาน
        serviceType,               // ประเภทบริการ
        row[3] || '',             // ประเภทประกัน
        row[5] || '',             // รายละเอียด
        row[6] || 0,              // ทุนประกัน
        row[7] || 0,              // เบี้ยประกัน
        row[8] || '',             // วันเริ่ม
        row[9] || '',             // วันสิ้นสุด
        row[10] || '',            // เลขกรมธรรม์
        row[13] || '',            // บริษัทประกัน
        'ใช้งาน',                  // สถานะ
        'Migrated from work order' // หมายเหตุ
      ]);
      
      migratedCount++;
      console.log(`Migrated work order ${workOrderId}`);
    }
  });
  
  console.log(`Migration complete. Migrated ${migratedCount} work orders.`);
  return { 
    success: true, 
    message: `Migration complete. Migrated ${migratedCount} work orders.`,
    count: migratedCount
  };
}

// ==================== Email Functions ====================
function sendEmailNotification(data) {
  return safeExecute('sendEmailNotification', function() {
    if (!data || !data.to || !data.subject || !data.body) {
      throw new Error('ข้อมูลอีเมลไม่ครบถ้วน');
    }
    
    // Validate email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(data.to)) {
      throw new Error('รูปแบบอีเมลผู้รับไม่ถูกต้อง');
    }
    
    // Validate CC email if provided
    if (data.cc && !emailRegex.test(data.cc)) {
      throw new Error('รูปแบบอีเมล CC ไม่ถูกต้อง');
    }
    
    try {
      // Prepare email options
      const emailOptions = {
        to: data.to,
        subject: data.subject,
        body: data.body,
        htmlBody: convertToHtml(data.body)
      };
      
      // Add CC if provided
      if (data.cc) {
        emailOptions.cc = data.cc;
      }
      
      // Send email
      MailApp.sendEmail(emailOptions);
      
      // Add to history with CC info
      if (data.workOrderId) {
        const ccInfo = data.cc ? ` (CC: ${data.cc})` : '';
        addHistory('ส่งอีเมลแจ้งงาน', `ส่งอีเมลไปยัง ${data.to}${ccInfo} สำหรับใบงาน ${data.workOrderId}`, 'ใบงาน', data.workOrderId);
      }
      
      return {
        success: true,
        message: 'ส่งอีเมลสำเร็จ'
      };
    } catch (error) {
      console.error('Error sending email:', error);
      throw new Error('ไม่สามารถส่งอีเมลได้: ' + error.toString());
    }
  });
}

// Helper function to convert plain text to HTML
function convertToHtml(text) {
  // Convert line breaks to <br>
  let html = text.replace(/\n/g, '<br>');
  
  // Convert === headers === to bold
  html = html.replace(/=== (.*?) ===/g, '<strong>$1</strong>');
  
  // Add basic styling
  html = `
    <div style="font-family: 'Sarabun', Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333;">
      <div style="background-color: #f5f5f5; padding: 20px; border-radius: 8px;">
        ${html}
      </div>
      <div style="margin-top: 20px; padding: 10px; background-color: #e3f2fd; border-left: 4px solid #2196f3;">
        <p style="margin: 0; font-size: 12px; color: #666;">
          อีเมลนี้ส่งจากระบบ Insurance Sale Management<br>
          หากมีข้อสงสัยกรุณาติดต่อกลับ
        </p>
      </div>
    </div>
  `;
  
  return html;
}

// ==================== SEARCH VEHICLES FUNCTION ====================
function searchVehicles(searchText) {
  return safeExecute('searchVehicles', function() {
    if (!searchText || searchText.trim() === '') {
      return { success: true, data: [], count: 0 };
    }
    
    const sheet = getOrCreateSheet('รถ');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: true, data: [], count: 0 };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const searchLower = searchText.toLowerCase();
    
    const results = values.filter(row => {
      const licensePlate = String(row[2] || '').toLowerCase();
      const brand = String(row[3] || '').toLowerCase();
      const model = String(row[4] || '').toLowerCase();
      
      return row[9] !== 'Deleted' && (
        licensePlate.includes(searchLower) || 
        brand.includes(searchLower) || 
        model.includes(searchLower)
      );
    }).map(row => ({
      id: String(row[0] || ''),
      customerId: String(row[1] || ''),
      licensePlate: String(row[2] || ''),
      brand: String(row[3] || ''),
      model: String(row[4] || ''),
      color: String(row[5] || ''),
      year: String(row[6] || ''),
      chassisNumber: String(row[7] || ''),
      createdDate: row[8] ? new Date(row[8]).toISOString() : '',
      status: String(row[9] || 'Active')
    }));
    
    return {
      success: true,
      data: results,
      count: results.length
    };
  }, { success: false, data: [], count: 0, error: 'Search failed' });
}

// ==================== PUBLIC WRAPPER FUNCTIONS FOR FRONTEND ====================
/**
 * markInstallmentPaid
 * Frontend helper – simply calls payInstallment with provided installmentId & amount.
 * @param {{installmentId:string, amount:number, referenceNumber?:string, bank?:string, notes?:string}} data 
 */
function markInstallmentPaid(data) {
  return safeExecute('markInstallmentPaid', function() {
    if (!data || !data.installmentId || !data.amount) {
      throw new Error('ข้อมูลไม่ครบถ้วน');
    }
    // Re-use main logic in payInstallment
    return payInstallment({
      installmentId: data.installmentId,
      amount: data.amount,
      referenceNumber: data.referenceNumber || '',
      bank: data.bank || '',
      notes: data.notes || ''
    });
  });
}

/**
 * loadInstallmentTracking
 * Returns installments due in current month (wrapper for getInstallmentsDueThisMonth)
 */
function loadInstallmentTracking() {
  return getInstallmentsDueThisMonth();
}
