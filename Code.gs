const DB_ID = '1dgVS30lADu1gSQfsWwt4Sj-qjFE3GnLVbdFmwEZfPB8'; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Hệ Thống Quản Lý Xe Công Vụ Đồng Nai')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initDB() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    
    let sh = getOrCreateSheet(ss, 'Users', ['username', 'password', 'role', 'name']);
    if (sh.getLastRow() === 1) {
      sh.appendRow(['admin', 'Dongnai@2026', 'ADMIN', 'Quản trị viên']);
      sh.appendRow(['vanphong', 'Dongnai@2026', 'APPROVER', 'Cán bộ duyệt lịch']);
      sh.appendRow(['canbo', 'Dongnai@2026', 'USER', 'Cán bộ đăng ký']);
    }
    
    sh = getOrCreateSheet(ss, 'Leaders', ['id', 'name', 'title', 'fixedVehicleId', 'autoAssign', 'policyNote']);
    if (sh.getLastRow() === 1) {
      sh.appendRow(['LD01', 'Bà Trần Thu Mai', 'Phó Chủ tịch Thường trực', 'V01', 'TRUE', 'Toàn quyền sử dụng xe']);
      sh.appendRow(['LD02', 'Ông Lê Văn Hoan', 'Phó Chủ tịch', 'V02', 'TRUE', 'Ưu tiên xe 60A 000.65']);
      sh.appendRow(['LD03', 'Ông Nguyễn Trung Dũng', 'Phó Chủ tịch', 'V02', 'TRUE', 'Ưu tiên xe 60A 000.65']);
      sh.appendRow(['LD04', 'Bà Hà Hồng Hạnh', 'Phó Bí thư Đảng ủy', 'V02', 'TRUE', 'Ưu tiên xe 60A 000.65']);
      sh.appendRow(['LD05', 'Lãnh đạo các ban', 'Các Ban chuyên môn', '', 'FALSE', 'Chờ xét duyệt thủ công']);
    }
    
    sh = getOrCreateSheet(ss, 'Vehicles', ['id', 'plate', 'brand', 'type']);
    if (sh.getLastRow() === 1) {
      sh.appendRow(['V01', '60A 006.36', 'Ford Everest', 'SUV']);
      sh.appendRow(['V02', '60A 000.65', 'Toyota Fortuner', 'SUV']);
      sh.appendRow(['V03', '60A 002.02', 'Xe 16 chỗ', 'Minibus']);
      sh.appendRow(['V_CANHAN', 'Phương tiện tự túc', 'Xe cá nhân', 'Cá nhân']);
    }

    sh = getOrCreateSheet(ss, 'Drivers', ['id', 'name', 'phone', 'license', 'vehicleId', 'servedLeaders', 'status']);
    if (sh.getLastRow() === 1) {
      sh.appendRow(['TX01', 'Nguyễn Lữ Phú', '0706195444', 'D', 'V01', 'Bà Trần Thu Mai & Các PCT, PBT', 'AVAILABLE']);
      sh.appendRow(['TX02', 'Phạm Ngọc Hiếu', '0918043222', 'E', 'V02', 'Ông Hoan, Ông Dũng, Bà Hạnh', 'AVAILABLE']);
      sh.appendRow(['TX_TUTUC', 'Tự Lái', '-', '-', 'V_CANHAN', 'Phương tiện cá nhân', 'AVAILABLE']);
    }

    getOrCreateSheet(ss, 'Schedules', ['id', 'date', 'start', 'end', 'buffer', 'leaderId', 'companions', 'content', 'from', 'to', 'passengers', 'reqVehicle', 'assignedVehicleId', 'assignedDriverId', 'status', 'reason', 'plan', 'resultText', 'isUpdated', 'sortIndex']);
  } catch (e) {
    throw new Error("Lỗi khi khởi tạo CSDL: " + e.message);
  }
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

// BẮT BUỘC ÉP CHUỖI JSON ĐỂ TRÁNH LỖI TREO APP
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName('Users');
    if(!sheet) return JSON.stringify({ success: false, message: 'Bảng Users chưa tồn tại. Vui lòng bấm chạy hàm initDB.' });
    
    const data = sheet.getDataRange().getDisplayValues();
    if(data.length < 2) return JSON.stringify({ success: false, message: 'Hệ thống chưa có tài khoản nào được tạo.' });

    for(let i=1; i<data.length; i++) {
      if(data[i][0] === username && data[i][1] === password) {
        return JSON.stringify({ success: true, username: username, role: data[i][2], name: data[i][3] });
      }
    }
    return JSON.stringify({ success: false, message: 'Sai tài khoản hoặc mật khẩu!' });
  } catch (e) {
    return JSON.stringify({ success: false, message: 'Lỗi truy cập Google Sheets: ' + e.message });
  }
}

function getDbData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const result = { success: true };
    const sheetNames = ['Users', 'Leaders', 'Vehicles', 'Drivers', 'Schedules'];
    
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        const data = sheet.getDataRange().getDisplayValues(); 
        const rows = [];
        if (data.length > 1) {
          const headers = data[0];
          for (let i = 1; i < data.length; i++) {
            if (!data[i][0]) continue; // Lọc dòng trống
            let obj = {};
            for (let j = 0; j < headers.length; j++) {
              obj[headers[j]] = data[i][j] || "";
            }
            rows.push(obj);
          }
        }
        result[name.toLowerCase()] = rows;
      } else {
        result[name.toLowerCase()] = [];
      }
    });
    return JSON.stringify(result);
  } catch (e) {
    return JSON.stringify({ success: false, error: "Lỗi tải dữ liệu: " + e.message });
  }
}

function saveRecord(sheetName, record, idField) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName(sheetName);
    if(!sheet) throw new Error("Không tìm thấy bảng " + sheetName);
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getDisplayValues();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf(idField)] == record[idField]) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const rowData = headers.map(h => record[h] !== undefined ? record[h] : '');
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return true;
  } catch (e) {
    throw new Error(e.message);
  }
}

function deleteRecord(sheetName, idValue, idField) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName(sheetName);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf(idField)] == idValue) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch(e) {
    throw new Error(e.message);
  }
}

function approveSchedulesBatch(ids) {
  try {
    const ss = SpreadsheetApp.openById(DB_ID);
    const sheet = ss.getSheetByName('Schedules');
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const idIdx = headers.indexOf('id');
    const updateIdx = headers.indexOf('isUpdated');
    
    if (idIdx === -1 || updateIdx === -1) return false;
    
    let hasChanges = false;
    for (let i = 1; i < data.length; i++) {
      if (ids.includes(data[i][idIdx])) {
        data[i][updateIdx] = 'FALSE';
        hasChanges = true;
      }
    }
    
    if (hasChanges) sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return true;
  } catch(e) {
    throw new Error(e.message);
  }
}

function getVisitCount() {
  try {
    const props = PropertiesService.getScriptProperties();
    let count = props.getProperty('SYSTEM_VISITS');
    if (!count) count = '0';
    count = parseInt(count, 10) + 1;
    props.setProperty('SYSTEM_VISITS', count.toString());
    return count;
  } catch (e) {
    return "Lỗi đếm";
  }
}