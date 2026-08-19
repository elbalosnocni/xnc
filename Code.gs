// ==========================================
// A. CHUẨN HÓA DATA STRUCTURE & CONSTANTS
// ==========================================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const TABLES = {
  USERS: 'Users',
  EMPLOYEES: 'Employees',
  DOCUMENTS: 'Documents',
  CONTRACTS: 'Contracts',
  APPENDICES: 'ContractAppendices',
  ENTRY_EXIT: 'EntryExit',
  TRAVEL: 'DomesticTravel',
  AUDIT: 'AuditLog',
  CONFIG: 'Config'
};

const ROLES = {
  ADMIN: 'ADMIN',
  HR_ADMIN: 'HR ADMIN',
  HR: 'HR',
  DIRECTOR: 'DIRECTOR',
  USER: 'USER'
};

function getSheet(tableName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(tableName);
  if (!sheet) {
    setupSystem();
    sheet = ss.getSheetByName(tableName);
  }
  return sheet;
}

// ==========================================
// B. ROUTER API (doGet & doPost)
// ==========================================
function doGet(e) {
  return handleRequest(e, "GET");
}

function doPost(e) {
  return handleRequest(e, "POST");
}

function handleRequest(e, method) {
  const responseHeader = ContentService.createTextOutput();
  responseHeader.setMimeType(ContentService.MimeType.JSON);

  try {
    let requestData = {};
    if (method === "POST" && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      requestData = e.parameter;
    }

    const action = requestData.action || e.parameter.action;
    
    // 1. Xử lý Đăng nhập công khai không cần Token
    if (action === "LOGIN") {
      const res = handleLogin(requestData.username, requestData.password);
      return responseHeader.setContent(JSON.stringify(res));
    }

    // 2. Xác thực Session / Token người dùng
    const token = requestData.token || e.parameter.token;
    const requester = validateSession(token);
    if (!requester) {
      return responseHeader.setContent(JSON.stringify({ status: "ERROR", message: "Phiên đăng nhập hết hạn hoặc không hợp lệ!" }));
    }

    let result = { status: "ERROR", message: "Hành động không hợp lệ!" };

    // 3. Phân luồng các API Action
    switch (action) {
      case "LOGOUT":
        result = handleLogout(token);
        break;
      case "GET_DASHBOARD_DATA":
        result = getDashboardData(requester);
        break;
      case "GET_PROFILE":
        result = getProfileData(requestData.employeeId || e.parameter.employeeId, requester);
        break;
      case "SAVE_PROFILE":
        result = handleSaveProfile(requestData.data || requestData, requester);
        break;
      case "SAVE_DOCUMENT":
        result = saveDocument(requestData.data || requestData, requester);
        break;
      case "UPLOAD_FILE":
        const fileParam = requestData.data || requestData;
        result = uploadEmployeeFile(fileParam.base64Data, fileParam.fileName, fileParam.mimeType, fileParam.employeeId, fileParam.subFolderType);
        break;
      case "ADD_ENTRY_EXIT":
        result = handleAddEntryExit(requestData.data || requestData, requester);
        break;
      case "GET_ENTRY_EXIT":
        result = handleGetEntryExit(requestData.employeeId || e.parameter.employeeId, requester);
        break;
      case "ADD_DOMESTIC_TRAVEL":
        result = handleAddDomesticTravel(requestData.data || requestData, requester);
        break;
      case "GET_DOMESTIC_TRAVEL":
        result = handleGetDomesticTravel(requestData.employeeId || e.parameter.employeeId, requester);
        break;
      case "GET_TIMELINE":
        result = handleGetTimeline(requestData.employeeId || e.parameter.employeeId, requester);
        break;
      case "DELETE_EMPLOYEE":
        result = handleDeleteEmployee(requestData.employeeId || requestData.data?.employeeId, requester);
        break;
    }

    return responseHeader.setContent(JSON.stringify(result));

  } catch (error) {
    return responseHeader.setContent(JSON.stringify({ status: "ERROR", message: error.toString() }));
  }
}

// ==========================================
// C. XÁC THỰC & QUẢN LÝ PHIÊN (AUTH & SESSION)
// ==========================================
function handleLogin(username, password) {
  const sheet = getSheet(TABLES.USERS);
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const [userId, email, passHash, role, employeeId, status] = rows[i];
    if ((String(email).toLowerCase() === String(username).toLowerCase() || String(employeeId) === String(username)) && String(passHash) === String(password)) {
      if (status !== "ACTIVE") {
        return { status: "ERROR", message: "Tài khoản của bạn đã bị khóa." };
      }
      
      const token = "TK-" + Utilities.getUuid();
      const userProps = PropertiesService.getScriptProperties();
      const sessionData = { userId, email, role, employeeId, createdAt: new Date().getTime() };
      
      userProps.setProperty(token, JSON.stringify(sessionData));
      
      // Lấy tên đầy đủ của nhân viên
      let fullName = email;
      const empSheet = getSheet(TABLES.EMPLOYEES);
      const empRows = empSheet.getDataRange().getValues();
      for (let j = 1; j < empRows.length; j++) {
        if (String(empRows[j][0]) === String(employeeId)) {
          fullName = empRows[j][1];
          break;
        }
      }

      writeAuditLog(userId, "LOGIN", "Users", userId, null, "User logged in");

      return {
        status: "SUCCESS",
        token: token,
        user: { id: userId, email: email, name: fullName, role: role, employeeId: employeeId }
      };
    }
  }
  return { status: "ERROR", message: "Tài khoản hoặc mật khẩu không chính xác!" };
}

function validateSession(token) {
  if (!token) return null;
  const userProps = PropertiesService.getScriptProperties();
  const raw = userProps.getProperty(token);
  if (!raw) return null;
  return JSON.parse(raw);
}

function handleLogout(token) {
  PropertiesService.getScriptProperties().deleteProperty(token);
  return { status: "SUCCESS", message: "Đã đăng xuất" };
}

// ==========================================
// D. LOGIC NGHIỆP VỤ CHÍNH & LƯU TRỮ
// ==========================================

function handleSaveProfile(data, requester) {
  const isHRAdmin = [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role);
  const isSelf = String(requester.employeeId) === String(data.employeeId);

  if (!isHRAdmin && !isSelf) {
    return { status: "ERROR", message: "Bạn không có quyền thực hiện thao tác này!" };
  }

  const sheet = getSheet(TABLES.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.employeeId)) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1 && !isHRAdmin) {
    return { status: "ERROR", message: "Không tìm thấy hồ sơ nhân viên để cập nhật!" };
  }

  const timestamp = new Date();

  if (rowIndex > 0) {
    // PARTIAL UPDATE
    const currentRow = rows[rowIndex - 1];
    
    const updatedRow = [
      currentRow[0],
      data.fullName !== undefined ? data.fullName : currentRow[1],
      data.nationality !== undefined ? data.nationality : currentRow[2],
      data.dob !== undefined ? formatDate(data.dob) : currentRow[3],
      data.position !== undefined ? data.position : currentRow[4],
      data.department !== undefined ? data.department : currentRow[5],
      data.phone !== undefined ? data.phone : currentRow[6],
      data.email !== undefined ? data.email : currentRow[7],
      currentRow[8], 
      currentRow[9], 
      (isHRAdmin && data.activeStatus !== undefined) ? data.activeStatus : currentRow[10],
      currentRow[11], 
      currentRow[12], 
      timestamp      
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    
    // Nếu có dữ liệu Hợp đồng/Giấy tờ gửi kèm trong form
    if (data.passportNo) {
      saveDocument({ employeeId: data.employeeId, docType: "PASSPORT", docNo: data.passportNo, expiryDate: data.passportExpiry, fileUrl: data.passportImg }, requester);
    }
    if (data.trcNo) {
      saveDocument({ employeeId: data.employeeId, docType: "TRC", docNo: data.trcNo, expiryDate: data.trcExpiry, fileUrl: data.trcImg }, requester);
    }
    if (data.wpNo) {
      saveDocument({ employeeId: data.employeeId, docType: "WORK_PERMIT", docNo: data.wpNo, expiryDate: data.wpExpiry, fileUrl: data.wpImg }, requester);
    }

    writeAuditLog(requester.userId, "UPDATE", "Employees", data.employeeId, JSON.stringify(currentRow), JSON.stringify(updatedRow));
    return { status: "SUCCESS", message: "Cập nhật hồ sơ thành công!" };

  } else {
    // TẠO MỚI
    const newFolderId = createEmployeeDriveFolder(data.employeeId, data.fullName);
    const newRow = [
      data.employeeId,
      data.fullName || "",
      data.nationality || "",
      formatDate(data.dob) || "",
      data.position || "",
      data.department || "",
      data.phone || "",
      data.email || "",
      "Exited",
      "Overseas",
      "ACTIVE",
      newFolderId,
      timestamp,
      timestamp
    ];

    sheet.appendRow(newRow);

    // Tạo tài khoản User mặc định
    const userSheet = getSheet(TABLES.USERS);
    userSheet.appendRow(["USR-" + Utilities.getUuid().substring(0,6), data.email || data.employeeId, "123456", ROLES.USER, data.employeeId, "ACTIVE", timestamp]);

    writeAuditLog(requester.userId, "CREATE", "Employees", data.employeeId, null, JSON.stringify(newRow));
    return { status: "SUCCESS", message: "Thêm mới nhân viên thành công!" };
  }
}

function saveDocument(docData, requester) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role) && String(requester.employeeId) !== String(docData.employeeId)) {
    return { status: "ERROR", message: "Không có quyền cập nhật giấy tờ!" };
  }

  const sheet = getSheet(TABLES.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(docData.employeeId) && rows[i][2] === docData.docType && rows[i][9] === true) {
      sheet.getRange(i + 1, 10).setValue(false);
    }
  }

  const docId = "DOC-" + Utilities.getUuid().substring(0, 8);
  const newDocRow = [
    docId,
    docData.employeeId,
    docData.docType,
    docData.docNo || "",
    formatDate(docData.issueDate) || "",
    formatDate(docData.expiryDate) || "",
    docData.issuer || "",
    docData.fileId || "",
    docData.fileUrl || "",
    true,
    "VALID",
    new Date(),
    requester.userId
  ];

  sheet.appendRow(newDocRow);
  writeAuditLog(requester.userId, "ADD_DOCUMENT", "Documents", docId, null, JSON.stringify(newDocRow));
  
  return { status: "SUCCESS", message: "Thêm mới giấy tờ thành công!", docId: docId };
}

function uploadEmployeeFile(base64Data, fileName, mimeType, employeeId, subFolderType) {
  try {
    const empSheet = getSheet(TABLES.EMPLOYEES);
    const emps = empSheet.getDataRange().getValues();
    let folderId = "";

    for (let i = 1; i < emps.length; i++) {
      if (String(emps[i][0]) === String(employeeId)) {
        folderId = emps[i][11];
        break;
      }
    }

    let parentFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    let targetFolder;
    const subFolders = parentFolder.getFoldersByName(subFolderType || "Misc");
    targetFolder = subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(subFolderType || "Misc");

    const bytes = Utilities.base64Decode(base64Data.split(",")[1]);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const file = targetFolder.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      status: "SUCCESS",
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      thumbnailUrl: `https://drive.google.com/thumbnail?sz=w800&id=${file.getId()}`
    };
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

function calculateEmployeeLocationStatus(employeeId) {
  const eeSheet = getSheet(TABLES.ENTRY_EXIT);
  const eeData = eeSheet.getDataRange().getValues().slice(1);
  
  const userEE = eeData
    .filter(row => String(row[1]) === String(employeeId))
    .sort((a, b) => new Date(b[3]) - new Date(a[3]));

  let status = "Exited";
  let location = "Overseas";

  if (userEE.length > 0) {
    const latestEE = userEE[0];
    if (latestEE[2] === "ENTRY") {
      status = "In Vietnam";
      location = latestEE[4] || "Vietnam";
      
      const travelSheet = getSheet(TABLES.TRAVEL);
      const travelData = travelSheet.getDataRange().getValues().slice(1);
      const now = new Date();
      
      const activeTravel = travelData.find(row => {
        const startDate = new Date(row[4]);
        const endDate = new Date(row[5]);
        return String(row[1]) === String(employeeId) && 
               row[6] === "APPROVED" && 
               now >= startDate && now <= endDate;
      });

      if (activeTravel) {
        status = "Traveling";
        location = activeTravel[3];
      }
    }
  }

  updateEmployeeStatusRecord(employeeId, status, location);
  return { status, location };
}

function updateEmployeeStatusRecord(employeeId, status, location) {
  const sheet = getSheet(TABLES.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(employeeId)) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 10).setValue(location);
      break;
    }
  }
}

// ==========================================
// E. BÁO CÁO DASHBOARD & DỮ LIỆU
// ==========================================
function getDashboardData(requester) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR, ROLES.DIRECTOR].includes(requester.role)) {
    return { status: "ERROR", message: "Không có quyền truy cập dữ liệu quản trị." };
  }

  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues().slice(1);
  
  const docSheet = getSheet(TABLES.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues().slice(1);

  const kpi = { total: 0, inVN: 0, business: 0, exited: 0, warning: 0, expContract: 0 };
  const employees = [];
  const expiryWarnings = [];
  const now = new Date();

  empRows.forEach(row => {
    if (row[10] === "ACTIVE") {
      kpi.total++;
      if (row[8] === "In Vietnam") kpi.inVN++;
      if (row[8] === "Traveling") kpi.business++;
      if (row[8] === "Exited") kpi.exited++;

      employees.push({
        employeeId: row[0],
        fullName: row[1],
        nationality: row[2],
        position: row[4],
        department: row[5],
        currentStatus: row[8],
        currentLocation: row[9]
      });
    }
  });

  // Kiểm tra cảnh báo hết hạn giấy tờ (45 ngày)
  docRows.forEach(doc => {
    if (doc[9] === true && doc[5]) { // IsCurrent === true
      const expDate = new Date(doc[5]);
      const diffTime = expDate - now;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays <= 45) {
        kpi.warning++;
        const emp = employees.find(e => e.employeeId === String(doc[1])) || {};
        expiryWarnings.push({
          employeeId: doc[1],
          fullName: emp.fullName || "N/A",
          document: doc[2],
          expiryDate: formatDate(doc[5]),
          daysRemaining: diffDays,
          severity: diffDays <= 0 ? "EXPIRED" : "WARNING"
        });
      }
    }
  });

  return { status: "SUCCESS", kpi, employees, expiryWarnings };
}

function getProfileData(employeeId, requester) {
  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues();
  let emp = null;

  for (let i = 1; i < empRows.length; i++) {
    if (String(empRows[i][0]) === String(employeeId)) {
      emp = empRows[i];
      break;
    }
  }

  if (!emp) return { status: "ERROR", message: "Không tìm thấy hồ sơ." };

  const docSheet = getSheet(TABLES.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues().slice(1);
  const docs = docRows.filter(d => String(d[1]) === String(employeeId) && d[9] === true);

  const passport = docs.find(d => d[2] === "PASSPORT") || [];
  const trc = docs.find(d => d[2] === "TRC") || [];
  const wp = docs.find(d => d[2] === "WORK_PERMIT") || [];

  return {
    status: "SUCCESS",
    profile: {
      employeeId: emp[0],
      fullName: emp[1],
      nationality: emp[2],
      dob: formatDate(emp[3]),
      position: emp[4],
      department: emp[5],
      phone: emp[6],
      email: emp[7],
      currentStatus: emp[8],
      currentLocation: emp[9],
      passportNo: passport[3] || "",
      passportExpiry: formatDate(passport[5]) || "",
      passportImg: passport[8] || "",
      trcNo: trc[3] || "",
      trcExpiry: formatDate(trc[5]) || "",
      trcImg: trc[8] || "",
      wpNo: wp[3] || "",
      wpExpiry: formatDate(wp[5]) || "",
      wpImg: wp[8] || ""
    }
  };
}

// ==========================================
// F. HÀM TRỢ GIÚP KHỔNG LỒ & HỆ THỐNG
// ==========================================
function setupSystem() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const schema = {
    [TABLES.USERS]: ['UserID', 'Email', 'PasswordHash', 'Role', 'EmployeeID', 'Status', 'CreatedAt'],
    [TABLES.EMPLOYEES]: ['EmployeeID', 'FullName', 'Nationality', 'DOB', 'Position', 'Department', 'Phone', 'Email', 'CurrentStatus', 'CurrentLocation', 'ActiveStatus', 'FolderId', 'CreatedAt', 'UpdatedAt'],
    [TABLES.DOCUMENTS]: ['DocID', 'EmployeeID', 'DocType', 'DocNo', 'IssueDate', 'ExpiryDate', 'Issuer', 'FileId', 'FileUrl', 'IsCurrent', 'Status', 'CreatedAt', 'CreatedBy'],
    [TABLES.CONTRACTS]: ['ContractID', 'EmployeeID', 'ContractNo', 'ContractType', 'StartDate', 'EndDate', 'Salary', 'AllowancesJson', 'Status', 'IsCurrent', 'FileId', 'FileUrl', 'CreatedAt', 'CreatedBy'],
    [TABLES.APPENDICES]: ['AppendixID', 'ContractID', 'EmployeeID', 'AppendixNo', 'EffectiveDate', 'EndDate', 'NewSalary', 'AllowancesJson', 'Content', 'FileId', 'FileUrl', 'CreatedAt', 'CreatedBy'],
    [TABLES.ENTRY_EXIT]: ['LogID', 'EmployeeID', 'Type', 'EventDate', 'PortName', 'Note', 'CreatedAt', 'CreatedBy'],
    [TABLES.TRAVEL]: ['TravelID', 'EmployeeID', 'FromLocation', 'ToLocation', 'StartDate', 'EndDate', 'Status', 'TicketFileId', 'TicketUrl', 'CreatedAt', 'CreatedBy'],
    [TABLES.AUDIT]: ['AuditID', 'UserID', 'Action', 'Module', 'TargetID', 'OldValue', 'NewValue', 'Timestamp', 'IPAddress']
  };

  Object.keys(schema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(schema[sheetName]);
      sheet.getRange(1, 1, 1, schema[sheetName].length).setFontWeight("bold").setBackground("#f3f3f3");
    }
  });

  setupDriveFolders();
  
  // Khởi tạo tài khoản Admin mặc định nếu chưa có
  const userSheet = ss.getSheetByName(TABLES.USERS);
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(["USR-ADMIN", "admin@company.com", "admin123", ROLES.ADMIN, "EMP-ADMIN", "ACTIVE", new Date()]);
  }
}

function createEmployeeDriveFolder(employeeId, fullName) {
  const rootFolders = DriveApp.getFoldersByName("Foreign Employee Management System");
  let rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder("Foreign Employee Management System");
  
  const empFolder = rootFolder.createFolder(`${employeeId} - ${fullName}`);
  empFolder.createFolder("Passport");
  empFolder.createFolder("Visa");
  empFolder.createFolder("TRC");
  empFolder.createFolder("Work Permit");
  empFolder.createFolder("Contracts");
  
  return empFolder.getId();
}

function setupDriveFolders() {
  const folderName = "Foreign Employee Management System";
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    DriveApp.createFolder(folderName);
  }
}

function writeAuditLog(userId, action, module, targetId, oldValue, newValue) {
  const sheet = getSheet(TABLES.AUDIT);
  sheet.appendRow([
    "AUD-" + Utilities.getUuid().substring(0, 8),
    userId || "SYSTEM",
    action,
    module,
    targetId,
    oldValue || "",
    newValue || "",
    new Date(),
    ""
  ]);
}

function formatDate(d) {
  if (!d) return "";
  const date = new Date(d);
  if (isNaN(date.getTime())) return String(d);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function handleAddEntryExit(d, req) {
  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const logId = "EE-" + Utilities.getUuid().substring(0, 8);
  sheet.appendRow([logId, d.employeeId, d.type, d.dateTime, d.airport, d.purpose, new Date(), req.userId]);
  calculateEmployeeLocationStatus(d.employeeId);
  return { status: "SUCCESS", message: "Đã ghi nhận thông tin Nhập/Xuất cảnh!" };
}

function handleGetEntryExit(empId, req) {
  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => ({
    type: r[2], dateTime: formatDate(r[3]), airport: r[4], destination: r[5]
  }));
  return { status: "SUCCESS", records };
}

function handleAddDomesticTravel(d, req) {
  const sheet = getSheet(TABLES.TRAVEL);
  const id = "TRV-" + Utilities.getUuid().substring(0, 8);
  sheet.appendRow([id, d.employeeId, d.fromLocation, d.toLocation, d.fromDate, d.toDate, "APPROVED", "", "", new Date(), req.userId]);
  calculateEmployeeLocationStatus(d.employeeId);
  return { status: "SUCCESS", message: "Đã lưu lịch trình công tác!" };
}

function handleGetDomesticTravel(empId, req) {
  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => ({
    fromLocation: r[2], toLocation: r[3], fromDate: formatDate(r[4]), toDate: formatDate(r[5]), status: r[6]
  }));
  return { status: "SUCCESS", records };
}

function handleGetTimeline(empId, req) {
  const ee = handleGetEntryExit(empId, req).records || [];
  const tr = handleGetDomesticTravel(empId, req).records || [];
  const timeline = [];
  
  ee.forEach(x => timeline.push({ title: x.type === "ENTRY" ? "Nhập cảnh VN" : "Xuất cảnh VN", date: x.dateTime, location: x.airport }));
  tr.forEach(x => timeline.push({ title: "Công tác nội địa", date: x.fromDate, endDate: x.toDate, location: x.fromLocation + " → " + x.toLocation }));
  
  timeline.sort((a,b) => new Date(b.date) - new Date(a.date));
  return { status: "SUCCESS", timeline };
}

function handleDeleteEmployee(empId, req) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN].includes(req.role)) {
    return { status: "ERROR", message: "Không có quyền xóa nhân viên!" };
  }
  const sheet = getSheet(TABLES.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(empId)) {
      sheet.getRange(i + 1, 11).setValue("INACTIVE");
      writeAuditLog(req.userId, "DELETE", "Employees", empId, "ACTIVE", "INACTIVE");
      return { status: "SUCCESS", message: "Đã vô hiệu hóa tài khoản nhân viên." };
    }
  }
  return { status: "ERROR", message: "Không tìm thấy mã nhân viên." };
}
