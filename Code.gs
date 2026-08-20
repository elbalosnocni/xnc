// ==========================================
// A. DATA STRUCTURE & CONSTANTS
// ==========================================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SESSION_TTL_SECONDS = 21600; // 6 Giờ
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024; // 10 MB / file
const ALLOWED_UPLOAD_MIME = new Set([
  'application/pdf',
  'image/jpeg',
  'image/png',
  'image/webp'
]);
let _SS = null;
const _SHEET_CACHE = {};

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
  if (_SHEET_CACHE[tableName]) return _SHEET_CACHE[tableName];
  if (!_SS) _SS = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = _SS.getSheetByName(tableName);
  if (!sheet) {
    throw new Error(`Bảng ${tableName} chưa tồn tại. Vui lòng chạy setupSystem() thủ công.`);
  }
  _SHEET_CACHE[tableName] = sheet;
  return sheet;
}

// ==========================================
// B. ROUTER API (doGet & doPost & apiCall)
// ==========================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Foreign Employee Management V2")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let result = { status: "ERROR", message: "Yêu cầu không hợp lệ!" };
  try {
    let requestData = {};
    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      requestData = e.parameter;
    }
    
    const action = String(requestData.action || "").trim().toUpperCase();
    const token = requestData.token || "";
    const data = requestData.data || requestData;
    const params = requestData.params || e.parameter || {};

    if (action === "LOGIN") {
      const username = requestData.username || data.username;
      const password = requestData.password || data.password;
      result = handleLogin(username, password);
    } else {
      const requester = validateSession(token);
      result = requester
        ? dispatchApiAction(action, token, requester, data, params)
        : { status: "ERROR", message: "Phiên đăng nhập hết hạn hoặc không hợp lệ!" };
    }
  } catch (error) {
    result = { status: "ERROR", message: error.message || String(error) };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function apiCall(action, token, data, params) {
  try {
    action = String(action || "").trim().toUpperCase();
    data = data || {};
    params = params || {};

    if (action === "LOGIN") {
      return handleLogin(data.username, data.password);
    }

    const requester = validateSession(token);
    if (!requester) {
      return { status: "ERROR", message: "Phiên đăng nhập hết hạn hoặc không hợp lệ!" };
    }

    return dispatchApiAction(action, token, requester, data, params);
  } catch (error) {
    console.error(error);
    return { status: "ERROR", message: error && error.message ? error.message : String(error) };
  }
}

function dispatchApiAction(action, token, requester, data, params) {
  let result = { status: "ERROR", message: "Hành động không hợp lệ!" };
  const employeeId = params.employeeId || data.employeeId || (data.data && data.data.employeeId);

  switch (action) {
    case "LOGOUT":
      result = handleLogout(token);
      break;
    case "GET_DASHBOARD_DATA":
      result = getDashboardData(requester);
      break;
    case "GET_PROFILE":
      result = getProfileData(employeeId, requester);
      break;
    case "SAVE_PROFILE":
      result = handleSaveProfile(data, requester);
      break;
    case "SAVE_DOCUMENT":
      if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role)) {
        result = { status: "ERROR", message: "Không có quyền cập nhật giấy tờ." };
      } else {
        result = saveDocument(data, requester);
      }
      break;
    case "UPLOAD_FILE":
      result = uploadEmployeeFile(data, requester);
      break;
    case "ADD_ENTRY_EXIT":
      result = handleAddEntryExit(data, requester);
      break;
    case "GET_ENTRY_EXIT":
      result = handleGetEntryExit(employeeId, requester);
      break;
    case "ADD_DOMESTIC_TRAVEL":
      result = handleAddDomesticTravel(data, requester);
      break;
    case "APPROVE_DOMESTIC_TRAVEL":
      result = handleApproveDomesticTravel(data, requester);
      break;
    case "GET_DOMESTIC_TRAVEL":
      result = handleGetDomesticTravel(employeeId, requester);
      break;
    case "GET_TIMELINE":
      result = handleGetTimeline(employeeId, requester);
      break;
    case "DELETE_EMPLOYEE":
      result = handleDeleteEmployee(employeeId, requester);
      break;
    case "IMPORT_EMPLOYEES_EXCEL":
      result = handleImportEmployeesExcel(data, requester);
      break;
    case "CHANGE_PASSWORD":
      result = handleChangePassword(data, requester, token);
      break;
    case "GET_ACCOUNT_USERS":
      result = getAccountUsers(requester);
      break;
    case "ADMIN_RESET_PASSWORD":
      result = adminResetPassword(data, requester);
      break;
    case "SET_ACCOUNT_STATUS":
      result = setAccountStatus(data, requester);
      break;
  }

  return result;
}

// ==========================================
// C. AUTH & AUTHORIZATION
// ==========================================
function hashPassword(pass) {
  const salt = "HR_GLOBAL_SALT_2026";
  const raw = pass + salt;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return hash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function handleLogin(username, password) {
  username = String(username || "").trim();
  password = String(password || "");
  if (!username || !password) {
    return { status: "ERROR", message: "Vui lòng nhập tài khoản và mật khẩu." };
  }

  const sheet = getSheet(TABLES.USERS);
  const rows = sheet.getDataRange().getValues();
  const inputHash = hashPassword(password);
  
  for (let i = 1; i < rows.length; i++) {
    const [userId, email, passHash, role, employeeId, status] = rows[i];
    
    const loginMatches = (String(email).toLowerCase() === String(username).toLowerCase() || String(employeeId) === String(username));
    const hashMatches = String(passHash) === inputHash;
    const legacyPlainMatches = String(passHash) === String(password);
    if (loginMatches && (hashMatches || legacyPlainMatches)) {
      
      if (String(status).toUpperCase() !== "ACTIVE") {
        return { status: "ERROR", message: "Tài khoản của bạn đã bị vô hiệu hóa!" };
      }

      const empSheet = getSheet(TABLES.EMPLOYEES);
      const empRows = empSheet.getDataRange().getValues();
      const emp = empRows.find(r => String(r[0]) === String(employeeId));
      if (emp && emp[10] !== "ACTIVE") {
        return { status: "ERROR", message: "Hồ sơ nhân viên đã dừng hoạt động!" };
      }
      
      // Migrate any legacy plaintext password to PasswordHash after a successful login.
      if (legacyPlainMatches && !hashMatches) {
        sheet.getRange(i + 1, 3).setValue(inputHash);
      }

      const token = "TK-" + Utilities.getUuid();
      const userProps = PropertiesService.getScriptProperties();
      const sessionData = { userId, email, role, employeeId, createdAt: Date.now() };
      
      userProps.setProperty(token, JSON.stringify(sessionData));
      
      let fullName = email;
      if (emp) fullName = emp[1];

      writeAuditLog(userId, "LOGIN", "Users", userId, null, "User logged in successfully");

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

  const session = JSON.parse(raw);
  const now = Date.now();
  
  if ((now - session.createdAt) / 1000 > SESSION_TTL_SECONDS) {
    userProps.deleteProperty(token);
    return null;
  }

  return session;
}

function handleLogout(token) {
  PropertiesService.getScriptProperties().deleteProperty(token);
  return { status: "SUCCESS", message: "Đã đăng xuất thành công!" };
}

function canAccessEmployee(requester, targetEmpId) {
  if ([ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR, ROLES.DIRECTOR].includes(requester.role)) {
    return true;
  }
  return String(requester.employeeId) === String(targetEmpId);
}

// ==========================================
// D. QUẢN LÝ PROFILE & HỢP ĐỒNG & GIẤY TỜ
// ==========================================
function handleSaveProfile(data, requester) {
  data = data || {};
  data.employeeId = String(data.employeeId || "").trim();
  data.fullName = String(data.fullName || "").trim();
  const isHRAdmin = [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role);
  const isSelf = String(requester.employeeId) === String(data.employeeId);

  if (!data.employeeId) return { status: "ERROR", message: "Thiếu Mã nhân viên." };
  if (!isHRAdmin && !isSelf) {
    return { status: "ERROR", message: "Bạn không có quyền cập nhật hồ sơ này!" };
  }
  if (isHRAdmin && !data.fullName) {
    return { status: "ERROR", message: "Họ và tên là bắt buộc." };
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

  const timestamp = new Date();

  if (rowIndex > 0) {
    const currentRow = rows[rowIndex - 1];
    const updatedRow = [
      currentRow[0],
      isHRAdmin ? (data.fullName || currentRow[1]) : currentRow[1],
      isHRAdmin ? (data.nationality || currentRow[2]) : currentRow[2],
      isHRAdmin ? (formatDate(data.dob) || currentRow[3]) : currentRow[3],
      isHRAdmin ? (data.position || currentRow[4]) : currentRow[4],
      isHRAdmin ? (data.department || currentRow[5]) : currentRow[5],
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
    writeAuditLog(requester.userId, "UPDATE", "Employees", data.employeeId, JSON.stringify(currentRow), JSON.stringify(updatedRow));
  } else {
    if (!isHRAdmin) {
      return { status: "ERROR", message: "Chỉ HR mới được thêm mới nhân viên!" };
    }
    if (rows.slice(1).some(r => String(r[0]).trim() === data.employeeId)) {
      return { status: "ERROR", message: "Mã nhân viên đã tồn tại." };
    }
    if (data.email) {
      const normalizedEmail = String(data.email).trim().toLowerCase();
      const userSheet = getSheet(TABLES.USERS);
      const userRows = userSheet.getDataRange().getValues();
      if (userRows.slice(1).some(r => String(r[1]).trim().toLowerCase() === normalizedEmail)) {
        return { status: "ERROR", message: "Email tài khoản đã tồn tại." };
      }
    }

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

    const userSheet = getSheet(TABLES.USERS);
    userSheet.appendRow(["USR-" + Utilities.getUuid().substring(0, 6), data.email || data.employeeId, hashPassword("123456"), ROLES.USER, data.employeeId, "ACTIVE", timestamp]);
    writeAuditLog(requester.userId, "CREATE", "Employees", data.employeeId, null, JSON.stringify(newRow));
  }

  if (isHRAdmin) {
    if (data.passportNo) saveDocument({ employeeId: data.employeeId, docType: "PASSPORT", docNo: data.passportNo, expiryDate: data.passportExpiry, fileId: data.passportFileId || "", fileUrl: data.passportImg }, requester);
    if (data.trcNo) saveDocument({ employeeId: data.employeeId, docType: "TRC", docNo: data.trcNo, expiryDate: data.trcExpiry, fileId: data.trcFileId || "", fileUrl: data.trcImg }, requester);
    if (data.wpNo) saveDocument({ employeeId: data.employeeId, docType: "WORK_PERMIT", docNo: data.wpNo, expiryDate: data.wpExpiry, fileId: data.wpFileId || "", fileUrl: data.wpImg }, requester);

    if (data.contractNo) {
      saveContractAndAppendix(data, requester);
    }
  }

  return { status: "SUCCESS", message: "Đã lưu thông tin hồ sơ thành công!" };
}

function saveContractAndAppendix(data, requester) {
  const contractSheet = getSheet(TABLES.CONTRACTS);
  const contractRows = contractSheet.getDataRange().getValues();
  let currentContractId = "";
  
  for (let i = 1; i < contractRows.length; i++) {
    if (String(contractRows[i][1]) === String(data.employeeId) && contractRows[i][9] === true) {
      if (contractRows[i][2] === data.contractNo) {
        currentContractId = contractRows[i][0];
      } else {
        contractSheet.getRange(i + 1, 10).setValue(false);
      }
    }
  }

  if (!currentContractId) {
    currentContractId = "CTR-" + Utilities.getUuid().substring(0, 8);
    const newContractRow = [
      currentContractId,
      data.employeeId,
      data.contractNo,
      "DEFINITE",
      formatDate(data.contractStartDate),
      formatDate(data.contractExpiry),
      Number(data.salary) || 0,
      JSON.stringify({ allowance: Number(data.allowance) || 0 }),
      "VALID",
      true,
      "",
      data.contractFileId || "",
      data.contractImg || "",
      new Date(),
      requester.userId
    ];
    contractSheet.appendRow(newContractRow);
    writeAuditLog(requester.userId, "ADD_CONTRACT", "Contracts", currentContractId, null, JSON.stringify(newContractRow));
  } else {
    for (let i = 1; i < contractRows.length; i++) {
      if (contractRows[i][0] === currentContractId) {
        contractSheet.getRange(i + 1, 5).setValue(formatDate(data.contractStartDate));
        contractSheet.getRange(i + 1, 6).setValue(formatDate(data.contractExpiry));
        contractSheet.getRange(i + 1, 7).setValue(Number(data.salary) || 0);
        contractSheet.getRange(i + 1, 8).setValue(JSON.stringify({ allowance: Number(data.allowance) || 0 }));
        if (data.contractFileId) contractSheet.getRange(i + 1, 11).setValue(data.contractFileId);
        if (data.contractImg) contractSheet.getRange(i + 1, 12).setValue(data.contractImg);
        break;
      }
    }
  }

  if (data.appendixNo) {
    const appSheet = getSheet(TABLES.APPENDICES);
    const appRows = appSheet.getDataRange().getValues();
    let appExists = false;

    for (let i = 1; i < appRows.length; i++) {
      if (String(appRows[i][1]) === currentContractId && appRows[i][3] === data.appendixNo) {
        appExists = true;
        break;
      }
    }

    if (!appExists) {
      const appKey = "APP-" + Utilities.getUuid().substring(0, 8);
      const newAppRow = [
        appKey,
        currentContractId,
        data.employeeId,
        data.appendixNo,
        formatDate(data.contractStartDate),
        formatDate(data.contractExpiry),
        Number(data.salary) || 0,
        JSON.stringify({ allowance: Number(data.allowance) || 0 }),
        "Cập nhật theo hồ sơ",
        data.appendixFileId || "",
        data.contractImg || "",
        new Date(),
        requester.userId
      ];
      appSheet.appendRow(newAppRow);
      writeAuditLog(requester.userId, "ADD_APPENDIX", "ContractAppendices", appKey, null, JSON.stringify(newAppRow));
    }
  }
}

function saveDocument(docData, requester) {
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

// ==========================================
// E. DRIVE FILE UPLOAD & BẢO MẬT
// ==========================================
function uploadEmployeeFile(param, requester) {
  try {
    const { base64Data, fileName, mimeType, employeeId, subFolderType } = param;
    
    if (!canAccessEmployee(requester, employeeId)) {
      return { status: "ERROR", message: "Bạn không có quyền upload file cho nhân viên này!" };
    }

    const empSheet = getSheet(TABLES.EMPLOYEES);
    const emps = empSheet.getDataRange().getValues();
    let folderId = "";

    for (let i = 1; i < emps.length; i++) {
      if (String(emps[i][0]) === String(employeeId)) {
        folderId = emps[i][11];
        break;
      }
    }

    if (!folderId) {
      return { status: "ERROR", message: "Nhân viên chưa có thư mục Drive." };
    }
    const parentFolder = DriveApp.getFolderById(folderId);
    const folderName = String(subFolderType || "Misc").replace(/[\\/]/g, "").trim() || "Misc";
    const subFolders = parentFolder.getFoldersByName(folderName);
    const targetFolder = subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(folderName);

    const encoded = String(base64Data).split(",")[1] || String(base64Data);
    const bytes = Utilities.base64Decode(encoded);
    if (bytes.length > MAX_UPLOAD_BYTES) {
      return { status: "ERROR", message: "File vượt quá giới hạn 10 MB." };
    }
    const safeName = String(fileName || "upload").replace(/[\\/:*?"<>|]/g, "_").trim() || "upload";
    const safeMime = String(mimeType || "").toLowerCase();
    if (!ALLOWED_UPLOAD_MIME.has(safeMime)) {
      return { status: "ERROR", message: "Định dạng file không được hỗ trợ. Chỉ nhận PDF, JPG, PNG, WEBP." };
    }
    const blob = Utilities.newBlob(bytes, safeMime, safeName);
    const file = targetFolder.createFile(blob);
    
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    return {
      status: "SUCCESS",
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };
  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

// ==========================================
// F. LỊCH TRÌNH & XUẤT NHẬP CẢNH & DỰ ĐỊNH
// ==========================================
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

function handleAddEntryExit(d, req) {
  if (!canAccessEmployee(req, d.employeeId)) return { status: "ERROR", message: "Không có quyền thực hiện." };
  
  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const logId = "EE-" + Utilities.getUuid().substring(0, 8);
  const note = JSON.stringify({ flightNo: d.flightNo || "", destination: d.destination || "", purpose: d.purpose || "" });
  
  sheet.appendRow([logId, d.employeeId, d.type, d.dateTime, d.airport, note, new Date(), req.userId]);
  calculateEmployeeLocationStatus(d.employeeId);
  return { status: "SUCCESS", message: "Đã ghi nhận thông tin Nhập/Xuất cảnh!" };
}

function handleGetEntryExit(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem thông tin." };

  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => {
    let noteObj = {};
    try { noteObj = JSON.parse(r[5] || "{}"); } catch(e) { noteObj = { purpose: r[5] }; }
    return {
      type: r[2],
      dateTime: formatDate(r[3]),
      airport: r[4],
      flightNo: noteObj.flightNo || "",
      destination: noteObj.destination || "",
      purpose: noteObj.purpose || ""
    };
  });
  return { status: "SUCCESS", records };
}

function handleAddDomesticTravel(d, req) {
  d = d || {};
  d.employeeId = String(d.employeeId || "").trim();
  d.fromLocation = String(d.fromLocation || "").trim();
  d.toLocation = String(d.toLocation || "").trim();
  if (!d.employeeId || !d.fromDate || !d.toDate || !d.fromLocation || !d.toLocation) {
    return { status: "ERROR", message: "Thiếu thông tin lịch trình." };
  }
  const start = new Date(d.fromDate);
  const end = new Date(d.toDate);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || end < start) {
    return { status: "ERROR", message: "Khoảng ngày đi lại không hợp lệ." };
  }
  if (!canAccessEmployee(req, d.employeeId)) return { status: "ERROR", message: "Không có quyền thực hiện." };

  const sheet = getSheet(TABLES.TRAVEL);
  const id = "TRV-" + Utilities.getUuid().substring(0, 8);
  
  const isHR = [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(req.role);
  const initialStatus = isHR ? "APPROVED" : "PENDING";

  sheet.appendRow([
    id,
    d.employeeId,
    d.fromLocation,
    d.toLocation,
    d.fromDate,
    d.toDate,
    initialStatus,
    "",
    "",
    new Date(),
    req.userId
  ]);
  
  calculateEmployeeLocationStatus(d.employeeId);
  return { status: "SUCCESS", message: isHR ? "Đã duyệt lịch trình công tác!" : "Đã gửi yêu cầu công tác (Chờ duyệt)!" };
}

function handleApproveDomesticTravel(d, req) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR, ROLES.DIRECTOR].includes(req.role)) {
    return { status: "ERROR", message: "Không có quyền duyệt yêu cầu!" };
  }
  const status = String(d && d.status ? d.status : "").toUpperCase();
  const travelId = String(d && d.travelId ? d.travelId : "").trim();
  if (!travelId || !["APPROVED", "REJECTED"].includes(status)) {
    return { status: "ERROR", message: "Trạng thái duyệt không hợp lệ." };
  }

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.travelId) {
      sheet.getRange(i + 1, 7).setValue(status);
      calculateEmployeeLocationStatus(rows[i][1]);
      return { status: "SUCCESS", message: `Đã cập nhật trạng thái yêu cầu thành: ${status}` };
    }
  }
  return { status: "ERROR", message: "Không tìm thấy mã yêu cầu!" };
}

function handleGetDomesticTravel(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem thông tin." };

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => ({
    travelId: r[0],
    fromLocation: r[2],
    toLocation: r[3],
    fromDate: formatDate(r[4]),
    toDate: formatDate(r[5]),
    status: r[6]
  }));
  return { status: "SUCCESS", records };
}

// ==========================================
// G. DASHBOARD & PROFILE RESPONSE
// ==========================================
function getDashboardData(requester) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR, ROLES.DIRECTOR].includes(requester.role)) {
    return { status: "ERROR", message: "Không có quyền truy cập dữ liệu quản trị." };
  }

  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues().slice(1);
  
  const docSheet = getSheet(TABLES.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues().slice(1);

  const ctrSheet = getSheet(TABLES.CONTRACTS);
  const ctrRows = ctrSheet.getDataRange().getValues().slice(1);

  const kpi = { total: 0, inVN: 0, business: 0, exited: 0, warning: 0, expContract: 0 };
  const employees = [];
  const expiryWarnings = [];
  const now = new Date();

  const activeContractsMap = {};
  ctrRows.forEach(c => {
    if (c[9] === true) {
      activeContractsMap[String(c[1])] = {
        contractNo: c[2],
        endDate: c[5]
      };
    }
  });

  empRows.forEach(row => {
    if (row[10] === "ACTIVE") {
      kpi.total++;
      if (row[8] === "In Vietnam") kpi.inVN++;
      if (row[8] === "Traveling") kpi.business++;
      if (row[8] === "Exited") kpi.exited++;

      const ctr = activeContractsMap[String(row[0])] || {};

      employees.push({
        employeeId: row[0],
        fullName: row[1],
        nationality: row[2],
        position: row[4],
        department: row[5],
        currentStatus: row[8],
        currentLocation: row[9],
        contractNo: ctr.contractNo || "-",
        contractExpiry: ctr.endDate ? formatDate(ctr.endDate) : "-"
      });
    }
  });

  docRows.forEach(doc => {
    if (doc[9] === true && doc[5]) {
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

  ctrRows.forEach(ctr => {
    if (ctr[9] === true && ctr[5]) {
      const expDate = new Date(ctr[5]);
      const diffTime = expDate - now;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays <= 45) {
        kpi.expContract++;
        const emp = employees.find(e => e.employeeId === String(ctr[1])) || {};
        expiryWarnings.push({
          employeeId: ctr[1],
          fullName: emp.fullName || "N/A",
          document: "CONTRACT (" + ctr[2] + ")",
          expiryDate: formatDate(ctr[5]),
          daysRemaining: diffDays,
          severity: diffDays <= 0 ? "EXPIRED" : "WARNING"
        });
      }
    }
  });

  return { status: "SUCCESS", kpi, employees, expiryWarnings };
}

function getProfileData(employeeId, requester) {
  if (!canAccessEmployee(requester, employeeId)) {
    return { status: "ERROR", message: "Bạn không có quyền xem thông tin hồ sơ này!" };
  }

  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues();
  let emp = null;

  for (let i = 1; i < empRows.length; i++) {
    if (String(empRows[i][0]) === String(employeeId)) {
      emp = empRows[i];
      break;
    }
  }

  if (!emp) return { status: "ERROR", message: "Không tìm thấy hồ sơ nhân viên." };

  const docSheet = getSheet(TABLES.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues().slice(1);
  const docs = docRows.filter(d => String(d[1]) === String(employeeId) && d[9] === true);

  const passport = docs.find(d => d[2] === "PASSPORT") || [];
  const trc = docs.find(d => d[2] === "TRC") || [];
  const wp = docs.find(d => d[2] === "WORK_PERMIT") || [];

  const ctrSheet = getSheet(TABLES.CONTRACTS);
  const ctrRows = ctrSheet.getDataRange().getValues().slice(1);
  const currentContract = ctrRows.find(c => String(c[1]) === String(employeeId) && c[9] === true) || [];

  const appSheet = getSheet(TABLES.APPENDICES);
  const appRows = appSheet.getDataRange().getValues().slice(1);
  const appendices = appRows
    .filter(a => String(a[2]) === String(employeeId))
    .sort((a, b) => new Date(b[4] || 0) - new Date(a[4] || 0));
  const latestApp = appendices.length > 0 ? appendices[0] : [];

  let allowance = 0;
  try {
    const allowanceObj = JSON.parse(currentContract[7] || "{}");
    allowance = allowanceObj.allowance || 0;
  } catch(e) {}

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
      passportFileId: passport[7] || "",
      passportImg: passport[8] || "",
      trcNo: trc[3] || "",
      trcExpiry: formatDate(trc[5]) || "",
      trcFileId: trc[7] || "",
      trcImg: trc[8] || "",
      wpNo: wp[3] || "",
      wpExpiry: formatDate(wp[5]) || "",
      wpFileId: wp[7] || "",
      wpImg: wp[8] || "",
      contractNo: currentContract[2] || "",
      appendixNo: latestApp[3] || "",
      contractStartDate: formatDate(currentContract[4]) || "",
      contractExpiry: formatDate(currentContract[5]) || "",
      contractFileId: currentContract[10] || "",
      contractImg: currentContract[11] || "",
      salary: currentContract[6] || 0,
      allowance: allowance
    }
  };
}


// ==========================================
// H. IMPORT EXCEL -> EMPLOYEES
// ==========================================

function normalizeImportText(value) {
  return String(value == null ? "" : value).trim();
}

function normalizeImportDate(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return formatDate(value);
  }
  const s = String(value).trim();
  if (!s) return "";
  const ymd = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (ymd) return `${ymd[1]}-${String(ymd[2]).padStart(2, "0")}-${String(ymd[3]).padStart(2, "0")}`;
  const dmy = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (dmy) return `${dmy[3]}-${String(dmy[2]).padStart(2, "0")}-${String(dmy[1]).padStart(2, "0")}`;
  return formatDate(s);
}

function handleImportEmployeesExcel(data, requester) {
  const allowedRoles = [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR];
  if (!allowedRoles.includes(requester.role)) {
    return { status: "ERROR", message: "Chỉ HR/HR ADMIN/ADMIN được import danh sách nhân viên." };
  }

  const rows = Array.isArray(data && data.rows) ? data.rows : [];
  if (!rows.length) return { status: "ERROR", message: "Không có dữ liệu Excel để import." };
  if (rows.length > 500) return { status: "ERROR", message: "Mỗi lần import tối đa 500 dòng." };

  const sheet = getSheet(TABLES.EMPLOYEES);
  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const existingById = {};
  const existingEmail = {};

  for (let i = 1; i < values.length; i++) {
    const id = normalizeImportText(values[i][0]);
    const email = normalizeImportText(values[i][7]).toLowerCase();
    if (id) existingById[id] = { rowIndex: i + 1, row: values[i] };
    if (email) existingEmail[email] = { rowIndex: i + 1, employeeId: id };
  }

  const result = {
    status: "SUCCESS",
    total: 0,
    created: 0,
    updated: 0,
    skipped: 0,
    documentsCreated: 0,
    documentsUpdated: 0,
    contractsCreated: 0,
    contractsUpdated: 0,
    appendicesCreated: 0,
    appendicesUpdated: 0,
    errors: []
  };

  const newRows = [];
  const updateRows = [];
  const importedIds = new Set();

  rows.forEach((item, idx) => {
    const excelRow = Number(item._excelRow || idx + 2);
    const employeeId = normalizeImportText(item.employeeId);
    const fullName = normalizeImportText(item.fullName);

    if (!employeeId || !fullName) {
      result.skipped++;
      result.errors.push({ row: excelRow, message: "Thiếu EmployeeID hoặc FullName." });
      return;
    }
    if (importedIds.has(employeeId)) {
      result.skipped++;
      result.errors.push({ row: excelRow, message: `Trùng EmployeeID ${employeeId} trong file.` });
      return;
    }
    importedIds.add(employeeId);

    const email = normalizeImportText(item.email).toLowerCase();
    if (email && existingEmail[email] && existingEmail[email].employeeId !== employeeId) {
      result.skipped++;
      result.errors.push({ row: excelRow, message: `Email ${email} đã thuộc EmployeeID ${existingEmail[email].employeeId}.` });
      return;
    }

    const current = existingById[employeeId];
    if (current) {
      const r = current.row.slice();
      r[0] = employeeId;
      r[1] = fullName;
      if (item.nationality !== undefined && item.nationality !== "") r[2] = normalizeImportText(item.nationality);
      if (item.dob !== undefined && item.dob !== "") r[3] = normalizeImportDate(item.dob);
      if (item.position !== undefined && item.position !== "") r[4] = normalizeImportText(item.position);
      if (item.department !== undefined && item.department !== "") r[5] = normalizeImportText(item.department);
      if (item.phone !== undefined && item.phone !== "") r[6] = normalizeImportText(item.phone);
      if (item.email !== undefined && item.email !== "") r[7] = email;
      if (item.currentStatus !== undefined && item.currentStatus !== "") r[8] = normalizeImportText(item.currentStatus);
      if (item.currentLocation !== undefined && item.currentLocation !== "") r[9] = normalizeImportText(item.currentLocation);
      if (item.activeStatus !== undefined && item.activeStatus !== "") r[10] = normalizeImportText(item.activeStatus).toUpperCase();
      r[13] = now;
      updateRows.push({ rowIndex: current.rowIndex, row: r, oldRow: current.row });
      result.updated++;
      result.total++;
      return;
    }

    const newFolderId = createEmployeeDriveFolder(employeeId, fullName);
    const newRow = [
      employeeId,
      fullName,
      normalizeImportText(item.nationality),
      normalizeImportDate(item.dob),
      normalizeImportText(item.position),
      normalizeImportText(item.department),
      normalizeImportText(item.phone),
      email,
      normalizeImportText(item.currentStatus || "Exited"),
      normalizeImportText(item.currentLocation || "Overseas"),
      normalizeImportText(item.activeStatus || "ACTIVE").toUpperCase(),
      newFolderId,
      now,
      now
    ];
    newRows.push(newRow);
    existingById[employeeId] = { rowIndex: sheet.getLastRow() + newRows.length, row: newRow };
    if (email) existingEmail[email] = { rowIndex: sheet.getLastRow() + newRows.length, employeeId: employeeId };
    result.created++;
    result.total++;
  });

  updateRows.forEach(x => {
    sheet.getRange(x.rowIndex, 1, 1, 14).setValues([x.row]);
    writeAuditLog(requester.userId, "IMPORT_UPDATE", "Employees", x.row[0], JSON.stringify(x.oldRow), JSON.stringify(x.row));
  });

  if (newRows.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 14).setValues(newRows);

    const userSheet = getSheet(TABLES.USERS);
    const userRows = userSheet.getDataRange().getValues();
    const userKeys = new Set();
    for (let i = 1; i < userRows.length; i++) {
      userKeys.add(String(userRows[i][1] || "").trim().toLowerCase());
      userKeys.add(String(userRows[i][4] || "").trim());
    }

    const newUsers = [];
    newRows.forEach(r => {
      const email = String(r[7] || "").trim().toLowerCase();
      const employeeId = String(r[0]).trim();
      const login = email || employeeId;
      if (!userKeys.has(email) && !userKeys.has(employeeId)) {
        newUsers.push([
          "USR-" + Utilities.getUuid().substring(0, 6),
          login,
          hashPassword("123456"),
          ROLES.USER,
          employeeId,
          "ACTIVE",
          now
        ]);
        userKeys.add(email);
        userKeys.add(employeeId);
      }
      writeAuditLog(requester.userId, "IMPORT_CREATE", "Employees", employeeId, null, JSON.stringify(r));
    });
    if (newUsers.length) {
      userSheet.getRange(userSheet.getLastRow() + 1, 1, newUsers.length, 7).setValues(newUsers);
    }
    writeAuditLog(requester.userId, "IMPORT_BATCH", "Employees", "BATCH", null, `Imported ${newRows.length} rows`);
  }

  // ============================================================
  // IMPORT PHẦN HỒ SƠ PHÁP LÝ + HỢP ĐỒNG + PHỤ LỤC
  // Frontend đã map sẵn các cột này nhưng phiên bản cũ chỉ ghi Employees.
  // ============================================================
  rows.forEach((item, idx) => {
    const excelRow = Number(item._excelRow || idx + 2);
    const employeeId = normalizeImportText(item.employeeId);
    if (!employeeId || !existingById[employeeId]) return;

    try {
      // Passport
      if (item.passportNo || item.passportExpiry || item.passportIssueDate || item.passportIssuer || item.passportFileId || item.passportFileUrl) {
        const docResult = upsertImportedDocument({
          employeeId: employeeId,
          docType: "PASSPORT",
          docNo: item.passportNo,
          issueDate: item.passportIssueDate,
          expiryDate: item.passportExpiry,
          issuer: item.passportIssuer,
          fileId: item.passportFileId,
          fileUrl: item.passportFileUrl,
          status: item.passportStatus || "VALID"
        }, requester);
        result[docResult.created ? "documentsCreated" : "documentsUpdated"]++;
      }

      // TRC
      if (item.trcNo || item.trcExpiry || item.trcIssueDate || item.trcIssuer || item.trcFileId || item.trcFileUrl) {
        const docResult = upsertImportedDocument({
          employeeId: employeeId,
          docType: "TRC",
          docNo: item.trcNo,
          issueDate: item.trcIssueDate,
          expiryDate: item.trcExpiry,
          issuer: item.trcIssuer,
          fileId: item.trcFileId,
          fileUrl: item.trcFileUrl,
          status: item.trcStatus || "VALID"
        }, requester);
        result[docResult.created ? "documentsCreated" : "documentsUpdated"]++;
      }

      // Work Permit
      if (item.wpNo || item.wpExpiry || item.wpIssueDate || item.wpIssuer || item.wpFileId || item.wpFileUrl) {
        const docResult = upsertImportedDocument({
          employeeId: employeeId,
          docType: "WORK_PERMIT",
          docNo: item.wpNo,
          issueDate: item.wpIssueDate,
          expiryDate: item.wpExpiry,
          issuer: item.wpIssuer,
          fileId: item.wpFileId,
          fileUrl: item.wpFileUrl,
          status: item.wpStatus || "VALID"
        }, requester);
        result[docResult.created ? "documentsCreated" : "documentsUpdated"]++;
      }

      // Contract + Appendix
      if (item.contractNo) {
        const contractResult = upsertImportedContract(item, requester);
        result[contractResult.created ? "contractsCreated" : "contractsUpdated"]++;

        if (item.appendixNo) {
          const appendixResult = upsertImportedAppendix(item, contractResult.contractId, requester);
          result[appendixResult.created ? "appendicesCreated" : "appendicesUpdated"]++;
        }
      } else if (item.appendixNo) {
        result.errors.push({ row: excelRow, message: "Có AppendixNo nhưng thiếu ContractNo nên không thể tạo phụ lục." });
      }
    } catch (e) {
      result.errors.push({ row: excelRow, message: e && e.message ? e.message : String(e) });
    }
  });

  return result;
}

/**
 * Import/upsert một giấy tờ theo EmployeeID + DocType + DocNo.
 * Nếu cùng loại giấy tờ đã có bản ghi hiện tại nhưng số giấy tờ thay đổi,
 * bản ghi cũ sẽ được đánh dấu IsCurrent=false và bản ghi mới trở thành hiện tại.
 */
function upsertImportedDocument(doc, requester) {
  const sheet = getSheet(TABLES.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  const employeeId = String(doc.employeeId || "").trim();
  const docType = String(doc.docType || "").trim().toUpperCase();
  const docNo = normalizeImportText(doc.docNo);

  let exactIndex = -1;
  let currentIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== employeeId || String(rows[i][2]).toUpperCase() !== docType) continue;
    if (rows[i][9] === true && currentIndex < 0) currentIndex = i + 1;
    if (docNo && String(rows[i][3]).trim() === docNo) exactIndex = i + 1;
  }

  if (exactIndex > 0) {
    const oldRow = rows[exactIndex - 1];
    // Bản ghi chính xác được cập nhật và trở thành bản ghi hiện tại.
    if (currentIndex > 0 && currentIndex !== exactIndex) {
      sheet.getRange(currentIndex, 10).setValue(false);
    }

    const updated = oldRow.slice();
    updated[1] = employeeId;
    updated[2] = docType;
    updated[3] = docNo;
    updated[4] = normalizeImportDate(doc.issueDate);
    updated[5] = normalizeImportDate(doc.expiryDate);
    updated[6] = normalizeImportText(doc.issuer);
    updated[7] = normalizeImportText(doc.fileId);
    updated[8] = normalizeImportText(doc.fileUrl);
    updated[9] = true;
    updated[10] = normalizeImportText(doc.status || "VALID").toUpperCase();
    updated[11] = oldRow[11] || new Date();
    updated[12] = requester.userId;
    sheet.getRange(exactIndex, 1, 1, 13).setValues([updated]);
    writeAuditLog(requester.userId, "IMPORT_UPDATE", "Documents", updated[0], JSON.stringify(oldRow), JSON.stringify(updated));
    return { created: false, docId: updated[0] };
  }

  // Không có bản ghi trùng số: đóng bản ghi hiện tại của cùng loại.
  if (currentIndex > 0) sheet.getRange(currentIndex, 10).setValue(false);

  const docId = "DOC-" + Utilities.getUuid().substring(0, 8);
  const newRow = [
    docId,
    employeeId,
    docType,
    docNo,
    normalizeImportDate(doc.issueDate),
    normalizeImportDate(doc.expiryDate),
    normalizeImportText(doc.issuer),
    normalizeImportText(doc.fileId),
    normalizeImportText(doc.fileUrl),
    true,
    normalizeImportText(doc.status || "VALID").toUpperCase(),
    new Date(),
    requester.userId
  ];
  sheet.appendRow(newRow);
  writeAuditLog(requester.userId, "IMPORT_CREATE", "Documents", docId, null, JSON.stringify(newRow));
  return { created: true, docId: docId };
}

/**
 * Import/upsert hợp đồng theo EmployeeID + ContractNo.
 */
function upsertImportedContract(item, requester) {
  const sheet = getSheet(TABLES.CONTRACTS);
  const rows = sheet.getDataRange().getValues();
  const employeeId = normalizeImportText(item.employeeId);
  const contractNo = normalizeImportText(item.contractNo);

  let exactIndex = -1;
  let currentIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== employeeId) continue;
    if (rows[i][9] === true && currentIndex < 0) currentIndex = i + 1;
    if (String(rows[i][2]).trim() === contractNo) exactIndex = i + 1;
  }

  const allowancesJson = buildAllowancesJson(item.allowance, item.allowancesJson);
  const status = normalizeImportText(item.contractStatus || "VALID").toUpperCase();

  if (exactIndex > 0) {
    const oldRow = rows[exactIndex - 1];
    if (currentIndex > 0 && currentIndex !== exactIndex) {
      sheet.getRange(currentIndex, 10).setValue(false);
    }

    const updated = oldRow.slice();
    updated[1] = employeeId;
    updated[2] = contractNo;
    updated[3] = normalizeImportText(item.contractType || "DEFINITE").toUpperCase();
    updated[4] = normalizeImportDate(item.contractStartDate);
    updated[5] = normalizeImportDate(item.contractExpiry);
    updated[6] = normalizeImportNumber(item.salary);
    updated[7] = allowancesJson;
    updated[8] = status;
    updated[9] = true;
    updated[10] = normalizeImportText(item.contractFileId);
    updated[11] = normalizeImportText(item.contractFileUrl);
    updated[12] = oldRow[12] || new Date();
    updated[13] = requester.userId;
    sheet.getRange(exactIndex, 1, 1, 14).setValues([updated]);
    writeAuditLog(requester.userId, "IMPORT_UPDATE", "Contracts", updated[0], JSON.stringify(oldRow), JSON.stringify(updated));
    return { created: false, contractId: updated[0] };
  }

  if (currentIndex > 0) sheet.getRange(currentIndex, 10).setValue(false);

  const contractId = "CTR-" + Utilities.getUuid().substring(0, 8);
  const newRow = [
    contractId,
    employeeId,
    contractNo,
    normalizeImportText(item.contractType || "DEFINITE").toUpperCase(),
    normalizeImportDate(item.contractStartDate),
    normalizeImportDate(item.contractExpiry),
    normalizeImportNumber(item.salary),
    allowancesJson,
    status,
    true,
    normalizeImportText(item.contractFileId),
    normalizeImportText(item.contractFileUrl),
    new Date(),
    requester.userId
  ];
  sheet.appendRow(newRow);
  writeAuditLog(requester.userId, "IMPORT_CREATE", "Contracts", contractId, null, JSON.stringify(newRow));
  return { created: true, contractId: contractId };
}

/**
 * Import/upsert phụ lục theo EmployeeID + ContractID + AppendixNo.
 */
function upsertImportedAppendix(item, contractId, requester) {
  const sheet = getSheet(TABLES.APPENDICES);
  const rows = sheet.getDataRange().getValues();
  const employeeId = normalizeImportText(item.employeeId);
  const appendixNo = normalizeImportText(item.appendixNo);

  let exactIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(contractId) &&
        String(rows[i][2]) === employeeId &&
        String(rows[i][3]).trim() === appendixNo) {
      exactIndex = i + 1;
      break;
    }
  }

  const allowancesJson = buildAllowancesJson(item.appendixAllowance, item.appendixAllowancesJson);
  const newSalary = normalizeImportNumber(item.appendixSalary !== undefined && item.appendixSalary !== "" ? item.appendixSalary : item.salary);
  const content = normalizeImportText(item.appendixContent);

  if (exactIndex > 0) {
    const oldRow = rows[exactIndex - 1];
    const updated = oldRow.slice();
    updated[1] = contractId;
    updated[2] = employeeId;
    updated[3] = appendixNo;
    updated[4] = normalizeImportDate(item.appendixEffectiveDate);
    updated[5] = normalizeImportDate(item.appendixEndDate);
    updated[6] = newSalary;
    updated[7] = allowancesJson;
    updated[8] = content;
    updated[9] = normalizeImportText(item.appendixFileId);
    updated[10] = normalizeImportText(item.appendixFileUrl);
    updated[11] = oldRow[11] || new Date();
    updated[12] = requester.userId;
    sheet.getRange(exactIndex, 1, 1, 13).setValues([updated]);
    writeAuditLog(requester.userId, "IMPORT_UPDATE", "ContractAppendices", updated[0], JSON.stringify(oldRow), JSON.stringify(updated));
    return { created: false, appendixId: updated[0] };
  }

  const appendixId = "APP-" + Utilities.getUuid().substring(0, 8);
  const newRow = [
    appendixId,
    contractId,
    employeeId,
    appendixNo,
    normalizeImportDate(item.appendixEffectiveDate),
    normalizeImportDate(item.appendixEndDate),
    newSalary,
    allowancesJson,
    content,
    normalizeImportText(item.appendixFileId),
    normalizeImportText(item.appendixFileUrl),
    new Date(),
    requester.userId
  ];
  sheet.appendRow(newRow);
  writeAuditLog(requester.userId, "IMPORT_CREATE", "ContractAppendices", appendixId, null, JSON.stringify(newRow));
  return { created: true, appendixId: appendixId };
}

function normalizeImportNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && isFinite(value)) return value;
  const s = String(value).trim().replace(/,/g, "");
  if (!s) return 0;
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function buildAllowancesJson(value, jsonValue) {
  const rawJson = normalizeImportText(jsonValue);
  if (rawJson) {
    try {
      JSON.parse(rawJson);
      return rawJson;
    } catch (e) {}
  }
  return JSON.stringify({ allowance: normalizeImportNumber(value) });
}

// ==========================================
// H. THỦ TỤC NGHỈ VIỆC & KHÓA TÀI KHOẢN
// ==========================================
function handleDeleteEmployee(empId, req) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN].includes(req.role)) {
    return { status: "ERROR", message: "Không có quyền thực hiện thao tác này!" };
  }

  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < empRows.length; i++) {
    if (String(empRows[i][0]) === String(empId)) {
      empSheet.getRange(i + 1, 11).setValue("INACTIVE");
      found = true;
      break;
    }
  }

  if (!found) return { status: "ERROR", message: "Không tìm thấy nhân viên." };

  const userSheet = getSheet(TABLES.USERS);
  const userRows = userSheet.getDataRange().getValues();
  for (let i = 1; i < userRows.length; i++) {
    if (String(userRows[i][4]) === String(empId)) {
      userSheet.getRange(i + 1, 6).setValue("INACTIVE");
      break;
    }
  }

  writeAuditLog(req.userId, "DEACTIVATE", "Employees", empId, "ACTIVE", "INACTIVE");
  return { status: "SUCCESS", message: "Đã vô hiệu hóa hồ sơ và khóa tài khoản thành công." };
}

// ==========================================
// I. ACCOUNT / PASSWORD MANAGEMENT
// ==========================================
function validateNewPassword(password) {
  password = String(password || "");
  if (password.length < 8) return { ok: false, message: "Mật khẩu mới phải có ít nhất 8 ký tự." };
  if (password.length > 128) return { ok: false, message: "Mật khẩu mới tối đa 128 ký tự." };
  return { ok: true };
}

function findUserById(userId) {
  const sheet = getSheet(TABLES.USERS);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(userId)) {
      return { sheet, rowIndex: i + 1, row: rows[i] };
    }
  }
  return null;
}

function invalidateUserSessions(userId) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(key => {
    if (!/^TK-/.test(key)) return;
    try {
      const session = JSON.parse(all[key]);
      if (String(session.userId) === String(userId)) props.deleteProperty(key);
    } catch (e) {}
  });
}

function handleChangePassword(data, requester, token) {
  data = data || {};
  const currentPassword = String(data.currentPassword || "");
  const newPassword = String(data.newPassword || "");
  const confirmPassword = String(data.confirmPassword || "");

  if (!currentPassword || !newPassword || !confirmPassword) {
    return { status: "ERROR", message: "Vui lòng nhập đầy đủ mật khẩu hiện tại và mật khẩu mới." };
  }
  if (newPassword !== confirmPassword) {
    return { status: "ERROR", message: "Mật khẩu mới và xác nhận mật khẩu không giống nhau." };
  }
  const valid = validateNewPassword(newPassword);
  if (!valid.ok) return { status: "ERROR", message: valid.message };
  if (currentPassword === newPassword) {
    return { status: "ERROR", message: "Mật khẩu mới phải khác mật khẩu hiện tại." };
  }

  const found = findUserById(requester.userId);
  if (!found) return { status: "ERROR", message: "Không tìm thấy tài khoản." };
  const storedHash = String(found.row[2] || "");
  if (storedHash !== hashPassword(currentPassword)) {
    return { status: "ERROR", message: "Mật khẩu hiện tại không chính xác." };
  }
  if (String(found.row[5]).toUpperCase() !== "ACTIVE") {
    return { status: "ERROR", message: "Tài khoản đang bị khóa." };
  }

  found.sheet.getRange(found.rowIndex, 3).setValue(hashPassword(newPassword));
  writeAuditLog(requester.userId, "CHANGE_PASSWORD", "Users", requester.userId, null, "Password changed");
  return { status: "SUCCESS", message: "Đổi mật khẩu thành công." };
}

function getAccountUsers(requester) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role)) {
    return { status: "ERROR", message: "Bạn không có quyền xem danh sách tài khoản." };
  }

  const userSheet = getSheet(TABLES.USERS);
  const userRows = userSheet.getDataRange().getValues().slice(1);
  const empSheet = getSheet(TABLES.EMPLOYEES);
  const empRows = empSheet.getDataRange().getValues().slice(1);
  const empMap = {};
  empRows.forEach(r => empMap[String(r[0])] = { fullName: r[1] || "", activeStatus: r[10] || "" });

  const users = userRows.map(r => ({
    userId: r[0],
    email: r[1],
    role: r[3],
    employeeId: r[4],
    status: String(r[5] || "").toUpperCase(),
    fullName: (empMap[String(r[4])] || {}).fullName || r[1] || String(r[4] || ""),
    employeeActiveStatus: (empMap[String(r[4])] || {}).activeStatus || ""
  }));

  return { status: "SUCCESS", users };
}

function adminResetPassword(data, requester) {
  if (requester.role !== ROLES.ADMIN) {
    return { status: "ERROR", message: "Chỉ ADMIN mới được reset mật khẩu cho tài khoản khác." };
  }
  data = data || {};
  const userId = String(data.userId || "").trim();
  const newPassword = String(data.newPassword || "");
  if (!userId || !newPassword) return { status: "ERROR", message: "Thiếu tài khoản hoặc mật khẩu mới." };
  if (String(userId) === String(requester.userId)) {
    return { status: "ERROR", message: "Hãy dùng chức năng Đổi mật khẩu cho chính tài khoản ADMIN." };
  }
  const valid = validateNewPassword(newPassword);
  if (!valid.ok) return { status: "ERROR", message: valid.message };

  const found = findUserById(userId);
  if (!found) return { status: "ERROR", message: "Không tìm thấy tài khoản." };
  found.sheet.getRange(found.rowIndex, 3).setValue(hashPassword(newPassword));
  invalidateUserSessions(userId);
  writeAuditLog(requester.userId, "ADMIN_RESET_PASSWORD", "Users", userId, null, "Password reset by ADMIN");
  return { status: "SUCCESS", message: "Đã reset mật khẩu và đăng xuất các phiên đang hoạt động của tài khoản." };
}

function setAccountStatus(data, requester) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN].includes(requester.role)) {
    return { status: "ERROR", message: "Chỉ ADMIN/HR ADMIN được khóa hoặc mở tài khoản." };
  }
  data = data || {};
  const userId = String(data.userId || "").trim();
  const newStatus = String(data.status || "").trim().toUpperCase();
  if (!userId || !["ACTIVE", "INACTIVE"].includes(newStatus)) {
    return { status: "ERROR", message: "Trạng thái tài khoản không hợp lệ." };
  }
  if (String(userId) === String(requester.userId) && newStatus === "INACTIVE") {
    return { status: "ERROR", message: "Không thể tự khóa tài khoản đang đăng nhập." };
  }

  const found = findUserById(userId);
  if (!found) return { status: "ERROR", message: "Không tìm thấy tài khoản." };
  const oldStatus = String(found.row[5] || "").toUpperCase();
  found.sheet.getRange(found.rowIndex, 6).setValue(newStatus);
  if (newStatus === "INACTIVE") invalidateUserSessions(userId);
  writeAuditLog(requester.userId, newStatus === "ACTIVE" ? "UNLOCK_ACCOUNT" : "LOCK_ACCOUNT", "Users", userId, oldStatus, newStatus);
  return { status: "SUCCESS", message: newStatus === "ACTIVE" ? "Đã mở lại tài khoản." : "Đã khóa tài khoản." };
}

// ==========================================
// I. SETUP HỆ THỐNG & HELPER
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
    [TABLES.AUDIT]: ['AuditID', 'UserID', 'Action', 'Module', 'TargetID', 'OldValue', 'NewValue', 'Timestamp', 'IPAddress'],
    [TABLES.CONFIG]: ['Key', 'Value', 'UpdatedAt']
  };

  Object.keys(schema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = schema[sheetName];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#f3f3f3");
  });

  const userSheet = ss.getSheetByName(TABLES.USERS);
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(["USR-ADMIN", "admin@company.com", hashPassword("admin123"), ROLES.ADMIN, "EMP-ADMIN", "ACTIVE", new Date()]);
  }
}

/**
 * ============================================================
 * GOOGLE DRIVE - EMPLOYEE FOLDER
 * ============================================================
 *
 * Cấu trúc:
 *
 * Foreign Employee Management System
 * ├── EMP001 - John Smith
 * │   ├── Passport
 * │   ├── Visa
 * │   ├── TRC
 * │   ├── Work Permit
 * │   └── Contracts
 *
 * DRIVE_FOLDER_ID:
 * - Nếu đã cấu hình: dùng đúng folder đó.
 * - Nếu chưa có: tự tìm folder theo tên.
 * - Nếu không tìm thấy: tự tạo folder mới.
 *
 * Hàm này KHÔNG tạo folder nhân viên trùng.
 */
function createEmployeeDriveFolder(employeeId, fullName) {
  try {
    if (!employeeId) {
      throw new Error("Missing employeeId");
    }

    // --------------------------------------------------------
    // 1. Lấy folder gốc
    // --------------------------------------------------------
    const props = PropertiesService.getScriptProperties();
    const configuredRootId = props.getProperty("DRIVE_FOLDER_ID");

    let rootFolder = null;

    if (configuredRootId) {
      try {
        rootFolder = DriveApp.getFolderById(configuredRootId);
      } catch (e) {
        console.warn(
          "DRIVE_FOLDER_ID không hợp lệ, chuyển sang tìm theo tên."
        );
      }
    }

    // Nếu chưa có ID hoặc ID không hợp lệ
    if (!rootFolder) {
      const ROOT_NAME = "Foreign Employee Management System";
      const folders = DriveApp.getFoldersByName(ROOT_NAME);

      if (folders.hasNext()) {
        rootFolder = folders.next();

        // Lưu lại ID để những lần sau dùng trực tiếp
        props.setProperty(
          "DRIVE_FOLDER_ID",
          rootFolder.getId()
        );
      } else {
        rootFolder = DriveApp.createFolder(ROOT_NAME);

        props.setProperty(
          "DRIVE_FOLDER_ID",
          rootFolder.getId()
        );
      }
    }

    // --------------------------------------------------------
    // 2. Chuẩn hóa tên nhân viên
    // --------------------------------------------------------
    const safeEmployeeId = String(employeeId).trim();

    const safeFullName = String(fullName || "Unknown")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "_");

    const folderName =
      `${safeEmployeeId} - ${safeFullName}`;

    // --------------------------------------------------------
    // 3. Tìm folder nhân viên đã tồn tại
    // --------------------------------------------------------
    const existingFolders =
      rootFolder.getFoldersByName(folderName);

    let employeeFolder;

    if (existingFolders.hasNext()) {
      employeeFolder = existingFolders.next();
    } else {
      employeeFolder =
        rootFolder.createFolder(folderName);
    }

    // --------------------------------------------------------
    // 4. Đảm bảo 5 folder con luôn tồn tại
    // --------------------------------------------------------
    const subFolders = [
      "Passport",
      "Visa",
      "TRC",
      "Work Permit",
      "Contracts"
    ];

    subFolders.forEach(function (name) {
      const folders =
        employeeFolder.getFoldersByName(name);

      if (!folders.hasNext()) {
        employeeFolder.createFolder(name);
      }
    });

    // --------------------------------------------------------
    // 5. Trả về ID folder nhân viên
    // --------------------------------------------------------
    return employeeFolder.getId();

  } catch (e) {
    console.error(
      "createEmployeeDriveFolder error:",
      e
    );

    return "";
  }
}


/**
 * ============================================================
 * FORMAT DATE
 * ============================================================
 *
 * Chuẩn hóa ngày về:
 * YYYY-MM-DD
 *
 * Hỗ trợ:
 * - Date object
 * - yyyy-mm-dd
 * - yyyy/mm/dd
 * - yyyy.mm.dd
 * - dd/mm/yyyy
 * - dd-mm-yyyy
 * - dd.mm.yyyy
 */
function formatDate(d) {
  if (d === null || d === undefined || d === "") {
    return "";
  }

  // ----------------------------------------------------------
  // 1. Nếu đã là Date object
  // ----------------------------------------------------------
  if (
    Object.prototype.toString.call(d) === "[object Date]"
  ) {
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(
        d,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      );
    }

    return "";
  }

  // ----------------------------------------------------------
  // 2. Chuyển sang string
  // ----------------------------------------------------------
  const s = String(d).trim();

  if (!s) {
    return "";
  }

  // ----------------------------------------------------------
  // 3. yyyy-mm-dd / yyyy/mm/dd / yyyy.mm.dd
  // ----------------------------------------------------------
  let match = s.match(
    /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/
  );

  if (match) {
    return [
      match[1],
      String(match[2]).padStart(2, "0"),
      String(match[3]).padStart(2, "0")
    ].join("-");
  }

  // ----------------------------------------------------------
  // 4. dd-mm-yyyy / dd/mm/yyyy / dd.mm.yyyy
  // ----------------------------------------------------------
  match = s.match(
    /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/
  );

  if (match) {
    return [
      match[3],
      String(match[2]).padStart(2, "0"),
      String(match[1]).padStart(2, "0")
    ].join("-");
  }

  // ----------------------------------------------------------
  // 5. Nếu không nhận dạng được
  // ----------------------------------------------------------
  return s;
}


/**
 * ============================================================
 * AUDIT LOG
 * ============================================================
 *
 * Ghi:
 * LOG ID
 * User ID
 * Action
 * Module
 * Target ID
 * Old Value
 * New Value
 * Timestamp
 * Note
 */
function writeAuditLog(
  userId,
  action,
  module,
  targetId,
  oldValue,
  newValue
) {
  try {
    const sheet = getSheet(TABLES.AUDIT);

    if (!sheet) {
      throw new Error("AuditLog sheet not found");
    }

    const logId =
      "LOG-" +
      Utilities.getUuid()
        .replace(/-/g, "")
        .substring(0, 8)
        .toUpperCase();

    sheet.appendRow([
      logId,
      userId || "SYSTEM",
      action || "",
      module || "",
      targetId || "",
      oldValue || "",
      newValue || "",
      new Date(),
      ""
    ]);

  } catch (e) {

    // Không làm hỏng chức năng chính chỉ vì AuditLog lỗi
    console.error(
      "writeAuditLog error:",
      e
    );
  }
}

function handleGetTimeline(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem timeline." };

  const ee = handleGetEntryExit(empId, req).records || [];
  const tr = handleGetDomesticTravel(empId, req).records || [];
  const timeline = [];
  
  ee.forEach(x => timeline.push({ title: x.type === "ENTRY" ? "Nhập cảnh VN" : "Xuất cảnh VN", date: x.dateTime, location: x.airport }));
  tr.forEach(x => timeline.push({ title: "Công tác nội địa", date: x.fromDate, endDate: x.toDate, location: x.fromLocation + " → " + x.toLocation }));
  
  timeline.sort((a,b) => new Date(b.date) - new Date(a.date));
  return { status: "SUCCESS", timeline };
}