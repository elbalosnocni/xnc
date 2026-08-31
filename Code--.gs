//Code.gs
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
    case "ADD_ENTRY_EXIT_TRIP":
      result = handleAddEntryExitTrip(data, requester);
      break;
    case "UPDATE_ENTRY_EXIT_TRIP":
      result = handleUpdateEntryExitTrip(data, requester);
      break;
    case "IMPORT_MASTER_WORKBOOK":
      result = handleImportMasterWorkbook(data, requester);
      break;
    case "UPDATE_ENTRY_EXIT":
      result = handleUpdateEntryExit(data, requester);
      break;
    case "DELETE_ENTRY_EXIT":
      result = handleDeleteEntryExit(data, requester);
      break;
    case "GET_ENTRY_EXIT":
      result = handleGetEntryExit(employeeId, requester);
      break;
    case "ADD_DOMESTIC_TRAVEL":
      result = handleAddDomesticTravel(data, requester);
      break;
    case "UPDATE_DOMESTIC_TRAVEL":
      result = handleUpdateDomesticTravel(data, requester);
      break;
    case "CANCEL_DOMESTIC_TRAVEL":
      result = handleCancelDomesticTravel(data, requester);
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
    case "BATCH_ADD_ENTRY_EXIT":
      result = handleBatchAddEntryExit(data, requester);
      break;
    case "BATCH_ADD_DOMESTIC_TRAVEL":
      result = handleBatchAddDomesticTravel(data, requester);
      break;
    case "BATCH_IMPORT_MOVEMENT":
      result = handleBatchImportMovement(data, requester);
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
    if (String(rows[i][0]).trim() === data.employeeId) {
      rowIndex = i + 1;
      break;
    }
  }

  const timestamp = new Date();

  if (rowIndex > 0) {
    const currentRow = rows[rowIndex - 1].slice();
    const updatedRow = currentRow.slice(0, 14);

    if (isHRAdmin) {
      updatedRow[1] = data.fullName || currentRow[1] || "";
      updatedRow[2] = data.nationality !== undefined ? String(data.nationality).trim() : (currentRow[2] || "");
      updatedRow[3] = data.dob !== undefined ? (formatDate(data.dob) || "") : (currentRow[3] || "");
      if (data.position !== undefined) updatedRow[4] = String(data.position || "").trim();
      if (data.department !== undefined) updatedRow[5] = String(data.department || "").trim();
    }
    if (data.phone !== undefined) updatedRow[6] = String(data.phone || "").trim();
    if (data.email !== undefined) {
      const newEmail = String(data.email || "").trim();
      updatedRow[7] = newEmail;
      
      // Đồng bộ email sang bảng Users nếu thay đổi email
      if (newEmail && newEmail.toLowerCase() !== String(currentRow[7]).trim().toLowerCase()) {
        const userSheet = getSheet(TABLES.USERS);
        const uRows = userSheet.getDataRange().getValues();
        for (let u = 1; u < uRows.length; u++) {
          if (String(uRows[u][4]).trim() === data.employeeId) {
            userSheet.getRange(u + 1, 2).setValue(newEmail);
            break;
          }
        }
      }
    }
    
    if (isHRAdmin && data.activeStatus !== undefined) updatedRow[10] = String(data.activeStatus || "").trim().toUpperCase();
    updatedRow[13] = timestamp;

    sheet.getRange(rowIndex, 1, 1, 14).setValues([updatedRow]);
    SpreadsheetApp.flush();
    writeAuditLog(requester.userId, "UPDATE", "Employees", data.employeeId, JSON.stringify(currentRow), JSON.stringify(updatedRow));
  } else {
    if (!isHRAdmin) return { status: "ERROR", message: "Chỉ HR mới được thêm mới nhân viên!" };

    if (rows.slice(1).some(r => String(r[0]).trim() === data.employeeId)) {
      return { status: "ERROR", message: "Mã nhân viên đã tồn tại." };
    }

    const normalizedEmail = String(data.email || "").trim().toLowerCase();
    if (normalizedEmail) {
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
    userSheet.appendRow([
      "USR-" + Utilities.getUuid().substring(0, 6),
      data.email || data.employeeId,
      hashPassword("123456"),
      ROLES.USER,
      data.employeeId,
      "ACTIVE",
      timestamp
    ]);
    writeAuditLog(requester.userId, "CREATE", "Employees", data.employeeId, null, JSON.stringify(newRow));
  }

  // Đồng bộ trạng thái vị trí làm việc
  calculateEmployeeLocationStatus(data.employeeId);

  const childResults = { documents: [], contract: null };
  if (isHRAdmin) {
    if (data.passportNo) childResults.documents.push(saveDocument({
      employeeId: data.employeeId,
      docType: "PASSPORT",
      docNo: data.passportNo,
      expiryDate: data.passportExpiry,
      fileId: data.passportFileId || "",
      fileUrl: data.passportImg || ""
    }, requester));

    if (data.trcNo) childResults.documents.push(saveDocument({
      employeeId: data.employeeId,
      docType: "TRC",
      docNo: data.trcNo,
      expiryDate: data.trcExpiry,
      fileId: data.trcFileId || "",
      fileUrl: data.trcImg || ""
    }, requester));

    if (data.wpNo) childResults.documents.push(saveDocument({
      employeeId: data.employeeId,
      docType: "WORK_PERMIT",
      docNo: data.wpNo,
      expiryDate: data.wpExpiry,
      fileId: data.wpFileId || "",
      fileUrl: data.wpImg || ""
    }, requester));

    if (data.contractNo) {
      childResults.contract = saveContractAndAppendix(data, requester);
    }
  }

  SpreadsheetApp.flush();

  const verify = getProfileData(data.employeeId, requester);
  if (verify.status !== "SUCCESS") {
    return { status: "ERROR", message: "Đã ghi nhưng không đọc lại được hồ sơ: " + verify.message };
  }

  return {
    status: "SUCCESS",
    message: "Đã lưu thông tin hồ sơ thành công!",
    profile: verify.profile,
    children: childResults
  };
}

function isCurrentValue(value) {
  return value === true || String(value).trim().toUpperCase() === "TRUE" || String(value).trim() === "1";
}

function saveContractAndAppendix(data, requester) {
  const contractSheet = getSheet(TABLES.CONTRACTS);
  const contractRows = contractSheet.getDataRange().getValues();
  let currentContractId = "";
  let currentRowIndex = -1;

  for (let i = 1; i < contractRows.length; i++) {
    if (String(contractRows[i][1]).trim() === String(data.employeeId).trim() && isCurrentValue(contractRows[i][9])) {
      if (String(contractRows[i][2]).trim() === String(data.contractNo).trim()) {
        currentContractId = String(contractRows[i][0]);
        currentRowIndex = i + 1;
      } else {
        contractSheet.getRange(i + 1, 10).setValue(false);
      }
    }
  }

  const contractType = String(data.contractType || "DEFINITE").trim() || "DEFINITE";
  const contractStatus = String(data.contractStatus || "VALID").trim() || "VALID";
  const allowancesJson = data.allowancesJson
    ? String(data.allowancesJson)
    : JSON.stringify({ allowance: Number(data.allowance) || 0 });

  if (!currentContractId) {
    currentContractId = "CTR-" + Utilities.getUuid().substring(0, 8);
    const newContractRow = [
      currentContractId,
      data.employeeId,
      data.contractNo,
      contractType,
      formatDate(data.contractStartDate),
      formatDate(data.contractExpiry),
      Number(data.salary) || 0,
      allowancesJson,
      contractStatus,
      true,
      data.contractFileId || "",
      data.contractImg || "",
      new Date(),
      requester.userId
    ];
    contractSheet.appendRow(newContractRow);
    writeAuditLog(requester.userId, "ADD_CONTRACT", "Contracts", currentContractId, null, JSON.stringify(newContractRow));
  } else {
    const oldRow = contractSheet.getRange(currentRowIndex, 1, 1, 14).getValues()[0];
    const updated = oldRow.slice();
    updated[2] = data.contractNo || oldRow[2];
    updated[3] = contractType;
    updated[4] = formatDate(data.contractStartDate) || oldRow[4] || "";
    updated[5] = formatDate(data.contractExpiry) || oldRow[5] || "";
    updated[6] = Number(data.salary) || 0;
    updated[7] = allowancesJson;
    updated[8] = contractStatus;
    updated[9] = true;
    if (data.contractFileId !== undefined && data.contractFileId !== "") updated[10] = data.contractFileId;
    if (data.contractImg !== undefined && data.contractImg !== "") updated[11] = data.contractImg;
    updated[12] = oldRow[12] || new Date();
    updated[13] = requester.userId;
    contractSheet.getRange(currentRowIndex, 1, 1, 14).setValues([updated]);
    writeAuditLog(requester.userId, "UPDATE_CONTRACT", "Contracts", currentContractId, JSON.stringify(oldRow), JSON.stringify(updated));
  }

  let appendixId = "";
  if (data.appendixNo) {
    const appSheet = getSheet(TABLES.APPENDICES);
    const appRows = appSheet.getDataRange().getValues();
    let appRowIndex = -1;

    for (let i = 1; i < appRows.length; i++) {
      if (String(appRows[i][1]).trim() === String(currentContractId).trim() &&
          String(appRows[i][3]).trim() === String(data.appendixNo).trim()) {
        appRowIndex = i + 1;
        appendixId = String(appRows[i][0]);
        break;
      }
    }

    const appendixRow = [
      appendixId || ("APP-" + Utilities.getUuid().substring(0, 8)),
      currentContractId,
      data.employeeId,
      data.appendixNo,
      formatDate(data.appendixEffectiveDate || data.contractStartDate),
      formatDate(data.appendixEndDate || data.contractExpiry),
      Number(data.appendixSalary !== undefined ? data.appendixSalary : data.salary) || 0,
      data.appendixAllowancesJson ? String(data.appendixAllowancesJson) : JSON.stringify({ allowance: Number(data.appendixAllowance !== undefined ? data.appendixAllowance : data.allowance) || 0 }),
      data.appendixContent || "Cập nhật theo hồ sơ",
      data.appendixFileId || "",
      data.appendixFileUrl || data.contractImg || "",
      appRowIndex > 0 ? (appRows[appRowIndex - 1][11] || new Date()) : new Date(),
      requester.userId
    ];

    if (appRowIndex > 0) {
      const oldApp = appSheet.getRange(appRowIndex, 1, 1, 13).getValues()[0];
      appSheet.getRange(appRowIndex, 1, 1, 13).setValues([appendixRow]);
      writeAuditLog(requester.userId, "UPDATE_APPENDIX", "ContractAppendices", appendixRow[0], JSON.stringify(oldApp), JSON.stringify(appendixRow));
    } else {
      appendixId = appendixRow[0];
      appSheet.appendRow(appendixRow);
      writeAuditLog(requester.userId, "ADD_APPENDIX", "ContractAppendices", appendixId, null, JSON.stringify(appendixRow));
    }
  }

  SpreadsheetApp.flush();
  return { status: "SUCCESS", contractId: currentContractId, appendixId: appendixId };
}

function saveDocument(docData, requester) {
  const sheet = getSheet(TABLES.DOCUMENTS);
  const rows = sheet.getDataRange().getValues();
  let currentRowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === String(docData.employeeId).trim() && rows[i][2] === docData.docType) {
      if (currentRowIndex < 0 && isCurrentValue(rows[i][9])) currentRowIndex = i + 1;
      else if (isCurrentValue(rows[i][9])) sheet.getRange(i + 1, 10).setValue(false);
    }
  }

  if (currentRowIndex > 0) {
    const oldRow = sheet.getRange(currentRowIndex, 1, 1, 13).getValues()[0];
    const updated = oldRow.slice();
    updated[2] = docData.docType;
    updated[3] = docData.docNo !== undefined ? String(docData.docNo || "").trim() : oldRow[3];
    updated[4] = docData.issueDate !== undefined ? (formatDate(docData.issueDate) || "") : oldRow[4];
    updated[5] = docData.expiryDate !== undefined ? (formatDate(docData.expiryDate) || "") : oldRow[5];
    updated[6] = docData.issuer !== undefined ? String(docData.issuer || "").trim() : oldRow[6];
    if (docData.fileId !== undefined && docData.fileId !== "") updated[7] = docData.fileId;
    if (docData.fileUrl !== undefined && docData.fileUrl !== "") updated[8] = docData.fileUrl;
    updated[9] = true;
    updated[10] = String(docData.status || oldRow[10] || "VALID").toUpperCase();
    updated[11] = oldRow[11] || new Date();
    updated[12] = requester.userId;
    sheet.getRange(currentRowIndex, 1, 1, 13).setValues([updated]);
    SpreadsheetApp.flush();
    writeAuditLog(requester.userId, "UPDATE_DOCUMENT", "Documents", updated[0], JSON.stringify(oldRow), JSON.stringify(updated));
    return { status: "SUCCESS", message: "Cập nhật giấy tờ thành công!", docId: updated[0] };
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
    String(docData.status || "VALID").toUpperCase(),
    new Date(),
    requester.userId
  ];

  sheet.appendRow(newDocRow);
  SpreadsheetApp.flush();
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

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const fileUrl = `https://drive.google.com/file/d/${fileId}/view`;
    const previewUrl = `https://drive.google.com/file/d/${fileId}/preview`;

    return {
      status: "SUCCESS",
      fileId: fileId,
      fileName: safeName,
      mimeType: safeMime,
      fileUrl: fileUrl,
      previewUrl: previewUrl
    };

  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

// ==========================================
// F. LỊCH TRÌNH & XUẤT NHẬP CẢNH & DỰ ĐỊNH
// ==========================================
function parseDateSafe(val) {
  if (!val) return new Date(0);
  if (Object.prototype.toString.call(val) === "[object Date]") return val;
  const parsed = new Date(val);
  return isNaN(parsed.getTime()) ? new Date(0) : parsed;
}

function calculateEmployeeLocationStatus(employeeId) {
  const now = new Date();
  const eeSheet = getSheet(TABLES.ENTRY_EXIT);
  const travelSheet = getSheet(TABLES.TRAVEL);

  const eeRows = eeSheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]).trim() === String(employeeId).trim())
    .map(r => ({
      type: String(r[2] || "").trim().toUpperCase(),
      date: parseDateSafe(r[3])
    }))
    .filter(x => x.date.getTime() > 0 && x.date <= now)
    .sort((a, b) => b.date - a.date);

  // Trạng thái hiện tại chỉ dựa trên sự kiện nhập/xuất gần nhất ĐÃ xảy ra.
  // Không dùng lịch Entry/Exit trong tương lai để làm thay đổi trạng thái hiện tại.
  let status = "Exited";
  let location = "Overseas";

  if (eeRows.length) {
    status = eeRows[0].type === "ENTRY" ? "In Vietnam" : "Exited";
    location = status === "In Vietnam" ? "Vietnam" : "Overseas";
  }

  // DomesticTravel chỉ được override khi đang ở Việt Nam và chuyến đã APPROVED
  // với khoảng thời gian bao phủ thời điểm hiện tại.
  if (status === "In Vietnam") {
    const travelRows = travelSheet.getDataRange().getValues().slice(1)
      .filter(r => String(r[1]).trim() === String(employeeId).trim())
      .map(r => ({
        from: parseDateSafe(r[4]),
        to: (() => {
          const d = parseDateSafe(r[5]);
          if (d.getTime() > 0) d.setHours(23, 59, 59, 999);
          return d;
        })(),
        status: String(r[6] || "").trim().toUpperCase(),
        toLocation: String(r[3] || "").trim(),
        fromLocation: String(r[2] || "").trim()
      }))
      .filter(x => x.status === "APPROVED" && x.from.getTime() > 0 && x.to.getTime() > 0 && now >= x.from && now <= x.to)
      .sort((a, b) => b.from - a.from);

    if (travelRows.length) {
      status = "Traveling";
      location = travelRows[0].toLocation || travelRows[0].fromLocation || "Vietnam";
    }
  }

  updateEmployeeStatusRecord(employeeId, status, location);
  return { status, location };
}

function updateEmployeeStatusRecord(employeeId, status, location) {
  const sheet = getSheet(TABLES.EMPLOYEES);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(employeeId).trim()) {
      if (String(rows[i][8] || "") !== String(status || "")) {
        sheet.getRange(i + 1, 9).setValue(status);
      }
      if (String(rows[i][9] || "") !== String(location || "")) {
        sheet.getRange(i + 1, 10).setValue(location);
      }
      break;
    }
  }
}

// --- 1. HÀM THÊM KHAI BÁO NHẬP XUẤT CẢNH (Đã fix nhận dateTime) ---
function handleAddEntryExit(data, req) {
  data = data || {};
  const employeeId = String(data.employeeId || "").trim();
  const type = String(data.type || "").trim().toUpperCase();
  const plannedDateTime = data.plannedDateTime || data.dateTime || data.date;
  const actualDateTime = data.actualDateTime || "";
  const eventDate = actualDateTime || plannedDateTime;

  if (!employeeId || !["ENTRY", "EXIT"].includes(type) || !eventDate || !data.airport) {
    return { status: "ERROR", message: "Thiếu hoặc sai thông tin nhập/xuất cảnh." };
  }
  if (!canAccessEmployee(req, employeeId)) {
    return { status: "ERROR", message: "Bạn không có quyền thực hiện." };
  }

  const parsedEvent = parseDateSafe(eventDate);
  if (parsedEvent.getTime() === 0) return { status: "ERROR", message: "Ngày giờ không hợp lệ." };

  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const id = "EE-" + Utilities.getUuid().substring(0, 8);
  const now = new Date();

  // Schema mới giữ nguyên 9 cột đầu để tương thích dữ liệu cũ,
  // đồng thời bổ sung CreatedBy / PlannedDateTime / ActualDateTime / TripID.
  const row = [
    id,
    employeeId,
    type,
    eventDate,
    String(data.airport || "").trim(),
    String(data.flightNo || "").trim(),
    String(data.destination || "").trim(),
    String(data.purpose || "").trim(),
    now,
    req.userId || "",
    plannedDateTime || "",
    actualDateTime || "",
    String(data.tripId || "").trim()
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);

  calculateEmployeeLocationStatus(employeeId);
  writeAuditLog(req.userId, "ADD_ENTRY_EXIT", TABLES.ENTRY_EXIT, id, null,
    JSON.stringify({ employeeId, type, eventDate, plannedDateTime, actualDateTime, airport: data.airport, flightNo: data.flightNo, destination: data.destination, purpose: data.purpose }));

  return { status: "SUCCESS", message: "Đã lưu khai báo nhập/xuất cảnh thành công!", entryExitId: id };
}

function handleUpdateEntryExit(data, req) {
  data = data || {};
  const id = String(data.entryExitId || data.logId || "").trim();
  const employeeId = String(data.employeeId || "").trim();
  if (!id || !employeeId) return { status: "ERROR", message: "Thiếu mã khai báo." };
  if (!canAccessEmployee(req, employeeId)) return { status: "ERROR", message: "Bạn không có quyền thực hiện." };

  const type = String(data.type || "").trim().toUpperCase();
  const plannedDateTime = data.plannedDateTime || data.dateTime || data.date;
  const actualDateTime = data.actualDateTime || "";
  const eventDate = actualDateTime || plannedDateTime;
  if (!["ENTRY", "EXIT"].includes(type) || !eventDate || !data.airport) {
    return { status: "ERROR", message: "Thiếu hoặc sai thông tin nhập/xuất cảnh." };
  }

  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== id) continue;
    if (String(rows[i][1]).trim() !== employeeId) return { status: "ERROR", message: "Khai báo không thuộc nhân viên này." };

    const oldRow = rows[i].slice();
    const updated = oldRow.slice();
    while (updated.length < 13) updated.push("");
    updated[2] = type;
    updated[3] = eventDate;
    updated[4] = String(data.airport || "").trim();
    updated[5] = String(data.flightNo || "").trim();
    updated[6] = String(data.destination || "").trim();
    updated[7] = String(data.purpose || "").trim();
    updated[8] = oldRow[8] || new Date();
    updated[9] = oldRow[9] || req.userId || "";
    updated[10] = plannedDateTime || "";
    updated[11] = actualDateTime || "";
    updated[12] = String(data.tripId || oldRow[12] || "").trim();

    sheet.getRange(i + 1, 1, 1, 13).setValues([updated.slice(0, 13)]);
    calculateEmployeeLocationStatus(employeeId);
    writeAuditLog(req.userId, "UPDATE_ENTRY_EXIT", TABLES.ENTRY_EXIT, id, JSON.stringify(oldRow), JSON.stringify(updated));
    return { status: "SUCCESS", message: "Đã cập nhật khai báo nhập/xuất cảnh." };
  }
  return { status: "ERROR", message: "Không tìm thấy khai báo nhập/xuất cảnh." };
}

function handleDeleteEntryExit(data, req) {
  data = data || {};
  const id = String(data.entryExitId || data.logId || "").trim();
  if (!id) return { status: "ERROR", message: "Thiếu mã khai báo." };

  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== id) continue;
    const employeeId = String(rows[i][1]).trim();
    if (!canAccessEmployee(req, employeeId)) return { status: "ERROR", message: "Bạn không có quyền thực hiện." };

    const oldRow = rows[i].slice();
    sheet.deleteRow(i + 1);
    calculateEmployeeLocationStatus(employeeId);
    writeAuditLog(req.userId, "DELETE_ENTRY_EXIT", TABLES.ENTRY_EXIT, id, JSON.stringify(oldRow), "DELETED");
    return { status: "SUCCESS", message: "Đã xóa khai báo nhập/xuất cảnh." };
  }
  return { status: "ERROR", message: "Không tìm thấy khai báo nhập/xuất cảnh." };
}

function handleGetEntryExit(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem thông tin." };

  const sheet = getSheet(TABLES.ENTRY_EXIT);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => {
    let flightNo = String(r[5] || "");
    let destination = String(r[6] || "");
    let purpose = String(r[7] || "");

    try {
      const noteObj = JSON.parse(String(r[5] || ""));
      if (noteObj && typeof noteObj === "object") {
        flightNo = noteObj.flightNo || "";
        destination = noteObj.destination || "";
        purpose = noteObj.purpose || "";
      }
    } catch (e) {}

    return {
      entryExitId: r[0],
      employeeId: r[1],
      type: r[2],
      dateTime: formatDateTime(r[3]),
      plannedDateTime: formatDateTime(r[10] || r[3]),
      actualDateTime: r[11] ? formatDateTime(r[11]) : "",
      airport: r[4],
      flightNo,
      destination,
      purpose,
      tripId: r[12] || ""
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

  const start = parseDateSafe(d.fromDate);
  const end = parseDateSafe(d.toDate);
  if (start.getTime() === 0 || end.getTime() === 0 || end < start) {
    return { status: "ERROR", message: "Khoảng ngày đi lại không hợp lệ." };
  }
  if (!canAccessEmployee(req, d.employeeId)) return { status: "ERROR", message: "Không có quyền thực hiện." };

  const sheet = getSheet(TABLES.TRAVEL);
  const id = "TRV-" + Utilities.getUuid().substring(0, 8);
  const isHR = [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(req.role);
  const initialStatus = isHR ? "APPROVED" : "PENDING";

  const row = [
    id, d.employeeId, d.fromLocation, d.toLocation,
    d.fromDate, d.toDate, initialStatus, "", "", new Date(), req.userId,
    String(d.purpose || "").trim()
  ];
  sheet.appendRow(row);
  calculateEmployeeLocationStatus(d.employeeId);
  writeAuditLog(req.userId, "ADD_DOMESTIC_TRAVEL", TABLES.TRAVEL, id, null, JSON.stringify(row));
  return {
    status: "SUCCESS",
    message: isHR ? "Đã lưu và duyệt lịch trình công tác!" : "Đã gửi yêu cầu công tác (Chờ duyệt)!",
    travelId: id
  };
}

function handleUpdateDomesticTravel(d, req) {
  d = d || {};
  const travelId = String(d.travelId || "").trim();
  if (!travelId) return { status: "ERROR", message: "Thiếu mã lịch đi lại." };

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== travelId) continue;

    const employeeId = String(rows[i][1]).trim();
    if (!canAccessEmployee(req, employeeId)) return { status: "ERROR", message: "Không có quyền thực hiện." };

    const fromDate = d.fromDate;
    const toDate = d.toDate;
    const startDate = parseDateSafe(fromDate);
    const endDate = parseDateSafe(toDate);
    if (!fromDate || !toDate || startDate.getTime() === 0 || endDate.getTime() === 0 || endDate < startDate ||
        !d.fromLocation || !d.toLocation) {
      return { status: "ERROR", message: "Thông tin lịch đi lại không hợp lệ." };
    }

    const oldRow = rows[i].slice();
    const updated = oldRow.slice();
    updated[2] = String(d.fromLocation).trim();
    updated[3] = String(d.toLocation).trim();
    updated[4] = fromDate;
    updated[5] = toDate;
    updated[11] = String(d.purpose || "").trim();
    // Không tự ý đổi PENDING/APPROVED khi chỉ sửa nội dung.
    sheet.getRange(i + 1, 1, 1, 11).setValues([updated.slice(0, 11)]);
    calculateEmployeeLocationStatus(employeeId);
    writeAuditLog(req.userId, "UPDATE_DOMESTIC_TRAVEL", TABLES.TRAVEL, travelId, JSON.stringify(oldRow), JSON.stringify(updated));
    return { status: "SUCCESS", message: "Đã cập nhật lịch đi lại." };
  }
  return { status: "ERROR", message: "Không tìm thấy lịch đi lại." };
}

function handleCancelDomesticTravel(d, req) {
  d = d || {};
  const travelId = String(d.travelId || "").trim();
  if (!travelId) return { status: "ERROR", message: "Thiếu mã lịch đi lại." };

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== travelId) continue;

    const employeeId = String(rows[i][1]).trim();
    if (!canAccessEmployee(req, employeeId)) return { status: "ERROR", message: "Không có quyền thực hiện." };

    const oldStatus = String(rows[i][6] || "");
    if (oldStatus.toUpperCase() === "CANCELLED") return { status: "SUCCESS", message: "Lịch đi lại đã được hủy trước đó." };

    sheet.getRange(i + 1, 7).setValue("CANCELLED");
    calculateEmployeeLocationStatus(employeeId);
    writeAuditLog(req.userId, "CANCEL_DOMESTIC_TRAVEL", TABLES.TRAVEL, travelId, oldStatus, "CANCELLED");
    return { status: "SUCCESS", message: "Đã hủy lịch đi lại." };
  }
  return { status: "ERROR", message: "Không tìm thấy lịch đi lại." };
}

function handleGetDomesticTravel(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem thông tin." };

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues().slice(1);
  const records = rows.filter(r => String(r[1]) === String(empId)).map(r => ({
    travelId: r[0],
    employeeId: r[1],
    fromLocation: r[2],
    toLocation: r[3],
    fromDate: formatDate(r[4]),
    toDate: formatDate(r[5]),
    status: String(r[6] || "").toUpperCase(),
    purpose: r[11] || ""
  }));
  return { status: "SUCCESS", records };
}


// ==========================================
// BATCH ADMIN / UNIFIED MOVEMENT IMPORT
// ==========================================
function requireBatchAdmin(req) {
  if (!req || ![ROLES.ADMIN, ROLES.HR_ADMIN].includes(String(req.role || "").toUpperCase())) {
    return { status: "ERROR", message: "Chỉ ADMIN / HR ADMIN được phép thực hiện thao tác hàng loạt." };
  }
  return null;
}

function handleBatchAddEntryExit(data, req) {
  const denied = requireBatchAdmin(req);
  if (denied) return denied;

  data = data || {};
  const employeeIds = Array.isArray(data.employeeIds) ? data.employeeIds.map(String).map(x => x.trim()).filter(Boolean) : [];
  const mode = String(data.mode || "ROUND_TRIP").toUpperCase();
  if (!employeeIds.length) return { status: "ERROR", message: "Chưa chọn nhân viên." };

  const results = [];
  const errors = [];

  employeeIds.forEach(employeeId => {
    try {
      if (["ROUND_TRIP", "ENTRY_ONLY"].includes(mode)) {
        if (!data.entryPlannedDateTime || !data.entryAirport) {
          throw new Error("Thiếu thông tin Entry.");
        }
        const r = handleAddEntryExit({
          employeeId,
          type: "ENTRY",
          plannedDateTime: data.entryPlannedDateTime,
          actualDateTime: data.entryActualDateTime || "",
          airport: data.entryAirport,
          flightNo: data.entryFlightNo || "",
          destination: data.entryDestination || "",
          purpose: data.purpose || "",
          tripId: data.tripId || ""
        }, req);
        if (r.status !== "SUCCESS") throw new Error(r.message || "Không lưu được Entry.");
        results.push({ employeeId, type: "ENTRY", id: r.entryExitId });
      }

      if (["ROUND_TRIP", "EXIT_ONLY"].includes(mode)) {
        if (!data.exitPlannedDateTime || !data.exitAirport) {
          throw new Error("Thiếu thông tin Exit.");
        }
        const r = handleAddEntryExit({
          employeeId,
          type: "EXIT",
          plannedDateTime: data.exitPlannedDateTime,
          actualDateTime: data.exitActualDateTime || "",
          airport: data.exitAirport,
          flightNo: data.exitFlightNo || "",
          destination: data.exitDestination || "",
          purpose: data.purpose || "",
          tripId: data.tripId || ""
        }, req);
        if (r.status !== "SUCCESS") throw new Error(r.message || "Không lưu được Exit.");
        results.push({ employeeId, type: "EXIT", id: r.entryExitId });
      }
    } catch (e) {
      errors.push({ employeeId, message: e.message || String(e) });
    }
  });

  return {
    status: errors.length && !results.length ? "ERROR" : "SUCCESS",
    message: `Đã xử lý ${results.length} bản ghi Entry/Exit${errors.length ? `, ${errors.length} lỗi` : ""}.`,
    created: results.length,
    errors,
    records: results
  };
}

function handleBatchAddDomesticTravel(data, req) {
  const denied = requireBatchAdmin(req);
  if (denied) return denied;

  data = data || {};
  const employeeIds = Array.isArray(data.employeeIds) ? data.employeeIds.map(String).map(x => x.trim()).filter(Boolean) : [];
  if (!employeeIds.length) return { status: "ERROR", message: "Chưa chọn nhân viên." };

  const results = [];
  const errors = [];
  employeeIds.forEach(employeeId => {
    try {
      const r = handleAddDomesticTravel({
        employeeId,
        fromDate: data.fromDate,
        toDate: data.toDate,
        fromLocation: data.fromLocation,
        toLocation: data.toLocation,
        purpose: data.purpose || ""
      }, req);
      if (r.status !== "SUCCESS") throw new Error(r.message || "Không lưu được lịch Travel.");
      results.push({ employeeId, id: r.travelId });
    } catch (e) {
      errors.push({ employeeId, message: e.message || String(e) });
    }
  });

  return {
    status: errors.length && !results.length ? "ERROR" : "SUCCESS",
    message: `Đã xử lý ${results.length} lịch Travel${errors.length ? `, ${errors.length} lỗi` : ""}.`,
    created: results.length,
    errors,
    records: results
  };
}

function handleBatchImportMovement(data, req) {
  const denied = requireBatchAdmin(req);
  if (denied) return denied;

  const rows = Array.isArray(data && data.rows) ? data.rows : [];
  if (!rows.length) return { status: "ERROR", message: "Không có dữ liệu Movement." };

  const errors = [];
  const created = [];
  let entryExitCreated = 0;
  let travelCreated = 0;

  rows.forEach((raw, idx) => {
    const row = raw || {};
    const rowNo = Number(row._excelRow || idx + 2);
    const movementType = String(row.movementType || row.module || row.type || "").trim().toUpperCase();
    const employeeId = String(row.employeeId || "").trim();

    try {
      if (!employeeId) throw new Error("Thiếu EmployeeID.");
      if (!["ENTRY_EXIT", "TRAVEL"].includes(movementType)) {
        throw new Error("MovementType phải là ENTRY_EXIT hoặc TRAVEL.");
      }

      if (movementType === "ENTRY_EXIT") {
        const entryPlanned = row.entryPlannedDateTime || "";
        const entryActual = row.entryActualDateTime || "";
        const exitPlanned = row.exitPlannedDateTime || "";
        const exitActual = row.exitActualDateTime || "";

        if (entryPlanned || entryActual) {
          const r = handleAddEntryExit({
            employeeId,
            type: "ENTRY",
            plannedDateTime: entryPlanned || entryActual,
            actualDateTime: entryActual,
            airport: row.entryAirport || row.airport,
            flightNo: row.entryFlightNo || row.flightNo,
            destination: row.entryDestination || row.destination,
            purpose: row.purpose,
            tripId: row.tripId
          }, req);
          if (r.status !== "SUCCESS") throw new Error(r.message || "Không tạo Entry.");
          entryExitCreated++;
          created.push({ row: rowNo, employeeId, type: "ENTRY", id: r.entryExitId });
        }

        if (exitPlanned || exitActual) {
          const r = handleAddEntryExit({
            employeeId,
            type: "EXIT",
            plannedDateTime: exitPlanned || exitActual,
            actualDateTime: exitActual,
            airport: row.exitAirport || row.airport,
            flightNo: row.exitFlightNo || row.flightNo,
            destination: row.exitDestination || row.destination,
            purpose: row.purpose,
            tripId: row.tripId
          }, req);
          if (r.status !== "SUCCESS") throw new Error(r.message || "Không tạo Exit.");
          entryExitCreated++;
          created.push({ row: rowNo, employeeId, type: "EXIT", id: r.entryExitId });
        }

        if (!(entryPlanned || entryActual || exitPlanned || exitActual)) {
          throw new Error("ENTRY_EXIT phải có ít nhất Entry hoặc Exit.");
        }
      } else {
        const r = handleAddDomesticTravel({
          employeeId,
          fromDate: row.fromDate,
          toDate: row.toDate,
          fromLocation: row.fromLocation,
          toLocation: row.toLocation,
          purpose: row.purpose || ""
        }, req);
        if (r.status !== "SUCCESS") throw new Error(r.message || "Không tạo Travel.");
        travelCreated++;
        created.push({ row: rowNo, employeeId, type: "TRAVEL", id: r.travelId });
      }
    } catch (e) {
      errors.push({ row: rowNo, employeeId, message: e.message || String(e) });
    }
  });

  return {
    status: errors.length && !created.length ? "ERROR" : "SUCCESS",
    message: `Import Movement hoàn tất: ${entryExitCreated} Entry/Exit, ${travelCreated} Travel${errors.length ? `, ${errors.length} dòng lỗi` : ""}.`,
    entryExitCreated,
    travelCreated,
    created: created.length,
    errors,
    records: created
  };
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
  const ctrSheet = getSheet(TABLES.CONTRACTS);
  const ctrRows = ctrSheet.getDataRange().getValues().slice(1);
  const eeSheet = getSheet(TABLES.ENTRY_EXIT);
  const eeRows = eeSheet.getDataRange().getValues().slice(1);
  const travelSheet = getSheet(TABLES.TRAVEL);
  const travelRows = travelSheet.getDataRange().getValues().slice(1);
  const now = new Date();

  // Luôn tính lại trạng thái hiện tại từ dữ liệu nguồn trước khi render Dashboard.
  // Như vậy Employees không phụ thuộc vào lần cuối người dùng sửa hồ sơ.
  empRows.forEach(r => {
    const empId = String(r[0] || "").trim();
    if (empId && String(r[10] || "").toUpperCase() === "ACTIVE") {
      calculateEmployeeLocationStatus(empId);
    }
  });

  // Đọc lại Employees sau khi recalculate.
  const freshEmpRows = empSheet.getDataRange().getValues().slice(1);

  const kpi = { total: 0, inVN: 0, business: 0, exited: 0, warning: 0, expContract: 0 };
  const employees = [];
  const expiryWarnings = [];
  const nextTravelMap = {};
  const nextEntryExitMap = {};

  // Entry/Exit: chỉ lấy sự kiện gần nhất trong tương lai, tính theo thời gian đầy đủ.
  eeRows.forEach(r => {
    const employeeId = String(r[1] || "").trim();
    const eventDate = parseDateSafe(r[10] || r[3]);
    if (!employeeId || eventDate.getTime() === 0 || eventDate < now) return;

    if (!nextEntryExitMap[employeeId] || eventDate < nextEntryExitMap[employeeId].rawDate) {
      let flightNo = String(r[5] || "");
      let destination = String(r[6] || "");
      let purpose = String(r[7] || "");
      try {
        const obj = JSON.parse(String(r[5] || ""));
        if (obj && typeof obj === "object") {
          flightNo = obj.flightNo || "";
          destination = obj.destination || "";
          purpose = obj.purpose || "";
        }
      } catch (e) {}

      nextEntryExitMap[employeeId] = {
        entryExitId: r[0],
        rawDate: eventDate,
        type: String(r[2] || "").toUpperCase(),
        dateTime: formatDateTime(r[3]),
        airport: r[4] || "",
        flightNo,
        destination,
        purpose
      };
    }
  });

  // DomesticTravel: chỉ lấy chuyến APPROVED có ngày bắt đầu từ hiện tại trở đi.
  travelRows.forEach(r => {
    const employeeId = String(r[1] || "").trim();
    const startDate = parseDateSafe(r[4]);
    const status = String(r[6] || "").toUpperCase();
    if (!employeeId || startDate.getTime() === 0 || status !== "APPROVED") return;
    const startDay = new Date(startDate);
    startDay.setHours(0, 0, 0, 0);
    const todayDay = new Date(now);
    todayDay.setHours(0, 0, 0, 0);
    if (startDay < todayDay) return;

    if (!nextTravelMap[employeeId] || startDate < nextTravelMap[employeeId].rawDate) {
      nextTravelMap[employeeId] = {
        travelId: r[0],
        rawDate: startDate,
        fromLocation: r[2] || "",
        toLocation: r[3] || "",
        fromDate: formatDate(r[4]),
        toDate: formatDate(r[5]),
        status
      };
    }
  });

  const activeContractsMap = {};
  ctrRows.forEach(c => {
    if (isCurrentValue(c[9])) {
      activeContractsMap[String(c[1]).trim()] = {
        contractNo: c[2],
        contractExpiry: formatDate(c[5])
      };
    }
  });

  freshEmpRows.forEach(r => {
    if (String(r[10] || "").toUpperCase() !== "ACTIVE") return;

    const empId = String(r[0]).trim();
    kpi.total++;

    const status = String(r[8] || "Exited");
    if (status === "In Vietnam") kpi.inVN++;
    else if (status === "Traveling") kpi.business++;
    else kpi.exited++;

    const ctr = activeContractsMap[empId] || {};
    employees.push({
      employeeId: empId,
      fullName: r[1],
      nationality: r[2],
      position: r[4],
      department: r[5],
      currentStatus: status,
      currentLocation: r[9] || "",
      contractNo: ctr.contractNo || "-",
      contractExpiry: ctr.contractExpiry || "-",
      nextEntryExit: nextEntryExitMap[empId] || null,
      nextTravel: nextTravelMap[empId] || null
    });
  });

  return {
    status: "SUCCESS",
    kpi,
    employees,
    expiryWarnings
  };
}

function getDriveFileInfo(fileId, fallbackUrl, fallbackMimeType) {
  const id = String(fileId || "").trim();
  const rawUrl = String(fallbackUrl || "").trim();
  let resolvedId = id;

  if (!resolvedId && rawUrl) {
    const m = rawUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
              rawUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
              rawUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) resolvedId = m[1];
  }

  let mimeType = String(fallbackMimeType || "").toLowerCase();
  let fileName = "";
  let fileUrl = rawUrl;
  let previewUrl = "";

  if (resolvedId) {
    try {
      const file = DriveApp.getFileById(resolvedId);
      mimeType = file.getMimeType() || mimeType;
      fileName = file.getName() || "";
      fileUrl = `https://drive.google.com/file/d/${resolvedId}/view`;
      previewUrl = `https://drive.google.com/file/d/${resolvedId}/preview`;
    } catch (e) {}
  }

  return { fileId: resolvedId, fileName, mimeType, fileUrl, previewUrl };
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
  const docs = docRows.filter(d => String(d[1]) === String(employeeId) && isCurrentValue(d[9]));

  const passport = docs.find(d => d[2] === "PASSPORT") || [];
  const trc = docs.find(d => d[2] === "TRC") || [];
  const wp = docs.find(d => d[2] === "WORK_PERMIT") || [];

  const ctrSheet = getSheet(TABLES.CONTRACTS);
  const ctrRows = ctrSheet.getDataRange().getValues().slice(1);
  const currentContract = ctrRows.find(c => String(c[1]) === String(employeeId) && isCurrentValue(c[9])) || [];

  const appSheet = getSheet(TABLES.APPENDICES);
  const appRows = appSheet.getDataRange().getValues().slice(1);
  const appendices = appRows
    .filter(a => String(a[2]) === String(employeeId))
    .sort((a, b) => parseDateSafe(b[4]) - parseDateSafe(a[4]));
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
      passport: getDriveFileInfo(passport[7], passport[8]),
      trcNo: trc[3] || "",
      trcExpiry: formatDate(trc[5]) || "",
      trcFileId: trc[7] || "",
      trcImg: trc[8] || "",
      trc: getDriveFileInfo(trc[7], trc[8]),
      wpNo: wp[3] || "",
      wpExpiry: formatDate(wp[5]) || "",
      wpFileId: wp[7] || "",
      wpImg: wp[8] || "",
      workPermit: getDriveFileInfo(wp[7], wp[8]),
      contractNo: currentContract[2] || "",
      appendixNo: latestApp[3] || "",
      contractStartDate: formatDate(currentContract[4]) || "",
      contractExpiry: formatDate(currentContract[5]) || "",
      contractFileId: currentContract[10] || "",
      contractImg: currentContract[11] || "",
      contract: getDriveFileInfo(currentContract[10], currentContract[11]),
      salary: currentContract[6] || 0,
      allowance: allowance
    }
  };
}


// ================= MASTER WORKBOOK BACKEND =================
function mt(v){return String(v==null?"":v).trim();}
function md(v){return normalizeImportDate(v);}
function mdt(v){
  if(!v)return "";
  if(Object.prototype.toString.call(v)==="[object Date]" && !isNaN(v.getTime()))
    return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm");
  const s=String(v).trim();
  if(!s)return "";
  const iso=s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if(iso)return `${iso[1]}-${iso[2]}-${iso[3]} ${iso[4]}:${iso[5]}`;
  return s.replace("T"," ").slice(0,16);
}
function mn(v){return normalizeImportNumber(v);}
function mb(v,d=true){if(v===true||/^(true|1)$/i.test(mt(v)))return true;if(v===false||/^(false|0)$/i.test(mt(v)))return false;return d;}

function validateMasterWorkbookServer(wb){
  const errors=[],required=["Employees","Documents","Contracts","ContractAppendices","EntryExit","DomesticTravel"];
  required.forEach(s=>{if(!Array.isArray(wb[s]))errors.push({sheet:s,row:1,message:"Missing dataset."});});
  if(errors.length)return{valid:false,errors};
  const ids=new Set();
  wb.Employees.forEach(r=>{const id=mt(r.employeeId).toUpperCase();if(!id)errors.push({sheet:"Employees",row:r._excelRow,message:"EmployeeID is required."});if(!mt(r.fullName))errors.push({sheet:"Employees",row:r._excelRow,message:"FullName is required."});if(id&&ids.has(id))errors.push({sheet:"Employees",row:r._excelRow,message:"Duplicate EmployeeID: "+id});if(id)ids.add(id);});
  ["Documents","Contracts","ContractAppendices","EntryExit","DomesticTravel"].forEach(s=>wb[s].forEach(r=>{const id=mt(r.employeeId).toUpperCase();if(!id||!ids.has(id))errors.push({sheet:s,row:r._excelRow,message:id?"Unknown EmployeeID: "+id:"EmployeeID is required."});}));
  wb.Contracts.forEach(r=>{if(!mt(r.contractNo))errors.push({sheet:"Contracts",row:r._excelRow,message:"ContractNo is required."});});
  wb.ContractAppendices.forEach(r=>{if(!mt(r.contractId))errors.push({sheet:"ContractAppendices",row:r._excelRow,message:"ContractID is required."});if(!mt(r.appendixNo))errors.push({sheet:"ContractAppendices",row:r._excelRow,message:"AppendixNo is required."});});
  wb.EntryExit.forEach(r=>{if(!["ENTRY","EXIT"].includes(mt(r.type).toUpperCase()))errors.push({sheet:"EntryExit",row:r._excelRow,message:"Type must be ENTRY or EXIT."});if(!r.plannedDateTime&&!r.actualDateTime)errors.push({sheet:"EntryExit",row:r._excelRow,message:"Date/time is required."});if(!mt(r.portName))errors.push({sheet:"EntryExit",row:r._excelRow,message:"PortName is required."});});
  wb.DomesticTravel.forEach(r=>{if(!r.startDate||!r.endDate)errors.push({sheet:"DomesticTravel",row:r._excelRow,message:"StartDate and EndDate are required."});if(r.startDate&&r.endDate&&md(r.endDate)<md(r.startDate))errors.push({sheet:"DomesticTravel",row:r._excelRow,message:"EndDate is before StartDate."});if(!mt(r.fromLocation)||!mt(r.toLocation))errors.push({sheet:"DomesticTravel",row:r._excelRow,message:"FromLocation and ToLocation are required."});});
  return{valid:errors.length===0,errors};
}
function handleImportMasterWorkbook(data,req){
  if(req.role!==ROLES.ADMIN)return{status:"ERROR",message:"Only ADMIN can synchronize the master workbook."};
  const wb=data?.workbook;if(!wb)return{status:"ERROR",message:"Workbook payload is missing."};
  const v=validateMasterWorkbookServer(wb);if(!v.valid)return{status:"ERROR",message:"Validation failed.",errors:v.errors};
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    const result={status:"SUCCESS",fileName:data.fileName||"",employees:{created:0,updated:0},documents:{created:0,updated:0},contracts:{created:0,updated:0},appendices:{created:0,updated:0},entryExit:{created:0,updated:0},domesticTravel:{created:0,updated:0}};
    result.employees=syncMasterEmployees(wb.Employees,req);result.documents=syncMasterDocuments(wb.Documents,req);result.contracts=syncMasterContracts(wb.Contracts,req);result.appendices=syncMasterAppendices(wb.ContractAppendices,req);result.entryExit=syncMasterEntryExit(wb.EntryExit,req);result.domesticTravel=syncMasterTravel(wb.DomesticTravel,req);
    SpreadsheetApp.flush();return result;
  }catch(e){return{status:"ERROR",message:e.message||String(e)}}finally{try{lock.releaseLock()}catch(_){}}
}
function syncMasterEmployees(rows,req){
  const s=getSheet(TABLES.EMPLOYEES),vals=s.getDataRange().getValues(),idx=new Map();for(let i=1;i<vals.length;i++)idx.set(mt(vals[i][0]),i+1);
  let created=0,updated=0;rows.forEach(r=>{const id=mt(r.employeeId),ri=idx.get(id);let folder=ri?vals[ri-1][11]:"";if(!folder)folder=createEmployeeDriveFolder(id,mt(r.fullName));const row=[id,mt(r.fullName),mt(r.nationality),md(r.dob),mt(r.position),mt(r.department),mt(r.phone),mt(r.email).toLowerCase(),mt(r.currentStatus||"Exited"),mt(r.currentLocation||"Overseas"),mt(r.activeStatus||"ACTIVE").toUpperCase(),folder,ri?(vals[ri-1][12]||new Date()):new Date(),new Date()];if(ri){s.getRange(ri,1,1,14).setValues([row]);updated++;}else{s.getRange(s.getLastRow()+1,1,1,14).setValues([row]);created++;}});return{created,updated};
}
function syncMasterDocuments(rows,req){
  const s=getSheet(TABLES.DOCUMENTS),vals=s.getDataRange().getValues(),byId=new Map(),byKey=new Map(),res={created:0,updated:0};for(let i=1;i<vals.length;i++){byId.set(mt(vals[i][0]),i+1);byKey.set(mt(vals[i][1])+"|"+mt(vals[i][2]).toUpperCase()+"|"+mt(vals[i][3]),i+1);}
  rows.forEach(r=>{const emp=mt(r.employeeId),typ=mt(r.docType).toUpperCase(),no=mt(r.docNo),ri=(r.docId&&byId.get(mt(r.docId)))||byKey.get(emp+"|"+typ+"|"+no),id=mt(r.docId)||(ri?vals[ri-1][0]:"DOC-"+Utilities.getUuid().substring(0,8)),row=[id,emp,typ,no,md(r.issueDate),md(r.expiryDate),mt(r.issuer),mt(r.fileId),mt(r.fileUrl),mb(r.isCurrent,true),mt(r.status||"VALID").toUpperCase(),ri?(vals[ri-1][11]||new Date()):new Date(),req.userId];if(ri){s.getRange(ri,1,1,13).setValues([row]);res.updated++;}else{s.getRange(s.getLastRow()+1,1,1,13).setValues([row]);res.created++;}});return res;
}
function syncMasterContracts(rows,req){
  const s=getSheet(TABLES.CONTRACTS),vals=s.getDataRange().getValues(),byId=new Map(),byKey=new Map(),res={created:0,updated:0};for(let i=1;i<vals.length;i++){byId.set(mt(vals[i][0]),i+1);byKey.set(mt(vals[i][1])+"|"+mt(vals[i][2]),i+1);}
  rows.forEach(r=>{const emp=mt(r.employeeId),no=mt(r.contractNo),ri=(r.contractId&&byId.get(mt(r.contractId)))||byKey.get(emp+"|"+no),id=mt(r.contractId)||(ri?vals[ri-1][0]:"CTR-"+Utilities.getUuid().substring(0,8)),row=[id,emp,no,mt(r.contractType||"DEFINITE").toUpperCase(),md(r.startDate),md(r.endDate),mn(r.salary),mt(r.allowancesJson)||JSON.stringify({allowance:0}),mt(r.status||"VALID").toUpperCase(),mb(r.isCurrent,true),mt(r.fileId),mt(r.fileUrl),ri?(vals[ri-1][12]||new Date()):new Date(),req.userId];if(ri){s.getRange(ri,1,1,14).setValues([row]);res.updated++;}else{s.getRange(s.getLastRow()+1,1,1,14).setValues([row]);res.created++;}});return res;
}
function syncMasterAppendices(rows,req){
  const s=getSheet(TABLES.APPENDICES),vals=s.getDataRange().getValues(),byId=new Map(),byKey=new Map(),res={created:0,updated:0};for(let i=1;i<vals.length;i++){byId.set(mt(vals[i][0]),i+1);byKey.set(mt(vals[i][1])+"|"+mt(vals[i][3]),i+1);}
  rows.forEach(r=>{const c=mt(r.contractId),no=mt(r.appendixNo),ri=(r.appendixId&&byId.get(mt(r.appendixId)))||byKey.get(c+"|"+no),id=mt(r.appendixId)||(ri?vals[ri-1][0]:"APP-"+Utilities.getUuid().substring(0,8)),row=[id,c,mt(r.employeeId),no,md(r.effectiveDate),md(r.endDate),mn(r.newSalary),mt(r.allowancesJson)||JSON.stringify({allowance:0}),mt(r.content),mt(r.fileId),mt(r.fileUrl),ri?(vals[ri-1][11]||new Date()):new Date(),req.userId];if(ri){s.getRange(ri,1,1,13).setValues([row]);res.updated++;}else{s.getRange(s.getLastRow()+1,1,1,13).setValues([row]);res.created++;}});return res;
}
function syncMasterEntryExit(rows,req){
  const s=getSheet(TABLES.ENTRY_EXIT),vals=s.getDataRange().getValues(),byId=new Map(),byKey=new Map(),res={created:0,updated:0};for(let i=1;i<vals.length;i++){byId.set(mt(vals[i][0]),i+1);byKey.set(mt(vals[i][1])+"|"+mt(vals[i][12])+"|"+mt(vals[i][2]).toUpperCase(),i+1);}
  rows.forEach(r=>{const emp=mt(r.employeeId),typ=mt(r.type).toUpperCase(),trip=mt(r.tripId),ri=(r.logId&&byId.get(mt(r.logId)))||byKey.get(emp+"|"+trip+"|"+typ),id=mt(r.logId)||(ri?vals[ri-1][0]:"EE-"+Utilities.getUuid().substring(0,8)),planned=mdt(r.plannedDateTime),actual=mdt(r.actualDateTime),event=actual||planned,row=[id,emp,typ,event,mt(r.portName),mt(r.flightNo),mt(r.destination),mt(r.purpose),ri?(vals[ri-1][8]||new Date()):new Date(),ri?(vals[ri-1][9]||req.userId):req.userId,planned,actual,trip];if(ri){s.getRange(ri,1,1,13).setValues([row]);res.updated++;}else{s.getRange(s.getLastRow()+1,1,1,13).setValues([row]);res.created++;}});return res;
}
function syncMasterTravel(rows,req){
  const s=getSheet(TABLES.TRAVEL),vals=s.getDataRange().getValues(),byId=new Map(),byKey=new Map(),res={created:0,updated:0};for(let i=1;i<vals.length;i++){byId.set(mt(vals[i][0]),i+1);byKey.set(mt(vals[i][1])+"|"+mt(vals[i][4])+"|"+mt(vals[i][5])+"|"+mt(vals[i][2])+"|"+mt(vals[i][3]),i+1);}
  rows.forEach(r=>{const emp=mt(r.employeeId),st=md(r.startDate),en=md(r.endDate),key=emp+"|"+st+"|"+en+"|"+mt(r.fromLocation)+"|"+mt(r.toLocation),ri=(r.travelId&&byId.get(mt(r.travelId)))||byKey.get(key),id=mt(r.travelId)||(ri?vals[ri-1][0]:"TR-"+Utilities.getUuid().substring(0,8)),row=[id,emp,mt(r.fromLocation),mt(r.toLocation),st,en,mt(r.status||"PENDING").toUpperCase(),mt(r.ticketFileId),mt(r.ticketUrl),ri?(vals[ri-1][9]||new Date()):new Date(),ri?(vals[ri-1][10]||req.userId):req.userId,mt(r.purpose)];if(ri){s.getRange(ri,1,1,12).setValues([row]);res.updated++;}else{s.getRange(s.getLastRow()+1,1,1,12).setValues([row]);res.created++;}});return res;
}
function handleAddEntryExitTrip(data,req){
  if(!data?.employeeId||!data.entry||!data.exit)return{status:"ERROR",message:"Invalid round-trip data."};
  const tripId=mt(data.tripId)||"TRIP-"+Utilities.getUuid().substring(0,8),e=handleAddEntryExit({...data.entry,employeeId:data.employeeId,tripId,type:"ENTRY"},req);if(e.status!=="SUCCESS")return e;
  const x=handleAddEntryExit({...data.exit,employeeId:data.employeeId,tripId,type:"EXIT"},req);if(x.status!=="SUCCESS"){handleDeleteEntryExit({entryExitId:e.entryExitId},req);return x;}return{status:"SUCCESS",message:"Round trip saved successfully.",tripId,entryExitIds:[e.entryExitId,x.entryExitId]};
}
function handleUpdateEntryExitTrip(data,req){
  if(!data?.employeeId||!data?.tripId||!data.entry||!data.exit)return{status:"ERROR",message:"Invalid round-trip data."};
  const rows=getSheet(TABLES.ENTRY_EXIT).getDataRange().getValues(),trip=rows.map((r,i)=>({r,i:i+1})).filter(x=>x.i>1&&mt(x.r[1])===mt(data.employeeId)&&mt(x.r[12])===mt(data.tripId)),entry=trip.find(x=>mt(x.r[2]).toUpperCase()==="ENTRY"),exit=trip.find(x=>mt(x.r[2]).toUpperCase()==="EXIT");if(!entry||!exit)return{status:"ERROR",message:"Trip must contain both Entry and Exit."};
  const a=handleUpdateEntryExit({...data.entry,employeeId:data.employeeId,tripId:data.tripId,type:"ENTRY",entryExitId:entry.r[0]},req);if(a.status!=="SUCCESS")return a;return handleUpdateEntryExit({...data.exit,employeeId:data.employeeId,tripId:data.tripId,type:"EXIT",entryExitId:exit.r[0]},req);
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
  const allowedRoles = [ROLES.ADMIN, ROLES.HR_ADMIN];
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

  rows.forEach((item, idx) => {
    const excelRow = Number(item._excelRow || idx + 2);
    const employeeId = normalizeImportText(item.employeeId);
    if (!employeeId || !existingById[employeeId]) return;

    try {
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
    if (isCurrentValue(rows[i][9]) && currentIndex < 0) currentIndex = i + 1;
    if (docNo && String(rows[i][3]).trim() === docNo) exactIndex = i + 1;
  }

  if (exactIndex > 0) {
    const oldRow = rows[exactIndex - 1];
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

function upsertImportedContract(item, requester) {
  const sheet = getSheet(TABLES.CONTRACTS);
  const rows = sheet.getDataRange().getValues();
  const employeeId = normalizeImportText(item.employeeId);
  const contractNo = normalizeImportText(item.contractNo);

  let exactIndex = -1;
  let currentIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== employeeId) continue;
    if (isCurrentValue(rows[i][9]) && currentIndex < 0) currentIndex = i + 1;
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
// I. THỦ TỤC NGHỈ VIỆC & KHÓA TÀI KHOẢN
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
      invalidateUserSessions(userRows[i][0]);
      break;
    }
  }

  writeAuditLog(req.userId, "DEACTIVATE", "Employees", empId, "ACTIVE", "INACTIVE");
  return { status: "SUCCESS", message: "Đã vô hiệu hóa hồ sơ và khóa tài khoản thành công." };
}

// ==========================================
// J. ACCOUNT / PASSWORD MANAGEMENT
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
  
  // Hủy tất cả các phiên đăng nhập khác của tài khoản này
  invalidateUserSessions(requester.userId);

  writeAuditLog(requester.userId, "CHANGE_PASSWORD", "Users", requester.userId, null, "Password changed");
  return { status: "SUCCESS", message: "Đổi mật khẩu thành công. Vui lòng đăng nhập lại!" };
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
// K. SETUP HỆ THỐNG & HELPER
// ==========================================
function setupSystem() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const schema = {
    [TABLES.USERS]: ['UserID', 'Email', 'PasswordHash', 'Role', 'EmployeeID', 'Status', 'CreatedAt'],
    [TABLES.EMPLOYEES]: ['EmployeeID', 'FullName', 'Nationality', 'DOB', 'Position', 'Department', 'Phone', 'Email', 'CurrentStatus', 'CurrentLocation', 'ActiveStatus', 'FolderId', 'CreatedAt', 'UpdatedAt'],
    [TABLES.DOCUMENTS]: ['DocID', 'EmployeeID', 'DocType', 'DocNo', 'IssueDate', 'ExpiryDate', 'Issuer', 'FileId', 'FileUrl', 'IsCurrent', 'Status', 'CreatedAt', 'CreatedBy'],
    [TABLES.CONTRACTS]: ['ContractID', 'EmployeeID', 'ContractNo', 'ContractType', 'StartDate', 'EndDate', 'Salary', 'AllowancesJson', 'Status', 'IsCurrent', 'FileId', 'FileUrl', 'CreatedAt', 'CreatedBy'],
    [TABLES.APPENDICES]: ['AppendixID', 'ContractID', 'EmployeeID', 'AppendixNo', 'EffectiveDate', 'EndDate', 'NewSalary', 'AllowancesJson', 'Content', 'FileId', 'FileUrl', 'CreatedAt', 'CreatedBy'],
    [TABLES.ENTRY_EXIT]: ['LogID', 'EmployeeID', 'Type', 'EventDate', 'PortName', 'FlightNo', 'Destination', 'Purpose', 'CreatedAt', 'CreatedBy', 'PlannedDateTime', 'ActualDateTime', 'TripID'],
    [TABLES.TRAVEL]: ['TravelID', 'EmployeeID', 'FromLocation', 'ToLocation', 'FromDate', 'ToDate', 'Status', 'TicketFileId', 'TicketUrl', 'CreatedAt', 'CreatedBy', 'Purpose'],
    [TABLES.AUDIT]: ['AuditID', 'UserID', 'Action', 'Module', 'TargetID', 'OldValue', 'NewValue', 'Timestamp', 'IPAddress'],
    [TABLES.CONFIG]: ['Key', 'Value', 'UpdatedAt']
  };

  Object.keys(schema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = schema[sheetName];
    const currentCols = Math.max(sheet.getLastColumn(), headers.length);
    if (sheet.getLastColumn() < headers.length) {
      sheet.insertColumnsAfter(Math.max(1, sheet.getLastColumn()), headers.length - Math.max(1, sheet.getLastColumn()));
    }
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

function createEmployeeDriveFolder(employeeId, fullName) {
  try {
    if (!employeeId) {
      throw new Error("Missing employeeId");
    }

    const props = PropertiesService.getScriptProperties();
    const configuredRootId = props.getProperty("DRIVE_FOLDER_ID");

    let rootFolder = null;

    if (configuredRootId) {
      try {
        rootFolder = DriveApp.getFolderById(configuredRootId);
      } catch (e) {
        console.warn("DRIVE_FOLDER_ID không hợp lệ, chuyển sang tìm theo tên.");
      }
    }

    if (!rootFolder) {
      const ROOT_NAME = "Foreign Employee Management System";
      const folders = DriveApp.getFoldersByName(ROOT_NAME);

      if (folders.hasNext()) {
        rootFolder = folders.next();
        props.setProperty("DRIVE_FOLDER_ID", rootFolder.getId());
      } else {
        rootFolder = DriveApp.createFolder(ROOT_NAME);
        props.setProperty("DRIVE_FOLDER_ID", rootFolder.getId());
      }
    }

    const safeEmployeeId = String(employeeId).trim();
    const safeFullName = String(fullName || "Unknown")
      .trim()
      .replace(/[\\/:*?"<>|]/g, "_");

    const folderName = `${safeEmployeeId} - ${safeFullName}`;
    const existingFolders = rootFolder.getFoldersByName(folderName);

    let employeeFolder;
    if (existingFolders.hasNext()) {
      employeeFolder = existingFolders.next();
    } else {
      employeeFolder = rootFolder.createFolder(folderName);
    }

    const subFolders = ["Passport", "Visa", "TRC", "Work Permit", "Contracts"];
    subFolders.forEach(function (name) {
      const folders = employeeFolder.getFoldersByName(name);
      if (!folders.hasNext()) {
        employeeFolder.createFolder(name);
      }
    });

    return employeeFolder.getId();

  } catch (e) {
    console.error("createEmployeeDriveFolder error:", e);
    return "";
  }
}

function formatDate(d) {
  if (!d) return "";
  if (Object.prototype.toString.call(d) === "[object Date]") {
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return "";
  }
  const s = String(d).trim();
  if (!s) return "";
  
  let match = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (match) return `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;

  match = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (match) return `${match[3]}-${String(match[2]).padStart(2, "0")}-${String(match[1]).padStart(2, "0")}`;

  return s;
}

function writeAuditLog(userId, action, module, targetId, oldValue, newValue) {
  try {
    const sheet = getSheet(TABLES.AUDIT);
    if (!sheet) throw new Error("AuditLog sheet not found");

    const logId = "LOG-" + Utilities.getUuid().replace(/-/g, "").substring(0, 8).toUpperCase();

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
    console.error("writeAuditLog error:", e);
  }
}

function handleGetTimeline(empId, req) {
  if (!canAccessEmployee(req, empId)) return { status: "ERROR", message: "Không có quyền xem thông tin." };

  const eeSheet = getSheet(TABLES.ENTRY_EXIT);
  const eeRows = eeSheet.getDataRange().getValues().slice(1);
  const travelSheet = getSheet(TABLES.TRAVEL);
  const travelRows = travelSheet.getDataRange().getValues().slice(1);

  const timeline = [];

  eeRows.filter(r => String(r[1]) === String(empId)).forEach(r => {
    let flightNo = String(r[5] || "");
    let destination = String(r[6] || "");
    let purpose = String(r[7] || "");
    try {
      const noteObj = JSON.parse(String(r[5] || ""));
      if (noteObj && typeof noteObj === "object") {
        flightNo = noteObj.flightNo || "";
        destination = noteObj.destination || "";
        purpose = noteObj.purpose || "";
      }
    } catch (e) {}

    timeline.push({
      id: r[0],
      kind: "ENTRY_EXIT",
      rawDate: parseDateSafe(r[3]),
      date: formatDateTime(r[3]),
      title: r[2] === "ENTRY" ? "✈️ Nhập cảnh Việt Nam" : "🛫 Xuất cảnh Việt Nam",
      status: r[2],
      location: `Cửa khẩu: ${r[4] || "-"} | Chuyến bay: ${flightNo || "-"} | Điểm đến: ${destination || "-"}`,
      purpose: purpose || "",
      tripId: r[12] || ""
    });
  });

  travelRows.filter(r => String(r[1]) === String(empId) && String(r[6] || "").toUpperCase() !== "CANCELLED").forEach(r => {
    timeline.push({
      id: r[0],
      kind: "DOMESTIC_TRAVEL",
      rawDate: parseDateSafe(r[4]),
      date: formatDate(r[4]),
      endDate: formatDate(r[5]),
      title: "🚗 Đi lại / Công tác nội địa",
      status: String(r[6] || "").toUpperCase(),
      location: `${r[2] || "-"} → ${r[3] || "-"}`,
      purpose: r[11] || ""
    });
  });

  timeline.sort((a, b) => b.rawDate - a.rawDate);
  return { status: "SUCCESS", timeline };
}

function formatDateTime(value) {
  if (!value) return "";
  const d = parseDateSafe(value);
  if (!d.getTime()) return String(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}
