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
// B. ROUTER API (doGet & doPost)
// ==========================================
function doGet(e) {
  // Khi truy cập Web App không có action, trả về giao diện index.html.
  // Khi có action, tiếp tục xử lý API GET như bình thường.
  const action = e && e.parameter ? String(e.parameter.action || "").trim() : "";
  if (!action) {
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("Foreign Employee Management V2")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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
    
    // 1. Login công khai
    if (action === "LOGIN") {
      const res = handleLogin(requestData.username, requestData.password);
      return responseHeader.setContent(JSON.stringify(res));
    }

    // 2. Xác thực Session
    const token = requestData.token || e.parameter.token;
    const requester = validateSession(token);
    if (!requester) {
      return responseHeader.setContent(JSON.stringify({ status: "ERROR", message: "Phiên đăng nhập hết hạn hoặc không hợp lệ!" }));
    }

    let result = { status: "ERROR", message: "Hành động không hợp lệ!" };

    // 3. Phân luồng API Action
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
      case "SAVE_DOCUMENT": {
        const data = requestData.data || requestData;
        if (![ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(requester.role)) {
          result = { status: "ERROR", message: "Không có quyền cập nhật giấy tờ." };
        } else {
          result = saveDocument(data, requester);
        }
        break;
      }
      case "UPLOAD_FILE":
        result = uploadEmployeeFile(requestData.data || requestData, requester);
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
      case "APPROVE_DOMESTIC_TRAVEL":
        result = handleApproveDomesticTravel(requestData.data || requestData, requester);
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
    
    // Kiểm tra match email/employeeId và password hash (hoặc plain cho tài khoản cũ chưa hash)
    if ((String(email).toLowerCase() === String(username).toLowerCase() || String(employeeId) === String(username)) && 
        (String(passHash) === inputHash || String(passHash) === String(password))) {
      
      if (status !== "ACTIVE") {
        return { status: "ERROR", message: "Tài khoản của bạn đã bị vô hiệu hóa!" };
      }

      // Kiểm tra xem trạng thái của Employee có ACTIVE không
      const empSheet = getSheet(TABLES.EMPLOYEES);
      const empRows = empSheet.getDataRange().getValues();
      const emp = empRows.find(r => String(r[0]) === String(employeeId));
      if (emp && emp[10] !== "ACTIVE") {
        return { status: "ERROR", message: "Hồ sơ nhân viên đã dừng hoạt động!" };
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
  
  // Kiểm tra TTL 6 Giờ
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

  // 1. CẬP NHẬT HOẶC TẠO MỚI EMPLOYEE
  if (rowIndex > 0) {
    const currentRow = rows[rowIndex - 1];
    
    // Nếu là USER thường, hạn chế sửa một số trường nhạy cảm
    const updatedRow = [
      currentRow[0],
      isHRAdmin ? (data.fullName || currentRow[1]) : currentRow[1],
      isHRAdmin ? (data.nationality || currentRow[2]) : currentRow[2],
      isHRAdmin ? (formatDate(data.dob) || currentRow[3]) : currentRow[3],
      isHRAdmin ? (data.position || currentRow[4]) : currentRow[4],
      isHRAdmin ? (data.department || currentRow[5]) : currentRow[5],
      data.phone !== undefined ? data.phone : currentRow[6],
      data.email !== undefined ? data.email : currentRow[7],
      currentRow[8], // CurrentStatus
      currentRow[9], // CurrentLocation
      (isHRAdmin && data.activeStatus !== undefined) ? data.activeStatus : currentRow[10],
      currentRow[11], // FolderId
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

    // Tạo tài khoản mặc định (Password được HASH)
    const userSheet = getSheet(TABLES.USERS);
    userSheet.appendRow(["USR-" + Utilities.getUuid().substring(0, 6), data.email || data.employeeId, hashPassword("123456"), ROLES.USER, data.employeeId, "ACTIVE", timestamp]);
    writeAuditLog(requester.userId, "CREATE", "Employees", data.employeeId, null, JSON.stringify(newRow));
  }

  // 2. LƯU CÁC GIẤY TỜ PHÁP LÝ (Passport, TRC, WP)
  if (isHRAdmin) {
    if (data.passportNo) saveDocument({ employeeId: data.employeeId, docType: "PASSPORT", docNo: data.passportNo, expiryDate: data.passportExpiry, fileUrl: data.passportImg }, requester);
    if (data.trcNo) saveDocument({ employeeId: data.employeeId, docType: "TRC", docNo: data.trcNo, expiryDate: data.trcExpiry, fileUrl: data.trcImg }, requester);
    if (data.wpNo) saveDocument({ employeeId: data.employeeId, docType: "WORK_PERMIT", docNo: data.wpNo, expiryDate: data.wpExpiry, fileUrl: data.wpImg }, requester);

    // 3. LƯU HỢP ĐỒNG & PHỤ LỤC (Fix lỗi lớn nhất của V1)
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
  
  // Tìm contract hiện tại đang IsCurrent = true
  for (let i = 1; i < contractRows.length; i++) {
    if (String(contractRows[i][1]) === String(data.employeeId) && contractRows[i][9] === true) {
      if (contractRows[i][2] === data.contractNo) {
        currentContractId = contractRows[i][0];
      } else {
        // Đổi sang hợp đồng mới -> Đánh dấu HĐ cũ IsCurrent = false
        contractSheet.getRange(i + 1, 10).setValue(false);
      }
    }
  }

  // Tạo Contract mới nếu chưa có
  if (!currentContractId) {
    currentContractId = "CTR-" + Utilities.getUuid().substring(0, 8);
    const newContractRow = [
      currentContractId,
      data.employeeId,
      data.contractNo,
      "DEFINITE", // Mặc định hoặc truyền thêm
      formatDate(data.contractStartDate),
      formatDate(data.contractExpiry),
      Number(data.salary) || 0,
      JSON.stringify({ allowance: Number(data.allowance) || 0 }),
      "VALID",
      true, // IsCurrent
      "",
      data.contractImg || "",
      new Date(),
      requester.userId
    ];
    contractSheet.appendRow(newContractRow);
    writeAuditLog(requester.userId, "ADD_CONTRACT", "Contracts", currentContractId, null, JSON.stringify(newContractRow));
  } else {
    // Cập nhật thông tin hợp đồng hiện tại
    for (let i = 1; i < contractRows.length; i++) {
      if (contractRows[i][0] === currentContractId) {
        contractSheet.getRange(i + 1, 5).setValue(formatDate(data.contractStartDate));
        contractSheet.getRange(i + 1, 6).setValue(formatDate(data.contractExpiry));
        contractSheet.getRange(i + 1, 7).setValue(Number(data.salary) || 0);
        contractSheet.getRange(i + 1, 8).setValue(JSON.stringify({ allowance: Number(data.allowance) || 0 }));
        if (data.contractImg) contractSheet.getRange(i + 1, 12).setValue(data.contractImg);
        break;
      }
    }
  }

  // Thêm Phụ lục hợp đồng nếu có nhập appendixNo
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
        "",
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
  
  // Đánh dấu bản ghi cùng DocType trước đó thành IsCurrent = false
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
    true, // IsCurrent
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

    const encoded = String(base64Data).split(",")[1] || "";
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
    
    // AN TOÀN BẢO MẬT: Bỏ ANYONE_WITH_LINK. Để Private.
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
  
  // Quy trình chuẩn: Đăng ký bởi USER phải là PENDING
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
  const status = String(d?.status || "").toUpperCase();
  const travelId = String(d?.travelId || "").trim();
  if (!travelId || !["APPROVED", "REJECTED"].includes(status)) {
    return { status: "ERROR", message: "Trạng thái duyệt không hợp lệ." };
  }

  const sheet = getSheet(TABLES.TRAVEL);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.travelId) {
      sheet.getRange(i + 1, 7).setValue(status); // APPROVED / REJECTED
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
// G. DASHBOARD & PROFILE RESPONSE (V2 FIX)
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

  // Indexing hợp đồng hiện hành
  const activeContractsMap = {};
  ctrRows.forEach(c => {
    if (c[9] === true) { // IsCurrent === true
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

  // 1. Kiểm tra Cảnh báo hết hạn Giấy tờ (Fix logic diffDays <= 45)
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

  // 2. Kiểm tra Cảnh báo Hợp đồng lao động (Fix KPI expContract luôn bằng 0)
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

  // Lấy Giấy tờ
  const docSheet = getSheet(TABLES.DOCUMENTS);
  const docRows = docSheet.getDataRange().getValues().slice(1);
  const docs = docRows.filter(d => String(d[1]) === String(employeeId) && d[9] === true);

  const passport = docs.find(d => d[2] === "PASSPORT") || [];
  const trc = docs.find(d => d[2] === "TRC") || [];
  const wp = docs.find(d => d[2] === "WORK_PERMIT") || [];

  // Lấy Hợp đồng hiện tại (Fix lỗi Backend không trả về Contract)
  const ctrSheet = getSheet(TABLES.CONTRACTS);
  const ctrRows = ctrSheet.getDataRange().getValues().slice(1);
  const currentContract = ctrRows.find(c => String(c[1]) === String(employeeId) && c[9] === true) || [];

  // Lấy Phụ lục mới nhất
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
      passportImg: passport[8] || "",
      trcNo: trc[3] || "",
      trcExpiry: formatDate(trc[5]) || "",
      trcImg: trc[8] || "",
      wpNo: wp[3] || "",
      wpExpiry: formatDate(wp[5]) || "",
      wpImg: wp[8] || "",
      contractNo: currentContract[2] || "",
      appendixNo: latestApp[3] || "",
      contractStartDate: formatDate(currentContract[4]) || "",
      contractExpiry: formatDate(currentContract[5]) || "",
      contractImg: currentContract[11] || "",
      salary: currentContract[6] || 0,
      allowance: allowance
    }
  };
}

// ==========================================
// H. THỦ TỤC NGHỈ VIỆC & KHÓA TÀI KHOẢN (V2 FIX)
// ==========================================
function handleDeleteEmployee(empId, req) {
  if (![ROLES.ADMIN, ROLES.HR_ADMIN].includes(req.role)) {
    return { status: "ERROR", message: "Không có quyền thực hiện thao tác này!" };
  }

  // 1. Vô hiệu hóa trong Employees
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

  // 2. Vô hiệu hóa tài khoản User tương ứng (Fix lỗi nghỉ việc vẫn đăng nhập được)
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
      sheet.appendRow(schema[sheetName]);
      sheet.getRange(1, 1, 1, schema[sheetName].length).setFontWeight("bold").setBackground("#f3f3f3");
    }
  });

  // Tạo tài khoản Admin mặc định có Mật khẩu đã Hash
  const userSheet = ss.getSheetByName(TABLES.USERS);
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(["USR-ADMIN", "admin@company.com", hashPassword("admin123"), ROLES.ADMIN, "EMP-ADMIN", "ACTIVE", new Date()]);
  }
}

function createEmployeeDriveFolder(employeeId, fullName) {
  const folderName = "Foreign Employee Management System";
  const rootFolders = DriveApp.getFoldersByName(folderName);
  let rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(folderName);
  
  const empFolder = rootFolder.createFolder(`${employeeId} - ${fullName}`);
  empFolder.createFolder("Passport");
  empFolder.createFolder("Visa");
  empFolder.createFolder("TRC");
  empFolder.createFolder("Work Permit");
  empFolder.createFolder("Contracts");
  
  return empFolder.getId();
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
