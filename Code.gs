const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  USERS: "Users",
  EMPLOYEES: "Employees",
  ENTRY_EXIT: "EntryExit",
  DOMESTIC: "DomesticTravel",
  DOCUMENTS: "Documents",
  AUDIT: "AuditLog"
};

const SESSION_TTL_SECONDS = 21600; // 6 giờ
const EXPIRING_DAYS = 30;

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isHRorAdmin(role) {
  const r = String(role || "").toUpperCase();
  return ["HR", "HR ADMIN", "DIRECTOR", "ADMIN"].includes(r);
}

function canEdit(role) {
  const r = String(role || "").toUpperCase();
  return ["HR", "HR ADMIN", "ADMIN"].includes(r);
}

function getRequester(userId) {
  if (!userId) return null;
  const sheet = SPREADSHEET.getSheetByName(SHEETS.USERS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const [id, email, password, role, name, employeeId, active] = data[i];
    if (active !== false && String(active).toUpperCase() !== "FALSE" &&
        (String(id) === String(userId) || String(email).toLowerCase() === String(userId).toLowerCase())) {
      return { id, email, role, name, employeeId };
    }
  }
  return null;
}

function createSession(user) {
  const token = Utilities.getUuid() + "-" + Utilities.getUuid();
  CacheService.getScriptCache().put("SESSION_" + token, JSON.stringify({
    id: user.id, createdAt: Date.now()
  }), SESSION_TTL_SECONDS);
  return token;
}

function getSessionUser(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get("SESSION_" + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  return getRequester(s.id);
}

function destroySession(token) {
  if (token) CacheService.getScriptCache().remove("SESSION_" + token);
}

function requireSession(token) {
  const user = getSessionUser(token);
  if (!user) throw new Error("Phiên đăng nhập không hợp lệ hoặc đã hết hạn. Vui lòng đăng nhập lại.");
  return user;
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || "";
    if (action === "GET_PROFILE" || action === "GET_ENTRY_EXIT" ||
        action === "GET_DOMESTIC_TRAVEL" || action === "GET_TIMELINE") {
      const requester = requireSession(p.token);
      const employeeId = p.employeeId;
      if (!employeeId) return respondJSON({status:"ERROR", message:"Thiếu employeeId"});
      if (!isHRorAdmin(requester.role) && String(requester.employeeId) !== String(employeeId)) {
        return respondJSON({status:"ERROR", message:"Bạn không có quyền truy cập dữ liệu này."});
      }
      if (action === "GET_PROFILE") return getProfile(employeeId);
      if (action === "GET_ENTRY_EXIT") return getEntryExit(employeeId);
      if (action === "GET_DOMESTIC_TRAVEL") return getDomesticTravel(employeeId);
      return getEmployeeTimeline(employeeId);
    }

    if (action === "GET_DASHBOARD_DATA" || action === "GET_EMPLOYEES") {
      const requester = requireSession(p.token);
      if (!isHRorAdmin(requester.role)) return respondJSON({status:"ERROR", message:"Bạn không có quyền."});
      return action === "GET_DASHBOARD_DATA" ? getDashboardData() : getEmployees();
    }

    return respondJSON({status:"ERROR", message:"Action không hợp lệ"});
  } catch (err) {
    return respondJSON({status:"ERROR", message:err.message || String(err)});
  }
}

function doPost(e) {
  try {
    const request = JSON.parse((e.postData && e.postData.contents) || "{}");
    const action = request.action;

    if (action === "LOGIN") return handleLogin(request);
    if (action === "LOGOUT") {
      destroySession(request.token);
      return respondJSON({status:"SUCCESS"});
    }

    const requester = requireSession(request.token);
    const data = request.data || {};

    switch (action) {
      case "SAVE_PROFILE":
        if (!canEdit(requester.role) && String(requester.employeeId) !== String(data.employeeId))
          return respondJSON({status:"ERROR", message:"Bạn không có quyền sửa hồ sơ này."});
        return saveProfile(data, requester.id, canEdit(requester.role));

      case "ADD_ENTRY_EXIT":
        if (!requester.employeeId || String(requester.employeeId) !== String(data.employeeId)) {
          if (!isHRorAdmin(requester.role))
            return respondJSON({status:"ERROR", message:"Bạn chỉ có thể đăng ký cho chính mình."});
        }
        validateEntryExit(data);
        return addEntryExit(data, requester.id);

      case "ADD_DOMESTIC_TRAVEL":
        if (!requester.employeeId || String(requester.employeeId) !== String(data.employeeId)) {
          if (!isHRorAdmin(requester.role))
            return respondJSON({status:"ERROR", message:"Bạn chỉ có thể đăng ký cho chính mình."});
        }
        validateDomesticTravel(data);
        return addDomesticTravel(data, requester.id);

      case "DELETE_EMPLOYEE":
        if (!canEdit(requester.role))
          return respondJSON({status:"ERROR", message:"Chỉ có Admin/HR mới có quyền xóa nhân viên."});
        return deleteEmployee(data.employeeId, requester.id);

      default:
        return respondJSON({status:"ERROR", message:"Action không hợp lệ"});
    }
  } catch (err) {
    return respondJSON({status:"ERROR", message:err.message || String(err)});
  }
}

function handleLogin(req) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.USERS);
  if (!sheet) return respondJSON({status:"ERROR", message:"Không tìm thấy Sheet Users"});
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [id, email, password, role, name, employeeId, active] = data[i];
    if (active !== false && String(active).toUpperCase() !== "FALSE" &&
        (String(email).toLowerCase() === String(req.username || "").toLowerCase() || String(id) === String(req.username || "")) &&
        String(password) === String(req.password || "")) {
      const user = {id, email, role, name, employeeId};
      const token = createSession(user);
      writeAudit(id, "LOGIN", employeeId, "User logged in");
      return respondJSON({status:"SUCCESS", user:user, token:token, expiresIn:SESSION_TTL_SECONDS});
    }
  }
  return respondJSON({status:"FAILED", message:"Thông tin đăng nhập không hợp lệ"});
}

function getDashboardData() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return respondJSON({status:"ERROR", message:"Employees sheet not found"});
  autoUpdateDomesticTravelStatuses();
  const rows = sheet.getDataRange().getValues().slice(1);
  let total=0, inVN=0, business=0, exited=0, expPassport=0, expVisa=0, expTRC=0, expWP=0, expContract=0;
  const employees = [], expiryWarnings = [];

  rows.forEach(row => {
    const activeStatus = String(row[31] || "ACTIVE").toUpperCase();
    if (activeStatus === "INACTIVE" || activeStatus === "DELETED") return;

    total++;
    const employee = parseEmployeeRow(row);
    employees.push(employee);

    const s = employee.currentStatus.toLowerCase();
    if (["in vietnam", "ở việt nam", "invietnam"].includes(s)) inVN++;
    else if (["traveling", "đi lại", "trong nước"].includes(s)) business++;
    else if (["exited", "xuất cảnh", "đã xuất cảnh"].includes(s)) exited++;

    const docs = [
      ["Passport", row[7]],
      ["Visa", row[9]],
      ["TRC", row[11]],
      ["Work Permit", row[13]],
      ["Hợp Đồng", row[25]]
    ];

    docs.forEach(([type, val]) => {
      const info = expiryInfo(val);
      if (!info) return;
      if (info.status !== "VALID") {
        expiryWarnings.push({
          employeeId: employee.employeeId,
          fullName: employee.fullName,
          document: type,
          expiryDate: formatDate(val),
          daysRemaining: info.days,
          severity: info.status
        });
      }
      if (type === "Passport" && info.status !== "VALID") expPassport++;
      if (type === "Visa" && info.status !== "VALID") expVisa++;
      if (type === "TRC" && info.status !== "VALID") expTRC++;
      if (type === "Work Permit" && info.status !== "VALID") expWP++;
      if (type === "Hợp Đồng" && info.status !== "VALID") expContract++;
    });
  });

  return respondJSON({
    status: "SUCCESS",
    kpi: { total, inVN, business, exited, expPassport, expVisa, expTRC, expWP, expContract, warning: expiryWarnings.length },
    employees,
    expiryWarnings
  });
}

function parseEmployeeRow(row) {
  return {
    id: row[0],
    employeeId: row[1],
    fullName: row[2],
    nationality: row[3],
    dob: formatDate(row[4]),
    passportNo: row[5],
    passportIssueDate: formatDate(row[6]),
    passportExpiry: formatDate(row[7]),
    visaNo: row[8],
    visaExpiry: formatDate(row[9]),
    trcNo: row[10],
    trcExpiry: formatDate(row[11]),
    wpNo: row[12],
    wpExpiry: formatDate(row[13]),
    position: row[14],
    department: row[15],
    phone: row[16],
    email: row[17],
    currentStatus: String(row[18] || "").trim(),
    currentLocation: String(row[19] || "").trim(),
    contractNo: row[20] || "",
    appendixNo: row[21] || "",
    salary: row[22] || 0,
    allowance: row[23] || 0,
    contractStartDate: formatDate(row[24]),
    contractExpiry: formatDate(row[25]),
    passportImg: row[26] || "",
    visaImg: row[27] || "",
    trcImg: row[28] || "",
    wpImg: row[29] || "",
    contractImg: row[30] || "",
    activeStatus: row[31] || "ACTIVE"
  };
}

function getEmployees() {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return respondJSON({status:"ERROR", message:"Employees sheet not found"});
  const data = sheet.getDataRange().getValues().slice(1);
  const result = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][31] || "ACTIVE").toUpperCase() !== "INACTIVE") {
      result.push(parseEmployeeRow(data[i]));
    }
  }
  return respondJSON({status:"SUCCESS", employees:result});
}

function getProfile(employeeId) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return respondJSON({status:"ERROR", message:"Employees sheet not found"});
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase()) {
      return respondJSON({status:"SUCCESS", profile: parseEmployeeRow(data[i])});
    }
  }
  return respondJSON({status:"FAILED", message:"Employee not found"});
}

function saveProfile(data, updatedById, allowStatus) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return respondJSON({status:"ERROR", message:"Employees sheet not found"});
  const rows = sheet.getDataRange().getValues();
  let foundIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === String(data.employeeId).trim().toLowerCase()) {
      foundIndex = i + 1;
      break;
    }
  }

  const values = [
    data.fullName || "", data.nationality || "", data.dob || "",
    data.passportNo || "", data.passportIssueDate || "", data.passportExpiry || "",
    data.visaNo || "", data.visaExpiry || "",
    data.trcNo || "", data.trcExpiry || "",
    data.wpNo || "", data.wpExpiry || "",
    data.position || "", data.department || "", data.phone || "", data.email || "",
    data.currentStatus || "In Vietnam", data.currentLocation || "Việt Nam",
    data.contractNo || "", data.appendixNo || "", data.salary || 0, data.allowance || 0,
    data.contractStartDate || "", data.contractExpiry || "",
    data.passportImg || "", data.visaImg || "", data.trcImg || "", data.wpImg || "", data.contractImg || "",
    data.activeStatus || "ACTIVE"
  ];

  if (foundIndex > 0) {
    sheet.getRange(foundIndex, 3, 1, values.length).setValues([values]);
    writeAudit(updatedById, "SAVE_PROFILE", data.employeeId, "Updated employee profile");
    return respondJSON({status:"SUCCESS", message:"Cập nhật hồ sơ nhân viên thành công"});
  } else {
    const newId = "EMP-" + Date.now();
    const newRow = [newId, data.employeeId, ...values];
    sheet.appendRow(newRow);
    writeAudit(updatedById, "ADD_EMPLOYEE", data.employeeId, "Created new employee");
    return respondJSON({status:"SUCCESS", message:"Thêm mới nhân viên thành công"});
  }
}

function deleteEmployee(employeeId, updatedById) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sheet) return respondJSON({status:"ERROR", message:"Employees sheet not found"});
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase()) {
      sheet.getRange(i + 1, 32).setValue("INACTIVE"); // Đánh dấu ngưng hoạt động
      writeAudit(updatedById, "DELETE_EMPLOYEE", employeeId, "Soft deleted employee");
      return respondJSON({status:"SUCCESS", message:"Đã ngưng hoạt động nhân viên thành công."});
    }
  }
  return respondJSON({status:"FAILED", message:"Employee không tồn tại."});
}

function validateEmployeeExists(employeeId) {
  if (!employeeId) throw new Error("Thiếu employeeId.");
  const sh = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES);
  if (!sh) throw new Error("Employees sheet not found");
  const data = sh.getDataRange().getValues();
  const ok = data.slice(1).some(r => String(r[1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase());
  if (!ok) throw new Error("Không tìm thấy nhân viên: " + employeeId);
}

function validateEntryExit(data) {
  validateEmployeeExists(data.employeeId);
  if (!["ENTRY","EXIT"].includes(String(data.type))) throw new Error("Type phải là ENTRY hoặc EXIT.");
  if (!data.dateTime) throw new Error("Vui lòng nhập ngày giờ.");
  if (!data.airport) throw new Error("Vui lòng nhập sân bay/cửa khẩu.");
  if (!data.destination) throw new Error("Vui lòng nhập điểm đến.");
}

function validateDomesticTravel(data) {
  validateEmployeeExists(data.employeeId);
  if (!data.fromDate || !data.toDate) throw new Error("Vui lòng nhập từ ngày và đến ngày.");
  if (new Date(data.toDate) < new Date(data.fromDate)) throw new Error("Đến ngày không được trước Từ ngày.");
  if (!data.fromLocation || !data.toLocation) throw new Error("Vui lòng nhập nơi đi và nơi đến.");
  if (!data.purpose) throw new Error("Vui lòng nhập mục đích.");
}

function addEntryExit(data, createdById) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.ENTRY_EXIT);
  if (!sheet) return respondJSON({status:"ERROR", message:"EntryExit sheet not found"});
  const recordId = "EE-" + Date.now();
  const status = "Completed";
  sheet.appendRow([recordId, data.employeeId, data.type, data.dateTime, data.airport, data.flightNo||"",
    data.ticketFile||"", data.destination||"", data.purpose||"", status, new Date(), createdById||""]);
  updateEmployeeStatus(data.employeeId, data.type === "ENTRY" ? "In Vietnam" : "Exited",
    data.type === "ENTRY" ? "Việt Nam" : data.destination || "");
  writeAudit(createdById, "ADD_ENTRY_EXIT", data.employeeId, data.type);
  return respondJSON({status:"SUCCESS", recordId});
}

function getEntryExit(employeeId) {
  return respondJSON({status:"SUCCESS", records:getEntryExitRaw(employeeId).map(r=>({
    recordId:r[0], employeeId:r[1], type:r[2], dateTime:formatDateTime(r[3]), airport:r[4],
    flightNo:r[5], ticketFile:r[6], destination:r[7], purpose:r[8], status:r[9], createdAt:formatDateTime(r[10])
  }))});
}

function addDomesticTravel(data, createdById) {
  const sheet = SPREADSHEET.getSheetByName(SHEETS.DOMESTIC);
  if (!sheet) return respondJSON({status:"ERROR", message:"DomesticTravel sheet not found"});
  const recordId = "DT-" + Date.now();
  sheet.appendRow([recordId, data.employeeId, data.fromDate, data.toDate, data.fromLocation, data.toLocation,
    data.purpose, data.transport||"", data.flightNo||"", data.hotel||"", data.hotelAddress||"", data.ticketFile||"",
    "Planned", new Date(), createdById||""]);
  autoUpdateDomesticTravelStatuses();
  writeAudit(createdById, "ADD_DOMESTIC_TRAVEL", data.employeeId, data.fromLocation + " → " + data.toLocation);
  return respondJSON({status:"SUCCESS", recordId});
}

function getDomesticTravel(employeeId) {
  return respondJSON({status:"SUCCESS", records:getDomesticRaw(employeeId).map(r=>({
    recordId:r[0], employeeId:r[1], fromDate:formatDate(r[2]), toDate:formatDate(r[3]), fromLocation:r[4],
    toLocation:r[5], purpose:r[6], transport:r[7], flightNo:r[8], hotel:r[9], hotelAddress:r[10],
    ticketFile:r[11], status:r[12]
  }))});
}

function getEmployeeTimeline(employeeId) {
  const events = [];
  getEntryExitRaw(employeeId).forEach(r => events.push({
    date: new Date(r[3]), type: r[2], title: r[2] === "ENTRY" ? "Nhập cảnh Việt Nam" : "Xuất cảnh Việt Nam",
    location: r[4], flight: r[5], purpose: r[8], status: r[9]
  }));
  getDomesticRaw(employeeId).forEach(r => events.push({
    date: new Date(r[2]), endDate: new Date(r[3]), type: "DOMESTIC", title: "Đi lại trong Việt Nam",
    location: r[4] + " → " + r[5], flight: r[8], hotel: r[9], purpose: r[6], status: r[12]
  }));
  events.sort((a,b) => b.date - a.date);
  return respondJSON({status:"SUCCESS", timeline:events.map(e=>({...e, date:formatDateTime(e.date), endDate:e.endDate?formatDateTime(e.endDate):""}))});
}

function getEntryExitRaw(employeeId) {
  const sh = SPREADSHEET.getSheetByName(SHEETS.ENTRY_EXIT); if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => String(r[1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase());
}

function getDomesticRaw(employeeId) {
  const sh = SPREADSHEET.getSheetByName(SHEETS.DOMESTIC); if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(r => String(r[1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase());
}

function autoUpdateDomesticTravelStatuses() {
  const sh = SPREADSHEET.getSheetByName(SHEETS.DOMESTIC);
  if (!sh) return;
  const data = sh.getDataRange().getValues(), now = new Date();
  for (let i = 1; i < data.length; i++) {
    const from = new Date(data[i][2]), to = new Date(data[i][3]), current = String(data[i][12] || "Planned");
    if (isNaN(from) || isNaN(to)) continue;
    let desired = current;
    if (now < from) desired = "Planned";
    else if (now >= from && now <= endOfDay(to)) desired = "In Progress";
    else if (now > endOfDay(to)) desired = "Completed";
    if (desired !== current) sh.getRange(i + 1, 13).setValue(desired);

    if (desired === "In Progress") updateEmployeeStatus(data[i][1], "Traveling", data[i][5]);
    if (desired === "Completed") {
      const latest = getLatestEntryExit(data[i][1]);
      if (!latest || latest.type === "ENTRY") updateEmployeeStatus(data[i][1], "In Vietnam", "Việt Nam");
    }
  }
}

function getLatestEntryExit(employeeId) {
  const records = getEntryExitRaw(employeeId);
  if (!records.length) return null;
  records.sort((a,b) => new Date(b[3]) - new Date(a[3]));
  return {type: records[0][2], date: new Date(records[0][3])};
}

function updateEmployeeStatus(employeeId, newStatus, newLocation) {
  const sh = SPREADSHEET.getSheetByName(SHEETS.EMPLOYEES); if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(employeeId).trim().toLowerCase()) {
      sh.getRange(i + 1, 19, 1, 2).setValues([[newStatus, newLocation]]);
      break;
    }
  }
}

function writeAudit(userId, action, employeeId, description) {
  const sh = SPREADSHEET.getSheetByName(SHEETS.AUDIT); if (!sh) return;
  sh.appendRow(["LOG-" + Date.now(), userId, action, employeeId, description, new Date()]);
}

function expiryInfo(value) {
  if (!value) return null;
  const d = new Date(value); if (isNaN(d.getTime())) return null;
  const days = Math.ceil((startOfDay(d) - startOfDay(new Date())) / 86400000);
  if (days < 0) return {status:"EXPIRED", days};
  if (days <= EXPIRING_DAYS) return {status:"EXPIRING", days};
  return {status:"VALID", days};
}

function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function endOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999); }

function formatDate(value) {
  if (!value) return "";
  const d = new Date(value); if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateTime(value) {
  if (!value) return "";
  const d = new Date(value); if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}
