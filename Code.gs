// ============================================================
// FOREIGN EMPLOYEE MANAGEMENT SYSTEM - OPTIMIZED FINAL
// Google Apps Script backend
// ============================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SESSION_TTL_SECONDS = 21600; // 6 hours
const MAX_UPLOAD_BYTES = 10 * 1024 * 1024; // 10 MB
const ROOT_DRIVE_FOLDER = 'Foreign Employee Management System';
const DEFAULT_USER_PASSWORD = '123456';
const EXPIRING_DAYS = 45;

const ALLOWED_UPLOAD_MIME = [
  'application/pdf',
  'image/jpeg',
  'image/png',
  'image/webp'
];

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

const SCHEMA = {
  Users: ['UserID','Email','PasswordHash','Role','EmployeeID','Status','CreatedAt'],
  Employees: ['EmployeeID','FullName','Nationality','DOB','Position','Department','Phone','Email','CurrentStatus','CurrentLocation','ActiveStatus','FolderId','CreatedAt','UpdatedAt'],
  Documents: ['DocID','EmployeeID','DocType','DocNo','IssueDate','ExpiryDate','Issuer','FileId','FileUrl','IsCurrent','Status','CreatedAt','CreatedBy'],
  Contracts: ['ContractID','EmployeeID','ContractNo','ContractType','StartDate','EndDate','Salary','AllowancesJson','Status','IsCurrent','FileId','FileUrl','CreatedAt','CreatedBy'],
  ContractAppendices: ['AppendixID','ContractID','EmployeeID','AppendixNo','EffectiveDate','EndDate','NewSalary','AllowancesJson','Content','FileId','FileUrl','CreatedAt','CreatedBy'],
  EntryExit: ['LogID','EmployeeID','Type','EventDate','PortName','Note','CreatedAt','CreatedBy'],
  DomesticTravel: ['TravelID','EmployeeID','FromLocation','ToLocation','StartDate','EndDate','Status','TicketFileId','TicketUrl','CreatedAt','CreatedBy'],
  AuditLog: ['AuditID','UserID','Action','Module','TargetID','OldValue','NewValue','Timestamp','IPAddress'],
  Config: ['Key','Value','UpdatedAt']
};

let _SS = null;
const _SHEET_CACHE = {};
const _HEADER_CACHE = {};

// ============================================================
// BASIC HELPERS
// ============================================================
function getSS() {
  if (!_SS) _SS = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _SS;
}

function getSheet(name) {
  if (_SHEET_CACHE[name]) return _SHEET_CACHE[name];
  const sheet = getSS().getSheetByName(name);
  if (!sheet) throw new Error('Sheet chưa tồn tại: ' + name + '. Hãy chạy setupSystem().');
  _SHEET_CACHE[name] = sheet;
  return sheet;
}

function getHeaderMap(sheetName) {
  if (_HEADER_CACHE[sheetName]) return _HEADER_CACHE[sheetName];
  const sheet = getSheet(sheetName);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  _HEADER_CACHE[sheetName] = map;
  return map;
}

function invalidateCaches() {
  Object.keys(_HEADER_CACHE).forEach(k => delete _HEADER_CACHE[k]);
  Object.keys(_SHEET_CACHE).forEach(k => delete _SHEET_CACHE[k]);
}

function readData(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function isTrue(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim() === '1' || String(v).trim().toUpperCase() === 'YES';
}

function isActiveValue(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  return s !== 'INACTIVE' && s !== 'FALSE' && s !== '0' && s !== 'NO';
}

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function uid(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
}

function safeText(v) {
  return String(v == null ? '' : v).trim();
}

function formatDate(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = safeText(d);
  if (!s) return '';
  let m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (m) return m[1] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[3]).padStart(2,'0');
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? s : Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function parseDateSafe(v) {
  if (!v) return new Date(0);
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? new Date(0) : v;
  const s = safeText(v);
  if (!s) return new Date(0);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const y=+m[1], mo=+m[2]-1, d=+m[3], h=+(m[4]||0), mi=+(m[5]||0), sec=+(m[6]||0);
    return new Date(y,mo,d,h,mi,sec);
  }
  const dt = new Date(s);
  return isNaN(dt.getTime()) ? new Date(0) : dt;
}

function formatDateTime(v) {
  const d = parseDateSafe(v);
  if (d.getTime() === 0) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function jsonParse(v, fallback) {
  try { return JSON.parse(String(v || '')); } catch (e) { return fallback; }
}

function normalizeRole(role) {
  return safeText(role).toUpperCase();
}

function isHR(req) {
  return req && [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR].includes(normalizeRole(req.role));
}

function isManagement(req) {
  return req && [ROLES.ADMIN, ROLES.HR_ADMIN, ROLES.HR, ROLES.DIRECTOR].includes(normalizeRole(req.role));
}

function isAccountAdmin(req) {
  return req && [ROLES.ADMIN, ROLES.HR_ADMIN].includes(normalizeRole(req.role));
}

function canAccessEmployee(req, employeeId) {
  if (!req) return false;
  if (isManagement(req)) return true;
  return safeText(req.employeeId) === safeText(employeeId);
}

// ============================================================
// SETUP
// ============================================================
function setupSystem() {
  const ss = getSS();
  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const expected = SCHEMA[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,expected.length).setValues([expected]);
    } else {
      const current = sh.getRange(1,1,Math.max(sh.getLastColumn(), expected.length),1).getValues().map(r=>safeText(r[0]));
      const row = sh.getRange(1,1,1,Math.max(sh.getLastColumn(), expected.length)).getValues()[0];
      const hasHeader = row.some(Boolean);
      if (!hasHeader) sh.getRange(1,1,1,expected.length).setValues([expected]);
    }
    sh.getRange(1,1,1,Math.max(sh.getLastColumn(), expected.length)).setFontWeight('bold').setBackground('#f3f4f6');
  });

  setupDriveRoot();

  const users = getSheet(TABLES.USERS);
  if (users.getLastRow() < 2) {
    users.getRange(2,1,1,7).setValues([[uid('USR'),'admin@company.com',hashPassword('admin123'),ROLES.ADMIN,'EMP-ADMIN','ACTIVE',new Date()]]);
  }
  invalidateCaches();
  return {status:'SUCCESS', message:'System setup completed.'};
}

function setupDriveRoot() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DRIVE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id).getId(); } catch(e) {}
  }
  const folders = DriveApp.getFoldersByName(ROOT_DRIVE_FOLDER);
  const root = folders.hasNext() ? folders.next() : DriveApp.createFolder(ROOT_DRIVE_FOLDER);
  props.setProperty('DRIVE_FOLDER_ID', root.getId());
  return root.getId();
}

function createEmployeeDriveFolder(employeeId, fullName) {
  const rootId = setupDriveRoot();
  const root = DriveApp.getFolderById(rootId);
  const id = safeText(employeeId);
  const name = safeText(fullName || 'Unknown').replace(/[\\/:*?"<>|]/g,'_');
  const targetName = id + ' - ' + name;
  const folders = root.getFoldersByName(targetName);
  const empFolder = folders.hasNext() ? folders.next() : root.createFolder(targetName);
  ['Passport','Visa','TRC','Work Permit','Contracts'].forEach(n => {
    if (!empFolder.getFoldersByName(n).hasNext()) empFolder.createFolder(n);
  });
  return empFolder.getId();
}

function writeAuditLog(userId, action, module, targetId, oldValue, newValue) {
  try {
    const sh = getSheet(TABLES.AUDIT);
    sh.getRange(sh.getLastRow()+1,1,1,9).setValues([[
      uid('AUD'), safeText(userId || 'SYSTEM'), safeText(action), safeText(module), safeText(targetId),
      oldValue == null ? '' : String(oldValue), newValue == null ? '' : String(newValue), new Date(), ''
    ]]);
  } catch(e) { console.error('Audit:', e); }
}

// ============================================================
// AUTHENTICATION / SESSION
// ============================================================
function hashPassword(pass) {
  const salt = 'HR_GLOBAL_SALT_2026';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pass || '') + salt, Utilities.Charset.UTF_8);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0')).join('');
}

function getUserByLogin(login) {
  const needle = safeText(login).toLowerCase();
  const rows = readData(TABLES.USERS);
  for (let i=0;i<rows.length;i++) {
    const r=rows[i];
    if (safeText(r[1]).toLowerCase()===needle || safeText(r[4]).toLowerCase()===needle || safeText(r[0]).toLowerCase()===needle) {
      return {row:r,rowIndex:i+2,sheet:getSheet(TABLES.USERS)};
    }
  }
  return null;
}

function findUserById(userId) {
  const rows=readData(TABLES.USERS);
  for(let i=0;i<rows.length;i++) if(safeText(rows[i][0])===safeText(userId)) return {row:rows[i],rowIndex:i+2,sheet:getSheet(TABLES.USERS)};
  return null;
}

function getUserPublic(userId) {
  const f=findUserById(userId);
  if(!f) return null;
  return {userId:f.row[0],email:f.row[1],role:f.row[3],employeeId:f.row[4],status:safeText(f.row[5]).toUpperCase()};
}

function handleLogin(username,password) {
  username=safeText(username); password=String(password||'');
  if(!username||!password) return {status:'ERROR',message:'Vui lòng nhập tài khoản và mật khẩu.'};
  const f=getUserByLogin(username);
  if(!f) return {status:'ERROR',message:'Tài khoản hoặc mật khẩu không chính xác!'};
  const stored=safeText(f.row[2]);
  const hashed=hashPassword(password);
  const legacy=stored===password;
  if(stored!==hashed && !legacy) return {status:'ERROR',message:'Tài khoản hoặc mật khẩu không chính xác!'};
  if(safeText(f.row[5]).toUpperCase()!=='ACTIVE') return {status:'ERROR',message:'Tài khoản của bạn đã bị khóa.'};

  const empId=safeText(f.row[4]);
  const emp=empId ? findEmployee(empId) : null;
  if(emp && safeText(emp.row[10]).toUpperCase()==='INACTIVE') return {status:'ERROR',message:'Hồ sơ nhân viên đã dừng hoạt động.'};
  if(legacy) f.sheet.getRange(f.rowIndex,3).setValue(hashed);

  const token='TK-'+Utilities.getUuid();
  CacheService.getScriptCache().put('SESSION_'+token,JSON.stringify({userId:f.row[0],createdAt:Date.now()}),SESSION_TTL_SECONDS);
  const name=emp ? safeText(emp.row[1]) : safeText(f.row[1]);
  writeAuditLog(f.row[0],'LOGIN','Users',f.row[0],null,'Login successful');
  return {status:'SUCCESS',token,user:{id:f.row[0],userId:f.row[0],email:f.row[1],name,role:f.row[3],employeeId:empId}};
}

function validateSession(token) {
  token=safeText(token); if(!token) return null;
  const raw=CacheService.getScriptCache().get('SESSION_'+token);
  if(!raw) return null;
  const s=jsonParse(raw,null); if(!s) return null;
  if(Date.now()-Number(s.createdAt||0)>SESSION_TTL_SECONDS*1000) { CacheService.getScriptCache().remove('SESSION_'+token); return null; }
  const user=getUserPublic(s.userId);
  if(!user || user.status!=='ACTIVE') return null;
  const emp=user.employeeId ? findEmployee(user.employeeId) : null;
  if(emp && safeText(emp.row[10]).toUpperCase()==='INACTIVE') return null;
  return {userId:user.userId,email:user.email,role:user.role,employeeId:user.employeeId};
}

function invalidateUserSessions(userId) {
  // CacheService cannot enumerate keys. Account status is checked on every request,
  // so locked users are rejected immediately. Existing token will naturally expire.
}

function handleLogout(token) {
  if(token) CacheService.getScriptCache().remove('SESSION_'+token);
  return {status:'SUCCESS',message:'Đã đăng xuất thành công.'};
}

// ============================================================
// API ROUTER
// ============================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Foreign Employee Management System').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body={};
  try { body=e&&e.postData&&e.postData.contents?JSON.parse(e.postData.contents):(e&&e.parameter)||{}; }
  catch(err){ return respondJSON({status:'ERROR',message:'Dữ liệu request không hợp lệ.'}); }
  const action=safeText(body.action).toUpperCase();
  const data=body.data||body;
  const params=body.params||{};
  if(action==='LOGIN') return respondJSON(handleLogin(body.username||data.username,body.password||data.password));
  const req=validateSession(body.token);
  if(!req) return respondJSON({status:'ERROR',message:'Phiên đăng nhập hết hạn hoặc không hợp lệ.'});
  return respondJSON(dispatchApiAction(action,body.token,req,data,params));
}

function apiCall(action,token,data,params) {
  try {
    action=safeText(action).toUpperCase(); data=data||{}; params=params||{};
    if(action==='LOGIN') return handleLogin(data.username,data.password);
    const req=validateSession(token);
    if(!req) return {status:'ERROR',message:'Phiên đăng nhập hết hạn hoặc không hợp lệ.'};
    return dispatchApiAction(action,token,req,data,params);
  } catch(e) { console.error(e); return {status:'ERROR',message:e.message||String(e)}; }
}

function dispatchApiAction(action,token,req,data,params) {
  const employeeId=safeText(params.employeeId||data.employeeId||(data.data&&data.data.employeeId));
  switch(action) {
    case 'LOGOUT': return handleLogout(token);
    case 'GET_DASHBOARD_DATA': return getDashboardData(req);
    case 'GET_PROFILE': return getProfileData(employeeId,req);
    case 'SAVE_PROFILE': return handleSaveProfile(data,req);
    case 'SAVE_DOCUMENT': return isHR(req)?saveDocument(data,req):{status:'ERROR',message:'Không có quyền cập nhật giấy tờ.'};
    case 'UPLOAD_FILE': return uploadEmployeeFile(data,req);
    case 'ADD_ENTRY_EXIT': return handleAddEntryExit(data,req);
    case 'ADD_ENTRY_EXIT_TRIP': return handleAddEntryExitTrip(data,req);
    case 'UPDATE_ENTRY_EXIT': return handleUpdateEntryExit(data,req);
    case 'UPDATE_ENTRY_EXIT_TRIP': return handleUpdateEntryExitTrip(data,req);
    case 'DELETE_ENTRY_EXIT': return handleDeleteEntryExit(data,req);
    case 'GET_ENTRY_EXIT': return handleGetEntryExit(employeeId,req);
    case 'ADD_DOMESTIC_TRAVEL': return handleAddDomesticTravel(data,req);
    case 'UPDATE_DOMESTIC_TRAVEL': return handleUpdateDomesticTravel(data,req);
    case 'CANCEL_DOMESTIC_TRAVEL': return handleCancelDomesticTravel(data,req);
    case 'APPROVE_DOMESTIC_TRAVEL': return handleApproveDomesticTravel(data,req);
    case 'GET_DOMESTIC_TRAVEL': return handleGetDomesticTravel(employeeId,req);
    case 'GET_TIMELINE': return handleGetTimeline(employeeId,req);
    case 'DELETE_EMPLOYEE': return handleDeleteEmployee(employeeId,req);
    case 'IMPORT_EMPLOYEES_EXCEL': return handleImportEmployeesExcel(data,req);
    case 'SYNC_MASTER_EXCEL': return handleSyncMasterExcel(data,req);
    case 'CHANGE_PASSWORD': return handleChangePassword(data,req);
    case 'GET_ACCOUNT_USERS': return getAccountUsers(req);
    case 'ADMIN_RESET_PASSWORD': return adminResetPassword(data,req);
    case 'SET_ACCOUNT_STATUS': return setAccountStatus(data,req);
    default: return {status:'ERROR',message:'Hành động không hợp lệ: '+action};
  }
}

// ============================================================
// EMPLOYEE / PROFILE
// ============================================================
function findEmployee(employeeId) {
  const rows=readData(TABLES.EMPLOYEES);
  for(let i=0;i<rows.length;i++) if(safeText(rows[i][0])===safeText(employeeId)) return {row:rows[i],rowIndex:i+2,sheet:getSheet(TABLES.EMPLOYEES)};
  return null;
}

function handleSaveProfile(data,req) {
  data=data||{};
  const employeeId=safeText(data.employeeId), fullName=safeText(data.fullName);
  if(!employeeId) return {status:'ERROR',message:'Thiếu Mã nhân viên.'};
  if(!canAccessEmployee(req,employeeId)) return {status:'ERROR',message:'Bạn không có quyền cập nhật hồ sơ này.'};
  const manager=isHR(req), self=safeText(req.employeeId)===employeeId;
  if(manager && !fullName) return {status:'ERROR',message:'Họ và tên là bắt buộc.'};
  if(!manager && !self) return {status:'ERROR',message:'Bạn không có quyền cập nhật hồ sơ này.'};

  const found=findEmployee(employeeId), now=new Date();
  if(found) {
    const old=found.row.slice();
    const r=old.slice(); while(r.length<14) r.push('');
    if(manager) {
      r[1]=fullName||r[1]; r[2]=safeText(data.nationality)||r[2]; r[3]=formatDate(data.dob)||r[3];
      r[4]=safeText(data.position)||r[4]; r[5]=safeText(data.department)||r[5];
      if(data.activeStatus!==undefined) r[10]=safeText(data.activeStatus).toUpperCase();
    }
    if(data.phone!==undefined) r[6]=safeText(data.phone);
    if(data.email!==undefined) r[7]=safeText(data.email);
    r[13]=now;
    found.sheet.getRange(found.rowIndex,1,1,14).setValues([r.slice(0,14)]);
    if(manager) saveProfileDocumentsAndContract(data,req);
    writeAuditLog(req.userId,'UPDATE','Employees',employeeId,JSON.stringify(old),JSON.stringify(r));
    return {status:'SUCCESS',message:'Đã lưu thông tin hồ sơ thành công.'};
  }

  if(!manager) return {status:'ERROR',message:'Chỉ HR/HR ADMIN/ADMIN được thêm nhân viên.'};
  if(findEmployee(employeeId)) return {status:'ERROR',message:'Mã nhân viên đã tồn tại.'};
  const email=safeText(data.email).toLowerCase();
  if(email && getUserByLogin(email)) return {status:'ERROR',message:'Email tài khoản đã tồn tại.'};
  const folderId=createEmployeeDriveFolder(employeeId,fullName);
  const row=[employeeId,fullName,safeText(data.nationality),formatDate(data.dob),safeText(data.position),safeText(data.department),safeText(data.phone),email,'Exited','Overseas','ACTIVE',folderId,now,now];
  const sh=getSheet(TABLES.EMPLOYEES); sh.getRange(sh.getLastRow()+1,1,1,14).setValues([row]);
  const us=getSheet(TABLES.USERS);
  const login=email||employeeId;
  us.getRange(us.getLastRow()+1,1,1,7).setValues([[uid('USR'),login,hashPassword(DEFAULT_USER_PASSWORD),ROLES.USER,employeeId,'ACTIVE',now]]);
  saveProfileDocumentsAndContract(data,req);
  writeAuditLog(req.userId,'CREATE','Employees',employeeId,null,JSON.stringify(row));
  return {status:'SUCCESS',message:'Đã tạo hồ sơ nhân viên thành công.'};
}

function saveProfileDocumentsAndContract(data,req) {
  const employeeId=safeText(data.employeeId);
  if(data.passportNo || data.passportExpiry || data.passportFileId || data.passportImg) saveDocument({employeeId,docType:'PASSPORT',docNo:data.passportNo,expiryDate:data.passportExpiry,fileId:data.passportFileId,fileUrl:data.passportImg},req);
  if(data.trcNo || data.trcExpiry || data.trcFileId || data.trcImg) saveDocument({employeeId,docType:'TRC',docNo:data.trcNo,expiryDate:data.trcExpiry,fileId:data.trcFileId,fileUrl:data.trcImg},req);
  if(data.wpNo || data.wpExpiry || data.wpFileId || data.wpImg) saveDocument({employeeId,docType:'WORK_PERMIT',docNo:data.wpNo,expiryDate:data.wpExpiry,fileId:data.wpFileId,fileUrl:data.wpImg},req);
  if(safeText(data.contractNo)) saveContractAndAppendix(data,req);
}

function saveDocument(data,req) {
  if(!canAccessEmployee(req,data.employeeId)||!isHR(req)) return {status:'ERROR',message:'Không có quyền cập nhật giấy tờ.'};
  const sh=getSheet(TABLES.DOCUMENTS), rows=readData(TABLES.DOCUMENTS), empId=safeText(data.employeeId), type=safeText(data.docType).toUpperCase();
  rows.forEach((r,i)=>{ if(safeText(r[1])===empId && safeText(r[2]).toUpperCase()===type && isTrue(r[9])) sh.getRange(i+2,10).setValue(false); });
  const row=[uid('DOC'),empId,type,safeText(data.docNo),formatDate(data.issueDate),formatDate(data.expiryDate),safeText(data.issuer),safeText(data.fileId),safeText(data.fileUrl),true,'VALID',new Date(),req.userId];
  sh.getRange(sh.getLastRow()+1,1,1,13).setValues([row]);
  writeAuditLog(req.userId,'ADD_DOCUMENT','Documents',row[0],null,JSON.stringify(row));
  return {status:'SUCCESS',message:'Đã cập nhật giấy tờ.',docId:row[0]};
}

function saveContractAndAppendix(data,req) {
  const sh=getSheet(TABLES.CONTRACTS), rows=readData(TABLES.CONTRACTS), empId=safeText(data.employeeId), contractNo=safeText(data.contractNo);
  let current=null;
  rows.forEach((r,i)=>{ if(safeText(r[1])===empId && isTrue(r[9])) { if(safeText(r[2])===contractNo) current={row:r,rowIndex:i+2}; else sh.getRange(i+2,10).setValue(false); } });
  const now=new Date();
  if(!current) {
    const id=uid('CTR');
    const row=[id,empId,contractNo,safeText(data.contractType||'DEFINITE'),formatDate(data.contractStartDate),formatDate(data.contractExpiry),Number(data.salary)||0,JSON.stringify({allowance:Number(data.allowance)||0}),'VALID',true,safeText(data.contractFileId),safeText(data.contractImg),now,req.userId];
    sh.getRange(sh.getLastRow()+1,1,1,14).setValues([row]); current={row,rowIndex:sh.getLastRow()};
    writeAuditLog(req.userId,'ADD_CONTRACT','Contracts',id,null,JSON.stringify(row));
  } else {
    const old=current.row.slice(), r=old.slice();
    r[4]=formatDate(data.contractStartDate)||r[4]; r[5]=formatDate(data.contractExpiry)||r[5]; r[6]=Number(data.salary)||0; r[7]=JSON.stringify({allowance:Number(data.allowance)||0});
    if(data.contractFileId) r[10]=data.contractFileId; if(data.contractImg) r[11]=data.contractImg; r[8]='VALID'; r[9]=true;
    sh.getRange(current.rowIndex,1,1,14).setValues([r.slice(0,14)]); current.row=r;
    writeAuditLog(req.userId,'UPDATE_CONTRACT','Contracts',r[0],JSON.stringify(old),JSON.stringify(r));
  }
  if(safeText(data.appendixNo)) {
    const ah=getSheet(TABLES.APPENDICES), ar=readData(TABLES.APPENDICES), contractId=current.row[0];
    const exists=ar.some(r=>safeText(r[1])===contractId&&safeText(r[3])===safeText(data.appendixNo));
    if(!exists) {
      const row=[uid('APP'),contractId,empId,safeText(data.appendixNo),formatDate(data.contractStartDate),formatDate(data.contractExpiry),Number(data.salary)||0,JSON.stringify({allowance:Number(data.allowance)||0}),'Cập nhật theo hồ sơ',safeText(data.appendixFileId),safeText(data.contractImg),new Date(),req.userId];
      ah.getRange(ah.getLastRow()+1,1,1,13).setValues([row]);
      writeAuditLog(req.userId,'ADD_APPENDIX','ContractAppendices',row[0],null,JSON.stringify(row));
    }
  }
}

// ============================================================
// DRIVE UPLOAD
// ============================================================
function uploadEmployeeFile(data,req) {
  try {
    data=data||{};
    const employeeId=safeText(data.employeeId);
    if(!employeeId||!canAccessEmployee(req,employeeId)) return {status:'ERROR',message:'Bạn không có quyền upload file cho nhân viên này.'};
    const emp=findEmployee(employeeId); if(!emp) return {status:'ERROR',message:'Không tìm thấy nhân viên.'};
    let folderId=safeText(emp.row[11]);
    if(!folderId) { folderId=createEmployeeDriveFolder(employeeId,emp.row[1]); emp.sheet.getRange(emp.rowIndex,12).setValue(folderId); }
    const parent=DriveApp.getFolderById(folderId);
    const sub=safeText(data.subFolderType||'Misc').replace(/[\\/:*?"<>|]/g,'')||'Misc';
    const fs=parent.getFoldersByName(sub); const folder=fs.hasNext()?fs.next():parent.createFolder(sub);
    const base64=String(data.base64Data||'').split(',').pop();
    if(!base64) return {status:'ERROR',message:'Không có dữ liệu file.'};
    const bytes=Utilities.base64Decode(base64);
    if(bytes.length>MAX_UPLOAD_BYTES) return {status:'ERROR',message:'File vượt quá giới hạn 10 MB.'};
    const mime=safeText(data.mimeType).toLowerCase();
    if(ALLOWED_UPLOAD_MIME.indexOf(mime)<0) return {status:'ERROR',message:'Chỉ hỗ trợ PDF, JPG, PNG, WEBP.'};
    const fileName=safeText(data.fileName||'upload').replace(/[\\/:*?"<>|]/g,'_')||'upload';
    const file=folder.createFile(Utilities.newBlob(bytes,mime,fileName));
    file.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);
    return {status:'SUCCESS',fileId:file.getId(),fileUrl:file.getUrl(),fileName:file.getName(),mimeType:file.getMimeType()};
  } catch(e) { return {status:'ERROR',message:e.message||String(e)}; }
}

function getDriveFileInfo(fileId,url) {
  let id=safeText(fileId), raw=safeText(url), mime='', name='';
  if(!id&&raw) { const m=raw.match(/\/file\/d\/([\w-]+)/)||raw.match(/\/d\/([\w-]+)/)||raw.match(/[?&]id=([\w-]+)/); if(m) id=m[1]; }
  if(id) { try { const f=DriveApp.getFileById(id); mime=f.getMimeType(); name=f.getName(); raw='https://drive.google.com/file/d/'+id+'/view'; } catch(e){} }
  return {fileId:id,fileUrl:raw,previewUrl:id?'https://drive.google.com/file/d/'+id+'/preview':raw,mimeType:mime,fileName:name};
}

// ============================================================
// ENTRY / EXIT - ONE FORM FOR ROUND TRIP
// EntryExit sheet remains compatible with the existing 8-column structure.
// Extra fields are stored in Note as JSON.
// ============================================================
function makeEENote(data) {
  return JSON.stringify({plannedDateTime:data.plannedDateTime||data.dateTime||'',actualDateTime:data.actualDateTime||'',flightNo:data.flightNo||'',destination:data.destination||'',purpose:data.purpose||'',tripId:data.tripId||''});
}

function parseEENote(row) {
  const parsed=jsonParse(row[5],null);
  if(parsed && typeof parsed==='object') return parsed;
  return {plannedDateTime:formatDateTime(row[3]),actualDateTime:'',flightNo:safeText(row[5]),destination:'',purpose:'',tripId:''};
}

function validateEE(data) {
  const type=safeText(data.type).toUpperCase();
  const dt=data.plannedDateTime||data.dateTime||data.date;
  if(!['ENTRY','EXIT'].includes(type)||!dt||!safeText(data.airport)) return 'Thiếu hoặc sai thông tin nhập/xuất cảnh.';
  if(parseDateSafe(dt).getTime()===0) return 'Ngày giờ không hợp lệ.';
  return '';
}

function appendEERow(sh,data,req) {
  const id=uid('EE'), dt=data.plannedDateTime||data.dateTime||data.date, now=new Date();
  const row=[id,safeText(data.employeeId),safeText(data.type).toUpperCase(),dt,safeText(data.airport),makeEENote(data),now,req.userId];
  sh.getRange(sh.getLastRow()+1,1,1,8).setValues([row]);
  return row;
}

function handleAddEntryExit(data,req) {
  data=data||{}; data.type=safeText(data.type).toUpperCase(); data.employeeId=safeText(data.employeeId);
  const err=validateEE(data); if(err) return {status:'ERROR',message:err};
  if(!canAccessEmployee(req,data.employeeId)) return {status:'ERROR',message:'Bạn không có quyền thực hiện.'};
  const row=appendEERow(getSheet(TABLES.ENTRY_EXIT),data,req);
  calculateEmployeeLocationStatus(data.employeeId);
  writeAuditLog(req.userId,'ADD_ENTRY_EXIT',TABLES.ENTRY_EXIT,row[0],null,JSON.stringify(data));
  return {status:'SUCCESS',message:'Đã lưu khai báo nhập/xuất cảnh thành công.',entryExitId:row[0]};
}

function handleAddEntryExitTrip(data,req) {
  data=data||{}; const empId=safeText(data.employeeId), tripId=safeText(data.tripId)||uid('TRIP');
  if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Bạn không có quyền thực hiện.'};
  const entry=Object.assign({},data.entry||{},{employeeId:empId,type:'ENTRY',tripId});
  const exit=Object.assign({},data.exit||{},{employeeId:empId,type:'EXIT',tripId});
  const e1=validateEE(entry), e2=validateEE(exit);
  if(e1||e2) return {status:'ERROR',message:e1||e2};
  const sh=getSheet(TABLES.ENTRY_EXIT), a=appendEERow(sh,entry,req), b=appendEERow(sh,exit,req);
  calculateEmployeeLocationStatus(empId);
  writeAuditLog(req.userId,'ADD_ENTRY_EXIT_TRIP',TABLES.ENTRY_EXIT,tripId,null,JSON.stringify({entry,exit}));
  return {status:'SUCCESS',message:'Đã lưu đầy đủ chuyến khứ hồi (Nhập + Xuất).',tripId,entryExitIds:[a[0],b[0]]};
}

function handleUpdateEntryExit(data,req) {
  data=data||{}; const id=safeText(data.entryExitId||data.logId), empId=safeText(data.employeeId);
  if(!id||!empId) return {status:'ERROR',message:'Thiếu mã khai báo.'};
  if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Bạn không có quyền thực hiện.'};
  const err=validateEE(data); if(err) return {status:'ERROR',message:err};
  const sh=getSheet(TABLES.ENTRY_EXIT), rows=readData(TABLES.ENTRY_EXIT);
  for(let i=0;i<rows.length;i++) if(safeText(rows[i][0])===id) {
    if(safeText(rows[i][1])!==empId) return {status:'ERROR',message:'Khai báo không thuộc nhân viên này.'};
    const old=rows[i].slice(), r=old.slice(); while(r.length<8) r.push('');
    r[2]=safeText(data.type).toUpperCase(); r[3]=data.plannedDateTime||data.dateTime||data.date; r[4]=safeText(data.airport); r[5]=makeEENote(data); r[7]=old[7]||req.userId;
    sh.getRange(i+2,1,1,8).setValues([r.slice(0,8)]); calculateEmployeeLocationStatus(empId);
    writeAuditLog(req.userId,'UPDATE_ENTRY_EXIT',TABLES.ENTRY_EXIT,id,JSON.stringify(old),JSON.stringify(r));
    return {status:'SUCCESS',message:'Đã cập nhật khai báo nhập/xuất cảnh.'};
  }
  return {status:'ERROR',message:'Không tìm thấy khai báo nhập/xuất cảnh.'};
}

function handleUpdateEntryExitTrip(data,req) {
  data=data||{}; const empId=safeText(data.employeeId), tripId=safeText(data.tripId);
  if(!empId||!tripId) return {status:'ERROR',message:'Thiếu EmployeeID hoặc Trip ID.'};
  if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Bạn không có quyền thực hiện.'};
  const entry=Object.assign({},data.entry||{},{employeeId:empId,type:'ENTRY',tripId}), exit=Object.assign({},data.exit||{},{employeeId:empId,type:'EXIT',tripId});
  const e1=validateEE(entry), e2=validateEE(exit); if(e1||e2) return {status:'ERROR',message:e1||e2};
  const sh=getSheet(TABLES.ENTRY_EXIT), rows=readData(TABLES.ENTRY_EXIT), matches=[];
  rows.forEach((r,i)=>{const n=parseEENote(r); if(safeText(r[1])===empId&&safeText(n.tripId)===tripId) matches.push({r,i});});
  if(matches.length<2) return handleAddEntryExitTrip(data,req);
  const oldValues=[];
  const byType={ENTRY:entry,EXIT:exit};
  matches.forEach(x=>{const old=x.r.slice(); oldValues.push(old); const d=byType[safeText(old[2]).toUpperCase()]; if(d){const r=old.slice(); r[2]=safeText(d.type).toUpperCase(); r[3]=d.plannedDateTime; r[4]=safeText(d.airport); r[5]=makeEENote(d); sh.getRange(x.i+2,1,1,8).setValues([r.slice(0,8)]);}});
  calculateEmployeeLocationStatus(empId); writeAuditLog(req.userId,'UPDATE_ENTRY_EXIT_TRIP',TABLES.ENTRY_EXIT,tripId,JSON.stringify(oldValues),JSON.stringify({entry,exit}));
  return {status:'SUCCESS',message:'Đã cập nhật chuyến khứ hồi.'};
}

function handleDeleteEntryExit(data,req) {
  const id=safeText(data&&data.entryExitId||data&&data.logId); if(!id) return {status:'ERROR',message:'Thiếu mã khai báo.'};
  const sh=getSheet(TABLES.ENTRY_EXIT), rows=readData(TABLES.ENTRY_EXIT);
  const target=rows.find(r=>safeText(r[0])===id);
  if(!target) return {status:'ERROR',message:'Không tìm thấy khai báo nhập/xuất cảnh.'};
  const empId=safeText(target[1]); if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Bạn không có quyền thực hiện.'};
  const note=parseEENote(target), tripId=safeText(note.tripId);
  const indexes=[];
  rows.forEach((r,i)=>{const n=parseEENote(r);if(safeText(r[1])===empId && tripId && safeText(n.tripId)===tripId) indexes.push(i+2);});
  if(!indexes.length) indexes.push(rows.findIndex(r=>safeText(r[0])===id)+2);
  const oldRows=indexes.map(rowNo=>sh.getRange(rowNo,1,1,8).getValues()[0]);
  indexes.sort((a,b)=>b-a).forEach(rowNo=>sh.deleteRow(rowNo));
  calculateEmployeeLocationStatus(empId);
  writeAuditLog(req.userId,'DELETE_ENTRY_EXIT',TABLES.ENTRY_EXIT,tripId||id,JSON.stringify(oldRows),'DELETED');
  return {status:'SUCCESS',message:indexes.length>1?'Đã xóa toàn bộ chuyến khứ hồi.':'Đã xóa khai báo nhập/xuất cảnh.'};
}

function handleGetEntryExit(empId,req) {
  if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Không có quyền xem thông tin.'};
  const rows=readData(TABLES.ENTRY_EXIT).filter(r=>safeText(r[1])===safeText(empId));
  const records=rows.map(r=>{const n=parseEENote(r); return {entryExitId:r[0],tripId:n.tripId||r[0],type:safeText(r[2]).toUpperCase(),plannedDateTime:n.plannedDateTime?formatDateTime(n.plannedDateTime):formatDateTime(r[3]),actualDateTime:n.actualDateTime?formatDateTime(n.actualDateTime):'',dateTime:formatDateTime(r[3]),airport:safeText(r[4]),flightNo:safeText(n.flightNo),destination:safeText(n.destination),purpose:safeText(n.purpose)};}).sort((a,b)=>parseDateSafe(b.plannedDateTime)-parseDateSafe(a.plannedDateTime));
  return {status:'SUCCESS',records};
}

function calculateEmployeeLocationStatus(employeeId) {
  const emp=findEmployee(employeeId); if(!emp) return {status:'Exited',location:'Overseas'};
  const rows=readData(TABLES.ENTRY_EXIT).filter(r=>safeText(r[1])===safeText(employeeId));
  const now=new Date();
  let latest=null;
  rows.forEach(r=>{
    const note=parseEENote(r);
    const candidate=parseDateSafe(note.actualDateTime||note.plannedDateTime||r[3]);
    if(candidate.getTime()===0 || candidate>now)return;
    if(!latest || candidate>latest.date) latest={row:r,date:candidate,note};
  });
  let status='Exited', location='Overseas';
  if(latest) {
    const type=safeText(latest.row[2]).toUpperCase();
    if(type==='ENTRY') { status='In Vietnam'; location=safeText(latest.row[4])||'Vietnam'; }
  }
  const travels=readData(TABLES.TRAVEL);
  const active=travels.find(r=>safeText(r[1])===safeText(employeeId)&&safeText(r[6]).toUpperCase()==='APPROVED'&&now>=parseDateSafe(r[4])&&now<=parseDateSafe(r[5]));
  if(active&&status==='In Vietnam'){status='Traveling';location=safeText(active[3]);}
  updateEmployeeStatusRecord(employeeId,status,location); return {status,location};
}

function updateEmployeeStatusRecord(employeeId,status,location) {
  const f=findEmployee(employeeId); if(!f)return;
  if(safeText(f.row[8])!==status) f.sheet.getRange(f.rowIndex,9).setValue(status);
  if(safeText(f.row[9])!==location) f.sheet.getRange(f.rowIndex,10).setValue(location);
}

// ============================================================
// DOMESTIC TRAVEL - existing 11-column structure
// Purpose is stored in TicketUrl only when no ticket exists? NO:
// To preserve the existing 11 columns, purpose is stored as JSON in TicketUrl.
// Existing plain TicketUrl values are still returned as URLs.
// ============================================================
function travelPurposeFromRow(r) {
  const raw=safeText(r[8]); const obj=jsonParse(raw,null);
  return obj&&typeof obj==='object'&&obj.__purpose!==undefined?safeText(obj.__purpose):'';
}
function travelTicketUrlFromRow(r) {
  const raw=safeText(r[8]); const obj=jsonParse(raw,null);
  return obj&&typeof obj==='object'&&obj.__ticketUrl!==undefined?safeText(obj.__ticketUrl):raw;
}
function makeTravelTicketUrl(ticketUrl,purpose) {
  if(purpose && !ticketUrl) return JSON.stringify({__purpose:purpose});
  if(purpose) return JSON.stringify({__ticketUrl:ticketUrl||'',__purpose:purpose});
  return ticketUrl||'';
}

function handleAddDomesticTravel(d,req) {
  d=d||{}; const empId=safeText(d.employeeId), from=safeText(d.fromLocation), to=safeText(d.toLocation), fromDate=formatDate(d.fromDate), toDate=formatDate(d.toDate);
  if(!empId||!fromDate||!toDate||!from||!to) return {status:'ERROR',message:'Thiếu thông tin lịch trình.'};
  const start=parseDateSafe(fromDate), end=parseDateSafe(toDate); if(start.getTime()===0||end.getTime()===0||end<start) return {status:'ERROR',message:'Khoảng ngày đi lại không hợp lệ.'};
  if(!canAccessEmployee(req,empId)) return {status:'ERROR',message:'Không có quyền thực hiện.'};
  const approved=isHR(req)?'APPROVED':'PENDING', sh=getSheet(TABLES.TRAVEL), id=uid('TRV'), row=[id,empId,from,to,fromDate,toDate,approved,safeText(d.ticketFileId),makeTravelTicketUrl(safeText(d.ticketUrl),safeText(d.purpose)),new Date(),req.userId];
  sh.getRange(sh.getLastRow()+1,1,1,11).setValues([row]); calculateEmployeeLocationStatus(empId); writeAuditLog(req.userId,'ADD_DOMESTIC_TRAVEL',TABLES.TRAVEL,id,null,JSON.stringify(row));
  return {status:'SUCCESS',message:approved==='APPROVED'?'Đã duyệt lịch trình công tác.':'Đã gửi yêu cầu công tác, chờ duyệt.',travelId:id};
}

function handleUpdateDomesticTravel(d,req) {
  d=d||{}; const id=safeText(d.travelId), empId=safeText(d.employeeId); if(!id||!empId)return {status:'ERROR',message:'Thiếu mã lịch trình.'};
  if(!canAccessEmployee(req,empId))return {status:'ERROR',message:'Không có quyền thực hiện.'};
  const sh=getSheet(TABLES.TRAVEL), rows=readData(TABLES.TRAVEL);
  for(let i=0;i<rows.length;i++)if(safeText(rows[i][0])===id){
    if(safeText(rows[i][1])!==empId)return {status:'ERROR',message:'Lịch trình không thuộc nhân viên này.'};
    if(safeText(rows[i][6]).toUpperCase()==='CANCELLED')return {status:'ERROR',message:'Lịch trình đã hủy.'};
    const fromDate=formatDate(d.fromDate),toDate=formatDate(d.toDate); if(!fromDate||!toDate||parseDateSafe(toDate)<parseDateSafe(fromDate))return {status:'ERROR',message:'Khoảng ngày không hợp lệ.'};
    const old=rows[i].slice(),r=old.slice(); r[2]=safeText(d.fromLocation);r[3]=safeText(d.toLocation);r[4]=fromDate;r[5]=toDate;
    if(d.ticketFileId!==undefined)r[7]=safeText(d.ticketFileId); if(d.ticketUrl!==undefined||d.purpose!==undefined)r[8]=makeTravelTicketUrl(safeText(d.ticketUrl),safeText(d.purpose));
    sh.getRange(i+2,1,1,11).setValues([r.slice(0,11)]);calculateEmployeeLocationStatus(empId);writeAuditLog(req.userId,'UPDATE_DOMESTIC_TRAVEL',TABLES.TRAVEL,id,JSON.stringify(old),JSON.stringify(r));
    return {status:'SUCCESS',message:'Đã cập nhật lịch trình.'};
  }
  return {status:'ERROR',message:'Không tìm thấy lịch trình.'};
}

function handleCancelDomesticTravel(d,req) {
  const id=safeText(d&&d.travelId), rows=readData(TABLES.TRAVEL), sh=getSheet(TABLES.TRAVEL); if(!id)return {status:'ERROR',message:'Thiếu mã lịch trình.'};
  for(let i=0;i<rows.length;i++)if(safeText(rows[i][0])===id){const empId=safeText(rows[i][1]);if(!canAccessEmployee(req,empId))return {status:'ERROR',message:'Không có quyền.'};const old=rows[i].slice();sh.getRange(i+2,7).setValue('CANCELLED');calculateEmployeeLocationStatus(empId);writeAuditLog(req.userId,'CANCEL_DOMESTIC_TRAVEL',TABLES.TRAVEL,id,old[6],'CANCELLED');return {status:'SUCCESS',message:'Đã hủy lịch trình.'};}
  return {status:'ERROR',message:'Không tìm thấy lịch trình.'};
}

function handleApproveDomesticTravel(d,req) {
  if(!isManagement(req))return {status:'ERROR',message:'Không có quyền duyệt yêu cầu.'};
  const id=safeText(d&&d.travelId),status=safeText(d&&d.status).toUpperCase();if(!id||!['APPROVED','REJECTED'].includes(status))return {status:'ERROR',message:'Trạng thái duyệt không hợp lệ.'};
  const sh=getSheet(TABLES.TRAVEL),rows=readData(TABLES.TRAVEL);for(let i=0;i<rows.length;i++)if(safeText(rows[i][0])===id){const old=rows[i][6];sh.getRange(i+2,7).setValue(status);calculateEmployeeLocationStatus(rows[i][1]);writeAuditLog(req.userId,'APPROVE_DOMESTIC_TRAVEL',TABLES.TRAVEL,id,old,status);return {status:'SUCCESS',message:'Đã cập nhật trạng thái: '+status};}
  return {status:'ERROR',message:'Không tìm thấy mã yêu cầu.'};
}

function handleGetDomesticTravel(empId,req) {
  if(!canAccessEmployee(req,empId))return {status:'ERROR',message:'Không có quyền xem thông tin.'};
  const records=readData(TABLES.TRAVEL).filter(r=>safeText(r[1])===safeText(empId)).map(r=>({travelId:r[0],employeeId:r[1],fromLocation:r[2],toLocation:r[3],fromDate:formatDate(r[4]),toDate:formatDate(r[5]),status:safeText(r[6]).toUpperCase(),ticketFileId:r[7]||'',ticketUrl:travelTicketUrlFromRow(r),purpose:travelPurposeFromRow(r)})).sort((a,b)=>parseDateSafe(b.fromDate)-parseDateSafe(a.fromDate));
  return {status:'SUCCESS',records};
}

// ============================================================
// DASHBOARD / PROFILE / TIMELINE
// ============================================================
function getDashboardData(req) {
  if(!isManagement(req))return {status:'ERROR',message:'Không có quyền truy cập dữ liệu quản trị.'};
  const empSheet=getSheet(TABLES.EMPLOYEES), empRows=readData(TABLES.EMPLOYEES), eeRows=readData(TABLES.ENTRY_EXIT), trRows=readData(TABLES.TRAVEL), ctrRows=readData(TABLES.CONTRACTS), docRows=readData(TABLES.DOCUMENTS), now=new Date();
  // Recalculate status once per active employee.
  empRows.forEach(r=>{if(safeText(r[0])&&safeText(r[10]).toUpperCase()==='ACTIVE')calculateEmployeeLocationStatus(r[0]);});
  const fresh=readData(TABLES.EMPLOYEES), activeContracts={};
  ctrRows.forEach(r=>{if(isTrue(r[9]))activeContracts[safeText(r[1])]={contractNo:r[2],endDate:r[5]};});
  const nextEE={},nextTravel={};
  eeRows.forEach(r=>{const id=safeText(r[1]),d=parseDateSafe(r[3]);if(!id||d<now)return;const n=parseEENote(r);if(!nextEE[id]||d<nextEE[id].rawDate)nextEE[id]={entryExitId:r[0],tripId:n.tripId||r[0],rawDate:d,type:safeText(r[2]).toUpperCase(),dateTime:formatDateTime(n.plannedDateTime||r[3]),airport:r[4]||'',flightNo:n.flightNo||'',destination:n.destination||'',purpose:n.purpose||''};});
  trRows.forEach(r=>{const id=safeText(r[1]),d=parseDateSafe(r[4]);if(!id||d.getTime()===0||safeText(r[6]).toUpperCase()!=='APPROVED'||d<new Date(now.getFullYear(),now.getMonth(),now.getDate()))return;if(!nextTravel[id]||d<nextTravel[id].rawDate)nextTravel[id]={travelId:r[0],rawDate:d,fromLocation:r[2]||'',toLocation:r[3]||'',fromDate:formatDate(r[4]),toDate:formatDate(r[5]),status:r[6]};});
  const kpi={total:0,inVN:0,business:0,exited:0,warning:0,expContract:0},employees=[],expiryWarnings=[],empMap={};
  fresh.forEach(r=>{const id=safeText(r[0]);empMap[id]={fullName:r[1]||''};if(safeText(r[10]).toUpperCase()!=='ACTIVE')return;kpi.total++;if(r[8]==='In Vietnam')kpi.inVN++;else if(r[8]==='Traveling')kpi.business++;else kpi.exited++;const c=activeContracts[id]||{};employees.push({employeeId:id,fullName:r[1],nationality:r[2],position:r[4],department:r[5],currentStatus:r[8],currentLocation:r[9],contractNo:c.contractNo||'-',contractExpiry:c.endDate?formatDate(c.endDate):'-',nextEntryExit:nextEE[id]||null,nextTravel:nextTravel[id]||null});});
  const pushWarning=(empId,doc,expiry,isContract)=>{if(!expiry)return;const d=parseDateSafe(expiry);if(d.getTime()===0)return;const days=Math.ceil((d-now)/86400000);if(days<=EXPIRING_DAYS){if(isContract)kpi.expContract++;else kpi.warning++;expiryWarnings.push({employeeId:empId,fullName:(empMap[empId]||{}).fullName||'N/A',document:doc,expiryDate:formatDate(expiry),daysRemaining:days,severity:days<=0?'EXPIRED':'WARNING'});}};
  docRows.forEach(r=>{if(isTrue(r[9]))pushWarning(safeText(r[1]),safeText(r[2]),r[5],false);});
  ctrRows.forEach(r=>{if(isTrue(r[9]))pushWarning(safeText(r[1]),'CONTRACT ('+safeText(r[2])+')',r[5],true);});
  expiryWarnings.sort((a,b)=>a.daysRemaining-b.daysRemaining);
  return {status:'SUCCESS',kpi,employees,expiryWarnings};
}

function getProfileData(employeeId,req) {
  if(!canAccessEmployee(req,employeeId))return {status:'ERROR',message:'Bạn không có quyền xem thông tin hồ sơ này.'};
  const emp=findEmployee(employeeId);if(!emp)return {status:'ERROR',message:'Không tìm thấy hồ sơ nhân viên.'};
  const docs=readData(TABLES.DOCUMENTS).filter(r=>safeText(r[1])===safeText(employeeId)&&isTrue(r[9]));
  const passport=docs.find(r=>safeText(r[2]).toUpperCase()==='PASSPORT')||[],trc=docs.find(r=>safeText(r[2]).toUpperCase()==='TRC')||[],wp=docs.find(r=>safeText(r[2]).toUpperCase()==='WORK_PERMIT')||[];
  const contracts=readData(TABLES.CONTRACTS).filter(r=>safeText(r[1])===safeText(employeeId)&&isTrue(r[9]));
  contracts.sort((a,b)=>parseDateSafe(b[5])-parseDateSafe(a[5]));const ctr=contracts[0]||[];
  const apps=readData(TABLES.APPENDICES).filter(r=>safeText(r[2])===safeText(employeeId)).sort((a,b)=>parseDateSafe(b[4])-parseDateSafe(a[4]));const app=apps[0]||[];
  const allowance=jsonParse(ctr[7],{}).allowance||0;
  return {status:'SUCCESS',profile:{employeeId:emp.row[0],fullName:emp.row[1],nationality:emp.row[2],dob:formatDate(emp.row[3]),position:emp.row[4],department:emp.row[5],phone:emp.row[6],email:emp.row[7],currentStatus:emp.row[8],currentLocation:emp.row[9],passportNo:passport[3]||'',passportExpiry:formatDate(passport[5]),passportFileId:passport[7]||'',passportImg:passport[8]||'',passport:getDriveFileInfo(passport[7],passport[8]),trcNo:trc[3]||'',trcExpiry:formatDate(trc[5]),trcFileId:trc[7]||'',trcImg:trc[8]||'',trc:getDriveFileInfo(trc[7],trc[8]),wpNo:wp[3]||'',wpExpiry:formatDate(wp[5]),wpFileId:wp[7]||'',wpImg:wp[8]||'',workPermit:getDriveFileInfo(wp[7],wp[8]),contractNo:ctr[2]||'',appendixNo:app[3]||'',contractStartDate:formatDate(ctr[4]),contractExpiry:formatDate(ctr[5]),contractFileId:ctr[10]||'',contractImg:ctr[11]||'',contract:getDriveFileInfo(ctr[10],ctr[11]),salary:ctr[6]||0,allowance}};
}

function handleGetTimeline(empId,req) {
  if(!canAccessEmployee(req,empId))return {status:'ERROR',message:'Không có quyền xem timeline.'};
  const timeline=[];
  handleGetEntryExit(empId,req).records.forEach(x=>timeline.push({id:x.entryExitId,kind:'ENTRY_EXIT',rawDate:x.plannedDateTime,title:x.type==='ENTRY'?'🛬 Nhập cảnh Việt Nam':'🛫 Xuất cảnh Việt Nam',date:x.plannedDateTime,location:'Cửa khẩu: '+(x.airport||'-')+' | Chuyến bay: '+(x.flightNo||'-')+' | Điểm đến: '+(x.destination||'-'),purpose:x.purpose||''}));
  handleGetDomesticTravel(empId,req).records.filter(x=>x.status!=='CANCELLED').forEach(x=>timeline.push({id:x.travelId,kind:'DOMESTIC_TRAVEL',rawDate:x.fromDate,title:'🚗 Đi lại / Công tác nội địa',date:x.fromDate,endDate:x.toDate,location:(x.fromLocation||'-')+' → '+(x.toLocation||'-'),status:x.status,purpose:x.purpose||''}));
  timeline.sort((a,b)=>parseDateSafe(b.rawDate)-parseDateSafe(a.rawDate));return {status:'SUCCESS',timeline};
}

// ============================================================
// IMPORT / SYNC MASTER EXCEL
// ============================================================
function normalizeImportObject(o) {
  const out={}; Object.keys(o||{}).forEach(k=>{const nk=String(k).trim();out[nk]=o[k];out[nk.toLowerCase()]=o[k];});return out;
}
function importValue(o,names) { for(const n of names){if(o[n]!==undefined&&safeText(o[n])!=='')return o[n];if(o[String(n).toLowerCase()]!==undefined&&safeText(o[String(n).toLowerCase()])!=='')return o[String(n).toLowerCase()];} return ''; }

function handleImportEmployeesExcel(data,req) {
  if(!isHR(req))return {status:'ERROR',message:'Chỉ HR/HR ADMIN/ADMIN được import.'};
  const rows=Array.isArray(data&&data.rows)?data.rows:[]; if(!rows.length)return {status:'ERROR',message:'Không có dữ liệu Excel.'}; if(rows.length>1000)return {status:'ERROR',message:'Mỗi lần import tối đa 1.000 dòng.'};
  const sh=getSheet(TABLES.EMPLOYEES), existing=readData(TABLES.EMPLOYEES), byId={};existing.forEach((r,i)=>byId[safeText(r[0])]={row:r,rowIndex:i+2});
  const now=new Date(), newRows=[], result={status:'SUCCESS',total:0,created:0,updated:0,skipped:0,errors:[]}, seen={};
  rows.forEach((raw,idx)=>{const o=normalizeImportObject(raw), rowNo=Number(raw._excelRow||idx+2),id=safeText(importValue(o,['EmployeeID','EmpID','employeeId'])),name=safeText(importValue(o,['FullName','Name','fullName']));result.total++;if(!id||!name){result.skipped++;result.errors.push({row:rowNo,message:'Thiếu EmployeeID hoặc FullName.'});return;}if(seen[id]){result.skipped++;result.errors.push({row:rowNo,message:'Trùng EmployeeID trong file: '+id});return;}seen[id]=true;const cur=byId[id];if(cur){const r=cur.row.slice();r[1]=name;r[2]=safeText(importValue(o,['Nationality']))||r[2];r[3]=formatDate(importValue(o,['DOB','DateOfBirth']))||r[3];r[4]=safeText(importValue(o,['Position']))||r[4];r[5]=safeText(importValue(o,['Department']))||r[5];r[6]=safeText(importValue(o,['Phone']))||r[6];r[7]=safeText(importValue(o,['Email'])).toLowerCase()||r[7];r[8]=safeText(importValue(o,['CurrentStatus','Status']))||r[8];r[9]=safeText(importValue(o,['CurrentLocation','Location']))||r[9];r[10]=safeText(importValue(o,['ActiveStatus']))?.toUpperCase()||r[10];r[13]=now;sh.getRange(cur.rowIndex,1,1,14).setValues([r.slice(0,14)]);result.updated++;}else{const folder=createEmployeeDriveFolder(id,name);const r=[id,name,safeText(importValue(o,['Nationality'])),formatDate(importValue(o,['DOB','DateOfBirth'])),safeText(importValue(o,['Position'])),safeText(importValue(o,['Department'])),safeText(importValue(o,['Phone'])),safeText(importValue(o,['Email'])).toLowerCase(),safeText(importValue(o,['CurrentStatus','Status']))||'Exited',safeText(importValue(o,['CurrentLocation','Location']))||'Overseas',safeText(importValue(o,['ActiveStatus']))?.toUpperCase()||'ACTIVE',folder,now,now];newRows.push(r);result.created++;}});
  if(newRows.length)sh.getRange(sh.getLastRow()+1,1,newRows.length,14).setValues(newRows);
  // Create missing user accounts in one batch.
  const us=getSheet(TABLES.USERS), ur=readData(TABLES.USERS), keys=new Set();ur.forEach(r=>{keys.add(safeText(r[1]).toLowerCase());keys.add(safeText(r[4]));});const users=[];newRows.forEach(r=>{const id=safeText(r[0]),email=safeText(r[7]).toLowerCase(),login=email||id;if(!keys.has(login.toLowerCase())&&!keys.has(id)){users.push([uid('USR'),login,hashPassword(DEFAULT_USER_PASSWORD),ROLES.USER,id,'ACTIVE',now]);keys.add(login.toLowerCase());keys.add(id);}});if(users.length)us.getRange(us.getLastRow()+1,1,users.length,7).setValues(users);
  writeAuditLog(req.userId,'IMPORT_EMPLOYEES','Employees','BATCH',null,JSON.stringify({total:result.total,created:result.created,updated:result.updated,skipped:result.skipped}));return result;
}

function handleSyncMasterExcel(data,req) {
  if(!isHR(req))return {status:'ERROR',message:'Chỉ HR/HR ADMIN/ADMIN được đồng bộ Master Excel.'};
  data=data||{};const result={status:'SUCCESS',employees:{total:0,created:0,updated:0,errors:[]},documents:{total:0,created:0,errors:[]},contracts:{total:0,created:0,errors:[]},appendices:{total:0,created:0,errors:[]},entryExit:{total:0,created:0,errors:[]},domesticTravel:{total:0,created:0,errors:[]}};
  const employeeRows=Array.isArray(data.employees)?data.employees:[]; if(employeeRows.length) {const r=handleImportEmployeesExcel({rows:employeeRows},req);result.employees={total:r.total||0,created:r.created||0,updated:r.updated||0,errors:r.errors||[]};}
  const employeeExists=id=>!!findEmployee(id);
  (Array.isArray(data.documents)?data.documents:[]).forEach(raw=>{result.documents.total++;const o=normalizeImportObject(raw),id=safeText(importValue(o,['EmployeeID','EmpID']));if(!employeeExists(id)){result.documents.errors.push({message:'Không tìm thấy EmployeeID '+id});return;}const x=saveDocument({employeeId:id,docType:importValue(o,['DocType','Type']),docNo:importValue(o,['DocNo','DocumentNo']),issueDate:importValue(o,['IssueDate']),expiryDate:importValue(o,['ExpiryDate']),issuer:importValue(o,['Issuer']),fileId:importValue(o,['FileId']),fileUrl:importValue(o,['FileUrl'])},req);if(x.status==='SUCCESS')result.documents.created++;else result.documents.errors.push({message:x.message});});
  (Array.isArray(data.contracts)?data.contracts:[]).forEach(raw=>{result.contracts.total++;const o=normalizeImportObject(raw),id=safeText(importValue(o,['EmployeeID','EmpID']));if(!employeeExists(id)){result.contracts.errors.push({message:'Không tìm thấy EmployeeID '+id});return;}const x=saveContractAndAppendix({employeeId:id,contractNo:importValue(o,['ContractNo']),contractType:importValue(o,['ContractType']),contractStartDate:importValue(o,['StartDate','ContractStartDate']),contractExpiry:importValue(o,['EndDate','ContractExpiry']),salary:importValue(o,['Salary']),allowance:importValue(o,['Allowance']),contractFileId:importValue(o,['FileId']),contractImg:importValue(o,['FileUrl'])},req);result.contracts.created++;});
  (Array.isArray(data.appendices)?data.appendices:[]).forEach(raw=>{result.appendices.total++;const o=normalizeImportObject(raw),id=safeText(importValue(o,['EmployeeID','EmpID']));if(!employeeExists(id)){result.appendices.errors.push({message:'Không tìm thấy EmployeeID '+id});return;}const f=findEmployee(id);const apps=readData(TABLES.APPENDICES), ctr=readData(TABLES.CONTRACTS).find(r=>safeText(r[1])===id&&isTrue(r[9]));if(!ctr){result.appendices.errors.push({message:'Chưa có Contract cho '+id});return;}const no=safeText(importValue(o,['AppendixNo']));if(!no){result.appendices.errors.push({message:'Thiếu AppendixNo'});return;}if(!apps.some(r=>safeText(r[1])===safeText(ctr[0])&&safeText(r[3])===no)){const sh=getSheet(TABLES.APPENDICES),row=[uid('APP'),ctr[0],id,no,formatDate(importValue(o,['EffectiveDate'])),formatDate(importValue(o,['EndDate'])),Number(importValue(o,['NewSalary']))||0,safeText(importValue(o,['AllowancesJson'])),safeText(importValue(o,['Content'])),safeText(importValue(o,['FileId'])),safeText(importValue(o,['FileUrl'])),new Date(),req.userId];sh.getRange(sh.getLastRow()+1,1,1,13).setValues([row]);result.appendices.created++;}});
  (Array.isArray(data.entryExit)?data.entryExit:[]).forEach(raw=>{result.entryExit.total++;const o=normalizeImportObject(raw),id=safeText(importValue(o,['EmployeeID','EmpID']));if(!employeeExists(id)){result.entryExit.errors.push({message:'Không tìm thấy EmployeeID '+id});return;}const type=safeText(importValue(o,['Type'])).toUpperCase(),dt=importValue(o,['EventDate','DateTime','PlannedDateTime']),x=handleAddEntryExit({employeeId:id,type,plannedDateTime:dt,actualDateTime:importValue(o,['ActualDateTime']),airport:importValue(o,['PortName','Airport']),flightNo:importValue(o,['FlightNo']),destination:importValue(o,['Destination']),purpose:importValue(o,['Purpose']),tripId:importValue(o,['TripID','TripId'])},req);if(x.status==='SUCCESS')result.entryExit.created++;else result.entryExit.errors.push({message:x.message});});
  (Array.isArray(data.domesticTravel)?data.domesticTravel:[]).forEach(raw=>{result.domesticTravel.total++;const o=normalizeImportObject(raw),id=safeText(importValue(o,['EmployeeID','EmpID']));if(!employeeExists(id)){result.domesticTravel.errors.push({message:'Không tìm thấy EmployeeID '+id});return;}const x=handleAddDomesticTravel({employeeId:id,fromLocation:importValue(o,['FromLocation']),toLocation:importValue(o,['ToLocation']),fromDate:importValue(o,['FromDate','StartDate']),toDate:importValue(o,['ToDate','EndDate']),purpose:importValue(o,['Purpose']),ticketFileId:importValue(o,['TicketFileId']),ticketUrl:importValue(o,['TicketUrl'])},req);if(x.status==='SUCCESS')result.domesticTravel.created++;else result.domesticTravel.errors.push({message:x.message});});
  writeAuditLog(req.userId,'SYNC_MASTER_EXCEL','MasterExcel','BATCH',null,JSON.stringify(result));return result;
}

// ============================================================
// EMPLOYEE DEACTIVATION
// ============================================================
function handleDeleteEmployee(empId,req) {
  if(![ROLES.ADMIN,ROLES.HR_ADMIN].includes(normalizeRole(req.role)))return {status:'ERROR',message:'Chỉ ADMIN/HR ADMIN được vô hiệu hóa nhân viên.'};
  const f=findEmployee(empId);if(!f)return {status:'ERROR',message:'Không tìm thấy nhân viên.'};
  f.sheet.getRange(f.rowIndex,11).setValue('INACTIVE');
  const us=getSheet(TABLES.USERS),rows=readData(TABLES.USERS);rows.forEach((r,i)=>{if(safeText(r[4])===safeText(empId))us.getRange(i+2,6).setValue('INACTIVE');});
  writeAuditLog(req.userId,'DEACTIVATE','Employees',empId,'ACTIVE','INACTIVE');return {status:'SUCCESS',message:'Đã vô hiệu hóa hồ sơ và khóa tài khoản.'};
}

// ============================================================
// PASSWORD / ACCOUNT MANAGEMENT
// ============================================================
function validateNewPassword(p) {p=String(p||'');if(p.length<8)return {ok:false,message:'Mật khẩu mới phải có ít nhất 8 ký tự.'};if(p.length>128)return {ok:false,message:'Mật khẩu mới tối đa 128 ký tự.'};return {ok:true};}

function handleChangePassword(data,req) {
  data=data||{};const old=String(data.currentPassword||''),np=String(data.newPassword||''),cf=String(data.confirmPassword||'');if(!old||!np||!cf)return {status:'ERROR',message:'Vui lòng nhập đầy đủ mật khẩu.'};if(np!==cf)return {status:'ERROR',message:'Mật khẩu mới và xác nhận không giống nhau.'};const v=validateNewPassword(np);if(!v.ok)return {status:'ERROR',message:v.message};if(old===np)return {status:'ERROR',message:'Mật khẩu mới phải khác mật khẩu hiện tại.'};const f=findUserById(req.userId);if(!f||safeText(f.row[2])!==hashPassword(old))return {status:'ERROR',message:'Mật khẩu hiện tại không chính xác.'};f.sheet.getRange(f.rowIndex,3).setValue(hashPassword(np));writeAuditLog(req.userId,'CHANGE_PASSWORD','Users',req.userId,null,'Password changed');return {status:'SUCCESS',message:'Đổi mật khẩu thành công.'};
}

function getAccountUsers(req) {
  if(!isAccountAdmin(req))return {status:'ERROR',message:'Bạn không có quyền xem danh sách tài khoản.'};
  const em={};readData(TABLES.EMPLOYEES).forEach(r=>em[safeText(r[0])]={fullName:r[1]||'',activeStatus:r[10]||''});
  const users=readData(TABLES.USERS).map(r=>({userId:r[0],email:r[1],role:r[3],employeeId:r[4],status:safeText(r[5]).toUpperCase(),fullName:(em[safeText(r[4])]||{}).fullName||r[1]||r[4]||'',employeeActiveStatus:(em[safeText(r[4])]||{}).activeStatus||''}));
  return {status:'SUCCESS',users};
}

function adminResetPassword(data,req) {
  if(normalizeRole(req.role)!==ROLES.ADMIN)return {status:'ERROR',message:'Chỉ ADMIN mới được reset mật khẩu.'};
  const id=safeText(data&&data.userId),np=String(data&&data.newPassword||'');if(!id||!np)return {status:'ERROR',message:'Thiếu tài khoản hoặc mật khẩu mới.'};if(id===safeText(req.userId))return {status:'ERROR',message:'Hãy dùng chức năng Đổi mật khẩu cho chính ADMIN.'};const v=validateNewPassword(np);if(!v.ok)return {status:'ERROR',message:v.message};const f=findUserById(id);if(!f)return {status:'ERROR',message:'Không tìm thấy tài khoản.'};f.sheet.getRange(f.rowIndex,3).setValue(hashPassword(np));writeAuditLog(req.userId,'ADMIN_RESET_PASSWORD','Users',id,null,'Password reset by ADMIN');return {status:'SUCCESS',message:'Đã reset mật khẩu.'};
}

function setAccountStatus(data,req) {
  if(!isAccountAdmin(req))return {status:'ERROR',message:'Chỉ ADMIN/HR ADMIN được khóa hoặc mở tài khoản.'};
  const id=safeText(data&&data.userId),status=safeText(data&&data.status).toUpperCase();if(!id||!['ACTIVE','INACTIVE'].includes(status))return {status:'ERROR',message:'Trạng thái tài khoản không hợp lệ.'};if(id===safeText(req.userId)&&status==='INACTIVE')return {status:'ERROR',message:'Không thể tự khóa tài khoản đang đăng nhập.'};const f=findUserById(id);if(!f)return {status:'ERROR',message:'Không tìm thấy tài khoản.'};const old=safeText(f.row[5]).toUpperCase();f.sheet.getRange(f.rowIndex,6).setValue(status);writeAuditLog(req.userId,status==='ACTIVE'?'UNLOCK_ACCOUNT':'LOCK_ACCOUNT','Users',id,old,status);return {status:'SUCCESS',message:status==='ACTIVE'?'Đã mở lại tài khoản.':'Đã khóa tài khoản.'};
}
