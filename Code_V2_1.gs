const CFG = {
  TZ: 'Asia/Ho_Chi_Minh',
  EXPIRING_DAYS: 30,
  MAX_UPLOAD_MB: 8,
  S: ['Users', 'Employees', 'Documents', 'Contracts', 'ContractAppendices', 'EntryExit', 'DomesticTravel', 'AuditLog', 'Sessions']
};

function doGet(e) {
  try { return get_(e.parameter || {}); } 
  catch (x) { return out_({ status: 'ERROR', message: x.message }); }
}

function doPost(e) {
  try { return post_(JSON.parse(e.postData.contents || '{}')); } 
  catch (x) { return out_({ status: 'ERROR', message: x.message }); }
}

function setupSheets() {
  let ss = SpreadsheetApp.getActive();
  let h = {
    Users: ['id', 'email', 'password', 'role', 'name', 'employeeId', 'active'],
    Employees: ['employeeId', 'fullName', 'nationality', 'dob', 'passportNo', 'passportExpiry', 'visaNo', 'visaExpiry', 'trcNo', 'trcExpiry', 'wpNo', 'wpExpiry', 'position', 'department', 'phone', 'email', 'currentStatus', 'currentLocation', 'active'],
    Documents: ['documentId', 'employeeId', 'documentType', 'documentNo', 'issueDate', 'expiryDate', 'fileId', 'fileName', 'mimeType', 'note', 'active', 'createdAt', 'updatedAt'],
    Contracts: ['contractId', 'employeeId', 'contractNo', 'contractType', 'startDate', 'endDate', 'salary', 'salaryCurrency', 'allowanceTotal', 'allowanceCurrency', 'allowanceNote', 'fileId', 'fileName', 'status', 'note', 'createdAt', 'updatedAt'],
    ContractAppendices: ['appendixId', 'contractId', 'employeeId', 'appendixNo', 'effectiveDate', 'endDate', 'salary', 'salaryCurrency', 'allowanceTotal', 'allowanceCurrency', 'allowanceNote', 'fileId', 'fileName', 'status', 'note', 'createdAt', 'updatedAt'],
    EntryExit: ['id', 'employeeId', 'type', 'dateTime', 'airport', 'flightNo', 'ticketFile', 'destination', 'purpose', 'status', 'createdAt', 'createdBy'],
    DomesticTravel: ['id', 'employeeId', 'fromDate', 'toDate', 'fromLocation', 'toLocation', 'purpose', 'transport', 'flightNo', 'hotel', 'hotelAddress', 'ticketFile', 'status', 'createdAt', 'createdBy'],
    AuditLog: ['timestamp', 'requesterId', 'action', 'employeeId', 'details'],
    Sessions: ['token', 'userId', 'createdAt', 'expiresAt', 'active']
  };

  Object.keys(h).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(h[n]);
    sh.setFrozenRows(1);
  });

  try { ss.getSheetByName('Sessions').hideSheet(); } catch (e) {}
  return out_({ status: 'SUCCESS', message: 'Sheets ready' });
}

function post_(r) {
  let a = String(r.action || '').toUpperCase();
  if (a === 'LOGIN') return out_(login_(r.username, r.password));
  if (a === 'LOGOUT') {
    if (r.token) upd_('Sessions', 'token', r.token, { active: false });
    return out_({ status: 'SUCCESS' });
  }

  let u = auth_(r.token), d = r.data || {}, id = String(d.employeeId || '');

  if (a === 'ADD_ENTRY_EXIT') {
    needEdit_(u, id);
    if (!d.dateTime || !d.airport || !d.destination) throw Error('Vui lòng nhập ngày giờ, cửa khẩu và điểm đến.');
    let x = { id: Utilities.getUuid(), employeeId: id, type: d.type, dateTime: d.dateTime, airport: d.airport, flightNo: d.flightNo || '', ticketFile: d.ticketFile || '', destination: d.destination, purpose: d.purpose || '', status: d.status || 'Planned', createdAt: new Date(), createdBy: u.id };
    app_('EntryExit', x);
    audit_(u.id, a, id, x);
    if (x.status === 'Completed') upd_('Employees', 'employeeId', id, { currentStatus: x.type === 'EXIT' ? 'Exited' : 'In Vietnam' });
    return out_({ status: 'SUCCESS', record: x });
  }

  if (a === 'ADD_DOMESTIC_TRAVEL') {
    needEdit_(u, id);
    if (!d.fromDate || !d.toDate || new Date(d.fromDate) > new Date(d.toDate)) throw Error('Ngày đi/đến không hợp lệ.');
    if (!d.fromLocation || !d.toLocation) throw Error('Vui lòng nhập nơi đi và nơi đến.');
    let x = { id: Utilities.getUuid(), employeeId: id, fromDate: d.fromDate, toDate: d.toDate, fromLocation: d.fromLocation, toLocation: d.toLocation, purpose: d.purpose || '', transport: d.transport || '', flightNo: d.flightNo || '', hotel: d.hotel || '', hotelAddress: d.hotelAddress || '', ticketFile: '', status: d.status || 'Planned', createdAt: new Date(), createdBy: u.id };
    app_('DomesticTravel', x);
    audit_(u.id, a, id, x);
    if (x.status === 'In Progress') upd_('Employees', 'employeeId', id, { currentStatus: 'Traveling', currentLocation: x.toLocation });
    return out_({ status: 'SUCCESS', record: x });
  }

  if (['ADD_EMPLOYEE', 'UPDATE_EMPLOYEE', 'DEACTIVATE_EMPLOYEE', 'ACTIVATE_EMPLOYEE', 'SAVE_DOCUMENT', 'SAVE_CONTRACT', 'SAVE_APPENDIX'].includes(a)) needHR_(u.role);

  if (a === 'ADD_EMPLOYEE') {
    if (!d.employeeId || !d.fullName || !d.nationality) throw Error('Employee ID, Họ tên và Quốc tịch là bắt buộc.');
    if (getEntity_('employee', d.employeeId)) throw Error('Employee ID đã tồn tại.');
    let x = emp_(d);
    app_('Employees', x);
    audit_(u.id, a, d.employeeId, x);
    return out_({ status: 'SUCCESS', employee: x });
  }

  if (a === 'UPDATE_EMPLOYEE') {
    let x = getEntity_('employee', id);
    if (!x) throw Error('Không tìm thấy nhân viên.');
    let n = emp_(Object.assign({}, x, d));
    upd_('Employees', 'employeeId', id, n);
    audit_(u.id, a, id, n);
    return out_({ status: 'SUCCESS', employee: n });
  }

  if (a === 'DEACTIVATE_EMPLOYEE') {
    if (!getEntity_('employee', id)) throw Error('Không tìm thấy nhân viên.');
    upd_('Employees', 'employeeId', id, { active: false });
    audit_(u.id, a, id, d);
    return out_({ status: 'SUCCESS' });
  }

  if (a === 'ACTIVATE_EMPLOYEE') {
    upd_('Employees', 'employeeId', id, { active: true });
    audit_(u.id, a, id, d);
    return out_({ status: 'SUCCESS' });
  }

  if (a === 'SAVE_DOCUMENT') {
    if (!getEntity_('employee', id)) throw Error('Không tìm thấy nhân viên.');
    if (!d.documentType || !d.expiryDate) throw Error('Loại giấy tờ và ngày hết hạn là bắt buộc.');
    let f = upload_(d);
    let x = { documentId: d.documentId || Utilities.getUuid(), employeeId: id, documentType: d.documentType, documentNo: d.documentNo || '', issueDate: d.issueDate || '', expiryDate: d.expiryDate, fileId: f.id || d.fileId || '', fileName: f.name || d.fileName || '', mimeType: f.mimeType || d.mimeType || '', note: d.note || '', active: true, createdAt: new Date(), updatedAt: new Date() };
    if (getEntity_('doc', x.documentId)) upd_('Documents', 'documentId', x.documentId, x);
    else app_('Documents', x);
    audit_(u.id, a, id, { documentId: x.documentId });
    return out_({ status: 'SUCCESS', document: x });
  }

  if (a === 'SAVE_CONTRACT') {
    if (!getEntity_('employee', id) || !d.contractNo || !d.startDate) throw Error('Nhân viên, số hợp đồng và ngày bắt đầu là bắt buộc.');
    let f = upload_(d), x = { contractId: d.contractId || Utilities.getUuid(), employeeId: id, contractNo: d.contractNo, contractType: d.contractType || '', startDate: d.startDate, endDate: d.endDate || '', salary: d.salary || '', salaryCurrency: d.salaryCurrency || 'VND', allowanceTotal: d.allowanceTotal || '', allowanceCurrency: d.allowanceCurrency || 'VND', allowanceNote: d.allowanceNote || '', fileId: f.id || d.fileId || '', fileName: f.name || d.fileName || '', status: d.status || 'Active', note: d.note || '', createdAt: new Date(), updatedAt: new Date() };
    if (getEntity_('contract', x.contractId)) upd_('Contracts', 'contractId', x.contractId, x);
    else app_('Contracts', x);
    audit_(u.id, a, id, x);
    return out_({ status: 'SUCCESS', contract: x });
  }

  if (a === 'SAVE_APPENDIX') {
    let c = getEntity_('contract', d.contractId);
    if (!c || String(c.employeeId) !== id || !d.appendixNo) throw Error('Hợp đồng hoặc số phụ lục không hợp lệ.');
    let f = upload_(d), x = { appendixId: d.appendixId || Utilities.getUuid(), contractId: d.contractId, employeeId: id, appendixNo: d.appendixNo, effectiveDate: d.effectiveDate || '', endDate: d.endDate || '', salary: d.salary || '', salaryCurrency: d.salaryCurrency || c.salaryCurrency || 'VND', allowanceTotal: d.allowanceTotal || '', allowanceCurrency: d.allowanceCurrency || c.allowanceCurrency || 'VND', allowanceNote: d.allowanceNote || '', fileId: f.id || '', fileName: f.name || '', status: d.status || 'Active', note: d.note || '', createdAt: new Date(), updatedAt: new Date() };
    if (getEntity_('appx', x.appendixId)) upd_('ContractAppendices', 'appendixId', x.appendixId, x);
    else app_('ContractAppendices', x);
    audit_(u.id, a, id, x);
    return out_({ status: 'SUCCESS', appendix: x });
  }

  throw Error('Action không được hỗ trợ');
}

function get_(p) {
  let a = String(p.action || '').toUpperCase();
  if (a === 'GET_DOCUMENT_IMAGE') {
    let u = auth_(p.token), d = getEntity_('doc', p.documentId);
    if (!d || !can_(u, d.employeeId)) throw Error('Không có quyền.');
    return fileOut_(d.fileId);
  }
  if (a === 'GET_CONTRACT_FILE') {
    let u = auth_(p.token), sheet = p.type === 'APPENDIX' ? 'ContractAppendices' : 'Contracts', key = p.type === 'APPENDIX' ? 'appendixId' : 'contractId', d = getBy_(sheet, key, p.id);
    if (!d || !can_(u, d.employeeId)) throw Error('Không có quyền.');
    return fileOut_(d.fileId);
  }

  let u = auth_(p.token);
  if (a === 'GET_DASHBOARD_DATA') {
    needHR_(u.role);
    return out_(dashboard_());
  }

  let id = String(p.employeeId || u.employeeId || '');
  if (!can_(u, id)) throw Error('Không có quyền.');
  if (a === 'GET_PROFILE') return out_({ status: 'SUCCESS', profile: getEntity_('employee', id) });
  if (a === 'GET_ENTRY_EXIT') return out_({ status: 'SUCCESS', records: rows_('EntryExit').filter(x => String(x.employeeId) === id) });
  if (a === 'GET_DOMESTIC_TRAVEL') {
    autoTravel_();
    return out_({ status: 'SUCCESS', records: rows_('DomesticTravel').filter(x => String(x.employeeId) === id) });
  }
  if (a === 'GET_EMPLOYEE_DETAIL') {
    return out_({
      status: 'SUCCESS',
      detail: {
        employee: enrich_(getEntity_('employee', id)),
        documents: rows_('Documents').filter(x => String(x.employeeId) === id),
        contracts: rows_('Contracts').filter(x => String(x.employeeId) === id),
        timeline: timeline_(id)
      }
    });
  }

  throw Error('GET action không được hỗ trợ');
}

function login_(email, pw) {
  let u = rows_('Users').find(x => String(x.email).toLowerCase() === String(email).toLowerCase() && String(x.password) === String(pw) && truth_(x.active));
  if (!u) throw Error('Sai tài khoản hoặc mật khẩu.');
  let t = Utilities.getUuid() + Utilities.getUuid(), n = new Date(), e = new Date(n.getTime() + 8 * 3600000);
  app_('Sessions', { token: t, userId: u.id, createdAt: n, expiresAt: e, active: true });
  return { status: 'SUCCESS', sessionToken: t, user: { id: u.id, email: u.email, role: u.role, name: u.name, employeeId: u.employeeId || '' } };
}

function auth_(t) {
  if (!t) throw Error('Phiên đăng nhập không hợp lệ.');
  let s = rows_('Sessions').find(x => String(x.token) === String(t) && truth_(x.active));
  if (!s || new Date(s.expiresAt) < new Date()) throw Error('Phiên đã hết hạn.');
  let u = rows_('Users').find(x => String(x.id) === String(s.userId) && truth_(x.active));
  if (!u) throw Error('Tài khoản không hoạt động.');
  return u;
}

function needHR_(r) {
  if (!['HR', 'HR ADMIN', 'DIRECTOR', 'ADMIN'].includes(String(r).toUpperCase())) throw Error('Không có quyền HR.');
}

function can_(u, id) {
  return ['HR', 'HR ADMIN', 'DIRECTOR', 'ADMIN'].includes(String(u.role).toUpperCase()) || String(u.employeeId) === String(id);
}

function needEdit_(u, id) {
  if (!can_(u, id)) throw Error('Không có quyền thao tác.');
}

function dashboard_() {
  autoTravel_();
  let es = rows_('Employees').filter(x => x.active === '' || x.active == null || truth_(x.active)).map(enrich_), w = [];
  es.forEach(e => ['passport', 'visa', 'trc', 'wp'].forEach(t => {
    let s = ds_(e[t + 'Expiry']);
    if (s === 'EXPIRED' || s === 'EXPIRING') w.push({ employeeId: e.employeeId, fullName: e.fullName, type: t.toUpperCase(), expiry: e[t + 'Expiry'], status: s });
  }));
  return { status: 'SUCCESS', kpi: { total: es.length, inVN: es.filter(x => x.currentStatus === 'In Vietnam').length, business: es.filter(x => x.currentStatus === 'Traveling').length, exited: es.filter(x => x.currentStatus === 'Exited').length, warning: w.length }, employees: es, warnings: w };
}

function enrich_(e) {
  e = e || {};
  e.active = e.active === '' || e.active == null ? true : truth_(e.active);
  e.warningCount = ['passport', 'visa', 'trc', 'wp'].filter(t => ['EXPIRED', 'EXPIRING'].includes(ds_(e[t + 'Expiry']))).length;
  return e;
}

function ds_(v) {
  if (!v) return 'NO_DATE';
  let d = new Date(v), n = new Date();
  d.setHours(0, 0, 0, 0);
  n.setHours(0, 0, 0, 0);
  let z = (d - n) / 86400000;
  return z < 0 ? 'EXPIRED' : z <= CFG.EXPIRING_DAYS ? 'EXPIRING' : 'VALID';
}

function autoTravel_() {
  let sh = SpreadsheetApp.getActive().getSheetByName('DomesticTravel');
  if (!sh || sh.getLastRow() < 2) return;
  let range = sh.getDataRange();
  let v = range.getValues(), h = v[0].map(String);
  let si = h.indexOf('status'), fi = h.indexOf('fromDate'), ti = h.indexOf('toDate'), ei = h.indexOf('employeeId'), li = h.indexOf('toLocation'), now = new Date();
  let changed = false;

  for (let r = 1; r < v.length; r++) {
    let f = new Date(v[r][fi]), t = new Date(v[r][ti]);
    if (isNaN(f) || isNaN(t)) continue;
    let s = now < f ? 'Planned' : now <= new Date(t.getFullYear(), t.getMonth(), t.getDate(), 23, 59, 59) ? 'In Progress' : 'Completed';
    if (v[r][si] !== s) {
      v[r][si] = s;
      changed = true;
    }
    if (s === 'In Progress') upd_('Employees', 'employeeId', v[r][ei], { currentStatus: 'Traveling', currentLocation: v[r][li] || '' });
    if (s === 'Completed') upd_('Employees', 'employeeId', v[r][ei], { currentStatus: 'In Vietnam' });
  }
  if (changed) range.setValues(v);
}

function timeline_(id) {
  return rows_('EntryExit').filter(x => String(x.employeeId) === id).map(x => Object.assign({ kind: 'IMMIGRATION' }, x))
    .concat(rows_('DomesticTravel').filter(x => String(x.employeeId) === id).map(x => Object.assign({ kind: 'TRAVEL' }, x)))
    .sort((a, b) => String(b.dateTime || b.fromDate).localeCompare(String(a.dateTime || a.fromDate)));
}

function upload_(d) {
  if (!d.fileBase64) return {};
  let cleanBase64 = String(d.fileBase64).replace(/^data:[^;]+;base64,/, '');
  let b = Utilities.base64Decode(cleanBase64);
  if (b.length > CFG.MAX_UPLOAD_MB * 1024 * 1024) throw Error(`File tối đa ${CFG.MAX_UPLOAD_MB}MB.`);
  let f = DriveApp.getFoldersByName('Foreign Employee Secure Files');
  let folder = f.hasNext() ? f.next() : DriveApp.createFolder('Foreign Employee Secure Files');
  let name = String(d.fileName || 'file').replace(/[^\w.\- ]/g, '_'), file = folder.createFile(Utilities.newBlob(b, d.mimeType || 'application/octet-stream', name));
  return { id: file.getId(), name: file.getName(), mimeType: file.getMimeType() };
}

function fileOut_(id) {
  if (!id) throw Error('Chưa có file.');
  let b = DriveApp.getFileById(id).getBlob();
  return out_({ status: 'SUCCESS', fileName: b.getName(), mimeType: b.getContentType(), base64: Utilities.base64Encode(b.getBytes()) });
}

function emp_(d) {
  return { employeeId: String(d.employeeId || ''), fullName: String(d.fullName || ''), nationality: String(d.nationality || ''), dob: String(d.dob || ''), passportNo: String(d.passportNo || ''), passportExpiry: String(d.passportExpiry || ''), visaNo: String(d.visaNo || ''), visaExpiry: String(d.visaExpiry || ''), trcNo: String(d.trcNo || ''), trcExpiry: String(d.trcExpiry || ''), wpNo: String(d.wpNo || ''), wpExpiry: String(d.wpExpiry || ''), position: String(d.position || ''), department: String(d.department || ''), phone: String(d.phone || ''), email: String(d.email || ''), currentStatus: String(d.currentStatus || 'In Vietnam'), currentLocation: String(d.currentLocation || ''), active: d.active === false ? false : true };
}

function rows_(n) {
  let s = SpreadsheetApp.getActive().getSheetByName(n);
  if (!s || s.getLastRow() < 2) return [];
  let v = s.getDataRange().getValues(), h = v[0].map(String);
  return v.slice(1).map(r => Object.fromEntries(h.map((x, i) => [x, r[i]])));
}

function getEntity_(type, id) {
  let m = { employee: ['Employees', 'employeeId'], doc: ['Documents', 'documentId'], contract: ['Contracts', 'contractId'], appx: ['ContractAppendices', 'appendixId'] }[type];
  return m ? getBy_(m[0], m[1], id) : null;
}

function getBy_(s, k, v) {
  return rows_(s).find(x => String(x[k]) === String(v)) || null;
}

function app_(s, o) {
  let sh = SpreadsheetApp.getActive().getSheetByName(s), h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  sh.appendRow(h.map(k => o[k] === undefined ? '' : o[k]));
}

function upd_(s, k, v, p) {
  let sh = SpreadsheetApp.getActive().getSheetByName(s);
  if (!sh) throw Error('Sheet không tồn tại.');
  let range = sh.getDataRange();
  let d = range.getValues(), h = d[0].map(String), ki = h.indexOf(k);

  for (let r = 1; r < d.length; r++) {
    if (String(d[r][ki]) === String(v)) {
      Object.keys(p).forEach(x => {
        let i = h.indexOf(x);
        if (i >= 0) d[r][i] = p[x];
      });
      range.setValues(d); // Cập nhật nguyên bảng 1 lần để tối ưu performance
      return;
    }
  }
  throw Error('Không tìm thấy bản ghi.');
}

function audit_(u, a, e, d) {
  app_('AuditLog', { timestamp: new Date(), requesterId: u, action: a, employeeId: e, details: JSON.stringify(d || {}) });
}

function truth_(v) {
  return v === true || ['true', '1', 'yes'].includes(String(v).toLowerCase());
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
