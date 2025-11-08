/**
 * Code.gs - 完全版（最新版）
 *
 * - ヘッダ名の正規化を強化して、"A-1_目標" 等を柔軟に検出します。
 * - 評価明細の更新時に列がずれないよう、既存行更新は A 列(1) から書き込みます。
 * - setHeaderCellByName は正規化キーで探す安全実装にしています。
 *
 * 保存前にバックアップを推奨します。
 */

/* ========== 設定 ========== */
const ADMIN_EMAILS = [
  'info@digisupo-kagoshima.com'
];
const FALLBACK_EXEC_URL = '';

/* ========== 共通ユーティリティ ========== */
function _normalizeForRpc(value) {
  if (value === null || value === undefined) return value;
  var t = typeof value;
  if (t === 'string' || t === 'number' || t === 'boolean') return value;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    try { return value.toISOString(); } catch (e) { return String(value); }
  }
  if (Array.isArray(value)) {
    return value.map(function(v){ try { return _normalizeForRpc(v); } catch(e){ return String(v); } });
  }
  if (t === 'object') {
    var out = {};
    for (var k in value) {
      if (!Object.prototype.hasOwnProperty.call(value, k)) continue;
      try { out[k] = _normalizeForRpc(value[k]); } catch(e){ out[k] = String(value[k]); }
    }
    return out;
  }
  try { return String(value); } catch(e) { return null; }
}

function _getExecUrlForTemplate() {
  try {
    if (typeof ScriptApp !== 'undefined' && ScriptApp.getService) {
      var u = ScriptApp.getService().getUrl();
      if (u) return u;
    }
  } catch (e) {}
  return FALLBACK_EXEC_URL || '';
}

function _isAdminEmail(email) {
  if (!email) return false;
  try {
    var e = String(email).toLowerCase().trim();
    return ADMIN_EMAILS.some(function(a){ return a && a.toLowerCase().trim() === e; });
  } catch (e) { return false; }
}

/* ========== ヘッダ名正規化（堅牢化） ========== */
function _normalizeHeaderNameForKey(name) {
  if (!name && name !== 0) return '';
  try {
    var s = String(name).normalize('NFKC');           // NFKC で全角英数字等を近似変換
    s = s.replace(/\u3000/g, ' ').trim();             // 全角スペースを半角に
    s = s.replace(/\s+/g, ' ');                       // 連続空白を1つに
    s = s.replace(/[–—−ー]/g, '-');                   // ハイフン類を半角ハイフンに統一
    s = s.replace(/[\u0000-\u001F]/g, '');            // 制御文字除去
    var key = s.replace(/\s+/g, '').toLowerCase();   // 空白除去して小文字化
    return key;
  } catch (e) {
    return String(name).replace(/\s+/g,'').toLowerCase();
  }
}

function buildHeaderIndexMap(sheet) {
  if (!sheet) return {};
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var raw = headerRow[i] ? String(headerRow[i]) : '';
    var key = _normalizeHeaderNameForKey(raw);
    if (!key) continue;
    map[key] = i;
    // 互換キーをいくつか追加しておく
    try {
      var alt1 = key.replace(/[-_]/g, ''); // ハイフン/アンダースコア除去
      if (alt1 && !(alt1 in map)) map[alt1] = i;
      var alt2 = key.replace(/-/g, '_');
      if (alt2 && !(alt2 in map)) map[alt2] = i;
      var alt3 = key.replace(/_/g, '-');
      if (alt3 && !(alt3 in map)) map[alt3] = i;
    } catch (e) {}
  }
  return map;
}

function getRowValueByHeader(rowArray, headerMap, headerName) {
  if (!rowArray || !headerMap) return '';
  var key = _normalizeHeaderNameForKey(headerName || '');
  var idx = headerMap[key];
  if (typeof idx === 'number' && rowArray.length > idx) return rowArray[idx];

  // フォールバックの試行
  var altKeys = [];
  try {
    var noHyphen = key.replace(/[-_]/g, '');
    if (noHyphen !== key) altKeys.push(noHyphen);
    var hyToUnder = key.replace(/-/g, '_');
    if (hyToUnder !== key) altKeys.push(hyToUnder);
    var underToHy = key.replace(/_/g, '-');
    if (underToHy !== key) altKeys.push(underToHy);
  } catch (e) {}
  for (var k = 0; k < altKeys.length; k++) {
    var a = altKeys[k];
    if (headerMap[a] !== undefined && typeof headerMap[a] === 'number' && rowArray.length > headerMap[a]) {
      return rowArray[ headerMap[a] ];
    }
  }
  return '';
}

/* ========== Web UI entry ========== */
function doGet(e) {
  try {
    e = e || {}; e.parameter = e.parameter || {};
    var params = e.parameter;
    var page = params.page || '';
    var idParam = params.id || params.evaluationId || '';

    var activeEmail = '';
    var usedSource = '';
    try {
      if (Session && Session.getActiveUser && Session.getActiveUser().getEmail) {
        activeEmail = Session.getActiveUser().getEmail() || '';
        if (activeEmail) usedSource = 'Session.getActiveUser';
      }
    } catch (err) {}
    if (!activeEmail) {
      try {
        if (ScriptApp && ScriptApp.getEffectiveUser && ScriptApp.getEffectiveUser().getEmail) {
          activeEmail = ScriptApp.getEffectiveUser().getEmail() || '';
          if (activeEmail) usedSource = 'ScriptApp.getEffectiveUser';
        }
      } catch (err) {}
    }
    var emailLower = (activeEmail || '').toString().toLowerCase().trim();
    Logger.log('doGet called page=%s id=%s sessionEmail=%s source=%s', page, idParam, activeEmail || '(none)', usedSource || '(none)');

    var role = 'Unknown';
    try { role = getUserRole(emailLower); } catch (err) { Logger.log('doGet getUserRole error: %s', err && err.stack ? err.stack : String(err)); }
    if (_isAdminEmail(emailLower)) role = 'Admin';
    Logger.log('doGet resolved role=%s for email=%s', role, emailLower);

    var template;
    var execUrl = _getExecUrlForTemplate();

    if (page === 'evaluation') {
      if (!idParam) return createErrorHtml('パラメータ不足', '評価IDが指定されていません。');
      template = HtmlService.createTemplateFromFile('評価入力画面');
      template.evaluationId = idParam;
      template.userEmail = emailLower;
      template.userRole = role;
      template.execUrl = execUrl;
      return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } else if (page === 'admin') {
      if (role !== 'Admin') {
        Logger.log('doGet admin denied for %s (role=%s)', emailLower, role);
        return createErrorHtml('権限エラー', '管理ダッシュボードへのアクセス権限がありません。ログインメール: ' + (emailLower || '(取得できません)'));
      }
      template = HtmlService.createTemplateFromFile('管理ダッシュボード');
      template.execUrl = execUrl;
      template.adminData = getAdminPageData();
      return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } else {
      template = HtmlService.createTemplateFromFile('メインダッシュボード');
      template.userEmail = emailLower;
      template.userRole = role;
      template.execUrl = execUrl;
      return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (error) {
    Logger.log('doGet top-level error: %s\n%s', error && error.message ? error.message : String(error), error && error.stack ? error.stack : '');
    return createErrorHtml('予期せぬエラー', 'アプリケーションでエラーが発生しました。管理者に連絡してください。');
  }
}

function createErrorHtml(title, message) {
  var safe = (message || '').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
  var html = '<html><head><meta charset="utf-8"><title>' + (title || 'エラー') + '</title>' +
             '<style>body{font-family:sans-serif;margin:30px;background:#f7f7f7} .box{background:#fff;padding:20px;border:1px solid #ddd;border-radius:6px}</style>' +
             '</head><body><div class="box"><h1>' + (title || '') + '</h1><div>' + safe + '</div></div></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ========== ユーザ / マスタ関連 ========== */
function getUserRole(email) {
  if (!email) return 'Unknown';
  if (ADMIN_EMAILS.map(a => a.toLowerCase()).includes(email)) return 'Admin';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('社員マスタ');
  if (!master) throw new Error('社員マスタが見つかりません。');
  var data = master.getDataRange().getValues();
  var emailCol = 2, idCol = 0;
  var eval1Col = 9, eval2Col = 10, eval3Col = 11;
  var isEvaluator = false, employeeId = null;
  for (var i = 1; i < data.length; i++) {
    var row = data[i] || [];
    try {
      if ((row[emailCol] || '').toLowerCase().trim() === email) employeeId = row[idCol];
      if ((row[eval1Col] || '').toLowerCase().trim() === email || (row[eval2Col] || '').toLowerCase().trim() === email || (row[eval3Col] || '').toLowerCase().trim() === email) {
        isEvaluator = true;
      }
    } catch (e) {}
  }
  if (employeeId && !isEvaluator) isEvaluator = isUserEvaluator(data, employeeId);
  if (isEvaluator) return 'Evaluator';
  if (employeeId) return 'Employee';
  return 'Unknown';
}

function findEmployeeByEmail(data, email) {
  if (!data || !Array.isArray(data)) return null;
  var emailCol = 2, idCol = 0, nameCol = 1, gradeCol = 7, eval1Col = 9, eval2Col = 10, eval3Col = 11;
  for (var i = 1; i < data.length; i++) {
    var row = data[i] || [];
    if ((row[emailCol] || '').toLowerCase().trim() === (email || '').toLowerCase().trim()) {
      return {
        id: row[idCol] || '',
        name: row[nameCol] || '',
        grade: row[gradeCol] || '',
        eval1Id: row[eval1Col] || '',
        eval2Id: row[eval2Col] || '',
        eval3Id: row[eval3Col] || '',
        row: i + 1
      };
    }
  }
  return null;
}

function findEmployeeById(data, id) {
  try {
    if (!data || !Array.isArray(data)) return null;
    var idCol = 0, nameCol = 1, emailCol = 2, deptCol = 3, genderCol = 4, dobCol = 5, joinedCol = 6, gradeCol = 7, numberCol = 8, eval1Col = 9, eval2Col = 10, eval3Col = 11, statusCol = 12;
    for (var i = 1; i < data.length; i++) {
      var row = data[i] || [];
      try {
        if ((row[idCol] || '') === id) {
          function fmt(cell) {
            try {
              if (!cell && cell !== 0) return '';
              if (Object.prototype.toString.call(cell) === '[object Date]') return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy/MM/dd');
              return String(cell);
            } catch (e) { return String(cell || ''); }
          }
          return {
            id: row[idCol] || '',
            name: row[nameCol] || '',
            email: (row[emailCol] || '').toLowerCase(),
            department: row[deptCol] || '',
            gender: row[genderCol] || '',
            dob: fmt(row[dobCol]),
            joined: fmt(row[joinedCol]),
            grade: row[gradeCol] || '',
            number: row[numberCol] || '',
            eval1Id: row[eval1Col] || '',
            eval2Id: row[eval2Col] || '',
            eval3Id: row[eval3Col] || '',
            status: row[statusCol] || '',
            row: i + 1
          };
        }
      } catch (eRow) { Logger.log('findEmployeeById row error: %s', eRow && eRow.message ? eRow.message : String(eRow)); }
    }
  } catch (e) { Logger.log('findEmployeeById error: %s', e && e.stack ? e.stack : String(e)); }
  return null;
}

function isUserEvaluator(data, employeeId) {
  if (!data || !Array.isArray(data)) return false;
  var eval1Col = 9, eval2Col = 10, eval3Col = 11;
  for (var i = 1; i < data.length; i++) {
    var row = data[i] || [];
    if ((row[eval1Col] || '') === employeeId || (row[eval2Col] || '') === employeeId || (row[eval3Col] || '') === employeeId) {
      return true;
    }
  }
  return false;
}

function getSubordinates(headerData, loginUserId, masterData) {
  try {
    if (!loginUserId || !masterData || !Array.isArray(masterData)) return [];
    var eval1Col = 9, eval2Col = 10, eval3Col = 11, idCol = 0;
    var set = {};
    for (var i = 1; i < masterData.length; i++) {
      var row = masterData[i] || [];
      var empId = (row[idCol] || '').toString();
      if (!empId) continue;
      var e1 = (row[eval1Col] || '').toString(), e2 = (row[eval2Col] || '').toString(), e3 = (row[eval3Col] || '').toString();
      if (e1 === loginUserId || e2 === loginUserId || e3 === loginUserId) set[empId] = true;
    }
    return Object.keys(set);
  } catch (e) {
    Logger.log('getSubordinates error: %s', e && e.stack ? e.stack : String(e));
    return [];
  }
}

/* ========== ダッシュボード関連 ========== */
function getDashboardData() {
  try {
    Logger.log('getDashboardData start');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName('社員マスタ');
    var header = ss.getSheetByName('評価ヘッダDB');
    if (!master || !header) throw new Error('必要なシートが見つかりません。');

    var masterData = master.getDataRange().getValues();
    var headerData = header.getDataRange().getValues();

    var loginEmail = '';
    try { loginEmail = (Session.getActiveUser && Session.getActiveUser().getEmail) ? Session.getActiveUser().getEmail() : ''; } catch (e) { loginEmail = ''; }
    if (!loginEmail) {
      try { loginEmail = (ScriptApp.getEffectiveUser && ScriptApp.getEffectiveUser().getEmail) ? ScriptApp.getEffectiveUser().getEmail() : ''; } catch (e) { loginEmail = ''; }
    }
    var loginEmailLower = (loginEmail || '').toLowerCase();
    Logger.log('getDashboardData called by %s', loginEmailLower);

    var loginUser = findEmployeeByEmail(masterData, loginEmailLower);
    if (!loginUser) {
      Logger.log('getDashboardData: login user not found: %s', loginEmailLower);
      return { success: true, data: { myTask: null, subordinateTasks: [] } };
    }
    var loginUserId = loginUser.id;

    var myTask = null;
    var evalueeIdCol = 4, statusCol = 5;
    for (var i = 1; i < headerData.length; i++) {
      var evalueeId = headerData[i][evalueeIdCol];
      var status = (headerData[i][statusCol] || '').toString();
      if (evalueeId === loginUserId && status !== '6_完了') {
        var required = (status === '1_本人入力中') ? '本人入力' : '評価者待ち';
        var isPending = (status !== '1_本人入力中');
        myTask = { evaluationId: headerData[i][0], period: headerData[i][1], status: status, requiredAction: required, isPending: isPending };
        break;
      }
    }

    var subordinateTasks = [];
    var subordinateIds = getSubordinates(headerData, loginUserId, masterData);
    if (subordinateIds.length > 0) {
      for (var j = 1; j < headerData.length; j++) {
        var evalueeId2 = headerData[j][evalueeIdCol];
        if (subordinateIds.indexOf(evalueeId2) !== -1) {
          var status2 = (headerData[j][statusCol] || '').toString();
          if (status2 === '6_完了') continue;
          var evaluee = findEmployeeById(masterData, evalueeId2);
          if (!evaluee) continue;
          var requiredAction = null;
          if (status2 === '2_評価者1入力中' && evaluee.eval1Id === loginUserId) requiredAction = '評価者1入力';
          else if (status2 === '3_評価者2入力中' && evaluee.eval2Id === loginUserId) requiredAction = '評価者2入力';
          else if (status2 === '4_評価者3入力中' && evaluee.eval3Id === loginUserId) requiredAction = '評価者3入力';
          if (requiredAction) subordinateTasks.push({ evaluationId: headerData[j][0], period: headerData[j][1], evalueeName: evaluee.name, status: status2, requiredAction: requiredAction });
        }
      }
    }

    return { success: true, data: { myTask: myTask, subordinateTasks: subordinateTasks } };

  } catch (error) {
    Logger.log('getDashboardData error: %s\n%s', error && error.message ? error.message : String(error), error && error.stack ? error.stack : '');
    return { success: false, message: 'ダッシュボードデータの取得に失敗しました: ' + (error && error.message ? error.message : String(error)) };
  }
}

/* ========== 評価データ取得 ========== */
function getEvaluationData(evaluationId, loginEmail) {
  evaluationId = (typeof evaluationId !== 'undefined' && evaluationId !== null) ? String(evaluationId).trim() : '';
  loginEmail = (typeof loginEmail !== 'undefined' && loginEmail !== null) ? String(loginEmail).toLowerCase().trim() : '';
  try {
    Logger.log('getEvaluationData start. evaluationId=%s, loginEmail=%s', evaluationId, loginEmail);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName('社員マスタ');
    var headerSheet = ss.getSheetByName('評価ヘッダDB');
    var detailSheet = ss.getSheetByName('評価明細DB');
    var commonMasterSheet = ss.getSheetByName('評価項目マスタ_共通');
    var gradeMasterSheet = ss.getSheetByName('評価項目マスタ_等級別');

    if (!masterSheet || !headerSheet || !detailSheet || !commonMasterSheet || !gradeMasterSheet) {
      var msg = '必要なシートが見つかりません。';
      Logger.log('getEvaluationData error: %s', msg);
      return { success: false, message: msg };
    }

    var headerData = headerSheet.getDataRange().getValues();
    var headerRow = null;
    var headerIndexMap = buildHeaderIndexMap(headerSheet);

    for (var i = 1; i < headerData.length; i++) {
      var id = headerData[i][0] ? String(headerData[i][0]).trim() : '';
      if (id === evaluationId) {
        var row = headerData[i];
        function rv(name, fallbackIndex) {
          var v = getRowValueByHeader(row, headerIndexMap, name);
          if (v !== '') return v;
          return (typeof fallbackIndex === 'number' && row.length > fallbackIndex) ? row[fallbackIndex] : '';
        }
        headerRow = {
          evaluationId: rv('評価ID', 0),
          period: rv('評価期間', 1),
          periodFrom: rv('対象期間From', 2),
          periodTo: rv('対象期間To', 3),
          evalueeId: rv('被評価者ID', 4),
          status: rv('ステータス', 5),
          comments: {
            evaluee: rv('評価者1所見', 8) || '',
            eval1: rv('評価者1所見', 9) || '',
            eval2: rv('評価者2所見', 10) || '',
            eval3: rv('評価者3所見', 11) || ''
          },
          goals: {
            'A-1': { goal: rv('A-1_目標', 12) || '', result: rv('A-1_結果', 13) || '' },
            'A-2': { goal: rv('A-2_目標', 14) || '', result: rv('A-2_結果', 15) || '' }
          },
          presidentScore: rv('社長評価点', 7) || ''
        };
        break;
      }
    }

    if (!headerRow) {
      var msg2 = "評価ID '" + evaluationId + "' が評価ヘッダDBに見つかりません。";
      Logger.log('getEvaluationData error: %s', msg2);
      return { success: false, message: msg2 };
    }

    if (!headerRow.evalueeId || String(headerRow.evalueeId).trim() === '') {
      Logger.log("getEvaluationData error: evalueeId empty for header=%s", JSON.stringify(headerRow));
      return { success: false, message: '評価ヘッダに被評価者IDが設定されていません。' };
    }

    var masterData = masterSheet.getDataRange().getValues();
    var employee = findEmployeeById(masterData, headerRow.evalueeId);
    if (!employee) {
      var msg3 = "社員ID '" + headerRow.evalueeId + "' が社員マスタに見つかりません。";
      Logger.log('getEvaluationData error: %s', msg3);
      return { success: false, message: msg3 };
    }

    var eval1 = findEmployeeById(masterData, employee.eval1Id);
    var eval2 = findEmployeeById(masterData, employee.eval2Id);
    var eval3 = findEmployeeById(masterData, employee.eval3Id);

    headerRow.evalueeName = employee.name || '';
    headerRow.evalueeDepartment = employee.department || '';
    headerRow.evalueeGender = employee.gender || '';
    headerRow.evalueeDob = employee.dob || '';
    headerRow.evalueeJoined = employee.joined || '';
    headerRow.evalueeGrade = employee.grade || '';
    headerRow.evalueeNumber = employee.number || '';
    headerRow.eval1Id = employee.eval1Id || '';
    headerRow.eval1Name = (eval1 && eval1.name) ? eval1.name : ' (未設定)';
    headerRow.eval2Id = employee.eval2Id || '';
    headerRow.eval2Name = (eval2 && eval2.name) ? eval2.name : ' (未設定)';
    headerRow.eval3Id = employee.eval3Id || '';
    headerRow.eval3Name = (eval3 && eval3.name) ? eval3.name : ' (未設定)';

    var commonMaster = commonMasterSheet.getDataRange().getValues();
    var gradeMaster = gradeMasterSheet.getDataRange().getValues();

    var commonMap = {}, gradeMap = {};
    if (Array.isArray(commonMaster)) {
      for (var cm = 1; cm < commonMaster.length; cm++) {
        var r = commonMaster[cm];
        if (!r || !r[0]) continue;
        commonMap[r[0]] = { id: r[0], category: r[1], subCategory: r[2], item: r[3], description: r[4], maxScore: r[5] != null ? r[5] : 10, isGoal: (String(r[0]).startsWith('A-')) };
      }
    }
    if (Array.isArray(gradeMaster)) {
      for (var gm = 1; gm < gradeMaster.length; gm++) {
        var gr = gradeMaster[gm];
        if (!gr || !gr[1]) continue;
        gradeMap[gr[1]] = { id: gr[1], grade: gr[0], item: gr[2], description: gr[2] || '', maxScore: gr[3] != null ? gr[3] : 10, category: '2.役割評価', subCategory: '等級別役割', isGoal: false };
      }
    }

    var items = [], itemsIndex = {};
    if (Array.isArray(commonMaster)) {
      for (var cm2 = 1; cm2 < commonMaster.length; cm2++) {
        var rr = commonMaster[cm2];
        if (!rr || !rr[0]) continue;
        var obj = { id: rr[0], category: rr[1], subCategory: rr[2], item: rr[3], description: rr[4], maxScore: rr[5] != null ? rr[5] : 10, isGoal: String(rr[0]).startsWith('A-') };
        items.push(obj); itemsIndex[obj.id] = true;
      }
    }
    var evalueeGrade = (employee.grade || '').toString();
    if (Array.isArray(gradeMaster)) {
      for (var g2 = 1; g2 < gradeMaster.length; g2++) {
        var rowg = gradeMaster[g2];
        var gradeKey = String(rowg[0] || '');
        var itemId = rowg[1];
        if (!itemId) continue;
        if (gradeKey === evalueeGrade && !itemsIndex[itemId]) {
          var obj2 = { id: itemId, category: '2.役割評価', subCategory: '等級別役割', item: rowg[2], description: rowg[2] || '', maxScore: rowg[3] != null ? rowg[3] : 10, isGoal: false };
          items.push(obj2); itemsIndex[itemId] = true;
        }
      }
    }

    var detailData = detailSheet.getDataRange().getValues();
    var details = {};
    if (Array.isArray(detailData)) {
      for (var d = 1; d < detailData.length; d++) {
        var rowd = detailData[d] || [];
        if (String(rowd[0] || '') !== evaluationId) continue;
        var itemId = rowd[1];
        if (!itemId) continue;
        details[itemId] = {
          score: {
            evaluee: (rowd[2] != null && rowd[2] !== '') ? rowd[2] : null,
            eval1: (rowd[4] != null && rowd[4] !== '') ? rowd[4] : null,
            eval2: (rowd[6] != null && rowd[6] !== '') ? rowd[6] : null,
            eval3: (rowd[8] != null && rowd[8] !== '') ? rowd[8] : null
          },
          achievement: {
            evaluee: (rowd[3] != null && rowd[3] !== '') ? rowd[3] : '-',
            eval1: (rowd[5] != null && rowd[5] !== '') ? rowd[5] : '-',
            eval2: (rowd[7] != null && rowd[7] !== '') ? rowd[7] : '-',
            eval3: (rowd[9] != null && rowd[9] !== '') ? rowd[9] : '-'
          }
        };
        if (!itemsIndex[itemId]) {
          var meta = gradeMap[itemId] || commonMap[itemId] || null;
          var obj;
          if (meta) {
            obj = { id: itemId, category: meta.category || '2.役割評価', subCategory: meta.subCategory || '', item: meta.item || itemId, description: meta.description || '', maxScore: meta.maxScore != null ? meta.maxScore : 10, isGoal: !!meta.isGoal };
          } else {
            obj = { id: itemId, category: '2.役割評価', subCategory: '等級別(明細)', item: itemId, description: '', maxScore: 10, isGoal: false };
          }
          items.push(obj); itemsIndex[itemId] = true;
        }
      }
    }

    var userRole = (_isAdminEmail(loginEmail) ? 'Admin' : determineUserRole(headerRow, employee, loginEmail));
    var result = { success: true, data: { header: headerRow, items: items, details: details, loggedIn: { email: loginEmail, role: userRole } } };
    var safe = _normalizeForRpc(result.data);
    Logger.log('getEvaluationData finished evaluationId=%s items=%s', evaluationId, items.length);
    return { success: true, data: safe };
  } catch (error) {
    Logger.log('getEvaluationData exception: %s\n%s', error && error.message ? error.message : String(error), error && error.stack ? error.stack : '');
    return { success: false, message: '評価データの読み込みに失敗しました: ' + (error && error.message ? error.message : String(error)) };
  }
}

/* ========== 評価データ保存 ========== */
function saveEvaluationData(userRole, isSubmit, data) {
  try {
    Logger.log('saveEvaluationData start role=%s isSubmit=%s', userRole, !!isSubmit);
    if (!data || typeof data !== 'object') throw new Error('不正なデータです。');
    var evaluationId = data.evaluationId || '';
    var details = data.details || {};
    var comments = data.comments || {};
    var goals = data.goals || {};
    var presidentScore = data.presidentScore;
    if (!evaluationId) throw new Error('評価IDがありません。');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var headerSheet = ss.getSheetByName('評価ヘッダDB');
    var detailSheet = ss.getSheetByName('評価明細DB');
    if (!headerSheet || !detailSheet) throw new Error('DBシートが見つかりません。');

    var headerIndexMap = buildHeaderIndexMap(headerSheet);

    var headerData = headerSheet.getDataRange().getValues();
    var headerRowNum = -1, currentStatus = '', evalueeId = '';
    for (var i = 1; i < headerData.length; i++) {
      if ((headerData[i][0] || '').toString().trim() === (evaluationId || '').toString().trim()) {
        headerRowNum = i + 1;
        currentStatus = headerData[i][5];
        evalueeId = headerData[i][4];
        break;
      }
    }
    if (headerRowNum === -1) throw new Error('評価ヘッダが見つかりません: ' + evaluationId);

    // robust setHeaderCellByName using normalization and headerIndexMap
    function setHeaderCellByName(rowNum, headerName, value) {
      if (!headerName || headerName === '') return false;
      try {
        var lookupKey = _normalizeHeaderNameForKey(headerName);
        var idx = headerIndexMap[lookupKey];
        if (typeof idx === 'number') {
          headerSheet.getRange(rowNum, idx + 1).setValue(value);
          return true;
        }
        var alt1 = lookupKey.replace(/[-_]/g, '');
        if (alt1 && headerIndexMap[alt1] !== undefined) {
          headerSheet.getRange(rowNum, headerIndexMap[alt1] + 1).setValue(value);
          return true;
        }
        var alt2 = lookupKey.replace(/-/g, '_');
        if (alt2 && headerIndexMap[alt2] !== undefined) {
          headerSheet.getRange(rowNum, headerIndexMap[alt2] + 1).setValue(value);
          return true;
        }
        var alt3 = lookupKey.replace(/_/g, '-');
        if (alt3 && headerIndexMap[alt3] !== undefined) {
          headerSheet.getRange(rowNum, headerIndexMap[alt3] + 1).setValue(value);
          return true;
        }
        return false;
      } catch (e) {
        Logger.log('setHeaderCellByName error for header=%s : %s', headerName, e && e.stack ? e.stack : String(e));
        return false;
      }
    }

    if (userRole === 'Evaluee') {
      setHeaderCellByName(headerRowNum, '評価者1所見', comments.eval1 || '');
    } else if (userRole === 'Evaluator1') {
      setHeaderCellByName(headerRowNum, '評価者1所見', comments.eval1 || '');
    } else if (userRole === 'Evaluator2') {
      setHeaderCellByName(headerRowNum, '評価者2所見', comments.eval2 || '');
    } else if (userRole === 'Evaluator3') {
      setHeaderCellByName(headerRowNum, '評価者3所見', comments.eval3 || '');
    } else if (userRole === 'Admin') {
      setHeaderCellByName(headerRowNum, '評価者1所見', comments.eval1 || '');
      setHeaderCellByName(headerRowNum, '評価者2所見', comments.eval2 || '');
      setHeaderCellByName(headerRowNum, '評価者3所見', comments.eval3 || '');
    }

    // goals 保存（ヘッダ名を正規化して探す）
    setHeaderCellByName(headerRowNum, 'A-1_目標', (goals['A-1'] && goals['A-1'].goal) || '');
    setHeaderCellByName(headerRowNum, 'A-1_結果', (goals['A-1'] && goals['A-1'].result) || '');
    setHeaderCellByName(headerRowNum, 'A-2_目標', (goals['A-2'] && goals['A-2'].goal) || '');
    setHeaderCellByName(headerRowNum, 'A-2_結果', (goals['A-2'] && goals['A-2'].result) || '');

    if (userRole === 'Admin') {
      setHeaderCellByName(headerRowNum, '社長評価点', presidentScore || '');
    }

    if (isSubmit) {
      var nextStatus = '';
      if (currentStatus === '1_本人入力中' && userRole === 'Evaluee') nextStatus = '2_評価者1入力中';
      else if (currentStatus === '2_評価者1入力中' && userRole === 'Evaluator1') nextStatus = '3_評価者2入力中';
      else if (currentStatus === '3_評価者2入力中' && userRole === 'Evaluator2') nextStatus = '4_評価者3入力中';
      else if (currentStatus === '4_評価者3入力中' && userRole === 'Evaluator3') nextStatus = '5_最終確認中';
      else if (currentStatus === '5_最終確認中' && userRole === 'Admin') nextStatus = '6_完了';
      if (nextStatus) headerSheet.getRange(headerRowNum, 6).setValue(nextStatus);
    }

    // detail の更新・追加
    var detailData = detailSheet.getDataRange().getValues();
    var existingRows = {};
    for (var r = 1; r < detailData.length; r++) {
      var row = detailData[r] || [];
      if ((row[0] || '').toString() === (evaluationId || '').toString()) {
        var itemId = row[1] != null ? row[1].toString() : '';
        if (itemId) existingRows[itemId] = r + 1;
      }
    }

    var updates = [];
    for (var item in details) {
      if (!details.hasOwnProperty(item)) continue;
      var d = details[item] || {};
      var sc = d.score || {};
      var ach = d.achievement || {};
      var eScore = sc.evaluee != null ? sc.evaluee : '';
      var eAchi = ach.evaluee != null ? ach.evaluee : '-';
      var v1 = sc.eval1 != null ? sc.eval1 : '';
      var a1 = ach.eval1 != null ? ach.eval1 : '-';
      var v2 = sc.eval2 != null ? sc.eval2 : '';
      var a2 = ach.eval2 != null ? ach.eval2 : '-';
      var v3 = sc.eval3 != null ? sc.eval3 : '';
      var a3 = ach.eval3 != null ? ach.eval3 : '-';

      if (existingRows[item]) {
        var rowNum = existingRows[item];
        // 修正: 書き込み開始列を 1 (A列) に変更（列ずれ防止）
        detailSheet.getRange(rowNum, 1, 1, 10).setValues([[ evaluationId, item, eScore, eAchi, v1, a1, v2, a2, v3, a3 ]]);
      } else {
        updates.push([ evaluationId, item, eScore, eAchi, v1, a1, v2, a2, v3, a3 ]);
      }
    }
    if (updates.length > 0) {
      var startRow = Math.max(detailSheet.getLastRow(), 1) + 1;
      detailSheet.getRange(startRow, 1, updates.length, updates[0].length).setValues(updates);
    }

    Logger.log('saveEvaluationData finished for evaluationId=%s', evaluationId);
    return { success: true };
  } catch (error) {
    Logger.log('saveEvaluationData error: %s\n%s', error && error.message ? error.message : String(error), error && error.stack ? error.stack : '');
    return { success: false, message: '保存に失敗しました: ' + (error && error.message ? error.message : String(error)) };
  }
}

/* ========== 管理用 / 補助 ========== */
function getAdminPageData() {
  try {
    var dashboard = getAdminDashboardData();
    var stats = getAdminStatsData();
    var master = getMasterData();
    if (!dashboard.success) throw new Error(dashboard.message);
    if (!stats.success) throw new Error(stats.message);
    if (!master.success) throw new Error(master.message);
    return { success: true, dashboardData: dashboard.dashboardData, statsData: stats.statsData, masterData: master.masterData };
  } catch (e) {
    Logger.log('getAdminPageData error: %s', e && e.stack ? e.stack : String(e));
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}

function getAdminStatsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var header = ss.getSheetByName('評価ヘッダDB');
    if (!header) throw new Error('評価ヘッダDBが見つかりません。');
    var h = header.getDataRange().getValues();
    var statusCol = 5;
    var total = 0, completed = 0, pending = 0, counts = {};
    for (var i = 1; i < h.length; i++) {
      var s = (h[i][statusCol] || '').toString().trim();
      if (!s) continue;
      total++;
      if (s === '6_完了') completed++; else pending++;
      counts[s] = (counts[s] || 0) + 1;
    }
    return { success: true, statsData: { totalCount: total, completedCount: completed, pendingCount: pending, statusCounts: counts } };
  } catch (e) {
    Logger.log('getAdminStatsData error: %s', e && e.stack ? e.stack : String(e));
    return { success: false, message: String(e) };
  }
}

function getAdminDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName('社員マスタ');
    var header = ss.getSheetByName('評価ヘッダDB');
    if (!master || !header) throw new Error('必要なシートが見つかりません。');
    var m = master.getDataRange().getValues();
    var h = header.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < h.length; i++) {
      var evalueeId = h[i][4];
      var status = (h[i][5] || '').toString().trim();
      if (!status) continue;
      var emp = findEmployeeById(m, evalueeId);
      out.push({ evaluationId: h[i][0] || '', period: h[i][1] || '', employeeId: evalueeId, employeeName: emp ? emp.name : '不明', grade: emp ? emp.grade : '-', status: status, totalScore: h[i][6] || '', totalRank: h[i][7] || '' });
    }
    return { success: true, dashboardData: out };
  } catch (e) {
    Logger.log('getAdminDashboardData error: %s', e && e.stack ? e.stack : String(e));
    return { success: false, message: String(e) };
  }
}

function getMasterData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('社員マスタ');
    if (!sheet) throw new Error('社員マスタが見つかりません。');
    var vals = sheet.getDataRange().getValues();
    if (vals.length <= 1) return { success: true, masterData: [] };
    var rows = vals.slice(1);
    var out = rows.map(function(r){ return {
      employeeId: r[0] || '', name: r[1] || '', email: (r[2] || '').toString().toLowerCase(), department: r[3] || '', gender: r[4] || '', dob: r[5] ? Utilities.formatDate(new Date(r[5]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '', joined: r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '', grade: r[7] || '', number: r[8] || '', eval1Id: r[9] || '', eval2Id: r[10] || '', eval3Id: r[11] || '', status: r[12] || '' }; });
    return { success: true, masterData: out };
  } catch (e) {
    Logger.log('getMasterData error: %s', e && e.stack ? e.stack : String(e));
    return { success: false, message: String(e) };
  }
}

/* ========== デバッグ / ヘルスチェック ========== */
function debugHeaderIndexMap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var header = ss.getSheetByName('評価ヘッダDB');
  if (!header) { Logger.log('評価ヘッダDB が見つかりません'); return; }
  var headerRow = header.getRange(1,1,1, header.getLastColumn()).getValues()[0];
  Logger.log('headerRow raw: %s', JSON.stringify(headerRow));
  var map = buildHeaderIndexMap(header);
  // convert to 1-based col for readability
  var pretty = {};
  for (var k in map) { if (map.hasOwnProperty(k)) pretty[k] = map[k] + 1; }
  Logger.log('normalized header map (key -> 1-based col): %s', JSON.stringify(pretty));
  var want = _normalizeHeaderNameForKey('A-1_目標');
  Logger.log('lookup key for "A-1_目標" => %s', want);
  Logger.log('col for A-1_目標 = %s', pretty[want] || '(not found)');
}

function _diagnoseHealth(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var page = params.page || '';
    var id = params.id || params.evaluationId || '';
    var execUrl = '';
    try { execUrl = (ScriptApp.getService) ? ScriptApp.getService().getUrl() : ''; } catch (e2) { execUrl = ''; }
    var info = { now: (new Date()).toISOString(), pageParam: page, idParam: id, userEmail: (Session && Session.getActiveUser) ? Session.getActiveUser().getEmail() : '(no session)', execUrl: execUrl, templatesAvailable: [] };
    var tryFiles = ['評価入力画面','管理ダッシュボード','メインダッシュボード'];
    tryFiles.forEach(function(name){
      try { HtmlService.createTemplateFromFile(name); info.templatesAvailable.push({ name: name, ok: true }); }
      catch (ex) { info.templatesAvailable.push({ name: name, ok: false, error: ex && ex.message ? ex.message : String(ex) }); }
    });
    return HtmlService.createHtmlOutput('<pre>' + JSON.stringify(info, null, 2) + '</pre>');
  } catch (e) {
    return HtmlService.createHtmlOutput('<pre>diagnose failed: ' + (e && e.stack ? e.stack : String(e)) + '</pre>');
  }
}

/* ========== 社員詳細モーダル ========== */
function getEmployeeModalHtml(employeeId) {
  try {
    if (!employeeId) return { success: false, message: '社員IDが指定されていません。' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('社員マスタ');
    if (!sheet) return { success: false, message: '社員マスタが見つかりません。' };
    var vals = sheet.getDataRange().getValues();
    var row = null;
    for (var i = 1; i < vals.length; i++) { if ((vals[i][0] || '') == employeeId) { row = vals[i]; break; } }
    if (!row) return { success: false, message: '社員が見つかりません (ID:' + employeeId + ')' };
    function fmt(cell){ try { if (!cell && cell !== 0) return ''; if (Object.prototype.toString.call(cell) === '[object Date]') return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy/MM/dd'); return String(cell); } catch(e){ return String(cell||''); } }
    var emp = { id: row[0]||'', name: row[1]||'', email: row[2]||'', department: row[3]||'', gender: row[4]||'', dob: fmt(row[5]), joined: fmt(row[6]), grade: row[7]||'', number: row[8]||'', eval1Id: row[9]||'', eval2Id: row[10]||'', eval3Id: row[11]||'', status: row[12]||'' };
    var html = '<div style="font-family:Arial,sans-serif;">' +
               '<table style="width:100%;border-collapse:collapse;">' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;width:30%;">社員ID</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.id) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;">氏名</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.name) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;">メール</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.email) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;">部署</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.department) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;">等級</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.grade) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;border-bottom:1px solid #eee;">生年月日</th><td style="padding:6px;border-bottom:1px solid #eee;">' + (emp.dob) + '</td></tr>' +
               '<tr><th style="text-align:left;padding:6px;">在籍</th><td style="padding:6px;">' + (emp.status) + '</td></tr>' +
               '</table></div>';
    return { success: true, html: html };
  } catch (e) {
    Logger.log('getEmployeeModalHtml error: %s', e && e.stack ? e.stack : String(e));
    return { success: false, message: '社員情報の取得に失敗しました: ' + (e && e.message ? e.message : String(e)) };
  }
}

/* ========== 補助: determineUserRole ========== */
function determineUserRole(headerRow, employee, loginEmail) {
  try {
    if (!employee || !loginEmail) return 'Unknown';
    if (_isAdminEmail(loginEmail)) return 'Admin';
    return 'Employee';
  } catch (e) {
    return 'Unknown';
  }
}