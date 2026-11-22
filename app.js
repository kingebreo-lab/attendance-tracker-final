const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: 'Attendance',
  EMPLOYEES_SHEET_NAME: 'Employees',
  SELFIE_FOLDER_ID: '',
  SELFIE_FOLDER_NAME: 'Attendance Selfies',
  SELFIE_RETENTION_DAYS: 30,
  // Store your actual key in Project Settings > Script Properties as GEMINI_API_KEY
  GEMINI_API_KEY: '',
  GEMINI_MODELS: ['gemini-3.5-flash-lite', 'gemini-3.8-flash', 'gemini-3.6-flash'],
  ENFORCE_AI_SELFIE_CHECK: true,
  ALLOWED_LOCATION_LABEL: 'SM City Batangas Premises & Parking Lots (Pallocan Kanluran, Batangas City)',
  ALLOWED_CENTER_LAT: 13.7554,
  ALLOWED_CENTER_LNG: 121.0695,
  ALLOWED_RADIUS_METERS: 350,
  TIME_ZONE: 'Asia/Manila',
  ADMIN_PIN: 'Globe',
  WORK_START: '09:00',
  WORK_END: '18:00'
};

function setup() {
  const props = PropertiesService.getScriptProperties();

  if (CONFIG.GEMINI_API_KEY) {
    props.setProperty('GEMINI_API_KEY', CONFIG.GEMINI_API_KEY.trim());
  }

  let id = CONFIG.SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID');
  let ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('Employee Attendance Responses');
    id = ss.getId();
    props.setProperty('SPREADSHEET_ID', id);
  }

  let attendance = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!attendance) attendance = ss.insertSheet(CONFIG.SHEET_NAME);
  const headers = [
    'Date',
    'Employee ID',
    'Employee Name',
    'Role for the Day',
    'Time In',
    'Time Out',
    'Remarks',
    'Last Updated',
    'Time In Selfie',
    'Time In Location',
    'Time Out Selfie',
    'Time Out Location',
    'AI Verification',
    'Privacy Consent'
  ];
  if (attendance.getLastRow() === 0) {
    attendance.appendRow(headers);
  } else if (attendance.getLastColumn() < headers.length) {
    attendance.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  attendance.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  attendance.setFrozenRows(1);
  attendance.autoResizeColumns(1, headers.length);

  let employees = ss.getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
  if (!employees) employees = ss.insertSheet(CONFIG.EMPLOYEES_SHEET_NAME);
  if (employees.getLastRow() === 0) {
    employees.appendRow(['Employee ID', 'Employee Name', 'Active', 'Employee PIN']);
  } else if (employees.getLastColumn() < 4) {
    employees.getRange(1, 4).setValue('Employee PIN');
  }

  const suppliedRoster = [
    ['10010607', 'Ailene Jaen', 'YES', '1234'], ['10004979', 'Andrew Mer Pascual', 'YES', '1234'], ['10007078', 'Anna Katrina Guarino', 'YES', '1234'],
    ['10001141', 'Genelynn Dimple Galvero', 'YES', '1234'], ['10002890', 'Janice De Castro', 'YES', '1234'], ['10002174', 'Jeromme Omar Paras', 'YES', '1234'],
    ['10007572', 'Jonald Garcia', 'YES', '1234'], ['10009386', 'Jyzhell Laydia', 'YES', '1234'], ['10002790', 'Karen Molinyawe', 'YES', '1234'],
    ['10022774', 'King Gerald Ebreo', 'YES', '1234'], ['10010068', 'Ma. Mylin Patrick Bautista', 'YES', '1234'], ['10010286', 'Mario Dolor', 'YES', '1234'],
    ['10001479', 'Mellow Jane Kalaw', 'YES', '1234'], ['10014620', 'Rebejean Mercado', 'YES', '1234'], ['10000036', 'Renelyn Bay', 'YES', '1234'],
    ['10020056', 'Renz Harlee Adame', 'YES', '1234'], ['10009504', 'Ronald Ebreo', 'YES', '1234'], ['RSM4781', 'Ruby Mangubat', 'YES', '1234']
  ];
  const existing = employees.getLastRow() > 1 ? employees.getRange(2, 1, employees.getLastRow() - 1, 4).getDisplayValues() : [];
  const rowById = {};
  existing.forEach((r, i) => {
    const key = String(r[0] || '').trim().toUpperCase();
    if (key) rowById[key] = { row: i + 2, active: r[2] || 'YES', pin: r[3] || '' };
  });
  suppliedRoster.forEach(person => {
    const key = person[0].toUpperCase();
    if (rowById[key]) {
      employees.getRange(rowById[key].row, 2).setValue(person[1]);
      if (!String(rowById[key].active).trim()) employees.getRange(rowById[key].row, 3).setValue('YES');
      if (!String(rowById[key].pin).trim()) employees.getRange(rowById[key].row, 4).setValue(person[3]);
    } else {
      employees.appendRow(person);
    }
  });
  employees.getRange(1, 1, 1, 4).setFontWeight('bold');
  employees.setFrozenRows(1);
  employees.autoResizeColumns(1, 4);

  getSelfieFolder_();
  ensureCleanupTrigger_();
  return ss.getUrl();
}

// ==========================================
// STANDALONE WEB APP REST + JSONP API
// ==========================================

function routeApiRequest_(action, args) {
  setupIfNeeded_();
  switch (String(action || '').trim()) {
    case 'ping':
      return { status: 'Attendance API Online', time: new Date().toISOString() };
    case 'lookupEmployee':
      return lookupEmployee(args[0]);
    case 'preVerifySelfie':
      return preVerifySelfie(args[0]);
    case 'submitAttendance':
      return submitAttendance(args[0]);
    case 'getEmployeeDashboard':
      return getEmployeeDashboard(args[0], args[1], args[2]);
    case 'getResponses':
      return getResponses(args[0]);
    case 'getSpreadsheetUrl':
      return getSpreadsheetUrl(args[0]);
    case 'getEmployeesList':
      return getEmployeesList(args[0]);
    case 'addNewEmployee':
      return addNewEmployee(args[0], args[1]);
    case 'updateEmployeeDetails':
      return updateEmployeeDetails(args[0], args[1]);
    default:
      throw new Error('Unknown API action: ' + action);
  }
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = String(params.callback || '').replace(/[^a-zA-Z0-9_$.]/g, '');

  try {
    const action = params.action || 'ping';
    const args = params.args ? JSON.parse(params.args) : [];
    const result = routeApiRequest_(action, Array.isArray(args) ? args : []);
    const payload = JSON.stringify({ ok: true, data: result });

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + payload + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const errPayload = JSON.stringify({ ok: false, error: err.message || String(err) });
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + errPayload + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(errPayload)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const body = JSON.parse(raw);
    const result = routeApiRequest_(body.action, Array.isArray(body.args) ? body.args : []);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  if (CONFIG.GEMINI_API_KEY && props.getProperty('GEMINI_API_KEY') !== CONFIG.GEMINI_API_KEY.trim()) {
    props.setProperty('GEMINI_API_KEY', CONFIG.GEMINI_API_KEY.trim());
  }
  const id = props.getProperty('SPREADSHEET_ID') || CONFIG.SPREADSHEET_ID;
  if (!id) { setup(); return; }
  const ss = SpreadsheetApp.openById(id);
  if (!ss.getSheetByName(CONFIG.SHEET_NAME) || !ss.getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME)) setup();
}

// ==========================================
// DATA RETENTION: AUTO-CLEANUP OLD SELFIES
// ==========================================

function ensureCleanupTrigger_() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some(t => t.getHandlerFunction() === 'cleanupOldSelfies');
    if (!exists) {
      ScriptApp.newTrigger('cleanupOldSelfies')
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .create();
    }
  } catch (e) {
    Logger.log('Could not auto-create trigger: ' + e.message);
  }
}

function cleanupOldSelfies() {
  const folder = getSelfieFolder_();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - CONFIG.SELFIE_RETENTION_DAYS);

  const files = folder.getFiles();
  let trashedCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoff) {
      file.setTrashed(true);
      trashedCount++;
    }
  }
  Logger.log('Trashed ' + trashedCount + ' selfie(s) older than ' + CONFIG.SELFIE_RETENTION_DAYS + ' days.');
}

// ==========================================
// GEOFENCE & LOCATION VALIDATION
// ==========================================

function calculateDistanceMeters_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function isWithinAllowedArea_(lat, lng, accuracy, resolvedAddress) {
  const dist = calculateDistanceMeters_(lat, lng, CONFIG.ALLOWED_CENTER_LAT, CONFIG.ALLOWED_CENTER_LNG);
  const accuracyBuffer = Math.min(75, Math.max(0, Number(accuracy) || 0));
  const withinRadius = dist <= (CONFIG.ALLOWED_RADIUS_METERS + accuracyBuffer);

  if (withinRadius) {
    return { allowed: true, distanceMeters: Math.round(dist) };
  }

  const addrLower = String(resolvedAddress || '').toLowerCase();
  const inBatangas = addrLower.includes('batangas');
  const matchesMallPremises =
    addrLower.includes('sm city batangas') ||
    addrLower.includes('sm batangas');

  return {
    allowed: inBatangas && matchesMallPremises && dist <= 500,
    distanceMeters: Math.round(dist)
  };
}

// ==========================================
// FAST GEMINI AI PERSON VERIFICATION
// ==========================================

function getGeminiKey_() {
  const props = PropertiesService.getScriptProperties();
  const propKey = String(props.getProperty('GEMINI_API_KEY') || '').trim();
  const configKey = String(CONFIG.GEMINI_API_KEY || '').trim();
  const placeholders = [
    '',
    'PASTE-YOUR-AIzaSy-GEMINI-API-KEY-HERE',
    'AIzaSyYourActualKeyHere'
  ];

  if (!placeholders.includes(configKey)) {
    if (propKey !== configKey) {
      try { props.setProperty('GEMINI_API_KEY', configKey); } catch (e) {}
    }
    return configKey;
  }
  if (!placeholders.includes(propKey)) return propKey;
  return '';
}

function computeSelfieHash_(base64Body) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base64Body);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function callGeminiFast_(parts, responseMimeType) {
  const apiKey = getGeminiKey_();
  if (!apiKey) return null;

  const models = Array.isArray(CONFIG.GEMINI_MODELS) && CONFIG.GEMINI_MODELS.length
    ? CONFIG.GEMINI_MODELS
    : ['gemini-3.5-flash-lite', 'gemini-3.8-flash'];

  let lastError = '';

  for (let m = 0; m < models.length; m++) {
    const model = String(models[m] || '').trim();
    if (!model) continue;

    const payload = {
      contents: [{ parts: parts }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 120
      }
    };
    if (responseMimeType) {
      payload.generationConfig.responseMimeType = responseMimeType;
    }
    if (model.startsWith('gemini-3')) {
      payload.generationConfig.thinkingConfig = { thinkingLevel: 'LOW' };
    }

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(model) + ':generateContent';

    for (let attempt = 1; attempt <= 2; attempt++) {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': apiKey },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = res.getResponseCode();
      if (code === 200) {
        const json = JSON.parse(res.getContentText());
        const candidate = json && json.candidates && json.candidates[0];
        const text = candidate && candidate.content && candidate.content.parts &&
          candidate.content.parts[0] && candidate.content.parts[0].text;
        if (text) return text;
      }

      lastError = 'AI service error (' + code + ' on ' + model + ')';
      if (code === 503 || code === 429) {
        if (attempt === 1) {
          Utilities.sleep(800);
          continue;
        }
        break;
      }
      if (code === 404 || code === 400) {
        break;
      }
      throw new Error(lastError);
    }
  }

  throw new Error(lastError || 'No compatible Gemini model responded.');
}

function preVerifySelfie(base64Data) {
  return verifySelfieWithAI_(base64Data);
}

function verifySelfieWithAI_(base64Data) {
  const apiKey = getGeminiKey_();
  if (!apiKey) {
    if (CONFIG.ENFORCE_AI_SELFIE_CHECK) {
      throw new Error('Person verification is enabled, but GEMINI_API_KEY is missing in Script Properties.');
    }
    return { verified: true, summary: 'Selfie captured (AI check skipped)' };
  }

  const parts = String(base64Data || '').split(',');
  if (parts.length < 2 || !parts[1]) {
    throw new Error('Invalid image format captured. Please retake the selfie.');
  }

  const mimeMatch = parts[0].match(/data:(image\/[a-zA-Z0-9.+-]+);base64/);
  const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
  const base64Body = parts[1];

  const cache = CacheService.getScriptCache();
  const hashKey = 'SELFIE_VERIFIED_' + computeSelfieHash_(base64Body).slice(0, 32);
  const cachedResult = cache.get(hashKey);
  if (cachedResult) {
    try { return JSON.parse(cachedResult); } catch (e) {}
  }

  const prompt =
    'Inspect this attendance selfie and return JSON: {"hasLivePerson": boolean, "reason": string}. ' +
    'hasLivePerson=true ONLY if a real live person with a visible face is taking the selfie. ' +
    'hasLivePerson=false if no person, object/wall/ceiling/floor, covered/unrecognizable face, or screen/print photo. ' +
    'Reason under 10 words.';

  try {
    const raw = callGeminiFast_([
      { text: prompt },
      { inlineData: { mimeType: mimeType, data: base64Body } }
    ], 'application/json');

    const cleaned = String(raw).replace(/```json/gi, '').replace(/```/g, '').trim();
    const parsed = JSON.parse(cleaned);

    if (!parsed.hasLivePerson) {
      throw new Error('Selfie Rejected: No live person detected (' + (parsed.reason || 'Please show your face clearly') + ').');
    }

    const result = {
      verified: true,
      summary: 'Person verified: ' + (parsed.reason || 'Live person detected')
    };
    cache.put(hashKey, JSON.stringify(result), 600);
    return result;
  } catch (err) {
    const msg = String(err.message || '');
    console.error('Selfie Verification Error:', err);

    if (msg.startsWith('Selfie Rejected:')) {
      throw err;
    }

    if (msg.includes('(503 ') || msg.includes('(429 ')) {
      const fallback = {
        verified: true,
        summary: 'Selfie saved (AI server busy)'
      };
      cache.put(hashKey, JSON.stringify(fallback), 300);
      return fallback;
    }

    if (CONFIG.ENFORCE_AI_SELFIE_CHECK) {
      throw new Error('Verification failed: ' + msg);
    }
    return { verified: true, summary: 'AI check unavailable' };
  }
}

// ==========================================
// EMPLOYEE LOOKUP (PIN ONLY ON DASHBOARD)
// ==========================================

function lookupEmployee(employeeId) {
  const id = String(employeeId || '').trim().toUpperCase();
  if (!id) return { found: false, message: 'Enter your Employee ID.' };

  const sheet = getSpreadsheet_().getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { found: false, message: 'Employee directory is empty. Contact the administrator.' };

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  for (const row of values) {
    if (String(row[0] || '').trim().toUpperCase() === id) {
      const name = String(row[1] || '').trim();
      const active = String(row[2] || 'YES').trim().toUpperCase();
      if (['NO', 'FALSE', 'INACTIVE'].includes(active)) {
        return { found: false, message: 'This Employee ID is inactive. Contact the administrator.' };
      }
      if (!name) {
        return { found: false, message: 'No name is assigned to this ID. Contact the administrator.' };
      }
      return { found: true, employeeId: row[0], employeeName: name };
    }
  }
  return { found: false, message: 'Employee ID not found. Please check your ID or contact the administrator.' };
}

function lookupEmployeeWithPin_(employeeId, employeePin) {
  const id = String(employeeId || '').trim().toUpperCase();
  const pin = String(employeePin || '').trim();
  if (!id) return { found: false, message: 'Enter your Employee ID.' };
  if (!pin) return { found: false, message: 'Enter your Employee PIN to view your records.' };

  const sheet = getSpreadsheet_().getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { found: false, message: 'Employee directory is empty. Contact the administrator.' };

  const cols = Math.max(4, sheet.getLastColumn());
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getDisplayValues();
  for (const row of values) {
    if (String(row[0] || '').trim().toUpperCase() === id) {
      const name = String(row[1] || '').trim();
      const active = String(row[2] || 'YES').trim().toUpperCase();
      const storedPin = String(row[3] || '').trim();

      if (['NO', 'FALSE', 'INACTIVE'].includes(active)) {
        return { found: false, message: 'This Employee ID is inactive. Contact the administrator.' };
      }
      if (!name) {
        return { found: false, message: 'No name is assigned to this ID. Contact the administrator.' };
      }
      if (!storedPin) {
        return { found: false, message: 'No PIN is set for this Employee ID. Contact the administrator.' };
      }
      if (storedPin !== pin) {
        return { found: false, message: 'Incorrect Employee PIN. Access denied.' };
      }
      return { found: true, employeeId: String(row[0]).trim(), employeeName: name };
    }
  }
  return { found: false, message: 'Employee ID not found. Please check your ID or contact the administrator.' };
}

function saveSelfieToDrive_(base64Data, dateStr, employeeId, action) {
  if (!base64Data || !base64Data.startsWith('data:image/')) {
    throw new Error('A live selfie photo taken by a person is required before submitting.');
  }
  const parts = base64Data.split(',');
  const mimeMatch = parts[0].match(/data:(image\/[a-zA-Z0-9.+-]+);base64/);
  const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
  const bytes = Utilities.base64Decode(parts[1]);
  const fileName = dateStr + '_' + employeeId + '_' + action + '.jpg';
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const folder = getSelfieFolder_();
  const file = folder.createFile(blob);
  return { url: file.getUrl(), file: file };
}

function getSelfieFolder_() {
  const props = PropertiesService.getScriptProperties();
  const fixedId = String(CONFIG.SELFIE_FOLDER_ID || '').trim();

  if (fixedId && fixedId !== 'PASTE-YOUR-GOOGLE-DRIVE-FOLDER-ID-HERE') {
    try {
      return DriveApp.getFolderById(fixedId);
    } catch (e) {
      throw new Error('Configured SELFIE_FOLDER_ID is invalid or inaccessible: ' + fixedId);
    }
  }

  let savedId = props.getProperty('SELFIE_FOLDER_ID');
  if (savedId) {
    try {
      return DriveApp.getFolderById(savedId);
    } catch (e) {
      throw new Error('Saved selfie folder was deleted or moved. Set CONFIG.SELFIE_FOLDER_ID in Code.gs.');
    }
  }

  const folders = DriveApp.getFoldersByName(CONFIG.SELFIE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.SELFIE_FOLDER_NAME);
  try {
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (e) {}
  props.setProperty('SELFIE_FOLDER_ID', folder.getId());
  return folder;
}

function reverseGeocodeServer_(lat, lng) {
  try {
    const res = Maps.newGeocoder().setRegion('ph').reverseGeocode(lat, lng);
    if (res && res.status === 'OK' && res.results && res.results.length > 0) {
      return res.results[0].formatted_address;
    }
  } catch (e) {}
  return '';
}

function validateAndFormatLocation_(loc) {
  if (!loc || loc.lat === undefined || loc.lng === undefined || loc.lat === null || loc.lng === null) {
    throw new Error('Location access is required. Please enable GPS/Location services and try again.');
  }
  const numLat = Number(loc.lat);
  const numLng = Number(loc.lng);
  if (isNaN(numLat) || isNaN(numLng)) {
    throw new Error('Invalid GPS coordinates detected.');
  }

  let address = String(loc.address || '').trim();
  let fenceCheck = isWithinAllowedArea_(numLat, numLng, loc.accuracy, address);

  if (!fenceCheck.allowed && !address) {
    address = reverseGeocodeServer_(numLat, numLng);
    fenceCheck = isWithinAllowedArea_(numLat, numLng, loc.accuracy, address);
  }

  if (!fenceCheck.allowed) {
    throw new Error(
      'Submission Blocked: You are outside ' + CONFIG.ALLOWED_LOCATION_LABEL +
      ' (~' + fenceCheck.distanceMeters + 'm away). Attendance can only be submitted inside the designated location.'
    );
  }

  const lat = numLat.toFixed(6);
  const lng = numLng.toFixed(6);
  const acc = loc.accuracy ? ' ±' + Math.round(Number(loc.accuracy)) + 'm' : '';
  const coordsPart = '(' + lat + ', ' + lng + acc + ')';
  return address
    ? address + ' ' + coordsPart
    : 'SM City Batangas Premises ' + coordsPart;
}

function submitAttendance(form) {
  const employeeId = String(form.employeeId || '').trim().toUpperCase();
  const role = String(form.role || '').trim();
  const action = String(form.action || '').trim().toUpperCase();
  const remarks = String(form.remarks || '').trim();
  const selfieData = String(form.selfie || '').trim();
  const privacyConsent = Boolean(form.privacyConsent);

  const allowedRoles = [
    'Admin',
    'Assistant Retail Manager',
    'CXA',
    'CXA / Head Cashier',
    'CXA / TM',
    'Inventory Planning Associate',
    'Retail Manager'
  ];

  if (!employeeId) throw new Error('Please enter your Employee ID.');
  if (!privacyConsent) throw new Error('Please agree to the Data Privacy Notice before submitting.');
  if (!role) throw new Error('Please select your Role for the Day.');
  if (!allowedRoles.includes(role)) throw new Error('Please select a valid Role for the Day.');
  if (!['IN', 'OUT'].includes(action)) throw new Error('Choose Time In or Time Out.');
  if (!selfieData) throw new Error('Please capture a live selfie before submitting.');

  const employee = lookupEmployee(employeeId);
  if (!employee.found) throw new Error(employee.message);

  const locationText = validateAndFormatLocation_(form.location);
  const aiCheck = verifySelfieWithAI_(selfieData);

  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const savedSelfie = saveSelfieToDrive_(selfieData, today, employeeId, action);
  const selfieUrl = savedSelfie.url;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);
    let row = -1;

    if (sheet.getLastRow() > 1) {
      const idCol = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
      const matches = idCol.createTextFinder(employeeId).matchEntireCell(true).findAll();
      for (let i = matches.length - 1; i >= 0; i--) {
        const rIdx = matches[i].getRow();
        const dateVal = sheet.getRange(rIdx, 1).getValue();
        const rowDate = dateVal instanceof Date
          ? Utilities.formatDate(dateVal, CONFIG.TIME_ZONE, 'yyyy-MM-dd')
          : String(dateVal).slice(0, 10);
        if (rowDate === today) {
          row = rIdx;
          break;
        }
      }
    }

    const stamp = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');

    if (row === -1) {
      if (action === 'OUT') throw new Error('You must record Time In before Time Out today.');
      sheet.appendRow([
        today,
        employeeId,
        employee.employeeName,
        role,
        stamp,
        '',
        remarks,
        stamp,
        selfieUrl,
        locationText,
        '',
        '',
        'IN: ' + aiCheck.summary,
        'Consented (' + stamp + ')'
      ]);
      return {
        message: 'Time In recorded for ' + employee.employeeName + ' (' + aiCheck.summary + ')',
        status: 'IN'
      };
    }

    const current = sheet.getRange(row, 1, 1, 14).getValues()[0];
    if (action === 'IN') throw new Error('This employee has already recorded Time In today. Only one Time In is allowed per day.');
    if (current[5]) throw new Error('This employee has already recorded Time Out today. Only one Time Out is allowed per day.');
    if (!current[4]) throw new Error('Time In is required before Time Out.');

    const prevAi = String(current[12] || '').trim();
    const combinedAi = [prevAi, 'OUT: ' + aiCheck.summary].filter(Boolean).join(' | ');

    sheet.getRange(row, 6).setValue(stamp);
    if (remarks) sheet.getRange(row, 7).setValue(remarks);
    sheet.getRange(row, 8).setValue(stamp);
    sheet.getRange(row, 11).setValue(selfieUrl);
    sheet.getRange(row, 12).setValue(locationText);
    sheet.getRange(row, 13).setValue(combinedAi);

    return {
      message: 'Time Out recorded for ' + employee.employeeName + ' (' + aiCheck.summary + ')',
      status: 'OUT'
    };
  } catch (err) {
    try {
      if (savedSelfie && savedSelfie.file) savedSelfie.file.setTrashed(true);
    } catch (cleanupErr) {}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// Strictly returns ONLY the authenticated employee's own records
function getEmployeeDashboard(employeeId, employeePin, month) {
  const employee = lookupEmployeeWithPin_(employeeId, employeePin);
  if (!employee.found) throw new Error(employee.message);

  const verifiedId = employee.employeeId.toUpperCase();
  const wantedMonth = String(month || Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM'));
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  const records = [];
  let present = 0, early = 0, totalMinutes = 0;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const rowEmpId = String(r[1] || '').trim().toUpperCase();
    const rowMonth = String(r[0] || '').slice(0, 7);

    if (rowEmpId !== verifiedId || rowMonth !== wantedMonth) continue;

    const earlyMin = earlyMinutes_(r[5]);
    if (r[4]) present++;
    if (earlyMin > 0) early++;
    if (r[4] && r[5]) {
      totalMinutes += Math.max(0, Math.round((new Date(r[5].replace(' ', 'T')) - new Date(r[4].replace(' ', 'T'))) / 60000));
    }
    records.push({
      date: r[0],
      employeeId: r[1] || employee.employeeId,
      employeeName: r[2] || employee.employeeName,
      role: r[3],
      timeIn: r[4],
      timeOut: r[5],
      remarks: r[6],
      timeInLocation: r[9] || '',
      timeOutLocation: r[11] || '',
      earlyMinutes: earlyMin,
      status: !r[4] ? 'Absent' : 'Present'
    });
  }

  records.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  const today = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const todayRec = records.find(x => x.date === today);
  const hours = Math.floor(totalMinutes / 60), mins = totalMinutes % 60;
  const summary = { present, early, totalHours: hours + ':' + String(mins).padStart(2, '0') };

  try {
    syncPersonalEmployeeSheet_(employee, wantedMonth, summary, records);
  } catch (syncErr) {
    console.warn('Personal sheet sync skipped:', syncErr.message);
  }

  return {
    employeeName: employee.employeeName,
    employeeId: employee.employeeId,
    month: wantedMonth,
    summary,
    today: todayRec || { status: 'No attendance yet' },
    records
  };
}

function syncPersonalEmployeeSheet_(employee, month, summary, records) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'PERSONAL_SHEET_' + employee.employeeId.toUpperCase();
  let fileId = props.getProperty(propKey);
  let ss = null;

  if (fileId) {
    try {
      ss = SpreadsheetApp.openById(fileId);
    } catch (e) {
      fileId = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Attendance - ' + employee.employeeName + ' (' + employee.employeeId + ')');
    fileId = ss.getId();
    props.setProperty(propKey, fileId);
    try {
      const file = DriveApp.getFileById(fileId);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    } catch (e) {}
  }

  const tabName = month;
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
  } else {
    sheet.clearContents();
  }

  const summaryRows = [
    ['Employee Name', employee.employeeName, 'Month', month],
    ['Employee ID', employee.employeeId, 'Days Present', summary.present],
    ['Total Hours Worked', summary.totalHours, 'Early Out Days', summary.early],
    ['', '', '', '']
  ];
  sheet.getRange(1, 1, summaryRows.length, 4).setValues(summaryRows);
  sheet.getRange('A1:A3').setFontWeight('bold');
  sheet.getRange('C1:C3').setFontWeight('bold');

  const headers = ['Date', 'Employee ID', 'Employee Name', 'Role', 'Time In', 'Time Out', 'In Location', 'Out Location', 'Remarks'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (records.length > 0) {
    const rows = records.map(r => [
      r.date || '',
      r.employeeId || employee.employeeId,
      r.employeeName || employee.employeeName,
      r.role || '',
      r.timeIn || '',
      r.timeOut || '',
      r.timeInLocation || '',
      r.timeOutLocation || '',
      r.remarks || ''
    ]);
    sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);
  return ss.getUrl();
}

function earlyMinutes_(time) {
  if (!time) return 0;
  const t = String(time).slice(11, 16);
  const a = t.split(':').map(Number), b = CONFIG.WORK_END.split(':').map(Number);
  return Math.max(0, (b[0] * 60 + b[1]) - (a[0] * 60 + a[1]));
}

// ==========================================
// ADMIN: ATTENDANCE LOGS & EMPLOYEE DIRECTORY
// ==========================================

function getResponses(adminPin) {
  if (String(adminPin || '') !== CONFIG.ADMIN_PIN) throw new Error('Incorrect administrator PIN.');
  const values = getSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME).getDataRange().getDisplayValues();
  if (values.length <= 1) return [];
  return values.slice(1).reverse().map(r => ({
    date: r[0],
    employeeId: r[1],
    employeeName: r[2],
    role: r[3],
    timeIn: r[4],
    timeOut: r[5],
    remarks: r[6],
    timeInLocation: r[9] || '',
    timeOutLocation: r[11] || ''
  }));
}

function getSpreadsheetUrl(adminPin) {
  if (String(adminPin || '') !== CONFIG.ADMIN_PIN) throw new Error('Incorrect administrator PIN.');
  return getSpreadsheet_().getUrl();
}

function getEmployeesList(adminPin) {
  if (String(adminPin || '') !== CONFIG.ADMIN_PIN) {
    throw new Error('Incorrect administrator PIN.');
  }
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues();
  return values.map(r => ({
    employeeId: r[0],
    employeeName: r[1],
    active: r[2] || 'YES',
    employeePin: r[3] || ''
  }));
}

function addNewEmployee(adminPin, employeeData) {
  if (String(adminPin || '') !== CONFIG.ADMIN_PIN) {
    throw new Error('Incorrect administrator PIN.');
  }

  const empId = String(employeeData.employeeId || '').trim().toUpperCase();
  const empName = String(employeeData.employeeName || '').trim();
  const empPin = String(employeeData.employeePin || '').trim();
  const activeStatus = String(employeeData.active || 'YES').trim().toUpperCase();

  if (!empId) throw new Error('Employee ID is required.');
  if (!empName) throw new Error('Employee Name is required.');
  if (!empPin || empPin.length < 4) {
    throw new Error('Employee PIN must be at least 4 characters.');
  }
  if (!['YES', 'NO'].includes(activeStatus)) {
    throw new Error('Active status must be YES or NO.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
    if (!sheet) {
      setup();
      sheet = ss.getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
    }

    if (sheet.getLastRow() > 1) {
      const existingIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
      for (let i = 0; i < existingIds.length; i++) {
        if (String(existingIds[i][0] || '').trim().toUpperCase() === empId) {
          throw new Error('Employee ID "' + empId + '" already exists in the directory.');
        }
      }
    }

    sheet.appendRow([empId, empName, activeStatus, empPin]);
    sheet.autoResizeColumns(1, 4);

    return {
      message: 'Added ' + empName + ' (' + empId + ') successfully.',
      employees: getEmployeesList(adminPin)
    };
  } finally {
    lock.releaseLock();
  }
}

function updateEmployeeDetails(adminPin, employeeData) {
  if (String(adminPin || '') !== CONFIG.ADMIN_PIN) {
    throw new Error('Incorrect administrator PIN.');
  }

  const targetId = String(employeeData.employeeId || '').trim().toUpperCase();
  const empName = String(employeeData.employeeName || '').trim();
  const empPin = String(employeeData.employeePin || '').trim();
  const activeStatus = String(employeeData.active || 'YES').trim().toUpperCase();

  if (!targetId) throw new Error('Employee ID is missing.');
  if (!empName) throw new Error('Employee Name is required.');
  if (!empPin || empPin.length < 4) {
    throw new Error('Employee PIN must be at least 4 characters.');
  }
  if (!['YES', 'NO'].includes(activeStatus)) {
    throw new Error('Active status must be YES or NO.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(CONFIG.EMPLOYEES_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      throw new Error('Employee directory is empty.');
    }

    const existingIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    let foundRow = -1;
    for (let i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0] || '').trim().toUpperCase() === targetId) {
        foundRow = i + 2;
        break;
      }
    }

    if (foundRow === -1) {
      throw new Error('Employee ID "' + targetId + '" was not found.');
    }

    sheet.getRange(foundRow, 2, 1, 3).setValues([[empName, activeStatus, empPin]]);
    sheet.autoResizeColumns(1, 4);

    return {
      message: 'Updated ' + empName + ' (' + targetId + ').',
      employees: getEmployeesList(adminPin)
    };
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || CONFIG.SPREADSHEET_ID;
  if (!id) throw new Error('Spreadsheet is not initialized. Run setup() once from the Apps Script editor.');
  return SpreadsheetApp.openById(id);
}

// Replace with your deployed Google Apps Script Web App URL:
const API_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';

const ALLOWED_ZONE = {
  label: 'SM City Batangas Premises & Parking Lots (Pallocan Kanluran, Batangas City)',
  lat: 13.7554,
  lng: 121.0695,
  radiusMeters: 350
};

let mediaStream = null;
let capturedSelfieData = '';
let preVerifyPromise = null;
let preVerifyError = '';
let currentLocation = null;
let locationWatchId = null;
let watchStartedAt = 0;
let lastGeocodedKey = '';
let lastResolvedAddress = '';
let usingNativeCamera = true;
let allAdminResponses = [];
let allAdminEmployees = [];

function isTransientNetworkError(err) {
  const msg = (err && err.message) ? err.message : String(err || '');
  return (
    msg.includes('Failed to fetch') ||
    msg.includes('NetworkError') ||
    msg.includes('Load failed') ||
    msg.includes('HTTP 502') ||
    msg.includes('HTTP 503') ||
    msg.includes('HTTP 504')
  );
}

// CORS-free JSONP transport for GET-sized payloads (bypasses browser CORS blocks)
function callViaJsonp(cleanUrl, methodName, args) {
  return new Promise((resolve, reject) => {
    const encodedArgs = encodeURIComponent(JSON.stringify(args || []));
    if (encodedArgs.length > 6500) {
      reject(new Error('Payload too large for JSONP fallback.'));
      return;
    }

    const cbName = '__gasCb_' + Date.now() + '_' + Math.floor(Math.random() * 100000);
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Server request timed out. Check your internet connection or Apps Script deployment.'));
    }, 15000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[cbName] = function(res) {
      cleanup();
      if (res && res.ok) resolve(res.data);
      else reject(new Error((res && res.error) || 'Server returned an error.'));
    };

    script.onerror = function() {
      cleanup();
      reject(new Error('Blocked by browser or Apps Script is not set to "Anyone". Open ' + cleanUrl + ' in Incognito to verify.'));
    };

    const sep = cleanUrl.includes('?') ? '&' : '?';
    script.src = cleanUrl + sep +
      'action=' + encodeURIComponent(methodName) +
      '&args=' + encodedArgs +
      '&callback=' + encodeURIComponent(cbName);
    document.head.appendChild(script);
  });
}

// Connects standalone frontend to Google Apps Script Web App API
async function runAppsScript(methodName, args = [], maxRetries = 2) {
  const cleanUrl = String(API_URL || '').trim();
  if (!cleanUrl || cleanUrl.includes('PASTE_YOUR_GOOGLE_APPS_SCRIPT')) {
    throw new Error('Backend API_URL is not configured in app.js.');
  }

  let attempt = 0;
  while (true) {
    attempt++;
    try {
      const response = await fetch(cleanUrl, {
        method: 'POST',
        mode: 'cors',
        redirect: 'follow',
        cache: 'no-store',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: methodName, args: args })
      });

      if (!response.ok) {
        throw new Error('Server error (HTTP ' + response.status + ')');
      }

      const rawText = await response.text();
      let json;
      try {
        json = JSON.parse(rawText);
      } catch (parseErr) {
        throw new Error('Apps Script returned an HTML login/error page instead of JSON. Ensure Web App access is set to "Anyone" and republished as a New Version.');
      }

      if (!json.ok) {
        throw new Error(json.error || 'Unknown server error.');
      }

      return json.data;
    } catch (err) {
      const canUseJsonp = methodName !== 'submitAttendance' && methodName !== 'preVerifySelfie';
      if (canUseJsonp && isTransientNetworkError(err)) {
        try {
          return await callViaJsonp(cleanUrl, methodName, args);
        } catch (jsonpErr) {
          throw jsonpErr;
        }
      }

      if (attempt <= maxRetries && isTransientNetworkError(err)) {
        await new Promise(resolve => setTimeout(resolve, 900 * attempt));
        continue;
      }
      throw err;
    }
  }
}

function friendlyErrorMessage(err) {
  const raw = (err && err.message) ? err.message : String(err || '');
  if (raw.includes('Failed to fetch') || raw.includes('NetworkError') || raw.includes('Load failed')) {
    return 'Browser blocked the connection to Google Apps Script. Disable ad/tracker blockers (Brave Shields/uBlock) for this site, or verify your Apps Script deployment is set to "Anyone" under a New Version.';
  }
  return raw;
}

function formatTime12(value) {
  if (!value) return '—';
  const s = String(value).trim();
  let m = s.match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) m = s.match(/(?:T|\s)(\d{1,2}):(\d{2})(?::\d{2})?/);
  if (!m) return s;
  let h = Number(m[1]), min = m[2];
  const ap = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  return h + ':' + min + ' ' + ap;
}

function showTab(which) {
  ['attendance', 'dashboard', 'responses'].forEach(x => {
    document.getElementById(x + 'Panel').classList.toggle('hidden', x !== which);
    document.getElementById(x + 'Tab').classList.toggle('active', x === which);
  });
  if (which !== 'attendance') {
    stopCameraStream();
    stopLocationWatcher();
  } else {
    startSmoothLocationTracker();
  }
  if (which !== 'dashboard') {
    document.getElementById('dashPin').value = '';
    document.getElementById('dashRows').innerHTML = '';
    document.getElementById('dashContent').classList.add('hidden');
    document.getElementById('dashNotice').style.display = 'none';
    document.getElementById('dashName').textContent = 'Enter your Employee ID and personal PIN to view only your own attendance records.';
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function notice(id, msg, ok) {
  const e = document.getElementById(id);
  e.textContent = msg;
  e.className = 'notice ' + (ok ? 'success' : 'error');
  e.style.display = 'block';
}

function stopCameraStream() {
  if (mediaStream) {
    mediaStream.getTracks().forEach(t => t.stop());
    mediaStream = null;
  }
}

function openNativeCamera() {
  stopCameraStream();
  usingNativeCamera = true;
  const input = document.getElementById('nativeCamInput');
  input.value = '';
  input.click();
}

function startBackgroundVerification(dataUrl) {
  preVerifyError = '';
  const status = document.getElementById('camStatus');
  status.textContent = 'Verifying live person in background...';
  status.style.color = '#92400e';

  preVerifyPromise = runAppsScript('preVerifySelfie', [dataUrl], 2)
    .then(res => {
      if (capturedSelfieData === dataUrl) {
        status.textContent = '✓ ' + (res && res.summary ? res.summary : 'Live person verified') + '. Ready to submit!';
        status.style.color = '#166534';
      }
      return true;
    })
    .catch(err => {
      const msg = friendlyErrorMessage(err);
      if (capturedSelfieData === dataUrl) {
        preVerifyError = msg;
        status.textContent = msg;
        status.style.color = '#991b1b';
      }
      return false;
    });
}

async function startCamera(silentFallback) {
  const status = document.getElementById('camStatus');
  const retryBtn = document.getElementById('retryCamBtn');
  const nativeBtn = document.getElementById('nativeCamBtn');
  const captureBtn = document.getElementById('captureBtn');
  const video = document.getElementById('cameraVideo');
  const camFrame = document.getElementById('camFrame');

  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    enableNativeCamera(
      silentFallback
        ? 'Tap "Open Front Camera" to take a live selfie.'
        : 'Live preview is unavailable in this browser. Tap "Open Front Camera" instead.',
      !silentFallback
    );
    return;
  }

  try {
    stopCameraStream();
    mediaStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: 'user' }, width: { ideal: 320 } },
      audio: false
    });
    usingNativeCamera = false;
    video.srcObject = mediaStream;
    await video.play();
    camFrame.classList.remove('hidden');
    video.classList.remove('hidden');
    document.getElementById('selfiePreview').classList.add('hidden');
    captureBtn.classList.remove('hidden');
    nativeBtn.classList.add('hidden');
    retryBtn.classList.add('hidden');
    status.textContent = 'Camera ready. Position your face clearly and tap Take Live Selfie.';
    status.style.color = '#166534';
  } catch (err) {
    enableNativeCamera(
      silentFallback
        ? 'Tap "Open Front Camera" to take a live selfie.'
        : 'Camera permission denied (' + (err.name || 'NotAllowedError') + '). Tap "Open Front Camera" to use your device camera.',
      !silentFallback
    );
  }
}

function enableNativeCamera(msg, isWarning) {
  usingNativeCamera = true;
  const status = document.getElementById('camStatus');
  status.textContent = msg;
  status.style.color = isWarning ? '#92400e' : '#64748b';
  document.getElementById('camFrame').classList.add('hidden');
  document.getElementById('captureBtn').classList.add('hidden');
  document.getElementById('nativeCamBtn').classList.remove('hidden');
  document.getElementById('retryCamBtn').classList.remove('hidden');
}

async function handleNativeSelfie(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const status = document.getElementById('camStatus');
  status.textContent = 'Processing selfie...';
  status.style.color = '#64748b';

  try {
    const canvas = document.getElementById('selfieCanvas');
    const ctx = canvas.getContext('2d');
    const maxW = 320;

    if (window.createImageBitmap) {
      const bitmap = await createImageBitmap(file);
      const scale = Math.min(1, maxW / bitmap.width);
      canvas.width = Math.round(bitmap.width * scale);
      canvas.height = Math.round(bitmap.height * scale);
      ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
      bitmap.close();
      finalizeSelfiePreview(canvas.toDataURL('image/jpeg', 0.55));
    } else {
      const reader = new FileReader();
      reader.onload = function(e) {
        const img = new Image();
        img.onload = function() {
          const scale = Math.min(1, maxW / img.width);
          canvas.width = Math.round(img.width * scale);
          canvas.height = Math.round(img.height * scale);
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          finalizeSelfiePreview(canvas.toDataURL('image/jpeg', 0.55));
        };
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    }
  } catch (err) {
    status.textContent = 'Could not read image. Please tap "Open Front Camera" again.';
    status.style.color = '#991b1b';
  }
}

function finalizeSelfiePreview(dataUrl) {
  capturedSelfieData = dataUrl;
  const preview = document.getElementById('selfiePreview');
  preview.src = capturedSelfieData;
  document.getElementById('camFrame').classList.remove('hidden');
  document.getElementById('cameraVideo').classList.add('hidden');
  preview.classList.remove('hidden');

  document.getElementById('nativeCamBtn').classList.add('hidden');
  document.getElementById('retryCamBtn').classList.add('hidden');
  document.getElementById('retakeBtn').classList.remove('hidden');

  startBackgroundVerification(capturedSelfieData);
}

function captureSelfie() {
  const video = document.getElementById('cameraVideo');
  if (!video.videoWidth) {
    notice('formNotice', 'Live stream is not ready. Tap "Open Front Camera" instead.', false);
    return;
  }
  const canvas = document.getElementById('selfieCanvas');
  const maxW = 320;
  const scale = Math.min(1, maxW / video.videoWidth);
  canvas.width = Math.round(video.videoWidth * scale);
  canvas.height = Math.round(video.videoHeight * scale);

  const ctx = canvas.getContext('2d');
  ctx.translate(canvas.width, 0);
  ctx.scale(-1, 1);
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

  capturedSelfieData = canvas.toDataURL('image/jpeg', 0.55);
  stopCameraStream();

  const preview = document.getElementById('selfiePreview');
  preview.src = capturedSelfieData;
  preview.classList.remove('hidden');
  video.classList.add('hidden');

  document.getElementById('captureBtn').classList.add('hidden');
  document.getElementById('retakeBtn').classList.remove('hidden');

  startBackgroundVerification(capturedSelfieData);
}

function retakeSelfie() {
  capturedSelfieData = '';
  preVerifyPromise = null;
  preVerifyError = '';
  document.getElementById('nativeCamInput').value = '';
  document.getElementById('selfiePreview').classList.add('hidden');
  document.getElementById('retakeBtn').classList.add('hidden');

  if (usingNativeCamera) {
    document.getElementById('camFrame').classList.add('hidden');
    document.getElementById('nativeCamBtn').classList.remove('hidden');
    document.getElementById('retryCamBtn').classList.remove('hidden');
    document.getElementById('camStatus').textContent = 'Tap "Open Front Camera" to take a new selfie.';
    document.getElementById('camStatus').style.color = '#64748b';
  } else {
    startCamera(false);
  }
}

function getDistanceMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function checkInsideAllowedZone(lat, lng, accuracy, address) {
  const dist = getDistanceMeters(lat, lng, ALLOWED_ZONE.lat, ALLOWED_ZONE.lng);
  const accBuffer = Math.min(75, Math.max(0, Number(accuracy) || 0));
  const withinRadius = dist <= (ALLOWED_ZONE.radiusMeters + accBuffer);

  const addrLower = String(address || '').toLowerCase();
  const inBatangas = addrLower.includes('batangas');
  const matchesMallPremises =
    addrLower.includes('sm city batangas') ||
    addrLower.includes('sm batangas');

  return {
    allowed: withinRadius || (inBatangas && matchesMallPremises && dist <= 500),
    distance: Math.round(dist)
  };
}

async function resolveSpecificAddress(lat, lng) {
  const gridKey = lat.toFixed(3) + ',' + lng.toFixed(3);
  if (gridKey === lastGeocodedKey && lastResolvedAddress) {
    return lastResolvedAddress;
  }
  try {
    const url = 'https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=' +
      encodeURIComponent(lat) + '&lon=' + encodeURIComponent(lng) + '&zoom=18&addressdetails=1';
    const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
    if (!res.ok) return lastResolvedAddress || '';
    const data = await res.json();
    const a = data.address || {};
    const parts = [
      a.building || a.amenity || a.shop || a.mall || a.office,
      [a.house_number, a.road].filter(Boolean).join(' '),
      a.suburb || a.neighbourhood || a.quarter || a.village,
      a.city || a.town || a.municipality,
      a.state || a.region
    ].filter(Boolean);
    const formatted = parts.length ? parts.join(', ') : (data.display_name || '');
    if (formatted) {
      lastGeocodedKey = gridKey;
      lastResolvedAddress = formatted;
    }
    return formatted;
  } catch (e) {
    return lastResolvedAddress || '';
  }
}

function stopLocationWatcher() {
  if (locationWatchId !== null && navigator.geolocation) {
    navigator.geolocation.clearWatch(locationWatchId);
    locationWatchId = null;
  }
}

function applyLocationReading(pos) {
  const pill = document.getElementById('locationStatus');
  const details = document.getElementById('locationDetails');
  const addrEl = document.getElementById('locationAddress');
  const coordsEl = document.getElementById('locationCoords');

  const lat = pos.coords.latitude;
  const lng = pos.coords.longitude;
  const accuracy = pos.coords.accuracy;

  if (
    currentLocation &&
    currentLocation.inZone &&
    (Date.now() - currentLocation.timestamp < 30000) &&
    accuracy > (currentLocation.accuracy * 2.5) &&
    accuracy > 120
  ) {
    return currentLocation;
  }

  details.classList.remove('hidden');
  coordsEl.textContent = 'GPS: ' + lat.toFixed(6) + ', ' + lng.toFixed(6) + ' (±' + Math.round(accuracy) + 'm)';

  const check = checkInsideAllowedZone(lat, lng, accuracy, lastResolvedAddress);
  const defaultLabel = 'SM City Batangas, Pallocan Kanluran, Batangas City';

  currentLocation = {
    lat: lat,
    lng: lng,
    accuracy: accuracy,
    address: check.allowed ? (lastResolvedAddress || defaultLabel) : (lastResolvedAddress || ''),
    inZone: check.allowed,
    distance: check.distance,
    timestamp: Date.now()
  };

  if (check.allowed) {
    pill.innerHTML = '<span>✓ Inside SM Batangas Premises (' + check.distance + 'm)</span>';
    pill.className = 'loc-pill ok';
    addrEl.textContent = currentLocation.address;
  } else {
    const stillWarmingUp = (Date.now() - watchStartedAt < 3500) && accuracy > 150;
    if (stillWarmingUp) {
      pill.innerHTML = '<span class="pulse-dot"></span><span>Refining indoor GPS signal...</span>';
      pill.className = 'loc-pill warn';
      addrEl.textContent = 'Pinpointing exact location inside SM City Batangas...';
    } else {
      pill.innerHTML = '<span>Outside Allowed Area (~' + check.distance + 'm away)</span>';
      pill.className = 'loc-pill err';
      addrEl.textContent = lastResolvedAddress || 'Outside designated SM City Batangas premises';
    }
  }

  resolveSpecificAddress(lat, lng).then(resolved => {
    if (!resolved || !currentLocation) return;
    currentLocation.address = resolved;
    addrEl.textContent = resolved;
    const recheck = checkInsideAllowedZone(currentLocation.lat, currentLocation.lng, currentLocation.accuracy, resolved);
    if (recheck.allowed && !currentLocation.inZone) {
      currentLocation.inZone = true;
      pill.innerHTML = '<span>✓ Inside SM Batangas Premises (' + recheck.distance + 'm)</span>';
      pill.className = 'loc-pill ok';
    }
  });

  return currentLocation;
}

function startSmoothLocationTracker(forceReset) {
  const pill = document.getElementById('locationStatus');
  if (!navigator.geolocation) {
    pill.textContent = 'Geolocation not supported by this browser';
    pill.className = 'loc-pill err';
    return;
  }

  if (forceReset) {
    stopLocationWatcher();
    currentLocation = null;
    pill.innerHTML = '<span class="pulse-dot"></span><span>Recalibrating GPS...</span>';
    pill.className = 'loc-pill';
  }

  if (locationWatchId !== null) return;
  watchStartedAt = Date.now();

  navigator.geolocation.getCurrentPosition(
    pos => {
      if (!currentLocation) applyLocationReading(pos);
    },
    () => {},
    { enableHighAccuracy: false, timeout: 3000, maximumAge: 60000 }
  );

  locationWatchId = navigator.geolocation.watchPosition(
    pos => applyLocationReading(pos),
    () => {
      if (!currentLocation) {
        pill.textContent = 'Location access denied or unavailable';
        pill.className = 'loc-pill err';
        const details = document.getElementById('locationDetails');
        details.classList.remove('hidden');
        document.getElementById('locationAddress').textContent = 'Unable to verify your location';
        document.getElementById('locationCoords').textContent = 'Enable High-Accuracy GPS in your phone settings, then tap Recalibrate GPS.';
      }
    },
    { enableHighAccuracy: true, timeout: 10000, maximumAge: 5000 }
  );
}

function refreshLocation(manualReset) {
  if (manualReset) {
    startSmoothLocationTracker(true);
  }
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error('Geolocation is not supported by this browser.'));
      return;
    }
    if (currentLocation && currentLocation.inZone && (Date.now() - currentLocation.timestamp < 90000) && !manualReset) {
      resolve(currentLocation);
      return;
    }
    startSmoothLocationTracker(false);
    navigator.geolocation.getCurrentPosition(
      pos => {
        const loc = applyLocationReading(pos);
        if (loc.inZone) resolve(loc);
        else reject(new Error('Submission Blocked: You are outside ' + ALLOWED_ZONE.label + ' (~' + loc.distance + 'm away).'));
      },
      () => {
        if (currentLocation && currentLocation.inZone) resolve(currentLocation);
        else reject(new Error('Location permission is required to verify you are inside SM City Batangas premises.'));
      },
      { enableHighAccuracy: true, timeout: 8000, maximumAge: 15000 }
    );
  });
}

function lookupEmployeeName() {
  const id = document.getElementById('employeeId').value.trim();
  const name = document.getElementById('employeeName');
  const s = document.getElementById('employeeLookupStatus');
  name.value = '';
  if (!id) return;
  s.textContent = 'Checking Employee ID...';
  s.style.color = '#64748b';

  runAppsScript('lookupEmployee', [id], 2)
    .then(r => {
      if (r.found) {
        name.value = r.employeeName;
        s.textContent = 'Employee verified: ' + r.employeeName;
        s.style.color = '#166534';
      } else {
        s.textContent = r.message;
        s.style.color = '#991b1b';
      }
    })
    .catch(e => {
      s.textContent = friendlyErrorMessage(e);
      s.style.color = '#991b1b';
    });
}

async function submitForm(e) {
  e.preventDefault();
  const b = document.getElementById('submitBtn');
  b.disabled = true;

  if (!document.getElementById('employeeName').value) {
    notice('formNotice', 'Please enter a valid Employee ID first.', false);
    b.disabled = false;
    b.textContent = 'Submit Attendance';
    return;
  }

  if (!document.getElementById('privacyConsent').checked) {
    notice('formNotice', 'Please check the Data Privacy Notice consent box before submitting.', false);
    b.disabled = false;
    b.textContent = 'Submit Attendance';
    return;
  }

  if (!capturedSelfieData) {
    notice('formNotice', 'Please capture a live selfie showing your face before submitting.', false);
    b.disabled = false;
    b.textContent = 'Submit Attendance';
    return;
  }

  if (preVerifyPromise) {
    b.textContent = 'Finishing selfie check...';
    const passed = await preVerifyPromise;
    if (!passed && preVerifyError) {
      notice('formNotice', preVerifyError, false);
      b.disabled = false;
      b.textContent = 'Submit Attendance';
      document.getElementById('formNotice').scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }
  }

  let loc = currentLocation;
  const isFreshLocation = loc && loc.inZone && (Date.now() - (loc.timestamp || 0) < 90000);

  if (!isFreshLocation) {
    b.textContent = 'Verifying GPS location...';
    try {
      loc = await refreshLocation(false);
    } catch (err) {
      notice('formNotice', err.message, false);
      b.disabled = false;
      b.textContent = 'Submit Attendance';
      document.getElementById('formNotice').scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }
  }

  if (!loc || !loc.inZone) {
    notice('formNotice', 'Submission Blocked: You must be inside ' + ALLOWED_ZONE.label + ' to submit attendance.', false);
    b.disabled = false;
    b.textContent = 'Submit Attendance';
    document.getElementById('formNotice').scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }

  b.textContent = 'Saving Attendance...';
  const p = {
    employeeId: document.getElementById('employeeId').value,
    role: document.getElementById('role').value,
    action: document.getElementById('action').value,
    remarks: document.getElementById('remarks').value,
    privacyConsent: document.getElementById('privacyConsent').checked,
    selfie: capturedSelfieData,
    location: loc
  };

  try {
    const r = await runAppsScript('submitAttendance', [p], 2);
    notice('formNotice', r.message, true);
    document.getElementById('action').value = r.status === 'IN' ? 'OUT' : 'IN';
    document.getElementById('remarks').value = '';
    retakeSelfie();
  } catch (err) {
    notice('formNotice', friendlyErrorMessage(err), false);
  } finally {
    b.disabled = false;
    b.textContent = 'Submit Attendance';
    document.getElementById('formNotice').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}

function buildCell(label, text) {
  const td = document.createElement('td');
  td.setAttribute('data-label', label);
  td.textContent = text;
  return td;
}

function buildLocationCell(label, locStr) {
  const td = document.createElement('td');
  td.setAttribute('data-label', label);
  if (!locStr) {
    td.textContent = '—';
    return td;
  }

  const wrap = document.createElement('div');
  wrap.className = 'proof-wrap';

  const addrDiv = document.createElement('div');
  addrDiv.className = 'proof-addr';
  const cleanAddr = locStr.replace(/\s*\([^)]*\)\s*$/, '').trim();
  addrDiv.textContent = cleanAddr || locStr;
  addrDiv.title = locStr;
  wrap.appendChild(addrDiv);

  const match = locStr.match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/);
  const coords = match ? (match[1] + ',' + match[2]) : locStr.split('(')[0].trim();
  const map = document.createElement('a');
  map.href = 'https://www.google.com/maps?q=' + encodeURIComponent(coords);
  map.target = '_blank';
  map.rel = 'noopener noreferrer';
  map.className = 'link-btn';
  map.textContent = 'Open Map';
  map.title = locStr;
  wrap.appendChild(map);

  td.appendChild(wrap);
  return td;
}

function renderRecordsTable(tbody, rows, fallbackId, fallbackName, emptyColspan, emptyMessage) {
  tbody.innerHTML = '';
  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.appendChild(buildCell('Date', r.date || '—'));
    tr.appendChild(buildCell('Employee ID', r.employeeId || fallbackId || '—'));
    tr.appendChild(buildCell('Name', r.employeeName || fallbackName || '—'));
    tr.appendChild(buildCell('Role', r.role || '—'));
    tr.appendChild(buildCell('Time In', formatTime12(r.timeIn)));
    tr.appendChild(buildCell('Time Out', formatTime12(r.timeOut)));
    tr.appendChild(buildLocationCell('In Location', r.timeInLocation));
    tr.appendChild(buildLocationCell('Out Location', r.timeOutLocation));
    tr.appendChild(buildCell('Remarks', r.remarks || '—'));
    tbody.appendChild(tr);
  });

  if (!rows.length) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row';
    tr.innerHTML = '<td colspan="' + emptyColspan + '" class="empty">' + emptyMessage + '</td>';
    tbody.appendChild(tr);
  }
}

async function loadDashboard() {
  const id = document.getElementById('dashId').value.trim();
  const pin = document.getElementById('dashPin').value.trim();
  const month = document.getElementById('dashMonth').value;
  const body = document.getElementById('dashRows');
  const loadBtn = document.getElementById('dashLoadBtn');

  body.innerHTML = '';
  document.getElementById('dashContent').classList.add('hidden');

  if (!id) { notice('dashNotice', 'Enter your Employee ID.', false); return; }
  if (!pin) { notice('dashNotice', 'Enter your Employee PIN.', false); return; }
  if (!month) { notice('dashNotice', 'Select a month.', false); return; }

  loadBtn.disabled = true;
  notice('dashNotice', 'Verifying PIN & loading your attendance records...', true);

  try {
    const r = await runAppsScript('getEmployeeDashboard', [id, pin, month], 2);
    document.getElementById('dashNotice').style.display = 'none';
    document.getElementById('dashContent').classList.remove('hidden');
    document.getElementById('dashName').textContent = 'Personal Attendance Sheet: ' + r.employeeName + ' (' + r.employeeId + ') • ' + r.month;

    const sum = r.summary || { present: 0, totalHours: '0:00', early: 0 };
    document.getElementById('sumPresent').textContent = sum.present;
    document.getElementById('sumHours').textContent = sum.totalHours;
    document.getElementById('sumEarly').textContent = sum.early;

    renderRecordsTable(
      body,
      r.records,
      r.employeeId,
      r.employeeName,
      9,
      'No attendance records found for your ID in ' + r.month + '.'
    );
  } catch (e) {
    document.getElementById('dashContent').classList.add('hidden');
    document.getElementById('dashName').textContent = 'Enter your Employee ID and personal PIN to view only your own attendance records.';
    notice('dashNotice', friendlyErrorMessage(e), false);
  } finally {
    loadBtn.disabled = false;
  }
}

function switchAdminView(view) {
  document.getElementById('adminLogsView').classList.toggle('hidden', view !== 'logs');
  document.getElementById('adminEmployeesView').classList.toggle('hidden', view !== 'employees');
  document.getElementById('subTabLogs').classList.toggle('active', view === 'logs');
  document.getElementById('subTabEmployees').classList.toggle('active', view === 'employees');
}

function filterAdminLogs() {
  const q = (document.getElementById('adminLogSearch').value || '').trim().toLowerCase();
  const dateVal = (document.getElementById('adminLogDateFilter').value || '').trim();

  const filtered = allAdminResponses.filter(r => {
    if (dateVal && String(r.date || '').slice(0, 10) !== dateVal) return false;
    if (!q) return true;
    const haystack = [
      r.date,
      r.employeeId,
      r.employeeName,
      r.role,
      r.remarks,
      r.timeInLocation,
      r.timeOutLocation
    ].join(' ').toLowerCase();
    return haystack.includes(q);
  });

  const countEl = document.getElementById('logsCountLabel');
  if (countEl) {
    countEl.textContent = filtered.length === allAdminResponses.length
      ? 'Showing all ' + allAdminResponses.length + ' record(s)'
      : 'Showing ' + filtered.length + ' of ' + allAdminResponses.length + ' record(s)';
  }

  const emptyMsg = allAdminResponses.length
    ? 'No attendance logs match your search filter.'
    : 'No attendance responses recorded yet.';
  renderRecordsTable(document.getElementById('responseRows'), filtered, '', '', 9, emptyMsg);
}

function clearAdminLogSearch() {
  document.getElementById('adminLogSearch').value = '';
  document.getElementById('adminLogDateFilter').value = '';
  filterAdminLogs();
}

function filterAdminEmployees() {
  const q = (document.getElementById('adminEmpSearch').value || '').trim().toLowerCase();
  const filtered = allAdminEmployees.filter(emp => {
    if (!q) return true;
    const statusText = String(emp.active || 'YES').toUpperCase() === 'YES' ? 'active' : 'inactive';
    const haystack = [emp.employeeId, emp.employeeName, statusText].join(' ').toLowerCase();
    return haystack.includes(q);
  });
  renderEmployeesTable(filtered);
}

function clearAdminEmpSearch() {
  document.getElementById('adminEmpSearch').value = '';
  filterAdminEmployees();
}

function editEmployeeRow(emp) {
  document.getElementById('empFormMode').value = 'edit';
  document.getElementById('empFormHeading').textContent = 'Edit Employee: ' + emp.employeeId;
  document.getElementById('empFormSubheading').textContent = 'Update the employee\'s name, PIN, or active status.';
  const idInput = document.getElementById('newEmpId');
  idInput.value = emp.employeeId;
  idInput.readOnly = true;

  document.getElementById('newEmpName').value = emp.employeeName;
  document.getElementById('newEmpPin').value = emp.employeePin;
  document.getElementById('newEmpActive').value = String(emp.active || 'YES').toUpperCase() === 'NO' ? 'NO' : 'YES';

  document.getElementById('addEmpBtn').textContent = 'Save Changes';
  document.getElementById('cancelEditEmpBtn').classList.remove('hidden');
  document.getElementById('addEmpNotice').style.display = 'none';
  document.getElementById('adminEmployeesView').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function resetEmployeeForm() {
  document.getElementById('empFormMode').value = 'add';
  document.getElementById('empFormHeading').textContent = 'Add New Employee';
  document.getElementById('empFormSubheading').textContent = 'Register a new employee ID and default personal PIN for attendance and dashboard access.';
  document.getElementById('addEmployeeForm').reset();
  const idInput = document.getElementById('newEmpId');
  idInput.readOnly = false;
  document.getElementById('newEmpPin').value = '1234';
  document.getElementById('newEmpActive').value = 'YES';
  document.getElementById('addEmpBtn').textContent = 'Add Employee';
  document.getElementById('cancelEditEmpBtn').classList.add('hidden');
}

function renderEmployeesTable(list) {
  const tbody = document.getElementById('employeeDirectoryRows');
  document.getElementById('empCount').textContent = list.length === allAdminEmployees.length
    ? list.length
    : list.length + ' of ' + allAdminEmployees.length;
  tbody.innerHTML = '';

  list.forEach(emp => {
    const tr = document.createElement('tr');
    tr.appendChild(buildCell('Employee ID', emp.employeeId || '—'));
    tr.appendChild(buildCell('Employee Name', emp.employeeName || '—'));

    const statusTd = document.createElement('td');
    statusTd.setAttribute('data-label', 'Status');
    const badge = document.createElement('span');
    const isYes = String(emp.active || 'YES').toUpperCase() === 'YES';
    badge.className = 'badge ' + (isYes ? 'badge-yes' : 'badge-no');
    badge.textContent = isYes ? 'ACTIVE' : 'INACTIVE';
    statusTd.appendChild(badge);
    tr.appendChild(statusTd);

    tr.appendChild(buildCell('Personal PIN', emp.employeePin || '—'));

    const actionTd = document.createElement('td');
    actionTd.setAttribute('data-label', 'Action');
    const editBtn = document.createElement('button');
    editBtn.type = 'button';
    editBtn.className = 'mini-btn';
    editBtn.textContent = 'Edit';
    editBtn.onclick = () => editEmployeeRow(emp);
    actionTd.appendChild(editBtn);
    tr.appendChild(actionTd);

    tbody.appendChild(tr);
  });

  if (!list.length) {
    const tr = document.createElement('tr');
    tr.className = 'empty-row';
    const emptyMsg = allAdminEmployees.length
      ? 'No employees match your search.'
      : 'No employees found in directory.';
    tr.innerHTML = '<td colspan="5" class="empty">' + emptyMsg + '</td>';
    tbody.appendChild(tr);
  }
}

async function loadResponses() {
  const pin = document.getElementById('adminPin').value;
  const btn = document.getElementById('adminUnlockBtn');
  if (!pin) {
    notice('responseNotice', 'Please enter the Administrator PIN.', false);
    return;
  }

  btn.disabled = true;
  btn.textContent = 'Verifying...';

  try {
    const rows = await runAppsScript('getResponses', [pin], 2);
    document.getElementById('responseNotice').style.display = 'none';
    document.getElementById('adminUnlockedArea').classList.remove('hidden');

    allAdminResponses = rows || [];
    filterAdminLogs();

    runAppsScript('getSpreadsheetUrl', [pin], 2).then(url => {
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.className = 'link-btn';
      a.textContent = 'Open Google Spreadsheet database';
      document.getElementById('sheetLink').innerHTML = '';
      document.getElementById('sheetLink').appendChild(a);
    }).catch(() => {});

    runAppsScript('getEmployeesList', [pin], 2).then(empList => {
      allAdminEmployees = empList || [];
      filterAdminEmployees();
    }).catch(() => {});
  } catch (e) {
    document.getElementById('adminUnlockedArea').classList.add('hidden');
    notice('responseNotice', friendlyErrorMessage(e), false);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Unlock Admin Panel';
  }
}

async function submitEmployeeForm(e) {
  e.preventDefault();
  const pin = document.getElementById('adminPin').value;
  const btn = document.getElementById('addEmpBtn');
  const mode = document.getElementById('empFormMode').value;

  const employeeData = {
    employeeId: document.getElementById('newEmpId').value,
    employeeName: document.getElementById('newEmpName').value,
    employeePin: document.getElementById('newEmpPin').value,
    active: document.getElementById('newEmpActive').value
  };

  btn.disabled = true;
  btn.textContent = mode === 'edit' ? 'Saving Changes...' : 'Adding Employee...';
  document.getElementById('addEmpNotice').style.display = 'none';

  try {
    const method = mode === 'edit' ? 'updateEmployeeDetails' : 'addNewEmployee';
    const res = await runAppsScript(method, [pin, employeeData], 2);
    notice('addEmpNotice', res.message, true);
    resetEmployeeForm();
    if (res.employees) {
      allAdminEmployees = res.employees;
      filterAdminEmployees();
    }
  } catch (err) {
    btn.textContent = mode === 'edit' ? 'Save Changes' : 'Add Employee';
    notice('addEmpNotice', friendlyErrorMessage(err), false);
  } finally {
    btn.disabled = false;
  }
}

(function init() {
  const ua = navigator.userAgent || '';
  if (/FBAN|FBAV|Messenger|Instagram|Viber|Line\//i.test(ua)) {
    document.getElementById('inAppBrowserBanner').classList.remove('hidden');
  }
  const d = new Date();
  document.getElementById('dashMonth').value = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  startSmoothLocationTracker(false);
})();
