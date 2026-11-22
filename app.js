// Replace with your deployed Google Apps Script Web App URL:
const API_URL = 'https://script.google.com/a/macros/globe.com.ph/s/AKfycbz5I56zVUtPUkjfgfyiDIM6FXgUloaBN2KmCICK0sgIVJoyuoEsONJeHy8wHLDtPgzH/exec';

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

// Connects standalone frontend to Google Apps Script Web App API
async function runAppsScript(methodName, args = [], maxRetries = 2) {
  if (!API_URL || API_URL.includes('PASTE_YOUR_GOOGLE_APPS_SCRIPT')) {
    throw new Error('Backend API_URL is not configured in app.js.');
  }

  let attempt = 0;
  while (true) {
    attempt++;
    try {
      const response = await fetch(API_URL, {
        method: 'POST',
        // text/plain prevents CORS preflight OPTIONS errors with Google Apps Script
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: methodName, args: args })
      });

      if (!response.ok) {
        throw new Error('Server error (HTTP ' + response.status + ')');
      }

      const json = await response.json();
      if (!json.ok) {
        throw new Error(json.error || 'Unknown server error.');
      }

      return json.data;
    } catch (err) {
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
  if (raw.includes('Failed to fetch') || raw.includes('NetworkError')) {
    return 'Unable to reach the attendance server. Please check your internet connection and try again.';
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
