const SHEETS_ = {
  USERS: 'users',
  MODULES: 'modules',
  CONTENT: 'content',
  PROGRESS: 'progress',
  SESSIONS: 'sessions',
  LOGINS: 'logins',
  APPROVED_EMAILS: 'approved_emails'
};

const SESSION_HEADERS = [
  'timestamp',
  'email',
  'code',
  'module_key',
  'session_id',
  'session_start',
  'questions_answered',
  'clock_minutes',
  'active_minutes',
  'interaction_rate',
  'device_type',
  'completion_reached'
];

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(data.action || '').trim();

    if (action === 'signUpUser') return signUpUser_(data);
    if (action === 'verifyUserLogin') return verifyUserLogin_(data);
    if (action === 'getUserDashboard') return getUserDashboard_(data);
    if (action === 'loadTraining') return loadTraining_(data);
    if (action === 'saveProgress') return saveProgress_(data);
    if (action === 'submitSession') return submitSession_(data);

    return jsonResponse(fail_('unknown_action'));
  } catch (err) {
    return jsonResponse(fail_(String(err)));
  }
}

function signUpUser_(data) {
  const email = normalizeEmail_(data.email);
  const name = String(data.name || '').trim();
  const password = String(data.password || '');

  if (!email || !name || !password) return jsonResponse(fail_('Missing signup fields.'));

  const approvedSheet = getOrCreateSheet_(SHEETS_.APPROVED_EMAILS);
  const approvedRows = getRows_(approvedSheet);
  const isApproved = approvedRows.some(r => normalizeEmail_(r.email) === email);
  if (!isApproved) {
    return jsonResponse(fail_('This email address is not authorized to create an account.'));
  }

  const usersSheet = getOrCreateSheet_(SHEETS_.USERS);
  const users = getRows_(usersSheet);
  if (users.some(u => normalizeEmail_(u.email) === email)) {
    return jsonResponse(fail_('An account already exists for this email.'));
  }

  const salt = makeSalt_();
  const passwordHash = hashPassword_(password, salt);

  ensureHeaders_(usersSheet, ['created_at', 'email', 'name', 'password_hash', 'salt', 'active']);
  usersSheet.appendRow([new Date(), email, name, passwordHash, salt, true]);

  return jsonResponse(ok_({ email: email, name: name }));
}

function verifyUserLogin_(data) {
  const email = normalizeEmail_(data.email);
  const password = String(data.password || '');
  if (!email || !password) return jsonResponse(fail_('Email and password are required.'));

  const usersSheet = getOrCreateSheet_(SHEETS_.USERS);
  const users = getRows_(usersSheet);
  const user = users.find(u => normalizeEmail_(u.email) === email);
  if (!user) return jsonResponse(fail_('Invalid email or password.'));

  const active = normalizeBool_(user.active);
  if (!active) return jsonResponse(fail_('This account is inactive.'));

  const expectedHash = String(user.password_hash || '');
  const salt = String(user.salt || '');
  const incomingHash = hashPassword_(password, salt);
  if (!expectedHash || incomingHash !== expectedHash) {
    logLogin_(email, false);
    return jsonResponse(fail_('Invalid email or password.'));
  }

  logLogin_(email, true);
  return jsonResponse(ok_({
    email: email,
    name: String(user.name || '')
  }));
}

function getUserDashboard_(data) {
  const email = normalizeEmail_(data.email);
  if (!email) return jsonResponse(fail_('Email is required.'));

  const users = getRows_(getOrCreateSheet_(SHEETS_.USERS));
  const user = users.find(u => normalizeEmail_(u.email) === email);
  if (!user) return jsonResponse(fail_('User not found.'));
  if (!normalizeBool_(user.active)) return jsonResponse(fail_('User is inactive.'));

  const moduleRows = getRows_(getOrCreateSheet_(SHEETS_.MODULES));
  const progressRows = getRows_(getOrCreateSheet_(SHEETS_.PROGRESS))
    .filter(r => normalizeEmail_(r.email) === email);
  const sessionsRows = getRows_(getOrCreateSheet_(SHEETS_.SESSIONS))
    .filter(r => normalizeEmail_(r.email) === email);

  const progressByKey = {};
  progressRows.forEach(r => {
    const key = String(r.module_key || '').trim();
    if (!key) return;
    progressByKey[key] = r;
  });

  const modules = moduleRows
    .filter(m => String(m.module_key || '').trim())
    .map(m => {
      const key = String(m.module_key || '').trim();
      const p = progressByKey[key] || {};
      return {
        module_key: key,
        module_title: String(m.module_title || m.title || key),
        module_subtitle: String(m.module_subtitle || m.subtitle || ''),
        progress_percent: Number(p.progress_percent || 0),
        completed: normalizeBool_(p.completed),
        last_activity: p.updated_at || p.last_activity || ''
      };
    });

  const modulesStarted = modules.filter(m => Number(m.progress_percent || 0) > 0).length;
  const modulesCompleted = modules.filter(m => !!m.completed).length;
  const totalSessions = sessionsRows.length;

  const avgActiveMinutes = average_(sessionsRows.map(r => Number(r.active_minutes || 0)));
  const avgInteractionRate = average_(sessionsRows.map(r => Number(r.interaction_rate || 0)));
  const lastActivity = mostRecent_(
    progressRows.map(r => r.updated_at).concat(sessionsRows.map(r => r.timestamp || r.session_start))
  );

  return jsonResponse(ok_({
    user: {
      email: email,
      name: String(user.name || '')
    },
    stats: {
      modules_started: modulesStarted,
      modules_completed: modulesCompleted,
      total_sessions: totalSessions,
      avg_active_minutes: avgActiveMinutes,
      avg_interaction_rate: avgInteractionRate,
      last_activity: lastActivity || ''
    },
    modules: modules
  }));
}

function loadTraining_(data) {
  const email = normalizeEmail_(data.email);
  const moduleKey = String(data.module_key || data.moduleKey || '').trim();
  if (!email) return jsonResponse(fail_('Email is required.'));
  if (!moduleKey) return jsonResponse(fail_('Module key is required.'));

  const users = getRows_(getOrCreateSheet_(SHEETS_.USERS));
  const user = users.find(u => normalizeEmail_(u.email) === email);
  if (!user) return jsonResponse(fail_('User not found.'));
  if (!normalizeBool_(user.active)) return jsonResponse(fail_('User is inactive.'));

  const modules = getRows_(getOrCreateSheet_(SHEETS_.MODULES));
  const moduleMeta = modules.find(m => String(m.module_key || '').trim() === moduleKey);
  if (!moduleMeta) return jsonResponse(fail_('Module not found.'));

  const contentRows = getRows_(getOrCreateSheet_(SHEETS_.CONTENT))
    .filter(r => String(r.module_key || '').trim() === moduleKey)
    .sort((a, b) => Number(a.section_order || 0) - Number(b.section_order || 0));

  if (!contentRows.length) return jsonResponse(fail_('No content rows found for this module.'));

  return jsonResponse(ok_({
    moduleKey: moduleKey,
    moduleTitle: String(moduleMeta.module_title || moduleMeta.title || moduleKey),
    moduleLabel: String(moduleMeta.module_label || 'Professional Development Module'),
    participantBadge: `${String(user.name || '').trim() || email} · ${email}`,
    rows: contentRows
  }));
}

function saveProgress_(data) {
  const email = normalizeEmail_(data.email || data.participantEmail);
  const moduleKey = String(data.module_key || data.moduleKey || '').trim();
  if (!email || !moduleKey) return jsonResponse(fail_('Missing email or module key.'));

  const sheet = getOrCreateSheet_(SHEETS_.PROGRESS);
  ensureHeaders_(sheet, ['updated_at', 'email', 'module_key', 'progress_percent', 'last_section', 'completed']);

  const rows = getRowsWithRowNum_(sheet);
  const found = rows.find(r => normalizeEmail_(r.email) === email && String(r.module_key || '').trim() === moduleKey);

  const payload = [
    new Date(),
    email,
    moduleKey,
    Number(data.progress_percent || 0),
    String(data.last_section || ''),
    normalizeBool_(data.completed)
  ];

  if (found) {
    sheet.getRange(found._rowNum, 1, 1, payload.length).setValues([payload]);
  } else {
    sheet.appendRow(payload);
  }

  return jsonResponse(ok_({}));
}

function submitSession_(data) {
  const sheet = getOrCreateSheet_(SHEETS_.SESSIONS);
  ensureSessionsHeader_(sheet);

  const email = String(data.email || data.participantEmail || '');
  const code = String(data.code || data.participantCode || '');
  const moduleKey = String(data.moduleKey || data.module_key || '');
  const sessionId = String(data.sessionId || data.session_id || '');
  const sessionStart = String(data.sessionStart || data.session_start || '');
  const questionsAnswered = Number(data.totalQuestions ?? data.questions_answered ?? 0);
  const clockMinutes = Number(data.clock_minutes ?? 0);
  const activeMinutes = Number(data.active_minutes ?? 0);
  const interactionRate = Number(data.interaction_rate ?? 0);
  const deviceType = String(data.device_type || 'desktop');
  const completionReached = normalizeBool_(data.completionReached);

  sheet.appendRow([
    new Date(),
    email,
    code,
    moduleKey,
    sessionId,
    sessionStart,
    questionsAnswered,
    clockMinutes,
    activeMinutes,
    interactionRate,
    deviceType,
    completionReached
  ]);

  return jsonResponse(ok_({}));
}

function ensureSessionsHeader_(sheet) {
  const existing = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, SESSION_HEADERS.length).getValues()[0]
    : [];

  const isMatch = SESSION_HEADERS.every((header, idx) => String(existing[idx] || '').trim() === header);
  if (!isMatch) {
    sheet.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]);
  }
}

function logLogin_(email, success) {
  const sheet = getOrCreateSheet_(SHEETS_.LOGINS);
  ensureHeaders_(sheet, ['timestamp', 'email', 'success']);
  sheet.appendRow([new Date(), email, !!success]);
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const same = headers.every((h, i) => String(existing[i] || '').trim() === h);
  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  return values.slice(1).map(function(row) {
    var out = {};
    headers.forEach(function(h, idx) { out[h] = row[idx]; });
    return out;
  });
}

function getRowsWithRowNum_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  return values.slice(1).map(function(row, idx) {
    var out = { _rowNum: idx + 2 };
    headers.forEach(function(h, col) { out[h] = row[col]; });
    return out;
  });
}

function average_(nums) {
  const valid = nums.filter(function(n) { return Number.isFinite(n); });
  if (!valid.length) return 0;
  const sum = valid.reduce(function(a, b) { return a + b; }, 0);
  return Number((sum / valid.length).toFixed(2));
}

function mostRecent_(values) {
  const dates = values
    .map(function(v) { return v ? new Date(v) : null; })
    .filter(function(d) { return d && !isNaN(d.getTime()); })
    .sort(function(a, b) { return b.getTime() - a.getTime(); });
  return dates.length ? dates[0].toISOString() : '';
}

function makeSalt_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function hashPassword_(password, salt) {
  const raw = String(password || '') + '::' + String(salt || '');
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytesToHex_(bytes);
}

function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    const v = (b < 0) ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeBool_(value) {
  if (typeof value === 'boolean') return value;
  const raw = String(value || '').trim().toLowerCase();
  return raw === 'true' || raw === '1' || raw === 'yes';
}

function ok_(data) {
  const payload = data || {};
  payload.ok = true;
  return payload;
}

function fail_(message) {
  return { ok: false, error: String(message || 'unknown_error') };
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
