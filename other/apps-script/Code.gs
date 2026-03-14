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

    if (action === 'loadTraining') return loadTraining(data);
    if (action === 'saveProgress') return saveProgress(data);
    if (action === 'submitSession') return submitSession(data);

    return jsonResponse({ ok: false, error: 'unknown_action' });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function submitSession(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('sessions') || ss.insertSheet('sessions');
  ensureSessionsHeader(sheet);

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
  const completionReached = normalizeBool(data.completionReached);

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

  return jsonResponse({ ok: true });
}

function ensureSessionsHeader(sheet) {
  const existing = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, SESSION_HEADERS.length).getValues()[0]
    : [];

  const isMatch = SESSION_HEADERS.every((header, idx) => String(existing[idx] || '').trim() === header);
  if (!isMatch) {
    sheet.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]);
  }
}

function normalizeBool(value) {
  if (typeof value === 'boolean') return value;
  const raw = String(value || '').trim().toLowerCase();
  return raw === 'true' || raw === '1' || raw === 'yes';
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
