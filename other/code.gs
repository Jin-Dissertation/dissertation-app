const QUESTION_BANK_SPREADSHEET_ID = "1MujF_VPCHDN1buPIO1T7XGu7oKazS3Gd0tdbhz-j3qM";

/* =========================================================
   CONFIG
   ========================================================= */

const SHEETS_ = {
  USERS: "users",
  MODULES: "modules",
  CONTENT: "content",
  PROGRESS: "progress",
  SESSIONS: "sessions",
  LOGINS: "logins",
  APPROVED_EMAILS: "approved_emails",
  DAILY_PRACTICE_TOTALS: "daily_practice_totals"
};

const QUESTION_BANK_ = {
  SPREADSHEET_ID: "1MujF_VPCHDN1buPIO1T7XGu7oKazS3Gd0tdbhz-j3qM",
  HISTORY_SHEET: "question_history",
  PRACTICE_TOPICS_SHEET: "practice_topics"
};

const GOOGLE_CLIENT_ID_ = "534534627380-bd6hjuj4ihbf6i7r2qsjnen6cqhm14ma.apps.googleusercontent.com";
const ALLOWED_DOMAIN_ = "apu.edu";
const DEFAULT_DAILY_GOAL_ = 20;
const APP_TIMEZONE_ = "America/Los_Angeles";

/* =========================================================
   WEB APP ENTRY
   ========================================================= */

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    const action = String(data.action || "").trim();

    switch (action) {
      case "verifyGoogleLogin":
        return json_(verifyGoogleLogin_(data));

      case "getUserDashboard":
        return json_(getUserDashboard_(data));

      case "loadTraining":
        return json_(loadTraining_(data));

      case "loadTrainingByAssignmentKey":
        return json_(loadTrainingByAssignmentKey_(data));

      case "updateDailyGoal":
        return json_(updateDailyGoal_(data));

      case "saveProgress":
        return json_(saveProgress_(data));

      case "submitSession":
        return json_(submitSession_(data));

      case "getPracticeQuestions":
        return json_(getPracticeQuestions_(data));

      case "submitPracticeResults":
        return json_(submitPracticeResults_(data));

      case "getPracticeTopics":
        return json_(getPracticeTopics_(data));

      default:
        return json_({
          ok: false,
          error: "Unknown action."
        });
    }
  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

/* =========================================================
   GOOGLE AUTH
   ========================================================= */

function verifyGoogleLogin_(data) {
  const idToken = String(data.id_token || "").trim();
  if (!idToken) return fail_("Missing Google ID token.");

  const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status !== 200) {
    return fail_("Google token verification failed.");
  }

  const payload = JSON.parse(text);

  const aud = String(payload.aud || "");
  const iss = String(payload.iss || "");
  const email = normalizeEmail_(payload.email);
  const emailVerified = String(payload.email_verified || "").toLowerCase() === "true";

  if (aud !== GOOGLE_CLIENT_ID_) {
    return fail_("Token audience mismatch.");
  }

  if (iss !== "accounts.google.com" && iss !== "https://accounts.google.com") {
    return fail_("Token issuer mismatch.");
  }

  if (!email || !emailVerified) {
    return fail_("Google account email was not verified.");
  }

  if (ALLOWED_DOMAIN_ && !email.endsWith("@" + ALLOWED_DOMAIN_)) {
    return fail_("Please sign in with your " + ALLOWED_DOMAIN_ + " account.");
  }

  const approvedRows = sheetObjects_(getSheet_(SHEETS_.APPROVED_EMAILS));
  const approvedEmails = approvedRows
    .map(r => normalizeEmail_(r.email))
    .filter(Boolean);

  if (!approvedEmails.includes(email)) {
    return fail_("This email address is not authorized.");
  }

  ensureUserExists_(email);
  logLoginAttempt_(email, true, "google_auth_success");

  return {
    ok: true,
    email: email,
    name: ""
  };
}

/* =========================================================
   DASHBOARD
   ========================================================= */

function getUserDashboard_(data) {
  const email = normalizeEmail_(data.email);
  if (!email) return fail_("Missing email.");

  const users = sheetObjects_(getSheet_(SHEETS_.USERS));
  const modules = sheetObjects_(getSheet_(SHEETS_.MODULES));
  const progressRows = sheetObjects_(getSheet_(SHEETS_.PROGRESS));
  const sessionRows = sheetObjects_(getSheet_(SHEETS_.SESSIONS));

  const user = users.find(r => normalizeEmail_(r.email) === email);
  if (!user) return fail_("User not found.");

  const dailyGoal = sanitizeDailyGoal_(user.daily_goal || DEFAULT_DAILY_GOAL_);

  const activeModules = modules.filter(m => toBool_(m.active));
  const practiceModulesRaw = activeModules.filter(m =>
    String(m.module_type || "").trim().toLowerCase() === "practice"
  );

  const practiceModules = practiceModulesRaw.map(m => {
    const moduleKey = String(m.module_key || "").trim();
    const progress = progressRows.find(p =>
      normalizeEmail_(p.email) === email &&
      String(p.module_key || "").trim() === moduleKey
    ) || {};

    return {
      module_key: moduleKey,
      module_title: String(m.module_title || moduleKey),
      module_subtitle: String(m.module_subtitle || ""),
      badge_text: String(m.badge_text || ""),
      estimated_minutes: Number(m.estimated_minutes || 0) || 0,
      course: String(m.course || "").trim(),
      module_type: String(m.module_type || "").trim(),
      active: toBool_(m.active),
      progress_percent: clampNumber_(Number(progress.progress_percent || 0), 0, 100),
      completed: toBool_(progress.completed),
      last_activity: progress.last_update || ""
    };
  });

  const practiceCourses = [...new Set(
    practiceModules
      .map(m => String(m.course || "").trim())
      .filter(Boolean)
  )].sort();

  const userSessions = sessionRows.filter(r => normalizeEmail_(r.email) === email);

  const totalQuestionsAnswered = userSessions.reduce((sum, r) => {
    return sum + Number(r.total_questions || r.questions_answered || 0);
  }, 0);

  const completedSessions = userSessions.filter(r => toBool_(r.completion_reached));
  const totalModulesCompleted = completedSessions.length;

const dailyPracticeRows = loadDailyPracticeTotalsForUser_(email);
const questionsToday = getQuestionsAnsweredForDate_(dailyPracticeRows, todayKey_());
const streakData = computeDailyPracticeStreakData_(dailyPracticeRows, dailyGoal);

  const modulesByKey = {};
  modules.forEach(m => {
    const key = String(m.module_key || "").trim();
    if (key) modulesByKey[key] = String(m.module_title || key);
  });

  const completionCounts = {};
  completedSessions.forEach(r => {
    const key = String(r.module_key || "").trim();
    if (!key) return;
    completionCounts[key] = (completionCounts[key] || 0) + 1;
  });

  const topModules = Object.keys(completionCounts)
    .map(key => ({
      module_key: key,
      module_title: modulesByKey[key] || key,
      completion_count: completionCounts[key]
    }))
    .sort((a, b) => {
      if (b.completion_count !== a.completion_count) {
        return b.completion_count - a.completion_count;
      }
      return a.module_title.localeCompare(b.module_title);
    })
    .slice(0, 3);

  const stats = {
    total_sessions: userSessions.length,
    avg_active_minutes: round2_(averageNumber_(
      userSessions.map(r => Number(r.active_minutes || 0))
    )),
    avg_interaction_rate: round2_(averageNumber_(
      userSessions.map(r => Number(r.interaction_rate || 0))
    )),
    last_activity: latestDateLikeValue_(
      practiceModules.map(m => m.last_activity).concat(
        userSessions.map(r => r.timestamp || r.session_end || r.session_start || "")
      )
    ),
    total_modules_completed: totalModulesCompleted,
    total_questions_answered: totalQuestionsAnswered,
    daily_goal: dailyGoal,
    questions_today: questionsToday,
    streak_days: streakData.streak_days,
    streak_secured_today: streakData.streak_secured_today,
    top_modules: topModules
  };

  return {
    ok: true,
    user: {
      email: email,
      name: ""
    },
    stats: stats,
    assignment_entry_enabled: true,
    practice_courses: practiceCourses,
    practice_modules: practiceModules
  };
}

/* =========================================================
   MODULE LOADING
   ========================================================= */

function loadTraining_(data) {
  const email = normalizeEmail_(data.email);
  const moduleKey = String(data.module_key || data.moduleKey || "").trim();

  if (!email) return fail_("Missing email.");
  if (!moduleKey) return fail_("Missing module key.");

  const users = sheetObjects_(getSheet_(SHEETS_.USERS));
  const modules = sheetObjects_(getSheet_(SHEETS_.MODULES));
  const contentRowsAll = sheetObjects_(getSheet_(SHEETS_.CONTENT));

  const user = users.find(r => normalizeEmail_(r.email) === email);
  if (!user) return fail_("User not found.");

  const meta = modules.find(r =>
    String(r.module_key || "").trim() === moduleKey &&
    toBool_(r.active)
  );

  if (!meta) {
    return fail_('Module "' + moduleKey + '" was not found or is inactive.');
  }

  const rows = contentRowsAll
    .filter(r => String(r.module_key || "").trim() === moduleKey)
    .sort((a, b) => {
      const aSection = Number(a.section_order || 0);
      const bSection = Number(b.section_order || 0);
      const aCard = Number(a.slot || a.card_order || 0);
      const bCard = Number(b.slot || b.card_order || 0);
      return aSection - bSection || aCard - bCard;
    });

  if (!rows.length) {
    return fail_('No content rows found for module "' + moduleKey + '".');
  }

  return trainingPayloadFromMetaAndRows_(meta, rows);
}

function loadTrainingByAssignmentKey_(data) {
  const email = normalizeEmail_(data.email);
  const assignmentKey = normalizeAssignmentKey_(data.assignment_key);

  if (!email) return fail_("Missing email.");
  if (!assignmentKey) return fail_("Missing assignment key.");

  const users = sheetObjects_(getSheet_(SHEETS_.USERS));
  const modules = sheetObjects_(getSheet_(SHEETS_.MODULES));
  const contentRowsAll = sheetObjects_(getSheet_(SHEETS_.CONTENT));

  const user = users.find(r => normalizeEmail_(r.email) === email);
  if (!user) return fail_("User not found.");

  const meta = modules.find(r =>
    normalizeAssignmentKey_(r.assignment_key) === assignmentKey &&
    toBool_(r.active)
  );

  if (!meta) {
    return fail_("That assignment key was not found or is not currently active.");
  }

  const moduleKey = String(meta.module_key || "").trim();

  const rows = contentRowsAll
    .filter(r => String(r.module_key || "").trim() === moduleKey)
    .sort((a, b) => {
      const aSection = Number(a.section_order || 0);
      const bSection = Number(b.section_order || 0);
      const aCard = Number(a.slot || a.card_order || 0);
      const bCard = Number(b.slot || b.card_order || 0);
      return aSection - bSection || aCard - bCard;
    });

  if (!rows.length) {
    return fail_('No content rows found for module "' + moduleKey + '".');
  }

  return trainingPayloadFromMetaAndRows_(meta, rows);
}

function trainingPayloadFromMetaAndRows_(meta, rows) {
  return {
    ok: true,
    moduleKey: String(meta.module_key || "").trim(),
    moduleTitle: String(meta.module_title || String(meta.module_key || "").trim()),
    moduleLabel: String(meta.module_subtitle || ""),
    participantBadge: String(meta.badge_text || ""),
    rows: rows
  };
}

/* =========================================================
   DAILY GOAL UPDATE
   ========================================================= */

function updateDailyGoal_(data) {
  const email = normalizeEmail_(data.email);
  const dailyGoal = sanitizeDailyGoal_(data.daily_goal);

  if (!email) return fail_("Missing email.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_(SHEETS_.USERS);
    const headerMap = getHeaderIndexMap_(sheet);
    const values = sheet.getDataRange().getValues();

    const emailCol = headerMap.email;
    const goalCol = headerMap.daily_goal;
    const activeCol = headerMap.active;

    if (!emailCol) return fail_('Missing "email" column in users sheet.');
    if (!goalCol) return fail_('Missing "daily_goal" column in users sheet.');

    for (let i = 1; i < values.length; i++) {
      const rowEmail = normalizeEmail_(values[i][emailCol - 1]);
      if (rowEmail === email) {
        sheet.getRange(i + 1, goalCol).setValue(dailyGoal);
        if (activeCol) sheet.getRange(i + 1, activeCol).setValue(true);
        return {
          ok: true,
          email: email,
          daily_goal: dailyGoal
        };
      }
    }

    appendUserRow_(sheet, {
      email: email,
      name: "",
      daily_goal: dailyGoal,
      salt: "",
      password_hash: "",
      active: true,
      created_at: new Date()
    });

    return {
      ok: true,
      email: email,
      daily_goal: dailyGoal
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   PROGRESS
   ========================================================= */

function saveProgress_(data) {
  const email = normalizeEmail_(data.email);
  const moduleKey = String(data.module_key || data.moduleKey || "").trim();
  const progressPercent = clampNumber_(
    Number(data.progress_percent != null ? data.progress_percent : data.progress || 0),
    0,
    100
  );
  const lastSection = String(data.last_section || data.lastSection || "").trim();
  const completed = (
    data.completed === true ||
    String(data.completed || "").toLowerCase() === "true" ||
    progressPercent >= 100
  );

  if (!email) return fail_("Missing email.");
  if (!moduleKey) return fail_("Missing module key.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_(SHEETS_.PROGRESS);
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const rowEmail = normalizeEmail_(values[i][0]);
      const rowModuleKey = String(values[i][1] || "").trim();

      if (rowEmail === email && rowModuleKey === moduleKey) {
        sheet.getRange(i + 1, 3).setValue(progressPercent);
        sheet.getRange(i + 1, 4).setValue(lastSection);
        sheet.getRange(i + 1, 5).setValue(completed);
        sheet.getRange(i + 1, 6).setValue(new Date());

        return {
          ok: true,
          updated: true,
          email: email,
          module_key: moduleKey,
          progress_percent: progressPercent,
          completed: completed
        };
      }
    }

    sheet.appendRow([
      email,
      moduleKey,
      progressPercent,
      lastSection,
      completed,
      new Date()
    ]);

    return {
      ok: true,
      updated: false,
      email: email,
      module_key: moduleKey,
      progress_percent: progressPercent,
      completed: completed
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   SESSION SUBMISSION
   ========================================================= */

function submitSession_(data) {
  const email = normalizeEmail_(data.participantEmail || data.email);
  const participantCode = String(data.participantCode || "").trim();
  const moduleKey = String(data.moduleKey || data.module_key || "").trim();
  const sessionId = String(data.sessionId || data.session_id || "").trim();
  const sessionStart = String(data.sessionStart || data.session_start || "").trim();
  const sessionEnd = String(data.sessionEnd || data.session_end || "").trim();
  const totalQuestions = Number(data.totalQuestions || data.questions_answered || 0);
  const completionReached = (
    data.completionReached === true ||
    String(data.completionReached || "").toLowerCase() === "true"
  );
  const clockMinutes = round2_(Number(data.clock_minutes || 0));
  const activeMinutes = round2_(Number(data.active_minutes || 0));
  const interactionRate = round2_(Number(data.interaction_rate || 0));
  const deviceType = String(data.device_type || "").trim() || "unknown";
  const events = Array.isArray(data.events) ? data.events : [];

  if (!email) return fail_("Missing email.");
  if (!moduleKey) return fail_("Missing module key.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_(SHEETS_.SESSIONS);
    const rows = sheetObjects_(sheet);

    const duplicate = rows.find(r =>
      normalizeEmail_(r.email) === email &&
      String(r.module_key || "").trim() === moduleKey &&
      String(r.session_id || "").trim() === sessionId &&
      sessionId
    );

    if (duplicate) {
      return {
        ok: true,
        duplicate: true,
        sessionId: sessionId
      };
    }

    sheet.appendRow([
      new Date(),
      email,
      participantCode,
      moduleKey,
      sessionId,
      sessionStart,
      sessionEnd,
      totalQuestions,
      completionReached,
      clockMinutes,
      activeMinutes,
      interactionRate,
      deviceType,
      JSON.stringify(events)
    ]);

    if (completionReached) {
      upsertProgressRow_(email, moduleKey, 100, "Completed", true);
    }

    return {
      ok: true,
      sessionId: sessionId
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
   PRACTICE QUESTIONS
   ========================================================= */

function getPracticeQuestions_(data) {
  const email = normalizeEmail_(data.email);
  const topics = Array.isArray(data.topics)
    ? data.topics.map(String).map(s => s.trim()).filter(Boolean)
    : [];
  const questionCount = Number(data.question_count || 0);

  if (!email) return fail_("Missing email.");
  if (!topics.length) return fail_("No topics provided.");
  if (!questionCount || questionCount < 1) return fail_("Invalid question count.");

  const bankSs = getQuestionBankSpreadsheet_();
  const dailyAnsweredCount = questionResults.filter(result => {
  const questionId = String(result.question_id || "").trim();
  return !!questionId;
}).length;
  const historyMap = loadStudentHistoryMap_(bankSs, email);
  const perTopicCounts = splitQuestionCountEvenly_(topics, questionCount);

  let questionSet = [];

  for (let i = 0; i < topics.length; i++) {
    const topic = topics[i];
    const targetCount = perTopicCounts[topic];

    const topicQuestions = loadTopicQuestions_(bankSs, topic);
    if (!topicQuestions.length) continue;

    const weightedPool = buildWeightedPool_(topicQuestions, historyMap);
    const chosen = weightedRandomSampleWithoutReplacement_(weightedPool, targetCount);
    questionSet = questionSet.concat(chosen);
  }

  questionSet = shuffleArray_(questionSet);

  return {
    ok: true,
    question_set: questionSet
  };
}

function submitPracticeResults_(data) {
  const email = normalizeEmail_(data.email);
  const questionResults = Array.isArray(data.question_results) ? data.question_results : [];

  if (!email) return fail_("Missing email.");
  if (!questionResults.length) return fail_("No question results were provided.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const bankSs = getQuestionBankSpreadsheet_();
    const sheet = getOrCreateQuestionHistorySheet_(bankSs);
    ensureQuestionHistoryHeaders_(sheet);

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h || "").trim());

    const headerMap = {};
    headers.forEach((header, i) => {
      if (header) headerMap[header] = i;
    });

    const existingRows = values.slice(1);
    const rowIndexMap = {};

    existingRows.forEach((row, i) => {
      const rowEmail = normalizeEmail_(row[headerMap.email]);
      const questionId = String(row[headerMap.question_id] || "").trim();
      if (!rowEmail || !questionId) return;
      rowIndexMap[rowEmail + "||" + questionId] = i + 2;
    });

    const appendedRows = [];
    const updates = [];
    let processedCount = 0;

    questionResults.forEach(result => {
      const questionId = String(result.question_id || "").trim();
      const topic = String(result.topic || "").trim();
      const isCorrect = (
        result.is_correct === true ||
        String(result.is_correct || "").trim().toLowerCase() === "true"
      );

      if (!questionId) return;

      const key = email + "||" + questionId;
      const existingRowNumber = rowIndexMap[key];

      if (existingRowNumber) {
        const currentRow = sheet.getRange(existingRowNumber, 1, 1, headers.length).getValues()[0];
        const currentTimesSeen = Number(currentRow[headerMap.times_seen] || 0);
        const currentScore = normalizeRecentStreakScore_(currentRow[headerMap.recent_streak_score]);

        const nextTimesSeen = currentTimesSeen + 1;
        const nextScore = nextRecentStreakScore_(currentScore, isCorrect);

        currentRow[headerMap.email] = email;
        currentRow[headerMap.question_id] = questionId;
        currentRow[headerMap.topic] = topic;
        currentRow[headerMap.times_seen] = nextTimesSeen;
        currentRow[headerMap.recent_streak_score] = nextScore;

        updates.push({
          rowNumber: existingRowNumber,
          rowValues: currentRow
        });
      } else {
        const rowValues = headers.map(() => "");
        rowValues[headerMap.email] = email;
        rowValues[headerMap.question_id] = questionId;
        rowValues[headerMap.topic] = topic;
        rowValues[headerMap.times_seen] = 1;
        rowValues[headerMap.recent_streak_score] = nextRecentStreakScore_(0, isCorrect);

        appendedRows.push(rowValues);
      }

      processedCount += 1;
    });

    updates.forEach(update => {
      sheet.getRange(update.rowNumber, 1, 1, update.rowValues.length).setValues([update.rowValues]);
    });

    if (appendedRows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appendedRows.length, headers.length).setValues(appendedRows);
    }

// Use processedCount (number of questions submitted) instead
const dailyTotalsResult = upsertDailyPracticeTotal_(email, processedCount);

return {
  ok: true,
  processed_count: processedCount,
  updated_rows: updates.length,
  appended_rows: appendedRows.length,
  daily_total_updated: true,
  questions_added_today: processedCount, // <- use processedCount
  questions_today: dailyTotalsResult.questions_today
};
  } finally {
    lock.releaseLock();
  }
}

function getPracticeTopics_(data) {
  const bankSs = getQuestionBankSpreadsheet_();
  return {
    ok: true,
    practice_topics: loadPracticeTopics_(bankSs)
  };
}

function getQuestionBankSpreadsheet_() {
  const id = String(QUESTION_BANK_.SPREADSHEET_ID || "").trim();
  if (!id) {
    throw new Error("Question bank spreadsheet ID is not configured.");
  }
  return SpreadsheetApp.openById(id);
}

function getQuestionBankSheet_(name) {
  const ss = getQuestionBankSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function loadPracticeTopics_(ss) {
  const sheet = ss.getSheetByName(QUESTION_BANK_.PRACTICE_TOPICS_SHEET);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1);

  const topics = rows
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });

      const topicKey = String(obj.topic_key || "").trim();
      if (!topicKey) return null;
      if (!toBool_(obj.active)) return null;

      return {
        topic_key: topicKey,
        topic_label: String(obj.topic_label || topicKey).trim(),
        course: String(obj.course || "").trim(),
        active: true,
        sort_order: Number(obj.sort_order || 9999) || 9999,
        question_count: countActiveQuestionsForTopic_(ss, topicKey)
      };
    })
    .filter(Boolean)
    .sort((a, b) => {
      if (a.sort_order !== b.sort_order) return a.sort_order - b.sort_order;
      if (a.course !== b.course) return a.course.localeCompare(b.course);
      return a.topic_label.localeCompare(b.topic_label);
    });

  return topics;
}

function countActiveQuestionsForTopic_(ss, topicName) {
  const sheet = ss.getSheetByName(topicName);
  if (!sheet) return 0;

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return 0;

  const headers = values[0].map(h => String(h || "").trim());
  const activeIndex = headers.indexOf("active");
  const questionIdIndex = headers.indexOf("question_id");

  let count = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const questionId = questionIdIndex >= 0 ? String(row[questionIdIndex] || "").trim() : "";
    if (!questionId) continue;

    if (activeIndex >= 0) {
      if (!parseBooleanLoose_(row[activeIndex])) continue;
    }

    count += 1;
  }

  return count;
}

function loadTopicQuestions_(ss, topicName) {
  const sheet = ss.getSheetByName(topicName);
  if (!sheet) {
    throw new Error("Missing topic sheet: " + topicName);
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1);
  const questions = [];

  rows.forEach(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    if (!parseBooleanLoose_(obj.active)) return;

    const questionId = String(obj.question_id || "").trim();
    if (!questionId) return;

    questions.push({
      topic: topicName,
      question_id: questionId,
      question_type: String(obj.question_type || "").trim(),
      stem: String(obj.stem || "").trim(),
      correct_blob: String(obj.correct_blob || "").trim(),
      distractors_blob: String(obj.distractors_blob || "").trim(),
      rationale: String(obj.rationale || "").trim(),
      media_url: String(obj.media_url || "").trim(),
      active: true
    });
  });

  return questions;
}

function loadStudentHistoryMap_(ss, email) {
  const sheet = ss.getSheetByName(QUESTION_BANK_.HISTORY_SHEET);
  if (!sheet) return {};

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return {};

  const headers = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1);
  const map = {};

  rows.forEach(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });

    const rowEmail = normalizeEmail_(obj.email);
    if (rowEmail !== email) return;

    const questionId = String(obj.question_id || "").trim();
    if (!questionId) return;

    map[questionId] = {
      email: rowEmail,
      question_id: questionId,
      topic: String(obj.topic || "").trim(),
      times_seen: Number(obj.times_seen || 0),
      recent_streak_score: normalizeRecentStreakScore_(obj.recent_streak_score)
    };
  });

  return map;
}

function buildWeightedPool_(questions, historyMap) {
  return questions.map(q => {
    const history = historyMap[q.question_id];
    const recentStreakScore = history
      ? normalizeRecentStreakScore_(history.recent_streak_score)
      : 0;

    return {
      question: q,
      weight: weightForRecentStreakScore_(recentStreakScore)
    };
  });
}

function getOrCreateQuestionHistorySheet_(ss) {
  let sheet = ss.getSheetByName(QUESTION_BANK_.HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QUESTION_BANK_.HISTORY_SHEET);
  }
  return sheet;
}

function ensureQuestionHistoryHeaders_(sheet) {
  const requiredHeaders = [
    "email",
    "question_id",
    "topic",
    "times_seen",
    "recent_streak_score"
  ];

  const lastCol = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim());

  const hasAnyHeaderContent = existingHeaders.some(Boolean);

  if (!hasAnyHeaderContent) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const normalizedExisting = existingHeaders.slice(0, requiredHeaders.length);
  const mismatch = requiredHeaders.some((header, i) => normalizedExisting[i] !== header);

  if (mismatch) {
    throw new Error(
      'The "' + QUESTION_BANK_.HISTORY_SHEET + '" sheet headers must be exactly: ' +
      requiredHeaders.join(" | ")
    );
  }
}

function normalizeRecentStreakScore_(value) {
  const n = Number(value);
  if (!isFinite(n)) return 0;
  if (n <= -2) return -2;
  if (n >= 2) return 2;
  if (n < 0) return Math.ceil(n);
  if (n > 0) return Math.floor(n);
  return 0;
}

function weightForRecentStreakScore_(score) {
  const normalized = normalizeRecentStreakScore_(score);
  const weights = {
    "-2": 3,
    "-1": 2,
    "0": 1,
    "1": 0.5,
    "2": 0.25
  };
  return weights[String(normalized)] || 1;
}

function nextRecentStreakScore_(currentScore, isCorrect) {
  const current = normalizeRecentStreakScore_(currentScore);

  if (isCorrect) {
    if (current === -2) return -1;
    if (current === -1) return 1;
    if (current === 0) return 1;
    if (current === 1) return 2;
    return 2;
  }

  if (current === 2) return 1;
  if (current === 1) return -1;
  if (current === 0) return -1;
  if (current === -1) return -2;
  return -2;
}

function splitQuestionCountEvenly_(topics, total) {
  const result = {};
  const base = Math.floor(total / topics.length);
  const remainder = total % topics.length;

  topics.forEach((topic, i) => {
    result[topic] = base + (i < remainder ? 1 : 0);
  });

  return result;
}

function weightedRandomSampleWithoutReplacement_(pool, count) {
  const items = pool.slice();
  const chosen = [];

  count = Math.min(count, items.length);

  for (let i = 0; i < count; i++) {
    const index = pickWeightedIndex_(items);
    chosen.push(items[index].question);
    items.splice(index, 1);
  }

  return chosen;
}

function pickWeightedIndex_(items) {
  let total = 0;
  items.forEach(item => {
    total += Number(item.weight || 0);
  });

  if (total <= 0) {
    return Math.floor(Math.random() * items.length);
  }

  const r = Math.random() * total;
  let running = 0;

  for (let i = 0; i < items.length; i++) {
    running += Number(items[i].weight || 0);
    if (r < running) return i;
  }

  return items.length - 1;
}

function parseBooleanLoose_(value) {
  if (value === true) return true;
  if (value === false) return false;
  const s = String(value || "").trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes";
}

function upsertDailyPracticeTotal_(email, questionsToAdd) {
  const sheet = getQuestionBankSheet_(SHEETS_.DAILY_PRACTICE_TOTALS);
  ensureDailyPracticeTotalsHeaders_(sheet);

  const today = todayKey_(); // always yyyy-MM-dd string
  const values = sheet.getDataRange().getValues();

  // Update existing row if it exists
  for (let i = 1; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][0]);
    const rowDate = String(values[i][1] || "").trim();
    const rowQuestions = Number(values[i][2] || 0);

    if (rowEmail === email && rowDate === today) {
      const nextTotal = rowQuestions + Number(questionsToAdd || 0);
      sheet.getRange(i + 1, 2).setValue(today); // <- write string, not Date
      sheet.getRange(i + 1, 3).setValue(nextTotal);
      return {
        ok: true,
        questions_today: nextTotal
      };
    }
  }

  // Append new row if not found
  const initialTotal = Number(questionsToAdd || 0);
  sheet.appendRow([email, today, initialTotal]); // <- date stored as string

  return {
    ok: true,
    questions_today: initialTotal
  };
}

function ensureDailyPracticeTotalsHeaders_(sheet) {
  const requiredHeaders = [
    "email",
    "date",
    "questions_answered"
  ];

  const lastCol = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || "").trim());

  const hasAnyHeaderContent = existingHeaders.some(Boolean);

  if (!hasAnyHeaderContent) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const normalizedExisting = existingHeaders.slice(0, requiredHeaders.length);
  const mismatch = requiredHeaders.some((header, i) => normalizedExisting[i] !== header);

  if (mismatch) {
    throw new Error(
      'The "' + SHEETS_.DAILY_PRACTICE_TOTALS + '" sheet headers must be exactly: ' +
      requiredHeaders.join(" | ")
    );
  }
}

/* =========================================================
   INTERNAL HELPERS
   ========================================================= */

function shuffleArray_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = a[i];
    a[i] = a[j];
    a[j] = temp;
  }
  return a;
}

function ensureUserExists_(email) {
  const sheet = getSheet_(SHEETS_.USERS);
  const headerMap = getHeaderIndexMap_(sheet);
  const values = sheet.getDataRange().getValues();

  const emailCol = headerMap.email;
  const goalCol = headerMap.daily_goal;
  const activeCol = headerMap.active;
  const createdCol = headerMap.created_at;

  if (!emailCol) throw new Error('Missing "email" column in users sheet.');

  for (let i = 1; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][emailCol - 1]);
    if (rowEmail === email) {
      if (goalCol && !values[i][goalCol - 1]) {
        sheet.getRange(i + 1, goalCol).setValue(DEFAULT_DAILY_GOAL_);
      }
      if (activeCol) {
        sheet.getRange(i + 1, activeCol).setValue(true);
      }
      if (createdCol && !values[i][createdCol - 1]) {
        sheet.getRange(i + 1, createdCol).setValue(new Date());
      }
      return;
    }
  }

  appendUserRow_(sheet, {
    email: email,
    name: "",
    daily_goal: DEFAULT_DAILY_GOAL_,
    salt: "",
    password_hash: "",
    active: true,
    created_at: new Date()
  });
}

function appendUserRow_(sheet, userObj) {
  const headers = getSheetHeaders_(sheet);
  const row = headers.map(header => {
    if (Object.prototype.hasOwnProperty.call(userObj, header)) {
      return userObj[header];
    }
    return "";
  });
  sheet.appendRow(row);
}

function upsertProgressRow_(email, moduleKey, progressPercent, lastSection, completed) {
  const sheet = getSheet_(SHEETS_.PROGRESS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][0]);
    const rowModuleKey = String(values[i][1] || "").trim();

    if (rowEmail === email && rowModuleKey === moduleKey) {
      sheet.getRange(i + 1, 3).setValue(clampNumber_(Number(progressPercent || 0), 0, 100));
      sheet.getRange(i + 1, 4).setValue(String(lastSection || ""));
      sheet.getRange(i + 1, 5).setValue(!!completed);
      sheet.getRange(i + 1, 6).setValue(new Date());
      return;
    }
  }

  sheet.appendRow([
    email,
    moduleKey,
    clampNumber_(Number(progressPercent || 0), 0, 100),
    String(lastSection || ""),
    !!completed,
    new Date()
  ]);
}

function loadDailyPracticeTotalsForUser_(email) {
  const sheet = getQuestionBankSheet_(SHEETS_.DAILY_PRACTICE_TOTALS);
  ensureDailyPracticeTotalsHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  return values.slice(1)
  .filter(row => normalizeEmail_(row[0]) === email)
  .map(row => ({
    email: normalizeEmail_(row[0]),
    date: dateKeyFromValue_(row[1]), // <- ensures yyyy-MM-dd string
    questions_answered: Number(row[2] || 0)
  }));
}

function getQuestionsAnsweredForDate_(dailyRows, dateKey) {
  const normalizedDateKey = dateKeyFromValue_(dateKey); // ensures yyyy-MM-dd
  const row = (dailyRows || []).find(r => dateKeyFromValue_(r.date) === normalizedDateKey);
  return row ? Number(row.questions_answered || 0) : 0;
}

function computeDailyPracticeStreakData_(dailyRows, dailyGoal) {
  const map = {};

  (dailyRows || []).forEach(r => {
    const dateKey = String(r.date || "").trim();
    if (!dateKey) return;
    map[dateKey] = Number(r.questions_answered || 0);
  });

  const today = todayKey_();
  const todayCount = Number(map[today] || 0);
  const streakSecuredToday = todayCount >= dailyGoal;

  let cursor = new Date(today + "T12:00:00");
  if (!streakSecuredToday) {
    cursor.setDate(cursor.getDate() - 1);
  }

  let streakDays = 0;

  while (true) {
    const key = Utilities.formatDate(cursor, APP_TIMEZONE_, "yyyy-MM-dd");
    const count = Number(map[key] || 0);

    if (count >= dailyGoal) {
      streakDays += 1;
      cursor.setDate(cursor.getDate() - 1);
    } else {
      break;
    }
  }

  return {
    streak_days: streakDays,
    streak_secured_today: streakSecuredToday
  };
}

function computeStreakData_(sessionRows, dailyGoal) {
  const questionsByDate = {};

  sessionRows.forEach(r => {
    const dateKey = firstDateKeyFromSession_(r);
    if (!dateKey) return;
    const q = Number(r.total_questions || r.questions_answered || 0);
    questionsByDate[dateKey] = (questionsByDate[dateKey] || 0) + q;
  });

  const today = todayKey_();
  const todayCount = Number(questionsByDate[today] || 0);
  const streakSecuredToday = todayCount >= dailyGoal;

  let cursor = new Date(today + "T12:00:00");
  if (!streakSecuredToday) {
    cursor.setDate(cursor.getDate() - 1);
  }

  let streakDays = 0;
  while (true) {
    const key = Utilities.formatDate(cursor, APP_TIMEZONE_, "yyyy-MM-dd");
    const count = Number(questionsByDate[key] || 0);
    if (count >= dailyGoal) {
      streakDays += 1;
      cursor.setDate(cursor.getDate() - 1);
    } else {
      break;
    }
  }

  return {
    streak_days: streakDays,
    streak_secured_today: streakSecuredToday
  };
}

function firstDateKeyFromSession_(sessionRow) {
  return (
    dateKeyFromValue_(sessionRow.timestamp) ||
    dateKeyFromValue_(sessionRow.session_end) ||
    dateKeyFromValue_(sessionRow.session_start) ||
    ""
  );
}

function sumQuestionsForDate_(sessionRows, dateKey) {
  return sessionRows.reduce((sum, r) => {
    const rowDate = firstDateKeyFromSession_(r);
    if (rowDate !== dateKey) return sum;
    return sum + Number(r.total_questions || r.questions_answered || 0);
  }, 0);
}

function todayKey_() {
  return Utilities.formatDate(new Date(), APP_TIMEZONE_, "yyyy-MM-dd");
}

function dateKeyFromValue_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_TIMEZONE_, "yyyy-MM-dd");
  }

  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) return "";

  return Utilities.formatDate(parsed, APP_TIMEZONE_, "yyyy-MM-dd");
}

function logLoginAttempt_(email, success, reason) {
  const sheet = getSheet_(SHEETS_.LOGINS);
  sheet.appendRow([
    new Date(),
    email,
    !!success,
    String(reason || "")
  ]);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: "' + name + '"');
  return sheet;
}

function getSheetHeaders_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || !values.length) return [];
  return values[0].map(h => String(h || "").trim());
}

function getHeaderIndexMap_(sheet) {
  const headers = getSheetHeaders_(sheet);
  const map = {};
  headers.forEach((header, i) => {
    if (header) map[header] = i + 1;
  });
  return map;
}

function sheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || !values.length) return [];

  const headers = values[0].map(h => String(h || "").trim());

  return values
    .slice(1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
}

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeAssignmentKey_(value) {
  return String(value || "").trim().toLowerCase();
}

function toBool_(value) {
  if (typeof value === "boolean") return value;
  const v = String(value || "").trim().toLowerCase();
  return v === "true" || v === "1" || v === "yes";
}

function sanitizeDailyGoal_(value) {
  const n = Number(value);
  if (!isFinite(n)) return DEFAULT_DAILY_GOAL_;
  const rounded = Math.round(n);
  if (rounded < 1) return DEFAULT_DAILY_GOAL_;
  return rounded;
}

function clampNumber_(n, min, max) {
  const x = Number(n);
  if (isNaN(x)) return min;
  return Math.min(max, Math.max(min, x));
}

function round2_(n) {
  return Math.round(Number(n || 0) * 100) / 100;
}

function averageNumber_(arr) {
  const nums = arr.map(Number).filter(n => !isNaN(n));
  if (!nums.length) return 0;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function latestDateLikeValue_(arr) {
  let bestDate = null;

  (arr || []).forEach(v => {
    if (!v) return;
    const d = Object.prototype.toString.call(v) === "[object Date]" ? v : new Date(v);
    if (isNaN(d.getTime())) return;
    if (!bestDate || d.getTime() > bestDate.getTime()) {
      bestDate = d;
    }
  });

  return bestDate ? bestDate.toISOString() : "";
}

function fail_(message) {
  return {
    ok: false,
    error: String(message || "Unknown error.")
  };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================
   TEST HELPERS
   ========================================================= */

function testDailyPracticeTotal() {
  const result = upsertDailyPracticeTotal_("tjin@apu.edu", 10);
  Logger.log(JSON.stringify(result, null, 2));
}

function testDailyPracticeDashboardData() {
  const rows = loadDailyPracticeTotalsForUser_("tjin@apu.edu");
  const questionsToday = getQuestionsAnsweredForDate_(rows, todayKey_());
  const streakData = computeDailyPracticeStreakData_(rows, DEFAULT_DAILY_GOAL_);

  Logger.log(JSON.stringify({
    rows: rows,
    questions_today: questionsToday,
    streakData: streakData
  }, null, 2));
}

function testPracticeQuestions() {
  const result = getPracticeQuestions_({
    email: "tjin@apu.edu",
    topics: ["intro"],
    question_count: 5
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testSubmitPracticeResults() {
  const result = submitPracticeResults_({
    email: "tjin@apu.edu",
    question_results: [
      {
        question_id: "intro_q_001",
        topic: "intro",
        is_correct: true
      },
      {
        question_id: "intro_q_002",
        topic: "intro",
        is_correct: false
      }
    ]
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testPracticeTopics() {
  const result = getPracticeTopics_({});
  Logger.log(JSON.stringify(result, null, 2));
}

function testVerifyGoogleLoginConfig() {
  Logger.log("GOOGLE_CLIENT_ID_: " + GOOGLE_CLIENT_ID_);
  Logger.log("QUESTION_BANK_.SPREADSHEET_ID: " + QUESTION_BANK_.SPREADSHEET_ID);
}

function simulatePracticeSubmission() {
  const email = "tjin@apu.edu"; // test student email
  const dailyGoal = DEFAULT_DAILY_GOAL_; // use current daily goal
  const simulatedResults = [];

  // Create enough questions to meet or exceed daily goal
  for (let i = 1; i <= dailyGoal; i++) {
    simulatedResults.push({
      question_id: `sim_q_${i}`, // unique question IDs
      topic: "intro",
      is_correct: true
    });
  }

  // Submit the simulated practice session
  const result = submitPracticeResults_({
    email: email,
    question_results: simulatedResults
  });

  Logger.log("=== Simulated Submission Result ===");
  Logger.log(JSON.stringify(result, null, 2));

  // Load updated daily totals and streak
  const rows = loadDailyPracticeTotalsForUser_(email);
  const questionsToday = getQuestionsAnsweredForDate_(rows, todayKey_());
  const streakData = computeDailyPracticeStreakData_(rows, dailyGoal);

  Logger.log("=== Updated Daily Totals & Streak ===");
  Logger.log(JSON.stringify({
    rows: rows,
    questions_today: questionsToday,
    streakData: streakData
  }, null, 2));
}

/**
 * Submits student feedback for a question.
 * Increments counter for the type and appends free-text comment.
 *
 * @param {Object} feedback - Feedback object
 *   { email, topic, question_id, type, comment (optional) }
 */
function submitQuestionFeedback(feedback) {
  // Replace with your content spreadsheet ID
  const ss = SpreadsheetApp.openById(QUESTION_BANK_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(feedback.topic);
  if (!sheet) throw new Error('Topic sheet not found: ' + feedback.topic);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const typeIndex = headers.indexOf(feedback.type); // e.g., confusing
  const commentsIndex = headers.indexOf('student_comments');
  if (typeIndex === -1) throw new Error('Feedback type not found: ' + feedback.type);

  // Find row by question_id
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(feedback.question_id).trim()) {
      rowIndex = i + 1; // Apps Script is 1-indexed
      break;
    }
  }
  if (rowIndex === -1) throw new Error('Question not found: ' + feedback.question_id);

  // Increment numeric counter
  const currentVal = parseInt(sheet.getRange(rowIndex, typeIndex + 1).getValue()) || 0;
  sheet.getRange(rowIndex, typeIndex + 1).setValue(currentVal + 1);

  // Append free-text comment
  if (feedback.comment) {
    const existingComments = sheet.getRange(rowIndex, commentsIndex + 1).getValue() || '';
    const newComments = existingComments ? existingComments + '||' + feedback.comment : feedback.comment;
    sheet.getRange(rowIndex, commentsIndex + 1).setValue(newComments);
  }

  return { ok: true };
}