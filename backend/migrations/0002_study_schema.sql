-- Core D1 study schema for the Cloudflare migration.
-- Contains structure only: no participant records, access codes, or research data.
-- Main AQG and training data are separated where their schemas differ.

CREATE TABLE access_codes (
  participant_id TEXT PRIMARY KEY,
  code_hash TEXT NOT NULL UNIQUE,
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
  allow_aqg INTEGER NOT NULL DEFAULT 1 CHECK (allow_aqg IN (0, 1)),
  allow_training INTEGER NOT NULL DEFAULT 1 CHECK (allow_training IN (0, 1)),
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE auth_failure_state (
  client_key_hash TEXT PRIMARY KEY,
  failure_count INTEGER NOT NULL DEFAULT 0 CHECK (failure_count >= 0),
  locked_until TEXT,
  updated_at TEXT NOT NULL
);

CREATE TABLE aqg_participant_counters (
  participant_id TEXT PRIMARY KEY,
  next_session_number INTEGER NOT NULL DEFAULT 1 CHECK (next_session_number >= 1)
);

CREATE TABLE aqg_session_context_counters (
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  next_context_number INTEGER NOT NULL DEFAULT 1 CHECK (next_context_number >= 1),
  PRIMARY KEY (participant_id, session_id)
);

CREATE TABLE training_participant_counters (
  participant_id TEXT PRIMARY KEY,
  next_session_number INTEGER NOT NULL DEFAULT 1 CHECK (next_session_number >= 1)
);

CREATE TABLE aqg_live_sessions (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  revision INTEGER NOT NULL DEFAULT 0 CHECK (revision >= 0),
  last_save_id TEXT,
  last_event_at TEXT,
  context_id TEXT,
  status TEXT NOT NULL DEFAULT 'in_progress',
  session_start TEXT,
  last_activity_at TEXT,
  submitted_at TEXT,
  llm_product TEXT,
  llm_model TEXT,
  llm_description TEXT,
  model_used TEXT,
  course_context TEXT,
  question_context TEXT,
  extra_instructions TEXT,
  desired_questions TEXT,
  final_response TEXT,
  feedback_text TEXT,
  audio_object_key TEXT,
  audio_duration_seconds REAL,
  active_seconds REAL,
  progress_json TEXT,
  events_json TEXT,
  session_close_type TEXT,
  session_close_at TEXT,
  app_version TEXT,
  mode TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id)
);

CREATE INDEX idx_aqg_live_participant_activity
  ON aqg_live_sessions(participant_id, last_activity_at);

CREATE TABLE aqg_submissions (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  context_id TEXT,
  status TEXT NOT NULL DEFAULT 'submitted',
  session_start TEXT,
  submitted_at TEXT NOT NULL,
  llm_product TEXT,
  llm_model TEXT,
  llm_description TEXT,
  course_context TEXT,
  question_context TEXT,
  extra_instructions TEXT,
  desired_questions TEXT,
  final_response TEXT,
  feedback_text TEXT,
  audio_object_key TEXT,
  audio_duration_seconds REAL,
  active_seconds REAL,
  session_close_type TEXT,
  session_close_at TEXT,
  app_version TEXT,
  mode TEXT,
  created_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id)
);

CREATE INDEX idx_aqg_submissions_participant_time
  ON aqg_submissions(participant_id, submitted_at);

CREATE TABLE aqg_events (
  record_id TEXT PRIMARY KEY,
  event_id TEXT NOT NULL UNIQUE,
  batch_id TEXT,
  event_timestamp TEXT NOT NULL,
  ms_since_previous INTEGER,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  context_id TEXT,
  event_type TEXT NOT NULL,
  button_id TEXT,
  detail_text TEXT,
  detail_json TEXT,
  device_label TEXT,
  created_at TEXT NOT NULL
);

CREATE INDEX idx_aqg_events_session_time
  ON aqg_events(participant_id, session_id, event_timestamp);

CREATE INDEX idx_aqg_events_batch
  ON aqg_events(batch_id);

CREATE TABLE aqg_feedback (
  record_id TEXT PRIMARY KEY,
  feedback_id TEXT NOT NULL UNIQUE,
  participant_id TEXT NOT NULL,
  session_id TEXT,
  submitted_at TEXT NOT NULL,
  text_feedback TEXT,
  audio_object_key TEXT,
  audio_original_filename TEXT,
  audio_mime_type TEXT,
  audio_duration_seconds REAL,
  notification_status TEXT NOT NULL DEFAULT 'pending',
  created_at TEXT NOT NULL
);

CREATE INDEX idx_aqg_feedback_participant_time
  ON aqg_feedback(participant_id, submitted_at);

CREATE TABLE aqg_latest_settings (
  participant_id TEXT PRIMARY KEY,
  source_session_id TEXT,
  course_context TEXT,
  question_context TEXT,
  extra_instructions TEXT,
  desired_questions TEXT,
  updated_at TEXT NOT NULL
);

CREATE TABLE training_live_sessions (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  session_seq INTEGER,
  revision INTEGER NOT NULL DEFAULT 0 CHECK (revision >= 0),
  last_save_id TEXT,
  last_event_at TEXT,
  session_start TEXT,
  last_activity_at TEXT,
  status TEXT NOT NULL DEFAULT 'in_progress',
  app_version TEXT,
  content_version TEXT,
  current_card_number INTEGER,
  device_trail TEXT,
  progress_json TEXT,
  progress_saved_at TEXT,
  last_event_type TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id)
);

CREATE INDEX idx_training_live_participant_activity
  ON training_live_sessions(participant_id, last_activity_at);

CREATE TABLE training_live_items (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  content_version TEXT,
  item_number INTEGER NOT NULL,
  response_text TEXT,
  response_ms INTEGER,
  updated_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id, item_number)
);

CREATE INDEX idx_training_live_items_session
  ON training_live_items(participant_id, session_id, item_number);

CREATE TABLE training_submissions (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  session_seq INTEGER,
  session_start TEXT,
  session_end TEXT,
  duration_ms INTEGER,
  duration_formatted TEXT,
  active_seconds REAL,
  total_questions INTEGER,
  correct_first INTEGER,
  event_count INTEGER,
  item_count INTEGER,
  app_version TEXT,
  content_version TEXT,
  current_card_number INTEGER,
  device_trail TEXT,
  submitted_at TEXT NOT NULL,
  details_json TEXT,
  created_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id)
);

CREATE INDEX idx_training_submissions_participant_time
  ON training_submissions(participant_id, submitted_at);

CREATE TABLE training_submission_items (
  record_id TEXT PRIMARY KEY,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  content_version TEXT,
  item_number INTEGER NOT NULL,
  response_text TEXT,
  response_ms INTEGER,
  created_at TEXT NOT NULL,
  UNIQUE (participant_id, session_id, item_number)
);

CREATE INDEX idx_training_submission_items_session
  ON training_submission_items(participant_id, session_id, item_number);

CREATE TABLE training_events (
  record_id TEXT PRIMARY KEY,
  event_id TEXT NOT NULL UNIQUE,
  batch_id TEXT,
  event_timestamp TEXT NOT NULL,
  ms_since_previous INTEGER,
  participant_id TEXT NOT NULL,
  session_id TEXT NOT NULL,
  session_seq INTEGER,
  event_type TEXT NOT NULL,
  section_index INTEGER,
  card_index INTEGER,
  detail_text TEXT,
  detail_json TEXT,
  device_label TEXT,
  created_at TEXT NOT NULL
);

CREATE INDEX idx_training_events_session_time
  ON training_events(participant_id, session_id, event_timestamp);

CREATE INDEX idx_training_events_batch
  ON training_events(batch_id);

CREATE TABLE training_feedback (
  record_id TEXT PRIMARY KEY,
  feedback_id TEXT NOT NULL UNIQUE,
  participant_id TEXT NOT NULL,
  session_id TEXT,
  session_seq INTEGER,
  submitted_at TEXT NOT NULL,
  section_index INTEGER,
  section_title TEXT,
  card_index INTEGER,
  feedback_source TEXT,
  text_feedback TEXT,
  device_label TEXT,
  notification_status TEXT NOT NULL DEFAULT 'pending',
  created_at TEXT NOT NULL
);

CREATE INDEX idx_training_feedback_participant_time
  ON training_feedback(participant_id, submitted_at);

CREATE TABLE notification_outbox (
  notification_id TEXT PRIMARY KEY,
  notification_type TEXT NOT NULL,
  participant_id TEXT,
  session_id TEXT,
  payload_json TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending',
  attempt_count INTEGER NOT NULL DEFAULT 0 CHECK (attempt_count >= 0),
  created_at TEXT NOT NULL,
  sent_at TEXT,
  last_error TEXT
);

CREATE INDEX idx_notification_outbox_status
  ON notification_outbox(status, created_at);
