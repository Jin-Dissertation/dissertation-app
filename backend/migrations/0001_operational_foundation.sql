-- Initial operational foundation for dissertation-study-data.
-- Contains no participant records or credentials.

CREATE TABLE request_receipts (
  request_id TEXT PRIMARY KEY,
  operation TEXT NOT NULL,
  request_hash TEXT NOT NULL,
  response_status INTEGER NOT NULL,
  response_json TEXT,
  created_at TEXT NOT NULL
);

CREATE TABLE mirror_changes (
  sequence INTEGER PRIMARY KEY AUTOINCREMENT,
  feed_generation TEXT NOT NULL,
  dataset TEXT NOT NULL,
  record_id TEXT NOT NULL,
  revision INTEGER NOT NULL,
  operation TEXT NOT NULL,
  changed_at TEXT NOT NULL
);

CREATE INDEX idx_mirror_changes_sequence
  ON mirror_changes(sequence);

CREATE TABLE nonparticipant_button_counts (
  button_id TEXT PRIMARY KEY,
  press_count INTEGER NOT NULL DEFAULT 0
    CHECK (press_count >= 0),
  updated_at TEXT NOT NULL
);
