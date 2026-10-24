CREATE TABLE IF NOT EXISTS runs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts INTEGER NOT NULL,          -- unix ms
  ok INTEGER NOT NULL,          -- 1 success, 0 failed
  gstin TEXT,
  name TEXT,
  device TEXT,
  ip TEXT
);
CREATE INDEX IF NOT EXISTS runs_ts ON runs (ts);
