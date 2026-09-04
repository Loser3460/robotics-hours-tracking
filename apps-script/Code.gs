/**
 * Robotics Lab Attendance Tracker — Apps Script backend.
 *
 * This file is the entire server. It is bound to the tracking Spreadsheet and
 * deployed as a Web App (Deploy > New deployment > Web app, execute as Me,
 * access Anyone). The static frontend on GitHub Pages talks to it as a JSON API.
 *
 * FIRST-TIME SETUP
 *   1. Run initializeSheets() once from the editor. Grant the permissions it asks
 *      for. It creates the tabs, the Drive folder, and the Config defaults, and
 *      logs the generated admin PIN.
 *   2. Run installAutoCloseTrigger() once to schedule the nightly session closer.
 *   3. Deploy as a Web App and paste the /exec URL into the frontend's config.js.
 *
 * ============================ READ THIS BEFORE EDITING =====================
 *
 * CORS / Content-Type
 *   Every POST from the browser sends Content-Type: text/plain with a JSON
 *   string body, and this file parses it with JSON.parse(e.postData.contents).
 *   That looks wrong and it is deliberate. text/plain is one of the three
 *   "simple request" content types, so the browser sends it WITHOUT a CORS
 *   preflight. Apps Script Web Apps cannot answer an OPTIONS preflight — there
 *   is no doOptions() — so the moment a request becomes non-simple it fails
 *   before doPost() is ever called. Setting application/json "properly" on the
 *   client breaks every write in the app. Do not do it.
 *   For the same reason, no custom request headers (no Authorization, no
 *   X-Anything) — those also trigger a preflight. The admin PIN travels in the
 *   JSON body instead.
 *
 * Timestamps
 *   The server assigns every timestamp. The tablet's clock is not trusted.
 *   The only exceptions are offline-queued scans (syncQueue) and admin
 *   corrections, which carry a supplied timestamp and are ALWAYS written with
 *   flagged = true so the admin UI can show them as unverified.
 *
 * Append-only log
 *   Events is immutable. Nothing here ever rewrites or deletes an Events row,
 *   and no total_hours / currently_in / last_seen is ever stored. Every
 *   statistic is recomputed from the log on read. Corrections append a new row
 *   whose note supersedes or voids an earlier event_id — see resolveEvents_().
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var VERSION = '1.0.0';

var TAB_STUDENTS  = 'Students';
var TAB_EVENTS    = 'Events';
var TAB_SUMMARIES = 'Summaries';
var TAB_CONFIG    = 'Config';

/**
 * Column order is the contract. Rows are read by index, so never reorder these
 * or insert a column in the middle — append at the end if you must extend.
 */
var HEADERS = {};
HEADERS[TAB_STUDENTS]  = ['student_id', 'name', 'grade', 'active', 'created_at'];
HEADERS[TAB_EVENTS]    = ['event_id', 'student_id', 'timestamp', 'direction', 'source', 'flagged', 'note'];
HEADERS[TAB_SUMMARIES] = ['summary_id', 'student_id', 'session_date', 'text', 'photo_urls', 'submitted_at'];
HEADERS[TAB_CONFIG]    = ['key', 'value'];

var DRIVE_FOLDER_NAME = 'Robotics Lab Summaries';

var DIR_IN   = 'in';
var DIR_OUT  = 'out';
var DIR_VOID = 'void';   // marker row for a voided event; never a real presence event

var SOURCE_TABLET    = 'tablet';
var SOURCE_OFFLINE   = 'offline';
var SOURCE_ADMIN     = 'admin';
var SOURCE_AUTOCLOSE = 'auto-close';

var LOCK_TIMEOUT_MS  = 25000;
var MAX_PHOTOS       = 4;
var MAX_PHOTO_BYTES  = 8 * 1024 * 1024;
var MAX_SUMMARY_CHARS = 4000;

// Per-execution caches. Apps Script globals live for one execution only, which
// is exactly the lifetime we want: fresh on every request, free within one.
var CONFIG_CACHE = null;
var ID_COUNTER = 0;

// ---------------------------------------------------------------------------
// Response helpers — every response is {ok:true,data:...} or {ok:false,error:"..."}
// ---------------------------------------------------------------------------

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok_(data) {
  return { ok: true, data: data === undefined ? null : data };
}

/** Throw a client-facing error. doPost turns it into {ok:false,error:message}. */
function fail_(message) {
  var e = new Error(message);
  e.apiError = true;
  throw e;
}

/**
 * Like fail_, but marks the response {retryable:true}.
 *
 * Only for failures where the SAME request would plausibly succeed a moment
 * later — lock contention, essentially. The tablet queues a retryable scan
 * instead of discarding it, which matters because contention is exactly when
 * two students are at the door at once.
 */
function failRetryable_(message) {
  var e = new Error(message);
  e.apiError = true;
  e.retryable = true;
  throw e;
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/**
 * Health check. Open the /exec URL in a browser to confirm the deployment is
 * live and pointed at an initialized spreadsheet. Returns no student data.
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabs = {};
    var missing = [];
    var names = [TAB_STUDENTS, TAB_EVENTS, TAB_SUMMARIES, TAB_CONFIG];
    for (var i = 0; i < names.length; i++) {
      var sh = ss.getSheetByName(names[i]);
      if (!sh) { missing.push(names[i]); tabs[names[i]] = null; continue; }
      tabs[names[i]] = Math.max(0, sh.getLastRow() - 1);   // data rows, excluding header
    }
    var cfg = missing.length ? {} : readConfig_();
    return json_(ok_({
      service: 'robotics-lab-attendance',
      version: VERSION,
      time: nowIso_(),
      timezone: tz_(),
      spreadsheet: ss.getName(),
      rows: tabs,
      missing_tabs: missing,
      initialized: missing.length === 0 && !!cfg.drive_folder_id && !!cfg.admin_pin,
      admin_pin_set: !!cfg.admin_pin,
      drive_folder_set: !!cfg.drive_folder_id
    }));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/**
 * The single write endpoint. Body is a JSON string sent as text/plain — see the
 * CORS note at the top of this file before changing anything about that.
 */
function doPost(e) {
  var payload;
  try {
    if (!e || !e.postData || !e.postData.contents) fail_('empty request body');
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      fail_('body is not valid JSON');
    }
    if (!payload || typeof payload !== 'object') fail_('body must be a JSON object');

    var action = String(payload.action || '').trim();
    if (!action) fail_('missing action');

    switch (action) {
      // Tablet / student actions
      case 'scan':          return json_(actionScan_(payload));
      case 'enroll':        return json_(actionEnroll_(payload));
      case 'syncQueue':     return json_(actionSyncQueue_(payload));
      case 'submitSummary': return json_(actionSubmitSummary_(payload));
      case 'lookupStudent': return json_(actionLookupStudent_(payload));

      // Admin actions — every one of these requires the PIN in the body.
      case 'getRoster':     return json_(actionGetRoster_(payload));
      case 'getTimesheet':  return json_(actionGetTimesheet_(payload));
      case 'getStudent':    return json_(actionGetStudent_(payload));
      case 'getSummaries':  return json_(actionGetSummaries_(payload));
      case 'deleteSummary': return json_(actionDeleteSummary_(payload));
      case 'editEvent':     return json_(actionEditEvent_(payload));
      case 'deleteEvent':   return json_(actionDeleteEvent_(payload));
      case 'setActive':     return json_(actionSetActive_(payload));
      case 'setName':       return json_(actionSetName_(payload));

      default: fail_('unknown action: ' + action);
    }
  } catch (err) {
    if (err && err.apiError) {
      var body = { ok: false, error: err.message };
      if (err.retryable) body.retryable = true;
      return json_(body);
    }
    // Unexpected: log the stack for the execution log, return something terse.
    console.error('doPost failed', err && err.stack ? err.stack : err,
                  'action=' + (payload && payload.action));
    return json_({ ok: false, error: 'server error: ' + String(err && err.message ? err.message : err) });
  }
}

// ---------------------------------------------------------------------------
// Locking
// ---------------------------------------------------------------------------

/**
 * Run fn while holding the script lock. Two tablets scanning at the same instant
 * must not both read "last event = in" and both append an OUT, and two appends
 * must not target the same row. Every read-decide-write path goes through here,
 * with the READ inside the lock — locking only the write would not help.
 */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    // Retryable on purpose: another scan is mid-write. Reporting this as a
    // plain error made the tablet drop the scan outright.
    failRetryable_('the sheet is busy, try again');
  }
  try {
    return fn();
  } finally {
    // Push pending writes out before releasing, so the next holder reads them.
    try { SpreadsheetApp.flush(); } catch (flushErr) { console.error(flushErr); }
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Sheet access
// ---------------------------------------------------------------------------

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) fail_('missing tab "' + name + '" — run initializeSheets() once from the Apps Script editor');
  return sh;
}

/**
 * Read a whole tab in ONE getValues() call and hand back plain objects.
 *
 * Apps Script charges roughly a fixed round-trip per Range call, so a loop of
 * getRange().getValue() over 2,000 events takes tens of seconds while a single
 * getDataRange().getValues() takes milliseconds. Read once, work in memory,
 * write once. Nothing in this file may call getRange() inside a row loop.
 *
 * Each object carries _row (1-based sheet row) for the rare in-place update —
 * Students.active — which is the only mutable cell in the whole schema.
 */
function readTable_(name) {
  if (TABLE_MEMO_.hasOwnProperty(name)) return TABLE_MEMO_[name];
  var rows = cacheGetTable_(name);
  if (!rows) {
    // Note the generation BEFORE touching the sheet, and refuse to cache if it
    // moved while we were reading — see cachePutTable_.
    var gen = tableGen_(name);
    rows = loadTable_(name);
    cachePutTable_(name, rows, gen);
  }
  TABLE_MEMO_[name] = rows;
  return rows;
}

/** The uncached read. Only readTable_ and the cache layer call this. */
function loadTable_(name) {
  var values = sheet_(name).getDataRange().getValues();
  var header = HEADERS[name];
  if (!values.length) return [];

  // Column order is the contract; refuse to guess if someone reordered them.
  for (var c = 0; c < header.length; c++) {
    if (String(values[0][c]).trim().toLowerCase() !== header[c]) {
      fail_('tab "' + name + '" column ' + (c + 1) + ' should be "' + header[c] +
            '" but is "' + values[0][c] + '" — column order is the contract');
    }
  }

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (isBlankRow_(row)) continue;
    var obj = { _row: r + 1 };
    for (var i = 0; i < header.length; i++) obj[header[i]] = row[i];
    out.push(obj);
  }
  return out;
}

function isBlankRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) return false;
  }
  return true;
}

/**
 * Append rows to a tab in one setValues() call. Caller must hold the lock.
 *
 * This is also the single choke point for cache invalidation on the append-only
 * tabs: scan, enroll, syncQueue, submitSummary, editEvent, deleteEvent and
 * deleteSummary all reach the sheet through here, so none of them has to
 * remember to drop the key. The two in-place cell writes (Students.active and
 * Students.name) do not, and invalidate for themselves.
 */
function appendRows_(name, rows) {
  if (!rows.length) return;
  var sh = sheet_(name);
  var width = HEADERS[name].length;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows);
  invalidateTable_(name);
}

// ---------------------------------------------------------------------------
// Table cache
//
// Reading a tab costs a Sheets round-trip — tens to hundreds of milliseconds,
// and every admin page load wants Students, Events and Summaries. CacheService
// is an order of magnitude faster, so the tabs are serialized into it and the
// sheet is only touched on a miss or after a write.
//
// Three things this layer must get right, in order of how badly they bite:
//
//   1. It is never a source of truth. Every failure — a miss, an eviction, a
//      quota error, a value that will not parse — falls through to the sheet.
//      Nothing here may throw.
//   2. A write invalidates before the next read. Every append goes through
//      appendRows_ above; the two setValue writes call invalidateTable_
//      themselves. A stale roster showing an archived student as active is the
//      failure this guards against.
//   3. Cached rows must be INDISTINGUISHABLE from freshly read ones. A sheet
//      cell can hold a Date, and JSON turns Dates into strings — which would
//      quietly change summaryDate_(), because it branches on `instanceof Date`
//      and formats a real Date in the local timezone but reads a string
//      verbatim. Dates are therefore tagged on the way in and revived on the
//      way out, so both paths hand back the same objects.
// ---------------------------------------------------------------------------

/** Per-execution memo: two readTable_(Events) in one request cost one parse. */
var TABLE_MEMO_ = {};

var CACHE_TTL_SEC = 21600;      // six hours, the CacheService maximum
var CACHE_CHUNK_CHARS = 32000;  // a cached value caps at 100KB; UTF-8 is <=3B/char
var CACHE_MAX_CHUNKS = 40;      // ~1.2MB of JSON; past that, just read the sheet

function cacheKey_(name) {
  return 'tbl.v1.' + name;
}

function genKey_(name) {
  return 'gen.v1.' + name;
}

/**
 * The tab's current generation — an opaque token that changes on every write.
 *
 * This exists to close a race that dropping the key alone does not. Reads take
 * no lock, so a reader can load the sheet, a writer can then commit and
 * invalidate, and the reader can afterwards store what it read — parking data
 * that was already stale in a six-hour cache. Comparing the generation across
 * the sheet read catches exactly that overlap.
 *
 * A missing token (evicted, or a cold script) reads as "unknowable", which
 * makes the caller skip caching this once. Losing a cache fill is free;
 * serving a stale roster is not.
 */
function tableGen_(name) {
  try {
    var cache = CacheService.getScriptCache();
    var g = cache.get(genKey_(name));
    if (!g) {
      g = Utilities.getUuid();
      cache.put(genKey_(name), g, CACHE_TTL_SEC);
    }
    return g;
  } catch (err) {
    return null;
  }
}

/** Date -> {__d: epoch millis}. Everything else is already JSON-safe. */
function encodeRows_(rows) {
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var src = rows[i], dst = {};
    for (var k in src) {
      if (!src.hasOwnProperty(k)) continue;
      var v = src[k];
      if (v instanceof Date) {
        var t = v.getTime();
        dst[k] = isNaN(t) ? '' : { __d: t };
      } else {
        dst[k] = v;
      }
    }
    out.push(dst);
  }
  return out;
}

function decodeRows_(rows) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    for (var k in row) {
      if (!row.hasOwnProperty(k)) continue;
      var v = row[k];
      if (v && typeof v === 'object' && typeof v.__d === 'number') row[k] = new Date(v.__d);
    }
  }
  return rows;
}

/**
 * Read a tab out of the cache, or null on any kind of miss.
 *
 * Chunked, because a value caps at 100KB and Events outgrows that. The manifest
 * holds "<count>|<stamp>" and each chunk key carries the same stamp, so a
 * partially evicted or concurrently rewritten set can never be reassembled into
 * a Frankenstein table — a stamp mismatch or a missing chunk is just a miss.
 */
function cacheGetTable_(name) {
  try {
    var cache = CacheService.getScriptCache();
    var key = cacheKey_(name);
    var manifest = cache.get(key);
    if (!manifest) return null;

    var parts = String(manifest).split('|');
    var count = Number(parts[0]);
    var stamp = parts[1];
    if (!stamp || !(count > 0)) return null;

    var keys = [];
    for (var c = 0; c < count; c++) keys.push(key + '.' + stamp + '.' + c);
    var got = cache.getAll(keys);

    var json = '';
    for (var i = 0; i < keys.length; i++) {
      var piece = got[keys[i]];
      if (piece === null || piece === undefined) return null;   // evicted mid-set
      json += piece;
    }
    return decodeRows_(JSON.parse(json));
  } catch (err) {
    console.warn('table cache read failed for ' + name + ': ' + err);
    return null;
  }
}

function cachePutTable_(name, rows, gen) {
  try {
    if (!gen) return;                          // generation unknown: do not cache
    var json = JSON.stringify(encodeRows_(rows));
    var count = Math.ceil(json.length / CACHE_CHUNK_CHARS) || 1;
    if (count > CACHE_MAX_CHUNKS) return;      // too big to be worth the round-trips

    var key = cacheKey_(name);
    var stamp = Utilities.getUuid().replace(/-/g, '').substring(0, 8);
    var map = {};
    for (var c = 0; c < count; c++) {
      map[key + '.' + stamp + '.' + c] = json.substring(c * CACHE_CHUNK_CHARS,
                                                        (c + 1) * CACHE_CHUNK_CHARS);
    }
    var cache = CacheService.getScriptCache();
    // Someone committed a write while we were reading the sheet, so what we are
    // holding may already be out of date. Drop it rather than cache it.
    if (cache.get(genKey_(name)) !== gen) return;

    // Chunks first, manifest last: a reader that catches us mid-write finds no
    // manifest and reads the sheet, rather than finding one that points at
    // chunks which are not there yet.
    cache.putAll(map, CACHE_TTL_SEC);
    cache.put(key, count + '|' + stamp, CACHE_TTL_SEC);
  } catch (err) {
    console.warn('table cache write failed for ' + name + ': ' + err);
  }
}

/**
 * Drop a tab from the cache. Called after every write, inside the lock, so the
 * next reader misses and reloads.
 *
 * Only the manifest is deleted. The orphaned chunks expire on their own and can
 * never be read again, because the next write picks a fresh stamp.
 */
function invalidateTable_(name) {
  delete TABLE_MEMO_[name];
  if (name === TAB_CONFIG) CONFIG_CACHE = null;
  try {
    // Make the write durable FIRST. Apps Script buffers setValues, so dropping
    // the key while the row is still pending would let the next reader load the
    // sheet, miss the new row, and cache that for six hours.
    SpreadsheetApp.flush();
  } catch (err) {
    console.error('flush before invalidate failed: ' + err);
  }
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(cacheKey_(name));
    // Bump the generation so an in-flight reader that started before this write
    // discards what it read instead of caching it.
    cache.put(genKey_(name), Utilities.getUuid(), CACHE_TTL_SEC);
  } catch (err) {
    console.warn('table cache invalidate failed for ' + name + ': ' + err);
  }
}

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

function readConfig_() {
  if (CONFIG_CACHE) return CONFIG_CACHE;
  var cfg = {};
  var rows = readTable_(TAB_CONFIG);
  for (var i = 0; i < rows.length; i++) {
    var k = String(rows[i].key || '').trim();
    if (k) cfg[k] = rows[i].value;
  }
  CONFIG_CACHE = cfg;
  return cfg;
}

function cfgStr_(key, fallback) {
  var v = readConfig_()[key];
  return (v === undefined || v === null || v === '') ? fallback : String(v).trim();
}

function cfgNum_(key, fallback) {
  var n = Number(cfgStr_(key, ''));
  return isFinite(n) && n > 0 ? n : fallback;
}

/** Script timezone, overridable from Config so date math has one source. */
function tz_() {
  return cfgStr_('timezone', Session.getScriptTimeZone());
}

// ---------------------------------------------------------------------------
// Small utilities
// ---------------------------------------------------------------------------

function newId_(prefix) {
  ID_COUNTER++;
  return prefix + '_' + Date.now().toString(36) + '_' +
         ID_COUNTER.toString(36) + Math.floor(Math.random() * 46656).toString(36);
}

/**
 * Repair a student ID READ FROM THE SHEET. Sheets stores 7-digit IDs as text,
 * but a column that was pasted in before the format was set can come back as
 * the number 234567, so pad it back out.
 *
 * This is for sheet values only — never for client input, which must be exactly
 * seven digits already. Padding an incoming "12" into "0000012" would turn a
 * misread QR code into a lookup against someone else's ID.
 */
function normId_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (/^\d+$/.test(s) && s.length < 7) s = ('0000000' + s).slice(-7);
  return s;
}

/** Validate a client-supplied student ID. Strict: the QR payload is 7 digits. */
function requireId_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!/^\d{7}$/.test(s)) fail_('student_id must be exactly 7 digits');
  return s;
}

function requireStr_(payload, field, max) {
  var v = payload[field];
  if (v === undefined || v === null || String(v).trim() === '') fail_('missing ' + field);
  var s = String(v).trim();
  if (max && s.length > max) fail_(field + ' is too long (max ' + max + ' characters)');
  return s;
}

function truthy_(v) {
  if (v === true) return true;
  if (typeof v === 'number') return v !== 0;
  return /^(true|yes|1|y)$/i.test(String(v).trim());
}

function nowIso_() {
  return new Date().toISOString();
}

/** Timestamps are stored as UTC ISO 8601 strings so they round-trip exactly. */
function toIso_(d) {
  return d.toISOString();
}

function parseTs_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v || '').trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Local calendar day of an instant, e.g. "2026-09-02". */
function dateKey_(d) {
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

/** Local wall-clock time of an instant, e.g. "18:42". */
function timeKey_(d) {
  return Utilities.formatDate(d, tz_(), 'HH:mm');
}

/** The instant of local "HH:mm" on the same local day as refDate. */
function localTimeOn_(refDate, hhmm) {
  var m = /^(\d{1,2}):(\d{2})$/.exec(String(hhmm || '').trim());
  if (!m) return null;
  var day = Utilities.formatDate(refDate, tz_(), 'yyyy/MM/dd');
  var d = new Date(day + ' ' + ('0' + m[1]).slice(-2) + ':' + m[2] + ':00');
  return isNaN(d.getTime()) ? null : d;
}

function minutesBetween_(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 60000);
}

// ---------------------------------------------------------------------------
// The event log
// ---------------------------------------------------------------------------

function toEvent_(row) {
  var ts = parseTs_(row.timestamp);
  var id = String(row.event_id || '').trim();
  if (!id || !ts) return null;                     // malformed row: ignore, never crash a scan
  return {
    event_id: id,
    student_id: normId_(row.student_id),
    ts: ts,
    timestamp: toIso_(ts),
    direction: String(row.direction || '').trim().toLowerCase(),
    source: String(row.source || '').trim(),
    flagged: truthy_(row.flagged),
    note: String(row.note || '').trim()
  };
}

/**
 * Turn raw Events rows into the effective log, oldest first.
 *
 * The sheet is append-only, so a correction is itself a row whose note starts
 * with a machine-readable marker naming the event it replaces:
 *
 *   supersedes:<event_id> — <reason>   a replacement event; the target drops out
 *   void:<event_id> — <reason>         a tombstone; both rows drop out
 *
 * Nothing is ever edited or deleted in the sheet — the original rows stay
 * visible there forever, they just stop counting here. Chains work: if B
 * supersedes A and C supersedes B, only C survives.
 */
function resolveEvents_(rows) {
  var kept = [];
  var dropped = {};
  for (var i = 0; i < rows.length; i++) {
    var ev = toEvent_(rows[i]);
    if (!ev) continue;
    var m = /^(void|supersedes):(\S+)/.exec(ev.note);
    if (m) {
      dropped[m[2]] = true;
      if (m[1] === 'void') dropped[ev.event_id] = true;   // the tombstone is not itself an event
    }
    kept.push(ev);
  }
  var out = [];
  for (var j = 0; j < kept.length; j++) {
    if (dropped[kept[j].event_id]) continue;
    if (kept[j].direction !== DIR_IN && kept[j].direction !== DIR_OUT) continue;
    out.push(kept[j]);
  }
  return sortByTs_(out);
}

function sortByTs_(events) {
  return events.sort(function (a, b) {
    var d = a.ts.getTime() - b.ts.getTime();
    return d !== 0 ? d : String(a.event_id).localeCompare(String(b.event_id));  // stable on ties
  });
}

function groupByStudent_(events) {
  var by = {};
  for (var i = 0; i < events.length; i++) {
    var sid = events[i].student_id;
    if (!by[sid]) by[sid] = [];
    by[sid].push(events[i]);
  }
  return by;
}

/**
 * Pair a single student's events into sessions.
 *
 *   closed   — an IN matched by an OUT: the normal case, the only one that counts
 *   open     — the last IN, still unmatched: the student is in the lab right now
 *   unclosed — an IN followed by another IN; the first never got an OUT
 *   orphan   — an OUT with no IN before it
 *   reversed — an OUT that lands BEFORE its IN, which would give negative time
 *
 * The last three are broken logs, not attendance. They are still returned, every
 * one of them carrying needs_review and a null duration, because the admin page
 * has to be able to show a coach what is wrong and let them fix it. Dropping
 * them here would make an orphan OUT invisible in every view — the row would
 * simply never appear, and nobody would know to correct it.
 *
 * 'reversed' cannot happen on the normal path: resolveEvents_ sorts ascending,
 * so an OUT always lands at or after the IN it closes, and a negative duration
 * is impossible by construction. The branch is kept as a guard — if the ordering
 * ever changes, a bad pair degrades into a flagged row with a null duration
 * instead of silently reporting "-142 min" or clamping to zero.
 */
function buildSessions_(events) {
  var sessions = [];
  var open = null;
  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    if (e.direction === DIR_IN) {
      if (open) sessions.push(makeSession_(open, null, 'unclosed'));
      open = e;
    } else {
      if (open) {
        var backwards = e.ts.getTime() < open.ts.getTime();
        sessions.push(makeSession_(open, e, backwards ? 'reversed' : 'closed'));
        open = null;
      } else {
        sessions.push(makeOrphan_(e));
      }
    }
  }
  if (open) sessions.push(makeSession_(open, null, 'open'));
  return sessions;
}

/** An OUT with nothing before it. There is no start time, so there is no duration. */
function makeOrphan_(outEv) {
  return {
    student_id: outEv.student_id,
    date: dateKey_(outEv.ts),
    status: 'orphan',
    in_event_id: null,
    out_event_id: outEv.event_id,
    in_time: null,
    out_time: outEv.timestamp,
    in_clock: null,
    out_clock: timeKey_(outEv.ts),
    minutes: null,
    flagged: !!outEv.flagged,
    needs_review: true,
    sources: [outEv.source]
  };
}

function makeSession_(inEv, outEv, status) {
  // Only a well-formed session has a duration. 'unclosed' and 'reversed' report
  // null rather than a guess, so nothing downstream can sum a broken row.
  var minutes = null;
  if (status === 'closed') minutes = Math.max(0, minutesBetween_(inEv.ts, outEv.ts));
  else if (status === 'open') minutes = Math.max(0, minutesBetween_(inEv.ts, new Date()));
  return {
    student_id: inEv.student_id,
    date: dateKey_(inEv.ts),
    status: status,
    in_event_id: inEv.event_id,
    out_event_id: outEv ? outEv.event_id : null,
    in_time: inEv.timestamp,
    out_time: outEv ? outEv.timestamp : null,
    in_clock: timeKey_(inEv.ts),
    out_clock: outEv ? timeKey_(outEv.ts) : null,
    minutes: minutes,
    flagged: !!(inEv.flagged || (outEv && outEv.flagged)),
    needs_review: status !== 'closed' && status !== 'open',
    sources: outEv ? [inEv.source, outEv.source] : [inEv.source]
  };
}

/** Totals are ALWAYS computed here from the log — never stored anywhere. */
function summarize_(sessions) {
  var closedMinutes = 0, closedCount = 0, flagged = 0, review = 0, openSession = null;
  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    if (s.status === 'closed') { closedMinutes += s.minutes; closedCount++; }
    if (s.status === 'open') openSession = s;
    if (s.flagged) flagged++;
    if (s.needs_review) review++;
  }
  return {
    total_minutes: closedMinutes,
    total_hours: Math.round(closedMinutes / 6) / 10,
    session_count: closedCount,
    flagged_sessions: flagged,
    sessions_needing_review: review,
    currently_in: !!openSession,
    current_session_minutes: openSession ? openSession.minutes : null,
    current_session_since: openSession ? openSession.in_time : null
  };
}

function eventRow_(ev) {
  return [ev.event_id, ev.student_id, ev.timestamp, ev.direction,
          ev.source, ev.flagged === true, ev.note || ''];
}

// ---------------------------------------------------------------------------
// Students
// ---------------------------------------------------------------------------

function studentIndex_(rows) {
  var by = {};
  for (var i = 0; i < rows.length; i++) {
    var sid = normId_(rows[i].student_id);
    if (!sid) continue;
    by[sid] = {
      student_id: sid,
      name: String(rows[i].name || '').trim(),
      grade: rows[i].grade === '' || rows[i].grade === null ? '' : String(rows[i].grade).trim(),
      active: truthy_(rows[i].active),
      created_at: rows[i].created_at instanceof Date ? toIso_(rows[i].created_at)
                                                     : String(rows[i].created_at || ''),
      _row: rows[i]._row
    };
  }
  return by;
}

function publicStudent_(s) {
  return { student_id: s.student_id, name: s.name, grade: s.grade,
           active: s.active, created_at: s.created_at };
}

// ---------------------------------------------------------------------------
// Action: scan
// ---------------------------------------------------------------------------

/**
 * A student tapped their card on the tablet.
 *
 * Everything — the roster lookup, the debounce check, the last-event lookup and
 * the append — happens inside one lock. If the debounce ran outside it, two
 * tablets reading the same "last event" a millisecond apart would both decide
 * to append and the student would be scanned twice.
 *
 * Returns one of:
 *   {status:"unknown"}    ID not on the roster; the tablet offers enrollment
 *   {status:"inactive"}   on the roster but archived
 *   {status:"debounced"}  scanned again within the debounce window; nothing written
 *   {status:"ok"}         event appended; carries name, direction, duration
 */
function actionScan_(p) {
  var studentId = requireId_(p.student_id);
  var source = String(p.source || SOURCE_TABLET).trim() || SOURCE_TABLET;

  return withLock_(function () {
    var students = studentIndex_(readTable_(TAB_STUDENTS));
    var student = students[studentId];
    if (!student) return ok_({ status: 'unknown', student_id: studentId });
    if (!student.active) {
      return ok_({ status: 'inactive', student_id: studentId, name: student.name });
    }

    var mine = groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)))[studentId] || [];
    var last = mine.length ? mine[mine.length - 1] : null;
    var now = new Date();

    var debounceSec = cfgNum_('scan_debounce_seconds', 15);
    if (last && (now.getTime() - last.ts.getTime()) < debounceSec * 1000) {
      // Double-read of the same card, or a student handing the card to a friend
      // too fast. Say so and write nothing.
      return ok_({
        status: 'debounced',
        student_id: studentId,
        name: student.name,
        direction: last.direction,
        seconds_ago: Math.round((now.getTime() - last.ts.getTime()) / 1000)
      });
    }

    // Presence is derived from the log, never from a stored flag.
    var direction = (last && last.direction === DIR_IN) ? DIR_OUT : DIR_IN;

    var ev = {
      event_id: newId_('evt'),
      student_id: studentId,
      timestamp: toIso_(now),          // server clock, always
      direction: direction,
      source: source,
      flagged: false,
      note: ''
    };
    appendRows_(TAB_EVENTS, [eventRow_(ev)]);

    var durationMinutes = null;
    if (direction === DIR_OUT && last) durationMinutes = Math.max(0, minutesBetween_(last.ts, now));

    return ok_({
      status: 'ok',
      student_id: studentId,
      name: student.name,
      direction: direction,
      event_id: ev.event_id,
      timestamp: ev.timestamp,
      clock: timeKey_(now),
      duration_minutes: durationMinutes,
      session_start: direction === DIR_OUT && last ? last.timestamp : null
    });
  });
}

// ---------------------------------------------------------------------------
// Action: enroll
// ---------------------------------------------------------------------------

/** Add a student to the roster. Called from the tablet after an unknown scan. */
function actionEnroll_(p) {
  var studentId = requireId_(p.student_id);
  var name = requireStr_(p, 'name', 80);
  var grade = p.grade === undefined || p.grade === null ? '' : String(p.grade).trim();

  return withLock_(function () {
    var students = studentIndex_(readTable_(TAB_STUDENTS));
    if (students[studentId]) {
      fail_('student_id ' + studentId + ' is already enrolled as ' + students[studentId].name);
    }
    var created = nowIso_();
    appendRows_(TAB_STUDENTS, [[studentId, name, grade, true, created]]);
    return ok_({
      student_id: studentId, name: name, grade: grade, active: true, created_at: created
    });
  });
}

// ---------------------------------------------------------------------------
// Action: syncQueue
// ---------------------------------------------------------------------------

/**
 * Replay scans the tablet captured while offline.
 *
 * These are the one case where a client timestamp is written, so every row here
 * gets flagged = true and source = "offline" — the admin UI shows them as
 * unverified. Replay is idempotent: a retry after a dropped response will not
 * duplicate rows, because we skip any queued scan whose
 * student/second/direction already exists in the raw sheet (raw, not resolved,
 * so a scan an admin already voided does not come back from the dead).
 *
 * Direction comes from the tablet when the queue recorded one — that is what
 * the student was actually told at the door, and rewriting it would falsify
 * their evening. When the queue has no direction, it is derived by merging the
 * scan into the student's existing timeline and taking the opposite of whatever
 * precedes it.
 */
function actionSyncQueue_(p) {
  var queue = p.scans || p.queue || p.events;
  if (!queue || !queue.length) fail_('syncQueue needs a non-empty scans array');
  if (queue.length > 500) fail_('too many queued scans in one request (max 500)');

  return withLock_(function () {
    var students = studentIndex_(readTable_(TAB_STUDENTS));
    var rawRows = readTable_(TAB_EVENTS);
    var existing = resolveEvents_(rawRows);
    var byStudent = groupByStudent_(existing);

    // Dedupe index over the RAW sheet, keyed to the second.
    var seen = {};
    for (var r = 0; r < rawRows.length; r++) {
      var raw = toEvent_(rawRows[r]);
      if (raw) seen[dedupeKey_(raw.student_id, raw.ts, raw.direction)] = true;
    }

    var accepted = [];
    var rejected = [];
    var duplicates = 0;
    var now = new Date();

    for (var i = 0; i < queue.length; i++) {
      var q = queue[i] || {};
      var sid = normId_(q.student_id);
      if (!/^\d{7}$/.test(sid)) { rejected.push({ index: i, reason: 'bad student_id' }); continue; }
      if (!students[sid])       { rejected.push({ index: i, student_id: sid, reason: 'not on roster' }); continue; }
      var ts = parseTs_(q.timestamp);
      if (!ts)                  { rejected.push({ index: i, student_id: sid, reason: 'unparseable timestamp' }); continue; }
      var dir = String(q.direction || '').trim().toLowerCase();
      accepted.push({
        index: i,
        student_id: sid,
        ts: ts,
        claimed: (dir === DIR_IN || dir === DIR_OUT) ? dir : null,
        ahead: ts.getTime() > now.getTime() + 60000    // tablet clock running fast
      });
    }

    // Merge each student's queued scans into their timeline, then walk it in
    // order so directions alternate correctly even for out-of-order arrivals.
    var pending = {};
    for (var a = 0; a < accepted.length; a++) {
      var sidA = accepted[a].student_id;
      if (!pending[sidA]) pending[sidA] = [];
      pending[sidA].push(accepted[a]);
    }

    var rows = [];
    var written = [];
    for (var sid2 in pending) {
      if (!pending.hasOwnProperty(sid2)) continue;
      var timeline = (byStudent[sid2] || []).slice();
      var queued = pending[sid2].slice().sort(function (x, y) { return x.ts.getTime() - y.ts.getTime(); });

      for (var k = 0; k < queued.length; k++) {
        var item = queued[k];
        var prev = lastBefore_(timeline, item.ts);
        var derived = (prev && prev.direction === DIR_IN) ? DIR_OUT : DIR_IN;
        var direction = item.claimed || derived;

        var key = dedupeKey_(sid2, item.ts, direction);
        if (seen[key]) { duplicates++; continue; }
        seen[key] = true;

        var notes = ['offline replay'];
        if (item.claimed && item.claimed !== derived) {
          // Another tablet was online and logged this student in between. Write
          // what the offline tablet recorded and say so; buildSessions_ shows
          // the result as needs_review rather than quietly picking a winner.
          notes.push('tablet recorded "' + item.claimed + '", timeline implies "' + derived + '"');
        }
        if (item.ahead) notes.push('client clock ahead of server');

        var ev = {
          event_id: newId_('evt'),
          student_id: sid2,
          timestamp: toIso_(item.ts),    // CLIENT timestamp — hence flagged below
          direction: direction,
          source: SOURCE_OFFLINE,
          flagged: true,                 // unverified clock; never write these unflagged
          note: notes.join('; ')
        };
        rows.push(eventRow_(ev));
        written.push({ index: item.index, event_id: ev.event_id, student_id: sid2,
                       direction: direction, timestamp: ev.timestamp });
        timeline.push({ ts: item.ts, direction: direction, event_id: ev.event_id, student_id: sid2 });
        timeline = sortByTs_(timeline);
      }
    }

    appendRows_(TAB_EVENTS, rows);
    return ok_({
      received: queue.length,
      written: rows.length,
      skipped_duplicates: duplicates,
      rejected: rejected,
      events: written
    });
  });
}

function dedupeKey_(studentId, ts, direction) {
  return studentId + '|' + Math.floor(ts.getTime() / 1000) + '|' + direction;
}

/** Latest event strictly before ts, in an ascending-sorted list. */
function lastBefore_(sortedEvents, ts) {
  var found = null;
  for (var i = 0; i < sortedEvents.length; i++) {
    if (sortedEvents[i].ts.getTime() < ts.getTime()) found = sortedEvents[i];
    else break;
  }
  return found;
}

// ---------------------------------------------------------------------------
// Action: lookupStudent
// ---------------------------------------------------------------------------

/**
 * Confirm an ID belongs to somebody, and describe their lab time on one day.
 *
 * The summary page needs this before a student writes anything: it shows the
 * name back so they know they typed the right number, and it prefills the
 * session they are writing about.
 *
 * Deliberately unauthenticated, like `scan` and `enroll`. It exposes no more
 * than `scan` already does — a name for an ID you had to know, plus that day's
 * in/out times — and unlike `scan` it writes nothing at all, so pointing the
 * summary page at it can never create attendance. It returns ONLY the requested
 * day: no history, no totals, no roster listing.
 */
function actionLookupStudent_(p) {
  var studentId = requireId_(p.student_id);
  var date = String(p.date || '').trim();
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) fail_('date must be YYYY-MM-DD');

  var student = studentIndex_(readTable_(TAB_STUDENTS))[studentId];
  if (!student) return ok_({ found: false, student_id: studentId });

  var out = {
    found: true,
    student_id: studentId,
    name: student.name,
    grade: student.grade,
    active: student.active
  };

  if (date) {
    var events = groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)))[studentId] || [];
    var sessions = [];
    var all = buildSessions_(events);
    for (var i = 0; i < all.length; i++) {
      if (all[i].date !== date) continue;
      sessions.push({
        status: all[i].status,
        in_time: all[i].in_time,
        out_time: all[i].out_time,
        in_clock: all[i].in_clock,
        out_clock: all[i].out_clock,
        minutes: all[i].minutes,
        flagged: all[i].flagged
      });
    }
    out.date = date;
    out.sessions = sessions;

    var summaries = resolveSummaries_(readTable_(TAB_SUMMARIES));
    var already = 0;
    for (var s = 0; s < summaries.length; s++) {
      if (normId_(summaries[s].student_id) !== studentId) continue;
      if (summaryDate_(summaries[s]) === date) already++;
    }
    out.summaries_on_date = already;
  }

  return ok_(out);
}

/**
 * Summaries with voided ones removed.
 *
 * Same append-only shape as the event log: deleting a summary appends a
 * tombstone whose text is "void:<summary_id> — reason", and both rows drop out
 * here while staying in the sheet for the record.
 */
function resolveSummaries_(rows) {
  var dropped = {};
  var kept = [];
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i].summary_id || '').trim();
    if (!id) continue;
    var m = /^void:(\S+)/.exec(String(rows[i].text || '').trim());
    if (m) {
      dropped[m[1]] = true;
      dropped[id] = true;          // the tombstone is not itself a summary
    }
    kept.push(rows[i]);
  }
  var out = [];
  for (var j = 0; j < kept.length; j++) {
    if (!dropped[String(kept[j].summary_id).trim()]) out.push(kept[j]);
  }
  return out;
}

/** session_date can come back as text or as a Date, depending on the cell. */
function summaryDate_(row) {
  return row.session_date instanceof Date
    ? dateKey_(row.session_date)
    : String(row.session_date || '').trim();
}

// ---------------------------------------------------------------------------
// Action: submitSummary
// ---------------------------------------------------------------------------

/**
 * A student submits what they worked on, with photos, from their phone.
 *
 * Photos go to Drive BEFORE the lock is taken. An upload can take several
 * seconds and the lock is shared with the scanner — holding it through a Drive
 * round-trip would stall the tablet at the door.
 */
function actionSubmitSummary_(p) {
  var studentId = requireId_(p.student_id);
  var text = requireStr_(p, 'text', MAX_SUMMARY_CHARS);
  var photos = p.photos || [];
  if (!(photos instanceof Array)) fail_('photos must be an array');
  if (photos.length > MAX_PHOTOS) fail_('too many photos (max ' + MAX_PHOTOS + ')');

  var students = studentIndex_(readTable_(TAB_STUDENTS));
  var student = students[studentId];
  if (!student) fail_('student_id ' + studentId + ' is not on the roster');
  if (!student.active) fail_(student.name + ' is archived and cannot submit summaries');

  var sessionDate = String(p.session_date || '').trim();
  if (sessionDate && !/^\d{4}-\d{2}-\d{2}$/.test(sessionDate)) fail_('session_date must be YYYY-MM-DD');
  if (!sessionDate) sessionDate = dateKey_(new Date());

  // --- Drive uploads, outside the lock ---
  var urls = [];
  if (photos.length) {
    var folder = summariesFolder_();
    for (var i = 0; i < photos.length; i++) {
      var blob = decodePhoto_(photos[i], studentId, sessionDate, i);
      var file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {
        // Some Workspace domains forbid link sharing. The file is still saved;
        // the admin will see it when signed in to the account that owns it.
        console.warn('could not set link sharing on ' + file.getId() + ': ' + shareErr);
      }
      urls.push('https://drive.google.com/uc?export=view&id=' + file.getId());
    }
  }

  // --- Sheet write, inside the lock ---
  return withLock_(function () {
    var row = {
      summary_id: newId_('sum'),
      submitted_at: nowIso_()          // server clock
    };
    appendRows_(TAB_SUMMARIES, [[
      row.summary_id, studentId, sessionDate, text, urls.join(', '), row.submitted_at
    ]]);
    return ok_({
      summary_id: row.summary_id,
      student_id: studentId,
      name: student.name,
      session_date: sessionDate,
      photo_urls: urls,
      submitted_at: row.submitted_at
    });
  });
}

/**
 * Accepts either a bare base64 string, a data: URL, or
 * {name, mimeType, data}. Returns a Drive-ready blob.
 */
function decodePhoto_(item, studentId, sessionDate, index) {
  var data, mime = 'image/jpeg', name = '';
  if (typeof item === 'string') {
    data = item;
  } else if (item && typeof item === 'object') {
    data = item.data || item.base64 || '';
    mime = String(item.mimeType || item.type || mime);
    name = String(item.name || '');
  } else {
    fail_('photo ' + (index + 1) + ' is not a string or object');
  }

  var m = /^data:([^;,]+);base64,(.*)$/.exec(String(data));
  if (m) { mime = m[1]; data = m[2]; }
  data = String(data).replace(/\s+/g, '');
  if (!data) fail_('photo ' + (index + 1) + ' is empty');
  if (!/^image\//.test(mime)) fail_('photo ' + (index + 1) + ' is not an image (' + mime + ')');

  var bytes;
  try {
    bytes = Utilities.base64Decode(data);
  } catch (decodeErr) {
    fail_('photo ' + (index + 1) + ' is not valid base64');
  }
  if (bytes.length > MAX_PHOTO_BYTES) {
    fail_('photo ' + (index + 1) + ' is too large (' + Math.round(bytes.length / 1048576) +
          ' MB, max ' + Math.round(MAX_PHOTO_BYTES / 1048576) + ' MB) — resize it on the phone first');
  }
  if (!name) {
    name = sessionDate + '_' + studentId + '_' + (index + 1) + extForMime_(mime);
  }
  return Utilities.newBlob(bytes, mime, name);
}

function extForMime_(mime) {
  if (/png/i.test(mime)) return '.png';
  if (/webp/i.test(mime)) return '.webp';
  if (/heic|heif/i.test(mime)) return '.heic';
  if (/gif/i.test(mime)) return '.gif';
  return '.jpg';
}

/** The Drive folder from Config, falling back to a lookup and finally a create. */
function summariesFolder_() {
  var id = cfgStr_('drive_folder_id', '');
  if (id) {
    try { return DriveApp.getFolderById(id); }
    catch (e) { console.warn('Config drive_folder_id ' + id + ' is unreachable: ' + e); }
  }
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  fail_('Drive folder is not set up — run initializeSheets() once from the editor');
}

// ---------------------------------------------------------------------------
// Admin actions
// ---------------------------------------------------------------------------

/**
 * The PIN travels in the JSON body, not in a header — an Authorization header
 * would trigger the CORS preflight that Apps Script cannot answer. See the note
 * at the top of this file.
 *
 * This is a shared PIN on a class roster, not real authentication. It keeps
 * students out of the admin page; it is not a secret worth defending against
 * anyone determined. Never return the PIN in a response.
 */
function requirePin_(p) {
  var expected = cfgStr_('admin_pin', '');
  if (!expected) fail_('admin PIN is not configured — set admin_pin in the Config tab');
  var given = String(p.pin === undefined || p.pin === null ? '' : p.pin).trim();
  if (given !== String(expected)) {
    Utilities.sleep(500);            // token slowdown against guessing a 6-digit PIN
    fail_('incorrect PIN');
  }
}

/** Roster with per-student totals, all recomputed from the log. */
function actionGetRoster_(p) {
  requirePin_(p);
  var includeInactive = p.include_inactive === undefined ? true : truthy_(p.include_inactive);

  // The admin page needs four columns — id, name, grade, active — and derives
  // every statistic from the timesheet it fetches anyway. Saying so lets us
  // skip reading Events, resolving them and building every student's sessions
  // for a payload that would be thrown away. The tablet's roster prime still
  // wants currently_in, so it does not pass this and gets the full shape.
  var slim = truthy_(p.slim);

  var students = studentIndex_(readTable_(TAB_STUDENTS));
  var byStudent = slim ? null : groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)));

  var roster = [];
  var inLab = 0;
  for (var sid in students) {
    if (!students.hasOwnProperty(sid)) continue;
    var s = students[sid];
    if (!s.active && !includeInactive) continue;
    // Four columns in slim mode, literally: created_at is on the record but no
    // view has ever drawn it.
    var entry = slim
      ? { student_id: s.student_id, name: s.name, grade: s.grade, active: s.active }
      : publicStudent_(s);
    if (!slim) {
      var events = byStudent[sid] || [];
      var stats = summarize_(buildSessions_(events));
      if (s.active && stats.currently_in) inLab++;
      for (var k in stats) if (stats.hasOwnProperty(k)) entry[k] = stats[k];
    }
    roster.push(entry);
  }
  roster.sort(function (a, b) { return a.name.localeCompare(b.name); });

  return ok_({
    students: roster,
    count: roster.length,
    currently_in_lab: slim ? null : inLab,
    generated_at: nowIso_()
  });
}

/**
 * Sessions across the whole team, optionally narrowed by date range or student.
 * from/to are inclusive local calendar days, YYYY-MM-DD.
 */
function actionGetTimesheet_(p) {
  requirePin_(p);
  var from = String(p.from || '').trim();
  var to   = String(p.to || '').trim();
  if (from && !/^\d{4}-\d{2}-\d{2}$/.test(from)) fail_('from must be YYYY-MM-DD');
  if (to && !/^\d{4}-\d{2}-\d{2}$/.test(to)) fail_('to must be YYYY-MM-DD');
  var onlyStudent = p.student_id ? requireId_(p.student_id) : null;
  // Opt-in, so an older deployed page that has not asked for it keeps getting
  // the full shape. The admin page asks.
  var slim = truthy_(p.slim);

  var students = studentIndex_(readTable_(TAB_STUDENTS));
  var byStudent = groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)));

  var sessions = [];
  var totals = [];
  for (var sid in byStudent) {
    if (!byStudent.hasOwnProperty(sid)) continue;
    if (onlyStudent && sid !== onlyStudent) continue;
    var student = students[sid];
    var all = buildSessions_(byStudent[sid]);
    var inRange = [];
    for (var i = 0; i < all.length; i++) {
      var s = all[i];
      if (from && s.date < from) continue;
      if (to && s.date > to) continue;
      s.name = student ? student.name : '(not on roster)';
      if (slim) {
        // The admin timesheet renders in_time/out_time itself and reads grade
        // off the roster entry, so shipping a local-clock copy and a duplicate
        // grade on every one of a season's sessions is dead weight.
        delete s.in_clock;
        delete s.out_clock;
      } else {
        s.grade = student ? student.grade : '';
      }
      inRange.push(s);
      sessions.push(s);
    }
    if (inRange.length && !slim) {
      var stats = summarize_(inRange);
      totals.push({
        student_id: sid,
        name: student ? student.name : '(not on roster)',
        grade: student ? student.grade : '',
        active: student ? student.active : false,
        total_minutes: stats.total_minutes,
        total_hours: stats.total_hours,
        session_count: stats.session_count,
        flagged_sessions: stats.flagged_sessions,
        sessions_needing_review: stats.sessions_needing_review
      });
    }
  }

  sessions.sort(function (a, b) {
    return a.in_time < b.in_time ? 1 : (a.in_time > b.in_time ? -1 : 0);   // newest first
  });
  totals.sort(function (a, b) { return b.total_minutes - a.total_minutes; });

  var grandMinutes = 0, flagged = 0, review = 0;
  for (var j = 0; j < sessions.length; j++) {
    if (sessions[j].status === 'closed') grandMinutes += sessions[j].minutes;
    if (sessions[j].flagged) flagged++;
    if (sessions[j].needs_review) review++;
  }

  return ok_({
    from: from || null,
    to: to || null,
    sessions: sessions,
    // Recomputed client-side in slim mode — it is a group-by over rows the
    // client already has, not something worth a second copy on the wire.
    totals_by_student: slim ? [] : totals,
    total_minutes: grandMinutes,
    total_hours: Math.round(grandMinutes / 6) / 10,
    flagged_sessions: flagged,
    sessions_needing_review: review,
    generated_at: nowIso_()
  });
}

/** One student: profile, stats, sessions, raw event log, summaries. */
function actionGetStudent_(p) {
  requirePin_(p);
  var studentId = requireId_(p.student_id);

  var students = studentIndex_(readTable_(TAB_STUDENTS));
  var student = students[studentId];
  if (!student) fail_('no student with id ' + studentId);

  var events = groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)))[studentId] || [];
  var sessions = buildSessions_(events).reverse();     // newest first
  var stats = summarize_(sessions);

  var eventList = [];
  for (var i = events.length - 1; i >= 0; i--) {
    eventList.push({
      event_id: events[i].event_id, timestamp: events[i].timestamp,
      clock: timeKey_(events[i].ts), date: dateKey_(events[i].ts),
      direction: events[i].direction, source: events[i].source,
      flagged: events[i].flagged, note: events[i].note
    });
  }

  var summaries = [];
  var summaryRows = resolveSummaries_(readTable_(TAB_SUMMARIES));
  for (var s = 0; s < summaryRows.length; s++) {
    if (normId_(summaryRows[s].student_id) !== studentId) continue;
    summaries.push({
      summary_id: String(summaryRows[s].summary_id || ''),
      session_date: summaryDate_(summaryRows[s]),
      text: String(summaryRows[s].text || ''),
      photo_urls: splitUrls_(summaryRows[s].photo_urls),
      submitted_at: summaryRows[s].submitted_at instanceof Date
        ? toIso_(summaryRows[s].submitted_at) : String(summaryRows[s].submitted_at || '')
    });
  }
  summaries.sort(function (a, b) { return a.submitted_at < b.submitted_at ? 1 : -1; });

  var out = { student: publicStudent_(student), sessions: sessions,
              events: eventList, summaries: summaries, generated_at: nowIso_() };
  for (var k in stats) if (stats.hasOwnProperty(k)) out[k] = stats[k];
  return ok_(out);
}

/**
 * Correct an event — by APPENDING a replacement, never by editing the row.
 *
 * The new row carries note "supersedes:<event_id> — <reason>", which makes
 * resolveEvents_() drop the original from every computed stat while the
 * original stays in the sheet as an audit trail. The corrected timestamp comes
 * from a human, so the replacement is flagged = true like any unverified time.
 *
 * Takes event_id plus at least one of timestamp (ISO or YYYY-MM-DD HH:mm) and
 * direction, plus an optional reason.
 */
/**
 * One student's sessions, rebuilt from the log the caller already has in hand.
 *
 * editEvent and deleteEvent both append a row and then owe the admin page an
 * answer precise enough that it never has to re-fetch. Returning just the new
 * event would not be enough: superseding a check-in can re-pair the sessions
 * around it, and only the server knows how. So they hand back this student's
 * whole rebuilt timeline, computed in memory from `raw` plus the row just
 * written — no second sheet read, and no guessing on the client.
 */
function studentSessionsAfter_(raw, appended, studentId) {
  var live = resolveEvents_(raw.concat(appended));
  var mine = groupByStudent_(live)[studentId] || [];
  var sessions = buildSessions_(mine);
  var students = studentIndex_(readTable_(TAB_STUDENTS));
  var student = students[studentId];
  var name = student ? student.name : '(not on roster)';
  for (var i = 0; i < sessions.length; i++) {
    sessions[i].name = name;
    delete sessions[i].in_clock;    // matches the slim timesheet shape
    delete sessions[i].out_clock;
  }
  sessions.sort(function (a, b) {
    return a.in_time < b.in_time ? 1 : (a.in_time > b.in_time ? -1 : 0);   // newest first
  });
  return sessions;
}

function actionEditEvent_(p) {
  requirePin_(p);
  var targetId = requireStr_(p, 'event_id', 64);
  var newDirection = String(p.direction || '').trim().toLowerCase();
  var hasDirection = !!newDirection;
  if (hasDirection && newDirection !== DIR_IN && newDirection !== DIR_OUT) {
    fail_('direction must be "in" or "out"');
  }
  var newTs = p.timestamp ? parseTs_(p.timestamp) : null;
  if (p.timestamp && !newTs) fail_('could not parse timestamp "' + p.timestamp + '"');
  if (!hasDirection && !newTs) fail_('editEvent needs a new timestamp, a new direction, or both');
  var reason = String(p.reason || p.note || '').trim();

  return withLock_(function () {
    var raw = readTable_(TAB_EVENTS);
    var live = resolveEvents_(raw);
    var target = findEvent_(live, targetId);
    if (!target) {
      fail_(existsInRaw_(raw, targetId)
        ? 'event ' + targetId + ' has already been corrected or voided'
        : 'no event with id ' + targetId);
    }

    var ev = {
      event_id: newId_('evt'),
      student_id: target.student_id,
      timestamp: toIso_(newTs || target.ts),
      direction: hasDirection ? newDirection : target.direction,
      source: SOURCE_ADMIN,
      flagged: true,                      // hand-entered time: unverified by definition
      note: 'supersedes:' + targetId + (reason ? ' — ' + reason : '')
    };
    appendRows_(TAB_EVENTS, [eventRow_(ev)]);

    return ok_({
      corrected: { event_id: target.event_id, timestamp: target.timestamp, direction: target.direction },
      replacement: ev,
      student_id: target.student_id,
      // Everything the admin page needs to patch its state without re-reading.
      sessions: studentSessionsAfter_(raw, [ev], target.student_id)
    });
  });
}

/**
 * Remove an event from the record — again by APPENDING, not deleting.
 *
 * Writes a tombstone row with direction "void" and note "void:<event_id>".
 * resolveEvents_() then drops both the tombstone and its target, so the event
 * stops counting while both rows remain in the sheet forever.
 */
function actionDeleteEvent_(p) {
  requirePin_(p);
  var targetId = requireStr_(p, 'event_id', 64);
  var reason = String(p.reason || p.note || '').trim();

  return withLock_(function () {
    var raw = readTable_(TAB_EVENTS);
    var live = resolveEvents_(raw);
    var target = findEvent_(live, targetId);
    if (!target) {
      fail_(existsInRaw_(raw, targetId)
        ? 'event ' + targetId + ' has already been corrected or voided'
        : 'no event with id ' + targetId);
    }

    var ev = {
      event_id: newId_('void'),
      student_id: target.student_id,
      timestamp: nowIso_(),               // when the void was recorded, server clock
      direction: DIR_VOID,
      source: SOURCE_ADMIN,
      flagged: true,
      note: 'void:' + targetId + (reason ? ' — ' + reason : '')
    };
    appendRows_(TAB_EVENTS, [eventRow_(ev)]);

    return ok_({
      voided: { event_id: target.event_id, timestamp: target.timestamp, direction: target.direction },
      tombstone_id: ev.event_id,
      student_id: target.student_id,
      sessions: studentSessionsAfter_(raw, [ev], target.student_id)
    });
  });
}

/**
 * Every summary, for the admin page. Fetched once at unlock alongside the
 * roster and the timesheet, so opening a student's detail view needs no
 * further request.
 */
function actionGetSummaries_(p) {
  requirePin_(p);
  var studentId = p.student_id ? requireId_(p.student_id) : null;
  var rows = resolveSummaries_(readTable_(TAB_SUMMARIES));
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var sid = normId_(rows[i].student_id);
    if (!sid) continue;
    if (studentId && sid !== studentId) continue;
    out.push({
      summary_id: String(rows[i].summary_id || ''),
      student_id: sid,
      session_date: summaryDate_(rows[i]),
      text: String(rows[i].text || ''),
      photo_urls: splitUrls_(rows[i].photo_urls),
      submitted_at: rows[i].submitted_at instanceof Date
        ? toIso_(rows[i].submitted_at) : String(rows[i].submitted_at || '')
    });
  }
  out.sort(function (a, b) { return a.submitted_at < b.submitted_at ? 1 : -1; });
  return ok_({ summaries: out, count: out.length, generated_at: nowIso_() });
}

/**
 * Remove a summary — by appending a tombstone, never by editing the sheet.
 *
 * This is the coach's only recourse when something inappropriate is posted, and
 * since anyone with the URL can submit a summary under any ID, it needs to be
 * genuinely effective: the photos are moved to the Drive trash too. Trash, not
 * a permanent delete, so a mistake is recoverable for 30 days.
 */
function actionDeleteSummary_(p) {
  requirePin_(p);
  var summaryId = requireStr_(p, 'summary_id', 64);
  var reason = String(p.reason || p.note || '').trim();
  var trashPhotos = p.trash_photos === undefined ? true : truthy_(p.trash_photos);

  return withLock_(function () {
    var rows = readTable_(TAB_SUMMARIES);
    var live = resolveSummaries_(rows);
    var target = null;
    for (var i = 0; i < live.length; i++) {
      if (String(live[i].summary_id).trim() === summaryId) { target = live[i]; break; }
    }
    if (!target) {
      var existed = false;
      for (var j = 0; j < rows.length; j++) {
        if (String(rows[j].summary_id).trim() === summaryId) existed = true;
      }
      fail_(existed ? 'summary ' + summaryId + ' has already been deleted'
                    : 'no summary with id ' + summaryId);
    }

    var urls = splitUrls_(target.photo_urls);
    var trashed = [];
    if (trashPhotos) {
      for (var u = 0; u < urls.length; u++) {
        var id = driveIdFromUrl_(urls[u]);
        if (!id) continue;
        try {
          DriveApp.getFileById(id).setTrashed(true);
          trashed.push(id);
        } catch (err) {
          console.warn('could not trash ' + id + ': ' + err);
        }
      }
    }

    appendRows_(TAB_SUMMARIES, [[
      newId_('void'),
      normId_(target.student_id),
      summaryDate_(target),
      'void:' + summaryId + (reason ? ' — ' + reason : ''),
      '',
      nowIso_()
    ]]);

    return ok_({
      deleted: summaryId,
      student_id: normId_(target.student_id),
      session_date: summaryDate_(target),
      photos_trashed: trashed.length,
      photos_total: urls.length
    });
  });
}

/** Pull the Drive file id out of whichever URL shape we stored. */
function driveIdFromUrl_(url) {
  var m = /[?&]id=([A-Za-z0-9_-]+)/.exec(url) || /\/d\/([A-Za-z0-9_-]+)/.exec(url);
  return m ? m[1] : null;
}

function splitUrls_(value) {
  var parts = String(value || '').split(',');
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    if (parts[i].trim()) out.push(parts[i].trim());
  }
  return out;
}

function findEvent_(events, eventId) {
  for (var i = 0; i < events.length; i++) if (events[i].event_id === eventId) return events[i];
  return null;
}

function existsInRaw_(rawRows, eventId) {
  for (var i = 0; i < rawRows.length; i++) {
    if (String(rawRows[i].event_id || '').trim() === eventId) return true;
  }
  return false;
}

/**
 * Archive or restore a student. This is the one mutable cell in the schema:
 * Students.active is a roster flag, not attendance data. History is untouched —
 * an archived student keeps every event and every hour they ever logged.
 */
function actionSetActive_(p) {
  requirePin_(p);
  var studentId = requireId_(p.student_id);
  if (p.active === undefined || p.active === null) fail_('missing active (true or false)');
  var active = truthy_(p.active);

  return withLock_(function () {
    var students = studentIndex_(readTable_(TAB_STUDENTS));
    var student = students[studentId];
    if (!student) fail_('no student with id ' + studentId);
    // Column 4 is `active` — see HEADERS[Students]; column order is the contract.
    sheet_(TAB_STUDENTS).getRange(student._row, 4).setValue(active);
    invalidateTable_(TAB_STUDENTS);      // an in-place write, so not via appendRows_

    // The whole record, not just the field that moved. The admin page patches
    // its local roster with this and never re-fetches, so anything it needs has
    // to come back here.
    student.active = active;
    return ok_({ student_id: studentId, name: student.name, active: active,
                 student: publicStudent_(student) });
  });
}

/**
 * Correct a student's name. Like `active`, `name` is roster data, not
 * attendance data — a misspelling at enrollment, a legal name change or a
 * student who goes by something else is a fact about the person, not about a
 * night in the lab. Nothing in Events stores a name (every row keys off
 * student_id, and names are joined in on read), so a rename is invisible to
 * every hour, session and statistic: the log is untouched and the totals come
 * out identical.
 *
 * There is no history of the old name. If a coach needs one, the fix is a
 * `note` on the Config tab, not a mutable column here.
 */
function actionSetName_(p) {
  requirePin_(p);
  var studentId = requireId_(p.student_id);
  var name = requireStr_(p, 'name', 80);

  return withLock_(function () {
    var students = studentIndex_(readTable_(TAB_STUDENTS));
    var student = students[studentId];
    if (!student) fail_('no student with id ' + studentId);

    var previous = student.name;
    if (name !== previous) {
      // Column 2 is `name` — see HEADERS[Students]; column order is the contract.
      sheet_(TAB_STUDENTS).getRange(student._row, 2).setValue(name);
      invalidateTable_(TAB_STUDENTS);    // an in-place write, so not via appendRows_
    }
    student.name = name;
    return ok_({ student_id: studentId, name: name, previous_name: previous,
                 changed: name !== previous, student: publicStudent_(student) });
  });
}

// ---------------------------------------------------------------------------
// Setup — run these by hand from the Apps Script editor
// ---------------------------------------------------------------------------

/**
 * One-time setup. Safe to run again at any point: it only ever creates what is
 * missing. It never clears a tab, never rewrites a Config value that is already
 * set, and never touches existing data rows.
 */
function initializeSheets() {
  return withLock_(function () {
    CONFIG_CACHE = null;
    var ss = ss_();
    var report = { created_tabs: [], repaired_headers: [], config_added: [], notes: [] };

    var tabs = [TAB_STUDENTS, TAB_EVENTS, TAB_SUMMARIES, TAB_CONFIG];
    for (var i = 0; i < tabs.length; i++) {
      var name = tabs[i];
      var header = HEADERS[name];
      var sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        report.created_tabs.push(name);
      }
      // Write headers only if row 1 is blank or does not match. Never touch row 2+.
      var current = sh.getLastColumn()
        ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), header.length)).getValues()[0]
        : [];
      var needsHeader = false;
      for (var c = 0; c < header.length; c++) {
        if (String(current[c] || '').trim().toLowerCase() !== header[c]) { needsHeader = true; break; }
      }
      if (needsHeader) {
        if (sh.getLastRow() > 1) {
          report.notes.push('tab "' + name + '" has data but wrong headers — headers rewritten, ' +
                            'CHECK that the existing columns line up with ' + header.join(', '));
        }
        sh.getRange(1, 1, 1, header.length).setValues([header]);
        report.repaired_headers.push(name);
      }
      sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    // Text formatting where Sheets would otherwise "helpfully" coerce values:
    // 7-digit IDs must not become numbers (leading zeros) and ISO timestamps
    // must not become spreadsheet dates.
    sheet_(TAB_STUDENTS).getRange('A:A').setNumberFormat('@');
    sheet_(TAB_STUDENTS).getRange('E:E').setNumberFormat('@');
    sheet_(TAB_EVENTS).getRange('B:C').setNumberFormat('@');
    sheet_(TAB_SUMMARIES).getRange('B:C').setNumberFormat('@');
    sheet_(TAB_SUMMARIES).getRange('F:F').setNumberFormat('@');
    sheet_(TAB_CONFIG).getRange('B:B').setNumberFormat('@');
    sheet_(TAB_EVENTS).setColumnWidth(3, 200);
    sheet_(TAB_SUMMARIES).setColumnWidth(4, 400);

    // Drop the default empty "Sheet1" once the real tabs exist.
    var leftover = ss.getSheetByName('Sheet1');
    if (leftover && leftover.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(leftover);
      report.notes.push('removed the empty default Sheet1');
    }

    // --- Drive folder ---
    var existingCfg = {};
    var cfgRows = readTable_(TAB_CONFIG);
    for (var r = 0; r < cfgRows.length; r++) {
      var key = String(cfgRows[r].key || '').trim();
      if (key) existingCfg[key] = String(cfgRows[r].value === null ? '' : cfgRows[r].value).trim();
    }

    var folder = null;
    if (existingCfg.drive_folder_id) {
      try { folder = DriveApp.getFolderById(existingCfg.drive_folder_id); }
      catch (e) { report.notes.push('configured drive_folder_id was unreachable, finding another'); }
    }
    if (!folder) {
      var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
      folder = it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
    }

    // --- Config defaults: only fill in keys that are missing ---
    var defaults = {
      admin_pin: String(Math.floor(100000 + Math.random() * 900000)),
      lab_open: '15:00',
      lab_close: '21:00',
      max_session_minutes: '360',
      scan_debounce_seconds: '15',
      timezone: Session.getScriptTimeZone(),
      drive_folder_id: folder.getId()
    };
    var newRows = [];
    for (var k in defaults) {
      if (!defaults.hasOwnProperty(k)) continue;
      if (existingCfg[k]) continue;
      newRows.push([k, defaults[k]]);
      report.config_added.push(k);
    }
    // drive_folder_id may exist but point somewhere stale; repair it in place.
    if (existingCfg.drive_folder_id && existingCfg.drive_folder_id !== folder.getId()) {
      for (var cr = 0; cr < cfgRows.length; cr++) {
        if (String(cfgRows[cr].key).trim() === 'drive_folder_id') {
          sheet_(TAB_CONFIG).getRange(cfgRows[cr]._row, 2).setValue(folder.getId());
          report.notes.push('repaired drive_folder_id');
        }
      }
    }
    appendRows_(TAB_CONFIG, newRows);
    CONFIG_CACHE = null;

    report.drive_folder = { name: folder.getName(), id: folder.getId(), url: folder.getUrl() };
    report.admin_pin = existingCfg.admin_pin || defaults.admin_pin;
    report.spreadsheet_url = ss.getUrl();

    Logger.log('--- Robotics Lab Attendance: initializeSheets ---');
    Logger.log('Tabs created:      ' + (report.created_tabs.join(', ') || 'none (already existed)'));
    Logger.log('Headers written:   ' + (report.repaired_headers.join(', ') || 'none'));
    Logger.log('Config keys added: ' + (report.config_added.join(', ') || 'none'));
    Logger.log('Drive folder:      ' + folder.getName() + '  ' + folder.getUrl());
    Logger.log('ADMIN PIN:         ' + report.admin_pin + '   (Config tab, key admin_pin — change it there)');
    for (var n = 0; n < report.notes.length; n++) Logger.log('Note: ' + report.notes[n]);
    Logger.log('Next: run installAutoCloseTrigger(), then Deploy > New deployment > Web app.');

    return report;
  });
}

/**
 * Install the nightly auto-close trigger. Idempotent — removes any trigger it
 * previously installed before adding the new one. Run again after changing
 * lab_close in Config.
 */
function installAutoCloseTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === 'autoCloseOpenSessions') {
      ScriptApp.deleteTrigger(existing[i]);
      removed++;
    }
  }
  // Apps Script time triggers fire within an hour of the requested time, so aim
  // for the hour after lab close. autoCloseOpenSessions backdates the OUT to
  // lab close itself, so the imprecision never lands in a student's hours.
  var close = cfgStr_('lab_close', '21:00');
  var hour = (parseInt(close.split(':')[0], 10) + 1) % 24;
  ScriptApp.newTrigger('autoCloseOpenSessions').timeBased().atHour(hour).everyDays(1).create();
  Logger.log('Auto-close trigger installed for ~' + hour + ':00 ' + tz_() +
             ' (removed ' + removed + ' previous)');
  return { removed: removed, hour: hour, lab_close: close };
}

// ---------------------------------------------------------------------------
// Nightly maintenance
// ---------------------------------------------------------------------------

/**
 * Close out anyone who forgot to scan on the way out. Runs on a nightly trigger.
 *
 * The synthetic OUT is backdated to lab close (or to the in-time plus
 * max_session_minutes, whichever comes first) and written with flagged = true
 * and an explaining note, so nobody banks nine hours for walking out the door.
 * Stale INs left over from earlier days get capped the same way.
 */
function autoCloseOpenSessions() {
  return withLock_(function () {
    var maxMinutes = cfgNum_('max_session_minutes', 360);
    var closeToday = localTimeOn_(new Date(), cfgStr_('lab_close', '21:00'));
    if (!closeToday) fail_('Config lab_close must look like "21:00"');
    var now = new Date();

    var students = studentIndex_(readTable_(TAB_STUDENTS));
    var byStudent = groupByStudent_(resolveEvents_(readTable_(TAB_EVENTS)));

    var rows = [];
    var closed = [];
    for (var sid in byStudent) {
      if (!byStudent.hasOwnProperty(sid)) continue;
      var events = byStudent[sid];
      var last = events[events.length - 1];
      if (!last || last.direction !== DIR_IN) continue;      // not in the lab; nothing to do

      var inMs = last.ts.getTime();
      var capMs = inMs + maxMinutes * 60000;
      var outMs;
      var reason;
      if (closeToday.getTime() > inMs) {
        outMs = Math.min(closeToday.getTime(), capMs);
        reason = outMs === capMs ? 'capped at max_session_minutes' : 'lab close';
      } else {
        // Scanned in after lab close (or the clock rolled past midnight):
        // close at the moment this job runs, still capped.
        outMs = Math.min(now.getTime(), capMs);
        reason = 'auto-close run';
      }
      if (outMs <= inMs) outMs = inMs;                        // never negative

      var outTs = new Date(outMs);
      var ev = {
        event_id: newId_('evt'),
        student_id: sid,
        timestamp: toIso_(outTs),      // synthetic, hence flagged below
        direction: DIR_OUT,
        source: SOURCE_AUTOCLOSE,
        flagged: true,
        note: 'auto-closed at ' + reason + ' — student did not scan out; ' +
              'session shown as ' + Math.max(0, minutesBetween_(last.ts, outTs)) + ' min, verify before counting'
      };
      rows.push(eventRow_(ev));
      closed.push({
        student_id: sid,
        name: students[sid] ? students[sid].name : '(not on roster)',
        in_time: last.timestamp,
        out_time: ev.timestamp,
        minutes: Math.max(0, minutesBetween_(last.ts, outTs)),
        reason: reason
      });
    }

    appendRows_(TAB_EVENTS, rows);
    Logger.log('autoCloseOpenSessions: closed ' + closed.length + ' open session(s)');
    for (var i = 0; i < closed.length; i++) {
      Logger.log('  ' + closed[i].name + '  ' + closed[i].in_time + ' -> ' + closed[i].out_time +
                 '  (' + closed[i].minutes + ' min, ' + closed[i].reason + ')');
    }
    return { closed: closed.length, sessions: closed, ran_at: nowIso_() };
  });
}

// ---------------------------------------------------------------------------
// Manual smoke test — run from the editor after setup.
// Uses the dummy ID 1234567. Never put a real student ID in this file.
// ---------------------------------------------------------------------------

function smokeTest() {
  var TEST_ID = '1234567';
  var pin = cfgStr_('admin_pin', '');
  var log = [];

  function step(label, payload) {
    var res = JSON.parse(doPost({ postData: { contents: JSON.stringify(payload) } }).getContent());
    log.push(label + ': ' + JSON.stringify(res));
    Logger.log(label + ' -> ' + JSON.stringify(res));
    return res;
  }

  Logger.log('health -> ' + doGet({}).getContent());

  var scan1 = step('scan (unknown or in)', { action: 'scan', student_id: TEST_ID });
  if (scan1.ok && scan1.data.status === 'unknown') {
    step('enroll', { action: 'enroll', student_id: TEST_ID, name: 'Test Student', grade: '11' });
    scan1 = step('scan again', { action: 'scan', student_id: TEST_ID });
  }
  step('scan immediately (expect debounced)', { action: 'scan', student_id: TEST_ID });
  step('roster', { action: 'getRoster', pin: pin });
  step('timesheet', { action: 'getTimesheet', pin: pin });
  step('student detail', { action: 'getStudent', pin: pin, student_id: TEST_ID });
  if (scan1.ok && scan1.data.event_id) {
    step('void that event', { action: 'deleteEvent', pin: pin, event_id: scan1.data.event_id,
                              reason: 'smoke test cleanup' });
  }
  step('bad pin (expect error)', { action: 'getRoster', pin: '000000' });
  step('bad id (expect error)', { action: 'scan', student_id: '12' });

  return log;
}
