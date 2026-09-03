# Robotics Lab Attendance Tracker

Lab hours tracking for a high school robotics team. Students scan a QR ID card on a
wall-mounted tablet when they enter and again when they leave. Admins get a timesheet,
per-student stats, and student-written session summaries with photos.

## Stack (decided — do not propose alternatives)

| Layer | Choice |
| --- | --- |
| Frontend | Plain static HTML/CSS/JS. No build step, no npm, no framework, no bundler. Vanilla ES modules only. |
| Hosting | GitHub Pages, serving the repo as-is. |
| Backend | One Google Apps Script Web App, acting as a JSON API. |
| Database | A Google Sheet (4 tabs). |
| Photo storage | A Google Drive folder, written to by Apps Script. |

Everything must stay on free tiers with no credit card. That rules out Firebase Storage
and Supabase. Do not introduce a dependency that requires a build step or a paid tier.

## Hard constraints

These are the things that will silently break the app if ignored.

### 1. POSTs must use `Content-Type: text/plain`

Every POST to the Apps Script Web App sends `Content-Type: text/plain` and a JSON string
body, parsed server-side with `JSON.parse(e.postData.contents)`. Any other content type
(including `application/json`) triggers a CORS preflight `OPTIONS` request, which Apps
Script does not answer — the request fails outright. Do not "fix" this by setting
`application/json`.

```js
fetch(API_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain' },
  body: JSON.stringify({ action: 'scan', student_id: '1234567' })
});
```

### 2. The server assigns every timestamp

Clients never send a timestamp. The tablet's clock is not trusted. Apps Script stamps
`timestamp` at write time.

The single exception: scans queued while offline carry a client-supplied timestamp, and
that event MUST be written with `flagged = true` and `source` identifying it as an
offline replay. Flagged events are visible as unverified in the admin UI.

### 3. Events are an append-only log

`Events` is an immutable event log. Never store a mutable `total_hours`, `currently_in`,
or `last_seen` field anywhere — not in the Sheet, not in a cache.

- Whether a student is in the lab = the `direction` of their most recent event.
- Total hours, session lists, and every statistic are recomputed from the log on read.
- Corrections are made by appending a new event with a `note`, never by editing a row.

### 4. QR only

The QR payload is a bare 7-digit number and nothing else — e.g. `1234567`. No URL, no
prefix, no JSON. Validate with `/^\d{7}$/` before hitting the API.

Restrict the barcode detector to QR:

```js
const detector = new BarcodeDetector({ formats: ['qr_code'] });
```

Do not scan for linear/1D formats. It wastes frames on a tablet that is already running
a live camera feed.

## Sheet schema

Four tabs. Column order is the contract — Apps Script reads by index, so do not reorder
or insert columns in the middle.

**`Students`**

| Column | Notes |
| --- | --- |
| `student_id` | 7-digit string. Keep as text, not a number — leading zeros matter. |
| `name` | |
| `grade` | |
| `active` | Inactive students stay in the sheet; their history is preserved. |
| `created_at` | Server-assigned. |

**`Events`** — append-only

| Column | Notes |
| --- | --- |
| `event_id` | |
| `student_id` | |
| `timestamp` | Server-assigned, except flagged offline replays. |
| `direction` | `in` or `out`. |
| `source` | Which client produced it (e.g. tablet scanner, offline queue, admin). |
| `flagged` | `true` when the timestamp is unverified or the event needs review. |
| `note` | Free text; used for corrections and admin annotations. |

**`Summaries`**

| Column | Notes |
| --- | --- |
| `summary_id` | |
| `student_id` | |
| `session_date` | |
| `text` | Student-written. |
| `photo_urls` | Drive URLs, written by Apps Script after upload. |
| `submitted_at` | Server-assigned. |

**`Config`** — `key`, `value`. Runtime settings that must not require a redeploy.

## Files

```
index.html              scanner (tablet)
admin.html              roster, timesheet, student detail
summary.html            student summary submission (phone)
css/styles.css          shared styles
js/api.js               shared API client — the only place fetch() is called
config.example.js       template; copy to config.js
config.js               GITIGNORED — holds the deployed Apps Script URL
apps-script/Code.gs     the entire backend
```

`js/api.js` is the only module that talks to the network. Pages call into it; they do not
call `fetch` themselves.

## Secrets and test data

- **Never commit a real student ID.** Use `1234567` as the dummy ID in all test code,
  examples, and documentation.
- **Never commit the deployed Apps Script URL.** It lives in `config.js`, which is
  gitignored. `config.example.js` is checked in as the template.
- This project sits inside a git repository rooted at the user's home directory. Verify
  what is staged before committing; do not `git add -A` from the repo root.
