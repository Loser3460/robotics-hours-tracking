/**
 * Reusable QR scanner for student ID cards.
 *
 * Owns the whole camera lifecycle — permission, stream, decode loop, teardown —
 * so a page only has to supply a <video> element and a callback. Used by the
 * wall tablet (index.html), which scans continuously, and by the phone summary
 * page (summary.html), which opens the camera briefly and closes it again.
 *
 *   const scanner = createScanner(video, (id) => console.log(id));
 *   await scanner.start();     // throws CameraError if the camera is unusable
 *   scanner.stop();            // ALWAYS call this — see the note on stop()
 *
 * Three decisions worth keeping:
 *
 * QR only. The detector is restricted to 'qr_code'. Asking for linear formats
 * as well makes every frame more expensive on a tablet that is already running
 * a live camera preview, and the cards are QR.
 *
 * Decode a small centred crop, not the frame. The camera is opened at 720p,
 * and on the jsQR path each attempt reads only the middle 60% of the visible
 * preview, downscaled — a ~288px square rather than a 1280x720 one. That is
 * roughly 25ms per attempt instead of 130ms, which is what turns "hold the
 * card still" into "it reads as you bring it up". The price is that a card
 * outside the crop is not seen at all, so the page MUST draw a box over the
 * region: roiScreenSize() gives its size, and index.html's reticle is it.
 *
 * The callback only ever fires for a valid payload. A QR code that is not a
 * bare 7-digit number is dropped silently, with no callback and no error —
 * students scan cereal boxes and each other's phones to see what happens, and
 * any response at all turns that into a game.
 */

// One definition of what a student ID looks like, shared with the API client.
import { isValidStudentId } from './api.js';

/**
 * jsQR 1.4.0, vendored at js/vendor/jsqr.js.
 *
 * Served from our own origin on purpose. The lab tablet is an iPad, which has
 * no BarcodeDetector, so jsQR is the ONLY way it decodes anything — fetching it
 * from a CDN would mean the scanner cannot read a card whenever the wifi is
 * down, which is exactly when the offline queue is supposed to save the day. A
 * local copy is also one less third-party script running on a page that handles
 * student data, and it is what the service worker precaches.
 *
 * Resolved against this module's own URL so it does not matter which page loads
 * it, or what subdirectory the site is served from.
 * Upstream: https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js
 * sha384-b5Ya4Bq3qCyz39m2ISh+4DxjAIljdeFwK/BsXLuj9gugaNwAcj/ia15fxNZL9Nlx
 */
const JSQR_URL = new URL('./vendor/jsqr.js', import.meta.url).href;

const DEFAULTS = {
  // ~16 looks per second. This used to be 120ms, which was moot when an
  // attempt itself took longer than that; now that an attempt is ~25ms the
  // interval is the thing setting how fast a card is picked up, and at 60ms
  // the loop is still idle three quarters of the time.
  decodeIntervalMs: 60,
  repeatLockoutMs: 6000,     // ignore the same card sitting in frame
  // Rear camera by default, which is right for a hand-held phone pointing at a
  // card. The wall tablet overrides this with 'user' — see index.html.
  facingMode: 'environment',
  jsqrWidth: 480,            // downscale the frame to this long edge before decoding
  // Fraction of the visible short edge that is actually decoded. jsQR only ever
  // sees this centred square, so the on-screen alignment box MUST match it —
  // see roiScreenSize() and the reticle in index.html.
  roi: 0.6,
  // Ask for 720p rather than whatever the camera can do. We are decoding, not
  // recording: a 1080p or 4K stream costs the same decode after downscaling but
  // makes the browser move several times as many bytes per frame.
  width: 1280,
  height: 720,
  feedback: true,            // beep + vibrate on a successful scan
};

// How many decode timings to keep for the rolling stats. One page-load's worth
// of scans is plenty to spot a regression.
const STAT_WINDOW = 120;

/** A camera that could not be opened. `kind` lets the page choose its wording. */
export class CameraError extends Error {
  constructor(message, kind) {
    super(message);
    this.name = 'CameraError';
    this.kind = kind;   // 'insecure' | 'unsupported' | 'denied' | 'notfound' | 'busy' | 'reader' | 'other'
  }
}

const CAMERA_MESSAGES = {
  NotAllowedError: ['Camera permission was denied. Allow the camera for this page in the browser settings, then try again.', 'denied'],
  NotFoundError: ['No camera was found on this device.', 'notfound'],
  NotReadableError: ['The camera is already in use by another app. Close it and try again.', 'busy'],
  OverconstrainedError: ['No camera matched what this page asked for.', 'notfound'],
};

// The jsQR <script> is fetched at most once per page, however many scanners run.
let jsqrPromise = null;

function loadJsQr() {
  if (!jsqrPromise) {
    jsqrPromise = new Promise((resolve, reject) => {
      const el = document.createElement('script');
      el.src = JSQR_URL;
      el.onload = () => (typeof window.jsQR === 'function'
        ? resolve(window.jsQR)
        : reject(new CameraError('the QR reader loaded but did not register', 'reader')));
      el.onerror = () => reject(new CameraError('could not load the QR reader', 'reader'));
      document.head.appendChild(el);
    });
    jsqrPromise.catch(() => { jsqrPromise = null; });   // let a later attempt retry
  }
  return jsqrPromise;
}

/**
 * The part of the video frame the student can actually see.
 *
 * Both pages show the preview with `object-fit: cover`, so the browser scales
 * the frame up until it fills the element and throws away the overflow on one
 * axis. Decoding a region the student cannot see is worse than useless: they
 * would be aiming at a box that does not correspond to what jsQR is reading.
 */
function visibleRect(video) {
  const vw = video.videoWidth, vh = video.videoHeight;
  const ew = video.clientWidth || vw, eh = video.clientHeight || vh;
  const scale = Math.max(ew / vw, eh / vh);        // object-fit: cover
  const w = Math.min(vw, ew / scale);
  const h = Math.min(vh, eh / scale);
  return { x: (vw - w) / 2, y: (vh - h) / 2, w, h, scale };
}

/**
 * The centred square, in video pixels, that we hand to jsQR.
 *
 * Sized off the *visible* rect rather than the raw frame so the region is
 * always fully on screen, which is what lets the reticle be honest: the box
 * the student aims at is exactly the pixels being decoded.
 */
function roiRect(video, fraction) {
  const vis = visibleRect(video);
  const side = Math.min(vis.w, vis.h) * fraction;
  return {
    x: video.videoWidth / 2 - side / 2,
    y: video.videoHeight / 2 - side / 2,
    side,
    scale: vis.scale,
  };
}

/**
 * Scan feedback: a short beep and a buzz.
 *
 * The point is to end the hover. Without it a student has no idea whether the
 * scan landed, so they keep the card in frame and inch it around, which is
 * both the thing they complain about and the thing that keeps the decode loop
 * busy. The confirmation sheet appears a network round-trip later; this fires
 * the instant the code is read.
 */
let audioCtx = null;
let audioUnlocked = false;

/**
 * iOS starts an AudioContext suspended and will only resume it inside a user
 * gesture. The tablet boots the scanner with no gesture at all, so arm it on
 * the first touch the kiosk ever gets and keep it for the rest of the day.
 */
function unlockAudioOnFirstGesture() {
  if (audioUnlocked) return;
  audioUnlocked = true;
  const unlock = () => {
    try { audioCtx?.resume(); } catch { /* nothing to resume */ }
  };
  for (const evt of ['pointerdown', 'touchend', 'keydown']) {
    window.addEventListener(evt, unlock, { once: true, passive: true });
  }
}

function beep() {
  try {
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if (!Ctx) return;
    if (!audioCtx) audioCtx = new Ctx();
    if (audioCtx.state === 'suspended') audioCtx.resume();
    const now = audioCtx.currentTime;
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    osc.type = 'sine';
    osc.frequency.setValueAtTime(1080, now);
    // Ramped, not switched: a square-edged gate on a sine clicks on the iPad.
    gain.gain.setValueAtTime(0.0001, now);
    gain.gain.exponentialRampToValueAtTime(0.28, now + 0.012);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.14);
    osc.connect(gain).connect(audioCtx.destination);
    osc.start(now);
    osc.stop(now + 0.15);
  } catch (err) {
    console.warn('scanner: could not beep', err);   // never block a scan on audio
  }
}

function buzz() {
  // Android only — iOS Safari has no Vibration API, which is why the beep
  // carries the feedback on the tablet and this is a bonus on phones.
  try { navigator.vibrate?.(60); } catch { /* not supported */ }
}

/**
 * Native BarcodeDetector where it exists, jsQR everywhere else.
 * iPad Safari has no BarcodeDetector at all, which is the case that matters
 * most here — the lab tablet is an iPad.
 */
async function buildDetector(video, opts) {
  if ('BarcodeDetector' in window) {
    try {
      const formats = await window.BarcodeDetector.getSupportedFormats();
      if (formats.includes('qr_code')) {
        const detector = new window.BarcodeDetector({ formats: ['qr_code'] });
        console.info('scanner: using native BarcodeDetector');
        // Handed the <video> directly: the native path is hardware-backed, so
        // cropping and downscaling in JS first would cost more than it saves.
        // It therefore reads the whole frame while the reticle marks only the
        // middle — a card inside the box still works, it is just not the only
        // place that works. That asymmetry is fine; the box is guidance.
        return async () => {
          const codes = await detector.detect(video);
          return codes.length ? codes[0].rawValue : null;
        };
      }
    } catch (err) {
      console.warn('scanner: BarcodeDetector unusable, falling back to jsQR', err);
    }
  }

  const jsQR = await loadJsQr();
  console.info('scanner: using jsQR fallback');

  // One reusable canvas, sized once. jsQR is pure JavaScript and its cost is
  // linear in pixels, so what we hand it is the whole optimisation:
  //
  //   full 1280x720 frame          ~810 ms per attempt
  //   downscaled to 480 long edge  ~130 ms
  //   + cropped to the 60% reticle  ~25 ms
  //
  // At 25 ms a decode fits inside one 60 Hz frame's budget several times over,
  // which is the difference between "hold it still for a few seconds" and
  // "it reads as you bring the card up".
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true, alpha: false });

  return () => {
    if (!video.videoWidth) return null;
    const roi = roiRect(video, opts.roi);
    if (roi.side < 1) return null;

    // The decode buffer: the reticle square, at the resolution it would have
    // had in a frame downscaled to jsqrWidth. Never upscale — a blurry card
    // does not get sharper, it just costs more pixels.
    const side = Math.max(1, Math.round(Math.min(roi.side, opts.jsqrWidth * opts.roi)));
    if (canvas.width !== side) { canvas.width = side; canvas.height = side; }

    // NOTE: drawImage reads the raw frame. The tablet's preview is mirrored,
    // but that mirroring is a CSS transform on the <video> element and CSS
    // never touches the pixels the canvas gets. Do not "match" it here — a
    // mirrored QR code does not decode at all.
    ctx.drawImage(video, roi.x, roi.y, roi.side, roi.side, 0, 0, side, side);
    const image = ctx.getImageData(0, 0, side, side);
    const code = jsQR(image.data, side, side, { inversionAttempts: 'dontInvert' });
    return code ? code.data : null;
  };
}

/**
 * Create a scanner bound to a <video> element.
 *
 * @param {HTMLVideoElement} video   where the preview is shown
 * @param {(id: string) => void} onCode  called with a validated 7-digit ID
 * @param {object} [options]  decodeIntervalMs, repeatLockoutMs, facingMode, jsqrWidth
 */
export function createScanner(video, onCode, options = {}) {
  if (!video) throw new Error('createScanner needs a video element');
  if (typeof onCode !== 'function') throw new Error('createScanner needs a callback');
  const opts = { ...DEFAULTS, ...options };

  let stream = null;
  let detect = null;
  let frame = 0;
  let running = false;
  let paused = false;
  let decoding = false;
  let lastDecodeAt = 0;
  let lastCode = null;
  let lastCodeAt = 0;

  // Rolling decode timings, so a regression here is measurable rather than a
  // matter of opinion. See the stats getter.
  let attempts = 0;
  let attemptsSinceHit = 0;
  let hits = 0;
  const timings = [];

  function record(ms) {
    attempts += 1;
    attemptsSinceHit += 1;
    timings.push(ms);
    if (timings.length > STAT_WINDOW) timings.shift();
  }

  function percentile(sorted, q) {
    if (!sorted.length) return 0;
    return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * q))];
  }

  /**
   * One decode attempt per animation frame at most, and never two at once.
   *
   * `decoding` is the important guard. The native detector is async and jsQR
   * can overrun the interval on a slow tablet; without it, attempts pile up
   * behind each other and every one of them is reading a stale frame by the
   * time it runs, so the loop gets slower exactly when it is already behind.
   */
  function tick(timestamp) {
    if (!running) return;
    frame = requestAnimationFrame(tick);
    if (paused || decoding || !detect) return;
    if (timestamp - lastDecodeAt < opts.decodeIntervalMs) return;
    lastDecodeAt = timestamp;

    decoding = true;
    const started = performance.now();
    Promise.resolve()
      .then(detect)
      .then((raw) => {
        record(performance.now() - started);
        if (raw != null) handle(raw);
      })
      .catch((err) => {
        record(performance.now() - started);
        console.warn('scanner: decode failed', err);
      })
      .finally(() => { decoding = false; });
  }

  function handle(raw) {
    const value = String(raw).trim();
    if (!isValidStudentId(value)) return;             // silently ignored, by design

    const now = Date.now();
    if (value === lastCode && now - lastCodeAt < opts.repeatLockoutMs) return;
    lastCode = value;
    lastCodeAt = now;
    hits += 1;

    // Before the callback, so the confirmation is instant even though the
    // scan itself still has a network round-trip ahead of it.
    if (opts.feedback) { beep(); buzz(); }

    const s = stats();
    console.info(`scanner: read a card after ${attemptsSinceHit} attempt(s) — ` +
      `median ${s.medianMs.toFixed(1)}ms, p95 ${s.p95Ms.toFixed(1)}ms per attempt ` +
      `over the last ${s.samples}`);
    attemptsSinceHit = 0;

    try {
      onCode(value);
    } catch (err) {
      console.error('scanner: callback threw', err);  // never let it kill the loop
    }
  }

  /** Rolling decode-cost summary. Cheap to call; safe before start(). */
  function stats() {
    const sorted = [...timings].sort((a, b) => a - b);
    return {
      attempts,
      attemptsSinceHit,
      hits,
      samples: sorted.length,
      lastMs: timings.length ? timings[timings.length - 1] : 0,
      medianMs: percentile(sorted, 0.5),
      p95Ms: percentile(sorted, 0.95),
    };
  }

  async function start() {
    if (running) return;

    if (!window.isSecureContext) {
      throw new CameraError(
        'The camera needs a secure connection. Open this page over https:// — cameras are blocked on http:// and file://.',
        'insecure');
    }
    if (!navigator.mediaDevices?.getUserMedia) {
      throw new CameraError('This browser cannot open a camera.', 'unsupported');
    }

    try {
      // Ask for 720p explicitly rather than letting the device pick. Left to
      // itself an iPad hands back the sensor's full resolution, and every one
      // of those extra pixels is copied per frame for a decode that throws all
      // of them away. `ideal` rather than `exact` so a camera that cannot do
      // 720p still opens instead of throwing OverconstrainedError.
      stream = await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: opts.facingMode },
          width: { ideal: opts.width },
          height: { ideal: opts.height },
          frameRate: { ideal: 30 },
        },
        audio: false,
      });
    } catch (err) {
      const [message, kind] = CAMERA_MESSAGES[err.name] || [err.message, 'other'];
      throw new CameraError(message, kind);
    }

    try {
      video.srcObject = stream;
      video.setAttribute('playsinline', '');   // iOS would otherwise go fullscreen
      video.muted = true;
      await video.play();
      detect = await buildDetector(video, opts);
    } catch (err) {
      stop();                                   // never leave the light on after a failed start
      throw err instanceof CameraError ? err : new CameraError(err.message, 'other');
    }

    if (opts.feedback) unlockAudioOnFirstGesture();

    running = true;
    paused = false;
    lastCode = null;
    frame = requestAnimationFrame(tick);
  }

  /**
   * Release the camera. Safe to call twice, and safe to call after a failed
   * start. A phone camera left running drains the battery fast and leaves the
   * recording indicator on, so every caller must reach this on every exit path
   * — success, cancel, navigating away.
   */
  function stop() {
    running = false;
    paused = false;
    if (frame) { cancelAnimationFrame(frame); frame = 0; }
    if (stream) {
      for (const track of stream.getTracks()) track.stop();
      stream = null;
    }
    if (video) {
      try { video.pause(); } catch { /* not playing */ }
      video.srcObject = null;
    }
    detect = null;
    decoding = false;
  }

  return {
    start,
    stop,
    /** Keep the stream open but stop decoding — for a confirmation overlay. */
    pause() { paused = true; },
    resume() { paused = false; lastCode = null; },
    /** Let the same card be scanned again immediately. */
    clearLockout() { lastCode = null; lastCodeAt = 0; },
    get running() { return running; },
    get paused() { return paused; },
    /**
     * Side of the decoded region in CSS pixels, for drawing an alignment box
     * that matches it. Under `object-fit: cover` the visible short edge always
     * maps to the element's short edge, so this reduces to a fraction of the
     * element — but ask rather than hard-coding it, so the box follows the ROI
     * if it is ever retuned.
     */
    roiScreenSize() {
      return opts.roi * Math.min(video.clientWidth, video.clientHeight);
    },
    /** Rolling decode timings — {attempts, hits, samples, lastMs, medianMs, p95Ms}. */
    get stats() { return stats(); },
  };
}
