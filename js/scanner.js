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
 * Two decisions worth keeping:
 *
 * QR only. The detector is restricted to 'qr_code'. Asking for linear formats
 * as well makes every frame more expensive on a tablet that is already running
 * a live camera preview, and the cards are QR.
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
  decodeIntervalMs: 120,     // ~8 looks per second; faster only burns battery
  repeatLockoutMs: 6000,     // ignore the same card sitting in frame
  facingMode: 'environment', // rear camera: the tablet is wall-mounted, the phone is held
  jsqrWidth: 480,            // downscale before software decoding
};

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
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  return () => {
    if (!video.videoWidth) return null;
    // jsQR is pure JavaScript; decoding a full 1280x720 frame every tick heats up
    // an older phone or iPad for no gain on a card held at the reticle.
    const scale = Math.min(1, opts.jsqrWidth / video.videoWidth);
    canvas.width = Math.round(video.videoWidth * scale);
    canvas.height = Math.round(video.videoHeight * scale);
    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
    const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const code = jsQR(image.data, image.width, image.height, { inversionAttempts: 'dontInvert' });
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

  function tick(timestamp) {
    if (!running) return;
    frame = requestAnimationFrame(tick);
    if (paused || decoding || !detect) return;
    if (timestamp - lastDecodeAt < opts.decodeIntervalMs) return;
    lastDecodeAt = timestamp;

    decoding = true;
    Promise.resolve()
      .then(detect)
      .then((raw) => { if (raw != null) handle(raw); })
      .catch((err) => console.warn('scanner: decode failed', err))
      .finally(() => { decoding = false; });
  }

  function handle(raw) {
    const value = String(raw).trim();
    if (!isValidStudentId(value)) return;             // silently ignored, by design

    const now = Date.now();
    if (value === lastCode && now - lastCodeAt < opts.repeatLockoutMs) return;
    lastCode = value;
    lastCodeAt = now;

    try {
      onCode(value);
    } catch (err) {
      console.error('scanner: callback threw', err);  // never let it kill the loop
    }
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
      stream = await navigator.mediaDevices.getUserMedia({
        video: {
          facingMode: { ideal: opts.facingMode },
          width: { ideal: 1280 },
          height: { ideal: 720 },
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
  };
}
