// Deployed Apps Script Web App URL for the lab attendance tracker.
//
// This file IS committed and the URL IS public, by decision. A static site on
// GitHub Pages cannot keep a secret: anything the browser needs to make the
// request, a visitor can read out of the network tab. Gitignoring this file
// would have hidden the URL from repo scrapers only, not from anyone using the
// page — so it is checked in and the exposure is handled where it is real.
//
// What that means: `scan` and `enroll` are unauthenticated, so anyone with this
// URL can append events. Admin actions stay PIN-gated, and the PIN lives only in
// the Config tab of the Sheet — never in this repo. Junk rows are expected
// maintenance, not a breach: Events is append-only, so a coach voids them from
// the admin page and the originals stay visible for audit.
//
// Redeploy: Apps Script editor -> Deploy -> Manage deployments -> edit -> New version
//   Execute as: Me   |   Who has access: Anyone
// A redeploy that creates a NEW deployment changes this URL; paste the new one here.

export const API_URL = 'https://script.google.com/macros/s/AKfycby2CheYCuJP_XqDDeeD394K9CsuALmitCck5-mwmdE4NxLtbCW-gKNWUSnuPffa8lg/exec';
