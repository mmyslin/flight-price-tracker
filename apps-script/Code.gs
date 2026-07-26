/**
 * Flight Tracker — Gmail → Google Sheet sync + JSON endpoint
 *
 * Tracks FUTURE flights only: what each upcoming flight actually cost
 * (cash + previous credits + miles/points), so price drops are easy to act on.
 *
 * Parsers are modeled on real United and Alaska emails from this inbox
 * (July 2026 formats): United "eTicket Itinerary and Receipt", United
 * cancellation notices, United future-flight-credit notices, and Alaska
 * "Your flight is booked" (Atmos era, incl. re-ticketed changes).
 *
 * syncFlights() only Gmail-searches the last SYNC_LOOKBACK_DAYS on a Sheet
 * that already has data (cancellations/credits/rebookings correlate against
 * the Sheet, not the search window) — keeps it well under the daily Gmail
 * quota. The very first run against an empty Sheet backfills a full year.
 *
 * Setup (one time):
 *   1. Create a Google Sheet, then Extensions → Apps Script, paste this file.
 *   2. Run setupSheets() once (grants Sheets scope, creates tabs).
 *   3. Run syncFlights() once (grants Gmail scope) to pull current bookings.
 *   4. Deploy → New deployment → Web app → Execute as: Me,
 *      Who has access: Anyone with the link. Copy the /exec URL into
 *      DATA_URL at the top of dashboard/index.html.
 *
 * Future automation (daily sync), uncomment and run once:
 *   // ScriptApp.newTrigger('syncFlights').timeBased().everyDays(1).atHour(6).create();
 */

const FLIGHTS_SHEET = 'Flights';
const SAVINGS_SHEET = 'Savings';

const FLIGHT_COLS = [
  'airline', 'confirmation', 'date', 'departTime', 'origin', 'destination',
  'flightNumbers', 'cashPaid', 'creditsApplied', 'milesPaid', 'awardFees',
  'status', 'notes', 'sourceEmail', 'lastSynced', 'currentPrice', 'currentMiles',
  'previousPrice', 'previousMiles', 'alertedPrice', 'priceCheckDate'
];
const SAVINGS_COLS = ['date', 'route', 'note', 'dollarsSaved', 'milesSaved'];
const ALERT_USD_THRESHOLD = 25;
const SYNC_LOOKBACK_DAYS = 3;

/* ============================ SETUP ============================ */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, FLIGHTS_SHEET, FLIGHT_COLS);
  ensureSheet_(ss, SAVINGS_SHEET, SAVINGS_COLS);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================ SYNC ============================= */

function syncFlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();
  const sh = ss.getSheetByName(FLIGHTS_SHEET);

  // Correlation (rebooking carry-forward, cancellations, credits) checks
  // this map first, so a narrow daily window still finds the *original*
  // booking even when its email has aged out of the search — only a fresh
  // Sheet (first-ever run) needs the expensive full-year backfill.
  const existing = readExistingFlights_(sh);
  const win = existing.size ? SYNC_LOOKBACK_DAYS + 'd' : '1y';

  // 1. Gather booking state per confirmation, processing messages oldest →
  //    newest so re-ticketed changes (same conf, new flight/date) win.
  const bookings = {};   // conf → flight object
  collectMessages_('from:united.com subject:"eTicket Itinerary and Receipt" newer_than:' + win)
    .forEach(m => { const f = parseUnitedReceipt_(m.body); if (f) remember_(bookings, existing, f, m); });
  collectMessages_('from:alaskaair.com subject:"Your flight is booked" newer_than:' + win)
    .forEach(m => { const f = parseAlaskaBooked_(stripHtml_(m.html)); if (f) remember_(bookings, existing, f, m); });

  // 2. Cancellations flip status — in the in-run booking if present, else
  //    directly on the already-synced Sheet row.
  collectMessages_('from:united.com (subject:"cancellation is complete" OR subject:"reservation has been canceled") newer_than:' + win)
    .forEach(m => {
      const conf = matchOne_(m.body, /Confirmation number:\s*([A-Z0-9]{6})/i) ||
                   matchOne_(m.subject, /\(([A-Z0-9]{6})\)/);
      markCanceled_(conf, bookings, sh, existing);
    });
  collectMessages_('from:alaskaair.com subject:(canceled OR cancelled) newer_than:' + win)
    .forEach(m => {
      const conf = matchOne_(stripHtml_(m.html), /Confirmation code:\s*\n?\s*([A-Z]{6})/i);
      markCanceled_(conf, bookings, sh, existing);
    });

  // 3. FUTURE FLIGHTS ONLY: drop past and canceled bookings.
  const todayIso = isoDate_(new Date());
  const future = Object.values(bookings).filter(f =>
    f.date && f.date >= todayIso && f.status !== 'canceled');

  future.forEach(f => upsertFlight_(ss, f));

  // 4. Date-change savings: United sends "future flight credit" for two
  //    unrelated reasons that need opposite handling, distinguished by the
  //    email's own template rather than same-run timing (a genuine
  //    rebooking's receipt can legitimately fall just outside the window on
  //    a given day while the credit notice — sent seconds apart — is still
  //    in it; requiring both in the SAME run flagged real rebookings as
  //    cancellations). A straight cancellation-to-credit itemizes "Ticket
  //    Value" / "Change fee" under "Transaction summary"; a rebooking
  //    fare-difference credit is just "Future Flight Credit Details" with a
  //    bare dollar amount — no itemization, because it's not the whole fare.
  const savings = ss.getSheetByName(SAVINGS_SHEET);
  const existingNotes = savings.getLastRow() > 1
    ? savings.getRange(2, 3, savings.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  collectMessages_('from:united.com subject:"future flight credit" newer_than:' + win)
    .forEach(m => {
      const conf = matchOne_(m.body, /Confirmation Number:?,?\s*([A-Z0-9]{6})/i);
      // Two templates seen: rebooking credit reads "$28.02"; a straight
      // cancellation-to-credit reads "Ticket Value USD93.4" (no $, 1 decimal).
      const amt = matchNum_(m.body, /(?:\$|USD)\s?([\d,]+\.\d{1,2})/);
      if (!conf || !amt) return;
      if (/Transaction summary/i.test(m.body)) {
        markCanceled_(conf, bookings, sh, existing);
        return;
      }
      const rebooked = bookings[conf] || existing.get(conf);
      if (!rebooked || rebooked.status === 'canceled') return;   // no active booking on record for this credit
      const note = 'Fare difference returned as credit on ' + conf;
      if (existingNotes.some(n => n.includes(conf))) return;
      savings.appendRow([isoDate_(m.date), rebooked.origin + '-' + rebooked.destination + ' (' + conf + ')', note, amt, 0]);
    });
}

/** Reads already-synced Flights rows into a Map keyed by confirmation, so
 * correlation (rebooking/cancellation/credit) doesn't depend on the current
 * Gmail search window covering the original booking email. */
function readExistingFlights_(sh) {
  const map = new Map();
  if (sh.getLastRow() < 2) return map;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, FLIGHT_COLS.length).getValues();
  rows.forEach((row, i) => {
    const f = { rowNum: i + 2 };
    FLIGHT_COLS.forEach((c, j) => {
      let v = row[j];
      if (v instanceof Date) v = (c === 'departTime') ? formatTime_(v) : isoDate_(v);
      f[c] = v;
    });
    if (f.confirmation) map.set(String(f.confirmation), f);
  });
  return map;
}

function markCanceled_(conf, bookings, sh, existing) {
  if (!conf) return;
  if (bookings[conf]) { bookings[conf].status = 'canceled'; return; }
  const row = existing.get(conf);
  if (row && row.status !== 'canceled') {
    sh.getRange(row.rowNum, FLIGHT_COLS.indexOf('status') + 1).setValue('canceled');
    row.status = 'canceled';
  }
}

/**
 * Emails a summary when an active flight's currentPrice (set daily by
 * scripts/check-prices.mjs) has dropped $25+ below what was actually paid
 * (cashPaid + creditsApplied). alertedPrice tracks the last price already
 * emailed for that row, so a still-cheap fare doesn't re-alert every day —
 * only a NEW, lower price (or the first time crossing the threshold) does.
 * Run this on its own daily trigger, timed after the price-check workflow.
 */
function checkPriceAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FLIGHTS_SHEET);
  if (sh.getLastRow() < 2) return;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, FLIGHT_COLS.length).getValues();
  const idx = name => FLIGHT_COLS.indexOf(name);
  const alerts = [];

  rows.forEach((row, i) => {
    const status = row[idx('status')];
    const currentPrice = row[idx('currentPrice')];
    if (status !== 'active' || currentPrice === '' || currentPrice == null) return;

    const totalPaid = Number(row[idx('cashPaid')] || 0) + Number(row[idx('creditsApplied')] || 0);
    const delta = Number(currentPrice) - totalPaid;
    const alertedPrice = row[idx('alertedPrice')];
    if (delta > -ALERT_USD_THRESHOLD || Number(currentPrice) === Number(alertedPrice)) return;

    const dateVal = row[idx('date')];
    alerts.push({
      rowNum: i + 2,
      confirmation: row[idx('confirmation')],
      origin: row[idx('origin')],
      destination: row[idx('destination')],
      date: dateVal instanceof Date ? isoDate_(dateVal) : dateVal,
      totalPaid: totalPaid,
      currentPrice: Number(currentPrice),
    });
  });
  if (!alerts.length) return;

  const lines = alerts.map(a =>
    a.origin + ' → ' + a.destination + ' on ' + a.date + ' (' + a.confirmation + '): ' +
    'paid $' + a.totalPaid.toFixed(2) + ', now $' + a.currentPrice.toFixed(2) +
    ' — save $' + (a.totalPaid - a.currentPrice).toFixed(2)
  );
  const subject = alerts.length === 1
    ? 'Flight price drop: ' + alerts[0].origin + '→' + alerts[0].destination + ' is $' +
      (alerts[0].totalPaid - alerts[0].currentPrice).toFixed(2) + ' cheaper'
    : 'Flight price drop: ' + alerts.length + ' flights are $' + ALERT_USD_THRESHOLD + '+ cheaper';
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, lines.join('; '));

  alerts.forEach(a => sh.getRange(a.rowNum, idx('alertedPrice') + 1).setValue(a.currentPrice));
}

function collectMessages_(query) {
  const out = [];
  GmailApp.search(query, 0, 100).forEach(thread =>
    thread.getMessages().forEach(msg => out.push({
      body: msg.getPlainBody(),
      html: msg.getBody(),
      subject: msg.getSubject(),
      date: msg.getDate(),
      id: msg.getId(),
    })));
  out.sort((a, b) => a.date - b.date);   // oldest first → newest state wins
  return out;
}

function remember_(bookings, existing, flight, msg) {
  flight.sourceEmail = 'https://mail.google.com/mail/u/0/#all/' + msg.id;
  flight.lastSynced = isoDate_(new Date());
  const prev = bookings[flight.confirmation] || existing.get(flight.confirmation);
  if (prev) {
    const rebookedItinerary = prev.date !== flight.date || prev.flightNumbers !== flight.flightNumbers;
    // A reissue can correct the payment breakdown (e.g. a different
    // cash/credit split for the same total) without touching date or
    // flight number. sourceEmail on the Sheet row is the message that
    // produced its current values — a different message id here means
    // this is new information, not a hand-edit to protect.
    const newEmail = !prev.sourceEmail || !prev.sourceEmail.endsWith(msg.id);
    if (rebookedItinerary || newEmail) {
      flight.rebooked = true;   // a newer email supersedes the Sheet — not a hand-edit conflict
    }
    if (rebookedItinerary) {
      flight.notes = ((flight.notes || '') + ' [rebooked from ' + prev.flightNumbers + ' on ' + prev.date + ']').trim();
      // Carry cost forward when a change email has no payment breakdown.
      if (!flight.cashPaid && !flight.creditsApplied && !flight.milesPaid) {
        flight.cashPaid = prev.cashPaid; flight.creditsApplied = prev.creditsApplied;
        flight.milesPaid = prev.milesPaid; flight.awardFees = prev.awardFees;
      }
    }
  }
  bookings[flight.confirmation] = flight;
}

/**
 * Upsert keyed by confirmation. Parser-owned fields are updated; a cell the
 * user hand-edited to something different gets flagged in notes instead of
 * silently clobbered (notes always refreshes; sourceEmail/lastSynced only
 * refresh once nothing conflicted — see the note above that block). A
 * rebooking or a same-itinerary reissue from a new message (flight.rebooked,
 * set in remember_) always overwrites — a newer email superseding stale
 * Sheet data isn't a hand-edit conflict.
 */
function upsertFlight_(ss, flight) {
  const sh = ss.getSheetByName(FLIGHTS_SHEET);
  const confCol = FLIGHT_COLS.indexOf('confirmation');
  const rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, FLIGHT_COLS.length).getValues() : [];
  const idx = rows.findIndex(r => String(r[confCol]) === flight.confirmation);

  if (idx === -1) {
    sh.appendRow(FLIGHT_COLS.map(c => flight[c] !== undefined ? flight[c] : ''));
    return;
  }
  const rowNum = idx + 2;
  const conflicts = [];
  // sourceEmail/lastSynced are applied last, only if nothing conflicted —
  // advancing them unconditionally would let a message "claim" the row
  // without its values ever landing, so a later run reprocessing that same
  // (still in-window) message sees sourceEmail already matching and never
  // retries the correction, permanently masking it.
  FLIGHT_COLS.forEach((col, i) => {
    if (col === 'sourceEmail' || col === 'lastSynced') return;
    const parsed = flight[col];
    if (parsed === undefined || parsed === '') return;
    const cur = rows[idx][i];
    const curStr = cur instanceof Date ? (col === 'departTime' ? formatTime_(cur) : isoDate_(cur)) : String(cur);
    if (flight.rebooked || col === 'notes') {
      sh.getRange(rowNum, i + 1).setValue(parsed);
    } else if (curStr === '' || curStr === String(parsed)) {
      sh.getRange(rowNum, i + 1).setValue(parsed);
    } else {
      conflicts.push(col + ': parsed "' + parsed + '" vs sheet "' + curStr + '"');
    }
  });
  if (conflicts.length) {
    const noteCell = sh.getRange(rowNum, FLIGHT_COLS.indexOf('notes') + 1);
    const existingNote = noteCell.getValue() || '';
    const conflictText = ' [sync conflict — ' + conflicts.join('; ') + ']';
    if (!existingNote.includes(conflictText.trim())) {
      noteCell.setValue((existingNote + conflictText).trim());
    }
  } else {
    sh.getRange(rowNum, FLIGHT_COLS.indexOf('sourceEmail') + 1).setValue(flight.sourceEmail);
    sh.getRange(rowNum, FLIGHT_COLS.indexOf('lastSynced') + 1).setValue(flight.lastSynced);
  }
}

/* =========================== PARSERS =========================== */

/**
 * United "eTicket Itinerary and Receipt". Single- or multi-segment; payment
 * variants seen in the wild:
 *  - award: "Special member price: 12,600 miles" / "Total: 12,600 miles + 5.60 USD"
 *  - credit: "Future flight credit: 93.40 USD" + "Confirmation #: XXXXXX"
 *            + "Future flight credit applied: -93.40 USD" + "Total: 0.00 USD"
 *  - mixed:  credit + "An additional amount of 3.97 USD ... charged to Visa"
 *  - reissue: "Previous Ticket Balance" + additional collection
 *  - plain cash: "Total: 285.02 USD" on a card
 *
 * HTML-only emails (no text/plain part) get getPlainBody()'s auto-converted
 * text, which wraps <b> runs in literal asterisks — e.g. "applied: *-93.40
 * USD*" — so amount regexes tolerate an optional `*` before the value.
 */
function parseUnitedReceipt_(body) {
  const conf = matchOne_(body, /Confirmation Number:\s*\n?\s*([A-Z0-9]{6})/);
  if (!conf) return null;

  const seg = body.match(
    /Flight \d+ of \d+ (UA ?\d{2,4})[\s\S]{0,80}?\n\s*\w{3}, (\w{3}) (\d{2}), (\d{4})[\s\S]{0,80}?(\d{2}:\d{2} [AP]M)\s+(\d{2}:\d{2} [AP]M)\s*\n\s*[^(\n]*\(([A-Z]{3})\)\s*[^(\n]*\(([A-Z]{3})\)/);
  if (!seg) return null;

  const f = {
    airline: 'United', confirmation: conf,
    flightNumbers: seg[1].replace(/UA ?/, 'UA '),
    date: toIso_(seg[4], seg[2], seg[3]),
    departTime: to24h_(seg[5]),
    origin: seg[7], destination: seg[8],
    cashPaid: 0, creditsApplied: 0, milesPaid: 0, awardFees: 0,
    status: 'active', notes: '',
  };

  const miles = matchNum_(body, /Special member price:\s*\*?([\d,]+) miles/) ||
                matchNum_(body, /Total:\s*\*?([\d,]+) miles/);
  const creditApplied = matchNum_(body, /Future flight credit applied:\s*\*?-([\d,]+\.\d{2})/);
  const creditSource = matchOne_(body, /Future flight credit:[\s\S]{0,60}?Confirmation #:\s*([A-Z0-9]{6})/);
  const additional = matchNum_(body, /An additional amount of \*?([\d,]+\.\d{2}) USD/);
  const totalUsd = matchNum_(body, /Total:\s*\*?([\d,]+\.\d{2}) USD/);
  const awardFee = matchNum_(body, /Total:\s*\*?[\d,]+ miles \+ \*?([\d,]+\.\d{2}) USD/);

  if (miles) {
    f.milesPaid = miles;
    f.cashPaid = awardFee;
    f.awardFees = awardFee;
    f.notes = 'Award ticket';
  } else if (/Previous Ticket Balance/.test(body)) {
    f.creditsApplied = round2_(totalUsd - additional);
    f.cashPaid = additional;
    f.notes = 'Paid with previous ticket balance' + (additional ? ' + $' + additional + ' card' : '');
  } else if (creditApplied) {
    f.creditsApplied = creditApplied;
    f.cashPaid = totalUsd;   // "Total: X USD" is the remainder charged to the card (0 when credit covered it all)
    f.notes = 'Paid with future flight credit' + (creditSource ? ' from ' + creditSource : '');
  } else {
    f.cashPaid = totalUsd;
  }
  return f;
}

/**
 * Alaska "Your flight is booked" (Atmos era). Change emails share the conf
 * and carry "(Previous Ticket ...)". Year is absent — inferred as the next
 * occurrence of that month/day.
 *  - "Confirmation code: XXXXXX"
 *  - "Flight 1 · Sat Aug 29" / "AS 3100"
 *  - "Departure date: Aug 29 at 5:22 PM"
 *  - "5000 points have been redeemed"
 *  - "$5.60 to be charged to the VISA card"
 *
 * Takes stripHtml_(msg.getBody()), not getPlainBody() — Alaska's actual
 * text/plain part is an ESP-generated mess (raw tracking URLs spliced
 * between labels and values, "&rarr;" left undecoded), so it's parsed from
 * the HTML instead, matching what the inbox actually shows.
 */
function parseAlaskaBooked_(body) {
  const conf = matchOne_(body, /Confirmation code:\s*\n?\s*([A-Z0-9]{6})/);
  if (!conf) return null;

  const dep = body.match(/Departure date:\s*(\w{3}) (\d{1,2}) at (\d{1,2}:\d{2} [AP]M)/);
  const fltNum = matchOne_(body, /\b(AS ?\d{2,4})\b/);
  const routeM = body.match(/\n([A-Z]{3})\s*\n+\s*→\s*\n+\s*([A-Z]{3})\s*\n/);
  if (!dep || !fltNum) return null;

  const f = {
    airline: 'Alaska', confirmation: conf,
    flightNumbers: fltNum.replace(/AS ?/, 'AS '),
    date: nextOccurrenceIso_(dep[1], dep[2]),
    departTime: to24h_(dep[3]),
    origin: routeM ? routeM[1] : '', destination: routeM ? routeM[2] : '',
    cashPaid: 0, creditsApplied: 0, milesPaid: 0, awardFees: 0,
    status: 'active', notes: '',
  };

  f.milesPaid = matchNum_(body, /([\d,]+) points have been redeemed/);
  const fee = matchNum_(body, /\$([\d,]+\.\d{2}) to be charged to the VISA/i) ||
              matchNum_(body, /New Ticket Value\s*\n?\s*\$([\d,]+\.\d{2})/);
  f.cashPaid = fee;
  if (f.milesPaid) { f.awardFees = fee; f.notes = 'Atmos points award'; }
  if (/\(Previous Ticket/.test(body)) f.notes = (f.notes + ' (re-ticketed change)').trim();
  return f;
}

/* ========================= JSON ENDPOINT ======================= */

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = {
    generatedAt: isoDate_(new Date()),
    flights: sheetToObjects_(ss.getSheetByName(FLIGHTS_SHEET), FLIGHT_COLS),
    savings: sheetToObjects_(ss.getSheetByName(SAVINGS_SHEET), SAVINGS_COLS),
  };
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects_(sh, cols) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues().map(row => {
    const o = {};
    cols.forEach((c, i) => {
      let v = row[i];
      // Sheets auto-converts "HH:MM"-looking strings to Time-of-day values
      // (Date objects on the 1899-12-30 epoch) — format those as HH:MM, not
      // as a date, or the epoch date leaks into the JSON.
      if (v instanceof Date) v = (c === 'departTime') ? formatTime_(v) : isoDate_(v);
      o[c] = v;
    });
    return o;
  });
}

/* =========================== HELPERS =========================== */

function matchOne_(s, re) { const m = s.match(re); return m ? m[1] : null; }
function matchNum_(s, re) { const m = s.match(re); return m ? Number(m[1].replace(/,/g, '')) : 0; }
/** Lightweight HTML→text: block tags become newlines, other tags drop, common entities decode. */
function stripHtml_(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<\/?(br|p|div|tr|td|th|li|h[1-6])\b[^>]*>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&rarr;/gi, '→')
    .replace(/&zwnj;/gi, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
function round2_(n) { return Math.round(n * 100) / 100; }
function isoDate_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function formatTime_(d) {
  return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}
const MONTHS_ = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
function toIso_(year, monthName, day) {
  const mi = MONTHS_.indexOf(monthName.slice(0, 3).toLowerCase());
  return year + '-' + String(mi + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
}
function nextOccurrenceIso_(monthName, day) {
  const mi = MONTHS_.indexOf(monthName.slice(0, 3).toLowerCase());
  const now = new Date();
  let candidate = new Date(now.getFullYear(), mi, Number(day));
  if (candidate < now && (now - candidate) > 30 * 24 * 3600 * 1000) {
    candidate = new Date(now.getFullYear() + 1, mi, Number(day));
  }
  return isoDate_(candidate);
}
function to24h_(t) {
  const m = t.replace(/\./g, '').match(/(\d{1,2}):(\d{2})\s?([AP])M/i);
  if (!m) return '';
  let h = Number(m[1]);
  if (m[3].toUpperCase() === 'P' && h !== 12) h += 12;
  if (m[3].toUpperCase() === 'A' && h === 12) h = 0;
  return String(h).padStart(2, '0') + ':' + m[2];
}
