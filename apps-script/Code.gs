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
 * Setup (one time):
 *   1. Create a Google Sheet, then Extensions → Apps Script, paste this file.
 *   2. Run setupSheets() once (grants Sheets scope, creates tabs).
 *      Optionally paste data/flights-seed.csv + savings-seed.csv into the tabs.
 *   3. Run syncFlights() once (grants Gmail scope).
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
  'status', 'notes', 'sourceEmail', 'lastSynced'
];
const SAVINGS_COLS = ['date', 'route', 'note', 'dollarsSaved', 'milesSaved'];

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

  // 1. Gather booking state per confirmation, processing messages oldest →
  //    newest so re-ticketed changes (same conf, new flight/date) win.
  const bookings = {};   // conf → flight object
  collectMessages_('from:united.com subject:"eTicket Itinerary and Receipt" newer_than:1y')
    .forEach(m => { const f = parseUnitedReceipt_(m.body); if (f) remember_(bookings, f, m); });
  collectMessages_('from:alaskaair.com subject:"Your flight is booked" newer_than:1y')
    .forEach(m => { const f = parseAlaskaBooked_(m.body); if (f) remember_(bookings, f, m); });

  // 2. Cancellations flip status.
  collectMessages_('from:united.com (subject:"cancellation is complete" OR subject:"reservation has been canceled") newer_than:1y')
    .forEach(m => {
      const conf = matchOne_(m.body, /Confirmation number:\s*([A-Z0-9]{6})/i) ||
                   matchOne_(m.subject, /\(([A-Z0-9]{6})\)/);
      if (conf && bookings[conf]) bookings[conf].status = 'canceled';
    });
  collectMessages_('from:alaskaair.com subject:(canceled OR cancelled) newer_than:1y')
    .forEach(m => {
      const conf = matchOne_(m.body, /Confirmation code:\s*\n?\s*([A-Z]{6})/i);
      if (conf && bookings[conf]) bookings[conf].status = 'canceled';
    });

  // 3. FUTURE FLIGHTS ONLY: drop past and canceled bookings.
  const todayIso = isoDate_(new Date());
  const future = Object.values(bookings).filter(f =>
    f.date && f.date >= todayIso && f.status !== 'canceled');

  future.forEach(f => upsertFlight_(ss, f));

  // 4. Date-change savings: a future-flight-credit notice whose conf is a
  //    still-active future booking means a fare difference came back to you.
  const savings = ss.getSheetByName(SAVINGS_SHEET);
  const existingNotes = savings.getLastRow() > 1
    ? savings.getRange(2, 3, savings.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  collectMessages_('from:united.com subject:"future flight credit" newer_than:1y')
    .forEach(m => {
      const conf = matchOne_(m.body, /Confirmation Number:?,?\s*([A-Z0-9]{6})/i);
      const amt = matchNum_(m.body, /\$\s?([\d,]+\.\d{2})/);
      if (!conf || !amt) return;
      const f = future.find(x => x.confirmation === conf);
      if (!f) return;   // credit from a cancellation, not a rebooking win
      const note = 'Fare difference returned as credit on ' + conf;
      if (existingNotes.some(n => n.includes(conf))) return;
      savings.appendRow([isoDate_(m.date), f.origin + '-' + f.destination + ' (' + conf + ')', note, amt, 0]);
    });
}

function collectMessages_(query) {
  const out = [];
  GmailApp.search(query, 0, 100).forEach(thread =>
    thread.getMessages().forEach(msg => out.push({
      body: msg.getPlainBody(),
      subject: msg.getSubject(),
      date: msg.getDate(),
      id: msg.getId(),
    })));
  out.sort((a, b) => a.date - b.date);   // oldest first → newest state wins
  return out;
}

function remember_(bookings, flight, msg) {
  flight.sourceEmail = 'https://mail.google.com/mail/u/0/#all/' + msg.id;
  flight.lastSynced = isoDate_(new Date());
  const prev = bookings[flight.confirmation];
  if (prev && (prev.date !== flight.date || prev.flightNumbers !== flight.flightNumbers)) {
    flight.notes = ((flight.notes || '') + ' [rebooked from ' + prev.flightNumbers + ' on ' + prev.date + ']').trim();
    // Carry cost forward when a change email has no payment breakdown.
    if (!flight.cashPaid && !flight.creditsApplied && !flight.milesPaid) {
      flight.cashPaid = prev.cashPaid; flight.creditsApplied = prev.creditsApplied;
      flight.milesPaid = prev.milesPaid; flight.awardFees = prev.awardFees;
    }
  }
  bookings[flight.confirmation] = flight;
}

/**
 * Upsert keyed by confirmation. Parser-owned fields are updated; a cell the
 * user hand-edited to something different gets flagged in notes instead of
 * silently clobbered (notes/sourceEmail/lastSynced always refresh).
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
  FLIGHT_COLS.forEach((col, i) => {
    const parsed = flight[col];
    if (parsed === undefined || parsed === '') return;
    const cur = rows[idx][i];
    const curStr = cur instanceof Date ? isoDate_(cur) : String(cur);
    if (col === 'notes' || col === 'sourceEmail' || col === 'lastSynced') {
      sh.getRange(rowNum, i + 1).setValue(parsed);
    } else if (curStr === '' || curStr === String(parsed)) {
      sh.getRange(rowNum, i + 1).setValue(parsed);
    } else {
      conflicts.push(col + ': parsed "' + parsed + '" vs sheet "' + curStr + '"');
    }
  });
  if (conflicts.length) {
    const noteCell = sh.getRange(rowNum, FLIGHT_COLS.indexOf('notes') + 1);
    noteCell.setValue(((noteCell.getValue() || '') + ' [sync conflict — ' + conflicts.join('; ') + ']').trim());
  }
}

/* =========================== PARSERS =========================== */

/**
 * United "eTicket Itinerary and Receipt". Single- or multi-segment; payment
 * variants seen in the wild:
 *  - award: "Special member price: 12,600 miles" / "Total: 12,600 miles + 5.60 USD"
 *  - credit: "Future flight credit: 93.40 USD" + "Confirmation #: REDACTEDCR2"
 *            + "Future flight credit applied: -93.40 USD" + "Total: 0.00 USD"
 *  - mixed:  credit + "An additional amount of 3.97 USD ... charged to Visa"
 *  - reissue: "Previous Ticket Balance" + additional collection
 *  - plain cash: "Total: 285.02 USD" on a card
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

  const miles = matchNum_(body, /Special member price:\s*([\d,]+) miles/) ||
                matchNum_(body, /Total:\s*([\d,]+) miles/);
  const creditApplied = matchNum_(body, /Future flight credit applied:\s*-([\d,]+\.\d{2})/);
  const creditSource = matchOne_(body, /Future flight credit:[\s\S]{0,60}?Confirmation #:\s*([A-Z0-9]{6})/);
  const additional = matchNum_(body, /An additional amount of ([\d,]+\.\d{2}) USD/);
  const totalUsd = matchNum_(body, /Total:\s*([\d,]+\.\d{2}) USD/);
  const awardFee = matchNum_(body, /Total:\s*[\d,]+ miles \+ ([\d,]+\.\d{2}) USD/);

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
 *  - "Confirmation code: REDACTED8"
 *  - "Flight 1 · Sat Aug 29" / "AS 3100"
 *  - "Departure date: Aug 29 at 5:22 PM"
 *  - "5000 points have been redeemed"
 *  - "$5.60 to be charged to the VISA card"
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
      if (v instanceof Date) v = isoDate_(v);
      o[c] = v;
    });
    return o;
  });
}

/* =========================== HELPERS =========================== */

function matchOne_(s, re) { const m = s.match(re); return m ? m[1] : null; }
function matchNum_(s, re) { const m = s.match(re); return m ? Number(m[1].replace(/,/g, '')) : 0; }
function round2_(n) { return Math.round(n * 100) / 100; }
function isoDate_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
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
