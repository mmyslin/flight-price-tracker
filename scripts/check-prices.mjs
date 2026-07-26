import { chromium } from 'playwright';
import { google } from 'googleapis';
import { mkdirSync } from 'fs';
import { pathToFileURL } from 'url';

const SHEET_ID = process.env.SHEET_ID;
const SHEET_TAB = 'Flights';
// Signature of the "Economy (exclude Basic)" search URL with no extra filters.
// If the cabin-class selection didn't take (e.g. Google changed the UI), the
// tfs= param won't end this way — used to catch bad reads before they're written.
const EXPECTED_URL_SUFFIX = 'wGYAQLIAQE';

// Look like a normal Chrome session: Playwright's default headless UA says
// "HeadlessChrome", which gets Google Flights served in a degraded/throttled
// variant where prices render slowly or never — the main source of the flaky
// "waiting for span[aria-label$=US dollars]" timeouts in CI.
const CONTEXT_OPTS = {
  userAgent:
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
  viewport: { width: 1440, height: 900 },
  locale: 'en-US',
  timezoneId: 'America/Los_Angeles',
};

const ATTEMPTS = 3;

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

async function loadFlights(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A1:S`,
  });
  const [header, ...rows] = res.data.values;
  const col = Object.fromEntries(header.map((name, i) => [name, i]));
  const required = ['airline', 'date', 'origin', 'destination', 'status', 'milesPaid', 'currentPrice'];
  const missing = required.filter((name) => col[name] === undefined);
  if (missing.length) {
    throw new Error(`Sheet header row is missing expected column(s): ${missing.join(', ')} (found: ${header.join(', ')})`);
  }
  const today = todayIso();

  return rows
    .map((row, i) => ({
      rowNum: i + 2,
      airline: row[col.airline] || '',
      date: row[col.date] || '',
      origin: row[col.origin] || '',
      destination: row[col.destination] || '',
      status: row[col.status] || '',
      milesPaid: Number(row[col.milesPaid] || 0),
      currentPrice: row[col.currentPrice] || '',
    }))
    .filter(
      (f) =>
        f.status === 'active' &&
        f.date >= today &&
        f.milesPaid === 0 &&
        f.origin &&
        f.destination
    );
}

// Reads the top listed fare's aria-label ("94 US dollars"), riding out
// Google's transient "Oops, something went wrong" error panel by doing what
// its own Reload button does. Returns null if no fare ever renders.
async function readTopFare(page) {
  const priceEl = page.locator('span[aria-label$="US dollars"]').first();
  for (let tries = 0; tries < 3; tries++) {
    const label = await priceEl.getAttribute('aria-label', { timeout: 20000 }).catch(() => null);
    if (label != null) return label;
    if (!(await page.getByText(/something went wrong/i).count())) return null;
    await page.reload({ waitUntil: 'domcontentloaded' });
    await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
  }
  return null;
}

async function checkPrice(browser, flight) {
  // "airlines" suffix matters: a bare "alaska" reads as the STATE to
  // Google's query parser, which then fails to parse the whole query and
  // dumps the session on the Flights homepage — no results, no prices.
  const query = `${flight.origin} to ${flight.destination} on ${flight.date} one way nonstop ${flight.airline.toLowerCase()} airlines`;
  const url = `https://www.google.com/travel/flights?q=${encodeURIComponent(query)}`;

  for (let attempt = 1; attempt <= ATTEMPTS; attempt++) {
    if (attempt > 1) await new Promise((r) => setTimeout(r, 2000 * attempt));
    // Fresh context per attempt: clean cookies and a clean bot-score, so a
    // degraded first serve doesn't poison the retry.
    const context = await browser.newContext(CONTEXT_OPTS);
    const page = await context.newPage();
    await page.addInitScript(() =>
      Object.defineProperty(navigator, 'webdriver', { get: () => undefined }));
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
      await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});

      const consentBtn = page.getByRole('button', { name: /reject all|i agree/i }).first();
      if (await consentBtn.count()) {
        await consentBtn.click().catch(() => {});
      }

      // Baseline (default "include Basic") price. Required: if no price has
      // rendered by now the page is a degraded serve — fail fast into a
      // retry with a fresh context instead of limping on to a later timeout.
      const baseline = await readTopFare(page);
      if (baseline == null) throw new Error('no price rendered on results page');

      const cabinDropdown = page.locator('[role="combobox"]').filter({ hasText: 'Economy' });
      await cabinDropdown.first().click({ timeout: 15000 });

      // Wait for the menu itself (any VISIBLE option — the page also holds
      // hidden role=option nodes in closed autocomplete listboxes), THEN
      // look for the target option. Waiting only on the filtered locator
      // can't distinguish "menu never opened" (transient, worth retrying)
      // from "menu open but no such option" (permanent for some airlines).
      await page.locator('[role="option"]:visible').first().waitFor({ timeout: 10000 });
      const excludeBasicOption = page
        .locator('[role="option"]:visible')
        .filter({ hasText: /exclude basic/i });

      if ((await excludeBasicOption.count()) === 0) {
        // No exclude-Basic tier offered for this airline: fall back to the
        // default lowest-economy fare rather than failing the run.
        console.warn(`  NOTE ${flight.origin}->${flight.destination} ${flight.date}: no "exclude Basic" option offered — using lowest economy fare`);
        const price = parseInt(baseline, 10);
        if (!Number.isFinite(price)) throw new Error(`couldn't parse price from "${baseline}"`);
        return price;
      }

      await excludeBasicOption.first().click({ timeout: 15000 });

      await page.waitForURL((u) => u.pathname.includes('/search'), { timeout: 15000 });

      const tfs = new URL(page.url()).searchParams.get('tfs') || '';
      if (!tfs.endsWith(EXPECTED_URL_SUFFIX)) {
        throw new Error(`unexpected tfs param after cabin selection: ${page.url()}`);
      }

      // The URL now encodes exclude-Basic (tfs check above), but the DOM
      // still shows the stale include-Basic list: the in-place refetch is
      // slow and sometimes dies outright with Google's "Oops, something
      // went wrong" panel. Reading the live DOM here is what wrote Basic
      // fares (e.g. $59 instead of the real $94) into the Sheet. A hard
      // reload of the filtered URL gives a clean server render that can
      // only ever contain exclude-Basic fares.
      await page.reload({ waitUntil: 'domcontentloaded' });
      await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});

      const ariaLabel = await readTopFare(page);
      if (ariaLabel == null) throw new Error('no price rendered after cabin filter');
      const price = parseInt(ariaLabel, 10);
      if (!Number.isFinite(price)) throw new Error(`couldn't parse price from "${ariaLabel}"`);

      return price;
    } catch (err) {
      console.warn(`  attempt ${attempt} failed for ${flight.origin}->${flight.destination} ${flight.date}: ${err.message.split('\n')[0]}`);
      try {
        mkdirSync('debug', { recursive: true });
        await page.screenshot({
          path: `debug/${flight.origin}-${flight.destination}-${flight.date}-attempt${attempt}.png`,
        });
      } catch { /* screenshots are best-effort */ }
      if (attempt === ATTEMPTS) return null;
    } finally {
      await context.close();
    }
  }
  return null;
}

async function main() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY),
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  const sheets = google.sheets({ version: 'v4', auth });

  const flights = await loadFlights(sheets);
  console.log(`Checking ${flights.length} upcoming cash-fare flight(s)...`);

  const browser = await chromium.launch();
  const updates = [];
  let failures = 0;

  for (const flight of flights) {
    await new Promise((r) => setTimeout(r, 1000 + Math.random() * 2000));
    const price = await checkPrice(browser, flight);
    if (price == null) {
      console.warn(`  SKIP ${flight.origin}->${flight.destination} ${flight.date}: could not read a price`);
      failures++;
      continue;
    }
    console.log(`  ${flight.origin}->${flight.destination} ${flight.date}: $${price} (was ${flight.currentPrice || 'blank'})`);
    updates.push({ rowNum: flight.rowNum, oldPrice: flight.currentPrice, newPrice: price });
  }

  await browser.close();

  if (updates.length) {
    const data = [];
    for (const u of updates) {
      if (u.oldPrice !== '') {
        data.push({ range: `${SHEET_TAB}!R${u.rowNum}`, values: [[u.oldPrice]] });
      }
      data.push({ range: `${SHEET_TAB}!P${u.rowNum}`, values: [[u.newPrice]] });
      data.push({ range: `${SHEET_TAB}!U${u.rowNum}`, values: [[todayIso()]] });
    }
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { valueInputOption: 'RAW', data },
    });
    console.log(`Updated ${updates.length} row(s) in the Sheet.`);
  }

  if (failures) {
    console.error(`${failures} flight(s) could not be checked — see warnings above.`);
    process.exitCode = 1;
  }
}

export { checkPrice };

// Only auto-run when executed directly, so a test harness can import
// checkPrice without kicking off a full Sheet-backed run.
if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((err) => {
    console.error(err);
    process.exitCode = 1;
  });
}
