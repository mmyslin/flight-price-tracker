import { chromium } from 'playwright';
import { google } from 'googleapis';

const SHEET_ID = process.env.SHEET_ID;
const SHEET_TAB = 'Flights';
// Signature of the "Economy (exclude Basic)" search URL with no extra filters.
// If the cabin-class selection didn't take (e.g. Google changed the UI), the
// tfs= param won't end this way — used to catch bad reads before they're written.
const EXPECTED_URL_SUFFIX = 'wGYAQLIAQE';

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

async function checkPrice(browser, flight) {
  const query = `${flight.origin} to ${flight.destination} on ${flight.date} one way nonstop ${flight.airline.toLowerCase()}`;
  const url = `https://www.google.com/travel/flights?q=${encodeURIComponent(query)}`;

  for (let attempt = 1; attempt <= 2; attempt++) {
    const page = await browser.newPage();
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });

      await page.waitForTimeout(500);
      const consentBtn = page.getByRole('button', { name: /reject all|i agree/i }).first();
      if (await consentBtn.count()) {
        await consentBtn.click().catch(() => {});
      }

      const priceEl = page.locator('span[aria-label$="US dollars"]').first();
      // Baseline (default "include Basic") price, used below to detect the
      // filtered re-fetch actually landing — the URL updates via fast
      // client-side routing well before the new fare data arrives, so
      // reading the price right after the URL check can still return this
      // stale, cheaper Basic-inclusive number.
      const baseline = await priceEl.getAttribute('aria-label', { timeout: 10000 }).catch(() => null);

      const cabinDropdown = page.locator('[role="combobox"]').filter({ hasText: 'Economy' });
      await cabinDropdown.first().click({ timeout: 15000 });

      const excludeBasicOption = page
        .locator('[role="option"]')
        .filter({ hasText: 'Economy (exclude Basic)' });
      await excludeBasicOption.first().click({ timeout: 10000 });

      await page.waitForURL((u) => u.pathname.includes('/search'), { timeout: 15000 });

      const tfs = new URL(page.url()).searchParams.get('tfs') || '';
      if (!tfs.endsWith(EXPECTED_URL_SUFFIX)) {
        throw new Error(`unexpected tfs param after cabin selection: ${page.url()}`);
      }

      if (baseline) {
        await page
          .locator(`span[aria-label$="US dollars"]:not([aria-label="${baseline}"])`)
          .first()
          .waitFor({ state: 'attached', timeout: 8000 })
          .catch(() => {});
      }
      await page.waitForLoadState('networkidle', { timeout: 8000 }).catch(() => {});

      const ariaLabel = await priceEl.getAttribute('aria-label', { timeout: 10000 });
      const price = parseInt(ariaLabel, 10);
      if (!Number.isFinite(price)) throw new Error(`couldn't parse price from "${ariaLabel}"`);

      return price;
    } catch (err) {
      console.warn(`  attempt ${attempt} failed for ${flight.origin}->${flight.destination} ${flight.date}: ${err.message}`);
      if (attempt === 2) return null;
    } finally {
      await page.close();
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

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
