// Live smoke test for checkPrice — run `node test-harness.mjs` before
// pushing scraper changes. Hits real Google Flights (no Sheet access);
// dates are ~30 days out so the queries never go stale.
import { chromium } from 'playwright';
import { checkPrice } from './check-prices.mjs';

const date = new Date(Date.now() + 30 * 24 * 3600 * 1000).toISOString().slice(0, 10);
const flights = [
  // "alaska" doubles as a state name — this case regressed once already
  // (query parser served the homepage instead of results).
  { airline: 'Alaska', date, origin: 'SFO', destination: 'SAN' },
  { airline: 'United', date, origin: 'SFO', destination: 'SAN' },
];

const browser = await chromium.launch();
let failed = 0;
for (const f of flights) {
  const t0 = Date.now();
  const price = await checkPrice(browser, f);
  if (price === null) failed++;
  console.log(`RESULT ${f.airline} ${f.origin}->${f.destination} ${f.date}: ${price === null ? 'FAILED' : '$' + price} (${((Date.now() - t0) / 1000).toFixed(1)}s)`);
}
await browser.close();
process.exitCode = failed ? 1 : 0;
