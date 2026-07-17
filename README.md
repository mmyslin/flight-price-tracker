# Flight Tracker

Personal dashboard of upcoming flights and what they actually cost — cash, miles,
and previous credits applied — to make rebook/cancel decisions easy when prices drop.

## Architecture

```
Gmail inbox ──(Apps Script, on-demand / daily)──▶ Google Sheet ──(doGet JSON)──▶ dashboard/index.html
                                                     ▲
                                       manual corrections / credit annotations
```

The Sheet is the source of truth: parser mistakes and messy credit breakdowns get
corrected there, and future features (current-price column, savings tally, alerting)
read/write the same Sheet.

## Layout

- `dashboard/index.html` — self-contained dashboard. Set `DATA_URL` at the top to
  the Apps Script `/exec` URL; empty = renders embedded seed data.
- `apps-script/Code.gs` — Gmail → Sheet sync (United + Alaska parsers), upsert
  that never clobbers manual edits, and the JSON endpoint.
- `scripts/check-prices.mjs` — headless-browser Google Flights price check, run
  daily by GitHub Actions (see below).

## Setup (one time)

1. Create a Google Sheet ("Flight Tracker").
2. Extensions → Apps Script → paste `Code.gs` → save.
3. Run `setupSheets()`, approve the authorization prompt.
4. Run `syncFlights()` to pull your current bookings from Gmail (or paste rows
   into the Flights tab by hand — see `FLIGHT_COLS` in `Code.gs` for columns).
5. Deploy → New deployment → **Web app**, Execute as *Me*, access *Anyone with the link*.
6. Copy the `/exec` URL into `DATA_URL` in `dashboard/index.html`.
7. Open the dashboard: `python3 -m http.server 8742 --directory dashboard` →
   http://localhost:8742 (or just open the file directly in a browser).

## Conventions

- **Total cost = cashPaid + creditsApplied** (credits are real value spent);
  miles shown separately, with `awardFees` counted inside cashPaid.
- Multi-segment bookings carry the whole cost on the **first segment** so sums
  don't double-count.
- `status`: active / flown / canceled — dashboard shows only future `active` flights.
- Rebooking wins go in the **Savings** tab; the dashboard tallies them.

## Automated price check (GitHub Actions)

`scripts/check-prices.mjs` runs daily via `.github/workflows/price-check.yml` (14:30
UTC ≈ 7:30am Pacific). For each upcoming `active` cash-fare flight (award flights with
`milesPaid` set are skipped — their price stays a manual/personal lookup), it opens
Google Flights headlessly, switches the cabin-class filter to **"Economy (exclude
Basic)"** (matches United's real Standard/Economy tier — see Conventions), reads the
lowest fare, and writes it to the Sheet's `currentPrice` column — shifting the prior
value into `previousPrice` first so the dashboard's trend arrow (↗/↘) stays accurate.
It never touches award-flight rows.

One-time setup to enable it:

1. In Google Cloud Console, create (or reuse) a project and enable the **Google
   Sheets API**.
2. Create a **service account**, then generate a JSON key for it (IAM & Admin →
   Service Accounts → Keys → Add key → JSON).
3. Open the Flight Tracker Sheet → Share → add the service account's email
   (`...@...iam.gserviceaccount.com`) as an **Editor**.
4. In the GitHub repo → Settings → Secrets and variables → Actions, add a secret
   named `GOOGLE_SERVICE_ACCOUNT_KEY` containing the full contents of the JSON key
   file.
5. Trigger a test run from the Actions tab (`Daily flight price check` → Run
   workflow) before relying on the schedule.

The workflow fails loudly (red X, GitHub email notification) if any flight's price
couldn't be read, so a broken Google Flights selector doesn't fail silently.

## Price-drop email alerts

`checkPriceAlerts()` (in `Code.gs`) runs on its own daily time-driven trigger
(set up in the Apps Script editor's Triggers page), scheduled after the price
check above so it always sees same-day fresh prices. For each `active` flight
where `currentPrice` has dropped **$25+** below what was actually paid
(`cashPaid + creditsApplied`), it emails a summary to the script owner's own
Google account (`Session.getEffectiveUser().getEmail()` — no address is
hardcoded in source). The `alertedPrice` column tracks the last price already
emailed for a row, so a still-cheap fare doesn't re-alert every day — only a
new, lower price (or first time crossing the threshold) triggers another
email. No additional setup beyond the one-time `MailApp` authorization prompt
the first time the function runs.

## Privacy note

The web-app endpoint is "anyone with the link" — an unguessable URL, but
technically public. It serves flight metadata only (no passenger names or
payment details). For stricter privacy, redeploy as an Apps Script HtmlService
app behind your Google login.

## Roadmap

- ~~current-price column, daily automated check~~ — done via GitHub Actions (see above).
- ~~price-drop alerting~~ — done via `checkPriceAlerts()` (see above).
- v2: automated savings tally on rebooking.
