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
- `data/flights-seed.csv` — extracted current flights, paste into the Flights tab.

## Setup (one time)

1. Create a Google Sheet ("Flight Tracker").
2. Extensions → Apps Script → paste `Code.gs` → save.
3. Run `setupSheets()`, approve the authorization prompt.
4. Paste `data/flights-seed.csv` into the Flights tab (or run `syncFlights()`).
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

## Privacy note

The web-app endpoint is "anyone with the link" — an unguessable URL, but
technically public. It serves flight metadata only (no passenger names or
payment details). For stricter privacy, redeploy as an Apps Script HtmlService
app behind your Google login.

## Roadmap

- v2: current-price column via a flight-pricing MCP (on-demand or daily trigger),
  price-drop alerting (MailApp), automated savings tally on rebooking.
