# Fuel Surcharge Board

Tracks the published fuel surcharge (FSC) of ACE Courier and other carriers serving
British Columbia, and shows how each compares with the matching ACE rate.

---

## Quick start

```powershell
.\Get-FuelSurcharges.ps1        # scrape every carrier and write the data files
.\Start-Preview.ps1             # view the public dashboard at http://localhost:8080
.\Register-FscTask.ps1 -RunNow  # refresh the local copy automatically every morning
```

Runs on Windows PowerShell 5.1 or PowerShell 7, using `curl` as a fallback for sites
that reject PowerShell's own web requests. No other dependencies.

---

## What it collects

| Carrier | Services tracked | Source |
|---|---|---|
| **ACE Courier** | BC, Alberta, FTL / Direct Drive | Public FAQ page |
| **Comox Pacific Express** | Under / 10,000 lb and over | Homepage, including next week's posted rate |
| **Overland West Freight Lines** | Under / 10,000 lb and over | Homepage, including next week's posted rate |
| **Hi-Way 9 Express** | Under / 10,000 lb and over | Fuel surcharge page |
| **Steele's Transfer** | Under / 10,000 lb and over | Fuel surcharge page |
| **Grimshaw Trucking** | Under / 10,000 lb and over | Fuel surcharge updates page |

The board covers direct competitors for heavy LTL freight, compared on BC rates. Parcel
and courier networks are deliberately left out. Each carrier's 10,000 lb-and-over rate is
its truckload rate, so it is compared with ACE's FTL / Direct Drive rate; the lighter rate
is compared with ACE's BC rate.

Competitors without a scrapable rate (Van-Kam, Bandstra, Manitoulin, Rosenau, Clark) are
tracked in the local, gitignored `data/manual.json` and `data/competitor-reports.json`.
A dated rate is compared with ACE's BC surcharge on the same date. Carriers marked
`"publish": true` in `manual.json` also appear on the public board as **reported**, with
their as-of date; the rest stay local.

---

## Diesel vs ACE surcharge trend

`Get-FuelTrend.ps1` builds the 12-month chart below the rate tables:

- **Diesel:** Natural Resources Canada weekly average retail diesel prices (taxes
  included). The BC figure is the average of the seven BC cities NRCan surveys:
  Abbotsford, Fort St. John, Kamloops, Kelowna, Prince George, Vancouver and Victoria.
- **ACE surcharge:** `data/ace-fsc-history.json`. Each change is confirmed from its
  effective date through `confirmed_through`; between entries the rate is not on record
  and the chart shows a gap. Loaded from ACE's surcharge schedule; new changes are added
  automatically from the daily scrape.
- **Correlation:** ACE's day-weighted average BC surcharge for each week against that
  week's BC diesel price, over weeks with ACE's rate on record for at least four days.
  Both series are weekly, which matches how often ACE changes its surcharge.

To fill a gap, add the change to `data/ace-fsc-history.json` with its effective date,
BC and Alberta rates, and the date it was last confirmed.

---

## How the comparison works

Each rate is measured against the matching ACE rate:

- Truckload / FTL services → **ACE FTL / Direct Drive**
- Everything else → **ACE British Columbia**

`delta_points` is the gap in percentage points; `delta_percent` is the relative gap.
Both are recorded because a claim like "14% lower" means very different things
depending on which one is meant.

---

## Files

| File | Committed | Contents |
|---|---|---|
| `site/index.html` | yes | The dashboard |
| `site/latest.json` | yes | **Public snapshot** — live published rates only, no notes or errors |
| `data/history.jsonl` | yes | Append-only log of published rates, one row per carrier / service / effective date |
| `data/latest.json` | no | Full snapshot including notes, errors and manual entries |
| `site/data.js` | no | The full snapshot for the local dashboard |
| `data/manual.json` | no | Manually entered rates |

The public snapshot is built from an allowlist: only rates with status `ok` or
`upcoming`, and only the fields the page displays. Served over http, the dashboard
reads `site/latest.json`. Opened directly from disk, it reads the full `site/data.js`
and also shows carriers with no published rate.

---

## Hosting

`.github/workflows/fsc.yml` runs the scraper daily at 06:00 Pacific on GitHub Actions,
commits the public snapshot and history, and publishes the dashboard to GitHub Pages.
Only `site/index.html`, `site/latest.json` and `site/trend.json` are uploaded to Pages.

If a carrier site blocks GitHub's servers, the office PC's daily task
(`Publish-FscBoard.ps1`) publishes `data/local-public.json`, and the cloud run uses its
rows for any carrier it can't reach while those rates are still in effect.

To enable it: **Settings → Pages → Source → GitHub Actions**, then
**Actions → Fuel surcharge check → Run workflow**.

---

## Local daily refresh

`Register-FscTask.ps1` creates a Windows scheduled task that runs the scraper every
morning under the current user, with no elevation.

```powershell
.\Register-FscTask.ps1                # daily at 07:15 (default)
.\Register-FscTask.ps1 -At 06:30      # pick a different time
.\Register-FscTask.ps1 -Unregister    # remove it
```

If the machine is off at the scheduled time, the run happens at the next opportunity.
Each run appends a line to `data/run.log`.

---

## When a scraper breaks

Carrier sites change. When a page no longer matches, that carrier's adapter throws, the
error is recorded in the local snapshot, and the carrier drops out of the public one
rather than showing a stale number.

Each adapter is a `Get-Fsc*` function in `Get-FuelSurcharges.ps1`, with the text it
expects shown in a comment. Fetch the page, look at the text around the number, and
adjust the pattern.

---

## Caveats

- **Rates change weekly.** Most carriers reset on Monday, based on a diesel index
  published the previous Thursday. A daily check catches every change.
- **A surcharge is a percentage of a base rate.** A lower percentage does not by itself
  mean a lower price; base rates, weight rules and accessorials usually matter more.
- **Confirm before shipping.** These are published list surcharges. Negotiated accounts
  and accessorials are not included.
