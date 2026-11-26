# ACE expense capture

Google Apps Script attached to the **Ace Expenses** sheet. Every 15 minutes it files
new receipts from Gmail and Drive, renames them, and logs them on **All Expenses**.

## Naming convention

```
YYYY-MM-DD - <Trip> - <Vendor> - <Total>.<ext>
2026-10-08 - Port Hardy Oct 7-9 - Karai Sushi - 39.62.jpg
```

- **Date** is the charge date (matches *Charge Date* on the sheet) and sorts files in order.
- **Trip** matches the trip folder name and *Submission Description*.
- **Total** tells two same-day receipts from one vendor apart and ties the file to its row.
- Batch number is **not** in the name: it isn't known when you capture. The batch folder
  (`Batch 007 – …`) and *Submission #* carry it.

## How you use it

| Receipt | Do this |
|---|---|
| Photo | Upload to the trip's folder in **Ace Expenses** (best: trip is certain), or to **Ace Expenses** itself and let the date pick the trip. |
| Email | Apply the Gmail label **ACE-Receipt** (or set a Gmail filter for vendors like BC Ferries / Booking.com). |
| New trip | Add a row to the **Trips** tab: Trip, Start, End. |

Then check rows with Status **CHECK** (yellow): vendor/amount unread, guessed, or no trip
matched. Fill in Submitted / Submission # at submission time as before.

## Install (once, ~5 min)

1. Open the **Ace Expenses** sheet → *Extensions → Apps Script*.
2. Replace `Code.gs` with this folder's `Code.gs`.
3. *Project Settings* → tick **Show "appsscript.json"**, replace it with this folder's file.
4. *Project Settings → Script properties* → add `ANTHROPIC_API_KEY` (from
   console.anthropic.com). Optional; without it the script uses Google OCR, which reads
   dates and totals passably and vendor names poorly.
5. Run `setup` from the editor and approve access. Reload the sheet for the **Expenses** menu.

Only files added **after** `setup` runs are touched; existing receipts keep their names.

## Cost

Model is `claude-opus-5-5` at low effort: roughly 2–4¢ per receipt. Change `CFG.MODEL` to
`claude-haiku-5-5` for well under 1¢ if its reads hold up on your receipts.
