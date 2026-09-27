# My-Finance-Tracker

# 💰 My Finance Tracker

A lightweight personal finance tracker built entirely on **Google Sheets, Google Forms, and Google Apps Script** — no external database or hosting required.

Log expenses, income, transfers, lending, borrowing, and investments in a few seconds from any device (including mobile), and see everything automatically organized into monthly sheets with a live, filterable dashboard.

## Features

- 📝 **Simple entry form** (Google Form) — works on desktop and mobile, no app install needed
- 📅 **Automatic monthly sheets** — each transaction is filed into the correct month's sheet (e.g. `Sep-2026`) based on the transaction date, not just the submission date
- 🔄 **Smart date handling** — leave the date blank to default to today, or enter a past date and it's correctly filed under that month
- 📊 **Live Dashboard** — auto-updating summary of total income, expenses, savings, and breakdowns by category, account, payment method, and expense nature
- 🔍 **Filterable trends** — dropdown filters for Year and Category, showing:
  - Filtered totals (expenses, income, net)
  - Month-by-month trend for the selected year
  - Year-over-year trend for the selected category
- 🧾 **Categorized transactions** — expense category, account/wallet, payment method, fixed vs. variable, recurring flag, and free-form notes
- 🗂️ **Invoice item detail sheets** (`_Items`) — structure in place for itemized invoice/receipt breakdowns
- ⚙️ **One-time setup** — a single script run creates the form, sheets, dashboard, and trigger; safe to re-run without duplicating anything

## How it works

1. `setupFinanceTracker()` is run once from the Apps Script editor attached to a Google Sheet.
2. It creates:
   - A Google Form ("💰 My Finance Tracker") with all transaction fields
   - Supporting sheets: `Dashboard`, `Settings`, `Categories`, `Accounts`
   - A form-submit trigger that processes each response automatically
3. On every form submission:
   - The transaction date and amount are parsed and validated
   - A unique Transaction ID is generated
   - The entry is appended to that month's transaction sheet (created automatically if it doesn't exist yet)
   - The Dashboard and filtered trend views refresh automatically

## Tech stack

- **Google Apps Script** (JavaScript, V8 runtime)
- **Google Forms** — data entry
- **Google Sheets** — storage, reporting, and dashboard
- Uses `FormApp`, `SpreadsheetApp`, `PropertiesService`, and installable/simple triggers (`onFormSubmit`, `onEdit`)

## Setup

1. Create a new Google Sheet.
2. Open **Extensions → Apps Script** and paste in `FinanceTracker.gs`.
3. Save, then run `setupFinanceTracker` from the function dropdown.
4. Authorize the requested permissions (Sheets + Forms access) when prompted.
5. Open the form link shown in the execution log, bookmark it on your phone, and start logging transactions.

## Roadmap / future modules

- Invoice OCR (auto-extract line items from receipt photos)
- Bank statement import & reconciliation
- Budget management with alerts
- AI-based spending analysis and insights

## License

Feel free to fork and adapt for personal use.
