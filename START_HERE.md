# Bollin Clinic Stock Manager — START HERE (project background)

Paste or upload this file at the start of a new chat so the assistant has full context without re-reading the whole history.

## What this project is
A complete inventory + asset management web app for **Bollin Clinic** (aesthetic surgery clinic, Altrincham UK). Built as a single-file web app backed by Google Sheets. Owner/admin: **Yasar** (nurse).

## Architecture (important)
- **Frontend:** one file, `index.html`, plain HTML/CSS/JS (no build step). Hosted on GitHub Pages.
- **Backend:** one file, `Code.gs`, a Google Apps Script Web App attached to a Google Sheet. All data lives in the Sheet's tabs.
- The frontend talks to the backend through one URL baked into `index.html` at `CONFIG.API_URL`. **This URL is the only thing that differs between environments.**

## Two environments (staging + production)
| | Repo | URL | Backend Sheet |
|---|---|---|---|
| **PRODUCTION** (live) | `database` | https://bollin.hashirhub.uk | `Bollin_Inventory_Backend` |
| **STAGING** (test) | `database-staging` | https://yasarshah65-sketch.github.io/database-staging/ | `Bollin_Inventory_STAGING` |

- The same `Code.gs` is used on **both** — it's environment-agnostic.
- Two index files per update: `index.html` (production URL) and `index_STAGING.html` (staging URL + amber "STAGING" badge). They are otherwise identical.
- GitHub Pages needs the file named exactly `index.html` in the repo, so when uploading the staging build, rename `index_STAGING.html` → `index.html` inside the `database-staging` repo.

## The fixed update workflow (do this EVERY update)
1. Assistant provides: `index.html` (production), `index_STAGING.html` (staging), `Code.gs`, and an updated START_HERE.
2. **Staging first:** upload staging index to `database-staging`; paste `Code.gs` into the *staging* Apps Script; run `migrate` if a tab/column changed; Deploy → Manage deployments → New version.
3. **Test** on the staging URL with dummy data.
4. **Happy → promote to production:** upload production index to `database`; paste the same `Code.gs` into the *production* Apps Script; run `migrate` if needed; New version.
5. Not happy → report back; production is untouched.

**Deploy reminders that are easy to forget:** after pasting Code.gs you must click Deploy → Manage deployments → pencil → **New version** (saving alone doesn't publish). Run `migrate` whenever a tab or column was added (safe to run anytime — only creates what's missing, never touches data). Run `setupTriggers` only when email/trigger schedules change (skip on staging to avoid test emails).

## Roles & login
- Accounts live in the **Users** tab (Username, Password, Role, DisplayName, Active). Admin types a plain password; it auto-converts to a secure hash on first login. Users change their own password in-app; only admin edits usernames/roles. Lockout fix = overwrite the Password cell with a new plain password.
- **admin** = everything. **staff** = all except adding/editing items, labels, admin sections (Stores, Obsolete, Assets), and editing logged-used lines. **common** (shared scrub login) = scan stock in/out, log instruments used, submit requests only — no values, no implants, no dispatch/receive.
- Seeded accounts after a fresh migrate: `yasar`/admin, `stockteam`/staff, `scrubs`/common — all password `ChangeMe123` (change immediately).

## Settings tab keys (comma-separate multiple emails; read live, no redeploy)
`app_name`, `alert_email`, `escalation_email`, `ack_hours` (48), `expiry_days` (90), `slow_days`, `activity_email` (midnight digest), `request_email` (new request), `implant_email` (new implant order), `sticker_margins` (saved sticker layout preset, JSON).

## Feature map (what exists)
- **Trackers:** Consumables, Meds, Garments, Instruments, Linen. Columns include Code, Barcode, Name, Category, Supplier, Location, Unit, Qty, ReorderLevel, UnitCost, Expiry, Batch, Notes, Status, Obsolete/ObsoleteBy/ObsoleteAt. Category + Location are dropdowns with "＋ add new".
- **Scanning:** Netum USB barcode scanner (keyboard-wedge). Unknown barcodes prompt a link/register dialog.
- **Sterilisation** (Dispatch Log workflow): tabs = 1·Log used → 2·Dispatch → 3·Receive → On hand → History. Instruments self-register on first scan; expiry auto-set to dispatch date + 1 year on receive; items ≤60 days highlighted. Admin can edit/delete a logged-used line before dispatch.
- **Implants:** order → receive → return extras; email on new order; flat landscape PDF report (Today/Week/Month/Custom).
- **Requests:** submitter form; stock-team response dropdown (Ordered, Awaiting Supplier Response, Back order, Out of stock, Received, In stock, Other); email on new request; blinking nav badge + dashboard tile.
- **Stocktake:** scan-to-count, variance, Unit cost + Variance £ columns, landscape PDF with totals.
- **Labels:** single-item barcode/QR print (name above, code below); multiselect Barcode sheet (10/page) and QR sheet (15/page) PDFs with cut guides; delete-barcode (✕) to regenerate.
- **Activity log** (Transactions tab): logs EVERY write action (stock, items add/edit with exact change text, sterilisation, implants, requests, stocktakes, stores, assets, gas, barcodes, alerts). Filters by activity, person, tracker + search. Midnight email digest. Landscape PDF report.
- **Stock value** explorer (admin/staff): sortable by value/expiry/supplier/qty/name.
- **Stores & transfers** (admin): manage stores per tracker, bulk multi-select move.
- **Obsolete stock** (admin): dead stock hidden from lists/alerts, shows who made it obsolete + when, restorable. Obsolete items are excluded from the low-stock and out-of-stock dashboard drill-downs, which also have tracker/store/category filters.
- **Assets** (admin only, invisible to others): machines with serial, function, supplier+contact, location, last/next maintenance, engineering statuses; ≤30-day due highlight; landscape PDF.
- **Item remarks:** every item's edit dialog has a Remarks field (stored in the tab's Notes column, e.g. "on back order"). Items with remarks are highlighted with a gold pill flag + row tint in all tracker tables and the low/out-of-stock drill-downs; clearing the remark returns them to normal.
- **Appearance / themes** (⚙ top-right popover, per-device via localStorage, key `bollin_prefs`): 7 themes (Day, Night, Warm, Rose, Forest, Mono, Midnight — includes true dark modes), 5 accent colours, 3 fonts (Plex/Inter/System), density (Cozy/Compact), corner radius (Default/Rounded/Sharp). Purely visual, no backend.
- **Procedure case costing** (nav item "Procedure costing", visible to ALL roles; financials admin-only): start a procedure with surgeon + procedure + date + optional patient ref, then scan any item (meds/consumables/instruments/garments — auto-detected) to add it to a local cart. Scanning the same item again bumps its qty; each cart line has +/− qty steppers. **Stock is NOT touched on scan — it is decremented only when the procedure is ended**, so mistakes/quantities can be fixed freely mid-procedure. "Add without barcode" opens a searchable picker showing all matches (never auto-picks the first). On End, every cart line decrements stock (move) and records a costed procedure line. Staff/common can start/end procedures but see NO costs, no running total, and cannot download the case-sheet PDF (admin only). History shows cost-by-surgeon (admin only) and per-procedure case-sheet PDF (admin only). Backed by Procedures + ProcedureLines tabs. Requires unit costs on items for meaningful totals.
- **Sticker printer** (admin only): 21-per-A4 label sheets (3×7, 64.5×40mm, left 14 / right 7 / top-bottom 4mm margins). Two modes — (a) patient stickers (name, DOB, PAT no., address, surgeon; details clear after print) and (b) barcode/QR stickers for meds/consumables/garments/linen with easy multiselect (search, location filter, category filter, tick-all). Adjustable margins with live trial-and-error, saveable margin preset (Settings key `sticker_margins`, JSON), ink colour picker, and per-sheet PDF download. Uses jsPDF.
- **Gas room daily checks:** one walk-through — oxygen manifold (left/right bank 3+3, ~142 full, pipeline ~4.2, in-use bank picker, delivery-booked tick), medical air manifold (weekly Mondays, ~108, pipeline ~7.3), helium cylinder counts, trolley O₂ counts + next delivery, vacuum pump duty/status, theatre 1&2 airflow off. Visual manifold panels with digital readouts + LEDs. Monthly consumption estimate for big cylinders.
- **Branding:** gold+teal Bollin logo; PWA manifest; premium Inter-based gold/teal navigation with line icons.

## Sheet tabs (created/maintained by migrate)
Consumables, Meds, Garments, Instruments, Linen, Transactions, Requests, BarcodeLinks, Alerts, Stocktakes, Settings, Users, Implants, DispatchLog, Stores, Assets, GasChecks, Procedures, ProcedureLines.

## Deployment specifics
- Custom domain: `bollin.hashirhub.uk` via Cloudflare CNAME (bollin → yasarshah65-sketch.github.io, DNS-only/grey cloud) + GitHub Pages custom domain + Enforce HTTPS. Staging uses plain github.io, no Cloudflare.
- Switching domains means everyone re-logs in (sessions are per-origin). Backend needs no change on domain switch.

## Recent fixes worth knowing
- **Procedure admin controls:** history rows (admin only) now have **Edit** — reopens a closed procedure: `procReopen_` first RETURNS all its consumed stock (Code→Barcode→exact-Name resolution, blank-code meds safe), deletes its ProcedureLines, rebuilds the cart into CartJSON and sets Status=Open, so re-ending consumes exactly once — and **Delete** (double-confirm) — `procDelete_` restores consumed stock, removes lines + the procedure row, fully logged. In the live console, admin has **✎ Edit details** (procUpdateMeta_: surgeon/procedure/date/patient ref — fixes staff typing patient names in the surgeon box) and every cart row shows an **In stock** column (live on-hand qty, all roles).
- **Cancel procedure:** ✕ top-right of the active console (all roles) with a confirm warning — discards the cart and deletes the Open row via `procCancel_` (refuses non-Open; local-only PROC-LOCAL ids just discarded). Nothing was consumed, so no stock changes.
- **Sterilisation on-hand:** rows are now clickable → detail dialog (`ohDlg`) with location/cycles/days-left; **admin can edit** name/code/barcode/type/tray-count/expiry (writes to Instruments via updateItem with original identifiers — duplicate-safe); other roles view-only. **Direct-add of already-sterile items** stays via "＋ Add instrument" (admin+staff): user enters the expiry there, while the normal used→dispatch→receive flow keeps its automatic expiry (dispatch + 1 year).
- **Procedure admin controls:** in history (admin only) each closed procedure has **Edit** and **Delete**. Edit = `procReopen`: returns all its consumed stock, deletes its ProcedureLines, rebuilds the cart from them (CartJSON), sets Status=Open — so re-ending re-consumes with corrections and can never double-decrement. Delete = `procDelete`: returns consumed stock, removes lines + the procedure row. Both log to the activity trail and report any lines that couldn't be restored. In an open procedure, admin has **✎ Edit details** (surgeon/procedure/date/patient ref → `procUpdateMeta`) for the staff-typed-patient-in-surgeon mistakes, and the cart shows a live **In stock** column (all roles).
- **Cancel procedure:** ✕ top-right of the active panel (all roles) with an explicit warning; discards the cart and deletes the Open row via `procCancel` (PROC-LOCAL ids are discarded locally). Nothing is consumed pre-end, so no stock changes.
- **Sterilisation on-hand:** rows are now clickable → detail dialog (where, cycles, days of sterility). **Admin-only editing** of name/code/barcode/type/tray-count/expiry via updateItem (fixes wrongly-logged items). The **“＋ Add already-sterile item”** button (admin+staff) direct-registers items onto the shelf with a user-entered expiry; the normal used→dispatch→receive flow keeps its automatic expiry. (data-loss + duplicates — IMPORTANT)
- **Procedure data loss on refresh (fixed — was the #1 financial risk):** the procedure cart lives in memory on `activeProc._cart`. Previously any `loadAll()` during a case rebuilt `state.procedures` from the sheet and could drop the cart. `loadAll` now captures the live cart first and re-attaches it to the matching procedure afterwards, so a refresh/sync never loses scanned items.
- **Procedure temporary save + autosave (security net):** the console has a **💾 Save progress** button and also autosaves ~1.2s after each change. The cart is persisted to a `CartJSON` column on the Procedures row (no stock touched). On reload, an Open procedure restores its saved cart. Stock is only decremented on **End procedure**.
- **Atomic End procedure:** ending a procedure now calls ONE server endpoint `procConsumeBatch` that decrements stock for every line, records every ProcedureLine, and closes the procedure in a single request — so a client refresh can't interrupt it half-way. If any line can't be consumed (e.g. item deleted), it reports which ones and keeps going; nothing is silently lost.
- **Blank-code items (e.g. "Bupivacaine with Adrenaline 0.25%") now consume correctly:** meds usually have empty Code AND Barcode. The batch endpoint resolves items by Code → Barcode → exact Name, so name-only items are found and decremented. `procConsume`/`procConsumeBatch` also look up UnitCost server-side (correct even for scrub/common logins whose costs are stripped client-side).
- **Edit-creates-duplicate (fixed at source):** `updateItem_` never appends (errors instead). As a hard backstop, `addItem_` now checks for an existing row (same non-blank Code, or Barcode, or — when blank — same Name) and converts an accidental add into an update, so duplicates are impossible even if the client mis-calls addItem.
- **Sticker printer now available to staff and common** (patient stickers + barcode/QR). Saving the shared margin preset stays admin-only.
- **Activity log PDF now respects on-screen filters:** the Download PDF in the activity log applies the activity / person / tracker / search filters (shown above the button) on top of the chosen date range — not just the date range. Both failing PDFs (activity log + procedure case sheet) were hardened to check the autoTable plugin loaded and to surface the real error as a toast instead of failing silently.
- **Sticker barcode/QR tab now has a Location filter** (tracker + category + location), and "tick all shown" respects all three.
- **Edit-creates-duplicate bug (fixed, esp. meds):** `updateItem_` resolves the row with a strict priority — (1) trust the supplied sheet row only if its current Code/Barcode/(or Name when no code exists) still matches the item being edited; (2) else re-find by original Code → Barcode → Name; (3) else return an error and NEVER append. This fixes meds specifically, which often have blank Code and Barcode and so previously fell through to a name search that could hit the wrong row or fail. The frontend refuses to save an edit until the item has a real sheet row (`_row`), and passes original identifiers. (Apps Script Web Apps do not sleep — this was a row-matching bug, not a timing one.)

## Known limitations / notes
- Passwords are stored hashed by design (shared sheet). Recovery = admin sets a new plain password.
- No "date added" on items (sheets don't record creation time) — would need a DateAdded column if wanted.
- Netum 1D scanners can't read QR codes (barcodes only).

## Style / working preferences
- Yasar prefers numbered multi-part update requests; expects each update fully implemented in one go without clarifying questions.
- Every update: give staging + production index + Code.gs + refreshed START_HERE, summarise per numbered item, end with deploy steps (staging-first) and a concrete first test.
- Files delivered to `/mnt/user-data/outputs/`.
