# Bizactivity Sheet Protection Setup

One-time setup to prevent accidental edits to the Job Reports sheet in
`bizactivity.xls`. Do this once per workbook. After setup, the launcher
and file watcher will automatically unlock, write, and re-lock the sheet
during normal sync operations — you won't notice anything different in
how they work.

## Why this matters

The Job Reports sheet is automatically kept in sync with the Whole Job
Docs Map sheets. If you accidentally type over a synced cell (like a
client name or gross sales figure), the next sync silently overwrites
your edit and your change is lost. Sheet protection blocks those
accidental typos while leaving the columns you actively use (the "P"
dropdowns) fully editable.

## Before you start

- Close `bizactivity.xls` if it's open in Excel.
- Install the latest `NewOrderLauncher.zip` release (the watcher needs
  to know the protection password). The launcher handles this
  automatically if you already have it running.
- Open `bizactivity.xls` fresh.

## Step 1 — Unlock the companion columns

These are the columns you need to keep editable: **H, Q, S, U, W, Y, AA,
AD, AF, AH**. You'll unlock them across all 12 month sections in one
operation.

1. Click the **Job Reports** sheet tab at the bottom.
2. Click on cell **H1** (anywhere in column H — the exact cell doesn't
   matter).
3. Hold **Ctrl** and click each of the following column letters at the
   top of the sheet to add them to your selection:
   **Q, S, U, W, Y, AA, AD, AF, AH**.
   You should now have 10 whole columns highlighted.
4. Right-click anywhere in the highlighted selection and choose
   **Format Cells…**
5. Click the **Protection** tab.
6. **Uncheck** the box labeled **Locked**.
7. Click **OK**.

Nothing visible happens — Excel just marks those cells as "unlocked"
internally. The lock only takes effect in the next step.

## Step 2 — Protect the sheet

1. Click the **Review** tab in the Excel ribbon.
2. Click **Protect Sheet**.
3. In the password box, type exactly:
   ```
   password
   ```
   (lowercase, no quotes)
4. Leave the checkbox list at its defaults — **"Select locked cells"**
   and **"Select unlocked cells"** should both stay checked. Don't
   enable any of the other options.
5. Click **OK**.
6. Excel asks you to re-enter the password. Type `password` again and
   click **OK**.

## Step 3 — Save the workbook

Press **Ctrl+S**. If Excel asks about the file format, keep the existing
`.xls` format — **don't** convert to `.xlsx`.

## Step 4 — Verify it worked

1. Try clicking into cell **B13** (a client name cell on the January
   section). Try to type something. Excel should block you with a
   message like *"The cell or chart you're trying to change is on a
   protected sheet."* That's what we want.
2. Click into **Q13** and try to open the dropdown. It should still
   work normally — the "P" option is selectable.
3. Click into **H13** and try to type a note. It should work.

If all three behave correctly, you're done.

## If you need to turn protection off

Go to **Review → Unprotect Sheet**, enter `password`, and click OK. The
sheet is now fully editable again. To re-enable, repeat Step 2.

## What the launcher does automatically

- Every time the launcher or watcher writes to Job Reports, it
  temporarily unprotects the sheet, writes its changes, and reprotects
  it before saving. You don't need to do anything.
- If `bizactivity.xls` is open when a sync tries to run, the watcher
  holds the change in a queue and retries once you close the workbook.
  You'll see a note in the log file (not in Excel itself).

## Troubleshooting

**"I can't edit a P dropdown anymore."**
You probably missed one of the columns in Step 1. Unprotect the sheet
(Review → Unprotect Sheet → `password`), then repeat Step 1 making sure
all 10 columns are selected before you open Format Cells. Reprotect
when done.

**"Excel is asking me for a password when I open the file."**
That's a different kind of password (workbook-open password). We only
use sheet protection, not workbook encryption. If you see this prompt,
call for help — something else got set up.

**"The launcher says it can't write to bizactivity."**
Check whether the workbook is open in Excel on your machine. If yes,
close it and wait ~30 seconds — the watcher will drain any pending
updates automatically. If the workbook is closed and the error
persists, check the log file under
`%LOCALAPPDATA%\UneedTShirtsNewOrder\logs\`.
