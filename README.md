# offhours-vba-macros

Standalone **VBA macros** for classifying Outlook emails by working hours and
categorizing off-hours support demand. There is no build system and no
dependencies — each `.bas` file is self-contained.

## Usage

1. Open Outlook or Excel.
2. `Alt+F11` to open the VBA editor.
3. `File → Import File…` and pick a `.bas` from `scripts/`.
4. Place the cursor inside the entry macro and press `F5` to run.

All Outlook macros are **read-only on the mailbox** by structure (no `Save`,
`Move`, `Delete`, `Send`, category or read-state changes). Output is written to
timestamped `.xlsx` files.

## Workflow / run order

The five `.bas` files have **no dependencies on each other** — each module is
self-contained, so there is no required *import* order (only one module of a
given name can live in a VBA project at a time). What matters is the **workflow
order**, which splits by host (Outlook vs Excel) and by purpose.

**Phase 1 — Timing scan (Outlook), pick ONE rule variant**

1. `find-offhours-emails.bas` → `FindOffHoursEmails` — the canonical
   **Paris-only** business-hours scan. Start here. (Run `ListMailboxes` first to
   confirm the mailbox display name.)
2. *(optional alternative)* `find-offhours-emails-india-counts-as-inhours.bas` →
   `FindOffHoursEmailsIndiaInHours` — same scan but the India IST shift counts as
   in-hours. Run instead of / alongside #1 to compare; its output is prefixed
   `offhours-emails-india` so it won't collide.

**Phase 1b — Excel-side re-timing (optional)**

3. `add-timing-columns.bas` → `AddTimingColumns` — run **in Excel** to apply the
   Paris-only rules to an already-exported `.xlsx`. Only needed if a sheet lacks
   the timing columns; skip if you used #1/#2.

**Phase 2 — Off-hours categorization (Outlook → Excel + local LLM)**

Run **both** versions on the same window and compare:

4. **v1** — `extract-emails-for-llm.bas` → `EmailsForLLM` (one row per message,
   single freeform LLM pass).
5. **v2** — `extract-threads-for-llm.bas` → `ExtractThreadsForLLM` (one row per
   conversation, deterministic triage in VBA + a two-pass LLM flow whose totals
   are counted in Excel).

**Before running the Phase-2 extractors**, populate the `Private Const` config at
the top of the module or the deterministic half won't fire:

- `SHARED_MAILBOX`, `FOLDER_NAME`, scan window, `OUTPUT_DIR`
- team rosters `APS` / `ADM` / `DEV_ADDRESSES_CSV` (start empty)
- v2 also needs `SUPPORT_ADDRESSES_CSV` + `AUTOMATED_SENDERS_CSV` (drives the
  Inbound/Outbound split). After the run, check the `Stats` sheet — if Outbound
  is 0, the support list is wrong.

**Short version:** `find-offhours-emails` (Paris) → optionally the India variant
or `add-timing-columns` → then the categorization pair `extract-emails-for-llm`
(v1) and `extract-threads-for-llm` (v2).

## Scripts

| File | Host | Entry macro | What it does |
|------|------|-------------|--------------|
| `find-offhours-emails.bas` | Outlook | `FindOffHoursEmails` | Scans a folder, writes a 3-sheet XLSX (`Emails`/`Stats`/`Users`). Business hours = **Paris 09:00–18:00**; India is a holiday-afternoon backstop. Metadata only — no body. |
| `find-offhours-emails-india-counts-as-inhours.bas` | Outlook | `FindOffHoursEmailsIndiaInHours` | Same scan, but business hours = **Paris OR India IST shift** on the clock. Adds `HourIST`/`IsParisHours`/`IsIndiaHours` columns. Metadata only. |
| `add-timing-columns.bas` | Excel | `AddTimingColumns` | Applies the Paris-only timing rules to any open `.xlsx`, appending classification columns from a chosen date column. |
| `extract-emails-for-llm.bas` | Outlook | `EmailsForLLM` | **v1** export, one row per message (Inbox + Sent). Exports cleaned bodies for local-LLM categorization. |
| `extract-threads-for-llm.bas` | Outlook | `ExtractThreadsForLLM` | **v2** export, one row per conversation (Inbox-only, direction-by-sender). Deterministic dedup/triage in VBA, condensed thread text for a two-pass LLM flow. |

## The off-hours rule

`IsOffHours = TRUE` when ANY of:

- Saturday or Sunday, OR
- time outside 09:00–18:00 Paris local, OR
- the date is a **Polish** holiday AND (it is also a Mumbai holiday OR the time
  is past India coverage end).

**Poland is the gating calendar.** France and Mumbai holidays are tracked as
columns for visibility but do not by themselves make an email off-hours.

## Configuration

All config lives in `Private Const` blocks at the top of each module: work
hours, holiday CSV constants (Poland/France/Mumbai), shared mailbox, folder
name, scan window, and output directory. Edit a holiday → mirror it across all
five files (each module is self-contained, no shared include in VBA).
