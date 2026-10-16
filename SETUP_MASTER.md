# Building your MASTER_DASHBOARD — one-time setup

**The split:** `RUN.command` refreshes `data.xlsx` (clean tables, no charts).
Your `MASTER_DASHBOARD.xlsx` holds your charts/layout and pulls from `data.xlsx`.
The code never touches your master, so your layout is yours forever.

`data.xlsx` tabs (each a named Excel Table): `KPIs`, `GreenLeaf`, `Auction`,
`Production`, `Expenses`, `Economics`, `YoY_Expenses`.

---

## Recommended on Mac: Power Query + refresh (reliable)

1. Run `RUN.command` once so `data.xlsx` exists next to it.
2. New workbook → save as **MASTER_DASHBOARD.xlsx** in the **same folder**.
3. **Data → Get Data → From File → From Workbook** → pick `data.xlsx`.
4. In the navigator, tick the tables you want (tick all, then "Load To… →
   Only Create Connection + Add to Data Model" is fine, or load each to a sheet).
5. Build your charts off those loaded tables (PivotCharts or normal charts).
6. **Make refresh automatic:** Data → Queries & Connections → right-click a query →
   **Properties → tick "Refresh data when opening the file."**

**To update from then on:** double-click `RUN.command` → open the master (it
refreshes on open). Or hit **Data → Refresh All** — that's Excel's built-in
one-click refresh button, and it's reliable on Mac.

---

## The custom "Update" button (optional — read the Mac caveat)

A custom on-sheet button running a VBA macro is what you pictured. It works, but
on **Mac** two things get flaky: VBA-triggered Power Query refresh often silently
does nothing, and VBA opening a file by path can hit Mac's file-access sandbox.
So on Mac I recommend the Refresh-All button above. If you still want a custom
button, the macro in `docs/UpdateMacro.bas` does a **direct pull** (reads
`data.xlsx` and drops each table into sheets named `data_<TabName>` in your
master — no Power Query). Save your master as **.xlsm**, paste the macro into the
VBA editor, assign it to a button. First run on Mac may ask you to grant file
access — allow it.

---

## Updating the LOOK vs the DATA

- **New data** (daily/weekly/monthly): drop the file in its folder → `RUN.command`
  → refresh the master. No code, no git.
- **New metric / new column in the data**: tell Claude — that's a code change to
  the pipe; Claude pushes it, you `git pull`.
- **Chart/layout tweaks**: all yours, in the master. Never overwritten.
