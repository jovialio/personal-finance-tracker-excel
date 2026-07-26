# Changelog

All notable changes to the Personal Finance Tracker workbook. The latest version is always available from the [releases page](https://github.com/jovialio/personal-finance-tracker-excel/releases/latest/download/Personal.Finance.Tracker.xlsx).

## v1.12

- **Cost basis resets on a full sell-and-rebuy.** On **Shares**, `Total Buy Cost` (and `Total Sell Proceeds`) now count only the transactions since a ticker's position last returned to zero — so the average cost of a re-acquired holding reflects what you paid *this* time round, not a blend with a previously closed position. Overall P&L on Shares follows suit (current-cycle), while **All Ticker P&L** keeps the lifetime figure. Driven by a new `Running Units` helper column on Transactions, built with plain `SUMPRODUCT`/`LOOKUP` (no `LET`) so it works in Excel and LibreOffice. Shares also gains an **Average Buy Cost** column showing the average price paid per unit in the current cycle — buy cost ÷ units *bought*, so a partial sell no longer distorts it. The reset formulas and the Transactions helper columns now span the full 10,000-transaction capacity (previously capped at 1,000). For a rebought ticker, enter its transactions in date order so the reset lands at the right point.

## v1.11

- **Start Here clarity pass.** Corrected the FX guidance (you set FX rates on Settings; the Prices FX column fills itself in), scoped the capacity note to what cell R1 actually shows, split the dense Transactions paste caution onto its own line, and removed a duplicated panel-capacity note. The in-sheet version-history list was also dropped to keep the tab uncluttered.

## v1.10

- **Plain paste guidance instead of a warning banner.** v1.9's `R1` self-check is reverted to the simple capacity indicator. In its place, Start Here now explains directly: **paste new transactions only into columns A–I** (never over the J–P helper formulas), and if J–P ever go blank, select a filled row and drag them down. Short pointers were added on the Transactions sheet and beside the Prices TICKER SYNC panel, which also gained a plain-language 'how to use this panel' note.

## v1.9

- **Transactions now catches the paste trap.** The helper columns (J–P: Key, units, cost, flow) are derived formulas the rest of the workbook depends on. Pasting a data block *over* them blanks them, which silently hides those transactions from the Dashboard, Shares and Prices. Cell `R1` on Transactions now shows a **red warning** — with the exact count of affected rows — whenever any row has an Exchange/Symbol/Currency but no Key, so the failure is impossible to miss. Start Here now says plainly: **paste only into columns A–I.**

## v1.8

- **Version numbering aligned.** The Start Here change log now uses the public release numbers (v1.x) that match the GitHub releases and the workbook filename, instead of a parallel internal build count. The previously-missing v1.2 entry is restored here and on Start Here.

## v1.7

- **A single CAPACITY & LIMITS section on Start Here.** The workbook's real limits — Excel's 1,048,576-row grid, the 10,000-trade XIRR ceiling and its live `R1` capacity banner, and the ~500-holding Prices panel — are now documented in one place, with the one rule for spotting a cap before it misleads you, instead of being scattered through the change notes.

## v1.6

- **The cash-flow engine moved into the ledger table.** `FX to SGD`, `Flow SGD` and `Dividend SGD` are now calculated columns on **Transactions**. Table columns grow with your data, so dividends, net-invested and every `SUMIFS` that used to read a fixed 1,000-row block are no longer capped at all.
- **FX is resolved from each transaction's own Currency**, not by looking its ticker up in Prices. Previously, deleting a sold-out ticker from Prices silently dropped the currency conversion on that ticker's past transactions — quietly skewing XIRR for anyone holding foreign-currency positions.
- **Capacity is now visible.** XIRR still needs one contiguous block (Excel won't accept a union reference there), so it keeps a mirror — raised to 10,000 transactions, with a banner in `Transactions!R1` reporting usage and warning loudly rather than failing silently.

## v1.5

- **Portfolio Analysis is now immune to inserted rows.** It previously read the ledger with 6,000 positional references (`Transactions!J2`, `J3`, …). Inserting rows into Transactions shifted those references, so the inserted rows were skipped entirely — the XIRR and the Stock Dashboard's return figure were then computed from only the rows that happened to still line up. Each row is now resolved with a position-independent `INDEX(Transactions!$J:$J, ROW()-8)` lookup, so typing, pasting **or inserting** rows all work correctly.

## v1.4

- **Pasting a batch of transactions now works.** The four helper columns on **Transactions** (Key, Signed Units, Buy Cost, Sell Proceeds) are pre-filled with formulas down to row 1000. Previously they existed only for the 20 sample rows, so pasting a block of transactions left **Key** blank on every pasted row — and with no Key those transactions were invisible to All Ticker P&L, Shares, the Stock Dashboard and the Prices sync panel. Blank rows now show nothing instead of `0`.

## v1.3

- **Prices "Ticker Sync" panel** (Level 3) — a live checklist beside the price list that keeps your prices in step with what you actually hold. It reads your current positions and flags, in green/red, any ticker you've bought that's missing from Prices (**ADD TO PRICES**) and any price row you've fully sold out of (**NOT HELD**), with **Missing** / **Stale** counters at the top. You still type prices, notes and categories yourself — the panel never edits your data, it just points you to the one row to add or remove. The stale sample tickers were also removed so the price list matches the sample holdings exactly.
- **Consistent input highlighting on Prices** — every cell you fill in (Key, Exchange, Symbol, Currency, Current Price, Notes, Category) now carries the same faint-yellow input shading, so it is obvious at a glance what is yours to type. Only the calculated **FX to SGD** column is left unshaded.

## v1.2

- **Release automation & CI guardrails.** Tagged `v*` releases now build and publish the workbook assets automatically, and every change is checked by CI (secret scan, workbook filename and package integrity, README links, and release-asset preparation). No changes to the workbook itself.

## v1.1

- **Savings Goal Calculator** (Level 2) — a goal-based compound-growth planner. Solve for any one unknown: the final amount you'll reach, the return you'd need, how many years to save, or the lump sum / annual amount to set aside for a goal. Contributions compound at the start of each year (annuity-due), and the expected-return cells pull from your Settings growth assumptions, with a one-cell override for quick what-ifs. Built entirely from Excel's own `FV` / `RATE` / `NPER` / `PMT` / `PV` functions — no macros, no add-ins.
