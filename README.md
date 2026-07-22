# Personal Finance Tracker (Excel)

A free Excel template that grows with you — from basic financial planning to full portfolio management. Built around a three-level journey: know where you stand, plan your future, then run your investments properly. No macros, no add-ins, no internet access — just formulas you can inspect.

![Dashboard](images/dashboard.png)

## Why this template

Most finance spreadsheets are either a simple budget or an intimidating investment model. This one is organised as a **learning path**: the tabs are colour-coded by level, and each level unlocks a visible payoff on the Dashboard. Fill in the green tabs in ~15 minutes and you get your net worth, allocation and red/amber/green health checks. Work up to the purple tabs and you get a transaction-ledger-driven portfolio tracker with true money-weighted returns (XIRR), dividend tracking and a monthly performance history.

![Start Here guide](images/start-here.png)

## The three levels

| Level | Tabs | What you get |
|---|---|---|
| 🟩 **1 — Foundations** (~15 min) | Settings · Balance Sheet · Cash Flow · Net Worth History | Net worth, asset allocation, emergency-fund months, measured savings rate — with red/amber/green status checks |
| 🟦 **2 — Life Planning** | Insurance · Housing · Loan Calculator · Retirement Plan · Savings Goal Calculator | Retirement readiness %, your FI number, mortgage amortisation, CPF refund on sale, projections to age 85 at 3/5/8% growth, plus a goal-based savings calculator |
| 🟪 **3 — Portfolio Management** | Transactions · Prices · Stock Dashboard · Portfolio History | XIRR (true annual return), dividends, top holdings, concentration & data-quality checks, contribution-adjusted monthly returns |
| ⬜ **Engine room** | Shares · All Ticker P&L · Portfolio Analysis | Fully automatic — look, don't edit |

## Quick start

1. [Download the latest workbook](https://github.com/jovialio/personal-finance-tracker-excel/releases/latest/download/Personal.Finance.Tracker.xlsx) and open it in Excel (LibreOffice works too).
2. Read the **Start Here** tab — it is the manual.
3. Everything ships with a coherent fictional example (a 42-year-old investing since 2019) so every formula shows a working result. Follow the *Resetting the Sample Data* checklist on Start Here to make it yours.
4. Yellow cells are yours to edit; everything else is formulas. Formula-heavy sheets are protected against accidents — no password, `Review → Unprotect Sheet` if you ever need to.

### The monthly routine (~10 minutes)

1. Update **Prices** (prices & FX).
2. Add a row to **Cash Flow** (income & expenses).
3. Add a row to **Net Worth History**.
4. Copy `A5:B5` of the LIVE row on **Portfolio History** and paste-values below — this turns snapshots into a track record.

## The stock ledger

One transactions table drives the entire portfolio layer. Four action codes:

| Code | Meaning | Units | Total after fees |
|---|---|---|---|
| `1` | Buy | bought | cost incl. fees |
| `-1` | Sell | sold | proceeds after fees |
| `0` | Fee / charge | 0 | fee amount |
| `2` | Dividend received | 0 | amount received |

Tickers use a `Exchange:Symbol:Currency` key (e.g. `USX:AMD:USD`). Anything can be a ticker — the sample tracks physical gold as `CA:GOLD:SGD`, and you can do the same for CPF or bonds. The ledger handles multiple currencies (SGD/USD/HKD out of the box; add more in Settings).

![Stock Dashboard](images/stock-dashboard.png)

## What gets calculated for you

- **Portfolio XIRR** — money-weighted annual return since inception, dividends included, from your actual dated cash flows.
- **Retirement projection to age 85** — three growth scenarios; savings stop at your retirement age and inflation-adjusted withdrawals begin; CPF/SRS/endowment unlocks flow in at the ages you set. Your retirement-age row highlights automatically.
- **Retirement readiness %** — projected assets at retirement vs your FI number (target income ÷ safe withdrawal rate).
- **Housing** — full amortisation schedule, your equity share, and (Singapore) the CPF refund with accrued interest owed on sale, so you see true cash-in-hand.
- **Status checks** — emergency-fund months, holding concentration, actual-vs-planned savings rate, all red/amber/green.

![Retirement Plan](images/retirement-plan.png)

## Singapore context

Built with Singapore in mind — CPF (OA/SA/Medisave), SRS, SSB, HDB/BSD terms are used and explained on the Settings tab. Everything is a plain labelled input, so replacing these with your local equivalents (401k/IRA, ISA/SIPP, EPF…) takes minutes.

## Good to know

- All numbers and tickers in the file are **fictional sample data**.
- Once filled in, the file is a complete map of your finances — **keep your copy private**.
- Historic FX for the XIRR uses current rates (documented on the Portfolio Analysis tab).
- The ~44% "top holding concentration" red flag in the sample is the check working as intended — small portfolios concentrate easily.
- Not yet built (PRs welcome): benchmark comparison vs an index, YTD return, monthly return bar chart, allocation drift tracking.

## What's new in v1.10

- **Plain paste guidance instead of a warning banner.** v1.9's `R1` self-check is reverted to the simple capacity indicator. In its place, Start Here now explains directly: **paste new transactions only into columns A–I** (never over the J–P helper formulas), and if J–P ever go blank, select a filled row and drag them down. Short pointers were added on the Transactions sheet and beside the Prices TICKER SYNC panel, which also gained a plain-language 'how to use this panel' note.

## What's new in v1.9

- **Transactions now catches the paste trap.** The helper columns (J–P: Key, units, cost, flow) are derived formulas the rest of the workbook depends on. Pasting a data block *over* them blanks them, which silently hides those transactions from the Dashboard, Shares and Prices. Cell `R1` on Transactions now shows a **red warning** — with the exact count of affected rows — whenever any row has an Exchange/Symbol/Currency but no Key, so the failure is impossible to miss. Start Here now says plainly: **paste only into columns A–I.**

## What's new in v1.8

- **Version numbering aligned.** The Start Here change log now uses the public release numbers (v1.x) that match the GitHub releases and the workbook filename, instead of a parallel internal build count. The previously-missing v1.2 entry is restored here and on Start Here.

## What's new in v1.7

- **A single CAPACITY & LIMITS section on Start Here.** The workbook's real limits — Excel's 1,048,576-row grid, the 10,000-trade XIRR ceiling and its live `R1` capacity banner, and the ~500-holding Prices panel — are now documented in one place, with the one rule for spotting a cap before it misleads you, instead of being scattered through the change notes.

## What's new in v1.6

- **The cash-flow engine moved into the ledger table.** `FX to SGD`, `Flow SGD` and `Dividend SGD` are now calculated columns on **Transactions**. Table columns grow with your data, so dividends, net-invested and every `SUMIFS` that used to read a fixed 1,000-row block are no longer capped at all.
- **FX is resolved from each transaction's own Currency**, not by looking its ticker up in Prices. Previously, deleting a sold-out ticker from Prices silently dropped the currency conversion on that ticker's past transactions — quietly skewing XIRR for anyone holding foreign-currency positions.
- **Capacity is now visible.** XIRR still needs one contiguous block (Excel won't accept a union reference there), so it keeps a mirror — raised to 10,000 transactions, with a banner in `Transactions!R1` reporting usage and warning loudly rather than failing silently.

## What's new in v1.5

- **Portfolio Analysis is now immune to inserted rows.** It previously read the ledger with 6,000 positional references (`Transactions!J2`, `J3`, …). Inserting rows into Transactions shifted those references, so the inserted rows were skipped entirely — the XIRR and the Stock Dashboard's return figure were then computed from only the rows that happened to still line up. Each row is now resolved with a position-independent `INDEX(Transactions!$J:$J, ROW()-8)` lookup, so typing, pasting **or inserting** rows all work correctly.

## What's new in v1.4

- **Pasting a batch of transactions now works.** The four helper columns on **Transactions** (Key, Signed Units, Buy Cost, Sell Proceeds) are pre-filled with formulas down to row 1000. Previously they existed only for the 20 sample rows, so pasting a block of transactions left **Key** blank on every pasted row — and with no Key those transactions were invisible to All Ticker P&L, Shares, the Stock Dashboard and the Prices sync panel. Blank rows now show nothing instead of `0`.

## What's new in v1.3

- **Prices “Ticker Sync” panel** (Level 3) — a live checklist beside the price list that keeps your prices in step with what you actually hold. It reads your current positions and flags, in green/red, any ticker you’ve bought that’s missing from Prices (**ADD TO PRICES**) and any price row you’ve fully sold out of (**NOT HELD**), with **Missing** / **Stale** counters at the top. You still type prices, notes and categories yourself — the panel never edits your data, it just points you to the one row to add or remove. The stale sample tickers were also removed so the price list matches the sample holdings exactly.

- **Consistent input highlighting on Prices** — every cell you fill in (Key, Exchange, Symbol, Currency, Current Price, Notes, Category) now carries the same faint-yellow input shading, so it is obvious at a glance what is yours to type. Only the calculated **FX to SGD** column is left unshaded.

## What's new in v1.2

- **Release automation & CI guardrails.** Tagged `v*` releases now build and publish the workbook assets automatically, and every change is checked by CI (secret scan, workbook filename and package integrity, README links, and release-asset preparation). No changes to the workbook itself.

## What's new in v1.1

- **Savings Goal Calculator** (Level 2) — a goal-based compound-growth planner. Solve for any one unknown: the final amount you'll reach, the return you'd need, how many years to save, or the lump sum / annual amount to set aside for a goal. Contributions compound at the start of each year (annuity-due), and the expected-return cells pull from your Settings growth assumptions, with a one-cell override for quick what-ifs. Built entirely from Excel's own `FV` / `RATE` / `NPER` / `PMT` / `PV` functions — no macros, no add-ins.

## Maintainer release flow

Future releases are automated from final-version `v*` tags such as `v1.2` or `v1.2.3`. Prerelease suffixes such as `v1.2-rc1` are intentionally rejected so the stable latest-download URL never points at an RC or beta workbook.

For a new version:

1. Rename the workbook in this repo to match the tag, for example `Personal Finance Tracker v1.2.xlsx`.
2. Commit the workbook and README changes to `main`.
3. Tag the release commit, for example `git tag v1.2 && git push origin main v1.2`.

The release workflow uploads two assets:

- `Personal.Finance.Tracker.v1.2.xlsx` — the version-pinned workbook.
- `Personal.Finance.Tracker.xlsx` — the stable latest-download alias used by README and blog links.

The stable download URL should not change between releases:

```text
https://github.com/jovialio/personal-finance-tracker-excel/releases/latest/download/Personal.Finance.Tracker.xlsx
```

Pull request CI runs a secret scan and validates the workbook filename, XLSX package integrity, README assets, stable download link, workflow YAML, and release asset preparation before merge.

## About

I'm Dennis — I built this to answer my own money questions, then kept refining it until it was worth sharing. I write about the thinking behind it, and what I'm building next, at [blog.synvest.life](https://blog.synvest.life).

## License

[MIT](LICENSE) — free to use, modify and share, including commercially, with attribution.

**Disclaimer:** this template is for personal organisation and education only. It is not financial advice, and its calculations are simplified models — verify anything important with a qualified professional.
