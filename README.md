# FP&A Variance Report Generator

Automates a recurring finance reporting workflow: validates actual-vs-budget
data, calculates variance, flags data-quality exceptions, and produces a
management-ready Excel workbook + PDF summary — from one command.

## 1. Purpose

Monthly variance reporting (actual vs budget) is a recurring, manual task in
most FP&A teams — pulling data, checking it for errors, calculating
variances, writing up what changed. This script does the mechanical parts
end to end and leaves only genuine judgment calls (flagged exceptions) for a
human to review.

## 2. Input

`sample_input.xlsx` — 5 sheets:

| Sheet | Columns |
|---|---|
| Actuals | month, cost_center, department, account_category, actual_amount |
| Budget | month, cost_center, department, account_category, budget_amount |
| Prior_Month | month, cost_center, department, account_category, prior_month_amount |
| Department | cost_center, department |
| Control_Totals | month, control_total_actual |

## 3. Schema validation (runs before anything else)

Before any data check runs, the script confirms the workbook matches the
required template:

- All 5 sheets present (clear error naming which is missing, if not)
- All required columns present in each sheet

Column matching is **case-insensitive and ignores leading/trailing spaces**
— so `Cost_Center`, `cost_center `, and `COST_CENTER` all match correctly.
Anything beyond that (a genuinely missing or misspelled column) fails
immediately with an error naming the exact sheet, the missing column, and
the columns that were actually found — never silently guessed. This is a
deliberate choice: in a finance tool, a script that stops with a clear
error is safer than one that silently misreads a column.

## 4. Validation rules (6 checks)

| Check | Catches | Action |
|---|---|---|
| Missing cost center | Blank cost_center | Excluded from clean dataset, logged |
| Missing account category | Blank account_category | Excluded, logged |
| Duplicate records | Same month+cost_center+category repeated | Excluded, logged |
| Invalid dates | Unparseable month value | Excluded, logged |
| Unexpected negative value | Revenue < 0 | **Flagged only, not excluded** — a sign error needs human judgment |
| Control total mismatch | Monthly sum ≠ control total (±₹1 tolerance) | Logged — may be explained by other flagged rows, or genuinely unexplained |

Full list of every exception found, with reasons, is in
`Exceptions_Log` inside `validated_output.xlsx`.

## 5. Business rule — favorable vs unfavorable

Direction of "good" depends on category type:

- **Revenue**: favorable if actual > budget (more income than planned)
- **Cost categories**: favorable if actual < budget (spent less than planned)

This is the one judgment call baked into the logic — documented here because
it's the first thing worth being able to explain.

## 6. Pipeline flow

```
Raw Excel (5 sheets)
      |
      v
  Validate  ---->  Exceptions_Log (13 issues logged, not silently fixed)
      |
      v
Clean dataset
      |
      v
Variance calc (abs, %, fav/unfav flag, MoM change)
      |
      v
Top 5 adverse / favorable drivers
      |
      v
validated_output.xlsx  +  management_summary.pdf
```

## 7. Outputs

**`validated_output.xlsx`** — 5 sheets:
- `Variance_Detail` — full row-level variance table
- `Exceptions_Log` — every flagged issue, sheet/row/reason
- `Top5_Adverse` / `Top5_Favorable` — biggest swings by magnitude
- `Monthly_Summary` — totals and line counts per month

**`management_summary.pdf`** — one page, text generated from the computed
numbers (not hand-written), covering total variance, top drivers, and a data
quality note on how many rows were flagged.

## 8. Limits — what this does NOT do

- Flagged rows are **not** auto-corrected (e.g. a negative Revenue sign is
  flagged, never silently flipped) — control totals that don't tie are
  logged, not forced to match. Final sign-off is still a human decision.
- No forecasting — this is a backward-looking variance tool, not a
  predictive one.

## 9. How to run

```bash
pip install -r requirements.txt
python variance_report_generator.py sample_input.xlsx
```

Outputs `validated_output.xlsx` and `management_summary.pdf` in the same
folder. Point it at any new file the same way:

```bash
python variance_report_generator.py october_actuals.xlsx
```

## 10. Tech used

Python, Pandas, openpyxl (Excel read/write), reportlab (PDF generation).

## 11. Future enhancement — AI-generated commentary (not implemented)

Current `management_summary.pdf` text is built with Python f-string
templates — rule-based, fixed structure, numbers plugged in from the
variance calc.

Planned upgrade: replace the template with a call to the Claude API.

**Process:**
1. After variance calc, package key stats as JSON (top adverse/favorable
   driver, total variance, exception count, monthly trend)
2. Send JSON + a prompt to the Claude API asking for a 3-4 sentence
   management commentary
3. Insert the returned text into the PDF in place of the template paragraph
4. Keep the template as a fallback if the API call fails or no key is set —
   the script should never break because of a missing/invalid key

**Why not built yet:** adds an external dependency (API key, network call,
cost per run) that isn't needed to prove the core skill — data validation,
variance logic, and automated reporting. Kept as a documented next step
rather than built in, to keep the deliverable self-contained and free to
run.

## Sample output

**Running the script:**

![Terminal run](/images/terminal_run_preview.png)

**Variance_Detail sheet (inside validated_output.xlsx):**

![Excel output](/images/excel_variance_detail_preview.png)

**management_summary.pdf:**

![PDF summary](/images/management_summary_preview.png)
