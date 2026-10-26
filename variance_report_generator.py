"""
FP&A Variance Report Generator
================================
Validates actual-vs-budget finance data, calculates variance metrics,
flags data-quality exceptions, and produces an Excel workbook + a
one-page PDF management summary — fully automated from a single
Excel input file.

Usage:
    python variance_report_generator.py sample_input.xlsx

Input workbook must contain 5 sheets:
    Actuals, Budget, Prior_Month, Department, Control_Totals
(see README.md for the exact column spec)
"""

import sys
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

INCOME_CATEGORIES = {"Revenue"}

REQUIRED_SHEETS = {
    "Actuals": ["month", "cost_center", "department", "account_category", "actual_amount"],
    "Budget": ["month", "cost_center", "department", "account_category", "budget_amount"],
    "Prior_Month": ["month", "cost_center", "department", "account_category", "prior_month_amount"],
    "Department": ["cost_center", "department"],
    "Control_Totals": ["month", "control_total_actual"],
}


# ---------------------------------------------------------------------
# STAGE 0 — SCHEMA VALIDATION
# ---------------------------------------------------------------------
def normalize_columns(df, expected_cols, sheet_name):
    """
    Matches columns case-insensitively and ignoring leading/trailing spaces
    (e.g. "Cost_Center", " cost_center " both match "cost_center"), then
    renames them to the exact canonical name the rest of the script expects.

    Anything beyond trivial formatting (a genuinely missing or misspelled
    column) fails LOUDLY with a clear error — it is never silently guessed,
    since a wrong guess on financial data is worse than a script that stops.
    """
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    rename = {}
    missing = []

    for expected in expected_cols:
        key = expected.strip().lower()
        if key in norm_map:
            rename[norm_map[key]] = expected
        else:
            missing.append(expected)

    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing required column(s): {missing}. "
            f"Columns found in file: {list(df.columns)}. "
            f"Check spelling/sheet template against README.md section 2."
        )

    return df.rename(columns=rename)


def validate_schema(sheets_dict):
    """Checks all 5 required sheets exist, then normalizes columns in each."""
    missing_sheets = [s for s in REQUIRED_SHEETS if s not in sheets_dict]
    if missing_sheets:
        raise ValueError(
            f"Input workbook is missing required sheet(s): {missing_sheets}. "
            f"Sheets found: {list(sheets_dict.keys())}. "
            f"Required sheets: {list(REQUIRED_SHEETS.keys())}."
        )

    validated = {}
    for sheet_name, expected_cols in REQUIRED_SHEETS.items():
        validated[sheet_name] = normalize_columns(sheets_dict[sheet_name], expected_cols, sheet_name)
    return validated


# ---------------------------------------------------------------------
# STAGE 1 — LOAD
# ---------------------------------------------------------------------
def load_workbook(path):
    raw_sheets = pd.read_excel(path, sheet_name=None)
    sheets = validate_schema(raw_sheets)
    return (sheets["Actuals"], sheets["Budget"], sheets["Prior_Month"],
            sheets["Department"], sheets["Control_Totals"])


# ---------------------------------------------------------------------
# STAGE 2 — VALIDATE
# ---------------------------------------------------------------------
def validate(actuals, control):
    """Run all 6 data-quality checks. Returns (clean_actuals, exceptions_df)."""
    exceptions = []

    def flag(sheet, row_idx, issue, detail):
        exceptions.append({"sheet": sheet, "row": row_idx, "issue": issue, "detail": detail})

    # Check 1 — missing cost center
    missing_cc = actuals[actuals["cost_center"].isna() | (actuals["cost_center"] == "")]
    for idx in missing_cc.index:
        flag("Actuals", idx, "missing_cost_center", "cost_center blank")

    # Check 2 — missing account category
    missing_cat = actuals[actuals["account_category"].isna() | (actuals["account_category"] == "")]
    for idx in missing_cat.index:
        flag("Actuals", idx, "missing_account_category", "account_category blank")

    # Check 3 — duplicate records
    dup_mask = actuals.duplicated(subset=["month", "cost_center", "account_category"], keep="first")
    dupes = actuals[dup_mask]
    for idx in dupes.index:
        flag("Actuals", idx, "duplicate_record", "same month/cost_center/category seen earlier")

    # Check 4 — invalid dates
    parsed_dates = pd.to_datetime(actuals["month"], errors="coerce")
    bad_dates = actuals[parsed_dates.isna()]
    for idx in bad_dates.index:
        flag("Actuals", idx, "invalid_date", f"unparseable month value: {actuals.loc[idx, 'month']}")

    # Check 5 — unexpected negative values (Revenue should never be negative;
    # cost categories CAN legitimately be negative, e.g. a credit/refund)
    neg_revenue = actuals[(actuals["account_category"] == "Revenue") & (actuals["actual_amount"] < 0)]
    for idx in neg_revenue.index:
        flag("Actuals", idx, "unexpected_negative", f"Revenue negative: {actuals.loc[idx, 'actual_amount']}")

    # Build clean dataset — excludes rows with structural issues.
    # Negative-revenue rows are FLAGGED but NOT excluded: a sign error is a
    # business judgment call, not a structural data error, so it stays in
    # for a human to review rather than being silently dropped or auto-fixed.
    bad_idx = set(missing_cc.index) | set(missing_cat.index) | set(bad_dates.index) | set(dupes.index)
    clean_actuals = actuals.drop(index=bad_idx)

    # Check 6 — control total mismatch (computed on clean data only, so one
    # bad row can't mask or fake a match)
    monthly_actual = clean_actuals.groupby("month")["actual_amount"].sum().reset_index()
    monthly_actual.columns = ["month", "sum_of_actuals"]
    check = control.merge(monthly_actual, on="month", how="left")
    check["diff"] = (check["control_total_actual"] - check["sum_of_actuals"]).round(2)
    mismatches = check[check["diff"].abs() > 1]  # tolerance Rs.1 for rounding
    for _, r in mismatches.iterrows():
        flag("Control_Totals", r["month"], "control_total_mismatch",
             f"control={r['control_total_actual']}, sum_actuals={r['sum_of_actuals']}, diff={r['diff']}")

    exceptions_df = pd.DataFrame(exceptions)
    return clean_actuals, exceptions_df


# ---------------------------------------------------------------------
# STAGE 3 — VARIANCE CALCULATIONS
# ---------------------------------------------------------------------
def fav_flag(row):
    """Revenue: favorable if actual > budget. Cost: favorable if actual < budget."""
    if row["account_category"] in INCOME_CATEGORIES:
        return "Favorable" if row["abs_variance"] > 0 else "Unfavorable"
    return "Favorable" if row["abs_variance"] < 0 else "Unfavorable"


def calculate_variance(clean_actuals, budget, prior):
    keys = ["month", "cost_center", "account_category"]
    variance = clean_actuals.merge(budget[keys + ["budget_amount"]], on=keys, how="left")
    variance = variance.merge(prior[keys + ["prior_month_amount"]], on=keys, how="left")

    variance["abs_variance"] = (variance["actual_amount"] - variance["budget_amount"]).round(2)
    variance["pct_variance"] = (variance["abs_variance"] / variance["budget_amount"].replace(0, pd.NA)).round(4)
    variance["mom_change"] = (variance["actual_amount"] - variance["prior_month_amount"]).round(2)
    variance["mom_pct_change"] = (variance["mom_change"] / variance["prior_month_amount"].replace(0, pd.NA)).round(4)
    variance["variance_flag"] = variance.apply(fav_flag, axis=1)

    variance_final = variance[["month", "cost_center", "department", "account_category",
                                "actual_amount", "budget_amount", "abs_variance", "pct_variance",
                                "variance_flag", "prior_month_amount", "mom_change", "mom_pct_change"]]
    return variance_final


def top_drivers(variance_final, n=5):
    def top(flag_val):
        return (variance_final[variance_final["variance_flag"] == flag_val]
                .assign(magnitude=lambda d: d["abs_variance"].abs())
                .sort_values("magnitude", ascending=False)
                .head(n)[["month", "cost_center", "account_category",
                          "actual_amount", "budget_amount", "abs_variance", "pct_variance"]])
    return top("Unfavorable"), top("Favorable")


def monthly_summary_table(variance_final):
    return variance_final.groupby("month").agg(
        total_actual=("actual_amount", "sum"),
        total_budget=("budget_amount", "sum"),
        total_variance=("abs_variance", "sum"),
        unfavorable_count=("variance_flag", lambda s: (s == "Unfavorable").sum()),
        favorable_count=("variance_flag", lambda s: (s == "Favorable").sum()),
    ).round(2).reset_index()


# ---------------------------------------------------------------------
# STAGE 4 — OUTPUTS
# ---------------------------------------------------------------------
def write_excel_output(path, variance_final, exceptions_df, top5_adverse, top5_favorable, monthly_summary):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        variance_final.to_excel(writer, sheet_name="Variance_Detail", index=False)
        exceptions_df.to_excel(writer, sheet_name="Exceptions_Log", index=False)
        top5_adverse.to_excel(writer, sheet_name="Top5_Adverse", index=False)
        top5_favorable.to_excel(writer, sheet_name="Top5_Favorable", index=False)
        monthly_summary.to_excel(writer, sheet_name="Monthly_Summary", index=False)


def write_pdf_summary(path, variance_final, exceptions_df, top5_adverse, top5_favorable,
                       monthly_summary, total_input_rows):
    latest_month = variance_final["month"].max()
    latest = variance_final[variance_final["month"] == latest_month]
    total_actual = latest["actual_amount"].sum()
    total_budget = latest["budget_amount"].sum()
    total_variance = total_actual - total_budget
    total_variance_pct = total_variance / total_budget if total_budget else 0

    top_adv = top5_adverse.iloc[0]
    top_fav = top5_favorable.iloc[0]
    n_exceptions = len(exceptions_df)
    exception_rate = n_exceptions / total_input_rows if total_input_rows else 0

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontSize=9, textColor=colors.grey))

    doc = SimpleDocTemplate(path, pagesize=A4, topMargin=20 * mm, bottomMargin=20 * mm,
                             leftMargin=20 * mm, rightMargin=20 * mm)
    story = [
        Paragraph("Monthly Variance Report — Management Summary", styles["Title"]),
        Paragraph(f"Period: {latest_month}  |  Auto-generated from validated actuals", styles["Small"]),
        Spacer(1, 12),
        Paragraph(
            f"Total actuals for {latest_month} were Rs. {total_actual:,.0f} against a budget of "
            f"Rs. {total_budget:,.0f}, a variance of Rs. {total_variance:,.0f} ({total_variance_pct*100:.1f}%). "
            f"This total nets revenue and cost lines together — see Variance_Detail for the line-item "
            f"picture before drawing a conclusion.", styles["Normal"]),
        Spacer(1, 8),
        Paragraph("Key drivers", styles["Heading2"]),
        Paragraph(f"<b>Largest adverse driver:</b> {top_adv['cost_center']} / {top_adv['account_category']} "
                  f"({top_adv['month']}): variance {top_adv['abs_variance']:,.0f}", styles["Normal"]),
        Paragraph(f"<b>Largest favorable driver:</b> {top_fav['cost_center']} / {top_fav['account_category']} "
                  f"({top_fav['month']}): variance {top_fav['abs_variance']:,.0f}", styles["Normal"]),
        Spacer(1, 8),
        Paragraph("Data quality", styles["Heading2"]),
        Paragraph(f"{n_exceptions} of {total_input_rows} input rows ({exception_rate*100:.1f}%) were flagged "
                  f"during validation and excluded or held for review before this analysis was run. "
                  f"See Exceptions_Log for the full list and reasons.", styles["Normal"]),
        Spacer(1, 12),
        Paragraph("Monthly trend", styles["Heading2"]),
    ]

    table_data = [["Month", "Total Actual", "Total Budget", "Variance", "Unfav. lines", "Fav. lines"]]
    for _, r in monthly_summary.iterrows():
        table_data.append([r["month"], f"{r['total_actual']:,.0f}", f"{r['total_budget']:,.0f}",
                            f"{r['total_variance']:,.0f}", int(r["unfavorable_count"]), int(r["favorable_count"])])
    t = Table(table_data, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
    ]))
    story.append(t)
    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "This summary was generated by an automated validation + variance pipeline (Pandas). "
        "Exceptions require human review before final sign-off.", styles["Small"]))

    doc.build(story)


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------
def run(input_path, excel_out="validated_output.xlsx", pdf_out="management_summary.pdf"):
    actuals, budget, prior, department, control = load_workbook(input_path)
    clean_actuals, exceptions_df = validate(actuals, control)
    variance_final = calculate_variance(clean_actuals, budget, prior)
    top5_adverse, top5_favorable = top_drivers(variance_final)
    monthly_summary = monthly_summary_table(variance_final)

    write_excel_output(excel_out, variance_final, exceptions_df, top5_adverse, top5_favorable, monthly_summary)
    write_pdf_summary(pdf_out, variance_final, exceptions_df, top5_adverse, top5_favorable,
                       monthly_summary, total_input_rows=len(actuals))

    print(f"Input rows:        {len(actuals)}")
    print(f"Clean rows:        {len(clean_actuals)}")
    print(f"Exceptions found:  {len(exceptions_df)}")
    print(f"Excel written:     {excel_out}")
    print(f"PDF written:       {pdf_out}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python variance_report_generator.py <input_workbook.xlsx> "
              "[output_excel.xlsx] [output_pdf.pdf]")
        sys.exit(1)
    input_path = sys.argv[1]
    excel_out = sys.argv[2] if len(sys.argv) > 2 else "validated_output.xlsx"
    pdf_out = sys.argv[3] if len(sys.argv) > 3 else "management_summary.pdf"
    run(input_path, excel_out, pdf_out)
