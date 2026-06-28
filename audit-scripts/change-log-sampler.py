"""
Change Log Sampler
==================
IT Audit Tool — GRC Toolkit
Author: Sandesh | ULM Accounting & CIS

Purpose:
    Ingests a change management log CSV, selects a statistically appropriate
    audit sample per AICPA/PCAOB guidance, pre-flags exceptions, and outputs
    a formatted Excel workpaper ready for audit testing.

Sampling Logic (AICPA guidance):
    Population 1–24    → Test all
    Population 25–99   → Sample 25
    Population 100+    → Sample 40
    Emergency changes  → Always include all (100% testing)

Input CSV columns (required):
    change_id, description, category, requested_by, approved_by,
    implemented_by, request_date, approval_date, implementation_date,
    status, is_emergency, tested

Usage:
    python change-log-sampler.py --input change_log.csv --output change_log_sample.xlsx
    python change-log-sampler.py  (uses sample data if no file provided)
    python change-log-sampler.py --sample-size 25 --seed 42
"""

import pandas as pd
import argparse
import random
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings("ignore")

REQUIRED_COLUMNS = [
    "change_id", "description", "category", "requested_by", "approved_by",
    "implemented_by", "request_date", "approval_date", "implementation_date",
    "status", "is_emergency", "tested"
]

# AICPA sample size thresholds
SAMPLE_THRESHOLDS = [
    (0,   24,  None),   # Test all
    (25,  99,  25),     # Sample 25
    (100, 999, 40),     # Sample 40
]


# ─────────────────────────────────────────────
# INPUT VALIDATION
# ─────────────────────────────────────────────
def validate_input(df):
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Input CSV is missing required column(s): {', '.join(missing)}\n"
            f"Required columns: {', '.join(REQUIRED_COLUMNS)}"
        )
    for col in ["change_id", "requested_by", "approved_by", "implemented_by", "status"]:
        null_count = df[col].isna().sum()
        if null_count > 0:
            print(f"   ⚠️  Warning: {null_count} null value(s) in '{col}' — these rows may produce exceptions.")
    return True


# ─────────────────────────────────────────────
# SAMPLE DATA GENERATOR
# ─────────────────────────────────────────────
def generate_sample_data():
    """Generate a realistic change log with seeded anomalies."""
    today = datetime.today()
    categories = ["Normal", "Standard", "Emergency", "Routine"]
    systems    = ["ERP", "Active Directory", "Financial System", "Payroll", "Network", "Database"]
    users      = ["john.smith", "sara.jones", "mike.chen", "lisa.patel", "tom.brown",
                  "anna.kim",  "raj.patel",  "emily.wu",  "dan.garcia", "kate.moore"]

    random.seed(99)
    data = []

    for i in range(1, 76):  # 75 changes → triggers 40-sample rule
        req_date  = today - timedelta(days=random.randint(10, 365))
        app_date  = req_date  + timedelta(days=random.randint(0, 3))
        impl_date = app_date  + timedelta(days=random.randint(1, 7))

        requester   = random.choice(users)
        approver    = random.choice([u for u in users if u != requester])
        implementer = random.choice([u for u in users if u != requester])

        is_emergency = "Yes" if i % 15 == 0 else "No"
        category     = "Emergency" if is_emergency == "Yes" else random.choice(["Normal", "Standard", "Routine"])

        data.append({
            "change_id":             f"CHG-{i:04d}",
            "description":           f"Update to {random.choice(systems)} — {category.lower()} change #{i}",
            "category":              category,
            "requested_by":          requester,
            "approved_by":           approver,
            "implemented_by":        implementer,
            "request_date":          req_date.strftime("%Y-%m-%d"),
            "approval_date":         app_date.strftime("%Y-%m-%d"),
            "implementation_date":   impl_date.strftime("%Y-%m-%d"),
            "status":                "Completed",
            "is_emergency":          is_emergency,
            "tested":                random.choice(["Yes", "Yes", "Yes", "No"]),
        })

    # ── Inject specific anomalies ──

    # Missing approval
    data[4]["approved_by"]  = ""
    data[4]["approval_date"] = ""

    # SOD violation: requester = implementer
    data[9]["implemented_by"] = data[9]["requested_by"]

    # Same-day approval and implementation (rushed change)
    data[14]["approval_date"]      = data[14]["request_date"]
    data[14]["implementation_date"] = data[14]["request_date"]

    # Implemented before approval
    base = today - timedelta(days=50)
    data[19]["approval_date"]       = (base + timedelta(days=5)).strftime("%Y-%m-%d")
    data[19]["implementation_date"] = (base + timedelta(days=2)).strftime("%Y-%m-%d")

    # Not tested
    data[24]["tested"] = "No"
    data[29]["tested"] = "No"

    # SOD: approver = implementer
    data[34]["approved_by"] = data[34]["implemented_by"]

    # Missing implementation date
    data[39]["implementation_date"] = ""
    data[39]["status"]              = "Open"

    return pd.DataFrame(data)


# ─────────────────────────────────────────────
# SAMPLING ENGINE
# ─────────────────────────────────────────────
def determine_sample_size(population, override=None):
    """Return appropriate sample size per AICPA guidance."""
    if override:
        return min(override, population)
    for low, high, size in SAMPLE_THRESHOLDS:
        if low <= population <= high:
            return population if size is None else size
    return 40  # default for very large populations


def select_sample(df, sample_size, seed):
    """
    Select audit sample:
    - All emergency changes (mandatory 100% testing)
    - Random selection from remainder to hit sample_size
    Returns (sample_df, method_description)
    """
    emergency = df[df["is_emergency"].str.strip().str.lower() == "yes"].copy()
    normal    = df[df["is_emergency"].str.strip().str.lower() != "yes"].copy()

    emergency["selection_method"] = "Mandatory — Emergency Change"
    selected_normal = pd.DataFrame()

    remaining_needed = max(0, sample_size - len(emergency))

    if remaining_needed > 0 and len(normal) > 0:
        random.seed(seed)
        n = min(remaining_needed, len(normal))
        selected_normal = normal.sample(n=n, random_state=seed).copy()
        selected_normal["selection_method"] = f"Random Sample (seed={seed})"

    sample = pd.concat([emergency, selected_normal]).sort_values("change_id").reset_index(drop=True)
    sample["sample_number"] = range(1, len(sample) + 1)

    method = (
        f"AICPA-guided random sampling | "
        f"Population: {len(df)} | "
        f"Sample Size: {sample_size} | "
        f"Emergency (mandatory): {len(emergency)} | "
        f"Random: {len(selected_normal)} | "
        f"Random seed: {seed}"
    )
    return sample, method


# ─────────────────────────────────────────────
# EXCEPTION DETECTION ENGINE
# ─────────────────────────────────────────────
def detect_exceptions(df):
    """Scan all changes (not just sample) for automatic red flags."""
    exceptions = []

    df["request_date"]        = pd.to_datetime(df["request_date"],        errors="coerce")
    df["approval_date"]       = pd.to_datetime(df["approval_date"],        errors="coerce")
    df["implementation_date"] = pd.to_datetime(df["implementation_date"],  errors="coerce")

    for _, row in df.iterrows():
        cid = row["change_id"]

        def flag(exc_type, risk, detail, rec):
            exceptions.append({
                "Change ID":      cid,
                "Description":    row["description"],
                "Category":       row["category"],
                "Exception Type": exc_type,
                "Risk Rating":    risk,
                "Detail":         detail,
                "Recommendation": rec,
            })

        # 1. Missing approval
        if pd.isna(row["approved_by"]) or str(row["approved_by"]).strip() == "":
            flag(
                "Missing Approval",
                "Critical",
                f"Change {cid} has no recorded approver.",
                "Obtain approval documentation or flag as unauthorized change. Escalate to change manager."
            )

        # 2. SOD — requester is implementer
        elif str(row["requested_by"]).strip().lower() == str(row["implemented_by"]).strip().lower():
            flag(
                "SOD Violation — Requester = Implementer",
                "High",
                f"{row['requested_by']} both requested and implemented change {cid}.",
                "Require independent implementation. Update SDLC policy to enforce segregation."
            )

        # 3. SOD — approver is implementer
        if str(row["approved_by"]).strip().lower() == str(row["implemented_by"]).strip().lower() \
                and str(row["approved_by"]).strip() != "":
            flag(
                "SOD Violation — Approver = Implementer",
                "High",
                f"{row['approved_by']} both approved and implemented change {cid}.",
                "Require independent approval from someone who did not implement the change."
            )

        # 4. Implemented before approval
        if pd.notna(row["approval_date"]) and pd.notna(row["implementation_date"]):
            if row["implementation_date"] < row["approval_date"]:
                flag(
                    "Implemented Before Approval",
                    "Critical",
                    f"Change {cid} implemented on {row['implementation_date'].date()} "
                    f"before approval on {row['approval_date'].date()}.",
                    "Investigate root cause. Determine if change was authorized retroactively. "
                    "Strengthen pre-implementation approval gate controls."
                )

        # 5. Same-day request, approval, and implementation (rushed)
        if pd.notna(row["request_date"]) and pd.notna(row["approval_date"]) \
                and pd.notna(row["implementation_date"]):
            if row["request_date"] == row["approval_date"] == row["implementation_date"] \
                    and str(row["is_emergency"]).strip().lower() != "yes":
                flag(
                    "Same-Day Change — Not Emergency",
                    "Medium",
                    f"Change {cid} was requested, approved, and implemented on the same day "
                    f"({row['request_date'].date()}) but is not flagged as an emergency.",
                    "Verify adequate review time was provided. Consider if this should be reclassified "
                    "as an emergency change or if the approval process was circumvented."
                )

        # 6. Not tested
        if str(row["tested"]).strip().lower() in ["no", "false", "0", ""]:
            flag(
                "Change Not Tested",
                "High",
                f"Change {cid} has no recorded testing evidence.",
                "Obtain testing documentation (UAT sign-off, test results). "
                "If not tested, assess risk and determine if rollback is required."
            )

        # 7. Missing implementation date on completed change
        if str(row["status"]).strip().lower() == "completed" \
                and pd.isna(row["implementation_date"]):
            flag(
                "Completed Change — No Implementation Date",
                "Medium",
                f"Change {cid} is marked Completed but has no implementation date recorded.",
                "Update change record with implementation date. Review change management "
                "process to ensure all fields are completed before closure."
            )

    return pd.DataFrame(exceptions) if exceptions else pd.DataFrame()


# ─────────────────────────────────────────────
# ADD WORKPAPER COLUMNS TO SAMPLE
# ─────────────────────────────────────────────
def add_workpaper_columns(sample_df, exception_ids):
    """Add audit testing columns to the sample for manual completion."""
    wp = sample_df.copy()

    # Pre-populate exception column based on auto-detection
    def get_exception_note(cid):
        if cid in exception_ids:
            return "⚠️ See Exceptions Sheet"
        return ""

    wp["Approval Documented?"]    = wp["change_id"].apply(
        lambda x: "No ⚠️" if x in exception_ids else "")
    wp["Testing Evidence?"]       = ""
    wp["SOD Verified?"]           = ""
    wp["Impl. Date Matches?"]     = ""
    wp["Auditor Notes"]           = wp["change_id"].apply(get_exception_note)
    wp["Workpaper Conclusion"]    = ""   # Effective / Exception
    wp["Auditor Initials"]        = ""
    wp["Review Date"]             = ""

    return wp


# ─────────────────────────────────────────────
# EXCEL REPORT WRITER
# ─────────────────────────────────────────────
RISK_COLORS = {
    "Critical": "C00000",
    "High":     "FF0000",
    "Medium":   "FFC000",
    "Low":      "92D050",
}

def style_workbook(path, has_exceptions):
    wb = load_workbook(path)

    header_fill  = PatternFill("solid", fgColor="1F3864")
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin_border  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    def format_sheet(ws, col_widths):
        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = center_align
            cell.border    = thin_border
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border    = thin_border
                cell.alignment = left_align
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 30

    sheet_configs = {
        "Full Change Log": [12, 40, 14, 16, 16, 16, 14, 14, 18, 12, 14, 10],
        "Selected Sample": [8, 12, 40, 14, 16, 16, 16, 14, 14, 18, 12, 14, 10, 22, 20, 18, 18, 35, 20, 16, 14],
        "Exceptions":      [12, 40, 14, 32, 12, 45, 50],
        "Summary":         [40, 20],
    }

    for sheet_name, widths in sheet_configs.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        format_sheet(ws, widths)
        ws.freeze_panes = "A2"
        if ws.dimensions != "A1:A1":
            ws.auto_filter.ref = ws.dimensions

    # Color-code risk ratings in Exceptions sheet
    if "Exceptions" in wb.sheetnames and has_exceptions:
        ws_exc = wb["Exceptions"]
        risk_col = next(
            (i for i, cell in enumerate(ws_exc[1], 1) if cell.value == "Risk Rating"), None
        )
        if risk_col:
            for row in ws_exc.iter_rows(min_row=2):
                cell = row[risk_col - 1]
                color = RISK_COLORS.get(str(cell.value))
                if color:
                    cell.fill = PatternFill("solid", fgColor=color)
                    cell.font = Font(
                        bold=True,
                        color="FFFFFF" if cell.value in ["Critical", "High"] else "000000"
                    )
                    cell.alignment = Alignment(horizontal="center", vertical="center")

    # Highlight workpaper conclusion column in sample sheet
    if "Selected Sample" in wb.sheetnames:
        ws_samp = wb["Selected Sample"]
        conc_col = next(
            (i for i, c in enumerate(ws_samp[1], 1) if c.value == "Workpaper Conclusion"), None
        )
        if conc_col:
            wp_fill = PatternFill("solid", fgColor="FFF2CC")  # Light yellow — needs completion
            for row in ws_samp.iter_rows(min_row=2):
                cell = row[conc_col - 1]
                if not cell.value:
                    cell.fill = wp_fill

    wb.save(path)


def write_report(df_all, df_sample, df_exceptions, method, output_path):
    """Write all sheets to formatted Excel workbook."""

    # Prepare date columns for display
    date_cols = ["request_date", "approval_date", "implementation_date"]
    def fmt_dates(df):
        d = df.copy()
        for col in date_cols:
            if col in d.columns:
                d[col] = pd.to_datetime(d[col], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        return d

    df_all_display    = fmt_dates(df_all)
    df_sample_display = fmt_dates(df_sample)

    # Summary data
    summary_rows = [
        ("Audit Period",              "Full CSV range"),
        ("Total Population",          len(df_all)),
        ("Emergency Changes",         len(df_all[df_all["is_emergency"].str.strip().str.lower() == "yes"])),
        ("Normal Changes",            len(df_all[df_all["is_emergency"].str.strip().str.lower() != "yes"])),
        ("Sample Size Selected",      len(df_sample)),
        ("Emergency (mandatory)",     len(df_sample[df_sample["selection_method"].str.contains("Emergency", na=False)])),
        ("Random Sample",             len(df_sample[df_sample["selection_method"].str.contains("Random", na=False)])),
        ("Pre-flagged Exceptions",    len(df_exceptions) if not df_exceptions.empty else 0),
        ("Sampling Method",           method),
        ("Script Run Date",           datetime.today().strftime("%Y-%m-%d")),
        ("Prepared By",               ""),
        ("Reviewed By",               ""),
    ]
    df_summary = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_all_display.to_excel(writer,    sheet_name="Full Change Log",  index=False)
        df_sample_display.to_excel(writer, sheet_name="Selected Sample",  index=False)
        if not df_exceptions.empty:
            df_exceptions.to_excel(writer, sheet_name="Exceptions", index=False)
        else:
            pd.DataFrame(columns=["Change ID", "Description", "Category",
                                   "Exception Type", "Risk Rating", "Detail", "Recommendation"]
                         ).to_excel(writer, sheet_name="Exceptions", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

    style_workbook(output_path, not df_exceptions.empty)
    return df_summary


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Change Log Sampler — IT Audit Tool")
    parser.add_argument("--input",       "-i", default=None,
                        help="Path to change log CSV file")
    parser.add_argument("--output",      "-o", default="change_log_sample.xlsx",
                        help="Output Excel file path")
    parser.add_argument("--sample-size", "-n", type=int, default=None,
                        help="Override sample size (default: AICPA-guided)")
    parser.add_argument("--seed",        "-s", type=int, default=42,
                        help="Random seed for reproducible sampling (default: 42)")
    args = parser.parse_args()

    print("=" * 60)
    print("  CHANGE LOG SAMPLER — GRC & IT Audit Toolkit")
    print("=" * 60)

    # Load data
    if args.input:
        print(f"\n📂 Loading: {args.input}")
        df = pd.read_csv(args.input)
        try:
            validate_input(df)
        except ValueError as e:
            print(f"\n❌ ERROR: {e}")
            return
    else:
        print("\n⚠️  No input file provided. Using sample data for demo.")
        df = generate_sample_data()

    print(f"   {len(df)} change records loaded.")

    # Determine sample size
    sample_size = determine_sample_size(len(df), args.sample_size)
    print(f"\n📐 Sampling Strategy (AICPA-guided):")
    print(f"   Population   : {len(df)}")
    print(f"   Sample Size  : {sample_size}")
    print(f"   Random Seed  : {args.seed}")

    # Select sample
    df_sample, method = select_sample(df, sample_size, args.seed)
    print(f"   Emergency    : {len(df_sample[df_sample['selection_method'].str.contains('Emergency', na=False)])} (mandatory 100%)")
    print(f"   Random       : {len(df_sample[df_sample['selection_method'].str.contains('Random', na=False)])}")

    # Detect exceptions across full population
    print(f"\n🔍 Scanning full population for exceptions...")
    df_exceptions = detect_exceptions(df)
    exception_ids = set(df_exceptions["Change ID"].tolist()) if not df_exceptions.empty else set()

    checks = [
        "Missing approval",
        "SOD violations (requester = implementer)",
        "SOD violations (approver = implementer)",
        "Implemented before approval",
        "Same-day non-emergency changes",
        "Changes not tested",
        "Completed changes missing implementation date",
    ]
    for c in checks:
        print(f"   ✓ {c}")

    # Add workpaper columns
    df_sample = add_workpaper_columns(df_sample, exception_ids)

    # Write report
    print(f"\n📊 Writing report: {args.output}")
    write_report(df, df_sample, df_exceptions, method, args.output)

    # Console summary
    print("\n" + "=" * 60)
    print("  RESULTS SUMMARY")
    print("=" * 60)
    print(f"   Population       : {len(df)}")
    print(f"   Sample Selected  : {len(df_sample)}")
    if not df_exceptions.empty:
        print(f"\n   Exceptions Found : {len(df_exceptions)}")
        for risk in ["Critical", "High", "Medium", "Low"]:
            count = len(df_exceptions[df_exceptions["Risk Rating"] == risk])
            if count > 0:
                print(f"   {risk:<12}   : {count}")
    else:
        print(f"\n   ✅ No exceptions detected in full population.")

    print(f"\n✅ Report saved: {args.output}")
    print(f"   Sheets: Full Change Log | Selected Sample | Exceptions | Summary")
    print("=" * 60)


if __name__ == "__main__":
    main()
