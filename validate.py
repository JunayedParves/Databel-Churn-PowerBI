"""Recompute the Databel churn report figures from the CSV with pandas.

Mirrors the Power Query steps and DAX logic in the semantic model so the
numbers shown in Power BI can be checked independently.

Usage:
    python validate.py                       # uses the data/ folder next to this script
    python validate.py "D:/path/to/folder"   # folder containing "Databel - Data.csv"
"""
import sys
from pathlib import Path

import pandas as pd

DATA_FOLDER = Path(__file__).resolve().parent / "data"
CSV_NAME = "Databel - Data.csv"
TOLERANCE = 0.0005  # rates are compared at 0.1% display precision

INT_COLS = [
    "Age",
    "Account Length (in months)",
    "Customer Service Calls",
    "Monthly Charge",
    "Number of Customers in Group",
]


def load(folder: str) -> pd.DataFrame:
    # Power Query: promote headers, whole-number columns typed, rest text
    df = pd.read_csv(Path(folder) / CSV_NAME, dtype=str, keep_default_na=False)
    for c in INT_COLS:
        df[c] = df[c].astype("int64")
    # Capitalize each word in Intl Plan (source is yes/no)
    df["Intl Plan"] = df["Intl Plan"].str.title()
    # Blank Churn Category / Reason stay blank for non-churners
    df["churned"] = df["Churn Label"] == "Yes"

    # Calculated columns (same breaks as the DAX)
    age = df["Age"]
    df["Age Group"] = pd.cut(age, [-1, 29, 39, 49, 64, 200],
                             labels=["Under 30", "30-39", "40-49", "50-64", "65+"]).astype(str)
    m = df["Account Length (in months)"]
    df["Tenure Group"] = pd.cut(m, [-1, 6, 12, 24, 48, 10_000],
                                labels=["0-6", "7-12", "13-24", "25-48", "49+"]).astype(str)
    csc = df["Customer Service Calls"]
    df["Service Calls Group"] = csc.where(csc < 5, 5).astype(str).replace("5", "5+")
    df["Senior Label"] = df["Senior"].map(lambda s: "Senior" if s == "Yes" else "Not senior")
    return df


def churn_rate(d: pd.DataFrame) -> float:
    return d["churned"].mean() if len(d) else float("nan")


def rate_by(df: pd.DataFrame, col) -> dict:
    return df.groupby(col)["churned"].mean().to_dict()


failures = []


def check(label, actual, expected, tol=TOLERANCE, fmt="{:.1%}"):
    ok = abs(actual - expected) <= tol
    shown = fmt.format(actual)
    exp = fmt.format(expected)
    print(f"  {'OK ' if ok else 'FAIL'}  {label:<38} {shown:>10}   (expected {exp})")
    if not ok:
        failures.append(label)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else DATA_FOLDER
    df = load(folder)
    churned = df[df["churned"]]

    print("Headline KPIs")
    check("Total Customers", len(df), 6687, 0, "{:,.0f}")
    check("Churned Customers", len(churned), 1796, 0, "{:,.0f}")
    check("Churn Rate", churn_rate(df), 0.269)
    check("Monthly Revenue Lost", churned["Monthly Charge"].sum(), 66094, 0, "${:,.0f}")
    check("Total Monthly Revenue", df["Monthly Charge"].sum(), 207500, 0, "${:,.0f}")
    check("Group Churn Rate", churn_rate(df[df["Group"] == "Yes"]), 0.065)
    check("Non-Group Churn Rate", churn_rate(df[df["Group"] == "No"]), 0.328)
    # Panel A filters out blank categories (27 churners have none), and
    # ALLSELECTED in Share of Churners respects that visual filter.
    categorised = churned[churned["Churn Category"] != ""]
    share = categorised["Churn Category"].value_counts(normalize=True)
    check("Competitor share of churners", share.get("Competitor", 0), 0.455)

    print("\nContract type")
    r = rate_by(df, "Contract Type")
    for k, v in {"Month-to-Month": .463, "One Year": .113, "Two Year": .028}.items():
        check(k, r[k], v)

    print("\nTenure group")
    r = rate_by(df, "Tenure Group")
    for k, v in {"0-6": .531, "7-12": .366, "13-24": .295, "25-48": .205, "49+": .097}.items():
        check(k, r[k], v)

    print("\nPayment method")
    r = rate_by(df, "Payment Method")
    for k, v in {"Paper Check": .380, "Direct Debit": .345, "Credit Card": .145}.items():
        check(k, r[k], v)

    print("\nCustomer service calls")
    r = rate_by(df, "Service Calls Group")
    for k, v in {"0": .089, "1": .313, "2": .366, "3": .875, "4": .997, "5+": 1.0}.items():
        check(k, r[k], v)

    print("\nIntl Plan x Intl Active")
    r = rate_by(df, ["Intl Plan", "Intl Active"])
    for k, v in {("Yes", "Yes"): .076, ("Yes", "No"): .712,
                 ("No", "Yes"): .403, ("No", "No"): .200}.items():
        check(f"Plan {k[0]} / Active {k[1]}", r[k], v)

    print("\nAge / senior")
    r = rate_by(df, "Age Group")
    check("Under 30", r["Under 30"], .230)
    check("65+", r["65+"], .385)
    r = rate_by(df, ["Senior Label", "Contract Type"])
    check("Senior + Month-to-Month", r[("Senior", "Month-to-Month")], .746)

    # Figures quoted in headline / footnote text boxes
    print("\nText-box figures (informational)")
    m2m, two = rate_by(df, "Contract Type")["Month-to-Month"], rate_by(df, "Contract Type")["Two Year"]
    print(f"  M2M / Two Year multiple                {m2m / two:>10.2f}")
    print(f"  Revenue Lost %                         {churned['Monthly Charge'].sum() / df['Monthly Charge'].sum():>10.1%}")
    three_plus = churned["Customer Service Calls"].ge(3).mean()
    print(f"  3+ calls share of churners             {three_plus:>10.1%}")
    ca = df[df["State"] == "CA"]
    print(f"  CA customers / churn rate              {len(ca):>6} / {churn_rate(ca):.1%}")
    top_states = (df.groupby("State")["churned"].agg(["size", "mean"])
                  .sort_values("mean", ascending=False).head(5))
    print("  Top 5 states by churn rate:")
    for s, row in top_states.iterrows():
        print(f"    {s}  n={int(row['size']):>4}  {row['mean']:.1%}")
    print("  Top 5 churn reasons:")
    for reason, n in churned["Churn Reason"].value_counts().head(5).items():
        print(f"    {n:>4}  {reason}")

    print()
    if failures:
        print(f"{len(failures)} check(s) FAILED: {', '.join(failures)}")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()
