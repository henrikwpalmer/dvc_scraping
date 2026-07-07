"""
Analyze case data by "Solution" (the 'solution' column) from the ENGLISH tab
of cases_output_ALL_COURTS_DONETSK.xlsx.

For each distinct Solution value, this produces:
  - Total number of cases
  - Number of those cases that have a link to the judicial decision
    (non-empty 'judicial_acts_link')
  - Number of unique courts that issued that Solution
  - First and last occurrence dates, based on 'decision_date'
    (format in the source is dd.mm.yyyy, but parsing is done defensively
    since the field is not guaranteed to be clean/consistent)
"""

import pandas as pd

SOURCE_FILE = "../output/cases_output_ALL_COURTS_DONETSK.xlsx"
SHEET_NAME = "ENGLISH"
OUTPUT_FILE = "../output/solution_summary_donetsk.xlsx"


def parse_decision_date(series: pd.Series) -> pd.Series:
    """
    Robustly parse the 'decision_date' column into real datetimes.

    The field is expected to look like dd.mm.yyyy, but since it may not
    always be formatted properly (stray whitespace, other separators,
    non-date text, etc.), we:
      1. Strip whitespace and coerce to string.
      2. Try the expected dd.mm.yyyy format first (dayfirst, exact).
      3. Fall back to a flexible pandas parse (dayfirst=True) for anything
         that didn't match, so odd-but-parseable formats aren't lost.
      4. Anything still unparseable becomes NaT and is excluded from the
         date-range calculation (but the case itself is NOT dropped from
         the counts).
    """
    cleaned = series.astype(str).str.strip()
    cleaned = cleaned.replace({"nan": None, "NaT": None, "": None})

    # First pass: strict expected format
    parsed = pd.to_datetime(cleaned, format="%d.%m.%Y", errors="coerce")

    # Second pass: for anything that failed, try a flexible dayfirst parse
    still_missing = parsed.isna() & cleaned.notna()
    if still_missing.any():
        parsed.loc[still_missing] = pd.to_datetime(
            cleaned[still_missing], dayfirst=True, errors="coerce"
        )

    return parsed


def main():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME)

    # Normalize the Solution field (strip whitespace) and treat truly blank
    # values as "No solution recorded" so they still show up in the summary.
    # Note: with newer pandas string dtypes, astype(str) can leave real NaNs
    # as float NaN rather than the text "nan", so missingness is checked
    # with pd.isna(...) directly rather than by matching string sentinels.
    is_blank_solution = df["solution"].isna() | (
        df["solution"].astype(str).str.strip() == ""
    )
    df["solution"] = df["solution"].astype(str).str.strip()
    df.loc[is_blank_solution, "solution"] = "No solution recorded"

    # A case "has a link to the judicial decision" if judicial_acts_link
    # is present and non-blank.
    is_blank_link = df["judicial_acts_link"].isna() | (
        df["judicial_acts_link"].astype(str).str.strip() == ""
    )
    df["has_judicial_acts_link"] = ~is_blank_link

    # Parse decision dates defensively.
    df["decision_date_parsed"] = parse_decision_date(df["decision_date"])
    n_unparsed = df["decision_date_parsed"].isna().sum() - df["decision_date"].isna().sum()
    if n_unparsed > 0:
        print(f"Note: {n_unparsed} decision_date value(s) could not be parsed "
              f"as dates and were excluded from date-range calculations.")

    summary = (
        df.groupby("solution")
        .agg(
            number_of_cases=("solution", "size"),
            cases_with_decision_link=("has_judicial_acts_link", "sum"),
            unique_courts=("court_name", pd.Series.nunique),
            first_decision_date=("decision_date_parsed", "min"),
            last_decision_date=("decision_date_parsed", "max"),
        )
        .reset_index()
        .rename(columns={"solution": "Solution"})
        .sort_values("number_of_cases", ascending=False)
    )

    # Friendly date formatting for display / export
    summary["first_decision_date"] = summary["first_decision_date"].dt.strftime("%d.%m.%Y")
    summary["last_decision_date"] = summary["last_decision_date"].dt.strftime("%d.%m.%Y")

    summary = summary.rename(
        columns={
            "number_of_cases": "Number of Cases",
            "cases_with_decision_link": "Cases with Decision Link",
            "unique_courts": "Unique Courts",
            "first_decision_date": "First Occurrence",
            "last_decision_date": "Last Occurrence",
        }
    )

    print(summary.to_string(index=False))

    summary.to_excel(OUTPUT_FILE, index=False)
    print(f"\nSaved summary to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
