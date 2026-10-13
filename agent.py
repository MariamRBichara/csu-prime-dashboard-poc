from pathlib import Path
import json
import pandas as pd


BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"


# =========================================================
# INTERNAL SOURCE C — POC CONTROL DATA
# =========================================================
#
# This is temporary for the PoC.
#
# Later this should be replaced by the controlled internal
# monthly HC / capacity source.
#
# Format:
# YYYY-MM
# =========================================================

SOURCE_C = {
    "2026-07": {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 23,
    },
    "2026-08": {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 21,
    },
}


# =========================================================
# GENERIC HELPERS
# =========================================================

def load_sheet_containing_columns(path, required_columns):

    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(
            f"Input file not found: {path}"
        )

    workbook = pd.read_excel(
        path,
        sheet_name=None,
    )

    for sheet_name, df in workbook.items():

        df.columns = [
            str(c).strip()
            for c in df.columns
        ]

        if all(
            col in df.columns
            for col in required_columns
        ):

            print(
                f"✓ Using worksheet "
                f"'{sheet_name}' from {path.name}"
            )

            return df.copy()

    raise ValueError(
        f"Could not identify the required worksheet "
        f"in {path.name}.\n"
        f"Required columns: {required_columns}"
    )


def numeric(series):

    return pd.to_numeric(
        series,
        errors="coerce",
    ).fillna(0)


def normalize_month_name(value):

    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.month

    text = str(value).strip()

    # Explicit MM-YY format
    # Example: 07-26

    parts = text.split("-")

    if len(parts) == 2:

        try:
            month = int(parts[0])

            if 1 <= month <= 12:
                return month

        except Exception:
            pass

    # Fallback

    try:

        parsed = pd.to_datetime(
            text,
            errors="raise",
        )

        return parsed.month

    except Exception:

        return None


def normalize_year(value):

    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.year

    text = str(value).strip()

    # Explicit MM-YY format
    # Example: 07-26 -> 2026

    parts = text.split("-")

    if len(parts) == 2:

        try:

            year = int(parts[1])

            if year < 100:
                year += 2000

            return year

        except Exception:
            pass

    # Fallback

    try:

        parsed = pd.to_datetime(
            text,
            errors="raise",
        )

        return parsed.year

    except Exception:

        return None


def normalize_completion(value):

    """
    Source A Completion % is already stored
    as a decimal ratio.

    Examples:

    0.96 = 96%
    1.10 = 110%

    Do NOT divide values above 1 by 100.
    """

    try:
        return float(value)

    except Exception:
        return 0


# =========================================================
# SOURCE A
# =========================================================

def load_source_a(source_a_path):

    required = [
        "Legal Entity Name",
        "Work Type Name",
        "Timesheets Billable",
        "Timesheets Non Billable",
        "All Timesheets",
        "Completion %",
        "Man Period Name",
    ]

    df = load_sheet_containing_columns(
        source_a_path,
        required,
    )

    df["MonthNum"] = (
        df["Man Period Name"]
        .apply(normalize_month_name)
    )

    df["YearNum"] = (
        df["Man Period Name"]
        .apply(normalize_year)
    )

    for col in [
        "Timesheets Billable",
        "Timesheets Non Billable",
        "All Timesheets",
    ]:

        df[col] = numeric(
            df[col]
        )

    df["CompletionDecimal"] = (
        df["Completion %"]
        .apply(normalize_completion)
    )

    df["Work Type Name"] = (
        df["Work Type Name"]
        .astype(str)
        .str.strip()
    )

    return df


# =========================================================
# SOURCE B
# =========================================================

def load_source_b(source_b_path):

    required = [
        "Person Number",
        "Mois",
        "Year",
        "Time Entry",
    ]

    df = load_sheet_containing_columns(
        source_b_path,
        required,
    )

    df["Mois"] = pd.to_numeric(
        df["Mois"],
        errors="coerce",
    )

    df["Year"] = pd.to_numeric(
        df["Year"],
        errors="coerce",
    )

    df["Time Entry"] = numeric(
        df["Time Entry"]
    )

    return df


# =========================================================
# DASHBOARD CALCULATION
# =========================================================

def calculate_dashboard(
    a,
    b,
    reporting_year,
    reporting_months,
):

    reporting_months = sorted(
        reporting_months
    )

    month_names = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December",
    }


    # =====================================================
    # REPORTING PERIOD
    # =====================================================

    period_a = a[
        (a["YearNum"] == reporting_year)
        &
        (
            a["MonthNum"]
            .isin(reporting_months)
        )
    ].copy()

    if period_a.empty:

        available_periods = (
            a[
                [
                    "Man Period Name",
                    "MonthNum",
                    "YearNum",
                ]
            ]
            .drop_duplicates()
            .to_dict("records")
        )

        raise ValueError(
            "No Source A data found for "
            "the requested reporting period.\n"
            f"Requested year: {reporting_year}\n"
            f"Requested months: {reporting_months}\n"
            f"Detected periods: {available_periods}"
        )


    # =====================================================
    # BILLABILITY TREND
    # =====================================================

    trend = {}

    year_a = a[
        a["YearNum"]
        == reporting_year
    ]

    available_months = sorted(
        int(x)
        for x in year_a[
            "MonthNum"
        ]
        .dropna()
        .unique()
    )

    for month in available_months:

        month_df = year_a[
            year_a["MonthNum"]
            == month
        ]

        total = (
            month_df[
                "All Timesheets"
            ]
            .sum()
        )

        if total <= 0:
            continue

        billable = (
            month_df[
                "Timesheets Billable"
            ]
            .sum()
        )

        non_billable = (
            month_df[
                "Timesheets Non Billable"
            ]
            .sum()
        )

        trend[
            month_names[month]
        ] = {

            "month_number":
                month,

            "all_timesheets":
                float(total),

            "billable_time":
                float(billable),

            "non_billable_time":
                float(non_billable),

            "billable_pct":
                float(
                    billable
                    / total
                ),

            "non_billable_pct":
                float(
                    non_billable
                    / total
                ),
        }


    # =====================================================
    # WORK TYPES
    # =====================================================

    worktype_lower = (
        period_a[
            "Work Type Name"
        ]
        .str.lower()
        .str.strip()
    )


    billable_worktype = (
        period_a.loc[
            worktype_lower.isin(
                {
                    "billable",
                    "own work capitalized",
                }
            ),
            "All Timesheets",
        ]
        .sum()
    )


    general = (
        period_a.loc[
            worktype_lower.isin(
                {
                    "administrative",
                    "non billable",
                }
            ),
            "All Timesheets",
        ]
        .sum()
    )


    leave = (
        period_a.loc[
            worktype_lower
            == "leave",
            "All Timesheets",
        ]
        .sum()
    )


    presales = (
        period_a.loc[
            worktype_lower.isin(
                {
                    "pre sales",
                    "pre-sales",
                    "presales",
                }
            ),
            "All Timesheets",
        ]
        .sum()
    )


    sickness = (
        period_a.loc[
            worktype_lower
            == "sickness",
            "All Timesheets",
        ]
        .sum()
    )


    actual_reported = (
        period_a[
            "All Timesheets"
        ]
        .sum()
    )


    # =====================================================
    # EXPECTED CAPACITY
    # =====================================================

    expected_capacity = 0

    grouped = period_a.groupby(
        [
            "Legal Entity Name",
            "MonthNum",
        ],
        dropna=False,
    )

    for _, group in grouped:

        entity_completion = (
            group[
                "CompletionDecimal"
            ]
            .sum()
        )

        entity_actual = (
            group[
                "All Timesheets"
            ]
            .sum()
        )

        if entity_completion > 0:

            expected_capacity += (
                entity_actual
                / entity_completion
            )


    if expected_capacity <= 0:

        raise ValueError(
            "Expected capacity is zero "
            "or invalid."
        )


    # =====================================================
    # COMPLETION / ARVI / ARVE
    # =====================================================

    completion_rate = (
        actual_reported
        / expected_capacity
    )

    productive_time = (
        billable_worktype
    )

    adjusted_capacity = (
        expected_capacity
        - leave
        - sickness
    )

    if adjusted_capacity <= 0:

        raise ValueError(
            "Adjusted capacity is "
            "zero or invalid."
        )

    arvi = (
        productive_time
        / expected_capacity
    )

    arve = (
        productive_time
        / adjusted_capacity
    )


    # =====================================================
    # SOURCE C
    # =====================================================

    source_c_rows = {}

    for month in reporting_months:

        key = (
            f"{reporting_year}-"
            f"{month:02d}"
        )

        if key not in SOURCE_C:

            raise ValueError(
                "Source C is missing "
                "validated control data "
                f"for {key}."
            )

        source_c_rows[
            month
        ] = SOURCE_C[key]


    # =====================================================
    # HC COMPLIANCE
    # =====================================================

    submitters = {}

    for month in reporting_months:

        month_b = b[
            (b["Mois"] == month)
            &
            (b["Year"] == reporting_year)
        ]

        submitters[
            month
        ] = int(
            month_b[
                "Person Number"
            ]
            .dropna()
            .nunique()
        )


    hc_tagged_total = sum(
        row["hc_tagged"]
        for row
        in source_c_rows.values()
    )


    hc_compliance = (
        sum(
            submitters.values()
        )
        /
        hc_tagged_total
    )


    # =====================================================
    # OVERBOOKING
    # =====================================================

    overbooking_occurrences = {}

    for month in reporting_months:

        month_b = b[
            (b["Mois"] == month)
            &
            (b["Year"] == reporting_year)
        ]

        person_totals = (
            month_b
            .groupby(
                "Person Number"
            )[
                "Time Entry"
            ]
            .sum()
        )

        expected_days = (
            source_c_rows[
                month
            ][
                "expected_working_days"
            ]
        )

        overbooking_occurrences[
            month
        ] = int(
            (
                person_totals
                > expected_days
            )
            .sum()
        )


    overbooking_total = sum(
        overbooking_occurrences.values()
    )


    first_month = (
        reporting_months[0]
    )


    # Current agreed convention:
    # total occurrences / HC tagged
    # of first month

    overbooking_pct = (
        overbooking_total
        /
        source_c_rows[
            first_month
        ][
            "hc_tagged"
        ]
    )


    # =====================================================
    # WORKTYPE RECONCILIATION
    # =====================================================

    displayed_total = (
        billable_worktype
        + general
        + leave
        + presales
    )

    non_displayed = (
        actual_reported
        - displayed_total
    )


    # =====================================================
    # REPORTING PERIOD LABEL
    # =====================================================

    last_month = (
        reporting_months[-1]
    )

    if (
        first_month
        == last_month
    ):

        reporting_period = (
            f"{month_names[first_month]} "
            f"{reporting_year}"
        )

    else:

        reporting_period = (
            f"{month_names[first_month]} "
            f"& "
            f"{month_names[last_month]} "
            f"{reporting_year}"
        )


    # =====================================================
    # RETURN RESULT
    # =====================================================

    return {

        "reporting_period":
            reporting_period,

        "reporting_year":
            reporting_year,

        "reporting_months":
            reporting_months,


        "kpis": {

            "hc_tagged":
                source_c_rows[
                    first_month
                ][
                    "hc_tagged"
                ],

            "total_hc":
                source_c_rows[
                    first_month
                ][
                    "total_hc"
                ],

            "completion_rate":
                float(
                    completion_rate
                ),

            "hc_compliance":
                float(
                    hc_compliance
                ),

            "arve":
                float(
                    arve
                ),

            "arvi":
                float(
                    arvi
                ),

            "overbooking":
                float(
                    overbooking_pct
                ),
        },


        "worktypes": {

            "Billable":
                float(
                    billable_worktype
                ),

            "General":
                float(
                    general
                ),

            "Leave":
                float(
                    leave
                ),

            "Pre-sales":
                float(
                    presales
                ),

            "Non-displayed":
                float(
                    non_displayed
                ),

            "Total":
                float(
                    actual_reported
                ),
        },


        "trend":
            trend,


        "diagnostics": {

            "expected_capacity":
                float(
                    expected_capacity
                ),

            "actual_reported":
                float(
                    actual_reported
                ),

            "productive_time":
                float(
                    productive_time
                ),

            "leave":
                float(
                    leave
                ),

            "sickness":
                float(
                    sickness
                ),

            "submitters":
                submitters,

            "overbooking_occurrences":
                overbooking_occurrences,

            "overbooking_total":
                int(
                    overbooking_total
                ),
        },
    }


# =========================================================
# SAVE METRICS
# =========================================================

def save_metrics(
    metrics,
    output_path=None,
):

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    if output_path is None:

        output_path = (
            OUTPUT_DIR
            / "calculated_metrics.json"
        )

    with open(
        output_path,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            metrics,
            f,
            indent=2,
        )

    return str(
        output_path
    )