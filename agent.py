from pathlib import Path
import json
import os
from collections import Counter

import pandas as pd


BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"

# Total HC is not available in Source A or Source B.
# It can be overridden at deployment level without changing code:
# CSU_TOTAL_HC=761
DEFAULT_TOTAL_HC = int(os.getenv("CSU_TOTAL_HC", "761"))

MONTH_NAMES = {
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

SOURCE_A_REQUIRED = [
    "Legal Entity Name",
    "Work Type Name",
    "Timesheets Billable",
    "Timesheets Non Billable",
    "All Timesheets",
    "Completion %",
    "Man Period Name",
]

SOURCE_B_REQUIRED = [
    "Person Number",
    "Mois",
    "Year",
    "Time Entry",
]

# Historical PoC controls that were already validated.
# These allow Jul/Aug 2026 to continue working even when a newer
# Source B upload contains only Sep/Oct.
HISTORICAL_CONTROLS = {
    (2026, 7): {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 23.0,
        "hc_compliant": 633,
        "overbooked": 21,
        "control_source": "validated historical control",
    },
    (2026, 8): {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 21.0,
        "hc_compliant": 623,
        "overbooked": 40,
        "control_source": "validated historical control",
    },
}


def _clean_header(value):
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split()).strip()


def load_sheet_containing_columns(path, required_columns):
    """
    Find and load the worksheet containing all required columns.
    Also tolerates title rows above the real header.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    excel = pd.ExcelFile(path, engine="openpyxl")
    required = {_clean_header(c) for c in required_columns}
    checked = []

    for sheet_name in excel.sheet_names:
        preview = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=20,
            engine="openpyxl",
        )

        for header_row in range(len(preview)):
            headers = [_clean_header(v) for v in preview.iloc[header_row].tolist()]
            header_set = {h for h in headers if h}

            if required.issubset(header_set):
                df = pd.read_excel(
                    path,
                    sheet_name=sheet_name,
                    header=header_row,
                    engine="openpyxl",
                )
                df.columns = [_clean_header(c) for c in df.columns]
                return df

        checked.append(sheet_name)

    raise ValueError(
        f"Could not identify the required worksheet in {path.name}. "
        f"Required columns: {required_columns}. "
        f"Checked sheets: {checked}"
    )


def numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def normalize_completion(series):
    """
    Source A Completion % is already a decimal ratio.
    Values greater than 1 are legitimate, e.g. 1.10 = 110%.
    """
    return pd.to_numeric(series, errors="coerce")


def normalize_month_name(value):
    """Return month number from Source A period values such as 07-26."""
    if pd.isna(value):
        return None

    text = str(value).strip()
    parts = text.split("-")

    if len(parts) >= 2 and parts[0].strip().isdigit():
        month = int(parts[0].strip())
        if 1 <= month <= 12:
            return month

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return int(parsed.month)

    return None


def normalize_year(value):
    """Return four-digit year from Source A period values such as 07-26."""
    if pd.isna(value):
        return None

    text = str(value).strip()
    parts = text.split("-")

    if len(parts) >= 2 and parts[1].strip().isdigit():
        year = int(parts[1].strip())
        if year < 100:
            year += 2000
        return year

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return int(parsed.year)

    return None


def load_source_a(source_a_path):
    df = load_sheet_containing_columns(source_a_path, SOURCE_A_REQUIRED).copy()

    for col in [
        "Timesheets Billable",
        "Timesheets Non Billable",
        "All Timesheets",
    ]:
        df[col] = numeric(df[col])

    df["Completion %"] = normalize_completion(df["Completion %"])
    df["_month"] = df["Man Period Name"].apply(normalize_month_name)
    df["_year"] = df["Man Period Name"].apply(normalize_year)

    df = df[df["_month"].notna() & df["_year"].notna()].copy()
    df["_month"] = df["_month"].astype(int)
    df["_year"] = df["_year"].astype(int)

    return df


def _normalize_source_b_month(value):
    if pd.isna(value):
        return None

    try:
        month = int(float(str(value).strip()))
    except (TypeError, ValueError):
        return None

    return month if 1 <= month <= 12 else None


def load_source_b(source_b_path):
    df = load_sheet_containing_columns(source_b_path, SOURCE_B_REQUIRED).copy()

    df["_month"] = df["Mois"].apply(_normalize_source_b_month)
    df["_year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["_time_entry"] = numeric(df["Time Entry"])
    df["_person"] = df["Person Number"].astype(str).str.strip()

    df = df[
        df["_month"].notna()
        & df["_year"].notna()
        & df["_person"].ne("")
        & df["_person"].ne("nan")
    ].copy()

    df["_month"] = df["_month"].astype(int)
    df["_year"] = df["_year"].astype(int)

    return df


def _infer_expected_working_days(person_totals):
    """
    Infer the normal monthly capacity from the dominant person-month total.
    This removes the need to hard-code a calendar value every month.
    """
    rounded = [
        round(float(v), 4)
        for v in person_totals
        if pd.notna(v) and float(v) > 0
    ]

    if not rounded:
        raise ValueError(
            "Could not infer expected working days because Source B has no "
            "positive Time Entry totals for the selected month."
        )

    frequencies = Counter(rounded)
    highest_frequency = max(frequencies.values())
    modes = [
        value
        for value, count in frequencies.items()
        if count == highest_frequency
    ]

    expected = max(modes)

    if expected <= 0 or expected > 31:
        raise ValueError(
            f"Inferred expected working days ({expected}) is outside "
            "a valid monthly range."
        )

    return float(expected)


def _dynamic_control_for_month(source_b, reporting_year, month):
    """
    Derive PoC controls for a current/future month from Source B.

    HC compliance follows the validated historical convention:
    number of people represented in the month / HC tagged.

    HC tagged is the unique annual pool represented in the uploaded
    Source B file. Total HC remains a deployment-level control because
    it is not contained in Source A or B.
    """
    year_data = source_b[source_b["_year"] == int(reporting_year)].copy()
    month_data = year_data[year_data["_month"] == int(month)].copy()

    if month_data.empty:
        return None

    tagged_people = {
        p
        for p in year_data["_person"].dropna().astype(str).str.strip()
        if p and p != "nan"
    }
    hc_tagged = len(tagged_people)

    if hc_tagged == 0:
        raise ValueError("Could not derive HC tagged from Source B.")

    person_totals = (
        month_data.groupby("_person", dropna=True)["_time_entry"]
        .sum()
        .astype(float)
    )

    expected_days = _infer_expected_working_days(person_totals.tolist())

    # Historical validated compliance used the number of people
    # represented in the reporting month, not the number who reached
    # the expected-day threshold.
    hc_compliant = int(person_totals.size)

    tolerance = 1e-9
    overbooked = int((person_totals > expected_days + tolerance).sum())

    return {
        "hc_tagged": hc_tagged,
        "total_hc": DEFAULT_TOTAL_HC,
        "expected_working_days": expected_days,
        "hc_compliant": hc_compliant,
        "overbooked": overbooked,
        "control_source": "derived from uploaded Source B",
    }


def _get_month_control(source_b, reporting_year, month):
    """
    Use validated historical control when available.
    Otherwise derive the control automatically from the uploaded Source B.
    """
    key = (int(reporting_year), int(month))

    if key in HISTORICAL_CONTROLS:
        return dict(HISTORICAL_CONTROLS[key])

    dynamic = _dynamic_control_for_month(
        source_b,
        reporting_year,
        month,
    )

    if dynamic is not None:
        return dynamic

    available = sorted(
        source_b.loc[
            source_b["_year"] == int(reporting_year),
            "_month",
        ]
        .dropna()
        .astype(int)
        .unique()
        .tolist()
    )

    available_text = (
        ", ".join(MONTH_NAMES[m] for m in available)
        if available
        else "none"
    )

    raise ValueError(
        f"Source B has no data for {MONTH_NAMES[int(month)]} "
        f"{reporting_year}. Available Source B months: {available_text}."
    )


def _expected_capacity(source_a_filtered):
    """
    Reconstruct expected capacity from Source A.

    For each legal-entity/month group:
        capacity = All Timesheets / Completion %
    We take the median candidate per group to avoid duplicate counting.
    """
    candidates = []

    for _, group in source_a_filtered.groupby(
        ["Legal Entity Name", "_year", "_month"],
        dropna=False,
    ):
        valid = group[
            group["Completion %"].notna()
            & (group["Completion %"] > 0)
            & (group["All Timesheets"] >= 0)
        ].copy()

        if valid.empty:
            continue

        capacity_candidates = (
            valid["All Timesheets"] / valid["Completion %"]
        ).replace(
            [float("inf"), float("-inf")],
            pd.NA,
        ).dropna()

        if not capacity_candidates.empty:
            candidates.append(float(capacity_candidates.median()))

    if not candidates:
        raise ValueError(
            "Could not derive expected capacity from Source A. "
            "Check Completion % and All Timesheets values."
        )

    return float(sum(candidates))


def _worktype_total(df, names):
    names_normalized = {
        str(name).strip().lower()
        for name in names
    }
    worktype = (
        df["Work Type Name"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    return float(
        df.loc[
            worktype.isin(names_normalized),
            "All Timesheets",
        ].sum()
    )


def _month_label(month):
    return MONTH_NAMES[int(month)]


def _reporting_period(reporting_year, reporting_months):
    months = sorted(set(int(m) for m in reporting_months))

    if len(months) == 1:
        return f"{_month_label(months[0])} {reporting_year}"

    names = [_month_label(m) for m in months]

    if len(names) == 2:
        month_text = f"{names[0]} & {names[1]}"
    else:
        month_text = ", ".join(names[:-1]) + f" & {names[-1]}"

    return f"{month_text} {reporting_year}"


def calculate_dashboard(
    source_a,
    source_b,
    reporting_year,
    reporting_months,
):
    reporting_year = int(reporting_year)
    reporting_months = sorted(
        set(int(m) for m in reporting_months)
    )

    if not reporting_months:
        raise ValueError("Select at least one reporting month.")

    invalid_months = [
        m for m in reporting_months
        if m not in MONTH_NAMES
    ]
    if invalid_months:
        raise ValueError(
            f"Invalid reporting month(s): {invalid_months}"
        )

    # -----------------------------------------------------
    # Validate selected months against Source A.
    # -----------------------------------------------------
    year_a = source_a[
        source_a["_year"] == reporting_year
    ].copy()

    available_a = sorted(
        year_a["_month"]
        .dropna()
        .astype(int)
        .unique()
        .tolist()
    )

    missing_a = [
        m for m in reporting_months
        if m not in available_a
    ]

    if missing_a:
        missing_text = ", ".join(
            _month_label(m)
            for m in missing_a
        )
        available_text = (
            ", ".join(
                _month_label(m)
                for m in available_a
            )
            if available_a
            else "none"
        )

        raise ValueError(
            f"Source A is missing {missing_text} {reporting_year}. "
            f"Available Source A months: {available_text}."
        )

    selected_a = year_a[
        year_a["_month"].isin(reporting_months)
    ].copy()

    if selected_a.empty:
        raise ValueError(
            "No Source A data found for the selected reporting period."
        )

    # -----------------------------------------------------
    # Source-derived billability trend for every month
    # available in Source A for the reporting year.
    # -----------------------------------------------------
    trend = {}

    for month in available_a:
        month_df = year_a[
            year_a["_month"] == month
        ]

        total = float(
            month_df["All Timesheets"].sum()
        )
        billable = float(
            month_df["Timesheets Billable"].sum()
        )

        billable_pct = (
            billable / total
            if total
            else 0.0
        )

        trend[_month_label(month)] = {
            "month_number": int(month),
            "billable_pct": float(billable_pct),
            "non_billable_pct": (
                float(1.0 - billable_pct)
                if total
                else 0.0
            ),
        }

    # -----------------------------------------------------
    # Source A metrics.
    # -----------------------------------------------------
    actual_reported = float(
        selected_a["All Timesheets"].sum()
    )
    expected_capacity = _expected_capacity(
        selected_a
    )

    billable_productive = _worktype_total(
        selected_a,
        ["Billable", "Own Work Capitalized"],
    )
    general = _worktype_total(
        selected_a,
        ["Non Billable", "Administrative"],
    )
    leave = _worktype_total(
        selected_a,
        ["Leave"],
    )
    pre_sales = _worktype_total(
        selected_a,
        ["Pre Sales", "Pre-sales", "Presales"],
    )
    sickness = _worktype_total(
        selected_a,
        ["Sickness"],
    )

    completion_rate = (
        actual_reported / expected_capacity
        if expected_capacity
        else 0.0
    )

    arvi = (
        billable_productive / expected_capacity
        if expected_capacity
        else 0.0
    )

    arve_denominator = (
        expected_capacity
        - leave
        - sickness
    )

    arve = (
        billable_productive / arve_denominator
        if arve_denominator > 0
        else 0.0
    )

    # -----------------------------------------------------
    # Controls.
    #
    # Jul/Aug 2026 use validated historical controls.
    # Newer months are automatically derived from Source B.
    # -----------------------------------------------------
    controls = {
        month: _get_month_control(
            source_b,
            reporting_year,
            month,
        )
        for month in reporting_months
    }

    # Display the control population for the latest selected month.
    latest_month = max(reporting_months)
    hc_tagged = int(
        controls[latest_month]["hc_tagged"]
    )
    total_hc = int(
        controls[latest_month]["total_hc"]
    )

    compliance_numerator = sum(
        int(controls[m]["hc_compliant"])
        for m in reporting_months
    )
    compliance_denominator = sum(
        int(controls[m]["hc_tagged"])
        for m in reporting_months
    )

    hc_compliance = (
        compliance_numerator
        / compliance_denominator
        if compliance_denominator
        else 0.0
    )

    overbooked_count = sum(
        int(controls[m]["overbooked"])
        for m in reporting_months
    )

    # Preserve the historical dashboard convention:
    # overbooking occurrences across the selected period
    # divided by one HC-tagged population.
    overbooking = (
        overbooked_count / hc_tagged
        if hc_tagged
        else 0.0
    )

    expected_working_days = {
        _month_label(m):
            float(
                controls[m][
                    "expected_working_days"
                ]
            )
        for m in reporting_months
    }

    control_sources = {
        _month_label(m):
            controls[m]["control_source"]
        for m in reporting_months
    }

    metrics = {
        "reporting_year": reporting_year,
        "reporting_months": reporting_months,
        "reporting_period": _reporting_period(
            reporting_year,
            reporting_months,
        ),
        "kpis": {
            "hc_tagged": hc_tagged,
            "total_hc": total_hc,
            "completion_rate": float(completion_rate),
            "hc_compliance": float(hc_compliance),
            "arve": float(arve),
            "arvi": float(arvi),
            "overbooking": float(overbooking),
        },
        "worktypes": {
            "Billable": float(billable_productive),
            "General": float(general),
            "Leave": float(leave),
            "Pre-sales": float(pre_sales),
            "Total": float(actual_reported),
        },
        "trend": trend,
        "details": {
            "actual_reported": float(actual_reported),
            "expected_capacity": float(expected_capacity),
            "billable_productive_time": float(billable_productive),
            "leave": float(leave),
            "sickness": float(sickness),
            "hc_compliance_numerator": int(compliance_numerator),
            "hc_compliance_denominator": int(compliance_denominator),
            "overbooked_count": int(overbooked_count),
            "expected_working_days": expected_working_days,
            "control_sources": control_sources,
        },
    }

    return metrics


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
    else:
        output_path = Path(
            output_path
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
            ensure_ascii=False,
        )

    return str(output_path)
