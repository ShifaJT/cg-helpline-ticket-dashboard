import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ============================================================
# CG HELPLINE — TICKET INFLOW & RESOLUTION DASHBOARD
# Streamlit version
# ============================================================

st.set_page_config(
    page_title="CG Helpline Ticket Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ============================================================
# CSS - KEEP THE DASHBOARD LOOK CLOSE TO THE EXCEL VERSION
# ============================================================

st.markdown("""
<style>
.stApp {
    background-color: #f5f7fb;
}

.block-container {
    padding-top: 1rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1600px;
}

.dashboard-header {
    background: #17324d;
    color: white;
    padding: 18px 24px;
    border-radius: 8px;
    margin-bottom: 16px;
}

.dashboard-header h1 {
    margin: 0;
    font-size: 25px;
}

.dashboard-header p {
    margin: 5px 0 0 0;
    font-size: 13px;
    opacity: .82;
}

.filter-header {
    background: #2f75b5;
    color: white;
    padding: 9px 14px;
    border-radius: 7px;
    font-weight: 700;
    margin-bottom: 10px;
}

.section-header {
    color: #17324d;
    font-size: 17px;
    font-weight: 800;
    margin-top: 18px;
    margin-bottom: 10px;
    padding-bottom: 7px;
    border-bottom: 2px solid #dfe5eb;
}

.kpi-card {
    color: white;
    padding: 15px 16px;
    border-radius: 9px;
    min-height: 105px;
}

.kpi-label {
    font-size: 13px;
    font-weight: 700;
}

.kpi-value {
    font-size: 28px;
    font-weight: 800;
    margin-top: 7px;
}

.blue {
    background: #2f75b5;
}

.green {
    background: #70ad47;
}

.red {
    background: #c00000;
}

.orange {
    background: #ed7d31;
}

.small-note {
    font-size: 12px;
    color: #667085;
}

.email-box {
    background: white;
    border: 1px solid #d9e1e8;
    border-radius: 8px;
    padding: 8px;
}

div[data-testid="stDataFrame"] {
    border: 1px solid #dfe5eb;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# REQUIRED RAW COLUMNS
# ============================================================

REQUIRED_COLUMNS = [
    "Ticket ID",
    "Status",
    "Created time",
    "Closed time",
    "Issue Type - L1",
    "Issue Type - L2",
]


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_headers(df):
    """Clean only whitespace around headers. Do not rename raw columns."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_file(uploaded_file):
    """Read CSV/XLSX/XLS."""
    try:
        filename = uploaded_file.name.lower()

        if filename.endswith(".csv"):
            return pd.read_csv(uploaded_file)

        return pd.read_excel(uploaded_file)

    except Exception as e:
        st.error(f"Could not read the file: {e}")
        return None


def read_paste(text):
    """
    Read Excel copy/paste.
    Excel normally pastes tab-separated values.
    """

    if not text or not text.strip():
        return None

    try:
        df = pd.read_csv(
            io.StringIO(text),
            sep="\t",
            dtype=str,
            keep_default_na=False,
        )

        if len(df.columns) == 1:
            df = pd.read_csv(
                io.StringIO(text),
                dtype=str,
                keep_default_na=False,
            )

        return df

    except Exception as e:

        st.error(
            "Could not read the pasted data. "
            "Make sure you copied the complete Excel table including headers."
        )

        return None


def parse_datetime_series(series):
    """
    Robust datetime parser.
    Handles normal Excel/export timestamps and mixed values.
    """

    result = pd.to_datetime(
        series,
        errors="coerce",
        dayfirst=False,
    )

    # If too many values failed, try dayfirst.
    if result.notna().sum() < max(1, int(len(series) * 0.5)):

        result2 = pd.to_datetime(
            series,
            errors="coerce",
            dayfirst=True,
        )

        if result2.notna().sum() > result.notna().sum():
            result = result2

    return result


def prepare_data(raw):
    """
    Prepare dashboard fields without modifying the raw export columns.
    """

    if raw is None:
        return None

    df = clean_headers(raw)

    # Required column check
    missing = [
        c for c in REQUIRED_COLUMNS
        if c not in df.columns
    ]

    if missing:

        st.error(
            "Required columns are missing:\n\n"
            + ", ".join(missing)
        )

        return None

    # Remove completely blank rows
    df = df.dropna(how="all").copy()

    # Ticket ID is the safest row identifier.
    df["Ticket ID"] = (
        df["Ticket ID"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df = df[df["Ticket ID"] != ""].copy()

    # --------------------------------------------------------
    # STATUS
    # --------------------------------------------------------

    df["Status Clean"] = (
        df["Status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # --------------------------------------------------------
    # ISSUE TYPE
    # --------------------------------------------------------

    df["Issue L1"] = (
        df["Issue Type - L1"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    df.loc[
        df["Issue L1"] == "",
        "Issue L1"
    ] = "Unspecified"

    # Issue Type - L2 is kept exactly as supplied in Raw Data,
    # while Issue L2 is a clean internal field used for analysis.
    df["Issue L2"] = (
        df["Issue Type - L2"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    df.loc[
        df["Issue L2"] == "",
        "Issue L2"
    ] = "Unspecified"

    # --------------------------------------------------------
    # CREATED / CLOSED
    # --------------------------------------------------------

    df["Created time"] = parse_datetime_series(
        df["Created time"]
    )

    df["Closed time"] = parse_datetime_series(
        df["Closed time"]
    )

    df["Created Date"] = (
        df["Created time"]
        .dt.normalize()
    )

    # --------------------------------------------------------
    # MONTH
    # --------------------------------------------------------

    df["Month"] = (
        df["Created time"]
        .dt.strftime("%b %Y")
    )

    # --------------------------------------------------------
    # WEEK
    # Monday-Sunday
    # --------------------------------------------------------

    df["Week Start"] = (
        df["Created Date"]
        -
        pd.to_timedelta(
            df["Created Date"].dt.weekday,
            unit="D",
        )
    )

    df["Week End"] = (
        df["Week Start"]
        +
        pd.Timedelta(days=6)
    )

    df["Week"] = (
        df["Week Start"].dt.strftime("%d %b")
        +
        " - "
        +
        df["Week End"].dt.strftime("%d %b %Y")
    )

    # --------------------------------------------------------
    # DAY
    # --------------------------------------------------------

    df["Day"] = (
        df["Created Date"]
        .dt.strftime("%d %b %Y")
    )

    # --------------------------------------------------------
    # CLOSURE HOURS
    #
    # IMPORTANT:
    # Only CLOSED tickets with both timestamps are included.
    # Negative values are not allowed.
    # --------------------------------------------------------

    df["Closure Hours"] = np.nan

    closed_mask = (
        df["Status Clean"].eq("CLOSED")
        &
        df["Created time"].notna()
        &
        df["Closed time"].notna()
    )

    df.loc[
        closed_mask,
        "Closure Hours"
    ] = (
        (
            df.loc[closed_mask, "Closed time"]
            -
            df.loc[closed_mask, "Created time"]
        )
        .dt.total_seconds()
        / 3600
    )

    df["Closure Hours"] = (
        df["Closure Hours"]
        .clip(lower=0)
    )

    # --------------------------------------------------------
    # SLA BUCKET
    # --------------------------------------------------------

    df["SLA"] = "Open"

    df.loc[
        closed_mask
        &
        (df["Closure Hours"] <= 24),
        "SLA"
    ] = "Within 24h"

    df.loc[
        closed_mask
        &
        (df["Closure Hours"] > 24),
        "SLA"
    ] = "After 24h"

    # --------------------------------------------------------
    # OPEN AGE
    # --------------------------------------------------------

    now = pd.Timestamp.now()

    open_mask = (
        df["Status Clean"].eq("OPEN")
        &
        df["Created time"].notna()
    )

    df["Open Age Hours"] = np.nan

    df.loc[
        open_mask,
        "Open Age Hours"
    ] = (
        (
            now
            -
            df.loc[open_mask, "Created time"]
        )
        .dt.total_seconds()
        / 3600
    )

    df["Open Age Hours"] = (
        df["Open Age Hours"]
        .clip(lower=0)
    )

    df = add_fcr_fields(df)

    return df



def ensure_analysis_columns(data):
    """
    Backward-compatible safety layer for Streamlit session state.

    If the app code is updated while the browser session still contains
    an older prepared dataframe, rebuild the internal Issue L1/L2 fields
    from the actual raw columns instead of returning an empty table.
    """
    if data is None:
        return data

    data = data.copy()

    # Recreate Issue L1 from the raw field if necessary.
    if "Issue L1" not in data.columns:
        if "Issue Type - L1" in data.columns:
            data["Issue L1"] = (
                data["Issue Type - L1"]
                .fillna("Unspecified")
                .astype(str)
                .str.strip()
            )
        else:
            data["Issue L1"] = "Unspecified"

    data.loc[
        data["Issue L1"].eq(""),
        "Issue L1"
    ] = "Unspecified"

    # Recreate Issue L2 from the raw field if necessary.
    if "Issue L2" not in data.columns:
        if "Issue Type - L2" in data.columns:
            data["Issue L2"] = (
                data["Issue Type - L2"]
                .fillna("Unspecified")
                .astype(str)
                .str.strip()
            )
        else:
            data["Issue L2"] = "Unspecified"

    data.loc[
        data["Issue L2"].eq(""),
        "Issue L2"
    ] = "Unspecified"

    data = add_fcr_fields(data)

    return data




def add_fcr_fields(data):
    """
    FCR proxy based on the information available in the ticket dump.

    The raw dump does not contain a first-contact-resolution flag or a
    complete contact-history field. Therefore:
      FCR = CLOSED + Created date == Closed date.

    This is a SAME-DAY FCR PROXY, not a proven call-history FCR.
    It is intentionally transparent so the dashboard never reports a
    misleading zero simply because Call Type/Inbound Count are blank.
    """
    x = data.copy()

    if "Created time" not in x.columns:
        x["Created time"] = pd.NaT

    if "Closed time" not in x.columns:
        x["Closed time"] = pd.NaT

    closed = x["Status Clean"].eq("CLOSED")

    valid_times = (
        x["Created time"].notna()
        & x["Closed time"].notna()
    )

    same_day = (
        valid_times
        & x["Created time"].dt.normalize().eq(
            x["Closed time"].dt.normalize()
        )
    )

    x["Same Day Resolution"] = closed & same_day

    # Operational FCR proxy:
    # ticket was raised and resolved during the same calendar day.
    x["FCR"] = x["Same Day Resolution"]

    # Keep these fields for visibility/diagnostics.
    if "Call Type" not in x.columns:
        x["Call Type"] = ""

    if "Inbound Count" not in x.columns:
        x["Inbound Count"] = np.nan

    x["Call Type Clean"] = (
        x["Call Type"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    x["Inbound Count Num"] = pd.to_numeric(
        x["Inbound Count"],
        errors="coerce"
    )

    return x


def fcr_stats(data):
    if data is None or data.empty:
        return {
            "fcr": 0,
            "eligible": 0,
            "rate": 0.0,
            "same_day": 0,
            "coverage": 0.0,
        }

    x = add_fcr_fields(data)

    closed = x["Status Clean"].eq("CLOSED")
    valid_closed = (
        closed
        & x["Created time"].notna()
        & x["Closed time"].notna()
    )

    eligible_count = int(valid_closed.sum())
    fcr_count = int(x["FCR"].sum())
    same_day_count = int(x["Same Day Resolution"].sum())

    return {
        "fcr": fcr_count,
        "eligible": eligible_count,
        "rate": (
            fcr_count / eligible_count * 100
            if eligible_count else 0.0
        ),
        "same_day": same_day_count,
        "coverage": (
            valid_closed.sum() / closed.sum() * 100
            if closed.sum() else 0.0
        ),
    }



def fcr_period_summary(data, period_column):
    if data.empty:
        return pd.DataFrame(
            columns=[
                period_column,
                "Tickets Raised",
                "Closed",
                "FCR",
                "FCR %",
                "Same-Day Resolution",
            ]
        )

    rows = []

    for period, group in data.groupby(
        period_column,
        dropna=False,
        sort=False
    ):
        x = fcr_stats(group)
        rows.append({
            period_column: period,
            "Tickets Raised": group["Ticket ID"].nunique(),
            "Closed": int(
                group["Status Clean"].eq("CLOSED").sum()
            ),
            "FCR": x["fcr"],
            "FCR %": round(x["rate"], 1),
            "Same-Day Resolution": x["same_day"],
        })

    result = pd.DataFrame(rows)

    return result


def format_hms(hours):
    """Convert decimal hours to H:MM:SS."""
    if hours is None or pd.isna(hours):
        return "0:00:00"

    total_seconds = max(0, int(round(float(hours) * 3600)))

    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)

    return f"{h}:{m:02d}:{s:02d}"



def get_stats(data):

    if data.empty:

        return {
            "raised": 0,
            "closed": 0,
            "open": 0,
            "within": 0,
            "after": 0,
            "rate": 0.0,
            "avg": 0.0,
            "median": 0.0,
            "p90": 0.0,
            "p95": 0.0,
            "p99": 0.0,
            "old_open": 0,
        }

    closed = data[
        data["Status Clean"].eq("CLOSED")
    ]

    opened = data[
        data["Status Clean"].eq("OPEN")
    ]

    within = closed[
        closed["Closure Hours"].notna()
        &
        (closed["Closure Hours"] <= 24)
    ]

    after = closed[
        closed["Closure Hours"].notna()
        &
        (closed["Closure Hours"] > 24)
    ]

    valid = closed[
        closed["Closure Hours"].notna()
        &
        (closed["Closure Hours"] >= 0)
    ]

    if not valid.empty:
        avg_hours = float(valid["Closure Hours"].mean())
        median_hours = float(valid["Closure Hours"].median())
        p90_hours = float(valid["Closure Hours"].quantile(0.90))
        p95_hours = float(valid["Closure Hours"].quantile(0.95))
        p99_hours = float(valid["Closure Hours"].quantile(0.99))
    else:
        avg_hours = 0.0
        median_hours = 0.0
        p90_hours = 0.0
        p95_hours = 0.0
        p99_hours = 0.0

    old_open = opened[
        opened["Open Age Hours"] > 72
    ]

    return {
        "raised": len(data),
        "closed": len(closed),
        "open": len(opened),
        "within": len(within),
        "after": len(after),
        "rate": (
            len(within) / len(closed)
            if len(closed)
            else 0.0
        ),
        "avg": avg_hours,
        "median": median_hours,
        "p90": p90_hours,
        "p95": p95_hours,
        "p99": p99_hours,
        "old_open": len(old_open),
    }


def show_kpi(label, value, colour_class):

    colours = {
        "blue": "#2f75b5",
        "green": "#70ad47",
        "red": "#c00000",
        "orange": "#ed7d31",
    }

    bg = colours.get(
        colour_class,
        "#2f75b5"
    )

    # st.html prevents Streamlit from displaying the HTML itself.
    st.html(
        f"""
        <div style="
            background:{bg};
            color:white;
            padding:15px 16px;
            border-radius:9px;
            min-height:105px;
            width:100%;
        ">

            <div style="
                font-size:13px;
                font-weight:700;
            ">
                {label}
            </div>

            <div style="
                font-size:28px;
                font-weight:800;
                margin-top:7px;
            ">
                {value}
            </div>

        </div>
        """
    )


def summary_by_group(data, group_column):

    if data.empty:
        return pd.DataFrame(
            columns=[
                group_column,
                "Raised",
                "Closed",
                "Open",
                "≤24h",
                ">24h",
                "24h %",
            ]
        )

    rows = []

    for value, x in data.groupby(
        group_column,
        dropna=False,
        sort=False
    ):

        s = get_stats(x)

        rows.append({
            group_column: value,
            "Raised": s["raised"],
            "Closed": s["closed"],
            "Open": s["open"],
            "≤24h": s["within"],
            ">24h": s["after"],
            "24h %": round(
                s["rate"] * 100,
                1
            ),
        })

    return pd.DataFrame(rows)




def issue_analysis(data):
    """
    Issue Type L1 + L2 analysis using unique Ticket ID.
    """
    cols = [
        "Issue Type - L1",
        "Issue Type - L2",
        "Tickets",
        "% of Tickets",
        "Closed",
        "Open",
        "≤24h",
        ">24h",
        "24h %",
    ]

    if data.empty:
        return pd.DataFrame(columns=cols)

    x = ensure_analysis_columns(data)

    x["Issue L1 Analysis"] = (
        x["Issue Type - L1"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    x.loc[
        x["Issue L1 Analysis"].eq(""),
        "Issue L1 Analysis"
    ] = "Unspecified"

    x["Issue L2 Analysis"] = (
        x["Issue Type - L2"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
    )

    x.loc[
        x["Issue L2 Analysis"].eq(""),
        "Issue L2 Analysis"
    ] = "Unspecified"

    rows = []

    for (l1, l2), group in x.groupby(
        ["Issue L1 Analysis", "Issue L2 Analysis"],
        dropna=False,
        sort=False
    ):
        s = get_stats(group)

        rows.append({
            "Issue Type - L1": l1,
            "Issue Type - L2": l2,
            "Tickets": group["Ticket ID"].nunique(),
            "% of Tickets": 0,
            "Closed": s["closed"],
            "Open": s["open"],
            "≤24h": s["within"],
            ">24h": s["after"],
            "24h %": round(s["rate"] * 100, 1),
        })

    result = pd.DataFrame(rows)

    total = result["Tickets"].sum()

    if total:
        result["% of Tickets"] = (
            result["Tickets"] / total * 100
        ).round(1)

    return result.sort_values(
        ["Tickets", "Issue Type - L1", "Issue Type - L2"],
        ascending=[False, True, True]
    ).reset_index(drop=True)


def issue_contributors(data):
    """
    Returns separate Top/Low contributor tables for Issue Type L1 and L2.
    """
    empty = pd.DataFrame(
        columns=["Issue Type", "Tickets", "% of Tickets"]
    )

    if data.empty:
        return empty, empty, empty, empty

    data = ensure_analysis_columns(data)

    def contributor_table(df, column):
        if column not in df.columns:
            return (
                pd.DataFrame(columns=["Issue Type", "Tickets", "% of Tickets"]),
                pd.DataFrame(columns=["Issue Type", "Tickets", "% of Tickets"])
            )

        x = (
            df.groupby(column)["Ticket ID"]
            .nunique()
            .reset_index(name="Tickets")
            .rename(columns={column: "Issue Type"})
        )

        x["Issue Type"] = (
            x["Issue Type"]
            .fillna("Unspecified")
            .astype(str)
            .str.strip()
        )

        x.loc[
            x["Issue Type"].eq(""),
            "Issue Type"
        ] = "Unspecified"

        total = x["Tickets"].sum()

        x["% of Tickets"] = (
            (x["Tickets"] / total * 100).round(1)
            if total
            else 0
        )

        top = x.sort_values(
            ["Tickets", "Issue Type"],
            ascending=[False, True]
        ).head(10).reset_index(drop=True)

        low = x.sort_values(
            ["Tickets", "Issue Type"],
            ascending=[True, True]
        ).head(10).reset_index(drop=True)

        return top, low

    l1_top, l1_low = contributor_table(
        data,
        "Issue L1"
    )

    l2_top, l2_low = contributor_table(
        data,
        "Issue L2"
    )

    return l1_top, l1_low, l2_top, l2_low



def group_period_analysis(data, period_column):
    """
    Period-wise Group analysis:
    shows total tickets and the number recorded in CG Helpline,
    Credit Escalation and High Returns/Reattempts.
    """
    if data.empty:
        return pd.DataFrame()

    x = data.copy()

    x["Group Analysis"] = (
        x["Group"]
        .fillna("")
        .astype(str)
        .str.strip()
        .apply(group_bucket)
    )

    rows = []

    for period, g in x.groupby(
        period_column,
        dropna=False,
        sort=False
    ):
        total = g["Ticket ID"].nunique()

        cg = g.loc[
            g["Group Analysis"].eq("CG Helpline"),
            "Ticket ID"
        ].nunique()

        credit = g.loc[
            g["Group Analysis"].eq("Credit Escalation"),
            "Ticket ID"
        ].nunique()

        returns = g.loc[
            g["Group Analysis"].eq("High Returns/Reattempts"),
            "Ticket ID"
        ].nunique()

        other = max(
            total - cg - credit - returns,
            0
        )

        rows.append({
            period_column: period,
            "Tickets Raised": total,
            "CG Helpline": cg,
            "Credit Escalation": credit,
            "High Returns/Reattempts": returns,
            "Other Groups": other,
        })

    return pd.DataFrame(rows)



def group_bucket(value):
    text = str(value).strip()
    if text == "" or text.lower() == "nan":
        return "Unassigned"
    n = text.lower()
    if "credit" in n and "escalat" in n:
        return "Credit Escalation"
    if "high returns" in n and "reattempt" in n:
        return "High Returns/Reattempts"
    if "cg helpline" in n or n == "cg":
        return "CG Helpline"
    return text

def group_analysis(data):
    if data.empty or "Group" not in data.columns:
        return pd.DataFrame(columns=["Group", "Tickets", "% of Tickets"])
    x = data.copy()
    x["Group Analysis"] = x["Group"].fillna("").astype(str).str.strip().apply(group_bucket)
    result = (x.groupby("Group Analysis")["Ticket ID"].nunique().reset_index(name="Tickets")
              .rename(columns={"Group Analysis":"Group"})
              .sort_values("Tickets", ascending=False).reset_index(drop=True))
    total = result["Tickets"].sum()
    result["% of Tickets"] = ((result["Tickets"] / total * 100).round(1) if total else 0)
    return result

def requested_group_counts(data):
    if data.empty or "Group" not in data.columns:
        return {"cg":0,"credit":0,"returns":0}
    x=data.copy()
    x["Group Analysis"]=x["Group"].fillna("").astype(str).str.strip().apply(group_bucket)
    return {
        "cg": x.loc[x["Group Analysis"].eq("CG Helpline"),"Ticket ID"].nunique(),
        "credit": x.loc[x["Group Analysis"].eq("Credit Escalation"),"Ticket ID"].nunique(),
        "returns": x.loc[x["Group Analysis"].eq("High Returns/Reattempts"),"Ticket ID"].nunique(),
    }


def apply_dashboard_filters(
    data,
    selected_month,
    selected_week,
    selected_day,
    selected_issue
):
    x = data.copy()

    if selected_month != "All":
        x = x[x["Month"].eq(selected_month)]

    if selected_week != "All":
        x = x[x["Week"].eq(selected_week)]

    if selected_day != "All":
        x = x[x["Day"].eq(selected_day)]

    if selected_issue != "All":
        x = x[x["Issue L1"].eq(selected_issue)]

    return x


def split_current_group(data):
    """
    Main dashboard = current CG Helpline queue.
    Non-CG current groups are treated as moved-out tickets per the
    workflow described by the user.
    """
    x = data.copy()

    if "Group" in x.columns:
        x["Group Analysis"] = (
            x["Group"]
            .fillna("")
            .astype(str)
            .str.strip()
            .apply(group_bucket)
        )
    else:
        x["Group Analysis"] = "Unassigned"

    cg = x[x["Group Analysis"].eq("CG Helpline")].copy()
    moved = x[~x["Group Analysis"].eq("CG Helpline")].copy()

    return cg, moved


def moved_group_summary(data):
    if data.empty:
        return pd.DataFrame(
            columns=[
                "Destination Group",
                "Moved Tickets",
                "Closed",
                "Open",
                "Closed ≤24h",
                "Closed >24h",
                "24h %",
            ]
        )

    rows = []

    for group, g in data.groupby(
        "Group Analysis",
        dropna=False
    ):
        s = get_stats(g)

        rows.append({
            "Destination Group": group,
            "Moved Tickets": g["Ticket ID"].nunique(),
            "Closed": s["closed"],
            "Open": s["open"],
            "Closed ≤24h": s["within"],
            "Closed >24h": s["after"],
            "24h %": round(s["rate"] * 100, 1),
        })

    return (
        pd.DataFrame(rows)
        .sort_values("Moved Tickets", ascending=False)
        .reset_index(drop=True)
    )


def moved_period_summary(data, period_column):
    if data.empty:
        return pd.DataFrame(
            columns=[
                period_column,
                "Moved Tickets",
                "Closed",
                "Open",
                "Closed ≤24h",
                "Closed >24h",
                "24h %",
            ]
        )

    rows = []

    for period, g in data.groupby(
        period_column,
        dropna=False,
        sort=False
    ):
        s = get_stats(g)

        rows.append({
            period_column: period,
            "Moved Tickets": g["Ticket ID"].nunique(),
            "Closed": s["closed"],
            "Open": s["open"],
            "Closed ≤24h": s["within"],
            "Closed >24h": s["after"],
            "24h %": round(s["rate"] * 100, 1),
        })

    return pd.DataFrame(rows)



def filter_description(
    month,
    week,
    day,
    issue
):

    parts = []

    if month != "All":
        parts.append(f"Month: {month}")

    if week != "All":
        parts.append(f"Week: {week}")

    if day != "All":
        parts.append(f"Day: {day}")

    if issue != "All":
        parts.append(f"Issue Type: {issue}")

    if not parts:
        return "All tickets"

    return " | ".join(parts)


def make_email_summary(
    data,
    month,
    week,
    day,
    issue
):

    s = get_stats(data)

    scope = (
        "CG Helpline current queue | "
        + filter_description(
            month,
            week,
            day,
            issue
        )
    )

    issue_summary = summary_by_group(
        data,
        "Issue L1"
    )

    top_issue = "No issue-type data available"
    top_issue_count = 0
    top_issue_share = 0.0

    if not issue_summary.empty:
        top = issue_summary.iloc[0]
        top_issue = str(top["Issue L1"])
        top_issue_count = int(top["Raised"])
        top_issue_share = (
            top_issue_count / s["raised"] * 100
            if s["raised"] else 0
        )

    second_issue_text = ""

    if len(issue_summary) > 1:
        second = issue_summary.iloc[1]
        second_issue_text = (
            f"The second-highest contributing issue was "
            f"{second['Issue L1']} with {int(second['Raised']):,} tickets."
        )

    if s["closed"]:
        sla_text = (
            f"Out of {s['closed']:,} closed tickets, {s['within']:,} "
            f"were closed within 24 hours and {s['after']:,} exceeded "
            f"the 24-hour SLA. The resulting 24-hour closure rate was "
            f"{s['rate'] * 100:.1f}%."
        )

        resolution_text = (
            f"The average resolution time was {format_hms(s['avg'])} and "
            f"the median was {format_hms(s['median'])}. The 90% percentile "
            f"was {format_hms(s['p90'])}, meaning 90% of valid closed "
            f"tickets were resolved within this time. The 95% percentile "
            f"was {format_hms(s['p95'])}, meaning 95% of valid closed "
            f"tickets were resolved within this time. The 99% percentile "
            f"was {format_hms(s['p99'])}, meaning 99% were resolved "
            f"within this time. 90% and 99% percentile help highlight the long-tail "
            f"cases that take considerably longer to resolve."
        )
    else:
        sla_text = "There are no valid closed tickets for SLA analysis."
        resolution_text = (
            "Average, median, P90 and 99% Percentile resolution times are unavailable "
            "because there are no valid closed-ticket resolution times."
        )

    backlog_text = (
        f"There are {s['open']:,} tickets currently open."
    )

    if s["old_open"]:
        backlog_text += (
            f" {s['old_open']:,} of these are older than 72 hours "
            f"and should be prioritised."
        )
    else:
        backlog_text += (
            " There are no open tickets older than 72 hours."
        )

    if data["Created time"].notna().any():
        min_date = data["Created time"].min()
        max_date = data["Created time"].max()
        date_text = (
            f"The report covers ticket inflow from "
            f"{min_date.strftime('%d %b %Y')} to "
            f"{max_date.strftime('%d %b %Y')}."
        )
    else:
        date_text = "The reporting period could not be determined."

    # Copy/paste-friendly table for Outlook/Gmail.
    findings = [
        ("Tickets raised", f"{s['raised']:,}", "Total ticket inflow"),
        ("Tickets closed", f"{s['closed']:,}", "Tickets closed"),
        ("Tickets open", f"{s['open']:,}", "Current backlog"),
        ("Closed within 24h", f"{s['within']:,}", "Met 24-hour SLA"),
        ("Closed >24h", f"{s['after']:,}", "Exceeded 24-hour SLA"),
        ("24h closure rate", f"{s['rate'] * 100:.1f}%", "SLA adherence"),
        ("Average resolution", f"{s['avg']:.2f} hrs", "Average closure time"),
        ("Median resolution", f"{s['median']:.2f} hrs", "Middle resolution time"),
        ("90% Percentile resolution", f"{s['p90']:.2f} hrs", "90% of valid closed tickets resolved within this time"),
        ("95% Percentile resolution", f"{s['p95']:.2f} hrs", "95% of valid closed tickets resolved within this time"),
        ("99% Percentile resolution", f"{s['p99']:.2f} hrs", "99% of valid closed tickets resolved within this time"),
        ("Open >72h", f"{s['old_open']:,}", "Ageing backlog"),
    ]

    w1 = max(len("Metric"), max(len(x[0]) for x in findings)) + 3
    w2 = max(len("Value"), max(len(x[1]) for x in findings)) + 3

    table_lines = [
        f"{'Metric'.ljust(w1)}{'Value'.ljust(w2)}Comments",
        "-" * (w1 + w2 + 65),
    ]

    for metric, value, comment in findings:
        table_lines.append(
            f"{metric.ljust(w1)}{value.ljust(w2)}{comment}"
        )

    findings_table = "\n".join(table_lines)

    top_issue_text = (
        f"{top_issue} was the highest-contributing issue type, with "
        f"{top_issue_count:,} tickets, representing {top_issue_share:.1f}% "
        f"of total inflow."
        if top_issue_count
        else "No issue-type contribution could be identified."
    )

    email = f"""Subject: CG Helpline Ticket Performance Update – {scope}

Hi Team,

Please find below the CG Helpline ticket performance update for {scope}.

This update provides a concise view of ticket inflow, closure performance, SLA adherence and resolution-time distribution. The key findings are presented in a table for easier review.

KEY FINDINGS

{findings_table}

MANAGEMENT OBSERVATIONS

1. Ticket inflow
{top_issue_text}
{second_issue_text}

2. Closure and SLA performance
{sla_text}

3. Resolution-time distribution
{resolution_text}

4. Backlog
{backlog_text}

5. Reporting period
{date_text}

RECOMMENDED FOCUS

• Prioritise ageing open tickets and cases approaching the 72-hour threshold.
• Review tickets exceeding the 24-hour SLA and identify recurring operational or issue-type drivers.
• Review P90, P95 and 99% Percentile resolution times to understand the long-tail cases requiring significantly more time to resolve.
• Focus on the highest-contributing issue types for potential process, routing or knowledge-base improvements.

Overall, the key opportunity is to improve 24-hour SLA adherence while reducing the long-tail resolution time reflected in the 90%, 95% and 99% percentile metrics.

Regards,
CG Helpline
"""

    return email


# ============================================================
# EXCEL EXPORT
# ============================================================

def build_excel_report(
    raw_data,
    filtered_data,
    issue_summary,
    month_summary,
    week_summary,
    day_summary,
    email_text,
    filter_text
):
    """
    Creates a professional Excel report containing:
    1. Dashboard
    2. Raw Data - unchanged raw columns/values
    3. Filtered Raw Data
    4. Issue Type Analysis
    5. Monthly Analysis
    6. Weekly Analysis
    7. Daily Analysis
    8. Email Findings

    The raw-data sheet keeps the original raw export columns.
    Calculated helper columns are kept out of Raw Data.
    """

    output = io.BytesIO()

    wb = Workbook()

    # --------------------------------------------------------
    # Theme
    # --------------------------------------------------------

    navy = "17324D"
    blue = "2F75B5"
    green = "70AD47"
    red = "C00000"
    orange = "ED7D31"
    light_blue = "D9EAF7"
    light_green = "E2F0D9"
    light_red = "FCE4D6"
    light_orange = "FCE4D6"
    white = "FFFFFF"
    grey = "F2F4F7"
    dark_grey = "667085"

    thin_grey = Side(
        style="thin",
        color="D9E1E8"
    )

    # --------------------------------------------------------
    # Remove default sheet
    # --------------------------------------------------------

    ws = wb.active
    ws.title = "Dashboard"

    # --------------------------------------------------------
    # Helper functions
    # --------------------------------------------------------

    def style_title(sheet, title, subtitle=None, end_col=8):
        sheet.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=end_col
        )

        cell = sheet.cell(1, 1)
        cell.value = title
        cell.fill = PatternFill(
            "solid",
            fgColor=navy
        )
        cell.font = Font(
            color=white,
            bold=True,
            size=18
        )
        cell.alignment = Alignment(
            vertical="center"
        )

        sheet.row_dimensions[1].height = 30

        if subtitle:
            sheet.merge_cells(
                start_row=2,
                start_column=1,
                end_row=2,
                end_column=end_col
            )

            sub = sheet.cell(2, 1)
            sub.value = subtitle
            sub.font = Font(
                color=dark_grey,
                italic=True,
                size=10
            )

    def style_header_row(sheet, row, start_col, end_col):
        for col in range(start_col, end_col + 1):
            cell = sheet.cell(row, col)
            cell.fill = PatternFill(
                "solid",
                fgColor=navy
            )
            cell.font = Font(
                color=white,
                bold=True
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            cell.border = Border(
                bottom=thin_grey
            )

    def add_dataframe(sheet, dataframe, start_row=1, start_col=1):
        """
        Write dataframe with headers and return end row/end col.
        """
        data = dataframe.copy()

        # Convert timestamps to Excel-friendly values
        for col in data.columns:
            if pd.api.types.is_datetime64_any_dtype(data[col]):
                data[col] = data[col].dt.to_pydatetime()

        for j, col in enumerate(data.columns, start_col):
            sheet.cell(
                start_row,
                j,
                str(col)
            )

        style_header_row(
            sheet,
            start_row,
            start_col,
            start_col + len(data.columns) - 1
        )

        for i, row in enumerate(
            data.itertuples(index=False, name=None),
            start_row + 1
        ):
            for j, value in enumerate(
                row,
                start_col
            ):
                if pd.isna(value):
                    value = ""
                sheet.cell(
                    i,
                    j,
                    value
                )

        end_row = start_row + len(data)
        end_col = start_col + len(data.columns) - 1

        return end_row, end_col

    def format_table(sheet, start_row, end_row, start_col, end_col, name):
        if end_row < start_row + 1:
            return

        ref = (
            f"{get_column_letter(start_col)}{start_row}:"
            f"{get_column_letter(end_col)}{end_row}"
        )

        table = Table(
            displayName=name,
            ref=ref
        )

        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        table.tableStyleInfo = style
        sheet.add_table(table)

    def auto_width(sheet, max_width=32):
        for column_cells in sheet.columns:
            try:
                letter = get_column_letter(
                    column_cells[0].column
                )
            except Exception:
                continue

            max_len = 0

            for cell in column_cells:
                try:
                    value = str(cell.value)
                    max_len = max(
                        max_len,
                        len(value)
                    )
                except Exception:
                    pass

            sheet.column_dimensions[
                letter
            ].width = min(
                max(max_len + 2, 10),
                max_width
            )

    # --------------------------------------------------------
    # DASHBOARD SHEET
    # --------------------------------------------------------

    s = get_stats(filtered_data)

    style_title(
        ws,
        "CG HELPLINE — TICKET INFLOW & RESOLUTION REPORT",
        f"Scope: {filter_text}",
        end_col=8
    )

    # KPI labels
    kpis = [
        ("Tickets Raised", s["raised"], blue),
        ("Closed", s["closed"], green),
        ("Open", s["open"], red),
        ("Closed ≤24h", s["within"], green),
        ("Closed >24h", s["after"], orange),
        ("24h Closure %", f"{s['rate'] * 100:.1f}%", green),
        ("Avg Closure Hrs", f"{s['avg']:.2f}", blue),
        ("Open >72h", s["old_open"], red),
    ]

    positions = [
        (4, 1),
        (4, 3),
        (4, 5),
        (4, 7),
        (7, 1),
        (7, 3),
        (7, 5),
        (7, 7),
    ]

    for (label, value, colour), (row, col) in zip(
        kpis,
        positions
    ):
        ws.merge_cells(
            start_row=row,
            start_column=col,
            end_row=row,
            end_column=col + 1
        )
        ws.merge_cells(
            start_row=row + 1,
            start_column=col,
            end_row=row + 1,
            end_column=col + 1
        )

        label_cell = ws.cell(row, col)
        label_cell.value = label
        label_cell.fill = PatternFill(
            "solid",
            fgColor=colour
        )
        label_cell.font = Font(
            color=white,
            bold=True
        )
        label_cell.alignment = Alignment(
            horizontal="center"
        )

        value_cell = ws.cell(row + 1, col)
        value_cell.value = value
        value_cell.fill = PatternFill(
            "solid",
            fgColor=colour
        )
        value_cell.font = Font(
            color=white,
            bold=True,
            size=18
        )
        value_cell.alignment = Alignment(
            horizontal="center"
        )

    # Email findings on Dashboard
    ws["A10"] = "IMPORTANT FINDINGS / EMAIL READY"
    ws["A10"].fill = PatternFill(
        "solid",
        fgColor=navy
    )
    ws["A10"].font = Font(
        color=white,
        bold=True,
        size=12
    )

    ws.merge_cells(
        "A11:H23"
    )

    email_cell = ws["A11"]
    email_cell.value = email_text
    email_cell.alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )
    email_cell.fill = PatternFill(
        "solid",
        fgColor="FFFFFF"
    )

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 15

    # --------------------------------------------------------
    # RAW DATA
    # --------------------------------------------------------

    raw_ws = wb.create_sheet("Raw Data")

    # Keep original raw columns only.
    raw_export = raw_data.copy()

    # Remove internal calculated columns if present.
    internal_columns = [
        "Status Clean",
        "Issue L1",
        "Issue L2",
        "Created Date",
        "Month",
        "Week Start",
        "Week End",
        "Week",
        "Day",
        "Closure Hours",
        "SLA",
        "Open Age Hours",
    ]

    raw_export = raw_export.drop(
        columns=[
            c for c in internal_columns
            if c in raw_export.columns
        ],
        errors="ignore"
    )

    raw_end_row, raw_end_col = add_dataframe(
        raw_ws,
        raw_export,
        start_row=1
    )

    format_table(
        raw_ws,
        1,
        raw_end_row,
        1,
        raw_end_col,
        "RawDataTable"
    )

    raw_ws.freeze_panes = "A2"
    raw_ws.auto_filter.ref = (
        f"A1:{get_column_letter(raw_end_col)}{raw_end_row}"
    )

    auto_width(
        raw_ws,
        max_width=35
    )

    # --------------------------------------------------------
    # FILTERED RAW DATA
    # --------------------------------------------------------

    filtered_ws = wb.create_sheet(
        "Filtered Raw Data"
    )

    filtered_export = filtered_data.drop(
        columns=[
            c for c in internal_columns
            if c in filtered_data.columns
        ],
        errors="ignore"
    )

    end_row, end_col = add_dataframe(
        filtered_ws,
        filtered_export,
        1
    )

    format_table(
        filtered_ws,
        1,
        end_row,
        1,
        end_col,
        "FilteredRawTable"
    )

    filtered_ws.freeze_panes = "A2"

    auto_width(
        filtered_ws,
        max_width=35
    )

    # --------------------------------------------------------
    # ISSUE ANALYSIS
    # --------------------------------------------------------

    issue_ws = wb.create_sheet(
        "Issue Analysis"
    )

    style_title(
        issue_ws,
        "ISSUE TYPE-WISE ANALYSIS",
        filter_text,
        end_col=7
    )

    issue_end_row, issue_end_col = add_dataframe(
        issue_ws,
        issue_summary,
        start_row=4
    )

    format_table(
        issue_ws,
        4,
        issue_end_row,
        1,
        issue_end_col,
        "IssueAnalysisTable"
    )

    auto_width(
        issue_ws,
        max_width=32
    )

    # --------------------------------------------------------
    # GROUP ANALYSIS
    # --------------------------------------------------------

    group_ws = wb.create_sheet("Group Analysis")
    style_title(group_ws, "CG / ESCALATION GROUP ANALYSIS", filter_text, end_col=3)
    group_end_row, group_end_col = add_dataframe(group_ws, group_analysis(filtered_data), 4)
    format_table(group_ws, 4, group_end_row, 1, group_end_col, "GroupAnalysisTable")
    auto_width(group_ws, max_width=32)

    # --------------------------------------------------------
    # ISSUE L1 & L2 ANALYSIS
    # --------------------------------------------------------

    issue_detail_ws = wb.create_sheet(
        "Issue L1 & L2 Analysis"
    )

    style_title(
        issue_detail_ws,
        "ISSUE TYPE L1 & L2 ANALYSIS",
        filter_text,
        end_col=9
    )

    issue_detail_end_row, issue_detail_end_col = add_dataframe(
        issue_detail_ws,
        issue_analysis(filtered_data),
        4
    )

    format_table(
        issue_detail_ws,
        4,
        issue_detail_end_row,
        1,
        issue_detail_end_col,
        "IssueL1L2Table"
    )

    auto_width(
        issue_detail_ws,
        max_width=34
    )

    # --------------------------------------------------------
    # TOP / LOW CONTRIBUTORS
    # --------------------------------------------------------

    contrib_ws = wb.create_sheet(
        "Top Low Issues"
    )

    style_title(
        contrib_ws,
        "TOP & LOW CONTRIBUTING ISSUE TYPES",
        filter_text,
        end_col=6
    )

    l1_top_xl, l1_low_xl, l2_top_xl, l2_low_xl = issue_contributors(
        filtered_data
    )

    # L1 top
    contrib_ws["A4"] = "TOP CONTRIBUTING — ISSUE TYPE L1"
    contrib_ws["A4"].font = Font(bold=True, color=white)
    contrib_ws["A4"].fill = PatternFill("solid", fgColor=green)

    l1_top_end_row, l1_top_end_col = add_dataframe(
        contrib_ws,
        l1_top_xl,
        5
    )

    format_table(
        contrib_ws,
        5,
        l1_top_end_row,
        1,
        l1_top_end_col,
        "TopL1IssuesTable"
    )

    # L1 low
    start_l1_low = l1_top_end_row + 3

    contrib_ws.cell(
        start_l1_low,
        1,
        "LOW CONTRIBUTING — ISSUE TYPE L1"
    ).font = Font(bold=True, color=white)

    contrib_ws.cell(
        start_l1_low,
        1
    ).fill = PatternFill("solid", fgColor=orange)

    l1_low_end_row, l1_low_end_col = add_dataframe(
        contrib_ws,
        l1_low_xl,
        start_l1_low + 1
    )

    format_table(
        contrib_ws,
        start_l1_low + 1,
        l1_low_end_row,
        1,
        l1_low_end_col,
        "LowL1IssuesTable"
    )

    # L2 top
    start_l2_top = l1_low_end_row + 3

    contrib_ws.cell(
        start_l2_top,
        1,
        "TOP CONTRIBUTING — ISSUE TYPE L2"
    ).font = Font(bold=True, color=white)

    contrib_ws.cell(
        start_l2_top,
        1
    ).fill = PatternFill("solid", fgColor=green)

    l2_top_end_row, l2_top_end_col = add_dataframe(
        contrib_ws,
        l2_top_xl,
        start_l2_top + 1
    )

    format_table(
        contrib_ws,
        start_l2_top + 1,
        l2_top_end_row,
        1,
        l2_top_end_col,
        "TopL2IssuesTable"
    )

    # L2 low
    start_l2_low = l2_top_end_row + 3

    contrib_ws.cell(
        start_l2_low,
        1,
        "LOW CONTRIBUTING — ISSUE TYPE L2"
    ).font = Font(bold=True, color=white)

    contrib_ws.cell(
        start_l2_low,
        1
    ).fill = PatternFill("solid", fgColor=orange)

    l2_low_end_row, l2_low_end_col = add_dataframe(
        contrib_ws,
        l2_low_xl,
        start_l2_low + 1
    )

    format_table(
        contrib_ws,
        start_l2_low + 1,
        l2_low_end_row,
        1,
        l2_low_end_col,
        "LowL2IssuesTable"
    )

    auto_width(
        contrib_ws,
        max_width=34
    )

    # --------------------------------------------------------
    # MONTHLY GROUP ANALYSIS
    # --------------------------------------------------------

    monthly_group_ws = wb.create_sheet(
        "Monthly Group Analysis"
    )

    style_title(
        monthly_group_ws,
        "MONTH-WISE GROUP BREAKUP",
        filter_text,
        end_col=6
    )

    monthly_group_end_row, monthly_group_end_col = add_dataframe(
        monthly_group_ws,
        group_period_analysis(filtered_data, "Month"),
        4
    )

    format_table(
        monthly_group_ws,
        4,
        monthly_group_end_row,
        1,
        monthly_group_end_col,
        "MonthlyGroupTable"
    )

    auto_width(
        monthly_group_ws,
        max_width=32
    )

    # --------------------------------------------------------
    # WEEKLY GROUP ANALYSIS
    # --------------------------------------------------------

    weekly_group_ws = wb.create_sheet(
        "Weekly Group Analysis"
    )

    style_title(
        weekly_group_ws,
        "WEEK-WISE GROUP BREAKUP",
        filter_text,
        end_col=6
    )

    weekly_group_data = group_period_analysis(
        filtered_data,
        "Week"
    )

    if not weekly_group_data.empty and "Week Start" in filtered_data.columns:

        week_order_xl = (
            filtered_data[["Week", "Week Start"]]
            .drop_duplicates()
            .sort_values("Week Start")
            [["Week"]]
        )

        weekly_group_data = (
            week_order_xl
            .merge(
                weekly_group_data,
                on="Week",
                how="left"
            )
        )

    weekly_group_end_row, weekly_group_end_col = add_dataframe(
        weekly_group_ws,
        weekly_group_data,
        4
    )

    format_table(
        weekly_group_ws,
        4,
        weekly_group_end_row,
        1,
        weekly_group_end_col,
        "WeeklyGroupTable"
    )

    auto_width(
        weekly_group_ws,
        max_width=38
    )

    # --------------------------------------------------------
    # MONTHLY
    # --------------------------------------------------------

    month_ws = wb.create_sheet(
        "Monthly Analysis"
    )

    style_title(
        month_ws,
        "MONTH-WISE ANALYSIS",
        filter_text,
        end_col=7
    )

    month_end_row, month_end_col = add_dataframe(
        month_ws,
        month_summary,
        4
    )

    format_table(
        month_ws,
        4,
        month_end_row,
        1,
        month_end_col,
        "MonthlyAnalysisTable"
    )

    auto_width(
        month_ws,
        max_width=30
    )

    # --------------------------------------------------------
    # WEEKLY
    # --------------------------------------------------------

    week_ws = wb.create_sheet(
        "Weekly Analysis"
    )

    style_title(
        week_ws,
        "WEEK-WISE ANALYSIS",
        filter_text,
        end_col=7
    )

    week_end_row, week_end_col = add_dataframe(
        week_ws,
        week_summary,
        4
    )

    format_table(
        week_ws,
        4,
        week_end_row,
        1,
        week_end_col,
        "WeeklyAnalysisTable"
    )

    auto_width(
        week_ws,
        max_width=35
    )

    # --------------------------------------------------------
    # DAILY
    # --------------------------------------------------------

    day_ws = wb.create_sheet(
        "Daily Analysis"
    )

    style_title(
        day_ws,
        "DAY-WISE ANALYSIS",
        filter_text,
        end_col=8
    )

    day_end_row, day_end_col = add_dataframe(
        day_ws,
        day_summary,
        4
    )

    format_table(
        day_ws,
        4,
        day_end_row,
        1,
        day_end_col,
        "DailyAnalysisTable"
    )

    auto_width(
        day_ws,
        max_width=25
    )

    # --------------------------------------------------------
    # EMAIL FINDINGS
    # --------------------------------------------------------

    email_ws = wb.create_sheet(
        "Email Findings"
    )

    style_title(
        email_ws,
        "IMPORTANT FINDINGS — EMAIL READY",
        filter_text,
        end_col=8
    )

    email_ws.merge_cells(
        "A4:H30"
    )

    email_ws["A4"] = email_text
    email_ws["A4"].alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )
    email_ws["A4"].font = Font(
        size=11
    )

    for col in range(1, 9):
        email_ws.column_dimensions[
            get_column_letter(col)
        ].width = 18

    # --------------------------------------------------------
    # Conditional formatting-like colour coding
    # --------------------------------------------------------

    for sheet in [
        issue_ws,
        month_ws,
        week_ws,
        day_ws
    ]:

        for row in sheet.iter_rows():

            for cell in row:

                if cell.value == "Open":
                    cell.fill = PatternFill(
                        "solid",
                        fgColor="F4CCCC"
                    )

                elif cell.value == "≤24h":
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=light_green
                    )

                elif cell.value == ">24h":
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=light_orange
                    )

    # --------------------------------------------------------
    # Freeze panes
    # --------------------------------------------------------

    for sheet in wb.worksheets:

        if sheet.title not in [
            "Dashboard",
            "Email Findings"
        ]:
            sheet.freeze_panes = "A5" if sheet.title in [
                "Issue Analysis",
                "Monthly Analysis",
                "Weekly Analysis",
                "Daily Analysis"
            ] else "A2"

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    wb.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# HEADER
# ============================================================

st.html(
    """
    <div style="background:linear-gradient(135deg,#17324D 0%,#234F73 60%,#2F75B5 100%);color:white;padding:24px 30px;border-radius:12px;margin-bottom:18px;box-shadow:0 5px 18px rgba(23,50,77,.16);">
        <div style="font-size:12px;font-weight:800;letter-spacing:1.2px;opacity:.78;margin-bottom:6px;">CG HELPLINE • OPERATIONS ANALYTICS</div>
        <div style="font-size:30px;font-weight:800;line-height:1.15;letter-spacing:-.4px;">Ticket Inflow &amp; Resolution Dashboard</div>
        <div style="font-size:14px;margin-top:8px;opacity:.88;">Inflow • SLA performance • backlog ageing • issue mix • escalation flow</div>
    </div>
    """
)


# ============================================================
# SESSION STATE
# ============================================================

if "data" not in st.session_state:

    st.session_state.data = None


# ============================================================
# RAW DATA
# ============================================================

st.markdown(
    '<div class="section-header">RAW DATA</div>',
    unsafe_allow_html=True
)

upload_col, help_col = st.columns(
    [2, 1]
)

with upload_col:

    uploaded_file = st.file_uploader(
        "Upload your raw ticket dump",
        type=[
            "xlsx",
            "xls",
            "csv"
        ],
        help="Use the same raw export you currently use in Excel.",
    )

with help_col:

    st.info(
        "You can upload the raw file OR copy the entire Excel table and paste it below."
    )


paste_data = st.text_area(
    "Paste complete raw dump",
    height=130,
    placeholder=(
        "Copy the complete raw table from Excel, including headers, "
        "and paste it here."
    ),
)


load_button = st.button(
    "Load / Refresh Dashboard",
    type="primary",
)


# ============================================================
# LOAD
# ============================================================

if load_button:

    raw_df = None

    if uploaded_file is not None:

        raw_df = read_file(
            uploaded_file
        )

    elif paste_data.strip():

        raw_df = read_paste(
            paste_data
        )

    if raw_df is not None:

        prepared = prepare_data(
            raw_df
        )

        if prepared is not None:

            st.session_state.data = prepared

            st.success(
                f"Successfully loaded {len(prepared):,} tickets."
            )


# ============================================================
# WAIT FOR DATA
# ============================================================

if st.session_state.data is None:

    st.warning(
        "Please upload or paste your raw ticket dump."
    )

    st.stop()


df = ensure_analysis_columns(
    st.session_state.data
)

# Store the repaired dataframe back into the session so the same
# raw upload works even after the app code has been updated.
st.session_state.data = df

# ============================================================
# CURRENT GROUP SPLIT
# ============================================================

# MAIN DASHBOARD = ONLY CURRENT CG HELPLINE GROUP
# OTHER GROUPS = CURRENTLY WITH DESTINATION QUEUES
if "Group" in df.columns:
    df["_Group Clean"] = (
        df["Group"]
        .fillna("")
        .astype(str)
        .str.strip()
    )
else:
    df["_Group Clean"] = ""

cg_all = df[
    df["_Group Clean"].str.casefold().eq("cg helpline")
].copy()

other_group_all = df[
    ~df["_Group Clean"].str.casefold().eq("cg helpline")
    & df["_Group Clean"].ne("")
].copy()


# ============================================================
# DATA HEALTH CHECK
# ============================================================

invalid_created = df["Created time"].isna().sum()

if invalid_created:

    st.warning(
        f"{invalid_created:,} ticket(s) have an invalid/missing Created time "
        "and will not appear correctly in month/week/day analysis."
    )


# ============================================================
# FILTERS
# ============================================================

st.markdown(
    '<div class="filter-header">FILTERS</div>',
    unsafe_allow_html=True
)

f1, f2, f3, f4 = st.columns(4)


# ------------------------------------------------------------
# MONTH
# ------------------------------------------------------------

# Keep each month only once. The previous version used
# drop_duplicates() on both Month and Created time, which meant
# every different date in the same month created another "Jul 2026"
# entry in the dropdown.
month_values = (
    cg_all[["Month", "Created time"]]
    .dropna(subset=["Month", "Created time"])
    .sort_values("Created time")
    .drop_duplicates(subset=["Month"], keep="first")
    ["Month"]
    .tolist()
)

with f1:

    selected_month = st.selectbox(
        "Month",
        ["All"] + month_values
    )


# ------------------------------------------------------------
# WEEK
# ------------------------------------------------------------

month_filtered = cg_all.copy()

if selected_month != "All":

    month_filtered = month_filtered[
        month_filtered["Month"]
        ==
        selected_month
    ]

week_values = (
    month_filtered[
        ["Week", "Week Start"]
    ]
    .dropna(subset=["Week"])
    .drop_duplicates()
    .sort_values("Week Start")
    ["Week"]
    .tolist()
)

with f2:

    selected_week = st.selectbox(
        "Week",
        ["All"] + week_values
    )


# ------------------------------------------------------------
# DAY
# ------------------------------------------------------------

week_filtered = month_filtered.copy()

if selected_week != "All":

    week_filtered = week_filtered[
        week_filtered["Week"]
        ==
        selected_week
    ]

day_values = (
    week_filtered[
        ["Day", "Created Date"]
    ]
    .dropna(subset=["Day"])
    .drop_duplicates()
    .sort_values("Created Date")
    ["Day"]
    .tolist()
)

with f3:

    selected_day = st.selectbox(
        "Day",
        ["All"] + day_values
    )


# ------------------------------------------------------------
# ISSUE TYPE
# ------------------------------------------------------------

issue_values = sorted(
    cg_all["Issue L1"]
    .dropna()
    .unique()
)

with f4:

    selected_issue = st.selectbox(
        "Issue Type - L1",
        ["All"] + list(issue_values)
    )


# ============================================================
# APPLY FILTERS
# ============================================================

# Main dashboard is CURRENT CG HELPLINE only.
filtered = apply_dashboard_filters(
    cg_all,
    selected_month,
    selected_week,
    selected_day,
    selected_issue
)

# Separate view for tickets currently in other queues.
filtered_other_groups = apply_dashboard_filters(
    other_group_all,
    selected_month,
    selected_week,
    selected_day,
    selected_issue
)


# ============================================================
# FILTER INFO
# ============================================================

st.markdown(
    f"""
    <div class="small-note">
    Showing <b>{filtered["Ticket ID"].nunique():,}</b> current CG Helpline tickets | 
    {filter_description(selected_month, selected_week, selected_day, selected_issue)}
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SUMMARY / KPI
# ============================================================

s = get_stats(filtered)

st.markdown(
    '<div class="section-header">SUMMARY</div>',
    unsafe_allow_html=True
)

r1 = st.columns(4)

with r1[0]:
    show_kpi(
        "TICKETS RAISED",
        f"{s['raised']:,}",
        "blue"
    )

with r1[1]:
    show_kpi(
        "CLOSED",
        f"{s['closed']:,}",
        "green"
    )

with r1[2]:
    show_kpi(
        "OPEN",
        f"{s['open']:,}",
        "red"
    )

with r1[3]:
    show_kpi(
        "CLOSED ≤24H",
        f"{s['within']:,}",
        "green"
    )

st.write("")

r2 = st.columns(4)

with r2[0]:
    show_kpi(
        "CLOSED >24H",
        f"{s['after']:,}",
        "orange"
    )

with r2[1]:
    show_kpi(
        "24H CLOSURE %",
        f"{s['rate'] * 100:.1f}%",
        "green"
    )

with r2[2]:
    show_kpi(
        "AVG CLOSURE HRS",
        format_hms(s["avg"]),
        "blue"
    )

with r2[3]:
    show_kpi(
        "OPEN >72H",
        f"{s['old_open']:,}",
        "red"
    )

st.write("")

# Resolution-time distribution
# 90%/95%/99% percentile are calculated only from valid closed-ticket resolution hours.
r3 = st.columns(3)

with r3[0]:
    show_kpi(
        "90% PERCENTILE RESOLUTION HRS",
        format_hms(s["p90"]),
        "blue"
    )

with r3[1]:
    show_kpi(
        "95% PERCENTILE RESOLUTION HRS",
        format_hms(s["p95"]),
        "orange"
    )

with r3[2]:
    show_kpi(
        "99% PERCENTILE RESOLUTION HRS",
        format_hms(s["p99"]),
        "red"
    )

st.caption(
    "90%/95%/99% percentile show the resolution-time point within which "
    "90%/95%/99% of valid closed tickets were resolved."
)



# ============================================================
# FCR / FIRST CONTACT RESOLUTION
# ============================================================

st.markdown(
    '<div class="section-header">FCR / FIRST CONTACT RESOLUTION</div>',
    unsafe_allow_html=True
)

fcr = fcr_stats(filtered)

f1, f2, f3, f4 = st.columns(4)

with f1:
    show_kpi(
        "FCR TICKETS",
        f"{fcr['fcr']:,}",
        "blue"
    )

with f2:
    show_kpi(
        "FCR %",
        f"{fcr['rate']:.1f}%",
        "green"
    )

with f3:
    show_kpi(
        "FCR ELIGIBLE",
        f"{fcr['eligible']:,}",
        "orange"
    )

with f4:
    show_kpi(
        "SAME-DAY RESOLUTION",
        f"{fcr['same_day']:,}",
        "green"
    )

st.caption(
    "FCR operational proxy = Closed ticket + populated Call Type + Inbound Count ≤ 1. "
    "The raw export does not contain a dedicated FCR flag/history."
)

if fcr["eligible"] == 0:
    st.warning(
        "No closed tickets have valid Created and Closed timestamps, "
        "so the FCR proxy cannot be calculated."
    )

fcr_week = fcr_period_summary(filtered, "Week")

if not fcr_week.empty:

    st.markdown("**FCR by Week**")

    st.dataframe(
        fcr_week,
        use_container_width=True,
        hide_index=True
    )

# ============================================================
# CG HELPLINE — CURRENT QUEUE
# ============================================================

st.markdown(
    '<div class="section-header">CG HELPLINE — CURRENT QUEUE</div>',
    unsafe_allow_html=True
)

st.caption(
    "The main Summary, SLA, FCR and Issue Type analysis uses ONLY tickets "
    "whose current Group is CG Helpline."
)



# ============================================================
# CURRENTLY IN OTHER GROUPS
# ============================================================


st.markdown("### CURRENT GROUP DISTRIBUTION — RAW GROUP COLUMN")

group_distribution = (
    df.groupby("_Group Clean")["Ticket ID"]
    .nunique()
    .reset_index(name="Tickets")
    .rename(columns={"_Group Clean": "Current Group"})
    .sort_values("Tickets", ascending=False)
    .reset_index(drop=True)
)

st.dataframe(
    group_distribution,
    use_container_width=True,
    hide_index=True
)

st.markdown(
    '<div class="section-header">CURRENT QUEUE MOVEMENT VIEW</div>',
    unsafe_allow_html=True
)

st.caption(
    "This section uses ONLY the raw Group column. "
    "The main dashboard contains only Group = CG Helpline. "
    "Every ticket whose current Group is different from CG Helpline is "
    "shown here as being in an Other Queue. "
    "≤24h / >24h is measured from ticket creation."
)

other = filtered_other_groups.copy()

other_stats = get_stats(other)

other_total = other["Ticket ID"].nunique()

other_within = other_stats["within"]
other_after = other_stats["after"]
other_closed = other_stats["closed"]
other_open = other_stats["open"]

other_eligible = other_within + other_after

other_rate = (
    other_within / other_eligible * 100
    if other_eligible else 0
)

o1, o2, o3, o4 = st.columns(4)

with o1:
    show_kpi(
        "TICKETS IN OTHER QUEUES",
        f"{other_total:,}",
        "orange"
    )

with o2:
    show_kpi(
        "CLOSED",
        f"{other_closed:,}",
        "green"
    )

with o3:
    show_kpi(
        "OPEN",
        f"{other_open:,}",
        "red"
    )

with o4:
    show_kpi(
        "CLOSED ≤24H",
        f"{other_within:,}",
        "blue"
    )

o5, o6, o7 = st.columns(3)

with o5:
    show_kpi(
        "CLOSED >24H",
        f"{other_after:,}",
        "orange"
    )

with o6:
    show_kpi(
        "24H CLOSURE %",
        f"{other_rate:.1f}%",
        "green"
    )

with o7:
    show_kpi(
        "DESTINATION QUEUES",
        f"{other['_Group Clean'].nunique():,}"
        if "_Group Clean" in other.columns
        else "0",
        "blue"
    )

st.caption(
    "Queue Count = number of distinct non-CG Helpline values in the raw Group column. "
    "It is NOT the number of tickets. The ticket count is shown in TICKETS IN OTHER QUEUES."
)

st.markdown("### OTHER QUEUE-WISE BREAKDOWN")

if not other.empty:

    queue_rows = []

    for queue, g in other.groupby(
        "_Group Clean",
        dropna=False
    ):
        s = get_stats(g)

        queue_rows.append({
            "Current Queue": queue,
            "Tickets": g["Ticket ID"].nunique(),
            "Closed": s["closed"],
            "Open": s["open"],
            "Closed ≤24h": s["within"],
            "Closed >24h": s["after"],
            "24h Closure %": round(
                s["rate"] * 100,
                1
            ),
        })

    queue_df = (
        pd.DataFrame(queue_rows)
        .sort_values(
            "Tickets",
            ascending=False
        )
        .reset_index(drop=True)
    )

    q1, q2 = st.columns([1.05, 0.95])

    with q1:
        st.dataframe(
            queue_df,
            use_container_width=True,
            hide_index=True
        )

    with q2:
        import plotly.express as px

        fig_queue = px.bar(
            queue_df.sort_values("Tickets"),
            x="Tickets",
            y="Current Queue",
            orientation="h",
            text="Tickets",
            title="Tickets Currently in Other Queues"
        )

        fig_queue.update_layout(
            height=360,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig_queue,
            use_container_width=True
        )

else:
    st.info(
        "No tickets are currently in an Other Queue for the selected filters."
    )

st.markdown("### Other Queues — Month-wise")

other_month = summary_by_group(
    other,
    "Month"
)

if not other_month.empty:
    st.dataframe(
        other_month,
        use_container_width=True,
        hide_index=True
    )

st.markdown("### Other Queues — Week-wise")

other_week = summary_by_group(
    other,
    "Week"
)

if not other_week.empty:
    st.dataframe(
        other_week,
        use_container_width=True,
        hide_index=True
    )

st.markdown("### Other Queues — Day-wise")

other_day = summary_by_group(
    other,
    "Day"
)

if not other_day.empty:
    st.dataframe(
        other_day,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# ISSUE TYPE-WISE
# ============================================================

st.markdown(
    '<div class="section-header">ISSUE TYPE-WISE BREAKUP — L1 & L2</div>',
    unsafe_allow_html=True
)

issue_detail = issue_analysis(filtered)

if "Issue Type - L2" in filtered.columns:
    l2_source = (
        filtered["Issue Type - L2"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    if l2_source.eq("").all():
        st.warning(
            "Issue Type - L2 is blank in the selected raw data. "
            "The dashboard cannot create an L2 breakdown unless the raw dump contains L2 values."
        )

if not issue_detail.empty:

    st.dataframe(
        issue_detail,
        use_container_width=True,
        hide_index=True
    )

    l1_top, l1_low, l2_top, l2_low = issue_contributors(
        filtered
    )

    st.markdown(
        '<div class="section-header">TOP & LOW CONTRIBUTING ISSUES</div>',
        unsafe_allow_html=True
    )

    # -------------------------
    # L1 VIEW
    # -------------------------

    st.markdown("### Issue Type - L1")

    l1c1, l1c2 = st.columns(2)

    with l1c1:

        st.markdown("**Top contributing L1 issues**")

        st.dataframe(
            l1_top,
            use_container_width=True,
            hide_index=True
        )

        if not l1_top.empty:

            import plotly.express as px

            fig_l1_top = px.bar(
                l1_top.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type",
                orientation="h",
                text="Tickets",
                title="Top L1 Contributors"
            )

            fig_l1_top.update_layout(
                height=380,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_l1_top,
                use_container_width=True
            )

    with l1c2:

        st.markdown("**Low contributing L1 issues**")

        st.dataframe(
            l1_low,
            use_container_width=True,
            hide_index=True
        )

        if not l1_low.empty:

            import plotly.express as px

            fig_l1_low = px.bar(
                l1_low.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type",
                orientation="h",
                text="Tickets",
                title="Low L1 Contributors"
            )

            fig_l1_low.update_layout(
                height=380,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_l1_low,
                use_container_width=True
            )

    # -------------------------
    # L2 VIEW
    # -------------------------

    st.markdown("### Issue Type - L2")

    l2c1, l2c2 = st.columns(2)

    with l2c1:

        st.markdown("**Top contributing L2 issues**")

        st.dataframe(
            l2_top,
            use_container_width=True,
            hide_index=True
        )

        if not l2_top.empty:

            import plotly.express as px

            fig_l2_top = px.bar(
                l2_top.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type",
                orientation="h",
                text="Tickets",
                title="Top L2 Contributors"
            )

            fig_l2_top.update_layout(
                height=420,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_l2_top,
                use_container_width=True
            )

    with l2c2:

        st.markdown("**Low contributing L2 issues**")

        st.dataframe(
            l2_low,
            use_container_width=True,
            hide_index=True
        )

        if not l2_low.empty:

            import plotly.express as px

            fig_l2_low = px.bar(
                l2_low.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type",
                orientation="h",
                text="Tickets",
                title="Low L2 Contributors"
            )

            fig_l2_low.update_layout(
                height=420,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_l2_low,
                use_container_width=True
            )


# ============================================================
# MONTH-WISE
# ============================================================

st.markdown(
    '<div class="section-header">MONTH-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

month_summary = summary_by_group(
    filtered,
    "Month"
)

if not month_summary.empty:

    month_summary["_sort"] = (
        month_summary["Month"]
        .apply(
            lambda x: pd.to_datetime(
                "01 " + x
            )
        )
    )

    month_summary = (
        month_summary
        .sort_values("_sort")
        .drop(columns="_sort")
    )

    left, right = st.columns(
        [1.1, 0.9]
    )

    with left:

        st.dataframe(
            month_summary,
            use_container_width=True,
            hide_index=True
        )

    with right:

        fig = px.line(
            month_summary,
            x="Month",
            y="Raised",
            markers=True,
            title="Monthly Ticket Inflow"
        )

        fig.update_layout(
            height=350,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )



# ============================================================
# MONTH-WISE GROUP MOVEMENT / ALLOCATION
# ============================================================

st.markdown(
    '<div class="section-header">MONTH-WISE GROUP BREAKUP</div>',
    unsafe_allow_html=True
)

month_group = group_period_analysis(
    filtered,
    "Month"
)

if not month_group.empty:

    st.dataframe(
        month_group,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# WEEK-WISE
# ============================================================

st.markdown(
    '<div class="section-header">WEEK-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

week_summary = summary_by_group(
    filtered,
    "Week"
)

if not week_summary.empty:

    week_order = (
        filtered[
            ["Week", "Week Start"]
        ]
        .drop_duplicates()
        .sort_values("Week Start")
    )

    week_summary = (
        week_order[["Week"]]
        .merge(
            week_summary,
            on="Week",
            how="left"
        )
    )

    left, right = st.columns(
        [1.1, 0.9]
    )

    with left:

        st.dataframe(
            week_summary,
            use_container_width=True,
            hide_index=True
        )

    with right:

        fig = px.line(
            week_summary,
            x="Week",
            y="Raised",
            markers=True,
            title="Weekly Ticket Inflow"
        )

        fig.update_layout(
            height=350,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )



# ============================================================
# WEEK-WISE GROUP MOVEMENT / ALLOCATION
# ============================================================

st.markdown(
    '<div class="section-header">WEEK-WISE GROUP BREAKUP</div>',
    unsafe_allow_html=True
)

week_group = group_period_analysis(
    filtered,
    "Week"
)

if not week_group.empty:

    # Keep the same chronological order as the weekly analysis.
    if "Week Start" in filtered.columns:
        week_order = (
            filtered[["Week", "Week Start"]]
            .drop_duplicates()
            .sort_values("Week Start")
            [["Week"]]
        )

        week_group = (
            week_order
            .merge(
                week_group,
                on="Week",
                how="left"
            )
        )

    st.dataframe(
        week_group,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# DAY-WISE
# ============================================================

st.markdown(
    '<div class="section-header">DAY-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

day_summary = summary_by_group(
    filtered,
    "Day"
)

if not day_summary.empty:

    day_summary["_sort"] = (
        pd.to_datetime(
            day_summary["Day"]
        )
    )

    day_summary = (
        day_summary
        .sort_values("_sort")
        .drop(columns="_sort")
    )

    day_summary.insert(
        1,
        "Day Name",
        pd.to_datetime(
            day_summary["Day"]
        ).dt.strftime("%A")
    )

    st.dataframe(
        day_summary,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# EMAIL FINDINGS
# ============================================================

st.markdown(
    '<div class="section-header">📧 IMPORTANT FINDINGS — EMAIL READY</div>',
    unsafe_allow_html=True
)

st.caption(
    "This section automatically generates an email summary based on the current Month / Week / Day / Issue Type filters."
)

email_text = make_email_summary(
    filtered,
    selected_month,
    selected_week,
    selected_day,
    selected_issue
)

st.code(
    email_text,
    language=None
)


# ============================================================
# EXPORT
# ============================================================

st.markdown(
    '<div class="section-header">EXPORT</div>',
    unsafe_allow_html=True
)

csv_data = filtered.to_csv(
    index=False
).encode("utf-8")

st.download_button(
    "Download Filtered Ticket Data",
    data=csv_data,
    file_name="filtered_ticket_data.csv",
    mime="text/csv"
)


# ============================================================
# DOWNLOAD PROFESSIONAL EXCEL REPORT
# ============================================================

st.markdown(
    '<div class="section-header">📊 SHAREABLE EXCEL REPORT</div>',
    unsafe_allow_html=True
)

st.write(
    "Download a professionally formatted Excel report containing the "
    "original raw data, filterable analysis, KPI dashboard, L1/L2 issue "
    "breakup, top/low contributors, group analysis, and monthly/weekly/daily views."
)

excel_filter_text = filter_description(
    selected_month,
    selected_week,
    selected_day,
    selected_issue
)

# Legacy L1 summary used by the Excel dashboard sheet.
# The detailed L1 + L2 analysis is also exported separately.
issue_summary = summary_by_group(
    filtered,
    "Issue L1"
)

excel_bytes = build_excel_report(
    raw_data=df,
    filtered_data=filtered,
    issue_summary=issue_summary,
    month_summary=month_summary,
    week_summary=week_summary,
    day_summary=day_summary,
    email_text=email_text,
    filter_text=excel_filter_text,
)

st.download_button(
    label="⬇️ Download Color-Coded Excel Report",
    data=excel_bytes,
    file_name="CG_Helpline_Ticket_Analysis_Report.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    type="primary",
)

st.caption(
    "The Excel report keeps the raw export on a separate Raw Data sheet "
    "and adds filterable analysis sheets without changing the original raw columns."
)


# ============================================================
# RAW DATA PREVIEW
# ============================================================

with st.expander("View Raw Data"):

    st.dataframe(
        df,
        use_container_width=True,
        height=400
    )
