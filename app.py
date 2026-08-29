import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ============================================================
# CG HELPLINE - TICKET INFLOW & RESOLUTION DASHBOARD
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

    return df


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
    ]

    avg_hours = (
        valid["Closure Hours"].mean()
        if not valid.empty
        else 0.0
    )

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
        "avg": float(avg_hours),
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

    x = data.copy()

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
    Returns top and low contributing Issue L1 categories.
    Low contributing = categories with the smallest ticket counts.
    """
    if data.empty:
        empty = pd.DataFrame(
            columns=[
                "Issue Type - L1",
                "Tickets",
                "% of Tickets"
            ]
        )
        return empty, empty

    x = (
        data.groupby("Issue L1")["Ticket ID"]
        .nunique()
        .reset_index(name="Tickets")
        .rename(columns={"Issue L1": "Issue Type - L1"})
    )

    total = x["Tickets"].sum()

    x["% of Tickets"] = (
        (x["Tickets"] / total * 100).round(1)
        if total
        else 0
    )

    top = x.sort_values(
        ["Tickets", "Issue Type - L1"],
        ascending=[False, True]
    ).head(10)

    low = x.sort_values(
        ["Tickets", "Issue Type - L1"],
        ascending=[True, True]
    ).head(10)

    return top.reset_index(drop=True), low.reset_index(drop=True)


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
    g = requested_group_counts(data)

    scope = filter_description(
        month,
        week,
        day,
        issue
    )

    # --------------------------------------------------------
    # Top issue
    # --------------------------------------------------------

    issue_summary = summary_by_group(
        data,
        "Issue L1"
    )

    top_issue_text = "No issue-type data available."

    if not issue_summary.empty:

        top = issue_summary.iloc[0]

        share = (
            top["Raised"] / s["raised"] * 100
            if s["raised"]
            else 0
        )

        top_issue_text = (
            f"{top['Issue L1']} had the highest inflow "
            f"with {int(top['Raised']):,} tickets "
            f"({share:.1f}% of total inflow)."
        )

    # --------------------------------------------------------
    # Open backlog
    # --------------------------------------------------------

    backlog_text = (
        f"There are {s['open']:,} open tickets."
    )

    if s["old_open"]:

        backlog_text += (
            f" {s['old_open']:,} open tickets are older than 72 hours."
        )

    # --------------------------------------------------------
    # SLA
    # --------------------------------------------------------

    if s["closed"]:

        sla_text = (
            f"24-hour closure performance is "
            f"{s['rate'] * 100:.1f}% "
            f"({s['within']:,} of {s['closed']:,} closed tickets "
            f"were closed within 24 hours)."
        )

    else:

        sla_text = (
            "No closed tickets are available for 24-hour SLA analysis."
        )

    # --------------------------------------------------------
    # Slow closures
    # --------------------------------------------------------

    slow_text = ""

    if s["after"]:

        slow_text = (
            f"{s['after']:,} closed tickets exceeded 24 hours."
        )

    else:

        slow_text = (
            "No closed tickets exceeded 24 hours."
        )

    # --------------------------------------------------------
    # Average closure
    # --------------------------------------------------------

    if s["closed"]:

        avg_text = (
            f"Average closure time is {s['avg']:.2f} hours."
        )

    else:

        avg_text = (
            "Average closure time is not available because there are no valid closed-ticket timestamps."
        )

    # --------------------------------------------------------
    # Date range
    # --------------------------------------------------------

    if data["Created time"].notna().any():

        min_date = data["Created time"].min()
        max_date = data["Created time"].max()

        date_text = (
            f"Ticket inflow period: "
            f"{min_date.strftime('%d %b %Y')} to "
            f"{max_date.strftime('%d %b %Y')}."
        )

    else:

        date_text = "Ticket inflow period could not be determined."

    # --------------------------------------------------------
    # Email
    # --------------------------------------------------------

    email = f"""Subject: CG Helpline Ticket Performance Update – {scope}

Hi Team,

Please find below the CG Helpline ticket performance summary for {scope}.

Key findings:
• Total tickets raised: {s['raised']:,}
• Tickets closed: {s['closed']:,}
• Tickets currently open: {s['open']:,}
• Closed within 24 hours: {s['within']:,}
• Closed after 24 hours: {s['after']:,}
• 24-hour closure rate: {s['rate'] * 100:.1f}%
• Average closure time: {s['avg']:.2f} hours
• Open tickets older than 72 hours: {s['old_open']:,}

Operational observations:
• {top_issue_text}
• {sla_text}
• {slow_text}
• {backlog_text}
• {avg_text}
• {date_text}

Recommended focus:
• Review the open tickets older than 72 hours and prioritize ageing cases.
• Review the issue types contributing the highest ticket inflow.
• Track tickets exceeding the 24-hour closure SLA and identify recurring causes.

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

    top_xl, low_xl = issue_contributors(
        filtered_data
    )

    contrib_ws["A4"] = "TOP CONTRIBUTING ISSUE TYPES"
    contrib_ws["A4"].font = Font(
        bold=True,
        color=white
    )
    contrib_ws["A4"].fill = PatternFill(
        "solid",
        fgColor=green
    )

    top_end_row, top_end_col = add_dataframe(
        contrib_ws,
        top_xl,
        5
    )

    format_table(
        contrib_ws,
        5,
        top_end_row,
        1,
        top_end_col,
        "TopIssuesTable"
    )

    start_low = top_end_row + 3

    contrib_ws.cell(
        start_low,
        1,
        "LOW CONTRIBUTING ISSUE TYPES"
    ).font = Font(
        bold=True,
        color=white
    )
    contrib_ws.cell(
        start_low,
        1
    ).fill = PatternFill(
        "solid",
        fgColor=orange
    )

    low_end_row, low_end_col = add_dataframe(
        contrib_ws,
        low_xl,
        start_low + 1
    )

    format_table(
        contrib_ws,
        start_low + 1,
        low_end_row,
        1,
        low_end_col,
        "LowIssuesTable"
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


df = st.session_state.data


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
    df[["Month", "Created time"]]
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

month_filtered = df.copy()

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
    df["Issue L1"]
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

filtered = df.copy()

if selected_month != "All":

    filtered = filtered[
        filtered["Month"]
        ==
        selected_month
    ]

if selected_week != "All":

    filtered = filtered[
        filtered["Week"]
        ==
        selected_week
    ]

if selected_day != "All":

    filtered = filtered[
        filtered["Day"]
        ==
        selected_day
    ]

if selected_issue != "All":

    filtered = filtered[
        filtered["Issue L1"]
        ==
        selected_issue
    ]


# ============================================================
# FILTER INFO
# ============================================================

st.markdown(
    f"""
    <div class="small-note">
    Showing <b>{len(filtered):,}</b> tickets | 
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
        f"{s['avg']:.2f}",
        "blue"
    )

with r2[3]:

    show_kpi(
        "OPEN >72H",
        f"{s['old_open']:,}",
        "red"
    )


# ============================================================
# CG / ESCALATION GROUP BREAKUP
# ============================================================

st.markdown('<div class="section-header">CG / ESCALATION GROUP BREAKUP</div>', unsafe_allow_html=True)

group_counts = requested_group_counts(filtered)

gc1, gc2, gc3 = st.columns(3)
with gc1:
    show_kpi("CG HELPLINE TICKETS", f"{group_counts['cg']:,}", "blue")
with gc2:
    show_kpi("CREDIT ESCALATION", f"{group_counts['credit']:,}", "orange")
with gc3:
    show_kpi("HIGH RETURNS / REATTEMPTS", f"{group_counts['returns']:,}", "red")

st.caption("Counts use unique Ticket ID and the Group recorded in the raw dump. If the export contains only the current Group, these are current group allocations rather than historical movement counts.")

group_summary = group_analysis(filtered)
if not group_summary.empty:
    gl, gr = st.columns([1.05, 0.95])
    with gl:
        st.dataframe(group_summary, use_container_width=True, hide_index=True)
    with gr:
        import plotly.express as px
        fig_group = px.bar(group_summary, x="Tickets", y="Group", orientation="h", text="Tickets", title="Tickets by Group")
        fig_group.update_layout(height=330, margin=dict(l=10,r=10,t=50,b=10))
        st.plotly_chart(fig_group, use_container_width=True)


# ============================================================
# ISSUE TYPE-WISE
# ============================================================

st.markdown(
    '<div class="section-header">ISSUE TYPE-WISE BREAKUP — L1 & L2</div>',
    unsafe_allow_html=True
)

issue_detail = issue_analysis(filtered)

if not issue_detail.empty:

    st.dataframe(
        issue_detail,
        use_container_width=True,
        hide_index=True
    )

    top_issues, low_issues = issue_contributors(
        filtered
    )

    st.markdown(
        '<div class="section-header">TOP & LOW CONTRIBUTING ISSUES</div>',
        unsafe_allow_html=True
    )

    ic1, ic2 = st.columns(2)

    with ic1:

        st.markdown("**🔥 Top contributing Issue Types - L1**")

        st.dataframe(
            top_issues,
            use_container_width=True,
            hide_index=True
        )

        if not top_issues.empty:

            import plotly.express as px

            fig_top = px.bar(
                top_issues.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type - L1",
                orientation="h",
                text="Tickets",
                title="Top Contributors"
            )

            fig_top.update_layout(
                height=380,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_top,
                use_container_width=True
            )

    with ic2:

        st.markdown("**Low contributing Issue Types - L1**")

        st.dataframe(
            low_issues,
            use_container_width=True,
            hide_index=True
        )

        if not low_issues.empty:

            import plotly.express as px

            fig_low = px.bar(
                low_issues.sort_values("Tickets"),
                x="Tickets",
                y="Issue Type - L1",
                orientation="h",
                text="Tickets",
                title="Low Contributors"
            )

            fig_low.update_layout(
                height=380,
                margin=dict(
                    l=10,
                    r=10,
                    t=50,
                    b=10
                )
            )

            st.plotly_chart(
                fig_low,
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
    "original raw data, filtered data, KPI dashboard, issue analysis, "
    "monthly/weekly/daily analysis and email-ready findings."
)

excel_filter_text = filter_description(
    selected_month,
    selected_week,
    selected_day,
    selected_issue
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
    "and adds analysis sheets without changing the original raw columns."
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
