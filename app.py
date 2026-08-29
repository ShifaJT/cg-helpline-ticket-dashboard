import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="CG Helpline Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>

.stApp {
    background-color: #f5f7fb;
}

.block-container {
    padding-top: 1.2rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1600px;
}

/* Header */

.dashboard-header {
    background: #17324d;
    color: white;
    padding: 20px 25px;
    border-radius: 8px;
    margin-bottom: 18px;
}

.dashboard-header h1 {
    margin: 0;
    font-size: 26px;
}

.dashboard-header p {
    margin: 5px 0 0 0;
    font-size: 13px;
    opacity: 0.8;
}

/* Section headings */

.section-header {
    color: #17324d;
    font-size: 18px;
    font-weight: 800;
    margin-top: 18px;
    margin-bottom: 10px;
    padding-bottom: 7px;
    border-bottom: 2px solid #dfe5eb;
}

/* Filter header */

.filter-header {
    background: #2f75b5;
    color: white;
    padding: 9px 14px;
    border-radius: 7px;
    font-weight: 700;
    margin-bottom: 10px;
}

/* KPI */

.kpi {
    color: white;
    padding: 15px;
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
    margin-top: 8px;
}

.kpi-blue {
    background: #2f75b5;
}

.kpi-green {
    background: #70ad47;
}

.kpi-red {
    background: #c00000;
}

.kpi-orange {
    background: #ed7d31;
}

/* Upload */

.upload-box {
    background: white;
    border: 1px solid #dfe5eb;
    border-radius: 9px;
    padding: 15px;
}

/* Dataframe */

div[data-testid="stDataFrame"] {
    border: 1px solid #dfe5eb;
}

/* Buttons */

.stButton > button {
    font-weight: 700;
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
    "Issue Type - L1"
]


# ============================================================
# FUNCTIONS
# ============================================================

def normalize_column_name(column):
    """
    Removes extra spaces from column names.
    """
    return str(column).strip()


def read_excel_file(uploaded_file):

    try:

        excel = pd.ExcelFile(uploaded_file)

        # Always use first sheet from raw dump
        df = pd.read_excel(
            uploaded_file,
            sheet_name=excel.sheet_names[0]
        )

        return df

    except Exception as e:

        st.error(f"Unable to read Excel file: {e}")

        return None


def read_uploaded_file(uploaded_file):

    if uploaded_file is None:
        return None

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):

        try:
            return pd.read_csv(uploaded_file)
        except Exception as e:
            st.error(f"Unable to read CSV file: {e}")
            return None

    elif filename.endswith(".xlsx") or filename.endswith(".xls"):

        return read_excel_file(uploaded_file)

    return None


def read_pasted_data(text):

    if not text.strip():
        return None

    try:

        # Excel copy/paste normally produces TAB separated data

        df = pd.read_csv(
            io.StringIO(text),
            sep="\t"
        )

        # If it wasn't tab separated,
        # try comma separated.

        if len(df.columns) <= 1:

            df = pd.read_csv(
                io.StringIO(text)
            )

        return df

    except Exception:

        try:

            df = pd.read_csv(
                io.StringIO(text),
                sep=","
            )

            return df

        except Exception as e:

            st.error(
                f"Unable to read pasted data: {e}"
            )

            return None


def prepare_data(df):

    if df is None:
        return None

    df = df.copy()

    # Clean headers

    df.columns = [
        normalize_column_name(c)
        for c in df.columns
    ]

    # Check required columns

    missing = [
        col
        for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing:

        st.error(
            "The following required columns are missing:\n\n"
            + ", ".join(missing)
        )

        return None

    # Remove completely empty rows

    df = df.dropna(
        how="all"
    ).copy()

    # --------------------------------------------------------
    # DATE COLUMNS
    # --------------------------------------------------------

    df["Created time"] = pd.to_datetime(
        df["Created time"],
        errors="coerce"
    )

    df["Closed time"] = pd.to_datetime(
        df["Closed time"],
        errors="coerce"
    )

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
    # CREATED DATE
    # --------------------------------------------------------

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
    #
    # Monday = beginning of week
    # --------------------------------------------------------

    df["Week Start"] = (
        df["Created Date"]
        - pd.to_timedelta(
            df["Created Date"].dt.weekday,
            unit="D"
        )
    )

    df["Week End"] = (
        df["Week Start"]
        + pd.Timedelta(days=6)
    )

    df["Week"] = (
        df["Week Start"].dt.strftime("%d %b")
        + " - "
        + df["Week End"].dt.strftime("%d %b %Y")
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
    #
    # Only calculate when:
    # Status = CLOSED
    # Created time exists
    # Closed time exists
    #
    # Negative values are converted to 0.
    # --------------------------------------------------------

    df["Closure Hours"] = np.nan

    closed_mask = (
        (df["Status Clean"] == "CLOSED")
        & df["Created time"].notna()
        & df["Closed time"].notna()
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

    # Never allow negative closure hours

    df["Closure Hours"] = (
        df["Closure Hours"]
        .clip(lower=0)
    )

    # --------------------------------------------------------
    # SLA
    # --------------------------------------------------------

    df["SLA"] = "Open"

    df.loc[
        closed_mask
        & (df["Closure Hours"] <= 24),
        "SLA"
    ] = "Within 24h"

    df.loc[
        closed_mask
        & (df["Closure Hours"] > 24),
        "SLA"
    ] = "After 24h"

    # --------------------------------------------------------
    # OPEN AGE
    #
    # For open tickets only.
    # --------------------------------------------------------

    now = pd.Timestamp.now()

    open_mask = (
        (df["Status Clean"] == "OPEN")
        & df["Created time"].notna()
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
            "rate": 0,
            "avg": 0,
            "old_open": 0
        }

    closed = data[
        data["Status Clean"] == "CLOSED"
    ]

    opened = data[
        data["Status Clean"] == "OPEN"
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

    valid_closure = closed[
        closed["Closure Hours"].notna()
    ]

    avg_hours = (
        valid_closure["Closure Hours"].mean()
        if not valid_closure.empty
        else 0
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
            if len(closed) > 0
            else 0
        ),
        "avg": avg_hours,
        "old_open": len(old_open)
    }


def show_kpi(label, value, colour):

    st.markdown(
        f"""
        <div class="kpi {colour}">
            <div class="kpi-label">
                {label}
            </div>

            <div class="kpi-value">
                {value}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def create_summary_table(
    data,
    group_column
):

    output = []

    for value in data[group_column].dropna().unique():

        subset = data[
            data[group_column] == value
        ]

        s = get_stats(subset)

        output.append({

            group_column: value,

            "Raised": s["raised"],

            "Closed": s["closed"],

            "Open": s["open"],

            "≤24h": s["within"],

            ">24h": s["after"],

            "24h %": round(
                s["rate"] * 100,
                1
            )
        })

    result = pd.DataFrame(output)

    return result


# ============================================================
# HEADER
# ============================================================

st.markdown("""
<div class="dashboard-header">

<h1>
CG HELPLINE — TICKET INFLOW & RESOLUTION DASHBOARD
</h1>

<p>
Paste your latest raw ticket dump or upload the Excel/CSV file.
</p>

</div>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

if "data" not in st.session_state:

    st.session_state.data = None


# ============================================================
# RAW DATA INPUT
# ============================================================

st.markdown(
    '<div class="section-header">RAW DATA</div>',
    unsafe_allow_html=True
)

upload_col, info_col = st.columns(
    [2, 1]
)

with upload_col:

    uploaded_file = st.file_uploader(
        "Upload your raw dump",
        type=[
            "xlsx",
            "xls",
            "csv"
        ]
    )

with info_col:

    st.info(
        "You can either upload the raw Excel file or copy the complete Excel table and paste it below."
    )


paste_data = st.text_area(
    "Paste complete raw dump",
    height=130,
    placeholder=(
        "Copy the complete raw data from Excel "
        "and paste it here, including the header row."
    )
)


load_button = st.button(
    "Load / Refresh Dashboard",
    type="primary"
)


# ============================================================
# LOAD DATA
# ============================================================

if load_button:

    raw_df = None

    # Uploaded file gets priority

    if uploaded_file is not None:

        raw_df = read_uploaded_file(
            uploaded_file
        )

    elif paste_data.strip():

        raw_df = read_pasted_data(
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
# STOP UNTIL DATA IS LOADED
# ============================================================

if st.session_state.data is None:

    st.warning(
        "Please upload or paste your raw ticket data."
    )

    st.stop()


df = st.session_state.data


# ============================================================
# DATA INFORMATION
# ============================================================

info1, info2, info3 = st.columns(3)

with info1:

    st.caption(
        f"Total raw tickets: {len(df):,}"
    )

with info2:

    st.caption(
        f"Date range: "
        f"{df['Created time'].min().strftime('%d %b %Y') if df['Created time'].notna().any() else 'N/A'}"
        f" → "
        f"{df['Created time'].max().strftime('%d %b %Y') if df['Created time'].notna().any() else 'N/A'}"
    )

with info3:

    st.caption(
        f"Issue types: {df['Issue L1'].nunique():,}"
    )


# ============================================================
# FILTERS
# ============================================================

st.markdown(
    '<div class="filter-header">FILTERS</div>',
    unsafe_allow_html=True
)

filter1, filter2, filter3, filter4 = st.columns(4)


# ------------------------------------------------------------
# MONTH
# ------------------------------------------------------------

all_months = sorted(
    df["Month"]
    .dropna()
    .unique(),
    key=lambda x: pd.to_datetime(
        "01 " + x
    )
)

with filter1:

    selected_month = st.selectbox(
        "Month",
        ["All"] + list(all_months)
    )


# ------------------------------------------------------------
# WEEK
# ------------------------------------------------------------

month_data = df.copy()

if selected_month != "All":

    month_data = month_data[
        month_data["Month"]
        == selected_month
    ]


week_values = (
    month_data[
        ["Week", "Week Start"]
    ]
    .drop_duplicates()
    .sort_values("Week Start")
    ["Week"]
    .tolist()
)

with filter2:

    selected_week = st.selectbox(
        "Week",
        ["All"] + week_values
    )


# ------------------------------------------------------------
# DAY
# ------------------------------------------------------------

week_data = month_data.copy()

if selected_week != "All":

    week_data = week_data[
        week_data["Week"]
        == selected_week
    ]


day_values = (
    week_data[
        ["Day", "Created Date"]
    ]
    .drop_duplicates()
    .sort_values("Created Date")
    ["Day"]
    .tolist()
)

with filter3:

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

with filter4:

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
        == selected_month
    ]


if selected_week != "All":

    filtered = filtered[
        filtered["Week"]
        == selected_week
    ]


if selected_day != "All":

    filtered = filtered[
        filtered["Day"]
        == selected_day
    ]


if selected_issue != "All":

    filtered = filtered[
        filtered["Issue L1"]
        == selected_issue
    ]


# ============================================================
# KPI
# ============================================================

s = get_stats(filtered)

st.markdown(
    '<div class="section-header">SUMMARY</div>',
    unsafe_allow_html=True
)


row1 = st.columns(4)

with row1[0]:

    show_kpi(
        "TICKETS RAISED",
        f"{s['raised']:,}",
        "kpi-blue"
    )

with row1[1]:

    show_kpi(
        "CLOSED",
        f"{s['closed']:,}",
        "kpi-green"
    )

with row1[2]:

    show_kpi(
        "OPEN",
        f"{s['open']:,}",
        "kpi-red"
    )

with row1[3]:

    show_kpi(
        "CLOSED ≤24H",
        f"{s['within']:,}",
        "kpi-green"
    )


st.write("")


row2 = st.columns(4)

with row2[0]:

    show_kpi(
        "CLOSED >24H",
        f"{s['after']:,}",
        "kpi-orange"
    )

with row2[1]:

    show_kpi(
        "24H CLOSURE %",
        f"{s['rate'] * 100:.1f}%",
        "kpi-green"
    )

with row2[2]:

    show_kpi(
        "AVG CLOSURE HRS",
        f"{s['avg']:.2f}",
        "kpi-blue"
    )

with row2[3]:

    show_kpi(
        "OPEN >72H",
        f"{s['old_open']:,}",
        "kpi-red"
    )


# ============================================================
# ISSUE TYPE BREAKUP
# ============================================================

st.markdown(
    '<div class="section-header">ISSUE TYPE-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

issue_summary = create_summary_table(
    filtered,
    "Issue L1"
)

if not issue_summary.empty:

    issue_summary = issue_summary.sort_values(
        "Raised",
        ascending=False
    )

    left, right = st.columns(
        [1.1, 0.9]
    )

    with left:

        st.dataframe(
            issue_summary,
            use_container_width=True,
            hide_index=True
        )

    with right:

        fig_issue = px.bar(
            issue_summary,
            x="Raised",
            y="Issue L1",
            orientation="h",
            title="Tickets by Issue Type",
            text="Raised"
        )

        fig_issue.update_layout(
            height=380,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig_issue,
            use_container_width=True
        )


# ============================================================
# MONTH-WISE
# ============================================================

st.markdown(
    '<div class="section-header">MONTH-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

month_summary = create_summary_table(
    filtered,
    "Month"
)

if not month_summary.empty:

    month_summary["_sort"] = month_summary[
        "Month"
    ].apply(
        lambda x: pd.to_datetime(
            "01 " + x
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

        fig_month = px.line(
            month_summary,
            x="Month",
            y="Raised",
            markers=True,
            title="Monthly Ticket Inflow"
        )

        fig_month.update_layout(
            height=350,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig_month,
            use_container_width=True
        )


# ============================================================
# WEEK-WISE
# ============================================================

st.markdown(
    '<div class="section-header">WEEK-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

week_summary = create_summary_table(
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
        week_order[
            ["Week"]
        ]
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

        fig_week = px.line(
            week_summary,
            x="Week",
            y="Raised",
            markers=True,
            title="Weekly Ticket Inflow"
        )

        fig_week.update_layout(
            height=350,
            margin=dict(
                l=10,
                r=10,
                t=50,
                b=10
            )
        )

        st.plotly_chart(
            fig_week,
            use_container_width=True
        )


# ============================================================
# DAY-WISE
# ============================================================

st.markdown(
    '<div class="section-header">DAY-WISE BREAKUP</div>',
    unsafe_allow_html=True
)

day_summary = create_summary_table(
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
# RAW DATA PREVIEW
# ============================================================

with st.expander(
    "View Raw Data"
):

    st.dataframe(
        df,
        use_container_width=True,
        height=400
    )


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.markdown(
    '<div class="section-header">EXPORT</div>',
    unsafe_allow_html=True
)

csv_data = filtered.to_csv(
    index=False
).encode("utf-8")

st.download_button(
    label="Download Filtered Ticket Data",
    data=csv_data,
    file_name="filtered_ticket_data.csv",
    mime="text/csv"
)
