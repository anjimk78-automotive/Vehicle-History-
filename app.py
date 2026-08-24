"""
Sales Analysis Streamlit Application
=====================================
Data source: Sales Data Sheet (Year, Month, Item No., Item Description,
Customer Code, Customer Name, Quantity, Sales Amt, Gross Profit (currency),
Gross Profit, Gross Profit %, Column1, Zone)

Four sections, selectable from a sidebar on the RIGHT of the screen:
    1. Zone Wise Sale Analysis
    2. Item Wise Sales Analysis
    3. Sales % Contribution Analysis
    4. Sales % with Time Analysis

Run with:  streamlit run app.py
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from data_utils import (
    ALL_SALES_LABEL,
    ALL_ZONES_LABEL,
    SALES_TYPE_OPTIONS,
    TIME_FRAME_OPTIONS,
    PERIOD_COL,
    filter_by_sales_type,
    filter_by_zone,
    get_period_order,
    load_dataframe,
    load_dataframe_from_url,
    prepare_dataframe,
)

st.set_page_config(page_title="Sales Analysis", layout="wide")

# ---------------------------------------------------------------------------
# Push the native Streamlit sidebar to the RIGHT of the screen (CSS only).
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
        [data-testid="stAppViewContainer"] > .main {
            order: 1;
        }
        section[data-testid="stSidebar"] {
            order: 2;
            border-left: 1px solid rgba(49, 51, 63, 0.2);
            border-right: none;
        }
        [data-testid="stAppViewContainer"] {
            display: flex;
            flex-direction: row;
        }
        [data-testid="collapsedControl"] {
            left: auto;
            right: 0.5rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar: data source + navigation
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("📊 Sales Analysis")

    with st.expander("📁 Data Source", expanded="data" not in st.session_state):
        st.caption(
            "Paste your normal Google Sheets share link below (the sheet's "
            "general access must be set to 'Anyone with the link' - Viewer). "
            "Or upload a CSV/XLSX export instead."
        )
        sheet_url = st.text_input("Google Sheet link (share or edit URL)", value="")
        uploaded_file = st.file_uploader("...or upload CSV or XLSX", type=["csv", "xlsx", "xls"])
        load_clicked = st.button("Load / Refresh Data", use_container_width=True)

    st.markdown("---")
    st.subheader("Sections")
    section = st.radio(
        "Go to:",
        [
            "1. Zone Wise Sale Analysis",
            "2. Item Wise Sales Analysis",
            "3. Sales % Contribution Analysis",
            "4. Sales % with Time Analysis",
        ],
        label_visibility="collapsed",
    )

# ---------------------------------------------------------------------------
# Load data (persist across reruns via session_state)
# ---------------------------------------------------------------------------
if load_clicked or ("data" not in st.session_state):
    try:
        if sheet_url.strip():
            raw = load_dataframe_from_url(sheet_url.strip())
        elif uploaded_file is not None:
            raw = load_dataframe(uploaded_file)
        elif "data" in st.session_state:
            raw = None
        else:
            raw = None

        if raw is not None:
            st.session_state["data"] = prepare_dataframe(raw)
    except Exception as e:
        st.sidebar.error(f"Could not load data: {e}")

if "data" not in st.session_state:
    st.info(
        "👈 Upload your Sales Data (CSV/XLSX export of **Sheet1**) or paste a "
        "public Google Sheet CSV link in the sidebar to get started."
    )
    st.stop()

df = st.session_state["data"]

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def zone_options(frame: pd.DataFrame):
    return [ALL_ZONES_LABEL] + sorted(frame["Zone"].dropna().unique().tolist())


def line_chart_with_labels(plot_df, x_col, y_col, title, y_label, color_col=None, text_fmt=None):
    """Plotly line chart with the value shown at each point."""
    if text_fmt is None:
        text_fmt = lambda v: f"{v:,.0f}"
    plot_df = plot_df.copy()
    plot_df["_label"] = plot_df[y_col].apply(text_fmt)

    if color_col:
        fig = px.line(
            plot_df, x=x_col, y=y_col, color=color_col, markers=True,
            text="_label", title=title,
        )
    else:
        fig = px.line(
            plot_df, x=x_col, y=y_col, markers=True, text="_label", title=title,
        )
    fig.update_traces(textposition="top center")
    fig.update_layout(
        xaxis_title="Time Frame",
        yaxis_title=y_label,
        hovermode="x unified",
        legend_title_text="",
    )
    return fig


# ===========================================================================
# SECTION 1: Zone Wise Sale Analysis
# ===========================================================================
if section.startswith("1."):
    st.header("1. Zone Wise Sale Analysis")

    c1, c2, c3 = st.columns(3)
    with c1:
        sales_type = st.selectbox("Sales Type", SALES_TYPE_OPTIONS, key="s1_type")
    with c2:
        zone = st.selectbox("Zone", zone_options(df), key="s1_zone")
    with c3:
        time_frame = st.selectbox("Time Frame", TIME_FRAME_OPTIONS, key="s1_tf")

    filtered = filter_by_sales_type(df, sales_type)
    filtered = filter_by_zone(filtered, zone)

    if filtered.empty:
        st.warning("No data for this selection.")
    else:
        period_col = PERIOD_COL[time_frame]
        period_order = get_period_order(filtered, time_frame)

        agg = (
            filtered.groupby(period_col, as_index=False)["Sales Amt"]
            .sum()
            .rename(columns={"Sales Amt": "Sales"})
        )
        agg[period_col] = pd.Categorical(agg[period_col], categories=period_order, ordered=True)
        agg = agg.sort_values(period_col)

        title = f"Sales Trend — {sales_type} — {zone} ({time_frame})"
        fig = line_chart_with_labels(agg, period_col, "Sales", title, "Sales Amt")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("View underlying data"):
            st.dataframe(agg, use_container_width=True)


# ===========================================================================
# SECTION 2: Item Wise Sales Analysis
# ===========================================================================
elif section.startswith("2."):
    st.header("2. Item Wise Sales Analysis")

    c1, c2, c3 = st.columns(3)
    with c1:
        zone = st.selectbox("Zone", zone_options(df), key="s2_zone")
    with c2:
        sales_type = st.selectbox("Sales Type", SALES_TYPE_OPTIONS, key="s2_type")
    with c3:
        time_frame = st.selectbox("Time Frame", TIME_FRAME_OPTIONS, key="s2_tf")

    scoped = filter_by_zone(df, zone)
    scoped = filter_by_sales_type(scoped, sales_type)

    item_options = sorted(scoped["Item Description"].dropna().unique().tolist())
    if not item_options:
        st.warning("No items available for this Zone / Sales Type combination.")
    else:
        item_desc = st.selectbox("Item Description", item_options, key="s2_item")
        item_df = scoped[scoped["Item Description"] == item_desc]

        period_col = PERIOD_COL[time_frame]
        period_order = get_period_order(item_df, time_frame)

        agg = (
            item_df.groupby(period_col, as_index=False)["Sales Amt"]
            .sum()
            .rename(columns={"Sales Amt": "Sales"})
        )
        agg[period_col] = pd.Categorical(agg[period_col], categories=period_order, ordered=True)
        agg = agg.sort_values(period_col)

        title = f"Sales Trend — {item_desc} — {zone} ({time_frame})"
        fig = line_chart_with_labels(agg, period_col, "Sales", title, "Sales Amt")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("View underlying data"):
            st.dataframe(agg, use_container_width=True)


# ===========================================================================
# SECTION 3: Sales % Contribution Analysis
# ===========================================================================
elif section.startswith("3."):
    st.header("3. Sales Percentage Contribution Analysis")
    st.caption(
        "Each cell = a customer's Sales Amt ÷ total Sales Amt of the selected "
        "Zone & Sales Type, for that period. Columns sum to 100%."
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        sales_type = st.selectbox("Sales Type", SALES_TYPE_OPTIONS, key="s3_type")
    with c2:
        time_frame = st.selectbox("Time Frame", TIME_FRAME_OPTIONS, key="s3_tf")
    with c3:
        zone = st.selectbox("Zone", zone_options(df), key="s3_zone")

    scoped = filter_by_sales_type(df, sales_type)
    scoped = filter_by_zone(scoped, zone)

    if scoped.empty:
        st.warning("No data for this selection.")
    else:
        period_col = PERIOD_COL[time_frame]
        period_order = get_period_order(scoped, time_frame)

        cust_period = (
            scoped.groupby(["Customer Display", period_col], as_index=False)["Sales Amt"]
            .sum()
        )
        period_totals = scoped.groupby(period_col)["Sales Amt"].sum()

        pivot = cust_period.pivot(index="Customer Display", columns=period_col, values="Sales Amt").fillna(0.0)
        pivot = pivot.reindex(columns=period_order)

        pct_table = pivot.div(period_totals.reindex(period_order), axis=1) * 100
        pct_table = pct_table.round(2)

        # Sort customers by their average contribution, descending
        pct_table = pct_table.loc[pct_table.mean(axis=1).sort_values(ascending=False).index]

        st.dataframe(
            pct_table.style.format("{:.2f}%").background_gradient(cmap="Blues", axis=None),
            use_container_width=True,
        )

        total_row = pd.DataFrame(
            [pct_table.sum(axis=0)], index=["Total (check = 100%)"]
        )
        st.caption("Column totals (sanity check):")
        st.dataframe(total_row.style.format("{:.2f}%"), use_container_width=True)


# ===========================================================================
# SECTION 4: Sales % with Time Analysis
# ===========================================================================
elif section.startswith("4."):
    st.header("4. Sales Percentage with Time Analysis")
    st.caption(
        "Line = selected customer's Sales Amt ÷ total Sales Amt for the "
        "selected Sales Type & Zone, tracked over the selected Time Frame."
    )

    c1, c2 = st.columns(2)
    with c1:
        time_frame = st.selectbox("Time Frame", TIME_FRAME_OPTIONS, key="s4_tf")
    with c2:
        zone = st.selectbox("Zone", zone_options(df), key="s4_zone")

    c3, c4 = st.columns(2)
    with c3:
        sales_type = st.selectbox("Sales Type", SALES_TYPE_OPTIONS, key="s4_type")

    scoped_for_customers = filter_by_zone(df, zone)
    scoped_for_customers = filter_by_sales_type(scoped_for_customers, sales_type)
    customer_map = (
        scoped_for_customers[["Customer Code", "Customer Display"]]
        .drop_duplicates()
        .sort_values("Customer Display")
    )

    with c4:
        if customer_map.empty:
            st.warning("No customers for this Zone / Sales Type.")
            st.stop()
        customer_display = st.selectbox(
            "Customer Code", customer_map["Customer Display"].tolist(), key="s4_cust"
        )

    scoped = filter_by_sales_type(df, sales_type)
    scoped = filter_by_zone(scoped, zone)

    period_col = PERIOD_COL[time_frame]
    period_order = get_period_order(scoped, time_frame)

    period_totals = scoped.groupby(period_col)["Sales Amt"].sum().reindex(period_order).fillna(0.0)
    cust_scoped = scoped[scoped["Customer Display"] == customer_display]
    cust_by_period = cust_scoped.groupby(period_col)["Sales Amt"].sum().reindex(period_order).fillna(0.0)

    pct_series = (cust_by_period / period_totals.replace(0, pd.NA)) * 100
    pct_series = pct_series.fillna(0.0)

    plot_df = pd.DataFrame({
        period_col: period_order,
        "Contribution %": pct_series.values,
    })
    plot_df[period_col] = pd.Categorical(plot_df[period_col], categories=period_order, ordered=True)

    title = f"{customer_display} — % of {sales_type} Sales in {zone} ({time_frame})"
    fig = line_chart_with_labels(
        plot_df, period_col, "Contribution %", title, "Contribution %",
        text_fmt=lambda v: f"{v:.1f}%",
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("View underlying data"):
        st.dataframe(plot_df, use_container_width=True)
