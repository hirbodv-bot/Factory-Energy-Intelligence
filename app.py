
import io
from datetime import date

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Factory Energy Intelligence",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PROFESSIONAL STYLING
# ============================================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        .hero {
            padding: 1.5rem 1.7rem;
            border: 1px solid rgba(49, 51, 63, 0.16);
            border-radius: 18px;
            margin-bottom: 1.2rem;
            background: linear-gradient(
                135deg,
                rgba(255, 193, 7, 0.16),
                rgba(255, 255, 255, 0.02)
            );
        }

        .hero-title {
            font-size: 2.35rem;
            font-weight: 800;
            line-height: 1.1;
            margin-bottom: 0.45rem;
        }

        .hero-subtitle {
            font-size: 1.02rem;
            opacity: 0.78;
            max-width: 900px;
        }

        .section-title {
            font-size: 1.35rem;
            font-weight: 750;
            margin-top: 0.35rem;
            margin-bottom: 0.5rem;
        }

        .insight-card {
            border: 1px solid rgba(49, 51, 63, 0.16);
            border-radius: 14px;
            padding: 1rem 1.1rem;
            margin-bottom: 0.75rem;
            background: rgba(127, 127, 127, 0.035);
        }

        .small-note {
            opacity: 0.72;
            font-size: 0.86rem;
        }

        [data-testid="stMetric"] {
            border: 1px solid rgba(49, 51, 63, 0.14);
            padding: 0.8rem 1rem;
            border-radius: 14px;
            background: rgba(127, 127, 127, 0.025);
        }

        [data-testid="stSidebar"] {
            border-right: 1px solid rgba(49, 51, 63, 0.12);
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# EXPECTED DATA STRUCTURE
# ============================================================
REQUIRED_COLUMNS = [
    "Date",
    "Production Output (units)",
    "Electricity Consumption (kWh)",
    "Operating Hours",
    "Downtime (hours)",
    "Defective Units",
    "Electricity Tariff (£/kWh)",
]


# ============================================================
# SAMPLE DATA
# ============================================================
@st.cache_data
def create_demo_data() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Date": pd.date_range("2026-06-01", periods=14, freq="D"),
            "Production Output (units)": [
                1180, 1240, 900, 1280, 1210, 1050, 1320,
                1270, 980, 1350, 1290, 1100, 1380, 1330
            ],
            "Electricity Consumption (kWh)": [
                2740, 2810, 2680, 2860, 2790, 2600, 2920,
                2880, 2710, 2980, 2900, 2750, 3010, 2960
            ],
            "Operating Hours": [
                8, 8, 8, 8, 8, 7.5, 8,
                8, 8, 8, 8, 8, 8, 8
            ],
            "Downtime (hours)": [
                0.4, 0.3, 1.8, 0.2, 0.5, 0.8, 0.1,
                0.3, 1.3, 0.2, 0.3, 0.9, 0.1, 0.2
            ],
            "Defective Units": [
                22, 18, 31, 16, 20, 27, 14,
                17, 29, 13, 15, 24, 11, 12
            ],
            "Electricity Tariff (£/kWh)": [0.25] * 14,
        }
    )


def dataframe_to_excel_bytes(
    calculated_data: pd.DataFrame,
    summary_data: pd.DataFrame,
    insights_data: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        calculated_data.to_excel(
            writer, sheet_name="Calculated_Data", index=False
        )
        summary_data.to_excel(
            writer, sheet_name="Scenario_Summary", index=False
        )
        insights_data.to_excel(
            writer, sheet_name="Engineering_Insights", index=False
        )
    return output.getvalue()


def sample_excel_bytes() -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        create_demo_data().to_excel(
            writer, sheet_name="Factory_Data", index=False
        )
    return output.getvalue()


# ============================================================
# DATA PROCESSING
# ============================================================
def validate_and_clean(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_columns:
        raise ValueError(
            "Missing required columns: " + ", ".join(missing_columns)
        )

    cleaned = df.copy()
    warnings: list[str] = []

    original_rows = len(cleaned)

    cleaned["Date"] = pd.to_datetime(cleaned["Date"], errors="coerce")

    numeric_columns = [c for c in REQUIRED_COLUMNS if c != "Date"]
    for column in numeric_columns:
        cleaned[column] = pd.to_numeric(cleaned[column], errors="coerce")

    invalid_rows = cleaned[REQUIRED_COLUMNS].isna().any(axis=1).sum()
    if invalid_rows:
        warnings.append(
            f"{invalid_rows} row(s) containing missing or invalid values were removed."
        )

    cleaned = cleaned.dropna(subset=REQUIRED_COLUMNS)

    non_positive_mask = (
        (cleaned["Production Output (units)"] <= 0)
        | (cleaned["Electricity Consumption (kWh)"] < 0)
        | (cleaned["Operating Hours"] <= 0)
        | (cleaned["Electricity Tariff (£/kWh)"] < 0)
    )

    non_positive_count = int(non_positive_mask.sum())
    if non_positive_count:
        warnings.append(
            f"{non_positive_count} row(s) containing non-physical values were removed."
        )

    cleaned = cleaned.loc[~non_positive_mask].copy()

    duplicate_count = int(cleaned.duplicated().sum())
    if duplicate_count:
        warnings.append(
            f"{duplicate_count} duplicate row(s) were removed."
        )
        cleaned = cleaned.drop_duplicates()

    cleaned = cleaned.sort_values("Date").reset_index(drop=True)

    if len(cleaned) < original_rows and not warnings:
        warnings.append(
            f"{original_rows - len(cleaned)} row(s) were removed during cleaning."
        )

    if cleaned.empty:
        raise ValueError("No valid records remain after data cleaning.")

    return cleaned, warnings


def calculate_metrics(
    df: pd.DataFrame,
    scenario_tariff: float,
    reduction_pct: float,
    operating_days_month: int,
    carbon_factor: float,
) -> tuple[pd.DataFrame, dict]:
    calculated = df.copy()

    calculated["Daily Electricity Cost (£)"] = (
        calculated["Electricity Consumption (kWh)"]
        * calculated["Electricity Tariff (£/kWh)"]
    )

    calculated["Energy per Product (kWh/unit)"] = (
        calculated["Electricity Consumption (kWh)"]
        / calculated["Production Output (units)"]
    )

    calculated["Electricity Cost per Product (£/unit)"] = (
        calculated["Daily Electricity Cost (£)"]
        / calculated["Production Output (units)"]
    )

    calculated["Downtime (%)"] = (
        calculated["Downtime (hours)"]
        / calculated["Operating Hours"]
        * 100
    )

    calculated["Defect Rate (%)"] = (
        calculated["Defective Units"]
        / calculated["Production Output (units)"]
        * 100
    )

    calculated["Production Rate (units/hour)"] = (
        calculated["Production Output (units)"]
        / calculated["Operating Hours"]
    )

    calculated["CO₂ Emissions (kgCO₂e)"] = (
        calculated["Electricity Consumption (kWh)"]
        * carbon_factor
    )

    energy_intensity_mean = calculated[
        "Energy per Product (kWh/unit)"
    ].mean()
    energy_intensity_std = calculated[
        "Energy per Product (kWh/unit)"
    ].std(ddof=0)

    threshold = energy_intensity_mean + 1.5 * energy_intensity_std
    calculated["Energy Anomaly"] = (
        calculated["Energy per Product (kWh/unit)"] > threshold
    )

    total_energy = calculated["Electricity Consumption (kWh)"].sum()
    total_production = calculated["Production Output (units)"].sum()
    total_cost = calculated["Daily Electricity Cost (£)"].sum()
    total_emissions = calculated["CO₂ Emissions (kgCO₂e)"].sum()

    sample_days = len(calculated)
    avg_daily_energy = total_energy / sample_days
    avg_daily_cost = total_cost / sample_days

    current_monthly_cost = avg_daily_cost * operating_days_month
    scenario_monthly_cost = (
        avg_daily_energy * scenario_tariff * operating_days_month
    )
    saving_monthly = scenario_monthly_cost * reduction_pct / 100
    reduced_monthly_cost = scenario_monthly_cost - saving_monthly
    annual_saving = saving_monthly * 12

    worst_index = calculated[
        "Energy per Product (kWh/unit)"
    ].idxmax()
    best_index = calculated[
        "Energy per Product (kWh/unit)"
    ].idxmin()

    summary = {
        "total_energy": total_energy,
        "total_production": total_production,
        "total_cost": total_cost,
        "total_emissions": total_emissions,
        "energy_per_product": total_energy / total_production,
        "cost_per_product": total_cost / total_production,
        "average_downtime": calculated["Downtime (%)"].mean(),
        "average_defect_rate": calculated["Defect Rate (%)"].mean(),
        "current_monthly_cost": current_monthly_cost,
        "scenario_monthly_cost": scenario_monthly_cost,
        "reduced_monthly_cost": reduced_monthly_cost,
        "monthly_saving": saving_monthly,
        "annual_saving": annual_saving,
        "worst_day": calculated.loc[worst_index, "Date"],
        "worst_energy_intensity": calculated.loc[
            worst_index, "Energy per Product (kWh/unit)"
        ],
        "best_day": calculated.loc[best_index, "Date"],
        "best_energy_intensity": calculated.loc[
            best_index, "Energy per Product (kWh/unit)"
        ],
        "anomaly_count": int(calculated["Energy Anomaly"].sum()),
        "average_production_rate": calculated[
            "Production Rate (units/hour)"
        ].mean(),
    }

    return calculated, summary


def generate_engineering_insights(
    df: pd.DataFrame,
    summary: dict,
) -> list[dict]:
    insights: list[dict] = []

    worst_row = df.loc[
        df["Energy per Product (kWh/unit)"].idxmax()
    ]

    correlation_downtime_energy = df[
        ["Downtime (%)", "Energy per Product (kWh/unit)"]
    ].corr().iloc[0, 1]

    correlation_production_energy = df[
        ["Production Output (units)", "Energy per Product (kWh/unit)"]
    ].corr().iloc[0, 1]

    insights.append(
        {
            "Priority": "High",
            "Finding": "Least energy-efficient operating day",
            "Evidence": (
                f"{summary['worst_day'].strftime('%d %B %Y')} recorded "
                f"{summary['worst_energy_intensity']:.2f} kWh/unit, "
                f"with {worst_row['Downtime (%)']:.1f}% downtime."
            ),
            "Recommended action": (
                "Review machine idle time, maintenance events, operating settings, "
                "and production interruptions for this date."
            ),
        }
    )

    if summary["anomaly_count"] > 0:
        anomaly_dates = ", ".join(
            df.loc[df["Energy Anomaly"], "Date"]
            .dt.strftime("%d %b")
            .tolist()
        )
        insights.append(
            {
                "Priority": "High",
                "Finding": "Potential energy anomalies detected",
                "Evidence": (
                    f"{summary['anomaly_count']} operating day(s) exceeded "
                    f"the statistical energy-intensity threshold: {anomaly_dates}."
                ),
                "Recommended action": (
                    "Compare these dates with maintenance logs, production changes, "
                    "equipment alarms, and shift records."
                ),
            }
        )

    if correlation_downtime_energy > 0.45:
        insights.append(
            {
                "Priority": "Medium",
                "Finding": "Energy intensity rises with downtime",
                "Evidence": (
                    f"The observed correlation between downtime and energy intensity "
                    f"is {correlation_downtime_energy:.2f}."
                ),
                "Recommended action": (
                    "Investigate whether equipment continues consuming significant "
                    "power during stoppages and consider idle-energy controls."
                ),
            }
        )

    if correlation_production_energy < -0.45:
        insights.append(
            {
                "Priority": "Medium",
                "Finding": "Lower output is associated with higher energy per unit",
                "Evidence": (
                    f"The observed correlation between production output and energy "
                    f"intensity is {correlation_production_energy:.2f}."
                ),
                "Recommended action": (
                    "Review production scheduling, minimum efficient batch sizes, "
                    "and base-load energy consumption."
                ),
            }
        )

    if summary["average_defect_rate"] > 2:
        insights.append(
            {
                "Priority": "Medium",
                "Finding": "Defect rate may be increasing embodied energy waste",
                "Evidence": (
                    f"The average defect rate is "
                    f"{summary['average_defect_rate']:.2f}%."
                ),
                "Recommended action": (
                    "Assess process settings, material quality, operator variation, "
                    "and inspection records to reduce wasted production energy."
                ),
            }
        )

    insights.append(
        {
            "Priority": "Opportunity",
            "Finding": "Energy-reduction scenario",
            "Evidence": (
                f"The selected scenario indicates an estimated annual saving of "
                f"£{summary['annual_saving']:,.0f}."
            ),
            "Recommended action": (
                "Validate the saving through machine-level metering and develop an "
                "implementation plan with capital cost, payback, and operational risk."
            ),
        }
    )

    return insights


# ============================================================
# HEADER
# ============================================================
st.markdown(
    """
    <div class="hero">
        <div class="hero-title">⚡ Factory Energy Intelligence</div>
        <div class="hero-subtitle">
            An interactive engineering decision-support application for analysing
            manufacturing energy cost, production efficiency, downtime, quality,
            carbon emissions, and improvement scenarios.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR CONTROLS
# ============================================================
with st.sidebar:
    st.header("Analysis Controls")

    data_source = st.radio(
        "Choose data source",
        ["Use demonstration data", "Upload company Excel data"],
        help=(
            "Use the demonstration dataset to explore the app immediately, "
            "or upload a workbook containing a sheet named Factory_Data."
        ),
    )

    uploaded_file = None
    if data_source == "Upload company Excel data":
        uploaded_file = st.file_uploader(
            "Upload Excel workbook",
            type=["xlsx", "xls"],
        )

        st.download_button(
            "Download Excel template",
            data=sample_excel_bytes(),
            file_name="Factory_Energy_Data_Template.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    st.divider()
    st.subheader("Scenario Assumptions")

    reduction_pct = st.slider(
        "Expected energy reduction (%)",
        min_value=0,
        max_value=40,
        value=10,
        step=1,
    )

    scenario_tariff = st.number_input(
        "Scenario electricity tariff (£/kWh)",
        min_value=0.0,
        max_value=5.0,
        value=0.25,
        step=0.01,
        format="%.2f",
    )

    operating_days_month = st.slider(
        "Operating days per month",
        min_value=1,
        max_value=31,
        value=22,
        step=1,
    )

    carbon_factor = st.number_input(
        "Electricity carbon factor (kgCO₂e/kWh)",
        min_value=0.0,
        max_value=2.0,
        value=0.20,
        step=0.01,
        format="%.3f",
        help=(
            "Use the emissions factor appropriate to the location, "
            "reporting year, and electricity supply arrangement."
        ),
    )

    st.divider()
    st.caption(
        "Developed as a practical demonstration of AI-assisted engineering "
        "software development and manufacturing analytics."
    )


# ============================================================
# LOAD DATA
# ============================================================
try:
    if data_source == "Use demonstration data":
        raw_df = create_demo_data()
        source_label = "Demonstration dataset"
    elif uploaded_file is not None:
        raw_df = pd.read_excel(uploaded_file, sheet_name="Factory_Data")
        source_label = uploaded_file.name
    else:
        st.info(
            "Upload an Excel workbook or select the demonstration dataset "
            "to begin."
        )
        st.stop()

    clean_df, data_warnings = validate_and_clean(raw_df)

except Exception as exc:
    st.error(f"Data loading failed: {type(exc).__name__}: {exc}")
    st.stop()


# ============================================================
# DATE FILTER
# ============================================================
min_date = clean_df["Date"].min().date()
max_date = clean_df["Date"].max().date()

date_filter = st.date_input(
    "Analysis period",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

if isinstance(date_filter, tuple) and len(date_filter) == 2:
    start_date, end_date = date_filter
else:
    start_date = end_date = date_filter

filtered_df = clean_df[
    (clean_df["Date"].dt.date >= start_date)
    & (clean_df["Date"].dt.date <= end_date)
].copy()

if filtered_df.empty:
    st.warning("No records fall within the selected analysis period.")
    st.stop()


calculated_df, summary = calculate_metrics(
    filtered_df,
    scenario_tariff=scenario_tariff,
    reduction_pct=reduction_pct,
    operating_days_month=operating_days_month,
    carbon_factor=carbon_factor,
)

insights = generate_engineering_insights(calculated_df, summary)
insights_df = pd.DataFrame(insights)


# ============================================================
# STATUS AND DATA QUALITY
# ============================================================
status_col1, status_col2, status_col3 = st.columns([1.4, 1, 1])

with status_col1:
    st.success(f"Analysis complete — {source_label}")

with status_col2:
    st.info(f"{len(calculated_df)} valid operating day(s)")

with status_col3:
    if data_warnings:
        st.warning(f"{len(data_warnings)} data-quality notice(s)")
    else:
        st.success("Data-quality checks passed")

if data_warnings:
    with st.expander("View data-quality notices"):
        for warning in data_warnings:
            st.write(f"• {warning}")


# ============================================================
# MAIN TABS
# ============================================================
dashboard_tab, trends_tab, insights_tab, scenario_tab, data_tab = st.tabs(
    [
        "Executive Dashboard",
        "Interactive Trends",
        "Engineering Insights",
        "Scenario Analysis",
        "Data & Export",
    ]
)


# ============================================================
# EXECUTIVE DASHBOARD
# ============================================================
with dashboard_tab:
    st.markdown(
        '<div class="section-title">Executive Performance Summary</div>',
        unsafe_allow_html=True,
    )

    row1 = st.columns(5)
    row1[0].metric(
        "Total Production",
        f"{summary['total_production']:,.0f}",
        help="Total units produced during the selected analysis period.",
    )
    row1[1].metric(
        "Electricity Use",
        f"{summary['total_energy']:,.0f} kWh",
    )
    row1[2].metric(
        "Energy per Product",
        f"{summary['energy_per_product']:.2f} kWh/unit",
    )
    row1[3].metric(
        "Cost per Product",
        f"£{summary['cost_per_product']:.3f}",
    )
    row1[4].metric(
        "CO₂ Emissions",
        f"{summary['total_emissions'] / 1000:,.2f} tCO₂e",
    )

    row2 = st.columns(5)
    row2[0].metric(
        "Average Downtime",
        f"{summary['average_downtime']:.1f}%",
    )
    row2[1].metric(
        "Average Defect Rate",
        f"{summary['average_defect_rate']:.2f}%",
    )
    row2[2].metric(
        "Production Rate",
        f"{summary['average_production_rate']:,.0f} units/h",
    )
    row2[3].metric(
        "Potential Monthly Saving",
        f"£{summary['monthly_saving']:,.0f}",
    )
    row2[4].metric(
        "Potential Annual Saving",
        f"£{summary['annual_saving']:,.0f}",
    )

    st.divider()

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        fig = px.line(
            calculated_df,
            x="Date",
            y="Energy per Product (kWh/unit)",
            markers=True,
            title="Energy Intensity Trend",
        )
        fig.add_hline(
            y=calculated_df["Energy per Product (kWh/unit)"].mean(),
            line_dash="dash",
            annotation_text="Period average",
        )
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="kWh per unit",
            hovermode="x unified",
        )
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        fig = px.scatter(
            calculated_df,
            x="Production Output (units)",
            y="Electricity Consumption (kWh)",
            size="Downtime (%)",
            color="Defect Rate (%)",
            hover_data=["Date"],
            title="Production–Energy Relationship",
        )
        fig.update_layout(
            xaxis_title="Production output (units)",
            yaxis_title="Electricity consumption (kWh)",
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown(
        '<div class="section-title">Key Operational Observation</div>',
        unsafe_allow_html=True,
    )

    st.warning(
        f"The least energy-efficient day was "
        f"{summary['worst_day'].strftime('%d %B %Y')} at "
        f"{summary['worst_energy_intensity']:.2f} kWh/unit. "
        f"The best day was {summary['best_day'].strftime('%d %B %Y')} at "
        f"{summary['best_energy_intensity']:.2f} kWh/unit."
    )


# ============================================================
# TRENDS
# ============================================================
with trends_tab:
    metric_options = {
        "Energy per Product": "Energy per Product (kWh/unit)",
        "Daily Electricity Cost": "Daily Electricity Cost (£)",
        "Production Output": "Production Output (units)",
        "Electricity Consumption": "Electricity Consumption (kWh)",
        "Downtime": "Downtime (%)",
        "Defect Rate": "Defect Rate (%)",
        "Production Rate": "Production Rate (units/hour)",
        "CO₂ Emissions": "CO₂ Emissions (kgCO₂e)",
    }

    selected_metrics = st.multiselect(
        "Select metrics to compare",
        options=list(metric_options.keys()),
        default=[
            "Energy per Product",
            "Production Output",
            "Downtime",
        ],
    )

    if selected_metrics:
        long_df = calculated_df.melt(
            id_vars=["Date"],
            value_vars=[metric_options[m] for m in selected_metrics],
            var_name="Metric",
            value_name="Value",
        )

        fig = px.line(
            long_df,
            x="Date",
            y="Value",
            color="Metric",
            markers=True,
            facet_row="Metric",
            title="Selected Operational Trends",
        )
        fig.update_yaxes(matches=None)
        fig.for_each_annotation(
            lambda annotation: annotation.update(
                text=annotation.text.split("=")[-1]
            )
        )
        fig.update_layout(height=max(450, 230 * len(selected_metrics)))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Select at least one metric.")

    heatmap_data = calculated_df[
        [
            "Production Output (units)",
            "Electricity Consumption (kWh)",
            "Energy per Product (kWh/unit)",
            "Downtime (%)",
            "Defect Rate (%)",
            "Production Rate (units/hour)",
        ]
    ].corr()

    fig = px.imshow(
        heatmap_data,
        text_auto=".2f",
        aspect="auto",
        title="Operational Correlation Matrix",
    )
    st.plotly_chart(fig, use_container_width=True)


# ============================================================
# ENGINEERING INSIGHTS
# ============================================================
with insights_tab:
    st.markdown(
        '<div class="section-title">Automated Engineering Insights</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        "These observations are generated from the uploaded dataset using "
        "engineering rules and statistical relationships. They support, but "
        "do not replace, site investigation and professional judgement."
    )

    for insight in insights:
        st.markdown(
            f"""
            <div class="insight-card">
                <strong>{insight['Priority']} — {insight['Finding']}</strong><br>
                <span>{insight['Evidence']}</span><br><br>
                <strong>Recommended action:</strong>
                <span>{insight['Recommended action']}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    anomaly_df = calculated_df[
        calculated_df["Energy Anomaly"]
    ].copy()

    if not anomaly_df.empty:
        st.subheader("Potential Anomaly Records")
        st.dataframe(
            anomaly_df[
                [
                    "Date",
                    "Production Output (units)",
                    "Electricity Consumption (kWh)",
                    "Energy per Product (kWh/unit)",
                    "Downtime (%)",
                    "Defect Rate (%)",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# SCENARIO ANALYSIS
# ============================================================
with scenario_tab:
    st.markdown(
        '<div class="section-title">Cost and Energy Improvement Scenario</div>',
        unsafe_allow_html=True,
    )

    scenario_col1, scenario_col2 = st.columns([1.1, 1])

    with scenario_col1:
        scenario_df = pd.DataFrame(
            {
                "Scenario": [
                    "Current estimated monthly cost",
                    "New tariff without efficiency improvement",
                    "New tariff with selected energy reduction",
                ],
                "Monthly Cost (£)": [
                    summary["current_monthly_cost"],
                    summary["scenario_monthly_cost"],
                    summary["reduced_monthly_cost"],
                ],
            }
        )

        fig = px.bar(
            scenario_df,
            x="Scenario",
            y="Monthly Cost (£)",
            text_auto=".3s",
            title="Monthly Electricity Cost Comparison",
        )
        fig.update_layout(xaxis_title="", yaxis_title="Monthly cost (£)")
        st.plotly_chart(fig, use_container_width=True)

    with scenario_col2:
        st.metric(
            "Selected Energy Reduction",
            f"{reduction_pct:.0f}%",
        )
        st.metric(
            "Estimated Monthly Saving",
            f"£{summary['monthly_saving']:,.2f}",
        )
        st.metric(
            "Estimated Annual Saving",
            f"£{summary['annual_saving']:,.2f}",
        )

        simple_payback_cost = st.number_input(
            "Indicative implementation cost (£)",
            min_value=0.0,
            value=25000.0,
            step=1000.0,
        )

        if summary["annual_saving"] > 0:
            payback_years = (
                simple_payback_cost / summary["annual_saving"]
            )
            st.metric(
                "Indicative Simple Payback",
                f"{payback_years:.2f} years",
            )
        else:
            st.metric("Indicative Simple Payback", "Not available")

    st.info(
        "This scenario is a preliminary screening assessment. A business case "
        "should include capital expenditure, maintenance, energy-price uncertainty, "
        "production variability, operational risk, and verified savings."
    )


# ============================================================
# DATA AND EXPORT
# ============================================================
with data_tab:
    st.markdown(
        '<div class="section-title">Calculated Dataset</div>',
        unsafe_allow_html=True,
    )

    display_columns = [
        "Date",
        "Production Output (units)",
        "Electricity Consumption (kWh)",
        "Daily Electricity Cost (£)",
        "Energy per Product (kWh/unit)",
        "Electricity Cost per Product (£/unit)",
        "Downtime (%)",
        "Defect Rate (%)",
        "Production Rate (units/hour)",
        "CO₂ Emissions (kgCO₂e)",
        "Energy Anomaly",
    ]

    st.dataframe(
        calculated_df[display_columns],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Date": st.column_config.DateColumn(format="DD MMM YYYY"),
            "Daily Electricity Cost (£)": st.column_config.NumberColumn(
                format="£%.2f"
            ),
            "Electricity Cost per Product (£/unit)":
                st.column_config.NumberColumn(format="£%.3f"),
            "Energy per Product (kWh/unit)":
                st.column_config.NumberColumn(format="%.3f"),
            "Downtime (%)":
                st.column_config.NumberColumn(format="%.2f%%"),
            "Defect Rate (%)":
                st.column_config.NumberColumn(format="%.2f%%"),
            "CO₂ Emissions (kgCO₂e)":
                st.column_config.NumberColumn(format="%.2f"),
        },
    )

    summary_df = pd.DataFrame(
        {
            "Metric": [
                "Total production (units)",
                "Total electricity consumption (kWh)",
                "Energy per product (kWh/unit)",
                "Electricity cost per product (£/unit)",
                "Average downtime (%)",
                "Average defect rate (%)",
                "Total CO₂ emissions (kgCO₂e)",
                "Estimated monthly saving (£)",
                "Estimated annual saving (£)",
                "Least efficient day",
            ],
            "Value": [
                summary["total_production"],
                summary["total_energy"],
                summary["energy_per_product"],
                summary["cost_per_product"],
                summary["average_downtime"],
                summary["average_defect_rate"],
                summary["total_emissions"],
                summary["monthly_saving"],
                summary["annual_saving"],
                summary["worst_day"].strftime("%d %B %Y"),
            ],
        }
    )

    excel_output = dataframe_to_excel_bytes(
        calculated_df,
        summary_df,
        insights_df,
    )

    export_col1, export_col2 = st.columns(2)

    with export_col1:
        st.download_button(
            "Download Complete Excel Analysis",
            data=excel_output,
            file_name="Factory_Energy_Intelligence_Results.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with export_col2:
        st.download_button(
            "Download Calculated Data as CSV",
            data=calculated_df.to_csv(index=False).encode("utf-8"),
            file_name="Factory_Energy_Calculated_Data.csv",
            mime="text/csv",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================
st.divider()
st.markdown(
    """
    <div class="small-note">
        <strong>Engineering disclaimer:</strong> This application provides a
        preliminary analytical assessment. Users must verify input data,
        equations, tariff structures, emissions factors, assumptions, and site
        conditions before relying on the results for operational, investment,
        environmental, or design decisions.
    </div>
    """,
    unsafe_allow_html=True,
)
