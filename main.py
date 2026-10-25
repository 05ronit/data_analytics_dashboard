"""Streamlit app for exploring CSV and Excel datasets."""

from __future__ import annotations

import io
import zipfile
from collections.abc import Sequence

import pandas as pd
import streamlit as st
from openpyxl.utils.exceptions import InvalidFileException
from pandas.api.types import is_datetime64_any_dtype, is_numeric_dtype
from xlrd.biffh import XLRDError


SUPPORTED_EXTENSIONS = ("csv", "xls", "xlsx", "xlsm")
DATE_NAME_HINTS = ("date", "time", "timestamp", "month")
REVENUE_NAME_HINTS = ("revenue", "sales", "amount", "invoicevalue", "totalvalue")


def read_uploaded_data(
    file_bytes: bytes, filename: str, sheet_name: str | None = None
) -> pd.DataFrame:
    """Read a supported tabular upload without writing it to disk."""
    extension = filename.rsplit(".", 1)[-1].lower()
    if extension == "csv":
        try:
            content = file_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            content = file_bytes.decode("cp1252")
        return pd.read_csv(io.StringIO(content), sep=None, engine="python")
    if extension in ("xls", "xlsx", "xlsm"):
        return pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name,
            engine="xlrd" if extension == "xls" else "openpyxl",
        )
    raise ValueError(
        "Unsupported file type. Upload a CSV, XLSX, or XLSM file."
    )


def get_excel_sheet_names(file_bytes: bytes, filename: str) -> list[str]:
    extension = filename.rsplit(".", 1)[-1].lower()
    engine = "xlrd" if extension == "xls" else "openpyxl"
    with pd.ExcelFile(io.BytesIO(file_bytes), engine=engine) as workbook:
        return workbook.sheet_names


def normalize_column_name(name: object) -> str:
    return "".join(character for character in str(name).casefold() if character.isalnum())


def find_column(columns: Sequence[object], hints: Sequence[str]) -> object | None:
    for column in columns:
        normalized_name = normalize_column_name(column)
        if any(hint in normalized_name for hint in hints):
            return column
    return None


def prepare_dashboard_data(source: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    """Copy input data, parse likely date columns, and infer retail revenue if possible."""
    data = source.copy()
    for column in data.columns:
        if is_datetime64_any_dtype(data[column]):
            continue
        if any(hint in normalize_column_name(column) for hint in DATE_NAME_HINTS):
            parsed = pd.to_datetime(data[column], errors="coerce")
            if parsed.notna().any():
                data[column] = parsed

    revenue_column = find_column(data.columns, REVENUE_NAME_HINTS)
    quantity_column = find_column(data.columns, ("quantity", "units"))
    price_column = find_column(data.columns, ("unitprice", "priceperunit"))
    derived_revenue = (
        revenue_column is None
        and quantity_column is not None
        and price_column is not None
    )
    if derived_revenue:
        quantity = pd.to_numeric(data[quantity_column], errors="coerce")
        unit_price = pd.to_numeric(data[price_column], errors="coerce")
        data["Revenue (calculated)"] = quantity * unit_price

    return data, derived_revenue


def render_automatic_charts(data: pd.DataFrame) -> None:
    numeric_columns = [
        column for column in data.columns if is_numeric_dtype(data[column])
    ]
    date_columns = [
        column for column in data.columns if is_datetime64_any_dtype(data[column])
    ]
    categorical_columns = [
        column
        for column in data.columns
        if not is_numeric_dtype(data[column])
        and not is_datetime64_any_dtype(data[column])
    ]

    st.subheader("Automatic highlights")
    chart_columns = st.columns(2)

    with chart_columns[0]:
        if date_columns and numeric_columns:
            date_column = date_columns[0]
            metric_column = find_column(numeric_columns, REVENUE_NAME_HINTS)
            if metric_column is None:
                metric_column = numeric_columns[0]
            trend = (
                data.groupby(date_column, dropna=True, sort=False)[metric_column]
                .sum()
                .sort_index()
                .to_frame()
            )
            st.markdown(f"**{metric_column} over {date_column}**")
            st.line_chart(trend)
        elif numeric_columns:
            metric_column = numeric_columns[0]
            values = data[metric_column].dropna()
            st.markdown(f"**Distribution of {metric_column}**")
            if not values.empty:
                histogram = values.value_counts(
                    bins=min(20, max(1, values.nunique())), sort=False
                )
                histogram_data = histogram.rename("Rows").rename_axis("Range").reset_index()
                histogram_data["Range"] = histogram_data["Range"].astype(str)
                st.bar_chart(histogram_data, x="Range", y="Rows")
            else:
                st.info("No values are available for this chart.")
        else:
            st.info("Add a numeric column to see a distribution or trend chart.")

    with chart_columns[1]:
        if categorical_columns:
            category_column = categorical_columns[0]
            counts = (
                data[category_column]
                .value_counts(dropna=False)
                .head(20)
                .rename("Rows")
                .rename_axis(category_column)
                .reset_index()
            )
            counts[category_column] = counts[category_column].fillna("(missing)").astype(str)
            st.markdown(f"**Most common values in {category_column}**")
            st.bar_chart(counts, x=category_column, y="Rows")
        else:
            st.info("Add a text or category column to see its most common values.")


def render_chart_builder(data: pd.DataFrame) -> None:
    numeric_columns = [
        column for column in data.columns if is_numeric_dtype(data[column])
    ]
    chart_types = ["Bar"]
    if numeric_columns:
        chart_types.extend(["Line", "Area", "Scatter", "Histogram"])

    st.subheader("Build a chart")
    chart_type = st.selectbox("Chart type", chart_types)
    if chart_type == "Histogram":
        metric_column = st.selectbox("Numeric column", numeric_columns)
        values = data[metric_column].dropna()
        if values.empty:
            st.info("This column has no numeric values to chart.")
            return
        histogram = values.value_counts(
            bins=min(20, max(1, values.nunique())), sort=False
        )
        histogram_data = histogram.rename("Rows").rename_axis("Range").reset_index()
        histogram_data["Range"] = histogram_data["Range"].astype(str)
        st.bar_chart(histogram_data, x="Range", y="Rows")
        return

    x_column = st.selectbox("Horizontal axis", list(data.columns))
    y_options = ["Row count", *numeric_columns]
    default_y = find_column(numeric_columns, REVENUE_NAME_HINTS)
    if default_y is None:
        default_y = numeric_columns[0] if numeric_columns else "Row count"
    y_column = st.selectbox(
        "Vertical axis",
        y_options,
        index=y_options.index(default_y),
    )

    if chart_type == "Scatter":
        if y_column == "Row count":
            st.info("Choose a numeric vertical-axis column for a scatter chart.")
            return
        scatter_data = data[[x_column, y_column]].dropna()
        if scatter_data.empty:
            st.info("There are no complete rows for these chart columns.")
            return
        st.scatter_chart(scatter_data, x=x_column, y=y_column)
        return

    aggregation = "Sum"
    if y_column != "Row count":
        aggregation = st.selectbox("Aggregation", ["Sum", "Mean", "Count"])

    if y_column == "Row count":
        chart_data = (
            data.groupby(x_column, dropna=False, sort=False)
            .size()
            .rename("Row count")
            .reset_index()
        )
        chart_y = "Row count"
    else:
        aggregate_function = {
            "Sum": "sum",
            "Mean": "mean",
            "Count": "count",
        }[aggregation]
        chart_data = (
            data.groupby(x_column, dropna=False, sort=False)[y_column]
            .agg(aggregate_function)
            .reset_index()
        )
        chart_y = y_column

    if chart_data.empty:
        st.info("There is no data to chart for these selections.")
        return
    if is_datetime64_any_dtype(chart_data[x_column]):
        chart_data = chart_data.sort_values(x_column)
    elif chart_type == "Bar":
        chart_data = chart_data.sort_values(chart_y, ascending=False).head(30)
    if chart_type in ("Line", "Area"):
        if chart_type == "Line":
            st.line_chart(
                chart_data,
                x=x_column,
                y=chart_y,
                width="stretch",
                height=380,
            )
        else:
            st.area_chart(
                chart_data,
                x=x_column,
                y=chart_y,
                width="stretch",
                height=380,
            )
    else:
        st.bar_chart(
            chart_data,
            x=x_column,
            y=chart_y,
            width="stretch",
            height=380,
        )


def main() -> None:
    st.set_page_config(
        page_title="Data Canvas | Interactive Dashboard",
        page_icon="📊",
        layout="wide",
    )
    st.markdown(
        """
        <style>
        .stApp { background: #f5f7fb; }
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e5eaf2;
            padding: 16px 18px;
            border-radius: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("Data Canvas")
    st.caption("Upload a CSV or Excel worksheet to explore it and build a dashboard.")

    with st.sidebar:
        st.header("Your data")
        uploaded_file = st.file_uploader(
            "Choose a CSV or Excel file",
            type=list(SUPPORTED_EXTENSIONS),
            help="Supported formats: CSV, XLS, XLSX, and XLSM. Excel workbooks can contain multiple sheets.",
        )

    if uploaded_file is None:
        sample_path = "retail_sales.csv"
        try:
            with open(sample_path, "rb") as sample_file:
                file_bytes = sample_file.read()
            filename = sample_path
            st.info(
                "Showing the included retail sales sample. Upload your own file in "
                "the sidebar to replace it."
            )
        except FileNotFoundError:
            st.info("Upload a CSV or Excel file in the sidebar to get started.")
            st.stop()
    else:
        file_bytes = uploaded_file.getvalue()
        filename = uploaded_file.name

    extension = filename.rsplit(".", 1)[-1].lower()
    sheet_name = None
    if extension in ("xls", "xlsx", "xlsm"):
        try:
            sheet_names = get_excel_sheet_names(file_bytes, filename)
        except (
            ValueError,
            OSError,
            ImportError,
            zipfile.BadZipFile,
            InvalidFileException,
            XLRDError,
        ) as error:
            st.error(f"Could not read this Excel workbook: {error}")
            st.stop()
        if not sheet_names:
            st.error("This workbook does not contain any worksheets.")
            st.stop()
        with st.sidebar:
            sheet_name = st.selectbox("Worksheet", sheet_names)

    try:
        source_data = read_uploaded_data(file_bytes, filename, sheet_name)
    except (
        ValueError,
        OSError,
        UnicodeDecodeError,
        pd.errors.ParserError,
        pd.errors.EmptyDataError,
        ImportError,
        zipfile.BadZipFile,
        InvalidFileException,
        XLRDError,
    ) as error:
        st.error(f"Could not read {filename}: {error}")
        st.stop()

    if source_data.empty or len(source_data.columns) == 0:
        st.error("The selected file or worksheet does not contain tabular data.")
        st.stop()

    data, derived_revenue = prepare_dashboard_data(source_data)
    if derived_revenue:
        st.caption("Revenue (calculated) was added as Quantity × Unit Price.")

    missing_cells = int(data.isna().sum().sum())
    duplicate_rows = int(data.duplicated().sum())
    numeric_columns = [
        column for column in data.columns if is_numeric_dtype(data[column])
    ]
    revenue_column = find_column(numeric_columns, REVENUE_NAME_HINTS)

    st.subheader(f"Dashboard · {filename}")
    metrics = st.columns(5 if revenue_column is not None else 4)
    metrics[0].metric("Rows", f"{len(data):,}")
    metrics[1].metric("Columns", f"{len(data.columns):,}")
    metrics[2].metric("Missing cells", f"{missing_cells:,}")
    metrics[3].metric("Duplicate rows", f"{duplicate_rows:,}")
    if revenue_column is not None:
        total_value = pd.to_numeric(data[revenue_column], errors="coerce").sum()
        metrics[4].metric(f"Total {revenue_column}", f"{total_value:,.2f}")

    with st.expander("Data quality", expanded=False):
        quality = pd.DataFrame(
            {
                "Missing values": data.isna().sum(),
                "Missing (%)": (data.isna().mean() * 100).round(1),
                "Data type": data.dtypes.astype(str),
            }
        )
        st.dataframe(quality, width="stretch")
        st.caption(
            "Rows are shown as uploaded; this dashboard does not remove or overwrite "
            "your source data."
        )

    render_automatic_charts(data)
    st.divider()
    render_chart_builder(data)

    with st.expander("Preview data", expanded=False):
        st.dataframe(data.head(500), width="stretch", height=360)
        st.caption("Preview is limited to the first 500 rows for responsiveness.")

    st.download_button(
        "Download dashboard data as CSV",
        data=data.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{filename.rsplit('.', 1)[0]}_dashboard.csv",
        mime="text/csv",
    )
    st.caption("Uploaded data is processed in memory for this Streamlit session.")


if __name__ == "__main__":
    main()
