# Data Canvas

An interactive Streamlit dashboard for CSV and Excel data. It includes the
retail sales dataset as a sample and lets you upload your own data to generate
automatic summaries, data-quality details, and interactive charts.

## Run locally

```powershell
python -m pip install -r requirements.txt
streamlit run main.py
```

Supported uploads are CSV, XLS, XLSX, and XLSM files. For Excel workbooks, choose a
worksheet in the sidebar. The dashboard includes row and column counts, missing
and duplicate values, automatic charts, a chart builder, a data preview, and a
CSV download of the data shown in the dashboard. If a dataset contains quantity
and unit-price columns but no revenue column, revenue is calculated for display.

Uploaded data is processed in memory for the active Streamlit session and is
not written to the project directory.

## Deploy on Streamlit Community Cloud

1. Push this project to a GitHub repository.
2. In Streamlit Community Cloud, create an app from that repository.
3. Set the app entry point to `main.py`.
4. Deploy. Streamlit installs the packages listed in `requirements.txt`.

The included `.streamlit/config.toml` sets the app theme and a 200 MB upload
limit.
