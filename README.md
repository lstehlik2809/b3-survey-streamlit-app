# B3 Survey Report App

Streamlit app for generating `B3_Report.docx` from an uploaded Excel workbook in the same structure as `B3Inputs.xlsx`.

Uploaded data is processed in a temporary per-request directory. The app does not write uploads or generated reports to a database or persistent storage.

## Local Run

```powershell
python -m pip install -r requirements.txt
python -m streamlit run src/app.py
```

## Streamlit Community Cloud Deployment

1. Push this repository to GitHub.

2. Go to `https://share.streamlit.io/`.

3. Click `Create app`.

4. Select the GitHub repository, branch, and entrypoint file:

```text
src/app.py
```

5. In advanced settings, select Python `3.11`.

6. Deploy the app.

The app includes:

- `requirements.txt` for Python dependencies.
- `packages.txt` for Linux system packages required by Graphviz and `pygraphviz`.
- `.streamlit/config.toml` to use the light Streamlit theme by default.
- `B3Inputs.xlsx` so users can download a sample input workbook.

## Notes

- The uploaded Excel file can have any filename. It only needs to match the required workbook structure.
- The workbook must contain `nodes` and `edges` sheets with the same columns as the sample file.
- Network charts use Graphviz `sfdp` when available. If Graphviz is unavailable, the app falls back to a NetworkX Kamada-Kawai layout.
- `B3Inputs.xlsx` is included intentionally as a sample workbook. Make sure it does not contain sensitive data before publishing this repository.
