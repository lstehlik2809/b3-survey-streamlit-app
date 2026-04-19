# B3 Survey Report App

Streamlit app for generating `B3_Report.docx` from an uploaded Excel workbook in the same format as `B3Inputs.xlsx`.

Uploaded data is processed in a temporary per-request directory. The app does not write uploads or generated reports to a database or persistent storage.

## Local Run

```powershell
python -m pip install -r requirements.txt
python -m streamlit run src/app.py
```

## Cloud Run Deployment

Project ID: `b3-survey`

1. Authenticate and select the project.

```powershell
gcloud auth login
gcloud config set project b3-survey
```

2. Enable the required services once per project.

```powershell
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com
```

3. Deploy from this folder.

```powershell
gcloud run deploy b3-survey `
  --source . `
  --region europe-west1 `
  --allow-unauthenticated `
  --memory 1Gi `
  --cpu 1 `
  --concurrency 5 `
  --min-instances 0 `
  --max-instances 2 `
  --timeout 300
```

Cloud Run will print the public service URL after deployment. With `--allow-unauthenticated`, anyone with that URL can open the app.

## Notes

- The Docker image installs Graphviz and `pygraphviz` so the network charts can use Graphviz `sfdp`. If Graphviz is unavailable, the app falls back to a NetworkX Kamada-Kawai layout.
- Instance settings are intentionally small for the expected maximum of about five users. `--concurrency 5` allows a few simultaneous requests on one instance, and `--max-instances 2` leaves room for overlap without uncontrolled scaling.
- The `.gcloudignore` and `.dockerignore` files exclude generated DOCX files and non-sample Excel files from deployment. `B3Inputs.xlsx` is included so users can download a sample input workbook.
