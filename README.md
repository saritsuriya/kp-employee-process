# Excel Processor for SAP

Streamlit app for processing employee master and HR update Excel files. The app generates:

- SAP file
- Resigned batch file
- New master file

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push this repository to GitHub.
2. Open https://share.streamlit.io.
3. Click **Create app**.
4. Select the repository and branch.
5. Set the main file path to `app.py`.
6. In advanced settings, select Python 3.13 if available to match the local test environment. Python 3.12 should also work for this app.
7. Click **Deploy**.

## Deployment notes

- `requirements.txt` is already in the repository root, which Streamlit Cloud uses to install dependencies.
- No Streamlit secrets are required for the current app.
- Do not commit employee Excel files. Upload them through the app UI when processing.
- For employee data, deploy from a private GitHub repository and keep the Streamlit app private unless public access is explicitly acceptable.
