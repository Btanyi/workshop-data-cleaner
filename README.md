# Student Data Cleaner

A Streamlit app for reviewing student incentive records, flagging missing or
duplicate student IDs, and downloading an Excel audit report.

## Run locally

Use Python 3.10 or newer:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

The app accepts CSV, XLS, and XLSX files up to 50 MB each. Uploaded data is
processed in the active Streamlit session and the generated workbook is
provided as a download; the app does not intentionally write uploaded records
to disk.

## Deploy from GitHub

1. Create a repository containing the files permitted by `.gitignore`.
2. Push the repository to GitHub. Confirm that no student records, exported
   workbooks, credentials, or local environment files are included.
3. In Streamlit Community Cloud, create an app from the repository and select
   `app.py` as the app entry point.
4. Streamlit installs the packages listed in `requirements.txt`.

The repository is intended to contain application source only. The root
`.gitignore` is intentionally allowlisted to keep unrelated portfolio
projects and data files out of this app repository.

## Student-data privacy

Publishing this source repository does not make a deployment private. A
publicly accessible app can be used by anyone who can reach its URL. Do not
upload identifiable, confidential, or otherwise protected student records to
a public deployment. Use deployment-level authentication and access controls
approved by your organization before processing real student data.

Uploaded files and derived data exist in server memory while the Streamlit
session is active. Users can see records they upload in that session and can
download the resulting report. Avoid shared devices and close the session when
finished. The app does not provide authentication, retention guarantees, or
compliance certification.

Excel exports escape formula-like cell values and column headings so uploaded
content is exported as text rather than as executable spreadsheet formulas.
