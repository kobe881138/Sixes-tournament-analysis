# Sixes Tournament Performance Analysis | v0.1.4

## Deployment type: PRIVATE GitHub repository, 5 matches preloaded

This deployment package preserves the tested **Tournament Dashboard v0.1.4**, including the Single Match page.

Five XLSX files are deliberately tracked so Streamlit Community Cloud can read them when it starts. The original v0.1.4 `.gitignore` excluded XLSX and **must not** be used unchanged for this deployment mode.

**Confidentiality:** Keep the GitHub repository private. Streamlit Community Cloud's private-app sharing can be limited to approved coach emails. Making a GitHub repository public would expose the underlying workbooks and code, including their full version history; do not do so unless authorized.

## Deploy

1. Unzip the archive. The contents of `private_repo_ready/` should be the **root** of a new private GitHub repository. Do not upload the ZIP file itself.
2. At GitHub > New repository, choose **Private**. Suggested name: `sixes-tournament-analysis`.
3. Upload all files and **retain folder paths**:
   - `tournament_dashboard.py` (app entrypoint)
   - `tournament_analysis.py`, `tournament_dashboard_data.py`
   - `analysis.py`, `dashboard_data.py`, `coding_exceptions.json`
   - `pages/1_Single_Match.py`
   - `requirements.txt`
   - `Match_Records/*.xlsx` (5 files, intentionally included)
4. At https://share.streamlit.io/, log in with the same GitHub account, select **Create app** > **Yup, I have an app** / Deploy from GitHub (wording may vary).
5. Select the repository, branch `main`, and main file **`tournament_dashboard.py`**. Choose Python **3.12** under advanced options (if prompted), then **Deploy**.
6. After deployment, keep the app **private** under Sharing > Who can view this app > Only specific people; add approved coach email addresses as viewers.

### Expected first view

- `TPE · 5 match(es) · 203 possessions`
- Overview / Offensive Trends / Offensive Patterns
- Single Match page via sidebar
- Data quality: 0 errors, score reconstruction matches all five games

### Common problems

- **No files / zero data**: check the GitHub repository actually contains `Match_Records/1005_TPE_AUS.xlsx` etc. The original v0.1.4 `.gitignore` ignored spreadsheets; this deployment package does not.
- **ImportError**: ensure all 5 main `.py` support modules and `pages/1_Single_Match.py` were uploaded as well as `requirements.txt`.
- **Wrong starting page**: use `tournament_dashboard.py` as main file, not `pages/1_Single_Match.py`.
- **QC errors**: verify the five committed XLSX are the corrected versions, not templates or older duplicates.
- **No public link**: private apps require viewer sign-in and invitation; this is expected.

### Update match results

Edit your XLSX locally, re-upload the corrected XLSX to the same GitHub path, and commit; the app redeploys automatically. Preserve your authoritative local files and avoid committing old duplicates or unfinished templates. Additional finished matches can be added to `Match_Records`.

### Local check

```bash
python3 -m pip install -r requirements.txt
streamlit run tournament_dashboard.py
```

App usage supports `Local Match_Records` and `Upload Excel`. The uploaded option does not permanently save the data to the repository.

The dashboard contains no login mechanism itself; access control is administered in Streamlit Community Cloud sharing settings.
