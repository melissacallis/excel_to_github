# excel_to_github
Dashboard from Excel Worksheet

`index.html` reads the **Client Status Board** sheet from
`Client_Needs_Tracker_Finalexcel.xlsx` in this repo.

## Keeping the workbook in sync with OneDrive

A GitHub Action (`.github/workflows/sync-from-onedrive.yml`) downloads the
workbook from OneDrive every 15 minutes and commits it only when it changed.
It replaces the old Power Automate flow, so there are no tokens to expire.

### One-time setup
1. In OneDrive, right-click the workbook → **Share** → set to
   **Anyone with the link can view** → **Copy link**.
2. In this repo: **Settings → Secrets and variables → Actions → New repository secret**.
   Name: `ONEDRIVE_URL`, value: the link you copied.
3. Go to the **Actions** tab → **Sync workbook from OneDrive** → **Run workflow**
   to test it right away.

If the link breaks or the file is renamed, the run fails with a clear message
and the existing workbook in the repo is left untouched.

Note: GitHub pauses scheduled workflows after 60 days with no repo activity.
If that happens, re-enable it from the Actions tab.
