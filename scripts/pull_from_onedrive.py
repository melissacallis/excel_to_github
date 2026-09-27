"""Download the tracker workbook from a OneDrive share link and save it into the repo.

Reads the share link from the ONEDRIVE_URL environment variable. Works with
personal OneDrive links (1drv.ms / onedrive.live.com) and OneDrive for
Business / SharePoint links. The file is only written if it is a real Excel
workbook that still contains the sheet the dashboard reads, so a broken or
expired link can never overwrite the good copy with an error page.
"""

import base64
import io
import os
import sys
import urllib.parse
import urllib.request
import zipfile

OUTPUT_FILE = "Client_Needs_Tracker_Finalexcel.xlsx"
REQUIRED_SHEET = "Client Status Board"  # the sheet index.html reads


def candidate_urls(share_url):
    """Direct-download URLs to try for a OneDrive/SharePoint share link."""
    parts = urllib.parse.urlsplit(share_url)
    query = urllib.parse.parse_qs(parts.query)
    query["download"] = ["1"]
    with_download = urllib.parse.urlunsplit(
        parts._replace(query=urllib.parse.urlencode(query, doseq=True))
    )
    encoded = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    shares_api = f"https://api.onedrive.com/v1.0/shares/u!{encoded}/root/content"
    return [with_download, shares_api]


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "excel-to-github-sync"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def is_valid_workbook(data):
    if not data.startswith(b"PK"):
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            workbook_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
    except (zipfile.BadZipFile, KeyError):
        return False
    return f'name="{REQUIRED_SHEET}"' in workbook_xml


def main():
    share_url = os.environ.get("ONEDRIVE_URL", "").strip()
    if not share_url:
        sys.exit("ONEDRIVE_URL secret is not set. Add it under Settings > Secrets and variables > Actions.")

    for url in candidate_urls(share_url):
        try:
            data = fetch(url)
        except Exception as exc:  # try the next method
            print(f"Download attempt failed: {exc}")
            continue
        if is_valid_workbook(data):
            with open(OUTPUT_FILE, "wb") as f:
                f.write(data)
            print(f"Saved {OUTPUT_FILE} ({len(data):,} bytes)")
            return
        print("Downloaded something, but it was not the tracker workbook; trying next method.")

    sys.exit(
        "Could not download the workbook. Check that the OneDrive link is shared as "
        "'Anyone with the link can view' and that the file still has a "
        f"'{REQUIRED_SHEET}' sheet."
    )


if __name__ == "__main__":
    main()
