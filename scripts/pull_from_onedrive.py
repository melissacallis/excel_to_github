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
import re
import sys
import urllib.parse
import urllib.request
import zipfile

OUTPUT_FILE = "Client_Needs_Tracker_Finalexcel.xlsx"
REQUIRED_SHEET = "Client Status Board"  # the sheet index.html reads


UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


def note(msg):
    """Print a line that also shows up as an annotation on the Actions run page."""
    print(f"::warning::{msg}")


def with_download(url):
    parts = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qs(parts.query)
    query["download"] = ["1"]
    return urllib.parse.urlunsplit(
        parts._replace(query=urllib.parse.urlencode(query, doseq=True))
    )


def shares_token(url):
    return "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")


def fetch(url):
    """Return (final_url, content_type, body)."""
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.geturl(), resp.headers.get("Content-Type", ""), resp.read()


def candidate_urls(share_url):
    """Direct-download URLs to try for a OneDrive/SharePoint share link."""
    urls = [with_download(share_url)]
    # Short 1drv.ms links redirect to the real share page; try download forms of that too.
    try:
        final, _, _ = fetch(share_url)
        note(f"Share link resolves to {final[:160]}")
        if final != share_url:
            urls.append(with_download(final))
            if "/:x:/" in final:
                urls.append(with_download(final.replace("/:x:/", "/:u:/")))
            parts = urllib.parse.urlsplit(final)
            m = re.search(r"/personal/([^/]+)/", parts.path)
            if m:
                token = parts.path.rstrip("/").rsplit("/", 1)[-1]
                urls.append(
                    f"https://{parts.netloc}/personal/{m.group(1)}/_layouts/15/download.aspx?share={token}"
                )
    except Exception as exc:
        note(f"Could not open share link: {exc}")
    urls.append(f"https://api.onedrive.com/v1.0/shares/{shares_token(share_url)}/root/content")
    urls.append(f"https://api.onedrive.com/v1.0/shares/{shares_token(share_url)}/driveItem/content")
    seen, out = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


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
            final, ctype, data = fetch(url)
        except Exception as exc:  # try the next method
            note(f"Download attempt failed ({url[:120]}): {exc}")
            continue
        note(f"Got {len(data):,} bytes, type '{ctype}', from {final[:120]}")
        if is_valid_workbook(data):
            with open(OUTPUT_FILE, "wb") as f:
                f.write(data)
            print(f"Saved {OUTPUT_FILE} ({len(data):,} bytes)")
            return
        if data.startswith(b"PK"):
            note(f"Got an Excel file, but it has no '{REQUIRED_SHEET}' sheet.")
        print("Downloaded something, but it was not the tracker workbook; trying next method.")

    sys.exit(
        "Could not download the workbook. Check that the OneDrive link is shared as "
        "'Anyone with the link can view' and that the file still has a "
        f"'{REQUIRED_SHEET}' sheet."
    )


if __name__ == "__main__":
    main()
