"""Save the Client Status Board rows that Power Automate sends into data/client_status.json.

Power Automate runs an Office Script on the workbook in OneDrive, then sends the
result to GitHub as a repository_dispatch event (event type "excel-update") with
client_payload = {"rows": <the script's result>}. The result can arrive either as
a JSON string or as an already-parsed list, so both are accepted.

The file is only written when the data looks like the status board, so a bad or
empty run can never wipe the website.
"""

import json
import os
import sys

OUTPUT_FILE = os.path.join("data", "client_status.json")
REQUIRED_COLUMNS = {"Member ID", "Client Initials", "Bed"}


def main():
    payload = json.loads(os.environ.get("PAYLOAD") or "{}")
    rows = payload.get("rows")
    if isinstance(rows, str):
        rows = json.loads(rows)

    if not isinstance(rows, list) or not rows:
        sys.exit("No rows received. Check that the flow sends client_payload {\"rows\": <script result>}.")
    if not all(isinstance(r, dict) for r in rows):
        sys.exit("Rows are not in the expected format (a list of objects).")
    missing = REQUIRED_COLUMNS - set(rows[0])
    if missing:
        sys.exit(f"Rows are missing columns: {', '.join(sorted(missing))}. Is the script reading 'Client Status Board'?")

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=1, ensure_ascii=False)
        f.write("\n")
    print(f"Saved {len(rows)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
