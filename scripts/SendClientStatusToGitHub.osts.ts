// Office Script for Excel on the web (Automate tab > New script).
// Paste everything below into the script editor, save it as "Send Client Status to GitHub".
// Power Automate runs it and sends the result to GitHub.
//
// It reads the "Client Status Board" sheet (headers on row 4) and returns the rows
// as JSON text, exactly as they display in Excel.

function main(workbook: ExcelScript.Workbook): string {
  const sheet = workbook.getWorksheet("Client Status Board");
  if (!sheet) {
    throw new Error('Sheet "Client Status Board" not found.');
  }
  const used = sheet.getUsedRange();
  const texts = used.getTexts();
  const headerIndex = 3 - used.getRowIndex(); // row 4 of the sheet
  const headers = texts[headerIndex].map(h => h.trim());

  const rows: { [key: string]: string | null }[] = [];
  for (let i = headerIndex + 1; i < texts.length; i++) {
    const row: { [key: string]: string | null } = {};
    let hasValue = false;
    headers.forEach((h, j) => {
      if (!h) return;
      const v = (texts[i][j] || "").trim();
      row[h] = v === "" ? null : v;
      if (v !== "") hasValue = true;
    });
    if (hasValue) rows.push(row);
  }
  return JSON.stringify(rows);
}
