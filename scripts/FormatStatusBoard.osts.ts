// Office Script: formats the "Client Status Board" sheet.
// Excel on the web > Automate > New Script > paste > save as "Format Status Board" > Run.
// Safe to run again any time. It only changes the look (colors, widths, borders,
// date/time display), never the data or formulas, and keeps headers on row 4 so
// the website update keeps working.

function main(workbook: ExcelScript.Workbook) {
  const ws = workbook.getWorksheet("Client Status Board");
  if (!ws) {
    throw new Error('Sheet "Client Status Board" not found.');
  }

  const navy = "#1F3A5F";
  const teal = "#2E7D8C";
  const autoFill = "#F2F5F8";
  const line = "#D5DDE5";
  const muted = "#A0AAB4";

  // last row that has a bed listed in column C
  const used = ws.getUsedRange();
  const usedLast = Math.max(used.getRowIndex() + used.getRowCount(), 5);
  const bedTexts = ws.getRange(`C5:C${usedLast}`).getTexts();
  let lastRow = 5;
  bedTexts.forEach((r, i) => { if (r[0].trim() !== "") lastRow = 5 + i; });
  const endRow = Math.max(lastRow, 40); // leave room for more beds
  const table = ws.getRange(`A4:O${lastRow}`);
  const data = ws.getRange(`A5:O${lastRow}`);

  ws.setShowGridlines(false);

  // ---- Title and instructions ----
  const title = ws.getRange("A1:O1");
  title.merge(false);
  title.getFormat().getFill().setColor(navy);
  title.getFormat().getFont().setColor("#FFFFFF");
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setSize(18);
  title.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  title.getFormat().setIndentLevel(1);
  title.getFormat().setRowHeight(34);

  const info = ws.getRange("A2:O2");
  info.merge(false);
  ws.getRange("A2").setValue(
    "Dark headers fill in automatically from the Client Needs Tracker. " +
    "Teal headers are filled in by hand. Changes here update the website automatically."
  );
  info.getFormat().getFont().setItalic(true);
  info.getFormat().getFont().setColor("#4A5563");
  info.getFormat().getFont().setSize(10);
  info.getFormat().setWrapText(true);
  info.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  info.getFormat().setIndentLevel(1);
  info.getFormat().setRowHeight(22);
  ws.getRange("A3:O3").getFormat().setRowHeight(6);

  // ---- Header row (row 4) ----
  const header = ws.getRange("A4:O4");
  header.getFormat().getFill().setColor(navy);
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFont().setBold(true);
  header.getFormat().setWrapText(true);
  header.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  header.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  header.getFormat().setRowHeight(36);
  // columns filled in by hand get a teal header
  ws.getRange("F4:L4").getFormat().getFill().setColor(teal);

  // ---- Data rows ----
  const df = data.getFormat();
  df.getFont().setSize(11);
  df.getFont().setColor("#1F2933");
  df.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  df.getFill().clear();
  ["A", "B", "C", "D", "E", "M", "N", "O"].forEach(col =>
    ws.getRange(`${col}5:${col}${lastRow}`).getFormat().getFill().setColor(autoFill)
  );
  ["A", "B", "C", "H", "I", "N", "O"].forEach(col =>
    ws.getRange(`${col}5:${col}${lastRow}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)
  );
  ["J", "M"].forEach(col => ws.getRange(`${col}5:${col}${lastRow}`).getFormat().setWrapText(true));
  ws.getRange(`B5:C${lastRow}`).getFormat().getFont().setBold(true);

  // date and time display
  ws.getRange(`H5:H${endRow}`).setNumberFormat("h:mm AM/PM");
  ws.getRange(`I5:I${endRow}`).setNumberFormat("mm/dd/yyyy");
  ws.getRange(`O5:O${endRow}`).setNumberFormat("mm/dd/yyyy");

  // light lines between rows, frame around the table
  const tf = table.getFormat();
  [ExcelScript.BorderIndex.insideHorizontal, ExcelScript.BorderIndex.insideVertical,
   ExcelScript.BorderIndex.edgeBottom, ExcelScript.BorderIndex.edgeLeft,
   ExcelScript.BorderIndex.edgeRight].forEach(side => {
    const b = tf.getRangeBorder(side);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setColor(line);
    b.setWeight(ExcelScript.BorderWeight.thin);
  });

  // thicker line where the female beds start
  const beds = ws.getRange(`C5:C${lastRow}`).getTexts();
  for (let i = 0; i < beds.length; i++) {
    if (beds[i][0].trim().toUpperCase().startsWith("F")) {
      const b = ws.getRange(`A${5 + i}:O${5 + i}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
      b.setStyle(ExcelScript.BorderLineStyle.continuous);
      b.setColor(navy);
      b.setWeight(ExcelScript.BorderWeight.medium);
      break;
    }
  }

  fitToText(ws, lastRow);
  colorBeds(ws, lastRow);

  // ---- Color rules (update automatically as you type) ----
  const area = ws.getRange(`A5:O${endRow}`);
  area.clearAllConditionalFormats();

  // open beds: grey out the whole row
  addRule(ws, `A5:O${endRow}`, '=AND($C5<>"",$B5="")', "", muted, false);

  // appointment status
  addRule(ws, `F5:F${endRow}`, '=AND($B5<>"",F5="Scheduled")', "#E3F4E8", "#1E6B3A", true);
  addRule(ws, `F5:F${endRow}`, '=AND($B5<>"",F5="Requested/Pending")', "#FFF3D6", "#8A5A00", true);

  // pharmacy status
  addRule(ws, `K5:K${endRow}`, '=AND($B5<>"",K5="Picked Up")', "#E3F4E8", "#1E6B3A", true);
  addRule(ws, `K5:K${endRow}`, '=AND($B5<>"",K5="Submitted/Pending")', "#FFF3D6", "#8A5A00", true);
  addRule(ws, `K5:K${endRow}`, '=AND($B5<>"",K5="Other/See Med Clerk")', "#E3EEFB", "#1F4E8C", true);

  // appointment date: today = red, next 3 days = amber
  addRule(ws, `I5:I${endRow}`, '=AND($B5<>"",ISNUMBER(I5),I5=TODAY())', "#FDE4E4", "#9B1C1C", true);
  addRule(ws, `I5:I${endRow}`, '=AND($B5<>"",ISNUMBER(I5),I5>TODAY(),I5<=TODAY()+3)', "#FFF3D6", "#8A5A00", true);

  // keep the title and headers visible while scrolling
  ws.getFreezePanes().unfreeze();
  ws.getFreezePanes().freezeRows(4);
}

function addRule(ws: ExcelScript.Worksheet, address: string, formula: string,
                 fill: string, font: string, bold: boolean) {
  const cf = ws.getRange(address).addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  const custom = cf.getCustom();
  custom.getRule().setFormula(formula);
  if (fill) custom.getFormat().getFill().setColor(fill);
  if (font) custom.getFormat().getFont().setColor(font);
  if (bold) custom.getFormat().getFont().setBold(true);
}

// Sizes columns to their text (Notes and Med Times wrap instead of growing forever),
// then sizes rows to fit.
function fitToText(ws: ExcelScript.Worksheet, lastRow: number) {
  const data = ws.getRange(`A5:O${lastRow}`);
  data.getFormat().setWrapText(false);
  data.getFormat().autofitColumns();
  "ABCDEFGHIJKLMNO".split("").forEach(col => {
    const f = ws.getRange(`${col}:${col}`).getFormat();
    const max = (col === "J" || col === "M") ? 300 : 200;
    f.setColumnWidth(Math.min(Math.max(f.getColumnWidth() + 14, 64), max));
  });
  ws.getRange(`J5:J${lastRow}`).getFormat().setWrapText(true);
  ws.getRange(`M5:M${lastRow}`).getFormat().setWrapText(true);
  const header = ws.getRange("A4:O4").getFormat();
  header.setWrapText(true);
  header.setRowHeight(36);
  data.getFormat().autofitRows();
  for (let r = 5; r <= lastRow; r++) {
    const f = ws.getRange(`A${r}:O${r}`).getFormat();
    if (f.getRowHeight() < 22) f.setRowHeight(22);
  }
}

// Blue borders for male beds (M - ...), pink for female beds (F - ...),
// with a thick colored edge on the left and a tinted Bed cell.
function colorBeds(ws: ExcelScript.Worksheet, lastRow: number) {
  const male = { line: "#3B6FB6", tint: "#E6EEFA" };
  const female = { line: "#C2407A", tint: "#FBE7F0" };
  const beds = ws.getRange(`C5:C${lastRow}`).getTexts();
  const sides = [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
                 ExcelScript.BorderIndex.edgeRight, ExcelScript.BorderIndex.insideVertical];
  beds.forEach((b, i) => {
    const bed = b[0].trim().toUpperCase();
    const c = bed.startsWith("M") ? male : bed.startsWith("F") ? female : null;
    if (!c) return;
    const r = 5 + i;
    const rowFmt = ws.getRange(`A${r}:O${r}`).getFormat();
    sides.forEach(side => {
      const border = rowFmt.getRangeBorder(side);
      border.setStyle(ExcelScript.BorderLineStyle.continuous);
      border.setColor(c.line);
      border.setWeight(ExcelScript.BorderWeight.thin);
    });
    const left = ws.getRange(`A${r}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
    left.setStyle(ExcelScript.BorderLineStyle.continuous);
    left.setColor(c.line);
    left.setWeight(ExcelScript.BorderWeight.thick);
    ws.getRange(`C${r}`).getFormat().getFill().setColor(c.tint);
  });
}
