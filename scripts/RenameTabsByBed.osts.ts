// Office Script: renames every "MAR - ..." and "Inventory - ..." tab to its bed number.
// Excel on the web > Automate > New Script > paste > save as "Rename Tabs by Bed" > Run.
//
// Example: "MAR - P01" -> "MAR - M-1A", "Inventory - P09" -> "Inventory - F-5A".
// Each tab already links to one row of the Client Needs Tracker (its Bed cell, column S),
// and the bed for that row is what's used for the name. Excel updates every formula
// that points at a renamed tab automatically, so nothing breaks.
// Safe to run again any time, e.g. after you change a bed label on the tracker.

function main(workbook: ExcelScript.Workbook) {
  const tracker = workbook.getWorksheet("Client Needs Tracker");
  if (!tracker) {
    throw new Error('Sheet "Client Needs Tracker" not found.');
  }

  type Plan = { sheet: ExcelScript.Worksheet; from: string; to: string };
  const plans: Plan[] = [];

  for (const sheet of workbook.getWorksheets()) {
    const name = sheet.getName();
    const prefix = name.startsWith("MAR - ") ? "MAR" : name.startsWith("Inventory - ") ? "Inventory" : "";
    if (!prefix) continue;

    // Find which tracker row this tab is tied to, from its header formulas (rows 1-3).
    const formulas = sheet.getRange("A1:AG3").getFormulas();
    let trackerRow = 0;
    for (const row of formulas) {
      for (const f of row) {
        const m = String(f).match(/'Client Needs Tracker'!\$?S\$?(\d+)/);
        if (m) { trackerRow = Number(m[1]); break; }
      }
      if (trackerRow) break;
    }
    if (!trackerRow) {
      console.log(`Skipped "${name}": couldn't find its link to the Client Needs Tracker.`);
      continue;
    }

    const bed = tracker.getRange(`S${trackerRow}`).getText().trim();
    if (!bed) {
      console.log(`Skipped "${name}": tracker row ${trackerRow} has no bed.`);
      continue;
    }

    // "M - 1A" -> "M-1A"; drop characters Excel doesn't allow in tab names.
    const cleanBed = bed.replace(/\s+/g, "").replace(/[\\\/\?\*\[\]:]/g, "");
    const target = `${prefix} - ${cleanBed}`.slice(0, 31);
    if (target !== name) plans.push({ sheet, from: name, to: target });
  }

  // Stop if two tabs would end up with the same name (e.g. two rows with the same bed).
  const targets = plans.map(p => p.to);
  const moving = new Set(plans.map(p => p.from));
  const staying = workbook.getWorksheets().map(s => s.getName()).filter(n => !moving.has(n));
  const dupes = targets.filter((t, i) => targets.indexOf(t) !== i || staying.includes(t));
  if (dupes.length) {
    throw new Error(`Two tabs would get the same name: ${[...new Set(dupes)].join(", ")}. ` +
      "Check the Bed column on the Client Needs Tracker for duplicates. Nothing was renamed.");
  }

  // Rename in two passes so swapping names between tabs can't collide.
  plans.forEach((p, i) => p.sheet.setName(`__rename_${i}`));
  plans.forEach(p => {
    p.sheet.setName(p.to);
    console.log(`${p.from}  ->  ${p.to}`);
  });

  console.log(plans.length ? `Renamed ${plans.length} tabs.` : "All tabs already match their beds.");
}
