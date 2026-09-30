// Office Script: builds the "Card Dashboard" sheet as cards that look like the website.
// Excel on the web > Automate > New Script > paste > save as "Build Card Dashboard" > Run.
//
// Everything on the cards is a formula that reads the Client Status Board (and the
// ✓/✗ needs on the Client Needs Tracker), so the cards update by themselves as you
// type. Run this again only if you add or remove beds, or want to reset the layout.
// WARNING: it clears the Card Dashboard sheet first.

const SB_HDR = "'Client Status Board'!$A$4:$Z$4";
const SB_DATA = "'Client Status Board'!$A$5:$Z$300";
const BED_COL = `INDEX(${SB_DATA},0,XMATCH("Bed",${SB_HDR}))`;

const C = {
  ink: "#1C2B29", muted: "#7C8D8A", line: "#DFE6E2", paper: "#F2F6F4", card: "#FFFFFF",
  tealDeep: "#0D3D33",
  maleBg: "#DCEAF6", maleFg: "#2C5F8A", femaleBg: "#E7DDF5", femaleFg: "#5E3A96",
  red: "#A32D2D", redBg: "#FBE9E9", amber: "#8A5A0B", amberBg: "#FDF1DE",
  green: "#276A2F", greenBg: "#E9F3E8", blue: "#1F4F8F", blueBg: "#DBE9FB",
  gray: "#7C8D8A", grayBg: "#EEF1F0", infoBg: "#E4F0FA",
};

const FIRST_CARD_ROW = 10;
const CARD_ROWS = 10;      // rows used by one card
const CARD_STEP = 11;      // card + one spacer row

function main(workbook: ExcelScript.Workbook) {
  const sb = workbook.getWorksheet("Client Status Board");
  if (!sb) throw new Error('Sheet "Client Status Board" not found.');
  let ws = workbook.getWorksheet("Card Dashboard");
  if (!ws) ws = workbook.addWorksheet("Card Dashboard");

  // count male / female beds on the status board
  const hdr = sb.getRange("A4:Z4").getTexts()[0].map(h => h.trim().toLowerCase());
  const bedIdx = hdr.indexOf("bed");
  if (bedIdx < 0) throw new Error('No "Bed" column found on row 4 of the Client Status Board.');
  const bedTexts = sb.getRange("A5:Z300").getTexts().map(r => r[bedIdx].trim().toUpperCase());
  const nMale = bedTexts.filter(b => /^M\s*-/.test(b)).length;
  const nFemale = bedTexts.filter(b => /^F\s*-/.test(b)).length;
  const nCards = Math.max(nMale, nFemale, 1);
  const lastRow = FIRST_CARD_ROW + nCards * CARD_STEP;

  // ---- reset the sheet ----
  const all = ws.getRange("A1:Z400");
  all.clearAllConditionalFormats();
  all.unmerge();
  all.clear(ExcelScript.ClearApplyTo.all);
  ws.setShowGridlines(false);
  ws.getRange(`A1:K${lastRow}`).getFormat().getFill().setColor(C.paper);
  ws.getRange(`A1:K${lastRow}`).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  ws.getRange(`A1:K${lastRow}`).getFormat().getFont().setName("Aptos");

  const widths: [string, number][] = [
    ["A", 12], ["B", 62], ["C", 112], ["D", 62], ["E", 124],
    ["F", 18], ["G", 62], ["H", 112], ["I", 62], ["J", 124], ["K", 12],
  ];
  widths.forEach(([col, w]) => ws.getRange(`${col}:${col}`).getFormat().setColumnWidth(w));

  // ---- title ----
  ws.getRange("1:1").getFormat().setRowHeight(8);
  const title = ws.getRange("B2:J2");
  title.merge(false);
  ws.getRange("B2").setValue("Crisis Respite Program — Client Status Board");
  styleBlock(title, C.tealDeep, "#FFFFFF", 18, true);
  title.getFormat().setIndentLevel(1);
  title.getFormat().setRowHeight(36);

  const sub = ws.getRange("B3:J3");
  sub.merge(false);
  ws.getRange("B3").setFormula('="Live view  ·  "&(B6+D6+G6)&" clients  ·  updates automatically from the Client Status Board"');
  styleBlock(sub, C.paper, C.muted, 10, false);
  sub.getFormat().getFont().setItalic(true);
  sub.getFormat().setIndentLevel(1);
  sub.getFormat().setRowHeight(20);
  ws.getRange("4:4").getFormat().setRowHeight(8);

  // ---- summary tiles ----
  const chipCount = (t: string) =>
    `=COUNTIF($D$${FIRST_CARD_ROW}:$D$${lastRow},"${t}")+COUNTIF($I$${FIRST_CARD_ROW}:$I$${lastRow},"${t}")`;
  const tiles: [string, string, string, string][] = [
    ["B", "C", "NEEDS ACTION", "red"], ["D", "E", "UPCOMING", "amber"],
    ["G", "H", "ON TRACK", "green"], ["I", "J", "OPEN", "gray"],
  ];
  tiles.forEach(([a, b, label, tone]) => {
    const [bg, fg] = toneColors(tone);
    const lab = ws.getRange(`${a}5:${b}5`);
    lab.merge(false);
    ws.getRange(`${a}5`).setValue(label === "OPEN" ? "OPEN BEDS" : label);
    styleBlock(lab, bg, fg, 9, true);
    lab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    const val = ws.getRange(`${a}6:${b}6`);
    val.merge(false);
    ws.getRange(`${a}6`).setFormula(chipCount(label));
    styleBlock(val, bg, fg, 22, true);
    val.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  });
  ws.getRange("5:5").getFormat().setRowHeight(18);
  ws.getRange("6:6").getFormat().setRowHeight(36);
  ws.getRange("7:7").getFormat().setRowHeight(12);

  // ---- section headers ----
  const maleHead = ws.getRange("B8:E8");
  maleHead.merge(false);
  ws.getRange("B8").setValue("MALE");
  styleBlock(maleHead, C.maleFg, "#FFFFFF", 12, true);
  maleHead.getFormat().setIndentLevel(1);
  const femaleHead = ws.getRange("G8:J8");
  femaleHead.merge(false);
  ws.getRange("G8").setValue("FEMALE");
  styleBlock(femaleHead, C.femaleFg, "#FFFFFF", 12, true);
  femaleHead.getFormat().setIndentLevel(1);
  ws.getRange("8:8").getFormat().setRowHeight(24);
  ws.getRange("9:9").getFormat().setRowHeight(8);

  // ---- cards ----
  for (let k = 0; k < nCards; k++) {
    const top = FIRST_CARD_ROW + k * CARD_STEP;
    if (k < nMale) buildCard(ws, true, k + 1, top);
    if (k < nFemale) buildCard(ws, false, k + 1, top);
    ws.getRange(`${top + CARD_ROWS}:${top + CARD_ROWS}`).getFormat().setRowHeight(12);
  }

  // ---- color rules (shared by all cards) ----
  const L = (col: string) => `${col}${FIRST_CARD_ROW}:${col}${lastRow}`;
  ["D", "I"].forEach(col => {                    // urgency chip in each card header
    rule(ws, L(col), `=${col}${FIRST_CARD_ROW}="NEEDS ACTION"`, C.redBg, C.red, true);
    rule(ws, L(col), `=${col}${FIRST_CARD_ROW}="UPCOMING"`, C.amberBg, C.amber, true);
    rule(ws, L(col), `=${col}${FIRST_CARD_ROW}="ON TRACK"`, C.greenBg, C.green, true);
    rule(ws, L(col), `=${col}${FIRST_CARD_ROW}="OPEN"`, C.grayBg, C.gray, true);
  });
  ["C", "H"].forEach(col => {                    // appointment status + needs line
    const c1 = `${col}${FIRST_CARD_ROW}`;
    rule(ws, L(col), `=${c1}="Scheduled"`, C.blueBg, C.blue, true);
    rule(ws, L(col), `=ISNUMBER(SEARCH("requested",${c1}))`, "#2563EB", "#FFFFFF", true);
    rule(ws, L(col), `=${c1}="Completed"`, C.greenBg, C.green, true);
    rule(ws, L(col), `=LEFT(${c1},1)="✗"`, C.amberBg, C.amber, false);
  });
  ["E", "J"].forEach(col => {                    // scripts status, time, bed days
    const c1 = `${col}${FIRST_CARD_ROW}`;
    rule(ws, L(col), `=OR(ISNUMBER(SEARCH("submitted",${c1})),ISNUMBER(SEARCH("pending",${c1})))`, "#FFE14D", "#6B5200", true);
    rule(ws, L(col), `=${c1}="Picked Up"`, C.greenBg, C.green, true);
    rule(ws, L(col), `=${c1}="Dropped off"`, C.blueBg, C.blue, true);
    rule(ws, L(col), `=${c1}="⚠ time needed"`, "", C.red, true);
    rule(ws, L(col), `=AND(ISNUMBER(${c1}),${c1}>5)`, C.redBg, C.red, true);
  });

  ws.activate();
  ws.getRange("A1").select();
}

// One card: 4 columns wide, 10 rows tall.
function buildCard(ws: ExcelScript.Worksheet, male: boolean, k: number, top: number) {
  const [a, b, c, d] = male ? ["B", "C", "D", "E"] : ["G", "H", "I", "J"];
  const bg = male ? C.maleBg : C.femaleBg;
  const fg = male ? C.maleFg : C.femaleFg;
  const bedRef = `$${b}$${top}`;
  const iniRef = `$${a}$${top}`;
  const lk = (field: string) =>
    `XLOOKUP(${bedRef},${BED_COL},INDEX(${SB_DATA},0,XMATCH("${field}",${SB_HDR})),"")`;
  const val = (field: string) =>
    `=IF(${iniRef}="—","",IFERROR(LET(v,${lk(field)},IF(v="","—",v)),"—"))`;
  const set = (addr: string, f: string) => ws.getRange(addr).setFormula(f);
  const label = (addr: string, text: string) => {
    const r = ws.getRange(addr);
    r.setValue(text);
    r.getFormat().getFont().setSize(9);
    r.getFormat().getFont().setBold(true);
    r.getFormat().getFont().setColor(C.muted);
    r.getFormat().setIndentLevel(1);
  };

  const card = ws.getRange(`${a}${top}:${d}${top + CARD_ROWS - 1}`);
  card.getFormat().getFill().setColor(C.card);
  card.getFormat().getFont().setColor(C.ink);
  card.getFormat().getFont().setSize(11);

  // header: initials | bed | urgency chip
  const g = male ? "M" : "F";
  set(`${b}${top}`, `=IFERROR(INDEX(FILTER(${BED_COL},LEFT(${BED_COL},1)="${g}"),${k}),"")`);
  set(`${a}${top}`, `=IFERROR(LET(v,${lk("Client Initials")},IF(OR(${bedRef}="",v=""),"—",v)),"—")`);
  const chip = ws.getRange(`${c}${top}:${d}${top}`);
  chip.merge(false);
  set(`${c}${top}`,
    `=IF(${iniRef}="—",IF(${bedRef}="","","OPEN"),LET(dt,${b}${top + 4},` +
    `IF(NOT(ISNUMBER(dt)),"NEEDS ACTION",IF(dt<TODAY(),"NEEDS ACTION",IF(dt<=TODAY()+2,"UPCOMING","ON TRACK")))))`);
  const av = ws.getRange(`${a}${top}`).getFormat();
  av.getFill().setColor(bg);
  av.getFont().setColor(fg);
  av.getFont().setBold(true);
  av.getFont().setSize(14);
  av.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  const bedCell = ws.getRange(`${b}${top}`).getFormat();
  bedCell.getFont().setBold(true);
  bedCell.getFont().setSize(14);
  bedCell.setIndentLevel(1);
  chip.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  chip.getFormat().getFont().setSize(10);
  chip.getFormat().getFont().setBold(true);
  ws.getRange(`${top}:${top}`).getFormat().setRowHeight(32);

  // field rows
  const rows: [string, string, string, string][] = [
    ["ID", "Member ID", "LOC", "LOC"],
    ["CM", "Case Manager", "REFERRAL", "Referral source"],
    ["APPT", "Appointment Status", "DR", "Doctor"],
    ["DATE", "Appointment Date", "TIME", ""],
    ["PHARMACY", "Pharmacy", "SCRIPTS", "Pharmacy Status"],
    ["ADMITTED", "Admit Date", "BED DAYS", ""],
  ];
  rows.forEach(([l1, f1, l2, f2], i) => {
    const r = top + 1 + i;
    label(`${a}${r}`, l1);
    set(`${b}${r}`, val(f1));
    label(`${c}${r}`, l2);
    if (f2) set(`${d}${r}`, val(f2));
    ws.getRange(`${r}:${r}`).getFormat().setRowHeight(19);
    ws.getRange(`${b}${r}`).getFormat().setIndentLevel(1);
    ws.getRange(`${d}${r}`).getFormat().setIndentLevel(1);
  });
  // time: flag scheduled appointments with no time
  set(`${d}${top + 4}`,
    `=IF(${iniRef}="—","",IFERROR(LET(t,${lk("Appointment Time")},` +
    `IF(t="",IF(${b}${top + 3}="Scheduled","⚠ time needed","—"),t)),"—"))`);
  // bed days since admit
  set(`${d}${top + 6}`, `=IF(ISNUMBER(${b}${top + 6}),TODAY()-${b}${top + 6},IF(${iniRef}="—","","—"))`);
  ws.getRange(`${b}${top + 4}`).setNumberFormat("mm/dd/yyyy");
  ws.getRange(`${b}${top + 6}`).setNumberFormat("mm/dd/yyyy");
  ws.getRange(`${d}${top + 4}`).setNumberFormat("h:mm AM/PM");
  ws.getRange(`${d}${top + 6}`).setNumberFormat("0");
  ws.getRange(`${b}${top + 1}`).setNumberFormat("0");
  [top + 4, top + 6].forEach(r => ws.getRange(`${b}${r}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left));
  ws.getRange(`${d}${top + 6}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  ws.getRange(`${b}${top + 1}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  // wide rows: meds, needs, notes
  const wide: [number, string, string][] = [
    [top + 7, "MEDS", val("Med Times")],
    [top + 8, "NEEDS",
      `=IF(${iniRef}="—","",IFERROR(LET(h,'Client Needs Tracker'!$J$4:$Q$4,` +
      `v,XLOOKUP(${bedRef},'Client Needs Tracker'!$S$5:$S$300,'Client Needs Tracker'!$J$5:$Q$300,""),` +
      `m,IFERROR(TEXTJOIN(", ",TRUE,FILTER(h,v="✗")),""),ok,IFERROR(TEXTJOIN(", ",TRUE,FILTER(h,v="✓")),""),` +
      `IF(m="",IF(ok="","—","✓ Has: "&ok),"✗ Needs: "&m&IF(ok="","","     ✓ Has: "&ok))),"—"))`],
    [top + 9, "NOTES", val("Notes")],
  ];
  wide.forEach(([r, l, f]) => {
    label(`${a}${r}`, l);
    const span = ws.getRange(`${b}${r}:${d}${r}`);
    span.merge(false);
    ws.getRange(`${b}${r}`).setFormula(f);
    span.getFormat().setWrapText(true);
    span.getFormat().setIndentLevel(1);
    ws.getRange(`${r}:${r}`).getFormat().setRowHeight(30);
  });
  const meds = ws.getRange(`${a}${top + 7}:${d}${top + 7}`).getFormat();
  meds.getFill().setColor(C.infoBg);
  ws.getRange(`${b}${top + 7}`).getFormat().getFont().setColor(C.blue);
  ws.getRange(`${b}${top + 7}`).getFormat().getFont().setBold(true);
  ws.getRange(`${b}${top + 9}`).getFormat().getFont().setItalic(true);

  // borders: light frame, line under the header, thick colored left edge
  const f = card.getFormat();
  [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
   ExcelScript.BorderIndex.edgeRight].forEach(side => border(f.getRangeBorder(side), C.line, ExcelScript.BorderWeight.thin));
  border(f.getRangeBorder(ExcelScript.BorderIndex.edgeLeft), fg, ExcelScript.BorderWeight.thick);
  border(ws.getRange(`${a}${top}:${d}${top}`).getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom),
    C.line, ExcelScript.BorderWeight.thin);

  // open bed: grey out the card text
  rule(ws, `${a}${top + 1}:${d}${top + CARD_ROWS - 1}`, `=${iniRef}="—"`, "", "#B8C2BF", false);
  rule(ws, `${a}${top}:${b}${top}`, `=${iniRef}="—"`, C.grayBg, C.gray, false);
}

function border(b: ExcelScript.RangeBorder, color: string, weight: ExcelScript.BorderWeight) {
  b.setStyle(ExcelScript.BorderLineStyle.continuous);
  b.setColor(color);
  b.setWeight(weight);
}

function styleBlock(r: ExcelScript.Range, fill: string, font: string, size: number, bold: boolean) {
  const f = r.getFormat();
  f.getFill().setColor(fill);
  f.getFont().setColor(font);
  f.getFont().setSize(size);
  f.getFont().setBold(bold);
  f.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function toneColors(tone: string): [string, string] {
  if (tone === "red") return [C.redBg, C.red];
  if (tone === "amber") return [C.amberBg, C.amber];
  if (tone === "green") return [C.greenBg, C.green];
  return [C.grayBg, C.gray];
}

function rule(ws: ExcelScript.Worksheet, address: string, formula: string,
              fill: string, font: string, bold: boolean) {
  const cf = ws.getRange(address).addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  const custom = cf.getCustom();
  custom.getRule().setFormula(formula);
  if (fill) custom.getFormat().getFill().setColor(fill);
  if (font) custom.getFormat().getFont().setColor(font);
  if (bold) custom.getFormat().getFont().setBold(true);
}
