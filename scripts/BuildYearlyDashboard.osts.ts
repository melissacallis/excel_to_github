// Office Script: builds a "Yearly Dashboard" sheet from the Yearly Log.
// Excel on the web > Automate > New Script > paste > save as "Build Yearly Dashboard" > Run.
//
// On the dashboard, pick a Year and a Period (Whole Year or a month) and everything
// updates by itself: the summary numbers and all three charts.
//   1. Bar chart: discharges to each Shelter, Treatment Center and Sober Living place
//   2. Column chart: admits per week (Week 1-5) for each month
//   3. Line chart: admits per month for each Referral source
// It reads the Yearly Log, plus this month's Discharge sheet if "Include Discharge
// sheet" is Yes (so clients not archived yet still count).
//
// It only creates/replaces the "Yearly Dashboard" sheet; nothing else is changed.
// Run it again if new places or referral sources show up, so they get added to the charts.

const SHEETS = ["Yearly Log", "Discharge"];
const FIRST = 5;          // first data row on the log sheets (headers on row 4)
const LAST = 3000;
const MONTHS = ["January", "February", "March", "April", "May", "June", "July",
  "August", "September", "October", "November", "December"];
const PALETTE = ["#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B",
  "#E377C2", "#7F7F7F", "#BCBD22", "#17BECF", "#393B79", "#AD494A"];
const TYPE_COLORS: { [k: string]: string } = {
  "Shelter": "#2C7FB8", "Treatment Center": "#E07B39", "Sober Living": "#3BA272",
};

// cells on the dashboard
const YEAR = "$C$4";
const PERIOD = "$F$4";
const INCLUDE = "$J$4";
const START = "$S$2";
const END = "$S$3";

// header name -> column letter, per log sheet
type ColMap = { [sheet: string]: { [header: string]: string } };

function main(workbook: ExcelScript.Workbook) {
  const log = workbook.getWorksheet("Yearly Log");
  if (!log) throw new Error('Sheet "Yearly Log" not found.');
  const disc = workbook.getWorksheet("Discharge");
  const sources: ExcelScript.Worksheet[] = disc ? [log, disc] : [log];
  const cols: ColMap = {};

  // ---- read headers and collect place names / referral sources ----
  console.log("Step: reading the Yearly Log");
  const places: { [type: string]: string[] } = { "Shelter": [], "Treatment Center": [], "Sober Living": [] };
  const referrals: string[] = [];
  sources.forEach(ws => {
    const hdr = ws.getRange("A4:AZ4").getTexts()[0].map(h => h.trim());
    const map: { [header: string]: string } = {};
    hdr.forEach((h, i) => { if (h) map[h.toLowerCase()] = colLetter(i); });
    cols[ws.getName()] = map;
    ["admit date", "days admitted", "referral source", "shelter", "treatment center",
      "sober living", "discharge date", "met discharge plan"].forEach(h => {
      if (!map[h]) throw new Error(`Column "${h}" not found on row 4 of "${ws.getName()}".`);
    });
    const used = ws.getUsedRange();
    const lastRow = Math.max(used.getRowIndex() + used.getRowCount(), FIRST);
    const data = ws.getRange(`A${FIRST}:AZ${lastRow}`).getTexts();
    const idx = (h: string) => hdr.findIndex(x => x.toLowerCase() === h);
    Object.keys(places).forEach(type => {
      const i = idx(type.toLowerCase());
      data.forEach(r => {
        const v = r[i].trim();
        if (v && v.toUpperCase() !== "N/A" && places[type].indexOf(v) < 0) places[type].push(v);
      });
    });
    const ri = idx("referral source");
    data.forEach(r => {
      const v = r[ri].trim();
      if (v && referrals.indexOf(v) < 0) referrals.push(v);
    });
  });
  if (!disc) cols["Discharge"] = {};
  Object.keys(places).forEach(t => places[t].sort());
  referrals.sort();

  // ---- fresh sheet ----
  console.log("Step: creating the Yearly Dashboard sheet");
  const oldSheet = workbook.getWorksheet("Yearly Dashboard");
  if (oldSheet) oldSheet.delete();
  const ws = workbook.addWorksheet("Yearly Dashboard");
  ws.setShowGridlines(false);
  ws.getRange("A1:P60").getFormat().getFill().setColor("#F2F6F4");
  ws.getRange("A1:Z200").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  ws.getRange("A:A").getFormat().setColumnWidth(12);
  ws.getRange("B:M").getFormat().setColumnWidth(74);
  ws.getRange("N:Q").getFormat().setColumnWidth(20);
  ws.getRange("R:R").getFormat().setColumnWidth(120);
  ws.getRange("S:AF").getFormat().setColumnWidth(80);

  // title
  const title = ws.getRange("B2:M2");
  title.merge(false);
  ws.getRange("B2").setValue("Crisis Respite Program — Yearly Dashboard");
  block(title, "#0D3D33", "#FFFFFF", 18, true);
  title.getFormat().setIndentLevel(1);
  ws.getRange("2:2").getFormat().setRowHeight(36);
  ws.getRange("3:3").getFormat().setRowHeight(10);

  // ---- controls ----
  console.log("Step: controls");
  label(ws, "B4", "YEAR");
  ws.getRange("C4").setValue(new Date().getFullYear());
  control(ws, "C4");
  label(ws, "E4", "PERIOD");
  const per = ws.getRange("F4:G4");
  per.merge(false);
  ws.getRange("F4").setValue("Whole Year");
  control(ws, "F4:G4");
  ws.getRange("F4").getDataValidation().setRule({
    list: { inCellDropDown: true, source: ["Whole Year"].concat(MONTHS).join(",") },
  });
  const inc = ws.getRange("H4:I4");
  inc.merge(false);
  label(ws, "H4", "INCLUDE DISCHARGE SHEET");
  ws.getRange("J4").setValue(disc ? "Yes" : "No");
  control(ws, "J4");
  ws.getRange("J4").getDataValidation().setRule({ list: { inCellDropDown: true, source: "Yes,No" } });
  const hint = ws.getRange("K4:M4");
  hint.merge(false);
  ws.getRange("K4").setValue("◀ change these and the dashboard updates");
  ws.getRange("K4").getFormat().getFont().setItalic(true);
  ws.getRange("K4").getFormat().getFont().setColor("#7C8D8A");
  ws.getRange("K4").getFormat().getFont().setSize(9);
  ws.getRange("4:4").getFormat().setRowHeight(26);
  ws.getRange("5:5").getFormat().setRowHeight(10);

  // period start / end (helper cells)
  ws.getRange("R1").setValue("CHART DATA — filled in by formulas, don't type here");
  ws.getRange("R1").getFormat().getFont().setBold(true);
  ws.getRange("R1").getFormat().getFont().setColor("#7C8D8A");
  ws.getRange("R2").setValue("Start");
  ws.getRange("R3").setValue("End");
  const monthList = "{" + MONTHS.map(m => `"${m}"`).join(",") + "}";
  ws.getRange("S2").setFormula(`=IF(${PERIOD}="Whole Year",DATE(${YEAR},1,1),DATE(${YEAR},MATCH(${PERIOD},${monthList},0),1))`);
  ws.getRange("S3").setFormula(`=IF(${PERIOD}="Whole Year",DATE(${YEAR},12,31),EOMONTH(${START},0))`);
  ws.getRange("S2:S3").setNumberFormat("mm/dd/yyyy");

  // ---- summary tiles ----
  console.log("Step: summary tiles");
  const inPeriod = (h: string): [string, string][] => [[h, `">="&${START}`], [h, `"<="&${END}`]];
  const discharges = countAll(cols, inPeriod("discharge date"));
  const tiles: [string, string, string, string, string][] = [
    ["B", "D", "ADMITS", `=${countAll(cols, inPeriod("admit date"))}`, "0"],
    ["E", "G", "DISCHARGES", `=${discharges}`, "0"],
    ["H", "J", "MET DISCHARGE PLAN",
      `=IFERROR(${countAll(cols, inPeriod("discharge date").concat([["met discharge plan", '"✓"']]))}/${discharges},0)`, "0%"],
    ["K", "M", "AVG DAYS ADMITTED",
      `=IFERROR(${sumAll(cols, "days admitted", inPeriod("discharge date"))}/${discharges},0)`, "0.0"],
  ];
  tiles.forEach(([a, b, text, f, fmt]) => {
    const lab = ws.getRange(`${a}6:${b}6`);
    lab.merge(false);
    ws.getRange(`${a}6`).setValue(text);
    block(lab, "#FFFFFF", "#7C8D8A", 9, true);
    lab.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    const val = ws.getRange(`${a}7:${b}7`);
    val.merge(false);
    ws.getRange(`${a}7`).setFormula(f);
    ws.getRange(`${a}7`).setNumberFormat(fmt);
    block(val, "#FFFFFF", "#0D3D33", 24, true);
    val.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    const frame = ws.getRange(`${a}6:${b}7`).getFormat();
    [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeRight].forEach(s => {
      const bd = frame.getRangeBorder(s);
      bd.setStyle(ExcelScript.BorderLineStyle.continuous);
      bd.setColor("#DFE6E2");
    });
  });
  ws.getRange("6:6").getFormat().setRowHeight(18);
  ws.getRange("7:7").getFormat().setRowHeight(40);

  // ---- table 1: discharges by place ----
  console.log("Step: placement chart data");
  let row = 6;
  const t1 = row;
  ws.getRange(`R${t1}:U${t1}`).setValues([["Place", "Shelter", "Treatment Center", "Sober Living"]]);
  const types = ["Shelter", "Treatment Center", "Sober Living"];
  let placeRows = 0;
  types.forEach((type, ti) => {
    places[type].forEach(place => {
      const r = t1 + 1 + placeRows;
      ws.getRange(`R${r}`).setValue(place);
      types.forEach((t, j) => {
        const cell = ws.getRange(`${colLetter(18 + j)}${r}`);
        if (j === ti) cell.setFormula(`=${countAll(cols, inPeriod("discharge date").concat([[type.toLowerCase(), `$R${r}`]]))}`);
        else cell.setValue(0);
      });
      placeRows++;
    });
  });
  if (placeRows === 0) {
    ws.getRange(`R${t1 + 1}:U${t1 + 1}`).setValues([["(no places yet)", 0, 0, 0]]);
    placeRows = 1;
  }
  header(ws, `R${t1}:U${t1}`);
  row = t1 + placeRows + 3;

  // ---- table 2: weekly admits by month ----
  console.log("Step: weekly admits chart data");
  const t2 = row;
  ws.getRange(`R${t2}:W${t2}`).setValues([["Month", "Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]]);
  for (let m = 1; m <= 12; m++) {
    const r = t2 + m;
    ws.getRange(`R${r}`).setValue(MONTHS[m - 1].substring(0, 3));
    for (let w = 1; w <= 5; w++) {
      const cnt = countAll(cols, [["admit date", '">="&s'], ["admit date", '"<="&e']]);
      ws.getRange(`${colLetter(17 + w)}${r}`).setFormula(
        `=LET(ms,DATE(${YEAR},${m},1),s,DATE(${YEAR},${m},${1 + 7 * (w - 1)}),` +
        `e,MIN(DATE(${YEAR},${m},${7 * w}),EOMONTH(ms,0)),` +
        `IF(OR(ms<${START},ms>${END}),NA(),IF(s>e,0,${cnt})))`);
    }
  }
  header(ws, `R${t2}:W${t2}`);
  row = t2 + 12 + 3;

  // ---- table 3: admits by referral source by month ----
  console.log("Step: referral chart data");
  const t3 = row;
  const refs = referrals.length ? referrals : ["(no referrals yet)"];
  const lastRefCol = colLetter(17 + refs.length);
  ws.getRange(`R${t3}:${lastRefCol}${t3}`).setValues([["Month"].concat(refs)]);
  for (let m = 1; m <= 12; m++) {
    const r = t3 + m;
    ws.getRange(`R${r}`).setValue(MONTHS[m - 1].substring(0, 3));
    refs.forEach((src, j) => {
      const colL = colLetter(18 + j);
      const cnt = countAll(cols, [["admit date", '">="&ms'], ["admit date", '"<="&me'],
        ["referral source", `${colL}$${t3}`]]);
      ws.getRange(`${colL}${r}`).setFormula(
        `=LET(ms,DATE(${YEAR},${m},1),me,EOMONTH(ms,0),IF(OR(ms<${START},ms>${END}),NA(),${cnt}))`);
    });
  }
  header(ws, `R${t3}:${lastRefCol}${t3}`);

  // ---- charts ----
  console.log("Step: placement chart");
  const c1 = ws.addChart(ExcelScript.ChartType.barStacked,
    ws.getRange(`R${t1}:U${t1 + placeRows}`), ExcelScript.ChartSeriesBy.columns);
  c1.setPosition("B9", "G30");
  style(c1, "Discharged to: Shelter / Treatment / Sober Living");
  c1.getSeries().forEach(s => {
    const color = TYPE_COLORS[s.getName()];
    if (color) s.getFormat().getFill().setSolidColor(color);
    s.setGapWidth(60);
  });
  c1.getAxes().getCategoryAxis().setReversePlotOrder(true);

  console.log("Step: weekly admits chart");
  const c2 = ws.addChart(ExcelScript.ChartType.columnClustered,
    ws.getRange(`R${t2}:W${t2 + 12}`), ExcelScript.ChartSeriesBy.columns);
  c2.setPosition("H9", "M30");
  style(c2, "Weekly admits by month");
  c2.getSeries().forEach((s, i) => s.getFormat().getFill().setSolidColor(["#9EC2B3", "#5FA58D", "#12735C", "#0D3D33", "#F2A541"][i]));

  console.log("Step: referral chart");
  const c3 = ws.addChart(ExcelScript.ChartType.lineMarkers,
    ws.getRange(`R${t3}:${lastRefCol}${t3 + 12}`), ExcelScript.ChartSeriesBy.columns);
  c3.setPosition("B32", "M56");
  style(c3, "Admits by referral source");
  c3.getSeries().forEach((s, i) => {
    const color = PALETTE[i % PALETTE.length];
    s.getFormat().getLine().setColor(color);
    s.setMarkerBackgroundColor(color);
    s.setMarkerForegroundColor(color);
    s.setMarkerSize(6);
  });

  ws.activate();
  ws.getRange("F4").select();
  console.log("Done! Check the \"Yearly Dashboard\" sheet.");
}

// COUNTIFS over the Yearly Log, plus the Discharge sheet when INCLUDE is "Yes".
// conds: [header name (lowercase), criteria expression]
function countAll(cols: ColMap, conds: [string, string][]): string {
  const part = (sheet: string) =>
    "COUNTIFS(" + conds.map(([h, crit]) => `${rng(cols, sheet, h)},${crit}`).join(",") + ")";
  const d = cols["Discharge"] && cols["Discharge"]["admit date"] ? `+IF(${INCLUDE}="Yes",${part("Discharge")},0)` : "";
  return `(${part("Yearly Log")}${d})`;
}

function sumAll(cols: ColMap, sumHeader: string, conds: [string, string][]): string {
  const part = (sheet: string) =>
    `SUMIFS(${rng(cols, sheet, sumHeader)},` + conds.map(([h, crit]) => `${rng(cols, sheet, h)},${crit}`).join(",") + ")";
  const d = cols["Discharge"] && cols["Discharge"]["admit date"] ? `+IF(${INCLUDE}="Yes",${part("Discharge")},0)` : "";
  return `(${part("Yearly Log")}${d})`;
}

function rng(cols: ColMap, sheet: string, h: string): string {
  const c = cols[sheet][h];
  return `'${sheet}'!$${c}$${FIRST}:$${c}$${LAST}`;
}

function colLetter(i: number): string {
  let s = "";
  let n = i + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function block(r: ExcelScript.Range, fill: string, font: string, size: number, bold: boolean) {
  const f = r.getFormat();
  f.getFill().setColor(fill);
  f.getFont().setColor(font);
  f.getFont().setSize(size);
  f.getFont().setBold(bold);
  f.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function label(ws: ExcelScript.Worksheet, addr: string, text: string) {
  const r = ws.getRange(addr);
  r.setValue(text);
  r.getFormat().getFont().setBold(true);
  r.getFormat().getFont().setSize(9);
  r.getFormat().getFont().setColor("#48605C");
  r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
}

function control(ws: ExcelScript.Worksheet, addr: string) {
  const f = ws.getRange(addr).getFormat();
  f.getFill().setColor("#FFF7D6");
  f.getFont().setBold(true);
  f.getFont().setSize(12);
  f.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeRight].forEach(s => {
    const b = f.getRangeBorder(s);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setColor("#E0B84C");
  });
}

function header(ws: ExcelScript.Worksheet, addr: string) {
  const f = ws.getRange(addr).getFormat();
  f.getFont().setBold(true);
  f.getFill().setColor("#DFE6E2");
}

function style(ch: ExcelScript.Chart, text: string) {
  ch.getTitle().setText(text);
  ch.getTitle().getFormat().getFont().setSize(13);
  ch.getTitle().getFormat().getFont().setBold(true);
  ch.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
  ch.getFormat().getBorder().setColor("#DFE6E2");
}
