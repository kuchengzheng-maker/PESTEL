const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE";                     // 13.3 x 7.5
p.author = "PESTEL — U.S. EV Market";
p.title  = "The American EV Price Floor";

// ---- tokens -------------------------------------------------------------
const DARK="0E2A3A", INK="13202B", MUT="5A6E7C", TINT="E9F0F4", WHITE="FFFFFF";
const BLUE="3081AE", OCHRE="D99A16", OXIDE="C0392B", TEAL="15A48E", PALE="BCD3E0";
const SERIF="Cambria", SANS="Calibri";
const W=13.3, M=0.62;

const axis = () => ({
  catAxisLabelColor: MUT, valAxisLabelColor: MUT,
  catAxisLabelFontFace: SANS, valAxisLabelFontFace: SANS,
  catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  valGridLine: { color: "DDE5EA", size: 1 }, catGridLine: { style: "none" },
  catAxisLineShow: false, valAxisLineShow: false,
  dataLabelFontFace: SANS, dataLabelFontSize: 11, dataLabelColor: INK,
  showValue: true, dataLabelPosition: "outEnd",
});

function head(s, kicker, title, sub){
  s.addText(kicker, { x:M, y:0.42, w:W-2*M, h:0.24, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:11, bold:true, charSpacing:2, color:BLUE });
  s.addText(title, { x:M, y:0.70, w:W-2*M, h:0.72, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:29, bold:true, color:INK });
  if(sub) s.addText(sub, { x:M, y:1.44, w:W-2*M-0.2, h:0.42, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:14, color:MUT });
}
function src(s, t){
  s.addText(t, { x:M, y:6.92, w:W-2*M, h:0.26, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:9, color:MUT, italic:true });
}
// repeated motif: a filled circle badge
function badge(s, x, y, label, col){
  s.addShape(p.ShapeType.ellipse, { x, y, w:0.46, h:0.46, fill:{color:col} });
  s.addText(label, { x, y:y+0.005, w:0.46, h:0.45, isTextBox:true, margin:0,
    align:"center", valign:"middle", fontFace:SANS, fontSize:14, bold:true, color:WHITE });
}
function card(s, x, y, w, h, fill){
  s.addShape(p.ShapeType.roundRect, { x, y, w, h, rectRadius:0.06,
    fill:{color:fill||TINT}, line:{color:"D6E2E9", width:0.75} });
}

// ===== 1. TITLE ==========================================================
let s = p.addSlide(); s.background = { color: DARK };
s.addText("PESTEL · U.S. ELECTRIC VEHICLE MARKET · FACTOR 1", { x:M, y:1.55, w:11, h:0.3,
  isTextBox:true, margin:0, fontFace:SANS, fontSize:12, bold:true, charSpacing:2.5, color:PALE });
s.addText("A normal price level,\nan abnormal price floor", { x:M, y:2.0, w:11.2, h:1.9,
  isTextBox:true, margin:0, fontFace:SERIF, fontSize:50, bold:true, color:WHITE, lineSpacing:52 });
s.addText("How the prohibition on Chinese EVs and the expiry of the $7,500 federal tax credit reshaped what an electric car costs an American — and who can still buy one.",
  { x:M, y:4.05, w:9.6, h:0.85, isTextBox:true, margin:0, fontFace:SANS, fontSize:15, color:PALE });
s.addShape(p.ShapeType.rect, { x:M, y:5.2, w:2.1, h:0.02, fill:{color:OCHRE} });
s.addText("Baseline September 2026", { x:M, y:5.42, w:6, h:0.3, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:12, color:PALE });
s.addNotes("Political factor of a PESTEL analysis of the U.S. EV market. Core thesis: the policy effect is a missing market segment, not a higher price level.");

// ===== 2. CORE FINDING ===================================================
s = p.addSlide();
head(s, "THE CORE FINDING", "Policy did not make EVs expensive. It made the cheap ones illegal.",
  "Two decisions moved together: one removed the lowest-cost supply, the other removed the largest demand-side subsidy.");
const tiles = [
  ["$28,995","Cheapest new EV\nin the United States", OXIDE],
  ["$18,051","Cheapest new EV\nin Canada", BLUE],
  ["5.8%","U.S. BEV share of\nnew sales, Q1 2026", BLUE],
  ["−28%","U.S. new BEV sales,\nQ1 2026 vs Q1 2025", OXIDE],
];
tiles.forEach((t,i)=>{
  const x = M + i*3.05;
  card(s, x, 2.28, 2.82, 1.72);
  s.addText(t[0], { x:x+0.22, y:2.44, w:2.4, h:0.72, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:34, bold:true, color:t[2] });
  s.addText(t[1], { x:x+0.22, y:3.18, w:2.45, h:0.66, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:11.5, color:MUT });
});
card(s, M, 4.36, W-2*M, 2.24, DARK);
s.addText("The United States has a normal EV market in the middle and at the top, and no bottom at all.",
  { x:M+0.42, y:4.66, w:11.6, h:0.5, isTextBox:true, margin:0, fontFace:SERIF, fontSize:20, bold:true, color:WHITE });
s.addText([
  { text:"Adjusted for local incomes and vehicle price levels, the U.S. is level with Germany — Americans are not paying more per car.", options:{bullet:true, breakLine:true, paraSpaceAfter:5} },
  { text:"The entire measurable policy effect sits in the A and B size classes, where the U.S. figure is not a higher number but a blank cell.", options:{bullet:true, breakLine:true, paraSpaceAfter:5} },
  { text:"That makes this a finding about distribution, not about price level.", options:{bullet:true} },
], { x:M+0.42, y:5.26, w:11.5, h:1.1, isTextBox:true, margin:0, fontFace:SANS, fontSize:12.5, color:PALE });
src(s, "Cox Automotive Q1 2026 EV sales report · InsideEVs · BYD Canada launch pricing · OECD household disposable income");
s.addNotes("Headline: distribution, not level. This framing survives every normalization test in the analysis.");

// ===== 3. THREE WALLS ====================================================
s = p.addSlide();
head(s, "MECHANISM", "The \"ban\" is three instruments — and only one is a ban",
  "Tariffs raise cost and can be engineered around. The connected-vehicle rule attaches to manufacturer ownership, so no price clears it.");
const walls = [
  ["1","Section 301 tariff","100% on Chinese-built EVs since Sept 2024, stacked on the 2.5% MFN duty.","PRICE INSTRUMENT — avoidable by assembling outside China", OCHRE],
  ["2","Section 232 tariff","25% on all imported vehicles since April 2025. Adds $5,000–8,900 to imported models.","PRICE INSTRUMENT — partly avoidable via USMCA content", OCHRE],
  ["3","BIS connected-vehicle rule","From model year 2027, bars any China-controlled manufacturer from selling a connected vehicle — regardless of where it is built.","PROHIBITION — no price makes the car legal", OXIDE],
];
walls.forEach((w,i)=>{
  const x = M + i*4.09;
  card(s, x, 2.30, 3.82, 3.55, i===2 ? "FBEDEB" : TINT);
  badge(s, x+0.30, 2.58, w[0], w[4]);
  s.addText(w[1], { x:x+0.30, y:3.20, w:3.25, h:0.6, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:17, bold:true, color:INK });
  s.addText(w[2], { x:x+0.30, y:3.86, w:3.25, h:1.25, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:12, color:INK });
  s.addText(w[3], { x:x+0.30, y:5.16, w:3.25, h:0.55, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:10, bold:true, color:w[4] });
});
s.addText("Wall 3 is decisive: BYD cannot tariff-engineer its way in by assembling in Mexico or Hungary, because the prohibition attaches to who owns the manufacturer, not to where the car was made.",
  { x:M, y:6.10, w:W-2*M, h:0.6, isTextBox:true, margin:0, fontFace:SANS, fontSize:13, italic:true, color:INK });
src(s, "BIS final rule, Jan 2025 (effective Mar 2025) · CRS, Section 232 Automotive Tariffs · USTR Section 301");
s.addNotes("The key analytical move of the whole section: separate price instruments from the prohibition.");

// ===== 4. ENTRY PRICE BY MARKET =========================================
s = p.addSlide();
head(s, "EVIDENCE 1 — THE FLOOR", "The cheapest electric car you can buy, by market",
  "Entry-level battery-electric vehicle, converted to USD excluding consumption tax.");
s.addChart(p.ChartType.bar, [{
  name:"Entry EV price (USD, ex-tax)",
  labels:["China","Canada","Europe (DE)","Mexico","United States"],
  values:[9205, 18051, 19533, 19583, 28995],
}], Object.assign(axis(), {
  x:M, y:2.15, w:7.7, h:4.25, barDir:"col",
  chartColors:[BLUE, BLUE, BLUE, BLUE, OXIDE],
  showLegend:false, showTitle:false, valAxisMaxVal:32000,
  dataLabelFormatCode:'$#,##0', barGapWidthPct:55,
}));
card(s, 8.62, 2.15, 4.06, 4.25);
s.addText("$9,000", { x:8.92, y:2.42, w:3.5, h:0.55, isTextBox:true, margin:0,
  fontFace:SERIF, fontSize:30, bold:true, color:OXIDE });
s.addText("higher than Canada's entry point — and the U.S. has no product at all below $28,995.",
  { x:8.92, y:3.02, w:3.5, h:0.75, isTextBox:true, margin:0, fontFace:SANS, fontSize:12.5, color:INK });
s.addText([
  { text:"Europe alone offers four competing cars under $20,000 ex-VAT.", options:{bullet:true, breakLine:true, paraSpaceAfter:7} },
  { text:"The gap is not a markup on a car Americans can buy. It is a class of car that does not exist here.", options:{bullet:true, breakLine:true, paraSpaceAfter:7} },
  { text:"Segments differ: a Dolphin Surf is a 30 kWh city car; a Bolt is a 65 kWh, 262-mile hatch.", options:{bullet:true} },
], { x:8.92, y:3.86, w:3.5, h:2.3, isTextBox:true, margin:0, fontFace:SANS, fontSize:11, color:MUT });
src(s, "InsideEVs · BYD China/Mexico/Europe listings · BYD Canada launch pricing · rates: CNY 6.72, MXN 17.6, EUR 0.86, CAD 1.385");
s.addNotes("Caveat stated on the slide: these are price-point comparisons, not value-for-money comparisons.");

// ===== 5. EXPORT ESCALATION ==============================================
s = p.addSlide();
head(s, "EVIDENCE 2 — CALIBRATION", "Chinese EVs stop being cheap at any border they cross",
  "Same vehicle, indexed to its Chinese price = 100. The BYD Seagull doubles on export; a locally-built Tesla barely moves.");
s.addChart(p.ChartType.bar, [
  { name:"BYD Seagull (exported from China)", labels:["China","Canada","Germany","Mexico"], values:[100,196,212,213] },
  { name:"Tesla Model Y (built locally in each market)", labels:["China","Canada","Germany","Mexico"], values:[100,104,113,113] },
], Object.assign(axis(), {
  x:M, y:2.28, w:7.9, h:3.9, barDir:"col",
  chartColors:[OCHRE, BLUE], showLegend:true, legendPos:"b", legendFontFace:SANS,
  legendFontSize:11, legendColor:INK, showTitle:false, valAxisMaxVal:240,
  dataLabelFormatCode:'0', barGapWidthPct:45,
}));
card(s, 8.78, 2.28, 3.90, 3.9, "FDF6E8");
s.addText("The correction", { x:9.06, y:2.52, w:3.4, h:0.34, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:11, bold:true, charSpacing:1.5, color:OCHRE });
s.addText("The rung missing from the U.S. ladder is a $20,000–23,000 car — not a $10,000 one.",
  { x:9.06, y:2.90, w:3.36, h:1.0, isTextBox:true, margin:0, fontFace:SERIF, fontSize:17, bold:true, color:INK });
s.addText("Every export market roughly doubles China's domestic price, including those where the car is perfectly legal. China's home price is therefore not a valid policy counterfactual — Canada, Mexico and Europe are.\n\nAnd where a maker builds locally, prices converge within 15% across all five markets, the U.S. included.",
  { x:9.06, y:3.94, w:3.36, h:2.1, isTextBox:true, margin:0, fontFace:SANS, fontSize:11.5, color:INK });
src(s, "BYD and Tesla list prices, Sept 2026, converted ex-consumption-tax. U.S. Seagull row omitted: not sold at any price.");
s.addNotes("This slide deliberately argues against the strongest version of the case. It is what makes the rest credible.");

// ===== 6. FLOOR VS CEILING ==============================================
s = p.addSlide();
head(s, "EVIDENCE 3 — THE LADDER", "They kept building variety. They built it upward.",
  "Bottom and top of each market's full electric range, across all five size classes.");
s.addChart(p.ChartType.bar, [
  { name:"Floor — cheapest EV", labels:["China","Europe (DE)","Canada","Mexico","United States"], values:[7480,16514,18051,19583,28995] },
  { name:"Ceiling — dearest EV",  labels:["China","Europe (DE)","Canada","Mexico","United States"], values:[59234,75532,61585,79290,127405] },
], Object.assign(axis(), {
  x:M, y:2.24, w:8.15, h:3.95, barDir:"col",
  chartColors:[BLUE, OCHRE], showLegend:true, legendPos:"b", legendFontFace:SANS,
  legendFontSize:11, legendColor:INK, showTitle:false,
  dataLabelFormatCode:'$#,##0', barGapWidthPct:45, dataLabelFontSize:9,
}));
card(s, 9.03, 2.24, 3.65, 3.95);
s.addText("The U.S. has both the highest floor and by far the highest ceiling.",
  { x:9.30, y:2.50, w:3.15, h:0.95, isTextBox:true, margin:0, fontFace:SERIF, fontSize:17, bold:true, color:INK });
s.addText([
  { text:"The Escalade IQ at $127,405 sits ~$48,000 above any other market's dearest EV.", options:{bullet:true, breakLine:true, paraSpaceAfter:8} },
  { text:"Rivian and Volvo occupy a $77k–126k band with no Canadian or Mexican counterpart.", options:{bullet:true, breakLine:true, paraSpaceAfter:8} },
  { text:"This is the 1981 Feenstra upgrading result showing up in a product catalogue.", options:{bullet:true} },
], { x:9.30, y:3.55, w:3.15, h:2.4, isTextBox:true, margin:0, fontFace:SANS, fontSize:11.5, color:INK });
src(s, "Appendix A price matrix, five size classes, five markets, Sept 2026");
s.addNotes("Feenstra 1985: under the Japanese VER two-thirds of the price rise was trim upgrading, not like-for-like increases.");

// ===== 7. ADOPTION DIVERGENCE ===========================================
s = p.addSlide();
head(s, "EVIDENCE 4 — THE OUTCOME", "The U.S. is the only one of these markets going backwards",
  "Change in EV sales, 2026 versus 2025.");
s.addChart(p.ChartType.bar, [{
  name:"Change in EV sales, 2026 vs 2025",
  labels:["Mexico","Canada","Europe","United States"],
  values:[44, 40, 35, -28],
}], Object.assign(axis(), {
  x:M, y:2.22, w:7.6, h:4.05, barDir:"col",
  chartColors:[BLUE, BLUE, BLUE, OXIDE],
  showLegend:false, showTitle:false, dataLabelFormatCode:'+0"%";-0"%"', barGapWidthPct:55,
}));
card(s, 8.52, 2.22, 4.16, 1.92);
s.addText("BEV share of new sales, 2026", { x:8.80, y:2.40, w:3.6, h:0.3, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:10.5, bold:true, charSpacing:1, color:BLUE });
s.addText([
  { text:"Europe  22%     Canada  ~10.5%", options:{ breakLine:true, paraSpaceAfter:4 } },
  { text:"United States  5.8%     Mexico  ~1.8%", options:{} },
], { x:8.80, y:2.76, w:3.6, h:0.9, isTextBox:true, margin:0, fontFace:SANS, fontSize:13, color:INK });
card(s, 8.52, 4.30, 4.16, 1.97, "FDF6E8");
s.addText("What this concedes", { x:8.80, y:4.48, w:3.6, h:0.3, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:10.5, bold:true, charSpacing:1, color:OCHRE });
s.addText("Mexico has near-cheapest EVs and the lowest BEV share of the four. Cheap supply is necessary but not sufficient — income and charging bind there. The U.S. has both, which is why affordable supply would convert most efficiently here.",
  { x:8.80, y:4.82, w:3.6, h:1.35, isTextBox:true, margin:0, fontFace:SANS, fontSize:11, color:INK });
src(s, "Cox Automotive · ICCT European Car Market Monitor · J.D. Power Canada · AMIA Mexico");
s.addNotes("Four markets, four policy regimes, one outlier — and the outlier combined exclusion with subsidy withdrawal in the same twelve months.");

// ===== 8. THE CREDIT CLIFF ==============================================
s = p.addSlide();
head(s, "EVIDENCE 5 — THE CREDIT", "The $7,500 credit expired on 30 September 2025",
  "U.S. battery-electric share of new-vehicle sales, by quarter. Demand was pulled forward into Q3, then the market fell through the floor.");
s.addChart(p.ChartType.line, [{
  name:"U.S. BEV share of new sales",
  labels:["Q3 2025","Q4 2025","Q1 2026","Q2 2026"],
  values:[7.5, 5.8, 5.8, 6.0],
}], Object.assign(axis(), {
  x:M, y:2.42, w:7.5, h:3.8, chartColors:[OXIDE], lineSize:3, lineSmooth:false,
  showLegend:false, showTitle:false, valAxisMinVal:0, valAxisMaxVal:9,
  dataLabelFormatCode:'0.0"%"', dataLabelPosition:"t",
  lineDataSymbol:"circle", lineDataSymbolSize:9, lineDataSymbolLineColor:WHITE,
}));
const cliff = [
  ["$7,500","new clean vehicle credit — gone"],
  ["$4,000","used clean vehicle credit — gone"],
  ["§45W","the lease channel that reached price-sensitive buyers — gone"],
];
cliff.forEach((c,i)=>{
  card(s, 8.42, 2.42 + i*1.30, 4.26, 1.14);
  s.addText(c[0], { x:8.68, y:2.55 + i*1.30, w:1.5, h:0.44, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:21, bold:true, color:OXIDE });
  s.addText(c[1], { x:8.68, y:3.00 + i*1.30, w:3.75, h:0.48, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:11, color:INK });
});
s.addText("The replacement — a loan-interest deduction — requires U.S. final assembly, which disqualifies the Nissan Leaf and Chevrolet Equinox EV: two of America's three cheapest EVs.",
  { x:M, y:6.34, w:W-2*M, h:0.5, isTextBox:true, margin:0, fontFace:SANS, fontSize:12.5, italic:true, color:INK });
src(s, "One Big Beautiful Bill Act, signed 4 July 2025 · Cox Automotive quarterly EV share");
s.addNotes("The 45W lease repeal matters more for affordability than the headline 30D number.");

// ===== 9. EV VS ITS GAS TWIN ============================================
s = p.addSlide();
head(s, "EVIDENCE 6 — THE SHOWROOM", "The credit had achieved sticker parity",
  "Chevrolet Equinox, gasoline versus electric, in the same showroom. Its removal reopened a $6,500 gap.");
s.addChart(p.ChartType.bar, [{
  name:"Price to the buyer",
  labels:["Equinox EV\nwith $7,500 credit\n(to Sept 2025)","Equinox\n(gasoline)","Equinox EV\ntoday"],
  values:[28995, 29995, 36495],
}], Object.assign(axis(), {
  x:M, y:2.30, w:7.5, h:4.0, barDir:"col",
  chartColors:[TEAL, MUT, OXIDE], showLegend:false, showTitle:false,
  dataLabelFormatCode:'$#,##0', valAxisMaxVal:40000, barGapWidthPct:60,
}));
card(s, 8.42, 2.30, 4.26, 4.0, DARK);
s.addText("Under the credit, the electric version of a mainstream American crossover undercut its own gasoline sibling.",
  { x:8.70, y:2.58, w:3.72, h:1.15, isTextBox:true, margin:0, fontFace:SERIF, fontSize:17, bold:true, color:WHITE });
s.addText("That is the threshold at which a technology stops needing advocacy and starts selling on arithmetic.\n\nToday the same electric model carries a $6,500 premium over the car parked beside it, recoverable only through years of fuel savings.\n\nNo currency conversion, no income adjustment, no foreign market — this comparison is identical for every American buyer.",
  { x:8.70, y:3.86, w:3.72, h:2.2, isTextBox:true, margin:0, fontFace:SANS, fontSize:11.5, color:PALE });
src(s, "Cars.com 2026 Equinox and Equinox EV pricing, including destination");
s.addNotes("The most consumer-legible slide in the deck. No international comparison needed.");

// ===== 10. DEMAND DOWN, DISCOUNTS DOWN ==================================
s = p.addSlide(); s.background = { color: DARK };
s.addText("EVIDENCE 7 — PRICING POWER", { x:M, y:0.52, w:11, h:0.28, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:11, bold:true, charSpacing:2, color:OCHRE });
s.addText("Demand collapsed 28% — and the discounts shrank anyway", { x:M, y:0.86, w:11.6, h:0.66,
  isTextBox:true, margin:0, fontFace:SERIF, fontSize:31, bold:true, color:WHITE });
s.addText("In a competitively disciplined market, losing a quarter of your volume forces you to discount. That is what competition is.",
  { x:M, y:1.56, w:11.4, h:0.4, isTextBox:true, margin:0, fontFace:SANS, fontSize:14, color:PALE });
const grid = [
  ["−28%","New BEV sales, Q1 2026 year on year", OXIDE],
  ["−24%","EV incentive spending by manufacturers", OXIDE],
  ["$49,855","Average new-vehicle transaction price — a 2026 high", OCHRE],
  ["+2.5 pts","GM North America EBIT margin, to 8.6%; guidance raised twice", OCHRE],
];
grid.forEach((g,i)=>{
  const x = M + (i%2)*6.06, y = 2.22 + Math.floor(i/2)*1.88;
  s.addShape(p.ShapeType.roundRect, { x, y, w:5.75, h:1.66, rectRadius:0.06,
    fill:{color:"16394B"}, line:{color:"27536A", width:0.75} });
  s.addText(g[0], { x:x+0.32, y:y+0.20, w:2.7, h:0.62, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:32, bold:true, color:g[2] });
  s.addText(g[1], { x:x+0.32, y:y+0.86, w:5.1, h:0.62, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:12.5, color:PALE });
});
s.addText("Automakers absorbed the tariffs into sticker prices and let the deals shrink instead — which is exactly why a sticker comparison shows nothing. The increase is hidden in the incentive line.",
  { x:M, y:6.16, w:11.9, h:0.6, isTextBox:true, margin:0, fontFace:SANS, fontSize:13, italic:true, color:WHITE });
src(s, "Kelley Blue Book July 2026 · Cox Automotive · GM Q2 2026 earnings");
s.addNotes("Protection is being converted into margin rather than into cheaper cars.");

// ===== 11. NO CONSISTENT PREMIUM ========================================
s = p.addSlide();
head(s, "THE COUNTER-TEST", "There is no uniform \"American premium\" — and that proves the point",
  "U.S. price versus the comparison market, same model. Negative means the car is cheaper in the United States.");
s.addChart(p.ChartType.bar, [{
  name:"U.S. price vs comparison market",
  labels:["Mustang Mach-E\nvs Germany","VW ID.4\nvs Germany","Tesla Model Y\nvs Germany","Kia EV9\nvs Canada"],
  values:[-31, -5, 2, 27],
}], Object.assign(axis(), {
  x:M, y:2.30, w:7.6, h:3.95, barDir:"col",
  chartColors:[BLUE, BLUE, OCHRE, OXIDE], showLegend:false, showTitle:false,
  dataLabelFormatCode:'+0"%";-0"%"', barGapWidthPct:55,
}));
card(s, 8.52, 2.30, 4.16, 3.95);
s.addText("A policy effect spread across the price range would show a consistent direction. It does not.",
  { x:8.80, y:2.56, w:3.62, h:1.0, isTextBox:true, margin:0, fontFace:SERIF, fontSize:17, bold:true, color:INK });
s.addText([
  { text:"Above class C the U.S. is a normal, competitive market — sometimes the cheapest on the list.", options:{bullet:true, breakLine:true, paraSpaceAfter:8} },
  { text:"Where it is dearer, the cause is ordinary market positioning, not the prohibition.", options:{bullet:true, breakLine:true, paraSpaceAfter:8} },
  { text:"In classes A and B the American figure is not a percentage at all. It is a blank cell.", options:{bullet:true} },
], { x:8.80, y:3.66, w:3.62, h:2.4, isTextBox:true, margin:0, fontFace:SANS, fontSize:11.5, color:INK });
src(s, "Manufacturer list prices Sept 2026, ex-consumption-tax. EV9 comparison is base trim to base trim; standard content may differ between markets.");
s.addNotes("This is the slide that makes the argument falsifiable — and it survives.");

// ===== 12. THE BARBELL (body type) ======================================
s = p.addSlide();
head(s, "BY BODY TYPE", "The American EV market is barbell-shaped",
  "Cheapest electric option in the U.S., by body type. Two categories have no product at all — for two different reasons.");
s.addChart(p.ChartType.bar, [{
  name:"Cheapest U.S. electric option",
  labels:["Supermini","Compact SUV","Mid-size SUV","Pickup","Full-size SUV"],
  values:[28995, 36495, 39990, 52800, 54900],
}], Object.assign(axis(), {
  x:M, y:2.34, w:7.5, h:3.9, barDir:"col",
  chartColors:[BLUE, BLUE, BLUE, OCHRE, OCHRE],
  showLegend:false, showTitle:false, dataLabelFormatCode:'$#,##0', barGapWidthPct:50,
}));
const gaps = [
  ["City car","No product at any price","Supply-side: the Chinese-EV prohibition. Europe has four rivals under $20,000."],
  ["Minivan / MPV","None on sale for model year 2026","Demand-side: VW withdrew the ID.Buzz after the credit expired. China's 7-seat BYD Xia is $27,233."],
];
gaps.forEach((g,i)=>{
  const y = 2.34 + i*2.02;
  card(s, 8.42, y, 4.26, 1.86, "FBEDEB");
  s.addText(g[0], { x:8.70, y:y+0.16, w:3.7, h:0.34, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:18, bold:true, color:OXIDE });
  s.addText(g[1], { x:8.70, y:y+0.55, w:3.7, h:0.34, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:12, bold:true, color:INK });
  s.addText(g[2], { x:8.70, y:y+0.92, w:3.7, h:0.82, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:11, color:MUT });
});
s.addText("Each policy produced a missing segment by a different mechanism — the clearest sign they are two separate forces pushing the same way, not one story told twice.",
  { x:M, y:6.36, w:W-2*M, h:0.5, isTextBox:true, margin:0, fontFace:SANS, fontSize:12.5, italic:true, color:INK });
src(s, "Edmunds · iSeeCars · Cars.com · CnEVPost · manufacturer listings, Sept 2026");
s.addNotes("Strong and innovative above $50,000, absent below $29,000. Neither policy set out to build a barbell; both did.");

// ===== 13. POWERTRAIN LADDER ============================================
s = p.addSlide();
head(s, "BY POWERTRAIN", "The subsidy was removed from one powertrain and not the other",
  "What an American pays to move off gasoline, same nameplate, same dealer.");
s.addChart(p.ChartType.bar, [{
  name:"Premium over the gasoline version",
  labels:["Battery-electric\nwith the credit\n(to Sept 2025)","Hybrid\n(never subsidised)","Battery-electric\ntoday"],
  values:[-1000, 4710, 6500],
}], Object.assign(axis(), {
  x:M, y:2.34, w:7.5, h:3.9, barDir:"col",
  chartColors:[TEAL, BLUE, OXIDE], showLegend:false, showTitle:false,
  dataLabelFormatCode:'+$#,##0;-$#,##0', barGapWidthPct:58,
}));
card(s, 8.42, 2.34, 4.26, 1.72, DARK);
s.addText("Toyota no longer sells a gasoline RAV4 or Camry.", { x:8.70, y:2.54, w:3.7, h:0.72,
  isTextBox:true, margin:0, fontFace:SERIF, fontSize:17, bold:true, color:WHITE });
s.addText("For 2026 every one is a hybrid. America's default powertrain has already left pure gasoline.",
  { x:8.70, y:3.26, w:3.7, h:0.62, isTextBox:true, margin:0, fontFace:SANS, fontSize:11.5, color:PALE });
card(s, 8.42, 4.22, 4.26, 2.02);
s.addText("It just moved to hybrid, not electric", { x:8.70, y:4.38, w:3.7, h:0.44,
  isTextBox:true, margin:0, fontFace:SERIF, fontSize:16, bold:true, color:INK });
s.addText("A hybrid costs ~$4,700 over gasoline and never needed a subsidy. A BEV costs ~$6,500 and lost the $7,500 that closed the gap.\n\nMexico shows the destination: 69% of its electrified sales are hybrids, against 14.5% battery-electric.",
  { x:8.70, y:4.84, w:3.72, h:1.28, isTextBox:true, margin:0, fontFace:SANS, fontSize:10.5, color:INK });
s.addText("Chevrolet Equinox gasoline $29,995 vs Equinox EV $36,495 · Honda CR-V $32,370 vs CR-V Hybrid $37,080 · Toyota RAV4 hybrid $33,350, PHEV $42,920",
  { x:M, y:6.36, w:W-2*M, h:0.5, isTextBox:true, margin:0, fontFace:SANS, fontSize:11, color:MUT });
src(s, "Edmunds · TrueCar · Motor1 · Cars.com · AMIA Mexico, Sept 2026");
s.addNotes("Remove support from one powertrain and leave the other untouched, and the hybrid takes the middle of the market by default.");

// ===== 14. FORWARD ======================================================
s = p.addSlide();
head(s, "OUTLOOK", "What to expect through 2030",
  "Two changes since the analysis was drafted matter more than anything on a watch list.");
const fwd = [
  ["≈85%","The floor stays empty to 2030","The codifying bill is bipartisan — Moreno (R) and Slotkin (D) — and framed as security, not climate. A chamber flipping in November does not touch it.", OXIDE],
  ["≈25%","A purchase credit restored and in effect before 2029","A flipped chamber still faces a veto. The loan-interest deduction sunsets in December 2028.", OCHRE],
  ["≈20%","A U.S. EV price floor below $25,000 before 2030","Ford's $30,000 pickup arrives in 2027 on LFP cells. Platform economics reach $30k, not $22k.", OCHRE],
  ["6–10%","BEV share range through 2030","Against 22%+ in Europe. Expect divergence, not recovery.", BLUE],
];
fwd.forEach((f,i)=>{
  const y = 2.22 + i*1.13;
  card(s, M, y, W-2*M, 1.01);
  s.addText(f[0], { x:M+0.28, y:y+0.28, w:1.42, h:0.48, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:22, bold:true, color:f[3] });
  s.addText(f[1], { x:M+1.82, y:y+0.14, w:4.3, h:0.75, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:13, bold:true, color:INK, valign:"middle" });
  s.addText(f[2], { x:M+6.30, y:y+0.14, w:5.5, h:0.75, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:11, color:MUT, valign:"middle" });
});
s.addText("The underweighted change: the EPA repealed the 2009 endangerment finding in February 2026, on top of the December 2025 CAFE rollback. Automakers no longer have a compliance reason to build EVs at all — structurally larger than the lost subsidy.",
  { x:M, y:6.86, w:W-2*M, h:0.52, isTextBox:true, margin:0, fontFace:SANS, fontSize:12, italic:true, color:INK });
s.addNotes("Probabilities are directional judgements, not model outputs. The near-certainties are near-certain because they are already law.");

// ===== 15. CONCLUSION ===================================================
s = p.addSlide(); s.background = { color: DARK };
s.addText("CONCLUSION", { x:M, y:0.72, w:11, h:0.3, isTextBox:true, margin:0,
  fontFace:SANS, fontSize:11, bold:true, charSpacing:2.5, color:OCHRE });
s.addText("The policy is succeeding at its stated goal and failing at the one usually attached to it.",
  { x:M, y:1.16, w:11.6, h:1.35, isTextBox:true, margin:0, fontFace:SERIF, fontSize:33, bold:true, color:WHITE });
const cols = [
  ["Protection is working","Domestic industry shielded. GM's North American margin up 2.5 points, guidance raised twice, no Chinese market share, tariff revenue collected.", TEAL],
  ["Electrification is not","The U.S. is the only market of the four where adoption fell in 2026 — −28%, against Europe +35%, Canada +40%, Mexico +44%.", OXIDE],
  ["The cost falls on one group","Buyers above $40,000 are served normally. Buyers below $29,000 are served not at all. Protection removed the segment rather than raising its price.", OCHRE],
];
cols.forEach((c,i)=>{
  const x = M + i*4.06;
  s.addShape(p.ShapeType.roundRect, { x, y:2.86, w:3.78, h:2.72, rectRadius:0.06,
    fill:{color:"16394B"}, line:{color:"27536A", width:0.75} });
  s.addShape(p.ShapeType.ellipse, { x:x+0.30, y:3.10, w:0.20, h:0.20, fill:{color:c[2]} });
  s.addText(c[0], { x:x+0.30, y:3.42, w:3.2, h:0.62, isTextBox:true, margin:0,
    fontFace:SERIF, fontSize:18, bold:true, color:WHITE });
  s.addText(c[1], { x:x+0.30, y:4.06, w:3.2, h:1.35, isTextBox:true, margin:0,
    fontFace:SANS, fontSize:12, color:PALE });
});
s.addText("These are separable goals. The open question is whether that trade is being made deliberately — or as an unexamined side effect.",
  { x:M, y:5.92, w:11.6, h:0.5, isTextBox:true, margin:0, fontFace:SANS, fontSize:15, color:WHITE });
s.addText("PESTEL · U.S. Electric Vehicle Market · Factor 1 of 6, Political · Baseline September 2026",
  { x:M, y:6.72, w:11.6, h:0.3, isTextBox:true, margin:0, fontFace:SANS, fontSize:10, color:MUT });

p.writeFile({ fileName: "/home/user/PESTEL/slides/political-findings.pptx" })
 .then(f => console.log("WROTE", f));
