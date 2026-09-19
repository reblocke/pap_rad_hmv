// PAP, RADs, and HMVs — 45-minute fellows' talk
// Minimalist black-and-white draft deck.
const pptxgen = require("pptxgenjs");
const path = require("node:path");
const pres = new pptxgen();

pres.layout = "LAYOUT_WIDE";            // 13.333 x 7.5
pres.author = "Brian Locke";
pres.title = "PAP, RADs, and Home Mechanical Ventilators";

// ---------- design tokens ----------
const INK   = "000000";   // titles, rules
const TEXT  = "1A1A1A";   // body
const MUTE  = "6E6E6E";   // captions, secondary
const RULE  = "C8C8C8";   // hairlines, table borders
const WASH  = "F2F2F2";   // table header / callout fill
const WHITE = "FFFFFF";

const HEAD = "Cambria";   // titles
const BODY = "Calibri";   // everything else

const W = 13.333, H = 7.5;
const M = 0.7;                 // side margin
const CW = W - 2 * M;          // content width = 11.933
const TOP = 1.32;              // first content baseline
const TOPS = 1.58;             // first content baseline when the slide has a subtitle
const FOOT = 6.93;

// horizontal-rule-only table border (no verticals) — reads clean in B&W
const HRULE = [
  { type: "solid", color: RULE, pt: 0.75 },
  { type: "none" },
  { type: "solid", color: RULE, pt: 0.75 },
  { type: "none" },
];

let N = 0;
function newSlide(section) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  if (section !== null) {
    N += 1;
    s.addText(section.toUpperCase(), {
      x: M, y: FOOT, w: CW - 0.6, h: 0.26, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 9, color: MUTE, charSpacing: 1.1,
    });
    s.addText(String(N), {
      x: W - M - 0.6, y: FOOT, w: 0.6, h: 0.26, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 9, color: MUTE, align: "right",
    });
  }
  return s;
}

function title(s, text, sub) {
  s.addText(text, {
    x: M, y: 0.42, w: CW, h: 0.62, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 27, bold: true, color: INK, valign: "top",
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.02, w: CW, h: 0.3, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 12.5, italic: true, color: MUTE,
    });
  }
}

function body(s, runs, opts) {
  s.addText(runs, Object.assign({
    isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13,
    color: TEXT, valign: "top", lineSpacingMultiple: 1.15,
  }, opts));
}

function bullets(s, items, opts) {
  const runs = items.map((t, i) => {
    const o = { text: typeof t === "string" ? t : t.text, options: { bullet: { indent: 13 }, breakLine: i < items.length - 1, paraSpaceAfter: 7 } };
    if (typeof t !== "string" && t.bold) o.options.bold = true;
    return o;
  });
  body(s, runs, opts);
}

function table(s, rows, opts) {
  s.addTable(rows, Object.assign({
    x: M, w: CW, border: HRULE, fontFace: BODY, fontSize: 11.5,
    color: TEXT, valign: "top", margin: [6, 14, 6, 0], autoPage: false,
  }, opts));
}

function hdr(t) { return { text: t, options: { bold: true, fill: { color: WASH }, color: INK } }; }
function b(t)   { return { text: t, options: { bold: true } }; }

// pull-quote block: hairline above and below, no fill, no color bar
function quote(s, text, attrib, o) {
  const x = o.x, y = o.y, w = o.w;
  const h = o.h || 0.95;
  s.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color: INK, width: 1 } });
  s.addText(text, {
    x, y: y + 0.13, w, h, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: o.fontSize || 14, italic: true, color: INK,
    valign: "top", lineSpacingMultiple: 1.12,
  });
  if (attrib) {
    s.addText(attrib, {
      x, y: y + 0.13 + h, w, h: 0.24, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 10, color: MUTE,
    });
  }
  s.addShape(pres.ShapeType.line, {
    x, y: y + 0.13 + h + (attrib ? 0.28 : 0.06), w, h: 0,
    line: { color: RULE, width: 0.75 },
  });
}

function caption(s, text, o) {
  s.addText(text, Object.assign({
    isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10,
    color: MUTE, italic: true, lineSpacingMultiple: 1.1,
  }, o));
}

// connected orthogonal waveform from [[x,y],...] in inches
function wave(s, pts, color, width) {
  for (let i = 0; i < pts.length - 1; i++) {
    const [x1, y1] = pts[i], [x2, y2] = pts[i + 1];
    if (Math.abs(y1 - y2) < 0.001) {
      s.addShape(pres.ShapeType.line, { x: Math.min(x1, x2), y: y1, w: Math.abs(x2 - x1), h: 0, line: { color, width } });
    } else {
      s.addShape(pres.ShapeType.line, { x: x1, y: Math.min(y1, y2), w: 0, h: Math.abs(y2 - y1), line: { color, width } });
    }
  }
}

function arrow(s, x, y, w, opts) {
  s.addShape(pres.ShapeType.line, Object.assign({
    x, y, w, h: 0, line: { color: INK, width: 1, endArrowType: "triangle" },
  }, opts || {}));
}

function box(s, text, o) {
  s.addShape(pres.ShapeType.rect, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: { color: o.fill || WHITE }, line: { color: o.line || INK, width: o.lw || 1 },
  });
  s.addText(text, {
    x: o.x + 0.08, y: o.y, w: o.w - 0.16, h: o.h, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: o.fontSize || 11, bold: o.bold || false,
    color: o.color || TEXT, align: "center", valign: "middle", lineSpacingMultiple: 1.05,
  });
}

function stat(s, big, label, o) {
  s.addText(big, {
    x: o.x, y: o.y, w: o.w, h: 0.78, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: o.size || 40, bold: true, color: INK, align: "center",
  });
  s.addText(label, {
    x: o.x, y: o.y + 0.8, w: o.w, h: 0.6, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 11, color: MUTE, align: "center", lineSpacingMultiple: 1.1,
  });
}

// =====================================================================
// TITLE
// =====================================================================
{
  const s = newSlide(null);
  s.background = { color: WHITE };
  s.addText("PAP, RADs, and Home Mechanical Ventilators", {
    x: M, y: 2.25, w: CW, h: 0.9, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 40, bold: true, color: INK,
  });
  s.addShape(pres.ShapeType.line, { x: M, y: 3.28, w: 4.2, h: 0, line: { color: INK, width: 1.25 } });
  s.addText("How an epidemic, a blower, and a sentence in a 1993 tax bill\nproduced three device categories", {
    x: M, y: 3.5, w: CW, h: 0.8, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 17, italic: true, color: TEXT, lineSpacingMultiple: 1.2,
  });
  s.addText("Pulmonary & Critical Care Fellows Conference  ·  45 minutes", {
    x: M, y: 5.5, w: CW, h: 0.3, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, color: TEXT,
  });
  caption(s, "Medicare policy verified September 2026  ·  DRAFT", { x: M, y: 5.85, w: CW, h: 0.3 });
  s.addNotes("45 min: 40 content + 5 discussion. 22 content slides. Sections: I clinical problem, II polio lineage, III sleep lineage, IV the regulatory fork, V convergence, VI current categories, VII close.");
}

// =====================================================================
// I. THE CLINICAL PROBLEM  (slides 1-2, 0-3 min)
// =====================================================================
{
  const s = newSlide("I · The clinical problem");
  title(s, "Three patients need positive pressure.", "They do not need the same thing.");
  const T = TOPS;
  const cw = 3.72, gap = 0.39;
  const cases = [
    ["A", "54M, symptomatic OSA\nAHI 38, normal awake ABG,\npreserved muscle strength", "Sleepy. Hypertensive.\nNothing acute."],
    ["B", "68F, COPD\nPaCO₂ 58 mmHg two weeks\nafter an exacerbation, on 2 L", "Progressive CO₂ rise.\nReadmission."],
    ["C", "61M, ALS\nFVC 42%, now needs support\nduring meals and conversation", "An unrecognized interruption\nmay be fatal."],
  ];
  cases.forEach((c, i) => {
    const x = M + i * (cw + gap);
    s.addShape(pres.ShapeType.line, { x, y: T, w: cw, h: 0, line: { color: INK, width: 1.25 } });
    s.addText(c[0], { x, y: T + 0.1, w: cw, h: 0.38, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: INK });
    s.addText(c[1], { x, y: T + 0.58, w: cw, h: 1.1, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: TEXT, lineSpacingMultiple: 1.2 });
    s.addShape(pres.ShapeType.line, { x, y: T + 1.88, w: cw, h: 0, line: { color: RULE, width: 0.75 } });
    caption(s, "A night without the device", { x, y: T + 1.98, w: cw, h: 0.24 });
    s.addText(c[2], { x, y: T + 2.28, w: cw, h: 0.75, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: TEXT, lineSpacingMultiple: 1.2 });
  });
  quote(s, "What must the system do for each patient?  And what happens if it stops?", null,
    { x: M, y: 5.45, w: CW, h: 0.4, fontSize: 15 });
  s.addNotes("Do NOT name devices yet. Return to these three in the final slide. The point of the third column is that consequence-of-interruption, not disease severity, is what separates the categories.");
}

{
  const s = newSlide("I · The clinical problem");
  title(s, "Four decisions, routinely collapsed into one");
  const bw = 2.62, bh = 0.72, y0 = TOP + 0.12;
  const steps = ["Diagnosis", "Treatment mode", "Platform capabilities", "Coverage pathway"];
  steps.forEach((t, i) => {
    const x = M + i * (bw + 0.48);
    box(s, t, { x, y: y0, w: bw, h: bh, bold: true, fontSize: 12.5 });
    if (i < 3) arrow(s, x + bw + 0.08, y0 + bh / 2, 0.32);
  });
  s.addText("The correction that has to come early", {
    x: M, y: y0 + 1.32, w: CW, h: 0.3, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: INK,
  });
  body(s, [
    { text: "PAP, RAD, and HMV are not mutually exclusive categories.", options: { bold: true, breakLine: true } },
    { text: "Medicare's OSA policy (LCD L33718) defines “PAP device” to mean both E0601 CPAP and E0470 bilevel-without-backup. The same E0470 code follows an entirely different coverage pathway when the diagnosis is not OSA.", options: {} },
  ], { x: M, y: y0 + 1.62, w: CW, h: 0.8 });
  s.addShape(pres.ShapeType.line, { x: M, y: 4.42, w: CW, h: 0, line: { color: RULE, width: 0.75 } });
  s.addText("Objectives", {
    x: M, y: 4.56, w: CW, h: 0.28, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: INK,
  });
  bullets(s, [
    "Explain why two engineering lineages converged in capability but not in category",
    "Distinguish mode from platform from billing code",
    "Select and justify a home support system from clinical need and consequence of interruption — not device prestige",
  ], { x: M + 0.05, y: 4.86, w: CW - 0.05, h: 1.4, fontSize: 12.5 });
  s.addNotes("The E0470 example is the whole talk in miniature: one code, two coverage universes, depending only on diagnosis.");
}

// =====================================================================
// II. LINEAGE ONE — POLIO  (slides 3-6, 3-13 min)
// =====================================================================
{
  const s = newSlide("II · Lineage one — polio");
  title(s, "Before you can treat hypoventilation, you must be able to see it");
  const ty = TOP + 0.5;
  s.addShape(pres.ShapeType.line, { x: M, y: ty, w: CW, h: 0, line: { color: INK, width: 1.25 } });
  const marks = [
    ["1832", "First negative-pressure\n“box” (Dalziel, Scotland)"],
    ["1907", "Dräger's portable device —\nbuilt for mine rescue,\nnot for clinics"],
    ["1928–31", "Drinker & Shaw's iron lung;\nprolonged support\nreported in Lancet"],
    ["1950s", "Electrode pH measurement\nlets Ibsen attribute polio\nmortality to hypercapnia"],
  ];
  const cw2 = 2.72;
  marks.forEach((m, i) => {
    const x = M + i * (cw2 + 0.35);
    s.addShape(pres.ShapeType.line, { x: x + 0.02, y: ty, w: 0, h: 0.2, line: { color: INK, width: 1.25 } });
    s.addText(m[0], { x, y: ty + 0.24, w: cw2, h: 0.34, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: INK });
    s.addText(m[1], { x, y: ty + 0.62, w: cw2, h: 1.1, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: TEXT, lineSpacingMultiple: 1.15 });
  });
  quote(s, "The problem became treatable when it became measurable.", null, { x: M, y: 4.62, w: 7.6, h: 0.34, fontSize: 15 });
  caption(s, "Hold that thought. On slide 18 a PaCO₂ number becomes a coverage threshold.", { x: M, y: 5.42, w: 7.6, h: 0.3 });
  s.addShape(pres.ShapeType.rect, { x: 8.55, y: 4.5, w: 4.08, h: 1.35, fill: { color: WHITE }, line: { color: RULE, width: 0.75, dashType: "dash" } });
  caption(s, "FIGURE — one iron-lung photograph: ward of tanks,\nheads protruding, lids lifted for nursing care.\n[source with clear licensing still needed]", { x: 8.73, y: 4.66, w: 3.72, h: 1.1 });
  s.addNotes("Positive pressure was in industrial use (Draeger, gas poisoning and mining) two decades before the negative-pressure clinical era. Worth one sentence: the technology direction of travel is not what people assume.");
}

{
  const s = newSlide("II · Lineage one — polio");
  title(s, "1949 and 1952: two inflection points, not one");
  const cw3 = 5.78, gap3 = 0.37;
  const items = [
    ["1948–49  ·  Los Angeles",
     "Albert Bower (physician) and Ray Bennett (engineer) devise a valve converting tank respirators to deliver intermittent positive pressure during inspiration.",
     "Survival 21% → 84% from 1949."],
    ["1952  ·  Copenhagen",
     "Not enough iron lungs — and the tanks could not ventilate bulbar polio. Tracheotomy and manual positive-pressure ventilation, delivered by medical students in shifts.",
     "Death rate ~80%, falling markedly."],
  ];
  items.forEach((it, i) => {
    const x = M + i * (cw3 + gap3);
    s.addShape(pres.ShapeType.line, { x, y: TOP, w: cw3, h: 0, line: { color: INK, width: 1.25 } });
    s.addText(it[0], { x, y: TOP + 0.1, w: cw3, h: 0.34, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 15, bold: true, color: INK });
    s.addText(it[1], { x, y: TOP + 0.52, w: cw3, h: 1.15, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: TEXT, lineSpacingMultiple: 1.18 });
    s.addText(it[2], { x, y: TOP + 1.72, w: cw3, h: 0.34, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, bold: true, color: INK });
  });
  caption(s, "The 1949 Bower–Bennett work predates Copenhagen and is almost always omitted.   ·   80% figure: Hind 2017; D'Cruz & Hart report “substantial reductions” without a number.",
    { x: M, y: TOP + 2.18, w: CW, h: 0.4 });
  quote(s,
    "Once a patient survives respiratory paralysis, who supplies the next breath — tomorrow, next month, and at home?",
    "This question drives the rest of the talk.",
    { x: M, y: 4.62, w: 6.1, h: 0.72, fontSize: 15 });
  // four-box care system diagram
  const bx = 7.15, bw4 = 1.25, by = 4.72;
  ["Ventilation", "Monitoring", "Trained\nstaff", "Ongoing\nsupport"].forEach((t, i) => {
    box(s, t, { x: bx + i * (bw4 + 0.13), y: by, w: bw4, h: 0.78, fontSize: 9.5, line: i === 0 ? INK : RULE });
  });
  caption(s, "The machine is one box of four.", { x: bx, y: by + 0.9, w: 5.4, h: 0.3 });
  s.addNotes("Ibsen 1954; Lassen Lancet 1953; Trubuhovich Crit Care Resusc 2007 for Bower/Bennett. The four-box diagram is the setup for the entire regulatory story: Medicare ends up paying for box one and, through FSS, partly for boxes two through four.");
}

{
  const s = newSlide("II · Lineage one — polio");
  title(s, "Survival created a care problem that philanthropy solved first");
  bullets(s, [
    "1946 — the March of Dimes funds and transports thousands of iron lungs",
    "Late 1950s — the MOD is operating 15 polio respiratory centers",
    "Rancho Los Amigos, Downey CA — the first system of ventilatory home care; polio survivors discharged home with iron lungs and portable chest cuirasses. These were among the earliest domiciliary NIV patients",
    "Salk's vaccine ends the epidemics. By the mid-1960s the MOD closes its respiratory centers for lack of funding",
  ], { x: M + 0.05, y: TOP, w: 7.5, h: 2.5, fontSize: 13 });
  quote(s, "Up until this time, patients with chronic respiratory failure requiring ventilator support were thought to be destined for lifelong hospitalization.",
    "Gusman & Wolfe, Sleep Med Clin 2025",
    { x: 8.45, y: TOP, w: 4.18, h: 1.5, fontSize: 13 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 4.5, w: CW, h: 0.92, fill: { color: WASH }, line: { color: WASH } });
  s.addText([
    { text: "Home ventilation required three things, and only one of them was a machine:  ", options: {} },
    { text: "equipment, a care system, and financing.", options: { bold: true } },
  ], { x: M + 0.22, y: 4.5, w: CW - 0.44, h: 0.92, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK, valign: "middle" });
  caption(s, "Attribution: Gusman & Wolfe source the MOD and Rancho Los Amigos claims to marchofdimes.org and an independentliving.org history page — a secondary account, not an archival record.",
    { x: M, y: 5.58, w: CW, h: 0.4 });
  s.addNotes("Keep this to about 90 seconds. The teaching point is the three-part requirement, not the philanthropy trivia.");
}

{
  const s = newSlide("II · Lineage one — polio");
  title(s, "The handoff — with an expiration date that never came");
  quote(s,
    "The NCDs for HMV were purposely left vague because the assumption was that these criteria would only be relevant to patients with chronic respiratory failure secondary to poliomyelitis. Since the vaccine was effectively eradicating polio, it was thought that these criteria would not be in use for long.",
    "Gusman & Wolfe, Sleep Med Clin 2025  ·  uncited in the original; expert recollection, not archival record",
    { x: M, y: TOP, w: CW, h: 1.02, fontSize: 14 });
  const rows = [
    [hdr("1965"), hdr("1960s–70s"), hdr("1977"), hdr("1981–82")],
    [
      "Medicare established. Chronic respiratory failure is left out — the polio community was told the March of Dimes would keep covering it.",
      "Medicare writes NCDs for home mechanical ventilation. They include in-home RTs, consumable supplies, and backup equipment. No LCDs are written.",
      "The LP3 portable volume ventilator is cleared for use outside the hospital (510(k) K770132). First target population: ventilator-dependent children.",
      "Katie Beckett. TEFRA 1982 §134 creates the state-plan option letting states disregard parental income for children who would otherwise need institutional care.",
    ],
  ];
  table(s, rows, { y: 3.28, colW: [2.98, 2.98, 2.98, 2.99], fontSize: 11 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 5.72, w: CW, h: 0.72, fill: { color: WASH }, line: { color: WASH } });
  s.addText("The ventilator benefit was designed generously, for a population everyone expected to disappear.", {
    x: M + 0.22, y: 5.72, w: CW - 0.44, h: 0.72, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, bold: true, color: INK, valign: "middle",
  });
  s.addNotes("Katie Beckett is Medicaid and is about access to home care — not the later RAD billing story. Keep it to 45 seconds. Distinguish the 1981 individual intervention from the 1982 statutory option. Goldberg/CHOP detail is King 2012 personal communication — mention only if asked.");
}

// =====================================================================
// III. LINEAGE TWO — THE BLOWER  (slides 7-9, 13-19 min)
// =====================================================================
{
  const s = newSlide("III · Lineage two — the blower");
  title(s, "1981: nasal CPAP is a splint, not a ventilator", "Sullivan, Issa, Berthon-Jones & Eves. Lancet 1981;1:862–865");
  const rows = [
    [hdr(""), hdr("Ventilator lineage"), hdr("Sleep-apnea device lineage")],
    [b("Original problem"), "Respiratory paralysis — the pump fails", "Intermittent upper-airway obstruction"],
    [b("Engineering answer"), "Negative pressure → positive pressure → portable home ventilators", "A blower, a mask, one pressure"],
    [b("Central concern"), "Reliability; continuity of support", "Effective airway treatment through a tolerable interface"],
    [b("Failure mode"), "Death", "Sleepiness"],
  ];
  table(s, rows, { y: TOPS, colW: [2.35, 4.75, 4.83], fontSize: 12 });
  // flat CPAP trace
  const gy = 4.95, gx = M, gw = 5.6;
  caption(s, "Airway pressure", { x: gx, y: gy - 0.3, w: 2.2, h: 0.25 });
  s.addShape(pres.ShapeType.line, { x: gx, y: gy + 0.95, w: gw, h: 0, line: { color: RULE, width: 0.75 } });
  s.addShape(pres.ShapeType.line, { x: gx, y: gy, w: 0, h: 0.95, line: { color: RULE, width: 0.75 } });
  wave(s, [[gx + 0.1, gy + 0.72], [gx + 0.5, gy + 0.72], [gx + 0.5, gy + 0.32], [gx + gw - 0.15, gy + 0.32]], INK, 1.5);
  caption(s, "One pressure, held. No inspiratory assistance. The airway is splinted, not the pump.", { x: gx, y: gy + 1.04, w: gw + 1.2, h: 0.3 });
  s.addShape(pres.ShapeType.rect, { x: 7.9, y: gy + 0.05, w: 4.73, h: 0.9, fill: { color: WHITE }, line: { color: RULE, width: 0.75, dashType: "dash" } });
  caption(s, "FIGURE — pair with an upper-airway schematic:\ncollapsed versus pneumatically splinted pharynx.", { x: 8.08, y: gy + 0.24, w: 4.4, h: 0.6 });
  s.addNotes("The conceptual advance is pneumatic splinting of the upper airway through a nasal interface — not replacement of ventilatory work. Contrast directly with the tank on slide 4.");
}

{
  const s = newSlide("III · Lineage two — the blower");
  title(s, "1990: bilevel is the bridge — and it descends from the blower", "Sanders & Kern. Chest 1990;98:317–324");
  // ---- TWO-LINEAGE CONVERGENCE FIGURE ----
  const yA = 2.52, yB = 4.22, xL = M + 0.05, xR = 9.35, xC = 11.15, yC = (yA + yB) / 2;
  s.addShape(pres.ShapeType.line, { x: xL, y: yA, w: xR - xL, h: 0, line: { color: INK, width: 1.5 } });
  s.addShape(pres.ShapeType.line, { x: xL, y: yB, w: xR - xL, h: 0, line: { color: INK, width: 1.5 } });
  // converging limbs
  s.addShape(pres.ShapeType.line, { x: xR, y: yA, w: xC - xR, h: yC - yA, line: { color: INK, width: 1.5 } });
  s.addShape(pres.ShapeType.line, { x: xR, y: yC, w: xC - xR, h: yB - yC, line: { color: INK, width: 1.5 }, flipV: true });
  s.addText("VENTILATOR LINEAGE", { x: xL, y: yA - 0.36, w: 4, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10, bold: true, color: INK, charSpacing: 1 });
  s.addText("SLEEP-APNEA DEVICE LINEAGE", { x: xL, y: yB + 0.56, w: 4.4, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10, bold: true, color: INK, charSpacing: 1 });
  const topMarks = [[1928, "Iron lung"], [1952, "Copenhagen"], [1977, "LP3 portable"], [1990, "Home ventilators"]];
  const botMarks = [[1981, "Nasal CPAP"], [1990, "Bilevel"], [1998, "S/T, VAPS"], [2010, "Leak compensation"]];
  topMarks.forEach((m, i) => {
    const x = xL + 0.35 + i * 2.15;
    s.addShape(pres.ShapeType.line, { x, y: yA - 0.16, w: 0, h: 0.16, line: { color: INK, width: 1.25 } });
    s.addText(String(m[0]), { x: x - 0.9, y: yA - 0.72, w: 1.8, h: 0.24, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10.5, bold: true, color: INK, align: "center" });
    s.addText(m[1], { x: x - 0.98, y: yA + 0.06, w: 1.96, h: 0.26, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10, color: MUTE, align: "center" });
  });
  botMarks.forEach((m, i) => {
    const x = xL + 0.35 + i * 2.15;
    s.addShape(pres.ShapeType.line, { x, y: yB, w: 0, h: 0.16, line: { color: INK, width: 1.25 } });
    s.addText(String(m[0]), { x: x - 0.9, y: yB - 0.3, w: 1.8, h: 0.24, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10.5, bold: true, color: INK, align: "center" });
    s.addText(m[1], { x: x - 0.98, y: yB + 0.2, w: 1.96, h: 0.26, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10, color: MUTE, align: "center" });
  });
  s.addText("Capability\nconverges.\nCategory\ndoes not.", {
    x: xC + 0.06, y: yC - 0.62, w: 1.5, h: 1.24, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 12, bold: true, color: INK, valign: "middle", lineSpacingMultiple: 1.1,
  });
  quote(s,
    "Cross-pollination from the development of continuous CPAP blowers and masks to treat obstructive sleep apnea led to the increasing use of smaller bilevel pressure support devices.",
    "Hind, Polkey & Simonds, AJRCCM 2017",
    { x: M, y: 5.28, w: 6.2, h: 0.72, fontSize: 12.5 });
  quote(s,
    "RADs evolved from CPAP machines for treating OSA. Home mechanical ventilators evolved from hospital ventilators and are designed to function as life-support devices.",
    "Gusman & Wolfe, Sleep Med Clin 2025",
    { x: 7.15, y: 5.28, w: 5.48, h: 0.72, fontSize: 12.5 });
  s.addNotes("BUILD THIS FIGURE FIRST. Two horizontal tracks converging but never merging. Do NOT draw a single ladder from CPAP to RAD to HMV — that is the mental model the talk exists to break.");
}

{
  const s = newSlide("III · Lineage two — the blower");
  title(s, "The population explodes — and the category cannot see the difference");
  quote(s,
    "There is no way to distinguish the number of E0470 patients who need ventilator support from those obstructive sleep apnea patients who were non-adherent to CPAP and were therefore moved to a bi-level device for comfort.",
    "King, Respir Care 2012",
    { x: M, y: TOP, w: CW, h: 0.72, fontSize: 14 });
  caption(s, "Medicare claims, 2010", { x: M, y: 3.05, w: 4, h: 0.28 });
  const sw = 2.82, sgap = 0.22;
  [["36,117", "E0470\nbilevel, no backup rate"],
   ["7,793", "E0471\nbilevel with backup rate"],
   ["4,071", "all true home\nventilators"],
   ["22", "E0460 negative pressure\n(chest cuirass, Porta-Lung)"]].forEach((d, i) => {
    stat(s, d[0], d[1], { x: M + i * (sw + sgap), y: 3.38, w: sw, size: i === 3 ? 40 : 36 });
  });
  s.addShape(pres.ShapeType.line, { x: M, y: 5.62, w: CW, h: 0, line: { color: RULE, width: 0.75 } });
  s.addText("Twenty-two chest cuirasses still on the fee schedule in 2010. The iron lung never quite left the statute.", {
    x: M, y: 5.76, w: CW, h: 0.34, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 14, italic: true, color: INK,
  });
  s.addNotes("BPAP use for severe COPD and neuromuscular disease became widespread through the 1990s (ONMAP). These smaller devices opened home ventilatory support to OSA, COPD, CSA and hypoventilation — four populations the single E0470 code cannot separate.");
}

// =====================================================================
// IV. THE REGULATORY FORK  (slides 10-13, 19-28 min)
// =====================================================================
{
  const s = newSlide("IV · The regulatory fork");
  title(s, "1992–93: a manufacturer's sentence, then an act of Congress");
  caption(s, "January 1, 1992  ·  HCPCS E0452 is created to describe bilevel devices. The manufacturer's own submission is the hinge of the whole story.", { x: M, y: TOP, w: CW, h: 0.3 });
  quote(s,
    "The word “intermittent” refers to devices that are designed to be used by the patient for only part of the day, usually during the hours of sleep.   …   The bi-level equipment requires very little maintenance and servicing.",
    "Manufacturer submission, quoted at 71 FR 4519–4520",
    { x: M, y: TOP + 0.34, w: CW, h: 0.72, fontSize: 13.5 });
  s.addText("1993  ·  OBRA §13543", {
    x: M, y: 3.32, w: CW, h: 0.3, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: INK,
  });
  body(s, "Amends SSA §1834(a)(3)(A) to exclude “intermittent assist devices with continuous positive airway pressure devices” — by name — from the frequent-and-substantial-servicing category. Every bilevel device is swept out of the ventilator payment category by an act of Congress.",
    { x: M, y: 3.62, w: 7.55, h: 1.0, fontSize: 13 });
  // the money
  s.addShape(pres.ShapeType.rect, { x: 8.5, y: 3.3, w: 4.13, h: 1.72, fill: { color: WASH }, line: { color: WASH } });
  s.addText("One device, one patient, five years", { x: 8.72, y: 3.44, w: 3.7, h: 0.26, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 10, color: MUTE });
  s.addText([{ text: "$38,530", options: { bold: true, breakLine: true } }, { text: "under frequent & substantial servicing", options: { fontSize: 10, color: MUTE } }],
    { x: 8.72, y: 3.74, w: 3.7, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 21, color: INK });
  s.addText([{ text: "$12,201", options: { bold: true, breakLine: true } }, { text: "under capped rental", options: { fontSize: 10, color: MUTE } }],
    { x: 8.72, y: 4.36, w: 3.7, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 21, color: INK });
  quote(s,
    "This difference in costs highlights the fact that the correct classification of these devices for Medicare payment purposes is a significant issue in terms of safeguarding the Medicare Trust Fund.",
    "71 FR 4520",
    { x: M, y: 4.82, w: 7.55, h: 0.72, fontSize: 12.5 });
  s.addText("The split was fiscal. It was not clinical — and it was not about sleep apnea by name.", {
    x: M, y: 6.28, w: CW, h: 0.34, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 13.5, bold: true, color: INK,
  });
  s.addNotes("FSS was 'intended to include items which require frequent servicing in order to avoid imminent danger to a beneficiary's health' (H.R. Conf. Rep. 103-213). E0453, 'therapeutic ventilator; suitable for use 12 hours or less per day,' was created the same day.");
}

{
  const s = newSlide("IV · The regulatory fork");
  title(s, "1998: where the volume fear actually shows up");
  quote(s,
    "These medical directors would become alarmed if outlays suddenly grew exponentially even when new technologies arrived in the marketplace. In response to the meteoric rise in BPAP prescriptions and subsequent charges, the DME MAC medical directors implemented a regional medical review policy…",
    "Gay & Owens, ONMAP. Chest 2021",
    { x: M, y: TOP, w: CW, h: 0.9, fontSize: 13.5 });
  const rows = [
    [hdr("(a)  Coverage criteria"), hdr("(b)  Payment category")],
    [
      "ACCP/NAMDRC consensus conference, 1998 → published Chest 1999;116:521–534. DMERC medical review policies implemented October 1, 1999.\n\nThe policy renamed the category “respiratory assist device, bi-level pressure capability” — coining the term fellows still trip over — and split backup-rate from non-backup-rate codes for the first time.",
      "Deliberately deferred. CMS “delayed our decision regarding the appropriate DME payment category for devices with the backup rate.”\n\nThat delay ran another six and a half years, to April 2006.",
    ],
  ];
  table(s, rows, { y: 3.42, colW: [5.97, 5.96], fontSize: 11.5 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 5.5, w: CW, h: 0.62, fill: { color: WASH }, line: { color: WASH } });
  s.addText("This — not 1993, and not 2006 — is the RAD designation's actual birthday.", {
    x: M + 0.22, y: 5.5, w: CW - 0.44, h: 0.62, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 13.5, bold: true, color: INK, valign: "middle",
  });
  caption(s, "Code chronology:  E0452 / E0453 (1992)  →  K0532 / K0533 / K0534 (Oct 1, 1999)  →  E0470 / E0471 / E0472 (Jan 1, 2004)", { x: M, y: 6.26, w: CW, h: 0.3 });
  s.addNotes("Two separate things happen in the same window and fellows conflate them: the coverage criteria (a) and the payment category (b). The criteria came first and the payment fight dragged on for years.");
}

{
  const s = newSlide("IV · The regulatory fork");
  title(s, "2006: CMS is asked the clinical question and declines to answer it");
  const rows = [
    [hdr("Comment"), hdr("CMS response")],
    ["The FDA classifies these as ventilators; their purpose and function require monitoring and servicing to avoid risk to the patient's health. Medicare payment policy is the only place they are called “respiratory assist devices.”",
     "They “are indeed referred to as ventilators in the statute, but are nonetheless excluded from the FSS category… Therefore, the law requires this change.” FDA classification “does not determine Medicare coverage and payment rules.”"],
    ["The rise in ventilator expenditures reflects the benefits of NPPV becoming known after 1995 — not miscoding.",
     { text: "“The reasons for the growth in expenditures for respiratory assist devices are not relevant to this final rule.”", options: { bold: true } }],
  ];
  table(s, rows, { y: TOP, colW: [5.97, 5.96], fontSize: 11.5 });
  caption(s, "Backed by OIG OEI-07-99-00440 (June 2001): supplier services “consist primarily of routine maintenance and patient monitoring,” visits fall short of suppliers' own protocols and decline over time. Effective April 1, 2006.",
    { x: M, y: 3.42, w: CW, h: 0.4 });
  const rows2 = [
    [hdr(""), hdr("RAD  ·  E0470 / E0471 / E0472"), hdr("Home ventilator  ·  E0465–E0468")],
    [b("Payment category"), "Capped rental", "Frequent & substantial servicing"],
    [b("Duration"), "13 months to purchase option (or 15 + semiannual servicing)", "Monthly rental indefinitely while medically necessary. No cap, no title transfer"],
    [b("Bundling"), "Device; supplies separately eligible", "All-inclusive — equipment, supplies, maintenance, repairs, replacement"],
    [b("What follows the patient home"), "“Patients on RADs typically do not receive home follow-up visits from a respiratory therapist”", "“Visited frequently for the first few months post-discharge, then at least every 30–90 days”"],
  ];
  table(s, rows2, { y: 3.92, colW: [2.35, 4.75, 4.83], fontSize: 11 });
  s.addNotes("ONMAP's nuance, say it out loud: FSS is an EQUIPMENT-servicing payment category, not a guaranteed schedule of professional RT visits. CMS 'will correctly argue that the statute and regulations apply to frequent and substantial servicing for equipment, not professional services.' The RT support is a culture that grew up around the payment, not a benefit the payment defines.");
}

{
  const s = newSlide("IV · The regulatory fork");
  title(s, "The loop closes: rationing the cheap device created the expensive one");
  // ---- CLOSED LOOP FIGURE ----
  const ly = TOP + 0.18, lh = 0.92, lbw = 2.75, lgap = 0.31;
  const nodes = [
    "1998 RAD criteria\nwritten restrictively",
    "RAD becomes harder\nto qualify for than\na ventilator",
    "Post-discharge COPD\nmigrates into the\nuncapped category",
    "85× claims growth\n2009 → 2015",
  ];
  nodes.forEach((t, i) => {
    const x = M + i * (lbw + lgap);
    box(s, t, { x, y: ly, w: lbw, h: lh, fontSize: 11 });
    if (i < 3) arrow(s, x + lbw + 0.04, ly + lh / 2, 0.26);
  });
  // return arrow underneath — drop from the centre of box 4, run left, arrow up into box 1
  const rx0 = M + lbw / 2, rx1 = M + 3 * (lbw + lgap) + lbw / 2;
  const ryB = ly + lh + 0.58;
  s.addShape(pres.ShapeType.line, { x: rx1, y: ly + lh + 0.06, w: 0, h: 0.52, line: { color: INK, width: 1 } });
  s.addShape(pres.ShapeType.line, { x: rx0, y: ryB, w: rx1 - rx0, h: 0, line: { color: INK, width: 1 } });
  s.addShape(pres.ShapeType.line, { x: rx0, y: ly + lh + 0.06, w: 0, h: 0.52, line: { color: INK, width: 1, beginArrowType: "triangle" } });
  caption(s, "NCD 240.9 (2025) is the answer to a problem the 1998 rationing created", { x: rx0, y: ryB + 0.08, w: rx1 - rx0, h: 0.3, align: "center" });
  const rows = [
    [hdr("E0464 / E0466"), hdr("2009"), hdr("2015"), hdr("Growth")],
    ["Claims", "2,528", "215,379", b("85×")],
    ["Spending", "$3.8 M", "$340 M", b("89×")],
    ["Beneficiaries", "415", "32,848", b("79×")],
  ];
  table(s, rows, { x: M, y: 4.05, w: 6.1, colW: [2.2, 1.3, 1.3, 1.3], fontSize: 11 });
  caption(s, "OIG OEI-12-15-00370, September 2016.   Ventilator ≈ $1,327–$1,561/month, uncapped.   RAD ≈ $209–$614/month, capped at 13–15 months.",
    { x: M, y: 5.95, w: 6.1, h: 0.5 });
  quote(s,
    "Since the RAD criteria are paradoxically more restrictive and challenging to fulfill than those for HMV, this has led to a trend toward providers inappropriately prescribing HMV.",
    "Gusman & Wolfe 2025",
    { x: 7.15, y: 4.05, w: 5.48, h: 0.72, fontSize: 12.5 });
  s.addText("Two competing risks, not one villain", {
    x: 7.15, y: 5.28, w: 5.48, h: 0.28, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: INK,
  });
  body(s, "Three suppliers accounted for 54% of the 2012–15 national increase. But growth alone does not make every additional prescription inappropriate — ONMAP calls it “not unexpected.” Under-access to appropriate NIV and use of unnecessarily costly equipment are both real.",
    { x: 7.15, y: 5.58, w: 5.48, h: 1.0, fontSize: 11.5 });
  s.addNotes("This is the punchline of the entire regulatory section. ONMAP: 'The clinicians honestly stated that it was easier to prescribe HMV than a device with S/T mode capability' — and they believed the lifelong clinical/technical support that came with HMV was essential. Migration coincided with COPD readmission penalties.");
}

// =====================================================================
// V. CONVERGENCE  (slides 14-16, 28-33 min)
// =====================================================================
{
  const s = newSlide("V · Convergence");
  title(s, "EPAP, pressure support, and backup breaths", "Three features. Three different problems.");
  // ---- WAVEFORM PANELS ----
  const py = TOPS + 0.46, ph = 1.9, pw = 3.72, pgap = 0.39;
  const base = py + ph, top1 = py + 0.38, mid = py + 1.18;
  function panel(i, label, sub) {
    const x = M + i * (pw + pgap);
    s.addShape(pres.ShapeType.line, { x, y: py - 0.02, w: 0, h: ph + 0.02, line: { color: RULE, width: 0.75 } });
    s.addShape(pres.ShapeType.line, { x, y: base, w: pw, h: 0, line: { color: RULE, width: 0.75 } });
    s.addText(label, { x, y: py - 0.42, w: pw, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: INK });
    caption(s, sub, { x, y: base + 0.08, w: pw, h: 0.5 });
    return x;
  }
  // CPAP — flat
  let x0 = panel(0, "CPAP", "One pressure, held.\nExpiratory pressure supports airway patency.");
  wave(s, [[x0 + 0.12, base - 0.18], [x0 + 0.45, base - 0.18], [x0 + 0.45, mid], [x0 + pw - 0.12, mid]], INK, 1.5);
  // Bilevel S — patient-triggered square wave
  x0 = panel(1, "Bilevel  S", "IPAP–EPAP difference supplies inspiratory assistance.\nEvery breath is patient-triggered.");
  let pts = [[x0 + 0.12, base - 0.18], [x0 + 0.35, base - 0.18], [x0 + 0.35, mid]];
  for (let k = 0; k < 3; k++) {
    const a = x0 + 0.55 + k * 1.02;
    pts.push([a, mid], [a, top1], [a + 0.42, top1], [a + 0.42, mid]);
  }
  pts.push([x0 + pw - 0.12, mid]);
  wave(s, pts, INK, 1.5);
  // Bilevel S/T — one machine-triggered breath
  x0 = panel(2, "Bilevel  S/T", "Timed breaths address insufficient spontaneous triggering.\nThe third breath here is machine-triggered.");
  pts = [[x0 + 0.12, base - 0.18], [x0 + 0.35, base - 0.18], [x0 + 0.35, mid]];
  for (let k = 0; k < 3; k++) {
    const a = x0 + 0.55 + k * 1.02;
    pts.push([a, mid], [a, top1], [a + 0.42, top1], [a + 0.42, mid]);
  }
  pts.push([x0 + pw - 0.12, mid]);
  wave(s, pts, INK, 1.5);
  s.addText("T", { x: x0 + 2.59 - 0.2, y: top1 - 0.34, w: 0.5, h: 0.26, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11, bold: true, color: INK, align: "center" });
  s.addShape(pres.ShapeType.line, { x: x0 + 2.59, y: top1 - 0.08, w: 0, h: 0.1, line: { color: INK, width: 1 } });
  s.addShape(pres.ShapeType.rect, { x: M, y: 5.35, w: CW, h: 1.12, fill: { color: WASH }, line: { color: WASH } });
  s.addText([
    { text: "Vocabulary, and nothing more than this today.   ", options: { bold: true } },
    { text: "VAPS (AVAPS, iVAPS) and ASV are control strategies. They are not synonyms for HMV, and they are not synonyms for each other. Rise time, cycling, and manufacturer algorithms are a separate lecture.", options: {} },
  ], { x: M + 0.22, y: 5.35, w: CW - 0.44, h: 1.12, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: INK, valign: "middle", lineSpacingMultiple: 1.15 });
  s.addNotes("Schematic square waves on purpose — clearer than a realistic trace. Walk left to right: what problem does each added feature solve?");
}

{
  const s = newSlide("V · Convergence");
  title(s, "Similar waveforms do not imply equivalent systems");
  const rows = [
    [hdr("Device"), hdr("Principal distinction"), hdr("Code")],
    ["CPAP / APAP", "Single-level positive pressure", b("E0601")],
    ["Bilevel RAD, no backup rate", "Separate IPAP and EPAP, no timed breaths", b("E0470")],
    ["Bilevel RAD with backup rate", "Timed support; further algorithms are model-dependent", b("E0471")],
    ["Home ventilator, noninvasive interface", "Ventilator-class system via mask or mouthpiece", b("E0466")],
    ["Home ventilator, invasive interface", "Ventilator-class system via tracheostomy", b("E0465")],
    ["Multi-function / dual-function ventilator", "Adds O₂ concentration, nebulization, suction, cough assist", b("E0467 / E0468")],
  ];
  table(s, rows, { y: TOP, colW: [4.1, 6.13, 1.7], fontSize: 11.5 });
  s.addShape(pres.ShapeType.line, { x: M, y: 4.62, w: CW, h: 0, line: { color: RULE, width: 0.75 } });
  s.addText("The FDA taxonomy is orthogonal to the CMS taxonomy — and worth thirty seconds", {
    x: M, y: 4.76, w: CW, h: 0.3, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: INK,
  });
  body(s, [
    { text: "FDA sorts by intended use and exhaust type — exhalation valve versus passive port. ", options: { breakLine: true } },
    { text: "NOU ", options: { bold: true } },
    { text: "devices “must be locatable in the event of a recall.”   ", options: { breakLine: false } },
    { text: "BZD ", options: { bold: true } },
    { text: "is “indicated only for the treatment of adult obstructive sleep apnea.”   ", options: { breakLine: true } },
    { text: "MNS ", options: { bold: true } },
    { text: "devices assume patients “should be expected to have no more than minor and transient adverse effects if ventilation/CPAP cannot be provided during extensive periods of time (eg, overnight).”", options: {} },
  ], { x: M, y: 5.08, w: CW, h: 1.1, fontSize: 12 });
  s.addText("That last sentence is the real category boundary — and the FDA drew it around consequence of interruption, which is what you actually care about.", {
    x: M, y: 6.24, w: CW, h: 0.34, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 13, italic: true, color: INK,
  });
  s.addNotes("Do not ask fellows to memorize FDA product codes. The point is only that two agencies drew two different lines through the same devices, and neither line is physiology.");
}

{
  const s = newSlide("V · Convergence");
  title(s, "Convergence is not equivalence");
  quote(s,
    "The current terminology of noninvasive ventilator modes is based on manufacturer whims, like the brand names of drugs. … Instead of a device-based approach, we should adopt a therapeutic approach based on clinically effective ventilator modes.",
    "Hatipoğlu, Lewarski & Chatburn, Respir Care 2024",
    { x: M, y: TOP, w: CW, h: 0.76, fontSize: 13.5 });
  s.addText("Applying the taxonomy", { x: M, y: 3.05, w: 5.9, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "Bilevel S/T = PC-IMV(2)s,s — and that mode exists on RADs with backup rate and on home ventilators alike. The mode you need for guideline-concordant high-intensity COPD ventilation sits on both sides of the reimbursement line",
    "CPAP and “Spontaneous” mode carry an identical tag: PC-CSVs",
    "ASV and iVAPS carry an identical tag: PC-IMV(2)a,a — despite being marketed as unrelated therapies",
  ], { x: M + 0.05, y: 3.35, w: 5.85, h: 2.2, fontSize: 11.5 });
  s.addText("But four differences remain real", { x: 6.85, y: 3.05, w: 5.78, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  const rows = [
    [hdr("Capability"), hdr("RAD"), hdr("HMV")],
    [b("Internal battery"), "No", "Yes — internal + external"],
    [b("Alarms"), "Limited", "Sophisticated"],
    [b("Mouthpiece ventilation"), "No", "Yes"],
    [b("Invasive interface"), "No (E0472 excepted)", "Yes"],
    ["Pressure ceiling", "~25–30 cm H₂O, model-dependent", "Higher"],
    ["Modes", "S, S/T, some PC and VAPS", "Multimodal; VAPS with auto-EPAP"],
  ];
  table(s, rows, { x: 6.85, y: 3.35, w: 5.78, colW: [2.1, 1.84, 1.84], fontSize: 10, margin: [4, 5, 4, 0] });
  s.addShape(pres.ShapeType.rect, { x: M, y: 5.62, w: 5.85, h: 0.78, fill: { color: WASH }, line: { color: WASH } });
  s.addText("What can this patient not safely or effectively do with the simpler system?", {
    x: M + 0.2, y: 5.62, w: 5.45, h: 0.78, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 13, bold: true, italic: true, color: INK, valign: "middle", lineSpacingMultiple: 1.1,
  });
  s.addNotes("Counterweight if time allows: alarms are a capability, not a guarantee. King's 2010 MAUDE review found at least 11 home-ventilation deaths reported; in at least 5 the ventilator did not alarm. One narrative: 'Husband unsure if alarm going off or not, as there were always alarms going off.' Sources disagree on the RAD ceiling — 25 (Gusman & Wolfe) vs 30 (Jimenez & Choi). Say model-dependent.");
}

// =====================================================================
// VI. CURRENT CATEGORIES  (slides 17-21, 33-42 min)
// =====================================================================
{
  const s = newSlide("VI · Current categories");
  title(s, "Coverage follows a diagnosis-specific pathway, not an escalation ladder", "Medicare policy checked September 2026");
  const rows = [
    [hdr("Clinical indication"), hdr("Policy to open")],
    ["OSA — CPAP or qualifying bilevel", "LCD L33718 (rev. 01/01/2024) + Article A52467"],
    ["Neuromuscular or severe thoracic cage disease", "LCD L33800 (rev. 06/09/2025)"],
    ["Central or complex sleep apnea", "LCD L33800"],
    ["Hypoventilation syndrome, including OHS", "LCD L33800"],
    [b("Chronic respiratory failure from COPD — RAD or HMV"), b("NCD 240.9 (effective 06/09/2025)")],
    ["Ventilator — neuromuscular or thoracic restrictive", "NCD 280.1 + ventilator coding guidance"],
  ];
  table(s, rows, { y: TOPS, colW: [6.2, 5.73], fontSize: 12 });
  s.addShape(pres.ShapeType.line, { x: M, y: 4.32, w: CW, h: 0, line: { color: RULE, width: 0.75 } });
  s.addText("Three things fellows get wrong", { x: M, y: 4.46, w: CW, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  const rows2 = [
    [b("1.  NCD 240.9 is narrower than its reputation"), b("2.  “Severe COPD” is gone from L33800"), b("3.  The PaCO₂ threshold is not one number")],
    ["It governs COPD only — and governs both RAD and HMV for COPD in one document. NMD, chest wall disease, OHS, CSA and OSA all remain under the LCDs. The 2025 NCD did not unify home NIV policy; it carved COPD out of it.",
     "Deleted effective 06/09/2025. The Group I / Group II construct no longer exists. If your qualification habit came from that LCD, it is out of date.",
     "≥ 52 mmHg under NCD 240.9 for COPD, for RAD and HMV alike. ≥ 45 mmHg under L33800 for restrictive thoracic disease and hypoventilation syndrome."],
  ];
  table(s, rows2, { y: 4.76, colW: [3.97, 3.98, 3.98], fontSize: 10.5 });
  s.addNotes("This is the navigation slide. Fellows do not need to memorize thresholds — they need to know which document to open and what clinical distinction it encodes.");
}

{
  const s = newSlide("VI · Current categories");
  title(s, "What NCD 240.9 actually requires", "Chronic respiratory failure consequent to COPD  ·  effective June 9, 2025");
  s.addText("RAD with backup rate — initial 6 months, all of:", { x: M, y: TOPS, w: 5.85, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "PaCO₂ ≥ 52 mmHg by ABG, awake, on prescribed FiO₂",
    "Sleep apnea is not the predominant cause — formal sleep testing not required. “Not predominant” is not “absent”",
    "Stability (no new or increased symptom for ≥ 2 days and no pharmacologic change in the prior 2 weeks) or persistence (hypercapnia ≥ 2 weeks post-hospitalization)",
  ], { x: M + 0.05, y: TOPS + 0.3, w: 5.8, h: 1.5, fontSize: 11.5 });
  body(s, [{ text: "By 6 months it must be used as high-intensity therapy: IPAP ≥ 15 cm H₂O and backup rate ≥ 14/min.", options: {} }],
    { x: M, y: TOPS + 1.88, w: 5.85, h: 0.5, fontSize: 11.5 });
  caption(s, "CMS clarified in April 2026 that “high intensity” applies to RADs only — never to HMV.", { x: M, y: TOPS + 2.34, w: 5.85, h: 0.3 });
  s.addText("HMV — must be used in a volume-targeted mode, plus:", { x: 6.85, y: TOPS, w: 5.78, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "PaCO₂ ≥ 52 mmHg, sleep apnea not predominant, and at least one of:",
    "FiO₂ ≥ 36% or ≥ 4 L nasally",
    "Support > 8 hours per 24 hours",
    "Alarms and internal battery needed because the patient “is unable to effectively breathe on their own for more than a few hours” and interruption “is likely to cause a life-threatening condition”",
    "Clinician judgment that no RAD outcome is achievable “because the patient's needs exceed the capabilities of a RAD”",
  ], { x: 6.9, y: TOPS + 0.3, w: 5.73, h: 2.4, fontSize: 11.5 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 4.72, w: CW, h: 0.82, fill: { color: WASH }, line: { color: WASH } });
  s.addText([
    { text: "The most important door in the NCD.   ", options: { bold: true } },
    { text: "A RAD is covered immediately at discharge for 6 months if the patient required a RAD or ventilator in the 24 hours before discharge and the clinician judges risk of rapid symptom exacerbation or PaCO₂ rise. ", options: {} },
    { text: "No ABG threshold on this pathway.", options: { bold: true } },
  ], { x: M + 0.22, y: 4.72, w: CW - 0.44, h: 0.82, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK, valign: "middle", lineSpacingMultiple: 1.12 });
  s.addText("Two asymmetries worth a sentence each", { x: M, y: 5.68, w: CW, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, bold: true, color: INK });
  body(s, "RAD continuing coverage requires adherence (≥ 4 h/24 h on ≥ 70% of days) and a clinical outcome — PaCO₂ < 46, stabilization, 20% reduction, or symptom improvement. HMV continuing coverage requires adherence only. And for volume-targeted HMV used > 8 h/24 h with an oronasal mask at night, a second daytime interface is covered.",
    { x: M, y: 5.98, w: CW, h: 0.8, fontSize: 11 });
  s.addNotes("The discharge pathway is the one fellows will use most and know least. No ABG needed. Months 7-12 require adherence in EACH paid rental month.");
}

{
  const s = newSlide("VI · Current categories");
  title(s, "The non-COPD pathways — and the one that matters most to you");
  s.addShape(pres.ShapeType.rect, { x: M, y: TOP, w: CW, h: 1.42, fill: { color: WASH }, line: { color: WASH } });
  s.addText("Restrictive thoracic / neuromuscular  ·  LCD L33800", { x: M + 0.22, y: TOP + 0.12, w: CW - 0.44, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  s.addText([
    { text: "The only pathway where the clinician freely chooses E0470 vs E0471, and the only one that does not require hypercapnia. ", options: { breakLine: true } },
    { text: "Documented NMD or severe thoracic cage abnormality, COPD not a significant contributor, plus ANY ONE of:  awake PaCO₂ ≥ 45 mmHg  ·  SpO₂ ≤ 88% for ≥ 5 min of nocturnal recording  ·  NMD only: MIP < 60 cm H₂O or FVC < 50% predicted", options: {} },
  ], { x: M + 0.22, y: TOP + 0.42, w: CW - 0.44, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: TEXT, lineSpacingMultiple: 1.15 });
  quote(s, "An ALS patient qualifies on MIP or FVC alone — normal ABG, no sleep study. Fellows routinely wait for hypercapnia they do not need.",
    null, { x: M, y: 2.92, w: CW, h: 0.38, fontSize: 14 });
  const rows = [
    [hdr("Pathway"), hdr("What it requires")],
    [b("Hypoventilation syndrome"), "PaCO₂ ≥ 45 awake, FEV₁/FVC ≥ 70%, plus either a ≥ 7 mmHg sleep or on-waking rise, or SpO₂ ≤ 88% for ≥ 5 min not caused by obstructive events.  E0471 here is a step-up only — it requires a covered E0470 already in use and failure on it."],
    [b("Central / complex sleep apnea"), "Attended facility PSG before initiation — home testing does not qualify. CAHI ≥ 5/h, centrals > 50% of events, and no evidence of hypoventilation, plus demonstrated improvement on the settings to be prescribed."],
    [b("OSA  ·  LCD L33718"), "E0470 requires E0601 tried and proven ineffective despite optimal mask fit and pressure.  E0471 is categorically denied if the primary diagnosis is OSA.  Adherence: ≥ 4 h/night on 70% of nights in a consecutive 30-day period, with in-person re-evaluation no sooner than day 31 and no later than day 91."],
  ];
  table(s, rows, { y: 3.62, colW: [3.1, 8.83], fontSize: 11 });
  s.addText("A gate that applies to every RAD pathway and gets missed: the record must document symptoms of sleep-associated hypoventilation — hypersomnolence, fatigue, morning headache, cognitive dysfunction, dyspnea. No symptoms, no RAD.", {
    x: M, y: 6.14, w: CW, h: 0.44, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 11.5, bold: true, color: INK, lineSpacingMultiple: 1.12,
  });
  s.addNotes("Say the ALS line out loud and pause. It is the single highest-yield practical correction in the talk.");
}

{
  const s = newSlide("VI · Current categories");
  title(s, "Matching the device to the clinical problem");
  const rows = [
    [hdr("Problem"), hdr("First-line"), hdr("Escalate when"), hdr("Evidence")],
    ["Uncomplicated OSA", "CPAP / APAP, not bilevel", "CPAP fails despite optimal fit and pressure", "Patil 2019 (AASM) suggests CPAP/APAP over bilevel"],
    ["Stable ambulatory OHS with severe OSA", "CPAP", "Persistent hypercapnia despite adherence → BPAP S/T", "Mokhlesi 2019 (ATS), conditional"],
    ["Hospitalized, suspected OHS", "Discharge on NIV pending evaluation and titration", "—", "Mokhlesi 2019 — do not generalize to stable outpatients"],
    ["COPD, persistent hypercapnia after recovery", "BPAP with backup rate, high-intensity, targeting real PaCO₂ reduction", "FiO₂ needs, > 8 h/day, alarms + battery, or RAD targets unachievable", "Köhnlein 2014; Murphy 2017; Macrea 2020 (ATS)"],
    ["Neuromuscular / chest wall", "BPAP (S or S/T) early — MIP/FVC criteria; do not wait for CO₂", "Daytime or mouthpiece support, > 10 h/day, dependence", "Ward 2005; Khan 2023 (CHEST)"],
    ["Central sleep apnea", "Depends on etiology — CPAP, backup-rate bilevel, or ASV", "—", "Badr 2025 (AASM); bilevel without backup is suggested against"],
  ];
  table(s, rows, { y: TOP, colW: [2.38, 3.3, 3.3, 2.95], fontSize: 10.5 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 4.88, w: CW, h: 0.82, fill: { color: WASH }, line: { color: WASH } });
  s.addText([
    { text: "Read the table down the first column, not across the last.   ", options: { bold: true } },
    { text: "Every row is a different physiologic problem. The device follows from the problem and from what happens if support stops — not from how sick the patient looks, and not from which code is easier to get approved.", options: {} },
  ], { x: M + 0.22, y: 4.88, w: CW - 0.44, h: 0.82, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK, valign: "middle", lineSpacingMultiple: 1.12 });
  caption(s, "Detailed qualification thresholds for each pathway are in the backup slides.", { x: M, y: 5.88, w: CW, h: 0.3 });
  s.addNotes("Do not linger. This is a reference slide; the fellows will photograph it. Spend the time on the COPD row and the neuromuscular row.");
}

{
  const s = newSlide("VI · Current categories");
  title(s, "COPD: selection and timing, not the device label");
  const rows = [
    [hdr("Trial"), hdr("Population and timing"), hdr("Result")],
    [b("Köhnlein 2014"), "Stable severe hypercapnic COPD; NIV targeted substantial PaCO₂ reduction", "One-year mortality 12% vs 33%.  ARR 21%, NNT 5"],
    [b("RESCUE  (Struik 2014)"), "Hypercapnia persisting > 48 h after stopping acute ventilatory support", "No improvement in readmission-or-death despite better CO₂"],
    [b("HOT-HMV  (Murphy 2017)"), "Persistent hypercapnia 2–4 weeks after resolution of acidosis", "Median time to readmission or death 1.4 → 4.3 months.  Adjusted HR 0.49 (0.31–0.77)"],
  ];
  table(s, rows, { y: TOP, colW: [2.75, 4.85, 4.33], fontSize: 11.5 });
  s.addShape(pres.ShapeType.line, { x: M, y: 4.05, w: CW, h: 0, line: { color: RULE, width: 0.75 } });
  s.addText("Two things to say explicitly", { x: M, y: 4.19, w: CW, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "HOT-HMV tested a home NIV strategy. It did not randomize an E0466 billing category against an E0471 billing category — do not let the trial name stand in for a code",
    "The contrast supports attention to persistence and patient selection. It does not isolate timing as the only explanation for the differing results",
    "ATS guidance: assess for OSA, reassess after recovery from an exacerbation, and target meaningful reduction or normalization of PaCO₂. Reassessing at 2–4 weeks is not a reason to withdraw support from someone who cannot be weaned",
  ], { x: M + 0.05, y: 4.49, w: CW - 0.05, h: 1.9, fontSize: 12 });
  s.addNotes("If you cite ONMAP's figure for Kohnlein, note their '36% relative reduction' does not reconcile with 12% vs 33%. Use ARR 21% / NNT 5 (D'Cruz & Hart).");
}

{
  const s = newSlide("VI · Current categories");
  title(s, "The prescription is a care plan, not a brand name");
  const fw = 1.75, fgap = 0.29, fy = TOP + 0.05;
  ["Indication\n& goal", "Mode and\nsettings", "Interface\n& oxygen", "Safety\nneeds", "Training", "Who reassesses,\nand when"].forEach((t, i) => {
    const x = M + i * (fw + fgap);
    box(s, t, { x, y: fy, w: fw, h: 0.72, fontSize: 10 });
    if (i < 5) arrow(s, x + fw + 0.04, fy + 0.36, 0.21);
  });
  s.addText("Three distinct follow-up questions, routinely collapsed into one", { x: M, y: fy + 1.0, w: 5.85, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "Is the patient using it?",
    "Is it providing effective support?",
    "Is the patient benefiting?",
  ], { x: M + 0.05, y: fy + 1.3, w: 5.8, h: 1.0, fontSize: 13 });
  caption(s, "The download answers only the first.", { x: M, y: fy + 2.3, w: 5.8, h: 0.3 });
  s.addText("Do not put these in a justification letter — they will not work", { x: 6.85, y: fy + 1.0, w: 5.78, h: 0.28, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, color: INK });
  bullets(s, [
    "Need for portability",
    "Need for an internal battery because home power is unstable",
    "Need for air travel",
    "Need for multimodal support — cough assist, nebulizer, suction",
    "Improvement on a hospital ventilator running settings the patient's own bilevel could deliver",
  ], { x: 6.9, y: fy + 1.3, w: 5.73, h: 1.7, fontSize: 11.5 });
  s.addShape(pres.ShapeType.rect, { x: M, y: 5.55, w: CW, h: 0.92, fill: { color: WASH }, line: { color: WASH } });
  s.addText([
    { text: "New, and worth flagging.   ", options: { bold: true } },
    { text: "Effective October 28, 2026, E0466 / E0467 / E0468 join the Face-to-Face Encounter and Written Order Prior to Delivery list. A home NIV ordered without a documented pre-delivery visit and written order is not payable — regardless of NCD 240.9. E0465 and the RAD codes are not on the list, and this is not prior authorization.", options: {} },
  ], { x: M + 0.22, y: 5.55, w: CW - 0.44, h: 0.92, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: INK, valign: "middle", lineSpacingMultiple: 1.12 });
  s.addNotes("For a dependent patient: an explicit interruption-response plan. Note CMS supports a second device only for the wheelchair-mounted plus bedside scenario — 'documentation that states the need for a second device as backup in case of device failure is not supported by CMS' — even though patients requiring >18 h/day are functionally ventilator-dependent.");
}

// =====================================================================
// VII. CLOSE  (slide 22, 42-45 min)
// =====================================================================
{
  const s = newSlide("VII · Close");
  title(s, "Return to the three patients");
  const rows = [
    [hdr(""), hdr("Physiologic goal"), hdr("Required capabilities"), hdr("Coverage route")],
    [b("A  ·  OSA"), "Splint the airway", "CPAP. Nothing more", "E0601, LCD L33718"],
    [b("B  ·  COPD"), "Reduce PaCO₂ meaningfully", "High-intensity BPAP S/T", "E0471, NCD 240.9 — or the discharge pathway if she leaves on it"],
    [b("C  ·  ALS"), "Follow a changing pump", "Starts as a RAD; becomes an HMV when daytime mouthpiece support, battery and alarms become the need", "L33800 on FVC/MIP criteria, then NCD 280.1"],
  ];
  table(s, rows, { y: TOP, colW: [1.75, 2.85, 4.55, 2.78], fontSize: 11 });
  s.addShape(pres.ShapeType.line, { x: M, y: 3.98, w: CW, h: 0, line: { color: INK, width: 1.25 } });
  s.addText("Three things to retain", { x: M, y: 4.12, w: CW, h: 0.3, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: INK });
  const tw = 3.72;
  [["A treatment mode is not a device class.", "Bilevel S/T exists on both sides of the reimbursement line."],
   ["Choose capabilities from clinical need and consequence of interruption.", "Battery, alarms, daytime interface, invasive route. Everything else has converged."],
   ["The categories are historical artifacts, not physiology.", "A 1992 marketing sentence, a 1993 statute, a 1998 spending alarm. Knowing that is what lets you argue with the denial letter."]].forEach((t, i) => {
    const x = M + i * (tw + 0.39);
    s.addText(String(i + 1), { x, y: 4.52, w: 0.4, h: 0.34, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 18, bold: true, color: INK });
    s.addText(t[0], { x, y: 4.9, w: tw, h: 0.62, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, bold: true, color: INK, lineSpacingMultiple: 1.12 });
    s.addText(t[1], { x, y: 5.56, w: tw, h: 0.8, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 11.5, color: TEXT, lineSpacingMultiple: 1.15 });
  });
  s.addNotes("Optional closing quote: 'It has been stated that ontogeny recapitulates phylogeny and perhaps so should technology recapitulate pathology. Technology in this arena, however, has taken off in such a quantum fashion that it assumed its own trajectory.' — Gay & Owens, ONMAP 2021. Then 5 minutes of discussion.");
}

pres.writeFile({ fileName: path.join(__dirname, "PAP_RAD_HMV_draft.pptx") })
  .then(f => console.log("wrote", f))
  .catch(error => {
    console.error("Failed to write presentation:", error);
    process.exitCode = 1;
  });
