// VWC Branded Case Study Generator
// Reads structured case study data and produces branded .docx files
// matching VWC brand guide specs (palette, typography, layout).
//
// Usage: node generate-case-studies.js
// Output: ./docx/NN-slug.docx (one per deal)

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber,
  HeadingLevel, ExternalHyperlink, TabStopType, TabStopPosition, PageBreak
} = require("docx");

// ---------- Brand spec ----------
const BRAND = {
  green:        "A3D55D",  // VWC Green (light/accent)
  greenDark:    "5C9627",  // For text/links/headlines on white (AA contrast)
  greenTint20:  "DAF0BC",  // Chart fills, callout backgrounds
  charcoal:     "40464B",  // Primary text
  charcoal80:   "666B70",  // Secondary text
  midGray:      "8A8F94",  // Captions
  lightGray:    "E5E7EA",  // Borders
  offWhite:     "F7F7F7",  // Section backgrounds
  white:        "FFFFFF",
  fontHead:     "Lato",         // brand fallback for headlines
  fontBody:     "Open Sans",    // brand fallback for body
};

const LOGO_PATH = path.resolve(__dirname, "../../../brand/logo/logo-small.png");
const OUTPUT_DIR = path.resolve(__dirname, "docx");

// ---------- Helpers ----------
const mm = (pt) => Math.round(pt * 20); // half-points
const dxa = (inches) => Math.round(inches * 1440);
const PAGE_W = 12240;
const PAGE_H = 15840;
const MARGIN = 1008; // 0.7"
const CONTENT_W = PAGE_W - MARGIN * 2; // 10224

function txt(s, opts = {}) {
  return new TextRun({
    text: s,
    font: opts.font || BRAND.fontBody,
    size: opts.size || 22,            // half-points (22 = 11pt)
    color: opts.color || BRAND.charcoal,
    bold: opts.bold || false,
    italics: opts.italics || false,
  });
}

function p(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.alignment,
    spacing: opts.spacing || { before: 40, after: 40, line: 280 },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
  });
}

function divider(color = BRAND.lightGray, size = 6) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    border: {
      bottom: { color, space: 1, style: BorderStyle.SINGLE, size },
    },
    children: [new TextRun({ text: "" })],
  });
}

// Cell with consistent padding
function cell(content, opts = {}) {
  const children = Array.isArray(content) ? content : [content];
  return new TableCell({
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    borders: opts.borders || {
      top: { style: BorderStyle.SINGLE, size: 4, color: BRAND.lightGray },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: BRAND.lightGray },
      left: { style: BorderStyle.SINGLE, size: 4, color: BRAND.lightGray },
      right: { style: BorderStyle.SINGLE, size: 4, color: BRAND.lightGray },
    },
    children: children.map((c) =>
      typeof c === "string"
        ? p(txt(c, opts.textOpts || {}))
        : (c instanceof TextRun ? p(c) : c)
    ),
  });
}

// ---------- Document builders ----------

function buildHeader() {
  // Logo top-left
  const logoData = fs.readFileSync(LOGO_PATH);
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 120 },
        children: [
          new ImageRun({
            type: "png",
            data: logoData,
            transformation: { width: 150, height: 25 },
            altText: { title: "VisionWise Capital", description: "VWC Logo", name: "vwc_logo" },
          }),
        ],
      }),
    ],
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        spacing: { before: 0, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: BRAND.green } },
        children: [new TextRun("")],
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 80, after: 0 },
        children: [
          txt("VisionWise Capital  ·  ", { size: 16, color: BRAND.midGray }),
          new ExternalHyperlink({
            link: "https://visionwisecapital.com",
            children: [
              new TextRun({
                text: "visionwisecapital.com",
                font: BRAND.fontBody,
                size: 16,
                color: BRAND.greenDark,
              }),
            ],
          }),
          txt("  ·  (949) 441-5580  ·  ", { size: 16, color: BRAND.midGray }),
          txt("Placing Multifamily Real Estate Within Your Reach", {
            size: 16, color: BRAND.midGray, italics: true,
          }),
        ],
      }),
    ],
  });
}

// Eyebrow + property title block
function buildTitle(d) {
  const eyebrowColor = d.statusBadge === "stabilized" ? BRAND.charcoal80
                     : d.statusBadge === "renovation" ? BRAND.charcoal80
                     : BRAND.greenDark;

  return [
    p([
      txt("CASE STUDY", { font: BRAND.fontHead, size: 16, color: eyebrowColor, bold: true }),
      txt("   ·   ", { font: BRAND.fontHead, size: 16, color: BRAND.midGray }),
      txt(d.statusLabel.toUpperCase(), { font: BRAND.fontHead, size: 16, color: eyebrowColor, bold: true }),
    ], { spacing: { before: 80, after: 20 } }),

    p(txt(d.property, { font: BRAND.fontHead, size: 44, bold: true, color: BRAND.charcoal }),
      { spacing: { before: 0, after: 20, line: 280 } }),

    p([
      txt(d.city, { size: 20, color: BRAND.charcoal80 }),
      txt("   ·   ", { size: 20, color: BRAND.midGray }),
      txt(`${d.units} units`, { size: 20, color: BRAND.charcoal80 }),
      txt("   ·   ", { size: 20, color: BRAND.midGray }),
      txt(`Acquired ${d.acquired}`, { size: 20, color: BRAND.charcoal80 }),
      ...(d.exited ? [
        txt("   ·   ", { size: 20, color: BRAND.midGray }),
        txt(`${d.statusLabel === "Sold" ? "Sold" : "Exit"} ${d.exited}`, { size: 20, color: BRAND.charcoal80 }),
      ] : []),
    ], { spacing: { before: 0, after: 80 } }),
  ];
}

// Headline KPI banner (return + profit)
function buildHeadlineBanner(d) {
  const isLoss = d.return.startsWith("-");
  const fill = isLoss ? "F2F2F2" : BRAND.greenTint20;
  const accent = isLoss ? BRAND.charcoal : BRAND.greenDark;

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [Math.round(CONTENT_W / 2), Math.round(CONTENT_W / 2)],
    rows: [
      new TableRow({
        children: [
          cell([
            p(txt(d.returnLabel.toUpperCase(), {
              font: BRAND.fontHead, size: 16, color: BRAND.charcoal80, bold: true,
            }), { spacing: { before: 40, after: 0 } }),
            p(txt(d.return, {
              font: BRAND.fontHead, size: 52, bold: true, color: accent,
            }), { spacing: { before: 0, after: 20 } }),
          ], {
            width: Math.round(CONTENT_W / 2),
            fill,
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
                       left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          }),
          cell([
            p(txt(d.profitLabel.toUpperCase(), {
              font: BRAND.fontHead, size: 16, color: BRAND.charcoal80, bold: true,
            }), { spacing: { before: 40, after: 0 } }),
            p(txt(d.profit, {
              font: BRAND.fontHead, size: 52, bold: true, color: accent,
            }), { spacing: { before: 0, after: 20 } }),
          ], {
            width: Math.round(CONTENT_W / 2),
            fill,
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
                       left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
          }),
        ],
      }),
    ],
  });
}

// Section header (small green line + bold charcoal heading)
function sectionHeader(label) {
  return [
    new Paragraph({
      spacing: { before: 140, after: 40 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BRAND.green } },
      children: [
        new TextRun({
          text: label.toUpperCase(),
          font: BRAND.fontHead,
          size: 20,
          bold: true,
          color: BRAND.charcoal,
        }),
      ],
    }),
  ];
}

// Snapshot table — 2 columns x N rows
function buildSnapshot(d) {
  const colW = Math.round(CONTENT_W / 2);
  const labelOpts = { textOpts: { size: 18, color: BRAND.charcoal80, bold: true, font: BRAND.fontHead }, width: Math.round(colW * 0.45), fill: BRAND.offWhite };
  const valueOpts = { textOpts: { size: 22, color: BRAND.charcoal, bold: true }, width: Math.round(colW * 0.55) };

  const rows = [
    [["Acquisition Price", d.acqPrice], ["Loan Amount", d.loan]],
    [["Per Unit", d.acqPerUnit], ["Loan-to-Cost", d.ltv]],
    [["CapEx", d.capex], ["Equity Invested", d.equity]],
    [["Total Cost", d.totalCost], [d.exitLabel || (d.isEstimate ? "Est. Market Value" : "Sale Price"), d.exitPrice]],
  ];

  if (d.exitPerUnit) rows.push([["Hold Period", d.hold || "—"], ["Exit Per Unit", d.exitPerUnit]]);

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [
      Math.round(colW * 0.45), Math.round(colW * 0.55),
      Math.round(colW * 0.45), Math.round(colW * 0.55),
    ],
    rows: rows.map((row) => new TableRow({
      children: [
        cell(row[0][0], labelOpts),
        cell(row[0][1], valueOpts),
        cell(row[1][0], labelOpts),
        cell(row[1][1], valueOpts),
      ],
    })),
  });
}

function buildBody(d) {
  const out = [];

  // Title block
  out.push(...buildTitle(d));

  // Headline banner
  out.push(buildHeadlineBanner(d));

  // The Opportunity
  out.push(...sectionHeader("The Opportunity"));
  out.push(p(txt(d.thesis, { size: 21 }), { spacing: { before: 40, after: 40, line: 280 } }));

  // The Strategy
  out.push(...sectionHeader("The Strategy"));
  out.push(p(txt(d.strategy, { size: 21 }), { spacing: { before: 40, after: 40, line: 280 } }));

  // Snapshot
  out.push(...sectionHeader("Deal Snapshot"));
  out.push(buildSnapshot(d));

  // Investor Takeaway (highlighted callout)
  out.push(...sectionHeader("Investor Takeaway"));
  out.push(new Paragraph({
    spacing: { before: 40, after: 40, line: 280 },
    indent: { left: 180, right: 180 },
    border: {
      left: { style: BorderStyle.SINGLE, size: 16, color: BRAND.green, space: 12 },
    },
    children: [txt(d.takeaway, { size: 21 })],
  }));

  // Disclaimer
  out.push(new Paragraph({
    spacing: { before: 160, after: 60 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: BRAND.lightGray } },
    children: [new TextRun("")],
  }));
  out.push(p(txt(
    "This communication is provided for informational purposes only and does not constitute an offer to sell or a solicitation of an offer to buy any securities. Any such offer or solicitation will be made only through definitive offering documents (including a private placement memorandum) provided to qualified investors. Past performance is not indicative of future results. Real estate investments involve risk, including loss of principal.",
    { size: 13, color: BRAND.midGray, italics: true }
  ), { spacing: { before: 0, after: 0, line: 220 } }));

  return out;
}

function buildDoc(d) {
  return new Document({
    creator: "VisionWise Capital",
    title: `${d.property} — Case Study`,
    description: `VWC Case Study: ${d.property}, ${d.city}`,
    styles: {
      default: {
        document: { run: { font: BRAND.fontBody, size: 22, color: BRAND.charcoal } },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.PORTRAIT },
          margin: { top: dxa(0.7), right: MARGIN, bottom: dxa(0.7), left: MARGIN, header: dxa(0.3), footer: dxa(0.3) },
        },
      },
      headers: { default: buildHeader() },
      footers: { default: buildFooter() },
      children: buildBody(d),
    }],
  });
}

// ---------- Case study data ----------
const data = require("./case-study-data.js");

(async () => {
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  for (const d of data) {
    const doc = buildDoc(d);
    const buf = await Packer.toBuffer(doc);
    const out = path.join(OUTPUT_DIR, `${d.num}-${d.slug}.docx`);
    fs.writeFileSync(out, buf);
    console.log(`✓ ${d.num} ${d.property}`);
  }
  console.log(`\nGenerated ${data.length} case studies → ${OUTPUT_DIR}`);
})();
