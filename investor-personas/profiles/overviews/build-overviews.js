// Builds: 10 per-persona Company Overview Word docs
// Run: NODE_PATH="<global npm node_modules>" node build-overviews.js
//
// Each output is a 3-4 page conversion document tailored to ONE persona,
// brand-aligned per company-intel/brand/brand-guide.md.
//
// Output directory: same folder as this script.

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, ShadingType, PageNumber, PageBreak, VerticalAlign,
} = require("docx");

// ============================================================
// BRAND CONSTANTS
// ============================================================
const PAGE_WIDTH = 12240, PAGE_HEIGHT = 15840;
const MARGIN = 1080;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2;

const COLOR = {
  green: "A3D55D",
  greenDark: "5C9627",
  greenTint: "DAF0BC",
  charcoal: "40464B",
  charcoalTint: "666B70",
  mid: "8A8F94",
  light: "E5E7EA",
  offWhite: "F7F7F7",
  white: "FFFFFF",
};

const HEADING_FONT = "Lato";
const BODY_FONT = "Open Sans";

const cellMargins = { top: 100, bottom: 100, left: 140, right: 140 };
const allBorders = (color = COLOR.light, size = 4) => ({
  top: { style: BorderStyle.SINGLE, size, color },
  bottom: { style: BorderStyle.SINGLE, size, color },
  left: { style: BorderStyle.SINGLE, size, color },
  right: { style: BorderStyle.SINGLE, size, color },
});

// ============================================================
// HELPERS
// ============================================================
const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: opts.spacing || { after: 140 },
  children: [new TextRun({
    text,
    font: opts.font || BODY_FONT,
    size: opts.size || 22,
    bold: !!opts.bold,
    italics: !!opts.italic,
    color: opts.color || COLOR.charcoal,
  })],
});

const pRich = (runs, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: opts.spacing || { after: 140 },
  children: runs.map(r => new TextRun({
    font: r.font || BODY_FONT,
    size: r.size || 22,
    text: r.text,
    bold: !!r.bold,
    italics: !!r.italic,
    color: r.color || COLOR.charcoal,
  })),
});

const h1 = (text, pageBreak = true) => new Paragraph({
  spacing: { before: 280, after: 200 },
  pageBreakBefore: pageBreak,
  children: [new TextRun({
    text, font: HEADING_FONT, size: 48, bold: true, color: COLOR.charcoal,
  })],
});

const h2 = (text) => new Paragraph({
  spacing: { before: 240, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR.green, space: 4 } },
  children: [new TextRun({
    text, font: HEADING_FONT, size: 32, bold: true, color: COLOR.charcoal,
  })],
});

const h3 = (text) => new Paragraph({
  spacing: { before: 180, after: 80 },
  children: [new TextRun({
    text, font: HEADING_FONT, size: 24, bold: true, color: COLOR.greenDark,
  })],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60 },
  children: [new TextRun({ text, font: BODY_FONT, size: 22, color: COLOR.charcoal })],
});

const spacer = (size = 120) => new Paragraph({
  spacing: { after: size }, children: [new TextRun({ text: "" })],
});

const quote = (text) => new Paragraph({
  spacing: { before: 120, after: 120 },
  indent: { left: 360 },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: COLOR.green, space: 12 } },
  children: [new TextRun({
    text, font: BODY_FONT, size: 22, italics: true, color: COLOR.charcoalTint,
  })],
});

const callout = (label, text) => new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [CONTENT_WIDTH],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: {
            top: { style: BorderStyle.NIL },
            bottom: { style: BorderStyle.NIL },
            right: { style: BorderStyle.NIL },
            left: { style: BorderStyle.SINGLE, size: 24, color: COLOR.green },
          },
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          shading: { fill: COLOR.greenTint, type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 280, right: 240 },
          children: [
            new Paragraph({
              spacing: { after: 80 },
              children: [new TextRun({
                text: label, font: HEADING_FONT, size: 18, bold: true, color: COLOR.charcoal,
              })],
            }),
            new Paragraph({
              spacing: { after: 0 },
              children: [new TextRun({
                text, font: BODY_FONT, size: 24, italics: true, color: COLOR.charcoal,
              })],
            }),
          ],
        }),
      ],
    }),
  ],
});

const textCell = (text, opts = {}) => new TableCell({
  borders: allBorders(opts.borderColor || COLOR.light, 4),
  width: { size: opts.width, type: WidthType.DXA },
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  margins: cellMargins,
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    spacing: { after: 0 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text,
      font: opts.font || BODY_FONT,
      size: opts.size || 20,
      bold: !!opts.bold,
      italics: !!opts.italic,
      color: opts.color || COLOR.charcoal,
    })],
  })],
});

const buildKVTable = (rows, leftWidth = 3000) => {
  const rightWidth = CONTENT_WIDTH - leftWidth;
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [leftWidth, rightWidth],
    rows: rows.map(([k, v], i) => new TableRow({
      children: [
        textCell(k, { width: leftWidth, fill: COLOR.offWhite, bold: true }),
        textCell(v, { width: rightWidth, fill: i % 2 === 0 ? COLOR.white : COLOR.offWhite }),
      ],
    })),
  });
};

const buildHeaderTable = (headers, dataRows, columnWidths, headerFill = COLOR.charcoal) => {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => textCell(h, {
          width: columnWidths[i], fill: headerFill, color: COLOR.white,
          bold: true, size: 20, font: HEADING_FONT,
        })),
      }),
      ...dataRows.map((row, ri) => new TableRow({
        children: row.map((c, i) => textCell(c, {
          width: columnWidths[i],
          fill: ri % 2 === 0 ? COLOR.white : COLOR.offWhite,
          size: 20,
        })),
      })),
    ],
  });
};

// ============================================================
// LOGO
// ============================================================
const LOGO_PATH = path.resolve(__dirname, "..", "..", "..", "company-intel", "brand", "logo", "logo-rgb.png");
const logoBuffer = fs.existsSync(LOGO_PATH) ? fs.readFileSync(LOGO_PATH) : null;

// ============================================================
// CASE STUDY LIBRARY (pulled from company-intel/track-record/case-studies/)
// ============================================================
const CASE_STUDIES = {
  cypress374: {
    title: "374 Cypress Drive — Laguna Beach (Cypress Garden Apartments)",
    status: "Stabilized · 10 units · Acquired September 2020",
    headlineReturn: "Estimated 72.30% levered project return",
    summary: "A 10-unit \"trophy\" coastal asset originally built in 1948 — 1,618 feet from Main Beach in Laguna. The acquisition is itself the story: VWC reviewed the property in April 2020 at $1M below the listing price. The team waited four months while the seller resisted. By July, terms had moved to VWC. The deal closed in September 2020 — VWC's first acquisition for Multifamily Fund IV.",
    lesson: "Patience on entry, not optimism on exit. The broker cited two reasons VWC won the deal: the completeness of the buyer package, and VWC's existing ownership of two other Laguna Beach properties within walking distance. Submarket density compounds.",
    numbers: [
      ["Acquisition", "$4,050,000 ($405K/unit)"],
      ["CapEx", "$391,749"],
      ["Total cost", "$5,014,749"],
      ["Loan", "$2,822,000 (56% LTV)"],
      ["Estimated market value", "$6,600,000 ($660K/unit)"],
      ["Estimated profit", "$1,585,251"],
      ["Levered return", "72.30%"],
    ],
  },
  charlene4921: {
    title: "4921 Charlene Circle — Huntington Beach",
    status: "Sold · 6 units · Acquired December 2020 · Sold December 2022",
    headlineReturn: "51.09% levered project return — 24-month cycle",
    summary: "The sister building to VWC's 4931 Charlene Circle property — literally next door, same building footprint, shared driveway. VWC walked into this deal with a verbatim playbook: same neighborhood, same product, same renovation scope, same property manager. From acquisition to sale, exactly 24 months. Improvements were complete by June 2021 in time for the summer leasing season.",
    lesson: "When you operate enough Southern California coastal multifamily, sister-property acquisitions become low-risk repeats of work you've already proven. Investors paid for an information advantage, not a market guess.",
    numbers: [
      ["Acquisition", "$2,090,000 ($348K/unit)"],
      ["CapEx", "$173,950"],
      ["Total cost", "$2,513,350"],
      ["Loan", "$1,365,000 (54% LTV)"],
      ["Sale price", "$3,100,000 ($517K/unit)"],
      ["Realized profit", "$586,650"],
      ["Levered return", "51.09%"],
    ],
  },
  encinitas: {
    title: "330 W I Street — Encinitas (North San Diego County)",
    status: "In renovation · 21 units · Acquired September 2024",
    headlineReturn: "Estimated 38.99% levered project return",
    summary: "VWC's largest acquisition to date — at $13.15M, roughly 2.3x the size of the next-largest VWC transaction. Encinitas is one of the most supply-constrained coastal submarkets in San Diego County. A heavy reposition with $1.2M projected CapEx — the largest CapEx commitment in VWC's history.",
    lesson: "A live demonstration that VWC's operating model — sub-50% LTV, value-add coastal multifamily, in-house construction — holds at scale. The framework didn't change as the firm grew. Only the deal size did.",
    numbers: [
      ["Acquisition", "$13,150,000 ($626K/unit)"],
      ["CapEx (projected)", "$1,200,000"],
      ["Total cost (projected)", "$15,484,500"],
      ["Loan", "$7,750,000 (50% LTV)"],
      ["Estimated market value", "$18,500,000 ($881K/unit)"],
      ["Estimated profit", "$3,015,500"],
      ["Levered return", "38.99%"],
    ],
  },
  washington: {
    title: "630 W Washington Avenue — San Diego",
    status: "Stabilized · 12 units · Acquired December 2021",
    headlineReturn: "Estimated -33.01% levered project return (paper loss, unrealized)",
    summary: "The only deal in the VWC portfolio currently marked at a paper loss. Acquired in the last weeks of 2021, before the rate environment shifted decisively in 2022. VWC has chosen to operate the asset through the cap-rate widening rather than force a disposition into a weak window. Conservative 50% LTV gave VWC the optionality to hold.",
    lesson: "Honest track records include the deals that didn't work. A 70%–80% LTV operator would face refinance pressure in this rate environment. VWC can wait. The story this asset tells is about leverage discipline — investors who avoided forced losses through 2022–2024 set their LTV ceiling before the cycle turned.",
    numbers: [
      ["Acquisition", "$3,530,900 ($294K/unit)"],
      ["CapEx", "$334,949"],
      ["Total cost", "$4,255,203"],
      ["Loan", "$2,118,600 (50% LTV)"],
      ["Estimated market value", "$3,550,000"],
      ["Estimated paper loss", "-$705,203"],
      ["Levered return (estimated)", "-33.01%"],
    ],
  },
  charlene4931: {
    title: "4931 Charlene Circle — Huntington Beach",
    status: "Sold · 6 units · Acquired August 2019 · Sold August 2022",
    headlineReturn: "43.42% levered project return",
    summary: "The first of VWC's three Charlene Circle properties. Standard value-add lifecycle in a stable Huntington Beach single-family neighborhood. The success here is what enabled the verbatim-playbook sister-property purchase at 4921 Charlene 16 months later.",
    lesson: "Operating experience is the moat. The first Charlene deal was discovery; the second was repetition. Investors who participated in both saw what programmatic execution looks like in practice.",
    numbers: [
      ["Acquisition", "$1,930,000 ($322K/unit)"],
      ["CapEx", "$171,416"],
      ["Total cost", "$2,488,516"],
      ["Loan", "$965,000 (50% LTV)"],
      ["Sale price", "$3,150,000"],
      ["Realized profit", "$661,484"],
      ["Levered return", "43.42%"],
    ],
  },
  cypress201: {
    title: "201 Cypress Drive — Laguna Beach",
    status: "Sold · 7 units · Acquired March 2021 · Sold May 2023",
    headlineReturn: "59.13% levered project return — 26-month cycle",
    summary: "A coastal Laguna asset adjacent to VWC's other Cypress Drive holdings. Acquired during the post-COVID coastal-flight tailwind, executed on schedule, sold into the early-2023 rebound. The submarket density VWC has built around Cypress Drive has now produced three consecutive deals at strong returns.",
    lesson: "Coastal Laguna is supply-constrained at any cycle. The combination of submarket knowledge, existing operating presence, and disciplined entry is what produces compound returns.",
    numbers: [
      ["Acquisition", "$3,150,000 ($450K/unit)"],
      ["CapEx", "$313,884"],
      ["Total cost", "$3,902,884"],
      ["Loan", "$2,047,500 (65% LTV)"],
      ["Sale price", "$5,000,000"],
      ["Realized profit", "$1,097,116"],
      ["Levered return", "59.13%"],
    ],
  },
};

// ============================================================
// PERSONA-SPECIFIC CONTENT
// ============================================================
// For each persona we need:
//  - filename slug
//  - audience subtitle (cover)
//  - opening "for [you]" framing
//  - "Why we're writing" paragraph
//  - "Why VWC matters to YOU" — bullets tailored to their pain points
//  - Two case studies that resonate
//  - Product mechanics block
//  - Next-step CTA
//  - Sign-off context

const personas = [
  {
    id: "P1",
    slug: "01-family-office",
    name: "Family Office Capital Partner",
    audience: "For Family Offices Considering Multifamily Real Estate",
    salutation: "For the Family Office Principal or Investment Committee",
    openingHook: "We are writing to a small number of Family Offices because we are looking for two or three long-term capital partners — not one hundred investors.",
    whyWriting:
      "VisionWise Capital has a 12-year track record acquiring and improving sub-institutional multifamily property in coastal Southern California. We have transacted on 24 deals representing 227 units, $103 million of total cost basis, and a 23.95% average levered project return — including the one paper-loss deal we will walk you through transparently. We are now structuring the next phase of growth around a small number of Family Office capital partnerships.",
    whatWeWantFromYou:
      "We are inviting two or three Family Offices to anchor a $30 million revolving Line of Credit. The structure pays you a coupon. We use the LOC to acquire single-asset multifamily LLCs all-cash. We refinance you out within 90 to 100 days of close-of-escrow. Then we recycle. The result for you: predictable, asset-backed income, with VWC absorbing acquisition execution risk.",
    whyMattersToYou: [
      "You get the operator without owning the operator. VWC carries 56 staff including a 30-year construction lead and ~35 in-house tradesmen — capability your office shouldn't need to build.",
      "You concentrate, not diversify. We are inviting two or three Family Offices, not soliciting at scale. Your relationship with our principal is direct.",
      "Your capital is real-property-backed. Single-asset LLCs at 50% LTV. Refinance-out within 100 days means your exposure on any one deal is short.",
      "We disclose losses transparently. The 630 W Washington deal in our portfolio is at a paper -33%. It is in our deck and on this page — because honest track records include the deals that didn't work.",
    ],
    productMechanicsTitle: "How a $30 million LOC partnership with VWC works",
    productMechanics: [
      ["Capital structure", "Revolving Line of Credit, $30M target — open to one to three Family Office partners"],
      ["Use of capital", "All-cash acquisitions of single-asset multifamily LLCs in coastal SoCal"],
      ["Take-out", "Refinance the LOC partner within 90–100 days of close-of-escrow"],
      ["Underwriting discipline", "5–75 unit sub-institutional multifamily, vintage 1940–2000, ≤50% LTV after month 6"],
      ["Reporting", "AppFolio Investor portal; quarterly written update from Sanford Coggins directly"],
      ["Capital goal context", "$10M raised in 2026, $25M in 2027 — disciplined growth, not roll-up"],
    ],
    caseStudies: ["cypress374", "encinitas"],
    nextStep: "We would like 30 minutes — in person if you are in Southern California, by phone otherwise — to walk through the LOC structure and answer your underwriting questions before we send you any documents. Please reply to this letter to schedule.",
    closingNote: "The VWC team has been disciplined enough to say no to most deals for 12 years. We will be equally disciplined in selecting our capital partners.",
  },

  {
    id: "P2",
    slug: "02-hnw-cre-veteran",
    name: "HNW CRE Veteran (Tired Landlord)",
    audience: "For Real Estate Owners Who Are Done Being Landlords",
    salutation: "For the Real Estate Owner Considering a Step Out of Active Management",
    openingHook: "You already know multifamily real estate works. The question is whether you still want to take 3 a.m. calls to make it work.",
    whyWriting:
      "VisionWise Capital is a private real estate investment firm based in Mission Viejo, California. For 12 years we have done what you have done — buy under-managed multifamily property in coastal Southern California, improve it, hold or sell it for capital appreciation. We offer accredited investors the same returns without the operating burden, through single-asset LLCs you can passively own.",
    whatWeWantFromYou: "We are looking for one conversation about whether the next step in your real estate life is being a partner instead of an operator. Our Sidecar single-asset LLC product has a $500,000 minimum, an 18-to-24 month cycle, and is 1031-eligible.",
    whyMattersToYou: [
      "Same asset class, same returns. Our 24-deal portfolio averages a 23.95% levered project return. The deals are exactly the kind of property you already know — 5- to 75-unit coastal SoCal multifamily.",
      "1031-eligible. If you are exiting a property, our Sidecar single-asset LLC qualifies as like-kind real property. Your 1031 chain continues; your tax deferral continues.",
      "No 3 a.m. calls. We carry 56 staff, including a 30-year construction manager and roughly 35 in-house tradesmen. The work that used to be yours is now ours.",
      "Conservative leverage. Most VWC deals close at 50% LTV. We sized our portfolio to survive cycles, not just to ride them.",
    ],
    productMechanicsTitle: "What the Sidecar single-asset LLC looks like",
    productMechanics: [
      ["Minimum", "$500,000"],
      ["Hold period", "18–24 months per property"],
      ["Tax structure", "Single-asset LLC — 1031-eligible like-kind real property"],
      ["Geography", "Coastal Orange County and San Diego County submarkets"],
      ["Buy box", "5–75 units, vintage 1940–2000"],
      ["Leverage", "≤50% LTV after month 6 (conservative by design)"],
      ["Reporting", "AppFolio Investor portal + quarterly written update from Sanford Coggins"],
    ],
    caseStudies: ["charlene4921", "cypress374"],
    nextStep: "We would welcome a 30-minute phone call. If you are exiting a current property and the 45-day identification clock matters, please tell us in your reply — we will move at the speed your QI needs.",
    closingNote: "Owning multifamily real estate is one of the best wealth-preservation tools we know of. Owning it actively is one of the most demanding ways to spend your weekends. We can help with the second half of that sentence.",
  },

  {
    id: "P3",
    slug: "03-1031-exchangor",
    name: "1031 Exchangor",
    audience: "For Investors Working Against a 1031 Identification Clock",
    salutation: "For the 1031 Exchangor and Their Qualified Intermediary",
    openingHook: "If your 45-day identification window is open, we have current or near-term sidecar deals you can identify today and close on schedule.",
    whyWriting:
      "VisionWise Capital structures all of our current acquisitions as single-asset LLCs holding like-kind real property. That makes them eligible replacement properties for a 1031 exchange. We have closed 24 deals over 12 years — every one on schedule, never causing a buyer or seller to miss a tax deadline.",
    whatWeWantFromYou: "If you have a 1031 identification clock running and a $500,000-plus exchange to complete, we want to share our current Sidecar pipeline so your QI can structure documentation. We also assist investors who do not yet have a Qualified Intermediary relationship.",
    whyMattersToYou: [
      "Closing certainty. 24 deals closed over 12 years. We do not move closing dates and we do not lose deals at escrow.",
      "Like-kind eligibility. Single-asset LLC, like-kind real property. Your QI can confirm with us.",
      "Multiple deals in pipeline. We maintain two to three sidecar deals in active or near-term close — useful for backup identification.",
      "We help if you don't have a QI. VWC Exchange Services connects investors to vetted accommodators when needed.",
    ],
    productMechanicsTitle: "How a 1031 exchange into a VWC Sidecar works",
    productMechanics: [
      ["Replacement property type", "Single-asset LLC, like-kind real property"],
      ["Minimum", "$500,000"],
      ["Hold period", "18–24 months per property (sidecar cycle)"],
      ["Geography", "Coastal Orange County and San Diego County"],
      ["Closing window", "Always operating on a 45/180-day-aware calendar"],
      ["QI coordination", "We coordinate directly with IPX1031, Asset Preservation, FAI, Madison 1031, Accruit, and others"],
      ["Reporting", "AppFolio Investor portal + quarterly written update"],
    ],
    caseStudies: ["charlene4921", "encinitas"],
    nextStep: "If your clock is running, please call (949) 441-5580 or reply with your QI's contact information. We will share the active offering summary within one business day.",
    closingNote: "We treat every 1031 exchange as a calendar problem before it is an investment decision. The investment thesis is straightforward; the calendar requires a sponsor who has run this play 24 times.",
  },

  {
    id: "P4",
    slug: "04-income-retiree",
    name: "Income-Seeking Accredited Retiree",
    audience: "For Accredited Investors Who Want a Quarterly Check, Not a Quarterly Headache",
    salutation: "For the Accredited Investor Building Income, Not Speculation",
    openingHook: "You have done the building. Your next 20 years should pay you for it — without the volatility of public markets.",
    whyWriting:
      "VisionWise Capital is a 12-year-old private real estate investment firm headquartered in Mission Viejo, California. We acquire and improve coastal Southern California multifamily property and pass tax-advantaged income and capital appreciation to accredited investors through single-asset LLCs. Our goal — and our investors' goal — is conservative, durable income from a real asset you can drive past.",
    whatWeWantFromYou: "We are inviting accredited investors to participate in our current and upcoming Direct-deal offerings, with a $100,000 minimum. If you have invested with us before, this letter is partly a portfolio update; if you have not, this is an introduction.",
    whyMattersToYou: [
      "Real asset, real income. Multifamily housing pays a coupon while it appreciates. Our portfolio has averaged a 23.95% levered project return over 24 deals.",
      "Not correlated to the stock market. Public REITs trade with stocks; direct-owned multifamily does not. The cycle here is real estate, not the daily news.",
      "Tax-advantaged. Depreciation pass-through and step-up basis at death make multifamily one of the most efficient legacy-building assets available.",
      "Conservative leverage. We finance most deals at 50% LTV. We will not be the operator forced to sell into a weak market.",
    ],
    productMechanicsTitle: "What a Direct-deal investment with VWC looks like",
    productMechanics: [
      ["Minimum", "$100,000 (accredited only)"],
      ["Hold period", "24–30 months"],
      ["Distributions", "Targeted quarterly — confirmed in offering documents"],
      ["Tax treatment", "K-1 with depreciation pass-through; step-up basis at death"],
      ["Geography", "Coastal Orange County and San Diego County"],
      ["Leverage", "≤50% LTV after month 6"],
      ["Reporting", "AppFolio Investor portal + quarterly written update from Sanford Coggins"],
    ],
    caseStudies: ["charlene4921", "cypress374"],
    nextStep: "Please reply to this letter or call (949) 441-5580 to receive the active offering summary. If you would like to meet in person in Orange County, we are happy to arrange that.",
    closingNote: "The best compounders in real estate are not the boldest. They are the ones who set their leverage cap before the cycle turns. That is the kind of operator we have tried to be for 12 years.",
  },

  {
    id: "P5",
    slug: "05-sdira",
    name: "Self-Directed IRA Investor",
    audience: "For Accredited Investors With Self-Directed IRA Capital",
    salutation: "For the Self-Directed IRA Investor",
    openingHook: "If your IRA cash is sitting at Equity Trust, STRATA, Quest, or Madison earning effectively nothing, we accept SDIRA money on every direct-deal we close.",
    whyWriting:
      "VisionWise Capital's Direct-deal multifamily offerings accept self-directed IRA capital through every major SDIRA custodian. We have a subscription packet calibrated for SDIRA paperwork, our minimum ($100,000) fits typical IRA balances, and our 24–30 month hold period aligns with how IRA investors think about time horizons.",
    whatWeWantFromYou: "If you hold meaningful cash in an SDIRA and are looking to deploy it into a real asset, we want a 30-minute phone call. We will walk you and your custodian through the subscription process step by step.",
    whyMattersToYou: [
      "We accept all major SDIRA custodians. Equity Trust, STRATA, Quest, Madison, Directed IRA, Entrust, Pacific Premier, Kingdom Trust.",
      "Subscription packet ready for IRA paperwork. Most sponsors hand you a packet and wish you luck. We move with your custodian directly.",
      "$100,000 minimum fits a typical IRA balance. You don't need to roll multiple accounts to participate.",
      "We disclose the UBIT/UDFI math up front. On the leveraged portion of any deal, IRA capital has tax considerations. We provide the calculation; your tax advisor confirms.",
    ],
    productMechanicsTitle: "What a Direct-deal investment from your SDIRA looks like",
    productMechanics: [
      ["Minimum", "$100,000 (accredited only)"],
      ["Hold period", "24–30 months"],
      ["Distributions", "Returned to your IRA account, tax-deferred"],
      ["Custodian compatibility", "All major SDIRA custodians supported"],
      ["UBIT/UDFI", "Calculation provided up front for the leveraged portion"],
      ["Geography", "Coastal Orange County and San Diego County"],
      ["Reporting", "AppFolio Investor portal + quarterly written update"],
    ],
    caseStudies: ["cypress374", "charlene4921"],
    nextStep: "Please reply with the name of your custodian or call (949) 441-5580. We can have your subscription packet to you within one business day, calibrated for your custodian's process.",
    closingNote: "Self-directed IRA capital deserves a sponsor that treats the paperwork as part of the deal, not as an afterthought. We have done this enough times to make it routine.",
  },

  {
    id: "P6",
    slug: "06-ria-advisor",
    name: "RIA / Wealth Advisor",
    audience: "For Registered Investment Advisors Considering Real Estate Alternatives",
    salutation: "For the Registered Investment Advisor and Investment Committee",
    openingHook: "Sanford Coggins spent 16 years as a Vice President of Wealth Management at Merrill Lynch. He knows what a clean placement looks like from your side of the desk. That is how VisionWise Capital is structured.",
    whyWriting:
      "VisionWise Capital is the rare real estate sponsor founded by a former Registered Investment Advisor. Our compensation is fully disclosed (GP carry plus management fee, no retros to advisors). Our reporting is institutional-grade through AppFolio Investor. Our track record is 12 years, 24 deals, and 23.95% average levered project return — including a -33% paper loss we disclose in every conversation.",
    whatWeWantFromYou: "We are looking to brief two or three RIA investment committees in the next 90 days. Our preferred entry point is an educational session, not a pitch. If your IC is exploring multifamily real estate as an alternative-asset allocation for accredited clients, we would like to be on the agenda.",
    whyMattersToYou: [
      "RIA-friendly compensation. GP carry plus management fee, fully disclosed. No retros, no hidden advisor compensation that creates compliance noise.",
      "Reporting your clients can read. AppFolio Investor portal generates client-friendly statements aligned with your quarterly review cycle.",
      "Operator quality. We discuss our worst deal (-33% paper loss on 630 W Washington, San Diego) on the first call. Your IC will see how we underwrite, not just how we sell.",
      "Founded by an RIA. Sanford Coggins is one of few professionals who has transitioned from RIA to RE sponsor. He speaks your language.",
    ],
    productMechanicsTitle: "How RIA placements with VWC work",
    productMechanics: [
      ["Products", "Direct-deal LLCs ($100K min) and Sidecar single-asset LLCs ($500K min)"],
      ["Geography", "Coastal Orange County and San Diego County multifamily"],
      ["Hold period", "Direct deal 24–30 months; Sidecar 18–24 months"],
      ["Compensation", "GP carry + management fee, fully disclosed; no retros"],
      ["Reporting", "AppFolio Investor portal — client statements available"],
      ["Underwriting", "≤50% LTV after month 6; 5–75 unit sub-institutional buy box"],
      ["Track record", "24 deals, 227 units, 23.95% avg levered project return"],
    ],
    caseStudies: ["washington", "encinitas"],
    nextStep: "Please reply to schedule a 45-minute IC briefing in person or via video. We will bring the deal-level deck, the worst-deal post-mortem, and a sample reporting package.",
    closingNote: "An RIA franchise lives on quality of placement, not volume. We take that seriously because we used to be one of you.",
  },

  {
    id: "P7",
    slug: "07-tech-medical-pro",
    name: "Tech / Medical High-Income Wealth Builder",
    audience: "For High-Income Professionals Building Tax-Advantaged Wealth",
    salutation: "For the Surgeon, Engineer, Attorney, or Executive Earning Past the Tax-Code Cliff",
    openingHook: "You are paid in full-tax W-2 dollars or partnership income. You have RSU concentration risk. You don't want to be a landlord. We solve all three.",
    whyWriting:
      "VisionWise Capital's Direct-deal multifamily offerings deliver depreciation pass-through to accredited investors at a $100,000 minimum. The structure is built for high-income professionals who need passive real estate exposure: the K-1 arrives, depreciation hits other passive income, and the 24- to 30-month hold matches the rhythm of your portfolio review, not your week.",
    whatWeWantFromYou: "If you are earning past the threshold where tax efficiency becomes the binding constraint, we want a 30-minute phone call. Bring your CPA. We will walk through the depreciation math, the K-1 mechanics, and which of our current deals fits your tax position.",
    whyMattersToYou: [
      "Depreciation pass-through. Multifamily real estate generates non-cash depreciation that flows through your K-1. It does not offset W-2 income unless you qualify as a Real Estate Professional, but it offsets other passive income, carries forward, and reduces gain on sale.",
      "Truly passive. We carry 56 staff. You file the K-1 and answer two emails a year.",
      "$100,000 minimum. Designed for the wealth-builder, not the family office. Direct-deal participation across multiple offerings lets you stage capital deployment.",
      "No RSU correlation. Coastal SoCal multifamily is uncorrelated to your employer's stock or the public tech market.",
    ],
    productMechanicsTitle: "What a Direct-deal investment looks like for a high-income professional",
    productMechanics: [
      ["Minimum", "$100,000 (accredited only)"],
      ["Hold period", "24–30 months"],
      ["Tax mechanics", "K-1 with depreciation pass-through; step-up basis at death; 1031 chain on exit"],
      ["Geography", "Coastal Orange County and San Diego County"],
      ["Leverage", "≤50% LTV after month 6"],
      ["Time required from you", "Sign subscription docs once; receive K-1 once a year"],
      ["Reporting", "AppFolio Investor portal + quarterly written update"],
    ],
    caseStudies: ["charlene4921", "encinitas"],
    nextStep: "Please reply or call (949) 441-5580. If you would like to bring your CPA on the first call, we welcome it — most of our investors do.",
    closingNote: "The best returns in real estate go to the operator. The best after-tax returns often go to the limited partner. Our job is to make sure those are the same dollars.",
  },

  {
    id: "P8",
    slug: "08-recently-liquid",
    name: "Recently Liquid Investor",
    audience: "For Investors With New Capital From a Liquidity Event",
    salutation: "For the Investor Who Just Sold the Business, Inherited the Estate, or Vested the RSUs",
    openingHook: "The hardest part of a liquidity event is not getting the money. It is deploying it without giving back the gain.",
    whyWriting:
      "VisionWise Capital is a 12-year-old private real estate investment firm in coastal Southern California. We work with investors who have meaningful new capital and a strong preference for not losing it. Our typical first-time investor brings their CPA and attorney to the second meeting. We welcome that. Our subscription documents look like documents you would receive from an institutional partner.",
    whatWeWantFromYou: "If you have recently received a meaningful liquidity distribution and are evaluating where to deploy it, we want a 60-minute first meeting with you and your advisors. We will walk through our underwriting framework, our worst deal, and our current pipeline. No pitch.",
    whyMattersToYou: [
      "Capital preservation first. We finance most deals at 50% LTV. The portfolio has weathered the 2022–2024 rate environment without forced sales.",
      "Bring your advisors. CPAs and estate attorneys are part of every meeting we want to have. Their questions sharpen our answers.",
      "We don't take all your money. The right amount of any one deal is rarely the largest amount. We will help you size sensibly.",
      "We tell you about the loss first. 630 W Washington, San Diego, -33% paper loss, currently held. We discuss it before we discuss the wins.",
    ],
    productMechanicsTitle: "How a first investment with VWC works",
    productMechanics: [
      ["Entry products", "Direct-deal LLC ($100K min) for accredited investors; Sidecar ($500K min) is 1031-eligible"],
      ["Hold period", "Direct deal 24–30 months; Sidecar 18–24 months"],
      ["Geography", "Coastal Orange County and San Diego County"],
      ["Leverage", "≤50% LTV after month 6"],
      ["Tax structure", "K-1 with depreciation pass-through; step-up basis at death"],
      ["Recommended bucket", "Allocate after reserving 12–18 months of living expenses outside this investment"],
      ["Reporting", "AppFolio Investor portal + quarterly written update from Sanford Coggins"],
    ],
    caseStudies: ["cypress374", "charlene4921"],
    nextStep: "Please reply to this letter or have your CPA or attorney introduce you. We are happy to send our worst-deal post-mortem and our subscription documents in advance — many investors prefer to read them before the meeting.",
    closingNote: "The most expensive thing a recently-liquid investor can do is move fast. We are happy to move at the pace your advisors set.",
  },

  {
    id: "P9",
    slug: "09-faith-based",
    name: "Faith-Based / Values-Aligned Investor",
    audience: "For Faithful Stewards Building Long-Term Real Estate Capital",
    salutation: "For the Christian Family, Foundation, or Business Owner Stewarding Real Estate Capital",
    openingHook: "Our company's stated value reads: \"Faith — a vibrant relationship with God as a follower of Jesus … all we have belongs to God; we are stewards of His resources.\" We don't lead with this in cold outreach. We don't pretend it isn't true, either.",
    whyWriting:
      "VisionWise Capital is a private real estate investment firm in Mission Viejo, California, founded by Sanford D. Coggins. Sanford spent 16 years at Merrill Lynch, where he co-led a $77 million pilot program that grew into a $550 million California-wide initiative serving ethnically diverse communities. He earned Merrill's Lifetime Achievement Award. We mention this not as résumé, but as character — the same ethic shapes how we treat investor capital and tenants today.",
    whatWeWantFromYou: "We are not making a cold pitch into faith communities. We are inviting introductions through trusted relationships — National Christian Foundation, Halftime Institute, C12, Pinnacle Forum, and the local Christian business networks where character-based diligence still happens.",
    whyMattersToYou: [
      "Stewardship is operational, not marketing. \"Committed to investors over projects\" is one of our two core 'we have' values. It shows up in conservative leverage, transparent loss disclosure, and patient acquisition.",
      "Honest about losses. The 630 W Washington deal in our portfolio is at a -33% paper loss. It is on the first page of every conversation. \"Earn future trust by accurately reporting past facts\" is our value — operationally.",
      "Multi-generational thinking. Hold periods and exit decisions optimize for compound returns to investors, not transaction volume to the sponsor.",
      "Subtle but unashamed. We don't put scripture on the cover of an offering memo. We won't sanitize our values out of who we are, either.",
    ],
    productMechanicsTitle: "What an investment with VWC looks like",
    productMechanics: [
      ["Entry products", "Direct-deal LLC ($100K min, accredited); Sidecar single-asset LLC ($500K min, 1031-eligible)"],
      ["Hold period", "Direct deal 24–30 months; Sidecar 18–24 months"],
      ["Geography", "Coastal Orange County and San Diego County multifamily"],
      ["Leverage", "≤50% LTV after month 6"],
      ["Reporting", "AppFolio Investor portal + quarterly written update from Sanford Coggins"],
      ["Track record", "24 deals, 227 units, 23.95% avg levered project return"],
      ["Loss disclosure", "1 paper-loss deal of 24, fully disclosed"],
    ],
    caseStudies: ["cypress374", "washington"],
    nextStep: "If we share trusted relationships in NCF, Halftime, C12, or a local Christian business network, please ask for a warm introduction rather than replying to this letter directly. If we don't, please reply.",
    closingNote: "We measure ourselves by character first, competency second. Both are required. Neither alone is enough.",
  },

  {
    id: "P10",
    slug: "10-co-sponsor",
    name: "Co-Sponsor / Operator's Personal Capital",
    audience: "For Real Estate Operators Diversifying Personal Capital",
    salutation: "For the GP, Syndicator, or Fix-and-Flip Operator Looking for an LP Position",
    openingHook: "We are not pitching another operator's deals to you. We are inviting you to put a small amount of your own capital with another operator in a market you don't compete in.",
    whyWriting:
      "VisionWise Capital is a 12-year, 24-deal coastal Southern California multifamily fix-and-flip operator. Our underwriting is rigorous, our leverage is conservative, our team is in-house, and our track record is transparent — including the deal that didn't work. We are exactly the kind of sponsor an operator looks at when choosing a passive position outside their home market.",
    whatWeWantFromYou: "If you operate in Texas, Florida, the Midwest, or the Pacific Northwest, you are precisely the operator we want as an LP. We accept $100,000-minimum Direct-deal LLC participation and $500,000-minimum Sidecar single-asset LLC participation. We will share our underwriting model, our renovation scope-of-work template, and our worst-deal post-mortem on the first call.",
    whyMattersToYou: [
      "Geographic diversification you cannot operate yourself. Coastal SoCal multifamily is supply-constrained at any cycle. You can underwrite us; you cannot replace us locally.",
      "≤50% LTV after month 6. The leverage discipline you've heard about — confirmed in our worst deal, 630 W Washington, where the conservative LTV gave us optionality to hold through a hostile rate environment.",
      "Vertically integrated 56-person team. Full-time construction manager, ~35 in-house tradesmen. You'll recognize what that does for renovation cost certainty.",
      "We respect your time on the first call. Underwriting model, scope-of-work template, deal-level returns, the loss. No story-time.",
    ],
    productMechanicsTitle: "What an operator-LP position looks like",
    productMechanics: [
      ["Entry products", "Direct-deal LLC ($100K min) and Sidecar single-asset LLC ($500K min)"],
      ["Hold period", "Direct deal 24–30 months; Sidecar 18–24 months"],
      ["Geography", "Coastal Orange County and San Diego County (i.e., not your market)"],
      ["Leverage", "≤50% LTV after month 6"],
      ["Underwriting access", "Sample model and renovation SOW shared on first call"],
      ["Reporting", "AppFolio Investor portal + quarterly written update"],
      ["Track record", "24 deals, 227 units, 23.95% avg levered project return"],
    ],
    caseStudies: ["cypress374", "washington"],
    nextStep: "Please reply with two or three specific underwriting questions. We will answer them in writing, then schedule a 30-minute call. Operators do better diligence than retail LPs, and we appreciate it.",
    closingNote: "The best LP feedback we get comes from other operators. Pitch is fastest when it doesn't have to happen.",
  },
];

// ============================================================
// DOCUMENT BUILDER (one document per persona)
// ============================================================
function buildPersonaOverview(persona) {
  // ---- COVER ----
  const cover = [
    ...(logoBuffer ? [
      new Paragraph({
        spacing: { before: 1600, after: 200 },
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({
          type: "png",
          data: logoBuffer,
          transformation: { width: 280, height: 64 },
          altText: { title: "VisionWise Capital", description: "VisionWise Capital logo", name: "VWC Logo" },
        })],
      }),
    ] : [
      new Paragraph({
        spacing: { before: 1800, after: 240 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "VisionWise Capital", font: HEADING_FONT, size: 64, bold: true, color: COLOR.charcoal })],
      }),
    ]),

    new Paragraph({
      spacing: { after: 800 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Multifamily Real Estate Investments", font: BODY_FONT, size: 24, italics: true, color: COLOR.mid })],
    }),

    new Paragraph({
      spacing: { before: 200, after: 800 },
      alignment: AlignmentType.CENTER,
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: COLOR.green, space: 1 } },
      children: [new TextRun({ text: "" })],
    }),

    new Paragraph({
      spacing: { after: 240 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "A company overview", font: HEADING_FONT, size: 56, bold: true, color: COLOR.charcoal })],
    }),

    new Paragraph({
      spacing: { after: 1200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: persona.audience,
        font: BODY_FONT, size: 26, color: COLOR.charcoalTint,
      })],
    }),

    new Paragraph({
      spacing: { before: 800, after: 60 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "VisionWise Capital LLC  ·  Mission Viejo, California",
        font: BODY_FONT, size: 22, color: COLOR.charcoal,
      })],
    }),

    new Paragraph({
      spacing: { after: 60 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "(949) 441-5580  ·  info@visionwisecapital.com  ·  visionwisecapital.com",
        font: BODY_FONT, size: 20, color: COLOR.mid,
      })],
    }),
  ];

  // ---- LETTER ----
  const letter = [
    h1(persona.salutation),

    callout("Why we are writing", persona.openingHook),
    spacer(),

    p(persona.whyWriting),
    spacer(80),
    p(persona.whatWeWantFromYou),

    h2("Why VisionWise Capital matters to you"),
    ...persona.whyMattersToYou.map(t => bullet(t)),

    h2("How VisionWise Capital operates"),
    p("We acquire under-managed multifamily property in coastal Southern California — five to seventy-five units, vintage 1940 to 2000 — and improve it to market standard. The playbook is three steps:"),
    bullet("Acquire patiently. We track properties for months at our price, not the seller's. The completeness of our buyer package wins deals where the highest offer does not."),
    bullet("Modernize in-house. We employ a full-time construction manager with 30 years of experience and roughly 35 in-house tradesmen. The work is not subcontracted to whoever returns the call."),
    bullet("Move rents to market. Eighteen to twenty-four months later, the property is at market rents and operating at stabilized cap rates. We hold for income or sell for capital appreciation."),
    p("Across 24 deals over 12 years, the portfolio has produced 23.95% average levered project returns — including a -33% paper-loss deal we will discuss before any other deal."),
  ];

  // ---- TRACK RECORD HEADLINE ----
  const trackRecord = [
    h2("Track record at a glance"),

    buildHeaderTable(
      ["Metric", "Value"],
      [
        ["Years operating", "12"],
        ["Total transactions", "24"],
        ["Total units acquired", "227"],
        ["Total cost basis", "$103.0M"],
        ["Total equity invested by LPs", "$53.3M"],
        ["Realized + estimated profit", "$12.8M"],
        ["Average levered project return", "23.95%"],
        ["Loss disclosure", "1 paper-loss deal of 24, fully transparent"],
      ],
      [3500, 6580]
    ),
  ];

  // ---- CASE STUDIES ----
  const caseStudyBlocks = persona.caseStudies.flatMap((csKey, i) => {
    const cs = CASE_STUDIES[csKey];
    return [
      h2(i === 0 ? "Case study one" : "Case study two"),
      h3(cs.title),
      p(cs.status, { color: COLOR.mid, italic: true, size: 20 }),
      callout("Headline", cs.headlineReturn),
      spacer(80),
      p(cs.summary),
      h3("What this deal teaches"),
      p(cs.lesson),
      h3("The numbers"),
      buildKVTable(cs.numbers, 3000),
      spacer(),
    ];
  });

  // ---- WHY THIS MATTERS / PRODUCT / NEXT STEP ----
  const closing = [
    h2(persona.productMechanicsTitle),
    buildKVTable(persona.productMechanics, 3000),
    spacer(),

    h2("Our values, briefly"),
    p("VisionWise Capital is governed by two pairs of values. The first pair — Entrepreneurial, Strategic Thinker, Resolute — describes how we work. The second pair — Faith and Clear Conscience — describes the standard we hold ourselves to."),
    quote("Committed to investors over projects."),
    p("That is not a tagline. It is one of the four enumerated commitments under our Clear Conscience value, and it is what produced the leverage discipline that protected the portfolio through 2022–2024.", { italic: true, color: COLOR.charcoalTint }),

    h2("Next step"),
    p(persona.nextStep),
    spacer(),

    p(persona.closingNote, { italic: true, color: COLOR.charcoalTint }),
    spacer(),
    spacer(),

    pRich([{ text: "With respect,", italic: true }]),
    spacer(80),
    pRich([
      { text: "Sanford D. Coggins", bold: true, font: HEADING_FONT, size: 26 },
    ]),
    pRich([
      { text: "CEO & Founder, VisionWise Capital LLC", italic: true, color: COLOR.mid, size: 20 },
    ]),
  ];

  // ---- DISCLAIMER ----
  const DISCLAIMER = "This communication is provided for informational purposes only and does not constitute an offer to sell or a solicitation of an offer to buy any securities. Any such offer or solicitation will be made only through definitive offering documents (including a private placement memorandum) provided to qualified investors. Past performance is not indicative of future results. Real estate investments involve risk, including loss of principal.";

  const disclaimer = [
    h1("Important disclosures"),
    p(DISCLAIMER, { size: 18, color: COLOR.charcoalTint }),
    spacer(),
    p("This document is provided as an introduction to VisionWise Capital LLC. Specific deal economics, expected returns, and tax treatment for any particular offering will be detailed in the offering's Private Placement Memorandum and subscription documents.", { size: 18, italic: true, color: COLOR.mid }),
  ];

  // ---- DOCUMENT ASSEMBLY ----
  return new Document({
    creator: "VisionWise Capital",
    title: `VWC Company Overview — ${persona.name}`,
    description: `Persona-tailored company overview for ${persona.audience}`,
    styles: {
      default: { document: { run: { font: BODY_FONT, size: 22, color: COLOR.charcoal } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 48, bold: true, font: HEADING_FONT, color: COLOR.charcoal },
          paragraph: { spacing: { before: 280, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: HEADING_FONT, color: COLOR.charcoal },
          paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: HEADING_FONT, color: COLOR.greenDark },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [
        { reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 540, hanging: 270 } }, run: { color: COLOR.greenDark, font: BODY_FONT } } }] },
        { reference: "numbers",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 540, hanging: 270 } }, run: { color: COLOR.greenDark, font: BODY_FONT, bold: true } } }] },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: `VisionWise Capital  ·  Company Overview  ·  ${persona.audience}`,
              font: BODY_FONT, size: 16, color: COLOR.mid, italics: true,
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 60 },
              border: { top: { style: BorderStyle.SINGLE, size: 6, color: COLOR.green, space: 4 } },
              children: [new TextRun({
                text: "VisionWise Capital LLC  ·  visionwisecapital.com  ·  (949) 441-5580",
                font: BODY_FONT, size: 16, color: COLOR.mid,
              })],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Page ", font: BODY_FONT, size: 16, color: COLOR.mid }),
                new TextRun({ children: [PageNumber.CURRENT], font: BODY_FONT, size: 16, color: COLOR.mid }),
                new TextRun({ text: "  of  ", font: BODY_FONT, size: 16, color: COLOR.mid }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: BODY_FONT, size: 16, color: COLOR.mid }),
              ],
            }),
          ],
        }),
      },
      children: [
        ...cover,
        ...letter,
        ...trackRecord,
        ...caseStudyBlocks,
        ...closing,
        ...disclaimer,
      ],
    }],
  });
}

// ============================================================
// BUILD ALL 10
// ============================================================
async function buildAll() {
  for (const persona of personas) {
    const doc = buildPersonaOverview(persona);
    const filename = `VWC Company Overview - ${persona.slug} - ${persona.name.replace(/[\/\\:*?"<>|()]/g, "")}.docx`;
    const outPath = path.join(__dirname, filename);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buffer);
    console.log("Wrote: " + filename);
  }
}

buildAll().catch(err => { console.error(err); process.exit(1); });
