// Builds: VWC Investor Personas - 2026-04-25.docx
// Run: NODE_PATH="<global npm node_modules>" node build-personas-docx.js
//
// Brand-aligned per company-intel/brand/brand-guide.md:
//  - Colors: VWC Green #A3D55D, Charcoal #40464B, Green Dark #5C9627, etc.
//  - Fonts: Lato (heading), Open Sans (body) — fall back to system on render
//  - Type scale: H1 28pt, H2 18pt, H3 13pt, body 11pt, caption 9pt
//  - Sentence case for headers (never ALL CAPS except brand mark)
//  - H2 with green bottom border
//  - Cover logo + Reg D disclaimer footer

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, ShadingType, PageNumber, PageBreak, VerticalAlign,
} = require("docx");

// ============================================================
// CONSTANTS — BRAND-ALIGNED
// ============================================================
const PAGE_WIDTH = 12240;
const PAGE_HEIGHT = 15840;
const MARGIN = 1080; // 0.75"
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2; // 10080

// Brand palette (from company-intel/brand/brand-guide.md)
const COLOR = {
  green: "A3D55D",       // VWC Green — fills, dividers, headlines on charcoal
  greenDark: "5C9627",   // Links, H3 headings, accents on white
  greenTint: "DAF0BC",   // Callout backgrounds, chart fills
  charcoal: "40464B",    // Primary text, H1/H2 headings (NEVER pure black)
  charcoalTint: "666B70",// Secondary text on white
  mid: "8A8F94",         // Captions, metadata
  light: "E5E7EA",       // Borders, table grid lines
  offWhite: "F7F7F7",    // Section backgrounds, table row stripes
  white: "FFFFFF",
};

// Brand fonts (web fallback per brand guide)
const HEADING_FONT = "Lato";
const BODY_FONT = "Open Sans";

const border = (color = COLOR.light, size = 4) =>
  ({ style: BorderStyle.SINGLE, size, color });

const allBorders = (color = COLOR.light, size = 4) => ({
  top: border(color, size), bottom: border(color, size),
  left: border(color, size), right: border(color, size),
});

const noBorders = {
  top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL },
  left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL },
};

const cellMargins = { top: 100, bottom: 100, left: 140, right: 140 };

// ============================================================
// PARAGRAPH HELPERS
// ============================================================
const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: opts.spacing || { after: 120 },
  children: [new TextRun({
    text,
    font: opts.font || BODY_FONT,
    size: opts.size || 22, // 11pt
    bold: !!opts.bold,
    italics: !!opts.italic,
    color: opts.color || COLOR.charcoal,
  })],
  ...opts.paragraphOpts,
});

const pRich = (runs, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: opts.spacing || { after: 120 },
  children: runs.map(r => new TextRun({
    font: r.font || BODY_FONT,
    size: r.size || 22,
    text: r.text,
    bold: !!r.bold,
    italics: !!r.italic,
    color: r.color || COLOR.charcoal,
  })),
});

// H1: 28pt bold Charcoal — page break before
const h1 = (text) => new Paragraph({
  spacing: { before: 320, after: 200 },
  pageBreakBefore: true,
  children: [new TextRun({
    text, font: HEADING_FONT, size: 56, bold: true, color: COLOR.charcoal,
  })],
});

const h1NoBreak = (text) => new Paragraph({
  spacing: { before: 320, after: 200 },
  children: [new TextRun({
    text, font: HEADING_FONT, size: 56, bold: true, color: COLOR.charcoal,
  })],
});

// H2: 18pt bold Charcoal with green bottom border
const h2 = (text) => new Paragraph({
  spacing: { before: 280, after: 180 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR.green, space: 4 } },
  children: [new TextRun({
    text, font: HEADING_FONT, size: 36, bold: true, color: COLOR.charcoal,
  })],
});

// H3: 13pt bold Green Dark
const h3 = (text) => new Paragraph({
  spacing: { before: 180, after: 80 },
  children: [new TextRun({
    text, font: HEADING_FONT, size: 26, bold: true, color: COLOR.greenDark,
  })],
});

const bullet = (text, opts = {}) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60 },
  children: [new TextRun({
    text, font: BODY_FONT, size: 22, color: COLOR.charcoal,
    bold: !!opts.bold,
  })],
});

const numbered = (text) => new Paragraph({
  numbering: { reference: "numbers", level: 0 },
  spacing: { after: 60 },
  children: [new TextRun({ text, font: BODY_FONT, size: 22, color: COLOR.charcoal })],
});

const caption = (text) => p(text, {
  size: 18, italic: true, color: COLOR.mid, font: BODY_FONT,
});

// ============================================================
// TABLE HELPERS
// ============================================================
const cell = (content, opts = {}) => new TableCell({
  borders: opts.noBorder ? noBorders : allBorders(opts.borderColor || COLOR.light, 4),
  width: { size: opts.width, type: WidthType.DXA },
  shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
  margins: cellMargins,
  verticalAlign: VerticalAlign.CENTER,
  children: Array.isArray(content) ? content : [content],
});

const textCell = (text, opts = {}) => cell(
  new Paragraph({
    spacing: { after: 0 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text,
      font: opts.font || BODY_FONT,
      size: opts.size || 20, // 10pt for table cells
      bold: !!opts.bold,
      italics: !!opts.italic,
      color: opts.color || COLOR.charcoal,
    })],
  }),
  opts
);

const buildKVTable = (rows, leftWidth = 3000) => {
  const rightWidth = CONTENT_WIDTH - leftWidth;
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [leftWidth, rightWidth],
    rows: rows.map(([k, v], i) => new TableRow({
      children: [
        textCell(k, { width: leftWidth, fill: COLOR.offWhite, bold: true, color: COLOR.charcoal }),
        textCell(v, { width: rightWidth, fill: i % 2 === 0 ? COLOR.white : COLOR.offWhite }),
      ],
    })),
  });
};

const buildHeaderTable = (headers, dataRows, columnWidths, headerFill = COLOR.charcoal, headerColor = COLOR.white) => {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => textCell(h, {
          width: columnWidths[i],
          fill: headerFill,
          color: headerColor,
          bold: true,
          size: 20,
          font: HEADING_FONT,
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

// Callout: Green Tint background, Green left border
const callout = (label, text) => {
  return new Table({
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
                  text: label, font: HEADING_FONT, size: 18, bold: true,
                  color: COLOR.charcoal,
                })],
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [new TextRun({
                  text, font: BODY_FONT, size: 22, italics: true, color: COLOR.charcoal,
                })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
};

const spacer = (size = 120) => new Paragraph({
  spacing: { after: size }, children: [new TextRun({ text: "" })],
});

// ============================================================
// PERSONA DATA
// ============================================================
const personas = [
  {
    id: "P1",
    name: "Family Office Capital Partner",
    tagline: "The $10M+ anchor relationship. Top strategic priority — owner-level commitment.",
    priority: 1,
    bestFit: "$30M revolving Line of Credit + anchor LP across single-asset LLCs",
    checkSize: "$10M – $20M+ per relationship (target: 2–3 Family Offices)",
    timeToCash: "3–9 months (relationship-led)",
    profile: [
      ["Entity type", "Single Family Office, Multi-Family Office, or institutional family-controlled trust"],
      ["AUM", "$50M – $2B+"],
      ["Principal age", "50–75; second-generation often 35–55"],
      ["Geography", "Southern California first; nationwide acceptable; Texas a wildcard via SDC's home network"],
      ["Investment committee", "1–5 members"],
      ["Existing alts allocation", "20%–60% of portfolio in alternatives"],
    ],
    psychographics: [
      "Seeks uncorrelated, tax-efficient returns",
      "Values relationship and transparency over flashy returns",
      "Long memory — one bad sponsor experience colors the next decade",
      "Cares deeply about alignment (sponsor co-invest, GP capital at risk)",
      "Often values-driven — many have a philanthropic mandate; some have explicit faith identity",
    ],
    jtbd: [
      "Deploy $10M–$50M of dry powder into an asset class that pays without daily news risk",
      "Find 2–3 trusted-sponsor relationships per asset class — concentrate, don't spread",
      "Maintain optionality (revolving credit beats one-shot equity)",
      "Receive conservative, predictable, tax-advantaged returns",
    ],
    painPoints: [
      "Too many syndicators pitching, too little operator-quality sponsorship",
      "Inconsistent quarterly reporting from sponsors",
      "Capital lockup conflicts with family liquidity events",
      "Scaled too fast with the wrong sponsor and got stuck",
    ],
    whereToFind: [
      "Family Office Exchange (FOX) — confirmed VWC LP-converting conference",
      "Buttonwood — confirmed converter",
      "Lido Family Office Symposium",
      "iConnections, Tiger 21, R360 (peer networks)",
      "Multi-Family-Office RIAs: Cresset, Aspiriant, Mariner, Hightower MFO, Pathstone, Geller Advisors",
      "Probate and business-sale trigger events (founder-liquidity moments)",
    ],
    objections: [
      ["Why not buy the assets ourselves?", "You'd need an in-house 56-person multifamily operating team. We have one. You get the operator without owning the operator."],
      ["What's your concentration risk?", "Single-asset LLCs, 5–75 units, ≤50% LTV after month 6, 18–24 month cycle. You underwrite one deal at a time."],
      ["Track record?", "227 units, 24 deals, 23.95% average levered project return — including a -33% loss we'll walk you through transparently."],
      ["Why a LOC and not equity?", "We refinance you out within 100 days. You earn a coupon, we get speed-to-close. Then we recycle."],
    ],
    hook: "$30M revolving line of credit, 90-day refinance cycle, backing a 12-year SoCal multifamily fix-and-flip operator with a 56-person vertically integrated team. We're inviting two or three Family Offices into this structure.",
    leadAssets: [
      "Track record one-pager (24 deals, with the honest -33% disclosure)",
      "Capital Raise structure brief (the $30M LOC mechanics)",
      "Encinitas Investment Summary as proof-of-product",
    ],
  },
  {
    id: "P2",
    name: 'HNW CRE Veteran ("Tired Landlord")',
    tagline: "Owned the duplex, sold it, doesn't want to be a landlord again — but loves the asset class.",
    priority: 4,
    bestFit: "Sidecar single-asset LLC ($500K minimum, 1031-eligible)",
    checkSize: "$500K – $2M per deal",
    timeToCash: "2–8 weeks",
    profile: [
      ["Age", "50–70"],
      ["Net worth", "$5M – $50M (excluding primary residence)"],
      ["Income", "$300K – $2M / year"],
      ["Profession", "Retired or semi-retired; often ex-business-owner or licensed professional"],
      ["Location", "SoCal-heavy; comfortable with OC / SD coastal"],
      ["Past behavior", "Owned 1–10 multifamily properties directly; has done 1031 exchanges"],
    ],
    psychographics: [
      "Tired of toilets, tenants, and trash — but still loves the asset class",
      "Wants passive exposure without the operating headache",
      "Values operator competence over slick marketing",
      "Tax-savvy — understands depreciation, 1031, cost segregation",
      "Skeptical by default — has seen sponsors fail",
    ],
    jtbd: [
      "Recycle existing real estate equity into passive vehicles",
      "Continue tax-deferral via 1031",
      "Stay in SoCal coastal (knows the market, won't leave it)",
      "Get back to family, hobbies, and health — escape the landlord trap",
    ],
    painPoints: [
      "3 a.m. calls, evictions, broker fees, plumber bills",
      "Hard to scale — each property eats time",
      "1031 pressure: 45 days to identify, 180 days to close",
      "Doesn't trust REITs (correlated to public markets, no tax shelter)",
    ],
    whereToFind: [
      "VWC's HubSpot CRM (existing 127-LP base)",
      "Marcus & Millichap, Berkadia, NMHC alumni networks (SDC's broker pedigree)",
      "1031 Qualified Intermediary referrals",
      "IMN multifamily forum, Bisnow events",
      "OC / SD apartment-owner direct mailing lists (VWC has scrubbed ~15,000)",
      "Local OC / SD CCIM, IREM, AAOC chapters",
    ],
    objections: [
      ["I'd rather own it myself.", "You've already done that. The 18-month sidecar gives you the upside without the toilet calls."],
      ["Why VWC vs. another sponsor?", "Vertically integrated 56-person team, in-house construction. We don't subcontract the value-add."],
      ["What's the IRR?", "23.95% average levered project return across 24 deals — including the loss. Per-deal numbers are in the offering summary."],
      ["Concentration risk in SoCal?", "Yes — that's the bet. You know coastal SoCal as well as we do. The sub-institutional 5–75 unit segment is where institutions can't compete."],
    ],
    hook: "You sold the 8-plex. You're done with 3 a.m. calls. But you're not done with the asset class. Here's what an 18-month sidecar in Encinitas looks like — same returns, none of the toilets.",
    leadAssets: [
      "Encinitas Investment Summary",
      "374 Cypress and 4921 Charlene case studies (the home-runs)",
      "1031 enablement one-pager",
    ],
  },
  {
    id: "P3",
    name: "1031 Exchangor",
    tagline: "They have 45 days. We have a deal closing now.",
    priority: 2,
    bestFit: "Sidecar single-asset LLC (1031-eligible)",
    checkSize: "$500K – $5M per exchange",
    timeToCash: "Days to weeks (forced by IRS deadlines)",
    profile: [
      ["Age", "45–80"],
      ["Net worth", "$2M – $50M"],
      ["Trigger", "Just sold or about to sell appreciated real estate"],
      ["Location", "National (IRS rules don't care)"],
      ["Past behavior", "Has done at least one prior 1031, or first-timer with strong CPA guidance"],
    ],
    psychographics: [
      "Solving a tax problem, not buying an investment",
      "Time-stressed — the IRS clock is the boss",
      "Values certainty of close above everything",
      "Often emotionally reluctant to sell (legacy property), looking for honorable redeployment",
    ],
    jtbd: [
      "Identify replacement property within 45 days",
      "Close on replacement within 180 days",
      "Avoid boot (taxable gain on un-reinvested cash)",
      "Avoid a sponsor blowing the close",
    ],
    painPoints: [
      "Deadline anxiety",
      "Few viable sponsors with deals closing on the right window",
      "Worry about non-1031-compliant structures masquerading as compliant",
    ],
    whereToFind: [
      "1031 Qualified Intermediaries: IPX1031, Asset Preservation, First American Exchange, FAI, Madison 1031, Accruit",
      "CPAs with HNW client books",
      "Estate attorneys",
      "VWC's own seller pipeline (sellers exiting via 1031)",
      "VWC Exchange Services capability — a hook for investors without an accommodator",
    ],
    objections: [
      ["Will we close in time?", "We've closed 24 deals on schedule. Here's our QI list and the closing timeline for the current sidecar."],
      ["Is this truly 1031-compliant?", "Single-asset LLC, like-kind real property, your ownership interest is real-property-eligible. Confirmed by our QI partners."],
      ["What if the deal falls through?", "We work with your QI on backup ID. We maintain 2–3 deals in pipeline for exactly this."],
    ],
    hook: "Need replacement property? We have a 21-unit Encinitas value-add closing in days. Single-asset LLC, 1031-eligible, $500K minimum. Here's the QI documentation.",
    leadAssets: [
      "Active offering summary (Encinitas, Knott)",
      "1031 process one-pager",
      "VWC Exchange Services brief",
    ],
  },
  {
    id: "P4",
    name: "Income-Seeking Accredited Retiree / Pre-Retiree",
    tagline: "They want the coupon, not the appreciation.",
    priority: 3,
    bestFit: "Direct deal ($100K minimum); legacy Fund IV holders recycling distributions",
    checkSize: "$100K – $1M per deal",
    timeToCash: "4–12 weeks",
    profile: [
      ["Age", "55–75"],
      ["Net worth", "$2M – $15M"],
      ["Income source", "Retirement income, pension, Social Security, prior business sale"],
      ["Profession", "Retired executive, doctor, attorney, business owner"],
      ["Location", "National, but SoCal-heavy in VWC's existing book"],
    ],
    psychographics: [
      "Income beats capital appreciation in life-stage priorities",
      "Capital preservation is non-negotiable",
      "Tired of public-market volatility",
      "Wants predictable distributions they can plan around",
      "Values stewardship language — they're thinking about legacy",
    ],
    jtbd: [
      "Generate $50K – $300K / year of tax-advantaged passive income",
      "Preserve principal for heirs",
      "Diversify away from public markets",
      "Stay aligned with personal values",
    ],
    painPoints: [
      "Rates are volatile; a bond ladder doesn't beat inflation",
      "Public REITs are correlated to stocks (defeats the purpose)",
      "Don't want to be active landlords at 65+",
      "Worry about sponsor solvency in a downturn",
    ],
    whereToFind: [
      "VWC's existing HubSpot list (~127 LPs, ~85% repeat — many here)",
      "Bisnow and IMN events",
      "RIA channel feeding income-seeking clients (P6)",
      "AppFolio Investor portal — current investors are the warm pool for additional placements",
    ],
    objections: [
      ["What's the yield?", "[Confirm specifics with SDC — pull from current offering]"],
      ["How illiquid is it?", "Single-asset LLC, 24–30 months. Plan around it — distributions begin in month [X]. This isn't an emergency-fund vehicle."],
      ["What if you go bust?", "Conservative leverage (≤50% LTV after month 6), 24-deal track record, in-house operating team. Here's the portfolio status snapshot."],
      ["What about my heirs?", "Step-up basis at death; the 1031 chain still works. Your CPA will love it."],
    ],
    hook: "For accredited investors who'd rather collect a quarterly check than watch CNBC. Direct-deal multifamily, tax-advantaged income, 12-year track record.",
    leadAssets: [
      "Quarterly distribution history (request from SDC)",
      '"Letter to Investors" blog (warm tone)',
      '"REITs vs. Direct Investing" blog (positioning piece)',
    ],
  },
  {
    id: "P5",
    name: "SDIRA Self-Directed Investor",
    tagline: "Their IRA is sitting in cash. We have the asset.",
    priority: 5,
    bestFit: "Direct deal ($100K minimum) — fits IRA economics",
    checkSize: "$100K – $500K per deal",
    timeToCash: "2–8 weeks once warmed via custodian webinar",
    profile: [
      ["Age", "35–70"],
      ["IRA balance", "$100K – $5M (often roll-overs from 401(k) at job changes)"],
      ["Custodian", "Equity Trust, STRATA, Quest, Madison, Directed IRA, Kingdom Trust"],
      ["Profession", "Wide range — engineers, dentists, pilots, ex-corporate executives"],
      ["Location", "National"],
    ],
    psychographics: [
      "Self-directed means wants control; hates Wall Street defaults",
      "Educated on alternatives — has read, attended seminars",
      "DIY mindset; values transparent docs",
      "Frustrated that IRA cash is doing nothing",
    ],
    jtbd: [
      "Move IRA cash into hard assets",
      "Keep the tax-deferred wrapper intact",
      "Find sponsors who accept SDIRA paperwork without friction",
      "Get clean K-1s and reporting back to the custodian",
    ],
    painPoints: [
      "Most sponsors don't accept SDIRA money (or charge friction)",
      "UBIT / UDFI confusion on leveraged deals",
      "Custodian paperwork delays close timelines",
      "Deal minimums often exceed available IRA cash",
    ],
    whereToFind: [
      "Equity Trust Company",
      "STRATA Trust Company",
      "Quest Trust Company",
      "Madison Trust",
      "Directed IRA / Directed Trust",
      "The Entrust Group",
      "Pacific Premier Trust",
      "Kingdom Trust (faith-aligned custodian)",
    ],
    objections: [
      ["Will UBIT / UDFI eat my returns?", "On the leveraged portion, yes — we provide the math. Many investors accept it; some use a Solo 401(k) instead."],
      ["Will my custodian work with you?", "We accept all major SDIRA custodians. Here's our subscription packet calibrated for your custodian."],
      ["What's the close timeline?", "Allow three weeks for custodian paperwork. We hit hard deadlines — 24 deals on schedule."],
    ],
    hook: 'Custodian-side hook: "Your audience is sitting on cash. We have a $100K-minimum single-asset LLC closing now, fits SDIRA structure, full subscription packet ready. Want us on your next webinar?"',
    leadAssets: [
      "SDIRA-specific subscription packet (build if not extant)",
      "Direct-deal one-pager",
      '"How to invest your IRA in real estate" educational piece',
    ],
  },
  {
    id: "P6",
    name: "RIA / Advisor Aggregating Client Capital",
    tagline: "One advisor relationship can deliver ten LPs. But it takes a year.",
    priority: 8,
    bestFit: "Direct deal + Sidecar bundle, placed across multiple advisor clients",
    checkSize: "$1M – $10M aggregated across an advisor's book",
    timeToCash: "6–18 months to first allocation (strategic build)",
    profile: [
      ["AUM", "$200M – $2B (sweet spot $500M – $1.5B)"],
      ["Geography", "SoCal first; SDC's Merrill / FPA / NAPFA alumni network"],
      ["Decision unit", "Lead advisor + investment committee"],
      ["Compensation model", "Fee-only or hybrid; needs RIA-friendly structure"],
      ["Past alts allocation", "Has done at least one alternative-investment placement"],
    ],
    psychographics: [
      "Compliance-anxious — won't risk franchise on a sponsor blowup",
      "Values operator track record + clean docs + audit-grade reporting",
      "Wants to differentiate practice with curated alts access",
      "SDC's RIA pedigree is a major trust signal — use it",
    ],
    jtbd: [
      "Differentiate practice with curated alts",
      "Avoid concentration risk for any one client",
      "Avoid surprise compliance issues",
      "Get clean reporting to feed client portals",
    ],
    painPoints: [
      "Most sponsors don't speak RIA — opaque docs, no IC briefings",
      "Hidden retro fees create compliance nightmares",
      "Sponsor reporting cadence inconsistent with quarterly client review",
    ],
    whereToFind: [
      "FPA chapters (VWC has a calibrated FPA deck)",
      "NAPFA, IMCA, CFA Society SoCal",
      "SEC IAPD database, filtered SoCal + alts-friendly",
      'LinkedIn Sales Navigator: "RIA" or "wealth advisor", CA, $300M+ AUM',
      "Merrill Lynch, Morgan Stanley, UBS alumni — SDC's 16-year pedigree",
    ],
    objections: [
      ["How are you compensated?", "GP carry plus management fee, fully disclosed. No retros to advisors. RIA-friendly comp."],
      ["Can you support quarterly client reporting?", "AppFolio Investor portal feeds client statements. Sample report attached."],
      ["What's your worst deal?", "630 W Washington, San Diego, -33%. Here's what happened, what we learned, and what we changed."],
    ],
    hook: "Sponsor founded by a 16-year Merrill VP. RIA-friendly structures, AppFolio-fed reporting, 12-year track record across 24 deals. Looking to brief your IC.",
    leadAssets: [
      "FPA deck (already calibrated for advisor audience)",
      "SDC bio (RIA-to-sponsor narrative)",
      "Track record one-pager + transparent loss disclosure",
    ],
  },
  {
    id: "P7",
    name: "Tech / Medical High-Income Wealth Builder",
    tagline: "Earning $500K – $2M per year, looking for tax-advantaged offsets.",
    priority: 7,
    bestFit: "Direct deal ($100K) entry; Sidecar ($500K) once they recycle once",
    checkSize: "$100K – $500K per deal",
    timeToCash: "4–10 weeks",
    profile: [
      ["Age", "35–55"],
      ["Income", "$300K – $2M / year (W-2 or partnership)"],
      ["Net worth", "$1M – $10M (accumulating)"],
      ["Profession", "Tech executive, software engineer (post-IPO), surgeon, anesthesiologist, dentist, partner attorney, consultant"],
      ["Location", "SoCal-heavy — but tech also reaches SF, Seattle, Austin"],
    ],
    psychographics: [
      "High-income, time-poor",
      "Tax burden is acute — wants depreciation",
      "Comfortable with risk; data-driven; reads finance blogs and podcasts",
      "Wants passive by definition (W-2 doesn't allow active landlord time)",
      "Often ESG / values-curious",
    ],
    jtbd: [
      "Reduce taxable income via depreciation pass-through",
      "Diversify outside W-2 / RSU concentration",
      "Build passive income streams for future optionality",
      "Avoid being a landlord while owning real assets",
    ],
    painPoints: [
      "W-2 income is fully taxed; hates it",
      "Stock concentration risk (RSUs, options)",
      "Time-poor — can't actively manage anything",
      "Distrust of polished-but-vague sponsors",
    ],
    whereToFind: [
      "LinkedIn Sales Navigator (job titles + SoCal + accredited-likely income tier)",
      "Professional associations: AMA, ADA, ABA, IEEE",
      "Tax-strategy podcasts and forums (BiggerPockets, Wealth Formula, Tax-Smart Real Estate Investors)",
      "Tech-IPO liquidity events (lockup expirations)",
      "Doctor-specific syndication-curious communities (Passive Income MD, Semi-Retired MD)",
    ],
    objections: [
      ["Will the depreciation actually offset my W-2 income?", "Passive losses don't offset W-2 unless you qualify as REPS, but they offset other passive income, carry forward, and offset gain on sale. Your CPA will work the numbers."],
      ["Why not buy a duplex myself?", "You'd spend five hours a week being a landlord. At your billable rate, that's the real cost."],
      ["How do I exit if I need cash?", "You don't, mid-deal. 24–30 month hold. Don't put rainy-day money in."],
    ],
    hook: "For high-W-2 professionals: $100K minimum, single-asset multifamily, depreciation pass-through, 24-month cycle. We do the landlording. You file the K-1.",
    leadAssets: [
      "Tax-benefits one-pager (depreciation, K-1, passive losses)",
      "Direct-deal sample (e.g., Pacific Crest)",
      "LP testimonial from a surgeon or engineer (request from SDC)",
    ],
  },
  {
    id: "P8",
    name: "Recently Liquid (Business Sale, Inheritance, RSU Vest)",
    tagline: "Just hit a liquidity event. $1M–$10M sitting at Schwab. Now what?",
    priority: 6,
    bestFit: "Sidecar ($500K) for meaningful first allocation, or Direct deal ($100K) toe-dip",
    checkSize: "$250K – $5M (often staged across multiple deals)",
    timeToCash: "6–16 weeks",
    profile: [
      ["Age", "35–70"],
      ["Trigger event", "Sold a business, inherited estate, IPO/RSU vest, divorce settlement, retirement lump sum"],
      ["Liquid capital", "$1M – $20M new cash"],
      ["Investment sophistication", "Low to moderate — often first-time private investor"],
      ["Location", "National"],
    ],
    psychographics: [
      "Bewildered by options — Wall Street sharks circling",
      "High anxiety about preserving the gain",
      "Values education and handholding",
      "Often introduced via a professional advisor (CPA, attorney, RIA)",
    ],
    jtbd: [
      "Deploy cash before inflation erodes it",
      "Diversify across asset classes",
      "Find sponsors that won't blow up the windfall",
      "Maintain flexibility to live off the capital",
    ],
    painPoints: [
      "Decision paralysis",
      "Spouse and family pressure",
      "Statistical likelihood of a bad first sponsor experience",
      "Tax surprise on the liquidity event itself",
    ],
    whereToFind: [
      "CPAs and M&A advisors flagging client liquidity events",
      "Estate attorneys at probate or inheritance moments",
      "Business-broker networks (BizBuySell, IBBA chapters)",
      "Wealth-transfer events (NCF, Ron Blue Trust for Christian families)",
      "Local Christian estate-planning networks (overlap with P9)",
    ],
    objections: [
      ["I don't know what I'm doing — how do I evaluate?", "Bring your CPA and attorney to the meeting. We'll walk all three of you through the same docs we'd give an institutional partner."],
      ["What if I need the money?", "Don't put recently-liquid money you might need into this. Reserve 12–18 months of living expenses outside of it. Your advisor will set the bucket sizes."],
      ["What if you go bust?", "Conservative leverage (≤50% LTV after month 6), 24-deal track record, in-house operating team. Here's the portfolio status snapshot."],
    ],
    hook: "Your client just sold the business. They have 12 months to deploy. We have a 24-month single-asset multifamily with full reporting and a 12-year sponsor track record.",
    leadAssets: [
      'Recently-liquid playbook (build): "What to do with the cash for the first 12 months"',
      "Track record one-pager",
      '"Letter to Investors" blog (warm tone)',
    ],
  },
  {
    id: "P9",
    name: "Faith-Based / Values-Aligned Investor",
    tagline: "They invest where their values invest. Subtle but unashamed fit.",
    priority: 9,
    bestFit: "Direct deal ($100K) entry; Sidecar ($500K) once relationship is established",
    checkSize: "$100K – $1M per deal",
    timeToCash: "3–6 months (relationship-led)",
    profile: [
      ["Profile types", "Christian-owned family foundations, HNW Christian families, denominational endowments (small), Christian business owner-operators"],
      ["Net worth", "$2M – $50M"],
      ["Age", "45–75"],
      ["Faith identity", "Practicing Christian — often involved with NCF, Halftime Institute, C12, Pinnacle Forum, Generous Giving"],
      ["Geography", "National, with strong nodes in OC, SD, Dallas, Atlanta, Nashville"],
    ],
    psychographics: [
      "Faith and stewardship are inseparable from money decisions",
      "Values relationship and reputation over raw returns",
      "Long-time-horizon, multi-generational thinking",
      "Allergic to greedy sponsors — but loves operators with character",
      "Comfortable with explicit faith vocabulary",
    ],
    jtbd: [
      "Steward family or foundation capital with measurable returns",
      "Align investment partners with values",
      "Generate income that funds philanthropic giving",
      "Build a legacy with operators they trust",
    ],
    painPoints: [
      "Few faith-aligned operators with institutional-quality track records",
      "Many Christian-branded sponsors are weaker operators",
      "Reluctance to be cold-pitched in faith communities",
    ],
    whereToFind: [
      "National Christian Foundation (NCF) — major donor-advised funds with allocations",
      "Halftime Institute alumni",
      "C12 Group — Christian CEO peer groups",
      "Pinnacle Forum, Generous Giving, Inspire Investing community",
      "Kingdom Trust (Christian SDIRA custodian — overlap with P5)",
      "Local Christian business networks (SoCal CEO Forum, Convene)",
    ],
    objections: [
      ["Are you faith-led, or just faith-friendly?", "Read the Core Values doc. We don't lead with faith publicly — but our team and our underwriting are run on it. Subtle but unashamed."],
      ["What's your charitable giving structure?", "[Confirm with SDC — there's likely a giving practice that should be documented for this audience.]"],
      ["How do you handle losses?", "Transparently. See 630 W Washington. Clear conscience is one of our two core 'we have' values."],
    ],
    hook: "VWC is a SoCal multifamily sponsor founded by a 16-year Merrill VP and faithful steward, with a 12-year track record and single-asset LLC structure. We're inviting introductions, not cold-pitching faith communities.",
    leadAssets: [
      "VWC Core Values document",
      "SDC bio (with the diverse-communities pilot story)",
      '"Letter to Investors" blog',
      '"Please Define Risk" article',
    ],
  },
  {
    id: "P10",
    name: "Co-Sponsor / Operator's Personal Capital",
    tagline: "Other operators investing their own money in deals they don't run.",
    priority: 10,
    bestFit: "Direct deal ($100K) or Sidecar ($500K)",
    checkSize: "$100K – $1M per deal",
    timeToCash: "2–6 weeks (operators decide fast)",
    profile: [
      ["Profile", "Other RE syndicators, GP-track operators, fix-and-flip pros"],
      ["Net worth", "$2M – $20M"],
      ["Investable", "$100K – $1M per deal of personal capital"],
      ["Geography", "National — the RE network is well-connected nationally"],
    ],
    psychographics: [
      "Knows the business — pitch is technical, not emotional",
      "Looks for operators in adjacent or non-competing markets",
      "Recognizes craft — appreciates VWC's vertical integration",
      "Gives credibility halo via word-of-mouth in their networks",
    ],
    jtbd: [
      "Diversify personal capital outside their own deals",
      "Learn from other operators' execution",
      "Build reciprocal LP/GP relationships",
      "Generate passive income while operating their own GP business",
    ],
    painPoints: [
      "Don't want to be in directly competing markets",
      "Suspicious of marketing-heavy sponsors (recognize the tells)",
      "Cycle-anxious — operators are most cycle-aware",
    ],
    whereToFind: [
      "BiggerPockets Pro / Inner Circle networks",
      "GoBundance (operator-investor peer group)",
      "Sumrok community (multifamily mentorship)",
      "Tarl Yarber, Michael Blank, Joe Fairless networks",
      'LinkedIn search: "Managing Partner" + "real estate" + outside SoCal coastal',
      "Conferences: IMN, Best Ever Conf, SUMMIT Multifamily",
    ],
    objections: [
      ["Why your market over mine?", "If you're in TX or FL, SoCal coastal is a perfect diversifier. You can't operate here from there."],
      ["Show me your underwriting.", "In-house team, full-time construction manager, ≤50% LTV after month 6. Here's a sample model."],
      ["What's your basis vs. comps?", "[Show recent deal comp work — they'll respect it.]"],
    ],
    hook: "Operator-to-operator: SoCal coastal multifamily sidecars. ≤50% LTV, 18-month cycle, vertically integrated 56-person team. $100K minimum for accredited, $500K for the 1031-eligible single-asset LLC.",
    leadAssets: [
      "Underwriting sample (request from SDC)",
      "Vertical-integration pitch (the 35-tradesmen story)",
      "Track record with honest loss disclosure",
    ],
  },
];

const channelPartners = [
  {
    id: "CP1",
    name: "CPA / Estate Attorney / 1031 Qualified Intermediary",
    role: "Referral partner — flags client liquidity events and 1031 deadlines",
    refers: "P3 (1031 Exchangors), P4 (Income Retirees), P8 (Recently Liquid)",
    wants: "Reliable sponsor + clean K-1 + zero compliance noise",
    compensation: "Typically referral fee or co-marketing — confirm with SDC what's allowed under SEC solicitor rules",
    outreach: "Quarterly portfolio update + invite to one private investor briefing per year. Build a 'trusted-sponsor' partnership pack.",
  },
  {
    id: "CP2",
    name: "Third-Party SDIRA Custodian",
    role: "Channel host for the SDIRA-holder audience (P5)",
    refers: "P5 (SDIRA Investors) at scale via webinars and platform promotion",
    wants: "Educational sponsors who don't sell — that's their compliance posture. They want speakers their audience trusts.",
    compensation: "No direct payment. Custodian benefits from client deployment fees once investments are placed.",
    outreach: "One-page sponsor brief + 30-minute intro deck for custodian education teams. Pitch hosting an educational VWC webinar. Targets: Equity Trust, STRATA, Quest, Madison, Directed IRA, Entrust, Pacific Premier, Kingdom Trust.",
  },
  {
    id: "CP3",
    name: "Christian Business Network Introducer",
    role: "Wedge into the faith-based investor community (P9)",
    refers: "P9 (Faith-Based Investors) — relationship-led only",
    wants: "Operators with character and competence in equal measure",
    compensation: "None (referral-driven). Reciprocity through relationship.",
    outreach: "Get warm introductions only — never cold outreach into faith communities. Targets: NCF (National Christian Foundation), Halftime Institute, C12 Group, Pinnacle Forum, Generous Giving.",
  },
];

// ============================================================
// LOGO
// ============================================================
const LOGO_PATH = path.resolve(__dirname, "..", "..", "company-intel", "brand", "logo", "logo-rgb.png");
const logoBuffer = fs.existsSync(LOGO_PATH) ? fs.readFileSync(LOGO_PATH) : null;

// ============================================================
// BUILD SECTIONS
// ============================================================

// ---- COVER PAGE ----
const cover = [
  // Logo (centered, ~3" wide)
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

  // Green divider
  new Paragraph({
    spacing: { before: 200, after: 800 },
    alignment: AlignmentType.CENTER,
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: COLOR.green, space: 1 } },
    children: [new TextRun({ text: "" })],
  }),

  new Paragraph({
    spacing: { after: 240 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Investor Persona Playbook", font: HEADING_FONT, size: 56, bold: true, color: COLOR.charcoal })],
  }),

  new Paragraph({
    spacing: { after: 1200 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "2026 / 2027 Capital Raise Strategy", font: BODY_FONT, size: 28, color: COLOR.charcoalTint })],
  }),

  new Paragraph({
    spacing: { before: 400, after: 200 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Prepared for", font: BODY_FONT, size: 22, color: COLOR.mid })],
  }),

  new Paragraph({
    spacing: { after: 100 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Sanford D. Coggins", font: HEADING_FONT, size: 32, bold: true, color: COLOR.charcoal })],
  }),

  new Paragraph({
    spacing: { after: 800 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "CEO & Founder, VisionWise Capital LLC", font: BODY_FONT, size: 22, italics: true, color: COLOR.mid })],
  }),

  new Paragraph({
    spacing: { before: 600, after: 60 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "April 25, 2026", font: BODY_FONT, size: 22, color: COLOR.charcoal })],
  }),

  new Paragraph({
    spacing: { after: 60 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Document version 1.0 — for SDC review and approval", font: BODY_FONT, size: 18, italics: true, color: COLOR.mid })],
  }),
];

// ---- EXECUTIVE SUMMARY ----
const execSummary = [
  h1("Executive summary"),

  p("This document defines the ten investor personas and three channel-partner profiles that will drive VWC's $10 million (2026) and $25 million (2027) capital raise targets. Each persona was synthesized from the Buyer Profile, Transaction History, Core Values, SDC's annotated strategy notes (“What I Understand About VWC,” 4/24/2026), and the founder's articulated vision."),

  h2("The vision"),

  callout(
    "From the founder",
    "“I want to be the In-N-Out for my type of services — catering to all types of clientele and maintaining a pipeline of investors.”"
  ),

  spacer(),

  p("This vision shapes the entire persona strategy:"),
  bullet("In-N-Out means simple menu, consistent quality, beloved across demographics — the kind of operator everyone names when asked who's the best."),
  bullet("“All types of clientele” means Family Offices and first-time accredited investors are both welcome — but with different products, different channels, and different conversations."),
  bullet("“Pipeline of investors” means this is not a one-deal sales motion. It's a continuous, segmented engine — each persona has its own pipeline, with content, cadence, and conversion ratios tracked separately."),

  h2("The strategic pivot"),

  p("VWC's capital raise has shifted in three material ways from prior years:"),
  numbered("From Fund to single-asset LLCs. Fund IV is closed and liquidating in 2026 / 2027. All forward raises are single-purpose LLCs, with the option to aggregate into a portfolio for a later PE-buyer exit (Blackstone-tier)."),
  numbered("From HNW retail to quasi-institutional. The current 127-LP base averages a $245,000 check; the new target is two to three Family Offices at $10 million-plus each — the most material strategic change."),
  numbered("From product pitch to capital partnership. The headline ask is a $30 million revolving Line of Credit from a Family Office equity partner, refinanced out within 90–100 days of close-of-escrow. This unlocks all-cash acquisitions and operational speed."),

  h2("How to read this document"),

  bullet("Section 2: Strategic framework — priority matrix and channel-by-persona map."),
  bullet("Section 3: Detailed personas — ten profiles, each with demographics, motivations, jobs-to-be-done, where to find them, top objections with responses, opening hook, and lead assets."),
  bullet("Section 4: Channel partners — three referral and gatekeeper profiles."),
  bullet("Section 5: 2026 / 2027 capital plan — how the personas combine to hit the targets."),
  bullet("Section 6: Open questions — items requiring SDC's input before the lead-gen engine can launch at full speed."),

  p("Each persona profile is designed to be a stand-alone briefing for the team member running that channel. The intent is for SDC to red-line, approve, or extend each persona before any outreach begins.", { italic: true, color: COLOR.charcoalTint }),
];

// ---- STRATEGIC FRAMEWORK ----
const strategicFramework = [
  h1("Strategic framework"),

  h2("Priority matrix"),

  p("Personas ranked by 2026 strategic priority — a function of capital concentration, time-to-cash, and strategic fit with the In-N-Out vision."),

  buildHeaderTable(
    ["Rank", "Persona", "Best-fit product", "Check size", "Time-to-cash"],
    [
      ["1", "P1 Family Office Capital Partner", "$30M LOC + anchor LP", "$10M – $20M+", "3–9 months"],
      ["2", "P3 1031 Exchangor", "Sidecar (1031-eligible)", "$500K – $5M", "Days–weeks"],
      ["3", "P4 Income-Seeking Retiree", "Direct deal", "$100K – $1M", "4–12 weeks"],
      ["4", "P2 HNW CRE Veteran", "Sidecar", "$500K – $2M", "2–8 weeks"],
      ["5", "P5 SDIRA Self-Directed Investor", "Direct deal", "$100K – $500K", "2–8 weeks"],
      ["6", "P8 Recently Liquid", "Sidecar / Direct deal", "$250K – $5M", "6–16 weeks"],
      ["7", "P7 Tech / Medical Wealth Builder", "Direct deal / Sidecar", "$100K – $500K", "4–10 weeks"],
      ["8", "P6 RIA Channel", "Bundle across deals", "$1M – $10M aggregated", "6–18 months"],
      ["9", "P9 Faith-Based Investor", "Direct deal / Sidecar", "$100K – $1M", "3–6 months"],
      ["10", "P10 Co-Sponsor", "Direct deal / Sidecar", "$100K – $1M", "2–6 weeks"],
    ],
    [780, 3300, 2400, 1800, 1800]
  ),

  spacer(),

  h2("Channel-by-persona map"),

  p("Where each persona is reached most efficiently:"),

  buildHeaderTable(
    ["Persona", "Primary channels"],
    [
      ["P1 Family Office", "Family Office Exchange (FOX), Buttonwood, Multi-Family-Office RIAs (Cresset, Aspiriant, Mariner, Pathstone), Tiger 21, R360, business-sale liquidity events"],
      ["P2 HNW CRE Vet", "VWC HubSpot CRM, Marcus & Millichap alumni, IMN, Bisnow, OC/SD direct mail (15K scrubbed), CCIM/IREM/AAOC chapters"],
      ["P3 1031 Exchangor", "Qualified Intermediaries (IPX1031, Asset Preservation, FAI, Madison 1031), CPAs, estate attorneys, VWC Exchange Services"],
      ["P4 Income Retiree", "VWC HubSpot warm pool (127 LPs), AppFolio Investor portal, Bisnow, IMN, RIA channel"],
      ["P5 SDIRA", "Equity Trust, STRATA, Quest, Madison Trust, Directed IRA, Entrust, Pacific Premier, Kingdom Trust webinar circuit"],
      ["P6 RIA", "FPA chapters, NAPFA, IMCA, CFA Society, SEC IAPD, LinkedIn Sales Nav, Merrill / MS / UBS alumni"],
      ["P7 Tech / Medical", "LinkedIn Sales Nav (titles + geography), AMA, ADA, ABA, IEEE, BiggerPockets, Wealth Formula, Passive Income MD"],
      ["P8 Recently Liquid", "CPAs, M&A advisors, estate attorneys, BizBuySell, IBBA, NCF / Ron Blue (Christian wealth-transfer)"],
      ["P9 Faith-Based", "NCF, Halftime Institute, C12 Group, Pinnacle Forum, Generous Giving, Kingdom Trust, Convene"],
      ["P10 Co-Sponsor", "BiggerPockets Pro, GoBundance, Sumrok, IMN, Best Ever Conf, SUMMIT Multifamily"],
    ],
    [2400, 7680]
  ),
];

// ---- PERSONA BUILDER ----
function buildPersona(persona) {
  return [
    h1(`${persona.id}. ${persona.name}`),

    callout("Tagline", persona.tagline),
    spacer(),

    buildHeaderTable(
      ["Strategic priority", "Best-fit product", "Check size", "Time-to-cash"],
      [[`#${persona.priority} of 10`, persona.bestFit, persona.checkSize, persona.timeToCash]],
      [1900, 4180, 2200, 1800],
      COLOR.greenDark,
      COLOR.white
    ),

    spacer(),

    h3("Profile"),
    buildKVTable(persona.profile, 2400),
    spacer(),

    h3("Psychographics"),
    ...persona.psychographics.map(t => bullet(t)),
    spacer(80),

    h3("Jobs-to-be-done"),
    ...persona.jtbd.map(t => bullet(t)),
    spacer(80),

    h3("Pain points"),
    ...persona.painPoints.map(t => bullet(t)),
    spacer(80),

    h3("Where to find them"),
    ...persona.whereToFind.map(t => bullet(t)),
    spacer(80),

    h3("Top objections and responses"),
    buildHeaderTable(
      ["Objection", "Response"],
      persona.objections,
      [3500, 6580]
    ),
    spacer(),

    h3("Opening hook"),
    callout("Outreach hook", persona.hook),
    spacer(),

    h3("Lead with these assets"),
    ...persona.leadAssets.map(t => bullet(t)),
  ];
}

const personaSections = personas.flatMap(buildPersona);

// ---- CHANNEL PARTNERS ----
const channelPartnerSection = [
  h1("Channel partners"),

  p("These are not customers but gatekeepers. Investing in these relationships compounds — every introducer reached can deliver multiple personas over time."),
  spacer(),

  ...channelPartners.flatMap(cp => [
    h2(`${cp.id}. ${cp.name}`),
    buildKVTable([
      ["Role", cp.role],
      ["Refers", cp.refers],
      ["What they want", cp.wants],
      ["Compensation", cp.compensation],
      ["Outreach approach", cp.outreach],
    ], 2200),
    spacer(),
  ]),
];

// ---- 2026/2027 CAPITAL PLAN ----
const capitalPlan = [
  h1("2026 / 2027 capital plan"),

  p("How the personas combine to deliver VWC's stated capital goals."),
  spacer(),

  h2("Capital goals"),
  buildHeaderTable(
    ["Year", "Goal", "Primary KPI"],
    [
      ["2026", "$10M raised", "Capital raised — not investor count"],
      ["2027", "$25M raised", "Capital raised — fewer, larger investors"],
    ],
    [1800, 3000, 5280]
  ),
  spacer(),

  h2("Allocation hypothesis — 2026 ($10M target)"),
  p("This is a working hypothesis. Allocate aggressively to P1 (Family Office) and the warm-pool re-engagement personas (P3, P4) — they convert the fastest at the lowest cost."),

  buildHeaderTable(
    ["Persona", "Capital hypothesis", "Strategic rationale"],
    [
      ["P1 Family Office", "$5M – $7M", "One Family Office anchor at $5M+ delivers half the year's target"],
      ["P3 1031 Exchangor", "$1M – $2M", "Time-pressured = fastest cash; ride the QI / CPA referral pipeline"],
      ["P4 Income Retiree (HubSpot warm)", "$1M – $1.5M", "Re-engage 127 existing LPs; cheapest cost per dollar raised"],
      ["P2 HNW CRE Vet", "$500K – $1M", "Sidecar fits perfectly; M&M alumni network"],
      ["P5 SDIRA via Custodian", "$500K – $1M", "Build first custodian webinar relationship in 2026"],
      ["P6 / P7 / P8 / P9 / P10", "Backfill", "Long-cycle — cultivate for the 2027 pipeline"],
    ],
    [2600, 2200, 5280]
  ),
  spacer(),

  h2("Allocation hypothesis — 2027 ($25M target)"),
  p("By 2027, the deeper-cycle personas come online. The In-N-Out vision is now visible: every persona has a working channel, and the capital flow becomes more diversified."),

  buildHeaderTable(
    ["Persona", "Capital hypothesis", "Strategic rationale"],
    [
      ["P1 Family Office", "$10M – $15M", "Scale to 2–3 Family Office relationships; LOC structure live"],
      ["P5 SDIRA via Custodian", "$3M – $5M", "Custodian webinar circuit institutionalized"],
      ["P6 RIA Channel", "$3M – $5M", "Year-1 educational nurture converts into placements"],
      ["P3 / P4 / P2", "$2M – $4M", "Warm pool continues to compound"],
      ["P7 / P8 / P9 / P10", "Backfill", "Healthy diversification"],
    ],
    [2600, 2200, 5280]
  ),
  spacer(),

  h2("The In-N-Out discipline"),
  p("Three principles to maintain consistency across all ten personas:"),
  numbered("Simple menu. Three products only — Fund IV (legacy), Sidecar, Direct deal — plus the LOC ask for P1. We do not custom-build for one investor."),
  numbered("Consistent quality. Every persona gets the same operator-quality due diligence, the same reporting cadence, the same character and competency we'd give a Family Office."),
  numbered("Pipeline orientation. Every channel is a continuous pipeline, not a one-shot sale. Track conversion ratios per persona quarterly."),
];

// ---- OPEN QUESTIONS ----
const openQuestions = [
  h1("Open questions for SDC"),

  p("These items require SDC's input before the lead-gen engine can launch at full speed. None of them block the personas in this document — they refine outreach mechanics and copy precision."),
  spacer(),

  h2("Triage — if we get one 30-minute call, ask these first"),
  numbered("506(b) vs. 506(c)? Gates whether cold outreach to accredited prospects is legally permitted."),
  numbered("$30M LOC structure — pari-passu equity, mezz-style preferred return, or pure debt? Family Offices will ask within five minutes."),
  numbered("Per-product preferred return and waterfall — needed for all collateral."),
  numbered("HubSpot tagging of 127 LPs — by deal, by total invested, by 1031/SDIRA/cash, by faith-aligned. Gates re-engagement campaigns."),
  numbered("Knott Townhome status and named active offerings — gates immediate outreach."),
  numbered("Leadership Team bios — six people including SDC. Gates capacity planning for who else takes investor meetings."),
  numbered("Confirmed 2026 conference calendar (TNDDA, FOX, Buttonwood, Fact Right) and SDC's intent (attend, speak, or sponsor)."),
  spacer(),

  h2("Compliance and solicitation"),
  bullet("Blue-sky filings — list of states with active filings vs. states to avoid."),
  bullet("Referral-fee compliance — what CPA, attorney, or RIA referral compensation structures are allowed under SEC solicitor rules?"),
  spacer(),

  h2("Deal economics"),
  bullet("GP co-invest percentage — how much skin does VWC put in each deal?"),
  bullet("Asset management, acquisition, and disposition fees — what does VWC charge?"),
  bullet("Distribution cadence — quarterly? Semi-annually? Begin month X post-close?"),
  spacer(),

  h2("Active pipeline and capital status"),
  bullet("Is the Encinitas (330 W I Street) deal still raising, or fully subscribed?"),
  bullet("$30M LOC — any Family Office LOIs or soft-circles already in flight?"),
  bullet("Pacific Crest (referenced as the Direct-deal example) — what is it specifically, and what's its status?"),
  spacer(),

  h2("Track-record clarifications"),
  bullet("Founding-date reconciliation: 2011 (SDC bio), 2014 (Buyer Profile), or first transaction Mar-2013?"),
  bullet("LP-level returns vs. project-level: average LP IRR after fees on closed deals — the number prospects actually care about."),
  bullet("Aggregate distributions paid to LPs to date."),
  bullet("Worst-deal full story (630 W Washington, -33%): what happened, what was the lesson, what changed in underwriting?"),
  spacer(),

  h2("Channel specifics"),
  bullet("Named RIA partner firms currently active or warm."),
  bullet("Custodian relationships in flight — has VWC been on an Equity Trust, STRATA, or Quest webinar before?"),
  bullet("Christian network presence: is SDC active in NCF, Halftime, C12, or Pinnacle today? Any warm introducer?"),
  spacer(),

  h2("Tech and operations"),
  bullet("HubSpot — existing pipelines, lists, or sequences in place, or starting from scratch?"),
  bullet("Email-sending domain authentication for cold outreach (SPF / DKIM / DMARC)."),
  bullet("LinkedIn Sales Navigator seat — does SDC or the team have one for prospecting?"),
  spacer(),

  h2("Brand and voice"),
  bullet("External use of faith vocabulary — existing line and phrase library, or do we draft and approve?"),
  bullet("Use of SDC's writing in outreach — can we repurpose blog content into email sequences, with credit?"),
  bullet("“As committed to Character as we are to Competency” — a public-facing line we can use, or internal?"),
];

// ---- CLOSING ----
const closing = [
  h1("Approval and next steps"),

  p("This document is intended for SDC's review. Recommended next steps:"),
  numbered("Red-line each persona — confirm, edit, or reject."),
  numbered("Resolve the seven triage questions in the Open Questions section."),
  numbered("Approve the 2026 capital allocation hypothesis (or adjust)."),
  numbered("Sign off on the In-N-Out vision as the operating principle for all outreach."),
  spacer(),

  callout("Founder vision (to be confirmed)",
    "“I want to be the In-N-Out for my type of services — catering to all types of clientele and maintaining a pipeline of investors.”"
  ),

  spacer(200),

  pRich([
    { text: "Approved by: ", bold: true },
    { text: "_______________________________________________________" },
  ]),
  spacer(),
  pRich([
    { text: "Sanford D. Coggins, CEO & Founder, VisionWise Capital LLC", italic: true, color: COLOR.mid, size: 20 },
  ]),
  spacer(),
  pRich([
    { text: "Date: ", bold: true },
    { text: "_______________________________" },
  ]),
];

// ---- DISCLAIMER (back page) ----
const DISCLAIMER_TEXT = "This communication is provided for informational purposes only and does not constitute an offer to sell or a solicitation of an offer to buy any securities. Any such offer or solicitation will be made only through definitive offering documents (including a private placement memorandum) provided to qualified investors. Past performance is not indicative of future results. Real estate investments involve risk, including loss of principal.";

const disclaimerSection = [
  h1("Important disclosures"),
  p(DISCLAIMER_TEXT, { size: 18, color: COLOR.charcoalTint }),
  spacer(),
  p("This document was prepared for the internal review of Sanford D. Coggins and the VisionWise Capital leadership team. It is not for external distribution.", { size: 18, italic: true, color: COLOR.mid }),
];

// ============================================================
// DOCUMENT
// ============================================================
const doc = new Document({
  creator: "VisionWise Capital",
  title: "VWC Investor Persona Playbook",
  description: "Investor persona playbook for SDC review",
  styles: {
    default: { document: { run: { font: BODY_FONT, size: 22, color: COLOR.charcoal } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 56, bold: true, font: HEADING_FONT, color: COLOR.charcoal },
        paragraph: { spacing: { before: 320, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: HEADING_FONT, color: COLOR.charcoal },
        paragraph: { spacing: { before: 280, after: 180 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: HEADING_FONT, color: COLOR.greenDark },
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
            text: "VWC Investor Persona Playbook  |  For SDC review  |  04.25.2026",
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
      ...execSummary,
      ...strategicFramework,
      ...personaSections,
      ...channelPartnerSection,
      ...capitalPlan,
      ...openQuestions,
      ...closing,
      ...disclaimerSection,
    ],
  }],
});

const outPath = path.join(__dirname, "VWC Investor Personas - 2026-04-25.docx");
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outPath, buffer);
  console.log("Wrote: " + outPath);
});
