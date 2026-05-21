# VWC Lead Generation — Claude Code Project Context

## Client: VisionWise Capital (VWC)

**What they do:** Private real estate investment firm (Southern California focus) that syndicates multifamily and residential deals to accredited investors. CEO and Founder: Sanford D. Coggins (SDC).

**Our goal:** Build a systematic investor lead generation engine — identify, qualify, and engage accredited investors who match VWC's buyer profile.

---

## Knowledge Vault & Key Vault (load at session start)

Two external locations back this project — check them at the start of any VWC session:

| What | Location | Use |
|---|---|---|
| **Obsidian knowledge vault** | `C:\Users\rjain\OneDrive - Technijian, Inc\Documents\obsidian\vwc-leadgen` | Curated VWC intelligence — start at `00 - Dashboard.md`. Holds `Personas.md` (10 investor personas + 3 channel partners), `Channels.md`, `Offerings.md`, `Capital Raise.md`, `Company.md`, `Lead Gen Strategy.md`, `SDC Voice.md`, `Brand & Templates.md`, session logs. Human-curated and richer than the raw PDFs in `company-intel/`. |
| **Key vault (API keys)** | `C:\Users\rjain\OneDrive - Technijian, Inc\Documents\VSCODE\keys\` | Credentials for the AI lead-gen enrichment stack. Lead-gen keys: `apollo.md`, `hunter.md`, `serpapi.md`, `batchdata.md`, `attom.md`, `opencorporates.md`, `sam-gov.md`, `bbc-leadgen.env`. **Never inline secret values into repo files, commits, or outbound content — reference each key by file name and read it at runtime only.** |

> Vault = the *why / who* (strategy, personas, voice). This repo = the *what* (deliverables, prospect data, outreach). Key vault = the *how* (tooling credentials that power outbound AI lead gen).

---

## Repo Purpose

This repo owns the **lead generation** side of VWC's go-to-market:
- Company intelligence and source documents (ingested from client)
- Investor persona definitions (ICP analysis)
- Lead sourcing strategy and prospect lists
- Outreach campaigns (email sequences, LinkedIn, phone scripts)
- Pipeline tracking (prospects → engaged → closed)
- Intelligence handoff to the SEO repo

**SEO repo** (separate): `C:\vscode\seo-sdlc\sdlc\clients\VWC\visionwisecapital.com`
— Keyword strategy, content briefs, and on-page work live there. Feed `seo-intel/` outputs from this repo into it.

---

## Folder Structure

```
company-intel/
  about/             Core values, CEO bio, EOS accountability, brand logo
  track-record/      Transaction history reports + press releases (closed deals)
  offerings/         Offering summaries for active/past deals
  presentations/     Investor pitch decks and PPTs
  content/
    articles/        Published thought leadership articles
    blogs/           Blog posts (SDC's voice — faith, values, real estate)

investor-personas/
  source-docs/       Raw VWC Buyer Profile PDF
  profiles/          Synthesized ICP persona documents (write here)

lead-gen/
  sources/           Where to find investors (LinkedIn, lists, events, referrals)
  prospect-lists/    Raw prospect data CSVs / sheets
  qualified-leads/   Vetted leads who match ICP

outreach/
  email-sequences/   Drip campaigns by persona/stage
  linkedin/          Connection request + message templates
  phone-scripts/     Call scripts

pipeline/
  prospects/         All leads being tracked
  engaged/           Leads in active conversation
  closed/            Won investors

seo-intel/           Research outputs for SEO repo handoff
  keywords/
  content-briefs/
```

---

## Key Source Documents (read these first for any analysis)

| Document | Location | What it tells you |
|---|---|---|
| VWC Buyer Profile.pdf | `investor-personas/source-docs/` | Who VWC's ideal investor is |
| VWC Core Values.docx | `company-intel/about/` | Brand voice, values, culture |
| Sanford D Coggins Bio.pdf | `company-intel/about/` | Founder story and credibility |
| VWC Transaction History (2025).pdf | `company-intel/track-record/` | Full deal track record |
| Offering Summaries (3 deals) | `company-intel/offerings/` | What deals look like |
| Pitch Decks | `company-intel/presentations/` | How VWC pitches investors |

---

## Rules

- Always read `investor-personas/source-docs/VWC Buyer Profile.pdf` before building any ICP or persona doc.
- All prospect data goes in `lead-gen/prospect-lists/` or `lead-gen/qualified-leads/` — never in `company-intel/`.
- Outreach copy must reflect SDC's voice (faith-based, values-driven, direct, relationship-first — evident in the blogs).
- `seo-intel/` outputs are for handoff to the SEO repo — keep them clean and structured.
