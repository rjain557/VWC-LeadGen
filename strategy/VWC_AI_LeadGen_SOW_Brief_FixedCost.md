# SOW Build Brief — VWC AI Lead Generation (FIXED COST)

**Purpose:** Input document for the Technijian legal repo (`/new-sow`) to generate a **fixed-cost** Statement of Work for the VisionWise Capital AI Lead Generation engagement.
**Prepared:** 2026-05-21 · by Technijian · feed to: `c:\vscode\tech-legal\tech-legal\clients\VWC\`

---

## ⚠️ Important — this REVISES an existing T&M SOW

A SOW for this work already exists: **`SOW-VWC-001-AI-Lead-Gen`** (Effective 2026-04-24), billed **Time & Materials** at $250/hr CTO Advisory, ~40 hrs, ~$10,000 estimate, first 2 hours free.

This brief converts that engagement to **FIXED COST**. Pick one before generating:

| Option | SOW Number | Effect |
|---|---|---|
| **A (recommended)** | `SOW-VWC-001-Rev2-AI-Lead-Gen` | Supersede the T&M SOW with a fixed-cost revision (cleaner — same engagement, new commercial model). |
| B | `SOW-VWC-003-AI-Lead-Gen-Fixed` | New SOW; mark 001 superseded. (002 is the SEO SOW.) |

Everything below assumes the **same lean "design → build → hand-off" scope** as SOW-001 (Technijian builds the engine; VWC procures tools and runs/approves outreach). Richer "done-for-you" work is captured as **optional fixed add-on modules** at the end — not in the base fixed price.

---

## SOW INTAKE PARAMETERS (for `/new-sow`)

| Field | Value |
|---|---|
| **Client code** | VWC |
| **Client legal name** | VisionWise Capital, LLC |
| **Address** | 27525 Puerta Real, Suite 300-164, Mission Viejo, CA 92691 |
| **Primary contact** | Sanford Coggins, Founder — info@visionwisecapital.com — (949) 441-5580 |
| **Project type** | AI-Lead-Gen (fixed cost) |
| **MSA reference** | None — standalone SOW. **NDA-VWC-001 in effect** (governs NPI/confidentiality). |
| **Effective date** | Date of last signature (target 2026-05-21) |
| **Compliance flags** | SEC Reg D 506(b)/506(c), SEC Marketing Rule, CAN-SPAM, TCPA, state solicitation. VWC is an RIA. |
| **Pricing model** | **FIXED COST** (Section 5.2 "Fixed Cost" type — Technijian completes at stated price regardless of hours) |
| **Signer (Technijian)** | Ravi Jain, CEO |
| **Resource** | Ravi Jain (CTO/AI Advisory); offshore dev support where applicable |

---

## 1. PROJECT OVERVIEW

### 1.1 Project Title
AI-Driven Lead Generation System — Design & Build (Fixed-Cost)

### 1.2 Project Description
VisionWise Capital, LLC ("VWC") is a Southern California multifamily real estate sponsor that raises capital from accredited investors, Registered Investment Advisers (RIAs), family offices, and high-net-worth individuals. VWC's growth depends on systematically identifying and engaging qualified, enumerable investor audiences.

Technijian will **design and build an AI-powered outbound lead-generation engine** that identifies target investor and partner audiences, enriches and AI-scores them, generates persona-aware outreach with human-in-the-loop approval, and routes qualified prospects into VWC's CRM (HubSpot). This SOW is delivered for a **fixed price** per Section 5; Technijian assumes the risk of effort variance within the defined scope.

The engine targets the enumerable ("buildable list") personas from VWC's go-to-market strategy: **P1 Family Office, P6 RIA, P2 HNW property owners, P7 high-income professionals, P10 co-sponsors**, and channel partners (**CP1** CPAs/1031 QIs/estate attorneys, **CP2** SDIRA custodians). Anonymous-searcher personas (P3/P5/P8) are served by the separate SEO/GEO engagement (`SOW-VWC-002-SEO`) and are out of scope here.

### 1.3 Locations
| Location Name | Code | Address | Billable |
|---|---|---|---|
| VWC — Headquarters | VWC-HQ | 27525 Puerta Real, Ste 300-164, Mission Viejo, CA 92691 | Yes |

All work performed remotely unless both Parties agree in writing to on-site presence.

---

## 2. SCOPE OF WORK

### 2.1 In Scope (fixed-price deliverables)
- **Discovery & Strategy** — confirm target personas, TAM, channel mix, success metrics, and compliance guardrails (506(b)/506(c), CAN-SPAM, TCPA, state solicitation).
- **Architecture & Tool Selection** — design the AI lead-gen stack: LLM selection, data/enrichment sources (Apollo, Hunter, ATTOM, BatchData, OpenCorporates, SEC IAPD, FINTRX), CRM integration (HubSpot), orchestration (n8n/Make/custom).
- **Data Pipeline & Enrichment** — build/configure ingestion + enrichment that pulls prospective investor/RIA/partner records, normalizes them, and appends firmographic/wealth-screening attributes.
- **AI Scoring & Qualification** — LLM-assisted fit/likelihood scoring with auditable reasoning summaries.
- **Personalized Outreach (templates + automation)** — persona-aware email + LinkedIn sequences with human-in-the-loop approval; deliverability setup (SPF/DKIM/DMARC), warm-up guidance, suppression-list management.
- **CRM & Workflow Integration** — push scored leads into HubSpot with full provenance, scoring rationale, next-best-action.
- **Dashboards & Reporting** — KPI dashboard (pipeline velocity, reply rate, meetings booked, conversion).
- **Knowledge Transfer** — operator runbook, architecture reference, recorded training so VWC operates the system independently.

### 2.2 Out of Scope (expressly excluded — Change Order required)
- Investment decisions, solicitation of investors, or any activity requiring a securities license.
- Compliance opinions or legal counsel on SEC/FINRA/state securities law. Technijian flags obvious considerations; **VWC engages qualified securities counsel.**
- Bulk content/marketing-campaign creation beyond templates and examples (see optional Module B).
- Done-for-you target-list building at scale (see optional Module A).
- Dimensional-mailer / analog program design & fulfillment (see optional Module C).
- Purchase of third-party licenses, enrichment data, LLM API usage, or CRM seats — **procured and paid directly by VWC** (Appendix: pass-through costs).
- Sending outreach at scale without VWC review/authorization per campaign.
- Web development, branding, or graphic design.
- Access to VWC investor records / subscription docs / regulated NPI except as strictly necessary and **governed by NDA-VWC-001**.
- 24x7 monitoring, incident response, or managed services (separate MSA/Managed-Services SOW).

### 2.3 Assumptions
- **Fixed price assumes the scope and deliverables in Sections 2.1 and 3.** Material change to persona count, data sources, CRM platform, or deliverables is a Change Order.
- VWC designates one empowered point of contact (default: Sanford Coggins) to approve scope, spend, vendor selection, and outreach go-live.
- VWC procures and pays for all third-party services directly (LLM APIs, enrichment, CRM, sending infrastructure).
- All system-generated outreach is reviewed and approved by VWC before sending; VWC is solely responsible for the legality and content of outbound communications.
- VWC provides timely access (CRM admin, vendor accounts, sample data) and reviews deliverables within the acceptance window (Section 8). Delays in VWC dependencies extend timelines without penalty to Technijian.
- Work product delivered "as-is" per Section 9 and the Technijian Standard Terms.

---

## 3. PROJECT PHASES (fixed price per phase)

> Est. hours are shown for reference only; pricing is **Fixed** — Technijian completes each phase at the stated price regardless of actual hours.

### Phase 1 — Discovery & Strategy
**Deliverables:** discovery notes + 1-page strategy brief; confirmed persona definitions (P1/P6/P2/P7/P10 + CP1/CP2); channel mix + success metrics; tool short-list + cost-of-ownership estimate.
**Est. hours:** 4.0 · **Fixed price: $1,000** *(includes the goodwill 2-hour discovery credit from SOW-001)*

### Phase 2 — Architecture & Tool Selection
**Deliverables:** architecture diagram (sources → enrichment → scoring → outreach → CRM → reporting); vendor recommendation memo (cost, pros/cons, compliance notes); prioritized build backlog.
**Est. hours:** 6.0 · **Fixed price: $1,750**

### Phase 3 — Build & Integration
**Deliverables:** configured ingestion + enrichment pipeline; LLM scoring model with audit-ready rationale; AI outreach sequences with human-in-the-loop review; bi-directional HubSpot integration with scored lead records; KPI dashboard.
**Est. hours:** 21.0 · **Fixed price: $6,000**

### Phase 4 — Pilot, Tuning & Go-Live
**Deliverables:** pilot results report (reply rate, meetings booked, false-positive rate); tuned scoring thresholds + outreach templates; go-live checklist + rollout plan.
**Est. hours:** 7.0 · **Fixed price: $2,000**

### Phase 5 — Knowledge Transfer
**Deliverables:** operator runbook; architecture reference; recorded training session(s).
**Est. hours:** 4.0 · **Fixed price: $1,250**

### Ongoing Advisory (Optional — not in fixed price)
Post-go-live fractional CTO/AI advisory on request. No retainer, no minimum. Billed T&M at the Section 5.1 rate, **or** see optional Module D (fixed monthly).

---

## 4. EQUIPMENT AND MATERIALS
None. Third-party software, data, and API subscriptions are procured by and paid directly by VWC (see Appendix — Pass-Through Costs).

---

## 5. PRICING AND PAYMENT

### 5.1 Rate Card (for Change Orders / optional T&M advisory)
| Role | Location | Rate |
|---|---|---|
| CTO / AI Advisory (Ravi Jain) | US (Remote) | $250.00 / hr |
| After-Hours Premium (weekends, holidays, after 6 PM PT) | US (Remote) | $350.00 / hr |
| Offshore Developer (integration support) | Offshore | $45.00 / hr |

### 5.2 Summary of Costs — FIXED
| Phase | Type | Est. Hours | Cost |
|---|---|---|---|
| Phase 1 — Discovery & Strategy | **Fixed** | 4.0 | $1,000 |
| Phase 2 — Architecture & Tool Selection | **Fixed** | 6.0 | $1,750 |
| Phase 3 — Build & Integration | **Fixed** | 21.0 | $6,000 |
| Phase 4 — Pilot, Tuning & Go-Live | **Fixed** | 7.0 | $2,000 |
| Phase 5 — Knowledge Transfer | **Fixed** | 4.0 | $1,250 |
| **TOTAL (Fixed Project Price)** | **Fixed** | **42.0** | **$12,000** |

> **Fixed Cost** — Technijian completes all Section 3 deliverables at **$12,000 total** regardless of actual hours. The fixed price is ~20% above the prior $10,000 T&M midpoint, reflecting Technijian's assumption of all overrun risk under fixed pricing. To reduce price, descope a phase (each is independently priced).

### 5.3 Payment Schedule (milestone-based)
| Milestone | Invoiced | Amount |
|---|---|---|
| On SOW execution (mobilization) | At signing | $4,800 (40%) |
| On acceptance of Phase 3 (Build & Integration) | Mid-project | $4,200 (35%) |
| On acceptance of Phase 5 (project completion) | Final | $3,000 (25%) |
| **Total** | | **$12,000** |

### 5.4 Payment Terms
Net 30 days from invoice date. Standard late-payment / acceleration / suspension / collection / lien provisions apply (no MSA → standalone provisions per template Section 5.4(b)).

---

## 6. CLIENT RESPONSIBILITIES
(a) Provide access to systems, tools, and personnel (CRM admin, vendor accounts, sample prospect data).
(b) Designate one empowered point of contact (default: Sanford Coggins) to approve scope, spend, vendor selection, outreach go-live.
(c) Review and approve deliverables within **five (5) business days** of submission.
(d) Procure and pay for all third-party services directly (LLM APIs, enrichment, CRM, sending infrastructure).
(e) Ensure all outreach content complies with applicable law (SEC 506(b)/506(c), CAN-SPAM, TCPA, state solicitation) before send. Technijian is not a law firm and provides no legal/compliance advice.
(f) Not use the system or its output for any activity requiring a securities license unless VWC (or a licensed affiliate) holds the requisite license.
(g) Inform internal stakeholders of workflow-affecting changes.

---

## 7. CHANGE MANAGEMENT
Out-of-scope work (added personas, additional data sources, alternate CRM, the optional modules below, or any change to the fixed deliverables) requires a **written Change Order** signed by both Parties before work begins. Emergency exception capped at the lesser of $2,500 or 10% of SOW value, with retrospective Change Order within 3 business days.

---

## 8. ACCEPTANCE
Deliverable-based. Technijian notifies on completion; VWC accepts or details deficiencies within **five (5) business days** (deemed accepted if no response). Deficiency-correction loop until resolution. (Use SOW-001's 5-business-day window for consistency with the active SEO SOW.)

---

## 9. GOVERNING TERMS (no MSA — standalone)
Carry over SOW-001 §9 verbatim:
- (a) Aggregate liability capped at fees paid/payable in the **6 months** preceding the claim.
- (b) No indirect/incidental/consequential/punitive damages (incl. lost profits, lost investors, lost opportunity).
- (c) California law; binding AAA arbitration in Orange County; injunctive relief available for IP/confidentiality.
- (d) Mutual confidentiality during the engagement and **3 years** after; **investor data governed by NDA-VWC-001.**
- (e) VWC owns Client Data; Technijian owns pre-existing tools/methodologies/templates; VWC receives a perpetual, non-exclusive license to the deliverables/configurations for internal business use.
- If the Parties later execute an MSA, the MSA governs and supersedes Section 9 on conflict.

---

## SIGNATURES
**Technijian, Inc.** — Ravi Jain, Chief Executive Officer
**VisionWise Capital, LLC** — Sanford Coggins, Founder

---

## Appendix A — Third-Party Pass-Through Costs (VWC-paid, NOT in fixed fee)
| Item | Est. cost | Frequency | Owner |
|---|---|---|---|
| FINTRX / Family Office List (P1 contacts) | several $K | annual | VWC |
| LinkedIn Sales Navigator (seats) | ~$100/seat/mo | monthly | VWC |
| Apollo / Hunter (contacts + verification) | $100–$500/mo | monthly | VWC |
| ATTOM / BatchData (property-owner data, P2) | usage-based | per use | VWC |
| LLM API usage (scoring + personalization) | usage-based | monthly | VWC |
| HubSpot seats | per HubSpot plan | monthly | VWC |
| Dedicated sending domain + warm-up tool | $50–$150/mo | monthly | VWC |
| Securities counsel review of outreach | attorney hourly | per campaign | VWC |

---

## Appendix B — Optional Add-On Modules (each separately priced, FIXED; Change Order or new SOW)
These capture the "done-for-you" scope from the VWC GTM strategy. Not part of the $12,000 base.

| Module | What Technijian does | Fixed price |
|---|---|---|
| **A — Target Universe Build & Enrichment** | Done-for-you named lists for P1/P6/P2/P7/P10 + CP1/CP2; AI fit-scoring; verification; dedupe vs. existing 127 LPs (data costs pass-through) | **$6,000** |
| **B — Persona Sequence Copywriting** | SDC-voice, compliance-ready email + LinkedIn sequences for the priority personas | **$3,500** |
| **C — Analog Program Design** | Dimensional-mailer design + 15k-owner "tired landlord" letter + QI/custodian partnership packs + conference one-pagers (FOX/Buttonwood/TNDDA/Fact Right) | **$3,500** |
| **D — Managed Campaign Operations** | Ongoing run/optimize: list refresh, sequence mgmt, A/B, analog batches, monthly reporting | **$3,500 / month** *(optional, recurring — monthly, no annual lock-in)* |

> All Appendix B prices are planning figures — confirm before generating the SOW.

---
*Generated by Technijian. Companion strategy: `VWC_Persona_GTM_Strategy.pdf` (this folder). Source records: `tech-legal\clients\VWC\` (SOW-001 T&M, NDA-VWC-001, CONTACTS.md).*
