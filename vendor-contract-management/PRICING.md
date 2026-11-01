# Keystone — Pricing (internal)

> Internal working note. GBP, excluding VAT. Target: UK mid-market organisations (roughly 100–2,000 staff).

## 1. Price list

### Licence (one-off)

| | **Keystone** |
|---|---|
| **Price** | **£2,999 one-off** per Microsoft 365 tenant |
| Licence | Perpetual. It keeps working for as long as the customer uses it |
| Users | **Unlimited** |
| Environments | Production + Dev/Test |
| Product | Everything: all 13 components, 8 screens, Rule Builder, Vendor 360, risk scoring, scorecards, obligations, approvals, BU confidentiality walls, multi-currency, spend import, e-signature integration pack* |
| Included | **12 months** of updates (new versions) and standard support |

\*E-signature connectors need the customer's own Power Automate Premium licence for the Keystone service account.

### Optional (all delivered remotely, no on-site)

| Item | Price |
|---|---|
| **Updates & support renewal** (after the first 12 months: new versions + standard support) | **£599 / year** |
| **Guided Installation Day** | **£1,150 fixed** |
| **Premium Support** (on top of updates & support) | **£1,200 / year** |
| **Consultancy** | **£1,050 / day** |
| **Data migration** (more than 500 contracts, messy registers, bulk document filing) | **from £1,950** (quoted) |

### Bundle
- **Keystone + Guided Installation Day: £3,999** (saves £150).

### Typical totals
| Package | Year 1 | Each later year |
|---|---|---|
| Licence only | £2,999 | £0 (or £599 to keep updates & support) |
| Licence + Installation Day | £3,999 | £0 / £599 |
| Licence + Installation Day + Premium Support | £5,199 | £599 + £1,200 |

## 2. Guided Installation Day (£1,150, remote)

**Before:** a 45-minute readiness call (licensing, environment/Dataverse capacity, DLP, admin roles, the decisions checklist in Installation Guide §2.5) and a data-quality review of the contract register.

**On the day:**
1. Service account, SharePoint site and solution installed and verified (Installation Guide Steps 1–6)
2. Settings, branding, holidays, business units and roles configured
3. Seeded rules reviewed with the business owner, then timings, recipients and approval thresholds set and published
4. Up to 3 custom rules built with the customer's admin
5. Up to 500 contracts imported from the Keystone template
6. Smoke test and go-live checklist
7. 2-hour admin training, recorded

**After:** 30 days of priority support (4-business-hour response) and a 30-minute follow-up call.

**Not included:** data cleansing, bulk document filing, custom development, integrations, end-user training. These are quoted as consultancy days.

## 3. Support

| | Standard (first 12 months included, then £599/yr) | Premium (£1,200 / year) |
|---|---|---|
| Channel | Email / portal | Email / portal + Teams call |
| Critical issue (app down, reminders not sending) | Next business day | 4 business hours |
| Other issues | 2 business days | Next business day |
| Upgrades | Customer installs them using the guide | Done for you, quarterly |
| Health check | — | Quarterly |
| Rule-building help | — | 1 hour per month |
| Named consultant | — | ✓ |

## 4. Market landscape

Checked November 2026 from public listings. Prices are as published or reported by review sites. **Confirm directly with each vendor before quoting them in sales conversations.**

### A. Contract management built on SharePoint / Microsoft 365 (closest competitors)
| Product | What it is | Pricing (public) | How Keystone differs |
|---|---|---|---|
| **CLM 365** (Apps365) | SharePoint-based CLM on AppSource, with AI features and a 14-day trial | Per user per year. Review sites list roughly **$29.99–$69.99/user/year** across three tiers, with a minimum of 5 users. Lower tiers cap contracts per year (e.g. 30 or 60). The vendor's own page shows inconsistent figures | Keystone has no per-user or per-contract caps. It's vendor-led as well as contract-led (Vendor 360, risk, scorecards), and has an admin Rule Builder with simulation |
| **Dock 365 Contract Management** | Contract management built for Microsoft 365 / SharePoint | Not published (quote only) | Transparent fixed price |
| **Consensus** (Barhead) | CLM on the Power Platform for Microsoft 365 | Not found publicly | Keystone stores data in SharePoint (no Dataverse data capacity needed) and uses standard connectors |
| **Contracts 365** | SaaS CLM on Power Platform / Dataverse, recently on Microsoft Marketplace | Not found publicly | Same as above |
| **SP Marketplace** (Contract Tracker) | No-code SharePoint app suite with a contract module | Not checked | A lighter list-style tool. Keystone is a dedicated, premium UX |

### B. Partner-built, bespoke Power Platform projects
Examples: Withum's "Power Apps Contract Management" (sold as a 6-week implementation), LTI's 4-week implementation offer, and many custom builds (e.g. Synapx case study with 90/60/30-day reminders). **These are projects, not products.** Typically several weeks of consultancy, so costs run well above Keystone's first-year price, and the customer then owns the maintenance.

### C. Microsoft's own route (DIY)
Microsoft documents a contract management pattern using **SharePoint Premium (formerly Syntex)** for AI extraction, plus SharePoint lists, Teams and Power Automate. There is also a Syntex *Contracts Management* site template. It needs Syntex / SharePoint Premium licensing for the people who use it, and it's a reference pattern you build and maintain yourself. There's no renewal timeline, vendor view, rules engine or packaged UX. It's **complementary**: Keystone could later read Syntex-extracted metadata (roadmap phase 4).

### D. Standalone CLM SaaS (outside Microsoft 365)
| Product | Pricing (public / reported) |
|---|---|
| **Concord** | ~$499 / $899 / $1,299 per month (5 users included), extra seats $49–$89/user/month |
| **ContractSafe** | Reported from ~$375–$450/month |
| **Juro** | Quote only. Third parties report ~$450–$1,000/month |
| Enterprise CLM (Icertis, Agiloft, DocuSign CLM, Ironclad, Conga) | Quote only. Typically enterprise-scale contracts |

These are strong on drafting, negotiation and e-signature. However, data sits outside the customer's tenant, there's another login and vendor to approve, and costs scale with seats.

### Where Keystone fits
**The gap:** most M365-native tools are per-user and contract-centric. DIY and partner builds are projects. Standalone CLM lives outside the tenant and charges per seat. Keystone's position is a **packaged, premium, vendor-and-contract product that lives entirely in the customer's tenant, at one fixed price for unlimited users, with policy that admins control.**

**Price check against the closest rival:** CLM 365's top reported tier is ~$70 (~£55) per user per year.
- **Year one:** Keystone's £2,999 equals about 55 CLM 365 users.
- **Over three years with updates renewed** (£4,197 total): Keystone is cheaper from roughly **25 users** upwards.
- **Without renewal:** it's cheaper still.

Keystone now undercuts per-user Microsoft 365 tools for almost any real procurement team, and costs a fraction of standalone contract management software.

## 5. Rationale and risks of a one-off price

**Why it works**
- **Impulse-level purchase:** £2,999 is usually within a department head's or IT manager's card or PO limit, with no procurement exercise and no annual budget line. That shortens sales cycles a lot.
- **Clear story:** one price, unlimited users, yours to keep. It's easy to say and hard for per-user rivals to match.
- **No hosting cost:** the software runs in the customer's tenant, so each extra customer costs you almost nothing.
- **Services become the profit engine:** the installation day (£1,150, roughly 38% of the licence) will be attached to most deals. The £599 renewal and Premium Support give recurring income.

**Risks to manage**
- **Less recurring revenue.** Income depends on new sales plus renewals. Aim for over 70% of customers renewing updates & support. Make renewal visible in the app (a gentle notice) and the default on invoices.
- **The platform keeps moving.** Microsoft changes Power Platform and PCF regularly, so you will keep shipping fixes. That's why updates are time-boxed to 12 months, not "free forever". Never promise lifetime updates.
- **Support load:** standard support stays asynchronous (email or portal only), so it doesn't eat consultancy time. Anything hands-on is consultancy at £1,050 a day.
- **Perceived value:** at this price, some buyers may doubt it's "premium". Lean on the demo, the installation day and the case studies.

**Unit economics (illustrative):** 20 sales in year one, with 60% taking the installation day and 70% renewing in year two:
- **Year one:** £59,980 licences + £13,800 installation days ≈ **£73,800**
- **Year two:** about **£8,400** recurring from those customers, plus new sales

## 5a. Pricing principle: volume first
The goal is the **largest possible customer base**, so every price is set to remove friction rather than maximise each deal:
- The licence sits below typical card and purchase-order sign-off limits.
- Services are priced near cost-plus. They're there to drive successful go-lives and renewals, not margin.
- Premium Support is £1,200 (about 40% of the licence). It shouldn't feel like a second product purchase.
- Each customer becomes a reference and a source of word-of-mouth. Ask for a case study at go-live.

## 6. Before publishing prices
- [ ] Run 5 pricing conversations. If buyers don't hesitate at all, test £3,950.
- [ ] Verify competitor prices directly. Directory figures drift and conflict.
- [ ] Get the code-components licensing position reviewed by a Microsoft licensing partner.
- [ ] Draft the perpetual EULA: licence per tenant, what the 12 months of updates and support cover, renewal terms, no refunds after the licence key is issued (or a 14-day window).
