# Keystone — Pricing (internal)

> Internal working note. GBP, excluding VAT. Target: UK mid-market organisations (roughly 100–2,000 staff).

## 1. Price list

### Subscription (one plan)

| | **Keystone** |
|---|---|
| **Price** | **£5,950 / year** per Microsoft 365 tenant (~£496/month), billed annually in advance |
| Users | **Unlimited** |
| Environments | Production + Dev/Test |
| Product | Everything: all 13 components, 8 screens, Rule Builder, Vendor 360, risk scoring, scorecards, obligations, approvals, BU confidentiality walls, multi-currency, spend import, e-signature integration pack* |
| Updates | All updates and new versions |
| Support | Standard support included (§3) |

\*E-signature connectors need the customer's own Power Automate Premium licence for the Keystone service account.

### Optional services (all delivered remotely, no on-site)

| Service | Price |
|---|---|
| **Guided Installation Day** | **£1,150 fixed** |
| **Premium Support** | **£2,400 / year** |
| **Consultancy** (custom rules, integrations, training, extensions) | **£1,050 / day** |
| **Data migration** (more than 500 contracts, messy registers, bulk document filing) | **from £1,950** (quoted) |

### Offers
- **Install bundle:** subscription + Guided Installation Day in year one = **£6,750** (saves £350).
- **Founding customer:** 25% off year one for the first 10 customers, in exchange for a case study and a reference call. That makes the subscription £4,460, or £5,260 with the install day.

### Typical first-year totals
| Package | Year 1 | Year 2+ |
|---|---|---|
| Subscription only | £5,950 | £5,950 |
| Subscription + Installation Day | £6,750 | £5,950 |
| Subscription + Installation Day + Premium Support | £9,150 | £8,350 |

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

| | Standard (included) | Premium (£2,400 / year) |
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

**Price check against the closest rival:** at CLM 365's top reported tier (~$70/user/year), Keystone's £5,950 breaks even at roughly 100 users. Below that, Keystone looks dearer on paper. Sell against that with:
- no contract caps or seat counting
- vendor management and risk included
- the Rule Builder
- the installation day

Above ~100 users, Keystone is cheaper and the gap widens.

## 5. Rationale for one plan
- **Simpler buying:** one number, no feature comparison, quicker approval (often within a department head's sign-off limit).
- **Simpler engineering:** no feature gating. The licence key only checks tenant and expiry.
- **Revenue grows through services:** support and consultancy carry the upsell, not plans.
- **Raising the price later:** keep the founding discount time-limited, then review the list price after the first 10 customers.

## 6. Before publishing prices
- [ ] Run 5 pricing conversations. If more than 3 call it cheap, raise it to £6,950–£7,500.
- [ ] Verify competitor prices directly. Directory figures drift and conflict.
- [ ] Get the code-components licensing position reviewed by a Microsoft licensing partner.
- [ ] Draft the EULA, SLA, DPA and cancellation terms shown in the FAQ.
