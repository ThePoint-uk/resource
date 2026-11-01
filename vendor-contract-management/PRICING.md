# Keystone — Suggested pricing (internal)

> Internal working note. Prices are GBP, excluding VAT, for UK mid-market buyers (roughly 100–2,000 staff). Test them with your first 3–5 customers before treating them as fixed.

## 1. Recommended price list

### Subscription (annual, per Microsoft 365 tenant, unlimited users)

| | **Standard** | **Enterprise** |
|---|---|---|
| **Price** | **£5,950 / year** | **£11,950 / year** |
| Monthly equivalent | ~£496 | ~£996 |
| Users | Unlimited | Unlimited |
| Production environments | 1 | 1 + licensed Dev/Test environments |
| All 13 Keystone controls, 8 screens | ✓ | ✓ |
| Rules engine (reminders, escalations, validations, approval routing, obligation templates) | ✓ | ✓ |
| Vendor 360, risk scoring, scorecards, spend import | ✓ | ✓ |
| Business-unit confidentiality walls | — | ✓ |
| E-signature integration pack (DocuSign / Adobe Acrobat Sign)* | — | ✓ |
| Multi-currency reporting | — | ✓ |
| Updates and new versions | ✓ | ✓ |
| Support | Standard | Premium included |
| Quarterly health check and upgrade done for you | — | ✓ |

\*E-signature connectors require the customer's own Power Automate Premium licence for the service account.

### Services (optional)

| Service | Price | What's included |
|---|---|---|
| **Guided Installation Day** (remote) | **£1,150 fixed** | See §2 |
| Guided Installation Day (on-site) | £1,450 + travel at cost | Same, delivered at the customer's office |
| Additional consultancy day | £1,050 / day | Custom rules, data cleansing, integrations, training |
| Data migration pack | from £1,950 | Messy registers, > 500 contracts, document bulk-filing |
| **Premium Support** (add-on to Standard) | **£2,400 / year** | See §3 |

### Launch offers
- **Founding customer:** 25% off the first year's subscription for the first 10 customers, in return for a case study and a reference call. Standard drops to £4,460.
- **Bundle:** Standard + Guided Installation Day in year one for **£6,750** (saves £350).
- **Pilot:** 60-day pilot for £1,500, credited in full against the first year if they buy.

## 2. Guided Installation Day — scope (£1,150)

Fixed price. One consultant, one working day of delivery, plus preparation before and support after.

**Before (about a week ahead)**
- 45-minute readiness call: licensing check, environment/Dataverse capacity, DLP, admin roles, decisions checklist (§2.5 of the Installation Guide)
- Customer sends their contract register for a quick data-quality review

**On the day**
1. Service account, SharePoint site and solution installed and verified (Installation Guide Steps 1–6)
2. Settings, branding, holidays, business units and roles configured
3. Seeded rules reviewed with the business owner. Timings, recipients and approval thresholds set and published
4. Up to **3 custom rules** built with the customer's admin
5. Contract register imported: **up to 500 contracts** from the Keystone template (the customer prepares the data, we fix import errors)
6. Smoke test and go-live checklist
7. **2-hour admin training** (rules, settings, imports, troubleshooting), recorded for the customer

**After**
- **30 days of priority support**: 4-business-hour response, plus one 30-minute follow-up call

**Not included** (quote separately): data cleansing, bulk document filing, custom development, integrations, end-user training sessions beyond the admin session.

## 3. Support levels

| | Standard (included) | Premium (+£2,400/yr, included in Enterprise) |
|---|---|---|
| Channel | Email / portal | Email / portal + Teams call |
| Response — app down / reminders not sending | Next business day | 4 business hours |
| Response — other issues | 2 business days | Next business day |
| Updates and release notes | ✓ | ✓ |
| Upgrades | Customer installs using the guide | Done for you, quarterly |
| Health check (connections, flow failures, rule hygiene) | — | Quarterly |
| Rule-building help | — | 1 hour per month |
| Named consultant | — | ✓ |

## 4. Rationale

1. **Value anchor.** The risk Keystone removes is a missed notice deadline. One unplanned auto-renewal of a £200k contract, or failing to renegotiate it by 5%, costs more than a year of Standard. Sales copy should lead with this.
2. **Market position.**
   - *Dedicated per-user CLM SaaS* is usually priced per user per month, often with platform fees. It typically reaches five figures a year once procurement, legal, finance and contract owners all need access.
   - *Spreadsheets and SharePoint lists* cost nothing, but they're where missed deadlines come from.
   - Keystone sits between them: a premium experience at a predictable tenant price, with **no per-user fees**. That removes the "who gets a seat" argument and lets every contract owner in.
3. **Low delivery cost.** There is no hosting cost to us. Everything runs in the customer's tenant. Margin is mainly support time, so price support explicitly (Premium) and keep Standard support asynchronous.
4. **Why a fixed installation day.** It removes the biggest buying objection ("our IT team doesn't have time"). It sharply improves activation (customers who go live renew), and it's a natural upsell into consultancy. £1,150 is in line with UK Power Platform consultant day rates and is easy to approve on a card or PO.
5. **Two plans, not three.** Mid-market buyers decide faster with fewer choices. Enterprise is justified by things large organisations need (BU walls, Dev/Test, premium support), not by locking away core features. Every customer gets the rules engine, because it's the differentiator.
6. **Annual subscription over perpetual.** It funds continuous updates, keeps up with Microsoft platform changes, and gives recurring revenue. If a buyer insists on perpetual: **£17,950 one-off + 22% a year maintenance** (Standard scope).

## 5. Before publishing prices
- [ ] Run 5 pricing conversations (show the page, ask for reactions to the price, don't negotiate). Adjust if more than 3 of 5 call it cheap.
- [ ] Check current Microsoft licensing guidance for code components in canvas apps, and get it reviewed.
- [ ] Draft the subscription terms (EULA, SLA, DPA) and the cancellation policy shown in the FAQ.
- [ ] Decide on VAT display (B2B: ex VAT is normal) and payment terms (annual in advance, invoice or card).
