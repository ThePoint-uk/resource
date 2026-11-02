# Haven — Annual Leave & Absence for Microsoft 365 (design v0.1)

> **Working name: "Haven". Check trademarks before launch.** Part of the same app suite as Keystone (vendor & contract management). It reuses the same architecture and shared components.
> Status: **features for confirmation** · Companion prototype: [`haven-leave-pcf.html`](./haven-leave-pcf.html). Click **PCF controls** to see each control and the flow that feeds it.

---

## 1. Market comparison (Microsoft 365 / Teams leave tools)

Checked 2 November 2026 from vendor sites and review directories. Prices are as published or reported, and they **conflict between sources**. Confirm them before using any in sales material.

| Product | Type | Where data lives | M365 / Teams integration (as described by vendor/listings) | Price (reported) |
|---|---|---|---|---|
| **Timetastic** | UK SaaS leave planner | Vendor cloud | Outlook calendar feeds, Teams integration on Pro | ~£1.20/user/month; Pro ~£2.00 (Teams) |
| **Vacation Tracker** | SaaS leave + attendance | Vendor cloud | Teams / Slack bot, web and mobile | ~$2–$4/user/month; free plan with core features |
| **absentify** | SaaS built around M365 | Vendor cloud | Automatic Outlook out-of-office with placeholders, Outlook / shared / Teams group calendar sync, approve in Teams or Outlook, multi-level approval, Power BI | ~$2.50–4/user/month reported. Shared calendars ~€5/calendar/month add-on |
| **WhosOff** | UK SaaS absence management | Vendor cloud | Outlook calendar subscription, Entra ID SSO, Teams via calendar feed | ~£1.55/user/month |
| **Time Off Manager 365** (Apps365 / HR365) | Runs in SharePoint and Teams, on AppSource | Customer's SharePoint | SharePoint / Teams native, AI-assisted | Inconsistent: free for very small teams, ~$4.99/employee/month quoted on blog |
| **Day Off** | Teams / Slack app | Vendor cloud | Request inside Teams | Per-user (not verified) |
| **BrightHR, Breathe, Factorial…** | Full HR suites | Vendor cloud | Calendar feeds, some Teams | Per-employee HR bundles |

**What the market does well:** simple request and approve, team calendars, Teams/Slack notifications, Outlook calendar sync, mobile apps.

**Gaps Haven targets:**
1. **Data stays in the customer's tenant.** Every leading SaaS except Time Off Manager 365 stores employee absence data (including sickness, which is special-category data under UK GDPR) in the vendor's cloud.
2. **No per-user pricing.** All competitors charge per user per month, and costs grow with headcount. Haven is a one-off licence with unlimited users.
3. **Staffing rules in the customer's hands.** Minimum cover, blackout periods and notice rules are usually fixed toggles in competitors. Haven uses the suite's Rule Builder, so admins write, simulate and publish their own rules.
4. **UK-first entitlement maths.** This means:
   - bank holidays by nation (England & Wales, Scotland, Northern Ireland)
   - pro-rata for part-time, starters and leavers
   - irregular-hours and part-year accrual at **12.07% of hours worked** (for leave years from 1 April 2024)
   - carry-over with expiry
   - TOIL
   - leaver owed / over-taken calculations for payroll

   *We haven't verified how well each competitor handles these. Check each one during sales prep.*
5. **Part of a suite.** One install pattern, one design language and one admin experience across Leave, Expenses, Keystone and the rest.

**Price position:** £1,499 one-off. Against ~£1.20–£3.20 per user per month competitors, Haven costs less in year one above roughly 40–100 users (depending on the competitor), and over three years above roughly 25–50 users. Below about 30 employees, per-user tools are cheaper. Sell to those buyers on data residency, rules and the suite.

---

## 2. Features — **please confirm**

Legend: **MVP** = launch (December target) · **v1.1** = within ~2 months of launch · **Later** = roadmap. Tick, strike or move items.

### 2.1 Employee
| # | Feature | Phase |
|---|---|---|
| E1 | Request leave in the app, in a Teams tab, or on mobile (Power Apps / Teams mobile). Full days, **half days (AM/PM)**, or **hours** | MVP |
| E2 | Live calculation of working days or hours requested, using **my work pattern and my nation's bank holidays** | MVP |
| E3 | **Balance impact before submitting:** remaining after this request, and whether it uses carry-over first | MVP |
| E4 | **Team clash check before submitting:** who else is off, and minimum-cover warnings from rules | MVP |
| E5 | Allowance wallet: entitlement, carry-over (with expiry), taken, booked, pending, remaining, TOIL | MVP |
| E6 | My year: a 12-month calendar of taken, booked and pending leave | MVP |
| E7 | Cancel or change a request. Approved leave needs re-approval to change | MVP |
| E8 | Nominate cover person. Shown to colleagues and included in the out-of-office text | MVP |
| E9 | One-click **Outlook automatic reply** for the leave dates (runs with the employee's own Outlook connection) | MVP |
| E10 | Leave appears in **my Outlook calendar** and the **team's shared calendar** | MVP |
| E11 | Report sickness (self-certification). Upload a fit note after 7 calendar days | MVP |
| E12 | Request TOIL earned (overtime) for approval | v1.1 |
| E13 | Buy / sell holiday scheme | Later |

### 2.2 Manager
| # | Feature | Phase |
|---|---|---|
| M1 | **Team planner:** people × days (week, month, quarter), bank holidays by nation, non-working days, pending shown hatched | MVP |
| M2 | **Daily cover row:** people on duty against the minimum for each day, with breaches highlighted | MVP |
| M3 | Approve or decline **in Teams (adaptive card)**, by email, or in the app. Decline needs a reason | MVP |
| M4 | Approval context: team overlap on those dates, the employee's balance after, and rule checks (notice, cover, blackout) | MVP |
| M5 | Delegate approvals while I'm away (automatic, to my cover or my own manager) | MVP |
| M6 | Escalation if a request is pending more than *N* working hours (rule) | MVP |
| M7 | Return-to-work task after sickness, with a short form | v1.1 |
| M8 | Team insights: untaken leave, people with no leave booked in 3+ months (wellbeing nudge), absence trends | v1.1 |

### 2.3 HR / admin
| # | Feature | Phase |
|---|---|---|
| H1 | **Leave policies:** base entitlement, service increments, bank holidays included or extra, pro-rata rounding, leave year start, units (days or hours) | MVP |
| H2 | **Work patterns:** days and hours per weekday, FTE. Rotating patterns later | MVP (rotating: Later) |
| H3 | **Irregular-hours / part-year accrual (12.07%)** from hours worked, by import or manual entry per pay period | v1.1 |
| H4 | Leave types: Annual, Sick, TOIL, Unpaid, Compassionate, Medical appointment, Study, Jury service, Parental (record only), Volunteering, Birthday. Each has a colour, deducts or not, needs approval or not, visibility | MVP |
| H5 | **Carry-over:** maximum, expiry date, reminders before expiry. Manual statutory exceptions (e.g. sickness or maternity) with reason | MVP |
| H6 | **Entitlement ledger:** every grant, accrual, adjustment, carry-over and booking is a ledger row. Balance = sum. Fully auditable | MVP |
| H7 | **Starter and leaver pro-rata.** Leaver report shows days owed or over-taken for payroll | MVP |
| H8 | Bank holidays for England & Wales, Scotland and Northern Ireland, seeded and editable. Company closure days, e.g. Christmas shutdown auto-booked | MVP |
| H9 | **Rule Builder** (shared with Keystone): minimum cover, maximum off at once, blackout periods, notice rules, auto-approve rules, carry-over expiry reminders, untaken-leave nudges, Bradford Factor triggers, approval escalation | MVP (Bradford: v1.1) |
| H10 | Employee sync from **Entra ID** (manager, department, job title, office) nightly. Opening balances imported from Excel | MVP |
| H11 | **Insights:** absence rate, sickness days, **Bradford Factor**, untaken-leave liability (£), carry-over exposure, leave by month vs last year | v1.1 |
| H12 | **Payroll export (CSV):** unpaid leave, sickness days, leaver holiday pay days, per pay period | v1.1 |
| H13 | Privacy controls: colleagues see "Away" (or the type, if allowed). Sickness reasons visible to HR only | MVP |
| H14 | Multiple policies per organisation (e.g. Scotland office, irregular-hours staff, senior staff) | MVP |
| H15 | Power BI dataset template | Later |

### 2.4 Out of scope for v1
Shift rostering, clocking in/out and timesheets, statutory pay calculations (SSP/SMP amounts), and full HR records. Haven exports days to payroll, but **doesn't calculate pay**.

---

## 3. Architecture

Same pattern as Keystone (see `../vendor-contract-management/DESIGN.md` §2):
- the **canvas app orchestrates**
- **PCFs present** the data
- **Power Automate flows are the API** over **SharePoint**
- flows run as a **service account**, with role checks in the flows
- the **Rule Builder** engine handles policy.

### 3.1 Shared suite core
To share code across the six apps, components are packaged as a **Suite Core** managed solution that each app solution depends on:

| Suite Core (shared, from Keystone) | Used in Haven as |
|---|---|
| `KpiStrip` | Balance summary, HR KPIs |
| `DataGrid` | Employees, adjustments, request history, payroll export preview |
| `SmartForm` | Policy editor, leave types, sickness and return-to-work forms |
| `Chart` | Absence trends, by-team breakdowns |
| `ApprovalPanel` | Approval queue (inbox mode) |
| `ActivityHub` | Request history and entitlement ledger timeline |
| `RuleBuilder` | Cover, blackout and notice rules, nudges, Bradford triggers |
| `RecordHeader` | Employee profile header |
| `@suite/core` npm package | Tokens, formatters, rule evaluator, action contract |

Core flows and lists also move to Suite Core: Rules, RuleRuns, MessageTemplates, Holidays, AppConfig, RoleAssignments and Activity. Each app adds only its domain lists and flows.

### 3.2 New PCFs for Haven (4)

| # | Control | What it does | Actions out |
|---|---|---|---|
| 1 | **Team Planner** `hvn.TeamPlanner` | Virtualised people × days grid (week / month / quarter). Leave-type colours. Pending shown hatched, half days split diagonally, non-working days and nation-specific bank holidays shaded. Sticky names. **Daily cover row** against the minimum, with breaches highlighted. Click-drag across a person's row to start a request, or click a booking for details. Filters by team, location and type. Compact "who's off this week" mode | `select`, `requestRange`, `open`, `navigate` |
| 2 | **Request Composer** `hvn.RequestComposer` | Type picker. Two-month range calendar with AM/PM half days or hours. Live working-day count using the pattern and bank holidays. **Balance before/after bar**. Clash and rule checks (from the `Leave.Check` flow, debounced). Cover person. Note. Fit-note upload for sickness. Shows the approver | `check` (debounced), `submit`, `cancel` |
| 3 | **Allowance Wallet** `hvn.BalanceCard` | Ring for the main allowance (taken / booked / pending / remaining), plus carry-over with an expiry countdown. Secondary bars per type (TOIL, medical, etc.). Tap through to ledger rows | `drill` |
| 4 | **Year Calendar** `hvn.YearCalendar` | 12-month mini-calendar heatmap of one person's leave, bank holidays and closure days. Leave-year aware (e.g. April–March) | `select`, `requestRange` |

**Total for Haven: 4 new and 8 reused.** Reusing components from Keystone is what keeps the build short.

### 3.3 Working-day engine (in `@suite/core`, mirrored in the flow)
`countLeave(range, halfStart, halfEnd, pattern, holidays[nation], closures, unit)` returns days or hours, with a per-day breakdown. The **same function** runs:
- in the Request Composer (instant feedback)
- in `Leave.Submit` (authoritative), via a lookup of pre-computed day rows. Flows can't run TypeScript, so `Nightly.Recalc` keeps a **Calendar** list (one row per date per nation, with a flag for bank holiday or closure). The submit flow sums the pattern hours/days from that.

### 3.4 Data model (SharePoint, Haven-specific)
| List | Key columns |
|---|---|
| **Employees** | EntraId (indexed), Name, Email, Manager (person), Team, Location, **Nation** (E&W / SCO / NI), PolicyId, WorkPatternId, FTE, StartDate, EndDate, LeaveYearStart (if it differs from policy), CoverApprover, IsActive |
| **LeavePolicies** | Name, Unit (Days / Hours), BaseEntitlement, BankHolidaysMode (Included / Additional), ServiceIncrements (JSON: years → +days, cap), ProRataRounding (Up to half day / Nearest half / Exact), LeaveYearStartMonth, CarryOverMax, CarryOverExpiryMonths, AccrualMethod (Upfront / Monthly / **Hours-worked 12.07%**), AllowNegative (days) |
| **WorkPatterns** | Name, Mon–Sun hours (number), DaysPerWeek (calc), RotationJSON (later) |
| **LeaveTypes** | Name, Colour, Icon, DeductsFrom (Annual / TOIL / None), RequiresApproval, AllowHalfDays, AllowHours, Visibility (Type / "Away" / Hidden), EvidenceRequiredAfterDays, Paid (Y/N), Active |
| **LeaveRequests** | RequestRef, **Employee** (indexed), LeaveType, **StartDate**, StartHalf, **EndDate**, EndHalf, Hours, Units (computed), **Status** (Draft · Pending · Approved · Declined · Cancelled · CancelPending), Approver, DecisionOn, DecisionNote, CoverPerson, Note, EvidenceUrl, CalendarEventIds, ETag |
| **LeaveLedger** | **Employee**, **LeaveYear**, EntryType (Entitlement · ServiceIncrement · ProRata · CarryOver · CarryOverExpiry · Accrual · Booking · Adjustment · TOILEarned · TOILExpiry), Units (+/−), RequestRef, Reason, CreatedBy. Balance = SUM per employee, year and type |
| **HoursWorked** | Employee, PeriodStart, PeriodEnd, Hours, Source. Feeds 12.07% accrual |
| **Calendar** | Date, Nation, IsBankHoliday, HolidayName, IsClosure. Generated 3 years ahead |
| **Closures** | Name, StartDate, EndDate, AppliesTo (all / location), AutoBookType |
| **ReturnToWork** (v1.1) | Request, Employee, Manager, CompletedOn, Answers JSON (HR-only) |
| Shared from Suite Core | Rules, RuleRuns, RuleOverrides, MessageTemplates, Holidays (source for Calendar), AppConfig, RoleAssignments, Activity, KpiSnapshots |

Indexes: LeaveRequests (Employee, StartDate, Status), LeaveLedger (Employee, LeaveYear). Volume: 500 staff ≈ 6–8k requests and 15k ledger rows a year, comfortably within limits. Activity and requests older than 6 years are archived.

### 3.5 Flows (Haven-specific)
| Flow | Type | Purpose |
|---|---|---|
| `Haven · Me.Get` | API | Balances (ledger sums), upcoming requests, year calendar, team this week, nudges |
| `Haven · Team.Get` | API | Planner payload for team(s) × date window, including the cover row (per-day on-duty vs rule minimum) |
| `Haven · Leave.Check` | API | Units, balance after, clashes, rule results (validation rules from Suite Core `Rules.Validate`). Called while composing |
| `Haven · Leave.Submit` / `Leave.Cancel` | API | Re-check, write the request (and a provisional ledger row if policy books pending), start the approval |
| `Haven · Approval.Orchestrate` | Automated | Routes to manager → delegate if away → second level (per rules). Teams card with team-overlap summary. Approve: ledger booking, calendar events, notify. Decline: notify with reason |
| `Haven · Calendar.Publish` | Child | Event on the team's **M365 group calendar** (Office 365 Groups connector) and an invite to the employee (*show-as Out of office behaviour to validate in spike*) |
| `Haven · OOO.Set` | API (runs **as the employee**) | Office 365 Outlook "Set up automatic replies", scheduled to leave dates, from a template with cover person. Uses the employee's own connection, so no tenant-wide mailbox permissions are needed |
| `Haven · People.Sync` | Scheduled 01:00 | Entra ID → Employees (Office 365 Users connector): manager, department, job title, office. New starters get a pro-rata entitlement, and leavers are flagged |
| `Haven · YearEnd.Rollover` | Scheduled (per policy leave-year start) | Carry-over (capped) → CarryOverExpiry rows. New entitlement + service increments. Report |
| `Haven · Accrual.Run` | Scheduled monthly / on HoursWorked import | Monthly accrual or 12.07% of hours worked → ledger |
| `Haven · Payroll.Export` | Scheduled per pay period / on demand | CSV to `/Exports/Payroll` |
| `Haven · Leaver.Statement` | Automated (EndDate set) | Accrued vs taken → days owed or over-taken → HR task + payroll line |
| Shared (Suite Core) | — | Rules.Scheduler / Validate / OnChange / Actions / Simulate / Publish, Nightly.Recalc (balances snapshot, Bradford, Calendar list), Digest.Send, Data.Import / Export |

> **Bank holiday updates:** fetching the GOV.UK bank-holidays feed would need the premium HTTP connector. Instead, holidays ship in the seed data and in each release, and admins can edit them. *Optional:* customers with Premium can turn on an automatic GOV.UK sync flow.

### 3.6 Seeded Haven rules (editable in the Rule Builder)
| ID | Type | Rule |
|---|---|---|
| L-01 | Validation (warn) | Request must be made at least 2 × its length in advance (statutory notice default). Managers can override |
| L-02 | Validation (block → manager override) | Team on-duty count must stay ≥ minimum (per team, default 60% of headcount) |
| L-03 | Validation (block) | Blackout periods (e.g. year-end close 28 Dec – 8 Jan for Finance) |
| L-04 | Validation (warn) | More than 10 consecutive working days needs a second-level approver |
| L-05 | Approval routing | Line manager → delegate if manager is away → +HR for Unpaid / Compassionate over 5 days |
| L-06 | Event | Auto-approve Medical appointment ≤ 3 hours and Birthday leave |
| L-07 | Reminder | Carry-over expiry: 60/30/7 days before, to the employee if unused carry-over > 0 |
| L-08 | Reminder | Wellbeing nudge: no annual leave taken or booked in 90 days → employee, then manager after 14 more days |
| L-09 | Reminder | Year-end: more than 5 days remaining 8 weeks before leave-year end → employee + manager |
| L-10 | Reminder | Pending request more than 2 working days → remind approver, more than 4 → escalate to their manager |
| L-11 | Event (v1.1) | Bradford Factor ≥ 200 (rolling 52 weeks) → HR alert. ≥ 400 → HR + manager |
| L-12 | Event (v1.1) | Sickness ended → return-to-work task to manager, due in 2 working days |

### 3.7 Security and privacy
| Role | Sees |
|---|---|
| Employee | Own requests, balances and ledger. Team planner with colleagues' leave shown as **"Away"** (or the type, if the policy allows) |
| Manager | Team details, including type. Sickness shown as "Sick" **without reason** |
| HR | Everything, including sickness reasons, fit notes and return-to-work answers |
| Admin | Configuration, rules, imports |

- **Special-category data:** sickness reasons, fit notes and return-to-work forms are stored in an HR-only list and library with unique permissions. Retention labels apply.
- **DPIA template** is included in the install pack.

### 3.8 Screens
| # | Screen | PCFs |
|---|---|---|
| 01 | **My leave** (employee home) | Allowance Wallet, KPI Strip, Year Calendar, Team Planner (compact "who's off"), nudge banner (rules) |
| 02 | **Request leave** | Request Composer |
| 03 | **Team planner** | Team Planner (full), filters |
| 04 | **Approvals** | Approval Panel (inbox) + Team Planner (strip mode) + rule checks |
| 05 | **Insights** (HR) | KPI Strip, Chart, Data Grid (Bradford, untaken leave) |
| 06 | **Admin** | Smart Form (policies, types, patterns), Data Grid, Rule Builder |
| 07 | **Mobile** | Phone layouts of 01/02 + Teams approval card |

---

## 4. Pricing (from the suite price list)
**£1,499 one-off**, unlimited users. 12 months of updates and support included, then £299/year optional.

Optional:
- Guided Installation (half day, includes policy set-up and opening balances import): £575
- Premium Support: £600/year
- Consultancy: £1,050/day

## 5. Build plan (to hit a December launch)
| Week | Scope |
|---|---|
| 1 | Suite Core extraction from Keystone (shared PCFs, lists, rules flows). Haven lists, seed data (policies, types, bank holidays 2026–2029) |
| 2 | Working-day engine + tests (E&W / SCO / NI, half days, part-time, starters/leavers). Request Composer, Allowance Wallet |
| 3 | Team Planner, Year Calendar. Me.Get / Team.Get / Leave.Check / Submit |
| 4 | Approval.Orchestrate + Teams card, calendar publish, OOO, People.Sync, YearEnd.Rollover. Seeded rules L-01–L-10 |
| 5 | Admin screens, opening-balance import, install pack, smoke tests, pilot customer |

## 6. Open questions
1. **Confirm the feature list and phases** (§2).
2. Is a December launch firm? If so, H3 (12.07% accrual), H11 and H12 stay in v1.1 as shown.
3. Default leave year: January–December or April–March? (Policy-level setting either way. This only affects seed data and demo.)
4. Should colleagues see the leave **type**, or only "Away", by default?
5. Product name: keep "Haven", or brand all six apps under one suite name?
6. **Spike items:** whether invited calendar events show the employee as Out of office in Teams. The behaviour of a per-user OOO flow in Teams mobile.

---
Sources for §1: vendor sites and directories, including [Capterra – Timetastic](https://www.capterra.co.uk/software/146655/timetastic), [Vacation Tracker vs Timetastic](https://vacationtracker.io/blog/vacation-tracker-vs-timetastic-comparing-top-hr-tools/), [absentify](https://absentify.com/en), [absentify pricing FAQ](https://support.absentify.com/en/article/understanding-absentify-pricing-a-detailed-guide-to-our-plans-l2g1pn/), [Software Advice – absentify](https://www.softwareadvice.com/hr/absentify-profile), [Capterra – WhosOff](https://www.capterra.ie/software/65528/whosoff-com), [WhosOff](https://www.whosoff.com/default.aspx), [Apps365 Time Off Manager pricing](https://www.apps365.com/sharepoint-time-off-manager-365/pricing/), [AppSource – Time Off Manager](https://marketplace.microsoft.com/en-us/product/WA200004982), [Day Off](https://day-off.app/?p=44509). UK rules: [GOV.UK/HMRC holiday entitlement](https://www.hmrc.gov.uk/holiday-entitlement-rights/print), [Personnel Today](https://www.personneltoday.com/?p=313907), [Lewis Silkin](https://www.lewissilkin.com/insights/2026/04/07/holiday-entitlement-and-pay).
