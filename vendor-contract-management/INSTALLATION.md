# Keystone — Installation Guide

**Product:** Keystone Vendor & Contract Management · **Version:** `{{VERSION}}` · **Applies to:** Microsoft 365 / Power Platform (commercial cloud)

> **Who this is for:** the IT administrator installing Keystone into your Microsoft 365 tenant.
> **Time needed:** 3–4 hours for installation and checks, plus time to load your data.
> **Prefer us to do it?** The optional *Guided Installation Day* covers everything in this guide with one of our consultants (see section 12).

---

## Contents
1. [How Keystone is built (what you are installing)](#1-how-keystone-is-built)
2. [Before you start: prerequisites checklist](#2-prerequisites-checklist)
3. [What's in the installation pack](#3-installation-pack)
4. [Step 1: Create the service account](#step-1--create-the-service-account)
5. [Step 2: Create the SharePoint site and lists](#step-2--create-the-sharepoint-site-and-lists)
6. [Step 3: Prepare the Power Platform environment](#step-3--prepare-the-power-platform-environment)
7. [Step 4: Import the Keystone solution](#step-4--import-the-keystone-solution)
8. [Step 5: Switch on and check the flows](#step-5--switch-on-and-check-the-flows)
9. [Step 6: Share the app and assign roles](#step-6--share-the-app-and-assign-roles)
10. [Step 7: Configure settings, rules and templates](#step-7--configure-settings-rules-and-templates)
11. [Step 8: Load your data](#step-8--load-your-data)
12. [Step 9: Verify the installation (smoke test)](#step-9--verify-the-installation)
13. [Step 10: Go live](#step-10--go-live)
14. [Upgrading to a new version](#upgrading)
15. [Uninstalling](#uninstalling)
16. [Troubleshooting](#troubleshooting)
17. [Appendices](#appendices)

---

## 1. How Keystone is built

Keystone runs entirely inside **your** Microsoft 365 tenant. Nothing is hosted by us, and your contract data never leaves your SharePoint.

| Component | Where it lives | Purpose |
|---|---|---|
| SharePoint site `/sites/Keystone` | Your SharePoint Online | All data: vendors, contracts, obligations, documents, rules, audit log |
| Keystone solution (managed) | Your Power Platform environment | Canvas app, 13 Keystone controls, ~25 Power Automate flows, connection references, environment variables |
| Service account | Your Microsoft Entra ID | Owns the connections. All reads and writes go through flows running as this account |
| Teams / Outlook | Your tenant | Reminders, escalations and approval cards |

Users never edit the SharePoint lists directly. They get access only through the app. This is what lets Keystone enforce roles, validation and a full audit trail.

---

## 2. Prerequisites checklist

Work through every item before installation day. Each one has a ☐ to tick.

### 2.1 People and admin roles
| ☐ | Who | Needed for |
|---|---|---|
| ☐ | **SharePoint Administrator** (or Global Admin) | Creating the site (Step 2) |
| ☐ | **Power Platform Administrator** (or Global Admin) | Environment settings and DLP (Step 3) |
| ☐ | **User Administrator** (or Global Admin) | Creating and licensing the service account (Step 1) |
| ☐ | **Keystone business owner** (e.g. Head of Procurement) | Decisions in Step 7: roles, reminder schedule, approval thresholds |

### 2.2 Licensing
| ☐ | Requirement | Notes |
|---|---|---|
| ☐ | Microsoft 365 / Office 365 licence with Power Apps and Power Automate **for each user** | E3/E5, Business Basic/Standard/Premium and similar plans include these rights for standard connectors. Keystone uses standard connectors only. |
| ☐ | Licence for the **service account** | M365 E3 or equivalent, including Exchange Online (mailbox for sending alerts) and Teams |
| ☐ | Keystone licence key | Sent with your order. Bound to your tenant ID |
| ☐ | *(Optional)* Power Automate Premium for the service account | Only if you use the optional e-signature connectors (DocuSign / Adobe Acrobat Sign) |

> **Confirm with your Microsoft licensing partner.** Our pre-install readiness call (included with the Guided Installation Day) covers this check.

### 2.3 Power Platform
| ☐ | Requirement | How to check |
|---|---|---|
| ☐ | An environment **with a Dataverse database**, used to hold the solution (not your data) | Power Platform admin centre › Environments › *environment* › Details shows a Dataverse URL. We recommend a dedicated **Production** environment called `Keystone`. This needs ~1 GB of available Dataverse capacity. If you have no spare capacity, Keystone can be installed into the Default environment. |
| ☐ | Your DLP policy lets these connectors sit together in the **Business** group: SharePoint, Approvals, Microsoft Teams, Office 365 Outlook, Office 365 Users | Power Platform admin centre › Policies › Data policies |
| ☐ | Teams app policy allows the **Power Automate** app (Flow bot), so adaptive cards are delivered | Teams admin centre › Teams apps › Manage apps |

### 2.4 Workstation for the installer
| ☐ | Requirement |
|---|---|
| ☐ | PowerShell **7.4 or later** |
| ☐ | **PnP.PowerShell** module 2.x+ (`Install-Module PnP.PowerShell -Scope CurrentUser`) |
| ☐ | A PnP Entra app registration for interactive sign-in. Run once: `Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Keystone Setup" -Tenant <yourtenant>.onmicrosoft.com`, and keep the client ID. Requires an admin who can consent. |
| ☐ | *(Optional)* Power Platform CLI (`pac`) for scripted import |

### 2.5 Decisions to make beforehand
| ☐ | Decision | Default if undecided |
|---|---|---|
| ☐ | Site URL | `https://<tenant>.sharepoint.com/sites/Keystone` |
| ☐ | Service account name | `svc-keystone@<yourdomain>` |
| ☐ | Entra security group(s) for users | `Keystone Users`, `Keystone Admins` |
| ☐ | Teams channel for admin and system alerts | A private channel in your IT or Procurement team |
| ☐ | Business units, categories, currency, time zone | One BU, GBP, GMT Standard Time |
| ☐ | Who holds each Keystone role (Admin, Procurement, Legal, Finance, Rules author, Rules publisher) | Installer = Admin only |

---

## 3. Installation pack

Unzip `Keystone_{{VERSION}}_InstallPack.zip` to a local folder. It contains:

| File | Purpose |
|---|---|
| `Keystone_{{VERSION}}_managed.zip` | The Power Platform solution (app, controls, flows) |
| `Provision-KeystoneSite.ps1` | Creates and configures the SharePoint site, lists, columns, indexes, permissions and seed data. Safe to run again |
| `templates/KeystoneSite.xml` | PnP provisioning template used by the script |
| `seed/` | Seeded rules (installed in *Testing* status), message templates, settings, UK bank holidays, categories |
| `import/Keystone_Import_Template.xlsx` | Workbook for loading vendors, contracts, obligations and certificates |
| `ReleaseNotes.md` | What's new, upgrade notes, known issues |
| `licence.txt` | Your licence key *(sent separately by email)* |

Check the download against the SHA-256 checksum in your order email:
```powershell
Get-FileHash .\Keystone_{{VERSION}}_InstallPack.zip -Algorithm SHA256
```

---

## Step 1 — Create the service account

**Time:** 15 minutes · **Role:** User Administrator

1. In the Microsoft 365 admin centre, create the user `svc-keystone@<yourdomain>`, display name **Keystone**. Alerts are sent as this name.
2. Assign the licence from §2.2.
3. **Sign-in security.** The account must stay signed in to its Power Automate connections. Apply MFA through Conditional Access in the normal way, and **don't set a password-expiry policy shorter than your connection review cycle**. If its connections expire, reminders stop (see Troubleshooting).
4. Add the account to the security group **Keystone Admins** (create the group now if needed). Also create **Keystone Users**.
5. In Power Platform admin centre › *your environment* › Users, add the service account with the **Environment Maker** security role.
6. Sign in once as the service account at <https://make.powerautomate.com> and at Teams, to complete first-run set-up.

✅ **Check:** you can sign in as `svc-keystone` and open Power Automate in the target environment.

---

## Step 2 — Create the SharePoint site and lists

**Time:** 20 minutes · **Role:** SharePoint Administrator

1. Open PowerShell 7 in the unzipped pack folder.
2. Run:
   ```powershell
   .\Provision-KeystoneSite.ps1 `
       -TenantUrl        "https://<tenant>.sharepoint.com" `
       -SiteAlias        "Keystone" `
       -ClientId         "<PnP app client ID from §2.4>" `
       -ServiceAccount   "svc-keystone@<yourdomain>" `
       -AdminGroup       "Keystone Admins" `
       -InitialAdmins    "you@<yourdomain>" `
       -TimeZone         "GMT Standard Time" `
       -Currency         "GBP"
   ```
3. Sign in when prompted. The script will:
   - create the site (communication site, no Microsoft 365 group)
   - create 20+ lists and the **Contract Documents** library, with all columns and indexes
   - break permission inheritance on core lists. Only the service account and **Keystone Admins** have access. Normal users have no direct access.
   - load seed data: settings, categories, holidays, message templates, and the **seeded rule library in *Testing* status** (nothing is emailed to users until you publish rules in Step 7)
   - write your licence placeholder and initial Admin role assignments
   - print a **`ListIds` JSON block**. **Copy it**, because you need it in Step 4.
4. The output also goes to `Provision-Keystone-<date>.log`.

✅ **Check:**
- [ ] `https://<tenant>.sharepoint.com/sites/Keystone` opens
- [ ] *Site contents* shows Vendors, Contracts, Obligations, Contract Documents, Rules, RuleRuns, AppConfig (and the others listed in Appendix B)
- [ ] The Rules list shows the seeded rules with Status = **Testing**
- [ ] A normal (non-admin) test user gets *Access denied* on the Contracts list. This is correct.

---

## Step 3 — Prepare the Power Platform environment

**Time:** 15 minutes · **Role:** Power Platform Administrator

1. **Create the environment** (skip if using an existing one): Power Platform admin centre › Environments › **New**. Name `Keystone`, type **Production**, region **United Kingdom** (or your data region), **Add a Dataverse data store = Yes**. Leave the security group empty for now. You'll restrict access later if needed.
2. **Allow code components.** Environment › Settings › Product › **Features** › switch on **"Allow publishing of canvas apps with code components"** › Save.
3. **DLP.** Confirm the policy covering this environment allows the five connectors in §2.3 in the same group.
4. **Time zone.** Environment › Settings › Product › Behaviour: confirm the time zone matches the `-TimeZone` value used in Step 2.

✅ **Check:** the code components feature shows **On**.

---

## Step 4 — Import the Keystone solution

**Time:** 30–40 minutes (mostly waiting) · **Sign in as:** the **service account**

> Import **as the service account**, so that it owns the connections and flows. If you import as yourself, the flows will run as you and stop when your account changes.

1. Go to <https://make.powerapps.com>, select the **Keystone** environment (top right).
2. **Solutions** › **Import solution** › **Browse** › choose `Keystone_{{VERSION}}_managed.zip` › **Next** › **Next**.
3. **Connections.** For each connection reference, choose **+ New connection** and sign in as the service account:

   | Connection reference | Connector |
   |---|---|
   | `kst_SharePoint` | SharePoint |
   | `kst_Approvals` | Approvals |
   | `kst_Teams` | Microsoft Teams |
   | `kst_Outlook` | Office 365 Outlook |
   | `kst_Users` | Office 365 Users |

4. **Environment variables.** Enter:

   | Variable | Value |
   |---|---|
   | `kst_SiteUrl` | `https://<tenant>.sharepoint.com/sites/Keystone` |
   | `kst_ListIds` | Paste the JSON block from Step 2 |
   | `kst_ServiceMailbox` | `svc-keystone@<yourdomain>` (or a shared mailbox the account can *Send As*) |
   | `kst_TeamsChannelId` | Teams admin channel, in the form `<teamId>\|<channelId>`. Get it from *Get link to channel* |
   | `kst_FeatureFlags` | Leave the default `{}` |
   | `kst_LicenceKey` | Paste the contents of `licence.txt` |

5. **Import.** It typically takes 10–20 minutes. Stay on the page until it says the import succeeded.

✅ **Check:** Solutions › **Keystone** opens and lists the *Keystone* app, controls, flows and environment variables. There should be no warning banner.

---

## Step 5 — Switch on and check the flows

**Time:** 15 minutes · **Sign in as:** service account

1. Solutions › Keystone › **Cloud flows**. All flows should show **On**. Switch on any that show *Off*, in this order:
   1. child flows (`Rules.Validate`, `Rules.OnChange`, `Rules.Actions`, `Contract.Provision`)
   2. API flows (`*.Get`, `*.Query`, `Record.Save`, …)
   3. scheduled and automated flows (`Nightly.Recalc`, `Rules.Scheduler`, `Digest.Send`, `Approval.Orchestrate`, `Spend.Import`, `Data.Import`)
2. Open **`Keystone · Nightly.Recalc`** › **Run** › **Run flow**. Wait for it to succeed (1–3 minutes on an empty site). This builds the first dashboard snapshot.
3. Open **`Keystone · Rules.Scheduler`** › **Run**. It should succeed and send nothing, because all rules are in Testing and there's no data yet.

✅ **Check:** both manual runs show **Succeeded** in run history.

---

## Step 6 — Share the app and assign roles

**Time:** 15 minutes · **Role:** Keystone admin

1. Solutions › Keystone › Apps › **Keystone** › **⋯ › Share**.
   - Add **Keystone Users** with permission **User**.
   - Add **Keystone Admins** as **Co-owner** (optional; only if your admins will help support the app).
   - If prompted to share data connections: none are needed, because the app reaches data only through the flows.
2. Copy the **web link** from *Details*. This is the address users open.
3. *(Recommended)* Add Keystone to Teams. In the app, open **⋯ › Add to Teams**, then publish it to your org app catalogue or pin it via an app setup policy.
4. Open Keystone yourself. You were made **Admin** in Step 2.
5. Go to **Rules admin › Settings › Roles** and assign people (or groups) to the roles agreed in §2.5. Changes take effect at each user's next action. No sign-out is needed.

✅ **Check:** a test user in **Keystone Users** can open the app and sees the Portfolio screen (empty).

---

## Step 7 — Configure settings, rules and templates

**Time:** 45–90 minutes · **Role:** Keystone Admin with the business owner

Open **Rules admin** (gear icon at the bottom of the left rail).

1. **Settings**
   - Organisation name, logo and brand colour (used in emails and Teams cards)
   - Currency, time zone, working days, quiet hours (default 19:00–07:00)
   - Bank holidays: UK England & Wales are pre-loaded. Add your region or company shutdown days.
   - *Expiring* window (default 90 days), alert cap per person per day (default 10), digest time (default 08:00)
   - Business units (with head and finance partner), and categories
2. **Message templates.** Review the wording of each template. Use **Preview** to see it against sample data.
3. **Seeded rules.** Every seeded rule is in **Testing**, so it currently sends only to its author. For each rule:
   1. open it, adjust timings, recipients and conditions to suit your policy
   2. click **Simulate** and check the matches and the 90-day forecast
   3. **Publish** when you're happy. Approval-routing and validation rules need a second Rules publisher to approve.

   Minimum set to publish before go-live: **R-001** notice reminders, **R-002** escalation, **R-005** obligation reminders, **R-010–R-012** validations, **R-020–R-025** approval routing.
4. **Owner overrides.** Decide whether contract owners may snooze or add reminders (per rule, under *Options*).

✅ **Check:** the rule library shows your chosen rules as **Published**, and each has a successful simulation.

---

## Step 8 — Load your data

**Time:** depends on data quality. Allow 1–2 hours for up to 500 contracts from a clean register.

1. Open `import/Keystone_Import_Template.xlsx`. Fill in the sheets **in this order**: *Vendors*, *Contracts*, *Obligations*, *Certificates*. Each sheet has a notes row explaining each column. Required columns are marked ★.
   - Use your existing references in `VendorRef` / `ContractRef` if you have them. Leave them blank to have them generated.
   - Dates in `dd/mm/yyyy`. Choice columns must match the drop-down values.
2. Upload the workbook to the SharePoint site folder **Imports › Data**.
3. `Keystone · Data.Import` starts automatically. When it finishes, a results file `…_results.xlsx` appears next to your upload, with one row per record: **Created**, **Updated** or **Rejected** (with the reason). Fix rejected rows and re-upload only those rows. Matching is on Ref, so re-uploading is safe.
4. **Documents.** Drag executed contracts into each contract's folder in **Contract Documents** (created automatically per contract), or upload them from the Documents panel in the app.
5. **Spend (optional).** Drop your monthly ERP export (CSV, format in Appendix D) into **Imports › Spend**.
6. Run **Nightly.Recalc** once more manually, so statuses, notice deadlines, risk scores and the dashboard reflect the imported data straight away.

✅ **Check:** the Portfolio screen shows your contract count, and the renewal timeline is populated.

---

## Step 9 — Verify the installation

Run this smoke test as a non-admin test user with the **Procurement** role. Tick each line.

| ☐ | Test | Expected |
|---|---|---|
| ☐ | Open Keystone | Portfolio loads in under ~3 seconds, KPIs show values |
| ☐ | Vendors › open any vendor | Vendor 360 shows contracts, scorecard, compliance |
| ☐ | New contract › create `TEST – delete me` with TCV £60k and no cost centre | Save is blocked with the message from rule **R-010** |
| ☐ | Add the cost centre and submit | Approval route shows BU head → Finance. A Teams approval card arrives for the BU head |
| ☐ | Open the test contract › Activity | Shows *Created* and *Submitted* entries with your name |
| ☐ | In a second browser, edit the same contract at the same time and save both | The second save shows **"changed by … at …"** (concurrency check) |
| ☐ | Rules admin › R-001 › Simulate | Forecast lists real upcoming reminders |
| ☐ | Run **Rules.Scheduler** manually | Run succeeds. RuleRuns shows Sent/Test rows |
| ☐ | As a user **not** in Keystone Users, open the app link | Access is refused |
| ☐ | As a normal user, open the Contracts list URL in SharePoint directly | *Access denied* |
| ☐ | Delete the test contract (Admin) | Removed. Activity keeps the audit entry |

---

## Step 10 — Go live

1. Publish any remaining rules that are still in Testing.
2. Send the launch message to users. A template is included in `seed/LaunchEmail.docx`.
3. Make sure that **Keystone Users** contains everyone who needs access.
4. Watch **Rules admin › run history** and the admin Teams channel for the first week.
5. Book a 30-day review: check reminder volumes, digest use and any rejected imports.

---

## Upgrading

1. Read `ReleaseNotes.md` for the new version, especially **"Schema changes"** and **"Rule changes"**.
2. If schema changes are listed, run the new script first:
   ```powershell
   .\Provision-KeystoneSite.ps1 -TenantUrl "https://<tenant>.sharepoint.com" -SiteAlias "Keystone" -ClientId "<id>" -Upgrade
   ```
   `-Upgrade` only adds columns, lists and indexes. It never deletes data, and it never overwrites your rules, settings or templates. New seeded rules arrive in **Testing**.
3. Signed in as the service account: Solutions › **Import solution** › choose the new managed zip › leave **Upgrade** selected (the default) › Import.
4. Re-run **Nightly.Recalc** manually, then repeat the Step 9 smoke test (first five rows).

> Keep all changes inside Keystone (rules, settings, templates). **Don't edit the managed solution's components, and don't add columns to the core lists by hand.** Unmanaged edits can block upgrades. Ask us about supported extension points.

---

## Uninstalling

1. Export your data first: Rules admin › Settings › **Export all data** (Excel + documents zip). Alternatively, keep the site.
2. Make.powerapps.com › Solutions › Keystone › **Delete**. This removes the app, controls and flows.
3. *(Optional)* Delete the SharePoint site. It stays in the site recycle bin for 93 days.
4. Remove the service account, groups and the PnP app registration if they're no longer needed.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Import fails: *"code components are not allowed"* | Feature not enabled | Step 3.2, then re-import |
| App opens but every screen spins or says *"Couldn't load"* | Wrong `kst_SiteUrl` / `kst_ListIds`, or flows off | Check the environment variables (Solutions › Keystone › Environment variables). Check flows are On |
| Flows fail with **403 Forbidden** | Service account lost access to the site | Re-run `Provision-KeystoneSite.ps1 -RepairPermissions` |
| Reminders stopped; admin channel shows *"connection invalid"* | Service-account connection expired (password change or MFA re-auth) | Make.powerautomate.com › Connections › fix each `kst_` connection, signed in as the service account |
| Teams cards not arriving; emails do | Power Automate app blocked in Teams | Allow it in Teams admin centre (§2.3) |
| Dashboard shows zeros after data import | Snapshot not refreshed | Run **Nightly.Recalc** manually |
| Notice in Rules admin: *"Updates & support expired"* | Your 12 months of included updates (or your renewal) has ended. Keystone keeps working | To upgrade to newer versions, renew updates & support and paste the new key in Rules admin › Settings › Licence (or update `kst_LicenceKey`) |
| Upgrade import refused: *"Version not covered by licence"* | The new version was released after your updates & support ended | Renew updates & support, enter the new key, then re-import |
| A user sees *"You don't have permission to do that"* | Missing Keystone role | Rules admin › Settings › Roles |
| Import results show *"Choice value not allowed"* | Spreadsheet value doesn't match the list choice | Use the drop-downs in the template. Don't paste values over them |
| Slow screens (> 5 s) | Usually throttling during big imports, or a very large list without indexes | Wait for the import to finish. Run `-Upgrade` to re-apply indexes. Contact support with the flow run ID |

When contacting support, include: Keystone version, environment name, the **correlation ID** (shown in every error message in the app) and the failing flow run URL.

---

## Appendices

### A. Permissions summary
| Principal | SharePoint site | Core lists | Reference lists | Power Platform |
|---|---|---|---|---|
| Service account | Owner | Full control | Full control | Environment Maker, owns flows and connections |
| Keystone Admins | Owner | Full control | Full control | App co-owner (optional) |
| Keystone Users | Visitor | **No access** | Read | App user |

### B. Lists created
Vendors, VendorContacts, Contracts, Contract Documents (library), Obligations, Risks, Certificates, Scorecards, SpendMonthly, Approvals, Activity, Categories, BusinessUnits, RoleAssignments, AppConfig, KpiSnapshots, SavedViews, Rules, RuleVersions, RuleRuns, RuleOverrides, MessageTemplates, Holidays, Imports (library).

### C. Environment variables
See Step 4. All can be changed later in Solutions › Keystone › Environment variables › *Current value*. Never change the *Default value* of a managed variable.

### D. Spend import CSV format
`VendorRef, ContractRef (optional), PeriodMonth (yyyy-mm), Amount, Currency, Source`. One row per vendor/contract per month. Invoice-level files are aggregated automatically if `InvoiceDate` is supplied instead of `PeriodMonth`.

### E. Network and data residency
Keystone makes no calls outside Microsoft 365. Data stays in your SharePoint geography. The Dataverse environment holds only the solution definition. Keystone collects no telemetry, unless you enable Application Insights with your own resource (`kst_FeatureFlags`).

---
---

# Publisher notes: producing the install pack

> **Internal. Remove this section before you ship the guide.** It lists what has to exist, and what you must capture, before the guide above is accurate for a release.

### P1. Build artefacts to produce for each release
| Artefact | How |
|---|---|
| `Keystone_<ver>_managed.zip` | Bump every PCF manifest version and the solution version. `dotnet build -c Release` in `/solution`, then `pac solution export --managed` from the build environment |
| `Provision-KeystoneSite.ps1` + `templates/KeystoneSite.xml` | Export a PnP template from a clean reference site (`Get-PnPSiteTemplate -Handlers Lists,Fields,ContentTypes`), then hand-edit. The script wraps `Invoke-PnPSiteTemplate`, adds permissions, indexes, seed data and the `ListIds` output. Must be **idempotent** and support `-Upgrade` and `-RepairPermissions` |
| `seed/` | JSON for rules (all `status: testing`), templates, settings, holidays (refresh UK bank holidays yearly from gov.uk), categories. Plus `LaunchEmail.docx` |
| `import/Keystone_Import_Template.xlsx` | Data validation drop-downs generated from the list choice values. Notes row per column. Sample rows on a separate sheet |
| `ReleaseNotes.md` | Sections: New, Fixed, **Schema changes**, **Rule changes**, Known issues, Upgrade steps |
| Licence key | Signed per customer with tenant ID + plan + expiry. Keep the private key offline. Record it in your CRM |
| Checksum | `Get-FileHash` of the final zip, sent in the order email |

### P2. Test the guide before each release
1. Use a **clean M365 developer or demo tenant**, not your build tenant.
2. Follow the guide literally as a new admin, timing each step. Update the "Time" lines.
3. Run the smoke test (Step 9). Then run an **upgrade** from the previous version on a second tenant.
4. Record every place you had to guess. Each one becomes a guide fix.

### P3. Screenshots to capture (add under each step)
Environment *Features* toggle (Step 3) · solution import connection screen (Step 4) · environment variables screen (Step 4) · flows list showing On (Step 5) · Share dialog (Step 6) · Rules admin › Settings and a Simulate result (Step 7) · import results file (Step 8). Use the demo tenant with fictional data, and blur tenant names.

### P4. Placeholders to replace
`{{VERSION}}` throughout, the SHA-256 in the order email, and support contact details (add a *Support* section with email, hours and response times matching the plan bought).

### P5. Facts to re-verify each release (Microsoft changes these)
- Admin centre menu paths and setting names (code components feature, DLP screens)
- PnP.PowerShell sign-in requirements (Entra app registration) and minimum PowerShell version
- Platform React / Fluent versions supported for virtual PCFs
- Licensing position for code components in canvas apps, and the standard/premium status of each connector used

### P6. Formats
Publish as PDF (customer pack) and as a web page on your help site. Keep this Markdown as the single source.
