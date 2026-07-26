# Unified Search — SPFx web part

A single, high-end search experience for the SharePoint intranet. One web part,
one page: delegated Microsoft Graph search across **files (SharePoint +
OneDrive), mail, calendar, people, sites and a configurable policies source**,
with dynamic refiners, inline file previews, sharing, and actionable people
cards.

**SharePoint-only by design** — the manifest supports `SharePointWebPart` and
`SharePointFullPage` hosts only. It is not exposed to Microsoft Teams.

Every query runs with the signed-in user's delegated token via the Graph
`/search/query` API, so results are always security-trimmed to what that user
can already see. Nothing in this solution elevates permissions.

## Features

| Area | Detail |
| --- | --- |
| Search bar | Full-width, as-you-type suggestions (recent searches, people, files), `/` keyboard shortcut, spelling "Did you mean…" via Graph query alteration |
| Verticals | All / Files / Mail / Calendar / People / Sites / Policies as a Fluent Pivot with live result counts; each can be switched off in the property pane |
| Refiners | Dynamic panel, position configurable **left / right / hidden**. Files & Policies use real Graph aggregations (file type, last modified). Mail: received date + attachments. People: department & office facets with counts |
| Files | Graph hit-highlighted snippets, inline **preview** (Graph preview endpoint, Office/PDF/300+ formats) in a Fluent Panel with previous/next, **Share** (Graph invite) and **Copy link** (organisation view link) |
| People | Fluent Personas with live **presence**, photos, click-to-**call** (`tel:`), **Teams chat** deep link, **email**, mini org chart (manager + direct reports) you can navigate, department chips that filter results |
| Promoted results | Keyword-triggered answer cards driven by an ordinary **SharePoint list** — content owners manage them, with per-entry enable switch and publish window; no admin centre, no redeploy |
| Zero-query | Recent searches, "pick up where you left off" (Graph insights), upcoming calendar |
| Theming | Inherits the host site design via SPFx `ThemeProvider` — all colours come from Fluent theme tokens, `supportsThemeVariants` + section background aware, `supportsFullBleed` for full-width sections |
| Copilot | Deliberately not in v1. The property pane documents the roadmap: grounded answer card, natural-language filters and chat-with-results via the Graph / Copilot Retrieval API |

## Prerequisites

- Node 18.17+ or 22.14+ (SPFx 1.23 toolchain)
- A SharePoint Online tenant with an App Catalog
- Tenant admin able to approve Graph permissions

## Build & package

```bash
npm install
npm run package        # gulp bundle --ship && gulp package-solution --ship
```

The package lands at `sharepoint/solution/unified-search.sppkg`.
(A pre-built package is committed there, so you can deploy without building.)

## Deploy

1. Upload `unified-search.sppkg` to the tenant App Catalog and deploy
   (tenant-wide or site-by-site; `skipFeatureDeployment` is enabled).
2. In the SharePoint admin centre → **Advanced → API access**, approve the
   pending delegated Graph permissions requested by the package:

   | Scope | Used for |
   | --- | --- |
   | `Files.Read.All` | File & policy search, previews |
   | `Files.ReadWrite.All` | Share (sharing invitations & links) |
   | `Sites.Read.All` | Site search, SharePoint content |
   | `Mail.Read` | Mail vertical |
   | `Calendars.Read` | Calendar vertical + zero-query "coming up" |
   | `User.Read.All` | People directory search, org chart |
   | `People.Read` | Relevance for people |
   | `Presence.Read.All` | Live presence dots |

   The web part degrades gracefully: any vertical whose permission isn't
   approved shows a targeted warning instead of breaking the page.
3. Add the **Unified Search** web part to a page — it works best as the only
   web part in a **full-width section** on a page set as the site (or hub)
   home. Configure in the property pane:
   - Layout: filter panel left/right/hidden, results per page
   - Result types: toggle each vertical; KQL template for the Policies tab
     (point it at your policy content type or library path)
   - Behaviour: scope (organisation / hub / site), previews, suggestions,
     spelling, open-in-new-tab

## Promoted results (list-driven)

Promoted results are keyword-triggered answer cards shown above results on the
**All** tab — the equivalent of Microsoft Search bookmarks, but owned by your
content team in a plain SharePoint list instead of the tenant admin centre.

Provision the list in one command:

```powershell
./provisioning/Create-PromotedResultsList.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/intranet"
```

| Column | Type | Purpose |
| --- | --- | --- |
| Title | Single line | Card headline |
| Keywords | Multi line | Trigger phrases — one per line or `;`-separated ("annual leave; holiday") |
| Description | Multi line | Card body text |
| LinkUrl | Hyperlink | Where the button goes |
| ButtonText | Single line | Button label (defaults to "Open") |
| Enabled | Yes/No | Kill switch per entry |
| StartDate / EndDate | Date | Optional publish window (e.g. only during open enrolment) |

Matching: an exact keyword hit outranks a partial one; the top two matches
render. The list is cached for five minutes per page session, and a missing or
inaccessible list fails silently — promoted results are always optional.

In the property pane ("Promoted results" group) set the list title, and
optionally a different site URL if the list lives on a central site — one list
can then serve every page that hosts the web part.

## Project layout

```
src/webparts/unifiedSearch/
  UnifiedSearchWebPart.ts        SPFx entry + property pane
  models.ts                      Shared interfaces
  services/
    GraphSearchService.ts        /search/query per vertical, refiners, spelling
    PeopleService.ts             Directory search, presence, photos, org chart
    FileService.ts               Preview URLs, share invitations, links, insights
  components/
    UnifiedSearch.tsx            Container: state, orchestration, layout
    SearchBar.tsx                Search box + suggestion callout
    Refiners.tsx                 Dynamic per-vertical filter panel
    ResultRows.tsx               File / mail / event / site renderers
    PeopleGrid.tsx               People cards + person callout (call/Teams/email/org)
    PreviewPanel.tsx             Preview panel + share dialog
    styles.ts                    Fluent theme-token driven styling
```

## Notes & design decisions

- **Hit highlighting** is rendered from Graph's `<c0>` markers as React nodes,
  never `innerHTML`, so indexed content can't inject markup.
- **Refiner tokens** returned by Graph aggregations are passed back verbatim as
  `aggregationFilters` — no client-side re-derivation.
- **Teams deep links** on person cards open the Teams client; the web part
  itself remains SharePoint-hosted only.
- The **Policies** vertical is ordinary driveItem search with an
  admin-controlled KQL template appended, so any policy taxonomy
  (content type, path, managed property) can drive it without code changes.
