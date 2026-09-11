# ThePoint MySite

A high-end, highly customisable SharePoint Framework (SPFx) web part that gives every user an
**iGoogle-style personal workspace** inside SharePoint, Teams and Microsoft 365. Users pick, drag,
resize and configure widgets; page owners control the defaults, the look and feel, and which widgets
are on offer.

![SPFx 1.23](https://img.shields.io/badge/SPFx-1.23.2-green) ![Node 22](https://img.shields.io/badge/Node-22-blue) ![React 17](https://img.shields.io/badge/React-17-61dafb)

## Widgets

| Widget | What it shows | Notable features |
| --- | --- | --- |
| **My profile** | Photo, presence, contact, work and about-me details | In-place editing of the fields the user owns (mobile, phones, about me, skills, interests, schools, projects, responsibilities, preferred name, my site). Directory-managed fields are shown with a lock. Field-by-field retry reports exactly which fields the tenant blocks. |
| **My team** | Manager, colleagues who share the manager, direct reports | Teams presence, chat / mail / call actions, tabs and "show all". |
| **My organisation** | Reporting line from the top of the org to the user's reports | Every node expands lazily to its direct reports; configurable depth. |
| **People I work with** | Colleagues ranked by Microsoft 365 relevance | Presence and quick actions. |
| **My Teams** | Teams the user belongs to | Deep links into each team, search, archived toggle. |
| **My apps** | Applications assigned to the user in **Microsoft Entra (My Apps)** plus Microsoft 365 launch tiles | Real app logos and direct sign-in links, search, tiles or list, configurable M365 app set, graceful fallback to the My Apps portal. |
| **Recent documents** | Recent, shared-with-me and trending files (Graph insights) | File-type icons, container names, optional thumbnails. |
| **Sites** | Followed sites, recently active sites and site search | Tiles or list, opens in new tab. |
| **My OneDrive** | Recent files, folder browsing with breadcrumbs, shared files, search | Storage quota bar, link to OneDrive. |
| **News** | SharePoint news across the organisation, hub or site | Cards or list, automatic layout by widget width. |
| **My calendar** | Upcoming events grouped by day | Week paging, in-progress highlight, one-click **Join** for Teams meetings. |
| **My tasks** | Microsoft To Do (add and complete inline) and Planner tasks | List picker, overdue highlighting, importance flags. |
| **My inbox** | Latest inbox messages with unread counter | Unread-only toggle, previews, attachment / importance markers. |
| **Quick links** | Personal bookmarks | Add / edit / reorder / remove, tiles or list, icon auto-detection. |
| **Notes** | A personal sticky note | Five colours, saved with the layout. |

Every widget is an isolated component wrapped in an error boundary, so a failing widget never takes
the dashboard down.

## User experience

* **Customise mode** – press *Customise* to drag widgets by their handle, change widths with the
  1 / 2 / 3 / 4 buttons, rename (double-click the title), collapse, open per-widget settings, remove,
  or add from a searchable, categorised catalogue. Keyboard users can move widgets with the arrow keys
  on the drag handle or the chevron buttons.
* **Auto-save** – every change is saved immediately to the browser and (optionally) synced to the
  user's OneDrive or a SharePoint list so the layout follows them across devices.
* **Import / export** – copy a layout as JSON to back it up or share it; paste to import.
* **Reset** – return to the page owner's default layout.
* **Refresh all** – clears the data cache and reloads every widget.
* **Theme aware** – honours the site theme, section backgrounds and dark mode; works in Teams
  personal apps and tabs.
* **Responsive** – container queries collapse the grid to two columns and then one on narrow hosts.

## Page-owner configuration (property pane)

**General**
* Greeting (time of day / welcome back / none), optional title, date and SharePoint search box.
* Header style: plain, accent block or gradient.
* Columns (1–4), density (comfortable / compact), card style (shadow / outline / flat), corner radius,
  accent colour override, open links in a new tab.

**Widgets**
* Tick / untick each widget to control what users may add (unticked widgets are also removed from
  existing dashboards).
* Default layout as JSON. Easiest route: edit the page → *Customise* → arrange → **Set as page default**
  → republish. Anyone who has not personalised (and anyone who resets) gets this layout.
* Show required permissions in the catalogue.

**Personalisation**
* Allow user customisation, or lock the layout.
* Storage for user layouts:
  * **This browser** – localStorage, no extra permissions.
  * **User's OneDrive** – small JSON file in the OneDrive app folder (`Files.ReadWrite`).
  * **SharePoint list** – item per user in a list on the current site. Create a list with the default
    `Title` column and a **multi-line, plain-text** column named `Layout`; grant users Contribute and set
    *Item-level permissions → Read/Edit items that were created by the user*.
* Data cache duration (minutes) for Graph / search calls.

## Permissions

The package requests these delegated Microsoft Graph scopes (approve them in the SharePoint admin
centre → Advanced → API access after deploying):

| Scope | Used by |
| --- | --- |
| `User.Read.All` | My team, My organisation, people search, photos |
| `User.ReadWrite` | My profile editing |
| `People.Read` | People I work with |
| `Presence.Read.All` | Presence dots on people |
| `Directory.Read.All` | My apps (app role assignments and service principals for logos / launch links) |
| `Sites.Read.All` | Recent documents (insights), followed sites |
| `Files.Read.All` | OneDrive shared-with-me |
| `Files.ReadWrite` | OneDrive browsing, OneDrive layout storage |
| `Calendars.Read` | My calendar |
| `Tasks.ReadWrite` | My tasks (To Do add / complete, Planner) |
| `Group.Read.All` | Planner plan titles |
| `Mail.Read` | My inbox |
| `Team.ReadBasic.All` | My Teams |

Widgets degrade gracefully: a widget whose permission has not been approved shows a clear message
with a retry button rather than breaking the page. If your organisation will not approve a scope,
simply untick that widget in the property pane.

## Build and deploy

```bash
cd thepoint-mysite
npm install
npm run build          # lint, test, bundle and package → sharepoint/solution/thepoint-mysite.sppkg
```

1. Upload `sharepoint/solution/thepoint-mysite.sppkg` to the tenant App Catalog and enable it
   (tenant-wide deployment is supported; `skipFeatureDeployment` is true).
2. Approve the pending API permission requests in the SharePoint admin centre.
3. Add **ThePoint MySite** (group *Other*) to a page, or pin it as a Teams personal app.

For local development run `npm run start` and open the hosted workbench on your tenant.

## Project structure

```
src/webparts/thePointMySite/
├── ThePointMySiteWebPart.ts        # SPFx entry point + property pane
├── models/                         # Shared types (widget definitions, instances, layout state)
├── services/
│   ├── GraphService.ts             # Typed Microsoft Graph wrapper with caching
│   ├── SharePointService.ts        # SharePoint REST (search, profile, list storage)
│   ├── LayoutStore.ts              # Layout persistence (local / OneDrive / list)
│   └── CacheService.ts             # TTL cache (sessionStorage + memory)
├── context/DashboardContext.tsx    # Dashboard state, add/move/remove/update, import/export
├── components/
│   ├── dashboard/                  # Header, grid + drag-and-drop, widget shell, catalogue, settings
│   └── shared/                     # Loading / error / empty states, rows, avatars
└── widgets/                        # One folder per widget + registry.ts
```

### Adding a widget

1. Create `widgets/<name>/<Name>Widget.tsx` exporting an `IWidgetDefinition`:
   `type`, `title`, `description`, `icon`, `category`, `defaultSize`, optional `settingsSchema`
   (drives the automatic settings panel) and the React `component` receiving `IWidgetProps`
   (`graph`, `sp`, `settings`, `updateSettings`, `size`, `density`, `editMode`, …).
2. Register it in `widgets/registry.ts`. It immediately appears in the catalogue, the property-pane
   widget list and the default layout normaliser.
3. If it needs a new Graph scope, add it to `config/package-solution.json`.

## Licence

© The Point. All rights reserved.
