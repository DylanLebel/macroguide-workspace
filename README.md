# Nordic Macro Forge

Nordic Macro Forge is the working home for Nordic Minesteel's SolidWorks and
PDM automation macros, plus the browser-based Macro Guide that makes those
tools easier for the team to find, understand, request changes for, and support.

The repository is intentionally practical: VBA macro source lives beside the
single-file guide app, deployment scripts, icon assets, and preview build tools.
That keeps the shop-floor tooling, the engineering automation, and the user
documentation moving together.

## What Is In Here

| Area | Purpose |
| --- | --- |
| `MacroGuide.html` | Single-file dark-mode web app for macro documentation, requests, changelog, crash tracking, polls, and admin tools. |
| `deploy/macro-guide-server.ps1` | Local PowerShell HTTP server that serves the guide and reads/writes shared JSON data. |
| `deploy/Open Macro Guide.vbs` | Launcher used by team members to start the local server and open the guide. |
| `deploy/Deploy-MacroGuide.ps1` | Deployment workflow for copying the guide/server/launcher to the shared PDM location. |
| `docs/index.html` | Generated GitHub Pages preview built from `MacroGuide.html`. |
| Macro folders | Exported SolidWorks VBA source grouped by macro/project. |

## Macro Inventory

| Macro | Summary |
| --- | --- |
| `PDMPDF` | Creates and audits drawing PDFs with PDM-aware revision handling, batch modes, assembly traversal, obsolete-file handling, and audit reporting. |
| `PDMDXF` | Exports DXF burn profiles from selected faces and puts them in the expected PDM folder with revision-aware naming. |
| `Prepare_for_CADLink_1234` | Prepares parts, assemblies, and drawings for CADLink by filling and validating properties such as dimensions, material, weight, type, author, and UOM. |
| `MoveToPDM` | Helps migrate legacy files into the PDM vault while preserving or remapping useful metadata. |
| `NewPL` | Evaluates flat plate/bar-stock parts for in-house manufacturing versus purchasing decisions. |
| `DateFixer` | Replaces bad dynamic drawing creation dates with the correct static revision-table date. |
| `DrawingTemplateUpdate` | Updates drawing sheet formats/templates in bulk. |
| `ModelToImperial` | Converts metric SolidWorks models to IPS units with supporting property and dimension cleanup. |

## Macro Guide Features

- Searchable macro cards with plain-language explanations and usage notes.
- Getting Started guide for installing and using the macros in SolidWorks.
- Changelog for macro updates and fixes.
- Request board for bugs, improvements, new macro ideas, comments, votes, and attachments.
- Team crash tracker with monthly standings, donut pay-up status, notifications, stats, and tie-break polling.
- Team polls for quick operational decisions.
- Admin tools for managing guide data, requests, users, and deployment state.
- Local-only server model: users open a localhost page; shared JSON files hold team data.

## How The Guide Runs

The launcher starts a hidden PowerShell process that hosts the guide on localhost:

```text
Open Macro Guide.vbs
  -> macro-guide-server.ps1
  -> http://localhost:8123
  -> MacroGuide.html
  -> shared JSON data files
```

The server binds only to loopback, verifies local request origins for mutating
API calls, and derives the caller from the Windows account running the server.

## Repository Layout

```text
.
|-- MacroGuide.html
|-- README.md
|-- DateFixer/
|-- DrawingTemplateUpdate/
|-- ModelToImperial/
|-- MoveToPDM/
|-- NewPL/
|-- PDMDXF/
|-- PDMPDF/
|-- Prepare_for_CADLink_1234/
|-- deploy/
|   |-- macro-guide-server.ps1
|   |-- Open Macro Guide.vbs
|   |-- Deploy-MacroGuide.ps1
|   |-- Build-GitHubPreview.ps1
|   |-- Build-MacroGuideIcon.ps1
|   |-- data/
|   `-- icon and preview assets
`-- docs/
    |-- index.html
    `-- .nojekyll
```

## Runtime Data Policy

Files under `deploy/data/` are runtime data, not source code. They may contain
real team activity, names, comments, crash logs, poll votes, usage logs, or
attachments. Keep those files local/shared only unless a specific sanitized
fixture is intentionally added.

The `.gitignore` keeps the active runtime files out of Git, including:

- request, crash, presence, usage, and notification JSON files
- poll, tie-break poll, and donut-duty status JSON files
- logs, locks, and uploaded attachments

## Local Checks

These checks are useful before committing guide or server changes:

```powershell
git diff --check
node -e "const fs=require('fs'); const html=fs.readFileSync('MacroGuide.html','utf8'); const re=/<script[^>]*>([\s\S]*?)<\/script>/g; let m,i=0,bad=0; while((m=re.exec(html))){ i++; try{ new Function(m[1]); } catch(e){ bad++; console.log('script #'+i+' SYNTAX ERROR: '+e.message); } } console.log('checked '+i+' script blocks, '+bad+' errors'); process.exit(bad ? 1 : 0);"
$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile('deploy\macro-guide-server.ps1',[ref]$tokens,[ref]$errors) | Out-Null; $errors
```

## Building The GitHub Preview

The GitHub Pages preview is generated from the source HTML:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File deploy\Build-GitHubPreview.ps1
```

Do not edit `docs/index.html` by hand. Regenerate it from `MacroGuide.html`.

## Deploying To The Team

Deployment copies the guide, server, launcher, and shortcut/icon assets to the
shared Macro Guide location used by the team:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File deploy\Deploy-MacroGuide.ps1
```

The deployment script backs up existing shared files before copying new ones.
After deployment, users pick up the new version by restarting the guide or using
the in-app refresh/update prompt.

## Development Notes

- Keep `MacroGuide.html` self-contained unless there is a strong deployment
  reason to split assets out.
- Prefer changing the source HTML first, then regenerating the preview.
- Keep runtime JSON out of commits.
- Use the existing PowerShell server patterns for API endpoints and JSON file
  locking.
- Keep UI work consistent with the rest of the dark, utility-focused guide:
  constrained content width, readable density, restrained accents, and controls
  that feel like tools rather than marketing panels.

## Maintainer

Maintained by Dylan Lebel for Nordic Minesteel Technologies.
