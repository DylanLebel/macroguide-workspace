# Nordic Minesteel — SolidWorks Macros

A collection of SolidWorks VBA macros used at Nordic Minesteel, along with an interactive **Macro Guide** web app for the team.

**Author & Maintainer:** Dylan Lebel
**Contact:** dlebel@nmtech.com

---

## Macros

| Macro | What It Does |
|-------|-------------|
| **PDMPDF** | Generates PDFs from SolidWorks drawings with PDM integration. Supports single files, full assembly audits, batch processing, and command-line mode. |
| **Prepare for CADLink** | Fills in required properties (dimensions, weight, material, type, etc.) and validates everything against CADLink's rules. Works from parts, assemblies, or drawings. |
| **MoveToPDM** | Migrates legacy files into the PDM vault, mapping old properties to the new format. |
| **NewPL** | Determines if a flat plate or bar stock part should be manufactured in-house or purchased, based on shop limits and material. |
| **PDMDXF** | Exports DXF files for flat pattern/laser cutting directly into PDM. Handles single parts and full assemblies. |
| **DateFixer** | Corrects the "Date Created" field on drawings by reading the actual date from the revision table. |
| **DrawingTemplateUpdate** | Swaps old drawing templates/sheet formats to the current company standard in bulk. |
| **ModelToImperial** | Converts metric models to imperial units (IPS) with proper dimension updates. |

## Macro Guide (Web App)

An interactive browser-based guide that helps the team understand and use the macros — no technical background required.

**Features:**
- Searchable macro cards with tips, use cases, and step-by-step info
- Getting Started guide for new employees (how to add macros to SolidWorks)
- Changelog with version tracking
- Request/ticket system where anyone can submit bug reports, feature requests, or change requests
- Upvoting so popular requests get visibility
- Admin controls for managing tickets and changelog

### How It Works

The guide runs as a lightweight local server (PowerShell — built into Windows, nothing to install).

**For users:** Double-click `Open Macro Guide.vbs` from PDM. That's it.

**What happens behind the scenes:**
1. The VBS launcher starts a local PowerShell HTTP server (hidden, no console window)
2. Opens `http://localhost:8123` in your browser
3. The server reads/writes data to JSON files on the shared Y: drive
4. Server auto-shuts down after inactivity

### File Layout

```
C:\AllMacros\                          ← Macro source code (this repo)
├── DateFixer\
├── DrawingTemplateUpdate\
├── ModelToImperial\
├── MoveToPDM\
├── NewPL\
├── PDMDXF\
├── PDMPDF\
├── Prepare_for_CADLink_1234\
├── MacroGuide.html                    ← The web app
└── deploy\                            ← Files that go on the shared drive
    ├── macro-guide-server.ps1
    └── Open Macro Guide.vbs

Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\   ← Deployed location
├── MacroGuide.html
├── macro-guide-server.ps1
├── Open Macro Guide.vbs               ← Also in PDM for easy access
└── data\
    ├── changelog.json
    └── tickets.json
```

### Deploying Updates

After making changes locally:

1. Edit files in `C:\AllMacros\`
2. Copy updated files to `Y:\Solidworks\Macros\Macro Data PDM\MacroGuide\`
3. Users restart the guide to pick up changes (just close the tab and re-open from PDM)

### Admin

The guide has admin mode for managing changelog entries and ticket statuses. Admin login is accessed via the lock icon in the bottom-right corner.
