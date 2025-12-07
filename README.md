# Universal Vortex Mod Exporter (UVME)

A utility that exports the installed-mod list from Vortex into a clean, structured, shareable format.

This tool reads the latest Vortex state snapshot and produces:

- HTML export with clickable download links
- JSON export for automation and scripting
- Excel (XLSX) export with hyperlink support

You may export either:

- All installed mods  
- Enabled mods only  

Supports all Vortex-managed games with readable display names and sorted load order (where available).

---

## Why Use This?

UVME is intended for situations where the in-app Vortex view is not ideal, such as:

- Sharing a full mod list for troubleshooting or support
- Documenting a setup before reinstalling or migrating systems
- Comparing load orders between profiles or machines
- Keeping a long-term, readable record of a modded setup

---

## How to Run

1. Download the release ZIP.
2. Extract anywhere (no installation required).
3. Run `Launcher.bat`.
4. Select your game from the numbered list.
5. Choose your export format.

The tool performs a read-only export and does not interact with Vortex directly while running.

---

## OPTIONAL: Run UVME from Inside Vortex

UVME runs perfectly on its own.  
However, it can also be added as a launchable tool inside Vortex for convenience.

### To add UVME as a Vortex tool:

1. Open **Vortex → Dashboard → Add Tool → New…**
2. Set **Target** to your UVME folder → `Launcher.bat`
3. Set **Start In** to the same folder
4. Save

You can now launch UVME with one click directly from inside Vortex.

---

## Requirements

- Windows 10 or Windows 11  
- PowerShell 5.1 or later  
- Vortex must have been opened at least once to generate mod state data  

---

## What This Tool Does Not Do

- It does not download any mods  
- It does not modify Vortex or the game  
- It does not back up mod files or configuration archives  

Only a structured list of mods is exported for documentation, troubleshooting, sharing, or future rebuilds.

---

## Nexus Mods Page

UVME is also available on Nexus Mods:

https://www.nexusmods.com/fallout4/mods/98790

---

## License

Free for personal use and distribution. No warranties or guarantees are provided.
