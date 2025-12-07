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

## How to Run

1. Download the release ZIP.
2. Extract anywhere (no installation required).
3. Run `Launcher.bat`.
4. Select your game from the numbered list.
5. Choose your export format.

The tool performs a read-only export and does not interact with Vortex directly while running.

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

## License

Free for personal use and distribution. No warranties or guarantees are provided.

