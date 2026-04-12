# EVE Mining Dashboard

A real-time mining activity tracker for EVE Online. Reads your game logs to display live stats, critical hit alerts, cargo tracking, compression handling, session history, and fleet reporting — all in a compact, always-on-top overlay.

---

## Features

| Feature | Description |
|---|---|
| **Multi-Character Support** | Automatically detects all characters from your EVE game logs |
| **Live Mining Stats** | Tracks total m3 mined, m3/s rate, and critical hit count in real time |
| **Session Control** | Start / Stop / Reset per character to measure specific mining runs |
| **Theoretical vs Actual Rate** | Compare configured ship yield against real-world performance |
| **Ship Profiles** | Multiple named fittings per character (modules, drones, implant, cargo) |
| **Cargo Tracking** | Visual cargo bar with time-to-full estimate; updates automatically on compression |
| **Compression Handling** | Detects in-space compression events and correctly adjusts the cargo bar |
| **Ore Breakdown** | Live per-ore-type volume summary for the current session |
| **Mining History** | Analyze historical data across all characters over a configurable date range |
| **Excel Export** | One-click .xlsx export with Summary, Ore Pivot, and Daily sheets per character |
| **Critical Hit Alerts** | Desktop notification + optional WAV sound on critical mining successes |
| **Fleet Mode** | Share session reports via Discord webhook or clipboard copy |
| **20+ Themes** | EVE faction colour themes (CONCORD, Gallente, Caldari, Amarr, and more) |
| **Always-On-Top Overlay** | Semi-transparent, borderless windows that stay above your game |
| **System Tray** | Minimize to tray; left-click to restore, right-click to exit |

---

## Requirements

- Python 3.10+ (developed on 3.14)
- Windows (system tray and desktop notifications use Windows APIs)

### Python packages

```
playsound==1.2.2
plyer>=2.1.0
openpyxl>=3.1.0
Pillow>=12.0.0
pystray>=0.19.0
```

Install into the project virtual environment:

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements_MD.txt
```

---

## Running from Source

```bash
.venv\Scripts\python.exe mining_dashboard.py
```

---

## Building a Standalone Executable

```bash
.venv\Scripts\pyinstaller --onefile --windowed --icon=mining_icon.ico --add-data "alert_crit.wav;." mining_dashboard.py
```

The compiled `.exe` will be in the `dist\` folder.

---

## Configuration

All settings are stored in `mining_config.json` (auto-created on first run, in the same folder as the script).

Open the **⚙ Config** dialog from the hub to configure:

- **Characters** — choose which pilots appear on the dashboard
- **Appearance** — theme, history days, window transparency, critical hit sound
- **Fleet** — Discord webhook URL for fleet reporting
- **Database** — view SDE version and download the latest ore data

Delete `mining_config.json` to fully reset to defaults.

---

## Ore Database (SDE)

Ore volumes and compression ratios are sourced from the EVE Static Data Export.
A bundled set of data is included out of the box.
Use **⚙ Config → DATABASE → UPDATE ORE DATA** after a major EVE expansion to pull the latest values.

---

## How to Use

See **HOW_TO.txt** for full usage instructions covering:
- Starting and stopping sessions
- Configuring ship fittings and cargo hold
- Compression and cargo bar behaviour
- Fleet mode and Discord reporting
- Mining history and Excel export