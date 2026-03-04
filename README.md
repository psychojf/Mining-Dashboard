# EVE Mining Dashboard

A real-time mining activity tracker for EVE Online. Monitors your game logs to display live stats, critical hit alerts, cargo tracking, session data, and historical mining trends — all in a modular, always-on-top overlay system.

## Features at a Glance

| Feature | Description |
|---|---|
| **Multi-Character Support** | Automatically detects all characters from your game logs |
| **Detachable Overlays** | Pop out individual characters into floating, resizable widgets that can be freely arranged on your screen |
| **State Persistence** | The app automatically remembers which characters are detached, alongside their exact window sizes and screen coordinates across restarts |
| **Cargo Tracking** | Visual neon progress bar with real-time "Time until full" estimates based on your configured ship capacity and active yield |
| **Live Mining Stats** | Tracks total m³ mined and critical hit count in real time |
| **Session Control** | Start/Stop/Empty/Reset buttons per character to measure specific mining runs |
| **Auto-Stop** | Session tracking automatically pauses when cargo is full, an asteroid depletes, or your target is lost |
| **Accordion UI** | Expand or collapse the live per-ore session breakdown to save valuable screen real estate |
| **Theoretical vs Actual Rates** | Compare your configured ship yield against real performance (m³/s and m³/hr) |
| **Ship Profiles** | Create multiple named fittings per character (e.g. "Ore Barge", "Ice Miner", "Gas Huffer") |
| **Drone Mining** | Configure mining drone count, yield, and cycle time per profile for accurate theoretical rates |
| **Implant Support** | Highwall MX-1005 implant toggle (+5% module yield) per profile |
| **Mining History** | Analyze historical mining data across all characters over a configurable number of days |
| **Excel Export** | Export history to EVE-styled Excel spreadsheets -- Summary, Daily Breakdown, Ore Pivot, or Full All-in-One |
| **Critical Hit Alerts** | Desktop notification + sound alert on critical mining successes |
| **Compression Tracking** | Logs in-space compression events from your game logs to accurately reflect remaining cargo space |
| **Fleet Reporting** | Share session reports to Discord via webhook or copy to clipboard -- with confirmation dialog and report preview |
| **SDE Auto-Update** | Ore database can be updated directly from CCP's latest Static Data Export -- no more outdated hardcoded values |
| **Always-On-Top Overlay** | Semi-transparent, borderless windows that stay pinned above your game |
| **Tooltip Hints** | Hover over buttons for contextual tips explaining what they do or why they're disabled |