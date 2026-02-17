import os
import re
import glob
import json
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from functools import lru_cache
import tkinter as tk
from tkinter import messagebox
import time

# Conditional imports
try:
    from playsound import playsound
    HAS_PLAYSOUND = True
except ImportError:
    HAS_PLAYSOUND = False

try:
    from plyer import notification
    HAS_NOTIFICATION = True
except ImportError:
    HAS_NOTIFICATION = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- CONFIG ---
DOCS = os.path.expanduser(r"~\Documents\EVE\logs\Gamelogs\*")
CRIT_SOUND_FILE = "alert_crit.wav"
CONFIG_FILE = "mining_config.json"
UPDATE_INTERVAL_MS = 1000
HISTORY_DAYS = 15
CRITICAL_HIT_KEYWORD = "Critical mining success"
MAX_MODULES = 5  # Maximum mining modules per ship

# --- EVE ONLINE COLOR PALETTE ---
BG = "#0b0e17"
BG_PANEL = "#111827"
BORDER = "#1e3a4a"
CYAN = "#3dd8e0"
RED = "#cc3325"
GREEN = "#2ecc40"
GOLD = "#ffd700"
DIM = "#5a7085"
WHITE = "#ffffff"

# colors per character
CHAR_ACCENTS = [CYAN, "#ff9f43", "#a29bfe", "#e056fd", "#26de81", "#fc5c65", "#45aaf2", "#fed330"]

# --- ORE VOLUME in m3 ---
_ORE_DATA = [
    # STANDARD ORES
    ("Veldspar", 0.1), ("Scordite", 0.15), ("Pyroxeres", 0.3),
    ("Plagioclase", 0.35), ("Omber", 0.6), ("Kernite", 1.2),
    
    # LOW-SEC ORES
    ("Jaspet", 2.0), ("Hemorphite", 3.0), ("Hedbergite", 3.0),
    
    # NULL-SEC ORES
    ("Gneiss", 5.0), ("Dark Ochre", 8.0),
    ("Spodumain", 16.0), ("Crokite", 16.0), ("Bistot", 16.0),
    ("Arkonor", 16.0), ("Mercoxit", 40.0),
    
    # MOON R4
    ("Zeolites", 100.0), ("Sylvite", 100.0), ("Bitumens", 100.0), ("Coesite", 100.0),
    
    # MOON R8
    ("Cobaltite", 100.0), ("Euxenite", 100.0), ("Titanite", 100.0), ("Scheelite", 100.0),
    
    # MOON R16
    ("Otavite", 100.0), ("Sperrylite", 100.0), ("Vanadinite", 100.0), ("Chromite", 100.0),
    
    # MOON R32
    ("Carnotite", 100.0), ("Zircon", 100.0), ("Pollucite", 100.0), ("Cinnabar", 100.0),
    
    # MOON R64
    ("Xenotime", 100.0), ("Monazite", 100.0), ("Loparite", 100.0), ("Ytterbite", 100.0),
    
    # ICE ORES
    ("Blue Ice", 1000.0), ("Clear Icicle", 1000.0), ("Glacial Mass", 1000.0),
    ("White Glaze", 1000.0), ("Glare Crust", 1000.0), ("Dark Glitter", 1000.0),
    ("Gelidus", 1000.0), ("Krystallos", 1000.0),
    
    # POCHVEN ORES
    ("Bezdnacine", 16.0), ("Rakovene", 16.0), ("Talassonite", 16.0),
    
    # SPECIAL/RARE ORES
    ("Ueganite", 5.0),        # Abyssal ore (compact!)
    ("Prismaticite", 16.0),   # Rare ore
    ("Ducinium", 16.0),       # Drifter ore
    
    # GAS CLOUDS - CYTOSEROCIN
    ("Amber Cytoserocin", 10.0), ("Azure Cytoserocin", 10.0),
    ("Celadon Cytoserocin", 10.0), ("Golden Cytoserocin", 10.0),
    ("Lime Cytoserocin", 10.0), ("Vermillion Cytoserocin", 10.0),
    ("Viridian Cytoserocin", 10.0),
    
    # GAS CLOUDS - MYKOSEROCIN
    ("Amber Mykoserocin", 10.0), ("Azure Mykoserocin", 10.0),
    ("Celadon Mykoserocin", 10.0), ("Golden Mykoserocin", 10.0),
    ("Lime Mykoserocin", 10.0), ("Vermillion Mykoserocin", 10.0),
    ("Viridian Mykoserocin", 10.0),
    
    # GAS CLOUDS - FULLERITES
    ("Fullerite-C28", 5.0), ("Fullerite-C32", 5.0), ("Fullerite-C50", 5.0),
    ("Fullerite-C60", 5.0), ("Fullerite-C70", 5.0), ("Fullerite-C72", 5.0),
    ("Fullerite-C84", 10.0), ("Fullerite-C320", 10.0), ("Fullerite-C540", 10.0),
]

# Generate all ore variations
ORE_VOLUMES: Dict[str, float] = {}
GRADE_SUFFIXES = ["", " II-Grade", " III-Grade", " IV-Grade"]
for ore_name, volume in _ORE_DATA:
    suffixes = GRADE_SUFFIXES[:-1] if ore_name == "Mercoxit" else GRADE_SUFFIXES
    for suffix in suffixes:
        ORE_VOLUMES[f"{ore_name}{suffix}"] = volume

# --- COMPRESSION RATIOS ---
COMPRESSION_RATIOS = {
    # Standard ores
    "Veldspar": 100, "Scordite": 100, "Pyroxeres": 100,
    "Plagioclase": 100, "Omber": 100, "Kernite": 100,
    "Jaspet": 100, "Hemorphite": 100, "Hedbergite": 100,
    "Gneiss": 100, "Dark Ochre": 100,
    "Spodumain": 100, "Crokite": 100, "Bistot": 100,
    "Arkonor": 100, "Mercoxit": 100,
    
    # Moon ores
    "Zeolites": 100, "Sylvite": 100, "Bitumens": 100, "Coesite": 100,
    "Cobaltite": 100, "Euxenite": 100, "Titanite": 100, "Scheelite": 100,
    "Otavite": 100, "Sperrylite": 100, "Vanadinite": 100, "Chromite": 100,
    "Carnotite": 100, "Zircon": 100, "Pollucite": 100, "Cinnabar": 100,
    "Xenotime": 100, "Monazite": 100, "Loparite": 100, "Ytterbite": 100,
    
    # Special ores
    "Ueganite": 100, "Prismaticite": 100, "Ducinium": 100,
    "Bezdnacine": 100, "Rakovene": 100, "Talassonite": 100,
    
    # Ice (1:1 - already compressed)
    "Blue Ice": 1, "Clear Icicle": 1, "Glacial Mass": 1,
    "White Glaze": 1, "Glare Crust": 1, "Dark Glitter": 1,
    "Gelidus": 1, "Krystallos": 1,
    
    # Gas clouds
    "Amber Cytoserocin": 100, "Azure Cytoserocin": 100,
    "Celadon Cytoserocin": 100, "Golden Cytoserocin": 100,
    "Lime Cytoserocin": 100, "Vermillion Cytoserocin": 100,
    "Viridian Cytoserocin": 100,
    "Amber Mykoserocin": 100, "Azure Mykoserocin": 100,
    "Celadon Mykoserocin": 100, "Golden Mykoserocin": 100,
    "Lime Mykoserocin": 100, "Vermillion Mykoserocin": 100,
    "Viridian Mykoserocin": 100,
    "Fullerite-C28": 100, "Fullerite-C32": 100, "Fullerite-C50": 100,
    "Fullerite-C60": 100, "Fullerite-C70": 100, "Fullerite-C72": 100,
    "Fullerite-C84": 100, "Fullerite-C320": 100, "Fullerite-C540": 100,
}

# Pre-compiled regex patterns - OPTIMIZED
MINING_LINE = re.compile(r'^\[.*?\]\s+\(mining\)', re.IGNORECASE)

REGULAR_MINE_PATTERN = re.compile(
    r"You mined <font size=12><color=[^>]+>(?P<amount>\d+)<color=[^>]+><font size=10> units of <color=[^>]+><font size=12>(?P<ore_type>[^\r\n<]+)",
    re.IGNORECASE
)

CRIT_MINE_PATTERN = re.compile(
    r"You mined an additional <color=[^>]+><font size=12>(?P<amount>\d+)<color=[^>]+><font size=10> units of <color=[^>]+><font size=12>(?P<ore_type>[^\r\n<]+)",
    re.IGNORECASE | re.DOTALL
)

COMPRESSION_PATTERN = re.compile(
    r'Successfully compressed (?P<ore_type>[^\s]+) into (?P<amount>[\d,]+) Compressed',
    re.IGNORECASE
)

# Character detection patterns (split on '_', char ID is 3rd segment)
LISTENER_LINE = re.compile(r'Listener:\s*(.+)', re.IGNORECASE)

# Timestamp pattern for date extraction from log lines
LOG_TIMESTAMP = re.compile(r'^\[\s*(\d{4}\.\d{2}\.\d{2})\s+\d{2}:\d{2}:\d{2}\s*\]')

# --- ORE CATEGORY COLORS FOR EXCEL ---
# Maps ore base name to a hex color for Excel styling
_ORE_CATEGORIES = {
    # Standard (green)
    "Veldspar": "2ecc40", "Scordite": "2ecc40", "Pyroxeres": "2ecc40",
    "Plagioclase": "2ecc40", "Omber": "2ecc40", "Kernite": "2ecc40",
    # Low-sec (yellow)
    "Jaspet": "f1c40f", "Hemorphite": "f1c40f", "Hedbergite": "f1c40f",
    # Null-sec (orange)
    "Gneiss": "ff9f43", "Dark Ochre": "ff9f43", "Spodumain": "ff9f43",
    "Crokite": "ff9f43", "Bistot": "ff9f43", "Arkonor": "ff9f43", "Mercoxit": "cc3325",
    # Moon R4 (light purple)
    "Zeolites": "a29bfe", "Sylvite": "a29bfe", "Bitumens": "a29bfe", "Coesite": "a29bfe",
    # Moon R8 (purple)
    "Cobaltite": "9b59b6", "Euxenite": "9b59b6", "Titanite": "9b59b6", "Scheelite": "9b59b6",
    # Moon R16 (magenta)
    "Otavite": "e056fd", "Sperrylite": "e056fd", "Vanadinite": "e056fd", "Chromite": "e056fd",
    # Moon R32 (hot pink)
    "Carnotite": "fd79a8", "Zircon": "fd79a8", "Pollucite": "fd79a8", "Cinnabar": "fd79a8",
    # Moon R64 (gold)
    "Xenotime": "ffd700", "Monazite": "ffd700", "Loparite": "ffd700", "Ytterbite": "ffd700",
    # Ice (ice blue)
    "Blue Ice": "74b9ff", "Clear Icicle": "74b9ff", "Glacial Mass": "74b9ff",
    "White Glaze": "74b9ff", "Glare Crust": "74b9ff", "Dark Glitter": "74b9ff",
    "Gelidus": "74b9ff", "Krystallos": "74b9ff",
    # Pochven (teal)
    "Bezdnacine": "00cec9", "Rakovene": "00cec9", "Talassonite": "00cec9",
    # Special (bright gold)
    "Ueganite": "ffeaa7", "Prismaticite": "ffeaa7", "Ducinium": "ffeaa7",
}

def _get_ore_excel_color(ore_name: str) -> str:
    # Return hex color for an ore name (checks base name match)
    for base_name, color in _ORE_CATEGORIES.items():
        if base_name.lower() in ore_name.lower():
            return color
    # Gas clouds
    if "cytoserocin" in ore_name.lower() or "mykoserocin" in ore_name.lower():
        return "55efc4"
    if "fullerite" in ore_name.lower():
        return "00b894"
    return "ffffff"

class MiningModule:
    def __init__(self, name: str = "", yield_per_cycle: float = 0.0, cycle_time: float = 0.0, enabled: bool = True):
        self.name = name
        self.yield_per_cycle = yield_per_cycle
        self.cycle_time = cycle_time
        self.enabled = enabled

    def get_m3_per_sec(self) -> float:
        if self.yield_per_cycle > 0 and self.cycle_time > 0:
            return self.yield_per_cycle / self.cycle_time
        return 0.0

    def is_configured(self) -> bool:
        return self.yield_per_cycle > 0 and self.cycle_time > 0

    def to_dict(self) -> Dict:
        return {
            "name": self.name,
            "yield_per_cycle": self.yield_per_cycle,
            "cycle_time": self.cycle_time,
            "enabled": self.enabled
        }

    @staticmethod
    def from_dict(data: Dict) -> 'MiningModule':
        return MiningModule(
            name=data.get("name", ""),
            yield_per_cycle=data.get("yield_per_cycle", 0.0),
            cycle_time=data.get("cycle_time", 0.0),
            enabled=data.get("enabled", True)
        )

class CharacterTracker:
    def __init__(self, char_id: str, char_name: str):
        self.char_id = char_id
        self.char_name = char_name
        self.log_path: Optional[str] = None
        self.log_pos: int = 0
        self.crit_count: int = 0
        self.total_m3: float = 0.0
        self.ore_summary: Dict[str, float] = {}
        self.compression_log: Dict[str, float] = {}
        
        # NEW: Multiple ship profiles support
        self.ship_profiles: Dict[str, List[MiningModule]] = {"Default": []}
        self.active_profile: str = "Default"
        
        self.session_start_time: float = time.time()
        self.session_start_m3: float = 0.0
        self.session_active: bool = False

    def get_active_modules(self) -> List[MiningModule]:
        # Get the modules for the currently active profile
        return self.ship_profiles.get(self.active_profile, [])

    def set_active_modules(self, modules: List[MiningModule]):
        # Set modules for the currently active profile
        self.ship_profiles[self.active_profile] = modules

    def get_total_theoretical_m3_per_sec(self) -> float:
        total = 0.0
        for module in self.get_active_modules():
            if module.enabled and module.is_configured():
                total += module.get_m3_per_sec()
        return total

    def get_active_module_count(self) -> int:
        return sum(1 for m in self.get_active_modules() if m.enabled and m.is_configured())

    def has_any_configured_module(self) -> bool:
        return any(m.is_configured() for m in self.get_active_modules())

    def get_profile_names(self) -> List[str]:
        # Get list of all profile names
        return list(self.ship_profiles.keys())

    def create_profile(self, name: str):
        # Create a new empty profile
        if name and name not in self.ship_profiles:
            self.ship_profiles[name] = []
            return True
        return False

    def delete_profile(self, name: str) -> bool:
        # Delete a profile (cannot delete if it's the only one)
        if name in self.ship_profiles and len(self.ship_profiles) > 1:
            if self.active_profile == name:
                # Switch to another profile before deleting
                for profile_name in self.ship_profiles:
                    if profile_name != name:
                        self.active_profile = profile_name
                        break
            del self.ship_profiles[name]
            return True
        return False

    def rename_profile(self, old_name: str, new_name: str) -> bool:
        # Rename a profile
        if old_name in self.ship_profiles and new_name and new_name not in self.ship_profiles:
            self.ship_profiles[new_name] = self.ship_profiles.pop(old_name)
            if self.active_profile == old_name:
                self.active_profile = new_name
            return True
        return False

class MiningDashboard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EVE Mining Dashboard")
        self.root.attributes("-topmost", True)
        self.root.configure(bg=BORDER)
        self.root.overrideredirect(True)
        self.root.attributes("-alpha", 0.85)
        self.root.resizable(False, False)

        self._drag_x = 0
        self._drag_y = 0

        # Load config
        self.app_config = self.load_config()
        self._apply_saved_app_settings()

        # Set position from saved geometry
        saved_geom = self.app_config.get("win_geom", "+100+100")
        try:
            if '+' in saved_geom:
                parts = saved_geom.split('+')
                if len(parts) >= 3:
                    self.root.geometry(f"+{parts[1]}+{parts[2]}")
                else:
                    self.root.geometry("+100+100")
            else:
                self.root.geometry("+100+100")
        except:
            self.root.geometry("+100+100")

        # Glob cache to avoid scanning gamelogs directory every tick
        self._glob_cache: List[str] = []
        self._glob_cache_time: float = 0.0
        self._glob_cache_ttl: float = 5.0  # refresh every 5 seconds

        # Discover characters from gamelogs
        self.all_characters = self.discover_all_characters()

        # Filter to only visible characters
        self.characters = self.get_visible_characters()

        # Load ship fitting data from config
        self.load_ship_configs()

        # Initialize log tracking per character
        for tracker in self.all_characters.values():
            tracker.log_path = self._get_latest_log_for_char(tracker.char_id)
            if tracker.log_path:
                tracker.log_pos = os.path.getsize(tracker.log_path)

        # Per-character UI widgets
        self.char_widgets: Dict[str, Dict] = {}

        # Store reference to chars_container for rebuilding
        self.chars_container = None

        self.setup_ui()

        # Initialize windows
        self.history_window = None
        self.ship_config_dialogs: Dict[str, tk.Toplevel] = {}
        self.config_dialog: Optional[tk.Toplevel] = None

        # Flag to control update loop
        self.update_loop_running = True

        # Bind drag events
        self.root.bind("<Button-1>", self._start_drag)
        self.root.bind("<B1-Motion>", self._do_drag)

        self.update_loop()
        self.root.mainloop()

    # --------------------- #
    #  CHARACTER DISCOVERY  #
    # --------------------- #

    def discover_all_characters(self) -> Dict[str, CharacterTracker]:
        char_names: Dict[str, str] = {}
        char_counts: Dict[str, int] = {}

        for filepath in self._get_all_log_files():
            char_id = self._get_char_id_from_file(filepath)
            if char_id:
                char_counts[char_id] = char_counts.get(char_id, 0) + 1
                if char_id not in char_names:
                    name = self._read_listener_name(filepath)
                    if name:
                        char_names[char_id] = name

        sorted_ids = sorted(
            char_names.keys(),
            key=lambda cid: char_counts.get(cid, 0),
            reverse=True
        )

        result: Dict[str, CharacterTracker] = {}
        for char_id in sorted_ids:
            result[char_id] = CharacterTracker(char_id, char_names[char_id])
        return result

    def get_visible_characters(self) -> Dict[str, CharacterTracker]:
        visible_chars = self.app_config.get("visible_characters", [])
        if not visible_chars:
            return self.all_characters.copy()
        result = {}
        for char_id, tracker in self.all_characters.items():
            if char_id in visible_chars:
                result[char_id] = tracker
        return result

    def save_visible_characters(self, visible_char_ids: List[str]):
        self.app_config["visible_characters"] = visible_char_ids
        self.save_config()
        self.characters = self.get_visible_characters()
        self.rebuild_dashboard()

    def rebuild_dashboard(self):
        if self.chars_container:
            for widget in self.chars_container.winfo_children():
                widget.destroy()
            self.char_widgets.clear()

        if not self.characters:
            tk.Label(
                self.chars_container,
                text="No characters selected\nClick ⚙ to select characters",
                fg=DIM, bg=BG,
                font=("Consolas", 9),
                justify="center"
            ).pack(pady=20)
        else:
            for i, (char_id, tracker) in enumerate(self.characters.items()):
                accent = CHAR_ACCENTS[i % len(CHAR_ACCENTS)]
                col_frame, widgets = self._create_char_column(self.chars_container, tracker, accent, char_id)
                col_frame.pack(side="left", fill="both", expand=True, padx=3)
                self.char_widgets[char_id] = widgets
                self.update_ship_indicator(char_id)

        self.root.update_idletasks()
        self.root.geometry("")

    def _get_all_log_files(self) -> List[str]:
        # Ensure DOCS ends with wildcard for glob
        pattern = DOCS.rstrip('\\').rstrip('/')
        if not pattern.endswith('*'):
            pattern = os.path.join(pattern, '*')
        main_files = glob.glob(pattern)
        old_pattern = os.path.join(os.path.dirname(pattern.rstrip('*').rstrip('\\').rstrip('/')), "OLD", "*")
        old_files = glob.glob(old_pattern)
        return [f for f in main_files + old_files if f.lower().endswith('.txt')]

    @staticmethod
    def _get_char_id_from_file(filepath: str) -> Optional[str]:
        # Split filename on '_' and check after the second underscore
        basename = os.path.splitext(os.path.basename(filepath))[0]
        parts = basename.split('_')
        if len(parts) >= 3:
            char_id = parts[2]
            if char_id.isdigit():
                return char_id
        return None

    @staticmethod
    def _read_listener_name(filepath: str) -> Optional[str]:
        try:
            with open(filepath, 'r', encoding='utf-8-sig', errors='ignore') as f:
                for i, line in enumerate(f):
                    if i > 15:
                        break
                    match = LISTENER_LINE.search(line)
                    if match:
                        return match.group(1).strip()
        except Exception:
            pass
        return None

    def _get_cached_log_files(self) -> List[str]:
        # Return cached glob results, refresh only after TTL expires
        now = time.time()
        if now - self._glob_cache_time > self._glob_cache_ttl:
            # Ensure DOCS ends with wildcard for glob
            pattern = DOCS.rstrip('\\').rstrip('/')
            if not pattern.endswith('*'):
                pattern = os.path.join(pattern, '*')
            self._glob_cache = [f for f in glob.glob(pattern) if f.lower().endswith('.txt')]
            self._glob_cache_time = now
        return self._glob_cache

    def _get_latest_log_for_char(self, char_id: str) -> Optional[str]:
        files = self._get_cached_log_files()
        char_files = [
            f for f in files
            if self._get_char_id_from_file(f) == char_id
        ]
        return max(char_files, key=os.path.getmtime) if char_files else None

    # ------ #
    #  DRAG  #
    # ------ #

    def _start_drag(self, event):
        widget = event.widget
        if isinstance(widget, tk.Button):
            return
        if isinstance(widget, tk.Label) and widget.cget("cursor") == "hand2":
            return
        self._drag_x = event.x
        self._drag_y = event.y

    def _do_drag(self, event):
        widget = event.widget
        if isinstance(widget, tk.Button):
            return
        if isinstance(widget, tk.Label) and widget.cget("cursor") == "hand2":
            return
        x = self.root.winfo_x() + event.x - self._drag_x
        y = self.root.winfo_y() + event.y - self._drag_y
        self.root.geometry(f"+{x}+{y}")

    # --------- #
    #  MAIN UI  #
    # --------- #

    def setup_ui(self) -> None:
        border_frame = tk.Frame(self.root, bg=BORDER, padx=1, pady=1)
        border_frame.pack(fill="both", expand=True)

        self.inner_frame = tk.Frame(border_frame, bg=BG)
        self.inner_frame.pack(fill="both", expand=True)

        # Top bar with title and buttons
        top_bar = tk.Frame(self.inner_frame, bg=BG, pady=8, padx=10)
        top_bar.pack(fill="x")

        tk.Label(
            top_bar,
            text="★ MINING DASHBOARD ★",
            fg=CYAN,
            bg=BG,
            font=("Consolas", 11, "bold")
        ).pack(side="left")

        # Close button
        close_btn = tk.Label(
            top_bar,
            text="✕",
            fg=DIM,
            bg=BG,
            font=("Consolas", 14, "bold"),
            cursor="hand2"
        )
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: self.on_close())
        close_btn.bind("<Enter>", lambda e: close_btn.config(fg=RED))
        close_btn.bind("<Leave>", lambda e: close_btn.config(fg=DIM))

        # Config gear icon
        self.config_icon = tk.Label(
            top_bar,
            text="⚙",
            fg=DIM,
            bg=BG,
            font=("Consolas", 13, "bold"),
            cursor="hand2"
        )
        self.config_icon.pack(side="right", padx=(0, 8))
        self.config_icon.bind("<Button-1>", lambda e: self.show_config_dialog())
        self.config_icon.bind("<Enter>", lambda e: self.config_icon.config(fg=CYAN))
        self.config_icon.bind("<Leave>", lambda e: self.config_icon.config(fg=DIM))

        # Character columns container
        self.chars_container = tk.Frame(self.inner_frame, bg=BG)
        self.chars_container.pack(fill="both", padx=5, pady=(5, 0))

        if not self.characters:
            tk.Label(
                self.chars_container,
                text="No characters selected\nClick ⚙ to select characters",
                fg=DIM, bg=BG,
                font=("Consolas", 9),
                justify="center"
            ).pack(pady=20)
        else:
            for i, (char_id, tracker) in enumerate(self.characters.items()):
                accent = CHAR_ACCENTS[i % len(CHAR_ACCENTS)]
                col_frame, widgets = self._create_char_column(self.chars_container, tracker, accent, char_id)
                col_frame.pack(side="left", fill="both", expand=True, padx=3)
                self.char_widgets[char_id] = widgets
                self.update_ship_indicator(char_id)

        # History button
        self.history_button = tk.Button(
            self.inner_frame,
            text="◈ HISTORY",
            command=self.show_history,
            bg=BG_PANEL,
            fg=CYAN,
            font=("Consolas", 9, "bold"),
            relief="flat",
            cursor="hand2",
            activebackground=BORDER,
            activeforeground=CYAN
        )
        self.history_button.pack(fill="x", padx=20, pady=(12, 15))

    def _create_char_column(self, parent, tracker: CharacterTracker, accent_color: str, char_id: str):
        col_outer = tk.Frame(parent, bg=BORDER, padx=1, pady=1)
        col_inner = tk.Frame(col_outer, bg=BG_PANEL, padx=10, pady=8)
        col_inner.pack(fill="both", expand=True)

        def show_context_menu(event):
            context_menu = tk.Menu(self.root, tearoff=0, bg=BG_PANEL, fg=WHITE,
                                   activebackground=BORDER, activeforeground=CYAN,
                                   relief="flat", bd=1)
            context_menu.add_command(label="⚙ Ship Config",
                                     command=lambda: self.show_ship_config(char_id))
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

        col_inner.bind("<Button-3>", show_context_menu)

        # Character name header with ship config indicator
        name_frame = tk.Frame(col_inner, bg=BG_PANEL)
        name_frame.pack(fill="x", pady=(0, 5))
        name_frame.bind("<Button-3>", show_context_menu)

        char_name_label = tk.Label(
            name_frame,
            text=tracker.char_name.upper(),
            fg=accent_color,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        )
        char_name_label.pack(side="left")
        char_name_label.bind("<Button-3>", show_context_menu)

        # Active profile indicator
        profile_label = tk.Label(
            name_frame,
            text=f"〈{tracker.active_profile}〉",
            fg=GOLD,
            bg=BG_PANEL,
            font=("Consolas", 8)
        )
        profile_label.pack(side="left", padx=(5, 0))
        profile_label.bind("<Button-3>", show_context_menu)

        # Ship config indicator
        ship_indicator = tk.Label(
            name_frame,
            text="◆",
            fg=DIM,
            bg=BG_PANEL,
            font=("Consolas", 10, "bold")
        )
        ship_indicator.pack(side="right")
        ship_indicator.bind("<Button-3>", show_context_menu)

        # Stats frame
        stats_frame = tk.Frame(col_inner, bg=BG_PANEL)
        stats_frame.pack(fill="x")
        stats_frame.bind("<Button-3>", show_context_menu)

        crit_label = tk.Label(
            stats_frame,
            text="Crits: 0",
            fg=GOLD,
            bg=BG_PANEL,
            font=("Consolas", 11, "bold")
        )
        crit_label.pack(anchor="w", pady=2)
        crit_label.bind("<Button-3>", show_context_menu)

        ore_label = tk.Label(
            stats_frame,
            text="Total: 0.0 m3",
            fg=GREEN,
            bg=BG_PANEL,
            font=("Consolas", 11, "bold")
        )
        ore_label.pack(anchor="w", pady=2)
        ore_label.bind("<Button-3>", show_context_menu)

        # Control buttons frame
        control_frame = tk.Frame(col_inner, bg=BG_PANEL)
        control_frame.pack(fill="x", pady=(5, 0))
        control_frame.bind("<Button-3>", show_context_menu)

        start_stop_btn = tk.Button(
            control_frame,
            text="▶ START",
            command=lambda: self.toggle_session(char_id),
            bg=BG,
            fg=GREEN,
            font=("Consolas", 8, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        )
        start_stop_btn.pack(side="left", padx=(0, 5))

        reset_btn = tk.Button(
            control_frame,
            text="↺ RESET",
            command=lambda: self.reset_session(char_id),
            bg=BG,
            fg=RED,
            font=("Consolas", 8, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        )
        reset_btn.pack(side="left")

        # Mining rate stats
        rate_frame = tk.Frame(col_inner, bg=BG_PANEL)
        rate_frame.pack(fill="x", pady=(5, 0))
        rate_frame.bind("<Button-3>", show_context_menu)

        theoretical_label = tk.Label(
            rate_frame,
            text="◈ Theoretical: -- m3/s",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9)
        )
        theoretical_label.pack(anchor="w", pady=1)
        theoretical_label.bind("<Button-3>", show_context_menu)

        actual_label = tk.Label(
            rate_frame,
            text="◉ Actual: -- m3/s",
            fg=accent_color,
            bg=BG_PANEL,
            font=("Consolas", 9)
        )
        actual_label.pack(anchor="w", pady=1)
        actual_label.bind("<Button-3>", show_context_menu)

        # Section separator
        separator = tk.Label(
            col_inner,
            text="── SESSION BREAKDOWN ──",
            fg=accent_color,
            bg=BG_PANEL,
            font=("Consolas", 8, "bold")
        )
        separator.pack(pady=(5, 3))
        separator.bind("<Button-3>", show_context_menu)

        # Summary box
        summary_outer = tk.Frame(col_inner, bg=BORDER, padx=1, pady=1)
        summary_outer.pack(fill="both", pady=(0, 3))
        summary_outer.bind("<Button-3>", show_context_menu)

        summary_box = tk.Label(
            summary_outer,
            text="Waiting...",
            fg=WHITE,
            bg=BG_PANEL,
            font=("Consolas", 9),
            justify="left",
            padx=8,
            pady=8
        )
        summary_box.pack(fill="both")
        summary_box.bind("<Button-3>", show_context_menu)

        widgets = {
            'crit': crit_label,
            'ore': ore_label,
            'summary': summary_box,
            'theoretical': theoretical_label,
            'actual': actual_label,
            'start_stop_btn': start_stop_btn,
            'ship_indicator': ship_indicator,
            'profile_label': profile_label
        }

        return col_outer, widgets

    # ---------------- #
    #  HISTORY WINDOW  #
    # ---------------- #

    def show_history(self) -> None:
        if self.history_window is None or not self.history_window.winfo_exists():
            self.history_button.config(state="disabled")
            self.history_window = tk.Toplevel(self.root)
            self.history_window.overrideredirect(True)
            self.history_window.configure(bg=BORDER)
            self.history_window.attributes("-topmost", True)
            self.history_window.attributes("-alpha", 0.85)

            self._history_drag_x = 0
            self._history_drag_y = 0

            def history_start_drag(event):
                if isinstance(event.widget, tk.Entry):
                    return
                self._history_drag_x = event.x
                self._history_drag_y = event.y

            def history_do_drag(event):
                if isinstance(event.widget, tk.Entry):
                    return
                x = self.history_window.winfo_x() + event.x - self._history_drag_x
                y = self.history_window.winfo_y() + event.y - self._history_drag_y
                self.history_window.geometry(f"+{x}+{y}")

            self.history_window.bind("<Button-1>", history_start_drag)
            self.history_window.bind("<B1-Motion>", history_do_drag)

            saved_geom = self.app_config.get("history_win_geom", "+400+100")
            try:
                if '+' in saved_geom:
                    parts = saved_geom.split('+')
                    if len(parts) >= 3:
                        self.history_window.geometry(f"+{parts[1]}+{parts[2]}")
            except:
                pass

            border_frame = tk.Frame(self.history_window, bg=BORDER, padx=1, pady=1)
            border_frame.pack(fill="both", expand=True)

            inner_frame = tk.Frame(border_frame, bg=BG)
            inner_frame.pack(fill="both", expand=True)

            top_bar = tk.Frame(inner_frame, bg=BG, pady=10, padx=10)
            top_bar.pack(fill="x")

            tk.Label(
                top_bar,
                text="★ MINING HISTORY ★",
                fg=CYAN,
                bg=BG,
                font=("Consolas", 12, "bold")
            ).pack(side="left")

            close_btn = tk.Label(
                top_bar,
                text="X",
                fg=DIM,
                bg=BG,
                font=("Consolas", 14, "bold"),
                cursor="hand2"
            )
            close_btn.pack(side="right")
            close_btn.bind("<Button-1>", lambda e: self.on_history_close())
            close_btn.bind("<Enter>", lambda e: close_btn.config(fg=RED))
            close_btn.bind("<Leave>", lambda e: close_btn.config(fg=DIM))

            control_outer = tk.Frame(inner_frame, bg=BORDER, padx=1, pady=1)
            control_outer.pack(fill="x", padx=10, pady=10)

            control_frame = tk.Frame(control_outer, bg=BG_PANEL, padx=12, pady=12)
            control_frame.pack(fill="x")

            days_frame = tk.Frame(control_frame, bg=BG_PANEL)
            days_frame.pack(fill="x", pady=(0, 10))

            tk.Label(
                days_frame,
                text="◆ Days to analyze:",
                fg=CYAN,
                bg=BG_PANEL,
                font=("Consolas", 9, "bold")
            ).pack(side="left", padx=(0, 10))

            max_days = self.get_max_history_days()

            self.history_days_var = tk.StringVar(value=str(HISTORY_DAYS))
            days_entry = tk.Entry(
                days_frame,
                textvariable=self.history_days_var,
                width=10,
                font=("Consolas", 10),
                bg=BG,
                fg=WHITE,
                insertbackground=CYAN,
                relief="flat",
                justify="center"
            )
            days_entry.pack(side="left", padx=5)
            days_entry.bind("<Return>", lambda e: self.calculate_and_display_history(text_widget))

            tk.Label(
                days_frame,
                text=f"(max: {HISTORY_DAYS})",
                fg=GOLD,
                bg=BG_PANEL,
                font=("Consolas", 9)
            ).pack(side="left", padx=5)

            refresh_button = tk.Button(
                control_frame,
                text="↻ REFRESH STATS",
                command=lambda: self.calculate_and_display_history(text_widget),
                bg=BG,
                fg=GREEN,
                font=("Consolas", 9, "bold"),
                relief="flat",
                cursor="hand2",
                activebackground=BORDER,
                activeforeground=GREEN
            )
            refresh_button.pack(side="left", fill="x", expand=True)

            export_button = tk.Button(
                control_frame,
                text="◈ EXPORT EXCEL",
                command=lambda: self.show_export_menu(export_button),
                bg=BG,
                fg=GOLD,
                font=("Consolas", 9, "bold"),
                relief="flat",
                cursor="hand2",
                activebackground=BORDER,
                activeforeground=GOLD,
                state="normal" if HAS_OPENPYXL else "disabled"
            )
            export_button.pack(side="left", fill="x", expand=True, padx=(5, 0))

            text_outer = tk.Frame(inner_frame, bg=BORDER, padx=1, pady=1)
            text_outer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            text_frame = tk.Frame(text_outer, bg=BG_PANEL)
            text_frame.pack(fill="both", expand=True)

            text_widget = tk.Text(
                text_frame,
                bg=BG_PANEL,
                fg=WHITE,
                font=("Consolas", 10),
                relief="flat",
                padx=10,
                pady=10,
                wrap="word",
                width=40,
                height=20
            )
            text_widget.pack(fill="both", expand=True)

            self.calculate_and_display_history(text_widget)

    def get_max_history_days(self) -> int:
        try:
            all_files = self._get_all_log_files()
            if not all_files:
                return 0
            oldest_file = min(all_files, key=os.path.getmtime)
            oldest_timestamp = os.path.getmtime(oldest_file)
            oldest_date = datetime.fromtimestamp(oldest_timestamp)
            days_available = (datetime.now() - oldest_date).days
            return max(1, days_available)
        except Exception:
            return HISTORY_DAYS

    def on_history_close(self):
        if self.history_window:
            self.app_config["history_win_geom"] = self.history_window.geometry()
            self.save_config()
            self.history_window.destroy()
            self.history_window = None
            self.history_button.config(state="normal")

    def calculate_and_display_history(self, text_widget: tk.Text):
        text_widget.config(state="normal")
        text_widget.delete("1.0", tk.END)

        try:
            days = int(self.history_days_var.get())
            if days < 1:
                days = 1
            max_days = self.get_max_history_days()
            if days > max_days:
                days = max_days
            self.history_days_var.set(str(days))
        except ValueError:
            days = HISTORY_DAYS
            self.history_days_var.set(str(days))

        threshold = datetime.now() - timedelta(days=days)

        per_char_ores: Dict[str, Dict[str, float]] = {}
        per_char_m3: Dict[str, float] = {}
        combined_m3 = 0.0

        all_files = self._get_all_log_files()
        for log_file in all_files:
            if os.path.getmtime(log_file) > threshold.timestamp():
                char_id = self._get_char_id_from_file(log_file)
                if not char_id or char_id not in self.all_characters:
                    continue
                if char_id not in per_char_ores:
                    per_char_ores[char_id] = {}
                    per_char_m3[char_id] = 0.0
                try:
                    with open(log_file, "r", encoding="utf-8-sig", errors="ignore") as f:
                        for line in f:
                            match = REGULAR_MINE_PATTERN.search(line) or CRIT_MINE_PATTERN.search(line)
                            if match:
                                units = float(match.group('amount').replace(",", ""))
                                volume, ore_name = self.get_ore_volume(match.group('ore_type'))
                                total_volume = units * volume
                                per_char_ores[char_id][ore_name] = per_char_ores[char_id].get(ore_name, 0) + total_volume
                                per_char_m3[char_id] = per_char_m3.get(char_id, 0) + total_volume
                                combined_m3 += total_volume
                except Exception:
                    continue

        w = 38
        result = ""
        total_str = f" ALL CHARS ({days}d): {combined_m3:,.1f} m3"
        pad = max(0, w - len(total_str))
        result += f"+{'=' * w}+\n"
        result += f"|{total_str}{' ' * pad}|\n"
        result += f"+{'=' * w}+\n\n"

        has_any_data = False
        for char_id, tracker in self.all_characters.items():
            char_name = tracker.char_name.upper()
            char_total = per_char_m3.get(char_id, 0.0)
            ores = per_char_ores.get(char_id, {})

            header = f" {char_name}: {char_total:,.1f} m3 "
            dashes = max(0, w - len(header)) // 2
            result += f"{'-' * dashes}{header}{'-' * dashes}\n"

            if ores:
                has_any_data = True
                for ore_name, volume in sorted(ores.items(), key=lambda x: x[1], reverse=True):
                    result += f"  * {ore_name}: {volume:,.1f} m3\n"
            else:
                result += "  No mining data.\n"
            result += "\n"

        if not has_any_data and not self.all_characters:
            result += "No mining data found in this period.\n"

        text_widget.insert("1.0", result)
        text_widget.config(state="disabled")

    # ========================= #
    #     EXCEL EXPORT         #
    # ========================= #

    def _gather_history_data(self, days: int):
        # Collect mining data from logs for the given number of days
        # Returns: (per_char_ores, per_char_m3, combined_m3, days_used)
        try:
            days = int(days)
            if days < 1:
                days = 1
            max_days = self.get_max_history_days()
            if days > max_days:
                days = max_days
        except ValueError:
            days = HISTORY_DAYS

        threshold = datetime.now() - timedelta(days=days)

        per_char_ores: Dict[str, Dict[str, float]] = {}
        per_char_m3: Dict[str, float] = {}
        combined_m3 = 0.0

        all_files = self._get_all_log_files()
        for log_file in all_files:
            if os.path.getmtime(log_file) > threshold.timestamp():
                char_id = self._get_char_id_from_file(log_file)
                if not char_id or char_id not in self.all_characters:
                    continue
                if char_id not in per_char_ores:
                    per_char_ores[char_id] = {}
                    per_char_m3[char_id] = 0.0
                try:
                    with open(log_file, "r", encoding="utf-8-sig", errors="ignore") as f:
                        for line in f:
                            match = REGULAR_MINE_PATTERN.search(line) or CRIT_MINE_PATTERN.search(line)
                            if match:
                                units = float(match.group('amount').replace(",", ""))
                                volume, ore_name = self.get_ore_volume(match.group('ore_type'))
                                total_volume = units * volume
                                per_char_ores[char_id][ore_name] = per_char_ores[char_id].get(ore_name, 0) + total_volume
                                per_char_m3[char_id] = per_char_m3.get(char_id, 0) + total_volume
                                combined_m3 += total_volume
                except Exception:
                    continue

        return per_char_ores, per_char_m3, combined_m3, days

    def _gather_daily_history_data(self, days: int):
        # Collect mining data with daily breakdown
        # Returns: (per_char_daily_ores, all_ore_names, all_dates, days_used)
        # per_char_daily_ores[char_id][date_str][ore_name] = m3
        try:
            days = int(days)
            if days < 1:
                days = 1
            max_days = self.get_max_history_days()
            if days > max_days:
                days = max_days
        except ValueError:
            days = HISTORY_DAYS

        threshold = datetime.now() - timedelta(days=days)

        per_char_daily: Dict[str, Dict[str, Dict[str, float]]] = {}
        all_ore_names = set()
        all_dates = set()

        all_files = self._get_all_log_files()
        for log_file in all_files:
            if os.path.getmtime(log_file) > threshold.timestamp():
                char_id = self._get_char_id_from_file(log_file)
                if not char_id or char_id not in self.all_characters:
                    continue
                if char_id not in per_char_daily:
                    per_char_daily[char_id] = {}
                try:
                    with open(log_file, "r", encoding="utf-8-sig", errors="ignore") as f:
                        for line in f:
                            match = REGULAR_MINE_PATTERN.search(line) or CRIT_MINE_PATTERN.search(line)
                            if match:
                                # Extract date from timestamp
                                ts_match = LOG_TIMESTAMP.match(line)
                                if ts_match:
                                    date_str = ts_match.group(1).replace(".", "-")
                                else:
                                    continue
                                
                                units = float(match.group('amount').replace(",", ""))
                                volume, ore_name = self.get_ore_volume(match.group('ore_type'))
                                total_volume = units * volume
                                
                                all_ore_names.add(ore_name)
                                all_dates.add(date_str)
                                
                                if date_str not in per_char_daily[char_id]:
                                    per_char_daily[char_id][date_str] = {}
                                per_char_daily[char_id][date_str][ore_name] = (
                                    per_char_daily[char_id][date_str].get(ore_name, 0) + total_volume
                                )
                except Exception:
                    continue

        sorted_dates = sorted(all_dates)
        sorted_ores = sorted(all_ore_names)
        return per_char_daily, sorted_ores, sorted_dates, days

    def _get_export_path(self, suffix: str, days: int) -> str:
        # Generate export filepath with auto-naming
        export_dir = self.app_config.get("app_settings", {}).get("export_dir", "")
        if not export_dir or not os.path.isdir(export_dir):
            export_dir = os.path.dirname(os.path.abspath(CONFIG_FILE))
        
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        filename = f"mining_{suffix}_{timestamp}_{days}d.xlsx"
        return os.path.join(export_dir, filename)

    def show_export_menu(self, button_widget):
        # Show export type selection popup
        if not HAS_OPENPYXL:
            messagebox.showwarning("Missing Library",
                "openpyxl is required for Excel export.\n\npip install openpyxl")
            return

        menu = tk.Menu(self.root, tearoff=0, bg=BG_PANEL, fg=WHITE,
                       activebackground=BORDER, activeforeground=CYAN,
                       relief="flat", bd=1, font=("Consolas", 9))
        
        menu.add_command(label="◆ Summary Export",
                         command=lambda: self._do_export("summary"))
        menu.add_command(label="◆ Daily Breakdown",
                         command=lambda: self._do_export("daily"))
        menu.add_command(label="◆ Ore Pivot (Chars x Ores)",
                         command=lambda: self._do_export("pivot"))
        menu.add_separator()
        menu.add_command(label="◆ Full Export (All Sheets)",
                         command=lambda: self._do_export("full"))
        
        try:
            x = button_widget.winfo_rootx()
            y = button_widget.winfo_rooty() + button_widget.winfo_height()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _do_export(self, export_type: str):
        # Execute the selected export type
        try:
            days = int(self.history_days_var.get())
        except (ValueError, AttributeError):
            days = HISTORY_DAYS

        try:
            if export_type == "summary":
                filepath = self._export_summary(days)
            elif export_type == "daily":
                filepath = self._export_daily_breakdown(days)
            elif export_type == "pivot":
                filepath = self._export_ore_pivot(days)
            elif export_type == "full":
                filepath = self._export_full(days)
            else:
                return

            if filepath:
                messagebox.showinfo("Export Complete",
                    f"Saved to:\n{filepath}",
                    parent=self.history_window or self.root)
        except Exception as e:
            messagebox.showerror("Export Error",
                f"Failed to export:\n{str(e)}",
                parent=self.history_window or self.root)

    def _apply_eve_header(self, ws, row, col, text, width=None, is_title=False):
        # Apply EVE Online styled header cell
        cell = ws.cell(row=row, column=col, value=text)
        if is_title:
            cell.font = Font(name="Consolas", size=12, bold=True, color="3DD8E0")
            cell.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
        else:
            cell.font = Font(name="Consolas", size=10, bold=True, color="3DD8E0")
            cell.fill = PatternFill(start_color="111827", end_color="111827", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            bottom=Side(style="thin", color="1E3A4A"),
            top=Side(style="thin", color="1E3A4A"),
            left=Side(style="thin", color="1E3A4A"),
            right=Side(style="thin", color="1E3A4A")
        )
        if width:
            ws.column_dimensions[get_column_letter(col)].width = width
        return cell

    def _apply_eve_data_cell(self, ws, row, col, value, fmt="number", ore_name=None, is_total=False):
        # Apply EVE Online styled data cell
        cell = ws.cell(row=row, column=col, value=value)
        
        if is_total:
            cell.font = Font(name="Consolas", size=10, bold=True, color="FFD700")
            cell.fill = PatternFill(start_color="1A1A2E", end_color="1A1A2E", fill_type="solid")
        elif ore_name:
            ore_color = _get_ore_excel_color(ore_name)
            cell.font = Font(name="Consolas", size=10, color=ore_color)
            cell.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
        else:
            cell.font = Font(name="Consolas", size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")

        cell.border = Border(
            bottom=Side(style="thin", color="1E3A4A"),
            left=Side(style="thin", color="1E3A4A"),
            right=Side(style="thin", color="1E3A4A")
        )
        
        if fmt == "number" and isinstance(value, (int, float)):
            cell.number_format = '#,##0.0'
            cell.alignment = Alignment(horizontal="right")
        elif fmt == "integer" and isinstance(value, (int, float)):
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal="right")
        else:
            cell.alignment = Alignment(horizontal="left")
        
        return cell

    def _apply_eve_ore_label(self, ws, row, col, ore_name):
        # Apply ore name with category color
        cell = ws.cell(row=row, column=col, value=ore_name)
        ore_color = _get_ore_excel_color(ore_name)
        cell.font = Font(name="Consolas", size=10, bold=True, color=ore_color)
        cell.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
        cell.alignment = Alignment(horizontal="left")
        cell.border = Border(
            bottom=Side(style="thin", color="1E3A4A"),
            left=Side(style="thin", color="1E3A4A"),
            right=Side(style="thin", color="1E3A4A")
        )
        return cell

    def _style_eve_sheet(self, ws):
        # Apply dark background to entire visible area
        ws.sheet_properties.tabColor = "3DD8E0"

    def _export_summary(self, days: int) -> str:
        # Export: Summary - one sheet per character + combined
        per_char_ores, per_char_m3, combined_m3, days = self._gather_history_data(days)
        filepath = self._get_export_path("summary", days)

        wb = Workbook()
        
        # -- COMBINED SHEET --
        ws = wb.active
        ws.title = "Combined"
        self._style_eve_sheet(ws)
        
        # Title row
        self._apply_eve_header(ws, 1, 1, f"EVE MINING SUMMARY  --  {days} DAYS", width=30, is_title=True)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        
        self._apply_eve_header(ws, 2, 1, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", width=30)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)

        self._apply_eve_header(ws, 4, 1, "Character", width=22)
        self._apply_eve_header(ws, 4, 2, "Total m3", width=18)
        self._apply_eve_header(ws, 4, 3, "% Share", width=12)
        
        row = 5
        for char_id, tracker in self.all_characters.items():
            char_total = per_char_m3.get(char_id, 0.0)
            pct = (char_total / combined_m3 * 100) if combined_m3 > 0 else 0
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            cell_name = ws.cell(row=row, column=1, value=tracker.char_name.upper())
            cell_name.font = Font(name="Consolas", size=10, bold=True, color=accent_color)
            cell_name.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
            cell_name.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                       left=Side(style="thin", color="1E3A4A"),
                                       right=Side(style="thin", color="1E3A4A"))
            
            self._apply_eve_data_cell(ws, row, 2, char_total)
            
            pct_cell = self._apply_eve_data_cell(ws, row, 3, pct)
            pct_cell.number_format = '0.0"%"'
            
            row += 1
        
        # Total row
        row += 1
        total_label = self._apply_eve_data_cell(ws, row, 1, "TOTAL", is_total=True)
        total_label.alignment = Alignment(horizontal="left")
        self._apply_eve_data_cell(ws, row, 2, combined_m3, is_total=True)
        total_pct = self._apply_eve_data_cell(ws, row, 3, 100.0, is_total=True)
        total_pct.number_format = '0.0"%"'
        
        # -- PER-CHARACTER SHEETS --
        for char_id, tracker in self.all_characters.items():
            ores = per_char_ores.get(char_id, {})
            if not ores:
                continue
            
            # Clean sheet name (max 31 chars, no special chars)
            sheet_name = tracker.char_name[:28].replace("/", "-").replace("\\", "-")
            ws = wb.create_sheet(title=sheet_name)
            self._style_eve_sheet(ws)
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            title_cell = self._apply_eve_header(ws, 1, 1, 
                f"{tracker.char_name.upper()}  --  MINING BREAKDOWN", width=30, is_title=True)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
            title_cell.font = Font(name="Consolas", size=12, bold=True, color=accent_color)
            
            char_total = per_char_m3.get(char_id, 0.0)
            self._apply_eve_header(ws, 2, 1, f"Total: {char_total:,.1f} m3", width=30)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=2)
            
            self._apply_eve_header(ws, 4, 1, "Ore Type", width=28)
            self._apply_eve_header(ws, 4, 2, "Volume (m3)", width=18)
            
            row = 5
            sorted_ores = sorted(ores.items(), key=lambda x: x[1], reverse=True)
            for ore_name, volume in sorted_ores:
                self._apply_eve_ore_label(ws, row, 1, ore_name)
                self._apply_eve_data_cell(ws, row, 2, volume, ore_name=ore_name)
                row += 1
            
            # Total
            row += 1
            self._apply_eve_data_cell(ws, row, 1, "TOTAL", is_total=True)
            self._apply_eve_data_cell(ws, row, 2, char_total, is_total=True)
        
        wb.save(filepath)
        return filepath

    def _export_daily_breakdown(self, days: int) -> str:
        # Export: Daily breakdown - days as rows, ores as columns, per character
        per_char_daily, sorted_ores, sorted_dates, days = self._gather_daily_history_data(days)
        filepath = self._get_export_path("daily", days)

        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet

        for char_id, tracker in self.all_characters.items():
            daily_data = per_char_daily.get(char_id, {})
            if not daily_data:
                continue
            
            sheet_name = tracker.char_name[:28].replace("/", "-").replace("\\", "-")
            ws = wb.create_sheet(title=sheet_name)
            self._style_eve_sheet(ws)
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            # Get ores that this character actually mined
            char_ores = set()
            for date_ores in daily_data.values():
                char_ores.update(date_ores.keys())
            char_ores_sorted = sorted(char_ores)
            
            if not char_ores_sorted:
                continue
            
            # Title
            title_cell = self._apply_eve_header(ws, 1, 1,
                f"{tracker.char_name.upper()}  --  DAILY BREAKDOWN", is_title=True)
            ws.merge_cells(start_row=1, start_column=1, 
                          end_row=1, end_column=min(len(char_ores_sorted) + 2, 10))
            title_cell.font = Font(name="Consolas", size=12, bold=True, color=accent_color)
            
            # Headers: Date | Ore1 | Ore2 | ... | TOTAL
            self._apply_eve_header(ws, 3, 1, "Date", width=14)
            for j, ore_name in enumerate(char_ores_sorted):
                header_cell = self._apply_eve_header(ws, 3, j + 2, ore_name, width=16)
                ore_color = _get_ore_excel_color(ore_name)
                header_cell.font = Font(name="Consolas", size=9, bold=True, color=ore_color)
            total_col = len(char_ores_sorted) + 2
            self._apply_eve_header(ws, 3, total_col, "DAILY TOTAL", width=16)
            
            # Data rows
            row = 4
            grand_total = 0.0
            ore_totals = {ore: 0.0 for ore in char_ores_sorted}
            
            for date_str in sorted_dates:
                if date_str not in daily_data:
                    continue
                
                date_ores = daily_data[date_str]
                self._apply_eve_data_cell(ws, row, 1, date_str, fmt="text")
                
                day_total = 0.0
                for j, ore_name in enumerate(char_ores_sorted):
                    vol = date_ores.get(ore_name, 0.0)
                    if vol > 0:
                        self._apply_eve_data_cell(ws, row, j + 2, vol, ore_name=ore_name)
                        ore_totals[ore_name] += vol
                        day_total += vol
                    else:
                        # Empty cell with dark bg
                        empty = ws.cell(row=row, column=j + 2, value="")
                        empty.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
                        empty.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                              left=Side(style="thin", color="1E3A4A"),
                                              right=Side(style="thin", color="1E3A4A"))
                
                self._apply_eve_data_cell(ws, row, total_col, day_total, is_total=True)
                grand_total += day_total
                row += 1
            
            # Totals row
            row += 1
            self._apply_eve_data_cell(ws, row, 1, "TOTAL", is_total=True)
            for j, ore_name in enumerate(char_ores_sorted):
                self._apply_eve_data_cell(ws, row, j + 2, ore_totals[ore_name], is_total=True)
            self._apply_eve_data_cell(ws, row, total_col, grand_total, is_total=True)
        
        # If no sheets were created, add a placeholder
        if len(wb.sheetnames) == 0:
            ws = wb.create_sheet(title="No Data")
            ws.cell(row=1, column=1, value="No mining data found in this period.")
        
        wb.save(filepath)
        return filepath

    def _export_ore_pivot(self, days: int) -> str:
        # Export: Ore pivot - ores as rows, characters as columns
        per_char_ores, per_char_m3, combined_m3, days = self._gather_history_data(days)
        filepath = self._get_export_path("pivot", days)

        wb = Workbook()
        ws = wb.active
        ws.title = "Ore Pivot"
        self._style_eve_sheet(ws)

        # Collect all ore names across all characters
        all_ores = set()
        for ores in per_char_ores.values():
            all_ores.update(ores.keys())
        sorted_ores = sorted(all_ores)

        # Active characters (those with data)
        active_chars = [(cid, t) for cid, t in self.all_characters.items() 
                        if cid in per_char_ores and per_char_ores[cid]]

        if not active_chars or not sorted_ores:
            ws.cell(row=1, column=1, value="No mining data found.")
            wb.save(filepath)
            return filepath

        # Title
        self._apply_eve_header(ws, 1, 1, 
            f"EVE MINING ORE PIVOT  --  {days} DAYS", is_title=True)
        ws.merge_cells(start_row=1, start_column=1, 
                      end_row=1, end_column=len(active_chars) + 2)

        # Headers: Ore Type | Char1 | Char2 | ... | TOTAL
        self._apply_eve_header(ws, 3, 1, "Ore Type", width=28)
        for j, (char_id, tracker) in enumerate(active_chars):
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            header = self._apply_eve_header(ws, 3, j + 2, tracker.char_name.upper(), width=18)
            header.font = Font(name="Consolas", size=10, bold=True, color=accent_color)
        total_col = len(active_chars) + 2
        self._apply_eve_header(ws, 3, total_col, "TOTAL", width=18)

        # Data rows
        row = 4
        char_totals = {cid: 0.0 for cid, _ in active_chars}
        grand_total = 0.0

        for ore_name in sorted_ores:
            self._apply_eve_ore_label(ws, row, 1, ore_name)
            
            ore_row_total = 0.0
            for j, (char_id, tracker) in enumerate(active_chars):
                vol = per_char_ores.get(char_id, {}).get(ore_name, 0.0)
                if vol > 0:
                    self._apply_eve_data_cell(ws, row, j + 2, vol, ore_name=ore_name)
                    char_totals[char_id] += vol
                    ore_row_total += vol
                else:
                    empty = ws.cell(row=row, column=j + 2, value="")
                    empty.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
                    empty.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                          left=Side(style="thin", color="1E3A4A"),
                                          right=Side(style="thin", color="1E3A4A"))
            
            self._apply_eve_data_cell(ws, row, total_col, ore_row_total, is_total=True)
            grand_total += ore_row_total
            row += 1

        # Character totals row
        row += 1
        self._apply_eve_data_cell(ws, row, 1, "TOTAL", is_total=True)
        for j, (char_id, tracker) in enumerate(active_chars):
            self._apply_eve_data_cell(ws, row, j + 2, char_totals[char_id], is_total=True)
        self._apply_eve_data_cell(ws, row, total_col, grand_total, is_total=True)

        wb.save(filepath)
        return filepath

    def _export_full(self, days: int) -> str:
        # Export: Full - all sheets in one workbook
        per_char_ores, per_char_m3, combined_m3, days_used = self._gather_history_data(days)
        per_char_daily, sorted_ores_daily, sorted_dates, _ = self._gather_daily_history_data(days)
        filepath = self._get_export_path("full", days_used)

        wb = Workbook()
        
        # ============================
        # SHEET 1: SUMMARY (Combined)
        # ============================
        ws = wb.active
        ws.title = "Summary"
        self._style_eve_sheet(ws)
        
        self._apply_eve_header(ws, 1, 1, 
            f"EVE MINING FULL REPORT  --  {days_used} DAYS", width=30, is_title=True)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
        
        self._apply_eve_header(ws, 2, 1, 
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", width=30)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)

        self._apply_eve_header(ws, 4, 1, "Character", width=22)
        self._apply_eve_header(ws, 4, 2, "Total m3", width=18)
        self._apply_eve_header(ws, 4, 3, "% Share", width=12)
        
        row = 5
        for char_id, tracker in self.all_characters.items():
            char_total = per_char_m3.get(char_id, 0.0)
            pct = (char_total / combined_m3 * 100) if combined_m3 > 0 else 0
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            cell_name = ws.cell(row=row, column=1, value=tracker.char_name.upper())
            cell_name.font = Font(name="Consolas", size=10, bold=True, color=accent_color)
            cell_name.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
            cell_name.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                       left=Side(style="thin", color="1E3A4A"),
                                       right=Side(style="thin", color="1E3A4A"))
            
            self._apply_eve_data_cell(ws, row, 2, char_total)
            pct_cell = self._apply_eve_data_cell(ws, row, 3, pct)
            pct_cell.number_format = '0.0"%"'
            row += 1
        
        row += 1
        self._apply_eve_data_cell(ws, row, 1, "TOTAL", is_total=True)
        self._apply_eve_data_cell(ws, row, 2, combined_m3, is_total=True)
        total_pct = self._apply_eve_data_cell(ws, row, 3, 100.0, is_total=True)
        total_pct.number_format = '0.0"%"'
        
        # Per-character ore breakdown below summary
        row += 3
        for char_id, tracker in self.all_characters.items():
            ores = per_char_ores.get(char_id, {})
            if not ores:
                continue
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            sep_cell = ws.cell(row=row, column=1, value=f"--- {tracker.char_name.upper()} ---")
            sep_cell.font = Font(name="Consolas", size=10, bold=True, color=accent_color)
            sep_cell.fill = PatternFill(start_color="111827", end_color="111827", fill_type="solid")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1
            
            for ore_name, volume in sorted(ores.items(), key=lambda x: x[1], reverse=True):
                self._apply_eve_ore_label(ws, row, 1, ore_name)
                self._apply_eve_data_cell(ws, row, 2, volume, ore_name=ore_name)
                row += 1
            row += 1

        # ============================
        # SHEET 2: ORE PIVOT
        # ============================
        ws2 = wb.create_sheet(title="Ore Pivot")
        self._style_eve_sheet(ws2)
        
        all_ores_set = set()
        for ores in per_char_ores.values():
            all_ores_set.update(ores.keys())
        all_ores_sorted = sorted(all_ores_set)
        
        active_chars = [(cid, t) for cid, t in self.all_characters.items() 
                        if cid in per_char_ores and per_char_ores[cid]]

        if active_chars and all_ores_sorted:
            self._apply_eve_header(ws2, 1, 1, "ORE PIVOT", is_title=True)
            ws2.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=len(active_chars) + 2)
            
            self._apply_eve_header(ws2, 3, 1, "Ore Type", width=28)
            for j, (char_id, tracker) in enumerate(active_chars):
                accent_idx = list(self.all_characters.keys()).index(char_id)
                accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
                h = self._apply_eve_header(ws2, 3, j + 2, tracker.char_name.upper(), width=18)
                h.font = Font(name="Consolas", size=10, bold=True, color=accent_color)
            total_col = len(active_chars) + 2
            self._apply_eve_header(ws2, 3, total_col, "TOTAL", width=18)
            
            row2 = 4
            for ore_name in all_ores_sorted:
                self._apply_eve_ore_label(ws2, row2, 1, ore_name)
                ore_total = 0.0
                for j, (char_id, _) in enumerate(active_chars):
                    vol = per_char_ores.get(char_id, {}).get(ore_name, 0.0)
                    if vol > 0:
                        self._apply_eve_data_cell(ws2, row2, j + 2, vol, ore_name=ore_name)
                        ore_total += vol
                    else:
                        empty = ws2.cell(row=row2, column=j + 2, value="")
                        empty.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
                        empty.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                              left=Side(style="thin", color="1E3A4A"),
                                              right=Side(style="thin", color="1E3A4A"))
                self._apply_eve_data_cell(ws2, row2, total_col, ore_total, is_total=True)
                row2 += 1
            
            # Totals
            row2 += 1
            self._apply_eve_data_cell(ws2, row2, 1, "TOTAL", is_total=True)
            for j, (char_id, _) in enumerate(active_chars):
                self._apply_eve_data_cell(ws2, row2, j + 2, per_char_m3.get(char_id, 0.0), is_total=True)
            self._apply_eve_data_cell(ws2, row2, total_col, combined_m3, is_total=True)

        # ============================
        # SHEET 3+: DAILY PER CHAR
        # ============================
        for char_id, tracker in self.all_characters.items():
            daily_data = per_char_daily.get(char_id, {})
            if not daily_data:
                continue
            
            sheet_name = f"Daily-{tracker.char_name[:24]}".replace("/", "-").replace("\\", "-")
            ws3 = wb.create_sheet(title=sheet_name)
            self._style_eve_sheet(ws3)
            
            accent_idx = list(self.all_characters.keys()).index(char_id)
            accent_color = CHAR_ACCENTS[accent_idx % len(CHAR_ACCENTS)].lstrip("#")
            
            char_ores = set()
            for date_ores in daily_data.values():
                char_ores.update(date_ores.keys())
            char_ores_sorted = sorted(char_ores)
            
            if not char_ores_sorted:
                continue
            
            title_cell = self._apply_eve_header(ws3, 1, 1,
                f"{tracker.char_name.upper()} - DAILY", is_title=True)
            ws3.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=min(len(char_ores_sorted) + 2, 10))
            title_cell.font = Font(name="Consolas", size=12, bold=True, color=accent_color)
            
            self._apply_eve_header(ws3, 3, 1, "Date", width=14)
            for j, ore_name in enumerate(char_ores_sorted):
                h = self._apply_eve_header(ws3, 3, j + 2, ore_name, width=16)
                ore_color = _get_ore_excel_color(ore_name)
                h.font = Font(name="Consolas", size=9, bold=True, color=ore_color)
            tcol = len(char_ores_sorted) + 2
            self._apply_eve_header(ws3, 3, tcol, "DAILY TOTAL", width=16)
            
            row3 = 4
            ore_totals = {ore: 0.0 for ore in char_ores_sorted}
            grand = 0.0
            
            for date_str in sorted_dates:
                if date_str not in daily_data:
                    continue
                date_ores = daily_data[date_str]
                self._apply_eve_data_cell(ws3, row3, 1, date_str, fmt="text")
                
                day_total = 0.0
                for j, ore_name in enumerate(char_ores_sorted):
                    vol = date_ores.get(ore_name, 0.0)
                    if vol > 0:
                        self._apply_eve_data_cell(ws3, row3, j + 2, vol, ore_name=ore_name)
                        ore_totals[ore_name] += vol
                        day_total += vol
                    else:
                        empty = ws3.cell(row=row3, column=j + 2, value="")
                        empty.fill = PatternFill(start_color="0B0E17", end_color="0B0E17", fill_type="solid")
                        empty.border = Border(bottom=Side(style="thin", color="1E3A4A"),
                                              left=Side(style="thin", color="1E3A4A"),
                                              right=Side(style="thin", color="1E3A4A"))
                
                self._apply_eve_data_cell(ws3, row3, tcol, day_total, is_total=True)
                grand += day_total
                row3 += 1
            
            row3 += 1
            self._apply_eve_data_cell(ws3, row3, 1, "TOTAL", is_total=True)
            for j, ore_name in enumerate(char_ores_sorted):
                self._apply_eve_data_cell(ws3, row3, j + 2, ore_totals[ore_name], is_total=True)
            self._apply_eve_data_cell(ws3, row3, tcol, grand, is_total=True)

        wb.save(filepath)
        return filepath

    # ------------------- #
    #  ORE VOLUME LOOKUP  #
    # ------------------- #

    @lru_cache(maxsize=256)
    def get_ore_volume(self, raw_name: str) -> Tuple[float, str]:
        clean_name = raw_name.strip().rstrip('.')
        if clean_name in ORE_VOLUMES:
            return ORE_VOLUMES[clean_name], clean_name
        clean_lower = clean_name.lower()
        for base_ore, volume in ORE_VOLUMES.items():
            if base_ore.lower() in clean_lower:
                return volume, clean_name
        return 1.0, clean_name

    def _apply_saved_app_settings(self):
        global DOCS, CRIT_SOUND_FILE, UPDATE_INTERVAL_MS, HISTORY_DAYS
        global CRITICAL_HIT_KEYWORD

        app_settings = self.app_config.get("app_settings", {})
        if not app_settings:
            return

        if "docs_path" in app_settings:
            DOCS = app_settings["docs_path"]
        if "crit_sound_file" in app_settings:
            CRIT_SOUND_FILE = app_settings["crit_sound_file"]
        if "update_interval_ms" in app_settings:
            UPDATE_INTERVAL_MS = max(250, int(app_settings["update_interval_ms"]))
        if "history_days" in app_settings:
            HISTORY_DAYS = max(1, int(app_settings["history_days"]))
        if "crit_keyword" in app_settings:
            CRITICAL_HIT_KEYWORD = app_settings["crit_keyword"]

    # -------- #
    #  CONFIG  #
    # -------- #

    def load_config(self) -> Dict:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_config(self) -> None:
        self.app_config["win_geom"] = self.root.winfo_geometry()
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.app_config, f, indent=2)
        except Exception:
            pass

    def on_close(self) -> None:
        self.update_loop_running = False
        if self.history_window and self.history_window.winfo_exists():
            self.on_history_close()
        self.save_config()
        self.root.destroy()

    # ---------------------- #
    #  LIVE MONITORING LOOP  #
    # ---------------------- #

    def update_loop(self) -> None:
        if not self.update_loop_running:
            return

        for char_id, tracker in self.characters.items():
            latest_log = self._get_latest_log_for_char(char_id)
            if latest_log and latest_log != tracker.log_path:
                tracker.log_path = latest_log
                tracker.log_pos = 0

            if tracker.log_path and os.path.exists(tracker.log_path):
                try:
                    with open(tracker.log_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                        f.seek(tracker.log_pos)
                        new_data = f.read()
                        new_pos = f.tell()
                        if new_data:
                            self._process_log_data(tracker, new_data)
                        # Only advance log_pos when session is active
                        # so events during inactive period are preserved
                        if tracker.session_active:
                            tracker.log_pos = new_pos
                except Exception:
                    pass

        self._update_ui_labels()
        self.root.after(UPDATE_INTERVAL_MS, self.update_loop)

    def _process_log_data(self, tracker: CharacterTracker, data: str) -> None:
        # log processing with compression tracking
        if not tracker.session_active:
            return

        crit_processed = False
        
        for line in data.splitlines():
            # Check for compression events
            compression_match = COMPRESSION_PATTERN.search(line)
            if compression_match:
                ore_type = compression_match.group('ore_type')
                compressed_amount = float(compression_match.group('amount').replace(",", ""))
                
                compression_ratio = COMPRESSION_RATIOS.get(ore_type, 100)
                original_units = compressed_amount * compression_ratio
                
                volume_per_unit, ore_name = self.get_ore_volume(ore_type)
                total_volume = original_units * volume_per_unit
                
                tracker.compression_log[ore_name] = tracker.compression_log.get(ore_name, 0) + total_volume
                continue
            
            # OPTIMIZATION: Fast check - skip if not a mining line
            if not MINING_LINE.match(line):
                continue
            
            # Regular mining
            regular_match = REGULAR_MINE_PATTERN.search(line)
            if regular_match:
                units = float(regular_match.group('amount').replace(",", ""))
                volume, ore_name = self.get_ore_volume(regular_match.group('ore_type'))
                total_volume = units * volume
                tracker.total_m3 += total_volume
                tracker.ore_summary[ore_name] = tracker.ore_summary.get(ore_name, 0) + total_volume

            # Critical hit
            if CRITICAL_HIT_KEYWORD in line and not crit_processed:
                crit_match = CRIT_MINE_PATTERN.search(line)
                if crit_match:
                    units = float(crit_match.group('amount').replace(",", ""))
                    ore_type_raw = crit_match.group('ore_type')
                    ore_type_clean = ore_type_raw.split('<')[0].split('\r')[0].split('\n')[0].strip()
                    volume, ore_name = self.get_ore_volume(ore_type_clean)
                    total_volume = units * volume
                    tracker.total_m3 += total_volume
                    tracker.ore_summary[ore_name] = tracker.ore_summary.get(ore_name, 0) + total_volume
                    tracker.crit_count += 1
                    crit_processed = True
                    self.trigger_crit_alert()

    def _update_ui_labels(self) -> None:
        for char_id, tracker in self.characters.items():
            if char_id not in self.char_widgets:
                continue
            w = self.char_widgets[char_id]
            w['crit'].config(text=f"Crits: {tracker.crit_count}")
            session_m3 = tracker.total_m3 - tracker.session_start_m3
            w['ore'].config(text=f"Total: {session_m3:,.1f} m3")
    
            if tracker.ore_summary:
                summary = "\n".join([
                    f"{ore_name}: {volume:,.1f} m3"
                    for ore_name, volume in tracker.ore_summary.items()
                ])
            else:
                summary = "Waiting..."
            w['summary'].config(text=summary)
    
            self._update_rate_stats(char_id, tracker, w)

    # -------- #
    #  ALERTS  #
    # -------- #

    def trigger_crit_alert(self) -> None:
        if HAS_NOTIFICATION:
            try:
                notification.notify(title="MINING", message="Critical Hit!", timeout=1)
            except Exception:
                pass

        if HAS_PLAYSOUND and os.path.exists(CRIT_SOUND_FILE):
            try:
                playsound(CRIT_SOUND_FILE, block=False)
            except Exception:
                pass

    # ===================== #
    #   SESSION CONTROL    #
    # ===================== #

    def toggle_session(self, char_id: str):
        tracker = self.all_characters[char_id]
        widgets = self.char_widgets[char_id]

        tracker.session_active = not tracker.session_active

        if tracker.session_active:
            # Process any backlog data accumulated while session was stopped
            # so crits/ore from the inactive period are not lost
            if tracker.log_path and os.path.exists(tracker.log_path):
                try:
                    with open(tracker.log_path, "r", encoding="utf-8-sig", errors="ignore") as f:
                        f.seek(tracker.log_pos)
                        backlog = f.read()
                        if backlog:
                            self._process_log_data(tracker, backlog)
                        tracker.log_pos = f.tell()
                except Exception:
                    pass
            # Set session baselines AFTER backlog processing
            # so rate calculation starts clean from this point
            tracker.session_start_time = time.time()
            tracker.session_start_m3 = tracker.total_m3
            widgets['start_stop_btn'].config(text="■ STOP", fg=RED)
            theoretical_m3_per_sec = tracker.get_total_theoretical_m3_per_sec()
            if theoretical_m3_per_sec > 0:
                widgets['actual'].config(
                    text=f"◉ Actual: {theoretical_m3_per_sec:.2f} m3/s ({theoretical_m3_per_sec * 3600:,.0f} m3/hr)"
                )
        else:
            widgets['start_stop_btn'].config(text="▶ START", fg=GREEN)

    def reset_session(self, char_id: str):
        tracker = self.all_characters[char_id]
        widgets = self.char_widgets[char_id]

        if tracker.session_active:
            tracker.session_active = False
            widgets['start_stop_btn'].config(text="▶ START", fg=GREEN)

        tracker.crit_count = 0
        tracker.total_m3 = 0.0
        tracker.ore_summary = {}
        tracker.compression_log = {}
        tracker.session_start_time = time.time()
        tracker.session_start_m3 = 0.0

        widgets['crit'].config(text="Crits: 0")
        widgets['ore'].config(text="Total: 0.0 m3")
        widgets['summary'].config(text="Waiting...")
        widgets['actual'].config(text="◉ Actual: -- m3/s")

    # ===================== #
    #     SHIP CONFIG      #
    # ===================== #

    def load_ship_configs(self):
        # Load ship configurations per character
        ship_configs = self.app_config.get("ship_configs", {})
        
        for char_id, tracker in self.all_characters.items():
            if char_id in ship_configs:
                cfg = ship_configs[char_id]
                
                if "profiles" in cfg:
                    tracker.ship_profiles = {}
                    for profile_name, profile_data in cfg["profiles"].items():
                        modules = []
                        for mod_data in profile_data.get("modules", []):
                            modules.append(MiningModule.from_dict(mod_data))
                        tracker.ship_profiles[profile_name] = modules
                    
                    tracker.active_profile = cfg.get("active_profile", "Default")
                    
                    # Ensure active profile exists
                    if tracker.active_profile not in tracker.ship_profiles:
                        if tracker.ship_profiles:
                            tracker.active_profile = list(tracker.ship_profiles.keys())[0]
                        else:
                            tracker.active_profile = "Default"
                            tracker.ship_profiles["Default"] = []
                
                elif "modules" in cfg:
                    modules_data = cfg.get("modules", [])
                    modules = []
                    for mod_data in modules_data:
                        modules.append(MiningModule.from_dict(mod_data))
                    tracker.ship_profiles = {"Default": modules}
                    tracker.active_profile = "Default"
                
                # yield/cycle
                elif "yield_per_cycle" in cfg:
                    old_yield = cfg.get("yield_per_cycle", 0.0)
                    old_cycle = cfg.get("cycle_time", 60.0)
                    if old_yield > 0:
                        module = MiningModule(
                            name="Module 1",
                            yield_per_cycle=old_yield,
                            cycle_time=old_cycle,
                            enabled=True
                        )
                        tracker.ship_profiles = {"Default": [module]}
                        tracker.active_profile = "Default"

    def save_ship_configs(self):
        # Save ship configurations with multiple profiles
        ship_configs = {}
        for char_id, tracker in self.all_characters.items():
            profiles_data = {}
            for profile_name, modules in tracker.ship_profiles.items():
                profiles_data[profile_name] = {
                    "modules": [m.to_dict() for m in modules]
                }
            
            ship_configs[char_id] = {
                "active_profile": tracker.active_profile,
                "profiles": profiles_data
            }
        
        self.app_config["ship_configs"] = ship_configs
        self.save_config()

    def show_ship_config(self, char_id: str):
        # Show ship configuration dialog
        if char_id in self.ship_config_dialogs and self.ship_config_dialogs[char_id].winfo_exists():
            self.ship_config_dialogs[char_id].lift()
            self.ship_config_dialogs[char_id].focus_force()
            return

        tracker = self.all_characters[char_id]

        dialog = tk.Toplevel(self.root)
        dialog.configure(bg=BORDER)
        dialog.overrideredirect(True)
        dialog.wm_attributes("-topmost", 1)
        dialog.attributes("-alpha", 0.85)
        dialog.resizable(False, False)
        self.ship_config_dialogs[char_id] = dialog

        _drag_x = [0]
        _drag_y = [0]

        def start_drag(event):
            if isinstance(event.widget, (tk.Entry, tk.OptionMenu)):
                return
            _drag_x[0] = event.x
            _drag_y[0] = event.y

        def do_drag(event):
            if isinstance(event.widget, (tk.Entry, tk.OptionMenu)):
                return
            x = dialog.winfo_x() + event.x - _drag_x[0]
            y = dialog.winfo_y() + event.y - _drag_y[0]
            dialog.geometry(f"+{x}+{y}")

        config_key = f"ship_config_{char_id}_geom"
        saved_geom = self.app_config.get(config_key, "+300+200")

        border_frame = tk.Frame(dialog, bg=BORDER, padx=1, pady=1)
        border_frame.pack(fill="both", expand=True)

        main_frame = tk.Frame(border_frame, bg=BG_PANEL, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        main_frame.bind("<Button-1>", start_drag)
        main_frame.bind("<B1-Motion>", do_drag)

        top_bar = tk.Frame(main_frame, bg=BG_PANEL)
        top_bar.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 15))
        top_bar.bind("<Button-1>", start_drag)
        top_bar.bind("<B1-Motion>", do_drag)

        title_label = tk.Label(
            top_bar,
            text=f"⚙ {tracker.char_name.upper()} — SHIP FITTING",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 10, "bold")
        )
        title_label.pack(side="left")
        title_label.bind("<Button-1>", start_drag)
        title_label.bind("<B1-Motion>", do_drag)

        def close_dialog():
            try:
                x = dialog.winfo_x()
                y = dialog.winfo_y()
                position = f"+{x}+{y}"
                self.app_config[config_key] = position
                self.save_config()
            except Exception:
                pass
            if char_id in self.ship_config_dialogs:
                del self.ship_config_dialogs[char_id]
            dialog.destroy()

        close_btn = tk.Label(
            top_bar,
            text="✕",
            fg=DIM,
            bg=BG_PANEL,
            font=("Consolas", 14, "bold"),
            cursor="hand2"
        )
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: close_dialog())
        close_btn.bind("<Enter>", lambda e: close_btn.config(fg=RED))
        close_btn.bind("<Leave>", lambda e: close_btn.config(fg=DIM))

        # ========================= #
        #    PROFILE MANAGEMENT    #
        # ========================= #
        
        profile_frame = tk.Frame(main_frame, bg=BG_PANEL)
        profile_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 15))
        
        tk.Label(
            profile_frame,
            text="◆ SHIP PROFILE:",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        ).pack(side="left", padx=(0, 10))

        # Profile selector
        current_profile = tk.StringVar(value=tracker.active_profile)
        profile_menu = tk.OptionMenu(
            profile_frame,
            current_profile,
            *tracker.get_profile_names()
        )
        profile_menu.config(
            bg=BG,
            fg=WHITE,
            font=("Consolas", 9),
            activebackground=BORDER,
            activeforeground=CYAN,
            highlightthickness=0,
            relief="flat"
        )
        profile_menu["menu"].config(
            bg=BG_PANEL,
            fg=WHITE,
            activebackground=BORDER,
            activeforeground=CYAN
        )
        profile_menu.pack(side="left", padx=5)

        # Profile management buttons
        btn_new = tk.Button(
            profile_frame,
            text="+ NEW",
            command=lambda: self.create_new_profile(tracker, current_profile, profile_menu, module_vars, update_preview, dialog),
            bg=BG,
            fg=GREEN,
            font=("Consolas", 8, "bold"),
            relief="flat",
            cursor="hand2",
            width=6
        )
        btn_new.pack(side="left", padx=2)

        btn_rename = tk.Button(
            profile_frame,
            text="✎ RENAME",
            command=lambda: self.rename_current_profile(tracker, current_profile, profile_menu, dialog),
            bg=BG,
            fg=CYAN,
            font=("Consolas", 8, "bold"),
            relief="flat",
            cursor="hand2",
            width=8
        )
        btn_rename.pack(side="left", padx=2)

        btn_delete = tk.Button(
            profile_frame,
            text="✕ DELETE",
            command=lambda: self.delete_current_profile(tracker, current_profile, profile_menu, module_vars, update_preview, dialog),
            bg=BG,
            fg=RED,
            font=("Consolas", 8, "bold"),
            relief="flat",
            cursor="hand2",
            width=8
        )
        btn_delete.pack(side="left", padx=2)

        # MINING MODULES
        modules_label = tk.Label(
            main_frame,
            text="◆ MINING MODULES",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        )
        modules_label.grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 10))

        tk.Label(main_frame, text="", bg=BG_PANEL, width=3).grid(row=3, column=0)
        tk.Label(main_frame, text="Module Name", fg=DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=3, column=1, padx=5)
        tk.Label(main_frame, text="Yield (m3/cycle)", fg=DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=3, column=2, padx=5)
        tk.Label(main_frame, text="Cycle Time (s)", fg=DIM, bg=BG_PANEL, font=("Consolas", 8)).grid(row=3, column=3, padx=5)

        module_vars = []

        def load_profile_into_ui(profile_name: str):
            """Load selected profile's modules into UI"""
            modules = tracker.ship_profiles.get(profile_name, [])
            
            # Pad to MAX_MODULES
            while len(modules) < MAX_MODULES:
                modules.append(MiningModule())
            
            # Update UI vars
            for i, (module, mv) in enumerate(zip(modules[:MAX_MODULES], module_vars)):
                mv['enabled'].set(module.enabled and module.is_configured())
                mv['name'].set(module.name if module.is_configured() else "")
                mv['yield'].set(str(module.yield_per_cycle) if module.yield_per_cycle > 0 else "")
                mv['cycle'].set(str(module.cycle_time) if module.cycle_time > 0 else "")
                
                if not (module.enabled and module.is_configured()):
                    mv['name_entry'].config(state="disabled")
                else:
                    mv['name_entry'].config(state="normal")
            
            update_preview()

        # Create module input fields
        active_modules = tracker.get_active_modules()
        while len(active_modules) < MAX_MODULES:
            active_modules.append(MiningModule())

        for i in range(MAX_MODULES):
            module = active_modules[i] if i < len(active_modules) else MiningModule()
            row = 4 + i

            enabled_var = tk.BooleanVar(value=module.enabled and module.is_configured())

            enabled_cb = tk.Checkbutton(
                main_frame,
                variable=enabled_var,
                bg=BG_PANEL,
                activebackground=BG_PANEL,
                selectcolor=WHITE,
                highlightthickness=0
            )
            enabled_cb.grid(row=row, column=0, padx=2, pady=3)

            name_display = module.name if module.is_configured() else ""
            name_var = tk.StringVar(value=name_display)
            name_entry = tk.Entry(
                main_frame,
                textvariable=name_var,
                width=12,
                font=("Consolas", 9),
                bg=BG,
                fg=WHITE,
                insertbackground=CYAN,
                disabledbackground=BG,
                disabledforeground=DIM
            )
            name_entry.grid(row=row, column=1, padx=5, pady=3)

            yield_display = str(module.yield_per_cycle) if module.yield_per_cycle > 0 else ""
            yield_var = tk.StringVar(value=yield_display)
            yield_entry = tk.Entry(
                main_frame,
                textvariable=yield_var,
                width=12,
                font=("Consolas", 9),
                bg=BG,
                fg=WHITE,
                insertbackground=CYAN
            )
            yield_entry.grid(row=row, column=2, padx=5, pady=3)

            cycle_display = str(module.cycle_time) if module.cycle_time > 0 else ""
            cycle_var = tk.StringVar(value=cycle_display)
            cycle_entry = tk.Entry(
                main_frame,
                textvariable=cycle_var,
                width=12,
                font=("Consolas", 9),
                bg=BG,
                fg=WHITE,
                insertbackground=CYAN
            )
            cycle_entry.grid(row=row, column=3, padx=5, pady=3)

            def update_name_state(name_e=name_entry, enabled_v=enabled_var, name_v=name_var):
                if enabled_v.get():
                    name_e.config(state="normal")
                else:
                    name_v.set("")
                    name_e.config(state="disabled")

            if not (module.enabled and module.is_configured()):
                name_entry.config(state="disabled")

            enabled_var.trace_add('write', lambda *args, fn=update_name_state: fn())

            module_vars.append({
                'enabled': enabled_var,
                'name': name_var,
                'yield': yield_var,
                'cycle': cycle_var,
                'name_entry': name_entry
            })

        sep_row = 4 + MAX_MODULES

        # Theoretical preview
        preview_frame = tk.Frame(main_frame, bg=BG, padx=10, pady=8)
        preview_frame.grid(row=sep_row, column=0, columnspan=4, sticky="ew", pady=(15, 10))

        preview_label = tk.Label(
            preview_frame,
            text="◈ Theoretical: -- m3/s (-- m3/hr)",
            fg=CYAN,
            bg=BG,
            font=("Consolas", 9, "bold")
        )
        preview_label.pack()

        def update_preview(*args):
            total_m3_per_sec = 0.0
            active_count = 0
            for mv in module_vars:
                if mv['enabled'].get():
                    try:
                        y = float(mv['yield'].get()) if mv['yield'].get() else 0.0
                        c = float(mv['cycle'].get()) if mv['cycle'].get() else 0.0
                        if y > 0 and c > 0:
                            total_m3_per_sec += y / c
                            active_count += 1
                    except ValueError:
                        pass

            if total_m3_per_sec > 0:
                preview_label.config(
                    text=f"◈ Theoretical: {total_m3_per_sec:.2f} m3/s ({total_m3_per_sec * 3600:,.0f} m3/hr) [{active_count} module{'s' if active_count > 1 else ''}]"
                )
            else:
                preview_label.config(text="◈ Theoretical: -- m3/s (configure modules)")

        for mv in module_vars:
            mv['enabled'].trace_add('write', update_preview)
            mv['yield'].trace_add('write', update_preview)
            mv['cycle'].trace_add('write', update_preview)

        # Profile change handler
        def on_profile_change(*args):
            new_profile = current_profile.get()
            if new_profile != tracker.active_profile:
                # Save current profile before switching
                save_current_profile_to_tracker()
                # Switch profile
                tracker.active_profile = new_profile
                # Load new profile into UI
                load_profile_into_ui(new_profile)

        current_profile.trace_add('write', on_profile_change)

        def save_current_profile_to_tracker():
            # Save current UI kept track of the profile used
            modules = []
            for mv in module_vars:
                mod = MiningModule(
                    name=mv['name'].get(),
                    yield_per_cycle=float(mv['yield'].get()) if mv['yield'].get() else 0.0,
                    cycle_time=float(mv['cycle'].get()) if mv['cycle'].get() else 0.0,
                    enabled=mv['enabled'].get()
                )
                modules.append(mod)
            tracker.ship_profiles[tracker.active_profile] = modules

        update_preview()

        # Buttons
        btn_frame = tk.Frame(main_frame, bg=BG_PANEL)
        btn_frame.grid(row=sep_row + 1, column=0, columnspan=4, pady=(10, 0))

        def save_and_close():
            try:
                # Save current profile
                save_current_profile_to_tracker()
                
                # Save all profiles to config
                self.save_ship_configs()

                if char_id in self.char_widgets:
                    self.update_ship_indicator(char_id)
                    self.update_profile_label(char_id)

                try:
                    x = dialog.winfo_x()
                    y = dialog.winfo_y()
                    position = f"+{x}+{y}"
                    self.app_config[config_key] = position
                    self.save_config()
                except Exception:
                    pass

                if char_id in self.ship_config_dialogs:
                    del self.ship_config_dialogs[char_id]
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid numbers")

        tk.Button(
            btn_frame,
            text="✔ SAVE",
            command=save_and_close,
            bg=BG,
            fg=GREEN,
            font=("Consolas", 9, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="✕ CANCEL",
            command=close_dialog,
            bg=BG,
            fg=RED,
            font=("Consolas", 9, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        ).pack(side="left", padx=5)

        dialog.update_idletasks()
        try:
            if '+' in saved_geom:
                parts = saved_geom.split('+')
                if len(parts) >= 3:
                    x_pos = parts[1]
                    y_pos = parts[2]
                    dialog.geometry(f"+{x_pos}+{y_pos}")
                else:
                    dialog.geometry("+300+200")
            else:
                dialog.geometry("+300+200")
        except Exception:
            dialog.geometry("+300+200")

        dialog.update()

        def initial_focus():
            if dialog.winfo_exists():
                dialog.lift()
                dialog.focus_force()

        dialog.after(150, initial_focus)

    def _ask_string_centered(self, title, prompt, parent_dialog, initialvalue=""):
        # Custom askstring dialog that centers on the ship config window
        result = [None]
        dlg = tk.Toplevel(parent_dialog)
        dlg.title(title)
        dlg.configure(bg=BG_PANEL)
        dlg.resizable(False, False)
        dlg.transient(parent_dialog)
        dlg.grab_set()

        tk.Label(dlg, text=prompt, bg=BG_PANEL, fg=WHITE, font=("Consolas", 10)).pack(padx=20, pady=(15, 5))

        entry = tk.Entry(dlg, font=("Consolas", 10), width=30, bg=BG, fg=WHITE, insertbackground=WHITE)
        entry.pack(padx=20, pady=5)
        if initialvalue:
            entry.insert(0, initialvalue)
            entry.select_range(0, tk.END)

        btn_frame = tk.Frame(dlg, bg=BG_PANEL)
        btn_frame.pack(pady=(5, 15))

        def on_ok(event=None):
            result[0] = entry.get().strip()
            dlg.destroy()

        def on_cancel(event=None):
            dlg.destroy()

        tk.Button(btn_frame, text="OK", command=on_ok, bg=BG, fg=GREEN,
                  font=("Consolas", 9, "bold"), relief="flat", width=8, cursor="hand2").pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", command=on_cancel, bg=BG, fg=RED,
                  font=("Consolas", 9, "bold"), relief="flat", width=8, cursor="hand2").pack(side="left", padx=5)

        entry.bind("<Return>", on_ok)
        entry.bind("<Escape>", on_cancel)

        dlg.update_idletasks()

        # Center on parent dialog
        pw = parent_dialog.winfo_width()
        ph = parent_dialog.winfo_height()
        px = parent_dialog.winfo_x()
        py = parent_dialog.winfo_y()
        dw = dlg.winfo_reqwidth()
        dh = dlg.winfo_reqheight()
        x = px + (pw - dw) // 2
        y = py + (ph - dh) // 2
        dlg.geometry(f"+{x}+{y}")
        dlg.wm_attributes("-topmost", 1)

        entry.focus_set()
        dlg.wait_window()
        return result[0] if result[0] else None

    def create_new_profile(self, tracker: CharacterTracker, current_profile_var: tk.StringVar, 
                          profile_menu: tk.OptionMenu, module_vars: List, update_preview_fn, parent_dialog=None):
        # Create a new ship profile
        parent = parent_dialog or self.root
        new_name = self._ask_string_centered(
            "New Profile",
            "Enter name for new ship profile:",
            parent
        )
        
        if new_name:
            if tracker.create_profile(new_name):
                # Update dropdown menu
                menu = profile_menu["menu"]
                menu.delete(0, "end")
                for profile in tracker.get_profile_names():
                    menu.add_command(
                        label=profile,
                        command=lambda value=profile: current_profile_var.set(value)
                    )
                
                # Switch to new profile
                current_profile_var.set(new_name)
                tracker.active_profile = new_name
                
                # Clear UI
                for mv in module_vars:
                    mv['enabled'].set(False)
                    mv['name'].set("")
                    mv['yield'].set("")
                    mv['cycle'].set("")
                    mv['name_entry'].config(state="disabled")
                
                update_preview_fn()
            else:
                messagebox.showerror("Error", "Profile name already exists or is invalid")

    def rename_current_profile(self, tracker: CharacterTracker, current_profile_var: tk.StringVar, 
                               profile_menu: tk.OptionMenu, parent_dialog=None):
        # Rename the current profile
        old_name = tracker.active_profile
        
        if len(tracker.ship_profiles) == 1:
            messagebox.showwarning("Cannot Rename", "You must have at least one profile",
                                   parent=parent_dialog or self.root)
            return
        
        parent = parent_dialog or self.root
        new_name = self._ask_string_centered(
            "Rename Profile",
            f"Rename '{old_name}' to:",
            parent,
            initialvalue=old_name
        )
        
        if new_name and new_name != old_name:
            if tracker.rename_profile(old_name, new_name):
                # Update dropdown menu
                menu = profile_menu["menu"]
                menu.delete(0, "end")
                for profile in tracker.get_profile_names():
                    menu.add_command(
                        label=profile,
                        command=lambda value=profile: current_profile_var.set(value)
                    )
                
                current_profile_var.set(new_name)
            else:
                messagebox.showerror("Error", "Profile name already exists or is invalid",
                                    parent=parent)

    def delete_current_profile(self, tracker: CharacterTracker, current_profile_var: tk.StringVar, 
                               profile_menu: tk.OptionMenu, module_vars: List, update_preview_fn, parent_dialog=None):
        # Delete the current profile
        profile_to_delete = tracker.active_profile
        parent = parent_dialog or self.root
        
        if len(tracker.ship_profiles) == 1:
            messagebox.showwarning("Cannot Delete", "You must have at least one profile",
                                   parent=parent)
            return
        
        result = messagebox.askyesno(
            "Delete Profile",
            f"Are you sure you want to delete profile '{profile_to_delete}'?",
            parent=parent
        )
        
        if result:
            if tracker.delete_profile(profile_to_delete):
                # Update dropdown menu
                menu = profile_menu["menu"]
                menu.delete(0, "end")
                for profile in tracker.get_profile_names():
                    menu.add_command(
                        label=profile,
                        command=lambda value=profile: current_profile_var.set(value)
                    )
                
                # Switch to remaining profile
                current_profile_var.set(tracker.active_profile)
                
                # Load new profile into UI
                modules = tracker.get_active_modules()
                while len(modules) < MAX_MODULES:
                    modules.append(MiningModule())
                
                for i, (module, mv) in enumerate(zip(modules[:MAX_MODULES], module_vars)):
                    mv['enabled'].set(module.enabled and module.is_configured())
                    mv['name'].set(module.name if module.is_configured() else "")
                    mv['yield'].set(str(module.yield_per_cycle) if module.yield_per_cycle > 0 else "")
                    mv['cycle'].set(str(module.cycle_time) if module.cycle_time > 0 else "")
                    
                    if not (module.enabled and module.is_configured()):
                        mv['name_entry'].config(state="disabled")
                    else:
                        mv['name_entry'].config(state="normal")
                
                update_preview_fn()

    def update_ship_indicator(self, char_id: str):
        tracker = self.all_characters[char_id]
        if char_id not in self.char_widgets:
            return
        widgets = self.char_widgets[char_id]
        if tracker.has_any_configured_module():
            widgets['ship_indicator'].config(fg=GREEN)
        else:
            widgets['ship_indicator'].config(fg=RED)

    def update_profile_label(self, char_id: str):
        # Update the active profile label for a character
        tracker = self.all_characters[char_id]
        if char_id not in self.char_widgets:
            return
        widgets = self.char_widgets[char_id]
        widgets['profile_label'].config(text=f"〈{tracker.active_profile}〉")

    # ========================= #
    #    RATE CALCULATIONS     #
    # ========================= #

    def _update_rate_stats(self, char_id: str, tracker: CharacterTracker, widgets: Dict):
        theoretical_m3_per_sec = tracker.get_total_theoretical_m3_per_sec()

        if theoretical_m3_per_sec > 0:
            widgets['theoretical'].config(
                text=f"◈ Theoretical: {theoretical_m3_per_sec:.2f} m3/s ({theoretical_m3_per_sec * 3600:,.0f} m3/hr)"
            )
        else:
            widgets['theoretical'].config(text="◈ Theoretical: -- m3/s (configure ship)")

        if not tracker.session_active:
            return

        actual_m3_per_sec = 0.0
        session_duration = time.time() - tracker.session_start_time
        if session_duration > 10 and tracker.total_m3 > 0:
            actual_m3_per_sec = (tracker.total_m3 - tracker.session_start_m3) / session_duration

        widgets['actual'].config(
            text=f"◉ Actual: {actual_m3_per_sec:.2f} m3/s ({actual_m3_per_sec * 3600:,.0f} m3/hr)"
        )

    # ========================= #
    #      CONFIG DIALOG       #
    # ========================= #

    def show_config_dialog(self):
        global DOCS, UPDATE_INTERVAL_MS, HISTORY_DAYS

        if self.config_dialog is not None and self.config_dialog.winfo_exists():
            self.config_dialog.lift()
            self.config_dialog.focus_force()
            return

        self.config_icon.config(fg=CYAN)
        self.config_icon.unbind("<Button-1>")
        self.config_icon.unbind("<Enter>")
        self.config_icon.unbind("<Leave>")

        dialog = tk.Toplevel(self.root)
        dialog.configure(bg=BORDER)
        dialog.overrideredirect(True)
        dialog.wm_attributes("-topmost", 1)
        dialog.attributes("-alpha", 0.85)
        dialog.resizable(False, False)
        self.config_dialog = dialog

        _drag_x = [0]
        _drag_y = [0]

        def start_drag(event):
            if isinstance(event.widget, tk.Entry):
                return
            _drag_x[0] = event.x
            _drag_y[0] = event.y

        def do_drag(event):
            if isinstance(event.widget, tk.Entry):
                return
            x = dialog.winfo_x() + event.x - _drag_x[0]
            y = dialog.winfo_y() + event.y - _drag_y[0]
            dialog.geometry(f"+{x}+{y}")

        config_key = "config_dialog_geom"
        saved_geom = self.app_config.get(config_key, "+250+150")

        border_frame = tk.Frame(dialog, bg=BORDER, padx=1, pady=1)
        border_frame.pack(fill="both", expand=True)

        main_frame = tk.Frame(border_frame, bg=BG_PANEL, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        main_frame.bind("<Button-1>", start_drag)
        main_frame.bind("<B1-Motion>", do_drag)

        top_bar = tk.Frame(main_frame, bg=BG_PANEL)
        top_bar.pack(fill="x", pady=(0, 15))
        top_bar.bind("<Button-1>", start_drag)
        top_bar.bind("<B1-Motion>", do_drag)

        title_label = tk.Label(
            top_bar,
            text="⚙ CONFIG",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 10, "bold")
        )
        title_label.pack(side="left")
        title_label.bind("<Button-1>", start_drag)
        title_label.bind("<B1-Motion>", do_drag)

        def close_dialog():
            try:
                x = dialog.winfo_x()
                y = dialog.winfo_y()
                self.app_config[config_key] = f"+{x}+{y}"
                self.save_config()
            except Exception:
                pass
            self.config_dialog = None
            self._enable_config_icon()
            dialog.destroy()

        close_btn = tk.Label(
            top_bar,
            text="✕",
            fg=DIM,
            bg=BG_PANEL,
            font=("Consolas", 14, "bold"),
            cursor="hand2"
        )
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: close_dialog())
        close_btn.bind("<Enter>", lambda e: close_btn.config(fg=RED))
        close_btn.bind("<Leave>", lambda e: close_btn.config(fg=DIM))

        # Main content frame
        content_frame = tk.Frame(main_frame, bg=BG_PANEL)
        content_frame.pack(fill="both", expand=True)

        # CHARACTER SELECTION
        tk.Label(
            content_frame,
            text="◆ CHARACTER SELECTION",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        ).pack(anchor="w", pady=(0, 10))

        visible_chars = self.app_config.get("visible_characters", [])
        if not visible_chars:
            visible_chars = list(self.all_characters.keys())

        checklist_outer = tk.Frame(content_frame, bg=BORDER, padx=1, pady=1)
        checklist_outer.pack(fill="x", pady=(0, 15))

        checklist_frame = tk.Frame(checklist_outer, bg=BG, padx=10, pady=10)
        checklist_frame.pack(fill="both")

        char_vars = {}
        for i, (char_id, tracker) in enumerate(self.all_characters.items()):
            var = tk.BooleanVar(value=char_id in visible_chars)
            char_vars[char_id] = var

            cb_frame = tk.Frame(checklist_frame, bg=BG)
            cb_frame.pack(fill="x", pady=3)

            cb = tk.Checkbutton(
                cb_frame,
                variable=var,
                bg=BG,
                activebackground=BG,
                selectcolor=WHITE,
                highlightthickness=0
            )
            cb.pack(side="left", padx=(0, 8))

            accent = CHAR_ACCENTS[i % len(CHAR_ACCENTS)]
            tk.Label(
                cb_frame,
                text=tracker.char_name,
                fg=accent,
                bg=BG,
                font=("Consolas", 10, "bold")
            ).pack(side="left")

        # Separator
        tk.Label(
            content_frame,
            text="-" * 55,
            fg=BORDER,
            bg=BG_PANEL,
            font=("Consolas", 8)
        ).pack(pady=8)

        # APPLICATION SETTINGS
        fields_frame = tk.Frame(content_frame, bg=BG_PANEL)
        fields_frame.pack(fill="x")

        app_settings = self.app_config.get("app_settings", {})

        def make_field(parent, row, label_text, default_value, width=35, note_text=None):
            lbl = tk.Label(
                parent,
                text=label_text,
                fg=WHITE,
                bg=BG_PANEL,
                font=("Consolas", 9),
                anchor="w"
            )
            lbl.grid(row=row, column=0, sticky="w", pady=4, padx=(0, 10))

            entry_frame = tk.Frame(parent, bg=BG_PANEL)
            entry_frame.grid(row=row, column=1, sticky="w", pady=4)

            var = tk.StringVar(value=str(default_value))
            entry = tk.Entry(
                entry_frame,
                textvariable=var,
                width=width,
                font=("Consolas", 9),
                bg=BG,
                fg=WHITE,
                insertbackground=CYAN,
                relief="flat"
            )
            entry.pack(side="left")

            if note_text:
                note = tk.Label(
                    entry_frame,
                    text=note_text,
                    fg=GOLD,
                    bg=BG_PANEL,
                    font=("Consolas", 9)
                )
                note.pack(side="left", padx=(8, 0))

            return var

        # Paths & Files
        tk.Label(
            fields_frame,
            text="◆ PATHS & FILES",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        docs_var = make_field(fields_frame, 1, "Gamelogs Path:",
                              app_settings.get("docs_path", DOCS), width=40)

        # Separator
        tk.Label(
            fields_frame,
            text="-" * 55,
            fg=BORDER,
            bg=BG_PANEL,
            font=("Consolas", 8)
        ).grid(row=2, column=0, columnspan=3, pady=8)

        # Timing & Limits
        tk.Label(
            fields_frame,
            text="◆ TIMING & LIMITS",
            fg=CYAN,
            bg=BG_PANEL,
            font=("Consolas", 9, "bold")
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 8))

        interval_var = make_field(fields_frame, 4, "Update Interval (ms):",
                                  app_settings.get("update_interval_ms", UPDATE_INTERVAL_MS),
                                  width=10, note_text="min: 250")

        history_var = make_field(fields_frame, 5, "Default History Days:",
                                 app_settings.get("history_days", HISTORY_DAYS),
                                 width=10)

        # Buttons
        btn_frame = tk.Frame(main_frame, bg=BG_PANEL)
        btn_frame.pack(pady=(15, 0))

        def save_and_close():
            global DOCS, UPDATE_INTERVAL_MS, HISTORY_DAYS

            try:
                new_interval = int(interval_var.get())
                if new_interval < 250:
                    new_interval = 250
                new_history = int(history_var.get())
                if new_history < 1:
                    new_history = 1
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter valid numbers for numeric fields.")
                return

            # Save character selection
            selected_chars = [char_id for char_id, var in char_vars.items() if var.get()]
            self.save_visible_characters(selected_chars)

            # Save application settings
            DOCS = docs_var.get().strip()
            UPDATE_INTERVAL_MS = new_interval
            HISTORY_DAYS = new_history

            self.app_config["app_settings"] = {
                "docs_path": DOCS,
                "crit_sound_file": CRIT_SOUND_FILE,
                "update_interval_ms": UPDATE_INTERVAL_MS,
                "history_days": HISTORY_DAYS,
                "max_modules": MAX_MODULES
            }

            try:
                x = dialog.winfo_x()
                y = dialog.winfo_y()
                self.app_config[config_key] = f"+{x}+{y}"
            except Exception:
                pass

            self.save_config()
            self.config_dialog = None
            self._enable_config_icon()
            dialog.destroy()

        tk.Button(
            btn_frame,
            text="✔ SAVE",
            command=save_and_close,
            bg=BG,
            fg=GREEN,
            font=("Consolas", 9, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        ).pack(side="left", padx=5)

        tk.Button(
            btn_frame,
            text="✕ CANCEL",
            command=close_dialog,
            bg=BG,
            fg=RED,
            font=("Consolas", 9, "bold"),
            relief="flat",
            cursor="hand2",
            width=10
        ).pack(side="left", padx=5)

        dialog.update_idletasks()
        try:
            if '+' in saved_geom:
                parts = saved_geom.split('+')
                if len(parts) >= 3:
                    dialog.geometry(f"+{parts[1]}+{parts[2]}")
                else:
                    dialog.geometry("+250+150")
            else:
                dialog.geometry("+250+150")
        except Exception:
            dialog.geometry("+250+150")

        dialog.update()

        def initial_focus():
            if dialog.winfo_exists():
                dialog.lift()
                dialog.focus_force()

        dialog.after(150, initial_focus)

    def _enable_config_icon(self):
        self.config_icon.config(fg=DIM)
        self.config_icon.bind("<Button-1>", lambda e: self.show_config_dialog())
        self.config_icon.bind("<Enter>", lambda e: self.config_icon.config(fg=CYAN))
        self.config_icon.bind("<Leave>", lambda e: self.config_icon.config(fg=DIM))

if __name__ == "__main__":
    MiningDashboard()