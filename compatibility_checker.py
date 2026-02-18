"""
Compatibility Checker (Spreadsheet-Driven)
==========================================
Runs door hardware compatibility checks using rules loaded from
the Excel spreadsheet via the RulesEngine.

Physical checks (rail width vs panic hardware) use the stile width
database from the spreadsheet.

Code compliance checks use the FenestrAI Rules sheet.
"""

import re
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field

from rules_engine import RulesEngine, Rule, StileWidth


# ─────────────────────────────────────────────
# Panic Hardware Requirements
# (These are physical specs that don't change —
#  could move to spreadsheet in future)
# ─────────────────────────────────────────────

PANIC_HARDWARE_REQUIREMENTS = {
    "Von Duprin 98": {"min_rail_width": 4.0, "min_door_thickness": 1.75, "types": ["rim"]},
    "Von Duprin 99": {"min_rail_width": 4.0, "min_door_thickness": 1.75, "types": ["rim"]},
    "Von Duprin 33/35": {"min_rail_width": 2.75, "min_door_thickness": 1.75, "types": ["rim"]},
    "Von Duprin 22": {"min_rail_width": 2.0, "min_door_thickness": 1.75, "types": ["touch_bar"]},
}

GLASS_THICKNESS_REQ = [
    (36, 84, 0.5),     # 3'x7' -> 1/2"
    (36, 96, 0.625),   # 3'x8' -> 5/8"
    (48, 84, 0.625),   # 4'x7' -> 5/8"
    (48, 96, 0.75),    # 4'x8' -> 3/4"
]


# ─────────────────────────────────────────────
# Issue Model
# ─────────────────────────────────────────────

@dataclass
class Issue:
    door_number: str
    severity: str           # critical, warning, info, advisory
    category: str
    description: str
    details: str
    solutions: List[str]
    cost_if_missed: str = ""
    rule_id: str = ""       # Links back to spreadsheet rule
    code_reference: str = ""

    def to_dict(self) -> Dict:
        return {
            "door_number": self.door_number,
            "severity": self.severity,
            "category": self.category,
            "description": self.description,
            "details": self.details,
            "solutions": self.solutions,
            "cost_if_missed": self.cost_if_missed,
            "rule_id": self.rule_id,
            "code_reference": self.code_reference,
        }


# ─────────────────────────────────────────────
# Checker
# ─────────────────────────────────────────────

class CompatibilityChecker:
    """
    Checks door-hardware compatibility using both:
    1. Physical compatibility rules (rail width, glass thickness)
    2. Code compliance rules from the spreadsheet
    """

    def __init__(self, rules_engine: RulesEngine):
        self.rules = rules_engine

    def check_door(self, door: Dict, hw_set: Optional[Dict] = None) -> List[Issue]:
        """Run all checks on a single door."""
        issues = []

        # Physical compatibility checks
        issues.extend(self._check_panic_rail_width(door, hw_set))
        issues.extend(self._check_glass_thickness(door))
        issues.extend(self._check_door_thickness(door, hw_set))
        issues.extend(self._check_vision_panel_interference(door, hw_set))

        # Spreadsheet-driven rule checks
        issues.extend(self._check_spreadsheet_rules(door, hw_set))

        return issues

    def check_all_doors(self, doors: List[Dict],
                        hw_sets: Dict[str, Dict]) -> List[Issue]:
        all_issues = []
        for door in doors:
            hw_set = hw_sets.get(door.get("hardware_set", ""))
            issues = self.check_door(door, hw_set)
            all_issues.extend(issues)
        return all_issues

    # ── Physical Checks ──

    def _check_panic_rail_width(self, door: Dict, hw_set: Optional[Dict]) -> List[Issue]:
        issues = []
        mat = self._normalize_material(door.get("material", ""))
        if mat != "aluminum" or not hw_set:
            return issues

        panic = self._find_panic(hw_set)
        if not panic:
            return issues

        panic_key = f"{panic.get('manufacturer', '')} {panic.get('series', '')}".strip()
        req = PANIC_HARDWARE_REQUIREMENTS.get(panic_key)

        # Look up actual rail width from spreadsheet stile database
        manufacturer = door.get("manufacturer", "")
        series = door.get("product_line", "") or door.get("series", "")
        rail_width = None

        if manufacturer and series:
            stile = self.rules.lookup_stile(manufacturer, series)
            if stile and stile.width:
                rail_width = stile.width

        # Fall back to product line names if no series match
        if rail_width is None:
            pl = (door.get("product_line", "") or "").lower().strip()
            fallback_widths = {"narrow stile": 2.0, "medium stile": 3.5, "wide stile": 5.0}
            rail_width = fallback_widths.get(pl)

        if rail_width is None and mat == "aluminum" and panic:
            issues.append(Issue(
                door_number=door.get("door_number", ""),
                severity="warning",
                category="Rail Width",
                description=f"Could not determine rail width — verify stile type for {panic_key}",
                details=f"Door {door.get('door_number', '')} is aluminum with panic hardware but no product line or series number specified.",
                solutions=["Verify product line is narrow, medium, or wide stile",
                          "Add manufacturer and series to door schedule"],
            ))
            return issues

        if req and rail_width and rail_width < req["min_rail_width"]:
            issues.append(Issue(
                door_number=door.get("door_number", ""),
                severity="critical",
                category="Rail Width",
                description=f'{rail_width}" rails too narrow for {panic_key} (requires {req["min_rail_width"]}")',
                details=f"Door {door.get('door_number', '')} has {rail_width}\" rails. {panic_key} requires minimum {req['min_rail_width']}\" rails.",
                solutions=[
                    f"Change to wide stile (5\" rails)",
                    f"Change panic to Von Duprin 33/35 (fits 2.75\"+ rails)",
                    "Use touch bar panic hardware (fits 2\"+ rails)",
                ],
                cost_if_missed="$6,500-11,000 per door if discovered during installation",
            ))

        return issues

    def _check_glass_thickness(self, door: Dict) -> List[Issue]:
        issues = []
        mat = self._normalize_material(door.get("material", ""))
        glazing = door.get("glazing", "")

        if mat not in ["aluminum"] and not glazing:
            return issues

        w = self._dim_to_inches(door.get("width", ""))
        h = self._dim_to_inches(door.get("height", ""))
        if not w or not h:
            return issues

        min_thk = None
        for max_w, max_h, thk in GLASS_THICKNESS_REQ:
            if w <= max_w and h <= max_h:
                min_thk = thk
                break
        if min_thk is None and w > 0 and h > 0:
            min_thk = 0.75

        glass_thk = self._parse_glass_thickness(glazing)
        if min_thk and glass_thk and glass_thk < min_thk:
            issues.append(Issue(
                door_number=door.get("door_number", ""),
                severity="critical",
                category="Glass Thickness",
                description=f'Glass too thin: {self._frac(glass_thk)}" specified, {self._frac(min_thk)}" required',
                details=f"Door {door.get('door_number', '')} ({door.get('width', '')} x {door.get('height', '')}) needs minimum {self._frac(min_thk)}\" glass per GANA.",
                solutions=[
                    f"Increase glass to {self._frac(min_thk)}\"",
                    "Verify with GANA glazing manual",
                    "Check hardware compatibility with thicker glass",
                ],
                cost_if_missed="$400-1,200 per door for glass replacement",
            ))

        return issues

    def _check_door_thickness(self, door: Dict, hw_set: Optional[Dict]) -> List[Issue]:
        issues = []
        if not hw_set:
            return issues

        thk = self._parse_thickness(door.get("thickness", ""))
        panic = self._find_panic(hw_set)
        if panic:
            panic_key = f"{panic.get('manufacturer', '')} {panic.get('series', '')}".strip()
            req = PANIC_HARDWARE_REQUIREMENTS.get(panic_key)
            if req and thk < req["min_door_thickness"]:
                issues.append(Issue(
                    door_number=door.get("door_number", ""),
                    severity="critical",
                    category="Door Thickness",
                    description=f'Door too thin ({thk}") for {panic_key}',
                    details=f"{panic_key} requires minimum {req['min_door_thickness']}\" thickness.",
                    solutions=["Verify door thickness specification",
                              "Check for thinner-profile panic device"],
                    cost_if_missed="$500-2,000 per door",
                ))
        return issues

    def _check_vision_panel_interference(self, door: Dict, hw_set: Optional[Dict]) -> List[Issue]:
        issues = []
        if not hw_set:
            return issues

        dt = (door.get("door_type", "") or "").upper()
        comments = (door.get("comments", "") or "").upper()
        has_vision = dt.startswith("V") or "VISION" in comments

        if not has_vision:
            return issues

        panic = self._find_panic(hw_set)
        if not panic:
            return issues

        panic_key = f"{panic.get('manufacturer', '')} {panic.get('series', '')}".strip()
        req = PANIC_HARDWARE_REQUIREMENTS.get(panic_key)
        if req and "rim" in req.get("types", []):
            issues.append(Issue(
                door_number=door.get("door_number", ""),
                severity="warning",
                category="Vision Panel",
                description=f"Vision panel may interfere with {panic_key} rim device",
                details=f"Door {door.get('door_number', '')} has a vision panel with rim-mounted panic hardware.",
                solutions=["Verify vision panel location vs panic device height",
                          "Consider concealed vertical rod instead of rim",
                          "Adjust vision panel to clear hardware"],
                cost_if_missed="$300-800 for field modifications",
            ))
        return issues

    # ── Spreadsheet Rule Checks ──

    def _check_spreadsheet_rules(self, door: Dict, hw_set: Optional[Dict]) -> List[Issue]:
        """Run applicable rules from the spreadsheet against this door."""
        issues = []

        mat = self._normalize_material(door.get("material", ""))
        location = door.get("room_name", "") or door.get("comments", "") or ""
        has_glazing = bool(door.get("glazing")) or (door.get("door_type", "") or "").upper().startswith("V")
        has_panic = bool(self._find_panic(hw_set)) if hw_set else False
        is_fire_rated = bool(door.get("fire_rating", "").strip())

        applicable_rules = self.rules.get_rules_for_door(
            door_material=mat,
            door_location=location,
            has_glazing=has_glazing,
            has_panic=has_panic,
            is_fire_rated=is_fire_rated,
        )

        for rule in applicable_rules:
            # Map spreadsheet severity to our severity levels
            sev = rule.severity.lower()
            if sev == "critical":
                severity = "critical"
            elif sev == "moderate":
                severity = "warning"
            elif sev == "warning":
                severity = "info"
            elif sev == "advisory":
                severity = "info"
            else:
                severity = "info"

            issues.append(Issue(
                door_number=door.get("door_number", ""),
                severity=severity,
                category=rule.category,
                description=f"[{rule.rule_id}] {rule.condition}: {rule.threshold}",
                details=rule.notes if rule.notes else rule.condition,
                solutions=[rule.fix_recommendation] if rule.fix_recommendation else [],
                rule_id=rule.rule_id,
                code_reference=rule.code_reference,
            ))

        return issues

    # ── Utilities ──

    def _find_panic(self, hw_set: Optional[Dict]) -> Optional[Dict]:
        if not hw_set:
            return None
        for c in hw_set.get("components", []):
            if c.get("type") in ["panic_hardware", "panic", "exit_device"]:
                return c
        return None

    def _normalize_material(self, mat: str) -> str:
        m = (mat or "").lower().strip()
        if "alum" in m:
            return "aluminum"
        if "wood" in m or "wd" in m:
            return "wood"
        if "metal" in m or "hm" in m or "hollow" in m or "steel" in m:
            return "hollow_metal"
        return m or "unknown"

    def _dim_to_inches(self, s: str) -> Optional[float]:
        if not s:
            return None
        m = re.match(r"(\d+)['\s]*[-\s]*(\d+)", s)
        if m:
            return int(m.group(1)) * 12 + int(m.group(2))
        try:
            return float(s)
        except ValueError:
            return None

    def _parse_glass_thickness(self, s: str) -> Optional[float]:
        if not s:
            return None
        m = re.search(r'(\d+)/(\d+)', s)
        if m:
            return int(m.group(1)) / int(m.group(2))
        m = re.search(r'(\d+\.\d+)', s)
        if m:
            return float(m.group(1))
        return None

    def _parse_thickness(self, t: str) -> float:
        if not t:
            return 1.75
        s = t.replace('"', '').replace("'", "").strip()
        m = re.match(r'(\d+)\s+(\d+)/(\d+)', s)
        if m:
            return int(m.group(1)) + int(m.group(2)) / int(m.group(3))
        try:
            return float(s)
        except ValueError:
            return 1.75

    def _frac(self, d: float) -> str:
        fracs = {0.25: '1/4', 0.375: '3/8', 0.5: '1/2', 0.625: '5/8', 0.75: '3/4'}
        return fracs.get(d, str(d))
