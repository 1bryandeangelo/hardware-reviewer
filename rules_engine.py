"""
Rules Engine
=============
Reads compatibility rules and manufacturer data from the Excel spreadsheet.
Bryan and his hardware specialist can update rules by uploading a new version
of the spreadsheet — no code changes needed.

The spreadsheet has two sheets:
  1. "FenestrAI Rules" — code compliance and hardware rules
  2. "Aluminum Door Stile Widths" — manufacturer-specific rail dimensions
"""

import os
import re
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


# ─────────────────────────────────────────────
# Data Models
# ─────────────────────────────────────────────

@dataclass
class Rule:
    """A single rule from the rules spreadsheet."""
    rule_id: str
    category: str
    condition: str
    threshold: str
    severity: str           # Critical, Warning, Advisory, Info
    code_reference: str
    trigger_element: str    # Door, Glazing, Hardware, Floor
    applies_to: str         # Both, Exterior, Interior
    confidence: str
    failure_likelihood: str
    fix_recommendation: str
    notes: str
    trigger_flags: str = ""  # e.g. "egress=true, ada_required=true"

    def matches_door_context(self, door_material: str, door_location: str = "",
                              has_glazing: bool = False, has_panic: bool = False,
                              is_fire_rated: bool = False, has_access_control: bool = False,
                              has_auto_operator: bool = False) -> bool:
        """
        Check if this rule is potentially relevant to a given door context.
        This is a broad filter — specific threshold checks happen in the checker.
        """
        trigger = self.trigger_element.lower() if self.trigger_element else ""
        condition = self.condition.lower() if self.condition else ""
        category = self.category.lower() if self.category else ""

        # Filter by trigger element
        if trigger == "glazing" and not has_glazing:
            return False

        # Filter by specific conditions
        if "auto operator" in condition and not has_auto_operator:
            return False
        if "fire rated" in condition and not is_fire_rated:
            return False
        if "fire door" in condition and not is_fire_rated:
            return False
        if "panic" in condition.lower() and not has_panic:
            return False
        if "access control" in condition and not has_access_control:
            return False
        if "vestibule" in condition and "vestibule" not in door_location.lower():
            return False
        if "stairwell" in condition and "stair" not in door_location.lower():
            return False

        # Filter by applies_to
        applies = self.applies_to.lower() if self.applies_to else "both"
        if applies == "exterior" and "exterior" not in door_location.lower():
            # Only skip if we know it's interior
            if "interior" in door_location.lower():
                return False

        return True


@dataclass
class StileWidth:
    """A manufacturer-specific aluminum door stile width entry."""
    vendor: str
    model: str
    series: str
    width: Optional[float]   # inches
    depth: Optional[float]   # inches

    def width_str(self) -> str:
        if self.width is None:
            return "unknown"
        return f'{self.width}"'


# ─────────────────────────────────────────────
# Rules Engine
# ─────────────────────────────────────────────

class RulesEngine:
    """
    Loads and manages rules from the Excel spreadsheet.
    
    Usage:
        engine = RulesEngine()
        engine.load("data/rules.xlsx")
        
        rules = engine.get_rules_for_category("Fire Rating")
        stile = engine.lookup_stile("Kawneer", "350T")
    """

    def __init__(self):
        self.rules: List[Rule] = []
        self.stile_widths: List[StileWidth] = []
        self.loaded = False
        self.source_file = ""
        self.load_errors: List[str] = []

    def load(self, filepath: str) -> bool:
        """Load rules from an Excel spreadsheet."""
        if not PANDAS_AVAILABLE:
            self.load_errors.append("pandas is required to read Excel files")
            return False

        if not os.path.exists(filepath):
            self.load_errors.append(f"File not found: {filepath}")
            return False

        self.rules = []
        self.stile_widths = []
        self.load_errors = []
        self.source_file = filepath

        try:
            xls = pd.ExcelFile(filepath)

            # Identify stile widths sheet
            stile_sheet = None
            for name in xls.sheet_names:
                if "stile" in name.lower() or "aluminum" in name.lower() or "width" in name.lower():
                    stile_sheet = name
                    break

            # Load rules from all sheets (old single-sheet or new multi-tab format)
            if "FenestrAI Rules" in xls.sheet_names:
                # Old format: single rules sheet
                self._load_rules_sheet(xls, sheet_name="FenestrAI Rules")
            else:
                # New format: load every sheet except stile widths as a rules tab
                for name in xls.sheet_names:
                    if name == stile_sheet:
                        continue
                    try:
                        self._load_rules_sheet(xls, sheet_name=name)
                    except Exception as e:
                        self.load_errors.append(f"Error loading tab '{name}': {str(e)}")

            # Load stile widths
            if stile_sheet:
                self._load_stile_sheet(xls, sheet_name=stile_sheet)

            self.loaded = True
            return True

        except Exception as e:
            self.load_errors.append(f"Failed to load spreadsheet: {str(e)}")
            return False

    def _load_rules_sheet(self, xls, sheet_name="FenestrAI Rules"):
        """Parse the rules sheet into Rule objects. Handles both old and new column formats."""
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")

        # Normalize column names - handle both old format and new simplified format
        col_map = {}
        for col in df.columns:
            cl = str(col).lower().strip()
            if "rule" in cl and "id" in cl:
                col_map["rule_id"] = col
            elif cl == "what to check" or "condition" in cl:
                col_map["condition"] = col
            elif cl == "fail when" or "threshold" in cl:
                col_map["threshold"] = col
            elif "severity" in cl:
                col_map["severity"] = col
            elif cl == "code ref" or ("code" in cl and "ref" in cl):
                col_map["code_reference"] = col
            elif "trigger" in cl and "element" in cl:
                col_map["trigger_element"] = col
            elif "applies" in cl:
                col_map["applies_to"] = col
            elif "confidence" in cl:
                col_map["confidence"] = col
            elif "failure" in cl or "likelihood" in cl:
                col_map["failure_likelihood"] = col
            elif cl == "how to fix" or "fix" in cl or "recommendation" in cl:
                col_map["fix_recommendation"] = col
            elif cl == "when to apply" or ("trigger" in cl and "flag" in cl):
                col_map["trigger_flags"] = col
            elif "note" in cl:
                col_map["notes"] = col

        for _, row in df.iterrows():
            rule_id = str(row.get(col_map.get("rule_id", ""), "")).strip()
            if not rule_id or rule_id.lower() == "rule id":
                continue

            rule = Rule(
                rule_id=rule_id,
                category=sheet_name if sheet_name != "FenestrAI Rules" else str(row.get(col_map.get("category", ""), "")).strip(),
                condition=str(row.get(col_map.get("condition", ""), "")).strip(),
                threshold=str(row.get(col_map.get("threshold", ""), "")).strip(),
                severity=str(row.get(col_map.get("severity", ""), "")).strip(),
                code_reference=str(row.get(col_map.get("code_reference", ""), "")).strip(),
                trigger_element=str(row.get(col_map.get("trigger_element", ""), "")).strip(),
                applies_to=str(row.get(col_map.get("applies_to", ""), "")).strip(),
                confidence=str(row.get(col_map.get("confidence", ""), "")).strip(),
                failure_likelihood=str(row.get(col_map.get("failure_likelihood", ""), "")).strip(),
                fix_recommendation=str(row.get(col_map.get("fix_recommendation", ""), "")).strip(),
                notes=str(row.get(col_map.get("notes", ""), "")).strip(),
                trigger_flags=str(row.get(col_map.get("trigger_flags", ""), "")).strip(),
            )
            self.rules.append(rule)

    def _load_stile_sheet(self, xls, sheet_name):
        """Parse the stile widths sheet."""
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")

        # Normalize columns
        col_map = {}
        for col in df.columns:
            cl = str(col).lower().strip()
            if "vendor" in cl or "manufacturer" in cl or cl == "mfr":
                col_map["vendor"] = col
            elif "model" in cl:
                col_map["model"] = col
            elif "series" in cl:
                col_map["series"] = col
            elif "width" in cl:
                col_map["width"] = col
            elif "depth" in cl:
                col_map["depth"] = col

        current_vendor = ""
        for _, row in df.iterrows():
            vendor = str(row.get(col_map.get("vendor", ""), "")).strip()
            if vendor and vendor.lower() not in ["vendor", "nan", ""]:
                current_vendor = vendor
            elif not vendor:
                vendor = current_vendor

            series = str(row.get(col_map.get("series", ""), "")).strip()
            if not series or series.lower() in ["series", "nan", ""]:
                continue

            width_str = str(row.get(col_map.get("width", ""), "")).strip()
            depth_str = str(row.get(col_map.get("depth", ""), "")).strip()

            self.stile_widths.append(StileWidth(
                vendor=vendor,
                model=str(row.get(col_map.get("model", ""), "")).strip(),
                series=series,
                width=self._parse_dimension(width_str),
                depth=self._parse_dimension(depth_str),
            ))

    def _parse_dimension(self, s: str) -> Optional[float]:
        """Parse a dimension string like '3.5\"' or '2.125\"' to float inches."""
        if not s or s.lower() in ["nan", ""]:
            return None
        s = s.replace('"', '').replace("'", "").strip()
        # Handle fractions
        m = re.match(r'(\d+)\s+(\d+)/(\d+)', s)
        if m:
            return int(m.group(1)) + int(m.group(2)) / int(m.group(3))
        try:
            return float(s)
        except ValueError:
            return None

    # ── Query Methods ──

    def get_all_rules(self) -> List[Rule]:
        return self.rules

    def get_rules_by_category(self, category: str) -> List[Rule]:
        cat = category.lower()
        return [r for r in self.rules if cat in r.category.lower()]

    def get_rules_by_severity(self, severity: str) -> List[Rule]:
        sev = severity.lower()
        return [r for r in self.rules if r.severity.lower() == sev]

    def get_rules_for_door(self, door_material: str = "", door_location: str = "",
                           has_glazing: bool = False, has_panic: bool = False,
                           is_fire_rated: bool = False, has_access_control: bool = False,
                           has_auto_operator: bool = False) -> List[Rule]:
        """Get all rules that could apply to a specific door context."""
        return [
            r for r in self.rules
            if r.matches_door_context(
                door_material, door_location, has_glazing, has_panic,
                is_fire_rated, has_access_control, has_auto_operator
            )
        ]

    def lookup_stile(self, vendor: str, series: str) -> Optional[StileWidth]:
        """Look up a specific manufacturer's stile width by vendor and series."""
        v = vendor.lower().strip()
        s = series.lower().strip()
        for sw in self.stile_widths:
            if sw.vendor.lower().strip() == v and sw.series.lower().strip() == s:
                return sw
        # Try partial match on series
        for sw in self.stile_widths:
            if sw.vendor.lower().strip() == v and s in sw.series.lower():
                return sw
        return None

    def lookup_stile_by_width(self, vendor: str, width: float, tolerance: float = 0.25) -> List[StileWidth]:
        """Find stile entries matching a vendor and approximate width."""
        v = vendor.lower().strip()
        return [
            sw for sw in self.stile_widths
            if sw.vendor.lower().strip() == v
            and sw.width is not None
            and abs(sw.width - width) <= tolerance
        ]

    def get_vendors(self) -> List[str]:
        """Get list of unique vendors in the stile database."""
        vendors = set()
        for sw in self.stile_widths:
            if sw.vendor:
                vendors.add(sw.vendor)
        return sorted(vendors)

    def get_stile_widths_for_vendor(self, vendor: str) -> List[StileWidth]:
        """Get all stile entries for a specific vendor."""
        v = vendor.lower().strip()
        return [sw for sw in self.stile_widths if sw.vendor.lower().strip() == v]

    def summary(self) -> Dict:
        """Get a summary of loaded rules and data."""
        categories = {}
        severities = {}
        for r in self.rules:
            categories[r.category] = categories.get(r.category, 0) + 1
            severities[r.severity] = severities.get(r.severity, 0) + 1

        return {
            "total_rules": len(self.rules),
            "categories": categories,
            "severities": severities,
            "stile_entries": len(self.stile_widths),
            "vendors": self.get_vendors(),
            "source": self.source_file,
            "errors": self.load_errors,
        }
