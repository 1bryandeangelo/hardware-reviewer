"""
Door Schedule Parser
====================
Extracts structured door data from PDF door schedules.

Uses pdfplumber for table extraction. Handles common variations in 
column naming, multi-page schedules, and merged cells.

Built for: Me
Author: Bryan (with Claude)
"""

import re
import json
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field, asdict

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


# ─────────────────────────────────────────────
# Data Models
# ─────────────────────────────────────────────

@dataclass
class DoorEntry:
    """Represents a single door from the schedule."""
    door_number: str
    width: str = ""
    height: str = ""
    thickness: str = ""
    material: str = ""          # aluminum, wood, hollow metal, etc.
    door_type: str = ""         # type code from schedule (VA1, VW1, FM1, etc.)
    manufacturer: str = ""
    product_line: str = ""      # narrow/medium/wide stile for aluminum
    frame_material: str = ""
    frame_type: str = ""
    glazing: str = ""
    fire_rating: str = ""
    hardware_set: str = ""
    finish: str = ""
    room_name: str = ""
    comments: str = ""
    raw_data: Dict = field(default_factory=dict)  # Original row data for debugging

    def to_dict(self) -> Dict:
        d = asdict(self)
        d.pop('raw_data', None)
        return {k: v for k, v in d.items() if v}

    def to_checker_format(self) -> Dict:
        """Convert to the format expected by the compatibility checker."""
        result = {
            "door_number": self.door_number,
            "material": self._normalize_material(),
            "hardware_set": self.hardware_set,
        }
        if self.product_line:
            result["product_line"] = self.product_line
        if self.manufacturer:
            result["manufacturer"] = self.manufacturer
        if self.thickness:
            result["thickness_in"] = self._parse_thickness()
        if self.width and self.height:
            result["width"] = self.width
            result["height"] = self.height
        if self.fire_rating:
            result["fire_rating"] = self.fire_rating
        if self.glazing:
            result["glazing"] = self.glazing
        if self.door_type:
            result["door_type"] = self.door_type
        return result

    def _normalize_material(self) -> str:
        """Normalize material names to standard format."""
        mat = self.material.lower().strip()
        if any(w in mat for w in ["alum"]):
            return "aluminum"
        elif any(w in mat for w in ["wood", "wd"]):
            return "wood"
        elif any(w in mat for w in ["metal", "hm", "hollow", "steel"]):
            return "hollow_metal"
        elif any(w in mat for w in ["fib", "frp"]):
            return "fiberglass"
        return mat if mat else "unknown"

    def _parse_thickness(self) -> float:
        """Parse thickness string to inches as float."""
        t = self.thickness.strip().replace('"', '').replace("'", '')
        # Handle "1 3/4" format
        match = re.match(r'(\d+)\s+(\d+)/(\d+)', t)
        if match:
            whole = int(match.group(1))
            num = int(match.group(2))
            den = int(match.group(3))
            return whole + num / den
        # Handle "1.75" format
        try:
            return float(t)
        except ValueError:
            return 1.75  # Default standard door thickness


@dataclass
class ParseResult:
    """Result of parsing a door schedule."""
    doors: List[DoorEntry]
    column_mapping: Dict[str, str]  # detected columns -> standard names
    warnings: List[str]
    raw_headers: List[str]
    page_count: int
    source_file: str = ""

    def summary(self) -> str:
        lines = [
            f"Parsed {len(self.doors)} doors from {self.page_count} page(s)",
            f"Source: {self.source_file}" if self.source_file else "",
            f"Columns detected: {', '.join(self.column_mapping.values())}",
        ]
        if self.warnings:
            lines.append(f"Warnings: {len(self.warnings)}")
            for w in self.warnings:
                lines.append(f"  - {w}")

        # Material breakdown
        materials = {}
        for d in self.doors:
            mat = d._normalize_material()
            materials[mat] = materials.get(mat, 0) + 1
        if materials:
            lines.append("Door materials:")
            for mat, count in sorted(materials.items()):
                lines.append(f"  {mat}: {count}")

        # Hardware sets
        hw_sets = set(d.hardware_set for d in self.doors if d.hardware_set)
        if hw_sets:
            lines.append(f"Hardware sets referenced: {', '.join(sorted(hw_sets))}")

        return "\n".join(l for l in lines if l)


# ─────────────────────────────────────────────
# Column Mapping Logic
# ─────────────────────────────────────────────

# Maps common header variations to standard field names
COLUMN_ALIASES = {
    "door_number": [
        "door", "door #", "door no", "door no.", "door num", "door number",
        "dr", "dr #", "dr no", "mark", "mark #", "no", "no.", "#",
        "opening", "opening #", "opening no",
    ],
    "width": [
        "width", "w", "wd", "door width", "dr width", "size w",
    ],
    "height": [
        "height", "h", "ht", "hgt", "door height", "dr height", "size h",
    ],
    "thickness": [
        "thickness", "thk", "thick", "door thickness", "dr thk",
        "door thk", "thkns",
    ],
    "material": [
        "material", "mat", "matl", "door material", "dr material",
        "door mat", "dr mat", "door matl",
    ],
    "door_type": [
        "type", "door type", "dr type", "detail", "detail #",
        "type/detail", "elev", "elevation",
    ],
    "frame_material": [
        "frame", "frame material", "frame mat", "frame matl",
        "fr material", "fr mat", "fr matl", "frame type",
    ],
    "fire_rating": [
        "fire", "fire rating", "fire rate", "fire rated", "fr",
        "fire label", "label", "rating", "fire rtg",
    ],
    "hardware_set": [
        "hardware", "hardware set", "hdw", "hdw set", "hw", "hw set",
        "hw grp", "hw group", "hardware group", "hdw grp", "hrdwr",
        "hrdwr.", "hrdwr set", "hrdwr. set", "hrdwr.set", "hard set",
        "hardware #", "hw #", "hdw #", "hrdwr #", "hrdwr. #",
        "hdwr", "hdwr.", "hdwr set",
    ],
    "glazing": [
        "glazing", "glass", "glaz", "glazing type", "glass type",
        "lite", "vision",
    ],
    "finish": [
        "finish", "fin", "door finish", "dr finish",
    ],
    "room_name": [
        "room", "room name", "location", "loc", "space", "area",
        "room/location",
    ],
    "comments": [
        "comments", "notes", "remarks", "comment", "note", "remark",
    ],
    "manufacturer": [
        "manufacturer", "mfr", "mfg", "manuf",
    ],
    "frame_finish": [
        "frame finish", "fr finish", "frame fin",
    ],
}


def _normalize_header(header: str) -> str:
    """Normalize a header string for matching."""
    if not header:
        return ""
    h = header.lower().strip()
    h = re.sub(r'[\n\r]+', ' ', h)  # Replace newlines with space
    h = re.sub(r'\s+', ' ', h)      # Collapse whitespace
    h = re.sub(r'[^a-z0-9# /.]', '', h)  # Keep only alphanumeric + a few chars
    return h.strip()


def map_columns(headers: List[str]) -> Dict[int, str]:
    """
    Map detected table headers to standard field names.
    
    Returns: Dict mapping column index -> standard field name
    """
    mapping = {}
    used_fields = set()

    normalized = [_normalize_header(h) for h in headers]

    for idx, norm_header in enumerate(normalized):
        if not norm_header:
            continue

        best_match = None
        best_score = 0

        for field_name, aliases in COLUMN_ALIASES.items():
            if field_name in used_fields:
                continue
            for alias in aliases:
                # Exact match
                if norm_header == alias:
                    best_match = field_name
                    best_score = 100
                    break
                # Header contains alias
                if alias in norm_header and len(alias) > best_score:
                    best_match = field_name
                    best_score = len(alias)
                # Alias contains header (for short headers like "#")
                if norm_header in alias and len(norm_header) >= 1 and best_score < 1:
                    best_match = field_name
                    best_score = 0.5

            if best_score == 100:
                break

        if best_match:
            mapping[idx] = best_match
            used_fields.add(best_match)

    return mapping


# ─────────────────────────────────────────────
# Size Parsing
# ─────────────────────────────────────────────

def parse_door_size(size_str: str) -> Tuple[str, str]:
    """
    Parse a combined door size string into width and height.
    
    Handles formats like:
      "3'-0\" x 7'-0\""
      "3'0\" x 7'0\""
      "3-0 x 7-0"
      "36 x 84"
      "3068" (30" wide x 68" tall - but 68 is 6'8")
    """
    if not size_str:
        return ("", "")

    s = size_str.strip()

    # Pattern: W'[-]H" x W'[-]H" (architectural notation)
    arch_pattern = r"(\d+['\s-]+\d+[\"\s]*)\s*[xX×]\s*(\d+['\s-]+\d+[\"\s]*)"
    m = re.search(arch_pattern, s)
    if m:
        return (m.group(1).strip(), m.group(2).strip())

    # Pattern: WW x HH (inches)
    inch_pattern = r"(\d+)\s*[xX×]\s*(\d+)"
    m = re.search(inch_pattern, s)
    if m:
        w_in = int(m.group(1))
        h_in = int(m.group(2))
        return (_inches_to_arch(w_in), _inches_to_arch(h_in))

    # Pattern: WWHH (4-digit shorthand, e.g. 3070 = 3'-0" x 7'-0")
    if re.match(r'^\d{4}$', s):
        w = int(s[:2])
        h = int(s[2:])
        return (_inches_to_arch(w), _inches_to_arch(h))

    return (s, "")


def _inches_to_arch(inches: int) -> str:
    """Convert inches to architectural format (e.g., 36 -> 3'-0\")."""
    feet = inches // 12
    remaining = inches % 12
    return f"{feet}'-{remaining}\""


# ─────────────────────────────────────────────
# Main Parser
# ─────────────────────────────────────────────

class DoorScheduleParser:
    """
    Parses door schedule PDFs into structured DoorEntry objects.
    
    Usage:
        parser = DoorScheduleParser()
        result = parser.parse_pdf("door_schedule.pdf")
        
        for door in result.doors:
            print(door.door_number, door.material, door.hardware_set)
    """

    def __init__(self, debug: bool = False):
        self.debug = debug
        self._log_lines = []

    def _log(self, msg: str):
        if self.debug:
            self._log_lines.append(msg)
            print(f"[DEBUG] {msg}")

    # ── PDF Extraction ──

    def parse_pdf(self, pdf_path: str) -> ParseResult:
        """
        Parse a door schedule PDF and return structured door data.

        Handles common architectural PDF patterns:
        - Multiple tables on the same page (room finish, material list, etc.)
        - Multi-row headers with merged cells ("Door" spanning Width/Height/etc.)
        - Headers split into a separate table from the data rows
        - "DOOR SCHEDULE" title row above the actual column headers
        - Continuation tables on subsequent pages

        Args:
            pdf_path: Path to the door schedule PDF

        Returns:
            ParseResult with extracted doors and metadata
        """
        if not PDFPLUMBER_AVAILABLE:
            raise ImportError("pdfplumber is required. Install with: pip install pdfplumber")

        all_rows = []
        headers = None
        expected_cols = None
        page_count = 0
        warnings = []

        with pdfplumber.open(pdf_path) as pdf:
            page_count = len(pdf.pages)

            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                if not tables:
                    self._log(f"Page {page_num + 1}: No tables found")
                    text_rows = self._extract_from_text(page)
                    if text_rows:
                        all_rows.extend(text_rows)
                    continue

                self._log(f"Page {page_num + 1}: {len(tables)} tables found")

                if headers is None:
                    # First page — need to find the door schedule
                    result = self._find_door_schedule_tables(tables)
                    if result is None:
                        self._log(f"Page {page_num + 1}: Could not identify door schedule")
                        continue

                    headers, data_rows = result
                    expected_cols = len(headers)
                    self._log(f"Found headers ({expected_cols} cols): {headers}")

                    for row in data_rows:
                        if self._is_data_row(row):
                            all_rows.append(row)
                else:
                    # Continuation pages — look for tables with matching column count
                    for table in tables:
                        if not table or len(table) < 1:
                            continue
                        # Check if first row matches headers (repeated header)
                        if self._rows_match(table[0], headers):
                            rows = table[1:]
                        elif expected_cols and len(table[0]) == expected_cols:
                            rows = table
                        else:
                            continue

                        for row in rows:
                            if self._is_data_row(row):
                                all_rows.append(row)

        if headers is None:
            return ParseResult(
                doors=[],
                column_mapping={},
                warnings=["No door schedule table found in PDF"],
                raw_headers=[],
                page_count=page_count,
                source_file=pdf_path,
            )

        # Map columns
        col_mapping = map_columns(headers)
        self._log(f"Column mapping: {col_mapping}")

        # Check for combined size column
        has_separate_wh = any(col_mapping.get(i) == "width" for i in col_mapping) and \
                          any(col_mapping.get(i) == "height" for i in col_mapping)
        size_col = None
        if not has_separate_wh:
            for idx, h in enumerate(headers):
                nh = _normalize_header(h)
                if nh in ["size", "door size", "opening size", "opng size"]:
                    size_col = idx
                    break

        # Build door entries
        doors = []
        for row_idx, row in enumerate(all_rows):
            door = self._row_to_door(row, col_mapping, size_col)
            if door:
                doors.append(door)
            else:
                self._log(f"Row {row_idx} skipped: {row}")

        # Detect warnings
        if not any(d.hardware_set for d in doors):
            warnings.append("No hardware set column detected — hardware compatibility checking will not be possible")
        if not any(d.material for d in doors):
            warnings.append("No material column detected — some compatibility checks may be limited")

        readable_mapping = {}
        for idx, field_name in col_mapping.items():
            if idx < len(headers):
                readable_mapping[headers[idx] or f"Column {idx}"] = field_name

        return ParseResult(
            doors=doors,
            column_mapping=readable_mapping,
            warnings=warnings,
            raw_headers=headers,
            page_count=page_count,
            source_file=pdf_path,
        )

    def _find_door_schedule_tables(self, tables: List[List[List]]) -> Optional[tuple]:
        """
        Find the door schedule among multiple tables on a page.

        Architectural drawings often have several tables on one sheet
        (door schedule, room finish schedule, material reference list, etc.)
        plus the headers may be in a separate table from the data.

        Strategy:
        1. Look for a table containing "DOOR SCHEDULE" — that's the header table
        2. Merge multi-row headers into a single header row
        3. If data rows are in the same table, use them
        4. If not, find the next table with matching column count — that's the data

        Returns:
            (merged_headers, data_rows) or None if not found
        """
        # Strategy 1: Find a table with "DOOR SCHEDULE" title
        for t_idx, table in enumerate(tables):
            if not table:
                continue
            for row in table[:3]:  # Check first 3 rows for a title
                for cell in row:
                    if cell and "DOOR SCHEDULE" in str(cell).upper().replace("\n", " "):
                        self._log(f"Found 'DOOR SCHEDULE' in table {t_idx}")
                        result = self._extract_from_door_schedule_table(table, t_idx, tables)
                        if result:
                            return result

        # Strategy 2: No "DOOR SCHEDULE" label — find the table with the most
        # recognized column headers (the old approach, improved)
        best_result = None
        best_score = 0

        for t_idx, table in enumerate(tables):
            if not table or len(table) < 2:
                continue

            # Try this table as a standalone schedule (headers + data in one table)
            headers = self._find_header_row(table)
            if headers:
                score = self._score_header_row(headers)
                if score > best_score:
                    header_idx = table.index(headers)
                    data_rows = table[header_idx + 1:]
                    if any(self._is_data_row(r) for r in data_rows):
                        best_score = score
                        best_result = (headers, data_rows)

            # Also try merging multi-row headers within this table
            merged = self._try_merge_multi_row_headers(table)
            if merged:
                m_headers, m_data = merged
                m_score = self._score_header_row(m_headers)
                if m_score > best_score and any(self._is_data_row(r) for r in m_data):
                    best_score = m_score
                    best_result = (m_headers, m_data)

        return best_result

    def _extract_from_door_schedule_table(self, header_table: List[List],
                                           header_table_idx: int,
                                           all_tables: List[List[List]]) -> Optional[tuple]:
        """
        Given a table that contains "DOOR SCHEDULE", extract headers and find data.
        The data might be in the same table or in an adjacent table.
        """
        # Merge multi-row headers
        merged_headers, remaining_rows = self._merge_header_rows(header_table)

        if not merged_headers:
            return None

        col_count = len(merged_headers)
        self._log(f"Merged headers ({col_count} cols): {merged_headers}")

        # Check if there are data rows in the same table
        data_rows = [r for r in remaining_rows if self._is_data_row(r)]
        if data_rows:
            self._log(f"Data found in same table: {len(data_rows)} rows")
            return (merged_headers, data_rows)

        # No data in the header table — look for the next table with matching column count
        for t_idx in range(header_table_idx + 1, len(all_tables)):
            candidate = all_tables[t_idx]
            if not candidate or len(candidate) < 1:
                continue
            if len(candidate[0]) == col_count:
                data_rows = [r for r in candidate if self._is_data_row(r)]
                if data_rows:
                    self._log(f"Data found in table {t_idx}: {len(data_rows)} rows")
                    return (merged_headers, data_rows)

        # Still no data — try tables with close column count (off by 1-2)
        for t_idx in range(header_table_idx + 1, len(all_tables)):
            candidate = all_tables[t_idx]
            if not candidate or len(candidate) < 1:
                continue
            if abs(len(candidate[0]) - col_count) <= 2 and len(candidate) >= 3:
                data_rows = [r for r in candidate if self._is_data_row(r)]
                if len(data_rows) >= 3:
                    self._log(f"Data found in table {t_idx} (close col count): {len(data_rows)} rows")
                    # Pad or trim rows to match header count
                    adjusted = []
                    for row in data_rows:
                        if len(row) < col_count:
                            adjusted.append(row + [''] * (col_count - len(row)))
                        elif len(row) > col_count:
                            adjusted.append(row[:col_count])
                        else:
                            adjusted.append(row)
                    return (merged_headers, adjusted)

        return None

    def _merge_header_rows(self, table: List[List]) -> tuple:
        """
        Merge multi-row headers into a single row.

        Handles patterns like:
          Row 0: ["DOOR SCHEDULE", None, None, ...]           (title)
          Row 1: ["Door\\nNumber", "Type", "Door", None, ..., "Frame", None, None]  (top-level groups)
          Row 2: [None, None, "Width", "Height", ..., "Hardware", "Material", "Finish"]  (sub-headers)

        Returns:
            (merged_header_row, remaining_data_rows)
        """
        if not table or len(table) < 2:
            return (None, [])

        # Find where headers end and data begins.
        # Header rows typically have: title text, column group names, sub-column names.
        # Data rows have: door numbers, dimensions, materials.

        # Step 1: Skip title rows (rows where first cell contains "SCHEDULE" or
        # the row is mostly empty/None with one label cell)
        header_rows = []
        data_start = 0

        for i, row in enumerate(table):
            row_text = ' '.join(str(c or '').strip() for c in row).upper()

            # Is this a title row? (contains "SCHEDULE" or "DOOR SCHEDULE")
            if 'SCHEDULE' in row_text and i < 3:
                data_start = i + 1
                continue

            # Is this a header-like row? Check if cells are mostly text labels
            # vs data rows which have numbers, dimensions, abbreviations
            non_empty = [str(c or '').strip() for c in row if c and str(c).strip()]
            if not non_empty:
                data_start = i + 1
                continue

            # Check if this looks like a data row (has door-number-like content in first cell)
            first_cell = str(row[0] or '').strip()
            if first_cell and re.match(r'^\d{2,4}[A-Za-z]?$', first_cell):
                # This looks like a door number — data starts here
                break

            # This is probably a header row
            header_rows.append(row)
            data_start = i + 1

        remaining = table[data_start:]

        if not header_rows:
            return (None, remaining)

        # Step 2: Merge header rows into one
        if len(header_rows) == 1:
            # Single header row — clean it up
            merged = [str(c or '').strip().replace('\n', ' ') for c in header_rows[0]]
            return (merged, remaining)

        # Multiple header rows — merge bottom-up (sub-headers take priority)
        col_count = max(len(r) for r in header_rows)
        merged = [''] * col_count

        # Start from the bottom row (most specific) and work up
        for row in reversed(header_rows):
            for i in range(min(len(row), col_count)):
                cell = str(row[i] or '').strip().replace('\n', ' ')
                if cell and not merged[i]:
                    merged[i] = cell

        return (merged, remaining)

    def _try_merge_multi_row_headers(self, table: List[List]) -> Optional[tuple]:
        """
        Try to detect and merge multi-row headers within a single table
        (even without a "DOOR SCHEDULE" title).
        """
        if len(table) < 3:
            return None

        merged, remaining = self._merge_header_rows(table)
        if not merged:
            return None

        # Verify the merged headers are actually good
        score = self._score_header_row(merged)
        if score >= 2:
            return (merged, remaining)

        return None

    def _score_header_row(self, headers: List) -> int:
        """Score how many recognized column names a header row contains."""
        score = 0
        for cell in headers:
            if not cell:
                continue
            norm = _normalize_header(str(cell))
            for aliases in COLUMN_ALIASES.values():
                if norm in aliases or any(a in norm for a in aliases if len(a) > 2):
                    score += 1
                    break
        return score

    def _find_header_row(self, table: List[List]) -> Optional[List]:
        """
        Find the header row in a table.
        Looks for the row that best matches known column names.
        """
        best_row = None
        best_score = 0

        for row in table[:5]:  # Check first 5 rows
            score = 0
            for cell in row:
                if not cell:
                    continue
                norm = _normalize_header(str(cell))
                for aliases in COLUMN_ALIASES.values():
                    if norm in aliases or any(a in norm for a in aliases if len(a) > 2):
                        score += 1
                        break
            if score > best_score:
                best_score = score
                best_row = row

        # Require at least 2 recognized columns
        return best_row if best_score >= 2 else None

    def _rows_match(self, row1: List, row2: List) -> bool:
        """Check if two rows are the same (header repeated on new page)."""
        if len(row1) != len(row2):
            return False
        matches = sum(1 for a, b in zip(row1, row2)
                      if _normalize_header(str(a or '')) == _normalize_header(str(b or '')))
        return matches >= len(row1) * 0.7

    def _is_data_row(self, row: List) -> bool:
        """Check if a row contains actual data (not empty/separator)."""
        non_empty = [c for c in row if c and str(c).strip()]
        return len(non_empty) >= 2

    def _row_to_door(self, row: List, col_mapping: Dict[int, str],
                     size_col: Optional[int] = None) -> Optional[DoorEntry]:
        """Convert a table row to a DoorEntry."""
        # Build a field dict from the row
        fields = {}
        raw = {}
        for idx, cell in enumerate(row):
            val = str(cell).strip() if cell else ""
            raw[f"col_{idx}"] = val
            if idx in col_mapping:
                field_name = col_mapping[idx]
                fields[field_name] = val

        # Must have a door number
        door_num = fields.get("door_number", "").strip()
        if not door_num:
            return None

        # Skip if the "door number" looks like a header or non-numeric junk
        if _normalize_header(door_num) in ["door", "door #", "door no", "#", "no"]:
            return None

        # Handle combined size column
        width = fields.get("width", "")
        height = fields.get("height", "")
        if not width and not height and size_col is not None:
            size_val = str(row[size_col]).strip() if size_col < len(row) and row[size_col] else ""
            if size_val:
                width, height = parse_door_size(size_val)

        door = DoorEntry(
            door_number=door_num,
            width=width,
            height=height,
            thickness=fields.get("thickness", ""),
            material=fields.get("material", ""),
            door_type=fields.get("door_type", ""),
            manufacturer=fields.get("manufacturer", ""),
            frame_material=fields.get("frame_material", ""),
            fire_rating=fields.get("fire_rating", ""),
            hardware_set=fields.get("hardware_set", ""),
            glazing=fields.get("glazing", ""),
            finish=fields.get("finish", ""),
            room_name=fields.get("room_name", ""),
            comments=fields.get("comments", ""),
            raw_data=raw,
        )

        # Infer material from door type code if not explicitly stated
        # If material IS provided, normalize it but don't override
        if door.material:
            # Keep the original — _normalize_material() handles conversion
            pass
        elif door.door_type:
            door.material = self._infer_material_from_type(door.door_type)

        return door

    def _infer_material_from_type(self, door_type: str) -> str:
        """
        Infer door material from type codes.
        Common conventions:
            VA = Vision Aluminum, VW = Vision Wood, VM = Vision Metal
            FA = Flush Aluminum, FW = Flush Wood, FM = Flush Metal
            SA = Storefront Aluminum
        """
        dt = door_type.upper().strip()
        if len(dt) >= 2:
            prefix = dt[:2]
            material_codes = {
                "VA": "aluminum", "SA": "aluminum", "AA": "aluminum",
                "VW": "wood", "FW": "wood", "SW": "wood",
                "VM": "hollow_metal", "FM": "hollow_metal", "HM": "hollow_metal",
                "SM": "hollow_metal",
            }
            # Try first two characters
            if prefix in material_codes:
                return material_codes[prefix]
            # Try just the second character (V=vision prefix, then material)
            if len(dt) >= 2:
                second = dt[1]
                single_codes = {"A": "aluminum", "W": "wood", "M": "hollow_metal"}
                if second in single_codes:
                    return single_codes[second]
        return ""

    def _extract_from_text(self, page) -> List[List]:
        """
        Fallback: try to extract door data from page text when no table is detected.
        This handles schedules that don't have clear table lines.
        """
        text = page.extract_text()
        if not text:
            return []

        rows = []
        for line in text.split('\n'):
            # Look for lines starting with a door number pattern
            match = re.match(r'^\s*(\d{2,4}[A-Z]?)\s+', line)
            if match:
                # Split line by multiple spaces (common in PDF text extraction)
                parts = re.split(r'\s{2,}', line.strip())
                if len(parts) >= 3:
                    rows.append(parts)

        return rows

    # ── CSV/Text Input ──

    def parse_csv(self, csv_path: str) -> ParseResult:
        """Parse a door schedule from CSV or TSV file."""
        if not PANDAS_AVAILABLE:
            raise ImportError("pandas is required for CSV parsing")

        # Auto-detect separator
        with open(csv_path, 'r') as f:
            first_line = f.readline()
        sep = '\t' if '\t' in first_line else ','

        df = pd.read_csv(csv_path, sep=sep, dtype=str).fillna('')
        return self._parse_dataframe(df, csv_path)

    def parse_excel(self, excel_path: str, sheet_name: int = 0) -> ParseResult:
        """Parse a door schedule from Excel file."""
        if not PANDAS_AVAILABLE:
            raise ImportError("pandas is required for Excel parsing")

        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str).fillna('')
        return self._parse_dataframe(df, excel_path)

    def _parse_dataframe(self, df, source: str) -> ParseResult:
        """Parse a pandas DataFrame into door entries."""
        headers = list(df.columns)
        col_mapping = map_columns(headers)

        # Convert index-based mapping to use column names
        idx_to_col = {i: headers[i] for i in col_mapping}

        doors = []
        warnings = []
        for _, row in df.iterrows():
            fields = {}
            for idx, field_name in col_mapping.items():
                col_name = headers[idx]
                fields[field_name] = str(row[col_name]).strip()

            door_num = fields.get("door_number", "").strip()
            if not door_num or _normalize_header(door_num) in ["door", "door #", ""]:
                continue

            door = DoorEntry(
                door_number=door_num,
                width=fields.get("width", ""),
                height=fields.get("height", ""),
                thickness=fields.get("thickness", ""),
                material=fields.get("material", ""),
                door_type=fields.get("door_type", ""),
                manufacturer=fields.get("manufacturer", ""),
                frame_material=fields.get("frame_material", ""),
                fire_rating=fields.get("fire_rating", ""),
                hardware_set=fields.get("hardware_set", ""),
                glazing=fields.get("glazing", ""),
                finish=fields.get("finish", ""),
                room_name=fields.get("room_name", ""),
                comments=fields.get("comments", ""),
            )

            if not door.material and door.door_type:
                door.material = self._infer_material_from_type(door.door_type)

            doors.append(door)

        readable_mapping = {headers[idx]: field_name for idx, field_name in col_mapping.items()}

        return ParseResult(
            doors=doors,
            column_mapping=readable_mapping,
            warnings=warnings,
            raw_headers=headers,
            page_count=1,
            source_file=source,
        )

    # ── Manual / Dict Input ──

    def parse_dict_list(self, doors_data: List[Dict]) -> ParseResult:
        """
        Parse door data from a list of dictionaries.
        Useful for testing or when data is already structured.
        """
        doors = []
        for d in doors_data:
            # Try to map whatever keys are provided
            door = DoorEntry(
                door_number=str(d.get("door_number", d.get("door", d.get("#", "")))),
                width=str(d.get("width", "")),
                height=str(d.get("height", "")),
                thickness=str(d.get("thickness", d.get("thk", ""))),
                material=str(d.get("material", d.get("mat", ""))),
                door_type=str(d.get("door_type", d.get("type", ""))),
                manufacturer=str(d.get("manufacturer", d.get("mfr", ""))),
                product_line=str(d.get("product_line", "")),
                frame_material=str(d.get("frame_material", d.get("frame", ""))),
                fire_rating=str(d.get("fire_rating", d.get("fire", ""))),
                hardware_set=str(d.get("hardware_set", d.get("hw_set", d.get("hw", "")))),
                glazing=str(d.get("glazing", d.get("glass", ""))),
                finish=str(d.get("finish", "")),
                room_name=str(d.get("room_name", d.get("room", d.get("location", "")))),
                comments=str(d.get("comments", d.get("notes", ""))),
            )
            if not door.material and door.door_type:
                door.material = self._infer_material_from_type(door.door_type)
            doors.append(door)

        return ParseResult(
            doors=doors,
            column_mapping={"direct_input": "dict"},
            warnings=[],
            raw_headers=[],
            page_count=0,
            source_file="manual_input",
        )


# ─────────────────────────────────────────────
# Export Utilities
# ─────────────────────────────────────────────

def export_to_checker_format(parse_result: ParseResult) -> List[Dict]:
    """Export parsed doors in the format expected by the compatibility checker."""
    return [door.to_checker_format() for door in parse_result.doors]


def export_to_json(parse_result: ParseResult, filepath: str):
    """Export parsed doors to a JSON file."""
    data = {
        "source": parse_result.source_file,
        "door_count": len(parse_result.doors),
        "column_mapping": parse_result.column_mapping,
        "warnings": parse_result.warnings,
        "doors": [door.to_dict() for door in parse_result.doors],
    }
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=2)


def export_to_csv(parse_result: ParseResult, filepath: str):
    """Export parsed doors to a CSV file."""
    if not PANDAS_AVAILABLE:
        raise ImportError("pandas is required for CSV export")

    rows = [door.to_dict() for door in parse_result.doors]
    df = pd.DataFrame(rows)
    df.to_csv(filepath, index=False)
