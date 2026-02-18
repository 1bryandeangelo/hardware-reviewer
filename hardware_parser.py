"""
Hardware Schedule Parser
=========================
Parses Section 08 71 00 Door Hardware specification PDFs to extract
hardware sets with all components.

Handles the standard DHI format:
  HEADING # XX - (DESCRIPTION)
  PROVIDE EACH SGL/PR DOOR(S) WITH THE FOLLOWING:
  QTY EA DESCRIPTION CATALOG_NUMBER FINISH MFR
  ...
  OPERATIONAL DESCRIPTION: ...
"""

import re
from typing import List, Dict, Optional
from dataclasses import dataclass, field

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False


@dataclass
class HardwareComponent:
    """A single hardware item within a hardware set."""
    qty: int
    unit: str                    # EA, SET, PR
    description: str             # e.g. "PANIC HARDWARE"
    catalog_number: str          # e.g. "CD-33A-NL-OP"
    finish: str                  # e.g. "626"
    manufacturer: str            # e.g. "VON"
    raw_line: str = ""           # Original text

    def to_dict(self) -> Dict:
        return {
            "qty": self.qty,
            "unit": self.unit,
            "description": self.description,
            "catalog_number": self.catalog_number,
            "finish": self.finish,
            "manufacturer": self.manufacturer,
        }


@dataclass
class HardwareSet:
    """A complete hardware set (heading) with all components."""
    set_number: str
    description: str
    door_type: str                           # SGL, PR, RU
    components: List[HardwareComponent] = field(default_factory=list)
    operational_description: str = ""
    notes: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict:
        return {
            "set_number": self.set_number,
            "description": self.description,
            "door_type": self.door_type,
            "components": [c.to_dict() for c in self.components],
            "operational_description": self.operational_description,
            "notes": self.notes,
        }

    def has_panic_hardware(self) -> bool:
        return any('PANIC' in c.description.upper() or 'EXIT' in c.description.upper()
                   for c in self.components)

    def has_closer(self) -> bool:
        return any('CLOSER' in c.description.upper() for c in self.components)

    def has_lockset(self) -> bool:
        return any(kw in c.description.upper() for c in self.components
                   for kw in ['LOCK', 'LOCKSET', 'PASSAGE SET', 'PRIVACY'])


@dataclass
class HardwareScheduleResult:
    """Result of parsing a hardware specification PDF."""
    hardware_sets: Dict[str, HardwareSet]
    source_file: str = ""
    total_sets: int = 0
    warnings: List[str] = field(default_factory=list)

    def summary(self) -> str:
        lines = [
            f"Hardware Schedule: {self.total_sets} sets parsed",
            f"Source: {self.source_file}" if self.source_file else "",
        ]
        for num, hw_set in sorted(self.hardware_sets.items()):
            panic = " [PANIC]" if hw_set.has_panic_hardware() else ""
            closer = " [CLOSER]" if hw_set.has_closer() else ""
            lines.append(f"  Set #{num}: {hw_set.description[:60]}{panic}{closer}")
        return "\n".join(l for l in lines if l)


class HardwareScheduleParser:
    """
    Parses Section 08 71 00 hardware specification PDFs.

    Usage:
        parser = HardwareScheduleParser()
        result = parser.parse_pdf("087100_Door_Hardware.pdf")
        for set_num, hw_set in result.hardware_sets.items():
            for comp in hw_set.components:
                print(f"  {comp.qty} {comp.description} - {comp.catalog_number}")
    """

    KNOWN_MANUFACTURERS = {
        'IVE', 'VON', 'SCH', 'GLY', 'LCN', 'ZER', 'SCE',
        'SAR', 'DOR', 'BES', 'HAG', 'STA', 'TRI', 'ROC',
        'ADA', 'PRE', 'RCI', 'NGP', 'REE', 'NOR', 'FAL',
        'RKW', 'CAL', 'KAB', 'PEM', 'PDQ', 'YAL',
        'DET', 'LGR',  # Detex, Securitron (alt finish code)
    }

    MFR_PATTERN = re.compile(r'\s+(\S+)\s+([A-Z]{2,4})\s*$')

    TYPE_KEYWORDS = {
        'panic_hardware': ['PANIC HARDWARE', 'EXIT HARDWARE', 'FIRE EXIT', 'EXIT DEVICE'],
        'lockset': ['MORTISE LOCK', 'STOREROOM LOCK', 'CLASSROOM LOCK', 'PRIVACY LOCK',
                    'PASSAGE SET', 'OFFICE LOCK', 'ENTRANCE LOCK', 'LOCKSET', 'EU MORTISE LOCK'],
        'closer': ['SURFACE CLOSER', 'DOOR CLOSER', 'CLOSER'],
        'hinge': ['HINGE'],
        'cylinder': ['CYLINDER', 'PERMANENT CORE', 'CORE'],
        'pull': ['DOOR PULL', 'OFFSET DOOR PULL', 'PUSH/PULL', 'PUSH PLATE', 'PUSH BAR'],
        'stop': ['OVERHEAD STOP', 'WALL STOP', 'FLOOR STOP', 'STOP'],
        'kick_plate': ['KICK PLATE', 'PROTECTION PLATE', 'ARMOR PLATE'],
        'threshold': ['THRESHOLD'],
        'weatherstrip': ['SEALS', 'SEAL', 'DOOR SWEEP', 'GASKETING', 'ASTRAGAL', 'DRIP CAP'],
        'flush_bolt': ['FLUSH BOLT'],
        'silencer': ['SILENCER'],
        'power_transfer': ['POWER TRANSFER'],
        'power_supply': ['POWER SUPPLY'],
        'contact': ['DOOR CONTACT'],
        'card_reader': ['CARD READER'],
        'coordinator': ['COORDINATOR'],
    }

    def __init__(self, debug: bool = False):
        self.debug = debug

    def _log(self, msg: str):
        if self.debug:
            print(f"[HW-DEBUG] {msg}")

    def parse_pdf(self, pdf_path: str) -> HardwareScheduleResult:
        """Parse a hardware specification PDF and extract all hardware sets."""
        if not PDFPLUMBER_AVAILABLE:
            raise ImportError("pdfplumber is required")

        full_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

        full_text = self._clean_text(full_text)
        hardware_sets = self._parse_hardware_sets(full_text)

        result = HardwareScheduleResult(
            hardware_sets=hardware_sets,
            source_file=pdf_path,
            total_sets=len(hardware_sets),
        )

        for num, hw_set in hardware_sets.items():
            if not hw_set.components:
                result.warnings.append(f"Set #{num} has no components parsed")

        return result

    def _clean_text(self, text: str) -> str:
        """Remove page headers/footers."""
        lines = text.split('\n')
        cleaned = []
        for line in lines:
            stripped = line.strip()
            if re.match(r'^08\s*71\s*00\s*-\s*\d+', stripped):
                continue
            if stripped == 'DOOR HARDWARE':
                continue
            if stripped == 'QTY DESCRIPTION CATALOG NUMBER FINISH MFR':
                continue
            cleaned.append(line)
        return '\n'.join(cleaned)

    def _parse_hardware_sets(self, text: str) -> Dict[str, HardwareSet]:
        """Parse all hardware sets from the document text."""
        sets = {}
        text = self._rejoin_wrapped_headings(text)

        # Try Format 1: "HEADING # XX - (DESCRIPTION)"
        heading_pattern = r'HEADING\s*#\s*(\d+)\s*-\s*\(([^)]+)\)'
        matches = list(re.finditer(heading_pattern, text))

        if matches:
            self._log(f"Detected format: HEADING # (found {len(matches)} sets)")
            for i, match in enumerate(matches):
                set_num = match.group(1).lstrip('0') or '0'
                description = re.sub(r'\s+', ' ', match.group(2).strip())
                start = match.end()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                content = text[start:end]
                hw_set = self._parse_single_set(set_num, description, content)
                sets[set_num] = hw_set
                self._log(f"Set #{set_num}: {len(hw_set.components)} components")
            return sets

        # Try Format 2: "Hardware Group No. XX"
        group_pattern = r'Hardware\s+Group\s+No\.\s*(\d+)'
        matches = list(re.finditer(group_pattern, text, re.IGNORECASE))

        if matches:
            self._log(f"Detected format: Hardware Group No. (found {len(matches)} sets)")
            for i, match in enumerate(matches):
                set_num = match.group(1).lstrip('0') or '0'
                start = match.end()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
                content = text[start:end]

                # Extract description from "For use on Door #(s):" line
                door_ref_match = re.search(r'For\s+use\s+on\s+Door\s*#?\(s\)\s*:\s*\n?\s*(.+?)(?:\n|$)', content, re.IGNORECASE)
                door_refs = door_ref_match.group(1).strip() if door_ref_match else ''
                description = f"Doors: {door_refs}" if door_refs else ""

                hw_set = self._parse_single_set(set_num, description, content)
                sets[set_num] = hw_set
                self._log(f"Set #{set_num}: {len(hw_set.components)} components - {description}")
            return sets

        # Try Format 3: Generic fallback - look for numbered sections with QTY EA patterns
        self._log("No standard format detected, trying generic pattern match")
        return sets

    def _rejoin_wrapped_headings(self, text: str) -> str:
        """Rejoin heading lines that wrapped across multiple lines."""
        lines = text.split('\n')
        result = []
        i = 0
        while i < len(lines):
            line = lines[i]
            if re.match(r'.*HEADING\s*#\s*\d+\s*-\s*\(', line) and ')' not in line.split('(', 1)[-1]:
                combined = line
                while i + 1 < len(lines) and ')' not in combined.split('(', 1)[-1]:
                    i += 1
                    combined += ' ' + lines[i].strip()
                result.append(combined)
            else:
                result.append(line)
            i += 1
        return '\n'.join(result)

    def _parse_single_set(self, set_num: str, description: str,
                          content: str) -> HardwareSet:
        """Parse a single hardware set."""
        door_type = "SGL"
        provide_match = re.search(r'PROVIDE\s+EACH\s+(SGL|PR|RU)\s', content)
        if provide_match:
            door_type = provide_match.group(1)

        components = []
        op_desc_lines = []
        notes = []
        lines = content.split('\n')
        i = 0
        in_op_desc = False

        while i < len(lines):
            line = lines[i].strip()
            i += 1

            if not line:
                continue

            if 'OPERATIONAL DESCRIPTION' in line:
                in_op_desc = True
                after = line.split(':', 1)[-1].strip() if ':' in line else ''
                if after:
                    op_desc_lines.append(after)
                continue

            if in_op_desc:
                if re.match(r'HEADING\s*#', line) or line.startswith('PROVIDE EACH'):
                    break
                op_desc_lines.append(line)
                continue

            if 'ALL WIRING' in line or 'DIVISION 26' in line:
                notes.append(line)
                continue

            comp_match = re.match(r'^(\d+)\s+(EA|SET|PR|BALANCE)\s+(.+?)$', line)
            if comp_match:
                qty = int(comp_match.group(1))
                unit = comp_match.group(2)
                rest = comp_match.group(3).strip()

                # Join continuation lines
                while i < len(lines):
                    next_line = lines[i].strip()
                    if not next_line:
                        break
                    if re.match(r'^\d+\s+(EA|SET|PR|BALANCE)', next_line):
                        break
                    if any(kw in next_line for kw in ['OPERATIONAL', 'ALL WIRING',
                           'HEADING #', 'PROVIDE EACH', 'DIVISION 26']):
                        break
                    rest += ' ' + next_line
                    i += 1

                component = self._parse_component_line(qty, unit, rest)
                components.append(component)

        return HardwareSet(
            set_number=set_num,
            description=description,
            door_type=door_type,
            components=components,
            operational_description=' '.join(op_desc_lines).strip(),
            notes=notes,
        )

    def _parse_component_line(self, qty: int, unit: str,
                              rest: str) -> HardwareComponent:
        """Parse component text into description, catalog, finish, manufacturer."""
        m = self.MFR_PATTERN.search(rest)
        if m and m.group(2) in self.KNOWN_MANUFACTURERS:
            finish = m.group(1)
            manufacturer = m.group(2)
            desc_and_catalog = rest[:m.start()].strip()
        else:
            finish = ""
            manufacturer = ""
            desc_and_catalog = rest.strip()

        description, catalog_number = self._split_desc_catalog(desc_and_catalog)

        return HardwareComponent(
            qty=qty, unit=unit, description=description,
            catalog_number=catalog_number, finish=finish,
            manufacturer=manufacturer, raw_line=f"{qty} {unit} {rest}",
        )

    def _split_desc_catalog(self, text: str) -> tuple:
        """Split combined text into description and catalog number."""
        for type_name, keywords in self.TYPE_KEYWORDS.items():
            for kw in sorted(keywords, key=len, reverse=True):
                if text.upper().startswith(kw):
                    desc = kw
                    catalog = text[len(kw):].strip()
                    return (desc, catalog)

        # Fallback: split at first word containing digits
        words = text.split()
        desc_parts = []
        cat_parts = []
        found_cat = False
        for word in words:
            if not found_cat and not re.search(r'\d', word) and word.upper() == word:
                desc_parts.append(word)
            else:
                found_cat = True
                cat_parts.append(word)

        if desc_parts and cat_parts:
            return (' '.join(desc_parts), ' '.join(cat_parts))
        elif desc_parts:
            return (' '.join(desc_parts), '')
        return (text, '')

    def classify_component(self, component: HardwareComponent) -> str:
        """Classify a component into a standard type category."""
        desc = component.description.upper()
        for type_name, keywords in self.TYPE_KEYWORDS.items():
            for kw in keywords:
                if kw in desc:
                    return type_name
        return 'other'
