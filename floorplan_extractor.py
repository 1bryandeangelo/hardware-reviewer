"""
Floor Plan Door Number Extractor
==================================
Extracts door numbers from architectural floor plan PDFs and compares
them against a door schedule to catch discrepancies.

Uses PyMuPDF (fitz) for text extraction with position data.
"""

import re
from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass, field

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


@dataclass
class ExtractedDoor:
    """A door number found on a floor plan."""
    number: str
    page: int
    x: float = 0
    y: float = 0
    original_text: str = ""


@dataclass
class ComparisonResult:
    """Result of comparing floor plan doors vs door schedule."""
    schedule_doors: List[str]
    floorplan_doors: List[str]
    matched: List[str]                  # In both
    on_schedule_not_plan: List[str]     # On schedule but missing from plans
    on_plan_not_schedule: List[str]     # On plans but not in schedule
    duplicates_on_plan: Dict[str, int]  # Door numbers appearing multiple times

    def to_dict(self) -> Dict:
        return {
            "schedule_count": len(self.schedule_doors),
            "floorplan_count": len(self.floorplan_doors),
            "matched_count": len(self.matched),
            "matched": sorted(self.matched),
            "on_schedule_not_plan": sorted(self.on_schedule_not_plan),
            "on_plan_not_schedule": sorted(self.on_plan_not_schedule),
            "duplicates": self.duplicates_on_plan,
            "all_match": len(self.on_schedule_not_plan) == 0 and len(self.on_plan_not_schedule) == 0,
        }


class FloorPlanExtractor:
    """
    Extracts door numbers from floor plan PDFs.

    Usage:
        extractor = FloorPlanExtractor()
        doors = extractor.extract_from_pdf("floor_plan.pdf")
        result = extractor.compare(doors, schedule_door_numbers)
    """

    # Door number patterns — ordered from most specific to broadest
    PATTERNS = [
        # Standard: 101, 102, 201A, 100B
        re.compile(r'\b([1-9]\d{2}[A-Za-z]?)\b'),
        # With prefix: D-101, D101
        re.compile(r'\b[Dd]-?(\d{2,4}[A-Za-z]?)\b'),
        # Letter prefix: A-1, B-12 (less common)
        re.compile(r'\b([A-Z]-\d{1,3}[A-Za-z]?)\b'),
    ]

    # Context to exclude — these are NOT door numbers
    EXCLUDE_PATTERNS = [
        re.compile(r"^\d+['\"]"),            # Dimensions: 3', 7", 12'-6"
        re.compile(r"^\d+/\d+"),             # Fractions: 1/4, 3/8
        re.compile(r"^\d{4,5}$"),            # Very large numbers (areas, codes)
        re.compile(r"^[A-Z]{2,}-\d+"),       # Spec references: ACC-01, TL-02
        re.compile(r"^\d+\s*[xX×]\s*\d+"),   # Dimensions: 24 X 24
        re.compile(r"^[A-Z]-\d{4}"),         # Model numbers: B-4288
    ]

    def __init__(self):
        pass

    def extract_from_pdf(self, pdf_path: str) -> List[ExtractedDoor]:
        """
        Extract door numbers from a floor plan PDF.

        Args:
            pdf_path: Path to the floor plan PDF

        Returns:
            List of ExtractedDoor objects found
        """
        if not FITZ_AVAILABLE:
            raise ImportError("PyMuPDF (fitz) is required. Install with: pip install PyMuPDF")

        doors = []
        seen = set()  # Track (number, page) to avoid duplicates

        doc = fitz.open(pdf_path)
        for page_num, page in enumerate(doc):
            # Get text with position info
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                if "lines" not in block:
                    continue
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        if not text:
                            continue

                        # Check each pattern
                        for pattern in self.PATTERNS:
                            for match in pattern.finditer(text):
                                candidate = match.group(1) if match.lastindex else match.group(0)
                                candidate = candidate.strip()

                                if not candidate:
                                    continue

                                # Skip if it matches an exclusion pattern
                                if self._should_exclude(candidate, text):
                                    continue

                                key = (candidate.upper(), page_num)
                                if key not in seen:
                                    seen.add(key)
                                    bbox = span["bbox"]
                                    doors.append(ExtractedDoor(
                                        number=candidate.upper(),
                                        page=page_num + 1,
                                        x=bbox[0],
                                        y=bbox[1],
                                        original_text=text,
                                    ))

        doc.close()
        return doors

    def _should_exclude(self, candidate: str, context: str) -> bool:
        """Check if a candidate door number should be excluded."""
        for pattern in self.EXCLUDE_PATTERNS:
            if pattern.match(candidate):
                return True
            if pattern.match(context):
                return True

        # Exclude pure single/double digit numbers that are likely labels
        if re.match(r'^\d{1,2}$', candidate):
            # These are ok if they look like door numbers in context
            # But on floor plans, standalone small numbers are usually
            # room numbers, detail markers, grid lines, etc.
            # Only keep if the number is >= 100
            try:
                if int(candidate) < 100:
                    return True
            except ValueError:
                pass

        return False

    def compare(self, extracted: List[ExtractedDoor],
                schedule_doors: List[str]) -> ComparisonResult:
        """
        Compare extracted floor plan door numbers against a door schedule.

        Args:
            extracted: Door numbers found on floor plans
            schedule_doors: Door numbers from the door schedule

        Returns:
            ComparisonResult with matches and discrepancies
        """
        plan_numbers = [d.number.upper() for d in extracted]
        schedule_set = set(d.upper() for d in schedule_doors)
        plan_set = set(plan_numbers)

        matched = sorted(schedule_set & plan_set)
        on_schedule_not_plan = sorted(schedule_set - plan_set)
        on_plan_not_schedule = sorted(plan_set - schedule_set)

        # Find duplicates
        from collections import Counter
        counts = Counter(plan_numbers)
        duplicates = {num: count for num, count in counts.items() if count > 1}

        return ComparisonResult(
            schedule_doors=sorted(schedule_set),
            floorplan_doors=sorted(plan_set),
            matched=matched,
            on_schedule_not_plan=on_schedule_not_plan,
            on_plan_not_schedule=on_plan_not_schedule,
            duplicates_on_plan=duplicates,
        )
