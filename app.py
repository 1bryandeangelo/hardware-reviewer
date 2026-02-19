"""
Door Review Tool — Web Server (v2)
====================================
Spreadsheet-driven door hardware compatibility checker.

Rules come from the Excel spreadsheet (uploaded via admin page).
Door schedules parsed from PDF/CSV/Excel uploads.

Run locally:  python app.py
Deploy:       Push to Render with requirements.txt

Endpoints:
  GET  /                    — Main app
  GET  /admin               — Admin page (upload new rules)
  POST /api/parse-schedule  — Upload & parse door schedule
  POST /api/run-review      — Run compatibility review
  POST /api/upload-rules    — Upload new rules spreadsheet
  GET  /api/rules-summary   — Current rules summary
  GET  /api/health          — Health check
"""

import os
import json
import shutil
import tempfile
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory

from door_schedule_parser import DoorScheduleParser
from rules_engine import RulesEngine
from compatibility_checker import CompatibilityChecker
from floorplan_extractor import FloorPlanExtractor
from hardware_parser import HardwareScheduleParser

# ─────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────

app = Flask(__name__, static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# Persistent data directory (survives restarts on Render with disk)
DATA_DIR = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data'))
RULES_FILE = os.path.join(DATA_DIR, 'rules.xlsx')
PROJECTS_DIR = os.path.join(DATA_DIR, 'projects')
UPLOAD_TEMP = tempfile.mkdtemp()

# Ensure projects directory exists
os.makedirs(PROJECTS_DIR, exist_ok=True)

# Load rules engine at startup
rules_engine = RulesEngine()
if os.path.exists(RULES_FILE):
    rules_engine.load(RULES_FILE)
    print(f"Loaded {len(rules_engine.rules)} rules, {len(rules_engine.stile_widths)} stile entries from {RULES_FILE}")
else:
    print(f"WARNING: No rules file found at {RULES_FILE}")
    print(f"Upload one via the admin page at /admin")


# ─────────────────────────────────────────────
# Frontend Routes
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/admin')
def admin():
    return send_from_directory('static', 'admin.html')

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)


# ─────────────────────────────────────────────
# API: Parse Door Schedule
# ─────────────────────────────────────────────

@app.route('/api/parse-schedule', methods=['POST'])
def parse_schedule():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400

    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext not in ('pdf', 'csv', 'tsv', 'txt', 'xlsx', 'xls'):
        return jsonify({"error": f"Unsupported format: .{ext}. Use PDF, CSV, TSV, or Excel."}), 400

    filepath = os.path.join(UPLOAD_TEMP, file.filename)
    file.save(filepath)

    try:
        parser = DoorScheduleParser()
        if ext == 'pdf':
            result = parser.parse_pdf(filepath)
        elif ext in ('csv', 'tsv', 'txt'):
            result = parser.parse_csv(filepath)
        elif ext in ('xlsx', 'xls'):
            result = parser.parse_excel(filepath)

        doors_json = []
        for door in result.doors:
            d = door.to_dict()
            d['_normalized_material'] = door._normalize_material()
            doors_json.append(d)

        return jsonify({
            "success": True,
            "source": file.filename,
            "door_count": len(result.doors),
            "page_count": result.page_count,
            "column_mapping": result.column_mapping,
            "warnings": result.warnings,
            "doors": doors_json,
        })

    except Exception as e:
        return jsonify({"error": f"Parsing failed: {str(e)}"}), 500
    finally:
        try: os.remove(filepath)
        except: pass


# ─────────────────────────────────────────────
# API: Run Review
# ─────────────────────────────────────────────

@app.route('/api/run-review', methods=['POST'])
def run_review():
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400

    doors = data.get("doors", [])
    hw_sets = data.get("hardware_sets", {})

    if not doors:
        return jsonify({"error": "No doors provided"}), 400

    try:
        checker = CompatibilityChecker(rules_engine)
        issues = checker.check_all_doors(doors, hw_sets)

        issues_json = [i.to_dict() for i in issues]

        # Summary
        doors_with_issues = set(i.door_number for i in issues)
        critical = len([i for i in issues if i.severity == "critical"])
        warnings = len([i for i in issues if i.severity == "warning"])
        info = len([i for i in issues if i.severity == "info"])

        return jsonify({
            "success": True,
            "summary": {
                "total_doors": len(doors),
                "total_issues": len(issues),
                "critical": critical,
                "warnings": warnings,
                "info": info,
                "doors_with_issues": len(doors_with_issues),
                "doors_ok": len(doors) - len(doors_with_issues),
            },
            "issues": issues_json,
            "rules_loaded": len(rules_engine.rules),
            "stile_entries": len(rules_engine.stile_widths),
        })

    except Exception as e:
        return jsonify({"error": f"Review failed: {str(e)}"}), 500


# ─────────────────────────────────────────────
# API: Upload Rules Spreadsheet
# ─────────────────────────────────────────────

@app.route('/api/upload-rules', methods=['POST'])
def upload_rules():
    global rules_engine

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext not in ('xlsx', 'xls'):
        return jsonify({"error": "Rules file must be .xlsx or .xls"}), 400

    # Save to temp first, validate, then replace
    temp_path = os.path.join(UPLOAD_TEMP, f"rules_temp.{ext}")
    file.save(temp_path)

    try:
        # Test load
        test_engine = RulesEngine()
        success = test_engine.load(temp_path)

        if not success:
            return jsonify({"error": f"Invalid rules file: {'; '.join(test_engine.load_errors)}"}), 400

        if len(test_engine.rules) == 0:
            return jsonify({"error": "No rules found in spreadsheet. Check that the sheet is named 'FenestrAI Rules' or similar."}), 400

        # Valid — replace current rules
        os.makedirs(DATA_DIR, exist_ok=True)
        shutil.copy2(temp_path, RULES_FILE)

        # Reload
        rules_engine = test_engine
        rules_engine.source_file = RULES_FILE

        summary = rules_engine.summary()

        return jsonify({
            "success": True,
            "message": f"Loaded {summary['total_rules']} rules and {summary['stile_entries']} stile entries",
            "summary": summary,
        })

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500
    finally:
        try: os.remove(temp_path)
        except: pass


# ─────────────────────────────────────────────
# API: Rules Summary
# ─────────────────────────────────────────────

@app.route('/api/rules-summary')
def rules_summary():
    if not rules_engine.loaded:
        return jsonify({"loaded": False, "message": "No rules loaded. Upload a spreadsheet via /admin"})

    summary = rules_engine.summary()
    summary["loaded"] = True

    # Include rules list for display
    rules_list = []
    for r in rules_engine.rules:
        rules_list.append({
            "rule_id": r.rule_id,
            "category": r.category,
            "condition": r.condition,
            "threshold": r.threshold,
            "severity": r.severity,
            "code_reference": r.code_reference,
            "trigger_element": r.trigger_element,
            "fix_recommendation": r.fix_recommendation,
        })
    summary["rules"] = rules_list

    # Include stile data
    stiles = []
    for s in rules_engine.stile_widths:
        stiles.append({
            "vendor": s.vendor,
            "model": s.model,
            "series": s.series,
            "width": s.width_str(),
            "depth": f'{s.depth}"' if s.depth else "",
        })
    summary["stile_widths"] = stiles

    return jsonify(summary)


# ─────────────────────────────────────────────
# API: Parse Hardware Schedule
# ─────────────────────────────────────────────

@app.route('/api/parse-hardware', methods=['POST'])
def parse_hardware():
    """
    Upload a Section 08 71 00 hardware spec PDF and extract all hardware sets.
    Returns structured hardware sets that can be used for compatibility checking.
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext != 'pdf':
        return jsonify({"error": "Hardware specification must be a PDF file"}), 400

    filepath = os.path.join(UPLOAD_TEMP, file.filename)
    file.save(filepath)

    try:
        parser = HardwareScheduleParser()
        result = parser.parse_pdf(filepath)

        # Convert to the format the compatibility checker expects
        hardware_sets = {}
        for set_num, hw_set in result.hardware_sets.items():
            components = []
            for comp in hw_set.components:
                component_data = comp.to_dict()
                component_data['type'] = parser.classify_component(comp)
                components.append(component_data)

            hardware_sets[set_num] = {
                "description": hw_set.description,
                "door_type": hw_set.door_type,
                "components": components,
                "operational_description": hw_set.operational_description,
                "has_panic": hw_set.has_panic_hardware(),
                "has_closer": hw_set.has_closer(),
                "has_lockset": hw_set.has_lockset(),
            }

        return jsonify({
            "success": True,
            "source": file.filename,
            "total_sets": result.total_sets,
            "hardware_sets": hardware_sets,
            "warnings": result.warnings,
        })

    except Exception as e:
        return jsonify({"error": f"Hardware parsing failed: {str(e)}"}), 500
    finally:
        try: os.remove(filepath)
        except: pass


# ─────────────────────────────────────────────
# API: Projects (Save / Load / List / Delete)
# ─────────────────────────────────────────────

@app.route('/api/projects', methods=['GET'])
def list_projects():
    """List all saved projects."""
    projects = []
    for fname in os.listdir(PROJECTS_DIR):
        if fname.endswith('.json'):
            filepath = os.path.join(PROJECTS_DIR, fname)
            try:
                with open(filepath, 'r') as f:
                    data = json.load(f)
                projects.append({
                    "id": fname.replace('.json', ''),
                    "name": data.get("name", "Untitled"),
                    "notes": data.get("notes", ""),
                    "door_count": len(data.get("doors", [])),
                    "hw_set_count": len(data.get("hardware_sets", {})),
                    "issue_count": len(data.get("issues", [])),
                    "created_at": data.get("created_at", ""),
                    "updated_at": data.get("updated_at", ""),
                })
            except:
                pass
    projects.sort(key=lambda p: p.get('updated_at', ''), reverse=True)
    return jsonify({"projects": projects})


@app.route('/api/projects', methods=['POST'])
def save_project():
    """Save or update a project."""
    data = request.get_json()
    if not data or not data.get('name'):
        return jsonify({"error": "Project name is required"}), 400

    # Generate ID from name if new, or use existing
    project_id = data.get('id')
    if not project_id:
        import re
        slug = re.sub(r'[^a-z0-9]+', '-', data['name'].lower()).strip('-')
        project_id = f"{slug}-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    filepath = os.path.join(PROJECTS_DIR, f"{project_id}.json")

    # Load existing to preserve created_at
    existing = {}
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r') as f:
                existing = json.load(f)
        except:
            pass

    project = {
        "id": project_id,
        "name": data['name'],
        "notes": data.get('notes', ''),
        "doors": data.get('doors', []),
        "hardware_sets": data.get('hardware_sets', {}),
        "issues": data.get('issues', []),
        "review_run": data.get('review_run', False),
        "parse_info": data.get('parse_info', None),
        "hw_parse_source": data.get('hw_parse_source', None),
        "floorplan_result": data.get('floorplan_result', None),
        "created_at": existing.get('created_at', datetime.now().isoformat()),
        "updated_at": datetime.now().isoformat(),
    }

    with open(filepath, 'w') as f:
        json.dump(project, f)

    return jsonify({"success": True, "id": project_id, "message": f"Project '{data['name']}' saved"})


@app.route('/api/projects/<project_id>', methods=['GET'])
def load_project(project_id):
    """Load a saved project."""
    filepath = os.path.join(PROJECTS_DIR, f"{project_id}.json")
    if not os.path.exists(filepath):
        return jsonify({"error": "Project not found"}), 404

    with open(filepath, 'r') as f:
        data = json.load(f)

    return jsonify({"success": True, "project": data})


@app.route('/api/projects/<project_id>', methods=['DELETE'])
def delete_project(project_id):
    """Delete a saved project."""
    filepath = os.path.join(PROJECTS_DIR, f"{project_id}.json")
    if not os.path.exists(filepath):
        return jsonify({"error": "Project not found"}), 404

    os.remove(filepath)
    return jsonify({"success": True, "message": "Project deleted"})


# ─────────────────────────────────────────────
# API: Floor Plan Comparison
# ─────────────────────────────────────────────

@app.route('/api/compare-floorplan', methods=['POST'])
def compare_floorplan():
    """
    Upload floor plan PDF(s) and compare door numbers against a door schedule.

    Expects:
      - file: Floor plan PDF
      - schedule_doors: JSON array of door numbers from the schedule
    """
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    schedule_doors_json = request.form.get('schedule_doors', '[]')

    try:
        import json
        schedule_doors = json.loads(schedule_doors_json)
    except:
        schedule_doors = []

    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext != 'pdf':
        return jsonify({"error": "Floor plan must be a PDF file"}), 400

    filepath = os.path.join(UPLOAD_TEMP, file.filename)
    file.save(filepath)

    try:
        extractor = FloorPlanExtractor()
        extracted = extractor.extract_from_pdf(filepath)

        # Build extracted list
        extracted_list = []
        for d in extracted:
            extracted_list.append({
                "number": d.number,
                "page": d.page,
                "x": round(d.x, 1),
                "y": round(d.y, 1),
                "original_text": d.original_text,
            })

        # Compare if schedule doors provided
        comparison = None
        if schedule_doors:
            result = extractor.compare(extracted, schedule_doors)
            comparison = result.to_dict()

        return jsonify({
            "success": True,
            "source": file.filename,
            "doors_found": len(extracted),
            "extracted": extracted_list,
            "comparison": comparison,
        })

    except ImportError as e:
        return jsonify({"error": "PyMuPDF is required for floor plan extraction. Install with: pip install PyMuPDF"}), 500
    except Exception as e:
        return jsonify({"error": f"Floor plan extraction failed: {str(e)}"}), 500
    finally:
        try: os.remove(filepath)
        except: pass


# ─────────────────────────────────────────────
# API: Generate PDF Report
# ─────────────────────────────────────────────

@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    """Generate a printable PDF report with selected content."""
    data = request.get_json()
    if not data:
        return jsonify({"error": "No data provided"}), 400

    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    import io

    # Options
    include_schedule = data.get('include_schedule', True)
    include_issues = data.get('include_issues', True)
    include_hardware = data.get('include_hardware', False)
    include_cost = data.get('include_cost', True)
    project_name = data.get('project_name', 'Door Review Report')
    project_notes = data.get('project_notes', '')
    doors = data.get('doors', [])
    issues = data.get('issues', [])
    hardware_sets = data.get('hardware_sets', {})

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch,
                            leftMargin=0.5*inch, rightMargin=0.5*inch)
    styles = getSampleStyleSheet()

    # Custom styles
    styles.add(ParagraphStyle('ReportTitle', parent=styles['Title'], fontSize=18, spaceAfter=6))
    styles.add(ParagraphStyle('SectionHead', parent=styles['Heading2'], fontSize=13,
                              textColor=colors.HexColor('#333'), spaceBefore=16, spaceAfter=8))
    styles.add(ParagraphStyle('SmallText', parent=styles['Normal'], fontSize=8, textColor=colors.grey))
    styles.add(ParagraphStyle('IssueText', parent=styles['Normal'], fontSize=9, leading=12))

    story = []

    # Header
    story.append(Paragraph(project_name, styles['ReportTitle']))
    if project_notes:
        story.append(Paragraph(project_notes, styles['Normal']))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", styles['SmallText']))
    story.append(Spacer(1, 12))

    # Summary
    critical = [i for i in issues if i.get('severity') == 'critical']
    warnings = [i for i in issues if i.get('severity') == 'warning']
    clean = len(doors) - len(set(i.get('door_number') for i in issues))

    summary_data = [
        ['Doors Reviewed', 'Critical Issues', 'Warnings', 'Clean'],
        [str(len(doors)), str(len(critical)), str(len(warnings)), str(max(0, clean))],
    ]
    summary_table = Table(summary_data, colWidths=[1.8*inch]*4)
    summary_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, 1), 18),
        ('TEXTCOLOR', (1, 1), (1, 1), colors.red if critical else colors.HexColor('#2e7d32')),
        ('TEXTCOLOR', (2, 1), (2, 1), colors.HexColor('#e65100') if warnings else colors.HexColor('#2e7d32')),
        ('TEXTCOLOR', (3, 1), (3, 1), colors.HexColor('#2e7d32')),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f5f5f5')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#ddd')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 16))

    # Issues section
    if include_issues and issues:
        story.append(Paragraph('Compatibility Issues', styles['SectionHead']))

        for issue in sorted(issues, key=lambda i: (0 if i.get('severity') == 'critical' else 1, i.get('door_number', ''))):
            sev = issue.get('severity', 'info').upper()
            door = issue.get('door_number', '?')
            desc = issue.get('description', '')
            details = issue.get('details', '')
            cost = issue.get('cost_if_missed', '')
            solutions = issue.get('solutions', [])

            sev_color = colors.red if sev == 'CRITICAL' else colors.HexColor('#e65100')
            text = f"<b><font color='{sev_color}'>[{sev}]</font> Door {door}</b>"
            if cost and include_cost:
                text += f"  <font color='grey'>({cost})</font>"
            text += f"<br/>{desc}"
            if details:
                text += f"<br/><font size='8' color='grey'>{details}</font>"
            if solutions:
                text += f"<br/><font size='8'><b>Solutions:</b> {' | '.join(solutions)}</font>"

            story.append(Paragraph(text, styles['IssueText']))
            story.append(Spacer(1, 6))

    # Cost impact
    if include_cost and critical:
        story.append(Spacer(1, 8))
        story.append(Paragraph('Cost Impact', styles['SectionHead']))
        story.append(Paragraph("<font color='#2e7d32'><b>Caught at submittal review: $0 to fix</b></font>", styles['IssueText']))
        story.append(Paragraph("<font color='red'><b>If discovered during fabrication/installation:</b></font>", styles['IssueText']))
        for issue in critical:
            if issue.get('cost_if_missed'):
                story.append(Paragraph(f"  Door {issue['door_number']}: {issue['cost_if_missed']}", styles['IssueText']))

    # Door schedule table
    if include_schedule and doors:
        story.append(PageBreak())
        story.append(Paragraph('Door Schedule', styles['SectionHead']))

        # Build table with key columns
        headers = ['Door #', 'Size', 'Material', 'HW Set']
        has_type = any(d.get('door_type') for d in doors)
        has_fire = any(d.get('fire_rating') for d in doors)
        has_finish = any(d.get('finish') for d in doors)
        has_frame = any(d.get('frame_material') for d in doors)
        if has_type: headers.insert(2, 'Type')
        if has_finish: headers.insert(-1, 'Finish')
        if has_frame: headers.insert(-1, 'Frame')
        if has_fire: headers.insert(-1, 'Fire')

        table_data = [headers]
        for d in doors:
            size = f"{d.get('width', '')} x {d.get('height', '')}".strip(' x')
            row = [d.get('door_number', ''), size, (d.get('material', '') or '').upper(), d.get('hardware_set', '')]
            if has_type: row.insert(2, d.get('door_type', ''))
            if has_finish: row.insert(-1, d.get('finish', ''))
            if has_frame: row.insert(-1, d.get('frame_material', ''))
            if has_fire: row.insert(-1, d.get('fire_rating', ''))
            table_data.append(row)

        col_count = len(headers)
        col_w = min(1.4, 7.0 / col_count) * inch
        sched_table = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
        sched_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#ccc')),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f9f9f9')]),
        ]))

        # Highlight rows with issues
        door_issues = {}
        for issue in issues:
            dn = issue.get('door_number', '')
            if dn not in door_issues:
                door_issues[dn] = issue.get('severity', 'info')
            elif issue.get('severity') == 'critical':
                door_issues[dn] = 'critical'

        for idx, d in enumerate(doors):
            dn = d.get('door_number', '')
            if dn in door_issues:
                row_idx = idx + 1
                bg = colors.HexColor('#fce4e4') if door_issues[dn] == 'critical' else colors.HexColor('#fff3e0')
                sched_table.setStyle(TableStyle([('BACKGROUND', (0, row_idx), (-1, row_idx), bg)]))

        story.append(sched_table)

    # Hardware sets
    if include_hardware and hardware_sets:
        story.append(PageBreak())
        story.append(Paragraph('Hardware Sets', styles['SectionHead']))

        for num in sorted(hardware_sets.keys(), key=lambda x: int(x) if x.isdigit() else 999):
            hw = hardware_sets[num]
            desc = hw.get('description', '')
            story.append(Paragraph(f"<b>Set #{num}</b> - {desc}", styles['IssueText']))
            comps = hw.get('components', [])
            for c in comps:
                cdesc = c.get('description', '')
                cat = c.get('catalog_number', '')
                mfr = c.get('manufacturer', '')
                qty = c.get('qty', 1)
                line = f"  {qty}x {cdesc}"
                if cat: line += f" {cat}"
                if mfr: line += f" ({mfr})"
                story.append(Paragraph(line, styles['SmallText']))
            story.append(Spacer(1, 8))

    # Build
    doc.build(story)
    buf.seek(0)

    from flask import send_file
    return send_file(buf, mimetype='application/pdf',
                     download_name=f"{project_name.replace(' ', '_')}_Report.pdf",
                     as_attachment=True)


# ─────────────────────────────────────────────
# API: Health Check
# ─────────────────────────────────────────────

@app.route('/api/health')
def health():
    return jsonify({
        "status": "ok",
        "version": "2.0",
        "rules_loaded": rules_engine.loaded,
        "rule_count": len(rules_engine.rules),
        "stile_count": len(rules_engine.stile_widths),
        "timestamp": datetime.now().isoformat(),
    })


# ─────────────────────────────────────────────
# Run
# ─────────────────────────────────────────────

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'

    print("=" * 50)
    print("Door Review Tool v2.0")
    print("=" * 50)
    print(f"App:   http://localhost:{port}")
    print(f"Admin: http://localhost:{port}/admin")
    print(f"Rules: {len(rules_engine.rules)} loaded")
    print("=" * 50)

    app.run(debug=debug, host='0.0.0.0', port=port)
