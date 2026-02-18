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

# ─────────────────────────────────────────────
# App Setup
# ─────────────────────────────────────────────

app = Flask(__name__, static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# Persistent data directory (survives restarts on Render with disk)
DATA_DIR = os.environ.get('DATA_DIR', os.path.join(os.path.dirname(__file__), 'data'))
RULES_FILE = os.path.join(DATA_DIR, 'rules.xlsx')
UPLOAD_TEMP = tempfile.mkdtemp()

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
