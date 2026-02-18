# Door Review Tool

Upload door schedules (PDF, CSV, Excel) → get instant hardware compatibility reports.

Rules are driven by an Excel spreadsheet that you and your hardware specialist maintain — no code changes needed to add new rules.

## Deploy to Render

1. Push this repo to GitHub
2. Go to [render.com](https://render.com) → New → Web Service
3. Connect your GitHub repo
4. Render auto-detects the `render.yaml` config
5. Add a custom domain: `hardwarereviewer.com`

**Important:** The `render.yaml` includes a persistent disk for storing the rules spreadsheet. This ensures your uploaded rules survive deployments.

## Local Development

```bash
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000

## Updating Rules

1. Go to `/admin` (or click "Admin" in the app)
2. Drag and drop your updated `.xlsx` spreadsheet
3. Rules are validated and loaded immediately
4. No restart or redeployment needed

## Files

```
├── app.py                    # Flask web server
├── door_schedule_parser.py   # PDF/CSV/Excel table extraction
├── rules_engine.py           # Reads rules from Excel spreadsheet
├── compatibility_checker.py  # Runs checks using rules engine
├── requirements.txt          # Python dependencies
├── render.yaml               # Render deployment config
├── data/
│   └── rules.xlsx            # Current rules spreadsheet (default)
└── static/
    ├── index.html            # Main review interface
    └── admin.html            # Admin panel for rule uploads
```

## Spreadsheet Format

The rules spreadsheet has two sheets:

**FenestrAI Rules** — Code compliance and hardware rules
- Rule ID, Category, Condition, Threshold, Severity
- Code Reference, Trigger Element, Applies To
- Fix Recommendation, Notes

**Aluminum Door Stile Widths** — Manufacturer stile dimensions
- Vendor, Model, Series, Width, Depth
