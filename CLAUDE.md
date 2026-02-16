# CLAUDE.md

## Project Overview

Manufacturing order management web application (手配発注マージシステム) for tracking parts procurement, receiving, and delivery schedules. Built with Flask + SQLite (dev) / SQL Server (prod). The UI is a single-page application served from `templates/index.html` with vanilla JavaScript.

## Tech Stack

- **Backend**: Python 3, Flask 2.3.3, Flask-SQLAlchemy 3.0.5
- **Frontend**: HTML5, CSS3, vanilla JavaScript (no framework)
- **Database**: SQLite (development), SQL Server via ODBC (production)
- **Data processing**: pandas 2.0.3, openpyxl 3.1.2
- **Other**: qrcode, Pillow, pyodbc, win32com (Windows COM for Excel), Flask-CORS

## Project Structure

```
app.py              # Main Flask app: models, routes, business logic (~6050 lines)
config.py           # Environment config (dev/prod/test) with SSL and ODBC settings
models.py           # Lazy import wrapper for models (avoids circular imports)
across_db.py        # SQL Server (Across DB) read-only connector for enterprise data
requirements.txt    # Python dependencies (pip)

utils/
  constants.py      # Application constants
  data_utils.py     # Data processing utilities (normalization, type conversion)
  mekki_utils.py    # Plating (メッキ) detection logic
  excel_styler.py   # Excel formatting helpers
  qr_generator.py   # QR code generation
  excel_gantt_chart.py  # Gantt chart Excel sheet generation
  email_sender.py   # Email notification system
  delivery_utils.py # Delivery data (deprecated stub)

services/
  cad_service.py    # CAD file path resolution from part numbers
  excel_export.py   # Excel export functionality
  cache_service.py  # Caching service

templates/
  index.html        # SPA main page (~3770 lines, all tabs in one file)

static/
  styles.css        # Global styles
  gantt-chart.js    # Gantt chart rendering (Chart.js)
  qr-scanner.js     # QR/barcode scanning (Html5Qrcode library)
  pallet-manager.js # Pallet management UI
  delivery-schedule.js  # Delivery schedule UI
  across-db.js      # Across DB frontend interface

instance/
  order_management.db  # SQLite database file
  migrate_db.py        # Manual database migration script

docs/               # Documentation
uploads/            # File upload storage
exports/            # Excel export output
cache/              # Cache storage
labels/             # Label output
```

## Key Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run development server
python app.py

# Run with specific environment
FLASK_ENV=development python app.py
FLASK_ENV=production python app.py

# Generate SSL certificate (development)
python generate_cert.py

# Check database integrity
python db_check.py

# Run database migrations
python instance/migrate_db.py
```

## Database Models

All models are defined in `app.py` (lines ~92-400). Key models:

- **Order** - Manufacturing order by seiban (serial number) + unit. Statuses: `受入準備前`, `納品中`, `納品完了`
- **OrderDetail** - Line items (parts) with supplier, delivery date, quantities. Supports parent-child hierarchy via `parent_id`
- **ReceivedHistory** - Receiving audit trail (tracks received quantities, timestamps, IP addresses)
- **EditLog** - Activity audit log (receive/unreceive actions)
- **PartCategory** - Part code master (3-letter prefix codes like NAA, NKA)
- **UserSettings** - Per-client settings stored by IP address

## API Design

RESTful JSON API under `/api/*` prefix. HTTP methods: GET, POST, DELETE. All responses are JSON. Key endpoint groups:

- `/api/orders` - Order CRUD
- `/api/orders/<id>/details` - Order detail management
- `/api/process` - Excel data import and merge
- `/api/receive`, `/api/unreceive` - Receiving workflow
- `/api/export-*` - Excel export endpoints
- `/api/across-db/*` - Across DB queries (SQL Server)
- `/api/cad/*` - CAD file lookup
- `/api/gantt-chart/*` - Gantt chart data
- `/api/pallet/*` - Pallet management

## Architecture Notes

- **Monolithic app.py**: Models, routes, and business logic are all in one file. When modifying, be aware of the file size (~6050 lines).
- **SPA pattern**: `index.html` contains all UI tabs. JavaScript files in `static/` handle tab-specific functionality via AJAX calls.
- **Data flow**: Excel/Across DB -> `/api/process` -> OrderDetail records -> Order aggregation -> Excel Export -> QR Scanning for receiving -> ReceivedHistory audit trail.
- **Japanese domain**: Variable names, UI text, database values, and comments are primarily in Japanese. Maintain this convention.
- **Order number normalization**: Order numbers are zero-padded and normalized. Use `DataUtils` for safe type conversions.
- **Plating detection**: `MekkiUtils` checks supplier codes and spec2 patterns to flag plating-related items.
- **Hierarchical parts**: OrderDetail supports parent-child relationships for multi-level BOMs.
- **Windows integration**: Production uses UNC paths (`\\server3\...`) for shared Excel files and `win32com` for COM automation. These will fail outside Windows.

## Configuration

Configured via `config.py` using `FLASK_ENV` environment variable:

| Setting | Development | Production | Test |
|---------|-------------|------------|------|
| DEBUG | True | False | True |
| Database | SQLite | SQLite (+ SQL Server ODBC) | SQLite (test db) |
| HTTPS | Self-signed cert | Cert from env vars | Disabled |
| ODBC | Disabled | Enabled | Disabled |

Key environment variables:
- `FLASK_ENV` - Environment selector (`development`/`production`/`testing`)
- `SECRET_KEY` - Flask secret key (required in production)
- `SSL_CERT_PATH`, `SSL_KEY_PATH` - SSL certificate paths

## Conventions

- **Language**: Comments, variable names, and UI strings are in Japanese. Keep this consistent.
- **Python style**: snake_case for functions and variables. No formal linter configured.
- **JavaScript style**: camelCase for functions and variables. No formal linter configured.
- **API endpoints**: kebab-case (e.g., `/api/across-db/test`)
- **Database tables**: snake_case (via SQLAlchemy model names)
- **HTML**: IDs and classes use kebab-case
- **Commit messages**: English, imperative mood (e.g., "Add box QR scanning feature")
- **No automated tests**: Testing is manual via the web UI. No pytest/unittest framework exists.
- **No CI/CD pipeline**: Deployment is manual.

## Common Pitfalls

- `app.py` is large (~6050 lines). Search carefully before making changes to avoid duplicate definitions.
- `win32com` and `pyodbc` imports will fail on non-Windows systems. The app handles `ImportError` gracefully in some places but not all.
- Database migrations are manual (`instance/migrate_db.py`). When adding columns, update the migration script.
- The `SQLALCHEMY_ENGINE_OPTIONS` pool settings are tuned to prevent connection exhaustion. Do not reduce `pool_size` without understanding the concurrent load.
- Excel export logic is complex with specific formatting requirements. Test output files visually when modifying.
- Order numbers may contain leading zeros. Always use normalization utilities from `DataUtils`.
