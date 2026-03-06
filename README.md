# Model Tiering Web App

A full-stack Flask application for model risk tiering and governance.

## Quick Start

```bash
pip install -r requirements.txt
python manage.py init-db
python manage.py run
```

Open: http://127.0.0.1:5000

## Default Users

| Email              | Password | Role      |
|--------------------|----------|-----------|
| admin@system.com   | admin123 | Admin     |
| dev@system.com     | dev123   | Developer |
| user@system.com    | user123  | User      |

## CLI Commands

```bash
python manage.py init-db        # Create tables + seed defaults
python manage.py create-user    # Interactively create a user
python manage.py run            # Start dev server
python manage.py db-path        # Print database path
```

## Features

- **Dashboard** – KPI cards, Chart.js doughnut & bar charts
- **Models** – Create, view, manage risk models
- **Parameters** – Define Materiality / Criticality / Complexity parameters with weights
- **Tiering** – Automated score calculation and tier assignment
- **Override** – Manual tier override with audit trail
- **Reports** – Filterable tiering reports with CSV export and saved filters
- **Upload** – Bulk import via Excel (.xlsx) with auto-tiering
- **Settings** – Admin configuration of group weights and tier ranges
- **Auth** – Role-based access (Admin / Developer / User)

## Tiering Formula

```
Group Score = Σ(parameter.weight × level)   [level: 1=Low, 2=Medium, 3=High]

Final Score = (Materiality_Score × mat_weight)
            + (Criticality_Score × crit_weight)
            + (Complexity_Score  × comp_weight)
```

## Excel Upload Format

Sheets: `Models`, `Parameters`, `Tiers`, `Settings`

## Tech Stack

- Python 3.9+ / Flask 3 (app factory pattern)
- SQLAlchemy ORM / SQLite
- Flask-Login
- Jinja2 templates
- Chart.js (CDN)
- openpyxl
