# Courses Manager

A lightweight Flask application that imports course registrations from an Excel workbook, allows staff to track which applicants have been contacted, and keeps the participant database in sync.

## Features

- **Single account sign-in.** A default administrative account is generated on first launch (username `admin`, password `changeme`). Override the credentials with the `COURSES_MANAGER_DEFAULT_USERNAME` and `COURSES_MANAGER_DEFAULT_PASSWORD` environment variables.
- **Automatic Excel import.** On every successful login the application reads `registrations.xlsx` (located next to `app.py`) and upserts its contents into the SQLite database.
- **Duplicate protection.** Participants are uniquely identified by the combination of their ID number and the course name. Attempts to import duplicates update the existing record instead of creating a new one.
- **Status tracking.** Mark a participant as `CALLED` directly from the dashboard. Called participants disappear from the list on subsequent reloads.
- **Chronological ordering.** Participants waiting to be contacted are sorted from the oldest to the most recent registration date.

## Requirements

- Python 3.10 or newer
- The dependencies listed in `requirements.txt`
- An Excel workbook named `registrations.xlsx` in the project root. Each worksheet must contain the columns listed below (case-insensitive, small spelling variations are accepted):
  - `Name`
  - `Surname`
  - `Date` (or `Registration Date`)
  - `ID` (participant identification number)
  - `HRDA no.` (optional)
  - `Name of the course`
  - `Κατηγορία απασχόλησης` (employment category)
  - `Email`
  - `Phone number`

Columns not listed above are ignored. Empty rows are skipped automatically.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the application

```bash
export FLASK_APP=app.py
# Optional overrides
# export COURSES_MANAGER_SECRET_KEY="super-secret-key"
# export COURSES_MANAGER_DEFAULT_USERNAME="manager"
# export COURSES_MANAGER_DEFAULT_PASSWORD="update-me"
flask --app app run
```

The site is now available at [http://localhost:5000](http://localhost:5000). Sign in with the configured credentials, review the participants, and mark them as called as you make contact.

Each login imports the latest data from `registrations.xlsx`. To refresh the list with new applicants, update the workbook and sign in again.

## Project structure

```
app.py              # Flask application and database logic
app.db              # SQLite database (created on first run)
registrations.xlsx  # Excel data source (must be provided)
templates/          # HTML templates
static/             # Stylesheet assets
```

## Notes

- The application keeps existing status values when synchronising with the Excel sheet. Records marked as `CALLED` remain hidden on subsequent imports.
- Remove or archive `app.db` if you need to start from a clean database snapshot.
- Inspect the server logs for warning messages when an Excel sheet is missing required headers.
