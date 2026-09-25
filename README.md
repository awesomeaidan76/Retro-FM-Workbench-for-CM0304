# Retro FM Database Workbench v0.4

A CM03/04-first conversion workbench intended to become:

CM03/04 database -> native extractor -> normalized SQLite -> matching -> validation -> review queue -> FM24/FM26/FMME export

## What this build does now

- Windows-friendly Tkinter GUI.
- Select CM03/04 as the source and FM24 or FM26 as the target.
- Inspect a CM03/04 database folder and identify likely `.dat` files.
- Create a normalized SQLite staging database.
- Import CSV, JSON and XLSX exports.
- Normalize common football-data field names.
- Run structural validation.
- Detect duplicate/ambiguous player records.
- Produce a review queue.
- Optionally ask a local Ollama model for a second opinion on two candidate records.
- Export a structured conversion bundle with `players.csv`, `staff.csv`, `clubs.csv`, `nations.csv`, `competitions.csv`, `issues.json`, `review_queue.json`, and `manifest.json`.

## Current CM03/04 parser status

The application DOES NOT claim to decode the proprietary `server_db.dat` / `people_db.dat` records yet.

Instead, it safely inventories the database and records its structure in the staging DB. The binary parser is isolated so it can be implemented and tested without changing the converter.

This is intentional: a bad binary parser can silently create tens of thousands of incorrect players. The workbench should never pretend that a guessed byte layout is correct.

## Running

Windows:
1. Install Python 3.11+ with Tkinter.
2. Double-click `run_retrofm.bat`.

Command line:
    python retrofm/app.py

Optional XLSX support:
    pip install openpyxl

Optional local AI review:
- Install/run Ollama separately.
- Pull a model, e.g. `ollama pull qwen2.5:7b`.
- Use the AI Review tab.

## Suggested first real test

1. Open CM03/04.
2. Point Source at its `data\db` folder.
3. Click Inspect Source.
4. Save/export a player list from an existing CM03/04 editor or other trusted tool.
5. Import that CSV/XLSX through the Import tab.
6. Run Validation and Duplicate Matching.
7. Export a conversion bundle.

## Next parser milestone

The native CM03/04 extractor should be built against the archived Nygreen CM4Pregame source structures and verified against an actual 03/04 database. The parser should extract, at minimum:

- people / players
- staff
- names
- nations
- clubs
- competitions
- contracts / registrations
- player history
- future transfers
- injuries / suspensions where present
- player attributes and positions

Then the existing normalization -> matching -> validation -> export pipeline can operate directly on the original game database.
