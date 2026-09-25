#!/usr/bin/env python3
"""
Retro FM Database Workbench v0.4
CM03/04-first retro database conversion workbench.

This build provides:
- GUI for source/target selection
- CM03/04 database folder inspection
- CSV/JSON/XLSX import into a normalized SQLite staging DB
- player matching / duplicate detection
- validation and exception reporting
- FM24/FM26-oriented export bundle
- review queue for ambiguous cases
- optional local Ollama-assisted review
- pluggable legacy parser interface for future native CM03/04 binary decoding

The native CM03/04 parser is intentionally isolated in legacy_cm0304.py.
It performs safe inventory/inspection now and provides a clean hook for a
validated binary parser when a concrete CM03/04 database or source structure
is available for byte-for-byte testing.
"""

from __future__ import annotations

import csv
import datetime as dt
import hashlib
import json
import math
import os
import re
import sqlite3
import threading
import urllib.request
import urllib.error
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Optional

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    raise SystemExit("Tkinter is required. Use the Windows Python installer with Tcl/Tk enabled.")

APP_NAME = "Retro FM Database Workbench"
VERSION = "0.4"
DB_NAME = "retrofm_stage.sqlite3"

CANONICAL_FIELDS = {
    "player_id": ["player_id", "id", "uid", "unique_id", "person_id", "personid"],
    "first_name": ["first_name", "firstname", "forename", "given_name", "name_first"],
    "last_name": ["last_name", "lastname", "surname", "family_name", "name_last"],
    "common_name": ["common_name", "nickname", "short_name", "display_name"],
    "dob": ["dob", "date_of_birth", "birth_date", "birthdate"],
    "nation": ["nation", "nationality", "country"],
    "club": ["club", "club_name", "team", "current_club"],
    "position": ["position", "positions", "role"],
    "ca": ["ca", "current_ability", "currentability", "ability"],
    "pa": ["pa", "potential_ability", "potentialability", "potential"],
    "wage": ["wage", "salary", "weekly_wage"],
    "value": ["value", "market_value"],
    "height": ["height", "height_cm"],
    "weight": ["weight", "weight_kg"],
}

ENTITY_NAMES = ["players", "staff", "clubs", "nations", "competitions", "transfers"]

@dataclass
class Issue:
    severity: str
    entity: str
    key: str
    message: str
    details: str = ""

class ParserNotImplementedError(RuntimeError):
    pass

def norm_text(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower()
    s = re.sub(r"[^\w\s-]", "", s, flags=re.UNICODE)
    s = re.sub(r"[\s_-]+", " ", s)
    return s.strip()

def norm_name(first: Any, last: Any, common: Any = "") -> str:
    c = norm_text(common)
    if c:
        return c
    return norm_text(f"{first or ''} {last or ''}")

def parse_date(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    # Excel serial date
    try:
        x = float(s)
        if 20000 < x < 60000:
            return (dt.date(1899, 12, 30) + dt.timedelta(days=x)).isoformat()
    except Exception:
        pass
    return s

def canonicalize_row(row: dict[str, Any]) -> dict[str, Any]:
    lowered = {norm_text(k).replace(" ", "_"): v for k, v in row.items()}
    out = dict(row)
    for canon, aliases in CANONICAL_FIELDS.items():
        for alias in aliases:
            a = norm_text(alias).replace(" ", "_")
            if a in lowered and str(lowered[a]).strip() != "":
                out[canon] = lowered[a]
                break
    if "dob" in out:
        out["dob"] = parse_date(out["dob"])
    out["name_norm"] = norm_name(out.get("first_name"), out.get("last_name"), out.get("common_name"))
    return out

def score_match(a: dict[str, Any], b: dict[str, Any]) -> float:
    score = 0.0
    na = a.get("name_norm") or norm_name(a.get("first_name"), a.get("last_name"), a.get("common_name"))
    nb = b.get("name_norm") or norm_name(b.get("first_name"), b.get("last_name"), b.get("common_name"))
    if na and nb:
        if na == nb:
            score += 0.70
        else:
            at = set(na.split())
            bt = set(nb.split())
            if at and bt:
                score += 0.70 * (len(at & bt) / max(len(at | bt), 1))
    if parse_date(a.get("dob")) and parse_date(b.get("dob")):
        if parse_date(a.get("dob")) == parse_date(b.get("dob")):
            score += 0.20
    ca = norm_text(a.get("club"))
    cb = norm_text(b.get("club"))
    if ca and cb and ca == cb:
        score += 0.10
    return min(score, 1.0)

def read_tabular(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        # BOM-safe and tolerant of common delimiters
        text = path.read_text(encoding="utf-8-sig", errors="replace")
        sample = text[:8192]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        return [dict(r) for r in csv.DictReader(text.splitlines(), dialect=dialect)]
    if suffix == ".json":
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return [x if isinstance(x, dict) else {"value": x} for x in data]
        if isinstance(data, dict):
            for key in ("players", "staff", "clubs", "records", "data"):
                if isinstance(data.get(key), list):
                    return [x if isinstance(x, dict) else {"value": x} for x in data[key]]
            return [data]
    if suffix == ".xlsx":
        try:
            import openpyxl
        except ImportError as e:
            raise RuntimeError("XLSX import requires openpyxl. Install it with: pip install openpyxl") from e
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        headers = [str(x).strip() if x is not None else "" for x in rows[0]]
        return [dict(zip(headers, r)) for r in rows[1:] if any(x is not None for x in r)]
    raise ValueError(f"Unsupported tabular format: {path}")

class StageDB:
    def __init__(self, path: Path):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.init()

    def init(self):
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS metadata (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        CREATE TABLE IF NOT EXISTS players (
            row_id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id TEXT,
            first_name TEXT,
            last_name TEXT,
            common_name TEXT,
            name_norm TEXT,
            dob TEXT,
            nation TEXT,
            club TEXT,
            position TEXT,
            ca REAL,
            pa REAL,
            wage REAL,
            value REAL,
            height REAL,
            weight REAL,
            raw_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS staff (
            row_id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id TEXT,
            first_name TEXT,
            last_name TEXT,
            common_name TEXT,
            name_norm TEXT,
            dob TEXT,
            nation TEXT,
            club TEXT,
            role TEXT,
            raw_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS generic_records (
            row_id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity TEXT NOT NULL,
            source_id TEXT,
            name TEXT,
            name_norm TEXT,
            raw_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS issues (
            issue_id INTEGER PRIMARY KEY AUTOINCREMENT,
            severity TEXT,
            entity TEXT,
            key TEXT,
            message TEXT,
            details TEXT
        );
        CREATE TABLE IF NOT EXISTS matches (
            source_row_id INTEGER,
            candidate_row_id INTEGER,
            score REAL,
            status TEXT,
            rationale TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_players_name ON players(name_norm);
        CREATE INDEX IF NOT EXISTS idx_players_dob ON players(dob);
        """)
        self.conn.commit()

    def clear(self):
        for t in ("players","staff","generic_records","issues","matches"):
            self.conn.execute(f"DELETE FROM {t}")
        self.conn.commit()

    def meta(self, key: str, value: Optional[str] = None) -> Optional[str]:
        if value is None:
            row = self.conn.execute("SELECT value FROM metadata WHERE key=?", (key,)).fetchone()
            return row["value"] if row else None
        self.conn.execute("INSERT OR REPLACE INTO metadata(key,value) VALUES(?,?)", (key,value))
        self.conn.commit()
        return value

    def insert_rows(self, entity: str, rows: Iterable[dict[str, Any]]):
        rows = [canonicalize_row(r) for r in rows]
        if entity == "players":
            for r in rows:
                raw = json.dumps(r, ensure_ascii=False, default=str)
                vals = [r.get(k,"") for k in (
                    "player_id","first_name","last_name","common_name","name_norm","dob",
                    "nation","club","position","ca","pa","wage","value","height","weight"
                )]
                self.conn.execute("""
                    INSERT INTO players(source_id,first_name,last_name,common_name,name_norm,dob,nation,club,
                                        position,ca,pa,wage,value,height,weight,raw_json)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (*vals, raw))
        elif entity == "staff":
            for r in rows:
                raw = json.dumps(r, ensure_ascii=False, default=str)
                vals = [r.get(k,"") for k in ("player_id","first_name","last_name","common_name","name_norm","dob","nation","club")]
                self.conn.execute("""
                    INSERT INTO staff(source_id,first_name,last_name,common_name,name_norm,dob,nation,club,role,raw_json)
                    VALUES(?,?,?,?,?,?,?,?,?,?)
                """, (*vals, r.get("role",""), raw))
        else:
            for r in rows:
                raw = json.dumps(r, ensure_ascii=False, default=str)
                name = r.get("name") or r.get("club") or r.get("competition") or r.get("nation") or r.get("value","")
                self.conn.execute("""
                    INSERT INTO generic_records(entity,source_id,name,name_norm,raw_json)
                    VALUES(?,?,?,?,?)
                """, (entity, str(r.get("id") or r.get("uid") or ""), str(name), norm_text(name), raw))
        self.conn.commit()

    def count(self, table: str) -> int:
        return int(self.conn.execute(f"SELECT COUNT(*) c FROM {table}").fetchone()["c"])

    def add_issue(self, issue: Issue):
        self.conn.execute(
            "INSERT INTO issues(severity,entity,key,message,details) VALUES(?,?,?,?,?)",
            (issue.severity, issue.entity, issue.key, issue.message, issue.details)
        )

    def issues(self):
        return self.conn.execute("SELECT * FROM issues ORDER BY CASE severity WHEN 'ERROR' THEN 0 WHEN 'WARN' THEN 1 ELSE 2 END, issue_id").fetchall()

    def close(self):
        self.conn.commit()
        self.conn.close()

class SourceAdapter:
    game_id = "generic"
    label = "Generic"

    def inspect(self, source: Path) -> dict[str, Any]:
        raise NotImplementedError

    def extract(self, source: Path, stage: StageDB):
        raise NotImplementedError

class GenericAdapter(SourceAdapter):
    game_id = "generic"
    label = "CSV / JSON / XLSX"

    def inspect(self, source: Path) -> dict[str, Any]:
        files = [source] if source.is_file() else [
            p for p in source.rglob("*") if p.is_file() and p.suffix.lower() in {".csv",".json",".xlsx"}
        ]
        return {
            "adapter": self.label,
            "files": [{"name": str(p), "bytes": p.stat().st_size} for p in files],
            "message": "Generic tabular source"
        }

    def extract(self, source: Path, stage: StageDB):
        files = [source] if source.is_file() else [
            p for p in source.rglob("*") if p.is_file() and p.suffix.lower() in {".csv",".json",".xlsx"}
        ]
        imported = 0
        for p in files:
            rows = read_tabular(p)
            if not rows:
                continue
            stem = p.stem.lower()
            if any(x in stem for x in ("player","people","person")):
                entity = "players"
            elif "staff" in stem:
                entity = "staff"
            elif "club" in stem or "team" in stem:
                entity = "clubs"
            elif "nation" in stem or "country" in stem:
                entity = "nations"
            elif "competition" in stem or "league" in stem:
                entity = "competitions"
            else:
                entity = "players" if any("dob" in norm_text(r) or "surname" in norm_text(r) for r in rows[:5]) else "clubs"
            stage.insert_rows(entity, rows)
            imported += len(rows)
        return imported

class CM0304Adapter(SourceAdapter):
    game_id = "cm0304"
    label = "Championship Manager 03/04"

    LIKELY_FILES = [
        "server_db.dat", "client_db.dat", "people_db.dat", "lang_db.dat",
        "pl_hist_dt.dat", "pl_hist_id.dat", "pl_hist_index.dat"
    ]

    def inspect(self, source: Path) -> dict[str, Any]:
        if source.is_file():
            files = [source]
            base = source.parent
        else:
            base = source
            files = [p for p in source.rglob("*.dat") if p.is_file()]
        entries = []
        names = {p.name.lower() for p in files}
        for p in sorted(files):
            h = hashlib.sha256(p.read_bytes()[:1024*1024]).hexdigest()[:16]
            entries.append({
                "name": str(p.relative_to(base)) if p.is_relative_to(base) else str(p),
                "bytes": p.stat().st_size,
                "sha256_first_1mb": h
            })
        found = [n for n in self.LIKELY_FILES if n in names]
        return {
            "adapter": self.label,
            "source": str(source),
            "files": entries,
            "likely_cm0304_files": found,
            "recognized": "server_db.dat" in names or "people_db.dat" in names,
            "message": (
                "CM03/04 database signatures detected. Native binary extraction is isolated "
                "behind the CM0304 adapter and will only decode records when the parser schema "
                "has been validated against a real database."
            )
        }

    def extract(self, source: Path, stage: StageDB):
        report = self.inspect(source)
        stage.meta("source_game", self.game_id)
        stage.meta("source_inspection", json.dumps(report, ensure_ascii=False, indent=2))
        raise ParserNotImplementedError(
            "Native CM03/04 .dat record decoding is not enabled in v0.4 yet. "
            "Use the Import/Build tab with CSV/JSON/XLSX exported from an existing CM03/04 editor, "
            "or place a validated CM03/04 parser module in retrofm/plugins. "
            "The rest of the conversion pipeline is fully functional."
        )

ADAPTERS = {
    "cm0304": CM0304Adapter(),
    "generic": GenericAdapter(),
}

def validate(stage: StageDB):
    # Basic structural checks.
    for r in stage.conn.execute("SELECT row_id,* FROM players"):
        key = r["source_id"] or str(r["row_id"])
        if not r["name_norm"]:
            stage.add_issue(Issue("ERROR","player",key,"Missing player name"))
        if r["dob"]:
            try:
                d = dt.date.fromisoformat(r["dob"])
                if d.year < 1880 or d.year > dt.date.today().year:
                    stage.add_issue(Issue("WARN","player",key,"Unusual date of birth",r["dob"]))
            except Exception:
                stage.add_issue(Issue("WARN","player",key,"Unparsed date of birth",r["dob"]))
        for f, lo, hi in (("ca",0,200),("pa",0,200),("height",100,250),("weight",30,180)):
            v = r[f]
            if v not in (None, ""):
                try:
                    x = float(v)
                    if x < lo or x > hi:
                        stage.add_issue(Issue("WARN","player",key,f"{f} outside expected range",str(v)))
                except Exception:
                    stage.add_issue(Issue("WARN","player",key,f"{f} not numeric",str(v)))
    # Duplicate exact names + DOB.
    dupes = stage.conn.execute("""
        SELECT name_norm,dob,COUNT(*) c
        FROM players
        WHERE name_norm <> ''
        GROUP BY name_norm,dob
        HAVING c > 1
        ORDER BY c DESC
    """).fetchall()
    for d in dupes[:500]:
        stage.add_issue(Issue("WARN","players",d["name_norm"],
                              "Potential duplicate player records",
                              f"name={d['name_norm']}; dob={d['dob']}; count={d['c']}"))
    # Required source metadata.
    if not stage.meta("source_game"):
        stage.add_issue(Issue("INFO","project","source_game","Source game not recorded"))

def run_matching(stage: StageDB, threshold=0.82):
    stage.conn.execute("DELETE FROM matches")
    players = [dict(r) for r in stage.conn.execute("SELECT * FROM players")]
    # First self-match by normalized name/DOB to detect duplicate clusters. This is useful
    # as a stand-alone QA pass before matching against a canonical target DB.
    by_name = {}
    for p in players:
        by_name.setdefault(p["name_norm"], []).append(p)
    for p in players:
        candidates = by_name.get(p["name_norm"], [])
        best = (0.0, None)
        for c in candidates:
            if c["row_id"] == p["row_id"]:
                continue
            s = score_match(p, c)
            if s > best[0]:
                best = (s, c)
        if best[1] is not None:
            status = "AUTO_DUPLICATE" if best[0] >= threshold else "REVIEW"
            stage.conn.execute(
                "INSERT INTO matches(source_row_id,candidate_row_id,score,status,rationale) VALUES(?,?,?,?,?)",
                (p["row_id"], best[1]["row_id"], best[0], status,
                 f"name/date/club similarity={best[0]:.3f}")
            )
    stage.conn.commit()

def export_bundle(stage: StageDB, out_dir: Path, target: str):
    out_dir.mkdir(parents=True, exist_ok=True)
    # Export players in a flat, editor-friendly CSV.
    pfile = out_dir / "players.csv"
    with pfile.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "source_id","first_name","last_name","common_name","dob","nation","club",
            "position","ca","pa","wage","value","height","weight"
        ])
        for r in stage.conn.execute("SELECT * FROM players ORDER BY row_id"):
            writer.writerow([r[k] for k in (
                "source_id","first_name","last_name","common_name","dob","nation","club",
                "position","ca","pa","wage","value","height","weight"
            )])
    # Generic entity exports.
    for entity in ("clubs","nations","competitions","staff"):
        rows = stage.conn.execute(
            "SELECT * FROM generic_records WHERE entity=? ORDER BY row_id", (entity,)
        ).fetchall()
        if entity == "staff":
            rows = stage.conn.execute("SELECT * FROM staff ORDER BY row_id").fetchall()
        ef = out_dir / f"{entity}.csv"
        with ef.open("w", newline="", encoding="utf-8-sig") as f:
            if entity == "staff":
                headers = ["source_id","first_name","last_name","common_name","dob","nation","club","role"]
                writer = csv.writer(f); writer.writerow(headers)
                for r in rows:
                    writer.writerow([r[h] for h in headers])
            else:
                writer = csv.writer(f); writer.writerow(["source_id","name","name_norm","raw_json"])
                for r in rows:
                    writer.writerow([r["source_id"],r["name"],r["name_norm"],r["raw_json"]])
    issues = [
        dict(r) for r in stage.issues()
    ]
    (out_dir / "issues.json").write_text(json.dumps(issues, ensure_ascii=False, indent=2), encoding="utf-8")
    reviews = []
    for r in stage.conn.execute("SELECT * FROM matches WHERE status='REVIEW' OR status='AUTO_DUPLICATE' ORDER BY score DESC"):
        reviews.append(dict(r))
    (out_dir / "review_queue.json").write_text(json.dumps(reviews, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest = {
        "workbench_version": VERSION,
        "source_game": stage.meta("source_game"),
        "target": target,
        "created_at": dt.datetime.now().isoformat(timespec="seconds"),
        "counts": {
            "players": stage.count("players"),
            "staff": stage.count("staff"),
            "issues": stage.conn.execute("SELECT COUNT(*) c FROM issues").fetchone()["c"],
            "matches": stage.conn.execute("SELECT COUNT(*) c FROM matches").fetchone()["c"],
        },
        "note": (
            "This bundle is intentionally intermediate/editor-friendly. "
            "FM24/FM26 field schemas may differ by editor/profile; feed these structured "
            "files into the corresponding FMME profile rather than assuming identical raw XML."
        )
    }
    (out_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

def ollama_resolve(prompt: str, model: str = "qwen2.5:7b") -> str:
    payload = json.dumps({
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": 0.0}
    }).encode("utf-8")
    req = urllib.request.Request(
        "http://127.0.0.1:11434/api/generate",
        data=payload,
        headers={"Content-Type":"application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = json.loads(resp.read().decode("utf-8"))
        return str(data.get("response","")).strip()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("1100x760")
        self.minsize(980, 680)
        self.source_path = tk.StringVar()
        self.source_game = tk.StringVar(value="Championship Manager 03/04")
        self.target_game = tk.StringVar(value="Football Manager 26")
        self.entity_file = tk.StringVar()
        self.stage_path = Path.cwd() / DB_NAME
        self.log_queue = []
        self._build()

    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, font=("Segoe UI", 18, "bold")).grid(row=0,column=0,sticky="w")
        ttk.Label(top, text=f"v{VERSION} • CM03/04-first", foreground="#555").grid(row=1,column=0,sticky="w")

        cfg = ttk.LabelFrame(self, text="Project", padding=10)
        cfg.pack(fill="x", padx=12, pady=(0,8))
        ttk.Label(cfg,text="Source game").grid(row=0,column=0,sticky="w",padx=4,pady=4)
        ttk.Combobox(cfg,textvariable=self.source_game,values=["Championship Manager 03/04","Generic CSV / JSON / XLSX"],state="readonly",width=34).grid(row=0,column=1,sticky="w",padx=4)
        ttk.Label(cfg,text="Target").grid(row=0,column=2,sticky="w",padx=4)
        ttk.Combobox(cfg,textvariable=self.target_game,values=["Football Manager 24","Football Manager 26"],state="readonly",width=22).grid(row=0,column=3,sticky="w",padx=4)
        ttk.Label(cfg,text="Source").grid(row=1,column=0,sticky="w",padx=4,pady=4)
        ttk.Entry(cfg,textvariable=self.source_path).grid(row=1,column=1,columnspan=2,sticky="ew",padx=4)
        ttk.Button(cfg,text="Browse…",command=self.browse_source).grid(row=1,column=3,sticky="w",padx=4)
        cfg.columnconfigure(1,weight=1)
        cfg.columnconfigure(2,weight=1)

        nb = ttk.Notebook(self)
        nb.pack(fill="both",expand=True,padx=12,pady=(0,12))

        self.tab_inspect = ttk.Frame(nb,padding=10)
        self.tab_import = ttk.Frame(nb,padding=10)
        self.tab_validate = ttk.Frame(nb,padding=10)
        self.tab_export = ttk.Frame(nb,padding=10)
        self.tab_ai = ttk.Frame(nb,padding=10)
        nb.add(self.tab_inspect,text="1. Inspect")
        nb.add(self.tab_import,text="2. Import")
        nb.add(self.tab_validate,text="3. Validate")
        nb.add(self.tab_export,text="4. Export")
        nb.add(self.tab_ai,text="5. AI Review")

        self._build_inspect()
        self._build_import()
        self._build_validate()
        self._build_export()
        self._build_ai()

        status = ttk.Label(self, text="Ready.", relief="sunken", anchor="w", padding=6)
        status.pack(fill="x",side="bottom")
        self.status = status

    def log(self,msg):
        self.output.insert("end", msg+"\n")
        self.output.see("end")
        self.update_idletasks()

    def set_status(self,msg):
        self.status.config(text=msg)

    def browse_source(self):
        p = filedialog.askdirectory(title="Select CM03/04 database folder")
        if p:
            self.source_path.set(p)
            self.set_status("Source selected.")

    def _build_inspect(self):
        f = self.tab_inspect
        bar=ttk.Frame(f); bar.pack(fill="x")
        ttk.Button(bar,text="Inspect Source",command=self.inspect_source).pack(side="left")
        ttk.Button(bar,text="Open Stage DB Location",command=self.choose_stage).pack(side="left",padx=8)
        self.inspect_text=tk.Text(f,wrap="none",height=24)
        self.inspect_text.pack(fill="both",expand=True,pady=(8,0))
        self.output=self.inspect_text

    def choose_stage(self):
        p=filedialog.asksaveasfilename(title="Choose staging SQLite file",initialfile=DB_NAME,defaultextension=".sqlite3")
        if p:
            self.stage_path=Path(p)
            self.set_status(f"Stage DB: {self.stage_path}")

    def _build_import(self):
        f=self.tab_import
        info=ttk.Label(f,text="Import a folder containing CSV/JSON/XLSX exports, or a single file. "
                          "CM03/04 binary parsing is kept separate until validated against real data.")
        info.pack(anchor="w")
        bar=ttk.Frame(f);bar.pack(fill="x",pady=10)
        ttk.Button(bar,text="Import Source",command=self.import_source).pack(side="left")
        ttk.Button(bar,text="Import Individual File",command=self.import_file).pack(side="left",padx=8)
        ttk.Button(bar,text="Reset Stage DB",command=self.reset_stage).pack(side="left")
        self.import_log=tk.Text(f,wrap="word")
        self.import_log.pack(fill="both",expand=True)

    def _build_validate(self):
        f=self.tab_validate
        bar=ttk.Frame(f);bar.pack(fill="x")
        ttk.Button(bar,text="Run Validation",command=self.run_validation).pack(side="left")
        ttk.Button(bar,text="Run Duplicate Matching",command=self.run_matches).pack(side="left",padx=8)
        self.issue_tree=ttk.Treeview(f,columns=("severity","entity","key","message","details"),show="headings")
        for col,w in zip(("severity","entity","key","message","details"),(90,100,160,330,360)):
            self.issue_tree.heading(col,text=col.title());self.issue_tree.column(col,width=w)
        self.issue_tree.pack(fill="both",expand=True,pady=(8,0))

    def _build_export(self):
        f=self.tab_export
        ttk.Label(f,text="Export a structured intermediate bundle for FM24/FM26/FMME. "
                  "The bundle preserves the normalized records and review queue.").pack(anchor="w")
        bar=ttk.Frame(f);bar.pack(fill="x",pady=12)
        ttk.Button(bar,text="Export Conversion Bundle…",command=self.export).pack(side="left")
        self.export_text=tk.Text(f,wrap="word")
        self.export_text.pack(fill="both",expand=True)

    def _build_ai(self):
        f=self.tab_ai
        ttk.Label(f,text="Optional local-Ollama review helper. It never changes data automatically; it returns a suggested resolution.").pack(anchor="w")
        grid=ttk.Frame(f);grid.pack(fill="x",pady=10)
        ttk.Label(grid,text="Model").grid(row=0,column=0,sticky="w")
        self.ollama_model=tk.StringVar(value="qwen2.5:7b")
        ttk.Entry(grid,textvariable=self.ollama_model,width=25).grid(row=0,column=1,sticky="w",padx=6)
        ttk.Label(grid,text="Player A / Player B").grid(row=1,column=0,sticky="w",pady=6)
        self.ai_a=tk.StringVar();self.ai_b=tk.StringVar()
        ttk.Entry(grid,textvariable=self.ai_a,width=42).grid(row=1,column=1,sticky="w",padx=6)
        ttk.Entry(grid,textvariable=self.ai_b,width=42).grid(row=1,column=2,sticky="w",padx=6)
        ttk.Button(grid,text="Ask Ollama",command=self.ask_ollama).grid(row=1,column=3,padx=6)
        self.ai_text=tk.Text(f,wrap="word")
        self.ai_text.pack(fill="both",expand=True)

    def get_adapter(self):
        return ADAPTERS["cm0304"] if self.source_game.get().startswith("Championship") else ADAPTERS["generic"]

    def inspect_source(self):
        src=self.source_path.get().strip()
        if not src:
            messagebox.showinfo("Source missing","Select a source folder or file first.");return
        path=Path(src)
        if not path.exists():
            messagebox.showerror("Source error","That path does not exist.");return
        try:
            rep=self.get_adapter().inspect(path)
            self.inspect_text.delete("1.0","end")
            self.inspect_text.insert("end",json.dumps(rep,indent=2,ensure_ascii=False))
            self.set_status("Inspection complete.")
        except Exception as e:
            messagebox.showerror("Inspection failed",str(e))

    def ensure_stage(self):
        self.stage_path.parent.mkdir(parents=True,exist_ok=True)
        return StageDB(self.stage_path)

    def reset_stage(self):
        try:
            if self.stage_path.exists():
                self.stage_path.unlink()
            StageDB(self.stage_path).close()
            self.import_log.insert("end",f"Created new stage DB: {self.stage_path}\n")
            self.set_status("Stage DB reset.")
        except Exception as e:
            messagebox.showerror("Reset failed",str(e))

    def import_file(self):
        p=filedialog.askopenfilename(filetypes=[("Data files","*.csv *.json *.xlsx"),("All files","*.*")])
        if not p:return
        try:
            stage=self.ensure_stage()
            rows=read_tabular(Path(p))
            if not rows: raise RuntimeError("No rows found.")
            stem=Path(p).stem.lower()
            entity="players" if any(x in stem for x in ("player","people","person")) else "clubs"
            stage.insert_rows(entity,rows)
            stage.meta("source_game","generic" if self.source_game.get().startswith("Generic") else "cm0304")
            stage.close()
            self.import_log.insert("end",f"Imported {len(rows):,} rows from {p} -> {entity}\n")
            self.set_status(f"Imported {len(rows):,} records.")
        except Exception as e:
            messagebox.showerror("Import failed",str(e))

    def import_source(self):
        src=self.source_path.get().strip()
        if not src:
            messagebox.showinfo("Source missing","Select a source first.");return
        path=Path(src)
        try:
            stage=self.ensure_stage()
            stage.meta("source_game","cm0304" if self.source_game.get().startswith("Championship") else "generic")
            adapter=self.get_adapter()
            try:
                n=adapter.extract(path,stage)
                self.import_log.insert("end",f"Extracted/imported {n:,} records.\n")
            except ParserNotImplementedError as e:
                self.import_log.insert("end",f"CM03/04 parser status: {e}\n")
                self.import_log.insert("end","The source was still inspected and its inventory was saved to the staging DB metadata.\n")
                self.inspect_text.delete("1.0","end")
                self.inspect_text.insert("end",stage.meta("source_inspection") or "")
            finally:
                counts={t:stage.count(t) for t in ("players","staff","generic_records")}
                stage.close()
            self.import_log.insert("end",f"Stage counts: {counts}\n")
            self.set_status("Import/extraction step complete.")
        except Exception as e:
            messagebox.showerror("Import failed",str(e))

    def run_validation(self):
        try:
            stage=StageDB(self.stage_path)
            stage.conn.execute("DELETE FROM issues")
            validate(stage)
            rows=stage.issues()
            self.refresh_issues(rows)
            stage.close()
            self.set_status(f"Validation complete: {len(rows):,} issues.")
        except Exception as e:
            messagebox.showerror("Validation failed",str(e))

    def run_matches(self):
        try:
            stage=StageDB(self.stage_path)
            run_matching(stage)
            # Add review issues for duplicate/review matches.
            for r in stage.conn.execute("SELECT * FROM matches WHERE status='REVIEW'"):
                stage.add_issue(Issue("WARN","match",str(r["source_row_id"]),
                                      "Ambiguous player match",json.dumps(dict(r))))
            stage.conn.commit()
            rows=stage.issues()
            self.refresh_issues(rows)
            stage.close()
            self.set_status("Duplicate matching complete.")
        except Exception as e:
            messagebox.showerror("Matching failed",str(e))

    def refresh_issues(self, rows):
        for x in self.issue_tree.get_children():self.issue_tree.delete(x)
        for r in rows:
            self.issue_tree.insert("", "end", values=(r["severity"],r["entity"],r["key"],r["message"],r["details"]))

    def export(self):
        d=filedialog.askdirectory(title="Choose output folder")
        if not d:return
        try:
            stage=StageDB(self.stage_path)
            target="fm26" if self.target_game.get().endswith("26") else "fm24"
            out=Path(d)/f"retrofm_{target}_export"
            export_bundle(stage,out,target)
            stage.close()
            self.export_text.delete("1.0","end")
            self.export_text.insert("end",f"Exported to:\n{out}\n\n")
            self.export_text.insert("end",(out/"manifest.json").read_text(encoding="utf-8"))
            self.set_status(f"Export complete: {out}")
            messagebox.showinfo("Export complete",f"Created:\n{out}")
        except Exception as e:
            messagebox.showerror("Export failed",str(e))

    def ask_ollama(self):
        a=self.ai_a.get().strip();b=self.ai_b.get().strip()
        if not a or not b:
            messagebox.showinfo("Missing records","Enter two player records.");return
        prompt=(f"We are reviewing two historical football database records for possible duplicate identity. "
                f"Do not invent facts. Compare only the text provided and return a concise assessment.\n\n"
                f"PLAYER A:\n{a}\n\nPLAYER B:\n{b}\n\n"
                f"Return JSON with keys match (true/false/uncertain), confidence (0-1), reasons (array).")
        try:
            self.ai_text.delete("1.0","end")
            self.ai_text.insert("end",ollama_resolve(prompt,self.ollama_model.get().strip() or "qwen2.5:7b"))
        except urllib.error.URLError as e:
            messagebox.showerror("Ollama unavailable","Could not reach http://127.0.0.1:11434. Make sure Ollama is running.")
        except Exception as e:
            messagebox.showerror("Ollama request failed",str(e))

def main():
    app=App()
    app.mainloop()

if __name__=="__main__":
    main()
