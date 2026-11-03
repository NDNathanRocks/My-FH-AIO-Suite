"""Command-line maintenance for the Inbox Email Exporter (no Outlook required).

  python outlook_export_tools.py status
  python outlook_export_tools.py backup
  python outlook_export_tools.py migrate
  python outlook_export_tools.py reconcile [--output-root PATH]
  python outlook_export_tools.py rollback-schema      # drops only the v1.6 manifest/attempt/run tables

Defaults come from config.json and outlook_export.db beside this script.
Use --db to point at a copy when testing.
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

from exporter_core.database import backup_database, migrate, pending_changes, rollback_schema, tables
from exporter_core.reconcile import reconcile
from exporter_core.reports import reconciliation_text, write_reconciliation

HERE = Path(__file__).resolve().parent


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["status", "backup", "migrate", "reconcile", "rollback-schema"])
    parser.add_argument("--db", type=Path, default=HERE / "outlook_export.db")
    parser.add_argument("--backup-dir", type=Path, default=HERE / "backup")
    parser.add_argument("--output-root", type=Path, default=None)
    parser.add_argument("--report-dir", type=Path, default=HERE / "logs")
    parser.add_argument("--yes", action="store_true", help="skip confirmation for rollback-schema")
    args = parser.parse_args(argv)
    if args.command != "migrate" and not args.db.exists():
        print(f"Database not found: {args.db}", file=sys.stderr)
        return 2
    if args.command == "status":
        with sqlite3.connect(f"file:{args.db.as_posix()}?mode=ro", uri=True) as c:
            print(json.dumps({"tables": sorted(tables(c)), "pending_migration": pending_changes(c),
                              "statuses": dict(c.execute("SELECT status, COUNT(*) FROM exports GROUP BY status").fetchall())}, indent=2))
    elif args.command == "backup":
        print(backup_database(args.db, args.backup_dir, "manual"))
    elif args.command == "migrate":
        print(json.dumps(migrate(args.db, args.backup_dir), indent=2))
    elif args.command == "reconcile":
        root = args.output_root
        if root is None:
            config = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
            root = Path(config["output_root"])
        migrate(args.db, args.backup_dir)
        with sqlite3.connect(args.db) as c:
            interrupted = c.execute("SELECT run_id, mode, started FROM export_runs WHERE finished IS NULL").fetchall()
        rec = reconcile(args.db, root, None, interrupted_runs=interrupted)
        paths = write_reconciliation(args.report_dir, rec, datetime.now().strftime("%Y%m%d_%H%M%S") + "_CLI")
        print(reconciliation_text(rec))
        print("\n" + "\n".join(f"{k}: {v}" for k, v in paths.items()))
        return 0 if rec.overall_status != "incomplete" else 1
    elif args.command == "rollback-schema":
        if not args.yes and input("Drop manifest/attempt/run tables (exports is kept)? Type YES: ") != "YES":
            print("Cancelled.")
            return 1
        print(json.dumps(rollback_schema(args.db, args.backup_dir), indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
