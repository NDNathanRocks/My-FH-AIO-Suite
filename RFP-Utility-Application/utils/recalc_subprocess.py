#!/usr/bin/env python3
"""Small CLI wrapper to run Excel recalculation in a separate process.

This script imports the real `_force_excel_recalc` from `utils.app_utils`
and executes it. It is designed to be launched as a separate Python process
so the main GUI process is not blocked or overwhelmed by COM-driven Excel CPU
usage (useful when packaged or under heavy load).
"""
import argparse
import sys
import os

def main(argv=None):
    parser = argparse.ArgumentParser(description="Run Excel recalc in subprocess.")
    parser.add_argument("--path", required=True, help="Workbook path")
    parser.add_argument("--sheet", default=None, help="Worksheet name")
    parser.add_argument("--refresh-external", type=int, choices=(0,1), default=1)
    parser.add_argument("--rows", default=None, help="Comma-separated 1-based row numbers")
    args = parser.parse_args(argv)

    row_numbers = None
    if args.rows:
        try:
            row_numbers = [int(x) for x in args.rows.split(",") if x.strip()]
        except Exception:
            row_numbers = None

    # Import here to avoid circular import during normal app startup
    try:
        from utils.app_utils import _force_excel_recalc
    except Exception as e:
        print(f"Failed to import _force_excel_recalc: {e}", file=sys.stderr)
        return 2

    try:
        _force_excel_recalc(args.path, sheet_name=args.sheet, refresh_external=bool(args.refresh_external), row_numbers=row_numbers)
    except Exception as e:
        print(f"Recalc subprocess failed: {e}", file=sys.stderr)
        return 3
    return 0

if __name__ == '__main__':
    sys.exit(main())
