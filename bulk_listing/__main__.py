"""Command line entry point for bulk product validation."""
import argparse
import json
from pathlib import Path

from .validator import validate_csv


def main() -> None:
    parser = argparse.ArgumentParser(description="Bulk listing data preflight")
    sub = parser.add_subparsers(dest="command", required=True)
    validate = sub.add_parser("validate", help="Validate a CSV, no actual upload")
    validate.add_argument("csv_file", type=Path)
    validate.add_argument("--output", type=Path, help="Save JSON report")
    args = parser.parse_args()
    if args.command == "validate":
        report = validate_csv(args.csv_file)
        text = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output:
            args.output.write_text(text + "\n", encoding="utf-8")
            print(f"Report saved to {args.output}")
        else:
            print(text)
        if report["error_count"]:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
