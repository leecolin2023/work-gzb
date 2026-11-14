"""Validate and rerun the archived report model with the repository's verifier."""
import argparse
import subprocess
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", required=True, type=Path, help="Write the rerun to a separate file.")
    args = parser.parse_args()
    folder = Path(__file__).resolve().parent
    case = folder.name.split("_", 1)[0]
    source = folder / f"params_{case}.json"
    if args.out.resolve() == source:
        parser.error("--out must differ from the saved parameter file")
    tool = folder.parents[2] / "buliang-valuation-review/scripts/verify_formulas.py"
    subprocess.run([sys.executable, str(tool), "--params", str(source),
                    "--out", str(args.out), "--sensitivity"], check=True)


if __name__ == "__main__":
    main()
