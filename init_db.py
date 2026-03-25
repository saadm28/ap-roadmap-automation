"""
Run once to create the database and tables.
Optionally seeds the bundled templates into the DB so the app uses DB-managed templates from the start.

Usage:
  python init_db.py                  # just create schema
  python init_db.py --seed-templates # also import bundled templates into DB
"""

import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import database as db


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--seed-templates", action="store_true", help="Import bundled templates into DB")
    args = parser.parse_args()

    print("Initialising database...")
    db.init_db()
    print("✓ Schema ready.")

    if args.seed_templates:
        templates_dir = Path(__file__).resolve().parent / "templates"
        mapping = {
            "Generic": templates_dir / "Generic_RoadMap.pptx",
            "Lawyers": templates_dir / "Lawyers_Roadmap.pptx",
        }
        for name, path in mapping.items():
            if path.exists():
                tid = db.save_template(name, path.read_bytes())
                print(f"✓ Template '{name}' seeded → id={tid}")
            else:
                print(f"  ⚠ Template '{name}' not found at {path}, skipping.")

    print("\nDone. Run: streamlit run app.py")


if __name__ == "__main__":
    main()
