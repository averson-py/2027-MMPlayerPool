"""Save Barttorvik team ratings to assets/torvik_ratings.json.

The dashboard reads this file when Barttorvik blocks its live download (Streamlit Cloud).
Used by the commissioner app's button and the scheduled GitHub Action:

    python torvik_ratings.py          # current season
    python torvik_ratings.py 2027     # a specific season
"""
import csv
import io
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import requests

TORVIK_CSV_URL = "https://barttorvik.com/{year}_team_results.csv"
SNAPSHOT_PATH = Path(__file__).parent / "assets" / "torvik_ratings.json"


def season_year(now=None):
    now = now or datetime.now()
    return now.year + 1 if now.month >= 10 else now.year


def download_torvik_ratings(year=None, path=SNAPSHOT_PATH):
    """Download the season's ratings and write them to `path`. Returns (team count, season year)."""
    year = year or season_year()
    resp = requests.get(
        TORVIK_CSV_URL.format(year=year),
        headers={"User-Agent": "MM-Player-Pool-Dashboard (personal use; refreshed once a day at most)"},
        timeout=30,
    )
    resp.raise_for_status()
    teams = {}
    for row in csv.DictReader(io.StringIO(resp.text)):
        try:
            teams[row["team"]] = [
                round(float(row["adjoe"]), 3), round(float(row["adjde"]), 3), round(float(row["adjt"]), 3),
                int(float(row["rank"])), row["record"], row["conf"],
            ]
        except (KeyError, TypeError, ValueError):
            continue
    if not teams:
        raise ValueError("Barttorvik returned no team ratings.")
    payload = {
        "year": year,
        "source": "https://barttorvik.com",
        "fetched": datetime.now().isoformat(timespec="seconds"),
        "teams": teams,
    }
    path = Path(path)
    path.parent.mkdir(exist_ok=True)
    path.write_text(json.dumps(payload, indent=1), encoding="utf-8")
    return len(teams), year


def push_snapshot(path=SNAPSHOT_PATH):
    """Commit and push only the snapshot file. Raises RuntimeError with git's message on failure."""
    path = Path(path)

    def git(*args):
        result = subprocess.run(["git", *args], cwd=path.parent, capture_output=True, text=True)
        if result.returncode:
            raise RuntimeError((result.stderr or result.stdout).strip() or f"git {args[0]} failed")
        return result.stdout

    if not git("status", "--porcelain", "--", path.name).strip():
        return False
    git("add", "--", path.name)
    git("commit", "-m", "Update Torvik ratings", "--", path.name)
    git("push")
    return True


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--push"]
    try:
        count, season = download_torvik_ratings(int(args[0]) if args else None)
        print(f"Saved Torvik ratings for {count} teams ({season}) to {SNAPSHOT_PATH}")
        if "--push" in sys.argv:
            print("Pushed." if push_snapshot() else "No changes to push.")
    except (requests.RequestException, RuntimeError) as exc:
        sys.exit(f"Failed: {exc}")
