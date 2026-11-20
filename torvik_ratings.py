"""Save Barttorvik team ratings to assets/torvik_ratings.json.

The dashboard reads this file when Barttorvik blocks its live download (Streamlit Cloud).
Used by the commissioner app's button and the scheduled GitHub Action:

    python torvik_ratings.py          # current season
    python torvik_ratings.py 2027     # a specific season
"""
import csv
import io
import json
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


if __name__ == "__main__":
    count, season = download_torvik_ratings(int(sys.argv[1]) if len(sys.argv) > 1 else None)
    print(f"Saved Torvik ratings for {count} teams ({season}) to {SNAPSHOT_PATH}")
