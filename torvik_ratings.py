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

from bracket_sim import torvik_key

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


def load_ratings(year=None):
    """{bracket_sim.torvik_key(team): (adjoe, adjde, adjt, rank, record, conf)} from the saved snapshot.

    Empty if the snapshot is missing or is for a different season.
    """
    try:
        payload = json.loads(SNAPSHOT_PATH.read_text(encoding="utf-8"))
        if payload.get("year") != (year or season_year()):
            return {}
        ratings = {}
        for name, values in payload["teams"].items():
            values = list(values) + [None] * (6 - len(values))
            ratings[torvik_key(name)] = (
                float(values[0]), float(values[1]), float(values[2]), values[3], values[4], values[5],
            )
        return ratings
    except (OSError, ValueError, KeyError, TypeError):
        return {}


def push_paths(paths, message, repo_dir=SNAPSHOT_PATH.parent.parent):
    """Commit and push only `paths` (relative to the repo). Returns False if nothing changed; raises RuntimeError on git errors."""

    def git(*args):
        result = subprocess.run(["git", *args], cwd=repo_dir, capture_output=True, text=True)
        if result.returncode:
            raise RuntimeError((result.stderr or result.stdout).strip() or f"git {args[0]} failed")
        return result.stdout

    paths = [str(p) for p in paths]
    if not git("status", "--porcelain", "--", *paths).strip():
        return False
    git("add", "--", *paths)
    git("commit", "-m", message, "--", *paths)
    git("push")
    return True


def push_snapshot():
    paths = [p for p in ("assets/torvik_ratings.json", "assets/net_rankings.json") if (SNAPSHOT_PATH.parent.parent / p).exists()]
    return push_paths(paths, "Update team ratings")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--push"]
    try:
        count, season = download_torvik_ratings(int(args[0]) if args else None)
        print(f"Saved Torvik ratings for {count} teams ({season}) to {SNAPSHOT_PATH}")
        if "--push" in sys.argv:
            print("Pushed." if push_snapshot() else "No changes to push.")
    except (requests.RequestException, RuntimeError) as exc:
        sys.exit(f"Failed: {exc}")
