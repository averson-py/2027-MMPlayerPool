"""Save the NCAA NET rankings to assets/net_rankings.json.

The ratings code blends these into the Barttorvik-based game odds once the NCAA has published the current
season's rankings (they are ignored while the page still shows last season).

    python net_rankings.py
"""
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import requests

from bracket_sim import blend_net_rankings
from torvik_ratings import load_ratings, season_year

NET_URL = "https://www.ncaa.com/rankings/basketball-men/d1/ncaa-mens-basketball-net-rankings"
NET_SNAPSHOT_PATH = Path(__file__).parent / "assets" / "net_rankings.json"


def _season_from_through(text):
    """Season (ending year) for a "Through Games Apr. 06 2026" style label."""
    match = re.search(r"([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})", text or "")
    if not match:
        return None
    month = datetime.strptime(match.group(1), "%b").month
    year = int(match.group(3))
    return year + 1 if month >= 10 else year


def download_net_rankings(path=NET_SNAPSHOT_PATH):
    """Download the NET table and write it to `path`. Returns (team count, season, through label)."""
    resp = requests.get(
        NET_URL,
        headers={"User-Agent": "MM-Player-Pool-Dashboard (personal use; refreshed once a day at most)"},
        timeout=30,
    )
    resp.raise_for_status()
    html = resp.text
    through = re.search(r"Through Games\s+([^<]+?)\s*<", html)
    through = through.group(1).strip() if through else ""
    ranks = {}
    for row in re.findall(r"<tr>(.*?)</tr>", html, re.S)[1:]:
        cells = [re.sub(r"<[^>]+>", "", c).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
        if len(cells) > 1 and cells[0].isdigit():
            ranks[cells[1]] = int(cells[0])
    season = _season_from_through(through)
    if not ranks or season is None:
        raise ValueError("Could not read the NET rankings table.")
    payload = {
        "season": season, "through": through, "source": NET_URL,
        "fetched": datetime.now().isoformat(timespec="seconds"), "ranks": ranks,
    }
    path = Path(path)
    path.parent.mkdir(exist_ok=True)
    path.write_text(json.dumps(payload, indent=1), encoding="utf-8")
    return len(ranks), season, through


def load_net_ranks(year=None):
    """({school: NET rank}, through label) from the saved snapshot; ({}, "") unless it is for this season."""
    try:
        payload = json.loads(NET_SNAPSHOT_PATH.read_text(encoding="utf-8"))
        if payload.get("season") != (year or season_year()):
            return {}, ""
        return {str(k): int(v) for k, v in payload["ranks"].items()}, str(payload.get("through", ""))
    except (OSError, ValueError, KeyError, TypeError):
        return {}, ""


def load_odds_ratings(year=None):
    """(Barttorvik ratings blended with NET if available, NET teams matched, NET 'through' label)."""
    ratings = load_ratings(year)
    net_ranks, through = load_net_ranks(year)
    blended, matched = blend_net_rankings(ratings, net_ranks)
    return blended, matched, through


if __name__ == "__main__":
    try:
        count, season, through = download_net_rankings()
    except (requests.RequestException, ValueError) as exc:
        sys.exit(f"Failed: {exc}")
    print(f"Saved NET rankings for {count} teams ({season} season, through {through}) to {NET_SNAPSHOT_PATH}")
