"""One-time image pull: saves team logos and drafted-player headshots into assets/.

Run after the draft completes, then commit and push the assets/ folder:

    python fetch_images.py ["path/to/workbook.xlsx"]

With no argument, you are offered the last workbook used (remembered in
.fetch_images_last_workbook) or a file dialog to pick a different one.

Logos: every team on the TeamList sheet.
Headshots: only players on the Draft sheet.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl
import requests

WORKBOOK = "2027 March Madness Player Pool Tracking Skeleton 2026 rnd5.xlsx"
ASSETS = Path(__file__).parent / "assets"
LAST_FILE = Path(__file__).parent / ".fetch_images_last_workbook"
TEAMS_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/"
    "mens-college-basketball/teams?limit=500"
)
ROSTER_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/"
    "mens-college-basketball/teams/{}/roster"
)
LOGO_URL = "https://a.espncdn.com/i/teamlogos/ncaa/500/{}.png"


# These must stay identical to the helpers in dashboard.py.
def norm_team(name):
    name = str(name).casefold().replace("saint", "st")
    return re.sub(r"[^a-z0-9]", "", name)


def norm_person(name):
    return re.sub(r"[^a-z0-9]", "", str(name).casefold())


def sheet_rows(wb, sheet):
    if sheet not in wb.sheetnames:
        return [], []
    rows = wb[sheet].iter_rows(values_only=True)
    header = next(rows, ())
    return [str(h).strip() if h is not None else "" for h in header], list(rows)


def get_json(url):
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    return resp.json()


def download(url, dest):
    try:
        resp = requests.get(url, timeout=20)
        if resp.status_code == 200 and "image" in resp.headers.get("content-type", ""):
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(resp.content)
            return True
    except Exception:
        pass
    return False


def roster_photos(team_id):
    photos = {}
    try:
        athletes = get_json(ROSTER_URL.format(team_id)).get("athletes", [])
    except Exception:
        return photos
    flat = []
    for entry in athletes:
        if isinstance(entry, dict) and "items" in entry:
            flat.extend(entry["items"])
        else:
            flat.append(entry)
    for athlete in flat:
        if not isinstance(athlete, dict):
            continue
        url = (athlete.get("headshot") or {}).get("href")
        if not url and athlete.get("id"):
            url = (
                "https://a.espncdn.com/i/headshots/mens-college-basketball/"
                f"players/full/{athlete['id']}.png"
            )
        if not url:
            continue
        for key in ("fullName", "displayName"):
            if athlete.get(key):
                photos.setdefault(norm_person(athlete[key]), (athlete.get("id"), url))
    return photos


def read_last_workbook():
    try:
        path = LAST_FILE.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    return path if path and Path(path).is_file() else None


def choose_workbook():
    """Command-line arg, else last used file (if confirmed), else a file dialog."""
    if len(sys.argv) > 1:
        return sys.argv[1]

    last = read_last_workbook()
    if last:
        answer = input(f"Use last workbook?\n  {last}\n[Y/n]: ").strip().lower()
        if answer in ("", "y", "yes"):
            return last

    try:
        import tkinter
        from tkinter import filedialog

        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        start = Path(last).parent if last else Path(__file__).parent
        chosen = filedialog.askopenfilename(
            title="Select the pool workbook",
            initialdir=str(start),
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        root.destroy()
    except Exception:
        chosen = input("Path to workbook: ").strip().strip('"')
    if not chosen:
        sys.exit("No workbook selected.")
    return chosen


def main():
    workbook_path = choose_workbook()
    if not Path(workbook_path).is_file():
        sys.exit(f"Workbook not found: {workbook_path}")
    LAST_FILE.write_text(str(Path(workbook_path).resolve()), encoding="utf-8")
    print(f"Using {workbook_path}")
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)

    # Team names from TeamList
    header, rows = sheet_rows(wb, "TeamList")
    if "Team" not in header:
        sys.exit("TeamList sheet with a Team column is required.")
    ti = header.index("Team")
    teams = {}
    for r in rows:
        if ti < len(r) and r[ti]:
            teams[norm_team(r[ti])] = str(r[ti]).strip()

    # Drafted players and their teams
    header, rows = sheet_rows(wb, "TeamData")
    pi, ci = header.index("Player"), header.index("Team")
    player_team = {}
    for r in rows:
        if pi < len(r) and r[pi] and ci < len(r) and r[ci]:
            player_team[str(r[pi]).strip().casefold()] = str(r[ci]).strip()

    _, rows = sheet_rows(wb, "Draft")
    drafted = sorted({str(r[2]).strip() for r in rows if len(r) > 2 and r[2]})

    print("Looking up ESPN team IDs...")
    espn = get_json(TEAMS_URL)["sports"][0]["leagues"][0]["teams"]
    ids = {}
    for entry in espn:
        t = entry.get("team", {})
        if not t.get("id"):
            continue
        for key in ("location", "shortDisplayName", "displayName", "name"):
            if t.get(key):
                ids.setdefault(norm_team(t[key]), t["id"])

    manifest = {"teams": {}, "players": {}}

    for key, name in sorted(teams.items()):
        team_id = ids.get(key)
        rel = f"logos/{team_id}.png" if team_id else None
        if rel and download(LOGO_URL.format(team_id), ASSETS / rel):
            manifest["teams"][key] = rel
        else:
            print(f"  MISSING logo: {name}")

    roster_cache = {}
    for player in drafted:
        team = player_team.get(player.casefold())
        team_id = ids.get(norm_team(team)) if team else None
        if not team_id:
            print(f"  MISSING headshot (no team): {player}")
            continue
        if team_id not in roster_cache:
            roster_cache[team_id] = roster_photos(team_id)
        found = roster_cache[team_id].get(norm_person(player))
        if not found:
            print(f"  MISSING headshot (not on roster): {player} ({team})")
            continue
        athlete_id, url = found
        rel = f"headshots/{athlete_id or norm_person(player)}.png"
        if download(url, ASSETS / rel):
            manifest["players"][f"{norm_team(team)}|{norm_person(player)}"] = rel
        else:
            print(f"  MISSING headshot (download failed): {player}")

    ASSETS.mkdir(exist_ok=True)
    (ASSETS / "manifest.json").write_text(
        json.dumps(manifest, indent=1, sort_keys=True), encoding="utf-8"
    )
    print(
        f"Saved {len(manifest['teams'])}/{len(teams)} logos and "
        f"{len(manifest['players'])}/{len(drafted)} headshots to {ASSETS}"
    )
    print("Now commit and push the assets/ folder.")


if __name__ == "__main__":
    main()
