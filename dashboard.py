
import colorsys
import csv
import io
import math
import os
import re
import secrets
from datetime import datetime
import html as html_lib
import json
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path


import openpyxl
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

from bracket_sim import (
    DEFAULT_TEMPO, GAME_FPTS_SPREAD, GAME_MARGIN_SD, SEED_STRENGTH, SIM_SEED, TORVIK_ALIASES,
    WIN_PROB_SIMS, blend_net_rankings, net_ranks_by_key, rating_key, read_bracket_teams, read_final_four_config,
    run_bracket,
)
from bracket_sim import norm_team_name as _norm_team_name
from bracket_sim import torvik_key as _torvik_key
from net_rankings import load_net_ranks


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
DEFAULT_WORKBOOK = "2027 March Madness Player Pool Tracking Skeleton 2026 rnd5.xlsx"
ONEDRIVE_FILE_PATH = "2027MMPlayerPoolStreamlit/2027 March Madness Player Pool Tracking.xlsx"
# Public read-only OneDrive sharing link. Store this in Streamlit Secrets as ONEDRIVE_SHARE_URL.
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
OAUTH_AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0"
GRAPH_SCOPES = "Files.Read User.Read offline_access"

DRAFT_ORDER_SHEET = "DraftOrder"
DRAFT_SHEET = "Draft"
TEAM_DATA_SHEET = "TeamData"
TEAMS_OUT_SHEET = "TeamsOut"
GAME_RESULTS_SHEET = "GameResults"
CONFIG_SHEET = "Config"
TEAM_LIST_SHEET = "TeamList"

DRAFTER_START_ROW = 2

SEED_AVERAGE_GAMES = {
    1: 4.34, 2: 3.35, 3: 2.84, 4: 2.57,
    5: 2.14, 6: 2.05, 7: 1.88, 8: 1.69,
    9: 1.63, 10: 1.58, 11: 1.53, 12: 1.50,
    13: 1.24, 14: 1.15, 15: 1.10, 16: 1.01,
}

# The original Tkinter app uses these same table colors.
BLUE = "#4472C4"
LIGHT_BLUE = "#DDEBF7"
LIGHT_GREEN = "#E2F0D9"
LIGHT_YELLOW = "#FFF2CC"
LIGHT_RED = "#F4CCCC"
LIGHT_ORANGE = "#FCE4D6"
LIGHT_PURPLE = "#D9D2E9"


# ---------------------------------------------------------------------
# Data / calculation layer
# This intentionally contains no Tkinter code, so it can run locally
# and later on Streamlit Community Cloud/Render.
# ---------------------------------------------------------------------
def sum_numeric_values(values):
    numeric_values = []
    for value in values:
        try:
            numeric_values.append(float(value))
        except (TypeError, ValueError):
            continue
    if not numeric_values:
        return None
    total = sum(numeric_values)
    return int(total) if total.is_integer() else total


def sum_existing_values(values):
    return sum_numeric_values(values)


def weighted_fantasy_points(points, rebounds, assists, weights):
    """Calculate fantasy points using the scoring weights from Config!A2:C2."""
    values = [points, rebounds, assists]
    numeric = []
    for value in values:
        try:
            numeric.append(float(value))
        except (TypeError, ValueError):
            numeric.append(None)

    weighted = []
    for value, weight in zip(numeric, weights):
        if value is not None:
            weighted.append(value * weight)

    if not weighted:
        return None

    total = sum(weighted)
    return int(total) if total.is_integer() else total


def format_stat(value):
    if value is None:
        return ""
    if isinstance(value, float):
        return f"{value:.1f}"
    return str(value)


def format_signed_stat(value):
    if value is None:
        return ""
    try:
        return f"{float(value):+.1f}"
    except (TypeError, ValueError):
        return ""


def draft_position_sort_key(value):
    try:
        round_number, pick_number = str(value).split(".", 1)
        return int(round_number), int(pick_number)
    except (TypeError, ValueError):
        return 999999, str(value)


def count_blank_values(values):
    return sum(
        value is None or (isinstance(value, str) and not value.strip())
        for value in values
    )


def count_values_not_out(values):
    count = 0
    for value in values:
        if value is None:
            continue
        if isinstance(value, str):
            normalized = value.strip().casefold()
            if not normalized or normalized == "out":
                continue
        count += 1
    return count


def count_game_statuses(values):
    games_played = 0
    games_not_input = 0

    for start in range(0, len(values), 3):
        game_values = values[start:start + 3]
        has_out_marker = any(
            isinstance(value, str) and value.strip().casefold() == "out"
            for value in game_values
        )
        if has_out_marker:
            continue

        def is_blank(value):
            return value is None or (
                isinstance(value, str) and not value.strip()
            )

        if all(is_blank(value) for value in game_values):
            games_not_input += 1
        else:
            games_played += 1

    return games_played, games_not_input


def average_games_for_seed(seed):
    if seed is None:
        return None
    try:
        seed_number = int(float(seed))
    except (TypeError, ValueError):
        match = re.search(r"\d+", str(seed))
        if match is None:
            return None
        seed_number = int(match.group(0))
    return SEED_AVERAGE_GAMES.get(seed_number)


def color_scale(value, minimum, maximum, reverse=False):
    if value is None or minimum is None or maximum is None:
        return "#FFFFFF"
    if maximum == minimum:
        return "#FFEB84"

    ratio = (float(value) - float(minimum)) / (float(maximum) - float(minimum))
    ratio = max(0.0, min(1.0, ratio))
    if reverse:
        ratio = 1.0 - ratio

    red = (248, 105, 107)
    yellow = (255, 235, 132)
    green = (99, 190, 123)

    if ratio <= 0.5:
        blend = ratio * 2
        start, end = red, yellow
    else:
        blend = (ratio - 0.5) * 2
        start, end = yellow, green

    rgb = tuple(
        round(start[i] + (end[i] - start[i]) * blend)
        for i in range(3)
    )
    return "#%02X%02X%02X" % rgb


def red_white_green_scale(value, minimum, maximum, reverse=False):
    if value is None or minimum is None or maximum is None:
        return "#FFFFFF"
    if maximum == minimum:
        return "#FFFFFF"

    ratio = (float(value) - float(minimum)) / (float(maximum) - float(minimum))
    ratio = max(0.0, min(1.0, ratio))
    if reverse:
        ratio = 1.0 - ratio

    red = (248, 105, 107)
    white = (255, 255, 255)
    green = (99, 190, 123)

    if ratio <= 0.5:
        blend = ratio * 2
        start, end = red, white
    else:
        blend = (ratio - 0.5) * 2
        start, end = white, green

    rgb = tuple(
        round(start[i] + (end[i] - start[i]) * blend)
        for i in range(3)
    )
    return "#%02X%02X%02X" % rgb


# Tournament bracket functions
TEAM_SEEDS = {}

ROUND_LOST_INDEX = [
    ("first four", -1),
    ("first round", 0),
    ("round of 64", 0),
    ("second round", 1),
    ("round of 32", 1),
    ("sweet", 2),
    ("elite", 3),
    ("final four", 4),
    ("semi", 4),
    ("title", 5),
    ("champ", 5),
]


def round_lost_index(round_text):
    """Map a TeamsOut round label to a bracket round index (-1 = First Four)."""
    text = str(round_text or "").strip().casefold()
    for key, idx in ROUND_LOST_INDEX:
        if key in text:
            return idx
    return None


def lost_in_round(team, round_idx, teams_out):
    """True only if the team was eliminated in this specific round."""
    if team is None or "TBD" in str(team):
        return False
    key = str(team).casefold()
    if key not in teams_out:
        return False
    lost_idx = round_lost_index(teams_out[key])
    if lost_idx is None:
        return True
    return lost_idx == round_idx

ESPN_TEAMS_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/"
    "mens-college-basketball/teams?limit=500"
)
ESPN_LOGO_URL = "https://a.espncdn.com/i/teamlogos/ncaa/500/{}.png"
ESPN_SCOREBOARD_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/"
    "mens-college-basketball/scoreboard"
)


@st.cache_data(ttl=86400, show_spinner=False)
def get_espn_team_ids():
    """Return {normalized team name: ESPN team id} from ESPN's teams API."""
    ids = {}
    try:
        resp = requests.get(ESPN_TEAMS_URL, timeout=10)
        resp.raise_for_status()
        teams = resp.json()["sports"][0]["leagues"][0]["teams"]
    except Exception:
        return ids

    for entry in teams:
        team = entry.get("team", {})
        team_id = team.get("id")
        if not team_id:
            continue
        for key in ("location", "shortDisplayName", "displayName", "name"):
            value = team.get(key)
            if value:
                ids.setdefault(_norm_team_name(value), team_id)
    return ids


ASSETS_DIR = Path(__file__).parent / "assets"
ASSETS_BASE_URL = (
    "https://raw.githubusercontent.com/averson-py/2027-MMPlayerPool/main/assets/"
)


@st.cache_data(show_spinner=False)
def load_image_manifest(_mtime):
    """Stored image index written by fetch_images.py (None if not run yet)."""
    try:
        data = json.loads((ASSETS_DIR / "manifest.json").read_text(encoding="utf-8"))
        return {"teams": data.get("teams", {}), "players": data.get("players", {})}
    except Exception:
        return None


def get_image_manifest():
    try:
        mtime = (ASSETS_DIR / "manifest.json").stat().st_mtime
    except OSError:
        return None
    return load_image_manifest(mtime)


def get_logo_url(team):
    if not team or "TBD" in str(team):
        return None
    manifest = get_image_manifest()
    if manifest is not None:
        rel = manifest["teams"].get(_norm_team_name(team))
        return ASSETS_BASE_URL + rel if rel else None
    team_id = get_espn_team_ids().get(_norm_team_name(team))
    return ESPN_LOGO_URL.format(team_id) if team_id else None


LOGO_ONERROR = (
    "if(!this.dataset.r){this.dataset.r=1;var i=this,s=this.src;"
    "setTimeout(function(){i.src=s+'?r=1'},1500)}"
    "else{this.style.visibility='hidden'}"
)


def _logo_img(team):
    url = get_logo_url(team)
    if not url:
        return '<span class="team-logo placeholder"></span>'
    return f'<img class="team-logo" src="{url}" alt="" onerror="{LOGO_ONERROR}">'


ESPN_ROSTER_URL = (
    "https://site.api.espn.com/apis/site/v2/sports/basketball/"
    "mens-college-basketball/teams/{}/roster"
)


def _norm_person_name(name):
    return re.sub(r"[^a-z0-9]", "", str(name).casefold())


@st.cache_data(ttl=86400, show_spinner=False)
def get_espn_roster_photos(team_id, league="mens-college-basketball"):
    """Return {normalized player name: headshot URL} for one ESPN team."""
    photos = {}
    try:
        resp = requests.get(
            ESPN_ROSTER_URL.format(team_id).replace("mens-college-basketball", league),
            timeout=10,
        )
        resp.raise_for_status()
        athletes = resp.json().get("athletes", [])
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
                f"https://a.espncdn.com/i/headshots/{league}/"
                f"players/full/{athlete['id']}.png"
            )
        if not url:
            continue
        for key in ("fullName", "displayName"):
            if athlete.get(key):
                photos.setdefault(_norm_person_name(athlete[key]), url)
    return photos


@st.cache_data(ttl=86400, show_spinner=False)
def photo_url_works(url):
    try:
        resp = requests.get(url, timeout=6, stream=True)
        ok = resp.status_code == 200 and "image" in resp.headers.get("content-type", "")
        resp.close()
        return ok
    except Exception:
        return False


# Prefix a value with "NBA:" to search an NBA roster instead of a college one.
PLAYER_PHOTO_TEAM_OVERRIDES = {
    "Elliot Cadeau": "Michigan",
    "Morez Johnson Jr.": "NBA:Dallas Mavericks",
    "Yaxel Lendeborg": "NBA:Golden State Warriors",
    "Trey McKenney": "Michigan",
    "Aday Mara": "NBA:Oklahoma City Thunder",
    "Alex Karaban": "NBA:Sacramento Kings",
    "Tarris Reed Jr.": "NBA:San Antonio Spurs",
    "Solo Ball": "UConn",
    "JT Toppin": "Texas Tech",
    "Caleb Wilson": "NBA:Chicago Bulls",
    "Mikel Brown Jr.": "NBA:Brooklyn Nets",
    "Joshua Jefferson": "NBA:Brooklyn Nets",
    "Braden Huff": "Gonzaga",
}


@st.cache_data(ttl=86400, show_spinner=False)
def get_espn_nba_team_ids():
    ids = {}
    try:
        resp = requests.get(
            "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams",
            timeout=10,
        )
        resp.raise_for_status()
        teams = resp.json()["sports"][0]["leagues"][0]["teams"]
    except Exception:
        return ids
    for entry in teams:
        team = entry.get("team", {})
        if not team.get("id"):
            continue
        for key in ("displayName", "shortDisplayName", "name"):
            if team.get(key):
                ids.setdefault(_norm_team_name(team[key]), team["id"])
    return ids


def get_player_photo_url(player_name, team):
    if not team or not player_name:
        return None
    override = {
        _norm_person_name(k): v for k, v in PLAYER_PHOTO_TEAM_OVERRIDES.items()
    }.get(_norm_person_name(player_name))
    if override and override.upper().startswith("NBA:"):
        nba_team = override[4:].strip()
        team_id = get_espn_nba_team_ids().get(_norm_team_name(nba_team))
        if not team_id:
            return None
        url = get_espn_roster_photos(team_id, "nba").get(_norm_person_name(player_name))
        return url if url and photo_url_works(url) else None
    if override:
        team = override
    manifest = get_image_manifest()
    if manifest is not None:
        rel = manifest["players"].get(
            f"{_norm_team_name(team)}|{_norm_person_name(player_name)}"
        )
        if rel:
            return ASSETS_BASE_URL + rel
        if not override:
            return None
    team_id = get_espn_team_ids().get(_norm_team_name(team))
    if not team_id:
        return None
    url = get_espn_roster_photos(team_id).get(_norm_person_name(player_name))
    return url if url and photo_url_works(url) else None


def lookup_seed(team, default):
    return TEAM_SEEDS.get(str(team).casefold(), default)


@st.cache_data(ttl=60, show_spinner=False)
def get_espn_tournament_games():
    """Return final NCAA tournament games from ESPN's scoreboard (cached 60s).

    Each game: {"teams": [aliases_a, aliases_b], "scores": [a, b]}
    """
    now = datetime.now()
    upcoming = now.year + 1 if now.month >= 10 else now.year

    def fetch_day(day):
        try:
            resp = requests.get(
                ESPN_SCOREBOARD_URL,
                params={"dates": day.strftime("%Y%m%d"), "groups": 100, "limit": 500},
                timeout=5,
            )
            resp.raise_for_status()
            return resp.json().get("events", [])
        except Exception:
            return []

    events = []
    for year in (upcoming, upcoming - 1):
        days = [
            d for d in pd.date_range(f"{year}-03-14", f"{year}-04-10")
            if d <= pd.Timestamp(now)
        ]
        if not days:
            continue
        with ThreadPoolExecutor(max_workers=8) as pool:
            events = [e for day_events in pool.map(fetch_day, days) for e in day_events]
        if events:
            break

    games = []
    for event in events:
        try:
            comp = event["competitions"][0]
            if event["status"]["type"]["state"] != "post":
                continue
            teams, scores = [], []
            for c in comp["competitors"][:2]:
                team = c.get("team", {})
                teams.append({
                    _norm_team_name(team[key])
                    for key in ("location", "shortDisplayName", "displayName", "name")
                    if team.get(key)
                })
                scores.append(int(float(c.get("score", 0) or 0)))
            if len(teams) == 2:
                games.append({"teams": teams, "scores": scores})
        except Exception:
            continue
    return games


def find_game_scores(team_a, team_b):
    """Return (score_a, score_b) for the final game between two teams, else None."""
    if "TBD" in str(team_a) or "TBD" in str(team_b):
        return None
    na, nb = _norm_team_name(team_a), _norm_team_name(team_b)
    for game in get_espn_tournament_games():
        t0, t1 = game["teams"]
        if na in t0 and nb in t1:
            return game["scores"][0], game["scores"][1]
        if na in t1 and nb in t0:
            return game["scores"][1], game["scores"][0]
    return None


def _score_html(score, other):
    if score is None:
        return '<span class="team-score"></span>'
    cls = "team-score score-win" if score > other else "team-score score-loss"
    return f'<span class="{cls}">{score}</span>'


def _win_prob_html(percent, favorite):
    cls = "team-score win-prob fav" if favorite else "team-score win-prob"
    return f'<span class="{cls}" title="Chance to win this game">{percent}%</span>'


@st.cache_data(show_spinner=False)
def _torvik_badge_css():
    """Style block that draws the toRvik logo on bracket cards (inlined so the bracket frame needs no files)."""
    import base64

    try:
        encoded = base64.b64encode((ASSETS_DIR / "torvik_logo.png").read_bytes()).decode()
    except OSError:
        return ""
    return (
        "<style>"
        ".game-label {display: flex; align-items: center; justify-content: space-between; min-height: 16px;}"
        ".torvik-badge {display: inline-block; width: 13px; height: 15px; "
        f"background: url(data:image/png;base64,{encoded}) center / contain no-repeat;}}"
        "</style>"
    )


def build_tournament_bracket(bracket_teams, teams_out=None):
    """
    Build tournament bracket structure from teams dict.
    bracket_teams: {seed: {region: [teams]}}
    teams_out: {team_name_lower: round_lost} - teams that were eliminated

    Supports First Four games when multiple teams share the same seed in a region.
    Uses TeamsOut data to determine winners. If a team is in TeamsOut, it lost in that round.
    Returns a list of rounds, where each round contains regions,
    and each region contains matchups.
    """
    if teams_out is None:
        teams_out = {}

    regions = ["EAST", "WEST", "MIDWEST", "SOUTH"]
    rounds = []

    # Create a reverse mapping: team name -> (seed, region) for quick lookup
    # This helps us find which seed a team belongs to when it's in teams_out
    team_to_seed_region = {}
    for seed in bracket_teams:
        for region in bracket_teams[seed]:
            for team in bracket_teams[seed][region]:
                team_to_seed_region[team.casefold()] = (seed, region)
                TEAM_SEEDS[team.casefold()] = seed

    # Track winners at each round to use in future rounds
    # round_winners[round_num][region] = {matchup_idx: winner_team_name}
    round_winners = {}
    round_seeds = {}

    # Track which seeds have First Four games (multiple teams at same seed)
    first_four_seeds = set()

    # First detect First Four games (same seed, same region, multiple teams)
    first_four = {}
    for region in regions:
        first_four[region] = []

    for seed in bracket_teams:
        for region in bracket_teams[seed]:
            teams = bracket_teams[seed][region]
            if len(teams) > 1:
                # Multiple teams at same seed in same region = First Four game
                first_four_seeds.add(seed)
                # Pair consecutive teams for First Four
                for i in range(0, len(teams), 2):
                    if i + 1 < len(teams):
                        team_a = teams[i]
                        team_b = teams[i + 1]

                        # Determine winner based on TeamsOut
                        team_a_lost = lost_in_round(team_a, -1, teams_out)
                        team_b_lost = lost_in_round(team_b, -1, teams_out)

                        first_four[region].append({
                            "team_a": team_a,
                            "seed_a": seed,
                            "team_b": team_b,
                            "seed_b": seed,
                            "round": "First Four",
                            "winner": team_b if team_a_lost else (team_a if team_b_lost else None),
                            "team_a_eliminated": team_a_lost,
                            "team_b_eliminated": team_b_lost,
                        })
                    elif i < len(teams):
                        # Odd team out gets a bye shown as TBD opponent
                        team_a = teams[i]
                        first_four[region].append({
                            "team_a": team_a,
                            "seed_a": seed,
                            "team_b": f"({seed}) TBD (FF)",
                            "seed_b": seed,
                            "round": "First Four",
                            "winner": team_a,
                            "team_a_eliminated": False,
                            "team_b_eliminated": False,
                        })

    # Add First Four round if it has any games
    if any(first_four.values()):
        rounds.append({"name": "First Four", "regions": first_four})

    # Helper function to get team for a seed in a region, considering First Four winners
    def get_team_for_seed(seed, region, round_num):
        # Round of 64 - show actual teams from bracket
        if round_num == 0:
            if seed in bracket_teams and region in bracket_teams[seed]:
                teams = bracket_teams[seed][region]
                if not teams:
                    return f"({seed}) TBD"

                if len(teams) == 1:
                    # Single team at this seed - show it
                    return teams[0]
                else:
                    # Multiple teams (First Four) - check which team is NOT eliminated (the First Four winner)
                    active_teams = [t for t in teams if not lost_in_round(t, -1, teams_out)]

                    if len(active_teams) == 1:
                        # One team won (other is eliminated)
                        return active_teams[0]
                    elif len(active_teams) == 0:
                        # Both eliminated (shouldn't happen) - return first
                        return teams[0]
                    else:
                        # Both still active (First Four game not yet played)
                        return f"({seed}) TBD"

            return f"({seed}) TBD"

        # Future rounds - look up winners from previous round
        if round_num - 1 in round_winners:
            prev_round_winners = round_winners[round_num - 1]
            if region in prev_round_winners:
                # Calculate which matchup index this seed came from
                # This is complex, so we'll use a different approach below
                pass

        return f"({seed}) TBD"

    # Helper to determine winner of a matchup
    def get_winner(team_a, team_b, round_num):
        # If either team is TBD or None, we can't determine a winner
        if team_a is None or "TBD" in str(team_a) or team_b is None or "TBD" in str(team_b):
            return None

        team_a_loss = lost_in_round(team_a, round_num, teams_out)
        team_b_loss = lost_in_round(team_b, round_num, teams_out)

        if team_a_loss and not team_b_loss:
            return team_b
        if team_b_loss and not team_a_loss:
            return team_a
        # Neither (or both) lost in this round - can't determine a winner yet
        return None

    # Helper to check if team lost in this round
    def is_eliminated(team_name, round_num):
        return lost_in_round(team_name, round_num, teams_out)

    # Main bracket rounds
    round_names = ["Round of 64", "Round of 32", "Sweet Sixteen", "Elite Eight", "Final Four", "Championship"]
    matchup_structure = [
        [(1, 16), (8, 9), (5, 12), (4, 13), (6, 11), (3, 14), (7, 10), (2, 15)],  # Round 1
        [(1, 8), (5, 4), (6, 3), (7, 2)],  # Round 2
        [(1, 5), (6, 7)],  # Round 3
        [(1, 6)],  # Round 4
        [(1, 1), (1, 1)],  # Round 5 (semifinal structure, placeholder)
        [(1, 1)],  # Round 6 (championship, placeholder)
    ]

    for round_num, matchups_seeds in enumerate(matchup_structure):
        round_obj = {"name": round_names[round_num], "regions": {}}
        round_winners[round_num] = {}
        round_seeds[round_num] = {}

        for region in regions:
            round_obj["regions"][region] = []
            round_winners[round_num][region] = {}
            round_seeds[round_num][region] = {}

            for matchup_idx, (seed_a, seed_b) in enumerate(matchups_seeds):
                team_a = None
                team_b = None

                # Round of 64 - get from bracket_teams
                if round_num == 0:
                    team_a = get_team_for_seed(seed_a, region, round_num)
                    team_b = get_team_for_seed(seed_b, region, round_num)
                else:
                    # Future rounds - get from previous round's winners
                    # The winner from matchup (matchup_idx * 2) plays winner from (matchup_idx * 2 + 1)
                    prev_matchup_a = matchup_idx * 2
                    prev_matchup_b = matchup_idx * 2 + 1

                    if region in round_winners[round_num - 1]:
                        prev_winners = round_winners[round_num - 1][region]
                        team_a = prev_winners.get(prev_matchup_a)
                        team_b = prev_winners.get(prev_matchup_b)

                    # If no winner found, show TBD with the best (lowest) seed that could advance
                    prev_seeds = round_seeds[round_num - 1].get(region, {})
                    if team_a is None:
                        seed_a = prev_seeds.get(prev_matchup_a, seed_a)
                        team_a = f"({seed_a}) TBD"
                    else:
                        seed_a = lookup_seed(team_a, seed_a)
                    if team_b is None:
                        seed_b = prev_seeds.get(prev_matchup_b, seed_b)
                        team_b = f"({seed_b}) TBD"
                    else:
                        seed_b = lookup_seed(team_b, seed_b)

                # Determine winner of this matchup
                winner = get_winner(team_a, team_b, round_num)
                round_winners[round_num][region][matchup_idx] = winner

                # Seed that advances from this game (best possible seed if undecided)
                if winner:
                    round_seeds[round_num][region][matchup_idx] = seed_a if winner == team_a else seed_b
                else:
                    try:
                        round_seeds[round_num][region][matchup_idx] = min(int(seed_a), int(seed_b))
                    except (TypeError, ValueError):
                        round_seeds[round_num][region][matchup_idx] = seed_a

                # Store matchup with elimination flags
                matchup_obj = {
                    "team_a": team_a,
                    "seed_a": seed_a,
                    "team_b": team_b,
                    "seed_b": seed_b,
                    "winner": winner,
                    "team_a_eliminated": is_eliminated(team_a, round_num),
                    "team_b_eliminated": is_eliminated(team_b, round_num),
                }

                round_obj["regions"][region].append(matchup_obj)

        rounds.append(round_obj)

    return rounds


def build_final_four_brackets(regional_winners, final_four_config, teams_out=None):
    """
    Build Final Four bracket with two semifinal games and one championship game.

    regional_winners: dict mapping region name to winning team (or TBD/None)
    final_four_config: dict with semifinal_1_region_1, semifinal_1_region_2, etc.
    teams_out: dict mapping team name to round lost
    """
    teams_out = teams_out or {}

    # Extract the configured region pairings
    sf1_r1 = final_four_config.get("semifinal_1_region_1", "EAST").upper()
    sf1_r2 = final_four_config.get("semifinal_1_region_2", "WEST").upper()
    sf2_r1 = final_four_config.get("semifinal_2_region_1", "MIDWEST").upper()
    sf2_r2 = final_four_config.get("semifinal_2_region_2", "SOUTH").upper()

    # Get regional winners
    team_sf1_r1 = regional_winners.get(sf1_r1, "TBD")
    team_sf1_r2 = regional_winners.get(sf1_r2, "TBD")
    team_sf2_r1 = regional_winners.get(sf2_r1, "TBD")
    team_sf2_r2 = regional_winners.get(sf2_r2, "TBD")

    # Ensure TBD is shown if None or empty
    if not team_sf1_r1 or "TBD" in str(team_sf1_r1):
        team_sf1_r1 = "TBD"
    if not team_sf1_r2 or "TBD" in str(team_sf1_r2):
        team_sf1_r2 = "TBD"
    if not team_sf2_r1 or "TBD" in str(team_sf2_r1):
        team_sf2_r1 = "TBD"
    if not team_sf2_r2 or "TBD" in str(team_sf2_r2):
        team_sf2_r2 = "TBD"

    rounds = []

    def game_winner(a, b, round_idx):
        if "TBD" in str(a) or "TBD" in str(b):
            return None
        a_lost = lost_in_round(a, round_idx, teams_out)
        b_lost = lost_in_round(b, round_idx, teams_out)
        if a_lost and not b_lost:
            return b
        if b_lost and not a_lost:
            return a
        return None

    # Semifinal Round
    semifinal_round = {
        "name": "Final Four - Semifinals",
        "regions": {
            "FINAL_FOUR": [
                {
                    "team_a": team_sf1_r1,
                    "seed_a": lookup_seed(team_sf1_r1, f"{sf1_r1[0]}"),
                    "team_b": team_sf1_r2,
                    "seed_b": lookup_seed(team_sf1_r2, f"{sf1_r2[0]}"),
                    "winner": game_winner(team_sf1_r1, team_sf1_r2, 4),
                    "team_a_eliminated": lost_in_round(team_sf1_r1, 4, teams_out),
                    "team_b_eliminated": lost_in_round(team_sf1_r2, 4, teams_out),
                },
                {
                    "team_a": team_sf2_r1,
                    "seed_a": lookup_seed(team_sf2_r1, f"{sf2_r1[0]}"),
                    "team_b": team_sf2_r2,
                    "seed_b": lookup_seed(team_sf2_r2, f"{sf2_r2[0]}"),
                    "winner": game_winner(team_sf2_r1, team_sf2_r2, 4),
                    "team_a_eliminated": lost_in_round(team_sf2_r1, 4, teams_out),
                    "team_b_eliminated": lost_in_round(team_sf2_r2, 4, teams_out),
                }
            ]
        }
    }
    rounds.append(semifinal_round)

    # Championship Round
    champ_a = game_winner(team_sf1_r1, team_sf1_r2, 4) or "TBD"
    champ_b = game_winner(team_sf2_r1, team_sf2_r2, 4) or "TBD"

    championship_round = {
        "name": "Championship",
        "regions": {
            "FINAL_FOUR": [
                {
                    "team_a": champ_a,
                    "seed_a": lookup_seed(champ_a, "SF1"),
                    "team_b": champ_b,
                    "seed_b": lookup_seed(champ_b, "SF2"),
                    "winner": game_winner(champ_a, champ_b, 5),
                    "team_a_eliminated": lost_in_round(champ_a, 5, teams_out),
                    "team_b_eliminated": lost_in_round(champ_b, 5, teams_out),
                }
            ]
        }
    }
    rounds.append(championship_round)

    return rounds


def render_regional_winner_card_html(region, winner, label=None, dark=False):
    """Render a winner card styled like the bracket game cards."""
    label = label or f"{region} REGIONAL WINNER"
    dark_css = (
        ".matchup-container {border-color: #374151; background-color: #1A2130;}"
        ".game-label {background: #1F2937; color: #9CA3AF;}"
        ".team-box {background: #1A2130;}"
        ".team-name {color: #E5E7EB;}"
        ".team-box.tbd, .team-box.tbd .team-seed, .team-box.tbd .team-name {color: #6B7280;}"
        if dark else ""
    )
    known = bool(winner) and "TBD" not in str(winner)
    seed = lookup_seed(winner, "") if known else ""
    name = str(winner) if known else "TBD"
    logo = _logo_img(winner) if known else '<span class="team-logo placeholder"></span>'
    tbd_class = "" if known else "tbd"
    winner_class = "winner" if known else ""

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                margin: 0;
                padding: 0 4px;
                height: 100vh;
                display: flex;
                align-items: center;
                justify-content: center;
                box-sizing: border-box;
                background: transparent;
            }}
            .matchup-container {{
                width: 100%;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                overflow: hidden;
                box-shadow: 0 1px 3px rgba(0,0,0,0.08);
                background-color: #f8f9fa;
            }}
            .game-label {{
                font-size: 10px;
                color: #888;
                background: #eef0f3;
                padding: 1px 6px;
                letter-spacing: 0.5px;
            }}
            .team-box {{
                background: #f8f9fa;
                padding: 10px 14px;
                font-size: 18px;
                display: flex;
                align-items: center;
                height: 52px;
                min-height: 52px;
            }}
            .team-info {{ display: flex; align-items: center; gap: 10px; flex: 1; }}
            .team-seed {{ font-weight: 700; color: #4472C4; min-width: 24px; font-size: 18px; }}
            .team-logo {{ width: 38px; height: 38px; object-fit: contain; flex-shrink: 0; }}
            .team-logo.placeholder {{ display: inline-block; }}
            .team-name {{ flex: 1; font-size: 18px; color: #333; }}
            .team-box.tbd {{ color: #999; font-style: italic; }}
            .team-box.tbd .team-seed, .team-box.tbd .team-name {{ color: #999; }}
            .team-box.winner .team-name {{ font-weight: 700; }}
            {dark_css}
        </style>
    </head>
    <body>
        <div class="matchup-container">
            <div class="game-label">{label}</div>
            <div class="team-box {tbd_class} {winner_class}">
                <div class="team-info">
                    <span class="team-seed">{seed}</span>
                    {logo}
                    <span class="team-name">{name}</span>
                </div>
            </div>
        </div>
    </body>
    </html>
    """


def build_team_players(data):
    """Return {normalized team name: [{n: player, o: owner}]} for drafted players only."""
    result = {}
    for row in data["draft_rows"]:
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        if not team:
            continue
        result.setdefault(_norm_team_name(team), []).append({
            "n": player_label(row["player_name"], stats),
            "i": get_player_photo_url(row["player_name"], team),
            "o": row["draftee"],
            "s": [
                list(row.get(key) or [None] * 6)
                for key in ("tourn_pts_rounds", "tourn_reb_rounds", "tourn_ast_rounds")
            ],
            "f": [
                weighted_fantasy_points(
                    (row.get("tourn_pts_rounds") or [None] * 6)[i],
                    (row.get("tourn_reb_rounds") or [None] * 6)[i],
                    (row.get("tourn_ast_rounds") or [None] * 6)[i],
                    data["fantasy_weights"],
                )
                for i in range(6)
            ],
        })
    return result


def render_bracket_html(rounds, region=None, team_players=None, my_team=None, dark=False):
    team_players = team_players or {}
    regions_list = ["EAST", "WEST", "MIDWEST", "SOUTH", "FINAL_FOUR"]

    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                margin: 0;
                padding: 10px;
                background: transparent;
            }

            .ncaa-bracket-wrapper {
                position: relative;
                overflow-x: auto;
                background: white;
                border-radius: 4px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            }

            .ncaa-bracket {
                display: flex;
                gap: 1px;
                padding: 10px;
                min-width: max-content;
            }

            .bracket-column {
                display: flex;
                flex-direction: column;
                align-items: center;
                min-width: 180px;
                padding: 0 15px;
                position: relative;
                flex: 1;
                min-height: 500px;
            }

            .column-matchups {
                flex: 1;
                width: 100%;
                display: flex;
                flex-direction: column;
                justify-content: space-around;
                align-items: center;
            }

            .column-matchups .matchup-container {
                margin: 6px 0;
            }

            .game-label {
                font-size: 10px;
                color: #888;
                background: #eef0f3;
                padding: 1px 6px;
                letter-spacing: 0.5px;
            }

            .bracket-column:last-child {
                border-right: none;
            }

            .bracket-column-title {
                text-align: center;
                font-size: 12px;
                font-weight: 600;
                color: #555;
                margin-bottom: 20px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                line-height: 1.4;
            }

            .matchup-container {
                position: relative;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                overflow: hidden;
                box-shadow: 0 1px 3px rgba(0,0,0,0.08);
                flex-shrink: 0;
                width: 100%;
                min-width: 160px;
                background-color: #f8f9fa;
            }

            .team-box {
                background: #f8f9fa;
                padding: 10px 12px;
                font-size: 13px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid #e8e8e8;
                transition: all 0.2s ease;
                height: 32px;
                min-height: 32px;
            }

            .team-box:last-child {
                border-bottom: none;
            }

            .team-box:hover {
                background: #eef2f7;
            }

            .matchup-container:hover {
                border-color: #d0d0d0;
                box-shadow: 0 2px 6px rgba(68, 114, 196, 0.15);
            }

            .team-info {
                display: flex;
                align-items: center;
                gap: 8px;
                flex: 1;
                min-width: 0;
            }

            .team-seed {
                font-weight: 700;
                color: #4472C4;
                min-width: 20px;
                font-size: 12px;
            }

            .team-logo {
                width: 22px;
                height: 22px;
                object-fit: contain;
                flex-shrink: 0;
                transform: translateZ(0);
            }

            .team-logo.placeholder {
                display: inline-block;
            }

            .team-box.eliminated .team-logo {
                filter: grayscale(100%);
                opacity: 0.5;
            }

            .team-name {
                flex: 1;
                min-width: 0;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                font-size: 13px;
                color: #333;
            }

            .team-score {
                font-weight: 600;
                color: #666;
                margin-left: 8px;
                min-width: 30px;
                flex-shrink: 0;
                text-align: right;
            }

            .team-box.tbd {
                background: #f9f9f9;
                color: #999;
                font-style: italic;
            }

            .team-score.score-win { font-weight: 800; color: #222; }
            .team-score.score-loss { color: #aaa; font-weight: 500; }
            .team-score.win-prob { font-size: 12px; font-weight: 500; color: #4472C4; }
            .team-score.win-prob.fav { font-weight: 700; }

            .team-box.tbd .team-seed {
                color: #ccc;
            }

            .team-box.tbd .team-name {
                color: #999;
            }

            .team-box.tbd .team-score {
                color: #ccc;
            }

            .team-box.eliminated .team-name {
                color: #999;
                font-weight: 400;
            }

            .team-box.eliminated .team-seed {
                color: #999;
                font-weight: 700;
            }

            .team-box.winner .team-name {
                font-weight: 700;
            }

            .team-box.winner .team-seed {
                font-weight: 700;
            }

            .connector-line {
                position: absolute;
                height: 2px;
                background: #d0d0d0;
                top: 50%;
                transform: translateY(-50%);
                pointer-events: none;
            }

            .vertical-connector {
                position: absolute;
                width: 2px;
                background: #d0d0d0;
                left: 100%;
                pointer-events: none;
            }
        </style>
    </head>
    <body>
        <style>
            .matchup-container { cursor: pointer; }
            #player-modal { display: none; position: fixed; inset: 0; z-index: 1000;
                background: rgba(0,0,0,0.35); align-items: center; justify-content: center; }
            #player-modal .box { background: #fff; border-radius: 10px; padding: 16px 20px;
                min-width: min(340px, 80%); max-width: 80%;
                max-height: calc(100% - 16px); overflow-y: auto; box-sizing: border-box;
                box-shadow: 0 8px 24px rgba(0,0,0,0.25); font-family: sans-serif; font-size: 13px; }
            #player-modal h3 { margin: 0 0 8px; font-size: 15px; white-space: nowrap; }
            #player-modal h4 { margin: 12px 0 4px; font-size: 13px; color: #4472C4; }
            #player-modal .odds { font-size: 12px; color: #4472C4; margin-bottom: 2px; }
            #player-modal .tv { font-size: 12px; color: #555; margin-bottom: 6px; line-height: 1.4; }
            #player-modal .row { padding: 0; height: 60px; overflow: hidden; border-bottom: 1px solid #eee; display: flex; align-items: flex-end; gap: 8px; }
            #player-modal .pp { width: 76px; height: 76px; object-fit: contain; object-position: center bottom; flex-shrink: 0; display: block; }
            #player-modal .info { flex: 1; min-width: 0; align-self: center; }
            #player-modal .owner { color: #888; }
            #player-modal .owner.me { color: #B85C00; font-weight: 800; background: #FFE0B2; border-radius: 4px; padding: 0 4px; }
            #player-modal .row.me { background: #FFF4E5; }
            #player-modal .none { color: #999; font-style: italic; }
            #player-modal .stats { color: #555; font-size: 12px; }
            #player-modal .close { float: right; cursor: pointer; font-size: 18px; color: #888; }
            @media (max-width: 480px) { #player-modal h3 { white-space: normal; } }
        </style>
        <div id="player-modal"><div class="box"></div></div>
        <div class="ncaa-bracket-wrapper">
            <svg class="bracket-connectors" style="position: absolute; top: 0; left: 0; pointer-events: none;"></svg>
            <div class="ncaa-bracket">
    """

    # If region specified, only render that region
    regions_to_render = [region] if region else regions_list

    game_counter = 1

    # Flatten all matchups by round across specified regions
    for round_idx, round_data in enumerate(rounds):
        round_name = round_data['name']

        # Get all matchups for this round across specified regions
        all_matchups = []
        matchup_regions = []
        for r in regions_to_render:
            if r in round_data['regions']:
                all_matchups.extend(round_data['regions'][r])
                matchup_regions.extend([r] * len(round_data['regions'][r]))

        # Hide the First Four column when the region has no First Four games
        if round_name == "First Four" and not all_matchups:
            continue

        display_name = "Opening Round" if round_name == "First Four" else round_name

        html += f"""
                <div class="bracket-column">
                    <div class="bracket-column-title">{display_name}</div>
                    <div class="column-matchups">
        """

        # Display matchups
        for matchup_idx, matchup in enumerate(all_matchups):
            team_a = matchup['team_a']
            team_b = matchup['team_b']
            seed_a = matchup['seed_a']
            seed_b = matchup['seed_b']
            winner = matchup.get('winner')

            is_tbd_a = "TBD" in str(team_a)
            is_tbd_b = "TBD" in str(team_b)
            is_eliminated_a = matchup.get('team_a_eliminated', False)
            is_eliminated_b = matchup.get('team_b_eliminated', False)
            is_winner_a = winner == team_a
            is_winner_b = winner == team_b

            tbd_class_a = "tbd" if is_tbd_a else ""
            eliminated_class_a = "eliminated" if is_eliminated_a else ""
            winner_class_a = "winner" if is_winner_a else ""
            classes_a = f"{tbd_class_a} {eliminated_class_a} {winner_class_a}".strip()

            tbd_class_b = "tbd" if is_tbd_b else ""
            eliminated_class_b = "eliminated" if is_eliminated_b else ""
            winner_class_b = "winner" if is_winner_b else ""
            classes_b = f"{tbd_class_b} {eliminated_class_b} {winner_class_b}".strip()

            game_num = game_counter
            game_counter += 1
            # Only show a score once the loser has been entered on TeamsOut.
            game_scores = (
                find_game_scores(team_a, team_b)
                if is_eliminated_a or is_eliminated_b else None
            )
            odds_badge = ""
            win_pct = {}
            spread = {}
            if game_scores:
                score_html_a = _score_html(game_scores[0], game_scores[1])
                score_html_b = _score_html(game_scores[1], game_scores[0])
            else:
                score_html_a = score_html_b = _score_html(None, None)
                # Unplayed game between two known teams: show each team's chance to win.
                if not (is_tbd_a or is_tbd_b or winner or is_eliminated_a or is_eliminated_b):
                    p_a = matchup_win_probability(team_a, seed_a, team_b, seed_b)
                    if p_a is not None:
                        percent_a = round(p_a * 100)
                        score_html_a = _win_prob_html(percent_a, percent_a >= 50)
                        score_html_b = _win_prob_html(100 - percent_a, percent_a < 50)
                        win_pct = {team_a: percent_a, team_b: 100 - percent_a}
                        margin = matchup_margin(team_a, seed_a, team_b, seed_b)
                        # Betting-line style: the favorite carries the negative number.
                        spread = {team_a: round(-margin * 2) / 2 + 0.0, team_b: round(margin * 2) / 2 + 0.0}
                        odds_badge = (
                            '<span class="torvik-badge" '
                            'title="Win probabilities from Barttorvik ratings (barttorvik.com)"></span>'
                        )
            feeds_to = matchup_idx // 2 if len(all_matchups) > 1 and round_name != "First Four" else -1

            align_to = -1
            if round_name == "First Four" and len(rounds) > 1:
                offset = 0
                for r in regions_to_render:
                    r64 = rounds[1]['regions'].get(r, [])
                    if r == matchup_regions[matchup_idx]:
                        for j, m in enumerate(r64):
                            if str(seed_a) in (str(m['seed_a']), str(m['seed_b'])):
                                align_to = offset + j
                                break
                        break
                    offset += len(r64)

            players_attr = html_lib.escape(json.dumps({
                "title": f"G{game_num} - {display_name}",
                "ri": {
                    "Round of 64": 0, "Round of 32": 1, "Sweet Sixteen": 2,
                    "Elite Eight": 3, "Final Four - Semifinals": 4, "Championship": 5,
                }.get(round_name, -1),
                "teams": [
                    {
                        "team": str(t),
                        "logo": get_logo_url(t),
                        "players": team_players.get(_norm_team_name(t), []),
                        "pct": win_pct.get(t),
                        "spread": spread.get(t),
                        "tv": torvik_team_info(t),
                    }
                    for t in (team_a, team_b) if "TBD" not in str(t)
                ],
            }), quote=True)

            html += f"""
                    <div class="matchup-container" data-round="{round_idx}" data-index="{matchup_idx}" data-feeds-to="{feeds_to}" data-align-to="{align_to}" data-players="{players_attr}">
                        <div class="game-label">G{game_num}{odds_badge}</div>
                        <div class="team-box {classes_a}">
                            <div class="team-info">
                                <span class="team-seed">{seed_a}</span>
                                {_logo_img(team_a)}
                                <span class="team-name">{str(team_a)[:28]}</span>
                            </div>
                            {score_html_a}
                        </div>
                        <div class="team-box {classes_b}">
                            <div class="team-info">
                                <span class="team-seed">{seed_b}</span>
                                {_logo_img(team_b)}
                                <span class="team-name">{str(team_b)[:28]}</span>
                            </div>
                            {score_html_b}
                        </div>
                    </div>
            """


        html += """
                    </div>
                </div>
        """

    html += """
            </div>
        </div>
        <script>
            function drawConnectors() {
                const svg = document.querySelector('.bracket-connectors');
                const wrapper = document.querySelector('.ncaa-bracket-wrapper');
                const columns = document.querySelectorAll('.bracket-column');

                if (columns.length < 2 || !wrapper) return;

                // Clear existing lines
                svg.innerHTML = '';

                // Align opening-round games with the Round of 64 game they feed
                const alignable = Array.from(document.querySelectorAll('.matchup-container[data-align-to]'))
                    .filter(m => m.dataset.alignTo !== '-1');
                alignable.forEach(m => { m.style.transform = ''; });
                alignable.forEach(m => {
                    const target = Array.from(document.querySelectorAll('.bracket-column')[1].querySelectorAll('.matchup-container'))[Number(m.dataset.alignTo)];
                    if (!target) return;
                    const a = m.getBoundingClientRect();
                    const b = target.getBoundingClientRect();
                    const dy = (b.top + b.height / 2) - (a.top + a.height / 2);
                    m.style.transform = 'translateY(' + dy + 'px)';
                });

                // Set SVG size to match wrapper
                svg.style.width = '0px';
                svg.style.height = '0px';
                svg.style.width = wrapper.scrollWidth + 'px';
                svg.style.height = wrapper.scrollHeight + 'px';
                svg.setAttribute('width', wrapper.scrollWidth);
                svg.setAttribute('height', wrapper.scrollHeight);

                // Function to draw connector lines between two columns
                function connectColumns(colA, colB) {
                    const matchupsA = colA.querySelectorAll('.matchup-container');
                    const matchupsB = colB.querySelectorAll('.matchup-container');

                    // Connect each pair of feeder games to the game they feed (by data-feeds-to)
                    for (let i = 0; i < matchupsB.length; i++) {
                        const feeders = Array.from(matchupsA).filter(m => m.dataset.feedsTo === String(i));
                        const mAa = feeders[0];
                        const mAb = feeders[1];
                        const mB = matchupsB[i];

                        if (mAa && mAb && mB) {
                            // Get positions relative to wrapper
                            const rectAa = mAa.getBoundingClientRect();
                            const rectAb = mAb.getBoundingClientRect();
                            const rectB = mB.getBoundingClientRect();
                            const wrapperRect = wrapper.getBoundingClientRect();

                            const offX = wrapper.scrollLeft - wrapperRect.left;
                            const offY = wrapper.scrollTop - wrapperRect.top;

                            const yAa = rectAa.top + offY + rectAa.height / 2;
                            const yAb = rectAb.top + offY + rectAb.height / 2;
                            const yB = rectB.top + offY + rectB.height / 2;
                            const xA = rectAa.right + offX;
                            const xB = rectB.left + offX;

                            // Midpoint x coordinate
                            const xMid = xA + (xB - xA) * 0.5;

                            // Line from first game right edge to midpoint
                            const line1 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
                            line1.setAttribute('x1', xA);
                            line1.setAttribute('y1', yAa);
                            line1.setAttribute('x2', xMid);
                            line1.setAttribute('y2', yAa);
                            line1.setAttribute('stroke', '#d0d0d0');
                            line1.setAttribute('stroke-width', '2');
                            svg.appendChild(line1);

                            // Line from second game right edge to midpoint
                            const line2 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
                            line2.setAttribute('x1', xA);
                            line2.setAttribute('y1', yAb);
                            line2.setAttribute('x2', xMid);
                            line2.setAttribute('y2', yAb);
                            line2.setAttribute('stroke', '#d0d0d0');
                            line2.setAttribute('stroke-width', '2');
                            svg.appendChild(line2);

                            // Vertical line from first game to winner position
                            const line3 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
                            line3.setAttribute('x1', xMid);
                            line3.setAttribute('y1', yAa);
                            line3.setAttribute('x2', xMid);
                            line3.setAttribute('y2', yAb);
                            line3.setAttribute('stroke', '#d0d0d0');
                            line3.setAttribute('stroke-width', '2');
                            svg.appendChild(line3);

                            // Line from midpoint to winner box
                            const line4 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
                            line4.setAttribute('x1', xMid);
                            line4.setAttribute('y1', yB);
                            line4.setAttribute('x2', xB);
                            line4.setAttribute('y2', yB);
                            line4.setAttribute('stroke', '#d0d0d0');
                            line4.setAttribute('stroke-width', '2');
                            svg.appendChild(line4);
                        }
                    }
                }

                // Connect every pair of adjacent rounds (handles First Four and later rounds)
                for (let c = 0; c < columns.length - 1; c++) {
                    connectColumns(columns[c], columns[c + 1]);
                }
            }

            window.addEventListener('load', () => {
                setTimeout(drawConnectors, 100);
            });

            function refreshLogos() {
                document.querySelectorAll('img.team-logo').forEach(img => {
                    const tries = Number(img.dataset.t || 0);
                    if ((!img.complete || img.naturalWidth === 0) && tries < 3) {
                        img.dataset.t = tries + 1;
                        const base = img.src.split('?')[0];
                        img.src = base + '?t=' + Date.now();
                    }
                });
            }
            let logoTimer = null;
            function queueLogoRefresh() {
                clearTimeout(logoTimer);
                logoTimer = setTimeout(refreshLogos, 200);
            }
            window.addEventListener('load', () => setTimeout(refreshLogos, 1500));
            window.addEventListener('scroll', queueLogoRefresh, true);
            window.addEventListener('touchend', queueLogoRefresh, true);

            const modal = document.getElementById('player-modal');
            const modalBox = modal.querySelector('.box');
            const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
            const MY_TEAM = __MY_TEAM__;
            const isMine = o => MY_TEAM && String(o).toLowerCase() === MY_TEAM.toLowerCase();
            document.querySelectorAll('.matchup-container').forEach(el => {
                el.addEventListener('click', () => {
                    let info;
                    try { info = JSON.parse(el.dataset.players); } catch (e) { return; }
                    if (!info.teams.length) return;
                    let h = '<span class="close" id="modal-close">&times;</span><h3>Drafted players in ' + esc(info.title) + '</h3>';
                    info.teams.forEach(t => {
                        const logo = t.logo ? '<img src="' + esc(t.logo) + '" style="width:20px;height:20px;object-fit:contain;vertical-align:middle;margin-right:6px" onerror="this.remove()">' : '';
                        h += '<h4>' + logo + esc(t.team) + '</h4>';
                        if (t.pct !== null && t.pct !== undefined) h += '<div class="odds">Win probability: <b>' + t.pct + '%</b></div>';
                        if (t.spread !== null && t.spread !== undefined) h += '<div class="odds">Predicted spread: <b>' + (t.spread > 0 ? '+' : '') + t.spread.toFixed(1) + '</b></div>';
                        if (t.tv) {
                            const net = t.tv.net;
                            h += '<div class="tv">Torvik' + (t.tv.rank ? ' #' + t.tv.rank : '')
                              + (t.tv.record ? ' &middot; ' + esc(t.tv.record) : '')
                              + ' &middot; Adj. O ' + t.tv.o + ' &middot; Adj. D ' + t.tv.d
                              + ' &middot; Net ' + (net > 0 ? '+' : '') + net + ' &middot; Tempo ' + t.tv.t + '</div>';
                        }
                        if (!t.players.length) h += '<div class="none">No drafted players</div>';
                        const fp = p => (info.ri >= 0 && p.f && typeof p.f[info.ri] === 'number') ? p.f[info.ri] : -1;
                        t.players.slice().sort((a, b) => fp(b) - fp(a)).forEach(p => {
                            const val = k => (info.ri >= 0 && p.s[k][info.ri] !== null && p.s[k][info.ri] !== undefined && p.s[k][info.ri] !== '') ? p.s[k][info.ri] : '-';
                            const photo = p.i ? '<img class="pp" src="' + esc(p.i) + '" referrerpolicy="no-referrer">' : '<span class="pp"></span>';
                            const mine = isMine(p.o);
                            h += '<div class="row' + (mine ? ' me' : '') + '">' + photo + '<div class="info">' + esc(p.n) + ' <span class="owner' + (mine ? ' me' : '') + '">&mdash; ' + esc(p.o) + '</span>'
                              + (info.ri >= 0 ? '<div class="stats"><b>FPts ' + (fp(p) >= 0 ? fp(p).toFixed(1) : '-') + '</b> &middot; PTS ' + esc(val(0)) + ' &middot; REB ' + esc(val(1)) + ' &middot; AST ' + esc(val(2)) + '</div>' : '') + '</div></div>';
                        });
                    });
                    modalBox.innerHTML = h;
                    modalBox.querySelectorAll('img.pp').forEach(im => im.addEventListener('error', () => { im.style.visibility = 'hidden'; }));
                    modal.style.display = 'flex';
                });
            });
            modal.addEventListener('click', e => {
                if (e.target === modal || e.target.id === 'modal-close') modal.style.display = 'none';
            });
            window.addEventListener('resize', drawConnectors);
            if (window.ResizeObserver) {
                new ResizeObserver(() => drawConnectors()).observe(document.body);
            }
        </script>
    </body>
    </html>
    """

    html = html.replace("__MY_TEAM__", json.dumps(my_team or ""))
    if 'class="torvik-badge"' in html:
        html = html.replace("</head>", _torvik_badge_css() + "</head>", 1)
    if dark:
        dark_css = """<style>
            .ncaa-bracket-wrapper {background: #141A25; box-shadow: none;}
            .game-label {background: #1F2937; color: #9CA3AF;}
            .bracket-column-title {color: #9CA3AF;}
            .matchup-container {border-color: #374151; background-color: #1A2130;}
            .team-box {background: #1A2130; border-bottom-color: #2D3748;}
            .team-box:hover {background: #243044;}
            .team-name {color: #E5E7EB;}
            .team-score {color: #9CA3AF;}
            .team-score.score-win {color: #F3F4F6;}
            .team-score.score-loss {color: #6B7280;}
            .team-score.win-prob {color: #8DB8E8;}
            .torvik-badge {filter: brightness(0) invert(0.85);}
            .team-box.tbd {background: #161C28; color: #6B7280;}
            .team-box.tbd .team-seed, .team-box.tbd .team-score, .team-box.tbd .team-name,
            .team-box.eliminated .team-name, .team-box.eliminated .team-seed {color: #6B7280;}
            .connector-line, .vertical-connector {background: #4B5563;}
            #player-modal .box {background: #1A2130; color: #E5E7EB;}
            #player-modal h3 {color: #E5E7EB;}
            #player-modal .row {border-bottom-color: #2D3748;}
            #player-modal .row.me {background: #3B2E1A;}
            #player-modal .owner, #player-modal .stats, #player-modal .close, #player-modal .tv {color: #9CA3AF;}
            #player-modal .odds, #player-modal h4 {color: #8DB8E8;}
        </style>"""
        html = html.replace('<div id="player-modal">', dark_css + '<div id="player-modal">', 1)
        html = html.replace("'#d0d0d0'", "'#4B5563'")
    return html


def load_data(file_path, updated_time=None):
    workbook = openpyxl.load_workbook(
        file_path,
        read_only=True,
        data_only=True,
        keep_vba=Path(file_path).suffix.lower() == ".xlsm",
        keep_links=False,
    )

    required_sheets = [DRAFT_ORDER_SHEET, DRAFT_SHEET, TEAM_DATA_SHEET, CONFIG_SHEET]
    missing = [s for s in required_sheets if s not in workbook.sheetnames]
    if missing:
        raise ValueError("Missing required sheet(s): " + ", ".join(missing))

    # Fantasy scoring weights. Config!A2:C2 correspond to PTS, REB, AST.
    config_ws = workbook[CONFIG_SHEET]
    fantasy_weights = []
    for cell in ("A2", "B2", "C2"):
        value = config_ws[cell].value
        try:
            fantasy_weights.append(float(value))
        except (TypeError, ValueError):
            raise ValueError(
                f"Config!{cell} must contain a numeric fantasy-point value."
            )

    # Payout settings. Config!D2:G2 correspond to base payout,
    # last-place multiplier, third-place multiplier, and second-place multiplier.
    # Historical average tournament games by seed.
    # Config!I2:I17 correspond to seeds 1 through 16.
    seed_average_games = {}
    for seed_number, row_number in enumerate(range(2, 18), start=1):
        value = config_ws[f"I{row_number}"].value
        try:
            seed_average_games[seed_number] = float(value)
        except (TypeError, ValueError):
            raise ValueError(
                f"Config!I{row_number} must contain a numeric historical average games value."
            )

    # Config!J2 is the tournament FPts weighting.
    try:
        tournament_fpts_weight = float(config_ws["J2"].value)
    except (TypeError, ValueError):
        raise ValueError("Config!J2 must contain a numeric tournament FPts weight.")

    payout_config = {}
    for cell in ("D2", "E2", "F2", "G2"):
        value = config_ws[cell].value
        try:
            payout_config[cell] = float(value)
        except (TypeError, ValueError):
            raise ValueError(f"Config!{cell} must contain a numeric payout value.")

    final_four_config = read_final_four_config(config_ws)
    drafters = []
    seen = set()
    ws = workbook[DRAFT_ORDER_SHEET]
    for values in ws.iter_rows(
        min_row=DRAFTER_START_ROW, min_col=1, max_col=1, values_only=True
    ):
        value = values[0]
        if value is None:
            continue
        draftee = str(value).strip()
        key = draftee.casefold()
        if draftee and key not in seen:
            drafters.append(draftee)
            seen.add(key)

    # Draft rows
    ws = workbook[DRAFT_SHEET]
    draft_rows = []
    for row_number, values in enumerate(
        ws.iter_rows(min_row=2, min_col=1, max_col=21, values_only=True),
        start=2,
    ):
        draft_position, draftee, player_name = values[:3]
        if draftee is None or player_name is None:
            continue

        tournament_values = values[3:21]
        _, games_not_input = count_game_statuses(tournament_values)
        games_played = count_values_not_out(values[3:9])

        draft_rows.append({
            "draft_row": row_number,
            "draft_position": draft_position,
            "draftee": str(draftee).strip(),
            "player_name": str(player_name).strip(),
            "games_not_input": games_not_input,
            "possible_games_remaining": count_blank_values(values[3:9]),
            "games_played": games_played,
            "tourn_pts": sum_numeric_values(tournament_values[0:6]),
            "tourn_reb": sum_numeric_values(tournament_values[6:12]),
            "tourn_ast": sum_numeric_values(tournament_values[12:18]),
            "tourn_pts_rounds": tournament_values[0:6],
            "tourn_reb_rounds": tournament_values[6:12],
            "tourn_ast_rounds": tournament_values[12:18],
        })

    # Player/team data
    ws = workbook[TEAM_DATA_SHEET]
    header_values = next(
        ws.iter_rows(min_row=1, max_row=1, values_only=True), ()
    )
    headers = {
        str(value).strip(): index
        for index, value in enumerate(header_values)
        if value is not None
    }

    required_headers = ["Player", "Team", "Seed", "PPG", "RPG", "APG"]
    missing = [h for h in required_headers if h not in headers]
    if missing:
        raise ValueError("Missing TeamData header(s): " + ", ".join(missing))

    status_index = 16  # TeamData column Q
    max_column = max(max(headers.values()) + 1, status_index + 1)
    player_stats = {}
    player_names = []

    for values in ws.iter_rows(min_row=2, max_col=max_column, values_only=True):
        def get_value(header):
            index = headers[header]
            return values[index] if index < len(values) else None

        player_name = get_value("Player")
        if player_name is None:
            continue

        display_name = str(player_name).strip()
        if not display_name:
            continue

        player_key = display_name.casefold()
        if player_key not in player_stats:
            player_names.append(display_name)

        ppg = get_value("PPG")
        rpg = get_value("RPG")
        apg = get_value("APG")
        # Season FPts are calculated from the season PPG/RPG/APG values
        # using the configurable PTS/REB/AST weights in Config!A2:C2.
        total_fantasy_points = weighted_fantasy_points(
            ppg, rpg, apg, fantasy_weights
        )

        player_stats[player_key] = {
            "Team": get_value("Team"),
            "Seed": get_value("Seed"),
            "Region": get_value("Region") if "Region" in headers else None,
            "Tot. FPts": total_fantasy_points,
            "PPG": ppg,
            "RPG": rpg,
            "APG": apg,
            "Status": (
                str(values[status_index]).strip()
                if status_index < len(values) and values[status_index] is not None
                else ""
            ),
        }

    player_names.sort(key=str.casefold)

    # Eliminated teams
    teams_out = {}  # {team_name_lower: round_lost}
    if TEAMS_OUT_SHEET in workbook.sheetnames:
        ws = workbook[TEAMS_OUT_SHEET]
        for values in ws.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
            team_name = values[0]
            round_lost = values[1] if len(values) > 1 else None
            if team_name is not None and str(team_name).strip():
                team_lower = str(team_name).strip().casefold()
                round_text = str(round_lost).strip() if round_lost else None
                teams_out[team_lower] = round_text

    # Workbook last-edited timestamp is maintained by the local stats updater
    # in GameResults!D2. This is the authoritative timestamp displayed by
    # the dashboard; it does not depend on OneDrive's file modified time.
    workbook_updated_time = None
    if GAME_RESULTS_SHEET not in workbook.sheetnames:
        raise ValueError(f"Missing required sheet: {GAME_RESULTS_SHEET}")

    ws = workbook[GAME_RESULTS_SHEET]
    raw_updated_time = ws["D2"].value
    # A brand-new/pre-draft workbook may not have a stats-updater timestamp yet.
    # Treat a blank D2 as "not updated yet" instead of rejecting the workbook.
    if raw_updated_time is None or str(raw_updated_time).strip() == "":
        workbook_updated_time = None
    elif isinstance(raw_updated_time, datetime):
        workbook_updated_time = raw_updated_time
    else:
        raw_updated_text = str(raw_updated_time).strip()
        try:
            workbook_updated_time = datetime.fromisoformat(
                raw_updated_text.replace("Z", "+00:00")
            )
        except ValueError:
            for date_format in (
                "%m/%d/%Y %I:%M %p",
                "%m/%d/%Y %H:%M",
                "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M",
            ):
                try:
                    workbook_updated_time = datetime.strptime(raw_updated_text, date_format)
                    break
                except ValueError:
                    continue

        if workbook_updated_time is None:
            raise ValueError(
                "GameResults!D2 does not contain a recognizable date/time. "
                "Expected an Excel date/time value or a standard date/time string."
            )

    # A blank D2 is valid for a pre-draft workbook, so only normalize
    # timezone information when a timestamp was actually supplied.
    if workbook_updated_time is not None and workbook_updated_time.tzinfo is not None:
        workbook_updated_time = workbook_updated_time.astimezone().replace(tzinfo=None)

    # Most recent game result
    latest_game_result = {"Round": "", "Winner": "", "Loser": ""}
    if GAME_RESULTS_SHEET in workbook.sheetnames:
        ws = workbook[GAME_RESULTS_SHEET]
        headers2 = {}
        for column in range(1, ws.max_column + 1):
            value = ws.cell(row=1, column=column).value
            if value is not None:
                headers2[str(value).strip()] = column

        required = ["Round", "Winner", "Loser"]
        if all(h in headers2 for h in required):
            for row in range(ws.max_row, 1, -1):
                vals = {
                    h: ws.cell(row=row, column=headers2[h]).value
                    for h in required
                }
                if any(v not in (None, "") for v in vals.values()):
                    latest_game_result = {
                        h: "" if vals[h] is None else str(vals[h]).strip()
                        for h in required
                    }
                    break

    bracket_teams = (
        read_bracket_teams(workbook[TEAM_LIST_SHEET]) if TEAM_LIST_SHEET in workbook.sheetnames else {}
    )

    return {
        "fantasy_weights": fantasy_weights,
        "seed_average_games": seed_average_games,
        "tournament_fpts_weight": tournament_fpts_weight,

        "payout_config": payout_config,
        "final_four_config": final_four_config,
        "drafters": drafters,
        "draft_rows": draft_rows,
        "player_stats": player_stats,
        "player_names": player_names,
        "teams_out": teams_out,
        "latest_game_result": latest_game_result,
        "updated_time": workbook_updated_time,
        "bracket_teams": bracket_teams,
    }


def player_total(row, fantasy_weights):
    return weighted_fantasy_points(
        row.get("tourn_pts"),
        row.get("tourn_reb"),
        row.get("tourn_ast"),
        fantasy_weights,
    )


COVERS_INJURIES_URL = "https://www.covers.com/sport/basketball/ncaab/injuries"
COVERS_TEAM_ALIASES = {"uconn": "connecticut"}
_NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v"}


def _injury_name_key(name):
    parts = [re.sub(r"[^a-z0-9]", "", p) for p in str(name).casefold().split()]
    parts = [p for p in parts if p and p not in _NAME_SUFFIXES]
    if len(parts) < 2:
        return None
    return parts[0][0], parts[-1]


def _injury_team_key(team):
    key = _norm_team_name(team)
    return COVERS_TEAM_ALIASES.get(key, key)


@st.cache_data(ttl=10800, show_spinner=False)
def get_covers_injuries():
    """{(first initial, last name): [(team key, 'out'|'questionable', injury type, updated date, notes)]} from Covers.com."""
    injuries = {}
    try:
        resp = requests.get(COVERS_INJURIES_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        resp.raise_for_status()
        html = resp.text
    except Exception:
        return injuries

    def clean(text):
        return re.sub(r"\s+", " ", html_lib.unescape(re.sub(r"<[^>]+>", " ", text or ""))).strip()

    sections = re.split(r'<a id="([^"]+)"></a>\s*<section>', html)
    for team, body in zip(sections[1::2], sections[2::2]):
        for chunk in re.split(r"(?=<a [^>]*class='player-link')", body)[1:]:
            player = re.search(r"class='player-link'[^>]*>\s*([^<]+?)\s*</a>", chunk, re.S)
            status = re.search(r"<b>([^<]+)</b>(?:\s*<br>\s*\(\s*([^)]*?)\s*\))?", chunk, re.S)
            if not player or not status:
                continue
            notes = re.search(r'injuryCopy">\s*(.*?)\s*</div>', chunk, re.S)
            status_text = clean(status.group(1))
            lowered = status_text.casefold()
            if lowered.startswith("out"):
                level = "out"
            elif lowered.startswith(("questionable", "doubtful", "day")):
                level = "questionable"
            else:
                continue
            key = _injury_name_key(re.sub(r"\s+", " ", player.group(1)))
            if key:
                injury_type = status_text.split(" - ", 1)[1] if " - " in status_text else ""
                injuries.setdefault(key, []).append((
                    _injury_team_key(team), level, injury_type, clean(status.group(2)), clean(notes.group(1) if notes else ""),
                ))
    return injuries


def _covers_match(name, team):
    key = _injury_name_key(name)
    matches = get_covers_injuries().get(key, []) if key else []
    if len(matches) > 1:
        team_key = _injury_team_key(team or "")
        matches = [m for m in matches if m[0] == team_key or m[0] in team_key or team_key in m[0]]
    return matches[0] if len(matches) == 1 else None


def injury_status(name, stats):
    """Workbook Status wins; otherwise fall back to the Covers.com injury report."""
    status = str(stats.get("Status") or "").strip().casefold()
    if status:
        return status
    match = _covers_match(name, stats.get("Team"))
    return match[1] if match else ""


def injury_tip_html(name, team, marker):
    """Hover text for a player's out/questionable marker: injury type, report date and notes when known."""
    label = "Out" if marker == "⛔" else "Questionable"
    match = _covers_match(name, team)
    if not match:
        return f"<b>{label}</b><br>No injury details available."
    _team, _level, injury_type, updated, notes = match
    text = f"<b>{label}</b>" + (f" – {html_lib.escape(injury_type)}" if injury_type else "")
    if updated:
        text += f"<br>Updated {html_lib.escape(updated)}"
    if notes:
        text += f"<br>{html_lib.escape(notes)}"
    return text


def player_label(name, stats):
    status = injury_status(name, stats)
    marker = {
        "q": "⚠️",
        "questionable": "⚠️",
        "o": "⛔",
        "out": "⛔",
    }.get(status)
    return f"{marker} {name}" if marker else name


def short_player_name(label):
    """'Darius Acuff Jr.' -> 'D. Acuff Jr.', keeping any status marker."""
    marker = ""
    for prefix in ("⚠️ ", "⛔ "):
        if label.startswith(prefix):
            marker, label = prefix, label[len(prefix):]
            break
    parts = label.split()
    if len(parts) < 2:
        return marker + label
    return f"{marker}{parts[0][0]}. {' '.join(parts[1:])}"


def make_next_opponents(data):
    """{normalized team: next opponent name, or 'TBD'} for teams still alive; cached on data."""
    cached = data.get("_next_opponents")
    if cached is not None:
        return cached
    result = {}
    if data.get("bracket_teams"):
        try:
            rounds = build_tournament_bracket(data["bracket_teams"], data["teams_out"])
            regional_rounds = rounds[:-2]
            regional_winners = {}
            for region in ("EAST", "WEST", "MIDWEST", "SOUTH"):
                elite_eight = regional_rounds[-1]["regions"][region]
                regional_winners[region] = (elite_eight[0].get("winner") if elite_eight else None) or "TBD"
            final_four_rounds = build_final_four_brackets(
                regional_winners, data["final_four_config"], data["teams_out"]
            )
            for round_data in regional_rounds + final_four_rounds:
                for matchups in round_data["regions"].values():
                    for matchup in matchups:
                        if matchup.get("winner"):
                            continue
                        for team, opponent in (
                            (matchup["team_a"], matchup["team_b"]),
                            (matchup["team_b"], matchup["team_a"]),
                        ):
                            key = _norm_team_name(team)
                            if "TBD" in str(team) or key in result:
                                continue
                            if str(team).strip().casefold() in data["teams_out"]:
                                continue
                            result[key] = "TBD" if "TBD" in str(opponent) else str(opponent)
        except Exception:
            result = {}
    data["_next_opponents"] = result
    return result


def next_opponent_for(data, team):
    if not team:
        return "—"
    return make_next_opponents(data).get(_norm_team_name(team), "—")


def make_team_rows(data, selected_draftee):
    rows = [
        row for row in data["draft_rows"]
        if row["draftee"].casefold() == selected_draftee.casefold()
    ]
    rows.sort(key=lambda row: draft_position_sort_key(row["draft_position"]))
    next_opponents = make_next_opponents(data)

    result = []
    for row in rows:
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        result.append({
            "Draft Position": row["draft_position"],
            "Player Name": player_label(row["player_name"], stats),
            "_photo": get_player_photo_url(row["player_name"], stats.get("Team")),
            "Team": stats.get("Team"),
            "Next Opponent": (
                next_opponents.get(_norm_team_name(stats.get("Team")), "—")
                if stats.get("Team") else "—"
            ),
            "Seed": stats.get("Seed"),
            "FPts": stats.get("Tot. FPts"),
            "PPG": stats.get("PPG"),
            "RPG": stats.get("RPG"),
            "APG": stats.get("APG"),
            "Tourn. FPts": player_total(row, data["fantasy_weights"]),
            "Tourn. PTS": row.get("tourn_pts"),
            "Tourn. REB": row.get("tourn_reb"),
            "Tourn. AST": row.get("tourn_ast"),
            "Projected FPts": player_projected_fpts(row, data),
            "_projected_tip": projected_fpts_explanation(row, data),
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
            "Expected Games Remaining": round(player_expected_games_remaining(row, data), 2),
            "_expected_tip": expected_games_explanation(row, data),
            "_team_out": (
                stats.get("Team") is not None
                and str(stats.get("Team")).strip().casefold() in data["teams_out"]
            ),
        })
    return result


def make_standings_round_history(data):
    """Return each draftee's ranking after each tournament round."""
    round_names = [
        "First Round",
        "Second Round",
        "Sweet Sixteen",
        "Elite Eight",
        "Final Four",
        "National Title",
    ]

    # Draft columns are grouped as 6 rounds of PTS, followed by 6 REB,
    # followed by 6 AST. For each round, add PTS + REB + AST for every
    # drafted player, then accumulate the result through the tournament.
    round_totals = {
        draftee: [0.0] * len(round_names)
        for draftee in data["drafters"]
    }

    for row in data["draft_rows"]:
        owner = row["draftee"]
        pts = row.get("tourn_pts_rounds", [None] * 6)
        reb = row.get("tourn_reb_rounds", [None] * 6)
        ast = row.get("tourn_ast_rounds", [None] * 6)

        for i in range(6):
            round_total = weighted_fantasy_points(
                pts[i], reb[i], ast[i], data["fantasy_weights"]
            )
            if round_total is not None:
                round_totals.setdefault(owner, [0.0] * len(round_names))[i] += round_total

    # Only include rounds that have at least one numeric stat recorded for
    # at least one drafted player. This prevents future/unplayed rounds from
    # appearing in the progression chart before any games have been played.
    played_round_indexes = []
    for i in range(len(round_names)):
        has_numeric_value = False
        for row in data["draft_rows"]:
            for key in ("tourn_pts_rounds", "tourn_reb_rounds", "tourn_ast_rounds"):
                values = row.get(key, [None] * len(round_names))
                if i < len(values):
                    value = values[i]
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        has_numeric_value = True
                        break
            if has_numeric_value:
                break
        if has_numeric_value:
            played_round_indexes.append(i)

    # Convert round-by-round scoring into cumulative totals, then rank
    # everyone after each played round. Rank 1 is the best (highest) total.
    result = []
    running_total = {owner: 0.0 for owner in round_totals}
    for i in played_round_indexes:
        round_name = round_names[i]
        for owner, values in round_totals.items():
            running_total[owner] += values[i]

        ordered = sorted(
            data["drafters"],
            key=lambda owner: (-running_total.get(owner, 0.0), owner.casefold()),
        )
        rankings = {owner: rank for rank, owner in enumerate(ordered, start=1)}

        row = {"Round": round_name}
        for owner in data["drafters"]:
            row[owner] = rankings.get(owner)
        result.append(row)

    return pd.DataFrame(result)


def make_team_round_changes(data, owner):
    """Return changes for one owner between the last two played rounds."""
    weights = data["fantasy_weights"]
    keys = ("tourn_pts_rounds", "tourn_reb_rounds", "tourn_ast_rounds")
    rows = data["draft_rows"]

    def has_value(i):
        for row in rows:
            for key in keys:
                values = row.get(key) or []
                if i < len(values) and isinstance(values[i], (int, float)) and not isinstance(values[i], bool):
                    return True
        return False

    played = [i for i in range(6) if has_value(i)]
    result = {"fpts_delta": None, "rank_change": None, "active_change": None, "round_name": None}
    if not played:
        return result
    last = played[-1]
    result["round_name"] = [
        "First Round", "Second Round", "Sweet Sixteen",
        "Elite Eight", "Final Four", "National Title",
    ][played[-2]] if len(played) >= 2 else "start"

    round_totals = {name: [0.0] * 6 for name in data["drafters"]}
    for row in rows:
        pts = row.get("tourn_pts_rounds") or [None] * 6
        reb = row.get("tourn_reb_rounds") or [None] * 6
        ast = row.get("tourn_ast_rounds") or [None] * 6
        for i in range(6):
            value = weighted_fantasy_points(pts[i], reb[i], ast[i], weights)
            if value is not None and row["draftee"] in round_totals:
                round_totals[row["draftee"]][i] += value

    def ranks_through(round_index):
        totals = {
            name: sum(values[: round_index + 1])
            for name, values in round_totals.items()
        }
        ordered = sorted(totals, key=lambda name: (-totals[name], name.casefold()))
        return {name: rank for rank, name in enumerate(ordered, start=1)}

    owner_key = next((n for n in round_totals if n.casefold() == owner.casefold()), None)
    if owner_key is None:
        return result

    result["fpts_delta"] = round_totals[owner_key][last]
    if len(played) >= 2:
        previous = played[-2]
        result["rank_change"] = ranks_through(previous)[owner_key] - ranks_through(last)[owner_key]

    current_active = 0
    previous_active = 0
    for row in rows:
        if row["draftee"].casefold() != owner.casefold():
            continue
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        if team is None:
            continue
        team_key = str(team).strip().casefold()
        if team_key not in data["teams_out"]:
            current_active += 1
            previous_active += 1
            continue
        lost_idx = round_lost_index(data["teams_out"][team_key])
        if lost_idx is None:
            lost_idx = last
        if lost_idx >= last:
            previous_active += 1
    result["active_change"] = current_active - previous_active
    return result


def team_round_fpts(data, owner):
    """Return one owner's FPts per tournament round and their top scorer as (name, FPts)."""
    totals = [0.0] * 6
    best = (None, None)
    for row in data["draft_rows"]:
        if row["draftee"].casefold() != owner.casefold():
            continue
        pts = row.get("tourn_pts_rounds") or [None] * 6
        reb = row.get("tourn_reb_rounds") or [None] * 6
        ast = row.get("tourn_ast_rounds") or [None] * 6
        for i in range(6):
            value = weighted_fantasy_points(pts[i], reb[i], ast[i], data["fantasy_weights"])
            if value is not None:
                totals[i] += float(value)
        total = player_total(row, data["fantasy_weights"])
        if total is not None and (best[1] is None or float(total) > best[1]):
            best = (row["player_name"], float(total))
    return totals, best


def owner_value(df, owner, column):
    """One owner's numeric value from a per-owner table, or None."""
    match = df[df["Team Owner"].astype(str).str.casefold() == owner.casefold()]
    if match.empty or pd.isna(match.iloc[0][column]):
        return None
    return float(match.iloc[0][column])


def make_standings(data):
    positions = {
        draftee.casefold(): position
        for position, draftee in enumerate(data["drafters"], start=1)
    }
    totals = {
        d.casefold(): {
            "Team Owner": d,
            "Draft Position": positions.get(d.casefold(), ""),
            "Total FPts": None,
            "Active Players": 0,
            "Poss. Games Remaining": 0,
            "MVP": "",
            "MVP FPts": None,
        }
        for d in data["drafters"]
    }

    for row in data["draft_rows"]:
        key = row["draftee"].casefold()
        if key not in totals:
            totals[key] = {
                "Team Owner": row["draftee"],
                "Draft Position": positions.get(key, ""),
                "Total FPts": None,
                "Active Players": 0,
                "Poss. Games Remaining": 0,
                "MVP": "",
                "MVP FPts": None,
            }

        t = totals[key]
        t["Total FPts"] = sum_existing_values([
            t["Total FPts"], player_total(row, data["fantasy_weights"])
        ])

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        if team is not None and str(team).strip().casefold() not in data["teams_out"]:
            t["Active Players"] += 1

        t["Poss. Games Remaining"] += row.get("possible_games_remaining") or 0

        pf = player_total(row, data["fantasy_weights"])
        if pf is not None and (t["MVP FPts"] is None or pf > t["MVP FPts"]):
            t["MVP"] = row["player_name"]
            t["MVP FPts"] = pf

    standings = list(totals.values())
    standings.sort(key=lambda x: (
        -(x["Total FPts"] if x["Total FPts"] is not None else 0),
        x["Draft Position"] if isinstance(x["Draft Position"], int) else 999999,
        x["Team Owner"].casefold(),
    ))
    for i, row in enumerate(standings, start=1):
        row["Rank"] = i

    # A pre-draft workbook can have no rows at all. Return an empty
    # DataFrame with the expected schema so downstream page rendering
    # does not raise a KeyError when selecting these columns.
    standings_columns = [
        "Rank", "Draft Position", "Team Owner", "Total FPts",
        "Active Players", "Poss. Games Remaining", "MVP", "MVP FPts"
    ]
    if not standings:
        return pd.DataFrame(columns=standings_columns)

    return pd.DataFrame(standings)[standings_columns]


def make_power_rankings(data):
    positions = {
        draftee.casefold(): position
        for position, draftee in enumerate(data["drafters"], start=1)
    }
    totals = {
        d.casefold(): {
            "Team Owner": d,
            "Draft Position": positions.get(d.casefold(), ""),
            "Total FPts": None,
            "Active Players": 0,
            "Games Played": 0,
        }
        for d in data["drafters"]
    }

    for row in data["draft_rows"]:
        key = row["draftee"].casefold()
        totals.setdefault(key, {
            "Team Owner": row["draftee"],
            "Draft Position": positions.get(key, ""),
            "Total FPts": None,
            "Active Players": 0,
            "Games Played": 0,
        })

        t = totals[key]
        current_fpts = player_total(row, data["fantasy_weights"])
        t["Total FPts"] = sum_existing_values([t["Total FPts"], current_fpts])

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        if team is not None and str(team).strip().casefold() not in data["teams_out"]:
            t["Active Players"] += 1

        t["Games Played"] += row.get("games_played") or 0

    result = []
    tournament_weight = data["tournament_fpts_weight"]

    for key, t in totals.items():
        team_projected_fpts = 0.0
        expected_games_total = 0.0

        for row in data["draft_rows"]:
            if row["draftee"].casefold() != key:
                continue

            stats = data["player_stats"].get(row["player_name"].casefold(), {})
            team = stats.get("Team")
            team_is_active = (
                team is not None
                and str(team).strip().casefold() not in data["teams_out"]
            )

            current_player_fpts = float(player_total(row, data["fantasy_weights"]) or 0.0)
            games_played = row.get("games_played") or 0

            tourn_fpts_per_game = (
                current_player_fpts / games_played
                if games_played > 0 else 0.0
            )

            # "Tot. FPts" in player_stats is the season FPts per game
            # calculated from the season PPG/RPG/APG fantasy components.
            season_fpts_per_game = float(stats.get("Tot. FPts") or 0.0)

            # If the player has not recorded any tournament games yet,
            # use the season FPts/game by itself. There is no tournament
            # sample to blend with the season rate.
            if games_played == 0:
                weighted_fpts_per_game = season_fpts_per_game
            else:
                weighted_fpts_per_game = (
                    tournament_weight * tourn_fpts_per_game
                    + (1.0 - tournament_weight) * season_fpts_per_game
                )

            expected_remaining_games = player_expected_games_remaining(row, data)

            expected_games_total += expected_remaining_games

            team_projected_fpts += (
                current_player_fpts
                + weighted_fpts_per_game * expected_remaining_games
            )

        total_fpts = t["Total FPts"]
        fpg = (
            total_fpts / t["Games Played"]
            if total_fpts is not None and t["Games Played"] > 0
            else None
        )

        result.append({
            **t,
            "Projected Score": round(team_projected_fpts, 1),
            "FPts per game": round(fpg, 1) if fpg is not None else None,
            "Expected Games Remaining": round(expected_games_total, 2),
        })

    result.sort(key=lambda x: (
        -(x["Projected Score"] if x["Projected Score"] is not None else 0),
        x["Draft Position"] if isinstance(x["Draft Position"], int) else 999999,
        x["Team Owner"].casefold(),
    ))

    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    power_columns = [
        "Rank", "Draft Position", "Team Owner", "Projected Score",
        "Total FPts", "Active Players", "FPts per game",
        "Games Played", "Expected Games Remaining"
    ]
    if not result:
        return pd.DataFrame(columns=power_columns)

    return pd.DataFrame(result)[power_columns]


def _formula_expected_games_remaining(row, data):
    """Seed-average expected remaining games for one player (0 if eliminated or at 6 games)."""
    stats = data["player_stats"].get(row["player_name"].casefold(), {})
    team = stats.get("Team")
    team_is_active = (
        team is not None
        and str(team).strip().casefold() not in data["teams_out"]
    )
    games_played = row.get("games_played") or 0

    seed_value = stats.get("Seed")
    try:
        seed_number = int(float(seed_value))
    except (TypeError, ValueError):
        match = re.search(r"\d+", str(seed_value))
        seed_number = int(match.group(0)) if match else None

    historical_avg_games = (
        data["seed_average_games"].get(seed_number)
        if seed_number is not None else None
    )

    if injury_status(row["player_name"], stats) in ("o", "out"):
        return 0.0
    if not team_is_active or games_played >= 6:
        return 0.0
    if historical_avg_games is None:
        return 1.0
    return max(historical_avg_games - games_played, 1.0)


def _simulated_player(data, row):
    """This player's results from the 10,000-run simulation, or None if it isn't available."""
    sim = data.get("_win_prob")
    if sim is None:
        return None
    return sim.attrs.get("players", {}).get(row.get("draft_row"))


def player_expected_games_remaining(row, data):
    """Expected remaining games for one drafted player: simulated when possible, else seed averages."""
    info = _simulated_player(data, row)
    if info is not None:
        return float(info["games"])
    return _formula_expected_games_remaining(row, data)


def expected_games_explanation(row, data):
    """HTML text describing how player_expected_games_remaining() got its value."""
    stats = data["player_stats"].get(row["player_name"].casefold(), {})
    team = stats.get("Team")
    team_is_active = (
        team is not None
        and str(team).strip().casefold() not in data["teams_out"]
    )
    games_played = row.get("games_played") or 0

    seed_value = stats.get("Seed")
    try:
        seed_number = int(float(seed_value))
    except (TypeError, ValueError):
        match = re.search(r"\d+", str(seed_value))
        seed_number = int(match.group(0)) if match else None
    avg_games = (
        data["seed_average_games"].get(seed_number)
        if seed_number is not None else None
    )

    if injury_status(row["player_name"], stats) in ("o", "out"):
        return "Player is listed as out, so 0 games are expected."
    if not team_is_active:
        return "Team is eliminated, so 0 games are expected."
    if games_played >= 6:
        return "Player has played the maximum 6 games, so 0 games are expected."
    info = _simulated_player(data, row)
    if info is not None:
        return (
            "<b>Expected Games Remaining</b><br>"
            f"Average number of more games this player's team plays across {WIN_PROB_SIMS:,} simulated "
            f"tournaments: <b>{info['games']:.2f}</b>.<br>"
            "Uses the real bracket and results so far, with modeled odds for games not yet played."
        )
    if avg_games is None:
        return "No average games found for this seed, so 1 game is assumed."
    remaining = avg_games - games_played
    return (
        f"<b>Expected Games Remaining</b><br>"
        f"Seed {seed_number} averages {avg_games:.2f} games.<br>"
        f"{avg_games:.2f} avg. - {games_played} played = {remaining:.2f}<br>"
        f"Never less than 1 while the team is alive:<br>"
        f"<b>max({remaining:.2f}, 1) = {max(remaining, 1.0):.2f}</b>"
    )


def player_projected_fpts(row, data):
    """Projected final tournament FPts: the simulation average when available, else a seed-average formula.

    The formula is tournament FPts so far plus (weighted FPts/game x expected games remaining), where
    FPts/game blends the tournament and season averages using Config!J2.
    """
    info = _simulated_player(data, row)
    if info is not None:
        return round(info["projected"], 1)
    stats = data["player_stats"].get(row["player_name"].casefold(), {})
    season_per_game = sum_existing_values([stats.get("Tot. FPts")])
    if season_per_game is None:
        return None

    games = row.get("games_played") or 0
    total = player_total(row, data["fantasy_weights"])

    if games <= 0 or total is None:
        weighted_per_game = season_per_game
        total = 0.0
    else:
        weight = data["tournament_fpts_weight"]
        weighted_per_game = (
            weight * (total / games) + (1.0 - weight) * season_per_game
        )

    expected_games = player_expected_games_remaining(row, data)
    return round(total + weighted_per_game * expected_games, 1)


def projected_fpts_explanation(row, data):
    """HTML text describing how player_projected_fpts() got its value."""
    info = _simulated_player(data, row)
    if info is not None:
        text = (
            "<b>Projected FPts</b><br>"
            f"Average final FPts across {WIN_PROB_SIMS:,} simulated tournaments.<br>"
        )
        if info["games"] < 0.005:
            text += (
                f"No more games expected, so the {info['current']:.1f} FPts already scored is the projection.<br>"
            )
        else:
            text += (
                f"{info['current']:.1f} FPts scored so far + {info['future']:.1f} from "
                f"{info['games']:.2f} expected more games (about {info['per_game']:.1f} FPts per game).<br>"
            )
            if info["low"] is not None:
                text += f"Middle 80% of outcomes: {info['low']:.0f} to {info['high']:.0f}.<br>"
        return text + f"<b>= {info['projected']:.1f}</b>"
    stats = data["player_stats"].get(row["player_name"].casefold(), {})
    season_per_game = sum_existing_values([stats.get("Tot. FPts")])
    if season_per_game is None:
        return "No season average is available, so no projection can be made."

    games = row.get("games_played") or 0
    total = player_total(row, data["fantasy_weights"])
    expected_games = player_expected_games_remaining(row, data)
    projected = player_projected_fpts(row, data)

    if games <= 0 or total is None:
        return (
            "<b>Projected FPts</b><br>"
            "No tournament games yet, so the season average is used alone.<br>"
            f"{season_per_game:.1f} season FPts/game x {expected_games:.2f} expected games<br>"
            f"<b>= {projected:.1f}</b>"
        )

    weight = data["tournament_fpts_weight"]
    tourn_per_game = total / games
    weighted_per_game = weight * tourn_per_game + (1.0 - weight) * season_per_game
    return (
        "<b>Projected FPts</b><br>"
        f"Blended FPts/game:<br>[{season_per_game:.1f} season x {round((1 - weight) * 100, 1):g}%] + "
        f"[{tourn_per_game:.1f} tourn. x {round(weight * 100, 1):g}%] = {weighted_per_game:.2f}<br>"
        f"{weighted_per_game:.2f} x {expected_games:.2f} expected games = "
        f"{weighted_per_game * expected_games:.1f}<br>"
        f"+ {total:.1f} FPts already scored<br>"
        f"<b>= {projected:.1f}</b>"
    )


def make_player_rankings(data, only_unavailable=False):
    result = []
    for row in data["draft_rows"]:
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        if only_unavailable and injury_status(row["player_name"], stats) not in (
            "q", "questionable", "o", "out"
        ):
            continue
        injury_order = (
            0 if injury_status(row["player_name"], stats) in ("o", "out") else 1
        )
        team = stats.get("Team")
        result.append({
            "_injury_order": injury_order,
            "Player Name": player_label(row["player_name"], stats),
            "_photo": get_player_photo_url(row["player_name"], team),
            "Team": team,
            "Next Opponent": next_opponent_for(data, team),
            "Total FPts": player_total(row, data["fantasy_weights"]),
            "Team Owner": row["draftee"],
            "Games Played": row.get("games_played", 0),
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
            "Expected Games Remaining": round(player_expected_games_remaining(row, data), 2),
            "_expected_tip": expected_games_explanation(row, data),
            "Projected FPts": player_projected_fpts(row, data),
            "_projected_tip": projected_fpts_explanation(row, data),
            "_team_out": (
                team is not None
                and str(team).strip().casefold() in data["teams_out"]
            ),
            "_draft_position": row["draft_position"],
        })

    result.sort(key=lambda p: (
        p["_injury_order"] if only_unavailable else 0,
        -(p["Total FPts"] if p["Total FPts"] is not None else 0),
        draft_position_sort_key(p["_draft_position"]),
    ))
    result = result[:10] if not only_unavailable else result
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    player_columns = [
        "Rank", "Player Name", "Team", "Next Opponent", "Total FPts", "Projected FPts",
        "Team Owner", "Games Played", "Poss. Games Remaining",
        "Expected Games Remaining", "_expected_tip", "_projected_tip", "_team_out", "_photo",
    ]
    if not result:
        return pd.DataFrame(columns=player_columns), result

    return pd.DataFrame(result)[player_columns], result


def make_performers(data, positive=True, exclude_unavailable=False):
    result = []
    for row in data["draft_rows"]:
        total = player_total(row, data["fantasy_weights"])
        games = row.get("games_played", 0)
        if games <= 0 or total is None:
            continue

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        if exclude_unavailable and injury_status(row["player_name"], stats) in ("q", "questionable", "o", "out"):
            continue
        season_avg = sum_existing_values([stats.get("Tot. FPts")])
        if season_avg is None:
            continue

        tournament_avg = total / games
        delta = tournament_avg - season_avg

        if positive and delta <= 0:
            continue
        if not positive and delta >= 0:
            continue

        team = stats.get("Team")
        result.append({
            "Player Name": player_label(row["player_name"], stats),
            "_photo": get_player_photo_url(row["player_name"], team),
            "Team": team,
            "Next Opponent": next_opponent_for(data, team),
            "Avg. FPts/Tourn. Game": round(tournament_avg, 1),
            "Delta From Average": round(delta, 1),
            "Team Owner": row["draftee"],
            "Games Played": games,
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
            "Expected Games Remaining": round(player_expected_games_remaining(row, data), 2),
            "_expected_tip": expected_games_explanation(row, data),
            "Projected FPts": player_projected_fpts(row, data),
            "_projected_tip": projected_fpts_explanation(row, data),
            "_team_out": (
                team is not None
                and str(team).strip().casefold() in data["teams_out"]
            ),
            "_draft_position": row["draft_position"],
        })

    result.sort(key=lambda p: (
        -p["Delta From Average"] if positive else p["Delta From Average"],
        draft_position_sort_key(p["_draft_position"]),
    ))
    result = result[:5]
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    performer_columns = [
        "Rank", "Player Name", "Team", "Next Opponent", "Avg. FPts/Tourn. Game", "Delta From Average",
        "Projected FPts", "Team Owner", "Games Played",
        "Poss. Games Remaining", "Expected Games Remaining", "_expected_tip", "_projected_tip", "_team_out", "_photo",
    ]
    if not result:
        return pd.DataFrame(columns=performer_columns), result

    return pd.DataFrame(result)[performer_columns], result


def make_draft_value(data, steals=True, exclude_unavailable=False):
    """Top 5 picks whose projected-FPts rank beats (steals) or trails (busts) their draft order."""
    drafted = sorted(
        data["draft_rows"], key=lambda r: draft_position_sort_key(r["draft_position"])
    )
    entries = []
    for row in drafted:
        projected = player_projected_fpts(row, data)
        if projected is not None:
            entries.append((row, projected))
    ranked = sorted(
        range(len(entries)), key=lambda i: (-entries[i][1], i)
    )
    projected_rank = {i: rank for rank, i in enumerate(ranked, start=1)}

    result = []
    for index, (row, projected) in enumerate(entries):
        pick_number = index + 1
        value = pick_number - projected_rank[index]
        if (steals and value <= 0) or (not steals and value >= 0):
            continue
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        if exclude_unavailable and injury_status(row["player_name"], stats) in ("q", "questionable", "o", "out"):
            continue
        team = stats.get("Team")
        result.append({
            "Player Name": player_label(row["player_name"], stats),
            "_photo": get_player_photo_url(row["player_name"], team),
            "Team": team,
            "Next Opponent": next_opponent_for(data, team),
            "Draft Pos.": row["draft_position"],
            "Overall Pick": pick_number,
            "Projected Rank": projected_rank[index],
            "Pick Value": f"{value:+d}",
            "_value": value,
            "Team Owner": row["draftee"],
            "Games Played": row.get("games_played", 0),
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
            "Expected Games Remaining": round(player_expected_games_remaining(row, data), 2),
            "_expected_tip": expected_games_explanation(row, data),
            "Projected FPts": projected,
            "_projected_tip": projected_fpts_explanation(row, data),
            "_team_out": (
                team is not None
                and str(team).strip().casefold() in data["teams_out"]
            ),
            "_pick": pick_number,
        })

    result.sort(key=lambda p: (-p["_value"] if steals else p["_value"], p["_pick"]))
    result = result[:5]
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    value_columns = [
        "Rank", "Player Name", "Team", "Next Opponent", "Draft Pos.", "Overall Pick",
        "Projected Rank", "Pick Value", "Projected FPts", "Team Owner", "Games Played",
        "Poss. Games Remaining", "Expected Games Remaining", "_expected_tip", "_projected_tip",
        "_team_out", "_photo",
    ]
    if not result:
        return pd.DataFrame(columns=value_columns)
    return pd.DataFrame(result)[value_columns]

def make_draft_ratings(data):
    positions = {
        draftee.casefold(): position
        for position, draftee in enumerate(data["drafters"], start=1)
    }
    totals = {
        d.casefold(): {
            "Team Owner": d,
            "Draft Position": positions.get(d.casefold(), ""),
            "Regions": set(),
            "Projected FPts": None,
        }
        for d in data["drafters"]
    }

    for row in data["draft_rows"]:
        key = row["draftee"].casefold()
        totals.setdefault(key, {
            "Team Owner": row["draftee"],
            "Draft Position": positions.get(key, ""),
            "Regions": set(),
            "Projected FPts": None,
        })

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        avg_games = average_games_for_seed(stats.get("Seed"))
        season_avg = sum_existing_values([stats.get("Tot. FPts")])

        region = stats.get("Region")
        if region:
            totals[key]["Regions"].add(str(region).strip().casefold())

        if avg_games is None or season_avg is None:
            continue

        projected = avg_games * season_avg
        totals[key]["Projected FPts"] = sum_existing_values([
            totals[key]["Projected FPts"], projected
        ])

    result = [{
        "Team Owner": t["Team Owner"],
        "Draft Position": t["Draft Position"],
        "Projected FPts": t["Projected FPts"],
        "Regions": len(t["Regions"]),
    } for t in totals.values()]

    result.sort(key=lambda x: (
        -(x["Projected FPts"] if x["Projected FPts"] is not None else 0),
        x["Draft Position"] if isinstance(x["Draft Position"], int) else 999999,
        x["Team Owner"].casefold(),
    ))
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    ratings_columns = [
        "Rank", "Draft Position", "Team Owner", "Projected FPts", "Regions"
    ]
    if not result:
        return pd.DataFrame(columns=ratings_columns)

    return pd.DataFrame(result)[ratings_columns]


# Odds constants and the bracket simulation live in bracket_sim.py (shared with the commissioner app).

TORVIK_CSV_URL = "https://barttorvik.com/{year}_team_results.csv"
# Written by the commissioner app's Fetch Torvik Team Ratings button; used when the live download is blocked.
TORVIK_SNAPSHOT = ASSETS_DIR / "torvik_ratings.json"


def torvik_season_year():
    now = datetime.now()
    return now.year + 1 if now.month >= 10 else now.year


@st.cache_data(ttl=21600, show_spinner=False)
def _fetch_torvik_ratings(year):
    """{normalised team: (adjoe, adjde, adjt, rank, record, conf)} from Barttorvik's season CSV, cached 6 hours."""
    resp = requests.get(
        TORVIK_CSV_URL.format(year=year),
        headers={"User-Agent": "MM-Player-Pool-Dashboard (small private pool; cached 6 hours)"},
        timeout=20,
    )
    resp.raise_for_status()
    ratings = {}
    for row in csv.DictReader(io.StringIO(resp.text)):
        try:
            ratings[_torvik_key(row["team"])] = (
                float(row["adjoe"]), float(row["adjde"]), float(row["adjt"]),
                int(float(row["rank"])), row["record"], row["conf"],
            )
        except (KeyError, TypeError, ValueError):
            continue
    return ratings


@st.cache_data(ttl=600, show_spinner=False)
def get_torvik_ratings(year):
    """(ratings, error, source): live Barttorvik ratings, else the repo snapshot, else nothing.

    source is "live", "snapshot:<date>" or ""; error is the live fetch failure (if any).
    """
    try:
        return _fetch_torvik_ratings(year), "", "live"
    except Exception as exc:
        reason = f"{type(exc).__name__}: {exc}"[:160]
    try:
        payload = json.loads(TORVIK_SNAPSHOT.read_text(encoding="utf-8"))
        if payload.get("year") == year:
            ratings = {}
            for name, values in payload["teams"].items():
                values = list(values) + [None] * (6 - len(values))
                ratings[_torvik_key(name)] = (
                    float(values[0]), float(values[1]), float(values[2]), values[3], values[4], values[5],
                )
            return ratings, reason, f"snapshot:{str(payload.get('fetched', ''))[:10]}"
    except Exception:
        pass
    return {}, reason, ""


@st.cache_data(ttl=600, show_spinner=False)
def get_odds_ratings(year):
    """(ratings for game odds, NET teams matched, NET data-through label).

    Barttorvik ratings, blended with the NCAA NET rankings once they are published for this season.
    """
    net_ranks, through = load_net_ranks(year)
    blended, matched = blend_net_rankings(get_torvik_ratings(year)[0], net_ranks)
    return blended, matched, through


def torvik_team_info(name):
    """Barttorvik rank, record and adjusted efficiency numbers for one team, or None."""
    ratings, _error, _source = get_torvik_ratings(torvik_season_year())
    key = _torvik_key(name)
    rating = ratings.get(key) or ratings.get(TORVIK_ALIASES.get(key, ""))
    if not rating:
        return None
    return {
        "rank": rating[3], "record": rating[4],
        "o": round(rating[0], 1), "d": round(rating[1], 1), "t": round(rating[2], 1),
        "net": round(rating[0] - rating[1], 1),
    }


def matchup_margin(team_a, seed_a, team_b, seed_b):
    """Predicted points margin for team A over team B at a neutral site (negative if B is favored); None if unknown."""
    ratings = get_odds_ratings(torvik_season_year())[0]
    tempos = [r[2] for r in ratings.values()]
    avg_tempo = sum(tempos) / len(tempos) if tempos else DEFAULT_TEMPO

    def strength(name, seed_value):
        key = _torvik_key(name)
        rating = ratings.get(key) or ratings.get(TORVIK_ALIASES.get(key, ""))
        if rating:
            return rating[0] - rating[1], rating[2]
        try:
            return SEED_STRENGTH[int(seed_value)] * 100.0 / avg_tempo, avg_tempo
        except (KeyError, TypeError, ValueError):
            return None

    a, b = strength(team_a, seed_a), strength(team_b, seed_b)
    if a is None or b is None:
        return None
    return (a[0] - b[0]) / 100.0 * a[1] * b[1] / avg_tempo


def matchup_win_probability(team_a, seed_a, team_b, seed_b):
    """Chance team A beats team B at a neutral site (Barttorvik ratings, seed fallback); None if unknown."""
    margin = matchup_margin(team_a, seed_a, team_b, seed_b)
    if margin is None:
        return None
    return 0.5 * (1.0 + math.erf(margin / (GAME_MARGIN_SD * math.sqrt(2.0))))


def _simulate_win_probabilities(data, n_sims, seed, pre_tournament=False):
    import numpy as np

    owners = data["drafters"]
    if not data["bracket_teams"] or not owners or not data["draft_rows"]:
        return None
    rng = np.random.default_rng(seed)
    teams_out = {} if pre_tournament else data["teams_out"]

    year = torvik_season_year()
    ratings, rating_error, rating_source = get_torvik_ratings(year)
    odds_ratings, net_matched, net_through = get_odds_ratings(year)
    lost_rounds = {name: round_lost_index(value) for name, value in teams_out.items()}
    bracket_result = run_bracket(data["bracket_teams"], data["final_four_config"], odds_ratings, n_sims, rng, lost_rounds)
    if bracket_result is None:
        return None
    appearances = bracket_result["appearances"]
    team_index = bracket_result["team_index"]
    rating_info = {
        "matched": bracket_result["matched"], "total": len(bracket_result["team_names"]),
        "year": year, "source": rating_source, "net_matched": net_matched, "net_through": net_through,
        "error": rating_error if not ratings else ("no team names matched" if not bracket_result["matched"] else ""),
    }
    rows = np.arange(n_sims)[:, None]

    weight = data["tournament_fpts_weight"]
    owner_pos = {name.casefold(): i for i, name in enumerate(owners)}
    current_totals = np.zeros(len(owners))
    active_players = np.zeros(len(owners), dtype=int)
    games_played = np.zeros(len(owners), dtype=int)
    expected_games = np.zeros(len(owners))
    sim_players = []
    player_info = {}
    fixed_future = np.zeros(len(owners))
    owner_players = {}
    for row in data["draft_rows"]:
        owner = owner_pos.get(row["draftee"].casefold())
        if owner is None:
            continue
        owner_players.setdefault(owner, []).append((row["player_name"], row.get("draft_row")))
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        total = None if pre_tournament else player_total(row, data["fantasy_weights"])
        current_totals[owner] += float(total or 0.0)
        season = sum_existing_values([stats.get("Tot. FPts")]) or 0.0
        games = 0 if pre_tournament else (row.get("games_played") or 0)
        games_played[owner] += games
        team_name = stats.get("Team")
        if team_name is not None and str(team_name).strip().casefold() not in teams_out:
            active_players[owner] += 1
        per_game = (
            weight * (float(total) / games) + (1.0 - weight) * season
            if games > 0 and total is not None else season
        )
        team_key = _norm_team_name(stats.get("Team")) if stats.get("Team") else None
        if team_key in team_index:
            blocked = not pre_tournament and (
                str(stats.get("Team")).strip().casefold() in teams_out
                or injury_status(row["player_name"], stats) in ("o", "out")
            )
            sim_players.append(
                (owner, team_index[team_key], games, per_game, blocked, row.get("draft_row"), float(total or 0.0))
            )
        else:
            fixed_games = (
                (average_games_for_seed(stats.get("Seed")) or 1.0) if pre_tournament
                else _formula_expected_games_remaining(row, data)
            )
            expected_games[owner] += fixed_games
            fixed_future[owner] += per_game * fixed_games
            player_info[row.get("draft_row")] = {
                "projected": float(total or 0.0) + per_game * fixed_games, "games": fixed_games,
                "current": float(total or 0.0), "future": per_game * fixed_games,
                "per_game": per_game, "low": None, "high": None,
            }

    totals = np.tile(current_totals + fixed_future, (n_sims, 1))
    if sim_players:
        owner_idx = np.array([p[0] for p in sim_players])
        team_idx = np.array([p[1] for p in sim_players])
        games_done = np.array([p[2] for p in sim_players])
        per_game = np.array([p[3] for p in sim_players])
        blocked = np.array([p[4] for p in sim_players])
        games_left = np.maximum(appearances[:, team_idx] - games_done[None, :], 0)
        games_left = np.where(blocked[None, :], 0, games_left)
        noise = rng.standard_normal(games_left.shape)
        future = np.maximum(
            games_left * per_game + GAME_FPTS_SPREAD * per_game * np.sqrt(games_left) * noise, 0.0
        )
        membership = np.zeros((len(sim_players), len(owners)))
        membership[np.arange(len(sim_players)), owner_idx] = 1.0
        totals = totals + future @ membership
        expected_games = expected_games + games_left.mean(axis=0) @ membership
        games_mean = games_left.mean(axis=0)
        future_mean = per_game * games_mean  # the noise averages out, so skip its sampling error here
        future_low, future_high = np.percentile(future, [10, 90], axis=0)
        for j, player in enumerate(sim_players):
            current_fpts = player[6]
            player_info[player[5]] = {
                "projected": current_fpts + float(future_mean[j]), "games": float(games_mean[j]),
                "current": current_fpts, "future": float(future_mean[j]), "per_game": float(per_game[j]),
                "low": current_fpts + float(future_low[j]), "high": current_fpts + float(future_high[j]),
            }

    order = np.argsort(-totals, axis=1, kind="stable")
    places = np.empty_like(order)
    places[rows, order] = np.arange(len(owners))[None, :]

    cfg = data["payout_config"]
    base = cfg["D2"]
    last_pay, third_pay, second_pay = base * cfg["E2"], base * cfg["F2"], base * cfg["G2"]
    first_pay = base * len(owners) - last_pay - third_pay - second_pay
    payout_by_place = np.zeros(len(owners))
    payout_by_place[0] = first_pay
    if len(owners) > 1:
        payout_by_place[1] = second_pay
    if len(owners) > 2:
        payout_by_place[2] = third_pay
    if len(owners) > 3:
        payout_by_place[-1] = last_pay
    place_probs = np.stack([(places == k).mean(axis=0) for k in range(len(owners))])
    expected_payout = payout_by_place @ place_probs

    avg_finish = (places + 1).mean(axis=0)
    mean_score = totals.mean(axis=0)
    low, high = np.percentile(totals, [10, 90], axis=0)
    win_counts = (places == 0).sum(axis=0)
    finish_tips = [
        f"<b>Avg. Finish</b><br>Average place across {n_sims:,} simulated tournaments: {avg_finish[i]:.2f}.<br>"
        f"1st: {place_probs[0][i] * 100:.1f}% · top 3: {place_probs[:3, i].sum() * 100:.1f}% "
        f"· last: {place_probs[-1][i] * 100:.1f}%"
        for i in range(len(owners))
    ]
    def score_tip(i):
        per_player = sorted(
            ((name, player_info[key]) for name, key in owner_players.get(i, []) if key in player_info),
            key=lambda item: -item[1]["projected"],
        )
        lines = "<br>".join(f"{html_lib.escape(name)}: {info['projected']:.1f}" for name, info in per_player)
        return (
            f"<b>Projected Score</b><br>Average final FPts across the simulations: "
            f"{current_totals[i]:.1f} scored so far + {mean_score[i] - current_totals[i]:.1f} still to come.<br>"
            f"Middle 80% of outcomes: {low[i]:.0f} to {high[i]:.0f}."
            + (f"<br>Projected FPts by player:<br>{lines}" if lines else "")
        )

    score_tips = [score_tip(i) for i in range(len(owners))]
    def games_tip(i):
        per_player = sorted(
            ((name, player_info[key]) for name, key in owner_players.get(i, []) if key in player_info),
            key=lambda item: -item[1]["games"],
        )
        lines = "<br>".join(
            f"{html_lib.escape(name)}: {info['games']:.2f}" for name, info in per_player if info["games"] >= 0.005
        )
        return (
            "<b>Expected Games Remaining</b><br>"
            f"Total additional games this owner's players' teams play, averaged across {n_sims:,} simulated "
            f"tournaments: <b>{expected_games[i]:.2f}</b>."
            + (f"<br>{lines}" if lines else "<br>No players with games left.")
        )

    games_tips = [games_tip(i) for i in range(len(owners))]
    win_tips = [
        f"<b>Win %</b><br>Finished 1st in {win_counts[i]:,} of {n_sims:,} simulated tournaments.<br>"
        f"Win % depends on the spread of outcomes, not just the average. Owners behind the leaders need a big "
        f"run to win, so a wider range ({low[i]:.0f} to {high[i]:.0f}) can give a higher Win % than a higher "
        "average score with a narrower range."
        for i in range(len(owners))
    ]

    result = pd.DataFrame({
        "Draft Position": range(1, len(owners) + 1),
        "Team Owner": owners,
        "Avg. Finish": avg_finish,
        "Projected Score": mean_score,
        "Total FPts": current_totals,
        "Win %": place_probs[0] * 100,
        "Top 3 %": place_probs[:3].sum(axis=0) * 100,
        "Last Place %": place_probs[-1] * 100,
        "Exp. Payout": [f"${value:,.2f}" for value in expected_payout],
        "Active Players": active_players,
        "FPts per game": [
            round(float(t) / g, 1) if g > 0 else None for t, g in zip(current_totals, games_played)
        ],
        "Games Played": games_played,
        "Expected Games Remaining": expected_games.round(2),
        "_finish_tip": finish_tips,
        "_score_tip": score_tips,
        "_win_tip": win_tips,
        "_expected_tip": games_tips,
    })
    result = result.sort_values(
        ["Avg. Finish", "Projected Score"], ascending=[True, False], kind="stable"
    ).reset_index(drop=True)
    result.insert(0, "Rank", range(1, len(result) + 1))
    result.attrs["players"] = player_info
    result.attrs["ratings"] = rating_info
    return result


def simulate_win_probabilities(data, n_sims=WIN_PROB_SIMS, seed=SIM_SEED, pre_tournament=False):
    """Monte Carlo finish odds per owner (cached on `data`); None if the bracket can't be simulated.

    pre_tournament ignores every game, result and injury entered so far.
    """
    cache_key = "_win_prob_pre" if pre_tournament else "_win_prob"
    if cache_key not in data:
        try:
            data[cache_key] = _simulate_win_probabilities(data, n_sims, seed, pre_tournament)
        except Exception:
            data[cache_key] = None
    return data[cache_key]


# ---------------------------------------------------------------------
# Streamlit presentation helpers
# ---------------------------------------------------------------------
def style_table(df, scale_columns=None, reverse_columns=None, medal_rank=True, pink_last=False, rank_rwg=False, performance_gradient=None, column_colors=None, zebra=False, cell_styles=None):
    scale_columns = scale_columns or {}
    reverse_columns = reverse_columns or {}

    # `_team_out` is an internal helper used only to identify eliminated
    # players. Remove it before creating the Styler so Streamlit cannot
    # render it as a visible column, while retaining the flags for styling.
    team_out_flags = None
    if "_team_out" in df.columns:
        team_out_flags = df["_team_out"].copy()
        df = df.drop(columns=["_team_out"])

    def apply_styles(row):
        styles = [""] * len(row)
        for col, (minimum, maximum) in scale_columns.items():
            if col in row.index:
                value = row[col]
                if isinstance(value, tuple):
                    value = value[0]
                bg = color_scale(
                    value, minimum, maximum,
                    reverse=reverse_columns.get(col, False)
                )
                styles[row.index.get_loc(col)] = f"background-color: {bg}"
        # Optional performance gradient applied ONLY to the Rank column.
        if performance_gradient in ("green", "red") and "Rank" in row.index:
            rank_values = pd.to_numeric(df["Rank"], errors="coerce")
            rank = pd.to_numeric(row["Rank"], errors="coerce")
            if pd.notna(rank) and rank_values.notna().any():
                lo = rank_values.min()
                hi = rank_values.max()
                if hi > lo:
                    # For both tables, Rank #1 gets the strongest color.
                    ratio = (hi - rank) / (hi - lo)
                    if performance_gradient == "green":
                        start = (235, 247, 238)
                        end = (112, 173, 71)
                    else:
                        start = (252, 235, 235)
                        end = (192, 80, 80)

                    rgb = tuple(
                        round(start[i] + (end[i] - start[i]) * ratio)
                        for i in range(3)
                    )
                    rank_color = "#%02X%02X%02X" % rgb
                    styles[row.index.get_loc("Rank")] = (
                        f"background-color: {rank_color}; color: #000000;"
                    )

        if rank_rwg and "Rank" in row.index:
            rank_values = pd.to_numeric(df["Rank"], errors="coerce")
            rank = pd.to_numeric(row["Rank"], errors="coerce")
            if pd.notna(rank) and rank_values.notna().any():
                lo = rank_values.min()
                hi = rank_values.max()
                if hi > lo:
                    midpoint = (lo + hi) / 2
                    if rank <= midpoint:
                        ratio = (rank - lo) / max(midpoint - lo, 1)
                        r = int(round(0 + 255 * ratio))
                        g = int(round(176 + 79 * ratio))
                        b = int(round(80 + 175 * ratio))
                    else:
                        ratio = (rank - midpoint) / max(hi - midpoint, 1)
                        r = 255
                        g = int(round(255 - 120 * ratio))
                        b = int(round(255 - 120 * ratio))
                    styles[row.index.get_loc("Rank")] = (
                        f"background-color: rgb({r}, {g}, {b})"
                    )

        if medal_rank and "Rank" in row.index and not rank_rwg:
            rank = row["Rank"]
            if pink_last and pd.notna(rank) and rank == df["Rank"].max():
                bg = "#FF87F7"
            else:
                bg = {
                    1: "#FFD700",
                    2: "#C0C0C0",
                    3: "#CD7F32",
                }.get(rank, "#FFFFFF")
            styles[row.index.get_loc("Rank")] = f"background-color: {bg}"

        # Performer gradient must be applied last so it overrides the
        # standard gold/silver/bronze rank colors. It affects ONLY Rank.
        if performance_gradient in ("green", "red") and "Rank" in row.index:
            rank_values = pd.to_numeric(df["Rank"], errors="coerce")
            rank = pd.to_numeric(row["Rank"], errors="coerce")
            if pd.notna(rank) and rank_values.notna().any():
                lo = rank_values.min()
                hi = rank_values.max()
                if hi > lo:
                    ratio = (hi - rank) / (hi - lo)
                    if performance_gradient == "green":
                        start = (235, 247, 238)
                        end = (112, 173, 71)
                    else:
                        start = (252, 235, 235)
                        end = (192, 80, 80)

                    rgb = tuple(
                        round(start[i] + (end[i] - start[i]) * ratio)
                        for i in range(3)
                    )
                    rank_color = "#%02X%02X%02X" % rgb
                    styles[row.index.get_loc("Rank")] = (
                        f"background-color: {rank_color}; color: #000000;"
                    )

        # Eliminated players: bold red text across the entire player row.
        # `_team_out` is kept outside the displayed dataframe.
        if team_out_flags is not None:
            try:
                is_out = bool(team_out_flags.loc[row.name])
            except (KeyError, TypeError):
                is_out = False
            if is_out:
                for idx in range(len(row)):
                    styles[idx] += "; color: #C00000; font-weight: 700"

        return styles

    styled = df.style.set_properties(
        **{"background-color": "#FFFFFF", "color": "#000000"}
    )
    if zebra:
        def zebra_styles(row):
            pos = df.index.get_loc(row.name)
            color = "#FFFFFF" if pos % 2 == 0 else "#EEF0F3"
            return [f"background-color: {color}"] * len(row)

        styled = styled.apply(zebra_styles, axis=1)
    my_team = st.session_state.get("my_team")
    if my_team and "Team Owner" in df.columns:
        def tint_my_row(row):
            mine = str(row["Team Owner"]).casefold() == str(my_team).casefold()
            return ["background-color: #FFF4E5" if mine else ""] * len(row)

        styled = styled.apply(tint_my_row, axis=1)
    styled = styled.apply(apply_styles, axis=1)

    # Display numeric values cleanly.
    # One-decimal fields show exactly one decimal; integer fields do not
    # expose pandas' trailing floating-point zeros.
    one_decimal_columns = {
        "FPts per game",
        "Avg. FPts/Tourn. Game",
        "Delta From Average",
        "Delta From Season Avg.",
        "Projected FPts",
        "Projected Score",
        "Total FPts",
        "MVP FPts",
        "Expected Games Remaining",
        "Win %",
        "Top 3 %",
        "Last Place %",
        "Avg. Final FPts",
    }
    two_decimal_columns = {"Avg. Finish", "Expected Games Remaining"}
    format_map = {}
    for col in df.columns:
        if col in two_decimal_columns:
            format_map[col] = lambda value: "" if pd.isna(value) else f"{float(value):.2f}"
        elif col in one_decimal_columns:
            format_map[col] = lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
        elif pd.api.types.is_float_dtype(df[col]):
            format_map[col] = lambda value: "" if pd.isna(value) else f"{float(value):g}"

    # Display positive performer deltas with an explicit "+" sign. This is
    # merged into format_map because a second styled.format(dict) call resets
    # the formatting of every column it does not list.
    if "Delta From Season Avg." in df.columns:
        format_map["Delta From Season Avg."] = lambda x: (
            f"+{x:.1f}" if pd.notna(x) and x > 0
            else (f"{x:.1f}" if pd.notna(x) else "")
        )

    if format_map:
        styled = styled.format(format_map)

    # Center all table headers and cell contents while preserving the
    # existing conditional formatting/background colors.
    styled = styled.set_properties(**{"text-align": "center"})
    styled = styled.set_table_styles([
        {"selector": "th", "props": [("text-align", "center")]}
    ])

    # Optional column-specific background colors. These are applied after the
    # row/scale styling so the requested Team View stat groups are actually
    # visible, while leaving the existing scale styling intact elsewhere.
    if column_colors:
        for col, bg in column_colors.items():
            if col in df.columns:
                styled = styled.set_properties(
                    subset=[col],
                    **{"background-color": bg}
                )

    if my_team and "Team Owner" in df.columns:
        def highlight_my_row(row):
            is_mine = str(row["Team Owner"]).casefold() == str(my_team).casefold()
            styles = [""] * len(row)
            if is_mine:
                styles = ["border-top: 2px solid #F28C28; border-bottom: 2px solid #F28C28"] * len(row)
                owner_idx = row.index.get_loc("Team Owner")
                styles[owner_idx] += "; background-color: #FFE0B2; font-weight: 800"
            return styles

        styled = styled.apply(highlight_my_row, axis=1)

    for column, css_list in (cell_styles or {}).items():
        if column in df.columns:
            styled = styled.apply(lambda _col, css=css_list: css, subset=[column])

    return styled


TEAM_VIEW_TAB = "👥 Team View"


def _standings_cell_clicked(select_key="standings_select"):
    event = st.session_state.get(select_key)
    cells = []
    try:
        cells = list(event["selection"]["cells"])
    except (KeyError, TypeError):
        pass
    for cell in cells:
        if isinstance(cell, dict):
            row_pos, col_name = cell.get("row"), cell.get("column")
        else:
            row_pos, col_name = cell[0], cell[1]
        if col_name == "Team Owner" and row_pos is not None:
            owner = st.session_state.get("_owners_" + select_key, [])[int(row_pos)]
            st.session_state["team_owner"] = owner
            st.session_state["_goto_tab"] = "Team View"
            return


def theme_chart(chart):
    """Give an Altair chart dark-mode colours when dark mode is on."""
    if not st.session_state.get("dark_mode"):
        return chart
    text, grid = "#E5E7EB", "#2D3748"
    return (
        chart.configure(background="transparent")
        .configure_view(stroke=None)
        .configure_axis(labelColor=text, titleColor=text, gridColor=grid, domainColor="#4B5563", tickColor="#4B5563")
        .configure_legend(labelColor=text, titleColor=text)
        .configure_text(color=text)
    )


def _owner_link_clicked(owner):
    st.session_state["team_owner"] = owner
    st.session_state["_goto_tab"] = "Team View"


# Reset on every script run so each table's hidden button keys are unique.
_owner_table_count = 0


def _next_owner_table_key():
    global _owner_table_count
    _owner_table_count += 1
    return f"t{_owner_table_count}"


def display_df(df, height=420, scale_columns=None, reverse_columns=None, hide_index=True, rank_rwg=False, pink_last=False, performance_gradient=None, column_colors=None, selectable=False, zebra=False, center=False, owner_links=None, medal_rank=True, html=False, cell_styles=None):
    if df.empty:
        st.info("No data available yet.")
        return

    cell_tips = {}
    for tip_key, tip_col in (
        ("_expected_tip", "Expected Games Remaining"),
        ("_projected_tip", "Projected FPts"),
        ("_finish_tip", "Avg. Finish"),
        ("_score_tip", "Projected Score"),
        ("_win_tip", "Win %"),
    ):
        if tip_key in df.columns:
            if tip_col in df.columns:
                cell_tips[tip_col] = df[tip_key].tolist()
            df = df.drop(columns=[tip_key])

    photo_map = None
    if "_photo" in df.columns and "Player Name" in df.columns:
        photo_map = dict(zip(df["Player Name"], df["_photo"]))
        df = df.drop(columns=["_photo"])

    if photo_map is not None and owner_links is None and "Team Owner" in df.columns:
        owner_links = _next_owner_table_key()
    html_mode = photo_map is not None or bool(owner_links) or html
    # HTML tables only: pair each value with its tooltip text and where the popup fits.
    if html_mode and cell_tips:
        df = df.copy()
        for tip_col, tips in cell_tips.items():
            if tip_col in df.columns:
                last = len(df) - 1
                df[tip_col] = [
                    (value, tip, "above" if position >= 2 else ("below" if last - position >= 2 else "side"))
                    for position, (value, tip) in enumerate(zip(df[tip_col], tips))
                ]

    try:
        _style_sig = (
            int(pd.util.hash_pandas_object(df.astype(str), index=True).sum()),
            tuple(df.columns),
            repr(scale_columns), repr(reverse_columns), pink_last, rank_rwg,
            repr(performance_gradient), repr(column_colors), zebra, repr(cell_styles),
            st.session_state.get("my_team") or "",
        )
    except Exception:
        _style_sig = None
    _style_cache = st.session_state.setdefault("_styled_cache", {})
    _cache_key = (selectable if isinstance(selectable, str) else None, html_mode, height)
    _hit = _style_cache.get(_cache_key)
    if _style_sig is not None and _hit and _hit[0] == _style_sig:
        styled = _hit[1]
    else:
        styled = style_table(
            df,
            scale_columns=scale_columns,
            reverse_columns=reverse_columns,
            pink_last=pink_last,
            rank_rwg=rank_rwg,
            performance_gradient=performance_gradient,
            column_colors=column_colors,
            zebra=zebra,
            medal_rank=medal_rank,
            cell_styles=cell_styles,
        )
        if _style_sig is not None and isinstance(selectable, str):
            _style_cache[_cache_key] = (_style_sig, styled)

    if html_mode:
        if photo_map is not None:
            team_by_player = dict(zip(df["Player Name"], df["Team"])) if "Team" in df.columns else {}

            def player_cell(value):
                url = photo_map.get(value)
                logo_url = get_logo_url(team_by_player.get(value)) if team_by_player.get(value) else None
                logo = (
                    f'<img src="{html_lib.escape(str(logo_url))}" class="player-team-logo" '
                    'loading="lazy" onerror="this.style.visibility=\'hidden\'">'
                    if logo_url else ''
                )
                img = (
                    f'<img src="{html_lib.escape(str(url))}" class="player-photo" '
                    'referrerpolicy="no-referrer" loading="lazy" '
                    'onerror="this.style.visibility=\'hidden\'">'
                    if url else '<span class="player-photo"></span>'
                )
                name, marker_html = str(value), ""
                for marker in ("⚠️", "⛔"):
                    if name.startswith(marker + " "):
                        name = name[len(marker) + 1:]
                        tip = injury_tip_html(name, team_by_player.get(value), marker)
                        marker_html = (
                            f'<span class="cell-tip injury-tip" tabindex="0">{marker}'
                            f'<span class="cell-tip-text">{tip}</span></span> '
                        )
                        break
                return (
                    '<div class="player-cell">' + logo + img
                    + '<span class="name-wrap"><span class="name-full">' + marker_html + html_lib.escape(name) + '</span>'
                    + '<span class="name-short">' + marker_html + html_lib.escape(short_player_name(name)) + '</span></span></div>'
                )

            styled = styled.format(player_cell, subset=["Player Name"], escape=None)
        if owner_links and "Team Owner" in df.columns:
            owners = list(dict.fromkeys(df["Team Owner"].tolist()))
            owner_index = {owner: i for i, owner in enumerate(owners)}

            def owner_cell(value):
                return (
                    f'<a class="owner-link" href="#" data-key="ownerlink_{owner_links}_{owner_index[value]}">'
                    f'{html_lib.escape(str(value))}</a>'
                )

            styled = styled.format(owner_cell, subset=["Team Owner"], escape=None)
        if cell_tips:
            def make_tip_cell(decimals):
                def tip_cell(value):
                    shown, tip, placement = value
                    if shown is None or pd.isna(shown):
                        return ""
                    return (
                        f'<span class="cell-tip {placement}" tabindex="0">'
                        f'{float(shown):.{decimals}f}<span class="cell-tip-text">{tip}</span></span>'
                    )

                return tip_cell

            for tip_col in cell_tips:
                if tip_col in df.columns:
                    styled = styled.format(
                        make_tip_cell(2 if tip_col in ("Avg. Finish", "Expected Games Remaining") else 1),
                        subset=[tip_col], escape=None,
                    )
        if "Team" in df.columns or "Next Opponent" in df.columns:
            def team_cell(value):
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    return ""
                url = get_logo_url(value)
                img = (
                    f'<img src="{html_lib.escape(str(url))}" class="team-logo-sm" '
                    'loading="lazy" onerror="this.style.visibility=\'hidden\'">'
                    if url else '<span class="team-logo-sm"></span>'
                )
                return (
                    '<div class="team-cell">' + img
                    + '<span>' + html_lib.escape(str(value)) + '</span></div>'
                )

            styled = styled.format(
                team_cell,
                subset=[c for c in ("Team", "Next Opponent") if c in df.columns],
                escape=None,
            )
        try:
            styled = styled.hide(axis="index")
        except AttributeError:
            styled = styled.hide_index()

        st.markdown(
            """
            <style>
            .photo-table-wrap {overflow-x: auto;}
            .photo-table-wrap table {border-collapse: collapse; width: 100%; font-size: 14px;}
            .photo-table-wrap th, .photo-table-wrap td {
                border: 1px solid #e0e0e0; padding: 4px 8px; white-space: nowrap;
            }
            .photo-table-wrap th {background: #4472C4 !important; color: #FFFFFF; font-weight: 700;}
            .player-cell {display: flex; align-items: center; position: relative; height: 64px; overflow: hidden; padding-left: 88px; text-align: left; width: 200px; transition: width 0.25s ease, padding-left 0.25s ease;}
            .player-cell span {white-space: normal; min-width: 0; overflow-wrap: anywhere; align-self: center;}
            .photo-table-wrap td:has(.player-cell) {vertical-align: bottom; padding-top: 0 !important; padding-bottom: 0 !important; overflow: visible;}
            .photo-table-wrap a.owner-link, .photo-table-wrap a.owner-link:hover, .photo-table-wrap a.owner-link:visited {color: inherit !important; text-decoration: none !important; border-bottom: 1px dotted #888; cursor: pointer;}
            [class*="st-key-ownerlinks_"] {display: none;}
            .team-cell {display: flex; align-items: center; gap: 8px; min-width: 160px; padding-left: 10px;}
            .team-cell span {white-space: normal; min-width: 0; overflow-wrap: anywhere;}
            .team-logo-sm {width: 36px; height: 36px; object-fit: contain; flex-shrink: 0; display: inline-block;}
            .player-photo {
                position: absolute; left: 0; bottom: 0; top: auto; height: 60px; width: auto; max-width: 84px;
                object-fit: cover; object-position: center bottom;
                background: transparent; display: block; z-index: 1;
                transition: height 0.25s ease, max-width 0.25s ease;
            }
            span.player-photo {width: 1px;}
            .name-wrap {display: grid; align-items: center;}
            .name-full, .name-short {grid-area: 1 / 1; transition: opacity 0.25s ease;}
            .player-cell .name-full {white-space: nowrap; overflow-wrap: normal;}
            .player-cell .name-short {opacity: 0; white-space: nowrap; overflow-wrap: normal; font-size: 12px;}
            .photo-table-wrap.scrolled .name-full {opacity: 0;}
            .photo-table-wrap.scrolled .name-short {opacity: 1;}
            .photo-table-wrap.scrolled .player-cell {width: 128px; padding-left: 50px;}
            .photo-table-wrap td:has(.player-cell) {transition: padding-left 0.25s ease;}
            .photo-table-wrap.scrolled td:has(.player-cell) {padding-left: 0;}
            .photo-table-wrap.scrolled .player-photo {height: 42px; max-width: 48px;}
            .player-team-logo {position: absolute; left: 14px; top: 7px; width: 20px; height: 20px; object-fit: contain; opacity: 0; transition: opacity 0.25s ease; z-index: 2;}
            .photo-table-wrap.scrolled.team-passed .player-team-logo {opacity: 1;}
            .cell-tip {position: relative; display: inline-block; cursor: help; border-bottom: 1px dotted #888; outline: none;}
            .cell-tip.injury-tip {border-bottom: none;}
            .cell-tip .cell-tip-text {
                display: none; position: absolute; right: 0; z-index: 20; width: max-content; min-width: 200px; max-width: 300px;
                white-space: normal; text-align: left; font-size: 12px; font-weight: 400;
                color: #000000; background: #FFFFFF; border: 1px solid #4472C4; border-radius: 6px;
                padding: 8px 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            }
            .cell-tip.above .cell-tip-text {bottom: 100%; margin-bottom: 6px;}
            .cell-tip.below .cell-tip-text {top: 100%; margin-top: 6px;}
            .cell-tip.side .cell-tip-text {right: 100%; top: 50%; transform: translateY(-50%); margin-right: 8px;}
            </style>
            """,
            unsafe_allow_html=True,
        )
        table_html = styled.to_html().replace("\n", "")
        if st.session_state.get("dark_mode"):
            table_html = re.sub(
                "([\U0001F300-\U0001FAFF\u26A0\u26D4\u2B50\u2705\u274C]\uFE0F?)",
                r'<span class="emoji-fix">\1</span>',
                table_html,
            )
        # Side popups are taller than a one-row table, so leave room beneath it.
        room_style = ' style="padding-bottom: 48px;"' if "cell-tip side" in table_html else ""
        # Freeze the Player Name column so it stays visible while scrolling sideways.
        freeze_css, freeze_class = "", ""
        visible_columns = [c for c in df.columns if not str(c).startswith("_")]
        freeze_name = next(
            (c for c in ("Player Name", "Team Owner", "Stat") if c in visible_columns), None
        )
        if freeze_name:
            name_col = visible_columns.index(freeze_name) + 1
            freeze_class += f" freeze-{name_col}"
            freeze_css = (
                f".photo-table-wrap.freeze-{name_col} th:nth-child({name_col}),"
                f".photo-table-wrap.freeze-{name_col} td:nth-child({name_col})"
                "{position: sticky; left: -1px; z-index: 3;}"
                f".photo-table-wrap.freeze-{name_col}.scrolled th:nth-child({name_col}),"
                f".photo-table-wrap.freeze-{name_col}.scrolled td:nth-child({name_col})"
                "{box-shadow: 2px 0 4px rgba(0,0,0,0.15);}"
            )
        st.markdown(
            f'<style>{freeze_css}</style>'
            f'<div class="photo-table-wrap{freeze_class}"{room_style}>{table_html}</div>',
            unsafe_allow_html=True,
        )
        if owner_links and "Team Owner" in df.columns:
            # Hidden buttons that the owner links click, so a click runs a normal Streamlit callback.
            with st.container(key=f"ownerlinks_{owner_links}"):
                for i, owner in enumerate(owners):
                    st.button(
                        str(owner), key=f"ownerlink_{owner_links}_{i}",
                        on_click=_owner_link_clicked, args=(owner,),
                    )
        return

    # Size the table to its actual row count so short tables do not
    # show a large block of blank cells underneath the data.
    row_height = 35
    header_height = 38
    compact_height = header_height + (len(df) * row_height) + 4

    # Give every supplied row enough room to render without a vertical
    # scrollbar. The height argument is retained only as a safety cap for
    # unusually large tables.
    actual_height = compact_height if (compact_height <= height or len(df) <= 60) else height

    if selectable:
        select_key = selectable if isinstance(selectable, str) else "standings_select"
        st.session_state["_owners_" + select_key] = df["Team Owner"].tolist()
        return st.dataframe(
            styled,
            use_container_width=True,
            hide_index=hide_index,
            height=actual_height,
            on_select=lambda k=select_key: _standings_cell_clicked(k),
            selection_mode="single-cell",
            key=select_key,
        )

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=hide_index,
        height=actual_height,
        column_config=(
            {c: st.column_config.Column(alignment="center") for c in df.columns}
            if center else None
        ),
    )


# ---------------------------------------------------------------------
# Public OneDrive helpers
# ---------------------------------------------------------------------
def get_onedrive_share_url():
    """Read the public OneDrive sharing URL from Streamlit Secrets."""
    try:
        if "onedrive" in st.secrets:
            url = st.secrets["onedrive"].get("share_url")
        else:
            url = st.secrets.get("ONEDRIVE_SHARE_URL")
    except (KeyError, AttributeError, FileNotFoundError):
        # No secrets file found - return None to use local mode
        return None

    if not url:
        return None

    url = str(url).strip()
    return url or None


@st.cache_data(ttl=60, show_spinner=False)
def download_onedrive_workbook_from_share(share_url):
    """Download an anonymously shared OneDrive workbook.

    The OneDrive link is an ``Anyone with the link can view`` link, so no
    Microsoft login or Microsoft Graph token is required.  OneDrive can
    return an HTML viewer page for a normal share URL; adding ``download=1``
    asks OneDrive for the underlying file instead.
    """
    from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

    parts = urlsplit(share_url)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["download"] = "1"
    download_url = urlunsplit(
        (parts.scheme, parts.netloc, parts.path, urlencode(query), parts.fragment)
    )

    response = requests.get(
        download_url,
        timeout=90,
        allow_redirects=True,
        headers={
            "User-Agent": "Mozilla/5.0 (Streamlit OneDrive workbook reader)",
        },
    )

    if not response.ok:
        raise RuntimeError(
            f"OneDrive returned HTTP {response.status_code} while downloading the shared workbook."
        )

    content = response.content
    content_type = response.headers.get("Content-Type", "unknown")

    # .xlsx files are ZIP containers and therefore begin with PK.
    if not content.startswith(b"PK"):
        # A public link should be anonymously downloadable. If OneDrive
        # instead returned an HTML page, surface a useful message rather
        # than feeding the HTML to openpyxl.
        preview = content[:200].decode("utf-8", errors="ignore").replace("\n", " ").strip()
        raise RuntimeError(
            "The OneDrive share link did not return the Excel workbook. "
            "Make sure the file is shared as 'Anyone with the link can view' "
            "and that downloads are allowed. "
            f"Content-Type: {content_type}."
            + (f" Response began with: {preview[:120]}" if preview else "")
        )

    temp_path = Path(".streamlit_onedrive_workbook.xlsx")
    temp_path.write_bytes(content)

    # The workbook itself contains the authoritative last-edited timestamp
    # in GameResults!D2. load_data() reads that cell after the download.
    return temp_path, None


# ---------------------------------------------------------------------
# App
# ---------------------------------------------------------------------
st.set_page_config(
    page_title="MM Player Pool Dashboard",
    page_icon="🏀",
    layout="wide",
    initial_sidebar_state="collapsed",
)


st.markdown(
    """
    <style>
    /* Hide Streamlit's top app toolbar/header and footer for a cleaner
       dashboard, especially on mobile. */
    header[data-testid="stHeader"],
    div[data-testid="stToolbar"],
    #MainMenu,
    footer,
    [class*="viewerBadge"],
    [class*="_viewerBadge"],
    [data-testid="stDecoration"],
    a[href*="streamlit.io"][target="_blank"][class*="badge"],
    .stAppDeployButton {
        display: none !important;
        visibility: hidden !important;
    }

    .block-container,
    div[data-testid="stMainBlockContainer"] {padding-top: 0 !important; padding-bottom: 2rem;}
    /* Phones leave safe-area margins outside the page; give them the page colour instead of white. */
    html, body {background: #F4F6FA;}
    /* Keep the main tab bar pinned while scrolling; nested tab lists inside panels are excluded. */
    .st-key-main_tabs [role="tablist"]:not([role="tabpanel"] *) {
        position: sticky; top: 0; z-index: 999; background: #F4F6FA;
    }
    /* Bracket region tabs stick just below the main tab bar (40px tall). */
    .st-key-region_tabs [role="tablist"] {
        position: sticky; top: 40px; z-index: 998; background: #F4F6FA;
    }
    /* Keep the owner name close to its roster table in the head-to-head section. */
    [class*="st-key-h2h_roster_"] {gap: 0.25rem !important;}
    /* Style-only markdown blocks above the tabs would otherwise add a flex gap above the header banner.
       Styler tables embed their own <style>, so tables and anything inside tabs are excluded. */
    div[data-testid="stElementContainer"]:has(style):not(:has(.pool-header, table, canvas)):not([data-testid="stTabs"] *) {display: none !important;}

    /* Collapse zero-height helper iframes (browser-storage scripts) so they
       do not leave a blank bar above the tabs. Kept in layout, not display:none,
       so their scripts keep running. */
    div[data-testid="stElementContainer"]:has(iframe[height="0"]),
    div[data-testid="element-container"]:has(iframe[height="0"]),
    .st-key-restore_my_team,
    div[data-testid="stElementContainer"]:has(iframe[title="streamlit_js_eval.streamlit_js_eval"]) {
        height: 0 !important;
        min-height: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        overflow: hidden !important;
    }
    /* Take the helper out of the layout flow so Streamlit's inter-element gap
       is not applied around it either. */
    div[data-testid="stLayoutWrapper"]:has(.st-key-restore_my_team),
    div[data-testid="stElementContainer"].st-key-restore_my_team {
        position: absolute !important;
        width: 0 !important;
        height: 0 !important;
        overflow: hidden !important;
    }

    /* Pool MVP photo and Most Recent Game logo float over the right side of the native metric card */
    div[data-testid="stColumn"]:has(.mvp-photo-float),
    div[data-testid="column"]:has(.mvp-photo-float),
    div[data-testid="stColumn"]:has(.logo-float),
    div[data-testid="column"]:has(.logo-float) {
        position: relative;
        container-type: inline-size;
    }
    div[data-testid="stElementContainer"]:has(.mvp-photo-float),
    div.element-container:has(.mvp-photo-float),
    div[data-testid="stElementContainer"]:has(.logo-float),
    div.element-container:has(.logo-float) {
        position: absolute;
        right: 14px;
        top: 50%;
        transform: translateY(-50%);
        width: auto !important;
        pointer-events: none;
    }
    .mvp-photo-float {height: 100px; width: auto; max-width: 120px; object-fit: contain; display: block;}
    div[data-testid="stElementContainer"]:has(.mvp-photo-float),
    div.element-container:has(.mvp-photo-float) {
        top: auto;
        bottom: 17px;
        right: 0;
        transform: none;
    }
    .logo-float {height: 80px; width: 80px; object-fit: contain; display: block; margin-right: 12px;}
    @media (max-width: 640px) {
        div[data-testid="stElementContainer"]:has(.mvp-photo-float),
        div.element-container:has(.mvp-photo-float) {
            right: -6px;
        }
    }
    @container (max-width: 340px) {
        .mvp-photo-float, .logo-float {display: none;}
    }
    /* Below ~1650px four cards per row are too narrow (<340px, where photos hide), so use 2 per row;
       below ~900px two per row is also too narrow, so use 1 per row. */
    @media (min-width: 901px) and (max-width: 1650px) {
        div[data-testid="stHorizontalBlock"]:has(.mvp-photo-float) > div[data-testid="stColumn"],
        div[data-testid="stHorizontalBlock"]:has(.logo-float) > div[data-testid="stColumn"],
        .st-key-team_summary div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            flex: 1 1 calc(50% - 1rem) !important;
            min-width: calc(50% - 1rem) !important;
        }
    }
    @media (min-width: 641px) and (max-width: 900px) {
        div[data-testid="stHorizontalBlock"]:has(.mvp-photo-float) > div[data-testid="stColumn"],
        div[data-testid="stHorizontalBlock"]:has(.logo-float) > div[data-testid="stColumn"],
        .st-key-team_summary div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }
    }

    /* Zero-height script frames: take them out of the layout flow so they
       do not add blank gaps between elements. */
    div[data-testid="stElementContainer"]:has(> iframe[height="0"]),
    div[data-testid="stElementContainer"]:has(iframe[height="0"]) {
        position: absolute !important;
        height: 0 !important;
        margin: 0 !important;
        overflow: hidden !important;
    }

    /* Team owner dropdown: tap to choose only, no typing/keyboard */
    input[aria-label="Select a team owner"],
    input[aria-label="Compare with"],
    input[aria-label^="👤 Choose your team"] {
        pointer-events: none !important;
        caret-color: transparent !important;
        user-select: none !important;
    }
    div[data-baseweb="select"]:has(input[aria-label="Select a team owner"]),
    div[data-baseweb="select"]:has(input[aria-label="Compare with"]),
    div[data-baseweb="select"]:has(input[aria-label^="👤 Choose your team"]) {
        cursor: pointer;
    }

    /* Enhanced metric card styling with colored left border and background */
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #e0e0e0;
        border-left: 4px solid #4472C4;
        border-radius: 8px;
        padding: 10px;
        box-shadow: 0 2px 4px rgba(68, 114, 196, 0.08);
        transition: box-shadow 0.3s ease;
    }

    div[data-testid="stMetric"]:hover {
        box-shadow: 0 4px 8px rgba(68, 114, 196, 0.15);
    }

    /* Enhanced title styling */
    h1 {
        color: #2E5C8A;
        border-bottom: 3px solid #F28C28;
        padding-bottom: 0.5rem;
    }

    /* Enhanced header styling with accent colors */
    h2 {
        color: #2E5C8A;
        border-bottom: 2px solid #4472C4;
        padding-bottom: 0.3rem;
    }

    h3 {
        color: #2E5C8A;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Runs before any data loads so the outer page's safe-area margins turn grey as early as possible;
# the script at the bottom of the page applies the same colour again from the dark-mode toggle.
components.html(
    """<script>
    try {
        const query = new URLSearchParams(window.parent.location.search).get('dark_mode');
        const dark = query ? ['true', '1'].includes(query) : window.parent.localStorage.getItem('mmpool_dark_mode') === '1';
        const pageBg = dark ? '#0E1117' : '#F4F6FA';
        let outer = window;
        while (outer !== outer.parent) {
            outer = outer.parent;
            try {
                outer.document.documentElement.style.background = pageBg;
                outer.document.body.style.background = pageBg;
            } catch (e) { break; }
        }
    } catch (e) {}
    </script>""",
    height=0,
)

LOGO_PATH = ASSETS_DIR / "logo.png"
if LOGO_PATH.is_file():
    import base64

    logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
    FONT_PATH = ASSETS_DIR / "freshman.ttf"
    font_face_css = ""
    if FONT_PATH.is_file():
        font_b64 = base64.b64encode(FONT_PATH.read_bytes()).decode()
        font_face_css = (
            "@font-face {font-family:'Freshman';"
            f"src:url(data:font/ttf;base64,{font_b64}) format('truetype');"
            "font-display:swap;}"
        )
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Alfa+Slab+One&family=Barlow:wght@600&display=swap');
        .pool-header-text {{display:none;}}
        div[data-testid="stMain"] {{overflow-x:hidden;}}
        .pool-header {{
            width:100vw;margin-left:calc(50% - 50vw);box-sizing:border-box;
            background:#4e9efd;padding:0.75rem 1rem;
        }}
        {font_face_css}
        @media (min-width: 1000px) {{
            .pool-header-text {{
                display:block;margin-left:1.5rem;
                font-family:'Freshman','Alfa Slab One','Rockwell Extra Bold','Georgia',serif;
                font-weight:400;font-style:normal;letter-spacing:0.02em;
                font-size:clamp(1.8rem,3vw,3rem);line-height:1.1;
                color:#FFFFFF;text-transform:uppercase;text-align:center;
                text-shadow:-1px -1px 0 #000,1px -1px 0 #000,-1px 1px 0 #000,1px 1px 0 #000,
                            0 -1px 0 #000,0 1px 0 #000,-1px 0 0 #000,1px 0 0 #000;
            }}
        }}
        </style>
        <div class="pool-header" style="display:flex;justify-content:center;align-items:center;
                    border-bottom:3px solid #F28C28;padding-bottom:0.75rem;
                    margin-bottom:0.5rem;">
            <img src="data:image/png;base64,{logo_b64}"
                 style="height:clamp(7rem,14vw,11rem);width:auto;max-width:90%;
                        object-fit:contain;">
            <div class="pool-header-text">
                <span style="font-style:italic;">2027 March Madness</span><br>
                <span style="font-family:'Barlow','Arial',sans-serif;font-weight:600;font-style:normal;font-size:1.0025em;letter-spacing:0.06em;">Player Pool Dashboard</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.title("🏀 2027 March Madness Player Pool Dashboard")

# Production mode: read the workbook directly from the public OneDrive
# sharing link. Visitors do NOT need to sign in with Microsoft.
#
# If the share URL is not configured, retain the local uploader flow so the
# app remains convenient to run during development.
workbook_path = Path(DEFAULT_WORKBOOK)
cloud_updated_time = None
onedrive_share_url = get_onedrive_share_url()

if onedrive_share_url:
    try:
        workbook_path, cloud_updated_time = download_onedrive_workbook_from_share(
            onedrive_share_url
        )
    except Exception as exc:
        st.error(f"Could not load the OneDrive workbook: {exc}")
        st.stop()
else:
    with st.sidebar:
        st.header("Data Source")
        uploaded = st.file_uploader(
            "Upload an Excel workbook",
            type=["xlsx", "xlsm"],
            help="For local testing, upload either the empty or populated example workbook.",
        )
        st.caption("Local development mode")

    if uploaded is not None:
        temp_path = Path(".streamlit_uploaded_workbook.xlsx")
        temp_path.write_bytes(uploaded.getvalue())
        workbook_path = temp_path

    if not workbook_path.exists():
        st.warning(
            f"Put `{DEFAULT_WORKBOOK}` in the same folder as `dashboard.py`, "
            "or upload an Excel workbook using the sidebar."
        )
        st.stop()

try:
    _stamp_path = Path(workbook_path)
    _stamp_stat = _stamp_path.stat()
    _data_stamp = (
        str(_stamp_path), _stamp_stat.st_mtime_ns, _stamp_stat.st_size, cloud_updated_time,
    )
    _cached_data = st.session_state.get("_data_cache")
    if _cached_data is not None and _cached_data[0] == _data_stamp:
        data = _cached_data[1]
    else:
        data = load_data(workbook_path, updated_time=cloud_updated_time)
        st.session_state["_data_cache"] = (_data_stamp, data)
except Exception as exc:
    st.error(f"Could not load the workbook: {exc}")
    st.stop()

latest = data["latest_game_result"]

if data["updated_time"] is None:
    updated_text = "Standings Updated: Not yet updated"
else:
    updated_text = (
        "Standings Updated: "
        + data["updated_time"].strftime("%m/%d/%Y %I:%M %p")
    )

# Restore the saved team and theme before rendering anything heavy, so the page draws once.
_need_team = not st.session_state.get("my_team") and not st.query_params.get("my_team")
_need_dark = "dark_mode" not in st.session_state and not st.query_params.get("dark_mode")
if (_need_team or _need_dark) and not st.session_state.get("_my_team_restored"):
    try:
        from streamlit_js_eval import streamlit_js_eval
        _saved = streamlit_js_eval(
            js_expressions="(() => { try { return JSON.stringify({team: localStorage.getItem('mmpool_my_team') || '', dark: localStorage.getItem('mmpool_dark_mode') || ''}); } catch (e) { return '{}'; } })()",
            key="restore_my_team",
        )
    except Exception:
        _saved = "{}"
    if _saved is None:
        st.stop()
    st.session_state["_my_team_restored"] = True
    try:
        _saved = json.loads(_saved or "{}")
    except ValueError:
        _saved = {}
    if _need_team and _saved.get("team") in data["drafters"]:
        st.session_state["my_team"] = _saved["team"]
    if _need_dark and _saved.get("dark") == "1":
        st.session_state["dark_mode"] = True

AUTO_REFRESH_SECONDS = 60


def _refresh_standings():
    get_espn_tournament_games.clear()
    download_onedrive_workbook_from_share.clear()
    st.session_state.pop("_data_cache", None)


updated_col, dark_col, auto_col, refresh_col = st.columns([2, 1, 1.2, 1.4], vertical_alignment="center")
updated_col.caption(updated_text)
dark_mode = dark_col.toggle("🌙 Dark mode", key="dark_mode", bind="query-params")
auto_refresh = auto_col.toggle("⏱ Auto refresh", key="auto_refresh")
if dark_mode:
    # Tables, charts and images are dark via an inverting filter; images are inverted back so photos/logos look normal.
    st.markdown(
        """
        <style>
        html, body {background: #0E1117 !important;}
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {background: #0E1117 !important;}
        [data-testid="stMain"] :is(p, li, label, h1, h2, h3, h4, h5, h6, strong, em) {color: #E5E7EB;}
        [data-testid="stMain"] h1, [data-testid="stMain"] h2, [data-testid="stMain"] h3 {color: #8DB8E8;}
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {color: #9CA3AF !important;}
        div[data-testid="stMetric"] {background-color: #1A2130; border-color: #2D3748; border-left-color: #4472C4;}
        div[data-testid="stMetric"] :is([data-testid="stMetricLabel"], [data-testid="stMetricValue"]) * {color: #E5E7EB;}
        [data-testid="stMarkdownContainer"] {color: #E5E7EB;}
        .st-key-main_tabs [role="tablist"]:not([role="tabpanel"] *) {background: #0E1117 !important;}
        .st-key-region_tabs [role="tablist"] {background: #0E1117 !important;}
        [role="tab"] {color: #CBD5E1;}
        [role="tab"][aria-selected="true"] {color: #4e9efd;}
        [data-testid="stSelectbox"] [role="group"] {background-color: #1E2633 !important; border-color: #374151 !important;}
        [data-testid="stSelectbox"] [role="group"] * {color: #E5E7EB !important;}
        [data-testid="stSelectbox"] input::placeholder {color: #9CA3AF !important; opacity: 1;}
        [role="listbox"], [role="listbox"] [role="option"], [data-testid="stSelectboxVirtualDropdown"] {background-color: #1E2633 !important; color: #E5E7EB !important;}
        [role="listbox"] [role="option"]:hover, [role="listbox"] [aria-selected="true"] {background-color: #2B3647 !important;}
        button[kind="secondary"], [data-testid="stBaseButton-secondary"] {background-color: #1E2633 !important; color: #E5E7EB !important; border-color: #374151 !important;}
        [data-testid="stAlert"] {background-color: #1E2A3A !important;}
        [data-testid="stAlert"] * {color: #E5E7EB !important;}
        [data-testid="stExpander"] details {background-color: #141A25; border-color: #2D3748;}
        [data-testid="stExpander"] summary * {color: #E5E7EB !important;}
        hr {border-color: #374151 !important;}
        .photo-table-wrap, [data-testid="stDataFrame"] {filter: invert(1) hue-rotate(180deg);}
        .photo-table-wrap img, .photo-table-wrap .player-photo, [data-testid="stDataFrame"] img {filter: invert(1) hue-rotate(180deg);}
        .photo-table-wrap .emoji-fix {display: inline-block; filter: invert(1) hue-rotate(180deg);}
        </style>
        """,
        unsafe_allow_html=True,
    )
refresh_col.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
if refresh_col.button("🔄 Refresh Standings", key="refresh_data", use_container_width=True):
    _refresh_standings()
    st.rerun()

if auto_refresh:
    @st.fragment(run_every=AUTO_REFRESH_SECONDS)
    def _auto_refresh_tick():
        now = datetime.now().timestamp()
        last = st.session_state.get("_auto_refresh_last")
        if last is None:
            st.session_state["_auto_refresh_last"] = now
        elif now - last >= AUTO_REFRESH_SECONDS - 1:
            st.session_state["_auto_refresh_last"] = now
            _refresh_standings()
            st.rerun()

    _auto_refresh_tick()
else:
    st.session_state.pop("_auto_refresh_last", None)

standings = make_standings(data)
power = make_power_rankings(data)
win_prob = simulate_win_probabilities(data)
# Simulated projections when the bracket is complete; otherwise the seed-average rankings.
projections = win_prob if win_prob is not None else power
players, player_rows = make_player_rankings(data)
over, over_rows = make_performers(data, True)
under, under_rows = make_performers(data, False)
ratings = make_draft_ratings(data)
# Draft ratings use the simulation with no tournament games, results or injuries counted.
pre_sim = simulate_win_probabilities(data, pre_tournament=True)
if pre_sim is not None and not ratings.empty:
    ratings = pre_sim[
        ["Rank", "Draft Position", "Team Owner", "Avg. Finish", "Projected Score", "Win %", "Top 3 %",
         "_finish_tip", "_score_tip", "_win_tip"]
    ].merge(ratings[["Team Owner", "Regions"]], on="Team Owner", how="left")

# Top-line summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("🧑‍💻 Pool Members", len(data["drafters"]))
c2.metric("📝 Drafted Players", len(data["draft_rows"]))
c3.metric("👥 Players Active", int(standings["Active Players"].sum()) if not standings.empty else 0)
c4.metric("❌ Teams Eliminated", len(data["teams_out"]))

pick_col, _pick_spacer = st.columns([1, 2])
with pick_col:
    my_team_choice = st.selectbox(
        "👤 Choose your team (highlights team in tables, should only have to select once)",
        [""] + data["drafters"],
        index=0,
        key="my_team",
        bind="query-params",
    )

tab0, tab1, tab2, tab5, tab3, tab4, tab6 = st.tabs([
    "🏠 Pool Overview",
    "🏆 Standings",
    "📊 Player Statistics",
    "⛹️‍♂️ Tournament Bracket",
    TEAM_VIEW_TAB,
    "🎯 Draft Recap",
    "ℹ️ League Info",
], key="main_tabs")

# Tabs don't rerun the app, so open a tab programmatically by clicking its button.
_goto_tab = st.session_state.pop("_goto_tab", None)
if _goto_tab:
    st.session_state["_goto_n"] = st.session_state.get("_goto_n", 0) + 1
    components.html(
        """<script>
        // run """ + str(st.session_state["_goto_n"]) + """
        const label = """ + json.dumps(_goto_tab) + """;
        let tries = 0;
        const timer = setInterval(() => {
            const tab = Array.from(window.parent.document.querySelectorAll('[role="tab"], [data-baseweb="tab"]'))
                .find(el => el.textContent.includes(label));
            if (tab) { tab.click(); clearInterval(timer); }
            else if (++tries > 40) clearInterval(timer);
        }, 50);
        </script>""",
        height=0,
    )

with tab0:
    st.subheader("🏟️ Pool Overview")
    st.caption("A quick look at the current state of the pool.")

    # Current leaders and headline stats. These are all derived from the
    # same standings/power-ranking data used elsewhere in the dashboard.
    # After the draft, standings rows exist even before any game stats have
    # been entered. In that state there is no meaningful leader/MVP/projected
    # leader yet, so keep these cards blank rather than attempting to format
    # None as a number.
    leader = None
    if not standings.empty and standings["Total FPts"].notna().any():
        leader_candidates = standings[standings["Total FPts"].notna()]
        if not leader_candidates.empty:
            leader = leader_candidates.iloc[0]

    mvp_row = None
    if not standings.empty:
        mvp_candidates = standings.dropna(subset=["MVP FPts"])
        if not mvp_candidates.empty:
            mvp_row = mvp_candidates.loc[mvp_candidates["MVP FPts"].idxmax()]

    projected_leader = None
    if not projections.empty and projections["Projected Score"].notna().any():
        sort_columns = [c for c in ("Win %", "Projected Score") if c in projections.columns]
        projected_candidates = projections[projections["Projected Score"].notna()].sort_values(
            sort_columns, ascending=False, kind="stable"
        )
        if not projected_candidates.empty:
            projected_leader = projected_candidates.iloc[0]
    active_total = int(standings["Active Players"].sum()) if not standings.empty else 0

    overview_history = make_standings_round_history(data)
    biggest_riser_name = "—"
    biggest_riser_text = "No completed rounds yet"
    biggest_faller_name = "—"
    biggest_faller_text = "No completed rounds yet"
    if not overview_history.empty and len(overview_history) >= 2:
        # Once a tournament round has been completed, compare the latest
        # round's standings rank with the previous completed round.
        previous = overview_history.iloc[-2]
        current = overview_history.iloc[-1]
        changes = []
        for owner in data["drafters"]:
            old_rank = previous.get(owner)
            new_rank = current.get(owner)
            if pd.notna(old_rank) and pd.notna(new_rank):
                changes.append((float(old_rank) - float(new_rank), owner))
        if changes:
            risers = [item for item in changes if item[0] > 0]
            fallers = [item for item in changes if item[0] < 0]
            if risers:
                movement = max(item[0] for item in risers)
                tied_risers = [
                    owner for change, owner in risers
                    if change == movement
                ]
                tied_risers.sort(key=str.casefold)
                biggest_riser_name = ", ".join(tied_risers)
                biggest_riser_text = f"+{int(movement)} spot{'s' if movement != 1 else ''} from {previous['Round']}"
            else:
                biggest_riser_name = "No riser"
                biggest_riser_text = "No one moved up"
            if fallers:
                movement = min(item[0] for item in fallers)
                movement_down = abs(movement)
                tied_fallers = [
                    owner for change, owner in fallers
                    if change == movement
                ]
                tied_fallers.sort(key=str.casefold)
                biggest_faller_name = ", ".join(tied_fallers)
                biggest_faller_text = f"-{int(movement_down)} spot{'s' if movement_down != 1 else ''} from {previous['Round']}"
            else:
                biggest_faller_name = "No faller"
                biggest_faller_text = "No one moved down"
    else:
        # Before the first round is complete, use the pre-tournament Draft
        # Ratings rank as the baseline and compare it with the current
        # tournament standings rank. Positive movement means a team is ranked higher now.
        draft_rank_by_owner = {}
        if not ratings.empty and {"Team Owner", "Rank"}.issubset(ratings.columns):
            draft_rank_by_owner = {
                str(row["Team Owner"]).casefold(): row["Rank"]
                for _, row in ratings.iterrows()
            }

        changes = []
        if not standings.empty and {"Team Owner", "Rank"}.issubset(standings.columns):
            for _, row in standings.iterrows():
                owner = str(row["Team Owner"])
                draft_rank = draft_rank_by_owner.get(owner.casefold())
                current_rank = row["Rank"]
                if pd.notna(draft_rank) and pd.notna(current_rank):
                    changes.append((float(draft_rank) - float(current_rank), owner))

        if changes:
            risers = [item for item in changes if item[0] > 0]
            fallers = [item for item in changes if item[0] < 0]
            if risers:
                movement = max(item[0] for item in risers)
                tied_risers = sorted(
                    [owner for change, owner in risers if change == movement],
                    key=str.casefold,
                )
                biggest_riser_name = ", ".join(tied_risers)
                biggest_riser_text = f"+{int(movement)} spot{'s' if movement != 1 else ''} vs. draft rating"
            else:
                biggest_riser_name = "No riser"
                biggest_riser_text = "No one ranked higher than draft rating"

            if fallers:
                movement = min(item[0] for item in fallers)
                movement_down = abs(movement)
                tied_fallers = sorted(
                    [owner for change, owner in fallers if change == movement],
                    key=str.casefold,
                )
                biggest_faller_name = ", ".join(tied_fallers)
                biggest_faller_text = f"- {int(movement_down)} spot{'s' if movement_down != 1 else ''} vs. draft rating"
            else:
                biggest_faller_name = "No faller"
                biggest_faller_text = "No one ranked lower than draft rating"

    c1, c2, c3, c4 = st.columns(4)
    if leader is not None:
        c1.metric("🏆 Current Leader", str(leader["Team Owner"]), f"{float(leader['Total FPts']):.1f} FPts")
    else:
        c1.metric("🏆 Current Leader", "—")
    if projected_leader is not None:
        projected_pill = f"{float(projected_leader['Projected Score']):.1f} Proj. FPts"
        if "Win %" in projected_leader.index and pd.notna(projected_leader["Win %"]):
            projected_pill = f"{float(projected_leader['Win %']):.1f}% win · {projected_pill}"
        c2.metric("🔮 Projected Winner", str(projected_leader["Team Owner"]), projected_pill)
    else:
        c2.metric("🔮 Projected Winner", "—")
    if mvp_row is not None:
        c3.metric("⭐ Pool MVP", str(mvp_row["MVP"]), f"{float(mvp_row['MVP FPts']):.1f} FPts")
        mvp_stats = data["player_stats"].get(str(mvp_row["MVP"]).casefold(), {})
        mvp_photo = get_player_photo_url(mvp_row["MVP"], mvp_stats.get("Team"))
        if mvp_photo:
            c3.markdown(
                f'<img class="mvp-photo-float" src="{html_lib.escape(mvp_photo)}" '
                'referrerpolicy="no-referrer" alt="">',
                unsafe_allow_html=True,
            )
    else:
        c3.metric("⭐ Pool MVP", "—")
    if latest["Winner"] and latest["Loser"]:
        # Use Streamlit's native metric card so this card has the exact same
        # height, padding, and responsive behavior as the other overview cards.
        # A leading minus sign makes the delta render with the red down arrow;
        # Streamlit strips the sign from the displayed text.
        c4.metric(
            "🏀 Most Recent Game Input",
            latest["Winner"],
            f"-- def. {latest['Loser']} [ {latest['Round']} ] --",
        )
        game_logo = get_logo_url(latest["Winner"])
        if game_logo:
            c4.markdown(
                f'<img class="logo-float" src="{html_lib.escape(game_logo)}" '
                'referrerpolicy="no-referrer" alt="">',
                unsafe_allow_html=True,
            )
    else:
        c4.metric("🏀 Most Recent Game", "No game recorded")

    st.markdown("### 💸 Current Payouts 💰")
    if not standings.empty:
        payout_rows = []

        payout_base = data["payout_config"]["D2"]
        last_payout = payout_base * data["payout_config"]["E2"]
        third_payout = payout_base * data["payout_config"]["F2"]
        second_payout = payout_base * data["payout_config"]["G2"]
        first_payout = (
            payout_base * len(data["drafters"])
            - last_payout
            - third_payout
            - second_payout
        )
        payout_by_rank = {
            1: first_payout,
            2: second_payout,
            3: third_payout,
            len(standings): last_payout,
        }

        # Show the top three teams plus the team currently in last place.
        payout_indices = list(range(min(3, len(standings))))
        if len(standings) > 3:
            payout_indices.append(len(standings) - 1)

        rank_emojis = {1: "🥇", 2: "🥈", 3: "🥉", (len(standings)):"🦨"}
        for index in payout_indices:
            row = standings.iloc[index]
            rank = int(row["Rank"])
            owner = row["Team Owner"]
            fpts = row["Total FPts"]

            change_text = "—"
            if not overview_history.empty and len(overview_history) >= 2:
                previous = overview_history.iloc[-2]
                current = overview_history.iloc[-1]
                old_rank = previous.get(owner)
                new_rank = current.get(owner)
                if pd.notna(old_rank) and pd.notna(new_rank):
                    movement = float(old_rank) - float(new_rank)
                    if movement > 0:
                        change_text = f"↑ {int(movement)}"
                    elif movement < 0:
                        change_text = f"↓ {int(abs(movement))}"

            payout_rows.append({
                "Rank": f"{rank_emojis.get(rank, rank)} {rank}",
                "Payout": f"${payout_by_rank.get(rank, 0):,.0f}",
                "Team Owner": owner,
                "FPts": f"{float(fpts):.1f}" if pd.notna(fpts) else "",
                "Change": change_text,
            })

        payout_display = pd.DataFrame(payout_rows)
        st.caption("Click a team owner's name to open their Team View.")
        display_df(payout_display, height=210, zebra=True, owner_links="payout", medal_rank=False)
    # Biggest riser / faller / highest team round / highest individual game
    c5, c6, c7, c8 = st.columns(4)
    if biggest_riser_name != "—":
        c5.metric("📈 Biggest Riser This Round", biggest_riser_name, biggest_riser_text)
    else:
        c5.metric("📈 Biggest Riser This Round", "—", biggest_riser_text)

    if biggest_faller_name != "—":
        c6.metric("📉 Biggest Faller This Round", biggest_faller_name, biggest_faller_text)
    else:
        c6.metric("📉 Biggest Faller This Round", "—", biggest_faller_text)

    # Highest single-round team score: sum the FPts from all five players
    # on each draftee's team for every tournament round.
    highest_round_name = "—"
    highest_round_text = "No round scores recorded yet"
    round_names = [
        "First Round",
        "Second Round",
        "Sweet Sixteen",
        "Elite Eight",
        "Final Four",
        "National Title",
    ]

    round_totals = {
        owner: [0.0] * len(round_names)
        for owner in data["drafters"]
    }

    for row in data["draft_rows"]:
        owner = row["draftee"]
        pts_rounds = row.get("tourn_pts_rounds", [None] * 6)
        reb_rounds = row.get("tourn_reb_rounds", [None] * 6)
        ast_rounds = row.get("tourn_ast_rounds", [None] * 6)

        for i in range(len(round_names)):
            round_fpts = weighted_fantasy_points(
                pts_rounds[i],
                reb_rounds[i],
                ast_rounds[i],
                data["fantasy_weights"],
            )
            if round_fpts is not None:
                round_totals.setdefault(
                    owner, [0.0] * len(round_names)
                )[i] += float(round_fpts)

    round_scores = [
        (score, owner, round_names[i])
        for owner, scores in round_totals.items()
        for i, score in enumerate(scores)
        if score > 0
    ]

    if round_scores:
        highest_score, highest_owner, highest_round = max(
            round_scores, key=lambda x: (x[0], x[1].casefold())
        )
        highest_round_name = highest_owner
        highest_round_text = f"{highest_score:.1f} FPts [ {highest_round} ]"

    c7.metric("🧨 Highest Single-Round Score", highest_round_name, highest_round_text)

    # Highest single-game individual performance. Each player's tournament
    # round stat line represents that player's performance in that round/game.
    highest_game_player = "—"
    highest_game_text = "No game scores recorded yet"
    individual_game_scores = []

    for row in data["draft_rows"]:
        pts_rounds = row.get("tourn_pts_rounds", [None] * 6)
        reb_rounds = row.get("tourn_reb_rounds", [None] * 6)
        ast_rounds = row.get("tourn_ast_rounds", [None] * 6)

        for i, round_name in enumerate(round_names):
            game_fpts = weighted_fantasy_points(
                pts_rounds[i],
                reb_rounds[i],
                ast_rounds[i],
                data["fantasy_weights"],
            )
            if game_fpts is not None:
                individual_game_scores.append(
                    (float(game_fpts), row["player_name"], round_name)
                )

    if individual_game_scores:
        highest_game_score, highest_game_player_name, highest_game_round = max(
            individual_game_scores,
            key=lambda x: (x[0], x[1].casefold())
        )
        highest_game_player = highest_game_player_name
        highest_game_text = f"{highest_game_score:.1f} FPts [ {highest_game_round} ]"

    c8.metric(
        "🔥 Highest Single-Game Performance",
        highest_game_player,
        highest_game_text,
    )
    if individual_game_scores:
        game_stats = data["player_stats"].get(highest_game_player.casefold(), {})
        game_photo = get_player_photo_url(highest_game_player, game_stats.get("Team"))
        if game_photo:
            c8.markdown(
                f'<img class="mvp-photo-float" src="{html_lib.escape(game_photo)}" '
                'referrerpolicy="no-referrer" alt="">',
                unsafe_allow_html=True,
            )


    if latest["Winner"] and latest["Loser"]:
        def _title_logo(team):
            url = get_logo_url(team)
            return (
                f'<img src="{html_lib.escape(str(url))}" referrerpolicy="no-referrer" '
                'style="height:1.3em;width:auto;vertical-align:-0.2em;margin:0 0.2em;" '
                'onerror="this.style.display=\'none\'">'
                if url else ""
            )

        st.subheader("📋 Most Recent Statistics Entered")
        st.markdown(
            "<div style='font-size:1.25rem; font-weight:700;'>"
            f"{_title_logo(latest['Winner'])} {html_lib.escape(str(latest['Winner']))}  def  "
            f"{_title_logo(latest['Loser'])} {html_lib.escape(str(latest['Loser']))} "
            f"[ {html_lib.escape(str(latest['Round']))} ]</div>",
            unsafe_allow_html=True,
        )
        game_round = round_lost_index(latest["Round"])
        game_teams = {str(latest["Winner"]).strip().casefold(), str(latest["Loser"]).strip().casefold()}
        recent_rows = []
        if game_round is not None and game_round >= 0:
            for row in data["draft_rows"]:
                stats = data["player_stats"].get(row["player_name"].casefold(), {})
                team = stats.get("Team")
                if team is None or str(team).strip().casefold() not in game_teams:
                    continue
                pts = (row.get("tourn_pts_rounds") or [None] * 6)[game_round]
                reb = (row.get("tourn_reb_rounds") or [None] * 6)[game_round]
                ast = (row.get("tourn_ast_rounds") or [None] * 6)[game_round]
                fpts = weighted_fantasy_points(pts, reb, ast, data["fantasy_weights"])
                if fpts is None:
                    continue
                recent_rows.append({
                    "Player Name": player_label(row["player_name"], stats),
                    "_photo": get_player_photo_url(row["player_name"], team),
                    "Team": team,
                    "Next Opponent": next_opponent_for(data, team),
                    "FPts": float(fpts),
                    "PTS": pts, "REB": reb, "AST": ast,
                    "Team Owner": row["draftee"],
                })
        if recent_rows:
            winner_key = str(latest["Winner"]).strip().casefold()
            recent_df = pd.DataFrame(recent_rows)
            recent_df["_loser"] = recent_df["Team"].astype(str).str.strip().str.casefold() != winner_key
            recent_df = recent_df.sort_values(
                ["_loser", "FPts"], ascending=[True, False], kind="stable"
            ).drop(columns=["_loser"])
            recent_df = recent_df[["Player Name", "_photo", "Team", "Next Opponent", "Team Owner", "FPts", "PTS", "REB", "AST"]]
            for col in ("FPts", "PTS", "REB", "AST"):
                recent_df[col] = recent_df[col].apply(
                    lambda x: f"{float(x):.1f}" if col == "FPts" and pd.notna(x)
                    else (f"{float(x):g}" if pd.notna(x) else "")
                )
            display_df(recent_df, height=min(80 + 38 * len(recent_df), 600), zebra=True)
        else:
            st.info("No drafted players appeared in the most recent game.")

    st.subheader("🚀 Standings Progression")
    st.caption("League ranking after each tournament round.")
    standings_history = make_standings_round_history(data)
    if not standings_history.empty:
        # Plot ranking rather than points so the chart is much easier to read.
        # Reverse the Y axis so Rank 1 appears at the top of the chart.
        chart_data = standings_history.melt(
            id_vars="Round",
            var_name="Team Owner",
            value_name="Rank",
        )
        chart_data["Round"] = pd.Categorical(
            chart_data["Round"],
            categories=[
                "First Round",
                "Second Round",
                "Sweet Sixteen",
                "Elite Eight",
                "Final Four",
                "National Title",
            ],
            ordered=True,
        )
        try:
            import altair as alt
            round_order = [
                "First Round", "Second Round", "Sweet Sixteen",
                "Elite Eight", "Final Four", "National Title"
            ]
            owner_order = list(data["drafters"])
            rank_max = max(len(owner_order), int(chart_data["Rank"].max()))

            # Generate exactly one deterministic, unique color per pool member.
            # The palette is sized from the actual number of pool members, so
            # it automatically expands/shrinks with the pool.
            def make_owner_colors(owners):
                count = len(owners)
                if count == 0:
                    return []

                colors = []
                used = set()

                for i in range(count):
                    # Evenly distribute hues around the color wheel.
                    hue = i / count
                    red, green, blue = colorsys.hsv_to_rgb(hue, 0.72, 0.82)
                    color = "#{:02X}{:02X}{:02X}".format(
                        round(red * 255),
                        round(green * 255),
                        round(blue * 255),
                    )

                    # Guard against an accidental RGB rounding collision.
                    # This keeps the guarantee of one unique hex color per owner.
                    if color in used:
                        step = 1 / max(count * 1000, 1)
                        attempt = 1
                        while color in used:
                            adjusted_hue = (hue + step * attempt) % 1.0
                            red, green, blue = colorsys.hsv_to_rgb(
                                adjusted_hue, 0.72, 0.82
                            )
                            color = "#{:02X}{:02X}{:02X}".format(
                                round(red * 255),
                                round(green * 255),
                                round(blue * 255),
                            )
                            attempt += 1

                    used.add(color)
                    colors.append(color)

                return colors

            owner_colors = make_owner_colors(owner_order)

            base = alt.Chart(chart_data).encode(
                x=alt.X(
                    "Round:N",
                    sort=round_order,
                    title=None,
                ),
                y=alt.Y(
                    "Rank:Q",
                    scale=alt.Scale(domain=[1, rank_max], reverse=True),
                    title=None,
                    axis=alt.Axis(values=list(range(1, rank_max + 1))),
                ),
                color=alt.Color(
                    "Team Owner:N",
                    title=None,
                    sort=owner_order,
                    legend=None,
                    scale=alt.Scale(domain=owner_order, range=owner_colors),
                ),
                tooltip=alt.value(None),
            )

            lines = base.mark_line()
            points = base.mark_point(filled=True, size=55)
            my_team_name = st.session_state.get("my_team")
            is_mine = f"datum['Team Owner'] == {json.dumps(my_team_name)}" if my_team_name else None
            if is_mine:
                lines = lines.encode(
                    strokeWidth=alt.condition(is_mine, alt.value(5), alt.value(2)),
                    opacity=alt.condition(is_mine, alt.value(1), alt.value(0.45)),
                )

            # Put a small rounded label behind every point so the owner's name
            # is visible without having to hover. The label background uses
            # the same series color as the corresponding line, while the
            # owner name is always white and bold for maximum contrast.
            label_bg = base.mark_rect(
                width=54,
                height=18,
                cornerRadius=5,
                opacity=0.92,
                strokeWidth=0,
            )
            if is_mine:
                label_bg = label_bg.encode(
                    stroke=alt.condition(is_mine, alt.value("#F28C28"), alt.value("transparent")),
                    strokeWidth=alt.condition(is_mine, alt.value(2.5), alt.value(0)),
                )
            labels = base.mark_text(
                fontSize=9,
                fontWeight="bold",
                baseline="middle",
            ).encode(
                text="Team Owner:N",
                color=alt.value("#FFFFFF"),
            )

            chart = alt.layer(lines, points, label_bg, labels).properties(height=500)
            st.altair_chart(theme_chart(chart), use_container_width=True)
        except ImportError:
            # Fallback if Altair is unavailable in the deployment environment.
            st.line_chart(
                standings_history.set_index("Round"),
                height=450,
                use_container_width=True,
            )

with tab1:
    st.subheader("🏆 Tournament Standings")
    standings_event = display_df(
        standings.rename(columns={"Draft Position": "Draft Pos."}),
        height=520,
        pink_last=True,
        zebra=True,
        owner_links="standings",
        scale_columns={
            "Draft Pos.": (
                standings["Draft Position"].min(),
                standings["Draft Position"].max(),
            ),
            "Total FPts": (
                standings["Total FPts"].min(),
                standings["Total FPts"].max(),
            ),
            "Active Players": (
                standings["Active Players"].min(),
                standings["Active Players"].max(),
            ),
            "Poss. Games Remaining": (
                standings["Poss. Games Remaining"].min(),
                standings["Poss. Games Remaining"].max(),
            ),
            "MVP FPts": (
                standings["MVP FPts"].min(),
                standings["MVP FPts"].max(),
            ),
        },
        reverse_columns={"Draft Pos.": True},
    )

    st.caption("Click a team owner's name to open their Team View.")
    st.session_state.pop("jump_to_team_view", None)
    st.subheader("🔮 Projected Rankings & Win Probability 🎲")
    if win_prob is None:
        st.info("Win probabilities need a complete bracket in the TeamList sheet; showing seed-average projections.")
        display_df(
            power.rename(columns={"Draft Position": "Draft Pos."}),
            height=520,
            rank_rwg=True,
            zebra=True,
            owner_links="power",
            scale_columns={
                "Draft Pos.": (power["Draft Position"].min(), power["Draft Position"].max()),
                "Active Players": (power["Active Players"].min(), power["Active Players"].max()),
                **{
                    col: (power[col].min(), power[col].max())
                    for col in ["Projected Score", "FPts per game", "Games Played", "Expected Games Remaining"]
                    if col in power.columns
                },
            },
            reverse_columns={"Draft Pos.": True},
        )
    else:
        rating_info = win_prob.attrs.get("ratings") or {}
        if rating_info.get("matched"):
            snapshot_note = (
                f"; ratings snapshot from {rating_info['source'].split(':', 1)[1] or 'an earlier date'}"
                if str(rating_info.get("source", "")).startswith("snapshot") else ""
            )
            odds_text = (
                f"Game odds use Barttorvik {rating_info['year']} adjusted offense, defense and tempo "
                f"({rating_info['matched']} of {rating_info['total']} bracket teams matched{snapshot_note})"
            )
            if rating_info.get("net_matched"):
                odds_text += (
                    f", blended with the NCAA NET rankings (through {rating_info.get('net_through') or 'latest games'}; "
                    f"{rating_info['net_matched']} teams matched)"
                )
        else:
            reason = rating_info.get("error")
            odds_text = "Game odds use seed-based win probabilities" + (
                f" (Barttorvik ratings unavailable: {reason})" if reason else ""
            )
        net_text = ""
        if rating_info.get("matched") and not rating_info.get("net_matched"):
            net_text = (
                " The NCAA NET rankings are published but not used in these projections (they appear in the "
                "League Info ratings table)."
                if rating_info.get("net_through") else
                " The NCAA NET rankings are not yet being used because they have not been published yet."
            )
        st.caption(
            f"Estimated from {WIN_PROB_SIMS:,} simulated tournaments. {odds_text} and real results so far.{net_text} "
            f"Player scoring varies about {GAME_FPTS_SPREAD:.0%} "
            "per game around the projected rate. Rank orders owners by average simulated finishing place "
            "(Avg. Finish); Projected Score is the average simulated final FPts. "
            "Hover over Avg. Finish, Projected Score or Win % for details."
        )
        display_df(
            win_prob.drop(columns=["FPts per game", "Games Played", "Exp. Payout"]).rename(columns={"Draft Position": "Draft Pos."}),
            height=520,
            rank_rwg=True,
            zebra=True,
            owner_links="winprob",
            scale_columns={
                "Draft Pos.": (win_prob["Draft Position"].min(), win_prob["Draft Position"].max()),
                **{
                    col: (win_prob[col].min(), win_prob[col].max())
                    for col in [
                        "Avg. Finish", "Win %", "Top 3 %", "Last Place %", "Projected Score", "Active Players",
                        "Expected Games Remaining",
                    ]
                },
            },
            reverse_columns={"Draft Pos.": True, "Avg. Finish": True, "Last Place %": True},
        )


with tab2:
    st.markdown(
        "<h3 style='margin:0;padding:0 0 0.25rem 0;'>Player Finder</h3>",
        unsafe_allow_html=True,
    )
    selected_player = st.selectbox(
        "Search for a player",
        [""] + data["player_names"],
        index=0,
    )

    if selected_player:
        matching = [
            row for row in data["draft_rows"]
            if row["player_name"].casefold() == selected_player.casefold()
        ]
        if matching:
            row = matching[0]
            stats = data["player_stats"].get(selected_player.casefold(), {})
            result = pd.DataFrame([{
                "Player Name": player_label(selected_player, stats),
                "_photo": get_player_photo_url(selected_player, stats.get("Team")),
                "Draft Pos.": row["draft_position"],
                "Draftee": row["draftee"],
                "Team": stats.get("Team"),
                "Next Opponent": next_opponent_for(data, stats.get("Team")),
                "Seed": stats.get("Seed"),
                "FPts": stats.get("Tot. FPts"),
                "PPG": stats.get("PPG"),
                "RPG": stats.get("RPG"),
                "APG": stats.get("APG"),
                "Tourn. FPts": player_total(row, data["fantasy_weights"]),
                "Tourn. PTS": row.get("tourn_pts"),
                "Tourn. REB": row.get("tourn_reb"),
                "Tourn. AST": row.get("tourn_ast"),
                "Projected FPts": player_projected_fpts(row, data),
                "_projected_tip": projected_fpts_explanation(row, data),
                "Poss. Games Remaining": row.get("possible_games_remaining", 0),
                "Expected Games Remaining": round(player_expected_games_remaining(row, data), 2),
                "_expected_tip": expected_games_explanation(row, data),
                "_team_out": (
                    stats.get("Team") is not None
                    and str(stats.get("Team")).strip().casefold() in data["teams_out"]
                ),
            }])
            for column in ["FPts", "PPG", "RPG", "APG", "Tourn. FPts"]:
                if column in result.columns:
                    result[column] = result[column].apply(
                        lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
                    )
            st.caption(f"Showing Player Finder statistics for {selected_player}.")
            display_df(
                result,
                height=90,
                zebra=True,
                column_colors={
                    "FPts": LIGHT_BLUE,
                    "PPG": LIGHT_BLUE,
                    "RPG": LIGHT_BLUE,
                    "APG": LIGHT_BLUE,
                    "Tourn. FPts": LIGHT_GREEN,
                    "Tourn. PTS": LIGHT_GREEN,
                    "Tourn. REB": LIGHT_GREEN,
                    "Tourn. AST": LIGHT_GREEN,
                },
            )
        else:
            st.info(f"{selected_player} was not drafted.")
    st.markdown(
        "<hr style='margin:0.25rem 0 0.5rem 0;border:0;border-top:1px solid #d9dbe0;'>",
        unsafe_allow_html=True,
    )

    st.subheader("Player Rankings — Top 10")
    st.caption("⛔ Out · ⚠️ Questionable")
    display_df(players, height=400, zebra=True)

    st.markdown("### <span style='color:#2E8B57; font-weight:800;'>↑</span> Top Overperformers 🦄", unsafe_allow_html=True)
    over_display = over.rename(columns={"Delta From Average": "Delta From Season Avg."}).copy()
    under_display = under.rename(columns={"Delta From Average": "Delta From Season Avg."}).copy()

    # Format these two dashboard columns explicitly as display strings.
    # This avoids Streamlit/pandas rendering the underlying floats with
    # six decimal places on some pandas/Streamlit versions.
    for display_df_ in (over_display, under_display):
        if "Avg. FPts/Tourn. Game" in display_df_.columns:
            display_df_["Avg. FPts/Tourn. Game"] = display_df_["Avg. FPts/Tourn. Game"].apply(
                lambda x: f"{float(x):.1f}" if pd.notna(x) else ""
            )

    display_df(
        over_display,
        height=240,
        zebra=True,
        performance_gradient="green",
    )

    st.markdown("### <span style='color:#2E8B57; font-weight:800;'>↑</span> Biggest Draft Steals 💎", unsafe_allow_html=True)
    st.caption(
        "Pick Value = overall pick number minus rank by Projected FPts. "
        "A bigger positive number means the player was drafted much later than they are projecting."
    )
    display_df(
        make_draft_value(data, steals=True),
        height=240,
        zebra=True,
        performance_gradient="green",
    )

    st.markdown("### <span style='color:#C62828; font-weight:800;'>↓</span> Top Underperformers 🪠", unsafe_allow_html=True)
    if st.checkbox("Hide Injured players", key="hide_unavailable_under"):
        under_filtered, _ = make_performers(data, False, exclude_unavailable=True)
        under_display = under_filtered.rename(columns={"Delta From Average": "Delta From Season Avg."}).copy()
        if "Avg. FPts/Tourn. Game" in under_display.columns:
            under_display["Avg. FPts/Tourn. Game"] = under_display["Avg. FPts/Tourn. Game"].apply(
                lambda x: f"{float(x):.1f}" if pd.notna(x) else ""
            )
    display_df(
        under_display,
        height=240,
        zebra=True,
        performance_gradient="red",
    )

    st.markdown("### <span style='color:#C62828; font-weight:800;'>↓</span> Biggest Draft Busts 🦨", unsafe_allow_html=True)
    hide_injured_busts = st.checkbox("Hide Injured players", key="hide_unavailable_busts")
    display_df(
        make_draft_value(data, steals=False, exclude_unavailable=hide_injured_busts),
        height=240,
        zebra=True,
        performance_gradient="red",
    )
    st.markdown(
        "<hr style='margin:0.25rem 0 0.5rem 0;border:0;border-top:1px solid #d9dbe0;'>",
        unsafe_allow_html=True,
    )
    st.subheader("🚑 Injured Players 🩼")
    st.caption("⛔ Out · ⚠️ Questionable")
    unavailable_df, _ = make_player_rankings(data, only_unavailable=True)
    unavailable_df = unavailable_df.drop(columns=["Rank"], errors="ignore")
    if unavailable_df.empty:
        st.info("No drafted players are currently listed as questionable or out.")
    else:
        display_df(
            unavailable_df,
            height=min(400, 40 + 35 * len(unavailable_df)),
            zebra=True,
        )

with tab3:
    st.subheader("Team View")

    my_team_now = st.session_state.get("my_team") or ""
    last_my_team = st.session_state.get("_last_my_team", "")
    if my_team_now != last_my_team:
        current_owner = st.session_state.get("team_owner") or ""
        if current_owner in ("", last_my_team):
            st.session_state["team_owner"] = my_team_now
        st.session_state["_last_my_team"] = my_team_now

    selected_draftee = st.selectbox(
        "Select a team owner",
        [""] + data["drafters"],
        index=0,
        key="team_owner",
    )
    components.html(
        """<script>
        const fix = () => {
            window.parent.document.querySelectorAll('input[aria-label="Select a team owner"]').forEach(el => {
                el.setAttribute('inputmode', 'none');
                el.setAttribute('readonly', 'readonly');
            });
        };
        fix();
        new MutationObserver(fix).observe(window.parent.document.body, {childList: true, subtree: true});
        </script>""",
        height=0,
    )

    if selected_draftee:
        team_rows = make_team_rows(data, selected_draftee)
        team_df = pd.DataFrame(team_rows)

        # Team View only: display FPts, PPG, RPG, APG, and Tournament FPts
        # to exactly one decimal place.
        for column in ["FPts", "PPG", "RPG", "APG", "Tourn. FPts"]:
            if column in team_df.columns:
                team_df[column] = team_df[column].apply(
                    lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
                )

        if team_df.empty:
            st.info("No drafted players found for this team owner.")
        else:
            owner_standing = standings[
                standings["Team Owner"].astype(str).str.casefold() == selected_draftee.casefold()
            ]
            if not owner_standing.empty:
                standing_row = owner_standing.iloc[0]
                total_fpts = standing_row["Total FPts"]
                owner_rank = standing_row["Rank"]
            else:
                total_fpts, owner_rank = None, None

            scored_players = [
                r for r in team_rows
                if r.get("Tourn. FPts") is not None and not pd.isna(r.get("Tourn. FPts"))
            ]
            best_player = max(scored_players, key=lambda r: float(r["Tourn. FPts"]), default=None)
            alive_count = sum(1 for r in team_rows if not r["_team_out"])

            st.markdown(
                "<style>.st-key-team_summary div[data-testid='stMetric'] "
                "{min-height: 130px; box-sizing: border-box;}</style>",
                unsafe_allow_html=True,
            )
            changes = make_team_round_changes(data, selected_draftee)
            fpts_delta = changes["fpts_delta"]
            rank_change = changes["rank_change"]
            active_change = changes["active_change"]
            rank_baseline = changes["round_name"]
            if rank_change is None and owner_rank is not None and pd.notna(owner_rank) and not ratings.empty:
                draft_match = ratings[ratings["Team Owner"].astype(str).str.casefold() == selected_draftee.casefold()]
                if not draft_match.empty:
                    rank_change = int(draft_match.iloc[0]["Rank"]) - int(owner_rank)
                    rank_baseline = "draft rating"
            with st.container(key="team_summary"):
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric(
                    "🧮 Total FPts",
                    f"{float(total_fpts):.1f}" if total_fpts is not None and pd.notna(total_fpts) else "—",
                    (f"{fpts_delta:+.1f} from {changes['round_name']}" if round(fpts_delta, 1) != 0 else "no change")
                    if fpts_delta is not None else None,
                    delta_color="normal" if fpts_delta is not None and round(fpts_delta, 1) != 0 else "off",
                )
                sc2.metric(
                    "📊 Pool Rank",
                    f"{int(owner_rank)} of {len(standings)}" if owner_rank is not None and pd.notna(owner_rank) else "—",
                    (f"{'-' if rank_change < 0 else ''}{abs(rank_change)} spot{'s' if abs(rank_change) != 1 else ''} from {rank_baseline}"
                     if rank_change else "no change")
                    if rank_change is not None else None,
                    delta_color="normal" if rank_change else "off",
                )
                if best_player is not None:
                    sc3.metric(
                        "🌟 Team MVP",
                        str(best_player["Player Name"]),
                        f"{float(best_player['Tourn. FPts']):.1f} FPts",
                    )
                    best_name = re.sub(r"^(?:⚠️|⛔)\s*", "", str(best_player["Player Name"]))
                    best_stats = data["player_stats"].get(best_name.casefold(), {})
                    best_photo = get_player_photo_url(best_name, best_stats.get("Team"))
                    if best_photo:
                        sc3.markdown(
                            f'<img class="mvp-photo-float" src="{html_lib.escape(best_photo)}" '
                            'referrerpolicy="no-referrer" alt="">',
                            unsafe_allow_html=True,
                        )
                else:
                    sc3.metric("🌟 Team MVP", "—")
                sc4.metric(
                    "👥 Players Active",
                    f"{alive_count} of {len(team_rows)}",
                    (f"{active_change:+d} from {changes['round_name']}" if active_change else "no change")
                    if active_change is not None else None,
                    delta_color="normal" if active_change else "off",
                )

            st.caption(
                f"Showing {len(team_df)} player(s) drafted by {selected_draftee}.  "
                "⛔ Out · ⚠️ Questionable"
            )
            display_df(
                team_df.rename(columns={"Draft Position": "Draft Pos."}),
                height=500,
                zebra=True,
                column_colors={
                    "FPts": LIGHT_BLUE,
                    "PPG": LIGHT_BLUE,
                    "RPG": LIGHT_BLUE,
                    "APG": LIGHT_BLUE,
                    "Tourn. FPts": LIGHT_GREEN,
                    "Tourn. PTS": LIGHT_GREEN,
                    "Tourn. REB": LIGHT_GREEN,
                    "Tourn. AST": LIGHT_GREEN,
                },
            )

            chart_rounds = [
                "First Round", "Second Round", "Sweet Sixteen",
                "Elite Eight", "Final Four", "National Title",
            ]
            chart_rows = []
            for draft_row in data["draft_rows"]:
                if draft_row["draftee"].casefold() != selected_draftee.casefold():
                    continue
                pts = draft_row.get("tourn_pts_rounds") or [None] * 6
                reb = draft_row.get("tourn_reb_rounds") or [None] * 6
                ast = draft_row.get("tourn_ast_rounds") or [None] * 6
                for i, round_name in enumerate(chart_rounds):
                    fpts = weighted_fantasy_points(pts[i], reb[i], ast[i], data["fantasy_weights"])
                    if fpts is not None:
                        chart_rows.append({
                            "Round": round_name,
                            "Player": draft_row["player_name"],
                            "FPts": float(fpts),
                        })

            st.subheader("FPts by Round")
            if chart_rows:
                round_chart_df = pd.DataFrame(chart_rows)
                try:
                    import altair as alt

                    round_chart_df = round_chart_df.sort_values(["Round", "Player"])
                    round_chart_df["End"] = round_chart_df.groupby("Round")["FPts"].cumsum()
                    round_chart_df["Start"] = round_chart_df["End"] - round_chart_df["FPts"]
                    round_chart_df["Mid"] = (round_chart_df["Start"] + round_chart_df["End"]) / 2
                    round_chart = alt.Chart(round_chart_df).mark_bar().encode(
                        x=alt.X("Round:N", sort=chart_rounds, title=None),
                        y=alt.Y("Start:Q", title="FPts"),
                        y2="End:Q",
                        color=alt.Color("Player:N", legend=alt.Legend(orient="bottom", title=None, columns=2, labelLimit=140, symbolLimit=0)),
                        tooltip=alt.value(None),
                    )
                    round_labels = alt.Chart(round_chart_df).mark_text(
                        dy=-8, fontWeight="bold"
                    ).encode(
                        x=alt.X("Round:N", sort=chart_rounds),
                        y=alt.Y("sum(FPts):Q"),
                        text=alt.Text("sum(FPts):Q", format=".1f"),
                        tooltip=alt.value(None),
                    )
                    segment_labels = alt.Chart(
                        round_chart_df[round_chart_df["FPts"] >= 3]
                    ).mark_text(color="white", fontSize=11).encode(
                        x=alt.X("Round:N", sort=chart_rounds),
                        y=alt.Y("Mid:Q"),
                        text=alt.Text("FPts:Q", format=".1f"),
                        tooltip=alt.value(None),
                    )
                    st.altair_chart(
                        theme_chart((round_chart + segment_labels + round_labels).properties(height=350)),
                        use_container_width=True,
                    )
                except ImportError:
                    st.bar_chart(
                        round_chart_df.groupby("Round")["FPts"].sum().reindex(chart_rounds).dropna()
                    )
            else:
                st.info("No tournament FPts recorded for this team yet.")

            st.subheader("⚔️ Head-to-Head Compare")
            rival = st.selectbox(
                "Compare with",
                [""] + [o for o in data["drafters"] if o.casefold() != selected_draftee.casefold()],
                index=0,
                key="h2h_rival",
            )
            if rival:
                mine_totals, mine_best = team_round_fpts(data, selected_draftee)
                rival_totals, rival_best = team_round_fpts(data, rival)

                def edge(a, b, higher=True):
                    if a is None or b is None:
                        return "—"
                    if a == b:
                        return "Tie"
                    return selected_draftee if (a > b) == higher else rival

                compare_rows = []
                compare_stats = [
                    ("Pool Rank", standings, "Rank", False, "{:.0f}"),
                    ("Total FPts", standings, "Total FPts", True, "{:.1f}"),
                ]
                if win_prob is not None:
                    compare_stats.append(("Win %", win_prob, "Win %", True, "{:.1f}%"))
                compare_stats += [
                    ("Projected Score", projections, "Projected Score", True, "{:.1f}"),
                    ("Draft Rating Rank", ratings, "Rank", False, "{:.0f}"),
                    ("Active Players", standings, "Active Players", True, "{:.0f}"),
                    ("Expected Games Remaining", projections, "Expected Games Remaining", True, "{:.2f}"),
                    ("FPts per Game", projections, "FPts per game", True, "{:.1f}"),
                ]
                for label, source, column, higher, fmt in compare_stats:
                    a = owner_value(source, selected_draftee, column)
                    b = owner_value(source, rival, column)
                    compare_rows.append({
                        "Stat": label,
                        selected_draftee: fmt.format(a) if a is not None else "—",
                        rival: fmt.format(b) if b is not None else "—",
                        "Edge": edge(a, b, higher),
                    })
                compare_rows.append({
                    "Stat": "Team MVP",
                    selected_draftee: f"{mine_best[0]} ({mine_best[1]:.1f})" if mine_best[0] else "—",
                    rival: f"{rival_best[0]} ({rival_best[1]:.1f})" if rival_best[0] else "—",
                    "Edge": edge(mine_best[1], rival_best[1]),
                })
                compare_rows.append({
                    "Stat": "Best Round",
                    selected_draftee: f"{max(mine_totals):.1f}" if max(mine_totals) > 0 else "—",
                    rival: f"{max(rival_totals):.1f}" if max(rival_totals) > 0 else "—",
                    "Edge": edge(max(mine_totals) or None, max(rival_totals) or None),
                })
                my_team_name = (st.session_state.get("my_team") or "").casefold()
                win_style = "background-color: #CDEBD3; font-weight: 700"
                compare_styles = {selected_draftee: [], rival: [], "Edge": []}
                for compare_row in compare_rows:
                    leader = compare_row["Edge"]
                    compare_styles[selected_draftee].append(win_style if leader == selected_draftee else "")
                    compare_styles[rival].append(win_style if leader == rival else "")
                    compare_styles["Edge"].append(
                        "background-color: #FFE0B2; font-weight: 800"
                        if my_team_name and leader.casefold() == my_team_name else ""
                    )
                wins = {selected_draftee: 0, rival: 0}
                for compare_row in compare_rows:
                    if compare_row["Edge"] in wins:
                        wins[compare_row["Edge"]] += 1
                st.caption(
                    f"{selected_draftee} leads {wins[selected_draftee]} categor{'y' if wins[selected_draftee] == 1 else 'ies'}, "
                    f"{rival} leads {wins[rival]}."
                )
                display_df(pd.DataFrame(compare_rows), height=420, zebra=True, html=True, cell_styles=compare_styles)

                played = [i for i in range(6) if mine_totals[i] > 0 or rival_totals[i] > 0]
                if played:
                    last_played = max(played)
                    rounds_shown = chart_rounds[: last_played + 1]
                    cumulative = []
                    for owner, totals in ((selected_draftee, mine_totals), (rival, rival_totals)):
                        running = 0.0
                        for i, round_name in enumerate(rounds_shown):
                            running += totals[i]
                            cumulative.append({"Round": round_name, "Owner": owner, "Cumulative FPts": running})
                    cumulative_df = pd.DataFrame(cumulative)
                    st.markdown("**Cumulative FPts by Round**")
                    try:
                        import altair as alt

                        cumulative_chart = alt.Chart(cumulative_df).mark_line(point=True).encode(
                            x=alt.X("Round:N", sort=chart_rounds, title=None, axis=alt.Axis(labelAngle=0)),
                            y=alt.Y("Cumulative FPts:Q", title="Cumulative FPts"),
                            color=alt.Color(
                                "Owner:N",
                                legend=alt.Legend(orient="bottom", title=None),
                                scale=alt.Scale(domain=[selected_draftee, rival], range=["#4472C4", "#F28C28"]),
                            ),
                            tooltip=alt.value(None),
                        ).properties(height=300)
                        st.altair_chart(theme_chart(cumulative_chart), use_container_width=True)
                    except ImportError:
                        st.line_chart(cumulative_df.pivot(index="Round", columns="Owner", values="Cumulative FPts").reindex(rounds_shown))

                    per_round_edges = [
                        edge(mine_totals[i], rival_totals[i]) for i in range(last_played + 1)
                    ]
                    per_round_df = pd.DataFrame({
                        "Round": rounds_shown,
                        selected_draftee: [f"{mine_totals[i]:.1f}" for i in range(last_played + 1)],
                        rival: [f"{rival_totals[i]:.1f}" for i in range(last_played + 1)],
                        "Edge": per_round_edges,
                    })
                    per_round_styles = {
                        selected_draftee: [win_style if e == selected_draftee else "" for e in per_round_edges],
                        rival: [win_style if e == rival else "" for e in per_round_edges],
                        "Edge": [
                            "background-color: #FFE0B2; font-weight: 800"
                            if my_team_name and e.casefold() == my_team_name else ""
                            for e in per_round_edges
                        ],
                    }
                    display_df(per_round_df, height=300, zebra=True, html=True, cell_styles=per_round_styles)

                roster_columns = ["Player Name", "_photo", "Team", "Tourn. FPts", "Projected FPts", "_projected_tip", "_team_out"]
                roster_cols = st.columns(2)
                for roster_col, owner in zip(roster_cols, (selected_draftee, rival)):
                    with roster_col, st.container(key=f"h2h_roster_{'a' if owner == selected_draftee else 'b'}"):
                        st.markdown(f"**{owner}**")
                        roster_df = pd.DataFrame(make_team_rows(data, owner))
                        if roster_df.empty:
                            st.info("No drafted players.")
                            continue
                        roster_df = roster_df[roster_columns]
                        roster_df["Tourn. FPts"] = roster_df["Tourn. FPts"].apply(
                            lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
                        )
                        display_df(roster_df, height=400, zebra=True)
            else:
                st.info("Pick a team owner above to compare against.")

with tab4:
    st.subheader("Draft Ratings")
    if "Avg. Finish" in ratings.columns:
        st.caption(
            f"Owners ranked by average finishing place across {WIN_PROB_SIMS:,} simulated tournaments that ignore "
            "all games, results and injuries entered so far, so this is the draft-day baseline. "
            "Click a team owner's name to open their Team View."
        )
    else:
        st.caption(
            "Projected FPts = seed average games × season-average FPts, summed by draftee. "
            "Click a team owner's name to open their Team View."
        )
    rating_metric_columns = [
        c for c in ("Avg. Finish", "Projected Score", "Win %", "Top 3 %", "Projected FPts") if c in ratings.columns
    ]
    display_df(
        ratings.rename(columns={"Draft Position": "Draft Pos."}),
        height=520,
        rank_rwg=True,
        zebra=True,
        owner_links="ratings",
        scale_columns={
            "Draft Pos.": (
                ratings["Draft Position"].min(),
                ratings["Draft Position"].max(),
            ),
            **{col: (ratings[col].min(), ratings[col].max()) for col in rating_metric_columns},
        },
        reverse_columns={"Draft Pos.": True, "Avg. Finish": True},
    )

    st.subheader("Drafted Players")
    st.caption("Listed in draft order.  ⛔ Out · ⚠️ Questionable")
    all_player_rows = []
    for drafter in data["drafters"]:
        for player_row in make_team_rows(data, drafter):
            all_player_rows.append({"Team Owner": drafter, **player_row})
    all_player_rows.sort(key=lambda r: draft_position_sort_key(r["Draft Position"]))
    all_players_df = pd.DataFrame(all_player_rows)
    if all_players_df.empty:
        st.info("No drafted players found.")
    else:
        all_players_df = all_players_df[[
            "Draft Position", "Player Name", "_photo", "Team", "Next Opponent", "Team Owner", "Seed",
            "FPts", "PPG", "RPG", "APG", "Tourn. FPts", "Tourn. PTS", "Tourn. REB",
            "Tourn. AST", "Projected FPts", "Poss. Games Remaining",
            "Expected Games Remaining", "_expected_tip", "_projected_tip", "_team_out",
        ]].rename(columns={"Draft Position": "Draft Pos."})
        for column in ["FPts", "PPG", "RPG", "APG", "Tourn. FPts"]:
            all_players_df[column] = all_players_df[column].apply(
                lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
            )
        display_df(
            all_players_df,
            height=500,
            zebra=True,
            column_colors={
                "FPts": LIGHT_BLUE,
                "PPG": LIGHT_BLUE,
                "RPG": LIGHT_BLUE,
                "APG": LIGHT_BLUE,
                "Tourn. FPts": LIGHT_GREEN,
                "Tourn. PTS": LIGHT_GREEN,
                "Tourn. REB": LIGHT_GREEN,
                "Tourn. AST": LIGHT_GREEN,
            },
        )

with tab6:
    st.subheader("ℹ️ League Info")

    st.markdown("### FPts Per Statistic")
    w_pts, w_reb, w_ast = data["fantasy_weights"]
    fc1, fc2, fc3 = st.columns(3)
    fc1.metric("PTS", f"{w_pts:g} FPts")
    fc2.metric("REB", f"{w_reb:g} FPts")
    fc3.metric("AST", f"{w_ast:g} FPts")

    st.markdown("### Payouts")
    cfg = data["payout_config"]
    buy_in = cfg["D2"]
    last_payout = buy_in * cfg["E2"]
    third_payout = buy_in * cfg["F2"]
    second_payout = buy_in * cfg["G2"]
    # First place gets whatever remains of the total pot.
    first_payout = buy_in * len(data["drafters"]) - last_payout - third_payout - second_payout
    buy_in_bg, pot_bg = ("#16301F", "#12273F") if st.session_state.get("dark_mode") else ("#E6F4EA", "#E3F0FB")
    st.markdown(
        "<style>.st-key-buy_in_card div[data-testid='stMetric'] "
        f"{{background-color: {buy_in_bg}; border-left-color: #2E8B57;}}"
        ".st-key-total_pot_card div[data-testid='stMetric'] "
        f"{{background-color: {pot_bg}; border-left-color: #2F80C9;}}</style>",
        unsafe_allow_html=True,
    )
    pc1, pc2, pc3, pc4, pc5, pc6 = st.columns(6)
    with pc1:
        with st.container(key="buy_in_card"):
            st.metric("Buy-in", f"${buy_in:,.0f}")
    with pc2:
        with st.container(key="total_pot_card"):
            st.metric("Total Pot", f"${buy_in * len(data['drafters']):,.0f}")
    pc3.metric("🥇 1st Place", f"${first_payout:,.0f}")
    pc4.metric("🥈 2nd Place", f"${second_payout:,.0f}")
    pc5.metric("🥉 3rd Place", f"${third_payout:,.0f}")
    pc6.metric("🦨 Last Place", f"${last_payout:,.0f}")

    st.markdown("### FPts Projection Calculations")
    tourn_w = data["tournament_fpts_weight"]
    st.markdown(
        f"**PROJECTIONS = Average Results of {WIN_PROB_SIMS:,} Simulated Tournaments**"
    )
    st.markdown(
        "Each simulated tournament plays out the bracket game by game. Results already entered are locked in; "
        "only games not yet played are simulated."
    )
    st.markdown("**HOW A GAME IS DECIDED** (Team ratings from [BartTorvik](https://barttorvik.com))")
    st.markdown(
        "1. **Net rating** = a team's Adjusted Offense minus Adjusted Defense, in points per 100 possessions.  \n"
        "2. **Expected possessions** = the two teams' Adjusted Tempos multiplied together, divided by the "
        "league-average tempo.  \n"
        "3. **Expected margin** = [ Team A Net Rating - Team B Net Rating ] x [ Expected Possessions ] / 100. "
        "Tournament games are at neutral sites, so there is no home-court adjustment.  \n"
        f"4. **Win probability** = a normal curve on that margin, with a standard deviation of {GAME_MARGIN_SD:g} "
        "points. A game is then decided by a random draw at that probability.  \n"
        "5. Any team without a Barttorvik rating falls back to a rough seed-based strength."
    )
    st.markdown(
        "*Ex. a team with a +10 net rating against a team with a +4 net rating over 68 expected possessions: "
        f"[ 10 - 4 ] x 68 / 100 = +4.1 points, which is about a {0.5 * (1 + math.erf(4.08 / (GAME_MARGIN_SD * math.sqrt(2)))):.0%} "
        "win probability.*"
    )
    st.markdown(
        f"Every player scores their blended FPts per game [ Season Avg. FPts x {1 - tourn_w:g} ] + "
        f"[ Tourn. Avg. FPts x {tourn_w:g} ] in each game their team plays (season average alone before their "
        f"first tournament game), with about {GAME_FPTS_SPREAD:.0%} game-to-game variation. Players listed as "
        "out score no more FPts."
    )
    st.markdown(
        "**PLAYER PROJECTED FPTS** = FPts Already Scored + Average Simulated FPts Still To Come"
    )
    st.markdown(
        "**EXPECTED GAMES REMAINING** = Average Number of Additional Games the Player's Team Plays in the Simulations "
        "(so it depends on the team's simulated opponents and win probabilities above; 0 for players who are out "
        "or whose team is eliminated)"
    )
    st.markdown(
        "**PROJECTED SCORE** = Average Final FPts for a Team Owner  \n"
        "**AVG. FINISH** = Average Finishing Place  \n"
        "**WIN %** = Share of Simulations Where the Owner Finished 1st"
    )
    st.markdown(
        "**DRAFT RATINGS** use the same simulation with no tournament games, results or injuries counted, "
        "so they show the draft-day baseline."
    )
    st.caption(
        "Average number of tournament games played by each seed (historical; only used for teams that "
        "aren't in the bracket)."
    )
    seed_games_df = pd.DataFrame(
        [
            {"Seed": seed, "Avg. Games Played": f"{games:.2f}"}
            for seed, games in data["seed_average_games"].items()
        ]
    )
    seed_table_col, _seed_spacer = st.columns([1, 2])
    with seed_table_col:
        display_df(seed_games_df, height=620, center=True)

    if data["bracket_teams"]:
        st.markdown("### Team Ratings: Barttorvik & NCAA NET")
        rating_year = torvik_season_year()
        torvik_ratings, _torvik_error, torvik_source = get_torvik_ratings(rating_year)
        net_ranks, net_through = load_net_ranks(rating_year)
        net_by_key = net_ranks_by_key(torvik_ratings, net_ranks)
        rating_rows = []
        for seed, region_teams in data["bracket_teams"].items():
            for region, teams in region_teams.items():
                for team in teams:
                    info = torvik_team_info(team)
                    net_rank = net_by_key.get(rating_key(torvik_ratings, team))
                    rating_rows.append({
                        "_sort": info["rank"] if info and info["rank"] else 9999,
                        "Seed": seed, "Region": region.title(), "Team": team,
                        "Torvik Rank": info["rank"] if info else "—",
                        "Record": info["record"] if info and info["record"] else "—",
                        "Adj. O": f"{info['o']:.1f}" if info else "—",
                        "Adj. D": f"{info['d']:.1f}" if info else "—",
                        "Net": f"{info['net']:+.1f}" if info else "—",
                        "Tempo": f"{info['t']:.1f}" if info else "—",
                        "NET Rank": net_rank if net_rank else "—",
                    })
        if rating_rows:
            ratings_df = pd.DataFrame(rating_rows).sort_values(["_sort", "Seed"]).drop(columns="_sort")
            source_note = (
                f"Barttorvik ratings snapshot from {torvik_source.split(':', 1)[1] or 'an earlier date'}"
                if torvik_source.startswith("snapshot") else "Barttorvik ratings (live)"
            ) if torvik_source else "Barttorvik ratings unavailable"
            net_note = (
                f"NCAA NET rankings through {net_through}." if net_ranks else
                f"The {rating_year} NCAA NET rankings haven't been published yet; the NET column fills in once "
                "they are released and fetched."
            )
            st.caption(
                f"Bracket teams ordered by Torvik rank. {source_note}: Adj. O and Adj. D are points per 100 "
                f"possessions, Net = Adj. O - Adj. D, Tempo = possessions per game. {net_note}"
            )
            display_df(ratings_df.reset_index(drop=True), height=620, zebra=True, center=True, html=True)

with tab5:
    st.subheader("Tournament Bracket")
    st.caption(
        "March Madness tournament bracket based on teams entered in the TeamList sheet. "
        "Percentages show each team's chance to win games that haven't been played yet."
    )

    if data["bracket_teams"]:
        try:
            # Build the bracket with teams_out data to determine winners
            rounds = build_tournament_bracket(data["bracket_teams"], data["teams_out"])

            # Create subtabs for regions and final four
            region_tabs = st.tabs(["🔵 EAST", "🟤 WEST", "🟢 MIDWEST", "🟡 SOUTH", "🏆 FINAL FOUR"], key="region_tabs")

            # Show each region's bracket (exclude Final Four and Championship)
            regions = ["EAST", "WEST", "MIDWEST", "SOUTH"]
            regional_rounds = rounds[:-2]  # All rounds except Final Four and Championship

            for tab_idx, (region_tab, region) in enumerate(zip(region_tabs[:4], regions)):
                with region_tab:
                    # Show Elite Eight winner box at the top
                    elite_eight_matchups = regional_rounds[-1]["regions"][region]
                    if elite_eight_matchups:
                        matchup = elite_eight_matchups[0]
                        team_a = matchup["team_a"]
                        team_b = matchup["team_b"]

                        # Determine winner
                        team_a_lost = lost_in_round(team_a, 3, data["teams_out"])
                        team_b_lost = lost_in_round(team_b, 3, data["teams_out"])

                        winner = None
                        if team_a_lost and not team_b_lost:
                            winner = team_b
                        elif team_b_lost and not team_a_lost:
                            winner = team_a
                        elif not team_a_lost and not team_b_lost:
                            winner = None  # Game not played yet

                        # Display winner box at the top
                        col1, col2, col3 = st.columns([1.3, 1, 1.3])
                        with col2:
                            components.html(
                                render_regional_winner_card_html(region, winner, dark=bool(st.session_state.get("dark_mode"))),
                                height=90,
                            )

                    st.markdown("---")

                    # Show bracket
                    bracket_html = render_bracket_html(regional_rounds, region=region, team_players=build_team_players(data), my_team=st.session_state.get("my_team"), dark=bool(st.session_state.get("dark_mode")))
                    components.html(bracket_html, height=1200, scrolling=True)
            # Final Four tab - show only Final Four and Championship
            with region_tabs[4]:
                # Compute regional winners
                regional_winners = {}
                for region in regions:
                    elite_eight_matchups = regional_rounds[-1]["regions"][region]
                    if elite_eight_matchups:
                        matchup = elite_eight_matchups[0]
                        team_a = matchup["team_a"]
                        team_b = matchup["team_b"]

                        # Determine winner
                        team_a_lost = lost_in_round(team_a, 3, data["teams_out"])
                        team_b_lost = lost_in_round(team_b, 3, data["teams_out"])

                        winner = None
                        if team_a_lost and not team_b_lost:
                            winner = team_b
                        elif team_b_lost and not team_a_lost:
                            winner = team_a
                        # else: winner is None (TBD)

                        regional_winners[region] = winner if winner else "TBD"

                # Build Final Four bracket with configured region pairings
                final_four_rounds = build_final_four_brackets(
                    regional_winners, 
                    data["final_four_config"], 
                    data["teams_out"]
                )
                champion = final_four_rounds[-1]["regions"]["FINAL_FOUR"][0].get("winner")
                col1, col2, col3 = st.columns([1.3, 1, 1.3])
                with col2:
                    components.html(
                        render_regional_winner_card_html(
                            "", champion, label="NATIONAL CHAMPION", dark=bool(st.session_state.get("dark_mode"))
                        ),
                        height=90,
                    )

                st.markdown("---")

                bracket_html = render_bracket_html(final_four_rounds, team_players=build_team_players(data), my_team=st.session_state.get("my_team"), dark=bool(st.session_state.get("dark_mode")))
                components.html(bracket_html, height=550, scrolling=False)
        except Exception as e:
            st.error(f"Error rendering bracket: {str(e)}")
            st.write(f"Details: {e}")
    else:
        st.warning("No teams found in the TeamList sheet.")
        st.info("Make sure your Excel file has a 'TeamList' sheet with columns:\n- **Region** (EAST, WEST, MIDWEST, SOUTH)\n- **Seed** (1-16)\n- **Team** (team name)")

st.divider()
st.caption(
    "The dashboard reads the Excel workbook from OneDrive. "
    "No workbook changes are made by the app."
)

components.html(
    """<script>
    const fixMyTeam = () => {
        window.parent.document.querySelectorAll('input[aria-label^="👤 Choose your team"], input[aria-label="Compare with"]').forEach(el => {
            el.setAttribute('inputmode', 'none');
            el.setAttribute('readonly', 'readonly');
        });
    };
    fixMyTeam();
    new MutationObserver(fixMyTeam).observe(window.parent.document.body, {childList: true, subtree: true});
    </script>""",
    height=0,
)

components.html(
    """<script>
    const KEY = 'mmpool_my_team';
    const current = """ + json.dumps(my_team_choice or "") + """;
    const dark = """ + json.dumps(bool(dark_mode)) + """;
    try {
        const p = window.parent;
        p.localStorage.removeItem('mmpool_auto_refresh');
        if (dark) p.localStorage.setItem('mmpool_dark_mode', '1');
        else p.localStorage.removeItem('mmpool_dark_mode');
        if (current) {
            p.localStorage.setItem(KEY, current);
            p.__mmpoolLast = current;
        } else if (p.__mmpoolLast) {
            p.localStorage.removeItem(KEY);
            p.__mmpoolLast = '';
        }
    } catch (e) {}
    // On Streamlit Cloud the app sits in an outer page that paints the phone's safe-area margins white.
    const pageBg = dark ? '#0E1117' : '#F4F6FA';
    let outer = window;
    while (outer !== outer.parent) {
        outer = outer.parent;
        try {
            outer.document.documentElement.style.background = pageBg;
            outer.document.body.style.background = pageBg;
        } catch (e) { break; }
    }
    </script>""",
    height=0,
)

# Player tables: once scrolled sideways, mark the table so the frozen name column can shrink.
components.html(
    """<script>
    const DARK = """ + json.dumps(bool(dark_mode)) + """;
    try {
        const p = window.parent, d = p.document;
        if (p.__mmFreezeOff) p.__mmFreezeOff();
        // One shared tooltip on the page body, so no table scroll container can clip it.
        let tip = d.getElementById('mm-cell-tip');
        if (!tip) { tip = d.createElement('div'); tip.id = 'mm-cell-tip'; d.body.appendChild(tip); }
        tip.style.cssText = 'position:fixed;display:none;z-index:99999;pointer-events:none;width:max-content;'
            + 'min-width:200px;max-width:300px;white-space:normal;text-align:left;font-size:12px;font-weight:400;'
            + 'line-height:1.4;border-radius:6px;padding:8px 10px;box-shadow:0 4px 12px rgba(0,0,0,0.3);'
            + 'border:1px solid #4472C4;' + (DARK ? 'background:#1A2130;color:#E5E7EB;' : 'background:#FFFFFF;color:#000000;');
        const hideTip = () => { tip.style.display = 'none'; };
        const showTip = el => {
            const src = el.querySelector('.cell-tip-text');
            if (!src) return;
            tip.innerHTML = src.innerHTML;
            tip.style.display = 'block';
            const r = el.getBoundingClientRect();
            const tw = tip.offsetWidth, th = tip.offsetHeight, gap = 8;
            let top = r.top - th - gap;
            if (top < gap) top = r.bottom + gap;
            if (top + th > p.innerHeight - gap) top = Math.max(gap, p.innerHeight - th - gap);
            const left = Math.min(Math.max(gap, r.right - tw), p.innerWidth - tw - gap);
            tip.style.top = top + 'px';
            tip.style.left = left + 'px';
        };
        const onTipOver = e => { const el = e.target.closest && e.target.closest('.cell-tip'); if (el) showTip(el); };
        const onTipOut = e => {
            const el = e.target.closest && e.target.closest('.cell-tip');
            if (el && !el.contains(e.relatedTarget)) hideTip();
        };
        d.addEventListener('mouseover', onTipOver);
        d.addEventListener('mouseout', onTipOut);
        d.addEventListener('focusin', onTipOver);
        d.addEventListener('focusout', onTipOut);
        d.addEventListener('scroll', hideTip, true);
        const onOwnerClick = e => {
            const a = e.target.closest && e.target.closest('a.owner-link');
            if (!a) return;
            e.preventDefault();
            const btn = d.querySelector('.st-key-' + a.dataset.key + ' button');
            if (btn) btn.click();
        };
        d.addEventListener('click', onOwnerClick, true);
        const onScroll = e => {
            const el = e.target;
            if (!el || !el.classList || !el.classList.contains('photo-table-wrap')) return;
            const scrolled = el.classList.contains('scrolled');
            // Only shrink when there is plenty of sideways room; shrinking a barely-too-wide table removes its scroll range and flickers.
            if (!scrolled && el.scrollLeft > 24 && el.scrollWidth - el.clientWidth > 110) el.classList.add('scrolled');
            else if (scrolled && el.scrollLeft < 4) el.classList.remove('scrolled');
            const heads = Array.from(el.querySelectorAll('thead th'));
            const nameTh = heads.find(h => h.textContent.trim() === 'Player Name');
            const teamIdx = heads.findIndex(h => h.textContent.trim() === 'Team');
            const teamLogo = teamIdx >= 0 ? el.querySelector('tbody tr td:nth-child(' + (teamIdx + 1) + ') .team-logo-sm') : null;
            const passed = !!(nameTh && teamLogo && el.classList.contains('scrolled')
                && teamLogo.getBoundingClientRect().right <= nameTh.getBoundingClientRect().right);
            el.classList.toggle('team-passed', passed);
        };
        d.addEventListener('scroll', onScroll, true);
        p.__mmFreezeOff = () => {
            d.removeEventListener('scroll', onScroll, true);
            d.removeEventListener('click', onOwnerClick, true);
            d.removeEventListener('mouseover', onTipOver);
            d.removeEventListener('mouseout', onTipOut);
            d.removeEventListener('focusin', onTipOver);
            d.removeEventListener('focusout', onTipOut);
            d.removeEventListener('scroll', hideTip, true);
        };
    } catch (e) {}
    </script>""",
    height=0,
)

# Switching tabs shrinks the page and the browser snaps to the top; keep the tab bar where it was.
components.html(
    """<script>
    try {
        const p = window.parent;
        if (!p.__mmTabAnchor2) {
            p.__mmTabAnchor2 = true;
            const s = p.document.createElement('script');
            s.textContent = `
            (function () {
                const TABS = '[role="tablist"], [data-baseweb="tab-list"]';
                const TAB = '[role="tab"], [data-baseweb="tab"]';
                const BOX = '[data-testid="stMainBlockContainer"], .block-container';
                let index = -1, offset = 0, scroller = null, until = 0, raf = 0;
                const lists = () => Array.from(document.querySelectorAll(TABS));
                const scrollParent = el => {
                    for (let n = el.parentElement; n; n = n.parentElement) {
                        const o = getComputedStyle(n).overflowY;
                        if ((o === 'auto' || o === 'scroll') && n.scrollHeight > n.clientHeight) return n;
                    }
                    return document.scrollingElement;
                };
                const stop = () => { until = 0; cancelAnimationFrame(raf); };
                const tick = () => {
                    const list = lists()[index];
                    if (!list || Date.now() > until) return stop();
                    const diff = list.getBoundingClientRect().top - offset;
                    if (Math.abs(diff) > 1) {
                        const want = scroller.scrollTop + diff;
                        scroller.scrollTop = want;
                        const short = want - scroller.scrollTop;
                        if (short > 1) {
                            const box = document.querySelector(BOX);
                            if (box) { box.style.minHeight = (box.offsetHeight + short) + 'px'; scroller.scrollTop = want; }
                        }
                    }
                    raf = requestAnimationFrame(tick);
                };
                document.addEventListener('pointerdown', e => {
                    const tab = e.target.closest && e.target.closest(TAB);
                    if (!tab) return;
                    const list = tab.closest(TABS);
                    if (!list) return;
                    index = lists().indexOf(list);
                    offset = list.getBoundingClientRect().top;
                    scroller = scrollParent(list);
                    const box = document.querySelector(BOX);
                    if (box) box.style.minHeight = '';
                    cancelAnimationFrame(raf);
                    until = Date.now() + 5000;
                    raf = requestAnimationFrame(tick);
                }, true);
                ['touchmove', 'wheel'].forEach(ev =>
                    document.addEventListener(ev, stop, {passive: true, capture: true}));
            })();`;
            p.document.head.appendChild(s);
        }
    } catch (e) {}
    </script>""",
    height=0,
)
