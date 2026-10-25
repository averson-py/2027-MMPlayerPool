
import colorsys
import os
import re
import secrets
from datetime import datetime
from pathlib import Path


import openpyxl
import pandas as pd
import requests
import streamlit as st


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

    # Draft order
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

    max_column = max(headers.values()) + 1
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
        }

    player_names.sort(key=str.casefold)

    # Eliminated teams
    teams_out = set()
    if TEAMS_OUT_SHEET in workbook.sheetnames:
        ws = workbook[TEAMS_OUT_SHEET]
        for values in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
            team_name = values[0]
            if team_name is not None and str(team_name).strip():
                teams_out.add(str(team_name).strip().casefold())

    # Workbook last-edited timestamp is maintained by the local stats updater
    # in GameResults!D2. This is the authoritative timestamp displayed by
    # the dashboard; it does not depend on OneDrive's file modified time.
    workbook_updated_time = None
    if GAME_RESULTS_SHEET not in workbook.sheetnames:
        raise ValueError(f"Missing required sheet: {GAME_RESULTS_SHEET}")

    ws = workbook[GAME_RESULTS_SHEET]
    raw_updated_time = ws["D2"].value
    if raw_updated_time is None or str(raw_updated_time).strip() == "":
        raise ValueError("GameResults!D2 is empty. The stats updater must write the workbook update date/time to cell D2.")

    if isinstance(raw_updated_time, datetime):
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

    if workbook_updated_time.tzinfo is not None:
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

    return {
        "fantasy_weights": fantasy_weights,
        "drafters": drafters,
        "draft_rows": draft_rows,
        "player_stats": player_stats,
        "player_names": player_names,
        "teams_out": teams_out,
        "latest_game_result": latest_game_result,
        "updated_time": workbook_updated_time,
    }


def player_total(row, fantasy_weights):
    return weighted_fantasy_points(
        row.get("tourn_pts"),
        row.get("tourn_reb"),
        row.get("tourn_ast"),
        fantasy_weights,
    )


def make_team_rows(data, selected_draftee):
    rows = [
        row for row in data["draft_rows"]
        if row["draftee"].casefold() == selected_draftee.casefold()
    ]
    rows.sort(key=lambda row: draft_position_sort_key(row["draft_position"]))

    result = []
    for row in rows:
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        result.append({
            "Draft Position": row["draft_position"],
            "Player Name": row["player_name"],
            "Team": stats.get("Team"),
            "Seed": stats.get("Seed"),
            "FPts": stats.get("Tot. FPts"),
            "PPG": stats.get("PPG"),
            "RPG": stats.get("RPG"),
            "APG": stats.get("APG"),
            "Tourn. FPts": player_total(row, data["fantasy_weights"]),
            "Tourn. PTS": row.get("tourn_pts"),
            "Tourn. REB": row.get("tourn_reb"),
            "Tourn. AST": row.get("tourn_ast"),
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

    # Convert round-by-round scoring into cumulative totals, then rank
    # everyone after each round. Rank 1 is the best (highest) total.
    result = []
    running_total = {owner: 0.0 for owner in round_totals}
    for i, round_name in enumerate(round_names):
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


def make_standings(data):
    positions = {
        draftee.casefold(): position
        for position, draftee in enumerate(data["drafters"], start=1)
    }
    totals = {}

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

    return pd.DataFrame(standings)[[
        "Rank", "Draft Position", "Team Owner", "Total FPts",
        "Active Players", "Poss. Games Remaining", "MVP", "MVP FPts"
    ]]


def make_power_rankings(data):
    positions = {
        draftee.casefold(): position
        for position, draftee in enumerate(data["drafters"], start=1)
    }
    totals = {}

    for row in data["draft_rows"]:
        key = row["draftee"].casefold()
        totals.setdefault(key, {
            "Team Owner": row["draftee"],
            "Draft Position": positions.get(key, ""),
            "Total FPts": None,
            "Active Players": 0,
            "Games Played": 0,
            "Poss. Games Remaining": 0,
        })
        t = totals[key]
        t["Total FPts"] = sum_existing_values([t["Total FPts"], player_total(row, data["fantasy_weights"])])

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        if team is not None and str(team).strip().casefold() not in data["teams_out"]:
            t["Active Players"] += 1

        t["Games Played"] += row.get("games_played") or 0
        t["Poss. Games Remaining"] += row.get("possible_games_remaining") or 0

    result = []
    for t in totals.values():
        if t["Games Played"] > 0:
            estimated = (
                (t["Total FPts"] or 0) / t["Games Played"] * t["Poss. Games Remaining"]
            )
        else:
            estimated = None

        fpg = (
            t["Total FPts"] / t["Games Played"]
            if t["Total FPts"] is not None and t["Games Played"] > 0
            else None
        )

        projected = None
        if t["Total FPts"] is not None or estimated is not None:
            projected = round((t["Total FPts"] or 0) + (estimated or 0), 1)

        result.append({**t, "FPts per game": round(fpg, 1) if fpg is not None else None, "Projected Score": projected})

    result.sort(key=lambda x: (
        -(x["Projected Score"] if x["Projected Score"] is not None else 0),
        x["Draft Position"] if isinstance(x["Draft Position"], int) else 999999,
        x["Team Owner"].casefold(),
    ))
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    return pd.DataFrame(result)[[
        "Rank", "Draft Position", "Team Owner", "Projected Score",
        "Total FPts", "Active Players", "FPts per game",
        "Games Played", "Poss. Games Remaining"
    ]]


def make_player_rankings(data):
    result = []
    for row in data["draft_rows"]:
        stats = data["player_stats"].get(row["player_name"].casefold(), {})
        team = stats.get("Team")
        result.append({
            "Player Name": row["player_name"],
            "Team": team,
            "Total FPts": player_total(row, data["fantasy_weights"]),
            "Team Owner": row["draftee"],
            "Games Played": row.get("games_played", 0),
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
            "_team_out": (
                team is not None
                and str(team).strip().casefold() in data["teams_out"]
            ),
            "_draft_position": row["draft_position"],
        })

    result.sort(key=lambda p: (
        -(p["Total FPts"] if p["Total FPts"] is not None else 0),
        draft_position_sort_key(p["_draft_position"]),
    ))
    result = result[:10]
    for i, row in enumerate(result, start=1):
        row["Rank"] = i

    return pd.DataFrame(result)[[
        "Rank", "Player Name", "Team", "Total FPts",
        "Team Owner", "Games Played", "Poss. Games Remaining",
        "_team_out",
    ]], result


def make_performers(data, positive=True):
    result = []
    for row in data["draft_rows"]:
        total = player_total(row, data["fantasy_weights"])
        games = row.get("games_played", 0)
        if games <= 0 or total is None:
            continue

        stats = data["player_stats"].get(row["player_name"].casefold(), {})
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
            "Player Name": row["player_name"],
            "Team": team,
            "Avg. FPts/Tourn. Game": round(tournament_avg, 1),
            "Delta From Average": round(delta, 1),
            "Team Owner": row["draftee"],
            "Games Played": games,
            "Poss. Games Remaining": row.get("possible_games_remaining", 0),
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

    return pd.DataFrame(result)[[
        "Rank", "Player Name", "Team", "Avg. FPts/Tourn. Game",
        "Delta From Average", "Team Owner", "Games Played",
        "Poss. Games Remaining", "_team_out",
    ]], result


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

    return pd.DataFrame(result)[[
        "Rank", "Draft Position", "Team Owner", "Projected FPts", "Regions"
    ]]


# ---------------------------------------------------------------------
# Streamlit presentation helpers
# ---------------------------------------------------------------------
def style_table(df, scale_columns=None, reverse_columns=None, medal_rank=True, pink_last=False, rank_rwg=False, performance_gradient=None, column_colors=None):
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

    styled = df.style.apply(apply_styles, axis=1)

    # Display numeric values cleanly. Calculations remain unchanged.
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
    }
    format_map = {}
    for col in df.columns:
        if col in one_decimal_columns:
            format_map[col] = lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
        elif pd.api.types.is_float_dtype(df[col]):
            format_map[col] = lambda value: "" if pd.isna(value) else f"{float(value):g}"

    if format_map:
        styled = styled.format(format_map)

    # Display positive performer deltas with an explicit "+" sign.
    if "Delta From Season Avg." in df.columns:
        styled = styled.format({
            "Delta From Season Avg.": lambda x: (
                f"+{x:.1f}" if pd.notna(x) and x > 0
                else (f"{x:.1f}" if pd.notna(x) else "")
            )
        })

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

    return styled


def display_df(df, height=420, scale_columns=None, reverse_columns=None, hide_index=True, rank_rwg=False, pink_last=False, performance_gradient=None, column_colors=None):
    if df.empty:
        st.info("No data available yet.")
        return

    styled = style_table(
        df,
        scale_columns=scale_columns,
        reverse_columns=reverse_columns,
        pink_last=pink_last,
        rank_rwg=rank_rwg,
        performance_gradient=performance_gradient,
        column_colors=column_colors,
    )

    # Size the table to its actual row count so short tables do not
    # show a large block of blank cells underneath the data.
    row_height = 35
    header_height = 38
    compact_height = header_height + (len(df) * row_height) + 4

    # Give every supplied row enough room to render without a vertical
    # scrollbar. The height argument is retained only as a safety cap for
    # unusually large tables.
    actual_height = compact_height if compact_height <= height else height

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=hide_index,
        height=actual_height,
    )


# ---------------------------------------------------------------------
# Public OneDrive helpers
# ---------------------------------------------------------------------
def get_onedrive_share_url():
    """Read the public OneDrive sharing URL from Streamlit Secrets."""
    if "onedrive" in st.secrets:
        url = st.secrets["onedrive"].get("share_url")
    else:
        url = st.secrets.get("ONEDRIVE_SHARE_URL")

    if not url:
        return None

    url = str(url).strip()
    return url or None


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
    footer {
        display: none !important;
    }

    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {
        border: 1px solid #d9d9d9;
        border-radius: 8px;
        padding: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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
    data = load_data(workbook_path, updated_time=cloud_updated_time)
except Exception as exc:
    st.error(f"Could not load the workbook: {exc}")
    st.stop()

latest = data["latest_game_result"]

st.caption(
    "Standings Updated: "
    + data["updated_time"].strftime("%m/%d/%Y %I:%M %p")
)

standings = make_standings(data)
power = make_power_rankings(data)
players, player_rows = make_player_rankings(data)
over, over_rows = make_performers(data, True)
under, under_rows = make_performers(data, False)
ratings = make_draft_ratings(data)

# Top-line summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("Pool Members", len(data["drafters"]))
c2.metric("Drafted Players", len(data["draft_rows"]))
c3.metric("Players Active", int(standings["Active Players"].sum()) if not standings.empty else 0)
c4.metric("Teams Eliminated", len(data["teams_out"]))

tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "🏠 Pool Overview",
    "🏆 Standings",
    "📊 Player Statistics",
    "👥 Team View",
    "🎯 Draft Ratings",
])

with tab0:
    st.subheader("Pool Overview")
    st.caption("A quick look at the current state of the pool.")

    # Current leaders and headline stats. These are all derived from the
    # same standings/power-ranking data used elsewhere in the dashboard.
    leader = standings.iloc[0] if not standings.empty else None
    mvp_row = None
    if not standings.empty:
        mvp_candidates = standings.dropna(subset=["MVP FPts"])
        if not mvp_candidates.empty:
            mvp_row = mvp_candidates.loc[mvp_candidates["MVP FPts"].idxmax()]

    projected_leader = power.iloc[0] if not power.empty else None
    active_total = int(standings["Active Players"].sum()) if not standings.empty else 0

    overview_history = make_standings_round_history(data)
    biggest_riser_name = "—"
    biggest_riser_text = "No completed rounds yet"
    biggest_faller_name = "—"
    biggest_faller_text = "No completed rounds yet"
    if not overview_history.empty and len(overview_history) >= 2:
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
                movement, owner = max(risers, key=lambda x: (x[0], x[1].casefold()))
                biggest_riser_name = owner
                biggest_riser_text = f"↑ {int(movement)} spot{'s' if movement != 1 else ''}"
            else:
                biggest_riser_name = "No riser"
                biggest_riser_text = "No one moved up"
            if fallers:
                movement, owner = min(fallers, key=lambda x: (x[0], x[1].casefold()))
                movement_down = abs(movement)
                biggest_faller_name = owner
                biggest_faller_text = f"↓ {int(movement_down)} spot{'s' if movement_down != 1 else ''}"
            else:
                biggest_faller_name = "No faller"
                biggest_faller_text = "No one moved down"

    c1, c2, c3, c4 = st.columns(4)
    if leader is not None:
        c1.metric("🏆 Current Leader", str(leader["Team Owner"]), f"{float(leader['Total FPts']):.1f} FPts")
    else:
        c1.metric("🏆 Current Leader", "—")
    if projected_leader is not None:
        c2.metric("📈 Projected Leader", str(projected_leader["Team Owner"]), f"{float(projected_leader['Projected Score']):.1f}")
    else:
        c2.metric("📈 Projected Leader", "—")
    if mvp_row is not None:
        c3.metric("⭐ Pool MVP", str(mvp_row["MVP"]), f"{float(mvp_row['MVP FPts']):.1f} FPts")
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
            f"-- def. {latest['Loser']} [ {latest['Round']} ]",
        )
    else:
        c4.metric("🏀 Most Recent Game", "No game recorded")

    st.markdown("### 🏁 Current Standings")
    if not standings.empty:
        overview_display = standings[[
            "Rank", "Team Owner", "Total FPts", "Active Players", "MVP", "MVP FPts"
        ]].head(10).copy()
        for column in ["Total FPts", "MVP FPts"]:
            overview_display[column] = overview_display[column].apply(
                lambda value: "" if pd.isna(value) else f"{float(value):.1f}"
            )
        display_df(overview_display, height=390, pink_last=True)

    # Biggest riser / faller
    c5, c6 = st.columns(2)
    if biggest_riser_name != "—":
        c5.metric("📈 Biggest Riser", biggest_riser_name, biggest_riser_text)
    else:
        c5.metric("📈 Biggest Riser", "—", biggest_riser_text)

    if biggest_faller_name != "—":
        c6.metric("📉 Biggest Faller", biggest_faller_name, biggest_faller_text)
    else:
        c6.metric("📉 Biggest Faller", "—", biggest_faller_text)

    st.subheader("📈 Standings Progression")
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
                tooltip=["Round:N", "Team Owner:N", "Rank:Q"],
            )

            lines = base.mark_line()
            points = base.mark_point(filled=True, size=55)

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
            labels = base.mark_text(
                fontSize=9,
                fontWeight="bold",
                baseline="middle",
            ).encode(
                text="Team Owner:N",
                color=alt.value("#FFFFFF"),
            )

            chart = alt.layer(lines, points, label_bg, labels).properties(height=500)
            st.altair_chart(chart, use_container_width=True)
        except ImportError:
            # Fallback if Altair is unavailable in the deployment environment.
            st.line_chart(
                standings_history.set_index("Round"),
                height=450,
                use_container_width=True,
            )

with tab1:
    st.subheader("Tournament Standings")
    display_df(
        standings,
        height=520,
        pink_last=True,
        scale_columns={
            "Draft Position": (
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
        reverse_columns={"Draft Position": True},
    )


    st.subheader("Power Rankings")
    display_df(
        power,
        height=520,
        rank_rwg=True,
        scale_columns={
            "Draft Position": (
                power["Draft Position"].min(),
                power["Draft Position"].max(),
            ),
            "Active Players": (
                power["Active Players"].min(),
                power["Active Players"].max(),
            ),
        },
        reverse_columns={"Draft Position": True},
    )


with tab2:
    st.subheader("Player Rankings — Top 10")
    display_df(players, height=400
    )

    st.markdown("### <span style='color:#2E8B57; font-weight:800;'>↑</span> Top Overperformers 🌟", unsafe_allow_html=True)
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
        performance_gradient="green",
    )

    st.markdown("### <span style='color:#C62828; font-weight:800;'>↓</span> Top Underperformers 💩", unsafe_allow_html=True)
    display_df(
        under_display,
        height=240,
        performance_gradient="red",
    )

    st.divider()
    st.subheader("Player Finder")
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
                "Draft Position": row["draft_position"],
                "Draftee": row["draftee"],
                "Team": stats.get("Team"),
                "Seed": stats.get("Seed"),
                "FPts": stats.get("Tot. FPts"),
                "PPG": stats.get("PPG"),
                "RPG": stats.get("RPG"),
                "APG": stats.get("APG"),
                "Tourn. FPts": player_total(row, data["fantasy_weights"]),
                "Tourn. PTS": row.get("tourn_pts"),
                "Tourn. REB": row.get("tourn_reb"),
                "Tourn. AST": row.get("tourn_ast"),
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

with tab3:
    st.subheader("Team View")

    selected_draftee = st.selectbox(
        "Select a team owner",
        [""] + data["drafters"],
        index=0,
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
            st.caption(
                f"Showing {len(team_df)} player(s) drafted by {selected_draftee}."
            )
            display_df(
                team_df,
                height=500,
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

with tab4:
    st.subheader("Draft Ratings")
    st.caption(
        "Projected FPts = seed average games × season-average FPts, summed by draftee."
    )
    display_df(
        ratings,
        height=520,
        rank_rwg=True,
        scale_columns={
            "Draft Position": (
                ratings["Draft Position"].min(),
                ratings["Draft Position"].max(),
            ),
            "Projected FPts": (
                ratings["Projected FPts"].min(),
                ratings["Projected FPts"].max(),
            ),
        },
        reverse_columns={"Draft Position": True},
    )

st.divider()
st.caption(
    "The dashboard reads the Excel workbook from OneDrive. "
    "No workbook changes are made by the app."
)
