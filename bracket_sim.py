"""Barttorvik-based game odds and the bracket Monte Carlo, shared by dashboard.py and commissioner.py.

Keeping one copy means the commissioner app's projected FPts always match the dashboard's.
"""
import difflib
import math
import re

import numpy as np

# Seed-based single-game strength (points of margin), used only for teams without Barttorvik ratings.
SEED_STRENGTH = {
    1: 25.0, 2: 21.0, 3: 18.0, 4: 16.0, 5: 14.0, 6: 13.0, 7: 11.5, 8: 10.0,
    9: 9.5, 10: 8.5, 11: 8.0, 12: 7.5, 13: 4.5, 14: 2.0, 15: 0.0, 16: -4.0,
}
WIN_PROB_SIMS = 10000
SIM_SEED = 2027
# Game-to-game variation (standard deviation as a share of a player's average).
GAME_FPTS_SPREAD = 0.30
REGIONS = ["EAST", "WEST", "MIDWEST", "SOUTH"]
# Seeds in bracket order within a region: each adjacent pair plays in the first round.
BRACKET_SLOT_SEEDS = [1, 16, 8, 9, 5, 12, 4, 13, 6, 11, 3, 14, 7, 10, 2, 15]
GAME_MARGIN_SD = 11.0
DEFAULT_TEMPO = 68.0
# Workbook team names (normalised) whose Barttorvik names differ.
TORVIK_ALIASES = {
    "longisland": "liu", "mcneese": "mcneesest", "miami": "miamifl", "olemiss": "mississippi",
    "uconn": "connecticut", "pitt": "pittsburgh", "louisianast": "lsu", "northcarolinast": "ncst",
    "brighamyoung": "byu", "southerncalifornia": "usc",
}


def norm_team_name(name):
    name = str(name).casefold().replace("saint", "st")
    return re.sub(r"[^a-z0-9]", "", name)


def torvik_key(name):
    text = str(name).casefold().replace("&", "and")
    text = re.sub(r"[.'()]", "", text)
    text = re.sub(r"\bsaint\b", "st", text)
    text = re.sub(r"\bstate\b", "st", text)
    return re.sub(r"[^a-z0-9]", "", text)


def read_final_four_config(config_ws):
    """Which regions meet in each national semifinal, from Config!M2:N3 (defaults if blank)."""
    cells = {
        "semifinal_1_region_1": ("M2", "EAST"), "semifinal_1_region_2": ("M3", "WEST"),
        "semifinal_2_region_1": ("N2", "MIDWEST"), "semifinal_2_region_2": ("N3", "SOUTH"),
    }
    config = {}
    for key, (cell, default) in cells.items():
        value = config_ws[cell].value
        config[key] = str(value).strip().upper() if value else default
    return config


def read_bracket_teams(team_list_ws):
    """{seed: {REGION: [teams]}} from the TeamList sheet's Seed, Region and Team columns."""
    bracket = {}
    header_values = next(team_list_ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    headers = {str(v).strip(): i for i, v in enumerate(header_values) if v is not None}
    if not {"Seed", "Region", "Team"} <= headers.keys():
        return bracket
    for values in team_list_ws.iter_rows(min_row=2, max_col=max(headers.values()) + 1, values_only=True):
        try:
            seed_value = values[headers["Seed"]]
            if seed_value is None:
                continue
            # Seeds like "11 - FF" carry a First Four suffix.
            seed_str = str(seed_value).strip()
            if " - " in seed_str:
                seed_str = seed_str.split(" - ")[0].strip()
            seed = int(float(seed_str))
            region = str(values[headers["Region"]]).strip().upper() if values[headers["Region"]] else None
            team = str(values[headers["Team"]]).strip() if values[headers["Team"]] else None
            if seed and region and team and 1 <= seed <= 16:
                bracket.setdefault(seed, {}).setdefault(region, []).append(team)
        except (TypeError, ValueError, IndexError):
            continue
    return bracket


NET_WEIGHT = 0.5
# NCAA NET school names that need help matching Barttorvik's.
NET_ALIASES = {
    "UNI": "Northern Iowa", "SFA": "Stephen F. Austin", "California Baptist": "Cal Baptist",
    "UNCW": "UNC Wilmington", "UIC": "Illinois Chicago", "UTRGV": "UT Rio Grande Valley",
    "ETSU": "East Tennessee St.", "Col. of Charleston": "Charleston", "LMU (CA)": "Loyola Marymount",
    "CSUN": "Cal St. Northridge", "App State": "Appalachian St.", "UT Martin": "Tennessee Martin",
    "Southern Miss.": "Southern Miss", "FGCU": "Florida Gulf Coast", "Lamar University": "Lamar",
    "SIUE": "SIU Edwardsville", "Omaha": "Nebraska Omaha", "Boston U.": "Boston University",
    "N.C. A&T": "North Carolina A&T", "UIW": "Incarnate Word", "Loyola Maryland": "Loyola MD",
    "FDU": "Fairleigh Dickinson", "NIU": "Northern Illinois", "Army West Point": "Army",
    "N.C. Central": "North Carolina Central", "UMES": "Maryland Eastern Shore", "ULM": "Louisiana Monroe",
    "Kansas City": "UMKC",
}
_NET_ABBREVIATIONS = {
    "ill": "illinois", "fla": "florida", "ky": "kentucky", "mich": "michigan", "tenn": "tennessee",
    "colo": "colorado", "wash": "washington", "ark": "arkansas", "ga": "georgia", "ala": "alabama",
    "mo": "missouri", "caro": "carolina", "ariz": "arizona", "ind": "indiana", "la": "louisiana",
    "conn": "connecticut", "miss": "mississippi", "val": "valley", "col": "college", "so": "southern",
}


def _net_key(name):
    text = re.sub(r"\s*\([^)]*\)", "", str(name))
    words = [_NET_ABBREVIATIONS.get(w[:-1].casefold(), w) if w.endswith(".") else w for w in text.split()]
    return torvik_key(" ".join(words))


def net_rating_key(name, rating_keys):
    """The ratings key for an NCAA NET school name, or None. Exact and alias matches first, then a close spelling."""
    key = torvik_key(NET_ALIASES[name]) if name in NET_ALIASES else _net_key(name)
    for candidate in (key, TORVIK_ALIASES.get(key, "")):
        if candidate in rating_keys:
            return candidate
    close = difflib.get_close_matches(key, rating_keys, n=1, cutoff=0.8)
    return close[0] if close else None


def blend_net_rankings(ratings, net_ranks, weight=NET_WEIGHT):
    """Pull each team's Barttorvik net rating toward the rating Barttorvik gives the team at its NET rank.

    NET only publishes ranks, so a team ranked Nth by NET is credited with the net rating of the Nth-best
    team by Barttorvik; `weight` is how far to move toward it. The shift is split evenly between offense
    and defense. Returns (ratings, number of NET teams matched); teams without a NET match are unchanged.
    """
    if not ratings or not net_ranks:
        return ratings, 0
    keys = list(ratings)
    ranked = {}
    for name, rank in net_ranks.items():
        key = net_rating_key(name, keys)
        if key:
            ranked[key] = rank
    ordered = sorted((r[0] - r[1] for r in ratings.values()), reverse=True)
    blended = dict(ratings)
    for key, rank in ranked.items():
        adjoe, adjde, *rest = ratings[key]
        target = ordered[min(max(rank, 1), len(ordered)) - 1]
        shift = weight * (target - (adjoe - adjde))
        blended[key] = (adjoe + shift / 2, adjde - shift / 2, *rest)
    return blended, len(ranked)


def find_rating(ratings, name):
    """(adjoe, adjde, adjt, rank, record, conf) for a workbook team name, or None."""
    key = torvik_key(name)
    return ratings.get(key) or ratings.get(TORVIK_ALIASES.get(key, ""))


def team_win_matrix(ratings, team_names, team_seeds):
    """Pairwise neutral-site win probabilities from Barttorvik adjusted offense, defense and tempo.

    Expected margin = net rating gap per 100 possessions x expected possessions / 100; win probability
    is a normal curve on that margin. Teams without ratings fall back to seed-based strength.
    Returns (matrix, number of teams matched to a rating).
    """
    tempos = [r[2] for r in ratings.values()]
    avg_tempo = float(np.mean(tempos)) if tempos else DEFAULT_TEMPO
    n = len(team_names)
    net = np.zeros(n)
    tempo = np.full(n, avg_tempo)
    matched = 0
    for i, (name, seed_value) in enumerate(zip(team_names, team_seeds)):
        rating = find_rating(ratings, name)
        if rating:
            net[i] = rating[0] - rating[1]
            tempo[i] = rating[2]
            matched += 1
        else:
            net[i] = SEED_STRENGTH.get(int(seed_value), 0.0) * 100.0 / avg_tempo
    possessions = np.outer(tempo, tempo) / avg_tempo
    margin = (net[:, None] - net[None, :]) / 100.0 * possessions
    win = 0.5 * (1.0 + np.vectorize(math.erf)(margin / (GAME_MARGIN_SD * math.sqrt(2.0))))
    return win, matched


def run_bracket(bracket, final_four_config, ratings, n_sims, rng, lost_rounds=None):
    """Simulate the whole bracket n_sims times.

    bracket: {seed: {REGION: [teams]}}; lost_rounds: {team name casefolded: round index the team lost in
    (None if the round is unknown)} for real results that override the model.
    Returns None if the bracket is incomplete, else a dict with `appearances` (n_sims x n_teams games
    played per simulation), `team_names`, `team_index` (norm_team_name -> column) and `matched`.
    """
    if not bracket:
        return None
    lost_rounds = lost_rounds or {}
    team_names, team_seeds, team_index = [], [], {}

    def team_id(team, seed_value):
        key = norm_team_name(team)
        if key not in team_index:
            team_index[key] = len(team_names)
            team_names.append(team)
            team_seeds.append(seed_value)
        return team_index[key]

    region_slots = {}
    for region in REGIONS:
        slots = []
        for slot_seed in BRACKET_SLOT_SEEDS:
            teams = bracket.get(slot_seed, {}).get(region, [])
            if not teams or len(teams) > 2:
                return None
            slots.append([team_id(t, slot_seed) for t in teams])
        region_slots[region] = slots

    n_teams = len(team_names)
    lost = np.full(n_teams, 99)
    for i, name in enumerate(team_names):
        lost_round = lost_rounds.get(str(name).strip().casefold())
        if lost_round is not None:
            lost[i] = lost_round
    win_matrix, matched = team_win_matrix(ratings, team_names, team_seeds)

    def play(a, b, rnd):
        a_out, b_out = lost[a] <= rnd, lost[b] <= rnd
        a_surv = (lost[a] > rnd) & (lost[a] < 99)
        b_surv = (lost[b] > rnd) & (lost[b] < 99)
        p_a = win_matrix[a, b]
        sampled = np.where(rng.random(a.shape) < p_a, a, b)
        # Real results win over the model: a known loser loses, a team known to play later advances.
        return np.where(
            a_out & ~b_out, b,
            np.where(b_out & ~a_out, a,
                     np.where(a_surv & ~b_surv, a, np.where(b_surv & ~a_surv, b, sampled))),
        )

    rows = np.arange(n_sims)[:, None]
    appearances = np.zeros((n_sims, n_teams), dtype=np.int16)

    def record(a, b):
        appearances[rows, a] += 1
        appearances[rows, b] += 1

    def run_region(slots):
        entrants = []
        for slot in slots:
            if len(slot) == 2:
                entrants.append(play(np.full(n_sims, slot[0]), np.full(n_sims, slot[1]), -1))
            else:
                entrants.append(np.full(n_sims, slot[0]))
        current = np.stack(entrants, axis=1)
        for rnd in range(4):
            a, b = current[:, 0::2], current[:, 1::2]
            record(a, b)
            current = play(a, b, rnd)
        return current[:, 0]

    winners = {region: run_region(region_slots[region]) for region in REGIONS}
    semi_pairs = [
        (final_four_config["semifinal_1_region_1"], final_four_config["semifinal_1_region_2"]),
        (final_four_config["semifinal_2_region_1"], final_four_config["semifinal_2_region_2"]),
    ]
    finalists = []
    for first_region, second_region in semi_pairs:
        if first_region not in winners or second_region not in winners:
            return None
        a, b = winners[first_region], winners[second_region]
        record(a[:, None], b[:, None])
        finalists.append(play(a, b, 4))
    record(finalists[0][:, None], finalists[1][:, None])

    return {
        "appearances": appearances, "team_names": team_names,
        "team_index": team_index, "matched": matched,
    }


def pre_tournament_expected_games(bracket, final_four_config, ratings, n_sims=WIN_PROB_SIMS, seed=SIM_SEED):
    """{norm_team_name: average games played} over simulated tournaments that ignore all results.

    Same seed and call order as the dashboard's pre-tournament simulation, so the numbers are identical.
    Returns None if the bracket can't be simulated.
    """
    result = run_bracket(bracket, final_four_config, ratings, n_sims, np.random.default_rng(seed))
    if result is None:
        return None
    means = result["appearances"].mean(axis=0)
    return {key: float(means[index]) for key, index in result["team_index"].items()}
