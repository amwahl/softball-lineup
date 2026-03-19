#!/usr/bin/env python3
"""
Simulate a softball season (batting stats + lineups) using algorithms inspired by Code.gs.
Creates outputs in `simulation_output/` as CSV and JSON.
"""
import csv
import json
import math
import os
import random
import argparse
from collections import defaultdict

POSITIONS = ['P', 'C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF']
MAX_PLAYERS = 12

OUTPUT_DIR = 'simulation_output'


def default_roster(n=12):
    roster = [f'Player {i+1}' for i in range(n)]
    return roster


def seed_players(roster, seed=None):
    rnd = random.Random(seed)
    profiles = {}
    for name in roster:
        # assign rough skill levels
        obp = rnd.uniform(0.25, 0.45)  # on-base percent
        slg = rnd.uniform(0.2, 0.6)    # slugging
        sb_rate = rnd.uniform(0.0, 0.2) # steal attempts per game
        profiles[name] = { 'obp': obp, 'slg': slg, 'sb_rate': sb_rate }
    return profiles


def compute_top_mid_overall_scores(batting_averages):
    # batting_averages: dict name -> {obp, slg, baserunning}
    def top_score(s): return (s['obp'] * 100) + (s.get('baserunning',0) * 5)
    def mid_score(s): return (s['slg'] * 100) + (s['obp'] * 30)
    def overall_score(s): return (s['obp'] * 50) + (s['slg'] * 50) + (s.get('baserunning',0) * 3)
    return top_score, mid_score, overall_score


def generate_batting_order(available_players, batting_averages):
    # Ported simplified logic from Code.gs
    with_data = []
    new_players = []
    for p in available_players:
        avg = batting_averages.get(p)
        if avg and avg.get('games',0) >= 3:
            with_data.append({'name': p, 'stats': avg})
        else:
            new_players.append(p)
    if not with_data:
        return [{'name': p, 'position': i+1} for i,p in enumerate(available_players)]

    top_score, mid_score, overall_score = compute_top_mid_overall_scores(batting_averages)
    topCandidates = sorted(with_data, key=lambda x: top_score(x['stats']), reverse=True)
    midCandidates = sorted(with_data, key=lambda x: mid_score(x['stats']), reverse=True)
    totalSlots = len(available_players)
    topSlots = min(3, totalSlots)
    midSlots = min(3, max(0, totalSlots - 3))

    assigned = set()
    order = [None] * totalSlots

    slot = 0
    for c in topCandidates:
        if slot >= topSlots: break
        if c['name'] in assigned: continue
        order[slot] = c
        assigned.add(c['name']); slot += 1

    slot = topSlots
    for c in midCandidates:
        if slot >= topSlots + midSlots: break
        if c['name'] in assigned: continue
        order[slot] = c
        assigned.add(c['name']); slot += 1

    remaining = [c for c in with_data if c['name'] not in assigned]
    remaining = sorted(remaining, key=lambda x: overall_score(x['stats']), reverse=True)
    slot = topSlots + midSlots
    for c in remaining:
        if slot >= totalSlots: break
        order[slot] = c
        assigned.add(c['name']); slot += 1

    for name in new_players:
        if slot >= totalSlots: break
        order[slot] = {'name': name, 'stats': batting_averages.get(name, {})}
        assigned.add(name); slot += 1

    # fill any gaps
    for i in range(totalSlots):
        if not order[i]:
            for p in available_players:
                if p not in assigned:
                    order[i] = {'name': p, 'stats': batting_averages.get(p,{})}
                    assigned.add(p); break

    # return simplified order
    return [{'name': e['name'], 'position': idx+1, 'obp': e.get('stats',{}).get('obp',0), 'slg': e.get('stats',{}).get('slg',0)} for idx,e in enumerate(order)]


def simulate_game(game_index, roster, profiles, batting_averages, innings=6, seed=None):
    rnd = random.Random((seed or 0) + game_index)
    # attendance: each player has 95% chance to attend
    available = [p for p in roster if rnd.random() < 0.95]
    if len(available) < 9:
        # force at least 9: add from roster
        needed = 9 - len(available)
        for p in roster:
            if p not in available:
                available.append(p)
                needed -= 1
                if needed <= 0: break
    batting_order = generate_batting_order(available, batting_averages)
    batting_order_names = [b['name'] for b in batting_order]

    # simulate plate appearances per inning roughly 3*team_runs/9 -> but simplify: each lineup spot gets ~ (innings*1.0 to 1.6) ABs
    per_player_ab = {name: 0 for name in available}
    per_player_bb = {name: 0 for name in available}
    per_player_hits = {name: 0 for name in available}
    per_player_types = {name: {'1B':0,'2B':0,'3B':0,'HR':0} for name in available}
    per_player_sb = {name: 0 for name in available}
    per_player_cs = {name: 0 for name in available}

    # Estimate at-bats per player: depending on batting pos and innings
    for idx, name in enumerate(batting_order_names):
        # leading spots get slightly more
        base = innings * (0.9 + (3 - min(idx,3)) * 0.05)
        ab = max(0, int(round(rnd.gauss(base, 0.8))))
        per_player_ab[name] = ab

        # simulate AB outcomes
        prof = profiles.get(name, {'obp':0.3, 'slg':0.35, 'sb_rate':0})
        obp = prof['obp']
        slg = prof['slg']
        # convert obp->walk probability roughly obp - avg_hit_prob; approximate avg hit prob from slg
        # crude mapping: hit_prob ~ slg / 1.2 (so slg 0.4 -> hit_prob 0.33)
        hit_prob = max(0.01, min(0.6, slg / 1.2))
        walk_prob = max(0.0, obp - hit_prob)
        # distribute hits by type according to slugging
        # assume HR ratio small
        for a in range(ab):
            r = rnd.random()
            if r < walk_prob:
                per_player_bb[name] += 1
            else:
                # hit occurs with probability hit_prob/(1-walk_prob) scaled
                if rnd.random() < hit_prob:
                    # choose type by simple heuristic from slg
                    # higher slg -> more extra-base hits
                    b = rnd.random()
                    if b < 0.75:
                        per_player_types[name]['1B'] += 1
                    elif b < 0.95:
                        per_player_types[name]['2B'] += 1
                    elif b < 0.995:
                        per_player_types[name]['3B'] += 1
                    else:
                        per_player_types[name]['HR'] += 1
                    per_player_hits[name] += 1
        # steals attempts
        sb_attempts = 0
        for _ in range(ab):
            if rnd.random() < prof.get('sb_rate',0.05):
                # success ~ 70%
                if rnd.random() < 0.7:
                    per_player_sb[name] += 1
                else:
                    per_player_cs[name] += 1

    # Build per-game batting rows
    batting_rows = []
    for pos, name in enumerate(batting_order_names, start=1):
        if name not in per_player_ab: continue
        ab = per_player_ab[name]
        ones = per_player_types[name]['1B']
        twos = per_player_types[name]['2B']
        threes = per_player_types[name]['3B']
        hrs = per_player_types[name]['HR']
        bb = per_player_bb[name]
        sb = per_player_sb[name]
        cs = per_player_cs[name]
        batting_rows.append({
            'Game': game_index+1,
            'Player': name,
            'AB': ab,
            '1B': ones, '2B': twos, '3B': threes, 'HR': hrs,
            'BB': bb, 'SB': sb, 'CS': cs,
            'BattingPos': pos
        })

    # fielding: assign players to positions by simple rotation
    # ensure P and C assigned from available
    fielding = {inning: {} for inning in range(1, innings+1)}
    # naive: for each inning, assign first 9 available players to the 9 positions shifting by inning
    for inning in range(1, innings+1):
        for i, p in enumerate(available[:9]):
            fielding[inning][POSITIONS[(i + inning - 1) % len(POSITIONS)]] = p

    return {
        'game_index': game_index+1,
        'available': available,
        'batting_order': batting_order,
        'batting_rows': batting_rows,
        'fielding': fielding,
        'innings': innings
    }


def aggregate_season(all_games):
    season = defaultdict(lambda: defaultdict(int))
    games_played = defaultdict(int)
    for g in all_games:
        for row in g['batting_rows']:
            p = row['Player']
            season[p]['AB'] += row['AB']
            season[p]['1B'] += row['1B']
            season[p]['2B'] += row['2B']
            season[p]['3B'] += row['3B']
            season[p]['HR'] += row['HR']
            season[p]['BB'] += row['BB']
            season[p]['SB'] += row['SB']
            season[p]['CS'] += row['CS']
            season[p]['PA'] = season[p]['AB'] + season[p]['BB']
            if row['AB'] > 0 or row['BB'] > 0:
                games_played[p] += 1
    # compute derived
    out = {}
    for p, stats in season.items():
        hits = stats['1B'] + stats['2B'] + stats['3B'] + stats['HR']
        tb = stats['1B'] + 2*stats['2B'] + 3*stats['3B'] + 4*stats['HR']
        ab = stats['AB']
        pa = stats.get('PA',0)
        obp = (hits + stats['BB']) / pa if pa > 0 else 0
        slg = tb / ab if ab > 0 else 0
        out[p] = {
            'games': games_played[p],
            'AB': ab, '1B': stats['1B'], '2B': stats['2B'], '3B': stats['3B'], 'HR': stats['HR'],
            'BB': stats['BB'], 'SB': stats['SB'], 'CS': stats['CS'],
            'OBP': round(obp,3), 'SLG': round(slg,3)
        }
    return out


def ensure_out_dir():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)


def write_outputs(all_games, season_stats):
    ensure_out_dir()
    # write per-game batting CSV
    with open(os.path.join(OUTPUT_DIR, 'per_game_batting.csv'), 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['Game','Player','AB','1B','2B','3B','HR','BB','SB','CS','BattingPos'])
        writer.writeheader()
        for g in all_games:
            for row in g['batting_rows']:
                writer.writerow(row)
    # per-game lineups
    with open(os.path.join(OUTPUT_DIR, 'lineups.json'), 'w') as f:
        json.dump([{'game': g['game_index'], 'batting_order': g['batting_order'], 'fielding': g['fielding']} for g in all_games], f, indent=2)
    # season summary
    with open(os.path.join(OUTPUT_DIR, 'season_stats.json'), 'w') as f:
        json.dump(season_stats, f, indent=2)
    print('Wrote outputs to', OUTPUT_DIR)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--games','-g', type=int, default=20)
    parser.add_argument('--innings','-i', type=int, default=6)
    parser.add_argument('--seed', type=int, default=42)
    parser.add_argument('--roster', type=str, default='')
    parser.add_argument('--batting-stats', type=str, default='')
    parser.add_argument('--season-history', type=str, default='')
    args = parser.parse_args()

    if args.roster and os.path.exists(args.roster):
        roster = []
        with open(args.roster, newline='') as f:
            reader = csv.reader(f)
            for row in reader:
                if row: roster.append(row[0])
        roster = roster[:MAX_PLAYERS]
    else:
        roster = default_roster(MAX_PLAYERS)

    profiles = seed_players(roster, seed=args.seed)

    # Load existing batting stats if provided to seed batting_averages
    def load_batting_stats(path):
        if not os.path.exists(path):
            return {}, []
        with open(path, newline='') as f:
            # attempt auto-detect delimiter
            sample = f.read(2048)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample.replace('\t', ','))
            except Exception:
                dialect = csv.excel
            reader = csv.DictReader(f, dialect=dialect)
            players = {}
            games_seen = defaultdict(set)
            for row in reader:
                # normalize keys
                keys = {k.strip(): v for k, v in row.items()}
                player = keys.get('Player') or keys.get('player')
                if not player: continue
                player = player.strip()
                # parse numeric fields safely
                def ni(k):
                    v = keys.get(k) or keys.get(k.replace('#','')) or ''
                    try: return int(v)
                    except: return 0
                game = ni('Game #') or ni('Game#') or ni('Game')
                ab = ni('AB')
                one = ni('1B')
                two = ni('2B')
                three = ni('3B')
                hr = ni('HR')
                bb = ni('BB')
                sb = ni('SB')
                cs = ni('CS')
                bpos = ni('BattingPos')

                p = players.setdefault(player, {'games':0,'ab':0,'singles':0,'doubles':0,'triples':0,'hr':0,'bb':0,'sb':0,'cs':0,'baserunning':0,'avgBattingPos':0})
                if game and game in games_seen[player]:
                    # skip duplicate entries for same game
                    continue
                if game: games_seen[player].add(game)
                p['games'] += 1
                p['ab'] += ab
                p['singles'] += one
                p['doubles'] += two
                p['triples'] += three
                p['hr'] += hr
                p['bb'] += bb
                p['sb'] += sb
                p['cs'] += cs
                p['baserunning'] = (p['sb'] * 1.5) - (p['cs'] * 2)
                # running average batting pos
                prev = p.get('avgBattingPos',0)
                p['avgBattingPos'] = ((prev * (p['games']-1)) + (bpos or 0)) / p['games']
            # derived rates
            for name,p in players.items():
                hits = p['singles'] + p['doubles'] + p['triples'] + p['hr']
                tb = p['singles'] + 2*p['doubles'] + 3*p['triples'] + 4*p['hr']
                pa = p['ab'] + p['bb']
                p['obp'] = (hits + p['bb']) / pa if pa > 0 else 0
                p['slg'] = tb / p['ab'] if p['ab'] > 0 else 0
            return players, list(players.keys())

    batting_averages = {p: {'games': 0, 'ab':0, 'singles':0, 'doubles':0,'triples':0,'hr':0,'bb':0,'sb':0,'cs':0,'baserunning':0,'avgBattingPos':0,'obp':0,'slg':0} for p in roster}
    if args.batting_stats and os.path.exists(args.batting_stats):
        loaded, names = load_batting_stats(args.batting_stats)
        # If roster was default, replace roster with names from batting stats
        if (not args.roster or args.roster == '') and names:
            roster = names[:MAX_PLAYERS]
        # seed batting_averages for players we have data for
        for name, stats in loaded.items():
            batting_averages[name] = stats

    all_games = []
    for gi in range(args.games):
        g = simulate_game(gi, roster, profiles, batting_averages, innings=args.innings, seed=args.seed)
        # after simulation, update batting_averages with this game's stats
        for row in g['batting_rows']:
            p = row['Player']
            ba = batting_averages.setdefault(p, {'games':0,'ab':0,'singles':0,'doubles':0,'triples':0,'hr':0,'bb':0,'sb':0,'cs':0,'baserunning':0,'avgBattingPos':0})
            ba['games'] += 1
            ba['ab'] += row['AB']
            ba['singles'] += row['1B']
            ba['doubles'] += row['2B']
            ba['triples'] += row['3B']
            ba['hr'] += row['HR']
            ba['bb'] += row['BB']
            ba['sb'] += row['SB']
            ba['cs'] += row['CS']
            ba['baserunning'] = (ba['sb'] * 1.5) - (ba['cs'] * 2)
            # rolling avg batting pos naive: average of positions
            ba['avgBattingPos'] = ((ba.get('avgBattingPos',0) * (ba['games']-1)) + row['BattingPos']) / ba['games']
            # compute derived rates for use by batting order algorithm
            hits = ba.get('singles',0) + ba.get('doubles',0) + ba.get('triples',0) + ba.get('hr',0)
            total_bases = ba.get('singles',0) + 2*ba.get('doubles',0) + 3*ba.get('triples',0) + 4*ba.get('hr',0)
            pa = ba.get('ab',0) + ba.get('bb',0)
            ba['obp'] = (hits + ba.get('bb',0)) / pa if pa > 0 else 0
            ba['slg'] = total_bases / ba.get('ab',1) if ba.get('ab',0) > 0 else 0
        all_games.append(g)

    season = aggregate_season(all_games)
    write_outputs(all_games, season)

if __name__ == '__main__':
    main()
