#!/usr/bin/env python3
"""
check_lineup_rules.py — Audit a generated lineup against the three hard rules.

Rules checked:
  1. No player is assigned to a position they are marked Restricted at.
     (Requires a restrictions CSV. Skipped if not provided.)
  2. The starting pitcher (inning 1, P) also pitches inning 2. No 1-inning starters.
  3. No player sits out two or more consecutive innings.

Usage:
  python3 check_lineup_rules.py --lineup lineup.csv [--restrictions roster_prefs.csv]

Lineup CSV format (matches the Game Entry sheet):
  Inning,P,C,1B,2B,3B,SS,LF,CF,RF,SitOut1,SitOut2,SitOut3
  1,Aleia,Lauren,Lily,Eleanor,Emily,Mia K,Aubrey,Lea,Mia S,Molly,,
  2,...

Restrictions CSV format (matches the Roster sheet):
  Player,P,C,1B,2B,3B,SS,LF,CF,RF
  Aleia,Preferred,Restricted,Okay,Okay,Okay,Okay,Okay,Okay,Okay
  ...

Exits 0 if all rules hold, 1 if any violation is found.
"""
import argparse
import csv
import sys
from collections import defaultdict

POSITIONS = ['P', 'C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF']
SIT_OUT_COLS = ['SitOut1', 'SitOut2', 'SitOut3']


def load_lineup(path):
    """Returns a list of dicts, one per inning, with keys = position/sit-out column names."""
    innings = []
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            innings.append({k: (v or '').strip() for k, v in row.items()})
    return innings


def load_restrictions(path):
    """Returns dict: player -> { position: 'Preferred' | 'Okay' | 'Restricted' }."""
    prefs = {}
    with open(path, newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get('Player') or '').strip()
            if not name:
                continue
            prefs[name] = {pos: (row.get(pos) or 'Okay').strip() for pos in POSITIONS}
    return prefs


def check_restrictions(innings, prefs):
    violations = []
    for row in innings:
        inning_num = row.get('Inning', '?')
        for pos in POSITIONS:
            player = row.get(pos, '')
            if not player:
                continue
            if player in prefs and prefs[player].get(pos) == 'Restricted':
                violations.append(
                    f"Inning {inning_num}: {player} is assigned {pos} but marked Restricted at {pos}"
                )
    return violations


def check_starter_two_innings(innings):
    if len(innings) < 2:
        return []
    starter = innings[0].get('P', '').strip()
    second = innings[1].get('P', '').strip()
    if not starter:
        return ["Inning 1 has no pitcher assigned"]
    if starter != second:
        return [f"Starting pitcher {starter!r} did not pitch inning 2 (replaced by {second!r})"]
    return []


def check_consecutive_sit_outs(innings):
    """A player who appears in any SitOut column in inning N must NOT appear in any SitOut column in inning N+1."""
    violations = []
    prev_sitters = set()
    for idx, row in enumerate(innings):
        inning_num = row.get('Inning', str(idx + 1))
        sitters = {row.get(col, '').strip() for col in SIT_OUT_COLS if row.get(col, '').strip()}
        if idx > 0:
            both = prev_sitters & sitters
            for p in sorted(both):
                violations.append(
                    f"{p} sat out inning {innings[idx-1].get('Inning', idx)} and inning {inning_num} (consecutive)"
                )
        prev_sitters = sitters
    return violations


def main():
    parser = argparse.ArgumentParser(description="Audit a softball lineup against the three hard rules.")
    parser.add_argument('--lineup', required=True, help='CSV of the lineup (one row per inning)')
    parser.add_argument('--restrictions', help='Optional CSV of player position preferences')
    args = parser.parse_args()

    innings = load_lineup(args.lineup)
    if not innings:
        print(f"ERROR: no rows found in {args.lineup}", file=sys.stderr)
        return 2

    all_violations = defaultdict(list)

    if args.restrictions:
        prefs = load_restrictions(args.restrictions)
        all_violations['Rule 1 (Restricted)'] = check_restrictions(innings, prefs)
    else:
        print("Note: --restrictions not provided — skipping Rule 1 (Restricted) check.\n")

    all_violations['Rule 2 (Starter ≥ 2 innings)'] = check_starter_two_innings(innings)
    all_violations['Rule 3 (No consecutive sit-outs)'] = check_consecutive_sit_outs(innings)

    total = sum(len(v) for v in all_violations.values())
    for rule, violations in all_violations.items():
        if violations:
            print(f"FAIL  {rule}:")
            for v in violations:
                print(f"  - {v}")
        else:
            print(f"PASS  {rule}")
    print()
    print(f"Total violations: {total}")
    return 1 if total > 0 else 0


if __name__ == '__main__':
    sys.exit(main())
