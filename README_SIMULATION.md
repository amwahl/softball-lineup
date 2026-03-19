Simulation README

- Script: simulate_season.py
- Purpose: Generate a season's worth of batting stats and lineups using the project's lineup/batting-order heuristics (simplified, offline).

Usage:

python3 simulate_season.py --games 20 --innings 6 --seed 42

Outputs (written to `softball-lineup/simulation_output`):
- per_game_batting.csv  — per-game batting rows (Game, Player, AB, 1B, 2B, 3B, HR, BB, SB, CS, BattingPos)
- lineups.json         — per-game batting order and fielding assignment
- season_stats.json    — aggregated season stats (OBP, SLG, counts)

Notes:
- The simulator is standalone and does not modify any Apps Script files or Google Sheets.
- It respects the request not to touch `softball-practice`.
- If you have a CSV roster file (one player name per line), pass it with `--roster roster.csv`.
