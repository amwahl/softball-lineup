# Softball Lineup Manager

A Google Apps Script tool for managing recreational softball lineups, fielding rotations, and batting orders. Designed for coaches who want fair playing time, smart position assignments, and stat-driven batting orders.

## Features

- **Roster Management** — Track up to 12 players with per-position preferences (Preferred / Okay / Restricted)
- **Depth Chart** — Rank players at each position to guide the lineup algorithm
- **Game Entry** — Record each game's lineup and batting stats with dropdowns, attendance checkboxes, and multiple sit-out columns
- **Lineup Suggester** — Auto-generate balanced field lineups that respect preferences, rotate positions fairly, and keep players at the same position for multiple innings
- **Batting Order** — Suggest optimal batting orders based on OBP, slugging, and speed stats
- **Season Dashboard** — View innings at each position, recency tracking, and cumulative batting stats

## Quick Setup (5 minutes)

1. **Create a new Google Sheet** at [sheets.google.com](https://sheets.google.com)

2. **Open the Script Editor**
   - Click **Extensions > Apps Script**

3. **Paste the code**
   - Delete any existing code in the editor
   - Copy the entire contents of `Code.gs` and paste it in
   - Click the save icon (or Ctrl+S)

4. **Run the initializer**
   - In the Apps Script editor, select `initializeAll` from the function dropdown at the top
   - Click the **Run** button (▶)
   - When prompted, click **Review Permissions** > choose your Google account > **Allow**
   - This grants the script permission to modify your spreadsheet
   - **Re-running is safe:** Initialize All Sheets preserves your existing roster names and preferences

5. **Return to your spreadsheet**
   - You'll see a new **⚾ Softball** menu at the far right of the menu bar (after Extensions and Help)
   - All 8 sheets will be created automatically
   - **Don't see the menu?** Go to Extensions > Apps Script, select `onOpen` from the function dropdown, click Run (▶), authorize when prompted, then close and reopen the spreadsheet

## Sheets Overview

| Sheet | Purpose |
|-------|---------|
| **Roster** | Enter player names and position preferences |
| **Depth Chart** | Rank players at each position (used by the lineup algorithm) |
| **Game Entry** | Record each game's lineup, batting stats, and attendance |
| **Season History** | Auto-populated fielding data (don't edit directly) |
| **Batting Stats** | Auto-populated per-game batting data (editable to fix errors) |
| **Dashboard** | Season stats at a glance — fielding and batting |
| **Lineup Suggester** | Auto-generate field lineups and batting orders |
| **How To Use** | In-app instructions |

## First Steps After Setup

1. Go to the **Roster** sheet
2. Enter your players' names in column B (up to 12)
3. Set position preferences for each player:
   - **Preferred** (green) — they love this position
   - **Okay** (yellow) — they can play here
   - **Restricted** (red) — never put them here

## How It Works

### Field Lineup Algorithm

- Scores each player-position combination based on preference, depth chart ranking, recency, and continuity
- Restricted positions are hard constraints — a player will never be assigned to a Restricted position
- **No-return rule for P (hard):** Once a player leaves Pitcher, they cannot return to that position later in the game
- **No-return rule for C (soft):** Once a player leaves Catcher, the algorithm strongly avoids putting them back but will allow it if needed
- **Bullpen warmup (soft):** The algorithm prefers pitchers who sat out the previous inning to warm up, but will assign an available pitcher without warmup if needed
- **P/C continuity:** Pitcher and Catcher get a stronger continuity bonus than field positions, since leaving is permanent
- Players get a bonus for staying at the same position across innings (builds comfort and confidence)
- **Sit-out cap:** No player sits out more than their fair share — the cap is calculated as `ceil(total sit-out slots / players)`, so with 12 players and 5 innings, no one sits more than 2
- Sit-outs rotate fairly — avoids consecutive sit-outs for the same player, and proactively sits out the next depth-chart pitcher to enable warmup
- Recency is per-player — absent games don't inflate "games since last played"

### Batting Order Algorithm

- **Spots 1-3 (top):** Highest OBP + baserunning speed — get on base and steal
- **Spots 4-6 (middle):** Highest slugging — power hitters who drive in runs
- **Spots 7+ (bottom):** Remaining players by overall composite score
- **Stability:** Players move at most 2 spots from their average position over the last 3 games
- **New players (<3 games):** Default to roster order until enough data is collected

## Stats Tracked

- **Fielding:** Innings at each position, sit-outs, games since last played each position
- **Batting:** AB, 1B, 2B, 3B, HR, BB, SB, CS, OBP, SLG

## Game Entry Layout

The Game Entry sheet is organized top to bottom:

1. **Game info** (rows 1-3) — Date, Opponent, Innings
2. **Attendance** (rows 5-17) — Checkbox + player name for each roster player; uncheck absent players
3. **Lineup grid** (rows 19-28) — Position assignments and sit-outs per inning
4. **Batting stats** (rows 30+) — Per-player at-bats, hits, walks, steals

## Updating the Code

To update `Code.gs` without losing data:

1. Open **Extensions > Apps Script** and paste the new code over the old
2. Save, then use **⚾ Softball > Rebuild Game Entry** to update the Game Entry layout
3. Season History, Batting Stats, Roster, and Depth Chart are all preserved

## Notes

- The `onEdit` trigger auto-updates dropdowns when you change roster names
- The **⚾ Softball** menu appears at the far right of the menu bar (after Help) on each open
- Dashboard colors: Yellow = 3+ games since, Red = 5+ games since playing a position
- **Attendance:** Uncheck absent players on Game Entry before saving — they are excluded from season history and don't affect recency scoring
- **Batting Stats corrections:** You can edit the Batting Stats sheet directly to fix errors, then Refresh Dashboard
- **Lineup Suggester:** Player names refresh automatically from the roster each time you generate a lineup
- **Rebuild Game Entry:** Use this menu option after code updates to refresh the Game Entry layout without affecting other sheets
