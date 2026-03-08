# Softball Lineup Manager

A Google Apps Script tool for managing recreational softball lineups, fielding rotations, and batting orders. Designed for coaches who want fair playing time, smart position assignments, and stat-driven batting orders.

## Features

- **Roster Management** — Track up to 12 players with per-position preferences (Preferred / Okay / Restricted)
- **Depth Chart** — Rank players at each position to guide the lineup algorithm
- **Game Entry** — Record each game's lineup and batting stats with dropdowns, attendance checkboxes, and multiple sit-out columns
- **Lineup Suggester** — Auto-generate balanced field lineups with rest options, relief pitcher suggestions, and fair sit-out rotation
- **Lineup Card** — Combined coach-friendly view: players in batting order, positions per inning, with OBP/SLG stats
- **Batting Order** — Suggest optimal batting orders based on OBP, slugging, and speed stats
- **Season Dashboard** — View innings at each position, recency tracking, and cumulative batting stats
- **Delete Last Game** — Undo a saved game if entered incorrectly

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
   - **Re-running is safe:** Initialize All Sheets preserves your existing roster names, preferences, and depth chart rankings

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
| **Lineup Suggester** | Auto-generate lineup card, field lineups, and batting orders |
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
- **Minimum 2-inning start:** Starting Pitcher and Catcher are locked in for at least the first 2 innings
- **P/C continuity:** Pitcher and Catcher get a stronger continuity bonus than field positions, since leaving is permanent
- **Field position rotation:** Players get a small bonus for a 2nd consecutive inning at the same position, but a growing penalty for 3+ innings to encourage rotation across the field
- **Outfield-only avoidance:** Players who have only played outfield (LF/CF/RF) for 2+ innings get a bonus toward infield positions to mix things up
- **Sit-out cap:** No player sits out more than their fair share — the cap is calculated as `ceil(total sit-out slots / players)`, so with 12 players and 5 innings, no one sits more than 2
- Sit-outs rotate fairly — avoids consecutive sit-outs for the same player, and proactively sits out the next depth-chart pitcher to enable warmup
- **Relief pitcher:** The output suggests a relief pitcher from the depth chart in case the starter needs to come out
- **Weekly IP tracking:** The lineup output shows each pitcher's rolling 7-day innings pitched count (prior + this game)
- **Position diversity:** Small bonus for positions a player has never or rarely played this season — preferences and depth chart rankings always take priority
- **Attendance equity:** Sit-out fairness uses per-game rate (not raw count), so missed games don't skew the rotation. Missing a game applies a small position-assignment penalty as a tiebreaker
- Recency is per-player — absent games don't inflate "games since last played"

### Rest Flags (P / C)

- On the Lineup Suggester, check **Rest P** or **Rest C** next to a player to hold them back from Pitcher or Catcher for that game
- Useful for friendlies, early tournament games, or resting arms for a later bracket game
- The player still plays all other positions — only the checked position is blocked
- **Validation:** If rest flags + roster restrictions leave too few pitchers or catchers, the system warns you before generating so you can adjust
- Rest flags reset when you change the roster but are preserved between lineup generations

### Lineup Card (Combined View)

When you run Suggest Lineup, the first output is a **Lineup Card** — a single grid combining batting order and fielding:

| # | Player | Inn 1 | Inn 2 | Inn 3 | ... | OBP | SLG |
|---|--------|-------|-------|-------|-----|-----|-----|
| 1 | Alice  | P     | P     | SS    | ... | .400 | .500 |
| 2 | Beth   | C     | C     | C     | ... | .350 | .400 |

- Rows are players in **batting order** — read top to bottom for your batting lineup
- Inning columns show each player's **field position** that inning
- **OUT** (gray background) = player sits out that inning
- **Green background** = player is at a Preferred position
- OBP and SLG stats appear at the right edge
- Summary info (sit-out cap, relief pitcher) appears directly below the card
- Position dropdowns on each inning cell allow manual edits
- The old **position-centric grid** (Suggested Field Lineup) is preserved below for easy copy-paste to Game Entry

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

## Save Game Validation

When you click Save Game, the system checks for errors before saving:

- **Duplicate players** — If the same player appears in multiple positions in one inning, the save is blocked with a specific error message
- **Absent players in lineup** — If a player marked absent (unchecked attendance) is assigned to a position, the save is blocked
- Fix the errors and save again

## Deleting a Game

If you saved a game with errors:

1. Click **⚾ Softball > Delete Last Game**
2. Confirm the game number and opponent
3. The game is removed from Season History and Batting Stats, and the Dashboard is refreshed

## Updating the Code

To update `Code.gs` without losing data:

1. Open **Extensions > Apps Script** and paste the new code over the old
2. Save, then run `rebuildGameEntry` from the function dropdown to update the Game Entry layout
3. To refresh the Lineup Suggester layout (e.g., for new Rest P/C columns), run `initializeStep2` from the function dropdown
4. Season History, Batting Stats, Roster, and Depth Chart are all preserved

## Notes

- The `onEdit` trigger auto-updates dropdowns when you change roster names
- The **⚾ Softball** menu appears at the far right of the menu bar (after Help) on each open
- Dashboard colors: Yellow = 3+ games since, Red = 5+ games since playing a position
- **Attendance:** Uncheck absent players on Game Entry before saving — they are excluded from season history and don't affect recency scoring
- **Batting Stats corrections:** You can edit the Batting Stats sheet directly to fix errors, then Refresh Dashboard
- **Lineup Suggester:** Player names refresh automatically from the roster each time you generate a lineup
- **Rebuild Game Entry:** Use this menu option after code updates to refresh the Game Entry layout without affecting other sheets
