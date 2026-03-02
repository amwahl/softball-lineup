// ============================================================
// SOFTBALL LINEUP MANAGER - Google Apps Script
// ============================================================

const POSITIONS = ['P', 'C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF'];
const MAX_PLAYERS = 12;
const PREF_OPTIONS = ['Preferred', 'Okay', 'Restricted'];
const BATTING_STATS = 'Batting Stats';

// ============================================================
// MENU & INITIALIZATION
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚾ Softball')
    .addItem('Initialize All Sheets', 'initializeAll')
    .addSeparator()
    .addItem('Save Game', 'saveGame')
    .addItem('Suggest Lineup', 'suggestLineup')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addToUi();
}

function initializeAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  initializeStep1();
  initializeStep2();

  // Delete default Sheet1 if it exists
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  ss.getSheetByName('Roster').activate();
  ui.alert('Setup Complete', 'All sheets have been created. Start by entering your roster on the Roster sheet.', ui.ButtonSet.OK);
}

function initializeStep1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createRosterSheet(ss);
  createGameEntrySheet(ss);
  createSeasonHistorySheet(ss);
  createBattingStatsSheet(ss);
}

function initializeStep2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createDashboardSheet(ss);
  createDepthChartSheet(ss);
  createLineupSuggesterSheet(ss);
  createHowToUseSheet(ss);
}

// ============================================================
// ROSTER / CONFIG SHEET
// ============================================================

function createRosterSheet(ss) {
  let sheet = ss.getSheetByName('Roster');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Roster');

  // Header row
  const headers = ['#', 'Player Name'];
  POSITIONS.forEach(p => headers.push(p));
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // Player rows - batch write numbers and defaults
  const playerNums = [];
  const defaultPrefs = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    playerNums.push([i + 1]);
    const row = [];
    for (let j = 0; j < POSITIONS.length; j++) row.push('Okay');
    defaultPrefs.push(row);
  }
  sheet.getRange(2, 1, MAX_PLAYERS, 1).setValues(playerNums);
  sheet.getRange(2, 3, MAX_PLAYERS, POSITIONS.length).setValues(defaultPrefs);

  // Position preference dropdowns - apply validation to entire range at once
  const prefValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PREF_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, MAX_PLAYERS, POSITIONS.length).setDataValidation(prefValidationRule);

  // Formatting - minimal column widths
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 180);

  // Conditional formatting for preferences
  const prefRange = sheet.getRange(2, 3, MAX_PLAYERS, POSITIONS.length);
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Preferred').setBackground('#b7e1cd').setFontColor('#137333').setRanges([prefRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Okay').setBackground('#fce8b2').setFontColor('#7f6003').setRanges([prefRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Restricted').setBackground('#f4c7c3').setFontColor('#a50e0e').setRanges([prefRange]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);
  sheet.getRange(2, 2, MAX_PLAYERS, 1).setFontSize(12);
}

// ============================================================
// GAME ENTRY SHEET
// ============================================================

function createGameEntrySheet(ss) {
  let sheet = ss.getSheetByName('Game Entry');
  if (sheet) { sheet.clear(); sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); }
  else sheet = ss.insertSheet('Game Entry');

  // Game info section
  sheet.getRange('A1').setValue('Date:').setFontWeight('bold');
  sheet.getRange('B1').setNumberFormat('mm/dd/yyyy');
  sheet.getRange('A2').setValue('Opponent:').setFontWeight('bold');
  sheet.getRange('A3').setValue('Innings:').setFontWeight('bold');
  sheet.getRange('B3').setValue(6);

  const inningsValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 9).setAllowInvalid(false).build();
  sheet.getRange('B3').setDataValidation(inningsValidation);

  // Instructions
  sheet.getRange('D1').setValue('← Fill in game details, then click ⚾ Softball (far right of the menu bar, after Help) > Save Game')
    .setFontColor('#666666').setFontStyle('italic');

  // Attendance checkboxes (left sidebar, rows 5+)
  const attendRow = 5;
  sheet.getRange(attendRow, 1).setValue('Attendance')
    .setFontWeight('bold').setFontSize(12);
  const players = getRosterNames();
  const attendCheckVals = [];
  const attendNameVals = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    attendCheckVals.push([i < players.length && players[i] ? true : false]);
    attendNameVals.push([i < players.length && players[i] ? players[i] : '']);
  }
  sheet.getRange(attendRow + 1, 1, MAX_PLAYERS, 1).insertCheckboxes();
  sheet.getRange(attendRow + 1, 1, MAX_PLAYERS, 1).setValues(attendCheckVals);
  sheet.getRange(attendRow + 1, 2, MAX_PLAYERS, 1).setValues(attendNameVals).setFontSize(11);

  // Lineup grid header (row after attendance section)
  const gridHeaderRow = attendRow + MAX_PLAYERS + 2; // 5 + 12 + 2 = 19
  const maxSitOut = MAX_PLAYERS - POSITIONS.length; // typically 3
  const gridHeaders = ['Inning'];
  POSITIONS.forEach(p => gridHeaders.push(p));
  for (let s = 1; s <= maxSitOut; s++) {
    gridHeaders.push('Sit Out ' + s);
  }
  sheet.getRange(gridHeaderRow, 1, 1, gridHeaders.length).setValues([gridHeaders])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // Inning rows (max 9) - batch write
  const gridDataRow = gridHeaderRow + 1;
  const inningNums = [[1],[2],[3],[4],[5],[6],[7],[8],[9]];
  sheet.getRange(gridDataRow, 1, 9, 1).setValues(inningNums).setHorizontalAlignment('center').setFontWeight('bold');

  // Batting stats section (after lineup grid)
  const statsStartRow = gridDataRow + 9 + 1; // after 9 inning rows + gap
  sheet.getRange(statsStartRow, 1).setValue('Batting Stats')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');

  const statsHeaders = ['Player', 'AB', '1B', '2B', '3B', 'HR', 'BB', 'SB', 'CS'];
  sheet.getRange(statsStartRow + 1, 1, 1, statsHeaders.length).setValues([statsHeaders])
    .setFontWeight('bold')
    .setBackground('#fbbc04')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // Player rows for stats - use roster names from attendance section above
  const statsNames = [];
  const statsDefaults = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    statsNames.push([i < players.length ? players[i] : '']);
    statsDefaults.push([0, 0, 0, 0, 0, 0, 0, 0]);
  }
  sheet.getRange(statsStartRow + 2, 1, MAX_PLAYERS, 1).setValues(statsNames).setFontSize(11);
  sheet.getRange(statsStartRow + 2, 2, MAX_PLAYERS, 8).setValues(statsDefaults).setHorizontalAlignment('center');

  // Numeric validation for stat columns
  const numRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 99).setAllowInvalid(false).build();
  sheet.getRange(statsStartRow + 2, 2, MAX_PLAYERS, 8).setDataValidation(numRule);

  sheet.setFrozenRows(3);
  updateGameEntryDropdowns();
}

function updateGameEntryDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Game Entry');
  const players = getRosterNames();

  if (players.length === 0) return;

  // Row calculations must match createGameEntrySheet layout
  const attendRow = 5;
  const gridDataRow = attendRow + MAX_PLAYERS + 2 + 1; // 20

  // Set dropdowns for position cells - apply to entire range at once
  const posRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(players, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(gridDataRow, 2, 9, POSITIONS.length).setDataValidation(posRule);

  // Sat out columns - apply to all sit-out columns
  const maxSitOut = MAX_PLAYERS - POSITIONS.length;
  const satOutRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(players, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(gridDataRow, POSITIONS.length + 2, 9, maxSitOut).setDataValidation(satOutRule);

  // Update player names in batting stats and attendance sections
  const statsStartRow = gridDataRow + 9 + 1;

  // Update attendance names
  const attendNameVals = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    attendNameVals.push([i < players.length ? players[i] : '']);
  }
  sheet.getRange(attendRow + 1, 2, MAX_PLAYERS, 1).setValues(attendNameVals);
  const nameValues = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    nameValues.push([i < players.length ? players[i] : '']);
  }
  sheet.getRange(statsStartRow + 2, 1, MAX_PLAYERS, 1).setValues(nameValues);
}

// ============================================================
// SEASON HISTORY SHEET (hidden data store)
// ============================================================

function createSeasonHistorySheet(ss) {
  let sheet = ss.getSheetByName('Season History');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Season History');

  const headers = ['Game #', 'Date', 'Opponent', 'Innings', 'Inning #', 'Player'];
  POSITIONS.forEach(p => headers.push(p));
  headers.push('Sat Out');

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('white');

  sheet.setFrozenRows(1);
}

// ============================================================
// BATTING STATS SHEET (hidden data store)
// ============================================================

function createBattingStatsSheet(ss) {
  let sheet = ss.getSheetByName(BATTING_STATS);
  if (sheet) sheet.clear(); else sheet = ss.insertSheet(BATTING_STATS);

  const headers = ['Game #', 'Date', 'Player', 'AB', '1B', '2B', '3B', 'HR', 'BB', 'SB', 'CS', 'BattingPos'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#fbbc04')
    .setFontColor('white');

  sheet.setFrozenRows(1);
}

// ============================================================
// SAVE GAME
// ============================================================

function saveGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const gameSheet = ss.getSheetByName('Game Entry');
  const historySheet = ss.getSheetByName('Season History');

  if (!gameSheet || !historySheet) {
    ui.alert('Error', 'Please run Initialize All Sheets first.', ui.ButtonSet.OK);
    return;
  }

  const date = gameSheet.getRange('B1').getValue();
  const opponent = gameSheet.getRange('B2').getValue();
  const innings = gameSheet.getRange('B3').getValue();

  if (!date || !opponent) {
    ui.alert('Missing Info', 'Please fill in the Date and Opponent fields.', ui.ButtonSet.OK);
    return;
  }

  // Determine next game number
  const historyData = historySheet.getDataRange().getValues();
  let maxGame = 0;
  for (let i = 1; i < historyData.length; i++) {
    if (historyData[i][0] > maxGame) maxGame = historyData[i][0];
  }
  const gameNum = maxGame + 1;

  // Collect lineup data - batch read all game data at once
  const players = getRosterNames();
  const rows = [];
  const maxSitOut = MAX_PLAYERS - POSITIONS.length;
  const attendRow = 5;
  const gridDataRow = attendRow + MAX_PLAYERS + 2 + 1; // 20
  const gameData = gameSheet.getRange(gridDataRow, 2, 9, POSITIONS.length + maxSitOut).getValues();

  // Read attendance checkboxes to identify absent players
  const attendData = gameSheet.getRange(attendRow + 1, 1, MAX_PLAYERS, 1).getValues();
  const absentPlayers = new Set();
  for (let i = 0; i < players.length; i++) {
    if (players[i] && attendData[i][0] === false) {
      absentPlayers.add(players[i]);
    }
  }

  for (let inning = 1; inning <= innings; inning++) {
    const inningData = gameData[inning - 1];
    for (let p = 0; p < players.length; p++) {
      const playerName = players[p];
      // Skip absent players — they don't get history rows
      if (absentPlayers.has(playerName)) continue;
      const dataRow = [gameNum, date, opponent, innings, inning, playerName];

      for (let j = 0; j < POSITIONS.length; j++) {
        dataRow.push(inningData[j] === playerName ? 1 : 0);
      }

      // Check sat out — check all sit-out columns or if not assigned any position
      let isSatOut = !POSITIONS.some((_, j) => inningData[j] === playerName);
      for (let s = 0; s < maxSitOut; s++) {
        const satOutVal = (inningData[POSITIONS.length + s] || '').toString();
        if (satOutVal === playerName) isSatOut = true;
      }
      dataRow.push(isSatOut ? 1 : 0);

      rows.push(dataRow);
    }
  }

  // Write to history
  if (rows.length > 0) {
    const startRow = historySheet.getLastRow() + 1;
    historySheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Save batting stats
  saveBattingStats(ss, gameSheet, gameNum, date, players);

  // Clear game entry for next use
  gameSheet.getRange('B1').clearContent();
  gameSheet.getRange('B2').clearContent();
  const maxSitOutClear = MAX_PLAYERS - POSITIONS.length;
  gameSheet.getRange(gridDataRow, 2, 9, POSITIONS.length + maxSitOutClear).clearContent();

  // Reset attendance checkboxes to all checked
  const resetChecks = [];
  for (let i = 0; i < MAX_PLAYERS; i++) resetChecks.push([true]);
  gameSheet.getRange(attendRow + 1, 1, MAX_PLAYERS, 1).setValues(resetChecks);

  // Clear batting stats section
  const statsStartRow = gridDataRow + 9 + 1;
  gameSheet.getRange(statsStartRow + 2, 2, MAX_PLAYERS, 8).setValue(0);

  // Refresh dashboard
  refreshDashboard();

  ui.alert('Game Saved!', 'Game #' + gameNum + ' vs ' + opponent + ' has been saved.\nDashboard has been refreshed.', ui.ButtonSet.OK);
}

function saveBattingStats(ss, gameSheet, gameNum, date, players) {
  const battingSheet = ss.getSheetByName(BATTING_STATS);
  if (!battingSheet) return;

  const attendRow = 5;
  const gridDataRow = attendRow + MAX_PLAYERS + 2 + 1;
  const statsStartRow = gridDataRow + 9 + 1;
  // Batch read: player names (col 1) + stats (cols 2-9)
  const statsData = gameSheet.getRange(statsStartRow + 2, 1, MAX_PLAYERS, 9).getValues();

  const battingRows = [];
  for (let i = 0; i < statsData.length; i++) {
    const playerName = statsData[i][0];
    if (!playerName || playerName.toString().trim() === '') continue;

    const ab = statsData[i][1] || 0;
    // Only save rows where the player had at least 1 AB or BB
    const bb = statsData[i][6] || 0;
    if (ab === 0 && bb === 0) continue;

    const battingPos = i + 1; // Default batting position = roster order
    battingRows.push([
      gameNum,
      date,
      playerName,
      ab,              // AB
      statsData[i][2] || 0, // 1B
      statsData[i][3] || 0, // 2B
      statsData[i][4] || 0, // 3B
      statsData[i][5] || 0, // HR
      bb,              // BB
      statsData[i][7] || 0, // SB
      statsData[i][8] || 0, // CS
      battingPos
    ]);
  }

  if (battingRows.length > 0) {
    const startRow = battingSheet.getLastRow() + 1;
    battingSheet.getRange(startRow, 1, battingRows.length, battingRows[0].length).setValues(battingRows);
  }
}

// ============================================================
// BATTING STATS COMPUTATION
// ============================================================

function computeBattingAverages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const battingSheet = ss.getSheetByName(BATTING_STATS);
  if (!battingSheet) return {};

  const data = battingSheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  // Columns: Game#, Date, Player, AB, 1B, 2B, 3B, HR, BB, SB, CS, BattingPos
  const playerStats = {};

  for (let i = 1; i < data.length; i++) {
    const player = data[i][2];
    if (!player) continue;

    if (!playerStats[player]) {
      playerStats[player] = {
        games: 0, ab: 0, singles: 0, doubles: 0, triples: 0, hr: 0,
        bb: 0, sb: 0, cs: 0,
        recentPositions: [] // [{gameNum, pos}] for stability calc
      };
    }

    const s = playerStats[player];
    s.games++;
    s.ab += data[i][3] || 0;
    s.singles += data[i][4] || 0;
    s.doubles += data[i][5] || 0;
    s.triples += data[i][6] || 0;
    s.hr += data[i][7] || 0;
    s.bb += data[i][8] || 0;
    s.sb += data[i][9] || 0;
    s.cs += data[i][10] || 0;
    s.recentPositions.push({ gameNum: data[i][0], pos: data[i][11] || 0 });
  }

  // Compute derived stats
  const averages = {};
  for (const player in playerStats) {
    const s = playerStats[player];
    const hits = s.singles + s.doubles + s.triples + s.hr;
    const totalBases = s.singles + (s.doubles * 2) + (s.triples * 3) + (s.hr * 4);
    const pa = s.ab + s.bb; // plate appearances (simplified, no HBP/SF)

    const obp = pa > 0 ? (hits + s.bb) / pa : 0;
    const slg = s.ab > 0 ? totalBases / s.ab : 0;
    const baserunning = (s.sb * 1.5) - (s.cs * 2); // SB bonus with CS penalty

    // Average batting position over last 3 games for stability
    const sorted = s.recentPositions.slice().sort((a, b) => b.gameNum - a.gameNum);
    const last3 = sorted.slice(0, 3);
    const avgPos = last3.length > 0
      ? last3.reduce((sum, r) => sum + r.pos, 0) / last3.length
      : 0;

    averages[player] = {
      games: s.games,
      ab: s.ab,
      hits: hits,
      obp: obp,
      slg: slg,
      sb: s.sb,
      cs: s.cs,
      baserunning: baserunning,
      avgBattingPos: avgPos
    };
  }

  return averages;
}

// ============================================================
// BATTING ORDER ALGORITHM
// ============================================================

function generateBattingOrder(availablePlayers, battingAverages) {
  // Separate players with enough data from new players
  const withData = [];
  const newPlayers = [];

  for (const player of availablePlayers) {
    const avg = battingAverages[player];
    if (avg && avg.games >= 3) {
      withData.push({ name: player, stats: avg });
    } else {
      newPlayers.push(player);
    }
  }

  // If no data, return roster order
  if (withData.length === 0) {
    return availablePlayers.map((name, idx) => ({
      name: name,
      position: idx + 1,
      obp: 0,
      slg: 0,
      sb: 0
    }));
  }

  // Score players for each slot category
  // Top of order (1-3): OBP + baserunning (get on base and steal)
  // Middle (4-6): slugging (power hitters)
  // Bottom (7+): overall composite
  const topScore = (s) => (s.obp * 100) + (s.baserunning * 5);
  const midScore = (s) => (s.slg * 100) + (s.obp * 30);
  const overallScore = (s) => (s.obp * 50) + (s.slg * 50) + (s.baserunning * 3);

  // Sort for top of order
  const topCandidates = withData.slice().sort((a, b) => topScore(b.stats) - topScore(a.stats));
  // Sort for middle
  const midCandidates = withData.slice().sort((a, b) => midScore(b.stats) - midScore(a.stats));

  const totalSlots = availablePlayers.length;
  const topSlots = Math.min(3, totalSlots);
  const midSlots = Math.min(3, Math.max(0, totalSlots - 3));

  // Assign slots greedily
  const assigned = new Set();
  const order = new Array(totalSlots).fill(null);

  // Top of order
  let slot = 0;
  for (const c of topCandidates) {
    if (slot >= topSlots) break;
    if (assigned.has(c.name)) continue;
    order[slot] = c;
    assigned.add(c.name);
    slot++;
  }

  // Middle of order
  slot = topSlots;
  for (const c of midCandidates) {
    if (slot >= topSlots + midSlots) break;
    if (assigned.has(c.name)) continue;
    order[slot] = c;
    assigned.add(c.name);
    slot++;
  }

  // Bottom: remaining with-data players by overall score
  const remaining = withData.filter(c => !assigned.has(c.name))
    .sort((a, b) => overallScore(b.stats) - overallScore(a.stats));
  slot = topSlots + midSlots;
  for (const c of remaining) {
    if (slot >= totalSlots) break;
    order[slot] = c;
    assigned.add(c.name);
    slot++;
  }

  // Fill remaining slots with new players (roster order)
  for (const name of newPlayers) {
    if (slot >= totalSlots) break;
    const avg = battingAverages[name] || { obp: 0, slg: 0, sb: 0 };
    order[slot] = { name: name, stats: avg };
    slot++;
  }

  // Fill any remaining nulls (shouldn't happen, but safety)
  for (let i = 0; i < totalSlots; i++) {
    if (!order[i]) {
      const unplaced = availablePlayers.find(p => !assigned.has(p));
      if (unplaced) {
        const avg = battingAverages[unplaced] || { obp: 0, slg: 0, sb: 0 };
        order[i] = { name: unplaced, stats: avg };
        assigned.add(unplaced);
      }
    }
  }

  // Stability: players move at most 2 spots from their average position over last 3 games
  for (let i = 0; i < order.length; i++) {
    if (!order[i]) continue;
    const avg = battingAverages[order[i].name];
    if (avg && avg.avgBattingPos > 0) {
      const targetPos = i + 1;
      const diff = Math.abs(targetPos - avg.avgBattingPos);
      if (diff > 2) {
        // Try to swap with someone closer to this player's avg position
        const idealSlot = Math.round(avg.avgBattingPos) - 1;
        const clampedSlot = Math.max(0, Math.min(order.length - 1, idealSlot));
        // Only swap if the other player also benefits or doesn't get worse
        if (order[clampedSlot]) {
          const otherAvg = battingAverages[order[clampedSlot].name];
          const otherIdeal = otherAvg && otherAvg.avgBattingPos > 0 ? Math.round(otherAvg.avgBattingPos) - 1 : clampedSlot;
          const currentOtherDist = Math.abs(clampedSlot - otherIdeal);
          const newOtherDist = Math.abs(i - otherIdeal);
          if (newOtherDist <= currentOtherDist + 1) {
            // Swap
            const temp = order[i];
            order[i] = order[clampedSlot];
            order[clampedSlot] = temp;
          }
        }
      }
    }
  }

  return order.map((entry, idx) => ({
    name: entry ? entry.name : '',
    position: idx + 1,
    obp: entry && entry.stats ? entry.stats.obp || 0 : 0,
    slg: entry && entry.stats ? entry.stats.slg || 0 : 0,
    sb: entry && entry.stats ? entry.stats.sb || 0 : 0
  }));
}

// ============================================================
// DASHBOARD SHEET
// ============================================================

function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName('Dashboard');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Dashboard');

  // Title
  sheet.getRange('A1').setValue('Season Dashboard')
    .setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Click ⚾ Softball (far right of menu bar, after Help) > Refresh Dashboard to update')
    .setFontColor('#666666').setFontStyle('italic');

  // Section 1: Innings at Each Position
  sheet.getRange('A4').setValue('Innings at Each Position')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
  const posHeaders1 = ['Player'];
  POSITIONS.forEach(p => posHeaders1.push(p));
  posHeaders1.push('Total Played', 'Sat Out');
  sheet.getRange(5, 1, 1, posHeaders1.length).setValues([posHeaders1])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');

  // Section 2: Games Since Last Played
  const sec2StartRow = 5 + MAX_PLAYERS + 2;
  sheet.getRange(sec2StartRow, 1).setValue('Games Since Last at Position')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
  const posHeaders2 = ['Player'];
  POSITIONS.forEach(p => posHeaders2.push(p));
  sheet.getRange(sec2StartRow + 1, 1, 1, posHeaders2.length).setValues([posHeaders2])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');

  // Section 3: Batting Stats
  const sec3StartRow = sec2StartRow + MAX_PLAYERS + 3;
  sheet.getRange(sec3StartRow, 1).setValue('Batting Stats')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
  const battingHeaders = ['Player', 'Games', 'AB', 'H', 'OBP', 'SLG', 'SB', 'CS'];
  sheet.getRange(sec3StartRow + 1, 1, 1, battingHeaders.length).setValues([battingHeaders])
    .setFontWeight('bold').setBackground('#fbbc04').setFontColor('white').setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 180);
  sheet.setFrozenRows(0);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');
  const historySheet = ss.getSheetByName('Season History');

  if (!dashboard || !historySheet) return;

  const players = getRosterNames();
  if (players.length === 0) return;

  const historyData = historySheet.getDataRange().getValues();
  if (historyData.length <= 1) return;

  // Parse history - columns: Game#, Date, Opponent, Innings, Inning#, Player, P, C, 1B, 2B, 3B, SS, LF, CF, RF, SatOut
  const posStartCol = 6; // index in historyData

  // Build stats per player
  const stats = {};

  for (let i = 1; i < historyData.length; i++) {
    const gameNum = historyData[i][0];
    const playerName = historyData[i][5];

    if (!stats[playerName]) {
      stats[playerName] = {
        positionInnings: new Array(POSITIONS.length).fill(0),
        satOutInnings: 0,
        lastGameAtPosition: new Array(POSITIONS.length).fill(0),
        totalInnings: 0,
        gamesPlayed: new Set()
      };
    }

    const s = stats[playerName];
    let playedThisInning = false;

    for (let j = 0; j < POSITIONS.length; j++) {
      if (historyData[i][posStartCol + j] === 1) {
        s.positionInnings[j]++;
        s.lastGameAtPosition[j] = gameNum;
        playedThisInning = true;
      }
    }

    if (playedThisInning) {
      s.totalInnings++;
    }
    s.gamesPlayed.add(gameNum);

    if (historyData[i][posStartCol + POSITIONS.length] === 1) {
      s.satOutInnings++;
    }
  }

  // Section 1: Innings at each position (row 6+) - batch write
  const sec1Data = [];
  for (let p = 0; p < players.length; p++) {
    const name = players[p];
    const s = stats[name];
    const row = [name];
    if (s) {
      for (let j = 0; j < POSITIONS.length; j++) row.push(s.positionInnings[j]);
      row.push(s.totalInnings, s.satOutInnings);
    } else {
      for (let j = 0; j < POSITIONS.length + 2; j++) row.push(0);
    }
    sec1Data.push(row);
  }
  if (sec1Data.length > 0) {
    const sec1Range = dashboard.getRange(6, 1, sec1Data.length, sec1Data[0].length);
    sec1Range.setValues(sec1Data);
    dashboard.getRange(6, 1, sec1Data.length, 1).setFontSize(11);
    dashboard.getRange(6, 2, sec1Data.length, POSITIONS.length + 2).setHorizontalAlignment('center');
    dashboard.getRange(6, POSITIONS.length + 2, sec1Data.length, 1).setFontWeight('bold');
  }

  // Section 2: Games since last played - batch write
  const sec2StartRow = 5 + MAX_PLAYERS + 2;
  const sec2DataRow = sec2StartRow + 2;
  const sec2Data = [];
  const sec2Backgrounds = [];

  for (let p = 0; p < players.length; p++) {
    const name = players[p];
    const s = stats[name];
    const row = [name];
    const bgRow = [null];
    if (s) {
      // Build ordinal map for this player so absent games don't inflate recency
      const sortedGames = Array.from(s.gamesPlayed).sort((a, b) => a - b);
      const gameOrdinal = {};
      sortedGames.forEach((g, idx) => gameOrdinal[g] = idx + 1);
      const playerTotalGames = sortedGames.length;

      for (let j = 0; j < POSITIONS.length; j++) {
        const lastGame = s.lastGameAtPosition[j];
        const lastOrd = lastGame > 0 && gameOrdinal[lastGame] ? gameOrdinal[lastGame] : 0;
        const gamesSince = lastOrd > 0 ? playerTotalGames - lastOrd : (playerTotalGames > 0 ? playerTotalGames : 0);
        row.push(gamesSince);
        if (gamesSince >= 5) bgRow.push('#f4c7c3');
        else if (gamesSince >= 3) bgRow.push('#fce8b2');
        else bgRow.push(null);
      }
    } else {
      for (let j = 0; j < POSITIONS.length; j++) { row.push(0); bgRow.push(null); }
    }
    sec2Data.push(row);
    sec2Backgrounds.push(bgRow);
  }
  if (sec2Data.length > 0) {
    const sec2Range = dashboard.getRange(sec2DataRow, 1, sec2Data.length, sec2Data[0].length);
    sec2Range.setValues(sec2Data);
    sec2Range.setBackgrounds(sec2Backgrounds);
    dashboard.getRange(sec2DataRow, 1, sec2Data.length, 1).setFontSize(11);
    dashboard.getRange(sec2DataRow, 2, sec2Data.length, POSITIONS.length).setHorizontalAlignment('center');
  }

  // Section 3: Batting stats - batch write
  const battingAverages = computeBattingAverages();
  const sec3StartRow = sec2StartRow + MAX_PLAYERS + 3;
  const sec3DataRow = sec3StartRow + 2;
  const sec3Data = [];

  for (let p = 0; p < players.length; p++) {
    const name = players[p];
    const avg = battingAverages[name];
    if (avg) {
      sec3Data.push([
        name,
        avg.games,
        avg.ab,
        avg.hits,
        avg.obp.toFixed(3),
        avg.slg.toFixed(3),
        avg.sb,
        avg.cs
      ]);
    } else {
      sec3Data.push([name, 0, 0, 0, '.000', '.000', 0, 0]);
    }
  }
  if (sec3Data.length > 0) {
    const sec3Range = dashboard.getRange(sec3DataRow, 1, sec3Data.length, sec3Data[0].length);
    sec3Range.setValues(sec3Data);
    dashboard.getRange(sec3DataRow, 1, sec3Data.length, 1).setFontSize(11);
    dashboard.getRange(sec3DataRow, 2, sec3Data.length, 7).setHorizontalAlignment('center');
  }
}

// ============================================================
// DEPTH CHART SHEET
// ============================================================

function createDepthChartSheet(ss) {
  let sheet = ss.getSheetByName('Depth Chart');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Depth Chart');

  // Title
  sheet.getRange('A1').setValue('Depth Chart')
    .setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Rank players per position (1st = top choice). Leave blank for unranked.')
    .setFontColor('#666666').setFontStyle('italic');

  // Header row (row 4): Rank label + position columns
  const headers = ['Rank'];
  POSITIONS.forEach(p => headers.push(p));
  sheet.getRange(4, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // Rank labels (rows 5-16)
  const rankLabels = [];
  const ordinals = ['1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th','11th','12th'];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    rankLabels.push([ordinals[i]]);
  }
  sheet.getRange(5, 1, MAX_PLAYERS, 1).setValues(rankLabels)
    .setFontWeight('bold').setHorizontalAlignment('center');

  // Player name dropdowns (populated when roster has names)
  const players = getRosterNames();
  if (players.length > 0) {
    const dropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(players, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(5, 2, MAX_PLAYERS, POSITIONS.length).setDataValidation(dropdownRule);
  }

  // Formatting
  sheet.setColumnWidth(1, 60);
  for (let c = 2; c <= POSITIONS.length + 1; c++) {
    sheet.setColumnWidth(c, 120);
  }

  sheet.setFrozenRows(4);
}

function updateDepthChartDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Depth Chart');
  if (!sheet) return;

  const players = getRosterNames();
  if (players.length === 0) return;

  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(players, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(5, 2, MAX_PLAYERS, POSITIONS.length).setDataValidation(dropdownRule);
}

// ============================================================
// LINEUP SUGGESTER SHEET
// ============================================================

function createLineupSuggesterSheet(ss) {
  let sheet = ss.getSheetByName('Lineup Suggester');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('Lineup Suggester');

  // Title and inputs
  sheet.getRange('A1').setValue('Lineup Suggester')
    .setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Click ⚾ Softball (far right of menu bar, after Help) > Suggest Lineup to generate')
    .setFontColor('#666666').setFontStyle('italic');

  sheet.getRange('A4').setValue('Innings:').setFontWeight('bold');
  sheet.getRange('B4').setValue(6);
  const inningsVal = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 9).setAllowInvalid(false).build();
  sheet.getRange('B4').setDataValidation(inningsVal);

  // Player availability checkboxes
  sheet.getRange('A6').setValue('Available Players:')
    .setFontWeight('bold').setFontSize(12);

  const players = getRosterNames();
  // Batch: insert checkboxes for entire range, then batch write values
  sheet.getRange(7, 1, MAX_PLAYERS, 1).insertCheckboxes();
  const checkVals = [];
  const nameVals = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    checkVals.push([i < players.length && players[i] ? true : false]);
    nameVals.push([i < players.length && players[i] ? players[i] : '']);
  }
  sheet.getRange(7, 1, MAX_PLAYERS, 1).setValues(checkVals);
  sheet.getRange(7, 2, MAX_PLAYERS, 1).setValues(nameVals).setFontSize(11);

  // Suggested lineup area (row 21+)
  const lineupStartRow = 7 + MAX_PLAYERS + 1;
  sheet.getRange(lineupStartRow, 1).setValue('Suggested Field Lineup')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');

  const gridHeaders = ['Inning'];
  POSITIONS.forEach(p => gridHeaders.push(p));
  gridHeaders.push('Sit Out');
  sheet.getRange(lineupStartRow + 1, 1, 1, gridHeaders.length).setValues([gridHeaders])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');

  // Batting order area (below field lineup)
  const battingStartRow = lineupStartRow + 2 + 9 + 2; // after max 9 innings + gap
  sheet.getRange(battingStartRow, 1).setValue('Suggested Batting Order')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');

  const battingHeaders = ['#', 'Player', 'OBP', 'SLG', 'SB'];
  sheet.getRange(battingStartRow + 1, 1, 1, battingHeaders.length).setValues([battingHeaders])
    .setFontWeight('bold').setBackground('#fbbc04').setFontColor('white').setHorizontalAlignment('center');
}

// ============================================================
// DEPTH CHART READER
// ============================================================

function getDepthChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Depth Chart');
  if (!sheet) return {};

  const data = sheet.getRange(5, 2, MAX_PLAYERS, POSITIONS.length).getValues();
  const depthChart = {};

  for (let j = 0; j < POSITIONS.length; j++) {
    const pos = POSITIONS[j];
    depthChart[pos] = [];
    for (let i = 0; i < MAX_PLAYERS; i++) {
      const name = data[i][j];
      if (name && name.toString().trim() !== '') {
        depthChart[pos].push(name.toString().trim());
      }
    }
  }

  return depthChart;
}

// ============================================================
// SUGGEST LINEUP ALGORITHM
// ============================================================

function suggestLineup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const suggesterSheet = ss.getSheetByName('Lineup Suggester');
  const historySheet = ss.getSheetByName('Season History');
  const rosterSheet = ss.getSheetByName('Roster');

  if (!suggesterSheet || !rosterSheet) {
    ui.alert('Error', 'Please run Initialize All Sheets first.', ui.ButtonSet.OK);
    return;
  }

  const innings = suggesterSheet.getRange('B4').getValue();
  if (!innings || innings < 1) {
    ui.alert('Error', 'Please enter a valid number of innings.', ui.ButtonSet.OK);
    return;
  }

  // Get available players - batch read
  const availablePlayers = [];
  const checkData = suggesterSheet.getRange(7, 1, MAX_PLAYERS, 2).getValues();
  for (let i = 0; i < MAX_PLAYERS; i++) {
    if (checkData[i][0] && checkData[i][1]) {
      availablePlayers.push(checkData[i][1]);
    }
  }

  if (availablePlayers.length < POSITIONS.length) {
    ui.alert('Not Enough Players',
      'You need at least ' + POSITIONS.length + ' available players to fill all positions. You have ' + availablePlayers.length + '.',
      ui.ButtonSet.OK);
    return;
  }

  // Get position preferences from roster
  const rosterData = rosterSheet.getRange(2, 2, MAX_PLAYERS, POSITIONS.length + 1).getValues();
  const preferences = {};
  for (let i = 0; i < rosterData.length; i++) {
    const name = rosterData[i][0];
    if (name && availablePlayers.indexOf(name) >= 0) {
      preferences[name] = {};
      for (let j = 0; j < POSITIONS.length; j++) {
        preferences[name][POSITIONS[j]] = rosterData[i][j + 1] || 'Okay';
      }
    }
  }

  // Get history stats for recency scoring
  // Uses per-player game counts so absent games don't inflate recency
  const gamesSinceAtPosition = {};
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();

    // First pass: collect each player's games in order
    const playerGameList = {}; // playerName -> sorted array of game numbers
    for (let i = 1; i < historyData.length; i++) {
      const gameNum = historyData[i][0];
      const playerName = historyData[i][5];
      if (availablePlayers.indexOf(playerName) < 0) continue;
      if (!playerGameList[playerName]) playerGameList[playerName] = new Set();
      playerGameList[playerName].add(gameNum);
    }

    // Build game-number-to-player-ordinal maps
    const playerGameIndex = {}; // playerName -> { gameNum -> ordinal (1-based) }
    for (const name of availablePlayers) {
      if (playerGameList[name]) {
        const sorted = Array.from(playerGameList[name]).sort((a, b) => a - b);
        playerGameIndex[name] = {};
        sorted.forEach((g, idx) => playerGameIndex[name][g] = idx + 1);
      }
    }

    // Second pass: find last game ordinal at each position per player
    const lastOrdinalAtPos = {};
    for (let i = 1; i < historyData.length; i++) {
      const gameNum = historyData[i][0];
      const playerName = historyData[i][5];
      if (availablePlayers.indexOf(playerName) < 0) continue;
      if (!lastOrdinalAtPos[playerName]) lastOrdinalAtPos[playerName] = {};
      const ordinal = playerGameIndex[playerName][gameNum];

      for (let j = 0; j < POSITIONS.length; j++) {
        if (historyData[i][6 + j] === 1) {
          if (!lastOrdinalAtPos[playerName][POSITIONS[j]] || ordinal > lastOrdinalAtPos[playerName][POSITIONS[j]]) {
            lastOrdinalAtPos[playerName][POSITIONS[j]] = ordinal;
          }
        }
      }
    }

    for (const name of availablePlayers) {
      const totalGamesPlayed = playerGameList[name] ? playerGameList[name].size : 0;
      gamesSinceAtPosition[name] = {};
      for (const pos of POSITIONS) {
        if (lastOrdinalAtPos[name] && lastOrdinalAtPos[name][pos]) {
          gamesSinceAtPosition[name][pos] = totalGamesPlayed - lastOrdinalAtPos[name][pos];
        } else {
          gamesSinceAtPosition[name][pos] = totalGamesPlayed > 0 ? totalGamesPlayed + 1 : 1;
        }
      }
    }
  }

  // Get depth chart rankings
  const depthChart = getDepthChart();

  // Also track how many total innings sat out historically per player
  const totalSatOut = {};
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    for (let i = 1; i < historyData.length; i++) {
      const playerName = historyData[i][5];
      const satOut = historyData[i][6 + POSITIONS.length];
      if (!totalSatOut[playerName]) totalSatOut[playerName] = 0;
      if (satOut === 1) totalSatOut[playerName]++;
    }
  }

  // Generate lineup inning by inning
  const lineup = []; // lineup[inning][posIndex] = playerName
  const sitOuts = []; // sitOuts[inning] = [playerNames]
  const inningCountThisGame = {}; // track innings played this game per player
  availablePlayers.forEach(p => inningCountThisGame[p] = 0);

  for (let inning = 0; inning < innings; inning++) {
    // Determine who sits out this inning
    // Players with most innings played this game (and historically least sat out) sit out
    const numSitOut = availablePlayers.length - POSITIONS.length;
    const sittingOut = [];

    if (numSitOut > 0) {
      // Identify who sat out last inning (for consecutive avoidance)
      const lastSitOuts = inning > 0 ? sitOuts[inning - 1] : [];

      // Pitcher-aware sit-out: find the next depth-chart pitcher who isn't
      // currently pitching, so they can warm up for a future transition
      let nextPitcher = null;
      if (depthChart && depthChart['P'] && inning > 0) {
        const currentPitcher = lineup[inning - 1][0];
        for (const candidate of depthChart['P']) {
          if (candidate !== currentPitcher && availablePlayers.indexOf(candidate) >= 0) {
            // This player hasn't pitched yet or has been away from P —
            // they're the most likely next pitcher
            const hasLeftP = lineup.some((inn, k) => inn[0] === candidate) &&
              lineup[inning - 1][0] !== candidate;
            if (!hasLeftP) {
              nextPitcher = candidate;
              break;
            }
          }
        }
      }

      // Sort by: most innings played this game first, then by least historical sat out
      const candidates = availablePlayers.slice().sort((a, b) => {
        // Primary: who has played the most innings this game
        const inningDiff = inningCountThisGame[b] - inningCountThisGame[a];
        if (inningDiff !== 0) return inningDiff;
        // Secondary: who has sat out the least historically
        return (totalSatOut[a] || 0) - (totalSatOut[b] || 0);
      });

      // Pick sit-outs, avoiding consecutive sit-outs when possible
      const consecutive = [];
      const nonConsecutive = [];
      for (const c of candidates) {
        if (lastSitOuts.indexOf(c) >= 0) {
          consecutive.push(c);
        } else {
          nonConsecutive.push(c);
        }
      }
      // Prefer non-consecutive first, fall back to consecutive
      const ordered = nonConsecutive.concat(consecutive);

      for (let s = 0; s < numSitOut && s < ordered.length; s++) {
        sittingOut.push(ordered[s]);
      }

      // Pitcher-aware: if the next pitcher isn't already sitting out and
      // there's room to swap them in, replace the last (least priority) sit-out
      if (nextPitcher && sittingOut.indexOf(nextPitcher) < 0) {
        // Only swap if the next pitcher isn't the current pitcher
        const currentPitcher = inning > 0 ? lineup[inning - 1][0] : null;
        if (nextPitcher !== currentPitcher) {
          // Replace the last sit-out (lowest priority) with the next pitcher
          sittingOut[sittingOut.length - 1] = nextPitcher;
        }
      }
    }

    sitOuts.push(sittingOut);

    const playing = availablePlayers.filter(p => sittingOut.indexOf(p) < 0);

    // Assign positions using a scoring system
    const assignment = assignPositions(playing, preferences, gamesSinceAtPosition, lineup, inning, depthChart);
    lineup.push(assignment);

    // Update inning counts
    playing.forEach(p => inningCountThisGame[p]++);
  }

  // Write suggested field lineup to sheet
  const lineupStartRow = 7 + MAX_PLAYERS + 1;

  // Clear previous suggestions
  const clearRange = suggesterSheet.getRange(lineupStartRow + 2, 1, 9, POSITIONS.length + 2);
  clearRange.clearContent().setBackground(null);

  // Batch write lineup data
  const lineupData = [];
  const lineupBackgrounds = [];
  for (let inning = 0; inning < innings; inning++) {
    const row = [inning + 1];
    const bgRow = [null];
    for (let j = 0; j < POSITIONS.length; j++) {
      const playerName = lineup[inning][j];
      row.push(playerName);
      if (preferences[playerName] && preferences[playerName][POSITIONS[j]] === 'Preferred') {
        bgRow.push('#b7e1cd');
      } else {
        bgRow.push(null);
      }
    }
    row.push(sitOuts[inning].join(', '));
    bgRow.push(null);
    lineupData.push(row);
    lineupBackgrounds.push(bgRow);
  }
  if (lineupData.length > 0) {
    const outputRange = suggesterSheet.getRange(lineupStartRow + 2, 1, lineupData.length, lineupData[0].length);
    outputRange.setValues(lineupData);
    outputRange.setBackgrounds(lineupBackgrounds);
    suggesterSheet.getRange(lineupStartRow + 2, 1, lineupData.length, 1).setHorizontalAlignment('center').setFontWeight('bold');

    // Add dropdowns for manual editing - apply to position columns at once
    const editRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(availablePlayers, true)
      .setAllowInvalid(true)
      .build();
    suggesterSheet.getRange(lineupStartRow + 2, 2, lineupData.length, POSITIONS.length).setDataValidation(editRule);
  }

  // Generate and write batting order
  const battingAverages = computeBattingAverages();
  const battingOrder = generateBattingOrder(availablePlayers, battingAverages);

  const battingStartRow = lineupStartRow + 2 + 9 + 2;
  // Clear previous batting order
  suggesterSheet.getRange(battingStartRow + 2, 1, MAX_PLAYERS, 5).clearContent();

  const battingData = [];
  for (const entry of battingOrder) {
    battingData.push([
      entry.position,
      entry.name,
      entry.obp > 0 ? entry.obp.toFixed(3) : '-',
      entry.slg > 0 ? entry.slg.toFixed(3) : '-',
      entry.sb > 0 ? entry.sb : '-'
    ]);
  }
  if (battingData.length > 0) {
    suggesterSheet.getRange(battingStartRow + 2, 1, battingData.length, 5).setValues(battingData);
    suggesterSheet.getRange(battingStartRow + 2, 1, battingData.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
    suggesterSheet.getRange(battingStartRow + 2, 3, battingData.length, 3).setHorizontalAlignment('center');
    suggesterSheet.getRange(battingStartRow + 2, 2, battingData.length, 1).setFontSize(11);
  }

  suggesterSheet.activate();
  ui.alert('Lineup Generated',
    'A suggested lineup has been generated for ' + innings + ' innings.\n\n' +
    'Field positions and batting order are shown below.\n' +
    'You can manually edit any cell using the dropdowns.\n' +
    'Copy this to the Game Entry sheet when ready.',
    ui.ButtonSet.OK);
}

function assignPositions(players, preferences, gamesSinceAtPosition, previousInnings, currentInning, depthChart) {
  // Score each player-position combination
  const numPlayers = players.length;
  const numPositions = POSITIONS.length;

  // Build cost matrix (lower = better assignment)
  const scores = [];
  for (let p = 0; p < numPlayers; p++) {
    scores[p] = [];
    for (let j = 0; j < numPositions; j++) {
      const playerName = players[p];
      const pos = POSITIONS[j];
      const pref = (preferences[playerName] && preferences[playerName][pos]) || 'Okay';

      let score = 0;

      // Hard constraint: restricted = very high cost
      if (pref === 'Restricted') {
        score = 10000;
      }

      // Bullpen warmup rule: prefer pitchers who sat out the previous inning
      // (continuing to pitch is fine, inning 0 is fine)
      // Soft penalty — strongly discouraged but allowed if no warmed-up pitcher is available
      if (score < 10000 && j === 0 && currentInning > 0) {
        const wasPitchingLast = (previousInnings[currentInning - 1][0] === playerName);
        const wasSittingOutLast = (previousInnings[currentInning - 1].indexOf(playerName) === -1);
        if (!wasPitchingLast && !wasSittingOutLast) {
          score += 200;
        }
      }

      // No-return rule for P and C: if a player previously played this position
      // but is no longer at it (i.e., they left), block them from returning
      if (score < 10000 && (j === 0 || j === 1) && currentInning > 0) {
        let everPlayedHere = false;
        let atPosInLastInning = false;
        for (let k = 0; k < currentInning; k++) {
          if (previousInnings[k].indexOf(playerName) === j) {
            everPlayedHere = true;
          }
        }
        if (everPlayedHere) {
          atPosInLastInning = (previousInnings[currentInning - 1].indexOf(playerName) === j);
          if (!atPosInLastInning) {
            score = 10000;
          }
        }
      }

      // Skip all bonuses if already blocked (Restricted or no-return rule)
      if (score < 10000) {
        if (pref === 'Preferred') {
          score -= 20; // bonus for preferred
        }
        // Okay = neutral (0)

        // Depth chart: ranked players get a bonus
        if (depthChart && depthChart[pos]) {
          const rank = depthChart[pos].indexOf(playerName);
          if (rank >= 0) {
            score -= (MAX_PLAYERS - rank) * 3; // 1st = -36, 2nd = -33, ..., 12th = -3
          }
        }

        // Recency: prioritize positions not played recently
        const gamesSince = (gamesSinceAtPosition[playerName] && gamesSinceAtPosition[playerName][pos]) || 0;
        score -= gamesSince * 5; // more games since = lower score = more priority

        // Position continuity: bonus for staying at same position
        if (currentInning > 0) {
          const prevAssignment = previousInnings[currentInning - 1];
          const prevPosIndex = prevAssignment.indexOf(playerName);
          if (prevPosIndex === j) {
            // Count consecutive innings at this position
            let consecutiveCount = 1;
            for (let k = currentInning - 2; k >= 0; k--) {
              if (previousInnings[k].indexOf(playerName) === j) {
                consecutiveCount++;
              } else {
                break;
              }
            }

            // P and C get a much stronger continuity bonus since leaving is permanent
            if (j === 0 || j === 1) {
              score -= 50; // P/C: strong incentive to keep pitcher/catcher in place
            } else if (consecutiveCount >= 2) {
              score -= 15; // 3rd+ consecutive inning: stronger bonus
            } else {
              score -= 10; // 2nd consecutive inning: standard bonus
            }
          }
        }
      }

      scores[p][j] = score;
    }
  }

  // Greedy assignment with backtracking avoidance
  // Sort positions by most constrained first (fewest valid players)
  const posOrder = POSITIONS.map((_, idx) => idx).sort((a, b) => {
    const validA = players.filter((_, pi) => scores[pi][a] < 10000).length;
    const validB = players.filter((_, pi) => scores[pi][b] < 10000).length;
    return validA - validB;
  });

  const assignment = new Array(numPositions).fill('');
  const assigned = new Set();

  for (const posIdx of posOrder) {
    let bestPlayer = -1;
    let bestScore = Infinity;

    for (let p = 0; p < numPlayers; p++) {
      if (assigned.has(p)) continue;
      if (scores[p][posIdx] >= 10000) continue; // never assign blocked players
      if (scores[p][posIdx] < bestScore) {
        bestScore = scores[p][posIdx];
        bestPlayer = p;
      }
    }

    if (bestPlayer >= 0) {
      assignment[posIdx] = players[bestPlayer];
      assigned.add(bestPlayer);
    }
  }

  return assignment;
}

// ============================================================
// HOW TO USE SHEET
// ============================================================

function createHowToUseSheet(ss) {
  let sheet = ss.getSheetByName('How To Use');
  if (sheet) sheet.clear(); else sheet = ss.insertSheet('How To Use');

  const instructions = [
    ['⚾ SOFTBALL LINEUP MANAGER - Instructions', ''],
    ['', ''],
    ['FINDING THE MENU', ''],
    ['Look for ⚾ Softball in the menu bar', 'It appears at the far right end, after Extensions and Help'],
    ['First time setup:', 'Go to Extensions > Apps Script, select onOpen, click Run (▶), and authorize when prompted'],
    ['If you don\'t see it after that', 'Close and reopen the spreadsheet — the menu loads automatically on each open'],
    ['', ''],
    ['GETTING STARTED', ''],
    ['1. Go to the Roster sheet', 'Enter your players\' names (up to 12 players)'],
    ['2. Set position preferences', 'For each player, set each position as Preferred (green), Okay (yellow), or Restricted (red)'],
    ['3. That\'s it!', 'You\'re ready to manage games'],
    ['', ''],
    ['ENTERING A GAME', ''],
    ['1. Go to the Game Entry sheet', 'Fill in the date, opponent, and number of innings'],
    ['2. Fill in the lineup grid', 'Use dropdowns to assign one player per position per inning'],
    ['3. Mark who sat out', 'Use the Sat Out column (players not assigned auto-count as sitting out)'],
    ['4. Enter batting stats (below lineup)', 'Fill in AB, hits (1B/2B/3B/HR), BB, SB, and CS for each player'],
    ['5. Save the game', 'Click ⚾ Softball (far right of menu bar) > Save Game (saves both lineup and batting stats)'],
    ['', ''],
    ['DEPTH CHART', ''],
    ['1. Go to the Depth Chart sheet', 'Rank players per position (1st = top choice, leave blank for unranked)'],
    ['2. Fill in rankings', 'Use the dropdowns to select which player is your 1st, 2nd, 3rd choice, etc. at each position'],
    ['3. How it works', 'Ranked players get a scoring bonus when the Lineup Suggester assigns positions'],
    ['4. Interaction with preferences', 'Restricted still blocks a player even if ranked 1st. Depth chart fine-tunes choices among Preferred/Okay players'],
    ['', ''],
    ['USING THE LINEUP SUGGESTER', ''],
    ['1. Go to the Lineup Suggester sheet', 'Check the boxes next to available players'],
    ['2. Set the number of innings', ''],
    ['3. Click ⚾ Softball (far right of menu bar) > Suggest Lineup', 'The algorithm will generate field positions AND a batting order'],
    ['4. Review and edit', 'Use dropdowns to make manual adjustments to field positions'],
    ['5. Batting order section', 'Shows suggested batting order based on OBP, slugging, and speed stats'],
    ['6. Copy to Game Entry', 'Transfer the final lineup to the Game Entry sheet for the actual game'],
    ['', ''],
    ['UNDERSTANDING THE BATTING ORDER', ''],
    ['Spots 1-3 (top of order):', 'Best OBP + speed — players who get on base and steal'],
    ['Spots 4-6 (middle):', 'Best slugging — power hitters who drive in runs'],
    ['Spots 7+ (bottom):', 'Remaining players by overall composite score'],
    ['Stability:', 'Players move at most 2 spots from their recent average position'],
    ['New players (< 3 games):', 'Default to roster order until enough data is collected'],
    ['', ''],
    ['VIEWING THE DASHBOARD', ''],
    ['1. Go to the Dashboard sheet', 'Click ⚾ Softball (far right of menu bar) > Refresh Dashboard to update stats'],
    ['2. Section 1: Innings at Each Position', 'Shows total innings each player has played at each position all season'],
    ['3. Section 2: Games Since Last Played', 'Yellow = 3+ games since, Red = 5+ games since playing that position'],
    ['4. Section 3: Batting Stats', 'Shows OBP, SLG, stolen bases, and caught stealing for each player'],
    ['', ''],
    ['TIPS', ''],
    ['• The Suggest Lineup algorithm:', 'Respects Restricted positions, keeps players at the same position for 2+ innings, and rotates sit-outs fairly'],
    ['• No-return rule for P and C:', 'Once a player leaves the Pitcher or Catcher position during a game, the algorithm will not assign them back to that position in a later inning'],
    ['• Bullpen warmup:', 'A new pitcher must have sat out the previous inning to warm up — the algorithm will only start a player at P if they were already pitching or were on the bench the inning before'],
    ['• Position continuity:', 'Players get a bonus for staying at the same position across innings (builds comfort)'],
    ['• Dashboard colors:', 'Help you spot players who need more time at certain positions'],
    ['• Mobile-friendly:', 'All dropdowns are large-format for easy phone/tablet use'],
    ['• Season History sheet:', 'Stores all game data - don\'t edit directly unless fixing errors'],
    ['• Batting Stats sheet:', 'Stores per-game batting data - don\'t edit directly unless fixing errors'],
  ];

  sheet.getRange(1, 1, instructions.length, 2).setValues(instructions);

  // Formatting
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');

  // Bold column A for all rows (step numbers and bullets will be bold)
  sheet.getRange(1, 1, instructions.length, 1).setFontWeight('bold');
  // Section header rows
  const sectionRows = [3, 8, 13, 20, 26, 34, 41, 47];
  sectionRows.forEach(row => {
    if (row <= instructions.length) {
      sheet.getRange(row, 1, 1, 2).setFontSize(13).setBackground('#e8f0fe');
    }
  });
}

// ============================================================
// HELPERS
// ============================================================

function getRosterNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roster = ss.getSheetByName('Roster');
  if (!roster) return [];

  const names = [];
  const data = roster.getRange(2, 2, MAX_PLAYERS, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0];
    if (name && name.toString().trim() !== '') {
      names.push(name.toString().trim());
    }
  }
  return names;
}

// Auto-refresh dropdowns when roster changes
function onEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() === 'Roster' && e.range.getColumn() === 2) {
    // Player name was edited - update dropdowns
    updateGameEntryDropdowns();
    updateSuggesterNames();
    updateDepthChartDropdowns();
  }
}

function updateSuggesterNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Lineup Suggester');
  if (!sheet) return;

  const players = getRosterNames();
  const checkValues = [];
  const nameValues = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    if (i < players.length) {
      checkValues.push([true]);
      nameValues.push([players[i]]);
    } else {
      checkValues.push([false]);
      nameValues.push(['']);
    }
  }
  sheet.getRange(7, 1, MAX_PLAYERS, 1).setValues(checkValues);
  sheet.getRange(7, 2, MAX_PLAYERS, 1).setValues(nameValues);
}
