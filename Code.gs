// ============================================================
// SOFTBALL LINEUP MANAGER - Google Apps Script
// ============================================================

const POSITIONS = ['P', 'C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF'];
const MAX_PLAYERS = 12;
const PREF_OPTIONS = ['Preferred', 'Okay', 'Restricted'];
const COACHING_MODES = ['Balanced', 'Competitive', 'Development'];
const BATTING_STATS = 'Batting Stats';
const MAX_GAMES = 3;

// Game Entry layout constants (single source of truth for row offsets)
const GE_ATTEND_ROW = 5;          // Attendance header row
const GE_ATTEND_DATA = 6;         // First attendance checkbox row (GE_ATTEND_ROW + 1)
const GE_GRID_HEADER = 19;        // Lineup grid header row (GE_ATTEND_ROW + MAX_PLAYERS + 2)
const GE_GRID_DATA = 20;          // First lineup data row (GE_GRID_HEADER + 1)
const GE_STATS_START = 30;        // Batting stats section header row (GE_GRID_DATA + 9 + 1)
const GE_MAX_SIT_OUT = MAX_PLAYERS - POSITIONS.length; // typically 3

// ============================================================
// MENU & INITIALIZATION
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚾ Softball')
    .addItem('Initialize All Sheets', 'initializeAll')
    .addItem('Rebuild Game Entry', 'rebuildGameEntry')
    .addSeparator()
    .addItem('Save Game', 'saveGame')
    .addItem('Delete Last Game', 'deleteLastGame')
    .addItem('Suggest Lineup', 'suggestLineup')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addToUi();
}

function rebuildGameEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createGameEntrySheet(ss);
  ss.getSheetByName('Game Entry').activate();
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

  // Preserve existing roster data (names and preferences) before clearing
  let existingData = null;
  if (sheet) {
    const data = sheet.getRange(2, 2, MAX_PLAYERS, POSITIONS.length + 1).getValues();
    const hasData = data.some(row => row[0] && row[0].toString().trim() !== '');
    if (hasData) existingData = data;
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Roster');
  }

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

  // Restore existing roster data if it was preserved
  if (existingData) {
    sheet.getRange(2, 2, MAX_PLAYERS, POSITIONS.length + 1).setValues(existingData);
  }
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

  // Attendance checkboxes (left sidebar)
  sheet.getRange(GE_ATTEND_ROW, 1).setValue('Attendance')
    .setFontWeight('bold').setFontSize(12);
  const players = getRosterNames();
  const attendCheckVals = [];
  const attendNameVals = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    attendCheckVals.push([i < players.length && players[i] ? true : false]);
    attendNameVals.push([i < players.length && players[i] ? players[i] : '']);
  }
  sheet.getRange(GE_ATTEND_DATA, 1, MAX_PLAYERS, 1).insertCheckboxes();
  sheet.getRange(GE_ATTEND_DATA, 1, MAX_PLAYERS, 1).setValues(attendCheckVals);
  sheet.getRange(GE_ATTEND_DATA, 2, MAX_PLAYERS, 1).setValues(attendNameVals).setFontSize(11);

  // Lineup grid header
  const gridHeaders = ['Inning'];
  POSITIONS.forEach(p => gridHeaders.push(p));
  for (let s = 1; s <= GE_MAX_SIT_OUT; s++) {
    gridHeaders.push('Sit Out ' + s);
  }
  sheet.getRange(GE_GRID_HEADER, 1, 1, gridHeaders.length).setValues([gridHeaders])
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // Inning rows (max 9) - batch write
  const inningNums = [[1],[2],[3],[4],[5],[6],[7],[8],[9]];
  sheet.getRange(GE_GRID_DATA, 1, 9, 1).setValues(inningNums).setHorizontalAlignment('center').setFontWeight('bold');

  // Batting stats section
  sheet.getRange(GE_STATS_START, 1).setValue('Batting Stats')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');

  const statsHeaders = ['Player', 'AB', '1B', '2B', '3B', 'HR', 'BB', 'SB', 'CS'];
  sheet.getRange(GE_STATS_START + 1, 1, 1, statsHeaders.length).setValues([statsHeaders])
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
  sheet.getRange(GE_STATS_START + 2, 1, MAX_PLAYERS, 1).setValues(statsNames).setFontSize(11);
  sheet.getRange(GE_STATS_START + 2, 2, MAX_PLAYERS, 8).setValues(statsDefaults).setHorizontalAlignment('center');

  // Numeric validation for stat columns
  const numRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 99).setAllowInvalid(false).build();
  sheet.getRange(GE_STATS_START + 2, 2, MAX_PLAYERS, 8).setDataValidation(numRule);

  sheet.setFrozenRows(3);
  updateGameEntryDropdowns();
}

function updateGameEntryDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Game Entry');
  const players = getRosterNames();

  if (players.length === 0) return;

  // Set dropdowns for position cells
  const posRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(players, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(GE_GRID_DATA, 2, 9, POSITIONS.length).setDataValidation(posRule);

  // Sat out columns
  const satOutRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(players, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(GE_GRID_DATA, POSITIONS.length + 2, 9, GE_MAX_SIT_OUT).setDataValidation(satOutRule);

  // Update attendance names
  const attendNameVals = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    attendNameVals.push([i < players.length ? players[i] : '']);
  }
  sheet.getRange(GE_ATTEND_DATA, 2, MAX_PLAYERS, 1).setValues(attendNameVals);

  // Update batting stats names
  const nameValues = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    nameValues.push([i < players.length ? players[i] : '']);
  }
  sheet.getRange(GE_STATS_START + 2, 1, MAX_PLAYERS, 1).setValues(nameValues);
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
  const gameData = gameSheet.getRange(GE_GRID_DATA, 2, 9, POSITIONS.length + GE_MAX_SIT_OUT).getValues();

  // Read attendance checkboxes to identify absent players
  const attendData = gameSheet.getRange(GE_ATTEND_DATA, 1, MAX_PLAYERS, 1).getValues();
  const absentPlayers = new Set();
  for (let i = 0; i < players.length; i++) {
    if (players[i] && attendData[i][0] === false) {
      absentPlayers.add(players[i]);
    }
  }

  // Validate: check for duplicate players in same inning and absent players in lineup
  const errors = [];
  for (let inning = 1; inning <= innings; inning++) {
    const inningData = gameData[inning - 1];
    const seen = new Set();
    for (let j = 0; j < POSITIONS.length; j++) {
      const name = inningData[j];
      if (!name || name.toString().trim() === '') continue;
      if (seen.has(name)) {
        errors.push('Inning ' + inning + ': ' + name + ' appears in multiple positions');
      }
      seen.add(name);
      if (absentPlayers.has(name)) {
        errors.push('Inning ' + inning + ': ' + name + ' is marked absent but assigned to ' + POSITIONS[j]);
      }
    }
  }
  if (errors.length > 0) {
    ui.alert('Lineup Errors', errors.join('\n'), ui.ButtonSet.OK);
    return;
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
      for (let s = 0; s < GE_MAX_SIT_OUT; s++) {
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
  gameSheet.getRange(GE_GRID_DATA, 2, 9, POSITIONS.length + GE_MAX_SIT_OUT).clearContent();

  // Reset attendance checkboxes to all checked
  const resetChecks = [];
  for (let i = 0; i < MAX_PLAYERS; i++) resetChecks.push([true]);
  gameSheet.getRange(GE_ATTEND_DATA, 1, MAX_PLAYERS, 1).setValues(resetChecks);

  // Clear batting stats section
  gameSheet.getRange(GE_STATS_START + 2, 2, MAX_PLAYERS, 8).setValue(0);

  // Refresh dashboard
  refreshDashboard();

  ui.alert('Game Saved!', 'Game #' + gameNum + ' vs ' + opponent + ' has been saved.\nDashboard has been refreshed.', ui.ButtonSet.OK);
}

function saveBattingStats(ss, gameSheet, gameNum, date, players) {
  const battingSheet = ss.getSheetByName(BATTING_STATS);
  if (!battingSheet) return;

  // Batch read: player names (col 1) + stats (cols 2-9)
  const statsData = gameSheet.getRange(GE_STATS_START + 2, 1, MAX_PLAYERS, 9).getValues();

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
// DELETE LAST GAME
// ============================================================

function deleteLastGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const historySheet = ss.getSheetByName('Season History');
  const battingSheet = ss.getSheetByName(BATTING_STATS);

  if (!historySheet) {
    ui.alert('Error', 'Season History sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const historyData = historySheet.getDataRange().getValues();
  if (historyData.length <= 1) {
    ui.alert('No Games', 'There are no saved games to delete.', ui.ButtonSet.OK);
    return;
  }

  // Find the max game number
  let maxGame = 0;
  for (let i = 1; i < historyData.length; i++) {
    if (historyData[i][0] > maxGame) maxGame = historyData[i][0];
  }

  // Get opponent for confirmation
  let opponent = '';
  for (let i = 1; i < historyData.length; i++) {
    if (historyData[i][0] === maxGame) { opponent = historyData[i][2]; break; }
  }

  const response = ui.alert('Delete Last Game',
    'Delete Game #' + maxGame + ' vs ' + opponent + '?\n\nThis cannot be undone.',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // Delete from Season History (bottom-up to preserve row indices)
  for (let i = historyData.length - 1; i >= 1; i--) {
    if (historyData[i][0] === maxGame) {
      historySheet.deleteRow(i + 1);
    }
  }

  // Delete from Batting Stats
  if (battingSheet) {
    const battingData = battingSheet.getDataRange().getValues();
    for (let i = battingData.length - 1; i >= 1; i--) {
      if (battingData[i][0] === maxGame) {
        battingSheet.deleteRow(i + 1);
      }
    }
  }

  refreshDashboard();
  ui.alert('Deleted', 'Game #' + maxGame + ' vs ' + opponent + ' has been deleted.\nDashboard has been refreshed.', ui.ButtonSet.OK);
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
        gameNums: new Set(), ab: 0, singles: 0, doubles: 0, triples: 0, hr: 0,
        bb: 0, sb: 0, cs: 0,
        recentPositions: [] // [{gameNum, pos}] for stability calc
      };
    }

    const s = playerStats[player];
    const gameNum = data[i][0];

    // Skip duplicate rows (same player + game number)
    if (s.gameNums.has(gameNum)) continue;
    s.gameNums.add(gameNum);

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
      games: s.gameNums.size,
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

function getCoachingModeConfig(mode) {
  const selected = COACHING_MODES.indexOf(mode) >= 0 ? mode : 'Balanced';
  const configs = {
    Balanced: {
      name: 'Balanced',
      description: 'Mixes best-fit positions with fair reps and steady development.',
      batting: {
        topObp: 100,
        topBaserunning: 5,
        midSlg: 100,
        midObp: 30,
        overallObp: 50,
        overallSlg: 50,
        overallBaserunning: 3,
        stabilityMaxMove: 2,
        shuffleJitter: 0.16,
        topScoreJitter: 6,
        midScoreJitter: 6,
        overallScoreJitter: 4
      },
      field: {
        preferredBonus: -20,
        depthWeight: 3,
        recencyWeight: 5,
        pitcherCatcherContinuityBonus: -50,
        fieldSecondInningBonus: -10,
        fieldThirdInningPenalty: 5,
        fieldFourthPlusPenalty: 30,
        outfieldOnlyBonus: 15,
        outfieldOnlyBonusMulti: 35,
        diversityNeverBonus: -5,
        diversityRareBonus: -2,
        missedGamePenalty: 1.5
      }
    },
    Competitive: {
      name: 'Competitive',
      description: 'Leans harder into best current fit, stronger depth-chart choices, and lineup stability.',
      batting: {
        topObp: 115,
        topBaserunning: 4,
        midSlg: 130,
        midObp: 25,
        overallObp: 45,
        overallSlg: 60,
        overallBaserunning: 2,
        stabilityMaxMove: 1,
        shuffleJitter: 0.07,
        topScoreJitter: 2,
        midScoreJitter: 2,
        overallScoreJitter: 1
      },
      field: {
        preferredBonus: -24,
        depthWeight: 4.5,
        recencyWeight: 2.5,
        pitcherCatcherContinuityBonus: -65,
        fieldSecondInningBonus: -14,
        fieldThirdInningPenalty: 8,
        fieldFourthPlusPenalty: 40,
        outfieldOnlyBonus: 8,
        outfieldOnlyBonusMulti: 18,
        diversityNeverBonus: -1,
        diversityRareBonus: 0,
        missedGamePenalty: 0.5
      }
    },
    Development: {
      name: 'Development',
      description: 'Pushes harder for overdue reps, broader position variety, and long-term growth.',
      batting: {
        topObp: 90,
        topBaserunning: 6,
        midSlg: 75,
        midObp: 40,
        overallObp: 47,
        overallSlg: 38,
        overallBaserunning: 5,
        stabilityMaxMove: 3,
        shuffleJitter: 0.24,
        topScoreJitter: 10,
        midScoreJitter: 10,
        overallScoreJitter: 7
      },
      field: {
        preferredBonus: -16,
        depthWeight: 2,
        recencyWeight: 8,
        pitcherCatcherContinuityBonus: -40,
        fieldSecondInningBonus: -6,
        fieldThirdInningPenalty: 2,
        fieldFourthPlusPenalty: 18,
        outfieldOnlyBonus: 24,
        outfieldOnlyBonusMulti: 48,
        diversityNeverBonus: -12,
        diversityRareBonus: -6,
        missedGamePenalty: 0.75
      }
    }
  };
  return configs[selected];
}

function getDeterministicScoreJitter(playerName, contextKey, magnitude) {
  if (!magnitude) return 0;
  const input = String(playerName || '') + '|' + String(contextKey || '');
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = ((hash * 31) + input.charCodeAt(i)) & 0x7fffffff;
  }
  const normalized = (hash % 10000) / 9999; // 0..1
  return (normalized * 2 - 1) * magnitude;
}

function generateBattingOrder(availablePlayers, battingAverages, modeConfig) {
  modeConfig = modeConfig || getCoachingModeConfig('Balanced');
  const battingConfig = modeConfig.batting;
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
  const topScore = (name, s) =>
    (s.obp * battingConfig.topObp) +
    (s.baserunning * battingConfig.topBaserunning) +
    getDeterministicScoreJitter(name, modeConfig.name + ':top', battingConfig.topScoreJitter);
  const midScore = (name, s) =>
    (s.slg * battingConfig.midSlg) +
    (s.obp * battingConfig.midObp) +
    getDeterministicScoreJitter(name, modeConfig.name + ':mid', battingConfig.midScoreJitter);
  const overallScore = (name, s) =>
    (s.obp * battingConfig.overallObp) +
    (s.slg * battingConfig.overallSlg) +
    (s.baserunning * battingConfig.overallBaserunning) +
    getDeterministicScoreJitter(name, modeConfig.name + ':overall', battingConfig.overallScoreJitter);

  // Sort for top of order
  const topCandidates = withData.slice().sort((a, b) => topScore(b.name, b.stats) - topScore(a.name, a.stats));
  // Sort for middle
  const midCandidates = withData.slice().sort((a, b) => midScore(b.name, b.stats) - midScore(a.name, a.stats));

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
    .sort((a, b) => overallScore(b.name, b.stats) - overallScore(a.name, a.stats));
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
      if (diff > battingConfig.stabilityMaxMove) {
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

// Shuffle batting order for variety across multi-game days.
// Players stay roughly in their tier but can drift ~1 spot across tier boundaries.
// Uses gameIndex as a seed for deterministic pseudo-random jitter.
function shuffleBattingOrderByTier(order, gameIndex, modeConfig) {
  modeConfig = modeConfig || getCoachingModeConfig('Balanced');
  // Seeded pseudo-random: simple LCG so same roster+gameIndex = same shuffle
  let seed = gameIndex * 2654435761; // large odd multiplier
  function nextRand() {
    seed = (seed * 1664525 + 1013904223) & 0x7fffffff;
    return seed / 0x7fffffff; // 0-1
  }

  // Add jitter to each player's position score: ±10% of lineup length
  const jitterRange = Math.max(1, Math.round(order.length * modeConfig.batting.shuffleJitter));
  const jittered = order.map((entry, idx) => {
    const jitter = (nextRand() * 2 - 1) * jitterRange; // random between -jitterRange and +jitterRange
    return { entry, sortKey: idx + jitter };
  });

  jittered.sort((a, b) => a.sortKey - b.sortKey);

  return jittered.map((item, idx) => ({
    name: item.entry.name,
    position: idx + 1,
    obp: item.entry.obp,
    slg: item.entry.slg,
    sb: item.entry.sb
  }));
}

function getDepthChartRank(depthChart, pos, playerName) {
  if (!depthChart || !depthChart[pos]) return 0;
  const idx = depthChart[pos].indexOf(playerName);
  return idx >= 0 ? idx + 1 : 0;
}

function buildBattingExplanation(entry, slot, battingAverages, modeConfig) {
  const stats = battingAverages[entry.name] || {};
  const bits = [];
  if ((slot <= 3) && stats.obp > 0) {
    bits.push('OBP ' + stats.obp.toFixed(3));
    if (stats.sb > 0) bits.push('speed profile ' + stats.sb + ' SB');
    bits.push('fits top-of-order in ' + modeConfig.name.toLowerCase() + ' mode');
  } else if ((slot >= 4 && slot <= 6) && stats.slg > 0) {
    bits.push('SLG ' + stats.slg.toFixed(3));
    if (stats.obp > 0) bits.push('OBP ' + stats.obp.toFixed(3));
    bits.push('profiles as a run producer');
  } else {
    if (stats.games >= 3) {
      bits.push('overall profile from ' + stats.games + ' tracked games');
    } else {
      bits.push('limited batting sample, held near roster/default slot');
    }
    if (stats.avgBattingPos > 0) {
      bits.push('recent avg slot ' + stats.avgBattingPos.toFixed(1));
    }
  }
  return entry.position + '. ' + entry.name + ' - ' + bits.join(', ');
}

function buildPositionExplanation(playerName, pos, preferences, depthChart, gamesSinceAtPosition, seasonPositionCounts) {
  const bits = [];
  const pref = preferences[playerName] ? preferences[playerName][pos] : '';
  if (pref === 'Preferred') bits.push('preferred there');
  else if (pref === 'Okay') bits.push('playable fit there');

  const rank = getDepthChartRank(depthChart, pos, playerName);
  if (rank > 0) bits.push('#' + rank + ' on depth chart');

  const since = gamesSinceAtPosition[playerName] ? gamesSinceAtPosition[playerName][pos] : 0;
  if (since >= 3) bits.push(since + ' games since last ' + pos);

  if (seasonPositionCounts && seasonPositionCounts[playerName]) {
    const posCount = seasonPositionCounts[playerName][pos] || 0;
    if (posCount === 0) bits.push('new season rep at ' + pos);
    else if (posCount <= 2) bits.push('limited season reps at ' + pos);
  }

  return playerName + ' at ' + pos + ': ' + (bits.length > 0 ? bits.join(', ') : 'best available fit');
}

function buildGameDecisionTrace(lineup, sitOuts, battingOrder, modeConfig, preferences,
    depthChart, gamesSinceAtPosition, seasonPositionCounts, sitOutCap, reliefPitcher, battingAverages) {
  const notes = [];
  notes.push('Mode: ' + modeConfig.name + ' - ' + modeConfig.description);

  const firstInning = lineup[0] || [];
  const anchorPositions = ['P', 'C', 'SS', 'CF'];
  const anchorNotes = [];
  anchorPositions.forEach(pos => {
    const idx = POSITIONS.indexOf(pos);
    if (idx < 0 || !firstInning[idx]) return;
    anchorNotes.push(buildPositionExplanation(
      firstInning[idx], pos, preferences, depthChart, gamesSinceAtPosition, seasonPositionCounts));
  });
  if (anchorNotes.length > 0) {
    notes.push('Opening defense: ' + anchorNotes.join(' | '));
  }

  const topOrderNotes = battingOrder.slice(0, Math.min(3, battingOrder.length))
    .map(entry => buildBattingExplanation(entry, entry.position, battingAverages, modeConfig));
  if (topOrderNotes.length > 0) {
    notes.push('Top of order: ' + topOrderNotes.join(' | '));
  }

  const middleOrderNotes = battingOrder.slice(3, Math.min(6, battingOrder.length))
    .map(entry => buildBattingExplanation(entry, entry.position, battingAverages, modeConfig));
  if (middleOrderNotes.length > 0) {
    notes.push('Middle order: ' + middleOrderNotes.join(' | '));
  }

  const sitOutCounts = {};
  sitOuts.forEach(group => {
    group.forEach(name => {
      sitOutCounts[name] = (sitOutCounts[name] || 0) + 1;
    });
  });
  const sitOutLeaders = Object.keys(sitOutCounts)
    .sort((a, b) => sitOutCounts[b] - sitOutCounts[a] || a.localeCompare(b))
    .slice(0, 3)
    .map(name => name + ' (' + sitOutCounts[name] + ')');
  if (sitOutCap > 0) {
    notes.push('Sit-out plan: cap ' + sitOutCap + ' per player' +
      (sitOutLeaders.length > 0 ? '; highest this game: ' + sitOutLeaders.join(', ') : ''));
  } else {
    notes.push('Sit-out plan: everyone stays on the field each inning.');
  }

  if (reliefPitcher) {
    notes.push('Relief pitcher: ' + reliefPitcher + ' is next in line from the depth chart if the starter needs to come out.');
  }

  return notes;
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

  // Clear all data rows to prevent stale data from removed players
  const sec2StartRow = 5 + MAX_PLAYERS + 2;
  const sec3StartRow = sec2StartRow + MAX_PLAYERS + 3;
  dashboard.getRange(6, 1, MAX_PLAYERS, POSITIONS.length + 3).clearContent().setBackground(null);
  dashboard.getRange(sec2StartRow + 2, 1, MAX_PLAYERS, POSITIONS.length + 1).clearContent().setBackground(null);
  dashboard.getRange(sec3StartRow + 2, 1, MAX_PLAYERS, 8).clearContent();

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

  // Preserve existing depth chart rankings before clearing
  let existingData = null;
  if (sheet) {
    const data = sheet.getRange(5, 2, MAX_PLAYERS, POSITIONS.length).getValues();
    const hasData = data.some(row => row.some(cell => cell && cell.toString().trim() !== ''));
    if (hasData) existingData = data;
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Depth Chart');
  }

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

  // Restore existing depth chart rankings if preserved
  if (existingData) {
    sheet.getRange(5, 2, MAX_PLAYERS, POSITIONS.length).setValues(existingData);
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

  sheet.getRange('C4').setValue('Games:').setFontWeight('bold');
  sheet.getRange('D4').setValue(1);
  const gamesList = [];
  for (let g = 1; g <= MAX_GAMES; g++) gamesList.push(String(g));
  const gamesVal = SpreadsheetApp.newDataValidation()
    .requireValueInList(gamesList, true).setAllowInvalid(false).build();
  sheet.getRange('D4').setDataValidation(gamesVal);

  sheet.getRange('H4').setValue('Mode:').setFontWeight('bold');
  sheet.getRange('I4').setValue('Balanced');
  const modeVal = SpreadsheetApp.newDataValidation()
    .requireValueInList(COACHING_MODES, true).setAllowInvalid(false).build();
  sheet.getRange('I4').setDataValidation(modeVal);
  sheet.getRange('J4').setValue('Balanced = mix fairness and fit; Competitive = best current lineup; Development = broaden reps')
    .setFontColor('#666666').setFontStyle('italic');

  // Player availability: A=checkbox, B=name, C=Rest P, D=Rest C, E=G1, F=G2, G=G3
  // G1 is always checked for available players. G2/G3 hidden when Games=1.
  // Rest P / Rest C stay in the familiar C/D position from the old layout.
  sheet.getRange('A6').setValue('Available Players:')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange('C6').setValue('Rest P').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange('D6').setValue('Rest C').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange('E6').setValue('G1').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange('F6').setValue('G2').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange('G6').setValue('G3').setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');

  const players = getRosterNames();
  sheet.getRange(7, 1, MAX_PLAYERS, 1).insertCheckboxes();
  sheet.getRange(7, 3, MAX_PLAYERS, 5).insertCheckboxes(); // Rest P, Rest C, G1, G2, G3
  const checkVals = [];
  const nameVals = [];
  const extraVals = []; // Rest P, Rest C, G1, G2, G3
  for (let i = 0; i < MAX_PLAYERS; i++) {
    const hasPlayer = i < players.length && players[i];
    checkVals.push([hasPlayer ? true : false]);
    nameVals.push([hasPlayer ? players[i] : '']);
    extraVals.push([false, false, hasPlayer ? true : false, false, false]);
  }
  sheet.getRange(7, 1, MAX_PLAYERS, 1).setValues(checkVals);
  sheet.getRange(7, 2, MAX_PLAYERS, 1).setValues(nameVals).setFontSize(11);
  sheet.getRange(7, 3, MAX_PLAYERS, 5).setValues(extraVals);

  // Hide G1/G2/G3 columns when Games=1 (default) — they're redundant with the master checkbox
  // G1 (col E), G2 (col F), G3 (col G) are all hidden initially; shown when Games > 1
  sheet.showColumns(3); sheet.showColumns(4); // ensure Rest P/C visible
  sheet.hideColumns(5); // G1
  sheet.hideColumns(6); // G2
  sheet.hideColumns(7); // G3

  // Lineup Card area (row 20+) — combined view written dynamically by suggestLineup()
  const lineupStartRow = 7 + MAX_PLAYERS + 1; // row 20
  sheet.getRange(lineupStartRow, 1).setValue('Lineup Card')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
  // Headers are dynamic (depend on innings) — written by suggestLineup()

  // Field Lineup area (below lineup card: card title + header + MAX_PLAYERS data + 2 summary + 1 gap)
  const fieldStartRow = lineupStartRow + MAX_PLAYERS + 5; // row 37
  sheet.getRange(fieldStartRow, 1).setValue('Suggested Field Lineup (for Game Entry copy-paste)')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');

  const gridHeaders = ['Inning'];
  POSITIONS.forEach(p => gridHeaders.push(p));
  for (let s = 1; s <= 3; s++) gridHeaders.push('Sit Out ' + s);
  sheet.getRange(fieldStartRow + 1, 1, 1, gridHeaders.length).setValues([gridHeaders])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');

  // Batting order area (below field lineup)
  const battingStartRow = fieldStartRow + 2 + 9 + 2; // after max 9 innings + gap
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

  // Ensure player names are current from roster
  updateSuggesterNames();

  const innings = suggesterSheet.getRange('B4').getValue();
  if (!innings || innings < 1) {
    ui.alert('Error', 'Please enter a valid number of innings.', ui.ButtonSet.OK);
    return;
  }

  const games = Number(suggesterSheet.getRange('D4').getValue()) || 1;
  const coachingMode = suggesterSheet.getRange('I4').getValue() || 'Balanced';
  const modeConfig = getCoachingModeConfig(coachingMode);

  // Read player data: col A=master checkbox, B=name, C=Rest P, D=Rest C, E=G1, F=G2, G=G3
  const checkData = suggesterSheet.getRange(7, 1, MAX_PLAYERS, 2).getValues();
  const extraData = suggesterSheet.getRange(7, 3, MAX_PLAYERS, 5).getValues(); // RestP,RestC,G1,G2,G3

  // Build per-game available players and rest flags
  const perGameAvailable = []; // perGameAvailable[g] = [playerNames]
  const restFlags = {}; // playerName -> { P: bool, C: bool }
  const allAvailablePlayers = []; // union across all games
  const allAvailableSet = new Set();

  for (let g = 0; g < games; g++) {
    perGameAvailable.push([]);
  }

  for (let i = 0; i < MAX_PLAYERS; i++) {
    if (!checkData[i][0] || !checkData[i][1]) continue;
    const name = checkData[i][1];
    restFlags[name] = { P: !!extraData[i][0], C: !!extraData[i][1] }; // cols 0=Rest P, 1=Rest C

    if (games === 1) {
      // Single game: master checkbox is the only attendance control
      perGameAvailable[0].push(name);
      allAvailableSet.add(name);
      allAvailablePlayers.push(name);
    } else {
      // Multi-game: use G1/G2/G3 columns (indices 2,3,4)
      for (let g = 0; g < games; g++) {
        if (extraData[i][2 + g]) { // index 2=G1, 3=G2, 4=G3
          perGameAvailable[g].push(name);
          if (!allAvailableSet.has(name)) {
            allAvailableSet.add(name);
            allAvailablePlayers.push(name);
          }
        }
      }
    }
  }

  // Validate minimum players per game
  for (let g = 0; g < games; g++) {
    if (perGameAvailable[g].length < POSITIONS.length) {
      ui.alert('Not Enough Players',
        'Game ' + (g + 1) + ': You need at least ' + POSITIONS.length + ' available players to fill all positions. You have ' + perGameAvailable[g].length + '.',
        ui.ButtonSet.OK);
      return;
    }
  }

  // Get position preferences from roster (using allAvailablePlayers for union)
  const rosterData = rosterSheet.getRange(2, 2, MAX_PLAYERS, POSITIONS.length + 1).getValues();
  const preferences = {};
  for (let i = 0; i < rosterData.length; i++) {
    const name = rosterData[i][0];
    if (name && allAvailableSet.has(name)) {
      preferences[name] = {};
      for (let j = 0; j < POSITIONS.length; j++) {
        preferences[name][POSITIONS[j]] = rosterData[i][j + 1] || 'Okay';
      }
      // Apply rest flags — override to Restricted
      if (restFlags[name]) {
        if (restFlags[name].P) preferences[name]['P'] = 'Restricted';
        if (restFlags[name].C) preferences[name]['C'] = 'Restricted';
      }
    }
  }

  // Warn about pitcher/catcher availability per game
  const warnings = [];
  for (let g = 0; g < games; g++) {
    const gLabel = games > 1 ? ' (Game ' + (g + 1) + ')' : '';
    const canPitch = perGameAvailable[g].filter(p => preferences[p] && preferences[p]['P'] !== 'Restricted');
    const canCatch = perGameAvailable[g].filter(p => preferences[p] && preferences[p]['C'] !== 'Restricted');
    if (canPitch.length === 0) {
      warnings.push('No players available to pitch' + gLabel + '!');
    } else if (canPitch.length < 2) {
      warnings.push('Only 1 player can pitch' + gLabel + ' (' + canPitch[0] + ').');
    }
    if (canCatch.length === 0) {
      warnings.push('No players available to catch' + gLabel + '!');
    } else if (canCatch.length < 2) {
      warnings.push('Only 1 player can catch' + gLabel + ' (' + canCatch[0] + ').');
    }
  }
  if (warnings.length > 0) {
    const proceed = ui.alert('Lineup Warning', warnings.join('\n') + '\n\nContinue anyway?', ui.ButtonSet.YES_NO);
    if (proceed !== ui.Button.YES) return;
  }

  // Get history stats for recency scoring (shared across all games)
  const gamesSinceAtPosition = {};
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();

    const playerGameList = {};
    for (let i = 1; i < historyData.length; i++) {
      const gameNum = historyData[i][0];
      const playerName = historyData[i][5];
      if (!allAvailableSet.has(playerName)) continue;
      if (!playerGameList[playerName]) playerGameList[playerName] = new Set();
      playerGameList[playerName].add(gameNum);
    }

    const playerGameIndex = {};
    for (const name of allAvailablePlayers) {
      if (playerGameList[name]) {
        const sorted = Array.from(playerGameList[name]).sort((a, b) => a - b);
        playerGameIndex[name] = {};
        sorted.forEach((g, idx) => playerGameIndex[name][g] = idx + 1);
      }
    }

    const lastOrdinalAtPos = {};
    for (let i = 1; i < historyData.length; i++) {
      const gameNum = historyData[i][0];
      const playerName = historyData[i][5];
      if (!allAvailableSet.has(playerName)) continue;
      if (!lastOrdinalAtPos[playerName]) lastOrdinalAtPos[playerName] = {};
      const ordinal = playerGameIndex[playerName] ? playerGameIndex[playerName][gameNum] : 0;

      for (let j = 0; j < POSITIONS.length; j++) {
        if (historyData[i][6 + j] === 1) {
          if (!lastOrdinalAtPos[playerName][POSITIONS[j]] || ordinal > lastOrdinalAtPos[playerName][POSITIONS[j]]) {
            lastOrdinalAtPos[playerName][POSITIONS[j]] = ordinal;
          }
        }
      }
    }

    for (const name of allAvailablePlayers) {
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

  // Get depth chart rankings (shared)
  const depthChart = getDepthChart();

  // Historical stats: sit-outs, season position counts, games attended, weekly IP
  const totalSatOut = {};
  const gamesAttended = {};
  const seasonPositionCounts = {};
  const weeklyIP = {};
  let teamTotalGames = 0;
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    const now = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const playerGameSets = {};
    for (let i = 1; i < historyData.length; i++) {
      const playerName = historyData[i][5];
      const gameNum = historyData[i][0];
      if (gameNum > teamTotalGames) teamTotalGames = gameNum;

      const satOut = historyData[i][6 + POSITIONS.length];
      if (!totalSatOut[playerName]) totalSatOut[playerName] = 0;
      if (satOut === 1) totalSatOut[playerName]++;

      // Track games attended per player
      if (!playerGameSets[playerName]) playerGameSets[playerName] = new Set();
      playerGameSets[playerName].add(gameNum);

      // Season position counts for diversity scoring
      if (!seasonPositionCounts[playerName]) {
        seasonPositionCounts[playerName] = {};
        POSITIONS.forEach(p => seasonPositionCounts[playerName][p] = 0);
      }
      for (let j = 0; j < POSITIONS.length; j++) {
        if (historyData[i][6 + j] === 1) {
          seasonPositionCounts[playerName][POSITIONS[j]]++;
        }
      }

      // Weekly innings pitched (last 7 days, display only)
      const gameDate = historyData[i][1];
      if (gameDate instanceof Date && gameDate >= weekAgo && historyData[i][6] === 1) {
        weeklyIP[playerName] = (weeklyIP[playerName] || 0) + 1;
      }
    }
    for (const name in playerGameSets) {
      gamesAttended[name] = playerGameSets[name].size;
    }
  }

  // Missed games per player (for small position equity penalty)
  const playerMissedGames = {};
  allAvailablePlayers.forEach(p => {
    playerMissedGames[p] = teamTotalGames - (gamesAttended[p] || 0);
  });

  // Compute batting averages once
  const battingAverages = computeBattingAverages();

  // ── Game loop: generate lineups for each game ──
  const crossGameSatOutCount = {};
  const crossGameGamesPlayed = {};
  allAvailablePlayers.forEach(p => {
    crossGameSatOutCount[p] = totalSatOut[p] || 0;
    crossGameGamesPlayed[p] = gamesAttended[p] || 0;
  });

  const allGamesLineups = [];  // [g] -> lineup[inning][posIndex]
  const allGamesSitOuts = [];  // [g] -> sitOuts[inning]
  const allGamesBatOrders = []; // [g] -> battingOrder[]
  const allGamesSitOutCaps = []; // [g] -> sitOutCap number
  const allGamesReliefPitchers = []; // [g] -> reliefPitcher name or null
  const allGamesDecisionNotes = []; // [g] -> explanation strings[]

  for (let g = 0; g < games; g++) {
    const availablePlayers = perGameAvailable[g];

    // FRESH per-game state (resets no-return P/C, warmup, 2-inning start)
    const lineup = [];
    const sitOuts = [];
    const inningCountThisGame = {};
    availablePlayers.forEach(p => inningCountThisGame[p] = 0);

    for (let inning = 0; inning < innings; inning++) {
      const numSitOut = availablePlayers.length - POSITIONS.length;
      const sittingOut = [];

      if (numSitOut > 0) {
        const lastSitOuts = inning > 0 ? sitOuts[inning - 1] : [];

        // Pitcher-aware sit-out
        let nextPitcher = null;
        if (depthChart && depthChart['P'] && inning > 0) {
          const currentPitcher = lineup[inning - 1][0];
          for (const candidate of depthChart['P']) {
            if (candidate !== currentPitcher && availablePlayers.indexOf(candidate) >= 0) {
              const hasLeftP = lineup.some((inn, k) => inn[0] === candidate) &&
                lineup[inning - 1][0] !== candidate;
              if (!hasLeftP) {
                nextPitcher = candidate;
                break;
              }
            }
          }
        }

        // Per-game sit-out cap
        const totalSitOutSlots = innings * numSitOut;
        const maxSitOutPerPlayer = Math.ceil(totalSitOutSlots / availablePlayers.length);

        // Sort by: most innings this game, then least cross-game sat out
        const candidates = availablePlayers.slice().sort((a, b) => {
          const inningDiff = inningCountThisGame[b] - inningCountThisGame[a];
          if (inningDiff !== 0) return inningDiff;
          // Sit-out rate (per games played) so missed games don't skew equity
          const rateA = (crossGameSatOutCount[a] || 0) / Math.max(crossGameGamesPlayed[a] || 1, 1);
          const rateB = (crossGameSatOutCount[b] || 0) / Math.max(crossGameGamesPlayed[b] || 1, 1);
          return rateA - rateB;
        });

        const sitOutsThisGame = {};
        availablePlayers.forEach(p => {
          sitOutsThisGame[p] = inning - inningCountThisGame[p];
        });
        const eligible = candidates.filter(c => sitOutsThisGame[c] < maxSitOutPerPlayer);
        const pool = eligible.length >= numSitOut ? eligible : candidates;

        const consecutive = [];
        const nonConsecutive = [];
        for (const c of pool) {
          if (lastSitOuts.indexOf(c) >= 0) {
            consecutive.push(c);
          } else {
            nonConsecutive.push(c);
          }
        }
        const ordered = nonConsecutive.concat(consecutive);

        for (let s = 0; s < numSitOut && s < ordered.length; s++) {
          sittingOut.push(ordered[s]);
        }

        if (nextPitcher && sittingOut.indexOf(nextPitcher) < 0) {
          const currentPitcher = inning > 0 ? lineup[inning - 1][0] : null;
          const nextPitcherSitOuts = inning - inningCountThisGame[nextPitcher];
          if (nextPitcher !== currentPitcher && nextPitcherSitOuts < maxSitOutPerPlayer) {
            sittingOut[sittingOut.length - 1] = nextPitcher;
          }
        }
      }

      sitOuts.push(sittingOut);

      const playing = availablePlayers.filter(p => sittingOut.indexOf(p) < 0);
      const assignment = assignPositions(
        playing, preferences, gamesSinceAtPosition, lineup, inning,
        depthChart, games, seasonPositionCounts, playerMissedGames, modeConfig);
      lineup.push(assignment);

      playing.forEach(p => inningCountThisGame[p]++);
    }

    // Accumulate sit-outs and games played into cross-game trackers
    for (const p of availablePlayers) {
      const sittingOutCount = innings - inningCountThisGame[p];
      crossGameSatOutCount[p] = (crossGameSatOutCount[p] || 0) + sittingOutCount;
      crossGameGamesPlayed[p] = (crossGameGamesPlayed[p] || 0) + 1;
    }

    // Per-game sit-out cap for display
    const numSitOut = availablePlayers.length - POSITIONS.length;
    const sitOutCap = numSitOut > 0 ? Math.ceil(innings * numSitOut / availablePlayers.length) : 0;

    // Per-game batting order — shuffle within tiers for games 2+
    let battingOrder = generateBattingOrder(availablePlayers, battingAverages, modeConfig);
    if (games > 1 && g > 0) {
      battingOrder = shuffleBattingOrderByTier(battingOrder, g, modeConfig);
    }

    // Per-game relief pitcher
    let reliefPitcher = null;
    const startingPitcher = lineup[0][0];
    if (depthChart && depthChart['P']) {
      for (const candidate of depthChart['P']) {
        if (candidate === startingPitcher) continue;
        if (availablePlayers.indexOf(candidate) < 0) continue;
        if (preferences[candidate] && preferences[candidate]['P'] === 'Restricted') continue;
        reliefPitcher = candidate;
        break;
      }
    }

    allGamesLineups.push(lineup);
    allGamesSitOuts.push(sitOuts);
    allGamesBatOrders.push(battingOrder);
    allGamesSitOutCaps.push(sitOutCap);
    allGamesReliefPitchers.push(reliefPitcher);
    allGamesDecisionNotes.push(buildGameDecisionTrace(
      lineup, sitOuts, battingOrder, modeConfig, preferences,
      depthChart, gamesSinceAtPosition, seasonPositionCounts, sitOutCap, reliefPitcher, battingAverages
    ));
  }

  // ── Output writing ──
  const lineupStartRow = 7 + MAX_PLAYERS + 1; // row 20
  const gameColors = ['#4285f4', '#34a853', '#fbbc04']; // blue, green, yellow
  const gameTitleFontColors = ['white', 'white', '#333333'];

  // Calculate generous clear area: enough for MAX_GAMES stacked outputs
  const maxCardRows = games * (MAX_PLAYERS + 13) + 5;
  const maxFieldRows = games * (11 + 3) + 5;
  const maxBattingRows = games * (MAX_PLAYERS + 3) + 5;
  const totalClearRows = maxCardRows + maxFieldRows + maxBattingRows + 10;
  const clearCols = Math.max(2 + innings + 2, POSITIONS.length + 4, 13);

  // Clear everything below the input section
  const clearRange = suggesterSheet.getRange(lineupStartRow, 1, totalClearRows, clearCols);
  clearRange.breakApart().clearContent().clearDataValidations().setBackground(null).setFontStyle(null).setFontWeight(null).setFontColor(null);

  let curRow = lineupStartRow;

  // ── Lineup Cards ──
  suggesterSheet.getRange(curRow, 1).setValue('Lineup Card')
    .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
  curRow++;

  for (let g = 0; g < games; g++) {
    const lineup = allGamesLineups[g];
    const sitOuts = allGamesSitOuts[g];
    const battingOrder = allGamesBatOrders[g];
    const sitOutCap = allGamesSitOutCaps[g];
    const reliefPitcher = allGamesReliefPitchers[g];
    const decisionNotes = allGamesDecisionNotes[g];
    const availablePlayers = perGameAvailable[g];
    const numSitOut = availablePlayers.length - POSITIONS.length;

    // Build stats lookup from batting order
    const statsLookup = {};
    for (const entry of battingOrder) {
      statsLookup[entry.name] = { obp: entry.obp, slg: entry.slg };
    }

    // Game title row (only when games > 1)
    if (games > 1) {
      suggesterSheet.getRange(curRow, 1).setValue('Game ' + (g + 1))
        .setFontSize(12).setFontWeight('bold').setBackground(gameColors[g]).setFontColor(gameTitleFontColors[g]);
      suggesterSheet.getRange(curRow, 1, 1, 2 + innings + 2).setBackground(gameColors[g]);
      curRow++;
    }

    // Card header
    const cardHeaders = ['#', 'Player'];
    for (let i = 1; i <= innings; i++) cardHeaders.push('Inn ' + i);
    cardHeaders.push('OBP', 'SLG');
    suggesterSheet.getRange(curRow, 1, 1, cardHeaders.length).setValues([cardHeaders])
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');
    curRow++;

    // Card data rows
    const cardData = [];
    const cardBackgrounds = [];
    for (let b = 0; b < battingOrder.length; b++) {
      const playerName = battingOrder[b].name;
      const row = [b + 1, playerName];
      const bgRow = [null, null];

      for (let inning = 0; inning < innings; inning++) {
        const posIndex = lineup[inning].indexOf(playerName);
        if (posIndex >= 0) {
          const pos = POSITIONS[posIndex];
          row.push(pos);
          bgRow.push(preferences[playerName] && preferences[playerName][pos] === 'Preferred' ? '#b7e1cd' : null);
        } else if (sitOuts[inning].indexOf(playerName) >= 0) {
          row.push('OUT');
          bgRow.push('#e0e0e0');
        } else {
          row.push('');
          bgRow.push(null);
        }
      }

      const stats = statsLookup[playerName] || { obp: 0, slg: 0 };
      row.push(stats.obp > 0 ? stats.obp.toFixed(3) : '-');
      row.push(stats.slg > 0 ? stats.slg.toFixed(3) : '-');
      bgRow.push(null, null);

      cardData.push(row);
      cardBackgrounds.push(bgRow);
    }

    if (cardData.length > 0) {
      const cardRange = suggesterSheet.getRange(curRow, 1, cardData.length, cardData[0].length);
      cardRange.setValues(cardData);
      cardRange.setBackgrounds(cardBackgrounds);
      suggesterSheet.getRange(curRow, 1, cardData.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
      suggesterSheet.getRange(curRow, 3, cardData.length, innings + 2).setHorizontalAlignment('center');
      suggesterSheet.getRange(curRow, 2, cardData.length, 1).setFontSize(11);

      const posValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(POSITIONS.concat(['OUT', '']), true)
        .setAllowInvalid(true)
        .build();
      suggesterSheet.getRange(curRow, 3, cardData.length, innings).setDataValidation(posValidation);
      curRow += cardData.length;
    }

    // Summary rows
    if (numSitOut > 0) {
      suggesterSheet.getRange(curRow, 1).setValue('Max sit-outs per player: ' + sitOutCap)
        .setFontStyle('italic');
      suggesterSheet.getRange(curRow, 1, 1, 4).mergeAcross();
      curRow++;
    }

    if (reliefPitcher) {
      const startingPitcher = lineup[0][0];
      const pitchesAllGame = lineup.every(inn => inn[0] === startingPitcher);
      const reliefLabel = pitchesAllGame
        ? 'Relief pitcher (if needed): ' + reliefPitcher + ' — starter pitches all ' + innings + ' innings'
        : 'Relief pitcher (if needed): ' + reliefPitcher;
      suggesterSheet.getRange(curRow, 1).setValue(reliefLabel)
        .setFontStyle('italic');
      suggesterSheet.getRange(curRow, 1, 1, Math.max(6, cardHeaders.length)).mergeAcross();
      curRow++;
    }

    // Weekly IP tracking for pitchers in this game
    const gamePitchers = {};
    for (let inn = 0; inn < innings; inn++) {
      const pitcher = lineup[inn][0];
      gamePitchers[pitcher] = (gamePitchers[pitcher] || 0) + 1;
    }
    const ipParts = [];
    for (const p of Object.keys(gamePitchers)) {
      const existing = weeklyIP[p] || 0;
      const thisGame = gamePitchers[p];
      ipParts.push(p + ': ' + existing + '+' + thisGame + '=' + (existing + thisGame));
    }
    if (ipParts.length > 0) {
      suggesterSheet.getRange(curRow, 1).setValue('Weekly IP (prior+this game): ' + ipParts.join(', '))
        .setFontStyle('italic');
      suggesterSheet.getRange(curRow, 1, 1, Math.max(6, cardHeaders.length)).mergeAcross();
      curRow++;
    }

    if (decisionNotes && decisionNotes.length > 0) {
      suggesterSheet.getRange(curRow, 1).setValue('Decision Notes')
        .setFontWeight('bold').setBackground('#e8f0fe');
      suggesterSheet.getRange(curRow, 1, 1, Math.max(6, cardHeaders.length)).mergeAcross();
      curRow++;
      for (let i = 0; i < decisionNotes.length; i++) {
        suggesterSheet.getRange(curRow, 1).setValue(decisionNotes[i]).setFontStyle('italic');
        suggesterSheet.getRange(curRow, 1, 1, Math.max(6, cardHeaders.length)).mergeAcross();
        curRow++;
      }
    }

    curRow++; // gap between games
  }

  // ── Field Lineup grids ──
  curRow++; // extra gap
  for (let g = 0; g < games; g++) {
    const lineup = allGamesLineups[g];
    const sitOuts = allGamesSitOuts[g];
    const availablePlayers = perGameAvailable[g];

    const fieldLabel = games > 1
      ? 'Game ' + (g + 1) + ' Field Lineup (copy to Game Entry)'
      : 'Suggested Field Lineup (for Game Entry copy-paste)';
    suggesterSheet.getRange(curRow, 1).setValue(fieldLabel)
      .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
    curRow++;

    // Grid header
    const gridHeaders = ['Inning'];
    POSITIONS.forEach(p => gridHeaders.push(p));
    for (let s = 1; s <= 3; s++) gridHeaders.push('Sit Out ' + s);
    suggesterSheet.getRange(curRow, 1, 1, gridHeaders.length).setValues([gridHeaders])
      .setFontWeight('bold').setBackground('#34a853').setFontColor('white').setHorizontalAlignment('center');
    curRow++;

    // Game title row in grid (only when games > 1)
    if (games > 1) {
      suggesterSheet.getRange(curRow - 2, 1, 1, gridHeaders.length)
        .setBackground(gameColors[g]);
      // Re-set the title text styling after background
      suggesterSheet.getRange(curRow - 2, 1).setFontColor(gameTitleFontColors[g]);
    }

    const lineupData = [];
    const lineupBackgrounds = [];
    for (let inning = 0; inning < innings; inning++) {
      const row = [inning + 1];
      const bgRow = [null];
      for (let j = 0; j < POSITIONS.length; j++) {
        const playerName = lineup[inning][j];
        row.push(playerName);
        bgRow.push(preferences[playerName] && preferences[playerName][POSITIONS[j]] === 'Preferred' ? '#b7e1cd' : null);
      }
      for (let s = 0; s < 3; s++) {
        row.push(s < sitOuts[inning].length ? sitOuts[inning][s] : '');
        bgRow.push(null);
      }
      lineupData.push(row);
      lineupBackgrounds.push(bgRow);
    }

    if (lineupData.length > 0) {
      const outputRange = suggesterSheet.getRange(curRow, 1, lineupData.length, lineupData[0].length);
      outputRange.setValues(lineupData);
      outputRange.setBackgrounds(lineupBackgrounds);
      suggesterSheet.getRange(curRow, 1, lineupData.length, 1).setHorizontalAlignment('center').setFontWeight('bold');

      const editRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(availablePlayers, true)
        .setAllowInvalid(true)
        .build();
      suggesterSheet.getRange(curRow, 2, lineupData.length, POSITIONS.length).setDataValidation(editRule);
      suggesterSheet.getRange(curRow, POSITIONS.length + 2, lineupData.length, 3).setDataValidation(editRule);
      curRow += lineupData.length;
    }

    curRow += 2; // gap between games
  }

  // ── Batting Order grids ──
  for (let g = 0; g < games; g++) {
    const battingOrder = allGamesBatOrders[g];

    const battingLabel = games > 1
      ? 'Game ' + (g + 1) + ' Batting Order'
      : 'Suggested Batting Order';
    suggesterSheet.getRange(curRow, 1).setValue(battingLabel)
      .setFontSize(13).setFontWeight('bold').setBackground('#e8f0fe');
    curRow++;

    if (games > 1) {
      suggesterSheet.getRange(curRow - 1, 1, 1, 5).setBackground(gameColors[g]);
      suggesterSheet.getRange(curRow - 1, 1).setFontColor(gameTitleFontColors[g]);
    }

    const battingHeaders = ['#', 'Player', 'OBP', 'SLG', 'SB'];
    suggesterSheet.getRange(curRow, 1, 1, battingHeaders.length).setValues([battingHeaders])
      .setFontWeight('bold').setBackground('#fbbc04').setFontColor('white').setHorizontalAlignment('center');
    curRow++;

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
      suggesterSheet.getRange(curRow, 1, battingData.length, 5).setValues(battingData);
      suggesterSheet.getRange(curRow, 1, battingData.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
      suggesterSheet.getRange(curRow, 3, battingData.length, 3).setHorizontalAlignment('center');
      suggesterSheet.getRange(curRow, 2, battingData.length, 1).setFontSize(11);
      curRow += battingData.length;
    }

    curRow += 2; // gap between games
  }

  suggesterSheet.activate();
  const gameLabel = games > 1 ? games + ' games of ' + innings + ' innings' : innings + ' innings';
  const firstReliefPitcher = allGamesReliefPitchers.find(r => r);
  const reliefMsg = firstReliefPitcher ? '\nRelief pitcher: ' + firstReliefPitcher : '';
  ui.alert('Lineup Generated',
    'A ' + modeConfig.name.toLowerCase() + ' lineup has been generated for ' + gameLabel + '.' +
    reliefMsg +
    '\n\nField positions, batting order, and decision notes are shown below.\n' +
    'You can manually edit any cell using the dropdowns.\n' +
    'Copy to the Game Entry sheet when ready.',
    ui.ButtonSet.OK);
}

function assignPositions(players, preferences, gamesSinceAtPosition, previousInnings, currentInning,
    depthChart, totalGames, seasonPositionCounts, playerMissedGames, modeConfig) {
  modeConfig = modeConfig || getCoachingModeConfig('Balanced');
  const fieldConfig = modeConfig.field;
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

      // No-return rule for P (hard block) and C (soft penalty):
      // if a player left the position, discourage or block returning
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
            if (j === 0) {
              score = 10000; // P: hard block — never return after leaving
            } else {
              score += 200; // C: strongly discouraged but allowed if needed
            }
          }
        }
      }

      // Minimum 2-inning starter rule for P and C:
      // In inning 1 (second inning), heavily penalize changing the starter
      if (score < 10000 && (j === 0 || j === 1) && currentInning === 1) {
        const starterAtPos = previousInnings[0][j];
        if (playerName !== starterAtPos) {
          score += 500; // very strong penalty — keep the starter for at least 2 innings
        }
      }

        // Skip all bonuses if already blocked (Restricted or no-return rule)
      if (score < 10000) {
        if (pref === 'Preferred') {
          score += fieldConfig.preferredBonus; // bonus for preferred
        }
        // Okay = neutral (0)

        // Depth chart: ranked players get a bonus
        if (depthChart && depthChart[pos]) {
          const rank = depthChart[pos].indexOf(playerName);
          if (rank >= 0) {
            score -= (MAX_PLAYERS - rank) * fieldConfig.depthWeight;
          }
        }

        // Recency: prioritize positions not played recently
        const gamesSince = (gamesSinceAtPosition[playerName] && gamesSinceAtPosition[playerName][pos]) || 0;
        score -= gamesSince * fieldConfig.recencyWeight; // more games since = lower score = more priority

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
              score += fieldConfig.pitcherCatcherContinuityBonus; // P/C: strong incentive to keep pitcher/catcher in place
            } else if (consecutiveCount >= 3) {
              score += fieldConfig.fieldFourthPlusPenalty; // 4th+ consecutive inning at same field position: penalty to encourage rotation
            } else if (consecutiveCount >= 2) {
              score += fieldConfig.fieldThirdInningPenalty; // 3rd consecutive inning: mild penalty to start encouraging moves
            } else {
              score += fieldConfig.fieldSecondInningBonus; // 2nd consecutive inning: small bonus for stability
            }
          }
        }

        // Outfield-only avoidance: if a player has only played OF positions (LF/CF/RF)
        // so far this game, give infield positions a bonus to mix them in
        // In multi-game mode, kick in earlier and with a stronger bonus
        const multiGame = (totalGames || 1) > 1;
        const ofThreshold = multiGame ? 1 : 2; // trigger after 1 inning in multi-game, 2 in single
        if (currentInning >= ofThreshold && j < 6) { // current position is infield (P/C/1B/2B/3B/SS)
          let allOutfield = true;
          let inningsPlayed = 0;
          for (let k = 0; k < currentInning; k++) {
            const prevIdx = previousInnings[k].indexOf(playerName);
            if (prevIdx >= 0) {
              inningsPlayed++;
              if (prevIdx < 6) { // was at an infield position
                allOutfield = false;
                break;
              }
            }
          }
          if (allOutfield && inningsPlayed >= ofThreshold) {
            score -= multiGame ? fieldConfig.outfieldOnlyBonusMulti : fieldConfig.outfieldOnlyBonus;
          }
        }

        // Position diversity: small bonus for positions rarely/never played this season
        // Intentionally small — preferences and depth chart should dominate
        if (seasonPositionCounts && seasonPositionCounts[playerName]) {
          const posInnings = seasonPositionCounts[playerName][pos] || 0;
          const totalInn = Object.values(seasonPositionCounts[playerName]).reduce((s, v) => s + v, 0);
          if (totalInn > 0 && posInnings === 0) {
            score += fieldConfig.diversityNeverBonus; // small bonus for never-played positions
          } else if (totalInn > 0 && posInnings / totalInn < 0.08) {
            score += fieldConfig.diversityRareBonus; // tiny bonus for rarely-played positions
          }
        }

        // Small attendance penalty: missing games slightly deprioritizes position assignments
        const missed = (playerMissedGames && playerMissedGames[playerName]) || 0;
        if (missed > 0) {
          score += missed * fieldConfig.missedGamePenalty;
        }
      }

      scores[p][j] = score;
    }
  }

  // In multi-game mode, add 15% jitter to non-blocked scores for position variety
  if ((totalGames || 1) > 1) {
    // Seeded PRNG from currentInning so each inning gets different but deterministic jitter
    let seed = (currentInning + 1) * 2654435761;
    function nextRand() {
      seed = (seed * 1664525 + 1013904223) & 0x7fffffff;
      return seed / 0x7fffffff;
    }
    // Find the score range among non-blocked entries for scaling
    let minScore = Infinity, maxScore = -Infinity;
    for (let p = 0; p < numPlayers; p++) {
      for (let j = 0; j < numPositions; j++) {
        if (scores[p][j] < 10000) {
          if (scores[p][j] < minScore) minScore = scores[p][j];
          if (scores[p][j] > maxScore) maxScore = scores[p][j];
        }
      }
    }
    const range = Math.max(maxScore - minScore, 1);
    const jitterScale = range * 0.15;
    for (let p = 0; p < numPlayers; p++) {
      for (let j = 0; j < numPositions; j++) {
        if (scores[p][j] < 10000) {
          scores[p][j] += (nextRand() * 2 - 1) * jitterScale;
        }
      }
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
    let fallbackPlayer = -1;
    let fallbackScore = Infinity;

    for (let p = 0; p < numPlayers; p++) {
      if (assigned.has(p)) continue;
      if (scores[p][posIdx] < 10000 && scores[p][posIdx] < bestScore) {
        bestScore = scores[p][posIdx];
        bestPlayer = p;
      }
      // Track best fallback in case all players are blocked (score >= 10000)
      if (scores[p][posIdx] < fallbackScore) {
        fallbackScore = scores[p][posIdx];
        fallbackPlayer = p;
      }
    }

    if (bestPlayer >= 0) {
      assignment[posIdx] = players[bestPlayer];
      assigned.add(bestPlayer);
    } else if (fallbackPlayer >= 0) {
      // No unblocked player — assign the least-bad option to avoid a blank
      assignment[posIdx] = players[fallbackPlayer];
      assigned.add(fallbackPlayer);
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
    ['⚾ SOFTBALL LINEUP MANAGER - Instructions', ''],   // 1
    ['', ''],                                              // 2
    ['FINDING THE MENU', ''],                              // 3
    ['Look for ⚾ Softball in the menu bar', 'It appears at the far right end, after Extensions and Help'],
    ['First time setup:', 'Go to Extensions > Apps Script, select onOpen, click Run (▶), and authorize when prompted'],
    ['If you don\'t see it after that', 'Close and reopen the spreadsheet — the menu loads automatically on each open'],
    ['', ''],                                              // 7
    ['GETTING STARTED', ''],                               // 8
    ['1. Go to the Roster sheet', 'Enter your players\' names (up to 12 players)'],
    ['2. Set position preferences', 'For each player, set each position as Preferred (green), Okay (yellow), or Restricted (red)'],
    ['3. That\'s it!', 'You\'re ready to manage games'],
    ['', ''],                                              // 12
    ['ENTERING A GAME', ''],                               // 13
    ['1. Go to the Game Entry sheet', 'Fill in the date, opponent, and number of innings'],
    ['2. Mark attendance', 'Uncheck absent players using the checkboxes in the left sidebar'],
    ['3. Fill in the lineup grid', 'Use dropdowns to assign one player per position per inning'],
    ['4. Mark who sat out', 'Use the Sit Out columns (up to 3 players can sit out per inning)'],
    ['5. Enter batting stats (below lineup)', 'Fill in AB, hits (1B/2B/3B/HR), BB, SB, and CS for each player'],
    ['6. Save the game', 'Click ⚾ Softball > Save Game — validates for errors before saving'],
    ['', ''],                                              // 20
    ['SAVE GAME VALIDATION', ''],                          // 21
    ['The system checks before saving:', 'Duplicate players in one inning and absent players in the lineup are blocked'],
    ['If you see an error:', 'Fix the issue on the Game Entry sheet and save again'],
    ['', ''],                                              // 24
    ['DELETING A GAME', ''],                               // 25
    ['Click ⚾ Softball > Delete Last Game', 'Removes the most recent game from Season History and Batting Stats'],
    ['Confirm the deletion', 'The game number and opponent are shown — this cannot be undone'],
    ['', ''],                                              // 28
    ['DEPTH CHART', ''],                                   // 29
    ['1. Go to the Depth Chart sheet', 'Rank players per position (1st = top choice, leave blank for unranked)'],
    ['2. Fill in rankings', 'Use the dropdowns to select which player is your 1st, 2nd, 3rd choice, etc. at each position'],
    ['3. How it works', 'Ranked players get a scoring bonus when the Lineup Suggester assigns positions'],
    ['4. Interaction with preferences', 'Restricted still blocks a player even if ranked 1st. Depth chart fine-tunes choices among Preferred/Okay players'],
    ['', ''],                                              // 34
    ['USING THE LINEUP SUGGESTER', ''],
    ['1. Go to the Lineup Suggester sheet', 'Check the boxes next to available players'],
    ['2. Rest P / Rest C columns', 'Check these to hold a player back from Pitcher or Catcher (great for friendlies or resting arms)'],
    ['3. Set the number of innings', ''],
    ['4. Set the number of games (1-3)', 'For single games leave at 1. For tournament days set to 2 or 3 (see Tournament Mode below)'],
    ['5. Choose a coaching mode', 'Balanced mixes fairness and fit, Competitive leans into best current lineup, Development pushes broader reps'],
    ['6. Click ⚾ Softball > Suggest Lineup', 'The algorithm generates field positions AND a batting order for each game'],
    ['7. Review the output', 'Sit-out cap, relief pitcher suggestion, and decision notes are shown per game below the lineup'],
    ['8. Edit if needed', 'Use dropdowns to make manual adjustments to field positions and sit-outs'],
    ['9. Batting order section', 'Shows suggested batting order based on OBP, slugging, and speed stats'],
    ['10. Copy to Game Entry', 'Each game has its own labeled Field Lineup grid for easy copy-paste'],
    ['', ''],
    ['TOURNAMENT MODE (MULTI-GAME)', ''],
    ['When to use:', 'Tournament days with 2-3 back-to-back games where you can\'t enter/save between games'],
    ['1. Set Games to 2 or 3', 'G1/G2/G3 columns appear so you can set per-game attendance (a player can miss specific games)'],
    ['2. Per-game attendance', 'G1/G2/G3 default to the master checkbox. Uncheck a player from a specific game if they\'re leaving early or arriving late'],
    ['3. What resets per game', 'No-return pitcher rule, 2-inning starter minimum, and position assignments all reset each game'],
    ['4. What carries across games', 'Sit-out fairness — players who sit more in game 1 sit less in game 2'],
    ['5. Infield rotation', 'The algorithm works harder to get everyone infield time across multiple games'],
    ['6. Batting order variety', 'Order is shuffled slightly each game — players stay near their tier but not in the exact same spot'],
    ['7. Output format', 'Lineup cards, field lineups, and batting orders are stacked per game with color-coded titles (blue/green/yellow)'],
    ['8. Relief pitcher', 'Each game gets its own relief pitcher suggestion based on that game\'s available players'],
    ['', ''],
    ['UNDERSTANDING THE BATTING ORDER', ''],
    ['Spots 1-3 (top of order):', 'Best OBP + speed — players who get on base and steal'],
    ['Spots 4-6 (middle):', 'Best slugging — power hitters who drive in runs'],
    ['Spots 7+ (bottom):', 'Remaining players by overall composite score'],
    ['Stability:', 'Players move at most 2 spots from their recent average position'],
    ['New players (< 3 games):', 'Default to roster order until enough data is collected'],
    ['', ''],
    ['VIEWING THE DASHBOARD', ''],
    ['1. Go to the Dashboard sheet', 'Click ⚾ Softball > Refresh Dashboard to update stats'],
    ['2. Section 1: Innings at Each Position', 'Shows total innings each player has played at each position all season'],
    ['3. Section 2: Games Since Last Played', 'Yellow = 3+ games since, Red = 5+ games since playing that position'],
    ['4. Section 3: Batting Stats', 'Shows OBP, SLG, stolen bases, and caught stealing for each player'],
    ['', ''],
    ['TIPS', ''],
    ['• Sit-out cap:', 'No player sits out more than their fair share per game — the cap is shown in the lineup output'],
    ['• Field position rotation:', 'Players rotate across field positions — the algorithm penalizes staying at the same non-P/C spot for 3+ innings'],
    ['• Outfield-only avoidance:', 'Players who have only played outfield for 2+ innings get a bonus toward infield positions'],
    ['• No-return rule for P (hard):', 'Once a player leaves Pitcher, they cannot return to that position later in the game (resets each game)'],
    ['• No-return rule for C (soft):', 'Once a player leaves Catcher, the algorithm strongly avoids putting them back but will allow it if needed'],
    ['• Bullpen warmup:', 'The algorithm prefers pitchers who sat out the previous inning to warm up, but will assign one without warmup if needed'],
    ['• Relief pitcher:', 'A suggested relief pitcher is shown per game in case the starter needs to come out'],
    ['• Rest P / Rest C:', 'Use these checkboxes on the Lineup Suggester to rest key players from P or C for specific games'],
    ['• Absent players:', 'Uncheck on Game Entry before saving — they are excluded from season history and don\'t affect recency scoring'],
    ['• Weekly IP tracking:', 'The lineup output shows each pitcher\'s rolling 7-day innings pitched (prior + this game) for arm management'],
    ['• Position diversity:', 'Small bonus nudges players toward positions they haven\'t played — preferences and depth chart always take priority'],
    ['• Attendance equity:', 'Sit-out fairness uses per-game rate so missed games don\'t skew the rotation. Missing games applies a small position penalty'],
    ['• Coaching modes:', 'Balanced mixes fairness and fit, Competitive favors strongest current alignment, Development pushes overdue reps and position variety'],
    ['• Decision notes:', 'Each suggested lineup includes a short explanation of the opening defense, batting-order logic, sit-outs, and relief pitcher choice'],
    ['• Tournament fairness:', 'Cross-game sit-out tracking ensures players who sit more in one game sit less in the next'],
    ['• Delete Last Game:', 'Use ⚾ Softball > Delete Last Game to undo the most recently saved game'],
    ['• Dashboard colors:', 'Yellow = 3+ games since, Red = 5+ games since playing a position'],
    ['• Season History sheet:', 'Stores all game data — don\'t edit directly unless fixing errors'],
    ['• Batting Stats sheet:', 'Stores per-game batting data — editable to fix errors, then Refresh Dashboard'],
    ['• Updating the code:', 'Paste new Code.gs, save, then run rebuildGameEntry and initializeStep2 — your data is preserved'],
  ];

  sheet.getRange(1, 1, instructions.length, 2).setValues(instructions);

  // Formatting
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');

  // Bold column A for all rows (step numbers and bullets will be bold)
  sheet.getRange(1, 1, instructions.length, 1).setFontWeight('bold');
  // Section header rows are the rows with text in column A and blank column B
  for (let i = 0; i < instructions.length; i++) {
    if (i === 0) continue;
    if (instructions[i][0] && instructions[i][1] === '') {
      sheet.getRange(i + 1, 1, 1, 2).setFontSize(13).setBackground('#e8f0fe');
    }
  }
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

  // Watch Games input on Lineup Suggester (cell D4) to show/hide per-game columns
  // Layout: C=Rest P, D=Rest C, E=G1, F=G2, G=G3
  if (sheet.getName() === 'Lineup Suggester' && e.range.getRow() === 4 && e.range.getColumn() === 4) {
    const games = Number(e.range.getValue());
    // Read master checkbox (col A) to default G1/G2/G3 attendance
    const masterValues = sheet.getRange(7, 1, MAX_PLAYERS, 1).getValues();
    if (games === 1) {
      // Single game: hide all game columns, master checkbox = attendance
      sheet.hideColumns(5); // G1
      sheet.hideColumns(6); // G2
      sheet.hideColumns(7); // G3
    } else {
      // Multi-game: show G1 + relevant game columns, default to master checkbox
      sheet.showColumns(5); // G1
      const g1Vals = masterValues.map(row => [row[0]]);
      sheet.getRange(7, 5, MAX_PLAYERS, 1).setValues(g1Vals);
      if (games >= 2) {
        sheet.showColumns(6); // G2
        sheet.getRange(7, 6, MAX_PLAYERS, 1).setValues(g1Vals);
      } else {
        sheet.hideColumns(6);
      }
      if (games >= 3) {
        sheet.showColumns(7); // G3
        sheet.getRange(7, 7, MAX_PLAYERS, 1).setValues(g1Vals);
      } else {
        sheet.hideColumns(7);
      }
    }
  }
}

function updateSuggesterNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Lineup Suggester');
  if (!sheet) return;

  const players = getRosterNames();
  // Read existing: col A=checkbox, B=name, C=Rest P, D=Rest C, E=G1, F=G2, G=G3
  const existingChecks = sheet.getRange(7, 1, MAX_PLAYERS, 1).getValues();
  const existingNames = sheet.getRange(7, 2, MAX_PLAYERS, 1).getValues();
  const existingExtra = sheet.getRange(7, 3, MAX_PLAYERS, 5).getValues();

  const checkValues = [];
  const nameValues = [];
  const extraValues = [];
  for (let i = 0; i < MAX_PLAYERS; i++) {
    if (i < players.length) {
      const nameChanged = existingNames[i][0] !== players[i];
      checkValues.push([nameChanged ? true : existingChecks[i][0]]);
      nameValues.push([players[i]]);
      if (nameChanged) {
        extraValues.push([false, false, true, false, false]); // RestP=f, RestC=f, G1=true, G2=f, G3=f
      } else {
        extraValues.push([existingExtra[i][0], existingExtra[i][1], existingExtra[i][2], existingExtra[i][3], existingExtra[i][4]]);
      }
    } else {
      checkValues.push([false]);
      nameValues.push(['']);
      extraValues.push([false, false, false, false, false]);
    }
  }
  sheet.getRange(7, 1, MAX_PLAYERS, 1).setValues(checkValues);
  sheet.getRange(7, 2, MAX_PLAYERS, 1).setValues(nameValues);
  sheet.getRange(7, 3, MAX_PLAYERS, 5).setValues(extraValues);
}
