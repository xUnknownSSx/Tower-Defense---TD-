// Map Path Coordinates [row, col]
const PATH = [
  [2,1], [2,2], [2,3], [2,4], [2,5], [2,6], [2,7], [2,8], [2,9], [2,10],
  [3,10], [4,10], [5,10], [6,10],
  [6,9], [6,8], [6,7], [6,6], [6,5], [6,4], [6,3], [6,2],
  [7,2], [8,2], [9,2], [10,2],
  [10,3], [10,4], [10,5], [10,6], [10,7], [10,8], [10,9], [10,10], [10,11], [10,12],
  [11,12], [12,12], [13,12], [14,12], [15,12], [16,12], [17,12], [18,12], [19,12]
];

// Colors
const COLOR_PATH = "#d9d9d9";
const COLOR_BG = "#202020";
const COLOR_SODA = "#0000ff";  
const COLOR_CANDY = "#ff00ff"; 

const COLOR_ANT = "#ff3333";   // Wave 1-2
const COLOR_RAT = "#cc0000";   // Wave 3-4
const COLOR_MOLD = "#00cc44";  // Wave 5+ Boss

function getGameSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName("Game") || ss.getActiveSheet();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Vending TD')
    .addItem('Reset Game', 'setupGame')
    .addItem('Next Step', 'gameTick')
    .addItem('Start Auto Play', 'startAutoRun')
    .addItem('Stop Auto Play', 'stopAutoRun')
    .addItem('Place Tower', 'placeTower')
    .addToUi();
}

function setupGame() {
  const sheet = getGameSheet();
  const scriptProps = PropertiesService.getScriptProperties();
  
  sheet.getRange("W2").setValue(100); // Cents
  sheet.getRange("W3").setValue(10);  // Lives
  sheet.getRange("W4").setValue(1);   // Wave
  
  const grid = sheet.getRange(1, 1, 20, 20);
  grid.setBackground(COLOR_BG);
  grid.clearContent();
  
  PATH.forEach(coord => {
    sheet.getRange(coord[0], coord[1]).setBackground(COLOR_PATH);
  });
  
  scriptProps.setProperty("IS_RUNNING", "false");
  scriptProps.setProperty("ACTIVE_ENEMIES", "[]"); 
  scriptProps.setProperty("ENEMIES_LEFT", "6");
  scriptProps.setProperty("SPAWN_COOLDOWN", "0");
}

function gameTick() {
  const sheet = getGameSheet();
  const scriptProps = PropertiesService.getScriptProperties();
  
  let cents = Number(sheet.getRange("W2").getValue());
  let lives = Number(sheet.getRange("W3").getValue());
  let wave = Number(sheet.getRange("W4").getValue());
  
  if (lives <= 0) {
    scriptProps.setProperty("IS_RUNNING", "false");
    SpreadsheetApp.getUi().alert("Game Over! The pests raided the vending machine!");
    return false;
  }

  let activeEnemies = JSON.parse(scriptProps.getProperty("ACTIVE_ENEMIES") || "[]");
  let enemiesLeft = parseInt(scriptProps.getProperty("ENEMIES_LEFT") || "6");
  let spawnCooldown = parseInt(scriptProps.getProperty("SPAWN_COOLDOWN") || "0");

  // 1. Clear current enemy cells from track
  activeEnemies.forEach(e => {
    if (e.pos >= 0 && e.pos < PATH.length) {
      const c = PATH[e.pos];
      sheet.getRange(c[0], c[1]).setBackground(COLOR_PATH).clearContent();
    }
  });

  const gridValues = sheet.getRange(1, 1, 20, 20).getValues();

  // 2. Move existing active enemies & apply turret damage
  let survivors = [];
  for (let i = 0; i < activeEnemies.length; i++) {
    let e = activeEnemies[i];
    e.pos++; // Step forward

    if (e.pos >= PATH.length) {
      // Reached drop tray!
      lives--;
      sheet.getRange("W3").setValue(lives);
      if (lives <= 0) {
         scriptProps.setProperty("IS_RUNNING", "false");
         SpreadsheetApp.getUi().alert("Game Over! The pests raided the vending machine!");
         return false;
      }
    } else {
      // Calculate damage from nearby turrets
      let damage = 0;
      let r = PATH[e.pos][0] - 1;
      let c = PATH[e.pos][1] - 1;

      for (let row = Math.max(0, r - 2); row <= Math.min(19, r + 2); row++) {
        for (let col = Math.max(0, c - 2); col <= Math.min(19, c + 2); col++) {
          let val = gridValues[row][col];
          if (val === "S") damage += 20; // Soda
          if (val === "C") damage += 12; // Candy
        }
      }

      e.hp -= damage;

      if (e.hp <= 0) {
        cents += 10 + wave;
        sheet.getRange("W2").setValue(cents);
      } else {
        survivors.push(e);
      }
    }
  }

  // 3. Multi-Enemy Spawner Logic (Wave 3+ Spawns Multiple Enemies at Once)
  if (enemiesLeft > 0) {
    if (spawnCooldown <= 0) {
      // Determine how many enemies to spawn together based on Wave
      let groupSize = 1; 
      if (wave >= 5) groupSize = 3;      // Wave 5+: Triple spawn
      else if (wave >= 3) groupSize = 2; // Wave 3-4: Double spawn

      groupSize = Math.min(groupSize, enemiesLeft);

      // Check if entry positions on the track are clear
      let spaceNeeded = groupSize;
      let canSpawn = true;
      for (let s = 0; s < spaceNeeded; s++) {
        if (survivors.some(e => e.pos === s)) {
          canSpawn = false;
          break;
        }
      }

      if (canSpawn) {
        for (let g = 0; g < groupSize; g++) {
          // Spawn enemies back-to-back at positions 0, 1, 2...
          survivors.push({ pos: g, hp: 30 + (wave * 15) });
          enemiesLeft--;
        }
        spawnCooldown = 3; // Wait 3 steps before releasing the next group
      }
    } else {
      spawnCooldown--;
    }
  }

  // 4. Wave Completion Check
  if (survivors.length === 0 && enemiesLeft === 0) {
    scriptProps.setProperty("IS_RUNNING", "false");
    scriptProps.setProperty("ACTIVE_ENEMIES", "[]");
    
    let bonus = 30 + (wave * 10);
    sheet.getRange("W2").setValue(cents + bonus);
    
    wave++;
    sheet.getRange("W4").setValue(wave);
    
    let nextWaveEnemies = 6 + (wave * 3);
    scriptProps.setProperty("ENEMIES_LEFT", nextWaveEnemies.toString());
    scriptProps.setProperty("SPAWN_COOLDOWN", "0");
    
    SpreadsheetApp.getUi().alert("Wave " + (wave - 1) + " Cleared!\n+" + bonus + " Bonus Cents!\nGet ready for Wave " + wave + "!");
    return false;
  }

  // 5. Draw all active enemies on board
  let enemyColor = COLOR_ANT;
  if (wave >= 5) enemyColor = COLOR_MOLD;
  else if (wave >= 3) enemyColor = COLOR_RAT;

  survivors.forEach(e => {
    const coord = PATH[e.pos];
    sheet.getRange(coord[0], coord[1])
         .setBackground(enemyColor)
         .setValue(e.hp)
         .setFontColor("#ffffff")
         .setFontSize(7)
         .setHorizontalAlignment("center");
  });

  // 6. Save Updated State
  scriptProps.setProperty("ACTIVE_ENEMIES", JSON.stringify(survivors));
  scriptProps.setProperty("ENEMIES_LEFT", enemiesLeft.toString());
  scriptProps.setProperty("SPAWN_COOLDOWN", spawnCooldown.toString());

  return true;
}

function startAutoRun() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty("IS_RUNNING", "true");

  for (let i = 0; i < 150; i++) {
    if (scriptProps.getProperty("IS_RUNNING") !== "true") {
      break;
    }

    let continueLoop = gameTick();
    SpreadsheetApp.flush();

    if (!continueLoop) {
      break;
    }

    Utilities.sleep(1500);
  }
}

function stopAutoRun() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty("IS_RUNNING", "false");
}

function placeTower() {
  const sheet = getGameSheet();
  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();
  
  if (row > 20 || col > 20) {
    SpreadsheetApp.getUi().alert("Please select a cell inside the 20x20 display grid!");
    return;
  }
  
  const isOnPath = PATH.some(coord => coord[0] === row && coord[1] === col);
  if (isOnPath) {
    SpreadsheetApp.getUi().alert("You cannot build on the vending machine track!");
    return;
  }

  let cents = sheet.getRange("W2").getValue();
  let towerType = sheet.getRange("W6").getValue();
  
  if (towerType === "Soda Can" && cents >= 30) {
    sheet.getRange("W2").setValue(cents - 30);
    cell.setBackground(COLOR_SODA).setValue("S").setFontColor("#ffffff");
  } else if (towerType === "Candy Bar" && cents >= 20) {
    sheet.getRange("W2").setValue(cents - 20);
    cell.setBackground(COLOR_CANDY).setValue("C").setFontColor("#ffffff");
  } else {
    SpreadsheetApp.getUi().alert("Not enough cents!");
  }
}