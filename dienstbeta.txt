/**
 * die übergeordneten Steuerungsfunktionen.
 * Version 12: Konsolidiert, entfernt redundante Hilfsfunktionen.
 * Fügt eine Funktion zum Erstellen einer Sicherheitskopie der gesamten Tabelle hinzu.
 * Fügt eine Funktion zum automatischen Anordnen der Blätter hinzu.
 * Verschiebt "Blätter anordnen" direkt ins Hauptmenü.
 * Automatisiert die Blattsortierung nach dem Anlegen der zukünftigen Blätter.
 * Automatisiert die Erstellung von Sicherheitskopien am 1. und 10. jedes Monats.
 */

// Globale Konstanten für die Blattnamen und Konfigurationen
const CONSTANTS = {
  CONFIG_SHEET: 'Konfiguration',
  DOCTORS_SHEET: 'Ärzte & Stammdaten',
  JUMPER_LIST_SHEET: 'Einspringliste',
  PLAN_SHEET_PREFIX: '', // Kein Präfix mehr für Plan-Blätter
  MONTHS_TO_KEEP: 6,
  MONTHS_IN_ADVANCE: 4, // Blätter für 4 Monate im Voraus anlegen
  HOLIDAY_STATE_CODE: 'RP'
};
const DUTY_RULES = {
  MAIN_COOLDOWN_DAYS: 4
};
const MS_PER_DAY = 24 * 60 * 60 * 1000;

function normalizeDate(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function diffInDays(laterDate, earlierDate) {
  if (!laterDate || !earlierDate) {
    return null;
  }
  const normalizedLater = normalizeDate(laterDate);
  const normalizedEarlier = normalizeDate(earlierDate);
  return Math.floor((normalizedLater.getTime() - normalizedEarlier.getTime()) / MS_PER_DAY);
}

function computeRelativeDayIndex(date, monthStart) {
  if (!date || !monthStart) {
    return null;
  }
  const diff = diffInDays(date, monthStart);
  return diff === null ? null : diff + 1;
}

function getPreviousYearMonth(year, month) {
  const d = new Date(year, month, 1);
  d.setMonth(d.getMonth() - 1);
  return { year: d.getFullYear(), month: d.getMonth() };
}

function loadRecentHistoryFromPreviousMonth(ss, year, month) {
  const history = {};
  const previous = getPreviousYearMonth(year, month);
  const sheetName = getFormattedSheetName(CONSTANTS.PLAN_SHEET_PREFIX, previous.year, previous.month);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return history;
  }

  const data = sheet.getDataRange().getValues();
  let planHeaderRowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row.includes('Dienstarzt')) {
      planHeaderRowIndex = i;
      break;
    }
  }
  if (planHeaderRowIndex === -1) {
    return history;
  }

  const header = data[planHeaderRowIndex];
  const dateColIndex = 0;
  const dutyDoctorColIndex = header.indexOf('Dienstarzt');
  const lateDoctorColIndex = header.indexOf('Spätdienstarzt');

  if (dutyDoctorColIndex === -1) {
    return history;
  }

  const monthStart = new Date(year, month, 1);
  let currentDate = null;

  for (let i = planHeaderRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    if (row[dateColIndex] instanceof Date) {
      currentDate = row[dateColIndex];
    }
    if (!currentDate) {
      continue;
    }

    const dutyTypeCell = (row[2] || '').toString().toLowerCase();
    let normalizedType = dutyTypeCell;
    if (dutyTypeCell.includes('24')) {
      normalizedType = '24h';
    }

    const mainDoctorRaw = row[dutyDoctorColIndex];
    const lateDoctorRaw = lateDoctorColIndex !== -1 ? row[lateDoctorColIndex] : null;

    if (mainDoctorRaw) {
      const name = mainDoctorRaw.toString().trim();
      if (name) {
        const entry = history[name] || {};
        if (!entry.lastMainDutyDate || currentDate.getTime() >= entry.lastMainDutyDate.getTime()) {
          entry.lastMainDutyDate = new Date(currentDate.getTime());
          entry.lastMainDutyType = normalizedType;
          entry.lastMainDutyDayIndex = computeRelativeDayIndex(currentDate, monthStart);
        }
        history[name] = entry;
      }
    }

    if (lateDoctorRaw) {
      const name = lateDoctorRaw.toString().trim();
      if (name) {
        const entry = history[name] || {};
        if (!entry.lastLateShiftDate || currentDate.getTime() >= entry.lastLateShiftDate.getTime()) {
          entry.lastLateShiftDate = new Date(currentDate.getTime());
          entry.lastLateShiftDayIndex = computeRelativeDayIndex(currentDate, monthStart);
        }
        history[name] = entry;
      }
    }
  }

  return history;
}

/**
 * Erstellt das Admin-Menü beim Öffnen des Dokuments.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Dienstplan Admin')
    .addItem('1. Einspringliste aktualisieren', 'updateJumperList')
    .addItem('2. Plan generieren', 'generateMonthlyPlan')
    .addItem('3. Nächste Monate anlegen', 'checkAndCreateFutureSheets')
    .addItem('4. Einzelnen Monat anlegen', 'createPastPlanSheet')
    .addItem('5. Alte Blätter löschen', 'cleanupOldSheets')
    .addSeparator();
    
  const automationMenu = ui.createMenu('Automatisierung')
    .addItem('Automatische Blatt-Verwaltung EINRICHTEN', 'setupAutomationTriggers')
    .addItem('Alle Automatisierungen LÖSCHEN', 'deleteAllTriggers');

  menu
    .addSubMenu(automationMenu)
    .addSeparator()
    .addItem('6. Basisdaten & Setup', 'initialSetup')
    .addSeparator()
    .addItem('7. Sicherheitskopie erstellen', 'createSpreadsheetBackup')
    .addSeparator()
    .addItem('8. Blätter sortieren', 'arrangeSheets')
    .addToUi();
}


/**
 * Führt das einmalige Setup aus. Baut die Struktur sicher auf.
 * Fügt eine doppelte Bestätigung hinzu, um versehentliche Datenresets zu verhindern.
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirmation1 = ui.alert(
    'Setup bestätigen',
    'Dies wird alle Konfigurations-, Ärzte- und Einspringerdaten zurücksetzen und die Basisstruktur neu erstellen. Möchten Sie fortfahren?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation1 !== ui.Button.YES) {
    ss.toast('Setup abgebrochen.', 'Hinweis', 5);
    return;
  }

  const confirmation2 = ui.alert(
    'WIRKLICH bestätigen?',
    'Alle vorhandenen Daten in "Konfiguration", "Ärzte & Stammdaten" und "Einspringliste" werden UNWIDERRUFLICH gelöscht und neu erstellt. Sind Sie ABSOLUT sicher?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation2 !== ui.Button.YES) {
    ss.toast('Setup abgebrochen.', 'Hinweis', 5);
    return;
  }

  try {
    createConfigSheet();
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createDoctorsSheet();
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createJumperListSheet();
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    if (ss.getSheets().length > 1) {
      const defaultSheet = ss.getSheetByName('Tabellenblatt1') || ss.getSheetByName('Sheet1');
      if (defaultSheet) {
        ss.deleteSheet(defaultSheet);
      }
    }

    ss.toast('Die Basis-Struktur wurde erfolgreich erstellt.', 'Setup erfolgreich!', 5);
  } catch (e) {
    handleError(e, 'Fehler beim initialen Setup.');
  }
}

/**
 * Steuert die Erstellung der Plan-Blätter für die Zukunft.
 */
function checkAndCreateFutureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let createdCount = 0;
    for (let i = 0; i <= CONSTANTS.MONTHS_IN_ADVANCE; i++) {
      const targetDate = new Date();
      targetDate.setMonth(targetDate.getMonth() + i);
      if (createMonthlyPlanSheet(targetDate.getFullYear(), targetDate.getMonth())) {
        createdCount++;
      }
    }
    if (createdCount > 0 && !isTriggered()) {
        ss.toast(`${createdCount} neue Plan-Blätter wurden erstellt.`, 'Erfolg', 5);
    } else if (createdCount === 0 && !isTriggered()) {
        ss.toast('Keine neuen Plan-Blätter erforderlich oder erstellt.', 'Hinweis', 5);
    }
  } catch (e) {
    handleError(e, 'Fehler beim Erstellen der zukünftigen Plan-Blätter.');
  }
}

/**
 * Fragt den Nutzer nach einem Monat und erstellt ein leeres Plan-Blatt für die Vergangenheit.
 */
function createPastPlanSheet() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Einzelnen Monat anlegen',
        'Bitte geben Sie den Monat und das Jahr im Format MM/JJ ein (z.B. 06/25):',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() == ui.Button.OK) {
        const text = response.getResponseText();
        const date = parseDateFromSheetName(text);
        
        if (date) {
            createMonthlyPlanSheet(date.getFullYear(), date.getMonth());
            SpreadsheetApp.getActiveSpreadsheet().toast(`Das leere Plan-Blatt für ${text} wurde erstellt.`, 'Erfolg', 5);
        } else {
            SpreadsheetApp.getActiveSpreadsheet().toast('Ungültiges Format. Bitte verwenden Sie MM/JJ.', 'Fehler', 5);
        }
    }
}

/**
 * Generiert den Dienstplan für den Monat, der 3 Monate in der Zukunft liegt.
 * Diese Funktion wird von einem automatischen Trigger aufgerufen.
 */
function generateFutureMonthlyPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const targetDate = new Date();
    targetDate.setMonth(targetDate.getMonth() + 3); // 3 Monate in der Zukunft
    targetDate.setDate(1); // Setze auf den 1. des Monats für korrekte Blattnamen-Generierung

    const sheetName = getFormattedSheetName(CONSTANTS.PLAN_SHEET_PREFIX, targetDate.getFullYear(), targetDate.getMonth());
    const targetSheet = ss.getSheetByName(sheetName);

    if (!targetSheet) {
      Logger.log(`Zielblatt '${sheetName}' für die Plangenerierung nicht gefunden. Versuche, es zu erstellen.`);
      createMonthlyPlanSheet(targetDate.getFullYear(), targetDate.getMonth());
      SpreadsheetApp.flush();
      const createdSheet = ss.getSheetByName(sheetName);
      if (!createdSheet) {
        handleError(new Error(`Konnte Zielblatt '${sheetName}' für die Plangenerierung nicht finden oder erstellen.`));
        return;
      }
      ss.setActiveSheet(createdSheet);
    } else {
      ss.setActiveSheet(targetSheet);
    }

    Logger.log(`Generiere Plan für aktives Blatt: ${ss.getActiveSheet().getName()}`);
    generateMonthlyPlan();
    ss.toast(`Dienstplan für ${sheetName} wurde automatisch generiert.`, 'Automatisierung', 5);

  } catch (e) {
    handleError(e, 'Fehler bei der automatischen Dienstplangenerierung für zukünftigen Monat.');
  }
}

/**
 * Erstellt eine Sicherheitskopie der gesamten Google-Tabelle im Google Drive.
 */
function createSpreadsheetBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const file = DriveApp.getFileById(ss.getId());
    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const backupName = `${ss.getName()}_Backup_${timestamp}`;
    
    // Erstellt eine Kopie im selben Ordner wie die Originaldatei
    const backupFile = file.makeCopy(backupName);
    
    ss.toast(`Sicherheitskopie '${backupName}' erfolgreich in Google Drive erstellt.`, 'Sicherung erstellt', 5);
    Logger.log(`Sicherheitskopie '${backupName}' erstellt unter ID: ${backupFile.getId()}`);

  } catch (e) {
    handleError(e, 'Fehler beim Erstellen der Sicherheitskopie.');
  }
}

/**
 * Ordnet die Blätter in der Tabelle neu an:
 * Konfiguration, Ärzte & Stammdaten, Einspringliste,
 * gefolgt von Monatsblättern (MM/JJ) absteigend sortiert.
 */
function arrangeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const fixedSheets = [];
  const monthlySheets = [];
  const otherSheets = []; // Für alle anderen Blätter, die nicht erkannt werden

  // 1. Blätter kategorisieren
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === CONSTANTS.CONFIG_SHEET ||
        sheetName === CONSTANTS.DOCTORS_SHEET ||
        sheetName === CONSTANTS.JUMPER_LIST_SHEET) {
      fixedSheets.push(sheet);
    } else if (parseDateFromSheetName(sheetName)) {
      monthlySheets.push(sheet);
    } else {
      otherSheets.push(sheet);
    }
  });

  // 2. Monatsblätter sortieren (absteigend, neuestes zuerst)
  monthlySheets.sort((a, b) => {
    const dateA = parseDateFromSheetName(a.getName()).getTime();
    const dateB = parseDateFromSheetName(b.getName()).getTime();
    return dateB - dateA; // Absteigend sortieren
  });

  // 3. Gewünschte Reihenfolge erstellen (Namen)
  const desiredOrderNames = [
    CONSTANTS.CONFIG_SHEET,
    CONSTANTS.DOCTORS_SHEET,
    CONSTANTS.JUMPER_LIST_SHEET
  ];
  monthlySheets.forEach(sheet => desiredOrderNames.push(sheet.getName()));
  otherSheets.forEach(sheet => desiredOrderNames.push(sheet.getName())); // Andere Blätter am Ende

  // 4. Blätter neu anordnen
  // Es ist am einfachsten, die Blätter von hinten nach vorne zu verschieben,
  // um Indexprobleme zu vermeiden.
  for (let i = desiredOrderNames.length - 1; i >= 0; i--) {
    const sheetName = desiredOrderNames[i];
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1); // moveActiveSheet ist 1-basiert
    }
  }
  
  ss.toast('Blätter erfolgreich neu angeordnet.', 'Anordnung abgeschlossen', 5);
  Logger.log('Blätter erfolgreich neu angeordnet.');
}


/**
 * Aufräumen alter Blätter.
 */
function cleanupOldSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isTriggered()) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert('Bestätigung', `Sollen wirklich alle Blätter, die älter als ${CONSTANTS.MONTHS_TO_KEEP} Monate sind, gelöscht werden?`, ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) return;
  }
  try {
    const allSheets = ss.getSheets();
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - CONSTANTS.MONTHS_TO_KEEP);
    let deletedCount = 0;
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (parseDateFromSheetName(sheetName)) {
        const sheetDate = parseDateFromSheetName(sheetName);
        if (sheetDate && sheetDate < cutoffDate) {
          ss.deleteSheet(sheet);
          deletedCount++;
        }
      }
    });
    if (deletedCount > 0 && !isTriggered()) {
        ss.toast(`${deletedCount} alte Blätter wurden gelöscht.`, 'Aufräumen abgeschlossen', 5);
    } else if (deletedCount === 0 && !isTriggered()) {
        ss.toast('Keine alten Blätter zum Löschen gefunden.', 'Aufräumen abgeschlossen', 5);
    }
  } catch (e) {
    handleError(e, 'Fehler beim Aufräumen der alten Blätter.');
  }
}

/**
 * Richtet Automatisierungs-Trigger ein.
 */
function setupAutomationTriggers() {
    deleteAllTriggers(true);

    // Trigger 1: Erstellt zukünftige Monatsblätter (für 4 Monate im Voraus)
    // Läuft am 1. jeden Monats um 2:00 Uhr morgens
    ScriptApp.newTrigger('checkAndCreateFutureSheets')
        .timeBased()
        .onMonthDay(1)
        .atHour(2)
        .create();

    // Trigger 2: Ordnet die Blätter nach dem Anlegen der neuen Blätter
    // Läuft am 1. jeden Monats um 2:15 Uhr morgens
    ScriptApp.newTrigger('arrangeSheets')
        .timeBased()
        .onMonthDay(1)
        .atHour(2)
        .nearMinute(15) // 15 Minuten nach dem Anlegen der Blätter
        .create();

    // Trigger 3: Aktualisiert die Einspringliste
    // Läuft am 1. jeden Monats um 2:30 Uhr morgens
    ScriptApp.newTrigger('updateJumperList')
        .timeBased()
        .onMonthDay(1)
        .atHour(2)
        .nearMinute(30)
        .create();


    // Trigger 5: Räumt alte Blätter auf
    // Läuft am 1. jeden Monats um 4:00 Uhr morgens
    ScriptApp.newTrigger('cleanupOldSheets')
        .timeBased()
        .onMonthDay(1)
        .atHour(4)
        .create();

    // NEU: Trigger 6: Erstellt eine Sicherheitskopie am 1. des Monats
    ScriptApp.newTrigger('createSpreadsheetBackup')
        .timeBased()
        .onMonthDay(1)
        .atHour(5) // Eine Stunde nach dem Aufräumen
        .create();

    // NEU: Trigger 7: Erstellt eine Sicherheitskopie am 10. des Monats
    ScriptApp.newTrigger('createSpreadsheetBackup')
        .timeBased()
        .onMonthDay(10)
        .atHour(5) // Um die gleiche Uhrzeit wie der erste Backup-Trigger
        .create();

    SpreadsheetApp.getActiveSpreadsheet().toast('Die Automatisierung für Monatsblätter, Dienstplan, Einspringliste, Aufräumen und Sicherungskopien wurde eingerichtet.', 'Automatisierung eingerichtet!', 5);
}

/**
 * Löscht alle Automatisierungs-Trigger.
 * @param {boolean} isSilent - Ob die Erfolgsmeldung unterdrückt werden soll.
 */
function deleteAllTriggers(isSilent = false) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() !== 'onEdit') {
            ScriptApp.deleteTrigger(trigger);
        }
    });
    if (!isSilent){
      SpreadsheetApp.getActiveSpreadsheet().toast('Alle automatischen Prozesse wurden gestoppt.', 'Automatisierung entfernt', 5);
    }
}

/**
 * @OnlyCurrentDoc
 *
 * planungslogik.gs: Enthält die Kernlogik für die Erstellung des Dienstplans.
 * Version mit 2x12h-Wunsch-Priorisierung und dynamischer Verteilung:
 * - NEU: Eine separate Logik am Anfang von `distributeMainDuties` behandelt 2x12h-Wochenend-Wünsche mit höchster Priorität.
 * - Der Hauptalgorithmus bewertet nach jeder Zuweisung die Lage neu und besetzt den jeweils schwierigsten Dienst (Engpass).
 */

/**
 * Hauptfunktion, die den gesamten Planungsprozess für den Monat des AKTIVEN Blattes steuert.
 */
function generateMonthlyPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  if (!parseDateFromSheetName(sheetName)) {
    ss.toast('Dies ist kein gültiges Plan-Blatt. Der Name sollte dem Format MM/JJ entsprechen.', 'Falsches Blatt', 5);
    return;
  }

  const planDate = parseDateFromSheetName(sheetName);
  
  try {
    ss.toast('Planung wird generiert...', 'Status', -1);

    createPlanLayout(activeSheet);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    const planningData = getPlanningData(planDate.getFullYear(), planDate.getMonth());
    if (!planningData) {
      ss.toast('Planung abgebrochen: Daten konnten nicht geladen werden.');
      return;
    }

    let plan = initializeEmptyPlan(planningData);
    
    planningData.dutyTargets = calculateDutyTargets(plan, planningData.doctors, planningData.jumperPointsMap);
    planningData.rawDutyTargets = { ...planningData.dutyTargets };

    // Der neue, leistungsfähigere Hauptalgorithmus
    plan = distributeMainDuties(plan, planningData);
    
    // Balancierung nach der Hauptverteilung
    plan = balanceMainDuties(plan, planningData);
    
    // Spätdienste
    plan = distributeLateShifts(plan, planningData);

    // Finales Schreiben
    writePlanToSheet(plan, planningData, planningData.stats);

    ss.toast(`Plan für ${sheetName} wurde erfolgreich generiert.`, 'Abgeschlossen', 5);

  } catch (e) {
    handleError(e, 'Ein schwerwiegender Fehler ist während der Planung aufgetreten.');
    ss.toast('Planung fehlgeschlagen. Siehe Logs für Details.', 'Fehler', 10);
  }
}

/**
 * Sammelt alle relevanten Daten für die Planung aus den verschiedenen Blättern.
 */
function getPlanningData(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getFormattedSheetName(CONSTANTS.PLAN_SHEET_PREFIX, year, month);
  const wishSheet = ss.getSheetByName(sheetName);

  if (!wishSheet) {
    handleError(new Error(`Das Plan-Blatt "${sheetName}" konnte nicht gefunden werden.`));
    return null;
  }

  const doctorsSheet = ss.getSheetByName(CONSTANTS.DOCTORS_SHEET);
  if (!doctorsSheet) {
    handleError(new Error(`Das Blatt "${CONSTANTS.DOCTORS_SHEET}" konnte nicht gefunden werden.`));
    return null;
  }
  const doctors = doctorsSheet.getRange('A2:F' + doctorsSheet.getLastRow()).getValues()
    .map(row => ({
      name: row[0].toString().trim(),
      factor: parseFloat(row[1] || 1.0),
      station: row[2],
      doesDuties: (row[3] || '').toString().toUpperCase() === 'JA',
      doesLateShifts: (row[4] || '').toString().toUpperCase() === 'JA',
      isJumper: (row[5] || '').toString().toUpperCase() === 'JA'
    }))
    .filter(d => d.name);

  const configSheet = ss.getSheetByName(CONSTANTS.CONFIG_SHEET);
  if (!configSheet) {
    handleError(new Error(`Das Blatt "${CONSTANTS.CONFIG_SHEET}" konnte nicht gefunden werden.`));
    return null;
  }
  const configValues = configSheet.getRange('A2:B' + configSheet.getLastRow()).getValues();
  const config = configValues.reduce((obj, row) => {
    if (row[0] && row[1]) obj[row[0]] = row[1];
    return obj;
  }, {});

  const ektPlanRaw = configSheet.getRange('E2:F15').getValues();
  const ektPlan = ektPlanRaw
    .filter(row => row[0] && row[1] && row[0].toLowerCase() !== 'station')
    .reduce((obj, row) => {
      obj[row[0]] = row[1];
      return obj;
    }, {});

  const wishesRaw = wishSheet.getDataRange().getValues();
  const wishes = {};

  doctors.forEach(doc => { wishes[doc.name] = {}; });

  const numDoctorColsInWishSheet = Math.floor((wishesRaw[0].length - 3) / 2);
  const doctorsToProcessForWishes = doctors.slice(0, numDoctorColsInWishSheet);

  let currentDate = null;
  for (let i = 2; i < wishesRaw.length; i++) {
    const row = wishesRaw[i];
    if (row[0] instanceof Date) currentDate = row[0];
    if (!currentDate) continue;

    const day = currentDate.getDate();
    const dutyType = (row[2] || '').toString().toLowerCase();

    for (let k = 0; k < doctorsToProcessForWishes.length; k++) {
      const doctor = doctorsToProcessForWishes[k];
      const doctorName = doctor.name;
      const dutyColIndex = 3 + (k * 2);
      const lateColIndex = dutyColIndex + 1;

      if (dutyColIndex >= row.length) continue;

      const rawDutyWish = row[dutyColIndex];
      const rawLateWish = row[lateColIndex];

      let dutyWish = '';
      if (rawDutyWish !== null && rawDutyWish !== undefined && rawDutyWish !== '') {
        const cleanedWish = rawDutyWish.toString().replace(/\s/g, '').toUpperCase();
        if (['N', 'W', '24H'].includes(cleanedWish)) {
          dutyWish = cleanedWish;
        }
      }

      let lateWish = '';
      if (rawLateWish !== null && rawLateWish !== undefined && rawLateWish !== '') {
        const cleanedWish = rawLateWish.toString().replace(/\s/g, '').toUpperCase();
        if (['N', 'W'].includes(cleanedWish)) {
          lateWish = cleanedWish;
        }
      }

      if (!wishes[doctorName][day]) wishes[doctorName][day] = {};

      if (dutyType.includes('24h')) {
        if (dutyWish) wishes[doctorName][day]['24h'] = dutyWish;
        if (lateWish) wishes[doctorName][day]['late'] = lateWish;
      } else {
        if (dutyWish) wishes[doctorName][day][dutyType] = dutyWish;
        if (lateWish && dutyType === 'tag') wishes[doctorName][day]['late'] = lateWish;
      }
    }
  }

  const jumperSheet = ss.getSheetByName(CONSTANTS.JUMPER_LIST_SHEET);
  let jumperPointsMap = {};
  if (jumperSheet) {
    const jumperData = jumperSheet.getRange('B2:C' + jumperSheet.getLastRow()).getValues();
    jumperData.forEach(row => {
      if (row[0]) {
        jumperPointsMap[row[0].toString().trim()] = row[1] || 0;
      }
    });
  }

  const holidays = getHolidaysForYear(year);
  const recentHistory = loadRecentHistoryFromPreviousMonth(ss, year, month);

  return { year, month, doctors, config, ektPlan, wishes, holidays, jumperPointsMap, recentHistory };
}

/**
 * Initialisiert ein leeres Dienstplan-Objekt für den Monat.
 */
function initializeEmptyPlan(planningData) {
  const daysInMonth = new Date(planningData.year, planningData.month + 1, 0).getDate();
  const plan = {};
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(planningData.year, planningData.month, day);
    const dayOfWeek = date.getDay();
    const isHoliday = planningData.holidays.some(h => h.getTime() === date.getTime());
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

    plan[day] = {
      date: date,
      duty_24h: null,
      duty_tag: null,
      duty_nacht: null,
      late: null,
      dutyTypes: (isWeekend || isHoliday) ? ['tag', 'nacht'] : ['24h'],
      isWeekendDay: (isWeekend || isHoliday),
      lateShiftPossible: !(isWeekend || isHoliday)
    };
  }
  return plan;
}

/**
 * Baut die Verfügbarkeitsmatrix für Ärzte basierend auf NOGO-Wünschen und EKT-Regeln.
 */
function buildAvailabilityMatrix(plan, planningData) {
  const availability = {};
  const weekDays = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];

  planningData.doctors.forEach(doc => {
    availability[doc.name] = {};
    for (const day in plan) {
      availability[doc.name][day] = { '24h': true, 'tag': true, 'nacht': true, 'late': true };
    }
  });

  for (const name in planningData.wishes) {
    for (const day in planningData.wishes[name]) {
      for (const dutyType in planningData.wishes[name][day]) {
        if (planningData.wishes[name][day][dutyType] === 'N') {
          availability[name][day][dutyType] = false;
        }
      }
    }
  }

  for (const station in planningData.ektPlan) {
    const ektDayName = planningData.ektPlan[station];
    const doctorsOnStation = planningData.doctors.filter(d => d.station === station);
    const ektDayOfWeekIndex = weekDays.indexOf(ektDayName);
    if (ektDayOfWeekIndex === -1) continue;

    const dayBeforeEktDayOfWeekIndex = (ektDayOfWeekIndex - 1 + 7) % 7;
    const dayBeforeEktDayName = weekDays[dayBeforeEktDayOfWeekIndex];

    for (const day in plan) {
      if (weekDays[plan[day].date.getDay()] === ektDayName || weekDays[plan[day].date.getDay()] === dayBeforeEktDayName) {
        doctorsOnStation.forEach(doc => {
          Object.keys(availability[doc.name][day]).forEach(key => availability[doc.name][day][key] = false);
        });
      }
    }
  }
  return availability;
}


/**
 * NEUE VERSION: Verteilt Hauptdienste iterativ, mit vorgeschalteter 2x12h-Wunsch-Priorisierung.
 */
function distributeMainDuties(plan, planningData) {
  const { doctors, wishes, config, jumperPointsMap, dutyTargets, recentHistory = {} } = planningData;
  const maxWeekends = config['Max. Wochenenddienste pro Monat'] || 2;
  const maxMainDutiesPerMonth = config['Max. Hauptdienste pro Monat'] || 5;

  const stats = {};
  doctors.forEach(doc => {
    const historyEntry = recentHistory[doc.name] || {};
    const statsEntry = {
      totalDuties: 0,
      mainDuties: 0,
      weekendsWorked: 0,
      workedWeekendKeys: new Set(),
      assignedMainDuties: {},
      lastMainDutyDate: historyEntry.lastMainDutyDate || null,
      lastMainDutyType: historyEntry.lastMainDutyType || null,
      lastLateShiftDate: historyEntry.lastLateShiftDate || null
    };

    if (historyEntry.lastMainDutyDate) {
      const lastWeekendKey = getWeekendKey(historyEntry.lastMainDutyDate);
      const dayOfWeek = historyEntry.lastMainDutyDate.getDay();
      if (dayOfWeek === 5 || dayOfWeek === 6 || dayOfWeek === 0) {
        statsEntry.workedWeekendKeys.add(lastWeekendKey);
      }
    }

    stats[doc.name] = statsEntry;
  });

  planningData.stats = stats;
  planningData.availability = buildAvailabilityMatrix(plan, planningData);

  Logger.log("Starte Hauptdienst-Verteilung mit 2x12h-Priorisierung.");

  let remainingSlots = [];
  for (const day in plan) {
    for (const dutyType of plan[day].dutyTypes) {
      remainingSlots.push({ day: parseInt(day, 10), type: dutyType });
    }
  }

  const weekendTagSlots = remainingSlots.filter(s => s.type === 'tag' && plan[s.day].isWeekendDay);

  for (const slot of weekendTagSlots) {
    const { day } = slot;
    const categorized = categorizeDoctorsForMainDuty(day, 'tag', plan, planningData, stats, maxWeekends, maxMainDutiesPerMonth, dutyTargets);
    const candidates = categorized.available.filter(doc => (wishes[doc.name]?.[day]?.['tag'] || '').includes('24H'));

    if (candidates.length > 0) {
      const chosenDoctor = candidates[0];
      const nightCategories = categorizeDoctorsForMainDuty(day, 'nacht', plan, planningData, stats, maxWeekends, maxMainDutiesPerMonth, dutyTargets);
      const canDoNight = !nightCategories.unavailable.some(d => d.name === chosenDoctor.name);

      if (canDoNight) {
        Logger.log(`Erfülle 2x12h-Wunsch für ${chosenDoctor.name} am ${day}.`);
        plan[day]['duty_tag'] = chosenDoctor.name;
        plan[day]['duty_nacht'] = chosenDoctor.name;

        stats[chosenDoctor.name].mainDuties += 2;
        stats[chosenDoctor.name].totalDuties += 2;
        stats[chosenDoctor.name].assignedMainDuties[day] = '24h_weekend';
        stats[chosenDoctor.name].lastMainDutyDate = plan[day].date;
        stats[chosenDoctor.name].lastMainDutyType = 'nacht';

        const weekendKey = getWeekendKey(plan[day].date);
        if (!stats[chosenDoctor.name].workedWeekendKeys.has(weekendKey)) {
          stats[chosenDoctor.name].weekendsWorked++;
          stats[chosenDoctor.name].workedWeekendKeys.add(weekendKey);
        }

        remainingSlots = remainingSlots.filter(s => !(s.day === day && (s.type === 'tag' || s.type === 'nacht')));
      }
    }
  }

  remainingSlots = remainingSlots.filter(slot => !plan[slot.day][`duty_${slot.type}`]);

  while (remainingSlots.length > 0) {
    const slotEvaluations = remainingSlots.map(slot => {
      const categories = categorizeDoctorsForMainDuty(slot.day, slot.type, plan, planningData, stats, maxWeekends, maxMainDutiesPerMonth, dutyTargets);
      const wishCandidates = categories.available.filter(doc => (wishes[doc.name]?.[slot.day]?.[slot.type] || '').includes('W'));
      const availablePoolSize = categories.available.length + categories.notfalls.length;
      const hasWish = wishCandidates.length > 0;
      return { ...slot, categories, wishCandidates, availablePoolSize, hasWish };
    });

    slotEvaluations.sort((a, b) => {
      if (a.hasWish && !b.hasWish) return -1;
      if (!a.hasWish && b.hasWish) return 1;
      return a.availablePoolSize - b.availablePoolSize;
    });

    const slotToProcess = slotEvaluations[0];
    if (!slotToProcess) {
      break;
    }

    const { day, type, categories, wishCandidates } = slotToProcess;
    const candidatePool = wishCandidates.length > 0
      ? wishCandidates
      : (categories.available.length > 0 ? categories.available : categories.notfalls);

    const chosenDoctor = chooseBestDoctor(candidatePool, stats, dutyTargets, planningData);

    if (chosenDoctor) {
      plan[day][`duty_${type}`] = chosenDoctor.name;
      updateStatsForNewAssignment(chosenDoctor.name, day, type, plan, planningData);
    } else {
      plan[day][`duty_${type}`] = 'UNBESETZT';
    }

    remainingSlots = remainingSlots.filter(s => !(s.day === day && s.type === type));
  }

  Logger.log("Dynamische Hauptdienst-Verteilung abgeschlossen.");
  return plan;
}


/**
 * Balanciert die Hauptdienste nach der initialen Verteilung, um die Varianz zu reduzieren.
 */
function balanceMainDuties(plan, planningData) {
  const { doctors, wishes, config, stats, dutyTargets } = planningData;
  const maxWeekends = config['Max. Wochenenddienste pro Monat'] || 2;
  const maxMainDutiesPerMonth = config['Max. Hauptdienste pro Monat'] || 5;

  Logger.log('Starte Hauptdienst-Balancierung (15 Iterationen)');

  for (let iteration = 0; iteration < 15; iteration++) {
    let swappedInIteration = false;
    planningData.availability = buildAvailabilityMatrix(plan, planningData);

    const doctorDutyStatus = doctors.filter(doc => doc.doesDuties).map(doc => ({
      doctor: doc,
      difference: (stats[doc.name].mainDuties || 0) - (dutyTargets[doc.name] || 0)
    })).sort((a, b) => a.difference - b.difference);

    const doctorsUnder = doctorDutyStatus.filter(d => d.difference < 0).map(d => d.doctor);
    const doctorsOver = doctorDutyStatus.filter(d => d.difference > 0).reverse().map(d => d.doctor);

    if (doctorsUnder.length === 0 || doctorsOver.length === 0) {
        break;
    }

    for (const doctorOver of doctorsOver) {
      if (swappedInIteration) break;
      const assignedDuties = Object.keys(stats[doctorOver.name].assignedMainDuties);

      for (const dayStr of assignedDuties) {
        if (swappedInIteration) break;
        const day = parseInt(dayStr);
        const dutyType = stats[doctorOver.name].assignedMainDuties[day];
        const dutyKey = `duty_${dutyType}`;
        
        const wish = wishes[doctorOver.name]?.[day]?.[dutyType] || '';
        if (wish.includes('W') || wish.includes('24H') || dutyType === '24h_weekend') continue;

        for (const doctorUnder of doctorsUnder) {
            const originalAssignee = plan[day][dutyKey];
            plan[day][dutyKey] = null;
            
            const categorized = categorizeDoctorsForMainDuty(day, dutyType, plan, planningData, stats, maxWeekends, maxMainDutiesPerMonth, dutyTargets);
            const canTake = categorized.available.some(d => d.name === doctorUnder.name) || categorized.notfalls.some(d => d.name === doctorUnder.name);

            plan[day][dutyKey] = originalAssignee;

            if (canTake) {
                updateStatsForSwap(doctorOver.name, doctorUnder.name, day, dutyType, plan, planningData);
                plan[day][dutyKey] = doctorUnder.name;
                swappedInIteration = true;
                break; 
            }
        }
      }
    }
    if (!swappedInIteration) {
        Logger.log(`Kein Tausch in Iteration ${iteration + 1} gefunden. Balancierung beendet.`);
        break;
    }
  }
  return plan;
}

/**
 * Hilfsfunktion, um die Statistiken nach einem Tausch zu aktualisieren.
 */
function updateStatsForSwap(oldDoctorName, newDoctorName, day, dutyType, plan, planningData) {
    updateStatsForRemoval(oldDoctorName, day, dutyType, plan, planningData);
    updateStatsForNewAssignment(newDoctorName, day, dutyType, plan, planningData);
}

/**
 * Hilfsfunktion, um die Statistiken für einen Arzt zu reduzieren, wenn sein Dienst entfernt wird.
 */
function updateStatsForRemoval(doctorName, day, dutyType, plan, planningData) {
    const { stats } = planningData;
    const dayInfo = plan[day];
    const isWeekendShift = dayInfo.isWeekendDay || (dayInfo.date.getDay() === 5 && dutyType === '24h');
    const weekendKey = getWeekendKey(dayInfo.date);

    stats[doctorName].mainDuties--;
    stats[doctorName].totalDuties--;
    delete stats[doctorName].assignedMainDuties[day];
    if (isWeekendShift && !isWeekendDutyStillNeeded(doctorName, day, plan)) {
        stats[doctorName].weekendsWorked--;
        stats[doctorName].workedWeekendKeys.delete(weekendKey);
    }
    recomputeLastMainDutyDate(doctorName, stats, plan, planningData);
}

/**
 * Hilfsfunktion, um die Statistiken für eine neue Zuweisung zu aktualisieren.
 */
function updateStatsForNewAssignment(doctorName, day, dutyType, plan, planningData) {
    const { stats, availability } = planningData;
    const dayInfo = plan[day];
    const isWeekendShift = dayInfo.isWeekendDay || (dayInfo.date.getDay() === 5 && dutyType === '24h');
    const weekendKey = getWeekendKey(dayInfo.date);

    stats[doctorName].mainDuties++;
    stats[doctorName].totalDuties++;
    stats[doctorName].assignedMainDuties[day] = dutyType;
    stats[doctorName].lastMainDutyDate = dayInfo.date;
    stats[doctorName].lastMainDutyType = dutyType;

    if (isWeekendShift && !stats[doctorName].workedWeekendKeys.has(weekendKey)) {
        stats[doctorName].weekendsWorked++;
        stats[doctorName].workedWeekendKeys.add(weekendKey);
    }
    // Blockiere Folgetag nach 24h-Dienst
    if (dutyType === '24h' && plan[day + 1]) {
        Object.keys(availability[doctorName][day + 1]).forEach(key => {
            availability[doctorName][day + 1][key] = false;
        });
    }
}
/**
 * @OnlyCurrentDoc
 *
 * sheet-erstellung.gs: Enthält alle Funktionen, die für die Erstellung und Formatierung
 * der Basis-Tabellenblätter zuständig sind.
 * Version 9: Bereinigt, entfernt redundante createMonthlyPlanSheet und Debug-Logs.
 */

// Globale Konstanten werden normalerweise in code.txt definiert.
// Stellen Sie sicher, dass CONSTANTS in Ihrem Projekt global verfügbar ist.
// Falls nicht, müssen Sie diese hier oder in einer anderen globalen Datei definieren.
// Beispiel:
/*
const CONSTANTS = {
  CONFIG_SHEET: 'Konfiguration',
  DOCTORS_SHEET: 'Ärzte & Stammdaten',
  JUMPER_LIST_SHEET: 'Einspringliste',
  PLAN_SHEET_PREFIX: '', // Kein Präfix mehr für Plan-Blätter
  MONTHS_TO_KEEP: 6,
  MONTHS_IN_ADVANCE: 4,
  HOLIDAY_STATE_CODE: 'RP'
};
*/

// Die Funktion createMonthlyPlanSheet wurde entfernt,
// da sie redundant in Plantabelle.txt existiert und von dort aufgerufen wird.
// Alle Debug-Logs wurden ebenfalls entfernt.


/**
 * Erstellt das 'Konfiguration' Blatt mit sauber getrennten Bereichen.
 */
function createConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONSTANTS.CONFIG_SHEET);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONSTANTS.CONFIG_SHEET, 0);
  }
  
  const requiredCols = 12;
  const maxCols = sheet.getMaxColumns();
  if (maxCols < requiredCols) {
    sheet.insertColumnsAfter(maxCols, requiredCols - maxCols);
  }

  sheet.getRange('A1').setValue('Globale Konfiguration').setFontWeight('bold');
  const configData = [['Parameter', 'Wert', 'Beschreibung'],['Punkte pro regulärem Dienst', 15, '...'],['Zusatzpunkte pro Einspringdienst', 10, '...'],['Max. Wochenenddienste pro Monat', 2, '...'],];
  sheet.getRange('A2:C5').setValues(configData);
  sheet.getRange('A2:C2').setFontWeight('bold');

  sheet.getRange('E1').setValue('EKT-Plan').setFontWeight('bold');
  const ektData = [['Station', 'EKT-Wochentag'],['Station 1', 'Montag'],['Station 3', 'Dienstag'],['Station 4', 'Mittwoch'],['Station 6', 'Donnerstag'],['Station 7', 'Freitag']];
  sheet.getRange('E2:F7').setValues(ektData);
  sheet.getRange('E2:F2').setFontWeight('bold');

  sheet.getRange('H1').setValue('Feiertage (Rheinland-Pfalz)').setFontWeight('bold');
  const currentYear = new Date().getFullYear();
  
  const holidaysCurrent = getHolidaysForYear(currentYear);
  if (holidaysCurrent.length > 0) {
    const holidaysCurrentData = holidaysCurrent.map(d => [d, d.toLocaleDateString('de-DE', {weekday: 'long'})]);
    sheet.getRange('H2').setValue(currentYear).setFontWeight('bold');
    sheet.getRange(3, 8, holidaysCurrentData.length, 2).setValues(holidaysCurrentData);
  }
  
  const holidaysNext = getHolidaysForYear(currentYear + 1);
  if (holidaysNext.length > 0) {
    const holidaysNextData = holidaysNext.map(d => [d, d.toLocaleDateString('de-DE', {weekday: 'long'})]);
    sheet.getRange('K2').setValue(currentYear + 1).setFontWeight('bold');
    sheet.getRange(3, 11, holidaysNextData.length, 2).setValues(holidaysNextData);
  }

  sheet.getRange('H:H').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('K:K').setNumberFormat('yyyy-mm-dd');

  sheet.autoResizeColumns(1, 12);
}

/**
 * Erstellt das 'Ärzte & Stammdaten' Blatt mit der neuen Ärzteliste und zusätzlichen Spalten.
 */
function createDoctorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONSTANTS.DOCTORS_SHEET);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONSTANTS.DOCTORS_SHEET, 1);
  }
  
  const headers = ['Name des Arztes', 'Arbeitszeit (Faktor)', 'Station', 'Macht Dienste', 'Macht Spätdienste', 'Ist Einspringer'];
  
  const doctorsList = [
      'Dwai', 'Lorscheid', 'Katzenbach', 'Scheller', 'Lanz', 'Hahn', 'Ost',
      'Abusaada', 'Altun', 'Figel', 'Deutsch', 'Knab', 'Herzog', 'Jaroni', 'Turner', 'Röthke', 'Salzgeber'
  ];
  
  const exampleData = doctorsList.map(name => [name, 1.0, '', 'JA', 'JA', 'JA']);
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, exampleData.length, exampleData[0].length).setValues(exampleData);
  
  sheet.autoResizeColumns(1, 6);
}

/**
 * Erstellt das 'Einspringliste' Blatt mit dem neuen Layout für das Ranking.
 */
function createJumperListSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONSTANTS.JUMPER_LIST_SHEET);
    if (sheet) {
        sheet.clear();
    } else {
        sheet = ss.insertSheet(CONSTANTS.JUMPER_LIST_SHEET, 2);
    }
    const headers = ['Platz', 'Name', 'Aktueller Punktestand'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.getRange('F1').setValue('Letzte Aktualisierung:').setFontWeight('bold');
    
    sheet.autoResizeColumns(1, 3);
    sheet.autoResizeColumn(6);
}

/**
 * @OnlyCurrentDoc
 *
 * hilfsunktionen.gs: Enthält alle Hilfsfunktionen (Feiertage, Parser, etc.).
 * Version 39: Ausnahme für 24H-Wunsch bei max. Wochenenden entfernt, um strikte Einhaltung zu gewährleisten.
 */

/**
 * Erzeugt einen formatierten Blattnamen im Format "Präfix MM/yy".
 * @param {string} prefix - Das Präfix (z.B. "Wünsche").
 * @param {number} year - Das Jahr.
 * @param {number} month - Der Monat (0-basiert).
 * @return {string} Der formatierte Blattname.
 */
function getFormattedSheetName(prefix, year, month) {
    const monthString = (month + 1).toString().padStart(2, '0');
    const yearString = year.toString().slice(-2);
    return prefix ? `${prefix} ${monthString}/${yearString}` : `${monthString}/${yearString}`;
}

/**
 * Parst das Datum aus einem Blattnamen im Format "MM/yy".
 * @param {string} name - Der Name des Blattes.
 * @return {Date|null} - Das Datum oder null, wenn kein Datum gefunden wurde.
 */
function parseDateFromSheetName(name) {
  const regex = /^(\d{2})\/(\d{2})$/;
  const match = name.match(regex);

  if (match && match.length === 3) {
    const month = parseInt(match[1], 10) - 1;
    const year = parseInt("20" + match[2], 10);
    if (month >= 0 && month <= 11 && year > 2000) {
        return new Date(year, month, 1);
    }
  }
  return null;
}

/**
 * Gibt eine Liste von Feiertagen für ein bestimmtes Jahr zurück (für Rheinland-Pfalz).
 * @param {number} year - Das Jahr, für das die Feiertage abgerufen werden sollen.
 * @return {Array<Date>} Eine sortierte Liste von Date-Objekten für die Feiertage.
 */
function getHolidaysForYear(year) {
  const holidays = [];
  // Feste Feiertage
  holidays.push(new Date(year, 0, 1)); // Neujahr
  holidays.push(new Date(year, 4, 1)); // Tag der Arbeit
  holidays.push(new Date(year, 9, 3)); // Tag der Deutschen Einheit
  holidays.push(new Date(year, 10, 1)); // Allerheiligen (in RP)
  holidays.push(new Date(year, 11, 25)); // 1. Weihnachtstag
  holidays.push(new Date(year, 11, 26)); // 2. Weihnachtstag
  holidays.push(new Date(year, 11, 24)); // Heiligabend
  holidays.push(new Date(year, 11, 31)); // Silvester
  // Bewegliche Feiertage (Ostern, Himmelfahrt, Pfingsten, Fronleichnam)
  // Berechnung des Osterdatums (Algorithmus von Gauß)
  const a = year % 19;
  const b = year % 4;
  const c = year % 7;
  const k = Math.floor(year / 100);
  const p = Math.floor((13 + 8 * k) / 25);
  const q = Math.floor(k / 4);
  const M = (15 - p + k - q) % 30;
  const N = (4 + k - q) % 7;
  const d = (19 * a + M) % 30;
  const e = (2 * b + 4 * c + 6 * d + N) % 7;
  let ostern = new Date(year, 2, 22 + d + e); // Ostersonntag

  // Sonderfälle für Osterdatum
  if ((d === 29 && e === 6) || (d === 28 && e === 6 && (11 * M + 11) % 30 < 19)) {
    ostern.setDate(ostern.getDate() - 7);
  }

  holidays.push(new Date(ostern.getTime() - 2 * 864e5)); // Karfreitag (-2 Tage)
  holidays.push(new Date(ostern.getTime() + 1 * 864e5)); // Ostermontag (+1 Tag)
  holidays.push(new Date(ostern.getTime() + 39 * 864e5)); // Christi Himmelfahrt (+39 Tage)
  holidays.push(new Date(ostern.getTime() + 50 * 864e5)); // Pfingstmontag (+50 Tage)
  holidays.push(new Date(ostern.getTime() + 60 * 864e5)); // Fronleichnam (+60 Tage) (in RP)

  return holidays.sort((a, b) => a.getTime() - b.getTime()); // Sortieren nach Datum
}

/**
 * Standardisierte Fehlerbehandlung.
 * @param {Error} e - Das Fehlerobjekt.
 * @param {string} [message='Ein Fehler ist aufgetreten.'] - Eine benutzerfreundliche Nachricht.
 */
function handleError(e, message = 'Ein Fehler ist aufgetreten.') {
  Logger.log(`${message} Details: ${e.message} \nStack: ${e.stack}`);
  if (!isTriggered()) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Fehler', `${message}\nDetails: ${e.message}`, ui.ButtonSet.OK);
  }
}

function violatesLateShiftFollowUpRule(doctorName, day, dutyType, plan, planningData, stats) {
  const recentHistory = planningData.recentHistory || {};
  const currentEntry = plan[day];
  if (!currentEntry) {
    return false;
  }

  const currentDate = currentEntry.date;
  const previousDayIndex = day - 1;

  const isAllowedFridayToSaturdayNight = previousDate => {
    if (!previousDate) {
      return false;
    }
    return previousDate.getDay() === 5 && currentDate.getDay() === 6 && dutyType === 'nacht';
  };

  if (previousDayIndex >= 1) {
    const previousEntry = plan[previousDayIndex];
    if (previousEntry && previousEntry.late === doctorName) {
      if (!isAllowedFridayToSaturdayNight(previousEntry.date)) {
        return true;
      }
    }
  } else {
    const historyEntry = recentHistory[doctorName];
    if (historyEntry && historyEntry.lastLateShiftDate) {
      const gapDays = diffInDays(currentDate, historyEntry.lastLateShiftDate);
      if (gapDays !== null && gapDays === 1) {
        if (!isAllowedFridayToSaturdayNight(historyEntry.lastLateShiftDate)) {
          return true;
        }
      }
    }
  }

  return false;
}

function recomputeLastMainDutyDate(doctorName, stats, plan, planningData) {
  const assignments = Object.keys(stats[doctorName].assignedMainDuties || {})
    .map(key => parseInt(key, 10))
    .filter(num => !Number.isNaN(num))
    .sort((a, b) => a - b);

  if (assignments.length > 0) {
    const lastDay = assignments[assignments.length - 1];
    stats[doctorName].lastMainDutyDate = plan[lastDay] ? plan[lastDay].date : null;
    stats[doctorName].lastMainDutyType = stats[doctorName].assignedMainDuties[lastDay] || null;
  } else {
    const recentHistory = planningData.recentHistory || {};
    const historyEntry = recentHistory[doctorName] || {};
    stats[doctorName].lastMainDutyDate = historyEntry.lastMainDutyDate || null;
    stats[doctorName].lastMainDutyType = historyEntry.lastMainDutyType || null;
  }
}

function computeMainDutyScore(doc, stats, dutyTargets, planningData) {
  const name = doc.name;
  const assignedMain = stats[name] ? stats[name].mainDuties || 0 : 0;
  const totalAssigned = stats[name] ? stats[name].totalDuties || 0 : 0;
  const target = dutyTargets[name] || 0;
  const factor = doc.factor || 1;
  const defaultMax = (planningData.config['Max. Hauptdienste pro Monat'] || 5) * factor;
  const effectiveTarget = target > 0 ? target : Math.max(1, defaultMax);
  const mainLoad = effectiveTarget > 0 ? assignedMain / effectiveTarget : assignedMain;
  const overTargetPenalty = assignedMain > effectiveTarget ? (assignedMain - effectiveTarget) * 0.25 : 0;
  const totalLoad = effectiveTarget > 0 ? totalAssigned / (effectiveTarget + 0.0001) : totalAssigned;
  const weekendCap = planningData.config['Max. Wochenenddienste pro Monat'] || 2;
  const weekendLoad = weekendCap > 0 ? (stats[name] ? stats[name].weekendsWorked || 0 : 0) / weekendCap : 0;
  return mainLoad + overTargetPenalty + (totalLoad * 0.15) + (weekendLoad * 0.1);
}

function chooseBestDoctor(candidates, stats, dutyTargets, planningData) {
  if (!candidates || candidates.length === 0) {
    return null;
  }
  const scored = candidates.map(doc => ({
    doc,
    score: computeMainDutyScore(doc, stats, dutyTargets, planningData)
  }));
  scored.sort((a, b) => a.score - b.score);
  const bestScore = scored[0].score;
  const tolerance = 0.0001;
  const bestDocs = scored.filter(entry => Math.abs(entry.score - bestScore) <= tolerance).map(entry => entry.doc);
  return shuffleArray(bestDocs)[0];
}

/**
 * Prüft, ob das Skript durch einen Trigger ausgeführt wird.
 * @return {boolean} True, wenn das Skript durch einen Trigger ausgeführt wird, sonst False.
 */
function isTriggered() {
    try {
        SpreadsheetApp.getUi();
        return false;
    } catch (e) {
        return true;
    }
}

/**
 * Wird bei jeder Bearbeitung eines Blattes automatisch ausgeführt.
 * Färbt die Zellen in den Wunsch-Blättern entsprechend der Eingabe.
 * @param {Object} e - Das Event-Objekt, das von Google Sheets übergeben wird.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (!parseDateFromSheetName(sheetName)) {
    return;
  }

  if (range.getColumn() < 4) { // Spalten A, B, C ignorieren
    return;
  }

  const value = (e.value || '').toString().replace(/\s/g, '').toUpperCase();
  let color = null;

  switch (value) {
    case 'W':
      color = '#c9ead5'; // Hellgrün
      break;
    case '24H':
    case 'WW': // Falls jemand noch WW verwendet
      color = '#8fbc8f'; // Dunkleres Grün
      break;
    case 'N':
      color = '#f4c7c3'; // Hellrot
      break;
    case '':
      color = null; // Keine Füllung
      break;
    default:
      color = '#d9d9d9'; // Grau für ungültige Eingaben
      break;
  }
  
  range.setBackground(color);
}

/**
 * Mischt die Elemente eines Arrays zufällig.
 * @param {Array} array - Das zu mischende Array.
 * @return {Array} Das gemischte Array.
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

/**
 * Erzeugt einen eindeutigen Schlüssel für ein Wochenende (Freitag des Wochenendes).
 * @param {Date} date - Ein Datum innerhalb des Wochenendes.
 * @return {string} Der lokalisierte Datumsstring des Freitags des Wochenendes.
 */
function getWeekendKey(date) {
  const d = new Date(date.getTime());
  const dayOfWeek = d.getDay(); // 0 = Sonntag, 1 = Montag, ..., 5 = Freitag, 6 = Samstag
  
  // Gehe zum Freitag des aktuellen Wochenendes
  // Wenn es Sonntag (0) ist, gehe 2 Tage zurück zum Freitag.
  // Wenn es Samstag (6) ist, gehe 1 Tag zurück zum Freitag.
  // Wenn es Freitag (5) ist, bleibe auf dem Freitag.
  // Für Wochentage vor Freitag, gehe zum letzten Freitag.
  const offset = dayOfWeek - 5; // Differenz zu Freitag (5)
  if (dayOfWeek < 5) { // Montag (1) bis Donnerstag (4)
    d.setDate(d.getDate() - (dayOfWeek + 2)); // Gehe zum vorherigen Freitag
  } else { // Freitag (5), Samstag (6), Sonntag (0)
    d.setDate(d.getDate() - offset); // Gehe zum Freitag dieser Woche
  }
  return d.toLocaleDateString();
}

/**
 * Prüft, ob ein Arzt am Vortag einen Hauptdienst hatte (24h, Tag oder Nacht).
 * @param {string} doctorName - Der Name des Arztes.
 * @param {number} currentDay - Der aktuelle Tag im Monat (1-basiert).
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @return {boolean} True, wenn der Arzt am Vortag einen Hauptdienst hatte, sonst False.
 */
function hasMainDutyOnPreviousDay(doctorName, currentDay, plan) {
  const prevDay = parseInt(currentDay) - 1;
  if (prevDay < 1 || !plan[prevDay]) {
    return false;
  }
  const prevDayPlan = plan[prevDay];
  if (prevDayPlan.duty_24h === doctorName ||
    prevDayPlan.duty_tag === doctorName ||
    prevDayPlan.duty_nacht === doctorName) {
    return true;
  }
  return false;
}

/**
 * Hilfsfunktion, um zu prüfen, ob ein Arzt am selben Wochenende noch andere Hauptdienste hat.
 * Wird verwendet, um die weekendWorked-Statistik korrekt zu dekrementieren.
 * @param {string} doctorName - Der Name des Arztes.
 * @param {number} currentDay - Der aktuelle Tag im Monat (1-basiert).
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @return {boolean} True, wenn der Arzt noch andere Hauptdienste am selben Wochenende hat, sonst False.
 */
function isWeekendDutyStillNeeded(doctorName, currentDay, plan) {
    const currentWeekendKey = getWeekendKey(plan[currentDay].date);
    for (let day = 1; day <= Object.keys(plan).length; day++) {
        const dayInfo = plan[day];
        if (dayInfo.isWeekendDay && getWeekendKey(dayInfo.date) === currentWeekendKey) {
            // Prüfe, ob der Arzt an diesem Wochenende noch irgendeinen Hauptdienst hat (außer dem gerade entfernten)
            if (day !== currentDay && (dayInfo.duty_24h === doctorName || dayInfo.duty_tag === doctorName || dayInfo.duty_nacht === doctorName)) {
                return true;
            }
        }
    }
    return false;
}

/**
 * Kategorisiert Ärzte für einen Hauptdienst-Slot in 'unavailable', 'notfalls', 'available'.
 * @param {number} day - Der aktuelle Tag im Monat (1-basiert).
 * @param {string} dutyType - Der Diensttyp ('24h', 'tag', 'nacht').
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @param {Object} planningData - Die Planungsdaten (enthält doctors, availability, wishes, ektPlan).
 * @param {Object} stats - Das Statistik-Objekt für Ärzte.
 * @param {number} maxWeekends - Maximale Wochenenddienste pro Monat aus der Konfiguration.
 * @param {number} maxMainDutiesPerMonth - Maximale Hauptdienste pro Monat aus der Konfiguration.
 * @param {Object} dutyTargets - Die Zieldienste pro Arzt (ungerundet).
 * @return {Object} Ein Objekt mit drei Arrays: unavailable, notfalls, available.
 */
function categorizeDoctorsForMainDuty(day, dutyType, plan, planningData, stats, maxWeekends, maxMainDutiesPerMonth, dutyTargets) {
  const { doctors, availability, wishes } = planningData;
  const dayInfo = plan[day];
  const currentWeekendKey = dayInfo.isWeekendDay ? getWeekendKey(dayInfo.date) : null;

  const unavailable = [];
  const notfalls = [];
  const available = [];

  doctors.filter(doc => doc.doesDuties).forEach(doc => {
    if (!stats[doc.name]) {
      Logger.log(`ERROR: stats[${doc.name}] is UNDEFINED. Skipping this doctor for categorization.`);
      unavailable.push(doc);
      return;
    }
    if (typeof stats[doc.name].mainDuties === 'undefined') {
      unavailable.push(doc);
      return;
    }
    if (typeof dutyTargets[doc.name] === 'undefined') {
      unavailable.push(doc);
      return;
    }

    const currentDate = dayInfo.date;
    const lastMainDutyDate = stats[doc.name].lastMainDutyDate || null;
    if (lastMainDutyDate) {
      const diffSinceLastMainDuty = diffInDays(currentDate, lastMainDutyDate);
      if (diffSinceLastMainDuty !== null && diffSinceLastMainDuty >= 0 && diffSinceLastMainDuty < DUTY_RULES.MAIN_COOLDOWN_DAYS) {
        unavailable.push(doc);
        return;
      }
    }

    if (violatesLateShiftFollowUpRule(doc.name, parseInt(day, 10), dutyType, plan, planningData, stats)) {
      unavailable.push(doc);
      return;
    }

    if (!availability[doc.name] || !availability[doc.name][day] || !availability[doc.name][day][dutyType]) {
      unavailable.push(doc);
      return;
    }

    if (hasMainDutyOnPreviousDay(doc.name, day, plan)) {
      unavailable.push(doc);
      return;
    }

    const nextDay = parseInt(day, 10) + 1;
    if (plan[nextDay] && (plan[nextDay].duty_24h === doc.name || plan[nextDay].duty_tag === doc.name || plan[nextDay].duty_nacht === doc.name)) {
      unavailable.push(doc);
      return;
    }

    let recentDutyCount = 0;
    for (let d = day - 1; d >= Math.max(1, day - 4); d--) {
      if (stats[doc.name].assignedMainDuties[d]) {
        recentDutyCount++;
      }
    }
    for (let d = day + 1; d <= Math.min(Object.keys(plan).length, day + 5); d++) {
      if (stats[doc.name].assignedMainDuties[d]) {
        recentDutyCount++;
      }
    }
    if (recentDutyCount > 0) {
      Logger.log(`DEBUG: Doctor ${doc.name} unavailable for ${dutyType} on day ${day} due to recent duty (count: ${recentDutyCount}).`);
      unavailable.push(doc);
      return;
    }

    if (stats[doc.name].mainDuties >= maxMainDutiesPerMonth) {
      unavailable.push(doc);
      return;
    }

    if (dayInfo.isWeekendDay && stats[doc.name].workedWeekendKeys.has(currentWeekendKey)) {
      const is24HWishForThisDuty = (wishes[doc.name]?.[day]?.[dutyType] || '').includes('24H');
      if (!(dutyType === 'nacht' && is24HWishForThisDuty && plan[day]['duty_tag'] === doc.name)) {
        unavailable.push(doc);
        return;
      }
    }

    const isWeekendShift = dayInfo.isWeekendDay || (dayInfo.date.getDay() === 5 && dutyType === '24h');
    if (isWeekendShift) {
      const weekendKey = getWeekendKey(dayInfo.date);
      if (stats[doc.name].weekendsWorked >= maxWeekends && !stats[doc.name].workedWeekendKeys.has(weekendKey)) {
        unavailable.push(doc);
        return;
      }
    }

    let isNotfalls = false;
    if (stats[doc.name].mainDuties > dutyTargets[doc.name]) {
      isNotfalls = true;
    }
    if (!isNotfalls && stats[doc.name].mainDuties >= dutyTargets[doc.name] - 1) {
      isNotfalls = true;
    }
    if (!isNotfalls && stats[doc.name].mainDuties >= (maxMainDutiesPerMonth * doc.factor) - 1) {
      isNotfalls = true;
    }

    if (isNotfalls) {
      notfalls.push(doc);
    } else {
      available.push(doc);
    }
  });

  return { unavailable, notfalls, available };
}

/**
 * Berechnet einen "Notfalls"-Score für einen Arzt. Höherer Score = "schlechterer" Notfall.
 * @param {string} doctorName - Der Name des Arztes.
 * @param {number} day - Der aktuelle Tag.
 * @param {string} dutyType - Der Diensttyp ('24h', 'tag', 'nacht', 'late').
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @param {Object} planningData - Die Planungsdaten (enthält doctors, wishes, etc.).
 * @param {Object} stats - Das Statistik-Objekt für Ärzte.
 * @param {number} maxMainDutiesPerMonth - Maximale Hauptdienste pro Monat aus der Konfiguration.
 * @param {Object} dutyTargets - Die Zieldienste pro Arzt (kann ungerundet sein).
 * @return {number} Der Notfalls-Score.
 */
function getNotfallsScore(doctorName, day, dutyType, plan, planningData, stats, maxMainDutiesPerMonth, dutyTargets) {
  const doc = planningData.doctors.find(d => d.name === doctorName);
  let score = 0;

  // Defensive Checks
  if (!stats[doctorName] || stats[doctorName].mainDuties === undefined || !dutyTargets[doctorName]) {
    Logger.log(`WARNING: getNotfallsScore called with incomplete stats/targets for ${doctorName}. Returning default score.`);
    return 0; // Hoher Score, um diesen Arzt zu benachteiligen
  }

  // Score für Überschreitung des Zieldienstes
  if (stats[doctorName].mainDuties > dutyTargets[doctorName]) {
    score += 101; // Moderater Malus für Überschreitung des Solls
  }

  // 1. Bereits ein Dienst in den letzten oder nächsten 5 Tagen eingetragen (KEIN SCORE MEHR, da harte Sperre)
  // Die Logik ist jetzt in categorizeDoctorsForMainDuty als harte Sperre implementiert.

  // 2. Noch weniger als 1 Dienst unter der ungerundeten zugeteilten Dienstanzahl
  if (stats[doctorName].mainDuties >= dutyTargets[doctorName] - 1) {
    score += 100;
  }
  // 3. Oder <= 1 Dienst unter der maximalen Dienstanzahl
  if (stats[doctorName].mainDuties >= (maxMainDutiesPerMonth * doc.factor) - 1) {
    score += 50;
  }
  return score;
}

/**
 * Berechnet die Zieldienste pro Arzt basierend auf Arbeitszeitfaktor und verfügbarer Dienstlast.
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @param {Array<Object>} doctors - Liste der Ärzte mit ihren Eigenschaften.
 * @param {Object} jumperPointsMap - Karte der Einspringerpunkte pro Arzt.
 * @return {Object} Ein Objekt, das für jeden Arzt den ungerundeten Zieldienst enthält.
 */
function calculateDutyTargets(plan, doctors, jumperPointsMap) {
  // NEU: totalMainDutiesAvailable basiert auf der tatsächlichen Anzahl der Dienst-Slots
  let totalMainDutiesAvailable = 0;
  for (const day in plan) {
    totalMainDutiesAvailable += plan[day].dutyTypes.length;
  }

  // NEU: totalFactor nur für Ärzte, die Hauptdienste leisten
  const doctorsWhoDoDuties = doctors.filter(doc => doc.doesDuties);
  const totalFactor = doctorsWhoDoDuties.reduce((sum, doc) => sum + doc.factor, 0);
  
  const targets = {}; 

  doctors.forEach(doc => {
    if (doc.doesDuties && totalFactor > 0) { // Nur für Ärzte, die Dienste leisten und wenn totalFactor > 0
      const rawTarget = (totalMainDutiesAvailable * doc.factor) / totalFactor;
      targets[doc.name] = rawTarget; // Store the raw (unrounded) target
    } else {
      targets[doc.name] = 0; // Setze 0 für Ärzte ohne Dienstverpflichtung oder wenn keine Dienste verfügbar sind
    }
  });

  return targets; // Now contains unrounded values
}

/**
 * @OnlyCurrentDoc
 *
 * einspringer-logik.gs: Enthält die Logik zur Berechnung und Aktualisierung der Einspringliste.
 * Version 30: Passt die Logik an das neue, kürzere Blatt-Namensformat an und macht das Auslesen der Plan-Tabellen robuster.
 * Version 31: Berücksichtigt 20 Punkte für Dienste am Freitag.
 */

/**
 * Hauptfunktion zum Aktualisieren der Einspringliste.
 */
function updateJumperList() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Einspringliste wird aktualisiert...', 'Status', -1);
    Logger.log('Starte updateJumperList');

    try {
        const doctorsSheet = ss.getSheetByName(CONSTANTS.DOCTORS_SHEET);
        if (!doctorsSheet) {
            throw new Error(`Blatt "${CONSTANTS.DOCTORS_SHEET}" nicht gefunden.`);
        }
        const doctorsData = doctorsSheet.getRange('A2:F' + doctorsSheet.getLastRow()).getValues();
        
        const jumperDoctors = doctorsData
            .filter(row => (row[5] || '').toString().toUpperCase() === 'JA')
            .map(row => row[0].toString().trim()); // Sicherstellen, dass Namen getrimmt werden

        if (jumperDoctors.length === 0) {
            ss.toast('Keine Ärzte als Einspringer markiert.', 'Hinweis', 5);
            Logger.log('Keine Ärzte als Einspringer markiert.');
            return;
        }

        const configSheet = ss.getSheetByName(CONSTANTS.CONFIG_SHEET);
        if (!configSheet) {
            throw new Error(`Blatt "${CONSTANTS.CONFIG_SHEET}" nicht gefunden.`);
        }
        const configValues = configSheet.getRange('A2:B' + configSheet.getLastRow()).getValues();
        const config = configValues.reduce((obj, row) => {
            if (row[0] && row[1]) obj[row[0]] = row[1];
            return obj;
        }, {});

        const pointsRegular = config['Punkte pro regulärem Dienst'] || 15;
        const pointsJumper = config['Zusatzpunkte pro Einspringdienst'] || 10;
        const pointsFridayDuty = 20; // Fester Wert für Freitagsdienste
        Logger.log(`Konfiguration: Punkte regulär=${pointsRegular}, Punkte Einspringer=${pointsJumper}, Punkte Freitag=${pointsFridayDuty}`);

        const points = {};
        jumperDoctors.forEach(name => { points[name] = 0; });
        Logger.log('Initialisierte Punkte für Einspringer: ' + JSON.stringify(points));

        const allSheets = ss.getSheets();
        const today = new Date();
        today.setMonth(today.getMonth() - 1);
        const threeMonthsAgo = new Date();
        threeMonthsAgo.setMonth(today.getMonth() - 3);
        threeMonthsAgo.setDate(1); // Start am ersten Tag des Monats für Vergleich

        const planSheets = allSheets.filter(sheet => {
            const sheetName = sheet.getName();
            const sheetDate = parseDateFromSheetName(sheetName);
            // Prüft, ob der Name ein gültiges Datum ist und im relevanten Zeitraum liegt.
            const isValidPlanSheet = sheetDate && sheetDate >= threeMonthsAgo && sheetDate < today;
            if (isValidPlanSheet) {
                Logger.log(`Gültiges Plan-Blatt gefunden: ${sheetName}`);
            }
            return isValidPlanSheet;
        });

        if (planSheets.length === 0) {
            ss.toast('Keine relevanten Plan-Blätter der letzten 3 Monate gefunden.', 'Hinweis', 5);
            Logger.log('Keine relevanten Plan-Blätter gefunden.');
            return;
        }

        planSheets.forEach(sheet => {
            Logger.log(`Verarbeite Blatt: ${sheet.getName()}`);
            const data = sheet.getDataRange().getValues();
            
            // Finde die Kopfzeile des Planbereichs (wo "Dienstarzt" steht)
            let planHeaderRowIndex = -1;
            for (let i = 0; i < data.length; i++) {
                const row = data[i];
                if (row.includes('Dienstarzt')) {
                    planHeaderRowIndex = i;
                    break;
                }
            }

            if (planHeaderRowIndex === -1) {
                Logger.log(`Kopfzeile 'Dienstarzt' in Blatt ${sheet.getName()} nicht gefunden. Überspringe Blatt.`);
                return; // Überspringe dieses Blatt, wenn die Kopfzeile nicht gefunden wird
            }

            const header = data[planHeaderRowIndex];
            const dateColIndex = 0; // Datum ist in der ersten Spalte
            const dutyDoctorColIndex = header.indexOf('Dienstarzt');
            const jumperMarkColIndex = header.indexOf('Eingesprungen (mit "E" markieren)');

            if (dutyDoctorColIndex === -1) {
                Logger.log(`Spalte 'Dienstarzt' in Blatt ${sheet.getName()} nicht gefunden. Überspringe Blatt.`);
                return;
            }

            // Beginne mit dem Auslesen der Daten nach der Kopfzeile des Plans
            for (let i = planHeaderRowIndex + 1; i < data.length; i++) {
                const row = data[i];
                const dutyDoctor = row[dutyDoctorColIndex] ? row[dutyDoctorColIndex].toString().trim() : '';
                const dutyDate = row[dateColIndex]; // Datum aus der Zeile abrufen
                const type = jumperMarkColIndex !== -1 ? (row[jumperMarkColIndex] || '').toString().toUpperCase() : '';

                if (dutyDoctor && points.hasOwnProperty(dutyDoctor) && dutyDate instanceof Date) {
                    // Überprüfen, ob es ein Freitag ist (getDay() gibt 5 für Freitag zurück)
                    if (dutyDate.getDay() === 5) { // Freitag
                        points[dutyDoctor] += pointsFridayDuty;
                    } else {
                        points[dutyDoctor] += pointsRegular;
                    }
                    
                    if (type === 'E') {
                        points[dutyDoctor] += pointsJumper;
                    }
                    Logger.log(`Arzt: ${dutyDoctor}, Datum: ${dutyDate.toLocaleDateString()}, Typ: ${type}, Aktuelle Punkte: ${points[dutyDoctor]}`);
                } else if (dutyDoctor && points.hasOwnProperty(dutyDoctor)) {
                    // Fallback für Zeilen ohne gültiges Datum, falls vorhanden
                    points[dutyDoctor] += pointsRegular;
                    if (type === 'E') {
                        points[dutyDoctor] += pointsJumper;
                    }
                    Logger.log(`Arzt: ${dutyDoctor}, Datum: Ungültig, Typ: ${type}, Aktuelle Punkte: ${points[dutyDoctor]}`);
                }
            }
        });

        const jumperSheet = ss.getSheetByName(CONSTANTS.JUMPER_LIST_SHEET);
        if (!jumperSheet) {
            throw new Error(`Blatt "${CONSTANTS.JUMPER_LIST_SHEET}" nicht gefunden.`);
        }
        jumperSheet.getRange(2, 1, jumperSheet.getMaxRows(), 3).clearContent().setBackground(null);
        jumperSheet.getRange('G1').clearContent();

        let outputData = jumperDoctors.map(name => [name, points[name]]);
        
        // Sortiere absteigend nach Punkten (höchster zuerst)
        outputData.sort((a, b) => b[1] - a[1]); 

        const rankedOutputData = outputData.map((row, index) => [index + 1, ...row]); // Platzierung hinzufügen

        if (rankedOutputData.length > 0) {
            const dataRange = jumperSheet.getRange(2, 1, rankedOutputData.length, 3);
            dataRange.setValues(rankedOutputData);
            Logger.log('Einspringerliste geschrieben: ' + JSON.stringify(rankedOutputData));
        } else {
            Logger.log('Keine Einspringerdaten zum Schreiben vorhanden.');
        }
        
        const numRows = rankedOutputData.length;
        if (numRows >= 1) {
            jumperSheet.getRange(2, 1, 1, 3).setBackground('#ffd700'); // Gold (Platz 1)
        }
        if (numRows >= 2) {
            jumperSheet.getRange(3, 1, 1, 3).setBackground('#c0c0c0'); // Silber (Platz 2)
        }
        if (numRows >= 3) {
            jumperSheet.getRange(4, 1, 1, 3).setBackground('#cd7f32'); // Bronze (Platz 3)
        }

        if (numRows > 3) {
            // Die letzten 3 (oder mehr) Plätze in absteigender Reihenfolge der "Schlechtigkeit" einfärben
            for (let i = 0; i < Math.min(numRows - 3, 3); i++) { // Färbe die 3 schlechtesten nach Bronze
                const rowToColor = numRows - i; // numRows ist der letzte Platz
                let color = '';
                if (i === 0) color = '#e57373'; // Schlechtester (letzter Platz)
                else if (i === 1) color = '#ef9a9a'; // Zweitschlechtester
                else if (i === 2) color = '#ffcdd2'; // Drittschlechtester
                jumperSheet.getRange(rowToColor + 1, 1, 1, 3).setBackground(color);
            }
        }
        
        jumperSheet.getRange('F1').setValue('Letzte Aktualisierung:');
        jumperSheet.getRange('G1').setValue(new Date()).setNumberFormat('yyyy-mm-dd HH:mm');

        ss.toast('Einspringliste wurde erfolgreich aktualisiert.', 'Abgeschlossen', 5);
        Logger.log('Einspringliste erfolgreich aktualisiert.');

    } catch (e) {
        handleError(e, 'Fehler beim Aktualisieren der Einspringliste.');
        ss.toast('Aktualisierung fehlgeschlagen.', 'Fehler', 10);
        Logger.log('Fehler in updateJumperList: ' + e.message);
    }
}

/**
 * @OnlyCurrentDoc
 *
 * plantabelle.gs: Enthält alle Funktionen, die für die Erstellung, Formatierung und
 * das Befüllen der monatlichen Plan-Tabellen zuständig sind.
 * Version 37: Behebt den Bug, bei dem bei wiederholter Plangenerierung ein neuer Plan angehängt wurde.
 * Version 39: Zeigt in der Statistik nur die Hauptdienste (ohne Spätdienste) an.
 * Version 41: Fügt Debug-Logs in createPlanLayout hinzu, um das Problem der fehlenden Plantabelle zu diagnostizieren.
 * Version 42: Fügt ungerundeten Soll-Wert zur Dienststatistik hinzu.
 */

/**
 * Erstellt das "Plan"-Blatt NUR mit dem Wunsch-Bereich.
 * @param {number} year - Das Jahr des zu erstellenden Blattes.
 * @param {number} month - Der Monat des zu erstellenden Blattes (0-basiert).
 * @return {boolean} True, wenn ein Blatt erstellt wurde.
 */
function createMonthlyPlanSheet(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getFormattedSheetName(CONSTANTS.PLAN_SHEET_PREFIX, year, month);
  if (ss.getSheetByName(sheetName)) return false;

  const doctorsSheet = ss.getSheetByName(CONSTANTS.DOCTORS_SHEET);
  if (!doctorsSheet) {
    if (!isTriggered()) {
      handleError(new Error('Das Blatt "Ärzte & Stammdaten" muss existieren.'));
    }
    return false;
  }
  const doctorNames = doctorsSheet.getRange('A2:A' + doctorsSheet.getLastRow()).getValues()
    .map(row => row[0]).filter(String);

  const sheet = ss.insertSheet(sheetName, 1);

  // --- WUNSCH-BEREICH ---
  const topHeader = ['', '', ''];
  const subHeader = ['Datum', 'Tag', 'Dienst-Typ'];
  doctorNames.forEach(name => {
    topHeader.push(name, '');
    subHeader.push('Dienst', 'Spät');
  });

  sheet.getRange(1, 1, 1, topHeader.length).setValues([topHeader]);
  sheet.getRange(2, 1, 1, subHeader.length).setValues([subHeader]);
  sheet.getRange('A1:C2').merge().setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  for (let i = 0; i < doctorNames.length; i++) {
    const col = 4 + i * 2;
    sheet.getRange(1, col, 1, 2).merge().setValue(doctorNames[i]).setFontWeight('bold').setHorizontalAlignment('center');
  }

  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const holidays = getHolidaysForYear(year);
  const weekDaysDE = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
  const data = [];
  let mergeRanges = [];

  for (let day = 1; day <= daysInMonth; day++) {
    const currentDate = new Date(year, month, day);
    const dayOfWeek = currentDate.getDay();
    const isHoliday = holidays.some(h => h.getTime() === currentDate.getTime());
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

    const rowTemplate = Array(doctorNames.length * 2).fill('');

    if (isWeekend || isHoliday) {
      data.push([currentDate, weekDaysDE[dayOfWeek], 'Tag', ...rowTemplate]);
      data.push(['', '', 'Nacht', ...rowTemplate]);
      const startRow = data.length + 2;
      mergeRanges.push(`A${startRow - 1}:A${startRow}`);
      mergeRanges.push(`B${startRow - 1}:B${startRow}`);
    } else {
      data.push([currentDate, weekDaysDE[dayOfWeek], '24h-Dienst', ...rowTemplate]);
    }
  }

  if (data.length > 0) {
    sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
    sheet.getRange('A:A').setNumberFormat('dd.mm.yyyy');

    mergeRanges.forEach(range => sheet.getRange(range).mergeVertically().setVerticalAlignment('middle'));

    let currentRow = 3;
    for (let day = 1; day <= daysInMonth; day++) {
      const currentDate = new Date(year, month, day);
      const dayOfWeek = currentDate.getDay();
      const isHoliday = holidays.some(h => h.getTime() === currentDate.getTime());
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

      if (isWeekend || isHoliday) {
        sheet.getRange(currentRow, 1, 2, topHeader.length).setBackground('#fce5cd');
        for (let i = 0; i < doctorNames.length; i++) {
          const lateShiftCol = 5 + i * 2;
          sheet.getRange(currentRow, lateShiftCol, 2, 1).setBackground('#d9d9d9');
        }
        currentRow += 2;
      } else {
        currentRow += 1;
      }
    }
  }

  sheet.setFrozenColumns(3);
  sheet.setFrozenRows(2);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidths(4, doctorNames.length * 2, 80);

  // --- NEUE HINZUFÜGUNG: Plan-Layout direkt nach dem Wunschbereich erstellen ---
  createPlanLayout(sheet);

  return true;
}

/**
 * Stellt sicher, dass der Plan-Bereich UNTER dem Wunsch-Bereich existiert.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Das aktive Tabellenblatt.
 */
function createPlanLayout(sheet) {
    Logger.log(`createPlanLayout aufgerufen für Blatt: ${sheet.getName()}`);
    const finder = sheet.createTextFinder('Dienstarzt');
    const foundCell = finder.findNext();
    if (foundCell) {
        Logger.log(`'Dienstarzt' Header bereits gefunden bei Zelle: ${foundCell.getA1Notation()}. Layout wird nicht neu erstellt.`);
        return; // Layout existiert bereits, nichts tun.
    }
    Logger.log(`'Dienstarzt' Header nicht gefunden. Erstelle neues Plan-Layout.`);
    
    const lastRow = sheet.getRange("A1").getDataRegion().getLastRow();
    Logger.log(`Letzte belegte Zeile im Blatt (Wunschbereich): ${lastRow}`);

    const planStartRow = lastRow + 2;
    const planHeaders = ['Datum', 'Tag', 'Dienst-Typ', 'Dienstarzt', 'Spätdienstarzt', 'Eingesprungen (mit "E" markieren)'];
    sheet.getRange(planStartRow, 1, 1, planHeaders.length).setValues([planHeaders]).setFontWeight('bold');
    Logger.log(`Plan-Header in Zeile ${planStartRow} erstellt.`);
    
    // Die Spalten Datum, Tag, Dienst-Typ aus dem Wunschbereich kopieren
    // Sicherstellen, dass der Bereich für planData korrekt ist, auch wenn der Wunschbereich leer ist.
    const numWishDataRows = sheet.getRange("A1").getDataRegion().getLastRow() - 2; // Anzahl der Datenzeilen im Wunschbereich
    if (numWishDataRows > 0) {
        const planData = sheet.getRange(3, 1, numWishDataRows, 3).getValues();
        sheet.getRange(planStartRow + 1, 1, planData.length, 3).setValues(planData);
        Logger.log(`Datum/Tag/Dienst-Typ Spalten aus Wunschbereich kopiert. ${planData.length} Zeilen.`);
    } else {
        Logger.log('Keine Daten im Wunschbereich gefunden, Datum/Tag/Dienst-Typ Spalten werden nicht kopiert.');
    }
    Logger.log(`[END] createPlanLayout für Blatt ${sheet.getName()} abgeschlossen.`);
}


/**
 * Schreibt den finalen Plan in den Bereich UNTER der Wunschtabelle.
 * Fügt anschließend eine Dienststatistik ein.
 * @param {Object} plan - Das generierte Dienstplan-Objekt.
 * @param {Object} planningData - Die Planungsdaten (enthält doctors, rawDutyTargets).
 * @param {Object} stats - Die Statistikdaten pro Arzt.
 */
function writePlanToSheet(plan, planningData, stats) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    const finder = sheet.createTextFinder('Dienstarzt');
    const headerCell = finder.findNext();
    if (!headerCell) {
        handleError(new Error("Konnte den Start des Plan-Bereichs nicht finden."));
        return;
    }
    
    const planStartRow = headerCell.getRow() + 1;
    const planStartCol = headerCell.getColumn();
    // Die Anzahl der Zeilen im Planbereich entspricht der Anzahl der Tage (oder 2x Tage für WE/Feiertage)
    // Wir müssen die tatsächliche Anzahl der Zeilen des Plans bestimmen.
    let numPlanDataRows = 0;
    for (const day in plan) {
        const entry = plan[day];
        numPlanDataRows += (entry.dutyTypes.includes('24h') ? 1 : 2);
    }

    // Leere nur die Datenzellen des Plans
    sheet.getRange(planStartRow, planStartCol, numPlanDataRows, 3).clearContent();

    const outputData = [];
    let rowIndex = 0;
    for (let day = 1; day <= Object.keys(plan).length; day++) {
        const entry = plan[day];
        if (entry.dutyTypes.includes('24h')) {
            outputData[rowIndex] = [entry.duty_24h || '', entry.late || '', '']; // Letzte Spalte für Einspringer leer lassen
            rowIndex++;
        } else {
            outputData[rowIndex] = [entry.duty_tag || '', '', '']; // Tag-Dienst, Spät leer, Einspringer leer
            outputData[rowIndex + 1] = [entry.duty_nacht || '', '', '']; // Nacht-Dienst, Spät leer, Einspringer leer
            rowIndex += 2;
        }
    }
    sheet.getRange(planStartRow, planStartCol, outputData.length, 3).setValues(outputData); // 3 Spalten schreiben

    // --- Dienststatistik hinzufügen ---
    const statsStartRow = planStartRow + numPlanDataRows + 2; // 2 Zeilen Abstand nach dem Plan
    // NEU: Header für ungerundeten Soll-Wert hinzugefügt
    const statsHeaders = ['Arzt', 'Anzahl Hauptdienste', 'Ungerundeter Soll-Wert']; 
    sheet.getRange(statsStartRow, 1, 1, statsHeaders.length).setValues([statsHeaders]).setFontWeight('bold');

    const statsData = [];
    // Sortiere Ärzte nach Namen für eine konsistente Anzeige
    const sortedDoctors = planningData.doctors.sort((a, b) => a.name.localeCompare(b.name));

    sortedDoctors.forEach(doctor => {
        const doctorName = doctor.name;
        const mainDutiesCount = stats[doctorName] ? stats[doctorName].mainDuties : 0; 
        // NEU: Ungerundeten Soll-Wert hinzufügen, auf 2 Nachkommastellen gerundet
        const rawTargetValue = planningData.rawDutyTargets[doctorName];
        const formattedTargetValue = Number.isFinite(rawTargetValue) ? rawTargetValue.toFixed(2).replace('.', ',') : '0,00';
        statsData.push([doctorName, mainDutiesCount, formattedTargetValue]);
    });

    if (statsData.length > 0) {
        sheet.getRange(statsStartRow + 1, 1, statsData.length, statsData[0].length).setValues(statsData);
        sheet.autoResizeColumns(1, statsData[0].length); // Spalten für Statistik anpassen
    }
}
/**
 * @OnlyCurrentDoc
 *
 * spätdienstlogik.gs: Enthält die Logik für die Verteilung der Spätdienste.
 * Ausgelagert aus planungslogik.gs für bessere Modularität.
 */

/**
 * Verteilt die Spätdienste basierend auf dem Pseudoalgorithmus.
 * @param {Object} plan - Das aktuelle Dienstplan-Objekt.
 * @param {Object} planningData - Die Planungsdaten (doctors, availability, wishes, stats, ektPlan, jumperPointsMap).
 * @returns {Object} Der angepasste Dienstplan.
 */
function distributeLateShifts(plan, planningData) {
  const { doctors, availability, wishes, stats, ektPlan, jumperPointsMap } = planningData;

  // 0. Stellenanteilige Berechnung der Spätdienstanzahl
  const totalLateShiftsAvailable = Object.values(plan).filter(day => day.lateShiftPossible).length;
  // totalFactor nur für Ärzte, die Spätdienste leisten
  const doctorsWhoDoLateShifts = doctors.filter(d => d.doesLateShifts);
  const totalFactor = doctorsWhoDoLateShifts.reduce((sum, doc) => sum + doc.factor, 0);
  
  const lateShiftTargets = {};
  doctors.forEach(doc => {
    // Sollwert nur für Ärzte, die Spätdienste leisten
    if (doc.doesLateShifts && totalFactor > 0) {
      lateShiftTargets[doc.name] = (totalLateShiftsAvailable * doc.factor) / totalFactor; // Ungerundet
    } else {
      lateShiftTargets[doc.name] = 0; // Setze 0 für Ärzte ohne Spätdienstverpflichtung oder wenn keine Spätdienste verfügbar sind
    }
  });

  // Initialisiere Spätdienst-Zähler in Stats
  doctors.forEach(doc => {
    if (!stats[doc.name].totalLateShifts) stats[doc.name].totalLateShifts = 0;
  });

  // 1. Einteilung der Spätdienste in zwei Dienstgruppen
  const allLateShiftSlots = [];
  for (const day in plan) {
    if (plan[day].lateShiftPossible) {
      allLateShiftSlots.push({ day: parseInt(day) });
    }
  }

  const lateShiftGroups = {
    'wishes': [],
    'open': []
  };

  allLateShiftSlots.forEach(slot => {
    const hasWish = doctors.some(doc => (wishes[doc.name]?.[slot.day]?.['late'] || '').includes('W'));
    if (hasWish) {
      lateShiftGroups['wishes'].push(slot);
    } else {
      lateShiftGroups['open'].push(slot);
    }
  });

  // Zufällig mischen
  lateShiftGroups['wishes'] = shuffleArray(lateShiftGroups['wishes']);
  lateShiftGroups['open'] = shuffleArray(lateShiftGroups['open']);

  // Verarbeite Spätdienstgruppen
  ['wishes', 'open'].forEach(groupKey => {
    const currentGroup = lateShiftGroups[groupKey];

    while (currentGroup.length > 0) {
      // 2. Zufälligen Dienst aus der Gruppe ziehen
      const randomIndex = Math.floor(Math.random() * currentGroup.length);
      const slot = currentGroup.splice(randomIndex, 1)[0];
      const day = slot.day;

      // Wenn der Slot bereits besetzt ist, überspringen (sollte nicht passieren, aber zur Sicherheit)
      if (plan[day].late && plan[day].late !== 'UNBESETZT') {
        continue;
      }

      // 2a, 2b, 2c: Ärzte kategorisieren
      const categorizedDoctors = categorizeDoctorsForLateShift(day, plan, planningData, stats, lateShiftTargets, jumperPointsMap);
      const { unavailable, notfalls, available } = categorizedDoctors;

      let chosenDoctor = null;

      if (groupKey === 'wishes') {
        // 2.1 für 'mit wunsch' Dienstgruppe
        const doctorsWithWish = available.filter(doc => (wishes[doc.name]?.[day]?.['late'] || '').includes('W'));
        if (doctorsWithWish.length > 0) {
          chosenDoctor = shuffleArray(doctorsWithWish)[0];
        } else {
          plan[day].late = 'UNBESETZT'; // Wenn 'verfügbar' leer ist: UNBESETZT (kein Fallback auf notfalls)
        }
      } else { // 'open' Dienstgruppe
        // 2.2.1 aus 'verfügbar' wird zwischen den Ärzten gelost
        if (available.length > 0) {
          chosenDoctor = shuffleArray(available)[0];
        } else if (notfalls.length > 0) {
          // 2.2.2 wenn keiner in 'verfügbar' wird aus 'notfalls' der Arzt mit der höchsten Differenz zu 'spätdienst_soll' eingetragen.
          // Sortiere notfalls nach höchster Differenz zu spätdienst_soll (je größer die Differenz, desto weiter unter Soll)
          notfalls.sort((a, b) => {
            const diffA = lateShiftTargets[a.name] - stats[a.name].totalLateShifts;
            const diffB = lateShiftTargets[b.name] - stats[b.name].totalLateShifts;
            if (diffA !== diffB) {
              return diffB - diffA; // Höhere Differenz (weiter unter Soll) zuerst
            }
            return (jumperPointsMap[a.name] || 0) - (jumperPointsMap[b.name] || 0); // Bei Gleichstand: weniger Einspringpunkte
          });
          chosenDoctor = shuffleArray(notfalls)[0]; // Verlosen bei Gleichstand
        } else {
          // Wenn alle Ärzte 'unverfügbar'
          plan[day].late = 'UNBESETZT';
        }
      }

      if (chosenDoctor) {
        plan[day].late = chosenDoctor.name;
        stats[chosenDoctor.name].totalDuties++;
        stats[chosenDoctor.name].totalLateShifts++;
        stats[chosenDoctor.name].lastLateShiftDate = plan[day].date;
      }
    }
  });

  planningData.stats = stats; // Aktualisierte Statistiken speichern
  return plan;
}

/**
 * Kategorisiert Ärzte für einen Spätdienst-Slot in 'unavailable', 'notfalls', 'available'.
 * @param {number} day - Der aktuelle Tag im Monat (1-basiert).
 * @param {Object} plan - Das gesamte Dienstplan-Objekt.
 * @param {Object} planningData - Die Planungsdaten (enthält doctors, availability, ektPlan).
 * @param {Object} stats - Das Statistik-Objekt für Ärzte.
 * @param {Object} lateShiftTargets - Die Zieldienste für Spätdienste pro Arzt (ungerundet).
 * @param {Object} jumperPointsMap - Die Einspringer-Punkte pro Arzt.
 * @return {Object} Ein Objekt mit drei Arrays: unavailable, notfalls, available.
 */
function categorizeDoctorsForLateShift(day, plan, planningData, stats, lateShiftTargets, jumperPointsMap) {
  const { doctors, availability, ektPlan } = planningData;
  const dayInfo = plan[day];

  const unavailable = [];
  const notfalls = [];
  const available = [];

  doctors.filter(doc => doc.doesLateShifts).forEach(doc => {
    // 2a. 'unverfügbar' festlegen
    // NOGO-Wunsch für Spätdienst
    if (!availability[doc.name] || !availability[doc.name][day] || !availability[doc.name][day]['late']) {
      Logger.log(`DEBUG: Doctor ${doc.name} unavailable for late shift on day ${day} due to availability matrix.`);
      unavailable.push(doc);
      return;
    }
    // Hat EKT (bereits in buildAvailabilityMatrix als harte Sperre)
    // Hat Dienst am Vortag (Hauptdienst)
    if (hasMainDutyOnPreviousDay(doc.name, day, plan)) {
      unavailable.push(doc);
      return;
    }
    // Hat Dienst an diesem Tag (Hauptdienst)
    if (plan[day].duty_24h === doc.name || plan[day].duty_tag === doc.name || plan[day].duty_nacht === doc.name) {
      unavailable.push(doc);
      return;
    }

    const nextDayIndex = parseInt(day, 10) + 1;
    const nextDayEntry = plan[nextDayIndex];
    if (nextDayEntry) {
      const nextDate = nextDayEntry.date;
      const currentDate = dayInfo.date;
      const currentDayOfWeek = currentDate.getDay();
      const nextDayOfWeek = nextDate.getDay();
      if (nextDayEntry.duty_24h === doc.name || nextDayEntry.duty_tag === doc.name) {
        unavailable.push(doc);
        return;
      }
      if (nextDayEntry.duty_nacht === doc.name) {
        const allowedCombo = currentDayOfWeek === 5 && nextDayOfWeek === 6;
        if (!allowedCombo) {
          unavailable.push(doc);
          return;
        }
      }
    }
    // Stationskollege hat Dienst (Hauptdienst)
    const dutyDoctorStation = doctors.find(d => d.name === plan[day].duty_24h || d.name === plan[day].duty_tag || d.name === plan[day].duty_nacht)?.station;
    if (dutyDoctorStation && doc.station === dutyDoctorStation) {
      unavailable.push(doc);
      return;
    }
    // Ist bereits für > 'spätdienst_soll' Spätdienste eingetragen
    if (stats[doc.name].totalLateShifts > lateShiftTargets[doc.name]) {
      unavailable.push(doc);
      return;
    }

    // 2b. 'notfalls' festlegen
    let isNotfalls = false;
    // Die Differenz der Zahl von bereits eingetragenen Spätdiensten zu 'spätdienst_soll' beträgt eins oder weniger
    if (stats[doc.name].totalLateShifts >= lateShiftTargets[doc.name] - 1) {
      isNotfalls = true;
    }

    if (isNotfalls) {
      notfalls.push(doc);
    } else {
      available.push(doc);
    }
  });

  return { unavailable, notfalls, available };
}
