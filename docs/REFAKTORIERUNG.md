# Refaktorierungsempfehlungen für `dienstbeta.gs`

## Ausgangslage
- Die Datei `dienstbeta.gs` umfasst weit über 2000 Zeilen und vereint Menüsteuerung, Datenaufbereitung, Trigger-Verwaltung und die eigentliche Planungslogik in einer Datei.
- Unterschiedliche Verantwortlichkeiten teilen sich globale Zustände (z.B. `SpreadsheetApp`, Konfigurationskonstanten, Statusobjekte der Planberechnung), wodurch Änderungen schwer nachvollziehbar werden.
- Reine Hilfsfunktionen wie `diffInDays` oder `computeRelativeDayIndex` stehen unmittelbar neben Trigger-Funktionen wie `onOpen` sowie komplexen Algorithmen wie `distributeMainDuties` oder `balanceMainDuties`.

## Sinnvolle Aufteilung in mehrere Dateien
1. **Grundlagen & Utilities**
   - `constants.gs`: Enthält `CONSTANTS`, `DUTY_RULES`, `MS_PER_DAY` sowie zukünftig weitere Konfigurationsobjekte.
   - `date-utils.gs`: Kapselt Datumshilfen (`normalizeDate`, `diffInDays`, `computeRelativeDayIndex`, `getPreviousYearMonth`).
   - `logging.gs` (optional): Gemeinsame Fehlerbehandlung (`handleError`) und wiederkehrende Logik wie `isTriggered`.

2. **Benutzeroberfläche & Setup**
   - `ui.gs`: Menüaufbau (`onOpen`) sowie `onEdit`, damit UI-Ereignisse klar getrennt bleiben.
   - `setup.gs`: `initialSetup`, `createConfigSheet`, `createDoctorsSheet`, `createJumperListSheet` und weitere Funktionen, die einmalige Tabellenstrukturen erzeugen.

3. **Blattverwaltung & Automatisierung**
   - `sheet-lifecycle.gs`: `checkAndCreateFutureSheets`, `createPastPlanSheet`, `createMonthlyPlanSheet`, `cleanupOldSheets`, `arrangeSheets`, `createPlanLayout`, `createSpreadsheetBackup`.
   - `automation.gs`: Trigger (`setupAutomationTriggers`, `deleteAllTriggers`, `generateFutureMonthlyPlan`) und Hilfsfunktionen wie `isTriggered`.

4. **Planungsdaten & Algorithmen**
   - `planning/data-loader.gs`: `getPlanningData`, `loadRecentHistoryFromPreviousMonth`, `calculateDutyTargets`, `buildAvailabilityMatrix`, `getHolidaysForYear`, `updateJumperList`.
   - `planning/main-duties.gs`: `distributeMainDuties` plus zugehörige Helfer (`categorizeDoctorsForMainDuty`, `updateStatsForNewAssignment`, `updateStatsForSwap`, `recomputeLastMainDutyDate`, `chooseBestDoctor`, etc.).
   - `planning/late-duties.gs`: `distributeLateShifts`, `categorizeDoctorsForLateShift`, `violatesLateShiftFollowUpRule`.
   - `planning/output.gs`: `writePlanToSheet` und Hilfsfunktionen, die ausschließlich für die Ausgabe zuständig sind.

Durch diese Struktur verbleiben pro Datei überschaubare Verantwortlichkeiten; zudem lassen sich reine Datenfunktionen deutlich leichter testen.

## Strategien zur Testbarkeit
- **Extrahiere pure Funktionen**: Funktionen wie `calculateDutyTargets`, `categorizeDoctorsForMainDuty` oder `buildAvailabilityMatrix` können nach der Aufteilung ohne direkte Abhängigkeit zu Apps Script in Node-Tests geprüft werden.
- **Lokale Testumgebung**: Richte ein `package.json` ein und nutze `clasp` plus `@types/google-apps-script`, um die Skripte lokal zu synchronisieren. Für Tests eignen sich `jest` oder `vitest` in Verbindung mit Bibliotheken wie [`gas-local`](https://github.com/google/clasp/tree/master/packages/gas-local) oder [`google-apps-script-mock`](https://github.com/PopGoesTheWza/google-apps-script-mock), um `SpreadsheetApp` & Co. zu simulieren.
- **Testschwerpunkte**:
  - Datumshilfsfunktionen (korrekte Normalisierung, Monatswechsel).
  - Berechnung der Soll-Dienstzahlen (`calculateDutyTargets`).
  - Auswahlheuristiken (`chooseBestDoctor`, `categorizeDoctorsForLateShift`).
  - Verarbeitung der Einspringerliste (`updateJumperList`).
- **Integrationstests**: Erstelle Sample-Daten (kleines Spreadsheet in JSON-Form) und speise sie in `getPlanningData` bzw. den Planungsfluss, während `SpreadsheetApp` durch Mocks ersetzt wird. So lässt sich die gesamte Kette bis `writePlanToSheet` prüfen, ohne Google Drive aufzurufen.

## Isolierung der Funktionalität
- Führe ein `SpreadsheetService`-Interface ein, das nur die benötigten Methoden (`getSheetByName`, `toast`, `getUi`, …) anbietet. Für Tests kann eine einfache In-Memory-Implementierung verwendet werden.
- Lass Datenfunktionen (`getPlanningData`, `loadRecentHistoryFromPreviousMonth`) keine `SpreadsheetApp`-Objekte akzeptieren, sondern das neue Interface. Dadurch bleiben sie unabhängig von der Laufzeitumgebung.
- Verwende klar definierte Datenstrukturen (z.B. Typdefinitionen in JSDoc) für `planningData`, damit beim Refactoring schneller sichtbar wird, welche Felder tatsächlich genutzt werden.
- Reduziere globale Mutationen: Übergib `planningData.stats` explizit an Funktionen, anstatt es implizit in `planningData` zu verändern. So lassen sich Tests gezielt auf bestimmte Ausschnitte beschränken.

## Nächste Schritte
- Schrittweise Migration: Beginne mit den Hilfsmodulen (`constants.gs`, `date-utils.gs`), da deren Entkopplung das geringste Risiko birgt.
- Nach jeder Aufteilung manuelle Rauchtests (Plan generieren, Trigger-Setup) durchführen.
- Ergänze neue Tests zeitnah, damit spätere größere Umbauten am Planungsalgorithmus auf einer abgesicherten Basis stattfinden.
