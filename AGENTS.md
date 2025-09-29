# Beitragshinweise für `dienstplan`

## Allgemeine Richtlinien
- Schreibe neue Kommentare und Dokumentation auf Deutsch, um konsistent mit dem bestehenden Code zu bleiben.
- Bevorzugt Einrückung mit zwei Leerzeichen in JavaScript/Apps-Script-Dateien.
- Verwende konsequent `const` und `let` statt `var`.
- Setze Semikolons am Zeilenende.
- Ergänze neue oder geänderte Funktionen nach Möglichkeit mit kurzen JSDoc-Kommentaren (ebenfalls Deutsch).
- Halte dich an bestehende Datenstrukturen; nimm Änderungen nur vor, wenn sie sorgfältig begründet sind.

## Tests und Validierung
- Dokumentiere im Pull-Request alle manuell ausgeführten Tests oder Skripte. Falls keine Tests existieren, erkläre kurz, wie du die Funktionalität geprüft hast.

## Architektur-Notizen
- Vorschläge zur Aufteilung und Teststrategie für `dienstbeta.gs` stehen in `docs/REFAKTORIERUNG.md`. Bitte dortige Hinweise prüfen, bevor umfangreiche Änderungen am Dienstplan-Algorithmus erfolgen.
- Bei Arbeiten an Google-Apps-Script-Dateien bevorzugt mehrere kleinere `.gs`-Dateien gemäß der empfohlenen Struktur anlegen, damit Menüs, Blattverwaltung und Planungslogik getrennt bleiben.
- Extrahiere neue, reine Hilfsfunktionen nach Möglichkeit in Utility-Dateien, um sie mit Node-basierten Tests abdecken zu können (z.B. via `jest` oder `vitest` in Verbindung mit `gas-local`).

## PR-Beschreibung
- Die PR-Zusammenfassung soll zwei Abschnitte enthalten: "Änderungen" (Stichpunkte) und "Tests" (Stichpunkte mit den ausgeführten Befehlen oder "keine").
