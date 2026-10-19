# KJR Zuschuss Autofill

Browser-Tool zum Erstellen von Teilnehmerlisten für Pfadfinderfahrten aus einer vorhandenen Excel-Mitgliederliste.

## Status

Das Repository enthält zunächst das Projektgerüst und eine GitHub-Pages-Startseite. Excel-Import, Teilnehmerauswahl und Word-Export sind noch nicht implementiert.

## Geplanter Ablauf

1. Excel-Datei lokal im Browser auswählen.
2. Das Blatt „Stamm“ als Standardquelle einlesen und Spalten anhand ihrer Überschriften erkennen.
3. Kinder mit Suche und Filtern nach Sippe und Stufe auswählen.
4. Das verpflichtende Startdatum der Fahrt angeben.
5. Angaben prüfen und eine ausgefüllte Word-Datei herunterladen.

## Fachliche Anforderungen

- Alter aus Geburtstag und Startdatum berechnen, nicht aus der vorhandenen Excel-Altersspalte.
- Titel, Veranstaltungsort und Enddatum sind optional und können später im Dokument ergänzt werden.
- Im Landkreis-Feld immer „Stadt“ eintragen.
- Originalvorlage und klassische Word-Formularfelder (`FORMTEXT`) möglichst erhalten und direkt befüllen.
- Unterschriftenfelder frei lassen.
- Fehlende Geburtstage oder Adressen vor dem Export anzeigen.
- Zusätzliche Seiten für größere Gruppen und lange Adressen berücksichtigen und testen.
- Keine ungeprüfte Zusammenführung von Gesamt- und Gruppenblättern; mögliche Dubletten vermeiden.

## Datenschutz

Die geplante Anwendung verarbeitet die ausgewählte Excel ausschließlich lokal im Browser. Es gibt kein Backend, kein Tracking und keine dauerhafte Speicherung personenbezogener Daten. Nur für die Liste benötigte Felder sollen übernommen werden; Gesundheitsdaten und Notfallkontakte sind nicht erforderlich.

Echte Mitgliederlisten, ausgefüllte Dokumente und andere personenbezogene Testdaten gehören niemals in dieses Repository. Die `.gitignore` blockiert typische Daten- und Dokumentdateien vorsorglich. Dies ersetzt keine Prüfung vor jedem Commit. Entwicklung und Tests erfolgen mit künstlichen Daten. Eine später benötigte leere Formularvorlage muss vor Veröffentlichung einschließlich Metadaten geprüft werden.

## Hosting

GitHub Pages veröffentlicht den Ordner `docs` aus dem Branch `main`. Für dieses statische Grundgerüst ist kein eigener Build-Schritt nötig.

## Lokal ansehen

Die Datei `docs/index.html` direkt im Browser öffnen. Die Startseite lädt keine externen Skripte, Schriftarten oder Stylesheets.
