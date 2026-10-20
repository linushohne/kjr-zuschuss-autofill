# KJR Zuschuss Autofill

Excel auswählen, Teilnehmende markieren, Fahrtbeginn eintragen, Word-Liste herunterladen.

Die Anwendung läuft auf [GitHub Pages](https://zuschuss.linushoehne.de/). Keine Anmeldung, kein Backend und kein Build-Schritt.

## Funktionen

- Excel-Import (`.xlsx` und `.xls`), ausschließlich aus dem Blatt „Stamm“.
- Erkennung der Spaltenüberschriften, auch bei unterschiedlicher Spaltenreihenfolge.
- Suche, Filter nach Sippe und Stufe sowie Auswahl aller angezeigten Personen.
- Anzeige nach Name, Sippe, Alter oder Adresse sortieren: Spaltenüberschrift anklicken, erneut klicken für die Gegenrichtung. Der Word-Export bleibt alphabetisch.
- Inaktive Mitglieder sind standardmäßig ausgeblendet. Unbekannte oder leere Aktiv-Werte gelten als aktiv.
- Fehlende Namen, Geburtstage und Adressen lassen sich für den aktuellen Export bearbeiten. Die Quelldatei bleibt unverändert.
- Alter zum verpflichtenden Fahrtbeginn; Titel, Ort und Fahrtende sind optional.
- Umschaltung zwischen Teilnehmer- und Betreuerliste. Betreuer werden ab 18 Jahren zum Fahrtbeginn angezeigt; ein Listenwechsel leert die Auswahl.
- Alphabetische Reihenfolge, fortlaufende Nummerierung und „X“ im Stadt-Kästchen.
- Word-Export anhand der Originalvorlage, mit erhaltenen `FORMTEXT`-Feldern und leeren Unterschriftenfeldern.
- Eigene Betreuer-Vorlage; Juleica-Nummer, Gültigkeit und Führungszeugnis-ID bleiben zum nachträglichen Ausfüllen leer.
- Automatische Seitenumbrüche mit wiederholtem Tabellenkopf; Teilnehmerzeilen werden möglichst nicht geteilt.
- Mindestens zwei Seiten mit leeren Formularzeilen unter den eingetragenen Personen.

Andere Tabellenblätter werden nicht übernommen. Fehlt „Stamm“, wird der Import abgebrochen. Exakt doppelte ausgewählte Datensätze blockieren den Export. Fehlende Angaben müssen vor dem Export ergänzt werden.

## Datenschutz

Datei und Personendaten werden ausschließlich im Browser verarbeitet. Nur Name, Geburtstag, Adresse, Sippe, Stufe und Aktiv-Status werden aus der gelesenen Arbeitsmappe übernommen. Es gibt keine Uploads, Analyse-Dienste, externen Laufzeit-Abhängigkeiten oder Speicherung in Local Storage. Zurücksetzen oder Neuladen verwirft die Daten aus der Anwendung. Bereits heruntergeladene Dateien bleiben auf dem Gerät.

GitHub liefert die statischen Anwendungsdateien aus und erhält dabei die üblichen Seitenaufrufe, nicht aber den Inhalt der ausgewählten Excel-Datei. Beim Export wird ausschließlich die leere Vorlage vom selben Server geladen.

Echte Mitgliederlisten und ausgefüllte Dokumente gehören niemals in dieses Repository. Die `.gitignore` schließt typische Datendateien aus; vor jedem Commit trotzdem prüfen. Tests verwenden künstliche Daten.

## Entwicklung

Voraussetzungen: Node.js 24 oder neuer und Python 3 für den lokalen Webserver.

```sh
npm ci
npm test
npm run dev
```

Anschließend `http://127.0.0.1:4173` öffnen. Die Anwendung muss über HTTP laufen, nicht als `file://`-Datei, da sie JavaScript-Module und die Word-Vorlage lädt.

`docs/` enthält die vollständige statische Anwendung. Änderungen auf `main` werden durch GitHub Pages veröffentlicht. Die Tests prüfen Import, Datumssysteme, Validierung, XML-Escaping, Formularfelder, Stadt-Zuordnung und längere Listen. Word-Testausgaben werden nur unter `/tmp/kjr-qa/` erzeugt.

## Abhängigkeiten und Vorlage

- SheetJS CE 0.20.3 (`docs/vendor/xlsx.mjs`, Apache-2.0).
- fflate 0.8.3 (`docs/vendor/fflate.js`, MIT).
- Lizenztexte liegen unter `docs/vendor/`. Beide Bibliotheken werden lokal ausgeliefert.
- `npm run vendor` erneuert die gebündelten Bibliotheken aus den festgelegten Versionen.
- `docs/teilnehmerliste.docx` und `docs/betreuerliste.docx` sind die leeren, bereitgestellten Originalvorlagen. Autorenmetadaten und externe Verknüpfungen wurden entfernt. Logos, Formatierung und Formularfelder bleiben erhalten.

Eine Ersatzvorlage kann mit `node scripts/prepare-template.mjs /pfad/zur/leeren-vorlage.docx docs/teilnehmerliste.docx` bereinigt werden. Das zweite Argument bestimmt die Zieldatei. Vor Veröffentlichung Inhalt und Metadaten prüfen; eine andere Formularstruktur erfordert eine Anpassung von `docs/word.js` und neue Layoutprüfungen.
