# crawler2

Ein lokaler Web-Crawler mit HTML-Oberfläche zum Durchsuchen von Schul-Webseiten nach Stichwörtern.

## Installation

1. Python 3.10 oder neuer installieren.
2. Abhängigkeiten installieren:

   ```bash
   pip install -r requirements.txt
   ```

## Nutzung

1. Anwendung starten:

   ```bash
   flask --app app run
   ```

   Alternativ kann `python app.py` verwendet werden, um den integrierten Entwicklungsserver zu starten.

2. Im Browser `http://127.0.0.1:5000` öffnen.
3. Im Formular die gewünschten Start-URLs (eine URL pro Zeile) und Stichwörter (ein Stichwort pro Zeile) eintragen.
4. Optional die maximale Seitenanzahl pro Start-URL sowie ein Veröffentlichungszeitfenster auswählen. Seiten ohne Datum werden – mit Ausnahme der Startseite – übersprungen, wenn ein Zeitraum gesetzt ist.
5. Formular absenden. Während der Suche zeigt die Oberfläche live an,
   welche Seite aktuell geprüft wird, welche Seiten zuletzt besucht
   wurden, wie viele Start-URLs bereits abgearbeitet sind, wie viele
   Treffer gefunden wurden, wie groß die verbleibende Warteschlange ist
   und zu wie viel Prozent der Crawl abgeschlossen ist. Treffer
   erscheinen fortlaufend in der Tabelle mit Ursprungs- und Ziel-URL.
6. Bei Bedarf kann der laufende Crawl per „Suche abbrechen“ gestoppt oder ein abgeschlossener Lauf über „Suche speichern“ lokal abgelegt werden. Der Button „Alte Suchen anzeigen“ öffnet ein Archiv mit allen gespeicherten Ergebnissen.

## Hinweise

- Es wird ein konservativer User-Agent verwendet und zwischen Anfragen automatisch pausiert.
- Gesperrte Pfade laut `robots.txt` werden nicht besucht.
- Standardmäßig bleibt der Crawler innerhalb der Domain der Start-URL. Diese Einstellung kann im Code angepasst werden.
- Pro Start-URL werden nur die Startseite und deren direkte internen Links durchsucht; weitere Ebenen werden übersprungen.
- Gespeicherte Suchläufe werden als JSON-Datei unter `saved_searches.json` im Projektverzeichnis abgelegt.
