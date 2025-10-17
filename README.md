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
3. Über das Menü „Stammdaten“ eine Excel-Datei (`.xlsx`) mit den Schulstammdaten hochladen. Erwartete Spalten: Schul ID, Träger, Name, Ort, Anzahl Schüler, Anzahl Klassen, Homepage, Aktiv. Die importierten Daten werden lokal in `stammdaten.json` abgelegt; nur als „Aktiv“ markierte Schulen stehen später in der Suche zur Auswahl.
4. Unter „Einstellungen → API-Schlüssel“ optional den OpenAI-API-Schlüssel hinterlegen und testen. Der Schlüssel wird nur im laufenden Prozess gespeichert und versorgt die KI-Funktionen zur Synonymerweiterung sowie zur Bewertung der Suchbegriffe.
5. Zurück auf der Seite „Suche“ die gewünschten Schulen per Checkbox auswählen (Alle/Keine sowie Feldfilter „Wert enthält …“ stehen als Hilfen zur Verfügung) und die Stichwörter (ein Suchbegriff pro Zeile) eintragen.
6. Bei aktivierter Option „Synonyme automatisch erweitern (KI)“ erscheint neben jedem Stichwort der Link „Suchbegriff erweitern“. Ein Klick öffnet ein Dialogfenster mit dem aktuell verwendeten Prompt, der sich vor dem Absenden anpassen lässt. Die vorgeschlagenen Synonyme landen direkt unter dem jeweiligen Begriff, können dort einzeln entfernt oder um frei hinzugefügte Varianten ergänzt werden – nur die verbleibenden Einträge werden später durchsucht.
7. Über den Button „Suchwörter bewerten“ lassen sich die eingetragenen Begriffe von der KI einschätzen. Die Bewertung (Suchwort, geschätzte Güte, alternative Vorschläge) erscheint direkt unter dem Eingabefeld und wird in gespeicherten Suchläufen abgelegt.
8. Optional die maximale Seitenanzahl pro Start-URL, die Anzahl paralleler Seitenabrufe (bis zu 150 gleichzeitig), ob `robots.txt` respektiert wird, sowie ein Veröffentlichungszeitfenster auswählen. Seiten ohne Datum werden – mit Ausnahme der Startseite – übersprungen, wenn ein Zeitraum gesetzt ist.
9. Formular absenden. Während der Suche zeigt die Oberfläche live an,
   welche Seite aktuell geprüft wird, welche Seiten zuletzt besucht
   wurden, wie viele Start-URLs bereits abgearbeitet sind, wie viele
   Treffer gefunden wurden, wie groß die verbleibende Warteschlange ist,
   zu wie viel Prozent der Crawl abgeschlossen ist und welche Begriffe
   inklusive Synonymen durchsucht werden. Treffer erscheinen
   fortlaufend in der Tabelle mit Ursprungs- und Ziel-URL sowie dem jeweils passenden Suchbegriff.
10. Bei Bedarf kann der laufende Crawl per „Suche abbrechen“ gestoppt oder ein abgeschlossener Lauf über „Suche speichern“ lokal abgelegt werden. Der Button „Alte Suchen anzeigen“ öffnet ein Archiv mit allen gespeicherten Ergebnissen.

## Hinweise

- Es wird ein konservativer User-Agent verwendet und zwischen Anfragen automatisch pausiert.
- Gesperrte Pfade laut `robots.txt` werden standardmäßig nicht besucht; über den Schalter „robots.txt respektieren“ kann die Vorgabe bei Bedarf aufgehoben werden.
- Standardmäßig bleibt der Crawler innerhalb der Domain der Start-URL. Diese Einstellung kann im Code angepasst werden.
- Pro Start-URL werden nur die Startseite und deren direkte internen Links durchsucht; weitere Ebenen werden übersprungen.
- Über das Feld „Parallele Seitenabrufe“ lassen sich bis zu 150 Seiten gleichzeitig laden, um die Ausführung zu beschleunigen.
- Gespeicherte Suchläufe werden als JSON-Datei unter `saved_searches.json` im Projektverzeichnis abgelegt.
- Die Schulstammdaten liegen als JSON-Datei unter `stammdaten.json`. Ein erneuter Excel-Import überschreibt die vorhandenen Einträge vollständig.
- Für die KI-gestützte Synonymerweiterung wird ein eigener OpenAI-API-Schlüssel benötigt, der über das Menü „Einstellungen“ gespeichert oder entfernt werden kann. Ohne Schlüssel läuft die Suche mit den eingegebenen Originalbegriffen weiter.
