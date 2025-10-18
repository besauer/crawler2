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
3. Über das Menü „Stammdaten“ eine Excel-Datei (`.xlsx`) mit den Schulstammdaten hochladen. Erwartete Spalten: Jahr, Schul_ID, Status, Anzahl Außenstellen, RB, KKZ, RKZ, Schulname, Straße, PLZ, Ort, Telefon, Fax, URL, Schüler/-innen insgesamt, Klassen insgesamt sowie die Schulart-Spalten (z. B. „31 – Berufsschulen“, „36 – Berufliche Gymnasien“). Die importierten Daten werden lokal in `stammdaten.json` abgelegt.
4. Unter „Einstellungen“ den OpenAI-API-Schlüssel sowie die Standardwerte für Prompt, Sprachmodell, Temperatur und maximale Synonymanzahl hinterlegen und bei Bedarf testen. Im selben Menü lassen sich außerdem die Standardwerte für die maximale Seitenanzahl pro Start-URL und die Zahl paralleler Seitenabrufe festlegen. Der Schlüssel wird nur im laufenden Prozess gespeichert und versorgt die KI-Funktionen zur Synonymerweiterung sowie zur Bewertung der Suchbegriffe; die übrigen Felder legen fest, mit welchen Vorgaben der Synonym-Dialog beziehungsweise das Suchformular vorbefüllt wird.
5. Zurück auf der Seite „Suche“ die gewünschten Schulen per Checkbox auswählen. Für die Stammdaten stehen neben „Alle/Keine“ separate Filterfelder (Text/Zahl) und eine Schularten-Liste bereit: Aktivierte Schularten müssen in den Stammdaten mit „true“ hinterlegt sein, anderenfalls werden die Datensätze ausgeblendet. Sobald ein Filter angewendet wird, blendet die Tabelle automatisch alle nicht passenden Schulen aus und markiert die verbleibenden Datensätze vor. Stichwörter werden im Feld „Suchstichwörter“ über „+ Hinzufügen“ eingetragen; jeder Begriff erscheint anschließend als Chip in der Verwaltungsliste.
6. Ein Klick auf einen Chip öffnet ein Quick-Action-Menü mit „＋ Synonyme finden“ und „✕ Löschen“. „Synonyme finden“ startet den Dialog: Prompt, Modell, Temperatur und Maximalanzahl lassen sich vor jedem Abruf anpassen; bereits verwendete Kombinationen werden lokal zwischengespeichert und bei erneuten Anfragen sofort vorgeschlagen. Die geladenen Synonyme erscheinen als Checkbox-Liste, können gefiltert, gesammelt markiert oder abgewählt werden und werden nach „Hinzufügen“ automatisch dem gewählten Stichwort zugeordnet (gleichzeitig lassen sich weitere Vorschläge über das Menü eines Synonym-Chips abrufen). Nur die angezeigten Chips werden später durchsucht; einzelne Einträge lassen sich jederzeit entfernen.
7. Über den Button „Suchwörter bewerten“ lassen sich die eingetragenen Begriffe von der KI einschätzen. Die Bewertung (Suchwort, geschätzte Güte, alternative Vorschläge) erscheint direkt unter dem Eingabefeld und wird in gespeicherten Suchläufen abgelegt.
8. Optional die maximale Seitenanzahl pro Start-URL, die Anzahl paralleler Seitenabrufe (bis zu 150 gleichzeitig, Standardwerte siehe Einstellungen), ob `robots.txt` respektiert wird, sowie ein Veröffentlichungszeitfenster auswählen. Seiten ohne Datum werden – mit Ausnahme der Startseite – übersprungen, wenn ein Zeitraum gesetzt ist.
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
- Die Standardkonfiguration für Synonym-Prompt, Modell, Temperatur, Maximalanzahl sowie die Crawl-Standardwerte (maximale Seiten pro Start-URL, parallele Abrufe) wird in `settings.json` abgelegt und kann sowohl über die Oberfläche als auch – bei Bedarf – direkt in der Datei angepasst werden.
- Abrufe der Synonymerweiterung werden in `synonym_cache.json` zwischengespeichert und bei identischem Begriff, Prompt, Modell und Temperatur wiederverwendet, um API-Kontingente zu schonen.
