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
3. Über das Menü „Stammdaten“ eine Excel-Datei (`.xlsx`) mit den Schulstammdaten hochladen. Erwartete Spalten: Jahr, Schul_ID, Status, Anzahl Außenstellen, RB, KKZ, RKZ, Schulname, Straße, PLZ, Ort, Telefon, Fax, URL, Schüler/-innen insgesamt, Klassen insgesamt sowie die Schulart-Spalten (z. B. „31 – Berufsschulen“, „36 – Berufliche Gymnasien“). Die importierten Daten werden automatisch in der SQLite-Datenbank `crawler_data.db` gespeichert.
4. Unter „Einstellungen“ den OpenAI-API-Schlüssel sowie – falls die Funktion „Verwandte & semantische Keywords finden“ genutzt werden soll – den Google-Keyword-Planner-Schlüssel hinterlegen. Zusätzlich können in derselben Ansicht die Zugangsdaten für die Google Custom Search API sowie die Standardparameter für die Datenqualitätsprüfung gepflegt werden (Modell, Temperatur, Confidence-Schwelle, maximale Treffer, Batch-Größe, Suchsprache/-region, Schwelle & Batchgröße der automatischen Korrektur, optionales Rate-Limit sowie Positiv-/Negativlisten für Domains). Alle Schlüssel und Konfigurationen werden dauerhaft in `crawler_data.db` abgelegt und beim Start automatisch geladen. Im unteren Bereich steht außerdem die Datenverwaltung bereit, um komplette Sicherungen (inkl. API-Schlüssel) zu exportieren oder – nach vorheriger Vorschau – wieder einzuspielen.
5. Im Menüpunkt „Datenqualität“ lassen sich die in den Stammdaten hinterlegten URLs überprüfen. Die Seite zeigt Status, letzte Prüfergebnisse, KI-Begründungen und ggf. vorgeschlagene Alternativ-URLs. Über den Aktionsbereich („Prüfung“ / „Automatische Korrektur“) kann festgelegt werden, ob nur eine Prüfung, nur die automatische Korrektur oder beides gleichzeitig für die aktuelle Auswahl gestartet werden soll; beide Prozesse laufen asynchron und beziehen sich ausschließlich auf die markierten Schulen – unabhängig vom bisherigen Status. Der Fortschritt erscheint als Live-Anzeige mit Fortschrittsbalken, Meldungen und Detail-Panel. Vorschläge lassen sich mit einem Klick übernehmen oder verwerfen, einzelne Datensätze können manuell korrigiert oder für eine erneute Prüfung vorgemerkt werden. Zusätzlich steht ein Bereich „Automatische Korrektur“ bereit: Ein Klick auf „Korrektur starten“ nutzt ebenfalls ausschließlich die aktuelle Auswahl, führt (falls nötig) Google-Suche und LLM-Bewertung aus und listet anschließend nur tatsächliche Änderungen im Format `<Schulname | URL_alt → URL_neu | Begründung>` inklusive Exportmöglichkeit. Alle Ergebnisse werden automatisch in `crawler_data.db` gespeichert und lassen sich weiterhin als JSON exportieren. Während der Prüfung wird dabei strikt folgender Ablauf eingehalten:

   - Startseite der Schule aufrufen und den sichtbaren Text inklusive Titel und Meta-Beschreibung erfassen.
   - Navigationspunkte nach einem Impressums-Link durchsuchen.
   - Das Impressum (oder eine gleichwertige Rechts-/Kontaktseite) öffnen und auswerten.
   - Im Impressum gezielt nach dem in den Stammdaten hinterlegten Schulort suchen, um Namensgleichheiten auszuschließen.

6. Zurück auf der Seite „Suche“ die gewünschten Schulen per Checkbox auswählen. Für die Stammdaten stehen neben „Alle/Keine“ separate Filterfelder (Text/Zahl) und eine Schularten-Liste bereit: Aktivierte Schularten müssen in den Stammdaten mit „true“ hinterlegt sein, anderenfalls werden die Datensätze ausgeblendet. Sobald ein Filter angewendet wird, blendet die Tabelle automatisch alle nicht passenden Schulen aus und markiert die verbleibenden Datensätze vor. Stichwörter werden im Feld „Suchstichwörter“ über „+ Hinzufügen“ eingetragen; jeder Begriff erscheint anschließend als Chip in der Verwaltungsliste.
7. Ein Klick auf einen Chip öffnet ein Quick-Action-Menü mit „＋ Synonyme finden“, „🆕 Verwandte & semantische Keywords finden“ und „✕ Löschen“. Beide Dialoge können vor jedem Abruf angepasst werden (Prompt, Modell, Temperatur, Maximalanzahl); bereits verwendete Kombinationen werden lokal gecacht und bei erneutem Aufruf vorgeschlagen. Die geladenen Vorschläge erscheinen jeweils als filterbare Checkbox-Liste, können gesammelt markiert oder abgewählt werden und werden nach „Hinzufügen“ automatisch dem gewählten Stichwort zugeordnet. Verwandte Keywords werden zunächst über den Google Keyword Planner gesammelt und anschließend von einem LLM gruppiert; die Ergebnisliste zeigt Cluster, Suchvolumen, Wettbewerb und Relevanzwerte an. Nur die angezeigten Chips werden später durchsucht; einzelne Einträge lassen sich jederzeit entfernen.
8. Über den Button „Suchwörter bewerten“ lassen sich die eingetragenen Begriffe von der KI einschätzen. Die Bewertung (Suchwort, geschätzte Güte, alternative Vorschläge) erscheint direkt unter dem Eingabefeld und wird in gespeicherten Suchläufen abgelegt.
9. Optional die maximale Seitenanzahl pro Start-URL, die Anzahl paralleler Seitenabrufe (bis zu 150 gleichzeitig, Standardwerte siehe Einstellungen), ob `robots.txt` respektiert wird, sowie ein Veröffentlichungszeitfenster auswählen. Seiten ohne Datum werden – mit Ausnahme der Startseite – übersprungen, wenn ein Zeitraum gesetzt ist.
10. Formular absenden. Während der Suche zeigt die Oberfläche live an,
   welche Seite aktuell geprüft wird, welche Seiten zuletzt besucht
   wurden, wie viele Start-URLs bereits abgearbeitet sind, wie viele
   Treffer gefunden wurden, wie groß die verbleibende Warteschlange ist,
   zu wie viel Prozent der Crawl abgeschlossen ist und welche Begriffe
   inklusive Synonymen durchsucht werden. Treffer erscheinen
   fortlaufend in der Tabelle mit Ursprungs- und Ziel-URL, dem jeweils passenden Suchbegriff
   sowie einem automatisch erzeugten Kontextfenster (20 Wörter davor und danach).
11. Bei Bedarf kann der laufende Crawl per „Suche abbrechen“ gestoppt oder ein abgeschlossener Lauf über „Suche speichern“ lokal abgelegt werden. Der Button „Alte Suchen anzeigen“ öffnet ein Archiv mit allen gespeicherten Ergebnissen.

## Hinweise

- Es wird ein konservativer User-Agent verwendet und zwischen Anfragen automatisch pausiert.
- Gesperrte Pfade laut `robots.txt` werden standardmäßig nicht besucht; über den Schalter „robots.txt respektieren“ kann die Vorgabe bei Bedarf aufgehoben werden.
- Standardmäßig bleibt der Crawler innerhalb der Domain der Start-URL. Diese Einstellung kann im Code angepasst werden.
- Pro Start-URL werden nur die Startseite und deren direkte internen Links durchsucht; weitere Ebenen werden übersprungen.
- Wiederkehrende Seitenteile (z. B. Navigationsleisten oder Footer) werden nach der ersten Begegnung automatisch
  aus der Stichwortsuche herausgefiltert, damit identische Strukturelemente keine mehrfachen Treffer erzeugen.
- Für jede Fundstelle wird ein Kontextfenster mit 20 Wörtern vor und nach dem Treffer gespeichert. Identische
  Wortfolgen auf derselben Seite werden ignoriert, damit Menü- oder Footer-Wiederholungen nicht mehrfach gemeldet werden.
- Über das Feld „Parallele Seitenabrufe“ lassen sich bis zu 150 Seiten gleichzeitig laden, um die Ausführung zu beschleunigen.
- Sämtliche Anwendungsdaten – einschließlich Stammdaten, Einstellungen, gespeicherter Suchläufe, Datenqualitätsstände, Keyword-/Synonym-Caches und API-Schlüssel – werden transparent in der lokalen SQLite-Datenbank `crawler_data.db` verwaltet.
- Für die KI-gestützte Synonymerweiterung wird ein eigener OpenAI-API-Schlüssel benötigt, der über das Menü „Einstellungen“ gespeichert oder entfernt werden kann. Ohne Schlüssel läuft die Suche mit den eingegebenen Originalbegriffen weiter.
- Die Standardkonfiguration für Synonym-Prompt, Modell, Temperatur, Maximalanzahl sowie die Crawl-Standardwerte (maximale Seiten pro Start-URL, parallele Abrufe) lässt sich jederzeit im Einstellungsdialog anpassen; Änderungen werden unmittelbar in der Datenbank abgelegt.
- Synonym- und Keyword-Ergebnisse werden automatisch im Datenbank-Cache gespeichert und bei identischem Parameter-Set wiederverwendet, um API-Kontingente zu schonen.
- Alle zentralen Daten werden zusätzlich alle fünf Minuten als JSON-Snapshot unter `backups/auto_backup.json` gesichert. Beim Start des Servers werden fehlende Inhalte aus diesem Backup wiederhergestellt.
- Über „Einstellungen → Datenverwaltung“ lassen sich vollständige Exporte erstellen. Die Exportdatei enthält – je nach Auswahl – Stammdaten, gespeicherte Suchen, Einstellungen, Datenqualitätsstände, Cache-Dateien sowie API-Schlüssel. Beim Import können die Bereiche ersetzt, zusammengeführt oder ignoriert werden; eine Dry-Run-Option zeigt vorab an, welche Datensätze neu angelegt oder aktualisiert würden.
