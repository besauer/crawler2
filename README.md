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
3. Über das Menü „Stammdaten“ eine Excel-Datei (`.xlsx`) mit den Schulstammdaten hochladen. Erwartete Spalten: Jahr, Schul_ID, Status, Anzahl Außenstellen, RB, KKZ, RKZ, Schulname, Straße, PLZ, Ort, Telefon, Fax, URL, Schüler/-innen insgesamt, Klassen insgesamt sowie die Schulart-Spalten (z. B. „31 – Berufsschulen“, „36 – Berufliche Gymnasien“). Die importierten Daten werden automatisch in der SQLite-Datenbank `crawler_data.db` gespeichert. Die Import-/Export-Schaltflächen in der Kopfzeile stehen auf allen Seiten zur Verfügung und öffnen denselben Dialog zur kompletten Datensicherung.
4. Unter „Einstellungen“ den OpenAI-API-Schlüssel sowie – falls die Funktion „Verwandte & semantische Keywords finden“ genutzt werden soll – den Google-Keyword-Planner-Schlüssel hinterlegen. Zusätzlich können in derselben Ansicht die Zugangsdaten für die Google Custom Search API sowie die Standardparameter für die Datenqualitätsprüfung gepflegt werden (Modell, Temperatur, Confidence-Schwelle, maximale Treffer, Batch-Größe – Standard 5 000 ohne Obergrenze –, Suchsprache/-region sowie Positiv-/Negativlisten für Domains). Für jede Datenqualitäts-Aktion steht ein eigener Prompt zur Verfügung („Prüfung starten“, „Auftrittstyp prüfen“, „Google-Suche (LLM-Auswertung)“) und kann direkt im Formular angepasst werden. Alle Schlüssel und Konfigurationen werden dauerhaft in `crawler_data.db` abgelegt und beim Start automatisch geladen. Im unteren Bereich steht außerdem die Datenverwaltung bereit, um komplette Sicherungen (inkl. API-Schlüssel) zu exportieren oder – nach vorheriger Vorschau – wieder einzuspielen.
5. Im Menüpunkt „Datenqualität“ lassen sich die in den Stammdaten hinterlegten URLs überprüfen. Die Seite zeigt Status, letzte Prüfergebnisse, KI-Begründungen und ggf. vorgeschlagene Alternativ-URLs. Die Prüfung läuft asynchron und bezieht sich ausschließlich auf die markierten Schulen. Neben den Buttons „Alle auswählen“, „Alle angezeigten auswählen“ und „Keine auswählen“ stehen eigene Aktionen bereit: „Prüfung starten“ für die reguläre Analyse, „Auftrittstyp prüfen“ für eine reine Klassifikation als exklusive oder geteilte Webseite (ohne Batch-Limit – alle markierten Datensätze werden berücksichtigt), „Google-Suche (100 Treffer)“ für eine erweiterte Trefferliste pro Schule sowie „„Falsch“-Vorschläge übernehmen“, das alle aktuell als „Falsch“ markierten Datensätze mit Vorschlag automatisch korrigiert (Statuswechsel auf „OK“ plus „Richtiger Wert übernommen“). Ein Abbrechen-Button stoppt laufende Jobs. Sichtbarkeitsfilter (Alle/OK/Falsch/Richtiger Wert) blenden nicht relevante Schulen aus; zusätzlich klassifiziert die KI jede Homepage als „Exklusive Webseite“ oder „Geteilter Auftritt“. Das Label erscheint direkt hinter der URL, kann per Klick angepasst werden (inklusive optionaler Begründung) und steht über einen eigenen Filter (Alle/Exklusiv/Geteilt/Unklar) für die Anzeige zur Verfügung. Der Button „Alle angezeigten auswählen“ berücksichtigt ausschließlich die gefilterte Menge. Während der Lauf aktiv ist, erscheinen Fortschrittsbalken, Meldungen und eine Live-Tabelle, die sämtliche aktuell gefundenen Empfehlungen samt Status, Sicherheit, Auftrittstyp und Begründung auflistet. Vorschläge lassen sich anschließend mit einem Klick übernehmen, verwerfen oder die URL kann manuell korrigiert werden. Die KI erhält für jede Analyse den vollständigen Seitenaufbau (Navigation, Überschriften, Inhaltssegmente, Impressum) und bewertet so auch komplexe Auftritte, ohne die URL überzubewerten. Alle Ergebnisse werden automatisch in `crawler_data.db` gespeichert und lassen sich weiterhin als JSON exportieren. Während der Prüfung wird dabei strikt folgender Ablauf eingehalten:

   - Startseite der Schule aufrufen und den sichtbaren Text inklusive Titel und Meta-Beschreibung erfassen.
   - Navigationspunkte nach einem Impressums-Link durchsuchen.
   - Das Impressum (oder eine gleichwertige Rechts-/Kontaktseite) öffnen und auswerten.
   - Im Impressum gezielt nach dem in den Stammdaten hinterlegten Schulort suchen, um Namensgleichheiten auszuschließen.

6. Im Menüpunkt „Suche anlegen“ besteht die Seite aus drei Abschnitten:
   - **Metadaten** – eine frei benennbare Kategorie (inkl. Kontextmenü zum Anlegen/Löschen weiterer Kategorien), der Suchname sowie eine optionale Beschreibung.
   - **Schulauswahl** – identisch zur Stammdatentabelle lassen sich alle aktiven Schulen filtern. Text- und Zahlenfilter, die Schularten-Checkboxliste sowie die Buttons „Alle auswählen“, „Keine auswählen“ und „Hard Reset (Filter)“ sorgen dafür, dass nur passende Datensätze sichtbar bleiben. Sichtbare Schulen werden automatisch vorselektiert und können einzeln abgewählt werden.
   - **Suche konfigurieren** – ein großes Textfeld für die Google-Stichwörter (ein Begriff pro Zeile), ein frei editierbarer Prompt, Modell- und Temperaturwahl sowie die maximale Trefferanzahl (Standard 200). Sobald das Stichwortfeld verlassen wird, erscheint unterhalb die zugehörige `site:`-Abfrage für die zuerst ausgewählte Schule. Zusätzlich stehen Buttons für „Suche testen“ (liefert die zusammengesetzten Query-Strings), „Probelauf mit erster Schule“ (führt einen kompletten Test inklusive Google-Suche, Seitenabruf und LLM-Bewertung aus, ohne zu speichern) sowie „Suche speichern“ bereit.
7. Im Menü „Suchen verwalten“ werden sämtliche Definitionen, Läufe und Ergebnisse verwaltet:
   - Die linke Spalte listet alle gespeicherten Suchen mit Kategorie, Name und Status sowie die zuletzt gestarteten Läufe. Ein Klick auf einen Eintrag öffnet die Detailansicht.
   - Rechts erscheinen Stammdaten der Suche (Kategorie, Beschreibung, Keywords, Schulmenge, letzter Lauf, letzter Fehler). Über die Button-Gruppe lassen sich Suchen starten, abbrechen, duplizieren, bearbeiten oder löschen. Der Button „Bearbeiten“ öffnet die Definition direkt im Formular „Suche anlegen“ und übernimmt Kategorie, Stichwörter, Prompt, Modell sowie die komplette Schulauswahl.
   - Laufende Jobs zeigen einen Fortschrittsbalken, kombinierte Statusmeldungen, das aktuell geprüfte Stichwort/URL sowie ein Log der letzten Ereignisse. Die Anzeige aktualisiert sich automatisch, bis der Lauf abgeschlossen oder abgebrochen ist.
   - Nach Abschluss werden die Ergebnisse geladen und lassen sich über die Buttons „Suchergebnisse für Schulen/Dimensionen/Stichwörter anzeigen“ öffnen. Die Schulansicht listet pro Treffer alle Metadaten und bewerteteten Dimensionen, die Dimensionsansicht zeigt die Werte 1–5 mit absoluten und relativen Treffern sowie einer Detailansicht pro Wert und die Stichwortansicht aggregiert nach Keyword. Einzelne Resultate können direkt bearbeitet oder gelöscht werden; die Bearbeiten-Ansicht erlaubt das Anpassen der Dimensionen inkl. Sicherheitswert, des Stichworts sowie einer optionalen Notiz.
   - Ergebnisse werden lokal gecacht und stehen auch für Exporte/Importe zur Verfügung.

8. Unter „Debugging“ werden sämtliche Fehlerereignisse gesammelt. Die Ansicht zeigt Zeitpunkt, Kategorie, Fehlermeldung, strukturierte Kontextdaten (z. B. HTTP-Methode, Pfad, IP-Adresse) sowie den vollständigen Stacktrace. Über die Buttons können die Protokolle neu geladen oder als JSON exportiert werden; das Paket wird außerdem in vollständige Daten-Exporte integriert.

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
