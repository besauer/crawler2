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
4. Optional die maximale Seitenanzahl pro Start-URL anpassen.
5. Formular absenden. Die Anwendung crawlt jede angegebene Domain systematisch, respektiert `robots.txt` und listet für jedes Stichwort-Paar die Ursprungs- und Ziel-URL auf.

## Hinweise

- Es wird ein konservativer User-Agent verwendet und zwischen Anfragen automatisch pausiert.
- Gesperrte Pfade laut `robots.txt` werden nicht besucht.
- Standardmäßig bleibt der Crawler innerhalb der Domain der Start-URL. Diese Einstellung kann im Code angepasst werden.
