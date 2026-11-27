# COMEX Silver Inventory Report Notificator

Sendet neue offizielle COMEX-Silberbestandsreports als ntfy-Push. Die
Meldungen enthalten Total, Eligible und Registered Silver, Aenderung
zum Vortag und aktuelle Menge in Feinunzen.

## Einrichtung

1. GitHub Actions aktivieren und das Repository-Secret `NTFY_TOPIC` auf
   den abonnierten ntfy.sh-Topic setzen (nur den Topic-Namen, keine URL).
2. Den Topic vertraulich halten. Dieses Repository ist oeffentlich und
   verwendet die bisherige ntfy.sh-Konfiguration ohne zusaetzliche Authentifizierung.
3. Workflow **Silver Alert** manuell starten; die Option `test_push`
   sendet eine ausdruecklich als TEST markierte Meldung.

## Betrieb

Der Zeitplan ist Minute 07 und 37 jeder Stunde in UTC, an allen Tagen.
GitHub Actions bietet keine garantierte Ausfuehrungszeit. Jeder
tatsaechliche Lauf ruft CME ab, auch ausserhalb der ueblichen
Publikationszeit und an Feiertagen. Das fruehere erlernte Zeitfenster
darf keine Abfragen mehr unterdruecken. Es gibt keine Meldung fuer ein
unveraendertes Berichtsdatum. Bei voruebergehenden Abruffehlern wird
nach 30, 90 und 180 Sekunden erneut versucht.

Der Workflow verwendet den aktuellen Stand von main und eine gemeinsame
Concurrency-Gruppe; der Zustand wird auch nach einem fehlgeschlagenen
Abruf oder Versand gespeichert. Pull Requests fuehren nur netzwerkfreie
Regressionstests aus. Normale Code-Pushes auf main fuehren Tests aus; ein
bewusst markierter Deployment-Commit mit `[ntfy-test]` sendet nach
erfolgreichen Regressionstests einen einmaligen Test-Push. State-Commits
loesen keinen Deployment-Test aus.

## Dauerhaftes Archiv

- `state/reports/YYYY-MM-DD.json`: erster beobachteter Report mit
  Originalzahlen (vorher/heute, Differenz, Prozent), Erkennungszeit in UTC,
  HTTP Last-Modified, Datei-SHA256 und Meldungstext.
- `delivery.status`: `pending` (noch nicht bestaetigt), `accepted`
  (ntfy hat eine ID zurueckgegeben) oder `legacy_cutoff` (Datum lag bei
  Einfuehrung des Archivs nicht nach dem alten Meldemarker; kein erneuter
  Versand und keine erfundene ID). Aenderungen derselben CME-Datei werden
  als Revision archiviert und nicht erneut gepusht.
- `state/last_report.txt`: hoechstes regulaer bestaetigtes Meldedatum
  (bestehender Wert bleibt erhalten).
- `state/observations.json`: bisherige Beobachtungen, weitergefuehrt
  ohne Einfluss auf den Zeitplan.
- `state/last_test_push.json`: letzter erfolgreicher TEST samt
  Bestandszahlen, ID und UTC-Annahmezeit; veraendert den regulaeren Marker nicht.

Die Meldungs-ID belegt die Annahme durch den ntfy-Server, nicht den Empfang
auf einem iPhone. Das Archiv enthaelt weder den Topic noch Zugangsdaten.
Ausstehende Nachrichten werden bei folgenden Laeufen, auch bei einem
CME-Ausfall, erneut versendet. Bereits bestaetigte Reports werden nicht
erneut versendet. Bei einem Prozessabbruch genau zwischen Serverannahme
und dem Git-Push oder einem unklaren Netzwerk-Timeout sind Duplikate
moeglich; die Zustellung ist nicht transaktional.

Historische Zahlen vor Einfuehrung des Archivs und Reports, die CME vor
dem ersten erfolgreichen Abruf bereits ersetzt hat, koennen nicht aus
der aktuellen Excel-Datei rekonstruiert werden. US-Feiertage koennen
ohne neue Berichte bleiben.

## Lokal

Python 3.12, Zeitzonendaten und Internet fuer Live-Abfragen vorausgesetzt.

```bash
python -m pip install xlrd==2.0.2 curl_cffi
python -m unittest discover -s tests -v
python silver_alert.py --file Silver_stocks.xls
```

Live-Aufrufe benoetigen `NTFY_TOPIC` und veraendern lokale State-Dateien.
Sie sollten nicht parallel zu GitHub Actions ausgefuehrt werden.
`--once` (auch Standard ohne Parameter) fragt CME einmal erfolgreich ab
und bearbeitet ausstehende Meldungen; `--test-push` prueft den Versand
mit einer klar gekennzeichneten Testmeldung.
