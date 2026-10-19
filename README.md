# COMEX Silver Inventory Alert

Automatische Push-Benachrichtigungen zu den täglichen COMEX-Silberlagerbeständen – mit Veränderungen zum Vortag und einem selbst ermittelten Veröffentlichungsfenster.

Das Python-Skript liest den offiziellen [CME-Lagerbericht](https://www.cmegroup.com/delivery_reports/Silver_stocks.xls) und sendet neue Berichte über [ntfy](https://ntfy.sh) aufs Handy. Der enthaltene GitHub-Actions-Workflow übernimmt den regelmäßigen Betrieb und speichert die Beobachtungen im Repository.

## Inhalt

- [Funktionsumfang](#funktionsumfang)
- [Beispielmeldung](#beispielmeldung)
- [Einrichtung mit GitHub Actions](#einrichtung-mit-github-actions)
- [Lokale Installation](#lokale-installation)
- [Verwendung](#verwendung)
- [Automatische Zeitplanung](#automatische-zeitplanung)
- [Konfiguration und gespeicherter Zustand](#konfiguration-und-gespeicherter-zustand)
- [Fehlerbehebung](#fehlerbehebung)

## Funktionsumfang

- **Drei Bestandskennzahlen:** Total, Eligible und Registered mit aktuellem Bestand, absoluter Veränderung und prozentualer Veränderung zum Vortag.
- **Benachrichtigungen über ntfy:** Ein neueres Berichtsdatum löst eine Meldung aus; das zuletzt erfolgreich gemeldete Datum verhindert wiederholte Meldungen im normalen Betrieb.
- **Adaptiver Abruf:** Das Skript grenzt die Veröffentlichungszeit anhand bisheriger Beobachtungen ein und passt die Abfrageintervalle an.
- **Berücksichtigung der Zeitumstellung:** Änderungen in Zürich und Chicago vergrößern vorübergehend das Suchfenster.
- **Lokale Auswertung:** Eine vorhandene XLS-Datei lässt sich ohne Netzwerkzugriff und ohne Versand prüfen.

| Kennzahl | Bedeutung |
| --- | --- |
| **Total** | Gesamter Silberbestand aus Eligible und Registered. |
| **Eligible** | Silber, das die Lieferanforderungen erfüllt, aber nicht als Registered geführt wird. |
| **Registered** | Silber, für das ein Lagerzertifikat zur Lieferung gegen Futures ausgestellt ist. |

## Beispielmeldung

Fiktive Werte zur Veranschaulichung:

```text
COMEX Silber, Report 16.10.2026
Total: ▼ -1.41% (-5.00M oz), jetzt 350.00M oz
Eligible: ▼ -1.09% (-3.00M oz), jetzt 272.00M oz
Registered: ▼ -2.50% (-2.00M oz), jetzt 78.00M oz
Erkannt um 22:15 Uhr Schweizer Zeit
```

`M oz` steht für Millionen Feinunzen, `K oz` für Tausend Feinunzen. Die Erkennungszeit bezeichnet den Abruf durch das Skript und kann nach der eigentlichen Veröffentlichung liegen.

## Einrichtung mit GitHub Actions

Für diesen Betriebsmodus ist kein dauerhaft laufender eigener Server erforderlich.

1. **Repository forken** und im eigenen Repository GitHub Actions aktivieren.
2. **ntfy-Topic abonnieren:** In der ntfy-App oder im Web einen eigenen Topic-Namen auf `https://ntfy.sh` abonnieren. Verwende einen langen, schwer zu erratenden Namen und behandle ihn vertraulich: Der aktuelle Versand nutzt keine zusätzliche Authentifizierung.
3. **Repository-Secret hinterlegen:** Unter **Settings → Secrets and variables → Actions → New repository secret** ein Secret namens `NTFY_TOPIC` anlegen. Als Wert ausschließlich den Topic-Namen eintragen, ohne Serveradresse oder führenden Schrägstrich.
4. **Schreibzugriff ermöglichen:** Der Workflow benötigt `contents: write`, um den aktualisierten Zustand zurück ins Repository zu schreiben. Repository- oder Organisationsregeln und der Schutz des Standardbranches müssen diese Pushes zulassen.
5. **Testmeldung auslösen:** Unter **Actions → Silver Alert → Run workflow** die Option `test_push` aktivieren und den Workflow starten. Dieser Aufruf lädt den aktuellen CME-Bericht und sendet ihn unabhängig vom zuletzt gemeldeten Datum.
6. **Automatischen Betrieb prüfen:** Die nächsten geplanten Läufe und ihre Protokolle im Actions-Tab kontrollieren. Bei Forks gegebenenfalls den geplanten Workflow gesondert aktivieren.

Der Workflow ist mit `7 * * * *` für Minute **07 jeder Stunde in UTC** konfiguriert. GitHub führt geplante Workflows auf dem Standardbranch aus; Starts können sich verzögern. Die einzelne Beobachtungssitzung läuft regulär bis zu 57 Minuten, das Actions-Joblimit beträgt 65 Minuten. Parallele Workflow-Läufe werden durch eine gemeinsame Concurrency-Gruppe vermieden.

Ein manueller Lauf ohne `test_push` verwendet `--once`: Er wertet den aktuellen Bericht aus und sendet bei einem neueren Berichtsdatum eine reguläre Benachrichtigung.

## Lokale Installation

Voraussetzungen:

- **Python 3.12**, entsprechend dem GitHub-Actions-Workflow.
- **pip** und Unterstützung für virtuelle Umgebungen (`venv`).
- **Zeitzonendaten** für `Europe/Zurich` und `America/Chicago`.
- Für Live-Abfragen und Versand: HTTPS-Zugriff auf `www.cmegroup.com` und `ntfy.sh`.

```bash
git clone https://github.com/JanM1503/comex_silver_inventory_report_notificator.git
cd comex_silver_inventory_report_notificator

python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install xlrd==2.0.2
python silver_alert.py --help
```

Unter Windows lässt sich die Umgebung in PowerShell mit `.venv\Scripts\Activate.ps1` aktivieren. Fehlen systemweite Zeitzonendaten, können sie innerhalb der virtuellen Umgebung mit `python -m pip install tzdata` ergänzt werden.

`xlrd` ist die einzige externe Python-Abhängigkeit des Skripts. Das Repository enthält derzeit keine Abhängigkeitsdatei; der Actions-Workflow installiert `xlrd` ohne Versionsbindung. Die lokale Anleitung verwendet die beim Einrichten geprüfte Version 2.0.2.

## Verwendung

Alle folgenden Befehle werden im Repository-Verzeichnis mit aktivierter Python-Umgebung ausgeführt.

### Lokale Datei auswerten

```bash
python silver_alert.py --file /pfad/zum/Silver_stocks.xls
```

Gibt Titel und Bestandszahlen aus. Dieser Modus benötigt kein `NTFY_TOPIC`, sendet keine Nachricht und verändert keine Zustandsdateien. Erwartet wird das XLS-Format des CME-Berichts; eine beliebige Excel-Datei oder eine XLSX-Datei genügt nicht.

### Topic für den Versand setzen

Für Live-Aufrufe muss `NTFY_TOPIC` in der Prozessumgebung vorhanden sein. In Bash lässt sich der Name ohne Anzeige bei der Eingabe setzen:

```bash
read -rsp 'ntfy-Topic: ' NTFY_TOPIC
export NTFY_TOPIC
```

Das Skript lädt keine `.env`-Datei automatisch.

### Live-Aufrufe

Diese Aufrufe können echte Benachrichtigungen an das konfigurierte Topic senden:

| Befehl | Verhalten |
| --- | --- |
| `python silver_alert.py` | Startet eine Sitzung mit adaptiven Abfragen. Speichert Beobachtungen und das zuletzt erfolgreich gemeldete Datum. |
| `python silver_alert.py --once` | Führt eine einzelne erfolgreiche Abfrage mit regulärer Auswertung aus. Kann senden und den Zustand aktualisieren; bei Abruffehlern erfolgen Wiederholungsversuche. |
| `python silver_alert.py --test-push` | Lädt den aktuellen Bericht und sendet eine Testmeldung unabhängig vom gespeicherten Meldedatum. Verändert den Zustand nicht. |

Im normalen Betrieb wird gesendet, wenn noch kein gültiges Meldedatum gespeichert ist oder der abgerufene Bericht ein neueres Datum trägt. Änderungen an derselben Berichtsdatei werden als Dateiänderungen protokolliert, lösen allein aber keine erneute Bestandsmeldung aus.

**Lokale Live-Läufe verändern Dateien unter `state/`.** Diese Dateien sind versioniert. Lokale Läufe und GitHub Actions sollten daher nicht unkoordiniert denselben Zustand verwenden. Der bereits im Repository enthaltene Zustand beeinflusst, ob ein Bericht als neu gilt; für eine reine Versandprobe eignet sich `--test-push`.

Ein lokaler Aufruf startet genau eine Sitzung. Die stündliche Wiederholung übernimmt ausschließlich der Scheduler, beispielsweise der enthaltene GitHub-Actions-Workflow.

## Automatische Zeitplanung

Das Skript verwendet eine regelbasierte Auswertung bisheriger Beobachtungen. Als genau eingegrenzt gilt eine Veröffentlichung, deren ermitteltes Zeitintervall höchstens 45 Minuten umfasst. Ein plausibler HTTP-Header `Last-Modified` kann diese Eingrenzung unterstützen.

| Phase | Voraussetzung | Abfrageverhalten |
| --- | --- | --- |
| **1 · Suchen** | Noch keine genau eingegrenzte Veröffentlichung. | Alle 15 Minuten innerhalb der laufenden Sitzung. |
| **2 · Verfeinern** | Eine bis vier genau eingegrenzte Veröffentlichungen. | Alle 5 Minuten im ermittelten Fenster mit je 2 Stunden Reserve; außerhalb einmal pro Sitzung. |
| **3 · Gelernt** | Mindestens fünf genau eingegrenzte Veröffentlichungen. | Alle 3 Minuten im ermittelten Fenster mit je 30 Minuten Reserve; bei Verspätung alle 15 Minuten bis 4 Stunden nach dem erweiterten Fenster. |

Das Veröffentlichungsfenster basiert auf den letzten sieben beobachteten neuen Berichten. In Phase 3 berücksichtigt das Skript zusätzlich die beobachteten Veröffentlichungstage aus bis zu 20 Berichten. Vor dem Fenster wartet es, sofern der Beginn noch in die Sitzung fällt; außerhalb der relevanten Zeiten beendet es die Sitzung.

Nach einer Zeitumstellung in Zürich oder Chicago wird die Reserve in Phase 3 für zehn Tage von 30 auf 90 Minuten je Seite erhöht. Intern rechnet das Skript in UTC; Uhrzeiten in Meldungen und Zusammenfassungen erscheinen in Schweizer Zeit.

Nach einer erfolgreichen Bestandsmeldung endet die Sitzung. Sobald die Lernschwelle erreicht ist, kann zusätzlich eine einmalige Meldung zum gefundenen Zeitfenster folgen.

## Konfiguration und gespeicherter Zustand

### Umgebungsvariable

| Name | Verwendung |
| --- | --- |
| `NTFY_TOPIC` | Topic-Name auf `ntfy.sh`. Für normale Sitzungen, `--once` und `--test-push` erforderlich; für `--file` nicht erforderlich. |

Berichtsadresse, ntfy-Server, Abfrageintervalle und Lernparameter sind Konstanten in [silver_alert.py](silver_alert.py). Für diese Einstellungen gibt es derzeit keine zusätzlichen Umgebungsvariablen oder Konfigurationsdateien.

### Dateien

```text
.
├── .github/workflows/silver-alert.yml  # Stündlicher Betrieb und Zustandssicherung
├── silver_alert.py                   # Download, Auswertung, Zeitplanung und Versand
└── state/
    ├── last_report.txt               # Datum des zuletzt erfolgreich gemeldeten Berichts
    └── observations.json             # Letzter Abruf, Datei-Hash, Ereignisse und Lernstatus
```

In `observations.json` bleiben bis zu 40 Ereignisse zu neuen Berichten und bis zu 40 sonstige Ereignisse erhalten. Das Meldedatum wird erst nach erfolgreichem Versand aktualisiert. Der Actions-Workflow versucht auch nach fehlgeschlagenen Schritten, vorhandene Zustandsänderungen zu committen und zu pushen.

Ein Zurücksetzen der Zustandsdateien entfernt Lernfortschritt beziehungsweise die Information über bereits gesendete Berichte und kann erneute Benachrichtigungen verursachen.

## Fehlerbehebung

| Symptom | Prüfung und Lösung |
| --- | --- |
| `NTFY_TOPIC ist nicht gesetzt.` | Prozessvariable prüfen; in GitHub Actions das Repository-Secret exakt `NTFY_TOPIC` nennen. Für eine lokale Dateiauswertung `--file` verwenden. |
| `ModuleNotFoundError: No module named 'xlrd'` | Virtuelle Umgebung aktivieren und `python -m pip install xlrd==2.0.2` im verwendeten Interpreter ausführen. |
| `ZoneInfoNotFoundError` | System-Zeitzonendaten bereitstellen oder `python -m pip install tzdata` ausführen. |
| HTTP-Fehler, Timeout oder Proxy-Fehler | HTTPS-Zugriff auf `www.cmegroup.com` und `ntfy.sh` sowie Proxy- und Netzwerkregeln prüfen. Ein `403` kann vom Proxy oder vom Zielserver stammen. |
| `Report Date nicht gefunden`, fehlende Zeilen oder unlesbare Excel-Datei | Prüfen, ob eine echte CME-XLS-Datei statt einer HTML-Fehlerseite vorliegt. Ein geändertes Berichtsformat kann eine Parseranpassung erfordern. |
| Keine neue Push-Meldung | Abonniertes Topic und `state/last_report.txt` prüfen. Bei unverändertem Berichtsdatum ist das im normalen Betrieb vorgesehen. `--test-push` erlaubt eine erneute Versandprobe. |
| Actions-Lauf startet nicht | Actions und geplante Workflows aktivieren; Standardbranch und GitHub-Ausführungsprotokolle prüfen. Geplante Starts können sich verzögern. |
| Zustand kann nicht gepusht werden | `contents: write`, Organisationsrichtlinien und Branch-Schutz prüfen. Fehlende Zustandssicherung kann bei späteren Läufen wiederholte Meldungen auslösen. |

Abruf- und Auswertungsfehler führen zu bis zu vier aufeinanderfolgenden Versuchen mit jeweils 60 Sekunden Pause. Bei anhaltenden Abruffehlern oder einem am Sitzungsende weiterhin fehlgeschlagenen Versand liefert das Skript Exit-Code `1`. Exit-Code `0` bedeutet einen erfolgreich abgeschlossenen Aufruf; dazu zählt auch ein Lauf ohne neuen Bericht.
