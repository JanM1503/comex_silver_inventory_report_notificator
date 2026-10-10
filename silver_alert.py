#!/usr/bin/env python3
"""COMEX Silber Lagerbestand: Push aufs Handy mit der Veraenderung zum Vortag.

Das Skript beobachtet die CME Datei Silver_stocks.xls und meldet jeden neuen Report
(Total, Eligible, Registered mit Veraenderung in Prozent) per ntfy aufs Handy.

Den Zeitpunkt der Veroeffentlichung findet es selbst heraus. Jede Stunde startet eine kurze
Sitzung (GitHub Actions). Je nach Wissensstand fragt sie unterschiedlich oft ab:

  Phase 1, Suchen:      noch kein Zeitpunkt bekannt, Abfrage alle 15 Minuten, rund um die Uhr
  Phase 2, Verfeinern:  ein grobes Zeitfenster ist bekannt, im Fenster (plus 2 Stunden Reserve)
                        alle 5 Minuten, ausserhalb nur eine Abfrage pro Stunde
  Phase 3, Gelernt:     Zeitfenster aus 5 genauen Beobachtungen, nur noch dort alle 3 Minuten,
                        danach bei Verspaetung alle 15 Minuten bis 4 Stunden nach dem Fenster

Alle Beobachtungen stehen in state/observations.json. Alles wird in UTC gerechnet. Nach einer
Zeitumstellung (Schweiz oder USA) wird das Fenster vorsichtshalber 10 Tage lang breiter gehalten.

Aufruf:
    python silver_alert.py                normale Sitzung (braucht NTFY_TOPIC)
    python silver_alert.py --once         genau eine Abfrage mit voller Auswertung
    python silver_alert.py --test-push    aktuellen Report einmal als Test senden
    python silver_alert.py --file X.xls   nur anzeigen, was aus einer lokalen Datei wuerde
"""
import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from datetime import time as dtime
from email.utils import parsedate_to_datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import xlrd

URL = "https://www.cmegroup.com/delivery_reports/Silver_stocks.xls"
NTFY_SERVER = "https://ntfy.sh"

ROOT = Path(__file__).resolve().parent
STATE_DIR = ROOT / "state"
LAST_REPORT_FILE = STATE_DIR / "last_report.txt"  # Datum des zuletzt gemeldeten Reports
OBS_FILE = STATE_DIR / "observations.json"  # Beobachtungen fuer das Lernen

UTC = timezone.utc
ZURICH = ZoneInfo("Europe/Zurich")
CHICAGO = ZoneInfo("America/Chicago")

# Reihenfolge der Meldung: Total, Eligible, Registered
ROWS = [
    ("Total", "COMBINED TOTAL"),
    ("Eligible", "TOTAL ELIGIBLE"),
    ("Registered", "TOTAL REGISTERED"),
]

SESSION_MAX = timedelta(minutes=57)
LEARN_EVENTS = 5  # genaue Beobachtungen bis zum Status "gelernt"
PRECISE = timedelta(minutes=45)  # so eng muss der Zeitpunkt eingegrenzt sein
RECENT = 7  # so viele letzte Reports bestimmen das Fenster
KEEP = 40  # so viele Eintraege pro Art bleiben im Protokoll

SCAN_INTERVAL = 15 * 60
REFINE_INTERVAL = 5 * 60
FINE_INTERVAL = 3 * 60
LATE_INTERVAL = 15 * 60
RETRY_SECONDS = 60
MAX_ERRORS = 4

MARGIN_REFINE = timedelta(hours=2)
MARGIN = timedelta(minutes=30)
MARGIN_DST = timedelta(minutes=90)
LATE_EXTRA = timedelta(hours=4)  # so lange wird nach dem Fenster weiter gesucht


# ---------------------------------------------------------------- Zeit, austauschbar fuer Tests
def _now() -> datetime:
    return datetime.now(UTC)


def _sleep(seconds: float) -> None:
    time.sleep(seconds)


def log(msg: str) -> None:
    print(msg, flush=True)


def iso(t: datetime) -> str:
    return t.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_iso(s: str) -> datetime:
    return datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=UTC)


def fmt_ch(t: datetime) -> str:
    return t.astimezone(ZURICH).strftime("%H:%M")


# ---------------------------------------------------------------- Datei holen und lesen
@dataclass
class Report:
    report_date: date
    data: dict  # Name -> (Vortag, heute)
    digest: str
    last_modified: "datetime | None"


def download(url: str):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "silver-stocks-personal-alert/1.0",
            "Accept": "application/vnd.ms-excel,*/*",
        },
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read(), resp.headers.get("Last-Modified")


def parse(data: bytes):
    """Gibt (Report Datum, {Name: (vortag, heute)}) zurueck."""
    try:
        book = xlrd.open_workbook(file_contents=data)
    except Exception as exc:
        raise ValueError(f"Datei ist keine lesbare Excel Datei: {exc}")
    sheet = book.sheet_by_index(0)

    report_date = None
    found = {}
    wanted = {key for _, key in ROWS}
    for r in range(sheet.nrows):
        values = sheet.row_values(r)
        texts = [str(v).strip() for v in values if str(v).strip()]
        if not texts:
            continue
        m = re.search(r"Report Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})", " ".join(texts))
        if m:
            report_date = date(int(m[3]), int(m[1]), int(m[2]))
        label = texts[0].upper()
        if label in wanted:
            nums = [v for v in values if isinstance(v, float)]
            # Reihenfolge: PREV TOTAL, RECEIVED, WITHDRAWN, NET CHANGE, ADJUSTMENT, TOTAL TODAY
            if len(nums) != 6:
                raise ValueError(f"Unerwartetes Format in Zeile '{label}': {nums}")
            prev, _rec, _wd, net, adj, today = nums
            if abs(prev + net + adj - today) > 1.0:
                raise ValueError(f"Zahlen in Zeile '{label}' passen nicht zusammen: {nums}")
            found[label] = (prev, today)

    if report_date is None:
        raise ValueError("Report Date nicht gefunden")
    result = {}
    for name, key in ROWS:
        if key not in found:
            raise ValueError(f"Zeile '{key}' nicht gefunden")
        result[name] = found[key]
    return report_date, result


def fetch_report() -> Report:
    raw, lm = download(URL)
    report_date, data = parse(raw)
    lm_dt = None
    if lm:
        try:
            lm_dt = parsedate_to_datetime(lm).astimezone(UTC)
        except (TypeError, ValueError):
            lm_dt = None
    return Report(report_date, data, hashlib.sha256(raw).hexdigest(), lm_dt)


# ---------------------------------------------------------------- Nachricht
def fmt_oz(x: float) -> str:
    return f"{x / 1e6:.2f}M oz"


def fmt_delta(x: float) -> str:
    sign = "+" if x >= 0 else "-"
    a = abs(x)
    if a >= 1e6:
        return f"{sign}{a / 1e6:.2f}M oz"
    return f"{sign}{a / 1e3:.1f}K oz"


def build_message(report_date, data, detected=None):
    lines = []
    for name, (prev, today) in data.items():
        delta = today - prev
        pct = delta / prev * 100 if prev else 0.0
        arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "►")
        lines.append(f"{name}: {arrow} {pct:+.2f}% ({fmt_delta(delta)}), jetzt {fmt_oz(today)}")
    if detected is not None:
        lines.append(f"Erkannt um {fmt_ch(detected)} Uhr Schweizer Zeit")
    title = f"COMEX Silber, Report {report_date:%d.%m.%Y}"
    return title, "\n".join(lines)


def send_ntfy(topic: str, title: str, body: str) -> None:
    req = urllib.request.Request(
        f"{NTFY_SERVER}/{topic}",
        data=body.encode("utf-8"),
        method="POST",
        headers={"Title": title, "Tags": "bar_chart"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        resp.read()


# ---------------------------------------------------------------- Zustand
def load_obs() -> dict:
    try:
        obs = json.loads(OBS_FILE.read_text())
    except (OSError, ValueError):
        obs = {}
    obs.setdefault("last_seen", None)
    obs.setdefault("last_probe_utc", None)
    obs.setdefault("events", [])
    obs.setdefault("learning_announced", False)
    return obs


def save_obs(obs: dict) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    reports = [e for e in obs["events"] if e["kind"] == "new_report"][-KEEP:]
    others = [e for e in obs["events"] if e["kind"] != "new_report"][-KEEP:]
    obs["events"] = sorted(reports + others, key=lambda e: e["detected_utc"])
    OBS_FILE.write_text(json.dumps(obs, indent=1, ensure_ascii=False) + "\n")


def notified_date():
    try:
        return date.fromisoformat(LAST_REPORT_FILE.read_text().strip())
    except (OSError, ValueError):
        return None


def mark_notified(d: date) -> None:
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    LAST_REPORT_FILE.write_text(d.isoformat() + "\n")


def register(obs: dict, rep: Report, t: datetime, prev_probe) -> "str | None":
    """Vergleicht mit der letzten Beobachtung und protokolliert Aenderungen."""
    seen = obs["last_seen"]
    kind = None
    if seen is not None:
        if rep.report_date > date.fromisoformat(seen["report_date"]):
            kind = "new_report"
        elif rep.digest != seen["hash"]:
            kind = "file_change"
    if kind:
        obs["events"].append(
            {
                "kind": kind,
                "report_date": rep.report_date.isoformat(),
                "detected_utc": iso(t),
                "lower_utc": iso(prev_probe) if prev_probe else None,
                "last_modified_utc": iso(rep.last_modified) if rep.last_modified else None,
            }
        )
    obs["last_seen"] = {"report_date": rep.report_date.isoformat(), "hash": rep.digest}
    return kind


# ---------------------------------------------------------------- Lernen und Zeitplan
def report_events(obs: dict) -> list:
    return [e for e in obs["events"] if e["kind"] == "new_report"]


def bounds(ev: dict):
    """Gibt (fruehester Zeitpunkt, spaetester Zeitpunkt, genau?) der Veroeffentlichung zurueck."""
    hi = parse_iso(ev["detected_utc"])
    lo = parse_iso(ev["lower_utc"]) if ev.get("lower_utc") else None
    if ev.get("last_modified_utc"):
        lm = parse_iso(ev["last_modified_utc"])
        if hi - timedelta(hours=36) <= lm <= hi:  # nur glaubwuerdige Zeitstempel
            lo = max(lo, lm) if lo else lm
    precise = lo is not None and hi - lo <= PRECISE
    if lo is None or hi - lo > LATE_EXTRA:
        lo = hi - LATE_EXTRA
    return lo, hi, precise


def precise_count(events: list) -> int:
    return sum(1 for e in events if bounds(e)[2])


def boundary_minute(events: list) -> int:
    """Tagesgrenze in Minuten ab Mitternacht UTC: 12 Stunden vom letzten Report entfernt."""
    if not events:
        return 12 * 60
    t = parse_iso(events[-1]["detected_utc"])
    return (t.hour * 60 + t.minute + 12 * 60) % (24 * 60)


def day_start(t: datetime, bmin: int) -> datetime:
    shifted = t - timedelta(minutes=bmin)
    return datetime.combine(shifted.date(), dtime(0, 0), tzinfo=UTC) + timedelta(minutes=bmin)


def release_window(events: list, ds: datetime, bmin: int):
    """Typisches Fenster der letzten Reports, auf den Tag ab ds uebertragen."""
    offs_lo, offs_hi = [], []
    for ev in events[-RECENT:]:
        lo, hi, _ = bounds(ev)
        d0 = day_start(hi, bmin)
        offs_lo.append(lo - d0)
        offs_hi.append(hi - d0)
    return ds + min(offs_lo), ds + max(offs_hi)


def recent_dst_change(ds: datetime) -> bool:
    """Wahr, wenn in der Schweiz oder in Chicago in den letzten 10 Tagen die Uhr umgestellt wurde."""
    ref = datetime.combine(ds.date(), dtime(12, 0), tzinfo=UTC)
    before = ref - timedelta(days=10)
    return any(ref.astimezone(tz).utcoffset() != before.astimezone(tz).utcoffset() for tz in (ZURICH, CHICAGO))


def decide(now: datetime, obs: dict, end: datetime):
    """Gibt ("stop", Grund), ("wait", Sekunden) oder ("probe", Sekunden bis zur naechsten Abfrage) zurueck.

    Bei ("probe", None) wird genau einmal abgefragt und die Sitzung dann beendet."""
    if now >= end:
        return "stop", "Sitzungsende erreicht"
    events = report_events(obs)
    bmin = boundary_minute(events)
    ds = day_start(now, bmin)
    if events and day_start(parse_iso(events[-1]["detected_utc"]), bmin) == ds:
        return "stop", "Der Report dieses Tages ist schon erkannt"

    precise = precise_count(events)
    if precise == 0:
        return "probe", SCAN_INTERVAL  # Phase 1: Suchen

    lo, hi = release_window(events, ds, bmin)
    if precise < LEARN_EVENTS:  # Phase 2: Verfeinern
        lo, hi = lo - MARGIN_REFINE, hi + MARGIN_REFINE
        if lo <= now <= hi:
            return "probe", REFINE_INTERVAL
        return "probe", None

    # Phase 3: Gelernt
    active = {day_start(parse_iso(e["detected_utc"]), bmin).weekday() for e in events[-20:]}
    if ds.weekday() not in active:
        return "stop", "An diesem Wochentag kommt normalerweise kein Report"
    margin = MARGIN_DST if recent_dst_change(ds) else MARGIN
    lo, hi = lo - margin, hi + margin
    if now < lo:
        return "wait", (lo - now).total_seconds()
    if now <= hi:
        return "probe", FINE_INTERVAL
    if now <= hi + LATE_EXTRA:
        return "probe", LATE_INTERVAL
    return "stop", "Fenster ohne neuen Report zu Ende"


def learned(obs: dict) -> bool:
    return precise_count(report_events(obs)) >= LEARN_EVENTS


def summary(obs: dict, now: datetime) -> str:
    events = report_events(obs)
    if not events:
        return "Noch kein neuer Report beobachtet (Phase 1, Suchen)."
    p = precise_count(events)
    phase = "Phase 3, gelernt" if p >= LEARN_EVENTS else "Phase 2, Verfeinern"
    text = f"Beobachtete Reports: {len(events)}, genau eingegrenzt: {p} von {LEARN_EVENTS} ({phase})."
    bmin = boundary_minute(events)
    lo, hi = release_window(events, day_start(now, bmin), bmin)
    text += f" Bisheriges Fenster: {fmt_ch(lo)} bis {fmt_ch(hi)} Uhr Schweizer Zeit."
    return text


def announce_learning(obs: dict, topic: str, now: datetime) -> None:
    if obs["learning_announced"] or not learned(obs):
        return
    events = report_events(obs)
    bmin = boundary_minute(events)
    lo, hi = release_window(events, day_start(now, bmin), bmin)
    body = (
        f"Der Report erscheint meist zwischen {fmt_ch(lo)} und {fmt_ch(hi)} Uhr Schweizer Zeit. "
        "Ab jetzt wird nur noch rund um diese Zeit geprüft."
    )
    try:
        send_ntfy(topic, "Silber Alert: Zeitfenster gefunden", body)
        obs["learning_announced"] = True
    except (urllib.error.URLError, OSError) as exc:
        log(f"Hinweis zum Zeitfenster konnte nicht gesendet werden: {exc}")


# ---------------------------------------------------------------- Sitzung
def run_session(topic: str, once: bool = False) -> int:
    obs = load_obs()
    start = _now()
    end = start + SESSION_MAX
    prev_probe = parse_iso(obs["last_probe_utc"]) if obs.get("last_probe_utc") else None
    notified = notified_date()
    errors = 0
    send_failed = False
    log(f"Sitzung ab {start:%Y-%m-%d %H:%M} UTC ({fmt_ch(start)} Uhr Schweizer Zeit)")

    while True:
        now = _now()
        remaining = (end - now).total_seconds()
        if once:
            action, arg = "probe", None
        else:
            action, arg = decide(now, obs, end)

        if action == "stop":
            log(f"Ende: {arg}")
            break
        if action == "wait":
            if arg >= remaining:
                log("Ende: Das Suchfenster beginnt erst nach dieser Sitzung")
                break
            log(f"Warte bis {fmt_ch(now + timedelta(seconds=arg))} Uhr Schweizer Zeit auf den Beginn des Fensters")
            _sleep(arg)
            continue

        try:
            rep = fetch_report()
        except (urllib.error.URLError, OSError, ValueError) as exc:
            errors += 1
            log(f"Abruf fehlgeschlagen ({errors} von {MAX_ERRORS}): {exc}")
            if errors >= MAX_ERRORS:
                save_obs(obs)
                return 1
            _sleep(RETRY_SECONDS)
            continue
        errors = 0

        kind = register(obs, rep, now, prev_probe)
        prev_probe = now
        obs["last_probe_utc"] = iso(now)
        what = {"new_report": "NEUER REPORT", "file_change": "Datei geändert"}.get(kind, "unverändert")
        lm = rep.last_modified.strftime("%Y-%m-%d %H:%M UTC") if rep.last_modified else "nicht angegeben"
        log(f"Abfrage {now:%H:%M} UTC ({fmt_ch(now)} Uhr CH): Report {rep.report_date}, {what}, Last-Modified: {lm}")

        if notified is None or rep.report_date > notified:
            title, body = build_message(rep.report_date, rep.data, now)
            try:
                send_ntfy(topic, title, body)
            except (urllib.error.URLError, OSError) as exc:
                log(f"Senden fehlgeschlagen, neuer Versuch bei der nächsten Abfrage: {exc}")
                send_failed = True
            else:
                notified = rep.report_date
                mark_notified(notified)
                send_failed = False
                announce_learning(obs, topic, now)
                save_obs(obs)
                log("Push gesendet.")
                log(summary(obs, now))
                return 0

        if once or arg is None:
            break
        if (end - _now()).total_seconds() <= arg:
            log("Ende: Sitzungszeit aufgebraucht")
            break
        _sleep(arg)

    save_obs(obs)
    log(summary(obs, _now()))
    return 1 if send_failed else 0


# ---------------------------------------------------------------- Aufruf
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="lokale xls Datei nur anzeigen")
    ap.add_argument("--once", action="store_true", help="genau eine Abfrage")
    ap.add_argument("--test-push", action="store_true", help="aktuellen Report einmal als Test senden")
    args = ap.parse_args()

    if args.file:
        try:
            report_date, data = parse(Path(args.file).read_bytes())
        except (OSError, ValueError) as exc:
            print(f"Auswertung fehlgeschlagen: {exc}", file=sys.stderr)
            return 1
        title, body = build_message(report_date, data)
        print(title)
        print(body)
        return 0

    topic = os.environ.get("NTFY_TOPIC", "").strip()
    if not topic:
        print("NTFY_TOPIC ist nicht gesetzt.", file=sys.stderr)
        return 1

    if args.test_push:
        try:
            rep = fetch_report()
            title, body = build_message(rep.report_date, rep.data)
            send_ntfy(topic, title, body)
        except (urllib.error.URLError, OSError, ValueError) as exc:
            print(f"Test fehlgeschlagen: {exc}", file=sys.stderr)
            return 1
        print("Test Push gesendet.")
        return 0

    return run_session(topic, once=args.once)


if __name__ == "__main__":
    sys.exit(main())
