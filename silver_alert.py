#!/usr/bin/env python3
"""COMEX Silber: jeden Abruf archivieren, neue Reports per ntfy melden.

Jeder Start fragt CME ab, unabhaengig vom Wochentag und der Uhrzeit.
Fehlgeschlagene Nachrichten bleiben im lokalen Reportarchiv ausstehend.
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
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import xlrd

URL = "https://www.cmegroup.com/delivery_reports/Silver_stocks.xls"
NTFY_SERVER = "https://ntfy.sh"

ROOT = Path(__file__).resolve().parent
STATE_DIR = ROOT / "state"
LAST_REPORT_FILE = STATE_DIR / "last_report.txt"  # Datum des zuletzt gemeldeten Reports
OBS_FILE = STATE_DIR / "observations.json"
REPORT_DIR = STATE_DIR / "reports"
TEST_RECEIPT = STATE_DIR / "last_test_push.json"

UTC = timezone.utc
ZURICH = ZoneInfo("Europe/Zurich")

# Reihenfolge der Meldung: Total, Eligible, Registered
ROWS = [
    ("Total", "COMBINED TOTAL"),
    ("Eligible", "TOTAL ELIGIBLE"),
    ("Registered", "TOTAL REGISTERED"),
]

KEEP = 40  # Begrenzung nur der Beobachtungen, nicht des Reportarchivs
RETRY_DELAYS = (30, 90, 180)


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


BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.cmegroup.com/clearing/operations-and-deliveries/nymex-delivery-notices.html",
}


def _download_impersonated(url: str):
    """Abruf mit dem TLS Fingerabdruck von Chrome (curl_cffi). CME blockt reine Python Abrufe mit 403."""
    try:
        from curl_cffi import requests as cr
    except ImportError as exc:
        raise OSError(f"curl_cffi nicht installiert: {exc}")
    try:
        resp = cr.get(url, impersonate="chrome", headers=BROWSER_HEADERS, timeout=60)
    except Exception as exc:  # curl_cffi hat eigene Fehlerklassen
        raise OSError(f"curl_cffi Abruf fehlgeschlagen: {exc}")
    if resp.status_code != 200:
        raise OSError(f"HTTP Error {resp.status_code} (curl_cffi)")
    return resp.content, resp.headers.get("Last-Modified")


def _download_urllib(url: str):
    req = urllib.request.Request(url, headers=BROWSER_HEADERS)
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read(), resp.headers.get("Last-Modified")


def download(url: str):
    """Erst Chrome Fingerabdruck, dann klassisch per urllib mit Browser Headern."""
    errors = []
    for fn in (_download_impersonated, _download_urllib):
        try:
            return fn(url)
        except (urllib.error.URLError, OSError) as exc:
            errors.append(f"{fn.__name__}: {exc}")
    raise OSError(" | ".join(errors))


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


def send_ntfy(topic: str, title: str, body: str) -> str:
    """Gibt ausschliesslich die Nachrichten-ID des Servers zurueck."""
    req = urllib.request.Request(
        f"{NTFY_SERVER}/{topic}",
        data=body.encode("utf-8"),
        method="POST",
        headers={"Title": title, "Tags": "bar_chart"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        try:
            payload = json.loads(resp.read())
        except (ValueError, UnicodeError):
            raise OSError("ntfy: Antwort ohne lesbare JSON-Bestaetigung") from None
    msg_id = payload.get("id") if isinstance(payload, dict) else None
    if not isinstance(msg_id, str) or not msg_id.strip():
        raise OSError("ntfy: Antwort ohne Nachrichten-ID")
    return msg_id


# ---------------------------------------------------------------- Persistenter Zustand
def atomic_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + ".tmp")
    temp.write_text(text, encoding="utf-8")
    temp.replace(path)


def write_json(path: Path, value: dict) -> None:
    atomic_text(path, json.dumps(value, indent=2, ensure_ascii=False, allow_nan=False) + "\n")


def read_json(path: Path) -> dict:
    # Beschaedigte Zustandsdateien nicht stillschweigend ueberschreiben.
    result = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(result, dict):
        raise ValueError(f"Ungueltiger JSON-Zustand: {path.name}")
    return result


def load_obs() -> dict:
    obs = read_json(OBS_FILE) if OBS_FILE.exists() else {}
    obs.setdefault("last_seen", None)
    obs.setdefault("last_probe_utc", None)
    obs.setdefault("events", [])
    return obs


def save_obs(obs: dict) -> None:
    reports = [e for e in obs["events"] if e["kind"] == "new_report"][-KEEP:]
    others = [e for e in obs["events"] if e["kind"] != "new_report"][-KEEP:]
    obs["events"] = sorted(reports + others, key=lambda e: e["detected_utc"])
    write_json(OBS_FILE, obs)


def notified_date():
    try:
        return date.fromisoformat(LAST_REPORT_FILE.read_text().strip())
    except FileNotFoundError:
        return None


def mark_notified(d: date) -> None:
    previous = notified_date()
    if previous is None or d > previous:
        atomic_text(LAST_REPORT_FILE, d.isoformat() + "\n")


def register(obs: dict, rep: Report, t: datetime) -> str:
    seen = obs["last_seen"]
    kind = "unveraendert"
    if seen is None:
        kind = "new_report"
    elif rep.report_date < date.fromisoformat(seen["report_date"]):
        return "aelterer Report"
    elif rep.report_date > date.fromisoformat(seen["report_date"]):
        kind = "new_report"
    elif rep.digest != seen["hash"]:
        kind = "file_change"
    if kind in ("new_report", "file_change"):
        obs["events"].append({
            "kind": kind,
            "report_date": rep.report_date.isoformat(),
            "detected_utc": iso(t),
            "lower_utc": obs["last_probe_utc"],
            "last_modified_utc": iso(rep.last_modified) if rep.last_modified else None,
        })
    obs["last_seen"] = {"report_date": rep.report_date.isoformat(), "hash": rep.digest}
    return kind


def snapshot(rep: Report, detected: datetime) -> dict:
    title, body = build_message(rep.report_date, rep.data, detected)
    return {
        "report_date": rep.report_date.isoformat(),
        "detected_utc": iso(detected),
        "last_modified_utc": iso(rep.last_modified) if rep.last_modified else None,
        "sha256": rep.digest,
        "unit": "troy_oz",
        "inventory": {
            name: {
                "previous": prev,
                "current": today,
                "change": today - prev,
                "change_percent": (today - prev) / prev * 100 if prev else None,
            }
            for name, (prev, today) in rep.data.items()
        },
        "title": title,
        "body": body,
    }


def archive_report(rep: Report, detected: datetime) -> None:
    path = REPORT_DIR / (rep.report_date.isoformat() + ".json")
    if path.exists():
        # Erste Beobachtung und versendete Zahlen unveraendert lassen.
        # Korrigierte CME-Dateien als weitere Version dokumentieren.
        entry = read_json(path)
        revisions = entry.get("revisions", [])
        if rep.digest != entry["sha256"] and not any(
            r["sha256"] == rep.digest for r in revisions
        ):
            entry.setdefault("revisions", []).append(snapshot(rep, detected))
            write_json(path, entry)
        return
    entry = snapshot(rep, detected)
    cutoff = notified_date()
    entry["delivery"] = {
        "status": "legacy_cutoff" if cutoff is not None and rep.report_date <= cutoff else "pending",
        "attempts": 0,
        "ntfy_id": None,
        "accepted_utc": None,
    }
    write_json(path, entry)
    log(f"Report archiviert: {rep.report_date} ({entry['delivery']['status']})")


def publish_pending(topic: str) -> bool:
    # Auch bei CME-Ausfall werden bereits archivierte Meldungen erneut versucht.
    # Datumsgerechte Reihenfolge, kein Ueberspringen eines fehlgeschlagenen Versands.
    for path in sorted(REPORT_DIR.glob("????-??-??.json")):
        entry = read_json(path)
        delivery = entry["delivery"]
        if delivery["status"] == "accepted":
            # Repariert den Marker, falls der vorige Prozess nach Archivierung endete.
            mark_notified(date.fromisoformat(entry["report_date"]))
            continue
        if delivery["status"] != "pending":
            continue
        delivery["attempts"] += 1
        delivery["last_attempt_utc"] = iso(_now())
        write_json(path, entry)
        try:
            msg_id = send_ntfy(topic, entry["title"], entry["body"])
        except (OSError, ValueError):
            # Keine URL, kein Topic und kein unbereinigter Antwortinhalt im oeffentlichen Log.
            delivery["last_error"] = "ntfy publish not confirmed"
            write_json(path, entry)
            log(f"Push nicht bestaetigt: {entry['report_date']}; erneuter Versuch beim naechsten Lauf.")
            return False
        delivery.update(status="accepted", ntfy_id=msg_id, accepted_utc=iso(_now()))
        delivery.pop("last_error", None)
        write_json(path, entry)
        mark_notified(date.fromisoformat(entry["report_date"]))
        log(f"Push vom ntfy-Server angenommen: {entry['report_date']}, ID {msg_id}")
    return True


def fetch_with_retry() -> Report:
    for attempt in range(len(RETRY_DELAYS) + 1):
        try:
            return fetch_report()
        except (OSError, ValueError) as exc:
            log(f"CME-Abruf fehlgeschlagen ({attempt + 1}/{len(RETRY_DELAYS) + 1}): {exc}")
            if attempt == len(RETRY_DELAYS):
                raise
            _sleep(RETRY_DELAYS[attempt])


def run_session(topic: str, once: bool = False) -> int:
    # --once bleibt kompatibel: jetzt ist jeder regulaere Lauf eine kurze Sitzung.
    obs = load_obs()
    fetch_ok = True
    try:
        rep = fetch_with_retry()
    except (OSError, ValueError):
        fetch_ok = False
    else:
        detected = _now()
        archive_report(rep, detected)
        kind = register(obs, rep, detected)
        obs["last_probe_utc"] = iso(detected)
        save_obs(obs)
        lm = iso(rep.last_modified) if rep.last_modified else "nicht angegeben"
        log(f"CME: Report {rep.report_date}, {kind}, Last-Modified: {lm}")
    sent_ok = publish_pending(topic)
    return 0 if fetch_ok and sent_ok else 1


def test_push(topic: str) -> int:
    try:
        rep = fetch_with_retry()
    except (OSError, ValueError):
        log("Test fehlgeschlagen: CME-Abruf")
        return 1
    entry = snapshot(rep, _now())
    title = "TEST - " + entry["title"]
    try:
        msg_id = send_ntfy(topic, title, entry["body"])
    except (OSError, ValueError):
        log("Test fehlgeschlagen: ntfy-Versand nicht bestaetigt")
        return 1
    entry["title"] = title
    entry["delivery"] = {
        "status": "accepted",
        "ntfy_id": msg_id,
        "accepted_utc": iso(_now()),
        "attempts": 1,
    }
    write_json(TEST_RECEIPT, entry)
    log(f"Test-Push vom ntfy-Server angenommen. ID {msg_id}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="lokale xls Datei nur anzeigen")
    ap.add_argument("--once", action="store_true", help="einmaliger regulaerer Abruf")
    ap.add_argument("--test-push", action="store_true", help="aktuellen Report als Test senden")
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
        return test_push(topic)
    return run_session(topic, once=args.once)


if __name__ == "__main__":
    sys.exit(main())
