#!/usr/bin/env python3
"""COMEX Silber Lagerbestand: Push aufs Handy mit der Veraenderung zum Vortag.

Das Skript laedt die CME Datei Silver_stocks.xls, liest die drei Summenzeilen
(TOTAL REGISTERED, TOTAL ELIGIBLE, COMBINED TOTAL) mit PREV TOTAL und TOTAL TODAY,
berechnet die Veraenderung in Prozent und sendet sie ueber ntfy aufs Handy.

Pro Report wird nur einmal gemeldet. Das Datum des zuletzt gemeldeten Reports steht in
state/last_report.txt.

Aufruf:
    python silver_alert.py                 normaler Lauf (braucht NTFY_TOPIC)
    python silver_alert.py --file X.xls    lokale Datei statt Download
    python silver_alert.py --dry-run       nur anzeigen, nichts senden, nichts speichern
    python silver_alert.py --force         auch senden, wenn der Report schon gemeldet wurde
"""
import argparse
import os
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

import xlrd

URL = "https://www.cmegroup.com/delivery_reports/Silver_stocks.xls"
NTFY_SERVER = "https://ntfy.sh"
STATE_FILE = Path(__file__).resolve().parent / "state" / "last_report.txt"

ROWS = [
    ("Total", "COMBINED TOTAL"),
    ("Registered", "TOTAL REGISTERED"),
    ("Eligible", "TOTAL ELIGIBLE"),
]


def download(url: str) -> bytes:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "silver-stocks-personal-alert/1.0",
            "Accept": "application/vnd.ms-excel,*/*",
        },
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def parse(data: bytes):
    """Gibt (Report Datum, {Name: (vortag, heute)}) zurueck."""
    try:
        book = xlrd.open_workbook(file_contents=data)
    except Exception as exc:
        raise ValueError(f"Datei ist keine lesbare Excel Datei: {exc}")
    sheet = book.sheet_by_index(0)

    report_date = None
    found = {}
    for r in range(sheet.nrows):
        values = sheet.row_values(r)
        texts = [str(v).strip() for v in values if str(v).strip()]
        if not texts:
            continue
        m = re.search(r"Report Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})", " ".join(texts))
        if m:
            report_date = datetime(int(m[3]), int(m[1]), int(m[2])).date()
        label = texts[0].upper()
        if label in {key for _, key in ROWS}:
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


def fmt_oz(x: float) -> str:
    return f"{x / 1e6:.2f}M oz"


def fmt_delta(x: float) -> str:
    sign = "+" if x >= 0 else "-"
    a = abs(x)
    if a >= 1e6:
        return f"{sign}{a / 1e6:.2f}M oz"
    return f"{sign}{a / 1e3:.1f}K oz"


def build_message(report_date, data):
    lines = []
    for name, (prev, today) in data.items():
        delta = today - prev
        pct = delta / prev * 100 if prev else 0.0
        arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "►")
        lines.append(f"{name}: {arrow} {pct:+.2f}% ({fmt_delta(delta)}), jetzt {fmt_oz(today)}")
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


def already_sent(report_date) -> bool:
    return STATE_FILE.exists() and STATE_FILE.read_text().strip() == report_date.isoformat()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", help="lokale xls Datei statt Download")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    try:
        raw = Path(args.file).read_bytes() if args.file else download(URL)
    except (urllib.error.URLError, OSError) as exc:
        print(f"Download fehlgeschlagen: {exc}", file=sys.stderr)
        return 1

    try:
        report_date, data = parse(raw)
    except ValueError as exc:
        print(f"Auswertung fehlgeschlagen: {exc}", file=sys.stderr)
        return 1

    if already_sent(report_date) and not args.force and not args.dry_run:
        print(f"Report {report_date} wurde schon gemeldet. Nichts zu tun.")
        return 0

    title, body = build_message(report_date, data)
    print(title)
    print(body)

    if args.dry_run:
        return 0

    topic = os.environ.get("NTFY_TOPIC", "").strip()
    if not topic:
        print("NTFY_TOPIC ist nicht gesetzt.", file=sys.stderr)
        return 1
    try:
        send_ntfy(topic, title, body)
    except (urllib.error.URLError, OSError) as exc:
        print(f"Senden an ntfy fehlgeschlagen: {exc}", file=sys.stderr)
        return 1

    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(report_date.isoformat() + "\n")
    print("Push gesendet.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
