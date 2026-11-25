#!/usr/bin/env python3
"""
E+S HubSpot-Backup

Sichert den kompletten CRM-Bestand aus HubSpot (Portal 26584208) und legt ihn als
verschlüsseltes ZIP im SharePoint der Leitung ab. Mit dem Inhalt kann E+S nach
einem Ausfall oder Angriff sofort in einem anderen CRM weiterarbeiten.

Gesichert werden:
  - Kontakte, Unternehmen, Deals, Tickets, Leads, Produkte, Deal-Positionen,
    Angebote und alle eigenen Objekte, jeweils mit ALLEN Feldern
  - Aktivitäten: Notizen, Anrufe, E-Mails, Meetings, Aufgaben, Nachrichten, Briefpost
  - alle Verknüpfungen (welcher Kontakt gehört zu welcher Firma, welchem Deal usw.)
    inklusive Beschriftungen
  - Feldbeschreibungen, Pipelines und Phasen, Besitzer, Listen mit Mitgliedern
  - Dateianhänge an Aktivitäten

Ausgabe im ZIP:
  csv/              eine CSV je Objekt, alle Felder, interne Werte (für den Import)
  excel/            eine Excel-Datei je Objekt, lesbare Werte (für Menschen)
  verknuepfungen/   eine CSV je Objektpaar
  struktur/         Felder, Pipelines, Besitzer, Listen, Verknüpfungsarten
  rohdaten/         Originalantwort von HubSpot als JSON Lines (volle Genauigkeit)
  Anhänge liegen in eigenen Archiven HubSpot-Anhaenge_*.zip (nur neue je Lauf)
  00_Uebersicht.xlsx und LIESMICH.txt

Alle Einstellungen kommen aus Umgebungsvariablen, siehe README.md.
Das Skript liest HubSpot nur. Es ändert dort nichts.
"""
from __future__ import annotations

import base64
import csv
import datetime as dt
import html
import json
import logging
import os
import re
import shutil
import sys
import time
import zipfile
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill

try:
    import pyzipper
except ImportError:  # Verschlüsselung dann nicht möglich
    pyzipper = None


# --------------------------------------------------------------------------
# Einstellungen
# --------------------------------------------------------------------------

def env(name: str, default: str | None = None) -> str | None:
    value = os.environ.get(name)
    return value if value not in (None, "") else default


BERLIN = ZoneInfo("Europe/Berlin")
NOW = dt.datetime.now(BERLIN)
STAMP = NOW.strftime("%Y-%m-%d_%H%M")
BACKUP_PREFIX = "HubSpot-Backup_"
BACKUP_NAME = f"{BACKUP_PREFIX}{STAMP}"
ANHANG_NAME = f"HubSpot-Anhaenge_{STAMP}"
ANHANG_INDEX = "anhaenge_index.json"

HUBSPOT_TOKEN = env("HUBSPOT_TOKEN")
WORK_DIR = Path(env("WORK_DIR", "./backup_arbeit"))
OUT_DIR = Path(env("OUT_DIR", "./backup_ausgabe"))
REQUESTS_PER_SECOND = float(env("HUBSPOT_RPS", "8"))
DOWNLOAD_FILES = env("DOWNLOAD_FILES", "1") == "1"
ZIP_PASSWORD = env("ZIP_PASSWORD")

# Ablage in SharePoint / OneDrive über Microsoft Graph
MS_TENANT_ID = env("MS_TENANT_ID")
MS_CLIENT_ID = env("MS_CLIENT_ID")
MS_CLIENT_SECRET = env("MS_CLIENT_SECRET")
SP_SHARE_LINK = env("SP_SHARE_LINK")   # Freigabelink des Zielordners
SP_DRIVE_ID = env("SP_DRIVE_ID")       # alternativ direkt Laufwerks-ID ...
SP_FOLDER_ID = env("SP_FOLDER_ID")     # ... und Ordner-ID
KEEP_DAILY = int(env("KEEP_DAILY", "30"))       # so viele letzte Stände bleiben
KEEP_MONTHLY = int(env("KEEP_MONTHLY", "12"))   # plus je Monat der erste Stand
DROP_ALERT_PERCENT = float(env("DROP_ALERT_PERCENT", "20"))
REQUIRE_UPLOAD = env("REQUIRE_UPLOAD", "1" if env("GITHUB_ACTIONS") else "0") == "1"

EXCEL_MAX_CELL = 32000
EXCEL_MAX_ROWS = 1_000_000

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("backup")


# --------------------------------------------------------------------------
# Objekte
# --------------------------------------------------------------------------

# (API-Name, deutscher Name, unverzichtbar)
STANDARD_OBJECTS = [
    ("contacts", "Kontakte", True),
    ("companies", "Unternehmen", True),
    ("deals", "Deals", True),
    ("tickets", "Tickets", False),
    ("leads", "Leads", False),
    ("products", "Produkte", False),
    ("line_items", "Deal-Positionen", False),
    ("quotes", "Angebote", False),
    ("notes", "Notizen", True),
    ("calls", "Anrufe", True),
    ("emails", "E-Mails", True),
    ("meetings", "Meetings", True),
    ("tasks", "Aufgaben", True),
    ("communications", "Nachrichten (SMS, WhatsApp, LinkedIn)", False),
    ("postal_mail", "Briefpost", False),
]
ENGAGEMENTS = {"notes", "calls", "emails", "meetings", "tasks", "communications", "postal_mail"}
RECORD_TARGETS = ["contacts", "companies", "deals", "tickets"]
COUNT_CHECK = ["contacts", "companies", "deals", "notes"]

# Feld, das einen Datensatz in Excel lesbar benennt
NAME_FIELDS = {
    "contacts": ["firstname", "lastname", "email"],
    "companies": ["name"],
    "deals": ["dealname"],
    "tickets": ["subject"],
    "leads": ["hs_lead_name"],
    "products": ["name"],
    "line_items": ["name"],
    "quotes": ["hs_title"],
}

# Diese Felder stehen vorne, der Rest folgt alphabetisch
PREFERRED_COLUMNS = {
    "contacts": ["firstname", "lastname", "email", "phone", "mobilephone", "company", "jobtitle",
                 "address", "zip", "city", "country", "hubspot_owner_id", "lifecyclestage",
                 "hs_lead_status", "createdate", "lastmodifieddate"],
    "companies": ["name", "domain", "phone", "address", "zip", "city", "state", "country",
                  "hubspot_owner_id", "lifecyclestage", "createdate", "hs_lastmodifieddate"],
    "deals": ["dealname", "amount", "pipeline", "dealstage", "closedate", "hubspot_owner_id",
              "createdate", "hs_lastmodifieddate"],
    "tickets": ["subject", "content", "hs_pipeline", "hs_pipeline_stage", "hs_ticket_priority",
                "hubspot_owner_id", "createdate"],
    "notes": ["hs_timestamp", "hs_note_body", "hubspot_owner_id"],
    "calls": ["hs_timestamp", "hs_call_title", "hs_call_body", "hs_call_direction",
              "hs_call_disposition", "hs_call_duration", "hubspot_owner_id"],
    "emails": ["hs_timestamp", "hs_email_subject", "hs_email_text", "hs_email_direction",
               "hs_email_from_email", "hs_email_to_email", "hubspot_owner_id"],
    "meetings": ["hs_timestamp", "hs_meeting_title", "hs_meeting_body", "hs_meeting_start_time",
                 "hs_meeting_end_time", "hs_meeting_outcome", "hubspot_owner_id"],
    "tasks": ["hs_timestamp", "hs_task_subject", "hs_task_body", "hs_task_status",
              "hs_task_priority", "hubspot_owner_id"],
}

PIPELINE_FIELDS = {"pipeline", "hs_pipeline"}
STAGE_FIELDS = {"dealstage", "hs_pipeline_stage"}


# --------------------------------------------------------------------------
# Hilfen
# --------------------------------------------------------------------------

class Report:
    """Sammelt Zahlen, Warnungen und ausgelassene Bereiche für LIESMICH und Log."""

    def __init__(self) -> None:
        self.counts: dict[str, int] = {}
        self.labels: dict[str, str] = {}
        self.extra: list[tuple[str, int, str]] = []   # (Bereich, Anzahl, Datei)
        self.warnings: list[str] = []
        self.skipped: list[str] = []
        self.incomplete = False

    def warn(self, text: str, essential: bool = False) -> None:
        log.warning(text)
        self.warnings.append(text)
        if essential:
            self.incomplete = True

    def skip(self, text: str, essential: bool = False) -> None:
        log.warning("Ausgelassen: %s", text)
        self.skipped.append(text)
        if essential:
            self.incomplete = True


class HubSpotError(RuntimeError):
    def __init__(self, status: int, text: str):
        super().__init__(f"HubSpot {status}: {text[:300]}")
        self.status = status


def chunks(items: list, size: int):
    for i in range(0, len(items), size):
        yield items[i:i + size]


def safe_name(text: str, limit: int = 80) -> str:
    text = re.sub(r"[^\w.\- äöüÄÖÜß]", "_", text or "").strip(" ._")
    return (text or "datei")[:limit]


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(path: Path, header: list[str], rows) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    n = 0
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        for row in rows:
            w.writerow(row)
            n += 1
    return n


def iter_jsonl(path: Path):
    if not path.exists():
        return
    with path.open(encoding="utf-8") as f:
        for line in f:
            if line.strip():
                yield json.loads(line)


TAG_RE = re.compile(r"<[^>]+>")
BR_RE = re.compile(r"<\s*(br|/p|/div|/li|/tr)[^>]*>", re.I)


def strip_html(text: str) -> str:
    text = BR_RE.sub("\n", text)
    text = TAG_RE.sub("", text)
    text = html.unescape(text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def parse_hs_datetime(value: str):
    try:
        if value.isdigit():
            moment = dt.datetime.fromtimestamp(int(value) / 1000, tz=dt.timezone.utc)
        else:
            moment = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
            if moment.tzinfo is None:
                moment = moment.replace(tzinfo=dt.timezone.utc)
        return moment.astimezone(BERLIN).replace(tzinfo=None)
    except (ValueError, OverflowError):
        return value


def parse_hs_date(value: str):
    try:
        if value.isdigit():
            return dt.datetime.fromtimestamp(int(value) / 1000, tz=dt.timezone.utc).date()
        return dt.date.fromisoformat(value[:10])
    except (ValueError, OverflowError):
        return value


def excel_text(value):
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        if len(value) > EXCEL_MAX_CELL:
            value = value[:EXCEL_MAX_CELL] + " […gekürzt, vollständig in csv/]"
    return value


# --------------------------------------------------------------------------
# HubSpot-Zugriff
# --------------------------------------------------------------------------

class HubSpot:
    BASE = "https://api.hubapi.com"

    def __init__(self, token: str, rps: float):
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        })
        self.interval = 1.0 / max(rps, 0.5)
        self.last = 0.0
        self.calls = 0

    def request(self, method: str, path: str, **kwargs) -> requests.Response:
        url = path if path.startswith("http") else self.BASE + path
        for attempt in range(8):
            pause = self.interval - (time.monotonic() - self.last)
            if pause > 0:
                time.sleep(pause)
            self.last = time.monotonic()
            self.calls += 1
            try:
                resp = self.session.request(method, url, timeout=120, **kwargs)
            except requests.RequestException as exc:
                log.warning("Verbindungsfehler (%s), neuer Versuch", exc.__class__.__name__)
                time.sleep(min(60, 2 ** attempt))
                continue
            if resp.status_code == 429 or resp.status_code >= 500:
                retry = resp.headers.get("Retry-After")
                time.sleep(float(retry) if retry and retry.replace(".", "").isdigit()
                           else min(60, 2 ** attempt))
                continue
            return resp
        raise HubSpotError(0, f"keine Antwort nach 8 Versuchen: {method} {path}")

    def json(self, method: str, path: str, ok=(200,), **kwargs):
        resp = self.request(method, path, **kwargs)
        if resp.status_code not in ok:
            raise HubSpotError(resp.status_code, resp.text)
        return resp.json() if resp.content else {}

    def paged(self, path: str, params: dict | None = None, key: str = "results"):
        params = dict(params or {})
        after = None
        while True:
            if after:
                params["after"] = after
            data = self.json("GET", path, params=params)
            yield from data.get(key, [])
            after = data.get("paging", {}).get("next", {}).get("after")
            if not after:
                break


# --------------------------------------------------------------------------
# Export-Bausteine
# --------------------------------------------------------------------------

class ObjectInfo:
    def __init__(self, otype: str, label: str, essential: bool, props: list[dict]):
        self.otype = otype
        self.label = label
        self.essential = essential
        self.props = props
        self.prop_by_name = {p["name"]: p for p in props}
        self.ids: list[str] = []
        self.name_fields = NAME_FIELDS.get(otype, [])

    @property
    def file_stem(self) -> str:
        label = self.label.replace("Eigenes Objekt: ", "").split(" (")[0]
        return safe_name(label).replace(" ", "_")


def export_owners(hs: HubSpot, root: Path, rep: Report) -> dict[str, str]:
    owners: dict[str, str] = {}
    rows = []
    try:
        for archived in ("false", "true"):
            for o in hs.paged("/crm/v3/owners", {"limit": 100, "archived": archived}):
                name = " ".join(x for x in (o.get("firstName"), o.get("lastName")) if x) or o.get("email") or o["id"]
                owners[str(o["id"])] = name
                teams = ", ".join(t.get("name", "") for t in o.get("teams") or [])
                rows.append([o["id"], o.get("userId"), o.get("firstName"), o.get("lastName"),
                             o.get("email"), teams, archived == "true"])
    except HubSpotError as exc:
        rep.skip(f"Besitzer (Verkäufer) ({exc.status}). Bereich crm.objects.owners.read fehlt?", essential=True)
        return owners
    n = write_csv(root / "struktur" / "besitzer.csv",
                  ["owner_id", "user_id", "vorname", "nachname", "email", "teams", "deaktiviert"], rows)
    rep.extra.append(("Besitzer (Verkäufer)", n, "struktur/besitzer.csv"))
    log.info("Besitzer: %d", n)
    return owners


def discover_custom_objects(hs: HubSpot, root: Path, rep: Report) -> list[tuple[str, str, bool, str]]:
    try:
        data = hs.json("GET", "/crm/v3/schemas")
    except HubSpotError as exc:
        if exc.status in (401, 403):
            rep.skip(f"Eigene Objekte ({exc.status}). Bereich crm.schemas.custom.read fehlt?")
        return []
    write_json(root / "struktur" / "eigene_objekte_schema.json", data)
    found = []
    for s in data.get("results", []):
        label = (s.get("labels") or {}).get("plural") or s.get("name")
        found.append((s["objectTypeId"], f"Eigenes Objekt: {label}", False, s.get("primaryDisplayProperty") or ""))
    log.info("Eigene Objekte: %d", len(found))
    return found


def export_object(hs: HubSpot, otype: str, label: str, essential: bool,
                  root: Path, rep: Report) -> ObjectInfo | None:
    log.info("Exportiere %s …", label)
    try:
        props = hs.json("GET", f"/crm/v3/properties/{otype}").get("results", [])
    except HubSpotError as exc:
        if exc.status in (400, 401, 403, 404):
            rep.skip(f"{label}: Felder nicht lesbar ({exc.status})", essential)
            return None
        raise
    info = ObjectInfo(otype, label, essential, props)
    write_json(root / "struktur" / "felder" / f"{info.file_stem}.json", props)
    write_csv(root / "struktur" / "felder" / f"{info.file_stem}.csv",
              ["name", "bezeichnung", "gruppe", "typ", "feldtyp", "berechnet", "optionen"],
              ([p["name"], p.get("label"), p.get("groupName"), p.get("type"), p.get("fieldType"),
                p.get("calculated", False),
                "; ".join(f'{o.get("value")}={o.get("label")}' for o in p.get("options") or [])]
               for p in props))

    # 1. alle IDs holen
    try:
        for rec in hs.paged(f"/crm/v3/objects/{otype}",
                            {"limit": 100, "archived": "false", "properties": "hs_object_id"}):
            info.ids.append(str(rec["id"]))
    except HubSpotError as exc:
        if exc.status in (400, 401, 403, 404):
            rep.skip(f"{label}: Datensätze nicht lesbar ({exc.status})", essential)
            return None
        raise

    # 2. Datensätze mit allen Feldern in 100er-Paketen lesen
    names = [p["name"] for p in props]
    raw = root / "rohdaten" / f"{info.file_stem}.jsonl"
    raw.parent.mkdir(parents=True, exist_ok=True)
    written = 0
    with raw.open("w", encoding="utf-8") as f:
        for part in chunks(info.ids, 100):
            data = hs.json("POST", f"/crm/v3/objects/{otype}/batch/read", ok=(200, 207),
                           json={"properties": names, "inputs": [{"id": i} for i in part]})
            for rec in data.get("results", []):
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                written += 1
    missing = len(info.ids) - written
    if missing > 0:
        rep.warn(f"{label}: {missing} Datensätze zwischen Auflistung und Abruf verschwunden "
                 f"(vermutlich gerade gelöscht)")
    rep.counts[otype] = written
    rep.labels[otype] = label
    log.info("%s: %d Datensätze, %d Felder", label, written, len(names))
    return info


def export_pipelines(hs: HubSpot, objects: dict[str, ObjectInfo], root: Path,
                     rep: Report) -> dict[str, dict[str, dict[str, str]]]:
    """Gibt je Objekt {'pipelines': {id: label}, 'stages': {id: 'Pipeline / Phase'}} zurück."""
    result: dict[str, dict[str, dict[str, str]]] = {}
    rows = []
    for otype in objects:
        if otype in ENGAGEMENTS or otype in ("contacts", "companies", "products", "line_items", "quotes"):
            continue
        try:
            data = hs.json("GET", f"/crm/v3/pipelines/{otype}")
        except HubSpotError as exc:
            if exc.status not in (400, 403, 404):
                rep.warn(f"Pipelines {otype}: {exc}")
            continue
        maps = {"pipelines": {}, "stages": {}}
        for p in data.get("results", []):
            maps["pipelines"][str(p["id"])] = p.get("label", "")
            for s in p.get("stages", []):
                maps["stages"][str(s["id"])] = f'{p.get("label", "")} / {s.get("label", "")}'
                meta = s.get("metadata") or {}
                rows.append([otype, p["id"], p.get("label"), s["id"], s.get("label"),
                             s.get("displayOrder"), meta.get("probability"), meta.get("isClosed"),
                             s.get("archived", False)])
        result[otype] = maps
        write_json(root / "struktur" / "pipelines" / f"{safe_name(otype)}.json", data)
    if rows:
        n = write_csv(root / "struktur" / "pipelines_und_phasen.csv",
                      ["objekt", "pipeline_id", "pipeline", "phase_id", "phase", "reihenfolge",
                       "wahrscheinlichkeit", "abgeschlossen", "archiviert"], rows)
        rep.extra.append(("Pipelines und Phasen", n, "struktur/pipelines_und_phasen.csv"))
    return result


def association_pairs(objects: dict[str, ObjectInfo]) -> list[tuple[str, str]]:
    present = set(objects)
    pairs = [
        ("contacts", "companies"), ("contacts", "deals"), ("contacts", "tickets"),
        ("companies", "companies"), ("companies", "deals"), ("companies", "tickets"),
        ("deals", "tickets"), ("deals", "line_items"),
        ("quotes", "deals"), ("quotes", "line_items"), ("quotes", "contacts"), ("quotes", "companies"),
        ("leads", "contacts"), ("leads", "companies"), ("leads", "deals"),
    ]
    for e in sorted(ENGAGEMENTS):
        for t in RECORD_TARGETS:
            pairs.append((e, t))
    for o in sorted(present):
        if o.startswith("2-"):
            for t in RECORD_TARGETS:
                pairs.append((o, t))
    return [(a, b) for a, b in pairs if a in present and b in present]


def export_associations(hs: HubSpot, objects: dict[str, ObjectInfo], root: Path, rep: Report):
    """Schreibt je Paar eine CSV und gibt assoc[typ][id][zieltyp] = [ids] zurück."""
    assoc: dict[str, dict[str, dict[str, list[str]]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    label_rows = []
    for a, b in association_pairs(objects):
        # Verknüpfungsarten (Beschriftungen)
        try:
            for lab in hs.json("GET", f"/crm/v4/associations/{a}/{b}/labels").get("results", []):
                label_rows.append([a, b, lab.get("category"), lab.get("typeId"), lab.get("label")])
        except HubSpotError:
            pass

        rows = []
        try:
            for part in chunks(objects[a].ids, 1000):
                data = hs.json("POST", f"/crm/v4/associations/{a}/{b}/batch/read", ok=(200, 207),
                               json={"inputs": [{"id": i} for i in part]})
                for res in data.get("results", []):
                    from_id = str(res["from"]["id"])
                    targets = list(res.get("to", []))
                    after = res.get("paging", {}).get("next", {}).get("after")
                    while after:
                        more = hs.json("GET", f"/crm/v4/objects/{a}/{from_id}/associations/{b}",
                                       params={"limit": 500, "after": after})
                        targets.extend(more.get("results", []))
                        after = more.get("paging", {}).get("next", {}).get("after")
                    for t in targets:
                        to_id = str(t["toObjectId"])
                        types = t.get("associationTypes") or []
                        rows.append([a, from_id, b, to_id,
                                     "; ".join(x["label"] for x in types if x.get("label")),
                                     "; ".join(f'{x.get("category")}:{x.get("typeId")}' for x in types)])
                        assoc[a][from_id][b].append(to_id)
                        if a != b:
                            assoc[b][to_id][a].append(from_id)
        except HubSpotError as exc:
            essential = objects[a].essential and objects[b].essential
            rep.skip(f"Verknüpfungen {a} → {b} ({exc.status})", essential)
            continue
        fname = f"{safe_name(a)}__{safe_name(b)}.csv"
        n = write_csv(root / "verknuepfungen" / fname,
                      ["von_objekt", "von_id", "zu_objekt", "zu_id", "beschriftung", "art"], rows)
        rep.extra.append((f"Verknüpfungen {objects[a].label} → {objects[b].label}", n, f"verknuepfungen/{fname}"))
        log.info("Verknüpfungen %s → %s: %d", a, b, n)
    write_csv(root / "struktur" / "verknuepfungsarten.csv",
              ["von_objekt", "zu_objekt", "kategorie", "type_id", "beschriftung"], label_rows)
    return assoc


def export_lists(hs: HubSpot, root: Path, rep: Report) -> None:
    lists = []
    offset = 0
    try:
        while True:
            data = hs.json("POST", "/crm/v3/lists/search", json={"offset": offset, "count": 500})
            lists.extend(data.get("lists", []))
            if not data.get("hasMore"):
                break
            offset = data.get("offset", offset + 500)
    except HubSpotError as exc:
        rep.skip(f"Listen/Segmente ({exc.status}). Bereich crm.lists.read fehlt?")
        return
    write_json(root / "struktur" / "listen.json", lists)
    members = []
    for lst in lists:
        lid = lst.get("listId")
        try:
            for m in hs.paged(f"/crm/v3/lists/{lid}/memberships", {"limit": 250}):
                members.append([lid, lst.get("name"), lst.get("objectTypeId"), m.get("recordId"),
                                m.get("membershipTimestamp")])
        except HubSpotError as exc:
            rep.warn(f"Mitglieder der Liste {lst.get('name')} nicht lesbar ({exc.status})")
    write_csv(root / "struktur" / "listen.csv", ["list_id", "name", "objekt", "typ", "angelegt"],
              ([l.get("listId"), l.get("name"), l.get("objectTypeId"), l.get("processingType"),
                l.get("createdAt")] for l in lists))
    n = write_csv(root / "struktur" / "listen_mitglieder.csv",
                  ["list_id", "liste", "objekt", "datensatz_id", "seit"], members)
    rep.extra.append(("Listen/Segmente", len(lists), "struktur/listen.csv"))
    rep.extra.append(("Listen-Mitgliedschaften", n, "struktur/listen_mitglieder.csv"))
    log.info("Listen: %d, Mitgliedschaften: %d", len(lists), n)


def export_files(hs: HubSpot, objects: dict[str, ObjectInfo], root: Path, rep: Report,
                 known: dict[str, str], new_dir: Path, archive_name: str) -> dict[str, str]:
    """Lädt nur Anhänge, die noch in keinem früheren Anhang-Archiv liegen.

    known:   Datei-ID -> Name des Archivs, in dem sie schon gesichert ist
    new_dir: Ordner für neue Anhänge (wird als eigenes Archiv hochgeladen)
    Rückgabe: neu gesicherte Datei-IDs -> archive_name
    """
    refs: dict[str, list[str]] = defaultdict(list)
    for otype, info in objects.items():
        if "hs_attachment_ids" not in info.prop_by_name:
            continue
        for rec in iter_jsonl(root / "rohdaten" / f"{info.file_stem}.jsonl"):
            value = (rec.get("properties") or {}).get("hs_attachment_ids") or ""
            for fid in re.split(r"[;,\s]+", value):
                if fid.strip():
                    refs[fid.strip()].append(f"{otype}:{rec['id']}")
    if not refs:
        return {}
    new_dir.mkdir(parents=True, exist_ok=True)
    rows, failed, denied = [], 0, 0
    added: dict[str, str] = {}
    plain = requests.Session()
    for fid, used_by in refs.items():
        if fid in known:
            rows.append([fid, "", "", "", known[fid], "", " ".join(used_by)])
            continue
        try:
            meta = hs.json("GET", f"/files/v3/files/{fid}/signed-url")
        except HubSpotError as exc:
            if exc.status in (401, 403):
                denied += 1
            failed += 1
            rows.append([fid, "", "", "", "", f"nicht abrufbar ({exc.status})", " ".join(used_by)])
            continue
        name = safe_name(meta.get("name") or fid)
        ext = meta.get("extension") or ""
        fname = f"{fid}_{name}" + (f".{ext}" if ext and not name.lower().endswith("." + ext.lower()) else "")
        try:
            with plain.get(meta["url"], stream=True, timeout=300) as r:
                r.raise_for_status()
                with (new_dir / fname).open("wb") as out:
                    for block in r.iter_content(1024 * 1024):
                        out.write(block)
            added[fid] = archive_name
            rows.append([fid, meta.get("name"), ext, meta.get("size"), archive_name, fname, " ".join(used_by)])
        except requests.RequestException as exc:
            failed += 1
            rows.append([fid, meta.get("name"), ext, meta.get("size"), "",
                         f"Download fehlgeschlagen ({exc.__class__.__name__})", " ".join(used_by)])
    write_csv(root / "struktur" / "dateien.csv",
              ["datei_id", "name", "endung", "groesse_bytes", "archiv", "datei_im_archiv", "haengt_an"], rows)
    saved = len(refs) - failed
    rep.extra.append(("Dateianhänge gesamt", saved, "Archive HubSpot-Anhaenge_*.zip (Liste: struktur/dateien.csv)"))
    rep.extra.append(("davon heute neu gesichert", len(added), f"{archive_name}.zip" if added else ""))
    if denied:
        rep.skip(f"{denied} Anhänge ohne Berechtigung. Bereiche files und files.ui_hidden.read fehlen?")
    elif failed:
        rep.warn(f"{failed} von {len(refs)} Anhängen nicht gesichert, siehe struktur/dateien.csv")
    log.info("Anhänge: %d bekannt, %d neu gesichert, %d fehlgeschlagen", len(refs) - len(added) - failed, len(added), failed)
    return added


# --------------------------------------------------------------------------
# Tabellen schreiben
# --------------------------------------------------------------------------

def column_order(info: ObjectInfo) -> list[str]:
    names = [p["name"] for p in info.props]
    preferred = [c for c in PREFERRED_COLUMNS.get(info.otype, info.name_fields) if c in info.prop_by_name]
    rest = sorted(n for n in names if n not in preferred)
    return preferred + rest


def display_name(info: ObjectInfo, props: dict) -> str:
    parts = [props.get(f) for f in info.name_fields if props.get(f)]
    if info.otype == "contacts" and len(parts) == 3:
        parts = parts[:2]
    return " ".join(str(p) for p in parts).strip()


def build_name_lookup(objects: dict[str, ObjectInfo], root: Path) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    for otype, info in objects.items():
        if otype in ENGAGEMENTS or not info.name_fields:
            continue
        names = {}
        for rec in iter_jsonl(root / "rohdaten" / f"{info.file_stem}.jsonl"):
            names[str(rec["id"])] = display_name(info, rec.get("properties") or {})
        lookup[otype] = names
    return lookup


def assoc_targets(info: ObjectInfo, assoc) -> list[str]:
    """Zieltypen, die als Spalte in die Objekttabelle kommen (ohne Aktivitäten)."""
    targets = set()
    for per_id in assoc.get(info.otype, {}).values():
        targets.update(per_id.keys())
    return sorted(t for t in targets if t not in ENGAGEMENTS)


class ValueFormatter:
    """Macht aus HubSpot-Rohwerten lesbare Excel-Werte."""

    def __init__(self, info: ObjectInfo, owners: dict[str, str], pipelines: dict):
        self.info = info
        self.owners = owners
        self.pipes = pipelines.get(info.otype, {"pipelines": {}, "stages": {}})
        self.options = {
            p["name"]: {str(o.get("value")): o.get("label") for o in p.get("options") or []}
            for p in info.props if p.get("options")
        }

    def __call__(self, name: str, value):
        if value is None or value == "":
            return None
        p = self.info.prop_by_name.get(name, {})
        ptype = p.get("type")
        if name in PIPELINE_FIELDS and value in self.pipes["pipelines"]:
            return self.pipes["pipelines"][value]
        if name in STAGE_FIELDS and value in self.pipes["stages"]:
            return self.pipes["stages"][value]
        if p.get("referencedObjectType") == "OWNER" or name == "hubspot_owner_id":
            return "; ".join(self.owners.get(v, v) for v in str(value).split(";"))
        if ptype == "enumeration" and name in self.options:
            opts = self.options[name]
            return excel_text("; ".join(str(opts.get(v, v)) for v in str(value).split(";")))
        if ptype == "datetime":
            return parse_hs_datetime(str(value))
        if ptype == "date":
            return parse_hs_date(str(value))
        if ptype == "number":
            try:
                num = float(value)
                return int(num) if num.is_integer() and abs(num) < 1e15 else num
            except ValueError:
                return value
        if ptype == "bool":
            return {"true": "Ja", "false": "Nein"}.get(str(value).lower(), value)
        if p.get("fieldType") == "html" or (isinstance(value, str) and "<" in value and ">" in value
                                              and name.endswith(("_body", "_html", "content"))):
            value = strip_html(value)
        return excel_text(value)


def write_object_tables(info: ObjectInfo, root: Path, assoc, lookup, owners, pipelines, rep: Report) -> None:
    raw = root / "rohdaten" / f"{info.file_stem}.jsonl"
    cols = column_order(info)
    targets = assoc_targets(info, assoc)
    my_assoc = assoc.get(info.otype, {})

    # CSV: alle Felder, Rohwerte
    nonempty: set[str] = set()
    csv_path = root / "csv" / f"{info.file_stem}.csv"
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["hubspot_id", "angelegt", "geaendert"] + cols + [f"verknuepft_{t}" for t in targets])
        for rec in iter_jsonl(raw):
            props = rec.get("properties") or {}
            for c in cols:
                if props.get(c) not in (None, ""):
                    nonempty.add(c)
            links = my_assoc.get(str(rec["id"]), {})
            w.writerow([rec["id"], rec.get("createdAt"), rec.get("updatedAt")]
                       + [props.get(c) for c in cols]
                       + [";".join(dict.fromkeys(links.get(t, []))) for t in targets])

    # Excel: nur gefüllte Spalten, lesbare Werte, Namen statt IDs bei Verknüpfungen
    xcols = [c for c in cols if c in nonempty]
    fmt = ValueFormatter(info, owners, pipelines)
    labels = {o: (rep.labels.get(o) or o) for o in targets}
    header = (["HubSpot-ID"]
              + [f'{info.prop_by_name[c].get("label") or c} ({c})' for c in xcols]
              + [f"Verknüpft: {labels[t]}" for t in targets])

    wb = Workbook(write_only=True)
    head_font = Font(bold=True, color="FFFFFF", name="Arial")
    head_fill = PatternFill("solid", fgColor="1F3864")

    def new_sheet(index: int):
        title = safe_name(info.label, 25) + (f" {index}" if index > 1 else "")
        ws = wb.create_sheet(title=title[:31])
        ws.freeze_panes = "B2"
        cells = []
        for h in header:
            cell = WriteOnlyCell(ws, value=h)
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cells.append(cell)
        ws.append(cells)
        return ws

    sheet_no, rows_in_sheet = 1, 0
    ws = new_sheet(sheet_no)
    for rec in iter_jsonl(raw):
        if rows_in_sheet >= EXCEL_MAX_ROWS:
            sheet_no += 1
            rows_in_sheet = 0
            ws = new_sheet(sheet_no)
        props = rec.get("properties") or {}
        links = my_assoc.get(str(rec["id"]), {})
        row = [int(rec["id"]) if str(rec["id"]).isdigit() else rec["id"]]
        row += [fmt(c, props.get(c)) for c in xcols]
        for t in targets:
            names = lookup.get(t, {})
            row.append(excel_text("; ".join(
                f"{names[i]} [{i}]" if names.get(i) else i for i in dict.fromkeys(links.get(t, [])))) or None)
        ws.append(row)
        rows_in_sheet += 1
    xpath = root / "excel" / f"{info.file_stem}.xlsx"
    xpath.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xpath)


def write_overview(root: Path, rep: Report, objects: dict[str, ObjectInfo], hs_calls: int, seconds: float) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Übersicht"
    bold = Font(bold=True, name="Arial")
    head_font = Font(bold=True, color="FFFFFF", name="Arial")
    head_fill = PatternFill("solid", fgColor="1F3864")
    ws.append(["E+S HubSpot-Backup"])
    ws["A1"].font = Font(bold=True, size=14, name="Arial")
    ws.append(["Stand", NOW.strftime("%d.%m.%Y %H:%M Uhr")])
    ws.append(["Status", "UNVOLLSTÄNDIG, siehe Hinweise" if rep.incomplete else "vollständig"])
    ws.append([])
    ws.append(["Bereich", "Anzahl", "CSV (Import)", "Excel (lesen)"])
    for c in ws[5]:
        c.font = head_font
        c.fill = head_fill
    for otype, info in objects.items():
        ws.append([info.label, rep.counts.get(otype, 0), f"csv/{info.file_stem}.csv",
                   f"excel/{info.file_stem}.xlsx"])
    for label, n, path in rep.extra:
        ws.append([label, n, path, ""])
    ws.append([])
    if rep.skipped or rep.warnings:
        ws.append(["Hinweise"])
        ws.cell(ws.max_row, 1).font = bold
        for t in rep.skipped:
            ws.append(["Ausgelassen", t])
        for t in rep.warnings:
            ws.append(["Warnung", t])
    for col, width in zip("ABCD", (48, 14, 52, 40)):
        ws.column_dimensions[col].width = width
    wb.save(root / "00_Uebersicht.xlsx")

    lines = [
        "E+S HUBSPOT-BACKUP",
        f"Stand: {NOW.strftime('%d.%m.%Y %H:%M Uhr')}",
        f"Status: {'UNVOLLSTÄNDIG' if rep.incomplete else 'vollständig'}",
        f"Laufzeit: {seconds / 60:.0f} Minuten, {hs_calls} Abfragen an HubSpot",
        "",
        "INHALT",
    ]
    for otype, info in objects.items():
        lines.append(f"  {info.label}: {rep.counts.get(otype, 0)}")
    for label, n, _ in rep.extra:
        lines.append(f"  {label}: {n}")
    if rep.skipped or rep.warnings:
        lines += ["", "HINWEISE"] + [f"  Ausgelassen: {t}" for t in rep.skipped] + [f"  Warnung: {t}" for t in rep.warnings]
    lines += [
        "",
        "ORDNER",
        "  csv/             eine Datei je Objekt mit ALLEN Feldern und Rohwerten. Für den Import in ein neues CRM.",
        "                   Spalte hubspot_id ist der eindeutige Schlüssel. Spalten verknuepft_* enthalten die",
        "                   IDs verknüpfter Datensätze, getrennt durch Semikolon.",
        "  excel/           dieselben Daten zum Lesen: nur gefüllte Spalten, Klartext statt interner Werte,",
        "                   Datum in deutscher Zeit, Namen bei Verknüpfungen.",
        "  verknuepfungen/  jede Verbindung zwischen zwei Datensätzen als eigene Zeile, mit Beschriftung.",
        "  struktur/        Feldbeschreibungen, Pipelines und Phasen, Besitzer, Listen, Anhangsverzeichnis.",
        "  rohdaten/        unveränderte Antworten von HubSpot (JSON Lines). Höchste Genauigkeit.",
        "  Anhänge          liegen NICHT in diesem ZIP, sondern in den Archiven HubSpot-Anhaenge_*.zip im selben",
        "                   Ordner. Jedes Archiv enthält die Anhänge, die seit dem vorherigen dazugekommen sind.",
        "",
        "WIEDERHERSTELLUNG IN EINEM ANDEREN CRM",
        "  1. Benutzer anlegen (struktur/besitzer.csv). hubspot_owner_id in den Daten zeigt auf owner_id.",
        "  2. Felder anlegen (struktur/felder/), Pipelines und Phasen anlegen (struktur/pipelines_und_phasen.csv).",
        "  3. Importieren in dieser Reihenfolge: Unternehmen, Kontakte, Deals, Tickets, Produkte, Positionen,",
        "     danach Notizen, Anrufe, E-Mails, Meetings, Aufgaben.",
        "  4. Die alte hubspot_id bei jedem Import in ein eigenes Feld (z. B. \"HubSpot-ID alt\") übernehmen.",
        "     Darüber lassen sich die Verknüpfungen aus verknuepfungen/ wieder zuordnen.",
        "  5. Anhänge: struktur/dateien.csv nennt je Datei das Archiv (Spalte archiv), den Dateinamen darin",
        "     und den Datensatz, an dem sie hängt (Spalte haengt_an).",
        "",
        "Nicht enthalten, weil HubSpot dafür keine Exportschnittstelle bietet oder es keine Kundendaten sind:",
        "Berichte und Dashboards, Workflows, Sequenzen, E-Mail-Vorlagen, Marketing-E-Mails, Formulare, Website.",
        "Die Zahlen hinter den Berichten (KPI-Felder) sind in den Daten enthalten.",
        "",
        "Datenschutz: Diese Datei enthält personenbezogene Daten. Nur im SharePoint-Ordner der Leitung",
        "aufbewahren, nicht weiterleiten, nicht auf private Geräte kopieren.",
    ]
    (root / "LIESMICH.txt").write_text("\n".join(lines) + "\n", encoding="utf-8")


def make_zip(src: Path, target: Path, password: str | None) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in src.rglob("*") if p.is_file())
    if password:
        if pyzipper is None:
            raise RuntimeError("ZIP_PASSWORD gesetzt, aber pyzipper fehlt (pip install pyzipper)")
        zf = pyzipper.AESZipFile(target, "w", compression=pyzipper.ZIP_DEFLATED,
                                 encryption=pyzipper.WZ_AES, allowZip64=True)
        zf.setpassword(password.encode("utf-8"))
    else:
        zf = zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED, allowZip64=True)
    with zf:
        for p in files:
            zf.write(p, arcname=f"{src.name}/{p.relative_to(src).as_posix()}")
    return target


# --------------------------------------------------------------------------
# SharePoint / OneDrive über Microsoft Graph
# --------------------------------------------------------------------------

class Graph:
    BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant: str, client_id: str, secret: str):
        self.tenant, self.client_id, self.secret = tenant, client_id, secret
        self.session = requests.Session()
        self.token_until = 0.0
        self.drive_id = None
        self.folder_id = None
        self.login()

    def login(self) -> None:
        """Holt eine frische Anmeldung. Microsoft-Anmeldungen laufen nach ca. 60 Minuten ab."""
        resp = requests.post(
            f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token",
            data={"client_id": self.client_id, "client_secret": self.secret,
                  "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"},
            timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"Anmeldung bei Microsoft fehlgeschlagen: {resp.status_code} {resp.text[:300]}")
        data = resp.json()
        self.session.headers["Authorization"] = f"Bearer {data['access_token']}"
        # 5 Minuten Puffer vor dem echten Ablauf
        self.token_until = time.monotonic() + int(data.get("expires_in", 3600)) - 300

    def ensure_login(self) -> None:
        if time.monotonic() >= self.token_until:
            log.info("Microsoft-Anmeldung wird erneuert")
            self.login()

    def call(self, method: str, path: str, ok=(200,), **kwargs):
        url = path if path.startswith("http") else self.BASE + path
        for attempt in range(6):
            self.ensure_login()
            resp = self.session.request(method, url, timeout=300, **kwargs)
            if resp.status_code == 401 and attempt == 0:
                self.login()
                continue
            if resp.status_code in (429, 500, 502, 503, 504):
                time.sleep(float(resp.headers.get("Retry-After", 2 ** attempt)))
                continue
            if resp.status_code not in ok:
                raise RuntimeError(f"Graph {method} {path}: {resp.status_code} {resp.text[:300]}")
            return resp
        raise RuntimeError(f"Graph {method} {path}: keine Antwort")

    def resolve_target(self, share_link: str | None, drive_id: str | None, folder_id: str | None) -> None:
        if drive_id and folder_id:
            self.drive_id, self.folder_id = drive_id, folder_id
            return
        if not share_link:
            raise RuntimeError("Ziel fehlt: SP_SHARE_LINK oder SP_DRIVE_ID und SP_FOLDER_ID setzen")
        share_id = "u!" + base64.urlsafe_b64encode(share_link.encode()).decode().rstrip("=")
        item = self.call("GET", f"/shares/{share_id}/driveItem").json()
        self.drive_id = item["parentReference"]["driveId"]
        self.folder_id = item["id"]
        log.info("Zielordner: %s (SP_DRIVE_ID=%s, SP_FOLDER_ID=%s)", item.get("name"), self.drive_id, self.folder_id)

    def _item_path(self, name: str) -> str:
        return f"/drives/{self.drive_id}/items/{self.folder_id}:/{quote(name)}:"

    def upload(self, path: Path, name: str | None = None) -> None:
        name = name or path.name
        session = self.call("POST", self._item_path(name) + "/createUploadSession",
                            json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}).json()
        url = session["uploadUrl"]
        size = path.stat().st_size
        chunk = 320 * 1024 * 32  # 10 MiB, Vielfaches von 320 KiB
        with path.open("rb") as f:
            start = 0
            while start < size:
                data = f.read(chunk)
                end = start + len(data) - 1
                for attempt in range(6):
                    try:
                        r = requests.put(url, data=data, timeout=600, headers={
                            "Content-Length": str(len(data)),
                            "Content-Range": f"bytes {start}-{end}/{size}"})
                    except requests.RequestException:
                        time.sleep(2 ** attempt)
                        continue
                    if r.status_code in (200, 201, 202):
                        break
                    if r.status_code in (429, 500, 502, 503, 504):
                        time.sleep(2 ** attempt)
                        continue
                    raise RuntimeError(f"Upload fehlgeschlagen: {r.status_code} {r.text[:300]}")
                else:
                    raise RuntimeError("Upload fehlgeschlagen: keine Antwort")
                start = end + 1
        log.info("Hochgeladen: %s (%.1f MB)", name, size / 1e6)

    def upload_text(self, name: str, text: str) -> None:
        self.call("PUT", self._item_path(name) + "/content", ok=(200, 201),
                  data=text.encode("utf-8"), headers={"Content-Type": "application/json"})

    def read_text(self, name: str) -> str | None:
        self.ensure_login()
        resp = self.session.get(self.BASE + self._item_path(name) + "/content", timeout=60)
        return resp.text if resp.status_code == 200 else None

    def list_backups(self) -> list[dict]:
        items, url = [], f"/drives/{self.drive_id}/items/{self.folder_id}/children?$select=id,name&$top=999"
        while url:
            data = self.call("GET", url).json()
            items += data.get("value", [])
            url = data.get("@odata.nextLink")
        return [i for i in items if i["name"].startswith(BACKUP_PREFIX) and i["name"].endswith(".zip")]

    def cleanup(self, keep_daily: int, keep_monthly: int) -> int:
        backups = sorted(self.list_backups(), key=lambda i: i["name"], reverse=True)
        keep = {i["id"] for i in backups[:keep_daily]}
        months: dict[str, dict] = {}
        for item in sorted(backups, key=lambda i: i["name"]):      # ältester zuerst
            month = item["name"][len(BACKUP_PREFIX):len(BACKUP_PREFIX) + 7]   # JJJJ-MM
            months.setdefault(month, item)
        for month in sorted(months, reverse=True)[:keep_monthly]:
            keep.add(months[month]["id"])
        removed = 0
        for item in backups:
            if item["id"] not in keep:
                self.call("DELETE", f"/drives/{self.drive_id}/items/{item['id']}", ok=(204,))
                removed += 1
        return removed


def check_drop(previous: dict, current: dict[str, int], rep: Report) -> bool:
    """True, wenn ein wichtiger Bestand auffällig geschrumpft ist."""
    alarm = False
    for otype in COUNT_CHECK:
        before, now = previous.get(otype), current.get(otype)
        if not before or now is None:
            continue
        drop = (before - now) / before * 100
        if drop >= DROP_ALERT_PERCENT:
            rep.warn(f"ACHTUNG: {rep.labels.get(otype, otype)} von {before} auf {now} gefallen "
                     f"(minus {drop:.0f} %). Bitte HubSpot prüfen.")
            alarm = True
    return alarm


# --------------------------------------------------------------------------
# Ablauf
# --------------------------------------------------------------------------

def main() -> int:
    started = time.monotonic()
    if not HUBSPOT_TOKEN:
        log.error("HUBSPOT_TOKEN fehlt")
        return 2
    upload_configured = all([MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET]) and (SP_SHARE_LINK or (SP_DRIVE_ID and SP_FOLDER_ID))
    if REQUIRE_UPLOAD and not upload_configured:
        log.error("Ablage in SharePoint ist nicht eingerichtet (MS_* und SP_* fehlen)")
        return 2
    if REQUIRE_UPLOAD and not ZIP_PASSWORD:
        log.error("ZIP_PASSWORD fehlt. Ohne Passwort wird kein Backup abgelegt.")
        return 2

    # Microsoft zuerst prüfen, damit ein Fehler nicht erst nach einer Stunde Export auffällt
    graph = None
    if upload_configured:
        graph = Graph(MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET)
        graph.resolve_target(SP_SHARE_LINK, SP_DRIVE_ID, SP_FOLDER_ID)

    root = WORK_DIR / BACKUP_NAME
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    rep = Report()
    hs = HubSpot(HUBSPOT_TOKEN, REQUESTS_PER_SECOND)

    owners = export_owners(hs, root, rep)
    custom = discover_custom_objects(hs, root, rep)
    custom_display = {o[0]: o[3] for o in custom}
    object_list = STANDARD_OBJECTS + [o[:3] for o in custom]

    objects: dict[str, ObjectInfo] = {}
    for otype, label, essential in object_list:
        info = export_object(hs, otype, label, essential, root, rep)
        if info:
            if otype in custom_display and custom_display[otype] in info.prop_by_name:
                info.name_fields = [custom_display[otype]]
            objects[otype] = info

    if "contacts" not in objects or "companies" not in objects:
        log.error("Kontakte oder Unternehmen nicht lesbar. Abbruch.")
        return 1

    pipelines = export_pipelines(hs, objects, root, rep)
    assoc = export_associations(hs, objects, root, rep)
    export_lists(hs, root, rep)
    if DOWNLOAD_FILES:
        # Anhänge werden nur einmal gesichert, danach nur neue (eigene Archive HubSpot-Anhaenge_*.zip)
        known: dict[str, str] = {}
        if graph:
            index_text = graph.read_text(ANHANG_INDEX)
            if index_text:
                try:
                    known = json.loads(index_text)
                except ValueError:
                    rep.warn("Anhangs-Verzeichnis in SharePoint unlesbar, alle Anhänge werden neu gesichert")
        new_dir = WORK_DIR / ANHANG_NAME
        if new_dir.exists():
            shutil.rmtree(new_dir)
        added = export_files(hs, objects, root, rep, known, new_dir, ANHANG_NAME)
        if added:
            files_zip = make_zip(new_dir, OUT_DIR / f"{ANHANG_NAME}.zip", ZIP_PASSWORD)
            log.info("Anhang-Archiv erstellt: %s (%.1f MB)", files_zip.name, files_zip.stat().st_size / 1e6)
            shutil.rmtree(new_dir, ignore_errors=True)
            if graph:
                graph.upload(files_zip)
                known.update(added)
                graph.upload_text(ANHANG_INDEX, json.dumps(known, ensure_ascii=False))
                if env("GITHUB_ACTIONS"):
                    files_zip.unlink(missing_ok=True)

    log.info("Schreibe CSV und Excel …")
    lookup = build_name_lookup(objects, root)
    for info in objects.values():
        write_object_tables(info, root, assoc, lookup, owners, pipelines, rep)

    # Plausibilität gegen den letzten Lauf
    alarm = False
    if graph:
        previous_text = graph.read_text("backup_zahlen.json")
        if previous_text:
            try:
                alarm = check_drop(json.loads(previous_text).get("counts", {}), rep.counts, rep)
            except ValueError:
                pass

    write_overview(root, rep, objects, hs.calls, time.monotonic() - started)
    write_json(root / "struktur" / "lauf.json", {
        "stand": NOW.isoformat(), "counts": rep.counts, "warnungen": rep.warnings,
        "ausgelassen": rep.skipped, "unvollstaendig": rep.incomplete, "abfragen": hs.calls})

    suffix = "_PRUEFEN" if (alarm or rep.incomplete) else ""
    zip_path = make_zip(root, OUT_DIR / f"{BACKUP_NAME}{suffix}.zip", ZIP_PASSWORD)
    log.info("ZIP erstellt: %s (%.1f MB)", zip_path.name, zip_path.stat().st_size / 1e6)

    if graph:
        graph.upload(zip_path)
        if alarm:
            log.warning("Bestand stark gefallen: alte Backups bleiben alle erhalten, Zählerstand wird nicht überschrieben")
        else:
            graph.upload_text("backup_zahlen.json", json.dumps(
                {"stand": NOW.isoformat(), "counts": rep.counts}, ensure_ascii=False, indent=2))
            removed = graph.cleanup(KEEP_DAILY, KEEP_MONTHLY)
            log.info("Alte Backups entfernt: %d", removed)
        if env("GITHUB_ACTIONS"):
            shutil.rmtree(root, ignore_errors=True)
            zip_path.unlink(missing_ok=True)

    summary = [f"Backup {BACKUP_NAME}{suffix}"] + [f"{rep.labels.get(k, k)}: {v}" for k, v in rep.counts.items()]
    summary += [f"Ausgelassen: {t}" for t in rep.skipped] + [f"Warnung: {t}" for t in rep.warnings]
    step_summary = env("GITHUB_STEP_SUMMARY")
    if step_summary:
        with open(step_summary, "a", encoding="utf-8") as f:
            f.write("## E+S HubSpot-Backup\n\n" + "\n".join(f"- {s}" for s in summary) + "\n")
    log.info("Fertig nach %.0f Minuten", (time.monotonic() - started) / 60)

    # Fehlercode sorgt dafür, dass GitHub eine Warn-Mail schickt
    return 1 if (alarm or rep.incomplete) else 0


if __name__ == "__main__":
    sys.exit(main())
