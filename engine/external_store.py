"""Independent external-price ledger; queries are read-only, snapshots append-only.

Prices are observations. Published dates do not invent validity periods, tax,
delivery scope, or a material/labour split. Raw source fields remain available.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import sqlite3
from contextlib import closing, contextmanager
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
DATABASE = ROOT / "data" / "외부가격_원장.db"
SCHEMA_VERSION = "construction-cost/external-price/v1"
_PRICE_SELECT = """SELECT p.*,s.metadata_json AS _snapshot_metadata_json,
 s.params_json AS _snapshot_params_json,
 s.operation AS _snapshot_operation,s.provider AS _snapshot_provider,
 q.source_network AS _source_network,q.test_only AS _test_only
 FROM prices p JOIN snapshots s ON p.snapshot_id=s.snapshot_id
 LEFT JOIN snapshot_provenance q ON q.snapshot_id=s.snapshot_id"""

_FIELDS = ["price_id", "snapshot_id", "provider", "source_origin", "item_id", "item_name",
           "specification", "unit", "unit_original", "price", "currency", "price_type",
           "published_at", "effective_start", "effective_end", "observed_at", "tax_basis",
           "delivery_condition", "region", "scope", "deleted", "department", "class_code",
           "task_category", "source_path", "source_sha256", "source_location", "cost_components",
           "raw_fields"]


def _json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def _utc():
    return datetime.now(timezone.utc).isoformat()


def sha256_file(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


@contextmanager
def _readonly():
    if not DATABASE.is_file():
        raise FileNotFoundError(f"External price ledger not imported: {DATABASE}")
    c = sqlite3.connect(DATABASE.resolve().as_uri() + "?mode=ro", uri=True)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA query_only=ON")
    try:
        yield c
    finally:
        c.close()


def ensure_store():
    """Writer entry point. Does not open or modify either pre-existing cost DB."""
    DATABASE.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(DATABASE, timeout=60)) as c, c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS ledger_metadata(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS snapshots(
          snapshot_id TEXT PRIMARY KEY,provider TEXT NOT NULL,operation TEXT NOT NULL,
          params_json TEXT NOT NULL,metadata_json TEXT NOT NULL,source_path TEXT NOT NULL,
          source_sha256 TEXT NOT NULL,observed_at TEXT NOT NULL,row_count INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS prices(
          price_id TEXT PRIMARY KEY,snapshot_id TEXT NOT NULL,provider TEXT NOT NULL,
          source_origin TEXT NOT NULL CHECK(source_origin='INDEPENDENT_EXTERNAL'),
          item_id TEXT,item_name TEXT,specification TEXT,unit TEXT,unit_original TEXT,
          price TEXT,currency TEXT NOT NULL,price_type TEXT NOT NULL,published_at TEXT,
          effective_start TEXT,effective_end TEXT,observed_at TEXT NOT NULL,
          tax_basis TEXT NOT NULL CHECK(tax_basis IN ('INCLUDED','EXCLUDED','UNKNOWN')),
          delivery_condition TEXT,region TEXT,scope TEXT,deleted INTEGER,department TEXT,
          class_code TEXT,task_category TEXT,source_path TEXT NOT NULL,source_sha256 TEXT NOT NULL,
          source_location TEXT NOT NULL,cost_components TEXT,raw_fields TEXT NOT NULL,
          FOREIGN KEY(snapshot_id) REFERENCES snapshots(snapshot_id)
        );
        CREATE INDEX IF NOT EXISTS prices_item ON prices(provider,item_id,published_at,deleted);
        CREATE INDEX IF NOT EXISTS prices_snapshot ON prices(snapshot_id);
        CREATE TABLE IF NOT EXISTS snapshot_provenance(
          snapshot_id TEXT PRIMARY KEY,source_network TEXT NOT NULL,
          test_only INTEGER NOT NULL CHECK(test_only IN (0,1)),
          FOREIGN KEY(snapshot_id) REFERENCES snapshots(snapshot_id)
        );
        CREATE TABLE IF NOT EXISTS historical_documents(
          document_id TEXT PRIMARY KEY,archive_path TEXT NOT NULL,archive_sha256 TEXT NOT NULL,
          member_name TEXT NOT NULL,extracted_path TEXT NOT NULL,sha256 TEXT NOT NULL,
          source_origin TEXT NOT NULL DEFAULT 'HISTORICAL_CASE',historical_period TEXT,
          document_format TEXT NOT NULL,metadata_json TEXT NOT NULL,text_path TEXT
        );
        """)
        c.execute("INSERT OR IGNORE INTO ledger_metadata VALUES (?,?)", ("schema_version", SCHEMA_VERSION))
        # Add provenance separately: source rows, raw responses and prior metadata stay immutable.
        for row in c.execute("SELECT snapshot_id,provider,operation,metadata_json FROM snapshots"):
            provenance = _snapshot_provenance(row[1], row[2], json.loads(row[3]))
            c.execute("INSERT OR IGNORE INTO snapshot_provenance VALUES (?,?,?)", (row[0], provenance["source_network"], int(provenance["test_only"])))
    return str(DATABASE)


def _text(v):
    return None if v is None or str(v).strip() == "" else str(v).strip()


def _pick(raw, *keys):
    for key in keys:
        if key in raw and _text(raw[key]) is not None:
            return _text(raw[key])
    return None


def _date(v):
    v = _text(v)
    if not v:
        return None
    try:
        if re.fullmatch(r"\d{8}", v):
            return datetime.strptime(v, "%Y%m%d").date().isoformat()
        # API timestamps preserve their reported calendar date, without mtime inference.
        return date.fromisoformat(v[:10]).isoformat()
    except ValueError:
        return None


def _decimal(v):
    v = _text(v)
    if v is None:
        return None
    try:
        d = Decimal(v.replace(",", ""))
        if not d.is_finite() or d < 0:
            return None
        return format(d, "f")
    except InvalidOperation:
        return None


def canonical_unit(v):
    original = _text(v)
    if original is None:
        return None
    aliases = {"M2": "m2", "M²": "m2", "㎡": "m2", "M3": "m3", "M³": "m3", "㎥": "m3",
               "M": "m", "KG": "kg", "TON": "ton", "TONS": "ton", "톤": "ton", "T": "ton",
               "EA": "ea", "개": "ea", "L": "l", "ℓ": "l", "리터": "l"}
    return aliases.get(original.upper(), original)


def _tax(value):
    v = _text(value)
    if v in ("INCLUDED", "포함", "부가가치세포함", "부가세포함", "Y"):
        return "INCLUDED"
    if v in ("EXCLUDED", "별도", "부가가치세별도", "부가세별도", "N"):
        return "EXCLUDED"
    return "UNKNOWN"


def _bool(v):
    v = _text(v)
    return True if v in ("Y", "true", "True", "1") else False if v in ("N", "false", "False", "0") else None


def normalize_pps(raw, provider="PPS_PRICE", source=None):
    """Official API fields or the supplied PPS TSV; no conditions are inferred."""
    source = dict(source or {})
    raw = dict(raw)
    shopping = provider == "PPS_SHOPPING"
    catalogue = provider == "PPS_CATALOG"
    unit = _pick(raw, "unit", "prdctUnit", "단위")
    # contractPrice is preserved; discounts/order prices are not silently substituted.
    price = None if catalogue else _decimal(_pick(raw, "prce", "cntrctPrceAmt", "가격", "price"))
    published = _date(_pick(raw, "nticeDt", "게시일자", "published_at"))
    if shopping:
        published = _date(_pick(raw, "chgDt", "rgstDt", "cntrctDate"))
    elif catalogue:
        published = _date(_pick(raw, "chgDt", "rgstDt"))
    costs = None
    if not shopping and not catalogue and all(_decimal(raw.get(k)) is not None for k in ("mtrlcst", "lbrcst", "gnrlexpns")):
        proposed = {"material": _decimal(raw["mtrlcst"]), "labor": _decimal(raw["lbrcst"]), "expense": _decimal(raw["gnrlexpns"])}
        if price is not None and sum(map(Decimal, proposed.values())) == Decimal(price):
            costs = proposed
    loc = str(source.get("source_location") or source.get("location") or "response.items")
    value = {
        "price_id": source.get("price_id"), "snapshot_id": source.get("snapshot_id"),
        "provider": provider, "source_origin": "INDEPENDENT_EXTERNAL",
        "item_id": _pick(raw, "prdctIdntNo", "물품식별번호", "item_id"),
        "item_name": _pick(raw, "prdctClsfcNoNm", "물품분류명", "item_name"),
        "specification": _pick(raw, "krnPrdctNm", "prdctSpecNm", "물품식별명", "specification"),
        "unit": canonical_unit(unit), "unit_original": unit, "price": price, "currency": "KRW",
        "price_type": "PRODUCT_METADATA" if catalogue else "PROCUREMENT_CONTRACT" if shopping else "PUBLIC_REFERENCE",
        "published_at": published,
        "effective_start": _date(_pick(raw, "cntrctBgnDate", "effective_start")),
        "effective_end": _date(_pick(raw, "cntrctEndDate", "effective_end")),
        "observed_at": source.get("observed_at") or _utc(),
        "tax_basis": _tax(_pick(raw, "vatYnNm", "vatAplDivNm", "부가가치세포함여부", "tax_basis")),
        "delivery_condition": _pick(raw, "dlvryCndtnNm", "prdctDlvryCndtnNm", "인도조건명", "delivery_condition"),
        "region": _pick(raw, "splyJrsdctRgnNm", "prdctSplyRgnNm", "공급관할지역명", "region"),
        "scope": _pick(raw, "etcCntnts", "scope") or _text(source.get("scope")),
        "deleted": _bool(_pick(raw, "dltYn", "삭제여부", "deleted")),
        "department": _pick(raw, "조달청담당부서명", "department") or _text(source.get("department")),
        "class_code": _pick(raw, "prdctClsfcNo", "물품분류번호", "class_code"),
        "task_category": _pick(raw, "bsnsDivNm", "포탈가격업무구분명", "task_category"),
        "source_path": str(source.get("source_path") or source.get("path") or ""),
        "source_sha256": source.get("source_sha256") or source.get("sha256") or "",
        "source_location": loc, "cost_components": costs, "raw_fields": raw,
    }
    if not value["price_id"]:
        identity = _json({"provider": provider, "source_sha256": value["source_sha256"], "location": loc, "raw": raw})
        value["price_id"] = "EXT:" + hashlib.sha256(identity.encode("utf-8")).hexdigest()[:32]
    return value


def _safe_params(params):
    secret_pattern = re.compile(r"key|token|password|secret|credential|auth", re.I)
    return {str(k): "[REDACTED]" if secret_pattern.search(str(k)) else v for k, v in (params or {}).items()}


def _snapshot_provenance(provider, operation, metadata):
    network = str(metadata.get("network") or "UNKNOWN").upper()
    test_only = bool(metadata.get("test_only") or metadata.get("purpose") == "TEST_ONLY" or network == "FIXTURE" or "fixture" in str(operation).lower())
    if provider == "PPS_CSV" and operation == "import_full_tsv_export" and network == "UNKNOWN":
        network = "FILE_IMPORT"
    return {"source_network": network, "test_only": test_only}


def _insert_rows(c, rows, snapshot_id, provider, source_path, source_sha256, observed_at):
    placeholders = ",".join("?" for _ in _FIELDS)
    sql = f"INSERT INTO prices ({','.join(_FIELDS)}) VALUES ({placeholders})"
    batch = []
    count = 0
    for index, row in enumerate(rows, 1):
        row = dict(row)
        if row.get("source_origin", "INDEPENDENT_EXTERNAL") != "INDEPENDENT_EXTERNAL":
            raise ValueError("A bill-derived price cannot be ingested into the independent external ledger")
        row["source_origin"] = "INDEPENDENT_EXTERNAL"
        row.update(snapshot_id=snapshot_id, provider=provider,
                   source_path=source_path,
                   source_sha256=source_sha256,
                   observed_at=row.get("observed_at") or observed_at)
        # A refreshed observation has its own snapshot identity even if content is equal.
        row["price_id"] = f"EXT:{snapshot_id.removeprefix('SNAP:')}:{index}"
        row.setdefault("source_location", f"response.items[{index - 1}]")
        row.setdefault("currency", "KRW")
        row.setdefault("price_type", "PUBLIC_REFERENCE")
        row.setdefault("tax_basis", "UNKNOWN")
        row.setdefault("raw_fields", {})
        if row.get("price") is not None:
            valid = _decimal(row["price"])
            if valid is None:
                raise ValueError("Nonfinite/negative/malformed external price")
            row["price"] = valid
        vals = [(_json(row.get(k)) if k in ("raw_fields", "cost_components") and row.get(k) is not None else row.get(k)) for k in _FIELDS]
        batch.append(vals)
        count += 1
        if len(batch) >= 2000:
            c.executemany(sql, batch)
            batch = []
    if batch:
        c.executemany(sql, batch)
    return count


def save_snapshot(provider, operation, params, raw_bytes, rows, metadata):
    """Append immutable raw API response and rows in one transaction. Idempotent by content."""
    ensure_store()
    metadata = dict(metadata or {})
    params = _safe_params(params)
    # API clients must sanitize echoed credentials before this immutable evidence boundary.
    for name in ("DATA_GO_KR_SERVICE_KEY", "OPINET_API_KEY"):
        secret = os.environ.get(name)
        if secret and secret.encode("utf-8") in raw_bytes:
            raise ValueError("Response contains credential; sanitize before preserving evidence")
    source_sha = hashlib.sha256(raw_bytes).hexdigest()
    ident = _json([provider, operation, params, source_sha, metadata.get("observed_at")])
    snapshot_id = "SNAP:" + hashlib.sha256(ident.encode("utf-8")).hexdigest()[:32]
    folder = ROOT / "evidence" / "api_raw" / provider
    folder.mkdir(parents=True, exist_ok=True)
    source_path = folder / (snapshot_id.removeprefix("SNAP:") + ".response")
    if source_path.exists():
        if sha256_file(source_path) != source_sha:
            raise ValueError("Existing immutable response hash conflict")
    else:
        with source_path.open("xb") as f:
            f.write(raw_bytes)
    observed_at = metadata.get("observed_at") or _utc()
    with closing(sqlite3.connect(DATABASE, timeout=60)) as c, c:
        old = c.execute("SELECT row_count,source_path,source_sha256 FROM snapshots WHERE snapshot_id=?", (snapshot_id,)).fetchone()
        if old:
            return {"snapshot_id": snapshot_id, "row_count": old[0], "source_path": old[1], "source_sha256": old[2]}
        c.execute("INSERT INTO snapshots VALUES (?,?,?,?,?,?,?,?,?)", (snapshot_id, provider, operation, _json(params), _json(metadata), str(source_path), source_sha, observed_at, 0))
        provenance = _snapshot_provenance(provider, operation, metadata)
        c.execute("INSERT INTO snapshot_provenance VALUES (?,?,?)", (snapshot_id, provenance["source_network"], int(provenance["test_only"])))
        count = _insert_rows(c, rows, snapshot_id, provider, str(source_path), source_sha, observed_at)
        c.execute("UPDATE snapshots SET row_count=? WHERE snapshot_id=?", (count, snapshot_id))
    return {"snapshot_id": snapshot_id, "row_count": count, "source_path": str(source_path), "source_sha256": source_sha}


def _unpack(row):
    if row is None:
        return None
    d = dict(row)
    for key in ("raw_fields", "cost_components"):
        if key in d:
            d[key] = json.loads(d[key]) if d[key] is not None else None
    if d.get("deleted") is not None:
        d["deleted"] = bool(d["deleted"])
    d["deletion_status"] = "DELETED" if d.get("deleted") is True else "NOT_DELETED" if d.get("deleted") is False else "NOT_EXPOSED"
    if "_snapshot_metadata_json" in d:
        metadata = json.loads(d.pop("_snapshot_metadata_json"))
        snapshot_operation = d.pop("_snapshot_operation")
        provenance = _snapshot_provenance(d.pop("_snapshot_provider"), snapshot_operation, metadata)
        network = d.pop("_source_network", None)
        test_only = d.pop("_test_only", None)
        # Fixture in immutable snapshot metadata cannot be neutralized by a derived flag.
        d["source_network"] = network or provenance["source_network"]
        d["test_only"] = provenance["test_only"] or bool(test_only)
        d["operation"] = snapshot_operation
        safe = _safe_params(json.loads(d.pop("_snapshot_params_json", "{}")))
        excluded = {"pageno", "numofrows", "type", "out"}
        sensitive = re.compile(r"key|token|password|secret|credential|auth|^code$", re.I)
        d["refresh_params"] = {k: v for k, v in safe.items() if k.lower() not in excluded and not sensitive.search(k)}
    return d


def get_price(price_id):
    with _readonly() as c:
        return _unpack(c.execute(_PRICE_SELECT + " WHERE p.price_id=?", (price_id,)).fetchone())


def search_prices(text="", provider="", limit=100):
    if not 1 <= int(limit) <= 1000:
        raise ValueError("limit must be 1..1000")
    where, params = [], []
    if provider:
        where.append("p.provider=?")
        params.append(provider)
    if text:
        where.append("(p.item_id LIKE ? ESCAPE '\\' OR p.item_name LIKE ? ESCAPE '\\' OR p.specification LIKE ? ESCAPE '\\')")
        term = "%" + str(text).replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
        params.extend([term] * 3)
    clause = " WHERE " + " AND ".join(where) if where else ""
    with _readonly() as c:
        total = c.execute("SELECT count(*) FROM prices p" + clause, params).fetchone()[0]
        rows = c.execute(_PRICE_SELECT + clause + " ORDER BY COALESCE(deleted,0),published_at DESC,price_id LIMIT ?", params + [int(limit)]).fetchall()
    return {"total": total, "rows": [_unpack(row) for row in rows]}


def candidates_for_price(price_id):
    selected = get_price(price_id)
    if selected is None or not selected.get("item_id"):
        return []
    with _readonly() as c:
        return [_unpack(row) for row in c.execute(_PRICE_SELECT + " WHERE p.provider=? AND p.item_id=? ORDER BY published_at DESC,p.price_id", (selected["provider"], selected["item_id"]))]


def _verified_opinet_tax_correction(rows):
    """Resolve an evidenced UNKNOWN→INCLUDED metadata correction, not an economic conflict."""
    if not rows or any(r["provider"] != "OPINET" for r in rows):
        return None
    economic_keys = ("item_id", "published_at", "price", "unit", "specification", "delivery_condition", "region", "scope", "department", "cost_components")
    economics = {_json({k: r.get(k) for k in economic_keys}) for r in rows}
    raw_payloads = {_json({k: v for k, v in r["raw_fields"].items() if k != "_tax_basis_evidence"}) for r in rows}
    if len(economics) != 1 or len(raw_payloads) != 1 or {r["tax_basis"] for r in rows} != {"UNKNOWN", "INCLUDED"}:
        return None
    verified = []
    for r in rows:
        if r["tax_basis"] != "INCLUDED":
            continue
        evidence = r["raw_fields"].get("_tax_basis_evidence")
        if not isinstance(evidence, dict):
            continue
        parsed = urlparse(str(evidence.get("url", "")))
        if parsed.scheme != "https" or parsed.hostname != "www.opinet.co.kr":
            continue
        if evidence.get("prior_tax_basis") != "UNKNOWN" or evidence.get("tax_basis") != "INCLUDED" or "부가세포함" not in str(evidence.get("claim", "")):
            continue
        path = Path(str(evidence.get("source_path", "")))
        try:
            if not path.is_file() or sha256_file(path) != evidence.get("source_sha256"):
                continue
        except OSError:
            continue
        verified.append(r)
    if not verified:
        return None
    most_recent = max(r["observed_at"] for r in verified)
    if most_recent < max(r["observed_at"] for r in rows if r["tax_basis"] == "UNKNOWN"):
        return None
    return [r for r in verified if r["observed_at"] == most_recent]


def _query_scope(row):
    """Search breadth is not a purchasing condition for an already exact item ID."""
    params = row.get("refresh_params") or {}
    discovery = {"prdctidntno", "prdctclsfcno", "dtilprdctclsfcno", "prdctclsfcnonm", "prdctclsfcnoengnm", "krnprdctnm", "text", "query", "prodcd"}
    return {k: _text(v) for k, v in params.items() if k.lower() not in discovery}


def _shopping_scope(row):
    if row.get("provider") != "PPS_SHOPPING":
        return {}
    raw = row.get("raw_fields") or {}
    # Contract/supplier identifies a source. Price/discount amounts remain version
    # values and must never be used to discard a later increase or changed condition.
    keys = ("shopngCntrctNo", "shopngCntrctSno", "cntrctNo", "cntrctSno", "cntrctCorpNm", "cntrctCorpBizno")
    return {k: _text(raw.get(k)) for k in keys}


def _compatible_source_scope(selected, candidate):
    keys = ("region", "unit", "specification", "department", "operation")
    for key in keys:
        left, right = _text(selected.get(key)), _text(candidate.get(key))
        if key == "unit":
            left, right = canonical_unit(left), canonical_unit(right)
        if left is not None and right is not None and left != right:
            return False
    for left, right in ((_query_scope(selected), _query_scope(candidate)), (_shopping_scope(selected), _shopping_scope(candidate))):
        for key in set(left) & set(right):
            if left[key] is not None and right[key] is not None and left[key] != right[key]:
                return False
    # Missing scope identity remains comparable so that ambiguity is reported;
    # an unknown region/supplier cannot quietly certify a known one.
    return True


def _economic_signature(row):
    keys = ("specification", "unit", "price", "tax_basis", "delivery_condition", "region", "scope", "department", "cost_components", "operation")
    result = {k: row.get(k) for k in keys}
    result["contract_identity"] = _shopping_scope(row)
    if row.get("provider") == "PPS_SHOPPING":
        raw = row.get("raw_fields") or {}
        result["conditional_price_terms"] = {k: _text(raw.get(k)) for k in ("orderCalclPrceAmt", "dscntAmt", "dscntBgnDate", "dscntEndDate", "prcrmntFeeInclsnCntnts")}
    return _json(result)


def resolve_latest(price_id, pricing_date, department=None, allow_test_only=False):
    """Validate an explicit ID against latest compatible data; never guess among conflicts."""
    day = date.fromisoformat(str(pricing_date)).isoformat()
    selected = get_price(price_id)
    if selected is None:
        return {"status": "BLOCKED", "selected": None, "conflicts": [], "reason": "PRICE_ID_NOT_FOUND"}
    if selected.get("test_only") and not allow_test_only:
        return {"status": "BLOCKED", "selected": None, "conflicts": [], "reason": "FIXTURE_PRICE_NOT_ALLOWED_FOR_ACTUAL"}
    if selected.get("source_network") not in (None, "LIVE", "FILE_IMPORT") and not allow_test_only:
        return {"status": "BLOCKED", "selected": None, "conflicts": [], "reason": "UNVERIFIED_SNAPSHOT_PROVENANCE"}
    candidates = [r for r in candidates_for_price(price_id) if r.get("published_at") and r["published_at"] <= day and (allow_test_only or not r.get("test_only"))]
    candidates = [r for r in candidates if _compatible_source_scope(selected, r)]
    if department is not None:
        candidates = [r for r in candidates if r.get("department") == department]
    if not candidates:
        return {"status": "BLOCKED", "selected": None, "conflicts": [], "reason": "NO_DATED_NONDELETED_PRICE"}
    latest_date = max(r["published_at"] for r in candidates)
    latest = [r for r in candidates if r["published_at"] == latest_date]
    # Availability is checked AFTER the economic date. A later deletion/missing price
    # invalidates the older value instead of resurrecting it as the latest usable row.
    scope_keys = ("specification", "unit", "region", "delivery_condition", "scope", "department")
    selected_scope = tuple(selected.get(k) for k in scope_keys)
    same_scope = [r for r in latest if tuple(r.get(k) for k in scope_keys) == selected_scope]
    tombstones = [r for r in same_scope if r.get("deleted") is True]
    if tombstones:
        return {"status": "BLOCKED", "selected": None, "conflicts": tombstones, "reason": "LATEST_PRICE_DELETED"}
    latest = [r for r in latest if r.get("deleted") is not True]
    if not latest:
        return {"status": "BLOCKED", "selected": None, "conflicts": same_scope, "reason": "LATEST_SCOPE_UNAVAILABLE"}
    if any(r.get("price") is None for r in latest):
        return {"status": "BLOCKED", "selected": None, "conflicts": latest, "reason": "LATEST_PRICE_MISSING"}
    signatures = {_economic_signature(r) for r in latest}
    correction = None
    if len(signatures) != 1:
        corrected = _verified_opinet_tax_correction(latest)
        if corrected:
            correction = {"reason": "VERIFIED_UNKNOWN_TO_INCLUDED_TAX_METADATA_CORRECTION",
                          "prior_price_ids": [r["price_id"] for r in latest if r["tax_basis"] == "UNKNOWN"],
                          "evidence": corrected[0]["raw_fields"]["_tax_basis_evidence"]}
            latest = corrected
            signatures = {_economic_signature(r) for r in latest}
    if len(signatures) != 1:
        return {"status": "BLOCKED", "selected": None, "conflicts": latest, "reason": "LATEST_PRICE_CONFLICT"}
    latest_row = next((r for r in latest if r["price_id"] == price_id), latest[0])
    if selected.get("deleted") is True or selected.get("published_at") != latest_date:
        return {"status": "BLOCKED", "selected": None, "conflicts": latest, "reason": "STALE_OR_DELETED_EXPLICIT_PRICE"}
    if _economic_signature(selected) != next(iter(signatures)):
        return {"status": "BLOCKED", "selected": None, "conflicts": latest, "reason": "SELECTED_SCOPE_DIFFERS_FROM_LATEST"}
    result = {"status": "OK", "selected": latest_row, "conflicts": [], "reason": "LATEST_DATED_COMPATIBLE_PRICE"}
    if correction:
        result["metadata_corrections"] = [correction]
    return result


def ledger_status():
    with _readonly() as c:
        return {"database": str(DATABASE), "schema_version": SCHEMA_VERSION,
                "row_count": c.execute("SELECT count(*) FROM prices").fetchone()[0],
                "providers": dict(c.execute("SELECT provider,count(*) FROM prices GROUP BY provider")),
                "snapshot_count": c.execute("SELECT count(*) FROM snapshots").fetchone()[0],
                "historical_documents": c.execute("SELECT count(*) FROM historical_documents").fetchone()[0],
                "quick_check": c.execute("PRAGMA quick_check").fetchone()[0]}


def latest_snapshot(provider, live_only=True):
    """Read persisted successful API evidence; never probe network or report fixtures as live."""
    if not DATABASE.is_file():
        return None
    with _readonly() as c:
        for row in c.execute("SELECT * FROM snapshots WHERE provider=? ORDER BY observed_at DESC,snapshot_id", (provider,)):
            metadata = json.loads(row["metadata_json"])
            if live_only and metadata.get("network") != "LIVE":
                continue
            return {"snapshot_id": row["snapshot_id"], "operation": row["operation"],
                    "observed_at": row["observed_at"], "metadata": metadata,
                    "row_count": row["row_count"], "source_path": row["source_path"],
                    "source_sha256": row["source_sha256"]}
    return None
