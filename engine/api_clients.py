"""Official PPS/Opinet clients. Credentials never enter the evidence store.

An API snapshot is an observation, not a quotation or an automatic price choice.
PPS pagination is bounded and reports completeness rather than implying a full
catalogue from a first-page sample. No transport retry or fallback price occurs.
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
KST = timezone(timedelta(hours=9))
SERVICES = {
    "PPS_PRICE": {
        "endpoint": "https://apis.data.go.kr/1230000/ao/PriceInfoService",
        "credential": "PPS", "documentation": "https://www.data.go.kr/data/15129415/openapi.do",
        "operations": {
            "토목": "getPriceInfoListFcltyCmmnMtrilEngrk",
            "건축": "getPriceInfoListFcltyCmmnMtrilBildng",
            "기계설비": "getPriceInfoListFcltyCmmnMtrilMchnEqp",
            "전기정보통신": "getPriceInfoListFcltyCmmnMtrilElctyIrmc",
        },
        "allowed": {"prdctClsfcNo", "prdctClsfcNoNm", "prdctIdntNo", "krnPrdctNm"},
        "price_type": "PUBLIC_REFERENCE",
    },
    "PPS_SHOPPING": {
        "endpoint": "https://apis.data.go.kr/1230000/at/ShoppingMallPrdctInfoService",
        "credential": "PPS", "documentation": "https://www.data.go.kr/data/15129471/openapi.do",
        "operations": {"MAS": "getMASCntrctPrdctInfoList", "일반단가": "getUcntrctPrdctInfoList", "제3자단가": "getThptyUcntrctPrdctInfoList"},
        "allowed": {"prdctClsfcNoNm", "prdctIdntNo", "cntrctCorpNm", "chgDtBgnDt", "chgDtEndDt", "rgstDtBgnDt", "rgstDtEndDt", "prodctCertYn", "inqryDiv"},
        "price_type": "PROCUREMENT_CONTRACT",
    },
    "PPS_CATALOG": {
        "endpoint": "https://apis.data.go.kr/1230000/ao/ThngListInfoService02",
        "credential": "PPS", "documentation": "https://www.data.go.kr/data/15129417/openapi.do",
        "operations": {"품목": "getThngPrdnmLocplcAccotListInfoInfoPrdlstSearch02", "속성": "getPrdctIndvAtrbInfoList02"},
        "allowed": {"dtilPrdctClsfcNo", "prdctIdntNo", "prdctClsfcNoEngNm", "prdctClsfcNoNm", "krnPrdctNm", "inqryBgnDt", "inqryEndDt", "chgPrdBgnDt", "chgPrdEndDt"},
        "price_type": "PRODUCT_METADATA",
    },
    "OPINET": {
        "endpoint": "https://www.opinet.co.kr/api", "credential": "OPINET",
        "documentation": "https://www.opinet.co.kr/user/custapi/custApiInfo.do",
        "operations": {"전국": "avgAllPrice.do", "시도": "avgSidoPrice.do", "최근7일": "avgRecentPrice.do"},
        "allowed": {"sido", "prodcd", "date"}, "price_type": "FUEL_RETAIL",
    },
}
ALIASES = {"PRICE": "PPS_PRICE", "PPS": "PPS_PRICE", "SHOPPING": "PPS_SHOPPING", "CATALOG": "PPS_CATALOG", "FUEL": "OPINET"}
SUCCESS_CODES = {"00", "0", "000", "0000", "NORMAL_SERVICE", "NORMAL_CODE"}
GATEWAY_ERRORS = {"SERVICE_KEY_IS_NULL", "SERVICE_ACCESS_DENIED_ERROR", "PERMISSION_DENIED",
                  "SERVICE_KEY_IS_NOT_REGISTERED_ERROR", "DEADLINE_HAS_EXPIRED_ERROR",
                  "LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR",
                  "LIMITED_NUMBER_OF_SERVICE_REQUESTS_PER_SECOND_EXCEEDS_ERROR",
                  "INVALID_REQUEST_PARAMETER_ERROR", "NO_OPENAPI_SERVICE_ERROR"}


class ApiError(Exception):
    """Only fixed, safe context is exposed; original exception URLs are discarded."""
    def __init__(self, code: str, provider: str = "", detail: str = ""):
        self.code, self.provider, self.detail = code, provider, detail
        super().__init__(code + (": " + detail if detail else ""))

    def as_dict(self):
        return {"status": "ERROR", "error_code": self.code, "provider": self.provider, "detail": self.detail, "fallback_applied": False}


def now_iso():
    return datetime.now(KST).isoformat(timespec="seconds")


def provider_name(provider):
    value = str(provider).upper()
    value = ALIASES.get(value, value)
    if value not in SERVICES:
        raise ApiError("UNSUPPORTED_PROVIDER")
    return value


def _xml_value(node):
    children = list(node)
    if not children:
        return (node.text or "").strip()
    out = {}
    for child in children:
        name = child.tag.rsplit("}", 1)[-1]
        value = _xml_value(child)
        if name in out:
            if not isinstance(out[name], list):
                out[name] = [out[name]]
            out[name].append(value)
        else:
            out[name] = value
    return out


def _list(value):
    if value in (None, "", []):
        return []
    if isinstance(value, list):
        if any(not isinstance(x, dict) for x in value):
            raise ApiError("INVALID_ITEM_SHAPE")
        return value
    if isinstance(value, dict):
        if "item" in value:
            return _list(value["item"])
        return [value] if value else []
    raise ApiError("INVALID_ITEM_SHAPE")


def parse_response(raw_bytes: bytes, provider: str):
    """Decode JSON/XML and require an explicit service success envelope."""
    provider = provider_name(provider)
    if len(raw_bytes) > 20_000_000:
        raise ApiError("RESPONSE_TOO_LARGE", provider)
    try:
        text = raw_bytes.decode("utf-8-sig")
        if text.lstrip().startswith("<"):
            # Entity declarations are not used by any documented response.
            if "<!DOCTYPE" in text.upper() or "<!ENTITY" in text.upper():
                raise ApiError("UNSAFE_XML", provider)
            node = ET.fromstring(text)
            value = {node.tag.rsplit("}", 1)[-1]: _xml_value(node)}
        else:
            value = json.loads(text, parse_float=str)
    except ApiError:
        raise
    except Exception:
        raise ApiError("INVALID_RESPONSE_FORMAT", provider) from None
    if not isinstance(value, dict):
        raise ApiError("INVALID_RESPONSE_ENVELOPE", provider)
    if provider == "OPINET":
        result = value.get("RESULT", value.get("result"))
        if not isinstance(result, dict) or "OIL" not in result:
            raise ApiError("SERVICE_RESPONSE_ERROR", provider)
        rows = _list(result["OIL"])
        return {"rows": rows, "total_count": len(rows), "page_no": 1, "num_of_rows": len(rows), "result_code": "SUCCESS"}
    envelope = value.get("response", value)
    if not isinstance(envelope, dict):
        raise ApiError("INVALID_RESPONSE_ENVELOPE", provider)
    header = envelope.get("header")
    if not isinstance(header, dict):
        # Gateway errors have OpenAPI_ServiceResponse/cmmMsgHeader instead.
        raise ApiError("SERVICE_RESPONSE_ERROR", provider)
    code = str(header.get("resultCode", ""))
    if code not in SUCCESS_CODES:
        safe_code = code if code.isalnum() and len(code) <= 20 else "UNKNOWN"
        raise ApiError("SERVICE_ERROR_" + safe_code, provider)
    body = envelope.get("body")
    if not isinstance(body, dict):
        raise ApiError("MISSING_RESPONSE_BODY", provider)
    rows = _list(body.get("items"))
    counts = {}
    for k, target in (("totalCount", "total_count"), ("pageNo", "page_no"), ("numOfRows", "num_of_rows")):
        try:
            counts[target] = int(body[k]) if body.get(k) not in (None, "") else None
        except (TypeError, ValueError):
            raise ApiError("INVALID_PAGINATION", provider) from None
        if counts[target] is not None and counts[target] < 0:
            raise ApiError("INVALID_PAGINATION", provider)
    if counts["total_count"] == 0 and rows:
        raise ApiError("INCONSISTENT_RESPONSE_COUNT", provider)
    return {"rows": rows, "result_code": code, **counts}


def _decimal(value):
    try:
        d = Decimal(str(value).replace(",", "").strip())
        return str(d) if d.is_finite() and d >= 0 else None
    except (InvalidOperation, ValueError):
        return None


def normalize_opinet(raw: dict, source: dict | None = None):
    source = source or {}
    product = str(raw.get("PRODCD", ""))
    names = {"B034": "고급휘발유", "B027": "보통휘발유", "D047": "자동차용경유", "C004": "실내등유", "K015": "자동차용부탄"}
    day = str(raw.get("TRADE_DT", raw.get("DATE", "")))
    published = None
    if len(day) == 8 and day.isdigit():
        try:
            published = datetime.strptime(day, "%Y%m%d").date().isoformat()
        except ValueError:
            pass
    canonical = json.dumps({"raw": raw, "operation": source.get("operation")}, ensure_ascii=False, sort_keys=True).encode()
    # Underscore-prefixed field is explicitly a normalization evidence overlay;
    # the API response bytes remain untouched in the immutable raw snapshot.
    raw_with_evidence = dict(raw)
    raw_with_evidence["_tax_basis_evidence"] = {
        "url": "https://www.opinet.co.kr/user/dopospdrg/dopOsPdrgSelect.do",
        "claim": "전체 주유소의 제품별 평균 판매가격(부가세포함), 원/리터",
        "source_path": str(ROOT / "evidence/api_docs/OPINET_avg_metadata.html"),
        "source_sha256": "2cc223dd68a1ef39de612449c5f83f7b0125a30f082dd80aa94ca47c611ff1ce",
        "observed_at": source.get("observed_at"),
        "prior_tax_basis": "UNKNOWN", "tax_basis": "INCLUDED",
    }
    return {
        "price_id": "API:OPINET:" + hashlib.sha256(canonical).hexdigest()[:24],
        "provider": "OPINET", "source_origin": "INDEPENDENT_EXTERNAL", "item_id": product,
        "item_name": raw.get("PRODNM") or names.get(product, product), "specification": product,
        "unit": "L", "price": _decimal(raw.get("PRICE")), "currency": "KRW", "price_type": "FUEL_RETAIL",
        "published_at": published, "effective_start": None, "effective_end": None,
        "observed_at": source.get("observed_at"), "tax_basis": "INCLUDED",
        "delivery_condition": "주유소 판매가격 평균; 현장 운반조건 미포함/미확인",
        "region": raw.get("SIDONM") or source.get("region") or "전국",
        "scope": "연료 소매 평균가격; 장비 시간당 운전비 또는 현장 납품 견적이 아님",
        "cost_components": None, "source_path": source.get("path"), "source_sha256": source.get("sha256"),
        "source_location": source.get("location"), "raw_fields": raw_with_evidence, "department": "연료", "deleted": False,
        "tax_evidence": "https://www.opinet.co.kr/user/dopospdrg/dopOsPdrgSelect.do ; 전체 주유소 제품별 평균 판매가격(부가세포함), 원/리터",
    }


def _credential(provider):
    try:
        from api_credentials import get_secret
        value = get_secret(SERVICES[provider]["credential"])
    except Exception:
        raise ApiError("CREDENTIAL_UNAVAILABLE", provider) from None
    if not value or not isinstance(value, str):
        raise ApiError("CREDENTIAL_UNAVAILABLE", provider)
    return value.strip()


def _reserve_call(provider):
    """Conservative local attempt budget; actual provider limits remain authoritative."""
    group = SERVICES[provider]["credential"]
    cap = 900 if group == "PPS" else 250
    folder = ROOT / "data"
    folder.mkdir(exist_ok=True)
    path, lock = folder / "api_usage.json", folder / "api_usage.lock"
    fd = None
    try:
        for _ in range(20):
            try:
                fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                break
            except FileExistsError:
                time.sleep(0.1)
        if fd is None:
            raise ApiError("CALL_BUDGET_BUSY", provider)
        os.close(fd)
        day = datetime.now(KST).date().isoformat()
        state = json.loads(path.read_text("utf-8")) if path.exists() else {}
        if state.get("date") != day:
            state = {"date": day, "PPS": 0, "OPINET": 0}
        if int(state.get(group, 0)) >= cap:
            raise ApiError("LOCAL_DAILY_BUDGET_EXHAUSTED", provider)
        state[group] = int(state.get(group, 0)) + 1
        state["last_attempt_at"] = now_iso()
        temp = path.with_name("api_usage_" + str(os.getpid()) + ".tmp")
        temp.write_text(json.dumps(state, ensure_ascii=False, indent=2), "utf-8")
        os.replace(temp, path)
    except ApiError:
        raise
    except Exception:
        raise ApiError("CALL_BUDGET_STATE_ERROR", provider) from None
    finally:
        if fd is not None:
            try:
                lock.unlink()
            except FileNotFoundError:
                pass


def _request(provider, operation, safe_params, timeout=20):
    secret = _credential(provider)
    query = dict(safe_params)
    # User keys may be already percent encoded; exactly one URL encoding pass.
    secret = urllib.parse.unquote(secret)
    query["code" if provider == "OPINET" else "serviceKey"] = secret
    url = SERVICES[provider]["endpoint"] + "/" + operation + "?" + urllib.parse.urlencode(query)
    _reserve_call(provider)
    try:
        request = urllib.request.Request(url, headers={"Accept": "application/json, application/xml", "User-Agent": "ConstructionCostCodex/2026.10"})
        with urllib.request.urlopen(request, timeout=timeout) as response:
            data = response.read(20_000_001)
        # Do not preserve server-reflected credentials, even as raw evidence.
        for variant in (secret, urllib.parse.quote(secret, safe=""), urllib.parse.quote_plus(secret)):
            if variant and variant.encode() in data:
                raise ApiError("CREDENTIAL_REFLECTED_IN_RESPONSE", provider)
        return data
    except ApiError:
        raise
    except urllib.error.HTTPError as exc:
        # Only allowlisted gateway error constants survive. Body/URL/messages
        # may echo credentials and must never become output or raw evidence.
        gateway_code = None
        try:
            value = json.loads(exc.read(100_000))
            candidate = value.get("OpenAPI_ServiceResponse", {}).get("cmmMsgHeader", {}).get("errMsg")
            if candidate in GATEWAY_ERRORS:
                gateway_code = candidate
        except Exception:
            pass
        raise ApiError("HTTP_" + str(exc.code), provider, gateway_code or "") from None
    except (TimeoutError, urllib.error.URLError):
        raise ApiError("NETWORK_OR_TIMEOUT", provider) from None
    except Exception:
        raise ApiError("TRANSPORT_ERROR", provider) from None


def _operation(provider, department, operation):
    service = SERVICES[provider]
    if operation:
        if operation not in service["operations"].values():
            raise ApiError("UNSUPPORTED_OPERATION", provider)
        return operation
    if provider == "PPS_PRICE":
        department = {"기계": "기계설비", "전기": "전기정보통신"}.get(department, department)
        if department not in service["operations"]:
            raise ApiError("UNSUPPORTED_DEPARTMENT", provider)
        return service["operations"][department]
    return next(iter(service["operations"].values()))


def fetch_prices(provider="PPS_PRICE", text="", item_id="", department="토목", page_size=10,
                 max_pages=1, params=None, operation=None, timeout=20, transport=None,
                 persist=True, sleep_seconds=1.1):
    """Fetch a bounded observation. transport enables fixture tests without keys.

    Result includes all saved price IDs plus complete=false whenever the query's
    later pages are not fetched. It never claims first-page records are latest.
    """
    provider = provider_name(provider)
    if not isinstance(page_size, int) or not 1 <= page_size <= 100:
        raise ApiError("PAGE_SIZE_OUT_OF_RANGE", provider)
    if not isinstance(max_pages, int) or not 1 <= max_pages <= 10:
        raise ApiError("MAX_PAGES_OUT_OF_RANGE", provider)
    service = SERVICES[provider]
    operation = _operation(provider, department, operation)
    safe_params = dict(params or {})
    if any(k not in service["allowed"] for k in safe_params):
        raise ApiError("UNSUPPORTED_OR_SECRET_PARAMETER", provider)
    if any(isinstance(v, (dict, list)) or len(str(v)) > 300 for v in safe_params.values()):
        raise ApiError("INVALID_PARAMETER_VALUE", provider)
    if provider == "OPINET":
        safe_params["out"] = "json"
        if item_id and operation != "avgAllPrice.do":
            safe_params["prodcd"] = item_id
    else:
        safe_params.update(type="json", numOfRows=str(page_size))
        if text:
            safe_params["prdctClsfcNoNm"] = text
        if item_id:
            safe_params["prdctIdntNo"] = item_id
    request_fn = transport or _request
    snapshots, normalized, pages, total, complete, seen = [], [], [], None, False, set()
    observed = now_iso()
    warnings = []
    for page_no in range(1, (1 if provider == "OPINET" else max_pages) + 1):
        query = dict(safe_params)
        if provider != "OPINET":
            query["pageNo"] = str(page_no)
        raw = request_fn(provider, operation, query, timeout=timeout)
        payload = parse_response(raw, provider)
        if provider != "OPINET" and payload["page_no"] not in (None, page_no):
            raise ApiError("PAGINATION_PAGE_MISMATCH", provider)
        if total is not None and payload["total_count"] is not None and total != payload["total_count"]:
            warnings.append("TOTAL_COUNT_CHANGED_DURING_FETCH")
        total = payload["total_count"]
        page_rows = payload["rows"]
        page_digest = hashlib.sha256(json.dumps(page_rows, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        if page_rows and page_digest in seen:
            raise ApiError("PAGINATION_REPEATED_PAGE", provider)
        seen.add(page_digest)
        source = {"observed_at": observed, "operation": operation, "department": department if provider == "PPS_PRICE" else None,
                  "location": "response.body.items; page=" + str(page_no), "region": safe_params.get("sido"),
                  "sha256": hashlib.sha256(raw).hexdigest()}
        if provider == "OPINET":
            records = [normalize_opinet(x, source) for x in page_rows]
        else:
            from external_store import normalize_pps
            records = [normalize_pps(x, provider=provider, source=source) for x in page_rows]
        fetched_count = sum(p["row_count"] for p in pages) + len(page_rows)
        complete = provider == "OPINET" or (total is not None and fetched_count >= total)
        if not page_rows and total not in (None, 0) and fetched_count < total:
            warnings.append("EMPTY_PAGE_BEFORE_TOTAL_COUNT")
        metadata = {"observed_at": observed, "department": department, "operation": operation,
                    "documentation": service["documentation"], "endpoint": service["endpoint"] + "/" + operation,
                    "result_code": payload["result_code"], "total_count": total, "page_no": page_no,
                    "complete": complete, "source_origin": "INDEPENDENT_EXTERNAL", "price_type": service["price_type"],
                    "network": "LIVE" if transport is None else "FIXTURE"}
        if persist:
            from external_store import ensure_store, save_snapshot
            ensure_store()
            saved = save_snapshot(provider, operation, query, raw, records, metadata)
            snapshots.append(saved["snapshot_id"])
            for record_index, record in enumerate(records, 1):
                record["price_id"] = "EXT:" + saved["snapshot_id"].removeprefix("SNAP:") + ":" + str(record_index)
                record["source_path"] = saved["source_path"]
                record["source_sha256"] = saved["source_sha256"]
                record["snapshot_id"] = saved["snapshot_id"]
        normalized.extend(records)
        pages.append({"page_no": page_no, "row_count": len(page_rows), "raw_sha256": source["sha256"]})
        if complete or not page_rows:
            break
        if page_no < max_pages:
            time.sleep(max(0, sleep_seconds))
    if provider == "OPINET" and item_id:
        normalized = [x for x in normalized if x["item_id"] == item_id]
    if not complete:
        warnings.append("PARTIAL_QUERY_NOT_A_COMPLETE_LATEST_CATALOG")
    if provider == "PPS_CATALOG":
        warnings.append("METADATA_ONLY_NOT_A_PRICE_SOURCE")
    return {"status": "OBSERVED" if normalized else "NO_MATCH", "provider": provider, "operation": operation,
            "snapshot_ids": snapshots, "prices": normalized, "returned_rows": len(normalized), "total_count": total,
            "complete": complete, "observed_at": observed, "query": safe_params, "pages": pages,
            "network": "LIVE" if transport is None else "FIXTURE", "warnings": sorted(set(warnings)), "fallback_applied": False}


def refresh_price_selection(selection, pricing_date=None):
    """Refresh an exact item's source. Caller owns scope/date selection policy."""
    if not isinstance(selection, dict):
        raise ApiError("INVALID_SELECTION")
    provider = provider_name(selection.get("provider", "PPS_PRICE"))
    item_id = selection.get("item_id")
    if not item_id:
        raise ApiError("EXACT_ITEM_ID_REQUIRED_FOR_REFRESH", provider)
    operation = selection.get("operation")
    if not operation:
        raise ApiError("SOURCE_OPERATION_REQUIRED_FOR_REFRESH", provider)
    if selection.get("test_only") or selection.get("network") == "FIXTURE" or selection.get("source_network") == "FIXTURE":
        raise ApiError("FIXTURE_NOT_A_LIVE_PRICE_SOURCE", provider)
    result = fetch_prices(provider, item_id=item_id, department=selection.get("department") or "토목",
                          max_pages=10, page_size=100, operation=operation,
                          params=selection.get("refresh_params"))
    result["requested_pricing_date"] = pricing_date
    result["automatic_price_selection"] = False
    return result


def api_status():
    """Local connection readiness, without consuming live request quota."""
    results = []
    for provider, service in SERVICES.items():
        try:
            _credential(provider)
            ready = True
        except ApiError:
            ready = False
        try:
            from external_store import latest_snapshot
            snapshot = latest_snapshot(provider, live_only=True)
        except Exception:
            snapshot = None
        results.append({"provider": provider, "credential_configured": ready,
                        "live_connection_verified": bool(snapshot), "verified_at": snapshot.get("observed_at") if snapshot else None,
                        "last_successful_observation": snapshot, "connection_status_scope": "LAST_SUCCESSFUL_OBSERVATION_ONLY",
                        "endpoint": service["endpoint"],
                        "operations": service["operations"], "documentation": service["documentation"]})
    return {"status": "LOCAL_READINESS_ONLY", "services": results, "no_network_calls": True,
            "daily_local_attempt_caps": {"PPS": 900, "OPINET": 250}, "observed_at": now_iso()}
