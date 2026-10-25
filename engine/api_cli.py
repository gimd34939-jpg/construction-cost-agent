"""Codex-facing API observations and immutable price ledger queries."""
from __future__ import annotations
import argparse
from datetime import datetime
import json
from pathlib import Path
import sys

from api_clients import ApiError, api_status, fetch_prices, provider_name, KST


def _query_options(parser, provider_default="PPS_PRICE"):
    parser.add_argument("--provider", default=provider_default, help="PPS_PRICE/PPS_SHOPPING/PPS_CATALOG/OPINET; local queries also PPS_CSV")
    parser.add_argument("--text", default="")
    parser.add_argument("--item-id", default="")
    parser.add_argument("--department", "--dept", default="토목")
    parser.add_argument("--page-size", type=int, default=10)
    parser.add_argument("--max-pages", type=int, default=1)
    parser.add_argument("--operation")
    parser.add_argument("--params", help="Non-secret additional documented query parameters as JSON")
    parser.add_argument("--as-of", default=datetime.now(KST).date().isoformat())
    parser.add_argument("--output", help="New absolute JSON result path; existing files are refused")


def build_parser():
    parser = argparse.ArgumentParser(description="공식 API 조회/독립 외부가격원장; 내역서 단가 대체 없음")
    subs = parser.add_subparsers(dest="command", required=True)
    status = subs.add_parser("status")
    status.add_argument("--output")
    fetch = subs.add_parser("fetch")
    _query_options(fetch)
    prices = subs.add_parser("prices")
    _query_options(prices, provider_default="")
    prices.add_argument("--live", action="store_true", help="Fetch one bounded live observation")
    prices.add_argument("--latest", action="store_true", help="Report explicit latest-date resolution; conflicting scopes remain blocked")
    prices.add_argument("--limit", type=int, default=100)
    item = subs.add_parser("item")
    item.add_argument("price_id", nargs="?")
    item.add_argument("--id", "--price-id", dest="named_id")
    item.add_argument("--as-of", default=datetime.now(KST).date().isoformat())
    item.add_argument("--latest", action="store_true")
    item.add_argument("--output")
    cases = subs.add_parser("cases", help="2018 과거 설계검토 사례 검색; 현재가격과 분리")
    cases.add_argument("--text", default="")
    cases.add_argument("--limit",type=int,default=20)
    cases.add_argument("--output")
    return parser


def _save_and_print(result, output):
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if output:
        path = Path(output)
        if not path.is_absolute() or path.exists():
            raise ApiError("OUTPUT_MUST_BE_NEW_ABSOLUTE_PATH")
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as f:
            f.write(rendered)
    print(rendered)


def run(args):
    from external_store import get_price, ledger_status, search_prices, resolve_latest
    if args.command == "cases":
        from historical_cases import search_cases
        return search_cases(args.text,args.limit)
    if args.command == "status":
        result = api_status()
        try:
            result["ledger"] = ledger_status()
        except Exception:
            result["ledger"] = {"status": "NOT_INITIALIZED"}
        return result
    if args.command == "item":
        price_id = args.named_id or args.price_id
        if not price_id:
            raise ApiError("PRICE_ID_REQUIRED")
        value = get_price(price_id)
        result = {"status": "FOUND" if value else "NOT_FOUND", "price": value, "network": "LOCAL_SAVED_OBSERVATION", "as_of": args.as_of}
        if value and args.latest:
            result["latest_resolution"] = resolve_latest(price_id, args.as_of)
        return result
    if args.command == "fetch" or args.live:
        provider = provider_name(args.provider or "PPS_PRICE")
        try:
            params = json.loads(args.params) if args.params else None
        except Exception:
            raise ApiError("INVALID_PARAMETER_JSON", provider) from None
        if params is not None and not isinstance(params, dict):
            raise ApiError("INVALID_PARAMETER_JSON", provider)
        result = fetch_prices(provider, args.text, args.item_id, args.department, args.page_size, args.max_pages, params, args.operation)
        result["as_of"] = args.as_of
        result["price_selection"] = "EXPLICIT_SCOPE_AND_DATE_CONFIRMATION_REQUIRED"
        return result
    provider = "" if not args.provider else ("PPS_CSV" if args.provider.upper() == "PPS_CSV" else provider_name(args.provider))
    result = search_prices(args.text or args.item_id, provider, args.limit)
    if args.item_id:
        result["rows"] = [x for x in result["rows"] if x.get("item_id") == args.item_id]
    result.update(status="SAVED_OBSERVATIONS", network="LOCAL_SAVED_OBSERVATION", as_of=args.as_of,
                  live_refreshed=False, price_selection="EXPLICIT_SCOPE_AND_DATE_CONFIRMATION_REQUIRED")
    if args.latest:
        result["latest_resolution"] = [resolve_latest(x["price_id"], args.as_of) for x in result["rows"]]
    return result


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "api":
        argv.pop(0)
    args = build_parser().parse_args(argv)
    # Refuse output overwrite BEFORE any live requests or mutations.
    if args.output and (not Path(args.output).is_absolute() or Path(args.output).exists()):
        _save_and_print(ApiError("OUTPUT_MUST_BE_NEW_ABSOLUTE_PATH").as_dict(), None)
        return 2
    try:
        result = run(args)
        _save_and_print(result, args.output)
        return 0
    except ApiError as exc:
        _save_and_print(exc.as_dict(), args.output)
        return 2
    except Exception:
        # Library exceptions can carry authentication URLs; never echo them.
        _save_and_print(ApiError("API_OR_LEDGER_OPERATION_FAILED").as_dict(), args.output)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
