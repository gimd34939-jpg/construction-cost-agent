"""Decimal costs and explicit source gates for SYNTHETIC/DEMO_ONLY records."""

from datetime import date
from decimal import Decimal, ROUND_DOWN, localcontext
import json
from pathlib import Path
import re


class InputError(ValueError):
    """The request itself cannot be interpreted safely."""


def decimal(value, label="number"):
    """Accept explicit decimal strings/integers, never float or nonfinite values."""
    if isinstance(value, bool) or not isinstance(value, (str, int, Decimal)):
        raise InputError(f"{label}: use an explicit decimal string")
    if not re.fullmatch(r"-?(?:\d+(?:\.\d*)?|\.\d+)", str(value)):
        raise InputError(f"{label}: invalid decimal")
    result = Decimal(str(value))
    if not result.is_finite() or abs(result) > Decimal("1000000000000"):
        raise InputError(f"{label}: outside supported range")
    if result < 0:
        raise InputError(f"{label}: negative values are unsupported")
    return result


def trunc(value):
    """Excel TRUNC(..., 0) for the supported nonnegative cost domain."""
    return value.quantize(Decimal("1"), rounding=ROUND_DOWN)


def text(value):
    return None if value is None else format(value, "f")


def load_fixtures():
    folder = Path(__file__).parent / "fixtures"
    return tuple(json.loads((folder / name).read_text(encoding="utf-8"))
                 for name in ("request.json", "prices.json"))


TARGETS = {
    "plywood_12mm": ("m2", "MATERIAL_SUPPLY_ONLY", "DEMO water-resistant plywood 12 mm", "EXCLUDED"),
    "carpenter": ("person_day", "LABOR_ONLY", "DEMO carpenter daily wage", "NOT_APPLICABLE"),
    "helper": ("person_day", "LABOR_ONLY", "DEMO helper daily wage", "NOT_APPLICABLE"),
}

# This reduced demo has one recipe. Do not relabel its result as a different
# work item or a wider cost scope merely by changing request display strings.
SUPPORTED_WORK = "DEMO 12 mm plywood installation"
SUPPORTED_SCOPE = "DEMO direct installation cost; VAT, overhead, profit and delivery excluded"


def _resolve(resource, records, applied):
    candidates = [record for record in records if record.get("resource") == resource]
    if len(candidates) != 1:
        return None, {"resource": resource, "code": "MISSING_OR_AMBIGUOUS_PRICE"}
    record = candidates[0]
    unit, scope, specification, delivery = TARGETS[resource]
    gates = (
        (record.get("classification") == "SYNTHETIC" and record.get("purpose") == "DEMO_ONLY", "NOT_DEMO_ONLY"),
        (bool(record.get("id")) and bool(record.get("source_id")), "MISSING_SOURCE"),
        (record.get("unit") == unit, "UNIT_CONFLICT"),
        (record.get("specification") == specification, "SPECIFICATION_CONFLICT"),
        (record.get("scope") == scope, "SCOPE_CONFLICT"),
        (record.get("delivery") == delivery, "DELIVERY_CONFLICT"),
        (record.get("tax_basis") == "EXCLUDED", "TAX_BASIS_UNCONFIRMED"),
    )
    for valid, code in gates:
        if not valid:
            return None, {"resource": resource, "code": code, "price_id": record.get("id")}
    try:
        start = date.fromisoformat(record["effective_start"])
        end = date.fromisoformat(record["effective_end"])
        if not start <= applied <= end:
            raise ValueError("outside explicit synthetic validity period")
    except (KeyError, TypeError, ValueError):
        return None, {"resource": resource, "code": "DATE_OUT_OF_SCOPE"}
    try:
        price = decimal(record.get("price"), resource)
    except InputError:
        return None, {"resource": resource, "code": "INVALID_PRICE"}
    return price, None


def calculate(request, catalog):
    """Calculate known components; unresolved components stay None, never zero."""
    if any(value.get("classification") != "SYNTHETIC" or value.get("purpose") != "DEMO_ONLY"
           for value in (request, catalog)):
        raise InputError("Only SYNTHETIC/DEMO_ONLY requests and catalogs are supported")
    if request.get("unit") != "m2":
        raise InputError("The reduced demo supports m2 only")
    if request.get("work") != SUPPORTED_WORK:
        raise InputError("The reduced demo supports its 12 mm plywood work item only")
    if request.get("scope") != SUPPORTED_SCOPE:
        raise InputError("The reduced demo supports its explicit direct-cost scope only")
    try:
        applied = date.fromisoformat(request["pricing_date"])
    except (KeyError, TypeError, ValueError) as error:
        raise InputError("An explicit ISO pricing_date is required") from error
    quantity = decimal(request.get("quantity"), "quantity")
    assumptions = request.get("assumptions", {})
    if assumptions.get("classification") != "SYNTHETIC":
        raise InputError("Coefficients must be explicitly synthetic")
    names = ("material_loss_rate", "fixings_rate", "carpenter_person_days",
             "helper_person_days", "daily_productivity_m2", "expense_rate_on_labor")
    coefficients = {name: decimal(assumptions.get(name), name) for name in names}
    if coefficients["daily_productivity_m2"] == 0:
        raise InputError("daily_productivity_m2 must be positive")
    prices, issues = {}, []
    for resource in TARGETS:
        prices[resource], issue = _resolve(resource, catalog.get("records", []), applied)
        if issue:
            issues.append(issue)
    with localcontext() as context:
        context.prec = 50
        supply = prices["plywood_12mm"]
        base = None if supply is None else trunc(supply * (1 + coefficients["material_loss_rate"]))
        fixings = None if base is None else trunc(base * coefficients["fixings_rate"])
        material = None if base is None else base + fixings
        labor = None
        if prices["carpenter"] is not None and prices["helper"] is not None:
            labor = trunc((prices["carpenter"] * coefficients["carpenter_person_days"]
                           + prices["helper"] * coefficients["helper_person_days"])
                          / coefficients["daily_productivity_m2"])
        expense = None if labor is None else trunc(labor * coefficients["expense_rate_on_labor"])
        parts = {"material": material, "labor": labor, "expense": expense}
        amounts = {name: None if value is None else trunc(value * quantity)
                   for name, value in parts.items()}
        complete = all(value is not None for value in parts.values())
        known_unit = sum((value for value in parts.values() if value is not None), Decimal("0"))
        known_amount = sum((value for value in amounts.values() if value is not None), Decimal("0"))
    return {
        "classification": "SYNTHETIC", "purpose": "DEMO_ONLY",
        "implementation": "PUBLIC_REDUCED_DEMO",
        "status": "COMPLETE" if complete else "INCOMPLETE",
        "work": request["work"], "pricing_date": applied.isoformat(),
        "quantity": text(quantity), "unit": "m2", "currency": "KRW",
        "tax_basis": "EXCLUDED", "scope": request["scope"],
        "unit_components": {name: text(value) for name, value in parts.items()},
        "amount_components": {name: text(value) for name, value in amounts.items()},
        "unit_total": text(known_unit) if complete else None,
        "total": text(known_amount) if complete else None,
        "known_unit_subtotal": text(known_unit), "known_amount_subtotal": text(known_amount),
        "intermediates": {"material_with_loss": text(base), "fixings": text(fixings)},
        "rounding": "TRUNC toward zero at material-with-loss, fixings, labor, expense and each quantity-times-unit component",
        "issues": issues,
        "sources": [{key: record.get(key) for key in
                     ("id", "resource", "source_id", "classification", "purpose", "unit", "effective_start", "effective_end")}
                    for record in catalog.get("records", [])],
        "limitations": ["All prices and coefficients are invented for this demo.",
                        "No private database, external API, drawing takeoff or native Excel verification is used.",
                        "This is a selected-scope demonstration, not a verified project contract total."],
    }
