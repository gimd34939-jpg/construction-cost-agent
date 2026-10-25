"""Auditable quantity arithmetic; no workbook writes, eval(), or guessed inputs.

All arithmetic uses Decimal and dimensional quantities in SI base units. Count
units are dimensionless multipliers, while their original unit is retained in
the trace. A complete arithmetic result is not an approval of a design or an
assertion that the user's evidence text came from an actual human instruction.
"""
from __future__ import annotations

import ast
import hashlib
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, localcontext, ROUND_DOWN, ROUND_HALF_UP
from pathlib import Path
from typing import Any

SCHEMA_VERSION = "construction-cost/quantity-result/v1"
_DECIMAL = re.compile(r"(?:0|[1-9]\d*)(?:\.\d+)?|\.\d+")
_VARIABLE = re.compile(r"[A-Za-z_][A-Za-z0-9_]{0,63}")
_SHA256 = re.compile(r"[0-9a-fA-F]{64}")
_MAX_INPUT = Decimal("1000000000000000")
_MAX_INTERMEDIATE = Decimal("1e30")
_ZERO_DIM = (0, 0, 0)  # length, mass, time


class QuantityError(ValueError):
    def __init__(self, code: str, message: str, variable: str | None = None):
        super().__init__(message)
        self.code, self.variable = code, variable

    def as_dict(self) -> dict:
        value = {"code": self.code, "message": str(self)}
        if self.variable is not None:
            value["variable"] = self.variable
        return value


@dataclass(frozen=True)
class Unit:
    canonical: str
    factor: Decimal
    dimensions: tuple[int, int, int]
    kind: str = "physical"


@dataclass(frozen=True)
class DimensionalValue:
    value: Decimal
    dimensions: tuple[int, int, int]


def _unit(canonical: str, factor: str, dimensions: tuple[int, int, int], kind: str = "physical") -> Unit:
    return Unit(canonical, Decimal(factor), dimensions, kind)


_UNITS: dict[str, Unit] = {}


def _register(names: tuple[str, ...], canonical: str, factor: str, dimensions: tuple[int, int, int], kind: str = "physical") -> None:
    unit = _unit(canonical, factor, dimensions, kind)
    _UNITS.update({name: unit for name in names})


_register(("1", "ratio", "무차원", "-"), "1", "1", _ZERO_DIM, "ratio")
_register(("%", "percent", "퍼센트"), "%", "0.01", _ZERO_DIM, "ratio")
_register(("EA", "ea", "count", "개", "개소", "회", "식", "본", "대", "조", "매", "세트", "set", "SET"), "EA", "1", _ZERO_DIM, "count")
_register(("천매",), "천매", "1000", _ZERO_DIM, "count")
_register(("D/M", "드럼"), "D/M", "1", _ZERO_DIM, "drum_count")
_register(("m", "M", "미터"), "m", "1", (1, 0, 0))
_register(("mm", "㎜"), "mm", "0.001", (1, 0, 0))
_register(("cm", "㎝"), "cm", "0.01", (1, 0, 0))
_register(("km", "㎞"), "km", "1000", (1, 0, 0))
_register(("m2", "m²", "m^2", "㎡", "M2"), "m2", "1", (2, 0, 0))
_register(("cm2", "cm²", "cm^2", "㎠"), "cm2", "0.0001", (2, 0, 0))
_register(("mm2", "mm²", "mm^2", "㎟"), "mm2", "0.000001", (2, 0, 0))
_register(("m3", "m³", "m^3", "㎥", "M3"), "m3", "1", (3, 0, 0))
_register(("cm3", "cm³", "cm^3", "㎤"), "cm3", "0.000001", (3, 0, 0))
_register(("mm3", "mm³", "mm^3", "㎣"), "mm3", "0.000000001", (3, 0, 0))
_register(("L", "l", "ℓ", "리터"), "L", "0.001", (3, 0, 0))
_register(("kg", "㎏", "킬로그램"), "kg", "1", (0, 1, 0))
_register(("g", "그램"), "g", "0.001", (0, 1, 0))
_register(("t", "ton", "Ton", "TON", "톤"), "t", "1000", (0, 1, 0))
_register(("hr", "h", "시간"), "hr", "3600", (0, 0, 1))
_register(("min", "분"), "min", "60", (0, 0, 1))
_register(("s", "sec", "초"), "sec", "1", (0, 0, 1))


def parse_unit(raw: Any) -> Unit:
    """Parse known units, optionally multiplied/divided (e.g. kg/m, t/m3)."""
    if not isinstance(raw, str) or not raw.strip() or len(raw) > 80:
        raise QuantityError("INVALID_UNIT", "단위를 비어 있지 않은 문자열로 명시해야 합니다.")
    text = raw.strip().replace(" ", "").replace("·", "*").replace("×", "*")
    if text in _UNITS:
        return _UNITS[text]
    if not re.fullmatch(r"[^/*]+(?:[/*][^/*]+)+", text):
        raise QuantityError("UNSUPPORTED_UNIT", f"지원하지 않는 단위입니다: {raw}")
    pieces = re.split(r"([/*])", text)
    if len(pieces) > 13 or text.count("/") > 1:
        raise QuantityError("UNSUPPORTED_UNIT", "복합 단위는 곱셈과 한 번의 나눗셈만 허용합니다.")
    factor, dimensions = Decimal(1), _ZERO_DIM
    rendered: list[str] = []
    sign = 1
    for index, piece in enumerate(pieces):
        if index % 2:
            sign = -1 if piece == "/" else 1
            rendered.append(piece)
            continue
        if piece not in _UNITS:
            raise QuantityError("UNSUPPORTED_UNIT", f"지원하지 않는 복합 단위 요소입니다: {piece}")
        unit = _UNITS[piece]
        factor = factor * unit.factor if sign > 0 else factor / unit.factor
        dimensions = tuple(a + sign * b for a, b in zip(dimensions, unit.dimensions))
        rendered.append(unit.canonical)
    return Unit("".join(rendered), factor, dimensions)


def decimal_text(value: Decimal) -> str:
    if value == 0:
        return "0"
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def _number(raw: Any, variable: str | None = None) -> Decimal:
    if isinstance(raw, bool) or raw is None or isinstance(raw, float):
        raise QuantityError("INVALID_NUMBER", "숫자는 정수 또는 일반 십진수 문자열이어야 합니다. 불리언·실수 객체는 허용하지 않습니다.", variable)
    if isinstance(raw, int):
        text = str(raw)
    elif isinstance(raw, Decimal):
        if not raw.is_finite():
            raise QuantityError("INVALID_NUMBER", "유한한 십진수만 허용합니다.", variable)
        text = format(raw, "f")
    elif isinstance(raw, str):
        text = raw
    else:
        raise QuantityError("INVALID_NUMBER", "지원하지 않는 숫자 형식입니다.", variable)
    if len(text) > 80 or not _DECIMAL.fullmatch(text):
        raise QuantityError("INVALID_NUMBER", "음수·지수·구분쉼표·공백·수식·비유한 숫자를 허용하지 않습니다.", variable)
    value = Decimal(text)
    if not value.is_finite() or value > _MAX_INPUT:
        raise QuantityError("NUMBER_OUT_OF_RANGE", "입력 범위는 0 이상 10^15 이하입니다.", variable)
    return value


def _bounded(value: Decimal, dimensions: tuple[int, int, int]) -> DimensionalValue:
    if not value.is_finite() or abs(value) > _MAX_INTERMEDIATE or any(abs(x) > 9 for x in dimensions):
        raise QuantityError("ARITHMETIC_OUT_OF_RANGE", "중간 계산값 또는 차원이 허용 범위를 벗어났습니다.")
    return DimensionalValue(value, dimensions)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _evidence(raw: Any, variable: str, cache: dict) -> tuple[dict, bool, list[str]]:
    if not isinstance(raw, dict):
        raise QuantityError("MISSING_EVIDENCE", "입력값의 근거를 명시해야 합니다.", variable)
    evidence_type = raw.get("type")
    if evidence_type == "USER_INSTRUCTION":
        instruction = raw.get("instruction")
        if not isinstance(instruction, str) or not instruction.strip() or len(instruction) > 10000:
            raise QuantityError("INVALID_USER_INSTRUCTION", "실제 사용자 지시 원문을 비어 있지 않은 문자열로 기록해야 합니다.", variable)
        return ({"variable": variable, "type": evidence_type, "status": "INSTRUCTION_RECORDED", "instruction": instruction}, False,
                ["USER_INSTRUCTION 문구의 실제 대화 일치 여부는 상위 코덱스 스킬이 확인해야 합니다."])
    if evidence_type == "TEST_ONLY":
        return ({"variable": variable, "type": evidence_type, "status": "TEST_ONLY", "note": raw.get("note", "검증용 가상 입력")}, True,
                ["TEST_ONLY 입력을 포함한 결과는 실제 설계·가격 근거로 적용할 수 없습니다."])
    if evidence_type not in ("SOURCE_FILE", "DRAWING"):
        raise QuantityError("UNSUPPORTED_EVIDENCE", "근거는 USER_INSTRUCTION, SOURCE_FILE, DRAWING 또는 TEST_ONLY여야 합니다.", variable)
    path_value = raw.get("path")
    expected_sha = raw.get("sha256")
    location = raw.get("location")
    if not isinstance(path_value, str) or not path_value.strip() or not Path(path_value).is_absolute():
        raise QuantityError("INVALID_SOURCE_PATH", "원문 근거에는 실제 절대 파일 경로가 필요합니다.", variable)
    if not isinstance(expected_sha, str) or not _SHA256.fullmatch(expected_sha):
        raise QuantityError("INVALID_SOURCE_SHA256", "원문의 SHA256 64자리 값을 기록해야 합니다.", variable)
    if not isinstance(location, str) or not location.strip():
        raise QuantityError("INVALID_SOURCE_LOCATION", "시트/셀·쪽/영역·객체 위치를 기록해야 합니다.", variable)
    path = Path(path_value)
    if not path.is_file():
        raise QuantityError("SOURCE_FILE_NOT_FOUND", "원문 근거 파일을 찾을 수 없습니다.", variable)
    try:
        stat = path.stat()
        if stat.st_size == 0:
            raise QuantityError("EMPTY_SOURCE_FILE", "내용이 없는 파일은 수량 근거가 될 수 없습니다.", variable)
        cache_key = (str(path.resolve()), stat.st_size, stat.st_mtime_ns)
        if cache_key not in cache:
            cache[cache_key] = _sha256(path)
        actual_sha = cache[cache_key]
    except OSError as exc:
        raise QuantityError("SOURCE_READ_ERROR", f"원문 파일을 읽을 수 없습니다: {type(exc).__name__}", variable) from exc
    if actual_sha.lower() != expected_sha.lower():
        raise QuantityError("SOURCE_SHA256_MISMATCH", "현재 원문 파일 SHA256이 입력 근거와 다릅니다.", variable)
    return ({"variable": variable, "type": evidence_type, "status": "FILE_IDENTITY_VERIFIED", "path": str(path),
             "sha256": actual_sha, "location": location}, False,
            ["파일 SHA 검증은 판본 일치 확인이며 해당 위치 숫자의 원문 대조는 자료 추출·상위 검증 단계에서 확인합니다."])


def _parse_expression(expression: Any, variables: dict) -> tuple[ast.Expression, list[str]]:
    if not isinstance(expression, str) or not expression.strip() or len(expression) > 1024:
        raise QuantityError("INVALID_EXPRESSION", "물량 산식은 1~1,024자 문자열이어야 합니다.")
    try:
        tree = ast.parse(expression, mode="eval")
    except (SyntaxError, RecursionError) as exc:
        raise QuantityError("INVALID_EXPRESSION", "산식 구문을 해석할 수 없습니다.") from exc
    nodes = list(ast.walk(tree))
    if len(nodes) > 128:
        raise QuantityError("EXPRESSION_TOO_COMPLEX", "산식 노드는 최대 128개입니다.")
    allowed = (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Name, ast.Load, ast.Constant,
               ast.Add, ast.Sub, ast.Mult, ast.Div, ast.UAdd, ast.USub)
    names: list[str] = []
    for node in nodes:
        if not isinstance(node, allowed):
            raise QuantityError("UNSAFE_EXPRESSION", "산식에는 변수·십진수·괄호와 + - * /만 사용할 수 있습니다.")
        if isinstance(node, ast.Name):
            if node.id not in variables:
                raise QuantityError("UNDECLARED_VARIABLE", f"산식 변수가 선언되지 않았습니다: {node.id}", node.id)
            if node.id not in names:
                names.append(node.id)
        if isinstance(node, ast.Constant):
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise QuantityError("UNSAFE_EXPRESSION", "산식의 상수는 일반 십진수만 허용합니다.")
            token = ast.get_source_segment(expression, node)
            _number(token)
    return tree, names


def _evaluate(node: ast.AST, expression: str, values: dict[str, DimensionalValue], depth: int = 0) -> DimensionalValue:
    if depth > 32:
        raise QuantityError("EXPRESSION_TOO_COMPLEX", "산식의 중첩이 너무 깊습니다.")
    if isinstance(node, ast.Expression):
        return _evaluate(node.body, expression, values, depth + 1)
    if isinstance(node, ast.Name):
        return values[node.id]
    if isinstance(node, ast.Constant):
        return DimensionalValue(_number(ast.get_source_segment(expression, node)), _ZERO_DIM)
    if isinstance(node, ast.UnaryOp):
        operand = _evaluate(node.operand, expression, values, depth + 1)
        return _bounded(-operand.value if isinstance(node.op, ast.USub) else operand.value, operand.dimensions)
    left = _evaluate(node.left, expression, values, depth + 1)
    right = _evaluate(node.right, expression, values, depth + 1)
    if isinstance(node.op, (ast.Add, ast.Sub)):
        if left.dimensions != right.dimensions:
            raise QuantityError("ADDITION_DIMENSION_MISMATCH", "덧셈·공제 항목의 차원이 서로 다릅니다.")
        value = left.value + right.value if isinstance(node.op, ast.Add) else left.value - right.value
        return _bounded(value, left.dimensions)
    if isinstance(node.op, ast.Mult):
        return _bounded(left.value * right.value, tuple(a + b for a, b in zip(left.dimensions, right.dimensions)))
    if right.value == 0:
        raise QuantityError("DIVISION_BY_ZERO", "분모가 0이므로 물량을 계산할 수 없습니다.")
    return _bounded(left.value / right.value, tuple(a - b for a, b in zip(left.dimensions, right.dimensions)))


def _measurement(raw: dict, measurements: list | None, variable: str) -> tuple[dict | None, dict | None]:
    measurement_id = raw.get("measurement_id")
    if not isinstance(measurement_id, str) or not measurement_id.strip():
        raise QuantityError("INVALID_MEASUREMENT_ID", "도면 측정 ID를 명시해야 합니다.", variable)
    if set(raw) != {"measurement_id"}:
        raise QuantityError("MEASUREMENT_OVERRIDE_FORBIDDEN", "도면 참조 입력은 measurement_id만 사용하며 값·단위를 덮어쓸 수 없습니다.", variable)
    matches = [item for item in (measurements or []) if isinstance(item, dict) and item.get("measurement_id") == measurement_id]
    if not matches:
        return None, {"variable": variable, "reason": "DRAWING_MEASUREMENT_NOT_FOUND", "measurement_id": measurement_id}
    if len(matches) != 1:
        raise QuantityError("DUPLICATE_MEASUREMENT_ID", "같은 ID의 도면 측정값이 여러 개 있습니다.", variable)
    measurement = matches[0]
    if measurement.get("status") != "VERIFIED" or measurement.get("ambiguous", False) is not False:
        return None, {"variable": variable, "reason": "DRAWING_MEASUREMENT_NOT_VERIFIED", "measurement_id": measurement_id,
                      "measurement_status": measurement.get("status"), "ambiguous": measurement.get("ambiguous", False)}
    source = measurement.get("source", measurement.get("evidence"))
    if not isinstance(source, dict):
        raise QuantityError("MISSING_DRAWING_SOURCE", "검증된 도면 측정값에도 파일·SHA·위치 근거가 필요합니다.", variable)
    evidence = dict(source)
    # Even synthetic drawing fixtures must identify an actual source file.
    evidence["type"] = "DRAWING"
    return {"value": measurement.get("value"), "unit": measurement.get("unit"), "evidence": evidence,
            "measurement_id": measurement_id,
            "test_only": source.get("type") == "TEST_ONLY" or measurement.get("test_only") is True}, None


def evaluate_quantity(rule: dict, inputs: dict, drawing_measurements: list | None = None) -> dict:
    """Return COMPLETE, INCOMPLETE, or BLOCKED with explicit evidence and units.

    Float objects, scientific notation, missing evidence, unknown variables and
    unsafe AST nodes are rejected. Omitting a variable returns INCOMPLETE/null;
    zero is a legitimate explicit quantity. A TEST_ONLY evidence input taints
    the whole result. No exception is expected for ordinary invalid requests.
    """
    result: dict = {"schema_version": SCHEMA_VERSION, "status": "BLOCKED", "value": None,
                    "unit": rule.get("output_unit") if isinstance(rule, dict) else None,
                    "expression": rule.get("expression") if isinstance(rule, dict) else None,
                    "missing": [], "errors": [], "inputs_used": [], "evidence_checks": [],
                    "rounding": None, "test_only": False, "warnings": []}
    try:
        if not isinstance(rule, dict) or not isinstance(inputs, dict):
            raise QuantityError("INVALID_INPUT_OBJECT", "rule과 inputs는 JSON 객체여야 합니다.")
        if drawing_measurements is not None and not isinstance(drawing_measurements, list):
            raise QuantityError("INVALID_DRAWING_MEASUREMENTS", "도면 측정값은 배열이어야 합니다.")
        variables = rule.get("variables")
        if not isinstance(variables, dict) or len(variables) > 32:
            raise QuantityError("INVALID_VARIABLES", "variables는 최대 32개 변수의 객체여야 합니다.")
        for name, definition in variables.items():
            if not isinstance(name, str) or not _VARIABLE.fullmatch(name) or not isinstance(definition, dict):
                raise QuantityError("INVALID_VARIABLES", "변수 이름·정의 형식이 올바르지 않습니다.")
            parse_unit(definition.get("unit"))
        extras = sorted(set(inputs) - set(variables), key=str)
        if extras:
            raise QuantityError("UNKNOWN_INPUT_VARIABLE", f"선언되지 않은 입력 변수가 있습니다: {', '.join(map(str, extras))}")
        expression = rule.get("expression")
        tree, used_names = _parse_expression(expression, variables)
        output_unit = parse_unit(rule.get("output_unit"))
        rounding = rule.get("rounding")
        if rounding is not None:
            if (not isinstance(rounding, dict) or set(rounding) != {"digits", "mode"}
                    or isinstance(rounding.get("digits"), bool) or not isinstance(rounding.get("digits"), int)
                    or not 0 <= rounding["digits"] <= 12
                    or rounding.get("mode") not in ("ROUND_DOWN", "ROUND_HALF_UP")):
                raise QuantityError("INVALID_ROUNDING", "rounding에는 0~12 정수 digits와 ROUND_DOWN/ROUND_HALF_UP mode를 명시해야 합니다.")
            result["rounding"] = dict(rounding)
        cache: dict = {}
        values: dict[str, DimensionalValue] = {}
        with localcontext() as context:
            context.prec = 50
            for name in used_names:
                if name not in inputs or inputs[name] is None:
                    result["missing"].append({"variable": name, "reason": "MISSING_INPUT", "unit": variables[name]["unit"]})
                    continue
                raw = inputs[name]
                try:
                    if not isinstance(raw, dict):
                        raise QuantityError("INVALID_VARIABLE_INPUT", "각 변수는 value·unit·evidence 객체 또는 measurement_id 참조여야 합니다.", name)
                    if "measurement_id" in raw:
                        raw, missing = _measurement(raw, drawing_measurements, name)
                        if missing:
                            result["missing"].append(missing)
                            continue
                    value = _number(raw.get("value"), name)
                    actual_unit = parse_unit(raw.get("unit"))
                    expected_unit = parse_unit(variables[name]["unit"])
                    if actual_unit.dimensions != expected_unit.dimensions:
                        raise QuantityError("INPUT_DIMENSION_MISMATCH", f"{name} 입력 단위의 차원이 선언 단위와 다릅니다.", name)
                    if actual_unit.dimensions == _ZERO_DIM and actual_unit.kind != expected_unit.kind:
                        raise QuantityError("INPUT_UNIT_KIND_MISMATCH", "개수와 비율 단위를 서로 대체할 수 없습니다.", name)
                    evidence, test_only, warnings = _evidence(raw.get("evidence"), name, cache)
                    result["evidence_checks"].append(evidence)
                    result["test_only"] = result["test_only"] or test_only or raw.get("test_only") is True
                    result["warnings"].extend(warnings)
                    si = value * actual_unit.factor
                    values[name] = _bounded(si, actual_unit.dimensions)
                    trace = {"variable": name, "value": decimal_text(value), "unit": raw["unit"],
                             "normalized_value": decimal_text(si / expected_unit.factor), "normalized_unit": variables[name]["unit"],
                             "si_value": decimal_text(si), "dimensions": list(actual_unit.dimensions), "evidence": evidence}
                    if raw.get("measurement_id"):
                        trace["measurement_id"] = raw["measurement_id"]
                    result["inputs_used"].append(trace)
                except QuantityError as exc:
                    if exc.variable is None:
                        exc.variable = name
                    result["errors"].append(exc.as_dict())
            result["warnings"] = list(dict.fromkeys(result["warnings"]))
            if result["errors"]:
                return result
            if result["missing"]:
                result["status"] = "INCOMPLETE"
                return result
            calculated = _evaluate(tree, expression, values)
            if calculated.dimensions != output_unit.dimensions:
                raise QuantityError("OUTPUT_DIMENSION_MISMATCH", "산식 결과의 차원이 출력 단위와 다릅니다.")
            raw_value = calculated.value / output_unit.factor
            if raw_value < 0:
                raise QuantityError("NEGATIVE_QUANTITY", "공제 결과가 음수입니다. 일반 공종의 물량으로 적용할 수 없습니다.")
            if raw_value > _MAX_INPUT:
                raise QuantityError("RESULT_OUT_OF_RANGE", "최종 물량이 10^15를 초과합니다.")
            result["unrounded_value"] = decimal_text(raw_value)
            result["dimensions"] = list(calculated.dimensions)
            if rounding:
                quantum = Decimal(1).scaleb(-rounding["digits"])
                mode = ROUND_DOWN if rounding["mode"] == "ROUND_DOWN" else ROUND_HALF_UP
                raw_value = raw_value.quantize(quantum, rounding=mode)
            result["value"] = decimal_text(raw_value)
            result["status"] = "COMPLETE"
            return result
    except QuantityError as exc:
        result["errors"].append(exc.as_dict())
    except (InvalidOperation, ArithmeticError, OverflowError, RecursionError) as exc:
        result["errors"].append({"code": "ARITHMETIC_ERROR", "message": f"정확한 십진 계산을 완료할 수 없습니다: {type(exc).__name__}"})
    except OSError as exc:
        result["errors"].append({"code": "SOURCE_READ_ERROR", "message": f"원문 파일 검증 중 오류: {type(exc).__name__}"})
    return result


__all__ = ["evaluate_quantity", "parse_unit", "decimal_text", "QuantityError"]
