"""Read-only source drawing extraction with explicit selectors and immutable provenance.

PDF bboxes use unrotated page points (PyMuPDF text/search coordinates).
DWG conversion touches a private copy only; CAD selections require exact handles.
No inferred drawing scale, global number selection, or automatic block counting.
"""
from __future__ import annotations

import argparse
from decimal import Decimal, InvalidOperation
import hashlib
import gzip
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import uuid


NUMBER = re.compile(r"(?<![\w.])[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?![\w.])")
AUTOCAD = Path("C:/Program Files/Autodesk/AutoCAD 2021/accoreconsole.exe")


def sha256(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(1048576), b""):
            h.update(chunk)
    return h.hexdigest()


def _write(path, obj):
    with Path(path).open("x", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def _write_index(path, obj):
    # CAD PDF vector paths can contain millions of points; compressed indexes keep source detail affordable.
    with gzip.open(path, "xt", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))


def _decimal(value):
    if isinstance(value, bool):
        raise ValueError("불리언은 수치가 아닙니다.")
    try:
        d = Decimal(str(value).replace(",", ""))
    except InvalidOperation as e:
        raise ValueError("유효한 십진수가 아닙니다.") from e
    if not d.is_finite() or abs(d) > Decimal("1e15"):
        raise ValueError("허용 범위 밖 수치입니다.")
    return format(d, "f")


def _jsonable(obj):
    if isinstance(obj, (str, bool, int, float)) or obj is None:
        return obj
    if isinstance(obj, dict):
        return {str(k): _jsonable(v) for k, v in obj.items()}
    try:
        return [_jsonable(x) for x in obj]
    except TypeError:
        return str(obj)


def _unit_in_text(unit, text):
    aliases = {"m": [r"(?<![A-Za-z])m(?![A-Za-z²³23])", "미터"],
               "mm": [r"(?<![A-Za-z])mm(?![A-Za-z])", "m/m"],
               "cm": [r"(?<![A-Za-z])cm(?![A-Za-z])"],
               "m2": ["㎡", r"(?<![A-Za-z])m²", r"(?<![A-Za-z])m\s*2(?![A-Za-z0-9])"], "㎡": ["㎡", r"(?<![A-Za-z])m²", r"(?<![A-Za-z])m\s*2(?![A-Za-z0-9])"],
               "m3": ["㎥", r"(?<![A-Za-z])m³", r"(?<![A-Za-z])m\s*3(?![A-Za-z0-9])"], "㎥": ["㎥", r"(?<![A-Za-z])m³", r"(?<![A-Za-z])m\s*3(?![A-Za-z0-9])"],
               "EA": [r"(?<![A-Za-z])EA(?![A-Za-z])", "개", "대", "조", "개소"], "개": ["개", r"(?<![A-Za-z])EA(?![A-Za-z])"],
               "대": ["대"], "kg": [r"(?<![A-Za-z])kg(?![A-Za-z])", "㎏"], "t": [r"(?<![A-Za-z])t(?![A-Za-z])", "ton", "톤"]}
    return any(re.search(p, text, re.I) for p in aliases.get(unit, [re.escape(unit)]))


def _extract_value(text, selector):
    kind = selector.get("kind", "number")
    if kind not in {"number", "count", "text"}:
        return None, "UNVERIFIED", "지원하지 않는 선택 종류입니다."
    pattern = selector.get("pattern")
    if pattern:
        if not isinstance(pattern, str) or len(pattern) > 512:
            return None, "UNVERIFIED", "정규식은 512자 이하 문자열이어야 합니다."
        try:
            rx = re.compile(pattern)
            if rx.groups > 1 and "value" not in rx.groupindex:
                return None, "UNVERIFIED", "수치 선택식은 value 그룹 또는 단일 캡처만 허용합니다."
            matches = list(rx.finditer(text))
        except re.error:
            return None, "UNVERIFIED", "유효하지 않은 정규식입니다."
        if len(matches) != 1:
            return None, "AMBIGUOUS" if len(matches) > 1 else "UNVERIFIED", f"선택식과 일치한 값이 {len(matches)}개입니다."
        m = matches[0]
        value = m.group("value") if "value" in rx.groupindex else m.group(1) if rx.groups else m.group(0)
    elif kind == "text":
        value = text.strip()
    else:
        vals = NUMBER.findall(text)
        if len(vals) != 1:
            return None, "AMBIGUOUS" if len(vals) > 1 else "UNVERIFIED", f"영역에 식별 가능한 숫자가 {len(vals)}개입니다."
        value = vals[0]
    if kind == "text":
        return value or None, "VERIFIED" if value else "UNVERIFIED", None
    try:
        value = _decimal(value)
    except ValueError as e:
        return None, "UNVERIFIED", str(e)
    if kind == "count" and (Decimal(value) < 0 or Decimal(value) % 1):
        return None, "UNVERIFIED", "개수는 0 이상 정수여야 합니다."
    return value, "VERIFIED", None


def _base_measurement(selector, source):
    return {"measurement_id": selector["id"], "status": "UNVERIFIED", "value": None,
            "unit": selector.get("unit"), "source": dict(source), "ambiguous": False,
            "kind": selector.get("kind", "number"), "raw_text": "", "reasons": []}


def _pdf(path, request, folder, source):
    import pymupdf as fitz
    measurements = []
    indexes = []
    with fitz.open(path) as doc:
        if doc.needs_pass:
            raise ValueError("암호화 PDF는 해제된 별도 사본이 필요합니다.")
        pages = request.get("pages") or list(range(1, len(doc) + 1))
        if any(isinstance(n, bool) or not isinstance(n, int) or n < 1 or n > len(doc) for n in pages):
            raise ValueError("PDF 쪽 번호는 1부터 시작하며 실제 쪽 범위에 있어야 합니다.")
        pages = sorted(set(pages))
        for number in pages:
            page = doc[number - 1]
            text = page.get_text()
            spans = []
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    spans.extend({"text": s["text"], "bbox": list(s["bbox"]), "font": s["font"],
                                  "size": s["size"], "direction": list(line["dir"])} for s in line["spans"])
            drawing_vectors = page.get_drawings()
            vector_fields = ("type", "rect", "items", "closePath", "fill", "color", "seqno") if request.get("include_vector_geometry") is True else ("type", "rect", "seqno")
            vectors = [{k: _jsonable(d[k]) for k in vector_fields if k in d} for d in drawing_vectors]
            index = {"page": number, "rotation": page.rotation, "coordinate_space": "UNROTATED_PAGE_POINTS",
                     "bounds": list(page.rect * page.derotation_matrix), "text": text, "spans": spans,
                     "vector_count": len(vectors), "vectors": vectors, "vector_detail": "FULL_PATHS" if request.get("include_vector_geometry") is True else "BOUNDING_BOXES", "image_count": len(page.get_images()),
                     "drawing_number_candidates": sorted(set(re.findall(r"\b[ACMS]-\d{3}\b", text))),
                     "text_status": "TEXT_AVAILABLE" if text.strip() else "SCANNED_OR_OUTLINED_REQUIRES_REVIEW"}
            index_path = folder / f"page_{number:03d}.json.gz"
            _write_index(index_path, index)
            indexes.append({"page": number, "path": str(index_path), "sha256": sha256(index_path),
                            "span_count": len(spans), "vector_count": len(vectors), "text_status": index["text_status"], "encoding": "gzip+json"})
        for selector in request.get("selectors", []):
            m = _base_measurement(selector, source)
            number = selector.get("page")
            if isinstance(number, bool) or not isinstance(number, int) or number not in pages:
                m["reasons"].append("선택 쪽이 인덱싱한 실제 PDF 쪽에 없습니다.")
                measurements.append(m)
                continue
            page = doc[number - 1]
            bbox = selector.get("bbox")
            bounds = page.rect * page.derotation_matrix
            if (not isinstance(bbox, list) or len(bbox) != 4 or
                    any(isinstance(x, bool) or not isinstance(x, (int, float)) or not math.isfinite(x) for x in bbox)):
                m["reasons"].append("쪽과 정확한 bbox를 명시해야 합니다. 전체 문서 자동 숫자 선택은 금지합니다.")
                measurements.append(m)
                continue
            rect = fitz.Rect(bbox)
            if rect.is_empty or rect.is_infinite or not bounds.contains(rect):
                m["reasons"].append("bbox는 회전 전 PDF 쪽 좌표의 범위 안에 있어야 합니다.")
                measurements.append(m)
                continue
            raw_text = page.get_text("text", clip=rect)
            m.update(raw_text=raw_text, page=number, bbox=bbox, coordinate_space="UNROTATED_PAGE_POINTS")
            m["source"]["location"] = f"PDF {number}쪽 / bbox {bbox} / UNROTATED_PAGE_POINTS"
            value, status, reason = _extract_value(raw_text, selector)
            m.update(value=value, status=status, ambiguous=status == "AMBIGUOUS")
            if reason:
                m["reasons"].append(reason)
            unit_text = raw_text
            if selector.get("unit_bbox"):
                ur = fitz.Rect(selector["unit_bbox"])
                if ur.is_empty or not bounds.contains(ur):
                    m.update(status="UNVERIFIED", value=None)
                    m["reasons"].append("unit_bbox가 실제 쪽 범위를 벗어납니다.")
                else:
                    m["unit_evidence"] = {"bbox": selector["unit_bbox"], "text": page.get_text("text", clip=ur)}
                    unit_text += "\n" + m["unit_evidence"]["text"]
            if m["status"] == "VERIFIED" and m["kind"] != "text":
                if not m["unit"] or not _unit_in_text(m["unit"], unit_text):
                    m.update(status="UNVERIFIED", value=None)
                    m["reasons"].append("명시 단위가 선택 영역/단위 근거 영역에서 확인되지 않습니다.")
            if m["status"] == "VERIFIED":
                m["validation"] = "EXACT_SOURCE_TEXT_AND_UNIT"
            measurements.append(m)
        return {"format": "PDF", "page_count": len(doc), "indexes": indexes, "measurements": measurements,
                "document_identity_evidence": {"page": 1, "text": doc[0].get_text()[:12000], "metadata": doc.metadata,
                    "note": "프로젝트 ID는 호출자가 지정한 식별자입니다. 표제·날짜·승인 판본은 이 원문 및 실제 승인자료와 대조해야 합니다."},
                "limitations": ["숫자와 단위의 원문 추출 검증이며 해당 치수의 공종 적용·승인 판본 판정은 별도입니다.",
                                "도면의 축척·그림 길이로 물량을 추정하지 않습니다. 스캔/윤곽선 문자는 자동 확정하지 않습니다."]}


def _short_path(path):
    import ctypes
    buf = ctypes.create_unicode_buffer(32768)
    size = ctypes.windll.kernel32.GetShortPathNameW(str(path), buf, len(buf))
    return buf.value if size else str(path)


def _dxf_index(path):
    raw = Path(path).read_bytes()
    if raw.startswith(b"AutoCAD Binary DXF"):
        raise ValueError("Binary DXF는 지원하지 않습니다. ASCII DXF 사본을 사용하세요.")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp949", errors="replace")
    lines = text.splitlines()
    if len(lines) % 2:
        lines = lines[:-1]
    pairs = []
    for n in range(0, len(lines), 2):
        try:
            code = int(lines[n].strip())
        except ValueError as e:
            raise ValueError("유효한 ASCII DXF 그룹코드가 아닙니다.") from e
        pairs.append((code, lines[n + 1].strip()))
    units = 0
    for n, pair in enumerate(pairs[:-1]):
        if pair == (9, "$INSUNITS"):
            units = int(pairs[n + 1][1])
            break
    entities = []
    section = None
    current = None
    def emit():
        if not current:
            return
        data = current["groups"]
        def get(code, default=None):
            return next((v for k, v in data if k == code), default)
        ent = {"type": current["type"], "handle": get(5), "layer": get(8), "layout": get(410), "paper_space": get(67, "0"),
               "raw_groups": data, "text": ""}
        if ent["type"] == "MTEXT":
            ent["text"] = "".join(v for k, v in data if k in (1, 3)).replace("\\P", "\n")
        elif ent["type"] in {"TEXT", "ATTRIB", "ATTDEF"}:
            ent["text"] = get(1, "")
        elif ent["type"] == "DIMENSION":
            ent["measurement"] = get(42)
            ent["override_text"] = get(1)
        if ent["type"] == "INSERT":
            ent["block_name"] = get(2)
        entities.append(ent)
    n = 0
    while n < len(pairs):
        code, val = pairs[n]
        if code == 0 and val == "SECTION":
            n += 1
            section = pairs[n][1]
        elif code == 0 and val == "ENDSEC":
            if section == "ENTITIES":
                emit()
                current = None
            section = None
        elif section == "ENTITIES":
            if code == 0:
                emit()
                current = {"type": val, "groups": []}
            elif current is not None:
                current["groups"].append((code, val))
        n += 1
    return {"insunits_code": units, "drawing_unit": {1: "inch", 2: "ft", 4: "mm", 5: "cm", 6: "m", 7: "km"}.get(units),
            "entities": entities, "entity_count": len(entities), "block_expansion": "NOT_PERFORMED",
            "limitations": ["블록 내부/외부참조/동적블록은 자동 개수 집계하지 않습니다.", "INSUNITS만으로 실제 도면 작성단위나 치수 배율을 확정하지 않습니다."]}


def _cad(path, request, folder, source):
    converted = None
    conversion = None
    if path.suffix.lower() == ".dwg":
        if not AUTOCAD.is_file():
            return {"format": "DWG", "status": "BLOCKED", "measurements": [], "reason": "AutoCAD Core Console을 찾을 수 없습니다."}
        cad_dir = folder / "cad_copy"
        cad_dir.mkdir()
        copied = cad_dir / "input.dwg"
        shutil.copy2(path, copied)
        output = cad_dir / "output.dxf"
        script = cad_dir / "convert.scr"
        # Source-folder LISP files are not copied or loaded. Conversion runs in an isolated output directory.
        script.write_text('(setvar "FILEDIA" 0)\n(setvar "CMDDIA" 0)\n(setvar "SECURELOAD" 2)\n(setvar "DEMANDLOAD" 0)\n(setvar "XLOADCTL" 0)\n'
                          '(command "_.DXFOUT" "output.dxf" "16")\n(command "_.QUIT" "_N")\n', encoding="ascii")
        kwargs = {"cwd": str(cad_dir), "stdout": subprocess.PIPE, "stderr": subprocess.STDOUT, "timeout": 100}
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        try:
            proc = subprocess.run([str(AUTOCAD), "/i", str(copied), "/s", str(script), "/l", "en-US"], **kwargs)
            data = proc.stdout
            exit_code = proc.returncode
            reason = None
        except subprocess.TimeoutExpired as e:
            data = e.stdout or b""
            exit_code = None
            reason = "CAD_CONVERSION_TIMEOUT"
        (cad_dir / "conversion.log").write_bytes(data)
        log_text = data.decode("utf-16" if data.count(b"\x00") > len(data) // 5 else "utf-8", errors="replace")
        font_warning = "누락된 글꼴" in log_text or "대체중" in log_text or "missing font" in log_text.lower()
        conversion = {"application": str(AUTOCAD), "copied_source": str(copied), "copy_input_sha256": sha256(copied),
                      "exit_code": exit_code, "log": str(cad_dir / "conversion.log"), "reason": reason,
                      "font_warning": font_warning, "visual_fidelity_verified": False,
                      "scope": "객체 텍스트·핸들·속성 추출; 누락 글꼴/외부참조/도형 인쇄의 시각적 동등성은 별도 검증 필요"}
        if not output.is_file() or output.stat().st_size < 100:
            return {"format": "DWG", "status": "BLOCKED", "measurements": [], "conversion": conversion,
                    "reason": reason or "CAD_DXF_OUTPUT_UNAVAILABLE"}
        converted = output
        conversion.update(output=str(output), output_sha256=sha256(output))
    dxf = converted or path
    index = _dxf_index(dxf)
    index_path = folder / "cad_entities.json.gz"
    _write_index(index_path, index)
    by_handle = {e["handle"]: e for e in index["entities"] if e["handle"]}
    measurements = []
    for selector in request.get("selectors", []):
        m = _base_measurement(selector, source)
        handle = selector.get("handle")
        if not handle or handle not in by_handle:
            m["reasons"].append("정확한 DXF 객체 핸들을 명시해야 합니다. 블록 전체 자동개수 계산은 금지합니다.")
            measurements.append(m)
            continue
        ent = by_handle[handle]
        m.update(handle=handle, entity_type=ent["type"], raw_text=ent["text"], drawing_unit=index["drawing_unit"])
        m["source"]["location"] = f"CAD handle={handle}; type={ent['type']}; layer={ent['layer']}; layout={ent['layout']}"
        if converted:
            m["source"].update(converted_path=str(dxf), converted_sha256=sha256(dxf))
        if ent["type"] not in {"TEXT", "MTEXT", "ATTRIB", "ATTDEF"}:
            m["reasons"].append("이 객체는 원문 텍스트 수치가 아닙니다. 치수 배율·블록/기하 의미를 확인하기 전 확정할 수 없습니다.")
        else:
            value, status, reason = _extract_value(ent["text"], selector)
            m.update(value=value, status=status, ambiguous=status == "AMBIGUOUS")
            if reason:
                m["reasons"].append(reason)
            if status == "VERIFIED" and m["kind"] != "text" and (not m["unit"] or not _unit_in_text(m["unit"], ent["text"])):
                m.update(status="UNVERIFIED", value=None)
                m["reasons"].append("객체 문자에서 명시된 단위를 확인할 수 없습니다.")
            if m["status"] == "VERIFIED":
                m["validation"] = "EXACT_ENTITY_TEXT_AND_UNIT"
        measurements.append(m)
    return {"format": path.suffix[1:].upper(), "conversion": conversion, "indexes": [{"path": str(index_path), "sha256": sha256(index_path), "encoding": "gzip+json"}],
            "entity_count": index["entity_count"], "drawing_unit": index["drawing_unit"], "measurements": measurements,
            "limitations": index["limitations"]}


def extract_drawing(request: dict, output_dir: Path) -> dict:
    """Extract source-verified annotations; no numeric selection occurs without explicit selectors."""
    if not isinstance(request, dict):
        raise ValueError("request는 객체여야 합니다.")
    for key in ("path", "sha256", "project_id", "revision"):
        if not isinstance(request.get(key), str) or not request[key].strip():
            raise ValueError(f"필수 문자열 누락: {key}")
    path = Path(request["path"]).resolve(strict=True)
    if not path.is_file():
        raise ValueError("도면은 파일이어야 합니다.")
    before = sha256(path)
    if before.lower() != request["sha256"].lower():
        raise ValueError("원본 SHA256이 요청과 다릅니다.")
    selectors = request.get("selectors", [])
    if not isinstance(selectors, list) or len(selectors) > 1000:
        raise ValueError("selectors는 1000개 이하 배열이어야 합니다.")
    ids = []
    for s in selectors:
        if not isinstance(s, dict) or not isinstance(s.get("id"), str) or not s["id"].strip():
            raise ValueError("모든 selector에 고유 id를 명시해야 합니다.")
        ids.append(s["id"])
    if len(ids) != len(set(ids)):
        raise ValueError("중복 measurement id입니다.")
    folder = Path(output_dir).resolve() / ("drawing_" + uuid.uuid4().hex[:12])
    folder.mkdir(parents=True, exist_ok=False)
    source = {"path": str(path), "sha256": before, "project_id": request["project_id"], "revision": request["revision"],
              "location": "UNSELECTED", "identity_status": "CALLER_IDENTIFIED_REQUIRES_DOCUMENT_TITLE_CHECK"}
    _write(folder / "request.json", request)
    try:
        if path.suffix.lower() == ".pdf":
            result = _pdf(path, request, folder, source)
        elif path.suffix.lower() in {".dwg", ".dxf"}:
            result = _cad(path, request, folder, source)
        else:
            result = {"status": "BLOCKED", "measurements": [], "reason": "UNSUPPORTED_DRAWING_FORMAT"}
    except Exception as e:
        result = {"status": "BLOCKED", "measurements": [], "reason": f"{type(e).__name__}: {e}"}
    after = sha256(path)
    if after != before:
        raise RuntimeError("원본 해시가 변경되었습니다. 결과를 사용할 수 없습니다.")
    result.update(source=source, source_before_sha256=before, source_after_sha256=after, original_unchanged=True)
    result.setdefault("status", "VERIFIED" if result.get("measurements") and all(m["status"] == "VERIFIED" for m in result["measurements"]) else
                      "PARTIAL" if result.get("measurements") else "INDEXED")
    result["report_path"] = str(folder / "drawing_extraction.json")
    _write(result["report_path"], result)
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    req = json.loads(Path(args.request).read_text(encoding="utf-8-sig"))
    print(json.dumps(extract_drawing(req, Path(args.output)), ensure_ascii=False, indent=2))
