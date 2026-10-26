"""Reproduce public checks without private data, network access, or Excel.

python tools/verify_public.py --output demo-output
Add --report docs/public-checks.json --preview assets/demo-result.svg to refresh
the committed evidence after inspecting a successful run.
"""

import argparse
import ast
from datetime import datetime, timedelta, timezone
import hashlib
import json
from pathlib import Path
import platform
import re
import sys
import unittest
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_links():
    checked, errors = 0, []
    for path in sorted(ROOT.rglob("*.md")):
        if ".git" in path.relative_to(ROOT).parts:
            continue
        fenced = False
        for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if line.lstrip().startswith(("```", "~~~")):
                fenced = not fenced
                continue
            if fenced:
                continue
            for raw in re.findall(r"!?\[[^\]]*\]\(([^)]+)\)", line):
                target = raw.strip().split(' "', 1)[0].strip("<>")
                url = urlsplit(target)
                if url.scheme or url.netloc or not url.path:
                    continue
                checked += 1
                if not (path.parent / unquote(url.path)).exists():
                    errors.append({"file": path.relative_to(ROOT).as_posix(),
                                   "line": line_number, "target": target})
    return {"checked": checked, "errors": errors,
            "scope": "Inline relative Markdown link/image targets exist; anchors and remote URLs are not checked."}


def check_sources():
    manifest = json.loads((ROOT / "docs/source-manifest.json").read_text(encoding="utf-8"))
    errors = []
    for entry in manifest["files"]:
        path = ROOT / entry["file"]
        # The source manifest explicitly normalizes source line endings to LF.
        content = path.read_text(encoding="utf-8").replace("\r\n", "\n").encode("utf-8")
        if hashlib.sha256(content).hexdigest() != entry["sha256"]:
            errors.append(entry["file"])
    return {"checked": len(manifest["files"]), "mismatches": errors,
            "scope": "Published engine files against the export manifest; private originals are not rechecked."}


def make_preview(report):
    demo = report["demo"]
    before, after = demo["baseline_total"], demo["repriced_total"]
    linked = demo["linked_cells_checked"]
    tests = report["tests"]["run"]
    return f'''<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="510" viewBox="0 0 1200 510" role="img" aria-labelledby="title desc">
  <title id="title">공개 합성 데모 실행 결과</title>
  <desc id="desc">합성 공급가격 변경 전 {before}원, 변경 후 {after}원. 가격 누락 시 전체 금액 null. 연결 셀 {linked}개 검사. 실제 Excel 화면이 아닌 실행 JSON의 요약이다.</desc>
  <rect width="1200" height="510" rx="18" fill="#102436"/>
  <g font-family="Malgun Gothic, Apple SD Gothic Neo, sans-serif">
    <text x="42" y="47" fill="#71d6c2" font-size="16">SYNTHETIC / DEMO_ONLY · 1㎡</text>
    <text x="42" y="94" fill="#ffffff" font-size="30" font-weight="700">가격 변경과 누락이 결과에 반영되는지 확인합니다</text>
    <text x="42" y="126" fill="#b8cddd" font-size="16">공개 코드 실행 JSON에서 생성한 요약 · 현장 견적·현재 공표가격·Excel 화면이 아닙니다</text>
    <g fill="#1c384c" stroke="#3e6177">
      <rect x="42" y="164" width="352" height="170" rx="12"/>
      <rect x="424" y="164" width="352" height="170" rx="12"/>
      <rect x="806" y="164" width="352" height="170" rx="12"/>
    </g>
    <g fill="#b8cddd" font-size="17">
      <text x="65" y="197">① 공급가격 12,000원</text>
      <text x="447" y="197">② 공급가격 13,000원</text>
      <text x="829" y="197">③ 합판 가격 누락</text>
    </g>
    <g fill="#ffffff" font-size="35" font-weight="700">
      <text x="65" y="254">{int(before):,}원</text>
      <text x="447" y="254">{int(after):,}원</text>
      <text x="829" y="254">null</text>
    </g>
    <g font-size="16">
      <text x="65" y="304" fill="#71d6c2">COMPLETE · 합성 직접비</text>
      <text x="447" y="304" fill="#71d6c2">COMPLETE · 수식 연결 확인</text>
      <text x="829" y="304" fill="#ffd493">INCOMPLETE · 부분합 별도</text>
    </g>
    <rect x="42" y="361" width="1116" height="99" rx="10" fill="#183043"/>
    <text x="65" y="399" fill="#ffffff" font-size="19">원본 SHA256 보존 · 연결값 {linked}셀 독립 대조 · 공개 테스트 {tests}개 통과</text>
    <text x="65" y="434" fill="#b8cddd" font-size="16">재현: python tools/verify_public.py --output demo-output</text>
    <text x="42" y="490" fill="#9cb4c7" font-size="14">실제 API 조회와 Microsoft Excel 재계산은 이번 공개 검증 범위에 포함하지 않습니다.</text>
  </g>
</svg>
'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="New directory for synthetic artifacts")
    parser.add_argument("--report", type=Path, help="Also refresh an explicitly selected report file")
    parser.add_argument("--preview", type=Path, help="Also refresh an explicitly selected SVG preview")
    args = parser.parse_args()
    if args.output.exists():
        parser.error("Use a new output directory to preserve previous runs")
    suite = unittest.defaultTestLoader.discover(str(ROOT / "tests"))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        return 1
    syntax_files, syntax_errors = [], []
    for folder in ("demo", "engine", "tests", "tools"):
        for path in sorted((ROOT / folder).rglob("*.py")):
            name = path.relative_to(ROOT).as_posix()
            syntax_files.append(name)
            try:
                ast.parse(path.read_text(encoding="utf-8"), filename=name)
            except SyntaxError:
                syntax_errors.append(name)
    links, sources = check_links(), check_sources()
    if syntax_errors or links["errors"] or sources["mismatches"]:
        print(json.dumps({"syntax_errors": syntax_errors, "links": links, "sources": sources}, indent=2))
        return 1
    from demo.__main__ import run
    demo = run(args.output)
    workbook = demo["workbook_verification"]
    code_paths = [path for folder in ("demo", "tests", "tools")
                  for path in sorted((ROOT / folder).rglob("*"))
                  if path.is_file() and path.suffix in (".py", ".json")]
    report = {
        "status": "PASS_REPRODUCIBLE_PUBLIC_CHECKS",
        "checked_on": datetime.now(timezone(timedelta(hours=9))).date().isoformat(),
        "python_version": platform.python_version(), "platform": platform.system(),
        "command": "python tools/verify_public.py --output NEW_DIRECTORY",
        "tests": {"run": result.testsRun, "passed": result.testsRun - len(result.skipped),
                  "failed": 0, "errors": 0, "skipped": len(result.skipped)},
        "python_syntax": {"checked": len(syntax_files), "errors": syntax_errors},
        "relative_links": links, "engine_source_manifest": sources,
        "demo": {"classification": demo["classification"], "purpose": demo["purpose"],
                 "baseline_total": demo["complete_case"]["total"],
                 "repriced_total": demo["price_change_case"]["total"],
                 "missing_price_status": demo["missing_price_case"]["status"],
                 "missing_price_total": demo["missing_price_case"]["total"],
                 "linked_cells_checked": workbook["check_count"],
                 "workbook_verification": workbook},
        "code_sha256": {path.relative_to(ROOT).as_posix(): sha256(path) for path in code_paths},
        "native_excel_rerun_for_review": False, "live_api_price_refresh_for_review": False,
        "private_data_or_historical_reports_revalidated": False,
        "scope": "Public synthetic demo, selected engine functions, source manifest, syntax and local link targets. No claim of complete engine execution, field approval, benefit measurement or exhaustive security audit.",
    }
    serialized = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    (args.output / "public-checks.json").write_text(serialized, encoding="utf-8")
    (args.output / "demo-result.svg").write_text(make_preview(report), encoding="utf-8")
    if args.report:
        args.report.write_text(serialized, encoding="utf-8")
    if args.preview:
        args.preview.write_text(make_preview(report), encoding="utf-8")
    print(f"PASS: {result.testsRun} tests; {len(syntax_files)} Python files; "
          f"{sources['checked']} engine hashes; {links['checked']} relative links")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
