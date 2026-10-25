# 실제 적산 엔진의 공개용 소스

이 폴더는 2026-10-06 외부가격/API 확장판의 실행 코드 중 핵심 모듈을 공개용으로 정리한 소스 스냅숏입니다. 실제 현장명은 `PROJECT_A`·`PROJECT_B`와 익명 표기로 바꾸고, 내부 `scripts/` 경로를 `engine/`로 옮겼습니다. 변경 범위와 파일 해시는 [소스 매니페스트](../docs/source-manifest.json)에 기록했습니다.

## 읽는 순서

| 관심 내용 | 시작 파일 |
|---|---|
| 요청과 작업 이력 | [codex_cli.py](codex_cli.py) |
| Decimal 계산·비목·미완성 결과 | [cost_engine.py](cost_engine.py) |
| 출처 및 가격 적용 정책 | [pricing_policy.py](pricing_policy.py) |
| 공식 API 수집 | [api_clients.py](api_clients.py) |
| 원문 스냅숏·최신판본 선정 | [external_store.py](external_store.py) |
| 가격과 시공 자원의 연결 | [external_pricing.py](external_pricing.py) |
| 공사 구성·물량·가격 비교 | [assembly_engine.py](assembly_engine.py) |
| 단위와 제한된 산식 평가 | [quantity_engine.py](quantity_engine.py) |
| XLSX의 좁은 수정 | [source_workbook_edit.py](source_workbook_edit.py) |
| 별도 Excel 재계산 | [native_excel_validation.py](native_excel_validation.py) |

## 실행 범위

바로 실행할 수 있는 공개 데모는 저장소 루트의 `python -m demo`입니다. 이 폴더의 전체 CLI는 별도의 적법한 자료와 환경이 필요합니다.

- 현장별 지식 DB·공사구성 규칙 DB·가격 원장과 해당 스키마의 자료
- 원문 파일과 그 SHA256·판본·위치 정보
- Python 패키지 `lxml`, `openpyxl`, `PyMuPDF`
- 조달청·오피넷 API를 호출하는 경우 본인 계정의 키
- Excel 재계산에는 Windows와 설치된 Microsoft Excel
- DWG 변환에는 설치된 AutoCAD Core Console
- 신규 XLSX 출력용 `.mjs`는 Codex 번들의 `@oai/artifact-tool`에 의존

자료 구축기, 실제 DB, 원문 PDF·도면, 인계 자료, 사용자의 키와 보호 저장소는 포함하지 않았습니다. `formula` 명령이 참조하는 비공개 자료 구축용 `semantic_enrichment` 모듈도 공개 범위에 없습니다. 따라서 이 소스 폴더만으로 원래 현장의 전체 실행 환경이 복원된다고 설명하지 않습니다.

`api_credentials.py`는 키를 담은 파일이 아닌 Windows 사용자별 DPAPI 저장 구현입니다. 실제 키를 저장소나 이슈에 입력하지 마세요. 공개 데모는 API 연결이나 인증을 사용하지 않습니다.
