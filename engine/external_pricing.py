"""Strict source-selected independent API/reference prices for the cost engine."""
from copy import deepcopy
import datetime, hashlib, json, re
from pathlib import Path
from decimal import Decimal, localcontext
from pricing_policy import PricingPolicyError, verify_evidence, validate_policy, historical_reference_authorized

def _norm(value): return re.sub(r'\s+', '', str(value or '')).casefold()

def _verified_job_replay(context, record, mapping, date):
    """Replay an unchanged completed calculation, never authorize a new quantity/date."""
    replay = context.get('replay_provenance')
    if replay is None: return None
    if not isinstance(replay, dict) or not replay.get('original_row_id'):
        raise PricingPolicyError('저장 계산 재생에는 원래 완료 결과 파일과 original_row_id가 필요합니다')
    proof = verify_evidence(replay.get('evidence'))
    try: original = json.loads(Path(proof['path']).read_text(encoding='utf-8-sig'))
    except (ValueError, OSError, UnicodeError): raise PricingPolicyError('저장 계산결과 JSON을 확인하세요')
    if (not isinstance(original, dict) or original.get('schema_version') != 'construction-cost/job-result/v1'
            or original.get('status') != 'CALCULATED_SELECTED_SCOPES'
            or original.get('price_data_mode') != 'CURRENT_REFRESH'):
        raise PricingPolicyError('재생 근거는 최초 가격 확인을 마친 CURRENT_REFRESH 완료 계산결과여야 합니다')
    if (original.get('input') or {}).get('purpose') == 'TEST_ONLY' and context.get('purpose') != 'TEST_ONLY':
        raise PricingPolicyError('시험 계산결과를 실제 작업 재생 근거로 사용할 수 없습니다')
    lines = [x for x in original.get('lines', []) if x.get('row_id') == replay['original_row_id']]
    if len(lines) != 1: raise PricingPolicyError('저장 결과에서 고유한 원래 계산 행을 찾을 수 없습니다')
    line = lines[0]
    import cost_engine as cost
    if original.get('database_sha256') != hashlib.sha256(cost.DB.read_bytes()).hexdigest():
        raise PricingPolicyError('저장 계산과 현재의 계산계수·공종 DB 판본이 다릅니다')
    if (line.get('pricing_date') != date or line.get('record_id') != context.get('record_id')
            or cost.canonical_unit(line.get('unit')) != cost.canonical_unit(context.get('unit'))
            or cost.decimal(line.get('quantity'), '원래 수량') != cost.decimal(context.get('quantity'), '재생 수량')):
        raise PricingPolicyError('저장 계산의 가격 적용일·공종 ID·단위·물량과 재생 요청이 다릅니다')
    original_input = line.get('original_input') or {}
    for key in ('conditions_confirmed','labor_additional_percent','allow_historical_prices','price_selections'):
        if original_input.get(key) != context.get(key):
            raise PricingPolicyError('저장 계산의 조건·노임 가산·자원 선택과 재생 요청이 다릅니다: '+key)
    prices = [x.get('price') or {} for x in line.get('resources', [])
              if (x.get('price') or {}).get('price_id') == record['price_id']]
    if (not prices or any(x.get('source_origin') != 'INDEPENDENT_EXTERNAL'
                         or x.get('api_mapping') != mapping
                         or (x.get('source') or {}).get('source_sha256') != record.get('source_sha256')
                         for x in prices)):
        raise PricingPolicyError('저장 계산의 선택 가격 ID·원문 SHA·가격 적용범위와 재생 요청이 다릅니다')
    return {**proof, 'original_row_id': replay['original_row_id'], 'original_pricing_date': date}

def resolve_api_price(selection, target, date, context=None):
    import external_store
    import cost_engine as cost
    context = context or {}
    validate_policy(context)
    if not isinstance(selection, dict) or not selection.get('api_price_id'):
        raise PricingPolicyError('독립 API 가격 ID가 필요합니다')
    if selection.get('_api_refresh_error'):
        raise PricingPolicyError(selection['_api_refresh_error'])
    record = external_store.get_price(selection['api_price_id'])
    if not record: raise PricingPolicyError('독립 API 가격 ID를 찾을 수 없습니다')
    if record.get('source_origin') != 'INDEPENDENT_EXTERNAL':
        raise PricingPolicyError('원문 내역서 관찰가격을 외부 API가격으로 적용할 수 없습니다')
    if record.get('test_only') and context.get('purpose') != 'TEST_ONLY':
        raise PricingPolicyError('시험 API 가격은 실제 작업에 사용할 수 없습니다')
    mapping = selection.get('api_mapping')
    if not isinstance(mapping, dict) or mapping.get('scope_confirmed') is not True or not mapping.get('scope_statement'):
        raise PricingPolicyError('API 상품과 산출 자원의 규격·공급/운반/설치 포함범위 대조가 필요합니다')
    scope_evidence = verify_evidence(mapping.get('evidence'))
    deletion_unexposed = record.get('deleted') is None and record.get('provider') in ('PPS_PRICE','PPS_SHOPPING','OPINET')
    if record.get('deleted') is not False and not deletion_unexposed:
        raise PricingPolicyError('삭제되지 않은 가격인지 확인되지 않았습니다')
    if not record.get('price') or record.get('currency', 'KRW') != 'KRW':
        raise PricingPolicyError('유효한 원화 API가격이 필요합니다')
    kind = target.get('kind', 'material')
    if kind != 'material':
        raise PricingPolicyError('이 API 경로는 자재·연료 공급가격만 지원합니다. 노임/기계 시간당비를 대체하지 않습니다')
    if record.get('price_type') not in ('PUBLIC_REFERENCE', 'PROCUREMENT_CONTRACT', 'FUEL_RETAIL'):
        raise PricingPolicyError('보조 추천상품이나 구입기준가를 기본 공급가격으로 사용할 수 없습니다')
    supplied_unit = cost.canonical_unit(record.get('unit'))
    expected_unit = cost.canonical_unit(target.get('resource_unit') or target.get('unit'))
    if supplied_unit != expected_unit:
        raise PricingPolicyError('API 자재 단위와 산출 단위가 다릅니다. 근거 있는 별도 환산이 필요합니다')
    name = target.get('source_leaf_resource_name') or target.get('resource_name') or target.get('name')
    specification = target.get('specification') or target.get('resource_specification') or ''
    named_mapping = mapping.get('target') or {}
    if _norm(record.get('item_name')) != _norm(name) or (specification and _norm(record.get('specification')) != _norm(specification)):
        if (_norm(named_mapping.get('item_name')) != _norm(name)
                or _norm(named_mapping.get('specification')) != _norm(specification)
                or cost.canonical_unit(named_mapping.get('unit')) != expected_unit):
            raise PricingPolicyError('API 상품과 공종/자원 품명·규격 차이의 명시 대상 매핑이 필요합니다')
    if not mapping.get('delivery_condition') or mapping.get('delivery_condition') not in ('INCLUDED', 'EXCLUDED', 'NOT_APPLICABLE'):
        raise PricingPolicyError('운반/납품 범위를 근거로 INCLUDED/EXCLUDED/NOT_APPLICABLE 중 확인하세요')
    if mapping.get('supply_only') is not True:
        raise PricingPolicyError('API 자재 공급가격에 시공비가 혼합되지 않았는지 확인하세요')
    delivery_raw = str(record.get('delivery_condition') or '')
    if '설치도' in delivery_raw or '설치포함' in delivery_raw:
        raise PricingPolicyError('API 가격에 설치가 포함되어 자재 공급단가로 분리 적용할 수 없습니다')
    known_delivery = ('EXCLUDED' if any(x in delivery_raw for x in ['공장상차','공장도','운반제외','운반비별도'])
                      else 'INCLUDED' if any(x in delivery_raw for x in ['납품장소하차','현장하차','납품장소도','현장도착도','납품장소차상도','운반포함']) else None)
    if known_delivery and mapping['delivery_condition'] != known_delivery:
        raise PricingPolicyError('API 원문의 운반/인도조건과 선택한 적용범위가 충돌합니다')
    if record.get('region') and _norm(mapping.get('region')) != _norm(record['region']):
        raise PricingPolicyError('API 적용지역을 원문과 같은 범위로 확인하세요')
    mode = context.get('price_data_mode', 'CURRENT_REFRESH')
    if mode not in ('CURRENT_REFRESH', 'SNAPSHOT_REPLAY'):
        raise PricingPolicyError('price_data_mode는 CURRENT_REFRESH 또는 SNAPSHOT_REPLAY입니다')
    # A stored snapshot is not an escape from current/as-of selection checks.
    latest = external_store.resolve_latest(record['price_id'], date, department=mapping.get('department'),allow_test_only=context.get('purpose')=='TEST_ONLY')
    latest_ok = latest.get('status') == 'OK' and (latest.get('selected') or {}).get('price_id') == record['price_id']
    if mode == 'CURRENT_REFRESH' and not latest_ok:
        raise PricingPolicyError(latest.get('reason') or '새 요청은 같은 조건의 최신 가격 ID를 선택해야 합니다')
    source = verify_evidence(record)
    applied = datetime.date.fromisoformat(date)
    observed = str(record.get('observed_at') or '')[:10]
    try: observed_day = datetime.date.fromisoformat(observed)
    except ValueError: raise PricingPolicyError('API 조회시각이 없습니다')
    if applied > observed_day:
        raise PricingPolicyError('API 조회 이후 날짜를 확인된 가격으로 적용할 수 없습니다')
    replay_proof = None
    historical = False
    historical_authorized = historical_reference_authorized(context, record['price_id'], date)
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date()
    if mode == 'SNAPSHOT_REPLAY':
        fresh_saved_reference = latest_ok and observed_day == applied == today
        if context.get('replay_provenance') is not None:
            replay_proof = _verified_job_replay(context, record, mapping, date)
        if not fresh_saved_reference:
            if replay_proof is None: replay_proof = _verified_job_replay(context, record, mapping, date)
            if not historical_authorized and replay_proof is None:
                raise PricingPolicyError('오래된 저장가격은 새 현재가격 계산에 사용할 수 없습니다. 실제 historical_reference 지시 또는 동일 적용일·공종·물량의 완료 계산결과 재생 근거가 필요합니다')
            historical = True
    start, end = record.get('effective_start'), record.get('effective_end')
    published = str(record.get('published_at') or '')[:10]
    price_type = record['price_type']
    if price_type == 'PROCUREMENT_CONTRACT':
        if not start or not end: raise PricingPolicyError('조달 계약가격의 적용시작·종료일 미확인')
        if not datetime.date.fromisoformat(start[:10]) <= applied <= datetime.date.fromisoformat(end[:10]):
            raise PricingPolicyError('조달 계약가격 적용기간 밖입니다')
        application_basis = 'VERIFIED_CONTRACT_PERIOD'
    elif price_type == 'PUBLIC_REFERENCE':
        try: publication_day = datetime.date.fromisoformat(published)
        except ValueError: raise PricingPolicyError('공표 자재가격의 게시/기준일이 없습니다')
        if publication_day > applied: raise PricingPolicyError('가격 적용일 이후 공표자료입니다')
        # Publication is a reference date, not an invented quotation validity period.
        if start and datetime.date.fromisoformat(start[:10]) > applied: raise PricingPolicyError('공표 가격 적용시작일 이전입니다')
        if end and datetime.date.fromisoformat(end[:10]) < applied: raise PricingPolicyError('공표 가격 적용종료일 이후입니다')
        application_basis = 'LATEST_PUBLISHED_REFERENCE_AS_OF'
    else:
        if observed_day != applied: raise PricingPolicyError('연료 관찰가격은 실제 조회일의 가격에만 적용합니다')
        try:fuel_day=datetime.date.fromisoformat(published)
        except ValueError:raise PricingPolicyError('연료 가격의 원문 DATE/TRADE_DT 기준일이 없습니다')
        if fuel_day!=applied:raise PricingPolicyError('연료 원문의 가격 기준일이 적용일과 다릅니다. 조회일로 현재가격을 추정하지 않습니다')
        application_basis = 'DATED_FUEL_OBSERVATION'
    if mode == 'SNAPSHOT_REPLAY':
        application_basis = ('VERIFIED_STORED_JOB_REPLAY' if replay_proof else
                             'EXPLICIT_HISTORICAL_EXTERNAL_REFERENCE' if historical else
                             'STORED_REFERENCE_SNAPSHOT_AS_OF')
    tax = record.get('tax_basis', 'UNKNOWN')
    if tax in ('INCLUDED','EXCLUDED') and mapping.get('tax_basis') in ('INCLUDED','EXCLUDED') and mapping['tax_basis'] != tax:
        raise PricingPolicyError('API 원문의 세금 포함범위와 선택한 세금범위가 충돌합니다')
    if tax == 'UNKNOWN': tax = mapping.get('tax_basis', 'UNKNOWN')
    if tax not in ('INCLUDED', 'EXCLUDED'):
        raise PricingPolicyError('API 가격의 부가세 포함/별도 미확인')
    raw = cost.decimal(record['price'], 'API가격')
    pack = cost.decimal(mapping.get('pack_quantity'), '근거 있는 거래당 환산량')
    if raw < 0 or pack <= 0: raise PricingPolicyError('API가격 또는 거래당 환산량 오류')
    vat = None
    if tax == 'INCLUDED':
        vat = cost.decimal(mapping.get('vat_percent'), '근거 있는 부가세율')
        if not 0 <= vat <= 100: raise PricingPolicyError('VAT 비율 오류')
    with localcontext() as decimal_context:
        decimal_context.prec = 50
        price = raw / pack
        if tax == 'INCLUDED': price /= 1 + vat / 100
    # A confirmed supply-only resource has one material component, not an invented installation split.
    parts = {'material': cost.s(price), 'labor': '0', 'expense': '0'}
    return {'price_id': record['price_id'], 'unit': expected_unit, 'price': cost.s(price),
            'cost_components': parts, 'historical': historical,
            'basis': '독립 외부 자재가격 / ' + application_basis,
            'source': {**record, **source}, 'source_origin': 'INDEPENDENT_EXTERNAL',
            'api_mapping': mapping, 'scope_evidence': scope_evidence,
            'price_data_mode': mode, 'historical_reference': context.get('historical_reference') if historical_authorized else None,
            'replay_provenance': replay_proof, 'latest_selection_status': latest.get('status'),
            'latest_selection_reason': latest.get('reason'),
            'application_basis': application_basis, 'date_status': 'REFERENCE_DATE' if not start else 'SOURCE_PERIOD',
            'deletion_status':'NOT_EXPOSED_BY_API' if deletion_unexposed else 'SOURCE_NOT_DELETED'}

def refresh_request(request):
    """Refresh exact selected products once per job; keep failures as missing data."""
    import external_store
    from api_clients import refresh_price_selection
    req = deepcopy(request)
    validate_policy(req)
    mode = req.get('price_data_mode', 'CURRENT_REFRESH')
    if mode == 'SNAPSHOT_REPLAY': return req
    if mode != 'CURRENT_REFRESH': raise PricingPolicyError('가격 데이터 실행모드를 확인하세요')
    cache = {}
    date = (req.get('project') or {}).get('pricing_date') or req.get('pricing_date')
    def visit(value):
        if isinstance(value, dict):
            pid = value.get('api_price_id')
            if pid:
                if pid not in cache:
                    old = external_store.get_price(pid)
                    try:
                        if not old: raise PricingPolicyError('갱신할 API 가격 ID가 없습니다')
                        if (old.get('test_only') is True or old.get('source_network')=='FIXTURE' or old.get('network')=='FIXTURE') and req.get('purpose')!='TEST_ONLY':
                            raise PricingPolicyError('시험 스냅샷 선택은 실제 API 갱신으로 실무 가격이 될 수 없습니다')
                        # Official downloaded reference tables are snapshots with their own
                        # publication/edition checks, not calls to a different API product feed.
                        if old.get('provider') not in ('PPS_PRICE', 'PPS_SHOPPING', 'OPINET'):
                            cache[pid] = {'status': 'REFERENCE_SNAPSHOT', 'selected_id': pid}
                        else:
                            fresh = refresh_price_selection(old, date or datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date().isoformat())
                            if fresh.get('status') not in ('OK', 'PASS', 'VERIFIED', 'SUCCESS', 'OBSERVED') or fresh.get('complete') is not True:
                                raise PricingPolicyError('실시간 API 갱신 미확인: ' + str(fresh.get('status', 'UNKNOWN')))
                            mapping = value.get('api_mapping') or {}
                            def same_scope(candidate):
                                if candidate.get('item_id')!=old.get('item_id') or candidate.get('deleted') is True:return False
                                if candidate.get('test_only') and req.get('purpose')!='TEST_ONLY':return False
                                for key in ('unit','specification','region','tax_basis','delivery_condition','scope','department','operation'):
                                    prior=old.get(key); new=candidate.get(key)
                                    if key=='unit':
                                        import cost_engine
                                        if cost_engine.canonical_unit(prior)!=cost_engine.canonical_unit(new):return False
                                    elif prior not in (None,'','UNKNOWN'):
                                        if _norm(prior)!=_norm(new):return False
                                    elif new not in (None,'','UNKNOWN'):
                                        # New scope information must agree with an explicit sourced choice.
                                        expected=mapping.get('source_scope' if key=='scope' else key)
                                        if key=='operation':expected=old.get('operation')
                                        if expected is None or _norm(expected)!=_norm(new):return False
                                return True
                            hydrated=[external_store.get_price(x['price_id']) or {**x,'operation':fresh.get('operation')} for x in fresh.get('prices', []) if x.get('price_id')]
                            eligible = [x for x in hydrated if same_scope(x)]
                            eligible.sort(key=lambda x: str(x.get('published_at') or ''),reverse=True)
                            if not eligible: raise PricingPolicyError('갱신 응답에 동일 상품의 사용 가능한 가격이 없습니다')
                            latest = external_store.resolve_latest(eligible[0]['price_id'], date or fresh['observed_at'][:10], department=mapping.get('department'),allow_test_only=req.get('purpose')=='TEST_ONLY')
                            if latest.get('status') != 'OK': raise PricingPolicyError(latest.get('reason') or '최신 가격 충돌')
                            cache[pid] = {'status': 'VERIFIED', 'selected_id': latest['selected']['price_id'], 'snapshot_ids': fresh.get('snapshot_ids', []), 'observed_at': fresh.get('observed_at')}
                    except Exception as exc:
                        # Do not copy request URLs or exception bodies that may contain credentials.
                        reason = str(exc) if isinstance(exc, PricingPolicyError) else 'API 갱신 실패; 안전한 API 상태/부족자료 보고서를 확인하세요'
                        cache[pid] = {'status': 'INCOMPLETE', 'reason': reason}
                report = cache[pid]
                value['_api_refresh'] = report
                if report.get('selected_id'):
                    value['_requested_api_price_id'] = pid; value['api_price_id'] = report['selected_id']
                else: value['_api_refresh_error'] = report['reason']
            for child in list(value.values()): visit(child)
        elif isinstance(value, list):
            for child in value: visit(child)
    visit(req)
    req['api_refresh_report'] = cache
    return req
