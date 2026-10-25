"""Independent prices are the default; original amounts remain audit data."""
from pathlib import Path
import datetime, hashlib, json, re

class PricingPolicyError(ValueError): pass
_database_hash_cache={}

def reference_authorized(context):
    if not isinstance(context, dict): return False
    ref = context.get('contract_reference')
    return (isinstance(ref, dict) and ref.get('enabled') is True
            and ref.get('type', 'USER_INSTRUCTION') == 'USER_INSTRUCTION'
            and isinstance(ref.get('instruction'), str) and bool(ref['instruction'].strip()))

def historical_reference_authorized(context, price_id=None, pricing_date=None):
    """A dated, selected historical external price needs an actual human instruction."""
    ref = (context or {}).get('historical_reference')
    valid = (isinstance(ref, dict) and ref.get('enabled') is True
             and ref.get('type') == 'USER_INSTRUCTION'
             and isinstance(ref.get('instruction'), str) and bool(ref['instruction'].strip())
             and isinstance(ref.get('price_ids'), list) and bool(ref['price_ids'])
             and all(isinstance(x, str) and x.strip() for x in ref['price_ids'])
             and isinstance(ref.get('pricing_date'), str))
    if not valid: return False
    try: datetime.date.fromisoformat(ref['pricing_date'])
    except ValueError: return False
    return ((price_id is None or price_id in ref['price_ids'])
            and (pricing_date is None or pricing_date == ref['pricing_date']))

def validate_policy(context):
    policy = (context or {}).get('pricing_policy', 'EXTERNAL_PRIMARY')
    if policy != 'EXTERNAL_PRIMARY':
        raise PricingPolicyError('기본 가격정책은 EXTERNAL_PRIMARY입니다. 내역서 참고는 실제 지시를 contract_reference에 기록하세요')
    ref = (context or {}).get('contract_reference')
    if ref is not None and not reference_authorized(context):
        raise PricingPolicyError('내역서 가격 참고에는 enabled=true와 실제 사용자 지시 원문이 필요합니다')
    if (context or {}).get('historical_reference') is not None and not historical_reference_authorized(context):
        raise PricingPolicyError('과거 외부가격 참고에는 실제 사용자 지시 원문·가격 ID 목록·적용일이 필요합니다')
    return policy

def inherit_policy(parent, child):
    validate_policy(parent)
    child = dict(child)
    # A nested calculation cannot invent or override the user's price policy.
    child['pricing_policy'] = 'EXTERNAL_PRIMARY'
    child.pop('contract_reference', None)
    if reference_authorized(parent): child['contract_reference'] = parent['contract_reference']
    child.pop('historical_reference', None)
    if historical_reference_authorized(parent): child['historical_reference'] = parent['historical_reference']
    child.pop('replay_provenance', None)
    if parent.get('replay_provenance') is not None: child['replay_provenance'] = parent['replay_provenance']
    return child

def require_reference(context):
    validate_policy(context)
    if not reference_authorized(context):
        raise PricingPolicyError('내역서 저장가격은 기본 가격근거가 아닙니다. 실제 사용자 명시가 있을 때만 contract_reference로 참고할 수 있습니다')

def verify_evidence(evidence):
    if not isinstance(evidence, dict): raise PricingPolicyError('파일 근거가 필요합니다')
    path = Path(evidence.get('path') or evidence.get('source_path') or '')
    expected = evidence.get('sha256') or evidence.get('source_sha256')
    location = evidence.get('location') or evidence.get('source_location')
    if not path.is_file() or not expected or not location:
        raise PricingPolicyError('가격·규격·포함범위 근거의 파일/SHA/위치가 필요합니다')
    with path.open('rb') as stream: actual = hashlib.file_digest(stream, 'sha256').hexdigest()
    if actual != expected: raise PricingPolicyError('가격·범위 원문 SHA가 다릅니다')
    return {'path': str(path.resolve()), 'sha256': actual, 'location': location}

def original_price_hashes(con):
    hashes = {row[0] for row in con.execute('SELECT sha256 FROM app_sources') if row[0]}
    for row in con.execute("SELECT DISTINCT json_extract(payload_json,'$.source_sha256') FROM app_catalog WHERE route='CONTRACT'"):
        if row[0]: hashes.add(row[0])
    return hashes

def validate_quote_origin(con, quote, context):
    validate_policy(context)
    digest = quote.get('source_sha256') or quote.get('sha256')
    hashes=original_price_hashes(con)
    sources={r[0] for r in con.execute('SELECT source_id FROM app_sources')}
    sources.update(r[0] for r in con.execute("SELECT DISTINCT json_extract(payload_json,'$.source_id') FROM app_catalog WHERE route='CONTRACT'") if r[0])
    paths={str(Path(r[0]).resolve()).casefold() for r in con.execute('SELECT path FROM app_sources') if r[0]}
    for _,_,dbpath in con.execute('PRAGMA database_list'):
        if not dbpath:continue
        p=Path(dbpath);paths.add(str(p.resolve()).casefold())
        stamp=(str(p),p.stat().st_size,p.stat().st_mtime_ns)
        if stamp not in _database_hash_cache:
            with p.open('rb') as stream:_database_hash_cache[stamp]=hashlib.file_digest(stream,'sha256').hexdigest()
        hashes.add(_database_hash_cache[stamp])
    links=[]
    lineage_hash_keys={'parent_source_sha256','upstream_sha256','source_sha256','original_source_sha256','parent_sha256','source_database_sha256','database_sha256','sha256'}
    lineage_id_keys={'original_price_id','observation_price_id','source_price_id','parent_price_id','price_id'}
    def trace(value,location='quote'):
        if isinstance(value,dict):
            for key,child in value.items():
                # Target quantities/specification evidence do not make a new supplier price historical.
                if key in ('target','edit_target','quantity_evidence','scope_evidence','quote_mapping','api_mapping'):continue
                here=location+'.'+key
                if isinstance(child,str):
                    if key in lineage_hash_keys and child in hashes:links.append(here+':'+child)
                    elif key=='source_id' and child in sources:links.append(here+':'+child)
                    elif key in lineage_id_keys and con.execute('SELECT 1 FROM app_price_observations WHERE price_id=?',(child,)).fetchone():links.append(here+':'+child)
                    elif key=='record_id' and con.execute("SELECT 1 FROM app_catalog WHERE record_id=? AND route='CONTRACT'",(child,)).fetchone():links.append(here+':'+child)
                    elif key in ('source_path','parent_source_path','upstream_path','database_path') and str(Path(child).resolve()).casefold() in paths:links.append(here+':'+child)
                if isinstance(child,(dict,list)):trace(child,here)
        elif isinstance(value,list):
            for i,child in enumerate(value):trace(child,location+'['+str(i)+']')
    trace(quote)
    raw_path=Path(quote.get('source_path') or quote.get('path') or '')
    if raw_path.is_file() and raw_path.stat().st_size<=4*1024*1024:
        try:trace(json.loads(raw_path.read_text(encoding='utf-8-sig')),'source_wrapper')
        except (ValueError,UnicodeError,OSError):pass
    original = digest in hashes or bool(links)
    if links:quote['detected_source_lineage']=sorted(set(links))
    if original or quote.get('source_origin') != 'INDEPENDENT_EXTERNAL':
        require_reference(context)
        return 'EXPLICIT_ORIGINAL_REFERENCE'
    return 'INDEPENDENT_EXTERNAL'

def validate_quote_identity(quote, target, selection, canonical_unit):
    """Same-unit unrelated commodities are never interchangeable prices."""
    norm=lambda value:re.sub(r'\s+','',str(value or '')).casefold()
    name=target.get('source_leaf_resource_name') or target.get('resource_name') or target.get('name')
    specification=target.get('resource_specification') or target.get('specification') or ''
    target_unit=target.get('resource_unit') or target.get('unit')
    if norm(quote.get('item_name'))==norm(name) and (not specification or norm(quote.get('specification'))==norm(specification)):
        return
    mapping=(selection or {}).get('quote_mapping')
    if not isinstance(mapping,dict) or mapping.get('scope_confirmed') is not True or not mapping.get('scope_statement'):
        raise PricingPolicyError('견적 품명·규격과 산출 대상이 다릅니다. 파일근거가 있는 quote_mapping이 필요합니다')
    chosen=mapping.get('target') or {}
    if norm(chosen.get('item_name'))!=norm(name) or norm(chosen.get('specification'))!=norm(specification) or canonical_unit(chosen.get('unit'))!=canonical_unit(target_unit):
        raise PricingPolicyError('견적 별칭/규격 매핑의 실제 대상 품명·규격·단위가 다릅니다')
    verify_evidence(mapping.get('evidence'))

def source_period_verified(source_id, date, root):
    """A fresh source-backed edition check may extend a dated official snapshot."""
    path = Path(root) / 'data/source_period_checks.json'
    if not path.is_file(): return False
    data = json.loads(path.read_text(encoding='utf-8-sig'))
    for check in data.get('sources', []):
        if check.get('source_id') != source_id or check.get('latest_confirmed') is not True: continue
        if check.get('status', 'VERIFIED') != 'VERIFIED': continue
        checked = check.get('checked_at', '')[:10]
        try:
            when = datetime.date.fromisoformat(checked)
            start = datetime.date.fromisoformat(check['start_date'])
            applied = datetime.date.fromisoformat(date)
            end = datetime.date.fromisoformat(check['end_date']) if check.get('end_date') else None
        except (ValueError, KeyError, TypeError): continue
        if when > datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date(): continue
        if applied < start or (end is not None and applied > end): continue
        # Latest confirmation is date-specific; never turn an open period into perpetuity.
        if applied > when: continue
        try: verify_evidence(check.get('evidence'))
        except PricingPolicyError: continue
        return True
    return False
