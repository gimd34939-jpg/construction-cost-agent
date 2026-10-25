"""Source-selected Decimal cost engine for contracts, market prices and fixed inputs."""
from pathlib import Path
from decimal import Decimal, ROUND_DOWN, localcontext
from collections import Counter
import sqlite3,json,hashlib,datetime,re,csv,io
from pricing_policy import (validate_policy, require_reference, reference_authorized,
 inherit_policy, validate_quote_origin, validate_quote_identity, source_period_verified, PricingPolicyError)
OUT=Path(__file__).resolve().parents[1]
DB=OUT/'data/공사단가_통합.db'
class InputError(ValueError):pass
def decimal(v,label='숫자'):
 if v is None or isinstance(v,bool) or isinstance(v,(list,dict)) or str(v).strip()=='':raise InputError(label+' 입력이 필요합니다')
 if not re.fullmatch(r'-?(?:\d+(?:\.\d*)?|\.\d+)',str(v).strip()):raise InputError(label+'에는 유한한 십진수를 입력하세요')
 x=Decimal(str(v).strip())
 if not x.is_finite() or abs(x)>Decimal('1000000000000000'):raise InputError(label+' 범위를 확인하세요')
 return x
def s(v):return format(v,'f')
def trunc(v):return v.quantize(Decimal(1),rounding=ROUND_DOWN)
def canonical_unit(u):
 return {'㎡':'m2','m²':'m2','M2':'m2','㎥':'m3','m³':'m3','M3':'m3','M':'m','㎏':'kg','㎜':'mm','TON':'t','Ton':'t','ton':'t','톤':'t','시간':'hr','ℓ':'L','l':'L','리터':'L','개':'ea','EA':'ea'}.get(str(u).strip(),str(u).strip())
def connect():
 c=sqlite3.connect(DB.as_uri()+'?mode=ro',uri=True);c.row_factory=sqlite3.Row;c.execute('PRAGMA query_only=ON');return c
def get_item(con,identifier):
 if isinstance(identifier,str) and identifier.startswith('API:'):
  from external_store import get_price
  external=get_price(identifier[4:])
  if not external:raise InputError('API 상품 ID를 찾을 수 없습니다')
  return {'record_id':identifier,'route':'EXTERNAL_MATERIAL','name':external['item_name'],'specification':external.get('specification') or '',
   'unit':external['unit'],'scope':external.get('scope'),'signed_quantity_allowed':False,'source_id':external.get('snapshot_id'),
   'source_path':external['source_path'],'source_sha256':external['source_sha256'],'source_origin':'INDEPENDENT_EXTERNAL'}
 row=con.execute('SELECT payload_json FROM app_catalog WHERE record_id=?',(identifier,)).fetchone()
 if row is None:
  if isinstance(identifier,str) and identifier.startswith('EXT:'):
   from assembly_engine import external_item
   external=external_item(identifier)
   if external is not None:return external
  if isinstance(identifier,str) and identifier.startswith('REV:'):
   from working_catalog import get_revision_item
   revision=get_revision_item(identifier)
   if revision is not None:return revision
  raise InputError('선택한 공종 ID를 찾을 수 없습니다')
 return json.loads(row[0])
def search(con,text='',route='',department='',limit=100,version='',scenario=''):
 if route not in ('','CONTRACT','MARKET','PUMSEM'):raise InputError('계산 경로를 확인하세요')
 clauses=[];args=[]
 if route:clauses.append('route=?');args.append(route)
 if department:clauses.append('department=?');args.append(department)
 if version:
  if version not in ['START','CHANGE','CHANGE_NEW']:raise InputError('계약 판본을 확인하세요')
  clauses.append("(route!='CONTRACT' OR json_extract(payload_json,'$.source_id')=?)");args.append(version)
 if scenario=='AUTO':
  clauses.append("(route!='CONTRACT' OR json_extract(payload_json,'$.scenario')=CASE WHEN json_extract(payload_json,'$.source_id')='START' OR department!='토목' THEN 1 ELSE 2 END)")
 elif scenario:
  if scenario not in ['1','2','3','4','5','6']:raise InputError('원본 구간을 확인하세요')
  clauses.append("(route!='CONTRACT' OR json_extract(payload_json,'$.scenario')=?)");args.append(int(scenario))
 if text:
  clauses.append('(name LIKE ? ESCAPE \'\\\' OR specification LIKE ? ESCAPE \'\\\' OR record_id LIKE ? ESCAPE \'\\\')')
  pattern='%'+text.replace('\\','\\\\').replace('%','\\%').replace('_','\\_')+'%';args.extend([pattern]*3)
 where=' WHERE '+' AND '.join(clauses) if clauses else ''
 count=con.execute('SELECT COUNT(*) FROM app_catalog'+where,args).fetchone()[0]
 args+=[max(1,min(int(limit),250))]
 rows=[]
 for r in con.execute('SELECT record_id,route,department,name,specification,unit,status,payload_json FROM app_catalog'+where+' ORDER BY route,department,name,record_id LIMIT ?',args):
  v=dict(r);details=json.loads(v.pop('payload_json'));v.update({k:details.get(k) for k in ['source_label','scenario_label','section']});rows.append(v)
 return {'total':count,'rows':rows}
def price_options(con,query='',category='',limit=80,context=None):
 validate_policy(context)
 options=[]
 # Official basic wages are supported by their published source period.
 if category in ('','labor'):
  for row in con.execute('SELECT * FROM "노임단가" WHERE 단가 IS NOT NULL'):
   r=dict(row)
   if query and query.replace(' ','') not in r['직종명'].replace(' ',''):continue
   options.append({'price_id':r['가격ID'],'category':'labor','item_name':r['직종명'],'specification':f'{r["적용시작일"]} / {r["일작업시간"]}시간 기본급','unit':'인','price':r['단가'],'source_id':r['문서ID'],'price_date':r['적용시작일'],'source_page':r['PDF쪽'],'scope':r['포함범위']})
 if not reference_authorized(context):
  return {'total_shown':len(options[:max(1,min(int(limit),250))]),'rows':options[:max(1,min(int(limit),250))],
   'pricing_policy':'EXTERNAL_PRIMARY','original_observations_suppressed':True,'external_prices_command':'api prices'}
 where=[];params=[]
 if query:where.append('(name LIKE ? OR specification LIKE ?)');params+=['%'+query+'%']*2
 if category:where.append('category=?');params.append(category)
 for row in con.execute('SELECT payload_json FROM app_price_observations'+(' WHERE '+' AND '.join(where) if where else '')+' ORDER BY category,name,price_id LIMIT ?',params+[max(1,min(int(limit),250))]):
  r=json.loads(row[0]);options.append({k:r.get(k) for k in ['price_id','category','item_name','specification','unit','price','source_id','price_date','scope','source_sheet','source_row','cost_components']})
 return {'total_shown':len(options),'rows':options[:max(1,min(int(limit),250))]}
def resolve_price(con,selection,resource,date,allow_historical,context=None):
 validate_policy(context)
 if not isinstance(selection,dict):raise InputError('자원 단가 선택이 필요합니다: '+resource.get('resource_name',''))
 if selection.get('api_price_id'):
  from external_pricing import resolve_api_price
  try:return resolve_api_price(selection,resource,date,context)
  except PricingPolicyError as exc:raise InputError(str(exc)) from exc
 pid=selection.get('price_id')
 if pid:
  r=con.execute('SELECT * FROM "노임단가" WHERE 가격ID=?',(pid,)).fetchone()
  if r:
   r=dict(r)
   if resource['kind']!='labor' or canonical_unit(resource['resource_unit'])!='인':raise InputError('노임 단위/자원 종류가 다릅니다')
   if int(r['일작업시간'])!=8:raise InputError('일 8시간 노임만 이 경로에서 적용할 수 있습니다')
   expected=resource.get('source_leaf_resource_name',resource.get('resource_name',''))
   if expected and re.sub(r'\s+','',expected)!=re.sub(r'\s+','',r['직종명']):raise InputError('품셈 직종과 선택 노임 직종이 다릅니다')
   if date<r['적용시작일'] or (r['적용종료일'] and date>r['적용종료일']) or (date>'2026-10-03' and not source_period_verified(r['문서ID'],date,OUT)):raise InputError('노임 적용기간 또는 최신 공표판본 근거를 확인하세요')
   price=decimal(r['단가'],'노임');return {'price_id':pid,'unit':'인','price':s(price),'cost_components':{'labor':s(price),'material':'0','expense':'0'},'basis':'공표 기본급','source':r,'historical':False}
  row=con.execute('SELECT payload_json FROM app_price_observations WHERE price_id=?',(pid,)).fetchone()
  if row is None:raise InputError('선택 가격 ID를 찾을 수 없습니다')
  r=json.loads(row[0])
  try:require_reference(context)
  except PricingPolicyError as exc:raise InputError(str(exc)) from exc
  expected={'labor':'labor','machine':'machine_hourly','material':'material'}[resource['kind']]
  if r['category']!=expected and not (resource['kind']=='material' and r['category']=='fuel'):raise InputError('단가 종류와 자원 종류가 다릅니다')
  if resource['kind']=='material' and any(tag in str(r.get('price_basis','')) for tag in ['RESOURCE_COMPONENT_PRICE','APPLIED_COMPARISON_MATERIAL']):raise InputError('복합공종의 재료비 구성값은 자재 공급단가로 사용할 수 없습니다')
  if canonical_unit(r['unit'])!=canonical_unit(resource['resource_unit']):raise InputError('단가 단위와 품셈 자원 단위가 다릅니다')
  if not allow_historical:raise InputError('과거 원문 가격은 과거가격 적용을 명시한 경우에만 계산합니다')
  price=decimal(r['price'],'원문 가격')
  if resource['kind']=='machine':
   parts=r.get('cost_components') or {}
   if not all(parts.get(k) is not None for k in ['labor','material','expense']):raise InputError('기계비 3비목 포함범위가 필요합니다')
   if sum(decimal(parts[k],k) for k in ['labor','material','expense'])!=price:raise InputError('기계비 구성 합계가 시간당 가격과 다릅니다')
  else:parts={k:s(price) if k==('labor' if resource['kind']=='labor' else 'material') else '0' for k in ['labor','material','expense']}
  return {'price_id':pid,'unit':r['unit'],'price':s(price),'cost_components':parts,'basis':r.get('price_basis','과거 원문 관찰가격'),'source':r,'historical':True}
 quote=selection.get('quote')
 if not isinstance(quote,dict):raise InputError('단가 선택 또는 견적 정보가 필요합니다')
 try:origin=validate_quote_origin(con,quote,context)
 except PricingPolicyError as exc:raise InputError(str(exc)) from exc
 if quote.get('test_only') is True and (context or {}).get('purpose')!='TEST_ONLY':raise InputError('시험 견적은 실제 작업에 사용할 수 없습니다')
 try:validate_quote_identity(quote,resource,selection,canonical_unit)
 except PricingPolicyError as exc:raise InputError(str(exc)) from exc
 if quote.get('currency','KRW')!='KRW':raise InputError('견적 가격은 원화(KRW)로 확인·환산한 값이 필요합니다')
 required=['item_name','specification','unit','price','source_path','source_sha256','source_location','start_date','end_date','tax_basis','scope']
 if any(quote.get(k) is None or str(quote.get(k)).strip()=='' for k in required):raise InputError('견적 필수 근거가 누락됐습니다')
 for field in ['start_date','end_date']:
  try:datetime.date.fromisoformat(quote[field])
  except Exception:raise InputError('견적 적용일 형식을 확인하세요')
 if not quote['start_date']<=date<=quote['end_date']:raise InputError('견적 적용기간 밖입니다')
 if canonical_unit(quote['unit'])!=canonical_unit(resource['resource_unit']):raise InputError('견적 계산단위가 다릅니다')
 path=Path(quote['source_path'])
 if not path.is_file():raise InputError('견적 원문 파일을 찾을 수 없습니다')
 with path.open('rb') as f:actual=hashlib.file_digest(f,'sha256').hexdigest()
 if actual!=quote['source_sha256']:raise InputError('견적 원문 SHA256이 다릅니다')
 price=decimal(quote['price'],'견적가격');pack=decimal(quote.get('pack_quantity','1'),'포장환산량')
 if price<0 or pack<=0:raise InputError('견적가격/포장환산량을 확인하세요')
 if quote['tax_basis'] not in ['EXCLUDED','INCLUDED']:raise InputError('부가세 포함/별도 확인이 필요합니다')
 if quote['tax_basis']=='INCLUDED':
  vat=decimal(quote.get('vat_percent'),'견적 세율')
  if not 0<=vat<=100:raise InputError('세율 범위를 확인하세요')
  price=price/(1+vat/100)
 price/=pack
 if resource['kind']=='machine':
  parts=quote.get('cost_components') or {}
  keys=['labor','material','expense']
  if not isinstance(parts,dict) or set(parts)!=set(keys):raise InputError('기계비 구성은 재료비·노무비·경비 3개만 입력하세요')
  values={k:decimal(parts[k],k) for k in keys}
  if any(v<0 for v in values.values()) or sum(values.values())!=price:raise InputError('세금·포장 환산 후 음수가 아닌 기계비 3비목 합계가 필요합니다')
  parts={k:s(v) for k,v in values.items()}
 else:parts={k:s(price) if k==('labor' if resource['kind']=='labor' else 'material') else '0' for k in ['labor','material','expense']}
 return {'price_id':'SOURCE_QUOTE:'+actual[:16],'unit':quote['unit'],'price':s(price),'cost_components':parts,'basis':'독립 원문 견적' if origin=='INDEPENDENT_EXTERNAL' else '명시 요청 원문 참고가격','source':quote,'historical':origin!='INDEPENDENT_EXTERNAL','source_origin':origin}
def calculate_line(con,line):
 if not isinstance(line,dict):raise InputError('공종 행 형식을 확인하세요')
 try:validate_policy(line)
 except PricingPolicyError as exc:raise InputError(str(exc)) from exc
 r=get_item(con,line.get('record_id'));q=decimal(line.get('quantity'),'수량')
 if r['route']=='EXTERNAL_MATERIAL' and not line.get('api_price_id'):raise InputError('API 상품 조회만으로 가격을 자동 선택하지 않습니다. api_price_id와 범위 근거를 명시하세요')
 if canonical_unit(line.get('unit'))!=canonical_unit(r['unit']):raise InputError('공종 수량 단위가 다릅니다')
 if q<0 and not (r['route']=='CONTRACT' and r.get('signed_quantity_allowed')):raise InputError('음수 수량은 원본 작업부산물 공제 행에만 허용됩니다')
 if q>0 and r['route']=='CONTRACT' and r.get('signed_quantity_allowed'):raise InputError('원본 작업부산물 공제 행의 수량은 0 이하여야 합니다')
 parts={k:Decimal(0) for k in ['material','labor','expense']};missing=[];resources=[]
 date=line.get('pricing_date')
 if not isinstance(date,str) or not date:raise InputError('실제 가격 적용일 pricing_date가 필요합니다. 구축일을 현재일로 대신 사용하지 않습니다')
 try:datetime.date.fromisoformat(date)
 except Exception:raise InputError('가격 적용일 형식을 확인하세요')
 if line.get('api_price_id'):
  if (line.get('api_mapping') or {}).get('target_role')!='MATERIAL_SUPPLY' or any(word in r['name'] for word in ['설치','타설','부설','철거']):raise InputError('API 직접단가는 근거로 확인된 별도 자재 공급행에만 적용합니다')
  from external_pricing import resolve_api_price
  try:p=resolve_api_price(line,{**r,'kind':'material'},date,line)
  except PricingPolicyError as exc:raise InputError(str(exc)) from exc
  unit_parts=p['cost_components'];unit_total=decimal(p['price']);parts={k:trunc(q*decimal(v)) for k,v in unit_parts.items()};amount=sum(parts.values());status='CALCULATED_HISTORICAL_MATERIAL_SCOPE' if p['historical'] else 'CALCULATED_INDEPENDENT_MATERIAL_SCOPE';resources=[{'input_id':'DIRECT_MATERIAL','resource_name':r['name'],'resource_unit':r['unit'],'resource_quantity':s(q),'price':p}]
 elif r['route']=='CONTRACT':
  try:require_reference(line)
  except PricingPolicyError as exc:raise InputError(str(exc)) from exc
  if r['status']!='VERIFIED_CONTRACT_RATE':raise InputError('원본 산술 불일치 행은 계산할 수 없습니다')
  factor=decimal(r.get('quantity_scale','1'))
  parts={k:trunc(q*decimal(r['unit_components'][k])*factor) for k in parts}
  unit_parts={k:s(decimal(r['unit_components'][k])*factor) for k in parts};unit_total=sum(decimal(v) for v in unit_parts.values());amount=sum(parts.values());status='CALCULATED_CONTRACT_SCOPE'
 elif r['route']=='MARKET':
  if line.get('conditions_confirmed') is not True:raise InputError('시장단가의 원문 포함·제외 조건을 확인해 주세요')
  if date<r['price_date'] or (r['effective_end'] and date>r['effective_end']) or (date>'2026-10-03' and not source_period_verified(r['source_document_id'],date,OUT)):raise InputError('시장단가 적용기간 또는 최신 공표판본 근거를 확인하세요')
  unit_total=decimal(r['unit_total'],'시장단가');amount=trunc(q*unit_total);unit_parts=None;parts=None;status='CALCULATED_MARKET_SCOPE'
 else:
  profile=r['profile']
  if profile['eligibility']!='EXPLICIT_FIXED_COEFFICIENT':raise InputError('이 품셈은 고정계수 경로로 계산할 수 없습니다')
  if line.get('conditions_confirmed') is not True:raise InputError('품셈 조건과 선택 규격을 확인해 주세요')
  if 'allow_historical_prices' in line and type(line['allow_historical_prices']) is not bool:raise InputError('과거가격 적용 선택은 참/거짓이어야 합니다')
  basis=decimal(profile['basis_quantity'],'품셈 기준량')
  if basis<=0:raise InputError('기준량은 양수여야 합니다')
  selected=line.get('price_selections') or {};known_unit={k:Decimal(0) for k in parts};historical=False
  labour_present=any(x['kind']=='labor' for x in profile['inputs'])
  extra=None
  if labour_present:
   try:
    extra=decimal(line.get('labor_additional_percent'),'추가 임금 가산율')
    if not 0<=extra<=1000:raise InputError('추가 임금 가산율 범위를 확인하세요')
   except InputError as e:missing.append({'reason':str(e),'input_id':'LABOR_ADDITIONAL_PERCENT'})
  for resource in profile['inputs']:
   try:p=resolve_price(con,selected.get(resource['input_id']),resource,date,bool(line.get('allow_historical_prices')),line)
   except InputError as e:missing.append({'input_id':resource['input_id'],'resource_name':resource.get('source_leaf_resource_name',resource.get('resource_name','')),'reason':str(e)});continue
   with localcontext() as ctx:
    ctx.prec=50;quantity_per_unit=decimal(resource['coefficient'])/basis
    multiplier=1+(extra or Decimal(0))/100 if resource['kind']=='labor' else Decimal(1)
    contributions={k:quantity_per_unit*decimal(p['cost_components'][k])*multiplier for k in parts}
   for k,v in contributions.items():known_unit[k]+=v
   historical|=p['historical'];resources.append({'input_id':resource['input_id'],'resource_name':resource.get('source_leaf_resource_name',resource.get('resource_name','')),'resource_unit':resource['resource_unit'],'resource_quantity':s(q*quantity_per_unit),'coefficient':resource['coefficient'],'price':p,'unit_cost_components':{k:s(v) for k,v in contributions.items()}})
  machine_parts=[x for x in resources if next(y for y in profile['inputs'] if y['input_id']==x['input_id'])['kind']=='machine']
  labor_names=[x.get('source_leaf_resource_name',x.get('resource_name','')).replace(' ','') for x in profile['inputs'] if x['kind']=='labor']
  fuel_names=[x.get('source_leaf_resource_name',x.get('resource_name','')).replace(' ','') for x in profile['inputs'] if x['kind']=='material']
  if any(decimal(x['price']['cost_components']['labor'])>0 for x in machine_parts) and any('운전' in n for n in labor_names):missing.append({'input_id':'DUPLICATE_OPERATOR_SCOPE','reason':'운전원 포함 기계비와 별도 운전원 투입이 겹칩니다. 포함범위가 맞는 다른 가격을 선택하세요.'})
  if any(decimal(x['price']['cost_components']['material'])>0 for x in machine_parts) and any(n in ['경유','휘발유','등유','중유'] for n in fuel_names):missing.append({'input_id':'DUPLICATE_FUEL_SCOPE','reason':'연료 구성비가 포함된 기계비와 별도 연료 투입을 함께 계산할 수 없습니다.'})
  parts={k:trunc(q*v) for k,v in known_unit.items()};unit_parts={k:s(v) for k,v in known_unit.items()};unit_total=None if missing else sum(known_unit.values());amount=None if missing else sum(parts.values());status='PARTIAL_MISSING_DATA' if missing else ('CALCULATED_HISTORICAL_RESOURCE_SCOPE' if historical else 'CALCULATED_SELECTED_RESOURCE_SCOPE')
 return {'row_id':str(line.get('row_id','')),'record_id':r['record_id'],'route':r['route'],'name':r['name'],'specification':r['specification'],'unit':r['unit'],'quantity':s(q),'signed_quantity_allowed':r.get('signed_quantity_allowed',False),'unit_components':unit_parts,'unit_total':s(unit_total) if unit_total is not None else None,'amount_components':{k:s(v) for k,v in parts.items()} if parts is not None else None,'amount':s(amount) if amount is not None else None,'known_amount':s(sum(parts.values())) if parts is not None else s(amount),'status':status,'missing':missing,'resources':resources,'pricing_date':date,'scope':r.get('scope'),'source':{k:r.get(k) for k in ['source_id','source_label','source_sheet','source_row','source_cells','source_sha256','price_date','conditions','source_document_id','source_page']},'original_input':line}
def calculate_job(job):
 if not isinstance(job,dict) or not isinstance(job.get('lines'),list) or not job['lines']:raise InputError('공종을 한 행 이상 추가해 주세요')
 if len(job['lines'])>3000:raise InputError('한 계산은 3000행까지 지원합니다')
 ids=[x.get('row_id') for x in job['lines']]
 if any(not isinstance(x,str) or not x for x in ids) or len(ids)!=len(set(ids)):raise InputError('각 행에 중복되지 않는 행 ID가 필요합니다')
 validate_policy(job)
 with connect() as con:lines=[calculate_line(con,{**inherit_policy(job,x),'pricing_date':x.get('pricing_date') or job.get('pricing_date'),'purpose':job.get('purpose',x.get('purpose')),'price_data_mode':job.get('price_data_mode','CURRENT_REFRESH')}) for x in job['lines']]
 complete=all(x['amount'] is not None for x in lines)
 groups=Counter(x['route'] for x in lines)
 total=sum(decimal(x['amount']) for x in lines) if complete else None
 return {'schema_version':'construction-cost/job-result/v1','pricing_policy':'EXTERNAL_PRIMARY','contract_reference':job.get('contract_reference'),'price_data_mode':job.get('price_data_mode','CURRENT_REFRESH'),'created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'title':str(job.get('title','공종 단가산출'))[:160],'status':'CALCULATED_SELECTED_SCOPES' if complete else 'INCOMPLETE_PRICE_DATA','line_count':len(lines),'routes':dict(groups),'selected_scope_total':s(total) if total is not None else None,'known_subtotal':s(sum(decimal(x['known_amount']) for x in lines if x['known_amount'] is not None)),'total_scope':'선택 행의 직접 구성금액 합계. 자동 제경비·보험·이윤·부가세·도급액 계산은 포함하지 않음. 원문 단가의 포함·제외 조건을 적용.','mixed_source_warning':len({(x['route'],x['source'].get('source_id')) for x in lines})>1,'database_sha256':hashlib.sha256(DB.read_bytes()).hexdigest(),'lines':lines,'input':job}
def result_csv(result):
 stream=io.StringIO(newline='');writer=csv.writer(stream)
 writer.writerow(['공종','규격','단위','수량','재료단가','노무단가','경비단가','합계단가','재료금액','노무금액','경비금액','합계금액','상태','출처'])
 for r in result['lines']:
  safe=lambda x: "'"+x if isinstance(x,str) and x[:1] in '=+-@' else x
  u=r['unit_components'] or {};a=r['amount_components'] or {}
  writer.writerow([safe(r['name']),safe(r['specification']),r['unit'],r['quantity'],*[u.get(k,'') for k in ['material','labor','expense']],r['unit_total'],*[a.get(k,'') for k in ['material','labor','expense']],r['amount'],r['status'],safe(json.dumps(r['source'],ensure_ascii=False))])
 writer.writerow(['선택 범위 합계','','','','','','','','','','',result['selected_scope_total'],result['status'],result['total_scope']]);return '\ufeff'+stream.getvalue()
