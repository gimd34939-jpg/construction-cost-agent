"""Codex-only DB, calculation and immutable-source revision workflow.

All numeric input is explicit. The knowledge DB is read-only; job history is
stored separately. Editing creates a new workbook and trace package.
"""
from pathlib import Path
from collections import Counter
import argparse,datetime,hashlib,importlib.util,json,os,sqlite3,subprocess,sys,uuid
from cost_engine import OUT,DB,connect,search,price_options,get_item,calculate_line,calculate_job,result_csv,decimal,s,trunc,canonical_unit,InputError
from pricing_policy import validate_policy, require_reference, inherit_policy, validate_quote_origin, PricingPolicyError

def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
 with Path(p).open('rb') as f:return hashlib.file_digest(f,'sha256').hexdigest()
def dump(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def emit(v):print(json.dumps(v,ensure_ascii=False,indent=2,default=str))
def job_folder(title):
 import re
 name=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',str(title))[:50].strip(' .') or '코덱스작업'
 p=OUT/'jobs'/(datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+name+'_'+uuid.uuid4().hex[:6]);p.mkdir(parents=True,exist_ok=False);return p
def history(job_id,kind,title,source,output,payload):
 path=OUT/'data/코덱스_작업이력.db';c=sqlite3.connect(path)
 c.executescript('CREATE TABLE IF NOT EXISTS jobs(job_id TEXT PRIMARY KEY,created_utc TEXT,kind TEXT,title TEXT,source_path TEXT,source_sha256 TEXT,output_path TEXT,payload_json TEXT);')
 c.execute('INSERT INTO jobs VALUES(?,?,?,?,?,?,?,?)',(job_id,datetime.datetime.now(datetime.timezone.utc).isoformat(),kind,title,source.get('path'),source.get('sha256'),str(output),json.dumps(payload,ensure_ascii=False)));c.commit();c.close()
def export_excel(p):
 base=Path(os.environ.get('USERPROFILE',str(Path.home())))/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
 node=base/'node/bin/node.exe'
 if not node.exists():raise InputError('Codex 번들 Node 실행기를 찾지 못했습니다')
 junction=OUT/'node_modules'
 if not junction.exists():
  ps=['powershell','-NoProfile','-NonInteractive','-Command',"New-Item -ItemType Junction -Path '"+str(junction).replace("'","''")+"' -Target '"+str(base/'node/node_modules').replace("'","''")+"' | Out-Null"]
  subprocess.run(ps,check=True,creationflags=0x08000000 if os.name=='nt' else 0)
 cmd=[str(node),str(OUT/'engine/export_workbook.mjs'),str(p/'계산결과.json')]
 r=subprocess.run(cmd,capture_output=True,text=True,encoding='utf-8',timeout=240,creationflags=0x08000000 if os.name=='nt' else 0)
 (p/'Excel생성로그.txt').write_text(r.stdout+'\n'+r.stderr,encoding='utf-8')
 if r.returncode:raise InputError('Excel 생성·재계산 검증 실패. 작업 폴더 로그를 확인하세요: '+str(p))
 return p/'내역산출.xlsx'
def complete_result(title,lines,job):
 complete=all(x['amount'] is not None for x in lines)
 return {'schema_version':'construction-cost/job-result/v1','created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'title':title,'status':'CALCULATED_SELECTED_SCOPES' if complete else 'INCOMPLETE_PRICE_DATA','line_count':len(lines),'routes':dict(Counter(x['route'] for x in lines)),'selected_scope_total':s(sum(decimal(x['amount']) for x in lines)) if complete else None,'known_subtotal':s(sum(decimal(x['known_amount']) for x in lines if x['known_amount'] is not None)),'total_scope':'선택 수정·추가 행의 직접 구성금액. 원본 전체 도급금액과 구별하며 원본 집계는 편집 검증에 기록.','mixed_source_warning':len({(x['route'],x['source'].get('source_id')) for x in lines})>1,'database_sha256':sha(DB),'lines':lines,'input':job}
def override_line(target,op,index,context=None):
 ov=op['unit_price_override'];parts=ov.get('components');basis=ov.get('basis')
 if not isinstance(parts,dict) or set(parts)!=set(['material','labor','expense']):raise InputError('변경 단가는 3비목을 모두 명시해야 합니다')
 if canonical_unit(ov.get('unit'))!=canonical_unit(target['unit']):raise InputError('변경 단가 단위가 원본 행과 다릅니다')
 if not isinstance(basis,dict) or basis.get('type') not in ['USER_INSTRUCTION','SOURCE_FILE','FORMULA_RESULT']:raise InputError('변경 단가의 사용자 지시 또는 원문 근거가 필요합니다')
 for k in ['scope','date']:
  if not isinstance(basis.get(k),str) or not basis[k].strip():raise InputError('가격 변경 근거 '+k+'가 필요합니다')
 try:datetime.date.fromisoformat(basis['date'])
 except Exception:raise InputError('근거 적용일 형식을 확인하세요')
 if basis['type']=='USER_INSTRUCTION':
  if not isinstance(basis.get('instruction'),str) or not basis['instruction'].strip():raise InputError('사용자가 명시한 가격 변경 지시 원문이 필요합니다')
 else:
  bp=Path(basis.get('path',''))
  if not bp.is_file() or sha(bp)!=basis.get('sha256') or not basis.get('location'):raise InputError('가격 근거 파일·SHA·위치를 확인하세요')
  with connect() as con:
   try:validate_quote_origin(con,{**basis,'source_sha256':basis.get('sha256')},context)
   except PricingPolicyError as exc:raise InputError(str(exc)) from exc
 values={k:decimal(v,'변경 '+k+' 단가') for k,v in parts.items()}
 if any(v<0 for v in values.values()):raise InputError('변경 단가는 음수가 될 수 없습니다')
 q=decimal(op.get('quantity',target['reference_quantity']),'변경 수량')
 if q<0 and not target.get('signed_quantity_allowed'):raise InputError('음수 수량을 적용할 수 없습니다')
 if q>0 and target.get('signed_quantity_allowed'):raise InputError('부산물 공제는 0 이하 수량이어야 합니다')
 factor=decimal(target.get('quantity_scale','1'));amounts={k:trunc(q*v*factor) for k,v in values.items()}
 return {'row_id':'EDIT:'+str(index),'record_id':target['record_id'],'route':'USER_OVERRIDE','name':target['name'],'specification':target['specification'],'unit':target['unit'],'quantity':s(q),'signed_quantity_allowed':target.get('signed_quantity_allowed',False),'unit_components':{k:s(v*factor) for k,v in values.items()},'unit_total':s(sum(values.values())*factor),'amount_components':{k:s(v) for k,v in amounts.items()},'amount':s(sum(amounts.values())),'known_amount':s(sum(amounts.values())),'status':'CALCULATED_EXPLICIT_USER_OVERRIDE','missing':[],'resources':[],'pricing_date':basis['date'],'scope':basis['scope'],'source':{k:target.get(k) for k in ['source_id','source_label','source_sheet','source_row','source_cells','source_sha256']},'override_basis':basis,'original_input':op}
def edit_request(req):
 validate_policy(req)
 from external_pricing import refresh_request
 req=refresh_request(req)
 source=req.get('source') or {};path=Path(source.get('path','')).resolve()
 if not path.is_file() or sha(path)!=source.get('sha256'):raise InputError('편집할 원본 파일과 SHA256을 확인하세요')
 if path.suffix.lower()!='.xlsx':raise InputError('이 자동 편집 경로는 .xlsx 파일을 지원합니다. .xls는 원본 보존 변환·별도검증이 필요합니다')
 operations=req.get('operations')
 if not isinstance(operations,list) or not 1<=len(operations)<=100:raise InputError('수정·추가 작업은 1~100개로 명시하세요')
 from working_catalog import signed_rows_for_source
 prepared=[];lines=[];sign_rows=signed_rows_for_source(source['sha256'])
 with connect() as con:
  for saved in con.execute("SELECT payload_json FROM app_catalog WHERE route='CONTRACT' AND json_extract(payload_json,'$.source_sha256')=?",(source['sha256'],)):
   saved=json.loads(saved[0])
   if saved.get('signed_quantity_allowed'):sign_rows.add((saved['source_sheet'],saved['source_row']))
  for i,op in enumerate(operations,1):
   if not isinstance(op,dict) or op.get('operation') not in ['update','add_after']:raise InputError('update 또는 add_after 작업을 지정하세요')
   target=get_item(con,op.get('target_record_id'))
   if target['route']!='CONTRACT' or target['status']!='VERIFIED_CONTRACT_RATE':raise InputError('검증된 실제 내역서 대상 행을 선택하세요')
   if target['source_sha256']!=source['sha256']:raise InputError('DB 원본 행과 편집 원본의 판본이 다릅니다')
   if target.get('signed_quantity_allowed'):sign_rows.add((target['source_sheet'],target['source_row']))
   if 'unit_price_override' in op:
    override_target=target
    if op['operation']=='add_after':
     if 'quantity' not in op:raise InputError('추가 공종의 수량을 명시하세요')
     display=op.get('display') or {}
     if not display.get('name'):raise InputError('추가 공종명을 명시하세요')
     newunit=op['unit_price_override'].get('unit')
     if not newunit:raise InputError('추가 공종 단위를 명시하세요')
     override_target={**target,'name':display['name'],'specification':display.get('specification',''),'unit':newunit,'quantity_scale':'0.01' if canonical_unit(newunit)=='%' else '1','signed_quantity_allowed':False}
    line=override_line(override_target,op,i,req)
   else:
    calc=op.get('calculation_line')
    if calc is None:
     try:require_reference(req)
     except PricingPolicyError as exc:raise InputError('내역서 단가 자동적용 차단: 독립 가격 calculation_line 또는 명시 가격근거가 필요합니다') from exc
     if op['operation']=='add_after' and 'quantity' not in op:raise InputError('추가 공종의 수량을 명시하세요')
     calc={'record_id':target['record_id'],'quantity':op.get('quantity',target['reference_quantity']),'unit':target['unit']}
    calc={**inherit_policy(req,calc),'pricing_date':calc.get('pricing_date') or req.get('pricing_date'),'row_id':'EDIT:'+str(i),'purpose':req.get('purpose'),'price_data_mode':req.get('price_data_mode','CURRENT_REFRESH')};line=calculate_line(con,calc)
   if line['amount'] is None:raise InputError('필수 단가·조건이 누락되어 원본 사본에 적용할 수 없습니다: '+json.dumps(line['missing'],ensure_ascii=False))
   if line['unit_components'] is None:raise InputError('실제 내역서 3비목의 구성 근거가 필요합니다. 전체 시장단가를 임의 분할하지 않습니다')
   if op['operation']=='update' and canonical_unit(line['unit'])!=canonical_unit(target['unit']):raise InputError('선택 계산 단위와 편집 대상 행 단위가 다릅니다')
   display=op.get('display') or {}
   if display and op['operation']!='add_after':raise InputError('기존 행 품명·규격 변경은 이 좁은 수정 경로에서 지원하지 않습니다')
   newunit=line['unit'] if op['operation']=='add_after' else target['unit']
   factor=decimal('0.01' if canonical_unit(newunit)=='%' else '1')
   prepared.append({'operation':op['operation'],'sheet':target['source_sheet'],'row':target['source_row'],'quantity':line['quantity'],'unit_components':{k:s(decimal(v)/factor) for k,v in line['unit_components'].items()},'name':display.get('name',line['name']) if op['operation']=='add_after' else target['name'],'specification':display.get('specification',line['specification']) if op['operation']=='add_after' else target['specification'],'unit':newunit,'quantity_scale':s(factor),'record_id':target['record_id'],'evidence':line,'expected_original':target})
   lines.append(line)
 title=str(req.get('title','코덱스 내역서 수정'))[:160];p=job_folder(title);dump(p/'요청.json',req)
 result=complete_result(title,lines,req);dump(p/'계산결과.json',result);(p/'내역산출.csv').write_text(result_csv(result),encoding='utf-8')
 dump(p/'적용계획.json',prepared)
 from source_workbook_edit import apply_edit
 revision=p/(path.stem+'_코덱스수정본.xlsx')
 report=apply_edit(path,source['sha256'],prepared,revision)
 dump(p/'편집검증.json',report)
 from working_catalog import register_workbook
 for insertion in report['insertions']:
  sign_rows={(sn,row+1 if sn==insertion['sheet'] and row>=insertion['new_row'] else row) for sn,row in sign_rows}
 registration=register_workbook(revision,title,sign_rows=sign_rows);dump(p/'수정본_재등록검증.json',registration)
 excel=export_excel(p)
 package={'status':'EDITED_SOURCE_COPY','revision_id':registration['revision_id'],'registration_report':str(p/'수정본_재등록검증.json'),'source_original':{'path':str(path),'sha256':source['sha256'],'unchanged':sha(path)==source['sha256']},'workbook':str(revision),'workbook_sha256':sha(revision),'calculation_workbook':str(excel),'report':str(p/'편집검증.json'),'request':str(p/'요청.json'),'operations':prepared,'validation':report}
 dump(p/'작업결과.json',package);history(p.name,'EDIT',title,source,revision,package);return package
def formula_record(con,key):
 rows=list(con.execute('SELECT payload_json FROM codex_explicit_formulas WHERE formula_key=? OR formula_id=?',(key,key)))
 if len(rows)!=1:raise InputError('산식 key를 정확히 선택하세요. 같은 원문ID의 여러 분기는 별도 key입니다')
 return json.loads(rows[0][0])
def numeric_request(req):
 source=req.get('source') or {};path=Path(source.get('path','')).resolve()
 if not path.is_file() or sha(path)!=source.get('sha256'):raise InputError('적산 원문 파일과 SHA256을 확인하세요')
 changes=req.get('changes')
 if not isinstance(changes,list) or not 1<=len(changes)<=100:raise InputError('숫자 변경은 1~100개를 명시하세요')
 for change in changes:
  evidence=change.get('evidence') or {}
  if not isinstance(evidence,dict) or not evidence.get('unit') or not evidence.get('scope'):raise InputError('적산 숫자의 단위·적용범위 근거가 필요합니다')
  if evidence.get('type')=='USER_INSTRUCTION':
   if not evidence.get('instruction'):raise InputError('적산 변경에 대한 사용자 지시 원문이 필요합니다')
  elif evidence.get('type') in ['SOURCE_FILE','FORMULA_RESULT']:
   pp=Path(evidence.get('path',''))
   if not pp.is_file() or sha(pp)!=evidence.get('sha256') or not evidence.get('location'):raise InputError('적산 숫자 변경 근거 원문·SHA·위치를 확인하세요')
  elif not (evidence.get('type')=='TEST_ONLY' and req.get('purpose')=='TEST_ONLY'):raise InputError('적산 숫자 변경의 근거 종류를 확인하세요')
 title=str(req.get('title','코덱스 적산 숫자 수정'))[:160];p=job_folder(title);dump(p/'요청.json',req)
 from legacy_cell_edit import apply_numeric_edit
 revised=p/(path.stem+'_코덱스숫자수정본.xlsx');report=apply_numeric_edit(path,source['sha256'],changes,revised);dump(p/'숫자편집검증.json',report)
 package={'status':report['status'],'source_original':{'path':str(path),'sha256':source['sha256'],'unchanged':sha(path)==source['sha256']},'workbook':str(revised),'workbook_sha256':sha(revised),'report':str(p/'숫자편집검증.json'),'request':str(p/'요청.json'),'recalculated_local_cells':report['recalculated_local_cells'],'changes':report['changes'],'conversion':report['conversion'],'scope_warnings':report['source_scope_warnings']}
 dump(p/'작업결과.json',package);history(p.name,'NUMERIC_EDIT',title,source,revised,package);return package
def main():
 if len(sys.argv)>1 and sys.argv[1]=='api':
  from api_cli import main as api_main
  return api_main(sys.argv[2:])
 if len(sys.argv)>1 and sys.argv[1] in {'assemblies','projects','compose','takeoff','compare','drawing'}:
  from assembly_cli import main as assembly_main
  return assembly_main(sys.argv[1:])
 p=argparse.ArgumentParser(description='Codex 건설단가 데이터 조회·계산·원본보존 수정')
 sub=p.add_subparsers(dest='command',required=True)
 sub.add_parser('status')
 q=sub.add_parser('search');q.add_argument('--text',default='');q.add_argument('--route',default='');q.add_argument('--department',default='');q.add_argument('--version',default='');q.add_argument('--scenario',default='');q.add_argument('--revision',default='');q.add_argument('--limit',type=int,default=100)
 q=sub.add_parser('item');q.add_argument('id')
 q=sub.add_parser('variant');q.add_argument('id')
 q=sub.add_parser('candidates');q.add_argument('--input-id',required=True);q.add_argument('--contract-reference-instruction')
 q=sub.add_parser('prices');q.add_argument('--text',default='');q.add_argument('--category',default='');q.add_argument('--limit',type=int,default=100);q.add_argument('--contract-reference-instruction')
 q=sub.add_parser('cell');q.add_argument('--source-id',required=True);q.add_argument('--sheet',required=True);q.add_argument('--address',required=True)
 q=sub.add_parser('model');q.add_argument('--sheet',required=True);q.add_argument('--field');q.add_argument('--equals');q.add_argument('--limit',type=int,default=100)
 q=sub.add_parser('formulas');q.add_argument('--text',default='')
 q=sub.add_parser('formula');q.add_argument('--key',required=True);q.add_argument('--values',required=True)
 q=sub.add_parser('calculate');q.add_argument('--request',required=True);q.add_argument('--excel',action='store_true')
 q=sub.add_parser('edit');q.add_argument('--request',required=True)
 q=sub.add_parser('numeric-edit');q.add_argument('--request',required=True)
 sub.add_parser('history')
 sub.add_parser('revisions')
 q=sub.add_parser('register');q.add_argument('--file',required=True);q.add_argument('--label')
 args=p.parse_args()
 if args.command=='status':
  with connect() as c:
   result={'database':str(DB),'sha256':sha(DB),'read_only':True,'tables':[r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")],'metadata':{r[0]:json.loads(r[1]) for r in c.execute('SELECT * FROM app_build_metadata')}}
 elif args.command=='calculate':
  from external_pricing import refresh_request
  job=refresh_request(read(args.request));result=calculate_job(job);out=job_folder(result['title']);dump(out/'요청.json',job);dump(out/'계산결과.json',result);(out/'내역산출.csv').write_text(result_csv(result),encoding='utf-8')
  if args.excel:export_excel(out)
  history(out.name,'CALCULATE',result['title'],{},out,result);result={'status':result['status'],'folder':str(out),'selected_scope_total':result['selected_scope_total'],'line_count':result['line_count']}
 elif args.command=='edit':
  package=edit_request(read(args.request));result={k:package[k] for k in ['status','revision_id','source_original','workbook','workbook_sha256','calculation_workbook','report','registration_report','request']};result['recalculated_local_cells']=package['validation']['recalculated_local_cells']
 elif args.command=='numeric-edit':
  package=numeric_request(read(args.request));result={k:package[k] for k in ['status','source_original','workbook','workbook_sha256','report','request','recalculated_local_cells','scope_warnings']}
 elif args.command=='register':
  from working_catalog import register_workbook
  metadata=register_workbook(args.file,args.label);result={k:metadata[k] for k in ['revision_id','path','sha256','record_count','verified_count','blocked_count','total_status','source_unchanged']}
 elif args.command=='revisions':
  from working_catalog import revisions
  result=revisions()
 elif args.command=='history':
  path=OUT/'data/코덱스_작업이력.db'
  if not path.exists():result=[]
  else:
   with sqlite3.connect(path.as_uri()+'?mode=ro',uri=True) as c:result=[dict(zip(['job_id','created_utc','kind','title','source_path','source_sha256','output_path'],r)) for r in c.execute('SELECT job_id,created_utc,kind,title,source_path,source_sha256,output_path FROM jobs ORDER BY created_utc DESC LIMIT 100')]
 else:
  with connect() as c:
   if args.command=='search':
    if args.revision:
     from working_catalog import search_revision
     result=search_revision(args.revision,args.text,args.department,args.scenario,args.limit)
    else:result=search(c,args.text,args.route,args.department,args.limit,args.version,args.scenario)
   elif args.command=='item':result=get_item(c,args.id)
   elif args.command=='variant':
    row=c.execute('SELECT payload_json FROM app_semantic_variants WHERE variant_id=?',(args.id,)).fetchone();result=json.loads(row[0]) if row else None
   elif args.command=='candidates':
    if not args.contract_reference_instruction:result={'rows':[],'status':'ORIGINAL_OBSERVATIONS_SUPPRESSED','pricing_policy':'EXTERNAL_PRIMARY','external_prices_command':'api prices'}
    else:result=[dict(r) for r in c.execute('SELECT * FROM codex_resource_price_candidates WHERE input_id=? ORDER BY price_id',(args.input_id,))]
   elif args.command=='prices':result=price_options(c,args.text,args.category,args.limit,{'contract_reference':{'enabled':True,'instruction':args.contract_reference_instruction}} if args.contract_reference_instruction else {})
   elif args.command=='cell':
    row=c.execute('SELECT value_json,formula,source_state FROM codex_workbook_cells WHERE source_id=? AND sheet_name=? AND address=?',(args.source_id,args.sheet,args.address.upper())).fetchone()
    result={'source_id':args.source_id,'sheet':args.sheet,'address':args.address.upper(),'cell':{'value':json.loads(row[0]),'formula':row[1],'state':row[2]} if row else None,'dependents':[dict(r) for r in c.execute('SELECT formula_sheet,formula_cell,state FROM codex_formula_dependencies WHERE source_id=? AND target_sheet=? AND target_cell=?',(args.source_id,args.sheet,args.address.upper()))]}
   elif args.command=='model':
    sql='SELECT payload_json FROM app_model_rows WHERE sheet_name=?';params=[args.sheet]
    if args.field:
     if not args.equals:raise InputError('--field 사용 시 --equals가 필요합니다')
     sql+=' AND json_extract(payload_json,?)=?';params+=['$."'+args.field.replace('"','')+'"',args.equals]
    sql+=' LIMIT ?';params+=[max(1,min(args.limit,500))];result=[json.loads(r[0]) for r in c.execute(sql,params)]
   elif args.command=='formulas':result=[{'key':r['formula_key'],'formula_id':r['formula_id'],'work_id':r['work_id'],'name':r['name'],'pdf_page':r['pdf_page'],'output_unit':r['output_unit'],'variable_contract':json.loads(r['payload_json'])['variable_contract']} for r in c.execute('SELECT * FROM codex_explicit_formulas WHERE name LIKE ? OR work_id LIKE ? ORDER BY formula_key',('%'+args.text+'%',)*2)]
   elif args.command=='formula':
    record=formula_record(c,args.key);typed=read(args.values);contract=record['variable_contract']
    if not isinstance(typed,dict) or set(typed)!=set(contract):raise InputError('산식 변수와 입력 항목이 정확히 일치해야 합니다')
    values={}
    for name,specification in contract.items():
     value=typed[name]
     if not isinstance(value,dict) or set(value)!=set(['value','unit']) or value['unit']!=specification['unit']:raise InputError('산식 변수 '+name+'의 값과 원문 계약 단위를 명시하세요: '+specification['unit'])
     values[name]=s(decimal(value['value'],name))
    path=OUT/'semantic_enrichment/formula_engine.py'
    spec=importlib.util.spec_from_file_location('codex_formula_engine',path);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    result=module.evaluate_formula(record,values);out=job_folder('명시산식_'+record['name']);dump(out/'산식계산.json',{'record':record,'typed_values':typed,'values':values,'result':s(result),'output_unit':record['output_unit']});result={'formula_key':record['formula_id']+'::'+record['name'],'result':s(result),'output_unit':record['output_unit'],'saved':str(out/'산식계산.json')}
 emit(result)
if __name__=='__main__':
 try:sys.exit(main() or 0)
 except (InputError,ValueError,FileNotFoundError) as exc:emit({'status':'BLOCKED_INPUT_OR_SOURCE','reason':str(exc)});sys.exit(2)
