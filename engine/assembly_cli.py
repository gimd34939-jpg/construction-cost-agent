"""Durable CLI jobs for source-backed construction assemblies."""
from pathlib import Path
import argparse,csv,io,json,os,subprocess,sys
import assembly_engine as engine
from codex_cli import read,dump,emit,sha,job_folder,history,edit_request

COMMANDS={'assemblies','projects','compose','takeoff','compare','drawing'}

def source_snapshot(value,found=None):
 found={} if found is None else found
 if isinstance(value,dict):
  raw=value.get('path') or value.get('source_path')
  expected=value.get('sha256') or value.get('source_sha256')
  if isinstance(raw,str) and isinstance(expected,str) and Path(raw).is_file():
   key=str(Path(raw).resolve())
   if key not in found:found[key]={'expected':expected,'before':sha(raw)}
  for child in value.values():source_snapshot(child,found)
 elif isinstance(value,list):
  for child in value:source_snapshot(child,found)
 return found

def drawing_inputs(req,folder):
 from drawing_extract import extract_drawing
 requests=req.get('drawings',[])
 if req.get('drawing'):requests=[req['drawing'],*requests]
 if not isinstance(requests,list):raise engine.AssemblyError('drawings는 배열이어야 합니다')
 known=read(engine.OUT/'data/drawing_sources.json') if (engine.OUT/'data/drawing_sources.json').exists() else []
 if isinstance(known,dict):known=known.get('sources',[])
 registry={x.get('sha256'):x for x in known}
 reports=[];measurements=[];issues=[]
 for request in requests:
  if request.get('project_id')!=req['project']['id']:raise engine.AssemblyError('도면과 작업 현장이 다릅니다')
  identity=registry.get(request.get('sha256'))
  if identity and identity['project_id']!=req['project']['id']:raise engine.AssemblyError('등록된 도면 표제의 현장과 요청이 다릅니다')
  if identity and identity.get('revision')!=request.get('revision'):raise engine.AssemblyError('등록된 도면 판본 식별자와 요청이 다릅니다')
  report=extract_drawing(request,folder)
  report['registered_identity']=identity
  if not identity:
   issues.append({'reason':'도면 표제·현장·판본의 검증된 등록 근거가 필요합니다','source':report['source']})
   for m in report['measurements']:
    m.update(status='UNVERIFIED',value=None,identity_pending=True)
  reports.append(report);measurements.extend(report['measurements'])
 ids=[x['measurement_id'] for x in measurements]
 if len(ids)!=len(set(ids)):raise engine.AssemblyError('도면들 사이 measurement_id가 중복됩니다')
 return reports,measurements,issues

def comparison_csv(result):
 output=io.StringIO(newline='');w=csv.writer(output)
 w.writerow(['공종ID','공종','규격','단위','물량','계약/저장가 금액','신규금액','차이','계약상태','신규상태'])
 safe=lambda v:("'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v)
 for x in result['rows']:
  w.writerow([safe(v) for v in [x['line_id'],x['name'],x['specification'],x['unit'],x['quantity'],x['contract'].get('amount'),x['current'].get('amount'),x['difference'],x['contract']['status'],x['current']['status']]])
 return output.getvalue()

def export_comparison(folder):
 deps=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
 junction=engine.OUT/'node_modules'
 if not junction.exists():
  subprocess.run(['powershell','-NoProfile','-NonInteractive','-Command',"New-Item -ItemType Junction -Path '"+str(junction).replace("'","''")+"' -Target '"+str(deps/'node/node_modules').replace("'","''")+"' | Out-Null"],check=True,creationflags=0x08000000)
 done=subprocess.run([str(deps/'node/bin/node.exe'),str(engine.OUT/'engine/export_comparison.mjs'),str(folder/'가격비교.json')],capture_output=True,text=True,encoding='utf-8',timeout=240,creationflags=0x08000000)
 (folder/'Excel생성로그.txt').write_text(done.stdout+'\n'+done.stderr,encoding='utf-8')
 if done.returncode:raise engine.AssemblyError('비교 Excel 생성 또는 수식 검증 실패: '+str(folder))
 return str(folder/'가격비교.xlsx')

def edit_plan(req,result,report_path):
 target=req.get('edit_target')
 if not target:return None
 side=target.get('price_side')
 if side not in ('contract','current'):raise engine.AssemblyError('edit_target.price_side는 contract/current로 명시하세요')
 if result['takeoff_status']!='TAKEOFF_COMPLETE' or not result['scope_confirmed']:
  raise engine.AssemblyError('구성 조건·물량·포함범위가 미확정이므로 편집할 수 없습니다')
 source=target.get('source') or {};engine.verify_file({**source,'location':'편집 대상 원문'})
 rows={x['line_id']:x for x in result['rows']};operations=[]
 requested=target.get('operations')
 if not isinstance(requested,list) or not requested:raise engine.AssemblyError('내역서에 적용할 행 operations를 명시하세요')
 seen=set()
 for op in requested:
  key=op.get('line_id')
  if key in seen:raise engine.AssemblyError('같은 산출 행을 중복 반영할 수 없습니다')
  seen.add(key)
  if key not in rows:raise engine.AssemblyError('존재하지 않는 산출 행입니다')
  row=rows[key];price=row[side]
  if price.get('amount') is None or not price.get('unit_components'):raise engine.AssemblyError('확정된 3비목 가격 근거가 필요합니다: '+key)
  base=engine.get_record(op.get('target_record_id'))
  if engine.record_project(base)!=result['project']['id'] or base.get('source_sha256')!=source['sha256']:
   raise engine.AssemblyError('편집 대상의 현장·원문 판본이 다릅니다')
  calculation_record=engine.get_record(row['contract'].get('record_id'))
  if base.get('scenario')!=result['project'].get('scenario') or base.get('department')!=calculation_record.get('department'):
   raise engine.AssemblyError('편집 대상의 총괄/연차 구간 또는 부서가 다릅니다')
  if op.get('operation')=='update' and row['contract'].get('record_id')!=base['record_id']:
   raise engine.AssemblyError('기존 행 수정은 계산 대상과 같은 행이어야 합니다')
  if op.get('operation') not in ('update','add_after'):raise engine.AssemblyError('update/add_after만 지원합니다')
  units=price['unit_components']
  factor=engine.cost.decimal(base.get('quantity_scale','1')) if op['operation']=='update' else engine.cost.decimal('1')
  operation={'operation':op['operation'],'target_record_id':base['record_id'],'quantity':row['quantity'],
    'unit_price_override':{'unit':base['unit'] if op['operation']=='update' else price['unit'],
      'components':{k:str(engine.cost.decimal(v)/factor) for k,v in units.items()},
      'basis':{'type':'FORMULA_RESULT','source_origin':price.get('source_origin','EXPLICIT_ORIGINAL_REFERENCE'),'path':str(report_path),'sha256':sha(report_path),'location':'rows[line_id='+key+'].'+side,'date':result['pricing_date'],'scope':result['scope']+' / '+price.get('scope_statement',price.get('basis','선택 3비목'))}}}
  if op['operation']=='add_after':operation['display']={'name':row['name'],'specification':row['specification']}
  operations.append(operation)
 return {'title':req.get('title','공사 구성 비교')+'_내역서반영','purpose':req.get('purpose','PROJECT_WORK'),'pricing_policy':'EXTERNAL_PRIMARY','contract_reference':req.get('contract_reference'),'price_data_mode':'SNAPSHOT_REPLAY','source':source,'operations':operations}

def run_job(command,req,excel=False,apply=False):
 folder=job_folder(req.get('title',command));dump(folder/'요청.json',req)
 before={'core':sha(engine.cost.DB),'assembly':sha(engine.RULE_DB)}
 originals=source_snapshot(req)
 package={'command':command,'folder':str(folder),'request':str(folder/'요청.json')}
 try:
  if command=='compare':
   from external_pricing import refresh_request
   req=refresh_request(req);dump(folder/'가격선택_적용요청.json',req)
  if command=='drawing':
   from drawing_extract import extract_drawing
   result=extract_drawing(req,folder);dump(folder/'도면추출.json',result)
  else:
   composition=engine.compose(req);dump(folder/'공종구성.json',composition)
   source_snapshot(composition,originals)
   result=composition
   if command in ('takeoff','compare'):
    reports,measurements,issues=drawing_inputs(req,folder)
    take=engine.takeoff(req,composition,measurements)
    if issues:take['missing'].extend(issues);take['status']='INCOMPLETE'
    conflicts_file=engine.OUT/'data/drawing_conflicts.json'
    conflicts=read(conflicts_file) if conflicts_file.exists() else []
    if isinstance(conflicts,dict):conflicts=conflicts.get('conflicts',[])
    selected_shas={r['source']['sha256'] for r in reports}
    take['drawing_conflicts']=[c for c in conflicts if c.get('source_sha256') in selected_shas]
    for conflict in take['drawing_conflicts']:
     component_ids=conflict.get('component_ids',[])
     for line in take['lines']:
      if line['component_id'] not in component_ids:continue
      resolution=(req.get('conflict_resolutions') or {}).get(conflict.get('conflict_id'))
      if resolution:
       if not resolution.get('reason'):raise engine.AssemblyError('충돌 적용 이유가 필요합니다')
       ev=resolution.get('evidence') or {}
       if ev.get('type')=='USER_INSTRUCTION':
        if not ev.get('instruction'):raise engine.AssemblyError('실제 사용자 지시 원문이 필요합니다')
       else:engine.verify_file(ev)
       from quantity_engine import evaluate_quantity
       expected=resolution.get('quantity')
       if not isinstance(expected,dict):raise engine.AssemblyError('충돌 해결 후 적용 물량 quantity(value/unit)를 명시하세요')
       verified=evaluate_quantity({'expression':'q','output_unit':line['unit'],'variables':{'q':{'unit':line['unit']}}},{'q':{**expected,'evidence':ev}})
       if verified.get('test_only') and req.get('purpose')!='TEST_ONLY':raise engine.AssemblyError('시험용 충돌 해결 근거는 실제 작업에 사용할 수 없습니다')
       if verified.get('status')!='COMPLETE' or line.get('quantity') is None or engine.cost.decimal(verified['value'])!=engine.cost.decimal(line['quantity']):
        raise engine.AssemblyError('충돌 해결 근거의 적용 물량과 실제 산출 물량이 다릅니다')
       line['conflict_resolution']=resolution
      else:
       line.update(quantity=None,quantity_status='INCOMPLETE')
       take['missing'].append({'line_id':line['line_id'],'reason':'도면 원문 근거 충돌','details':conflict});take['status']='INCOMPLETE'
    take['drawing_reports']=reports;dump(folder/'물량산출.json',take);result=take
   if command=='compare':
    result=engine.compare(req,take);result['drawing_conflicts']=take.get('drawing_conflicts',[])
    source_snapshot(result,originals)
    dump(folder/'가격비교.json',result);(folder/'가격비교.csv').write_text(comparison_csv(result),encoding='utf-8-sig')
    if excel:package['comparison_workbook']=export_comparison(folder)
    if req.get('edit_target'):
     plan=edit_plan(req,result,folder/'가격비교.json');dump(folder/'편집요청.json',plan);package['edit_plan']=str(folder/'편집요청.json')
     if apply:package['edited_copy']=edit_request(plan)
    elif apply:raise engine.AssemblyError('--edit에는 edit_target이 필요합니다')
  package.update(status=result['status'],result=result)
  if package.get('comparison_workbook') or package.get('edited_copy'):
   from native_excel_validation import verify_package
   package['native_excel_validation']=verify_package(package,folder/'native_excel')
   if package['native_excel_validation']['status']!='PASS':package['status']='VERIFICATION_FAILED'
 except Exception as exc:
  package.update(status='BLOCKED',reason=str(exc))
 finally:
  package['database_preservation']={'before':before,'after':{'core':sha(engine.cost.DB),'assembly':sha(engine.RULE_DB)}}
  package['database_preservation']['unchanged']=before==package['database_preservation']['after']
  if not package['database_preservation']['unchanged']:package.update(status='BLOCKED',reason='기준 DB 해시 변경')
  for path,entry in originals.items():
   entry['after']=sha(path) if Path(path).is_file() else None
   entry['unchanged']=entry['before']==entry['after']==entry['expected']
  package['originals_preservation']=originals
  if any(not x['unchanged'] for x in originals.values()):package.update(status='BLOCKED',reason='원문 파일 판본 또는 보존 SHA 불일치')
  dump(folder/'작업결과.json',package)
  history(folder.name,command.upper(),req.get('title',command),{},folder,package)
 return package

def main(argv=None):
 p=argparse.ArgumentParser(description='공사 구성·물량·동일 범위 가격 비교')
 sub=p.add_subparsers(dest='command',required=True)
 a=sub.add_parser('assemblies');a.add_argument('--text',default='');a.add_argument('--id')
 sub.add_parser('projects')
 for command in ('compose','takeoff','compare','drawing'):
  a=sub.add_parser(command);a.add_argument('--request',required=True)
  if command=='compare':a.add_argument('--excel',action='store_true');a.add_argument('--edit',action='store_true')
 args=p.parse_args(argv)
 if args.command=='assemblies':
  if args.id:
   with engine.rules() as c:r=c.execute('SELECT payload_json FROM assembly_types WHERE type_id=?',(args.id,)).fetchone()
   result=json.loads(r[0]) if r else {'status':'NOT_FOUND'}
  else:
   result=engine.list_types(args.text)
   capability_path=engine.OUT/'data/capability_status.json'
   if capability_path.exists():
    capability={x['type_id']:x for x in read(capability_path)['types']}
    result=[{**x,'capability':capability.get(x['type_id'])} for x in result]
 elif args.command=='projects':result=engine.project_list()
 else:result=run_job(args.command,read(args.request),getattr(args,'excel',False),getattr(args,'edit',False))
 emit(result);return 2 if isinstance(result,dict) and result.get('status') in ('BLOCKED','VERIFICATION_FAILED') else 0

if __name__=='__main__':
 try:raise SystemExit(main())
 except (ValueError,OSError,KeyError) as e:emit({'status':'BLOCKED','reason':str(e)});raise SystemExit(2)
