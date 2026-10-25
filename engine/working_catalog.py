"""Read-only intake of subsequent workbook revisions into a separate work DB."""
from pathlib import Path
from decimal import Decimal,localcontext
import datetime,hashlib,json,sqlite3
from build_reference_catalog import load,val,normal,num,trunc,textnum,cell,ast,membership_engine
from cost_engine import OUT,InputError

HISTORY=OUT/'data/코덱스_작업이력.db'
def sha(p):
 with Path(p).open('rb') as f:return hashlib.file_digest(f,'sha256').hexdigest()
def ensure(c):
 c.executescript('''CREATE TABLE IF NOT EXISTS revision_sources(revision_id TEXT PRIMARY KEY,path TEXT NOT NULL,sha256 TEXT NOT NULL,created_utc TEXT,metadata_json TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS revision_catalog(record_id TEXT PRIMARY KEY,revision_id TEXT NOT NULL,route TEXT,name TEXT,specification TEXT,department TEXT,unit TEXT,status TEXT,payload_json TEXT NOT NULL);
 CREATE INDEX IF NOT EXISTS revision_search ON revision_catalog(revision_id,department,name);''')
def register_workbook(path,label=None,sign_rows=None):
 p=Path(path).resolve()
 if not p.is_file() or p.suffix.lower()!='.xlsx':raise InputError('등록은 실제 XLSX 원문 파일을 읽기 전용으로 사용합니다')
 before=sha(p);rid='REV:'+before[:20];sheets=load(p);records=[];issues=[];memberships=[]
 inherited_signs=set(sign_rows or [])
 if HISTORY.exists():
  with sqlite3.connect(HISTORY.as_uri()+'?mode=ro',uri=True) as prior:
   try:
    for entry in prior.execute('SELECT payload_json FROM revision_catalog WHERE revision_id=?',(rid,)):
     entry=json.loads(entry[0])
     if entry.get('signed_quantity_allowed'):inherited_signs.add((entry['source_sheet'],entry['source_row']))
   except sqlite3.OperationalError:pass
 for sn,cells in sheets.items():
  if sn not in ['건축내역서','기계내역서','토목내역서']:continue
  layout=membership_engine._layout(cells);nc=layout['name'];pc=layout['price']
  hr=None;qc=uc=None
  for row in range(1,9):
   vals={c['col']:normal(c['value']) for c in cells.values() if c['row']==row}
   q=[col for col,v in vals.items() if v=='수량'];u=[col for col,v in vals.items() if v=='단위']
   if len(q)==len(u)==1:hr=row;qc=q[0];uc=u[0];break
  if hr is None:raise InputError('등록할 내역서 머리글을 확정할 수 없습니다: '+sn)
  rows=set()
  for rr in range(hr+2,max(c['row'] for c in cells.values())+1):
   name=val(cells,nc,rr);unit=val(cells,uc,rr)
   if not isinstance(name,str) or not str(unit or '').strip() or normal(unit) in ['단위','수량','규격']:continue
   try:
    amounts=[num(val(cells,pc+2*j+1,rr)) for j in range(3)];total=num(val(cells,pc+7,rr));ut=num(val(cells,pc+6,rr));rawq=val(cells,qc,rr)
    q=Decimal(0) if rawq is None and all(v==0 for v in amounts) else num(rawq)
    states={};parts={}
    for j,k in enumerate(['material','labor','expense']):
     value=val(cells,pc+2*j,rr)
     if value is None and amounts[j]==0:parts[k]=Decimal(0);states[k]='SOURCE_BLANK_COMPONENT_WITH_ZERO_SAVED_AMOUNT'
     else:parts[k]=num(value);states[k]='SOURCE_NUMERIC_COMPONENT'
   except ValueError as exc:issues.append({'sheet':sn,'row':rr,'reason':str(exc)});continue
   quantity_evidence=[];f=cell(cells,qc,rr).get('formula')
   if f:
    try:
     with localcontext() as ctx:
      ctx.prec=50;fresh=ast.Evaluator(sheets,sn,recursive=True).evaluate(ast.AST(f,sheets.keys()).parse())
     if isinstance(fresh,Decimal) and abs(fresh-q)<=Decimal('0.00000000001'):
      if fresh!=q:quantity_evidence.append({'cached':textnum(q),'fresh':textnum(fresh),'formula':f})
      q=fresh
    except Exception:pass
   factor=Decimal('0.01') if normal(unit)=='%' else Decimal(1)
   expected=[trunc(q*parts[k]*factor) for k in ['material','labor','expense']]
   checks={'component_amounts':expected==amounts,'unit_sum':sum(parts.values())==ut,'amount_sum':sum(amounts)==total,'nonnegative_unit_prices':all(v>=0 for v in parts.values())}
   status='VERIFIED_CONTRACT_RATE' if all(checks.values()) else 'BLOCKED_SOURCE_ARITHMETIC'
   scenario=val(cells,pc+9,rr);scenario=int(num(scenario)) if scenario not in [None,''] else None
   record={'record_id':f'{rid}:{sn}:{rr}','route':'CONTRACT','source_id':rid,'source_label':label or p.stem,'department':sn.replace('내역서',''),'name':name.strip(),'specification':str(val(cells,nc+1,rr) or ''),'unit':str(unit).strip(),'reference_quantity':textnum(q),'reference_quantity_raw':rawq,'quantity_is_source_blank':rawq is None,'signed_quantity_allowed':q<0,'quantity_scale':textnum(factor),'unit_components':{k:textnum(v) for k,v in parts.items()},'component_source_states':states,'unit_total':textnum(ut),'reference_amounts':dict(zip(['material','labor','expense'],map(textnum,amounts))),'reference_total':textnum(total),'scenario':scenario,'scenario_label':'등록 수정본의 원문 구간 '+str(scenario),'section':None,'source_path':str(p),'source_sha256':before,'source_sheet':sn,'source_row':rr,'source_cells':f'원문 {sn} {rr}행','price_date':None,'price_basis':'등록 내역서의 실제 저장 단가. 적용일 별도 확인.','scope':'등록 파일 원문 해당 행의 범위','rounding':'TRUNC_COMPONENT_AMOUNT_0_THEN_SUM','checks':checks,'status':status,'quantity_recompute_evidence':quantity_evidence}
   records.append(record);rows.add(rr)
   if (sn,rr) in inherited_signs:record['signed_quantity_allowed']=True;record['signed_quantity_basis']='INHERITED_SOURCE_BYPRODUCT_ROLE'
  membership=membership_engine.direct_membership(sheets,rid,sn,rows,pc);memberships.append({'sheet':sn,**membership});direct=set(membership['terminal_rows'])
  for r in records:
   if r['source_sheet']==sn:r['in_source_direct_total']=r['source_row'] in direct
 if not records:raise InputError('지원하는 검증 가능한 내역 항목이 없습니다')
 assert sha(p)==before,'등록 중 원본 변경'
 metadata={'revision_id':rid,'path':str(p),'sha256':before,'label':label or p.stem,'record_count':len(records),'verified_count':sum(r['status']=='VERIFIED_CONTRACT_RATE' for r in records),'blocked_count':sum(r['status']!='VERIFIED_CONTRACT_RATE' for r in records),'issues':issues,'source_memberships':memberships,'source_unchanged':True,'status':'REGISTERED_WITH_ROW_VALIDATION','total_status':'VERIFIED_SOURCE_DIRECT_MEMBERSHIP' if all(x['status']=='VERIFIED_SOURCE_DIRECT_MEMBERSHIP' for x in memberships) else 'UNRESOLVED_SOURCE_TOTAL'}
 with sqlite3.connect(HISTORY) as c:
  ensure(c);c.execute('INSERT OR REPLACE INTO revision_sources VALUES(?,?,?,?,?)',(rid,str(p),before,datetime.datetime.now(datetime.timezone.utc).isoformat(),json.dumps(metadata,ensure_ascii=False,default=str)))
  c.execute('DELETE FROM revision_catalog WHERE revision_id=?',(rid,))
  c.executemany('INSERT INTO revision_catalog VALUES(?,?,?,?,?,?,?,?,?)',[(r['record_id'],rid,r['route'],r['name'],r['specification'],r['department'],r['unit'],r['status'],json.dumps(r,ensure_ascii=False,default=str)) for r in records])
 return metadata
def get_revision_item(identifier):
 if not HISTORY.exists():return None
 with sqlite3.connect(HISTORY.as_uri()+'?mode=ro',uri=True) as c:
  try:row=c.execute('SELECT payload_json FROM revision_catalog WHERE record_id=?',(identifier,)).fetchone()
  except sqlite3.OperationalError:return None
 return json.loads(row[0]) if row else None
def signed_rows_for_source(source_sha256):
 if not HISTORY.exists():return set()
 with sqlite3.connect(HISTORY.as_uri()+'?mode=ro',uri=True) as c:
  try:rows=c.execute('SELECT r.payload_json FROM revision_catalog r JOIN revision_sources s ON s.revision_id=r.revision_id WHERE s.sha256=?',(source_sha256,)).fetchall()
  except sqlite3.OperationalError:return set()
 return {(v['source_sheet'],v['source_row']) for row in rows if (v:=json.loads(row[0])).get('signed_quantity_allowed')}
def revisions():
 if not HISTORY.exists():return []
 with sqlite3.connect(HISTORY.as_uri()+'?mode=ro',uri=True) as c:
  try:return [{'revision_id':r[0],'path':r[1],'sha256':r[2],'created_utc':r[3],'record_count':json.loads(r[4])['record_count']} for r in c.execute('SELECT * FROM revision_sources ORDER BY created_utc DESC')]
  except sqlite3.OperationalError:return []
def search_revision(revision_id,text='',department='',scenario='',limit=100):
 if not HISTORY.exists():raise InputError('등록한 작업 판본이 없습니다')
 sql=' FROM revision_catalog WHERE revision_id=?';params=[revision_id]
 if text:sql+=' AND (name LIKE ? OR specification LIKE ?)';params+=['%'+text+'%']*2
 if department:sql+=' AND department=?';params.append(department)
 if scenario:
  if scenario not in ['1','2','3','4','5','6']:raise InputError('등록 판본의 구간 번호를 지정하세요')
  sql+=" AND json_extract(payload_json,'$.scenario')=?";params.append(int(scenario))
 with sqlite3.connect(HISTORY.as_uri()+'?mode=ro',uri=True) as c:
  total=c.execute('SELECT COUNT(*)'+sql,params).fetchone()[0]
  rows=[json.loads(r[0]) for r in c.execute('SELECT payload_json'+sql+' ORDER BY department,name,record_id LIMIT ?',params+[max(1,min(limit,250))])]
 return {'revision_id':revision_id,'total':total,'rows':[{k:r.get(k) for k in ['record_id','route','department','name','specification','unit','status','reference_quantity','unit_total','reference_total','scenario','source_path']} for r in rows]}
