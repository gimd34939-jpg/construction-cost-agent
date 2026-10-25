"""Read-only full-text discovery of historical design review cases."""
from pathlib import Path
import sqlite3, hashlib, json

ROOT=Path(__file__).resolve().parents[1]
def search_cases(text='',limit=20):
    if not 1<=int(limit)<=100: raise ValueError('사례 검색 limit는1~100입니다')
    c=sqlite3.connect((ROOT/'data/외부가격_원장.db').as_uri()+'?mode=ro',uri=True)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA query_only=ON')
    try: documents=[dict(r) for r in c.execute('SELECT * FROM historical_documents ORDER BY member_name')]
    finally: c.close()
    matches=[];total=0;issues=[]
    for doc in documents:
        text_path=Path(doc['text_path'])
        if not text_path.is_file():
            issues.append({'document_id':doc['document_id'],'reason':'추출 텍스트 없음'}); continue
        metadata=json.loads(doc['metadata_json'])
        with text_path.open(encoding='utf-8-sig') as f:
            for line in f:
                row=json.loads(line)
                values=[str(x['value']) for x in row.get('cells',[])] if 'cells' in row else [str(x) for x in row.get('values',[])]
                rendered=' | '.join(values) or row.get('text','')
                if text and text.casefold() not in rendered.casefold() and text.casefold() not in doc['member_name'].casefold(): continue
                total+=1
                if len(matches)>=int(limit): continue
                source=Path(doc['extracted_path'])
                if hashlib.sha256(source.read_bytes()).hexdigest()!=doc['sha256']:
                    issues.append({'document_id':doc['document_id'],'reason':'사례 원문 SHA 불일치'}); continue
                matches.append({'document_id':doc['document_id'],'source_origin':'HISTORICAL_CASE',
                  'historical_period':doc['historical_period'],'member_name':doc['member_name'],
                  'source_path':str(source),'source_sha256':doc['sha256'],
                  'archive_path':doc['archive_path'],'archive_sha256':doc['archive_sha256'],
                  'location':{k:row[k] for k in ['sheet','row','page','physical_line'] if k in row},
                  'excerpt':rendered[:1200], 'saved_values_only':metadata.get('cell_values_are_saved_values_not_formula_recalculation',False)})
    return {'status':'HISTORICAL_REFERENCE_ONLY','total':total,'returned':len(matches),'rows':matches,
      'issues':issues,'current_price_source':False,
      'scope':'2018년 사례 검색. 현재가격·현행 기준·설계 승인·원문 수식 재계산 검증을 뜻하지 않음'}
