"""Source-preserving explicit numeric edits for XLS/XLSX calculation documents."""
from pathlib import Path
from decimal import Decimal
import datetime,hashlib,io,json,subprocess,uuid,zipfile
from lxml import etree as ET
import source_workbook_edit as edit

CONVERSION_SCRIPT=r'''param([Parameter(Mandatory=$true)][string]$Source,[Parameter(Mandatory=$true)][string]$Target,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop'
$xl=$null;$wb=$null;$blank=$null
$r=[ordered]@{status='RUNNING';source=$Source;target=$Target;calculation='manual';update_links=0;automation_security=3;read_only_original=$false;application_quit=$false;error=$null}
try{
 $xl=New-Object -ComObject Excel.Application
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.ScreenUpdating=$false;$xl.AskToUpdateLinks=$false;$xl.AutomationSecurity=3
 $r.excel_version=$xl.Version;$r.excel_build=$xl.Build
 $blank=$xl.Workbooks.Add();$xl.Calculation=-4135;$xl.CalculateBeforeSave=$false
 $wb=$xl.Workbooks.Open($Source,0,$true)
 $r.read_only_original=$wb.ReadOnly
 if(-not$wb.ReadOnly){throw 'ORIGINAL_MUST_BE_READONLY'}
 if($wb.HasVBProject){throw 'MACRO_WORKBOOK_CONVERSION_NOT_SUPPORTED'}
 $r.original_sheets=@();$r.original_shape_counts=[ordered]@{}
 foreach($s in $wb.Worksheets){$r.original_sheets+=$s.Name;$r.original_shape_counts[$s.Name]=$s.Shapes.Count;[Runtime.InteropServices.Marshal]::FinalReleaseComObject($s)|Out-Null}
 $wb.CheckCompatibility=$false
 $wb.SaveAs($Target,51)
 $r.converted_sheets=@();$r.converted_shape_counts=[ordered]@{}
 foreach($s in $wb.Worksheets){$r.converted_sheets+=$s.Name;$r.converted_shape_counts[$s.Name]=$s.Shapes.Count;[Runtime.InteropServices.Marshal]::FinalReleaseComObject($s)|Out-Null}
 $r.sheet_names_preserved=(($r.original_sheets|ConvertTo-Json -Compress)-ceq($r.converted_sheets|ConvertTo-Json -Compress))
 $r.shape_counts_preserved=(($r.original_shape_counts|ConvertTo-Json -Compress)-ceq($r.converted_shape_counts|ConvertTo-Json -Compress))
 if(-not$r.sheet_names_preserved -or -not$r.shape_counts_preserved){throw 'CONVERSION_SHEET_OR_SHAPE_MISMATCH'}
 $wb.Close($false);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)|Out-Null;$wb=$null
 $blank.Close($false);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($blank)|Out-Null;$blank=$null
 $r.status='CONVERTED_XLS_COPY'
}catch{$r.status='FAIL';$r.error=[ordered]@{message=$_.Exception.Message;stack=$_.ScriptStackTrace};throw}
finally{
 if($null-ne$wb){try{$wb.Close($false)}catch{};[Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)|Out-Null}
 if($null-ne$blank){try{$blank.Close($false)}catch{};[Runtime.InteropServices.Marshal]::FinalReleaseComObject($blank)|Out-Null}
 if($null-ne$xl){try{$xl.Quit();$r.application_quit=$true}catch{};[Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)|Out-Null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 $r|ConvertTo-Json -Depth 30|Set-Content -LiteralPath $Report -Encoding UTF8
}'''

def apply_numeric_edit(source_path,expected_sha256,changes,output_path):
    source=Path(source_path).resolve();target=Path(output_path).resolve()
    if source==target or target.exists():raise edit.EditError('원본/기존 결과는 덮어쓸 수 없습니다')
    if not source.is_file()or edit._sha(source)!=expected_sha256:raise edit.EditError('원본 SHA256이 입력 근거와 다릅니다')
    if source.suffix.lower()not in['.xls','.xlsx']or target.suffix.lower()!='.xlsx':raise edit.EditError('XLS/XLSX 원문에서 새 XLSX 사본만 만들 수 있습니다')
    if not isinstance(changes,list)or not changes:raise edit.EditError('명시한 숫자 입력 변경이 필요합니다')
    # Validate request syntax before conversion or any output creation.
    for c in changes:
        if not isinstance(c,dict)or any(c.get(k)is None for k in['sheet','cell','value','expected_value','role','evidence']):raise edit.EditError('숫자 변경의 필수 원문/값/역할/근거가 누락됐습니다')
        if c['role']not in['quantity','resource_price','coefficient']:raise edit.EditError('지원되지 않는 숫자 역할')
        if not isinstance(c['cell'],str)or not edit.re.fullmatch(r'[A-Z]{1,3}[1-9]\d*',c['cell']):raise edit.EditError('단일 A1 셀 주소가 필요합니다')
        new=edit._number(c['value']);old=edit._number(c['expected_value'])
        if new<0:raise edit.EditError('숫자 입력 변경은 0 이상만 지원합니다')
        if not isinstance(c['evidence'],dict)or not c['evidence']:raise edit.EditError('원문 숫자 변경 근거가 필요합니다')
    if len({(c['sheet'],c['cell'])for c in changes})!=len(changes):raise edit.EditError('같은 셀 변경이 중복됐습니다')
    conversion=None;intermediate=source
    if source.suffix.lower()=='.xls':
        folder=target.parent/('형식변환근거_'+uuid.uuid4().hex[:10]);folder.mkdir(parents=True,exist_ok=False)
        script=folder/'convert_readonly.ps1';script.write_text(CONVERSION_SCRIPT,encoding='utf-8-sig')
        intermediate=folder/'원문_형식변환사본.xlsx';report=folder/'형식변환검증.json'
        ps=Path('C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe')
        run=subprocess.run([str(ps),'-NoProfile','-ExecutionPolicy','Bypass','-File',str(script),'-Source',str(source),'-Target',str(intermediate),'-Report',str(report)],capture_output=True,text=True,encoding='utf-8',errors='replace',creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),timeout=180)
        if run.returncode or not intermediate.is_file():
            if edit._sha(source)!=expected_sha256:raise edit.EditError('원본 보존 실패')
            raise edit.EditError('읽기 전용 XLS 형식 변환 실패; 변환 진단 '+str(report))
        conversion=json.loads(report.read_text(encoding='utf-8-sig'))
        conversion.update(original_sha256=expected_sha256,converted_sha256=edit._sha(intermediate),conversion_report=str(report))
        if conversion['status']!='CONVERTED_XLS_COPY' or not conversion['read_only_original']:raise edit.EditError('형식 변환 보존 검사 실패')
    book=edit.Book(intermediate);changed=set();operations=[]
    for change in changes:
        sn=change['sheet'];addr=change['cell']
        if sn not in book.cells or addr not in book.cells[sn]:raise edit.EditError('원문 숫자 셀을 찾을 수 없습니다')
        cell=book.cells[sn][addr]
        if addr in book.formulas[sn]:raise edit.EditError('수식 셀을 숫자로 덮어쓰지 않습니다: '+sn+'!'+addr)
        if cell.get('t')in['e','str','s','inlineStr','b']:raise edit.EditError('숫자 상수 셀만 편집합니다: '+sn+'!'+addr)
        old=book.value(cell)
        if not isinstance(old,Decimal)or old!=edit._number(change['expected_value']):raise edit.EditError('원문 현재 값이 예상 값과 다릅니다: '+sn+'!'+addr)
        for f in book.trees[sn].iter(edit._q('f')):
            if f.get('t')=='array'and f.get('ref'):
                l,t,r,b=edit.range_boundaries(f.get('ref'));col,row=edit.coordinate_from_string(addr);cc=edit.column_index_from_string(col)
                if l<=cc<=r and t<=row<=b:raise edit.EditError('배열 수식 입력 영역을 편집하지 않습니다')
        new=edit._number(change['value']);book.set_cell(sn,addr,new);changed.add((sn,addr))
        operations.append({'sheet':sn,'cell':addr,'before':str(old),'after':str(new),'role':change['role'],'evidence':change['evidence']})
    cached=edit._recalculate(book,changed)
    removed=set()
    if'xl/calcChain.xml'in book.parts:
        removed.add('xl/calcChain.xml')
        rels=edit._xml(book.parts['xl/_rels/workbook.xml.rels'])
        for rel in list(rels):
            if rel.get('Type','').endswith('/calcChain'):rels.remove(rel)
        book.parts['xl/_rels/workbook.xml.rels']=edit._serialize(rels);book.changed.add('xl/_rels/workbook.xml.rels')
        content=edit._xml(book.parts['[Content_Types].xml'])
        for node in list(content):
            if node.get('PartName')=='/xl/calcChain.xml':content.remove(node)
        book.parts['[Content_Types].xml']=edit._serialize(content);book.changed.add('[Content_Types].xml')
    calc=book.workbook.find(edit._q('calcPr'))
    if calc is None:calc=ET.SubElement(book.workbook,edit._q('calcPr'))
    calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1');calc.set('calcMode','auto');book.changed.add('xl/workbook.xml')
    for sn,p in book.paths.items():
        if p in book.changed:
            for addr,f in book.formulas[sn].copy().items():book.normalize_formula(sn,addr,f)
            book.parts[p]=edit._serialize(book.trees[sn])
    book.parts['xl/workbook.xml']=edit._serialize(book.workbook)
    stream=io.BytesIO()
    with zipfile.ZipFile(stream,'w')as z:
        for info in book.infos:
            if info.filename not in removed:z.writestr(info,book.parts[info.filename])
    if edit._sha(source)!=expected_sha256:raise edit.EditError('변경 처리 중 원본 SHA가 달라졌습니다')
    target.parent.mkdir(parents=True,exist_ok=True)
    with target.open('xb')as f:f.write(stream.getvalue())
    if edit._sha(source)!=expected_sha256:raise edit.EditError('원본 보존 검증 실패')
    return {'status':'EDITED_NUMERIC_SOURCE_COPY','source_path':str(source),'source_sha256_before':expected_sha256,'source_sha256_after':edit._sha(source),'source_unchanged':True,'output_path':str(target),'output_sha256':edit._sha(target),'conversion':conversion,'conversion_source_path':str(intermediate),'conversion_source_sha256':edit._sha(intermediate),'changes':operations,'recalculated_local_cells':len(cached),'recalculated_cells':cached,'changed_ooxml_parts':sorted(book.changed),'removed_parts':sorted(removed),'preserved_parts':sorted(set(book.parts)-book.changed-removed),'source_scope_warnings':['명시된 숫자 상수와 변경 의존 수식만 수정; 수식/외부링크 입력 덮어쓰기는 차단','기존 원문 참조 오류·저장값 불일치와 적용 시기·단위·현장 조건을 자동 승인/교정하지 않음','원본 자료의 기존 검증: 착공/변경본 참조 오류 28건 및 UNIT 저장값 불일치 2건은 별도 보고서 참조'],'created_utc':datetime.datetime.now(datetime.timezone.utc).isoformat()}
