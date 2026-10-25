import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const input=path.resolve(process.argv[2]),dir=path.dirname(input);
const result=JSON.parse(await fs.readFile(input,'utf8')),dest=path.join(dir,'가격비교.xlsx');
try{await fs.access(dest);throw Error('기존 결과 덮어쓰기 금지');}catch(e){if(e.code!=='ENOENT')throw e;}
if(!result.rows.length)throw Error('비교할 공종이 없습니다');
const wb=Workbook.create(),sh=wb.worksheets.add('가격비교'),ev=wb.worksheets.add('물량과가격근거');
const lit=v=>typeof v==='string'&&/^[=+@-]/.test(v)?"'"+v:v;
const num=v=>{if(v===null||v===undefined)return null;let x=Number(v);if(!Number.isFinite(x)||Math.abs(x)>=1e15)throw Error('Excel 숫자 범위 초과');return x;};
const checks=[],expected={};
sh.getRange('A1:Q1').merge();sh.getRange('A1').values=[[(result.purpose==='TEST_ONLY'?'검증용 가상 사례 · 실제 견적 아님 | ':'')+result.type_id+' | 동일 물량 가격 비교']];
sh.getRange('A2:Q2').merge();sh.getRange('A2').values=[[result.project.id+' / '+result.project.source_id+' / 가격 적용일 '+result.pricing_date+' · '+result.scope]];
sh.getRange('A3:H3').merge();sh.getRange('A3').values=[['계약/설계 참고안 합계 · 실제 사용자 참고 요청이 있을 때만 계산']];
sh.getRange('J3:M3').merge();sh.getRange('J3').values=[['신규 산정안 합계 / 차이']];
sh.getRange('A4:Q4').merge();sh.getRange('A4').values=[['독립 외부가격을 우선 적용합니다. 원문가격 참고 요청이 없으면 왼쪽 참고안은 계산하지 않습니다. 빈 신규금액은 미확정이며 0원이 아닙니다.']];
sh.getRange('A6:Q6').values=[['공종','규격','단위','물량','계약/설계 기준','기존 재료단가','기존 노무단가','기존 경비단가','기존 금액','신규 재료단가','신규 노무단가','신규 경비단가','신규 전체단가','신규 금액','금액 차이','신규 상태','공종 ID']];
const sourceRows=[];
for(let i=0;i<result.rows.length;i++){
 const line=result.rows[i],r=i+7,a=line.contract,b=line.current,ap=a.unit_components||{},bp=b.unit_components||{};
 sh.getRange('A'+r+':Q'+r).values=[[line.name,line.specification,line.unit,num(line.quantity),a.basis||a.status,num(ap.material),num(ap.labor),num(ap.expense),null,num(bp.material),num(bp.labor),num(bp.expense),num(b.unit_total),null,null,b.amount===null?'미확정':'계산 가능',line.line_id].map(lit)];
 const valid='AND(ISNUMBER(D'+r+'),D'+r+'>=0)';
 if(a.amount!==null){
  sh.getRange('I'+r).formulas=[['=IF(AND('+valid+',COUNT(F'+r+':H'+r+')=3),TRUNC(D'+r+'*F'+r+',0)+TRUNC(D'+r+'*G'+r+',0)+TRUNC(D'+r+'*H'+r+',0),"")']];
  expected['I'+r]=num(a.amount);
 }
 if(b.amount!==null){
  const formula=b.unit_components?'TRUNC(D'+r+'*J'+r+',0)+TRUNC(D'+r+'*K'+r+',0)+TRUNC(D'+r+'*L'+r+',0)':'TRUNC(D'+r+'*M'+r+',0)';
  const priceGuard=b.unit_components?'COUNT(J'+r+':L'+r+')=3':'ISNUMBER(M'+r+')';
  sh.getRange('N'+r).formulas=[['=IF(AND('+valid+','+priceGuard+'),'+formula+',"")']];expected['N'+r]=num(b.amount);
 }
 sh.getRange('O'+r).formulas=[['=IF(AND(ISNUMBER(I'+r+'),ISNUMBER(N'+r+')),N'+r+'-I'+r+',"")']];
 expected['O'+r]=line.difference===null?'':num(line.difference);
 const briefSource=s=>s?(s.path||s.source_path||s.source_document_id||'')+' / '+(s.location||((s.source_sheet||'')+'!'+(s.source_row||'')))+' / SHA '+(s.sha256||s.source_sha256||'별도 JSON 참조'):'가격비교.json 참조';
 const qe=line.quantity_evidence;
 sourceRows.push([line.line_id,'물량',qe.expression+' = '+line.quantity+' '+line.unit+' / '+(qe.method||'')+' / 원문 위치·단위환산·입력값은 가격비교.json의 quantity_evidence 참조',null,null,qe.status]);
 sourceRows.push([line.line_id,a.basis||'기존가격',briefSource(a.source),num(a.amount),num(a.unit_total),a.status]);
 sourceRows.push([line.line_id,b.basis||'신규가격',b.source?briefSource(b.source):JSON.stringify(b.missing||[]),num(b.amount),num(b.unit_total),b.status]);
}
const last=result.rows.length+6;
for(const [side,col] of [['contract','I'],['current','N']]){
 if(result.totals[side].amount!==null){
  sh.getRange(col+'3').formulas=[['=IF(COUNT('+col+'7:'+col+last+')='+result.rows.length+',SUM('+col+'7:'+col+last+'),"미완성")']];
  expected[col+'3']=num(result.totals[side].amount);
 }else{sh.getRange(col+'3').values=[['미완성']];expected[col+'3']='미완성';}
}
sh.getRange('O3').formulas=[['=IF(AND(ISNUMBER(I3),ISNUMBER(N3)),N3-I3,"")']];expected.O3=result.difference_total===null?'':num(result.difference_total);
sourceRows.push(['적용 범위','직접 구성금액',result.scope,null,null,result.status]);
sourceRows.push(['미완성 사유','확인 필요',JSON.stringify(result.missing),null,null,'누락을 0원으로 처리하지 않음']);
ev.getRange('A1:F'+(sourceRows.length+1)).values=[['공종 ID','구분','원문·근거 또는 누락항목','금액','단가','상태'],...sourceRows].map(row=>row.map(lit));
for(const [sheet,end] of [[sh,'Q'+last],[ev,'F'+(sourceRows.length+1)]]){
 sheet.showGridLines=false;sheet.getRange('A1:'+end).format={font:{name:'맑은 고딕',size:10,color:'#19342D'},rowHeight:34,columnWidth:16,verticalAlignment:'center',wrapText:true};
 sheet.freezePanes.freezeRows(sheet===sh?6:1);
}
sh.getRange('A1:Q1').format={fill:'#164D40',font:{name:'맑은 고딕',size:17,bold:true,color:'#FFFFFF'},rowHeight:42};
sh.getRange('A2:Q4').format={fill:'#EFF5F1',rowHeight:32};
sh.getRange('A6:Q6').format={fill:'#26745C',font:{name:'맑은 고딕',size:10,bold:true,color:'#FFFFFF'},rowHeight:42};
sh.getRange('A6:A'+last).format.columnWidth=33;sh.getRange('B6:B'+last).format.columnWidth=28;sh.getRange('C6:C'+last).format.columnWidth=9;sh.getRange('E6:E'+last).format.columnWidth=22;sh.getRange('Q6:Q'+last).format.columnWidth=28;
sh.getRange('D7:D'+last).setNumberFormat('#,##0.####');sh.getRange('F7:O'+last).setNumberFormat('#,##0.######;[Red](#,##0.######);0');
sh.getRange('I7:I'+last).setNumberFormat('#,##0');sh.getRange('N7:O'+last).setNumberFormat('#,##0;[Red](#,##0);0');
sh.getRange('I3:O3').setNumberFormat('#,##0"원"');
ev.getRange('A1:F1').format={fill:'#164D40',font:{name:'맑은 고딕',size:10,bold:true,color:'#FFFFFF'},rowHeight:34};
ev.getRange('A2:A'+(sourceRows.length+1)).format.columnWidth=28;ev.getRange('C2:C'+(sourceRows.length+1)).format.columnWidth=90;
ev.getRange('A2:F'+(sourceRows.length+1)).format.rowHeight=78;
ev.getRange('B2:B'+(sourceRows.length+1)).format.columnWidth=28;
ev.getRange('F2:F'+(sourceRows.length+1)).format.columnWidth=24;
ev.getRange('D2:E'+(sourceRows.length+1)).setNumberFormat('#,##0.######');
wb.recalculate();
for(const [cell,wanted] of Object.entries(expected)){
 const actual=sh.getRange(cell).values[0][0],pass=typeof wanted==='number'?typeof actual==='number'&&Math.abs(actual-wanted)<1e-6:actual===wanted;
 checks.push({sheet:sh.name,cell,expected:wanted,actual,pass});
}
if(checks.some(c=>!c.pass))throw Error('독립 계산과 수식 불일치 '+JSON.stringify(checks.filter(c=>!c.pass)));
const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(dest);
const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(dest));
for(const c of checks){const actual=reopened.worksheets.getItem(c.sheet).getRange(c.cell).values[0][0];if(typeof c.expected==='number'?typeof actual!=='number'||Math.abs(actual-c.expected)>1e-6:!(actual===c.expected||(c.expected===''&&actual===null)))throw Error('재열기 불일치 '+c.cell);}
await fs.writeFile(path.join(dir,'비교Excel검증.json'),JSON.stringify({status:'PASS',checks,source_result_sha256:crypto.createHash('sha256').update(await fs.readFile(input)).digest('hex'),workbook_sha256:crypto.createHash('sha256').update(await fs.readFile(dest)).digest('hex')},null,2));
for(const [sheet,range,filename] of [['가격비교','A1:Q'+Math.min(last,16),'비교미리보기.png'],['물량과가격근거','A1:F'+Math.min(sourceRows.length+1,10),'근거미리보기.png']]){
 const img=await wb.render({sheetName:sheet,range,scale:1.2,format:'png'});await fs.writeFile(path.join(dir,filename),new Uint8Array(await img.arrayBuffer()));
}
console.log(JSON.stringify({status:'PASS',workbook:dest,checks:checks.length}));
