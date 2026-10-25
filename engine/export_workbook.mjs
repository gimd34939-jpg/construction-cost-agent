import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import {Workbook,SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const input=path.resolve(process.argv[2]);
const result=JSON.parse(await fs.readFile(input,'utf8'));
const dest=path.join(path.dirname(input),'내역산출.xlsx');
try{await fs.access(dest);throw Error('기존 결과를 덮어쓸 수 없습니다');}catch(e){if(e.code!=='ENOENT')throw e;}
const wb=Workbook.create(), main=wb.worksheets.add('내역산출'), sources=wb.worksheets.add('단가근거');
const literal=v=>typeof v==='string'&&/^[=+@]/.test(v)?"'"+v:v;
const numeric=v=>{if(v===null||v===undefined||v==='')return null;let n=Number(v);if(!Number.isFinite(n)||Math.abs(n)>=1e15)throw Error('Excel의 유효 숫자 범위를 초과한 값입니다');return n;};
const columns=['순번','공종','규격','단위','수량','재료단가','노무단가','경비단가','합계단가','재료금액','노무금액','경비금액','합계금액','경로','상태','출처'];
main.getRange('A1:P1').merge();main.getRange('A1').values=[[literal(result.title)]];
main.getRange('A2:P2').merge();main.getRange('A2').values=[['노란색 수량을 변경하면 선택한 단가로 재계산합니다. 공종·가격·조건 변경과 실제 내역서 반영은 코덱스에 요청하세요.']];
main.getRange('A3:J3').merge();main.getRange('A3').values=[['선택 범위의 직접 구성금액 합계 · 제경비/보험/이윤/세금은 별도']];
main.getRange('K3:L3').merge();main.getRange('K3').values=[['선택 범위 합계']];
main.getRange('A5:P5').values=[columns];
const expected={};
for(let i=0;i<result.lines.length;i++){
 let l=result.lines[i],r=i+6,u=l.unit_components||{},a=l.amount_components||{},complete=l.amount!==null;
 let signed=l.signed_quantity_allowed===true&&['CONTRACT','USER_OVERRIDE'].includes(l.route);
 let quantityGuard=`IF(ISNUMBER(E${r}),AND(E${r}${signed?'<=0':'>=0'},ABS(E${r})<=1000000000000000,D${r}="${l.unit.replaceAll('"','""')}"),FALSE)`;
 const externalSource=l.resources.find(x=>x.price?.source_origin==='INDEPENDENT_EXTERNAL')?.price?.source;
 const sourceLabel=l.route==='EXTERNAL_MATERIAL'&&externalSource?`${externalSource.provider} / ${externalSource.source_location}`:(l.source.source_sheet?`${l.source.source_label} / ${l.source.source_sheet}!${l.source.source_row}`:`${l.source.source_document_id||'품셈'} / ${l.source.source_page||'자원 근거'} PDF`);
 let row=[i+1,literal(l.name),literal(l.specification),l.unit,numeric(l.quantity),numeric(u.material),numeric(u.labor),numeric(u.expense),numeric(l.unit_total),null,null,null,null,{CONTRACT:'계약 원문',MARKET:'시장단가',PUMSEM:'품셈 투입',USER_OVERRIDE:'명시 가격변경',EXTERNAL_MATERIAL:'독립 외부 자재'}[l.route],complete?'선택 범위 계산':'미확정 · 전체 금액 없음',literal(sourceLabel)];
 main.getRange(`A${r}:P${r}`).values=[row];
 if(l.unit_components){
  for(let [price,amount,key] of [['F','J','material'],['G','K','labor'],['H','L','expense']]){
   main.getRange(`${amount}${r}`).formulas=[[`=IF(AND(${quantityGuard},ISNUMBER(${price}${r})),TRUNC(E${r}*${price}${r},0),"")`]];
   expected[`${amount}${r}`]=numeric(a[key]);
  }
  if(complete){main.getRange(`I${r}`).formulas=[[`=SUM(F${r}:H${r})`]];main.getRange(`M${r}`).formulas=[[`=IF(COUNT(J${r}:L${r})=3,SUM(J${r}:L${r}),"")`]];expected[`I${r}`]=numeric(l.unit_total);expected[`M${r}`]=numeric(l.amount);}
 }else if(complete){main.getRange(`M${r}`).formulas=[[`=IF(AND(${quantityGuard},ISNUMBER(I${r})),TRUNC(E${r}*I${r},0),"")`]];expected[`M${r}`]=numeric(l.amount);}
 main.getRange(`O${r}`).formulas=[[`=IF(${quantityGuard},"${complete?'선택 범위 계산':'미확정 · 전체 금액 없음'}","수량 확인 필요")`]];expected[`O${r}`]=complete?'선택 범위 계산':'미확정 · 전체 금액 없음';
}
const last=result.lines.length+5;
main.getRange('M3').formulas=[[`=IF(COUNT(M6:M${last})=${result.lines.length},SUM(M6:M${last}),"미완성")`]];
expected.M3=result.selected_scope_total===null?'미완성':numeric(result.selected_scope_total);
const sourceRows=[];
const sourceText=(value)=>{
 if(!value)return '';
 const p=value.source_path||value.path, h=value.source_sha256||value.sha256;
 return [value.provider||value.source_document_id, p, h?`SHA256 ${h}`:null, value.source_location||value.location].filter(Boolean).join('\n');
};
for(let line of result.lines){
 const independent=line.resources.find(x=>x.price?.source_origin==='INDEPENDENT_EXTERNAL')?.price;
 const lineEvidence=line.route==='EXTERNAL_MATERIAL'&&independent?independent.source:line.source;
 const lineScope=line.route==='EXTERNAL_MATERIAL'?(line.original_input.api_mapping?.scope_statement||line.scope||''):line.scope||'';
 sourceRows.push([line.name,line.specification,line.route,lineScope,sourceText(lineEvidence),line.missing.length?line.missing.map(m=>m.reason).join(' / '):'필요 구성 입력 확보']);
 if(line.override_basis)sourceRows.push([line.name,'명시된 단가 변경 근거',line.override_basis.type,line.override_basis.scope,JSON.stringify(line.override_basis),'지시/근거는 계산결과와 요청 파일에 보존']);
 for(let r of line.resources)sourceRows.push([line.name,r.resource_name,r.price.price_id,`${r.resource_quantity} ${r.resource_unit} × ${r.price.price}원`,r.price===independent&&line.route==='EXTERNAL_MATERIAL'?'위 공급가격 원문과 같음':sourceText(r.price.source),r.price.basis]);
}
sourceRows.push(['사용 범위','선택한 공종과 가격에 한정','',result.total_scope,'요청 또는 입력 파일·계산결과.json은 같은 결과 폴더에 보존','새 계약총액은 연결 원문 수식과 범위를 확인']);
sourceRows.push(['입력 수정','수량만 Excel에서 수정','', '단가·포함범위·기준량·가격일 변경은 코덱스에 요청하고 수정본·계산근거를 함께 생성','Excel 수량 수정은 요청·계산결과 파일을 자동 갱신하지 않음','기존 파일은 덮어쓰지 않음']);
sources.getRange(`A1:F${sourceRows.length+1}`).values=[['공종','규격·자원','계산/가격 경로','계산 및 포함범위','출처 원문','검토 상태'],...sourceRows].map(row=>row.map(literal));
for(let [sh,end] of [[main,`P${last}`],[sources,`F${sourceRows.length+1}`]]){
 sh.showGridLines=false;sh.getRange(`A1:${end}`).format={font:{name:'맑은 고딕',size:10,color:'#18362F'},rowHeight:34,verticalAlignment:'center',wrapText:true,columnWidth:16};sh.freezePanes.freezeRows(sh===main?5:1);
}
main.getRange('A1:P1').format={fill:'#163F33',font:{name:'맑은 고딕',size:18,bold:true,color:'#FFFFFF'},rowHeight:40};
main.getRange('A2:P2').format={fill:'#EDF4EF',rowHeight:32};main.getRange('A3:P3').format={fill:'#EDF4EF',rowHeight:32,font:{name:'맑은 고딕',size:11,bold:true,color:'#18362F'}};
main.getRange('A5:P5').format={fill:'#146447',font:{name:'맑은 고딕',size:10,bold:true,color:'#FFFFFF'},rowHeight:36,wrapText:true,horizontalAlignment:'center'};
main.getRange(`A6:A${last}`).format.columnWidth=6;main.getRange(`B5:B${last}`).format.columnWidth=36;main.getRange(`C5:C${last}`).format.columnWidth=34;main.getRange(`D5:D${last}`).format.columnWidth=8;
main.getRange(`E6:E${last}`).format.fill='#FFF4D5';main.getRange(`E6:E${last}`).setNumberFormat('#,##0.####');main.getRange(`F6:I${last}`).setNumberFormat('#,##0.######');main.getRange(`J6:M${last}`).setNumberFormat('#,##0;[Red](#,##0);0');main.getRange('M3').setNumberFormat('#,##0"원"');
main.getRange(`N5:N${last}`).format.columnWidth=13;main.getRange(`O5:O${last}`).format.columnWidth=25;main.getRange(`P5:P${last}`).format.columnWidth=38;
main.getRange(`A6:P${last}`).format.borders={insideHorizontal:{style:'thin',color:'#E3E9E4'}};
main.getRange(`E6:M${last}`).format.horizontalAlignment='right';
sources.getRange('A1:F1').format={fill:'#163F33',font:{name:'맑은 고딕',size:10,bold:true,color:'#FFFFFF'},rowHeight:36};
sources.getRange(`A2:B${sourceRows.length+1}`).format.columnWidth=35;sources.getRange(`C2:C${sourceRows.length+1}`).format.columnWidth=40;sources.getRange(`D2:D${sourceRows.length+1}`).format.columnWidth=65;sources.getRange(`E2:E${sourceRows.length+1}`).format.columnWidth=65;sources.getRange(`F2:F${sourceRows.length+1}`).format.columnWidth=40;
sources.getRange(`E2:E${sourceRows.length+1}`).format.wrapText=true;
sources.getRange(`A2:F${sourceRows.length+1}`).format.autofitRows();
wb.recalculate();
const checks=[];
for(let [cell,wanted] of Object.entries(expected)){
 let got=main.getRange(cell).values[0][0];let pass=typeof wanted==='number'?typeof got==='number'&&Math.abs(got-wanted)<0.000001:got===wanted;
 checks.push({cell,expected:wanted,actual:got,status:pass?'PASS':'FAIL'});
}
if(checks.some(c=>c.status!=='PASS'))throw Error('저장 전 수식 대조 불일치 '+JSON.stringify(checks.filter(c=>c.status!=='PASS')));
const blob=await SpreadsheetFile.exportXlsx(wb);await blob.save(dest);
const reopened=await SpreadsheetFile.importXlsx(await FileBlob.load(dest));
for(let c of checks){let got=reopened.worksheets.getItem('내역산출').getRange(c.cell).values[0][0];if(typeof c.expected==='number'?Math.abs(got-c.expected)>0.000001:got!==c.expected)throw Error('저장 후 재가져오기 불일치 '+c.cell);}
const sha=crypto.createHash('sha256').update(await fs.readFile(dest)).digest('hex');
await fs.writeFile(path.join(path.dirname(input),'Excel검증.json'),JSON.stringify({status:'PASS',file_sha256:sha,formula_checks:checks.length,checks,source_result_sha256:crypto.createHash('sha256').update(await fs.readFile(input)).digest('hex')},null,2));
if(process.argv.includes('--preview')){
 const img=await wb.render({sheetName:'내역산출',range:`A1:P${Math.min(last,16)}`,scale:1.4,format:'png'});await fs.writeFile(path.join(path.dirname(input),'미리보기.png'),new Uint8Array(await img.arrayBuffer()));
}
console.log(JSON.stringify({status:'PASS',path:dest,checks:checks.length,sha256:sha}));
