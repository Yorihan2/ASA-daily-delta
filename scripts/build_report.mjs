import fs from 'node:fs/promises';
import path from 'node:path';
import Module from 'node:module';
import { createRequire } from 'node:module';
if (process.env.NODE_PATH) Module.globalPaths.push(...process.env.NODE_PATH.split(';'));
const require = createRequire(import.meta.url);
const { Workbook, SpreadsheetFile } = require('@oai/artifact-tool');

const args=process.argv.slice(2);
const arg=name=>{const i=args.indexOf(`--${name}`);return i>=0?args[i+1]:undefined;};
const newPath=arg('current'),oldPath=arg('previous'),campaignPath=arg('campaigns'),outPath=arg('output');
if(!newPath||!outPath)throw new Error('用法：node build_report.mjs --current <最新日明细.csv> --output <输出.xlsx> [--previous <含前一日广告组字段的报告.csv>] [--campaigns <系列汇总.csv>]');
function parseLine(line){const a=[];let f='',q=false;for(let i=0;i<line.length;i++){const c=line[i];if(c==='"'){if(q&&line[i+1]==='"'){f+='"';i++;}else q=!q;}else if(c===','&&!q){a.push(f);f='';}else f+=c;}a.push(f);return a;}
async function load(p,headerStarts){const t=(await fs.readFile(p,'utf8')).replace(/^\uFEFF/,'');const ls=t.split(/\r?\n/).filter(x=>x.trim());const hi=ls.findIndex(x=>x.includes(headerStarts));if(hi<0)throw Error(`找不到表头：${p}`);const headers=parseLine(ls[hi]);return{lines:ls,headers,rows:ls.slice(hi+1).map(l=>Object.fromEntries(headers.map((h,i)=>[h,parseLine(l)[i]??''])))};}
const n=v=>Number(String(v??'0').replace(/[$,%\s,]/g,''))||0;
const newData=await load(newPath,'日期,支出,'), oldData=oldPath?await load(oldPath,'日期,支出,'):null, campaignData=campaignPath?await load(campaignPath,'广告系列 ID'):null;
const parseDate=s=>{const m=String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);return m?new Date(Number(m[3]),Number(m[1])-1,Number(m[2])):new Date(s);};
const dateList=[...new Set(newData.rows.map(r=>r['日期']).filter(Boolean))].sort((a,b)=>parseDate(a)-parseDate(b));
if(dateList.length<2)throw new Error('最新报告至少需要包含连续两天，才能比较昨日与前日。');
const previousDate=dateList.at(-2),currentDate=dateList.at(-1);
const dateLabel=s=>{const d=parseDate(s);return `${d.getMonth()+1}/${d.getDate()}`;};
const old14=oldData?.rows.filter(r=>r['日期']===previousDate&&r['广告组 ID']&&r['广告组名称'])??[];
const mk=(source,keyword,term)=>JSON.stringify([source,keyword,term]);
const oldGroupMap=new Map();
for(const r of old14){const k=mk(r['搜索词匹配来源'],r['关键词'],r['搜索词']);if(!oldGroupMap.has(k))oldGroupMap.set(k,new Map());const id=r['广告组 ID'],name=r['广告组名称'];if(!oldGroupMap.get(k).has(id))oldGroupMap.get(k).set(id,{id,name,spend:0});oldGroupMap.get(k).get(id).spend+=n(r['支出']);}
const dailyMap=new Map();
for(const r of newData.rows.filter(r=>[previousDate,currentDate].includes(r['日期']))){const key=mk(r['搜索词匹配来源'],r['关键词'],r['搜索词']);if(!dailyMap.has(key))dailyMap.set(key,{source:r['搜索词匹配来源'],keyword:r['关键词'],keywordId:r['关键词 ID']??'',term:r['搜索词'],byDate:{}});const g=dailyMap.get(key);g.byDate[r['日期']]??={spend:0,clicks:0,installs:0};g.byDate[r['日期']].spend+=n(r['支出']);g.byDate[r['日期']].clicks+=n(r['点击次数']);g.byDate[r['日期']].installs+=n(r['安装次数（总计）']??r['安装次数 (总计)']);}
const allDeltas=[...dailyMap.entries()].map(([key,g])=>{const a=g.byDate[previousDate]??{spend:0,clicks:0,installs:0},b=g.byDate[currentDate]??{spend:0,clicks:0,installs:0};const matches=[...(oldGroupMap.get(key)?.values()??[])].sort((x,y)=>y.spend-x.spend);return{...g,a,b,delta:b.spend-a.spend,matches};});
const details=allDeltas.filter(x=>x.delta>0.005).sort((x,y)=>y.delta-x.delta);
const sumMap=new Map();
for(const d of allDeltas){const key=JSON.stringify([d.source,d.keywordId,d.keyword]);if(!sumMap.has(key))sumMap.set(key,{source:d.source,keyword:d.keyword,keywordId:d.keywordId,a:{spend:0,clicks:0,installs:0},b:{spend:0,clicks:0,installs:0},queryDeltas:[],groups:new Map()});const z=sumMap.get(key);for(const m of ['spend','clicks','installs']){z.a[m]+=d.a[m];z.b[m]+=d.b[m];}z.queryDeltas.push({term:d.term,delta:d.delta});for(const g of d.matches){if(!z.groups.has(g.id))z.groups.set(g.id,g.name);}}
const summaries=[...sumMap.values()].map(x=>({...x,delta:x.b.spend-x.a.spend,groups:[...x.groups.values()],queryDeltas:x.queryDeltas.sort((a,b)=>b.delta-a.delta)})).filter(x=>x.delta>0.005).sort((a,b)=>b.delta-a.delta);
const campaignRows=campaignData?.rows??[];
for(const s of summaries){const matched=s.groups.length===1?s.groups[0]:null;s.campaign=matched?campaignRows.find(c=>c['广告系列名称']===matched):null;}
const total14=newData.rows.filter(r=>r['日期']===previousDate).reduce((s,r)=>s+n(r['支出']),0);const total15=newData.rows.filter(r=>r['日期']===currentDate).reduce((s,r)=>s+n(r['支出']),0);const positiveTotal=details.reduce((s,x)=>s+x.delta,0);
const wb=Workbook.create();const summary=wb.worksheets.add('关键词汇总');const detail=wb.worksheets.add('搜索词增量');
const font='Arial',black='#000000',white='#FFFFFF',red='#C00000';
function setup(sheet,title,note,endCol){sheet.showGridLines=false;sheet.getRange(`A1:${endCol}1`).merge();sheet.getRange('A1').values=[[title]];sheet.getRange(`A1:${endCol}1`).format={font:{name:font,size:16,bold:true,color:black},fill:white,verticalAlignment:'center'};sheet.getRange(`A1:${endCol}1`).format.rowHeight=30;sheet.getRange(`A2:${endCol}2`).merge();sheet.getRange('A2').values=[[note]];sheet.getRange(`A2:${endCol}2`).format={font:{name:font,size:10,color:black},fill:white,wrapText:true,verticalAlignment:'center'};sheet.getRange(`A2:${endCol}2`).format.rowHeight=34;}
setup(summary,`Apple Ads | ${dateLabel(currentDate)} 花费增加关键词`,`对比 ${dateLabel(previousDate)} 与 ${dateLabel(currentDate)}。总花费：$${total14.toFixed(2)} → $${total15.toFixed(2)}（净增 $${(total15-total14).toFixed(2)}）；本表列出花费增加的词。广告组仅参考前一日报告的同词/搜索词匹配；最新报告不含广告组/系列字段时，不能视作最新日期的确认归属。`, 'M');
const sumHeaders=['匹配来源','关键词','关键词 ID',`${dateLabel(previousDate)} 花费`,`${dateLabel(currentDate)} 花费`,'增加额',`${dateLabel(previousDate)} 点击`,`${dateLabel(currentDate)} 点击`,`${dateLabel(previousDate)} 总安装`,`${dateLabel(currentDate)} 总安装`,'增加主要搜索词','前日报告广告组参考','Campaign ID'];
summary.getRange('A4:M4').values=[sumHeaders];
const sumRows=summaries.map(s=>[s.source,s.keyword==='SEARCH_MATCH'?'自动搜索匹配':s.keyword,s.keywordId,s.a.spend,s.b.spend,s.delta,s.a.clicks,s.b.clicks,s.a.installs,s.b.installs,s.queryDeltas.filter(q=>q.delta>0).slice(0,5).map(q=>`${q.term} (+$${q.delta.toFixed(2)})`).join('；'),s.groups.length===0?'未匹配':s.groups.length===1?s.groups[0]:`多组匹配（${s.groups.length}组）：${s.groups.slice(0,5).join('；')}${s.groups.length>5?'…':''}`,s.campaign?.['广告系列 ID']??'未确认']);
const sumLast=4+sumRows.length;summary.getRange(`A5:M${sumLast}`).values=sumRows;const sumTable=summary.tables.add(`A4:M${sumLast}`,true,'KeywordSpendIncreaseSummary');sumTable.style='TableStyleLight1';
setup(detail,'Apple Ads | 搜索词级花费增量',`按匹配来源、关键词和搜索词拆分。红色关键词表示该组合 ${dateLabel(currentDate)} 花费高于 ${dateLabel(previousDate)}；“低量隐藏词”是 Apple 汇总桶，实际查询不可见。组别参考来自前一日报告，多组或未匹配表示无法确认最新日期归属。`, 'M');
const detHeaders=['匹配来源','关键词','关键词 ID','搜索词 / 隐藏词',`${dateLabel(previousDate)} 花费`,`${dateLabel(currentDate)} 花费`,'增加额',`${dateLabel(previousDate)} 点击`,`${dateLabel(currentDate)} 点击`,`${dateLabel(previousDate)} 总安装`,`${dateLabel(currentDate)} 总安装`,'前日报告广告组参考','广告组 ID 参考'];
detail.getRange('A4:M4').values=[detHeaders];
const detailRows=details.map(d=>[d.source,d.keyword==='SEARCH_MATCH'?'自动搜索匹配':d.keyword,d.keywordId,d.term==='(Low volume terms)'?'低量隐藏词（未披露）':d.term,d.a.spend,d.b.spend,d.delta,d.a.clicks,d.b.clicks,d.a.installs,d.b.installs,d.matches.length===0?'未匹配':d.matches.length===1?d.matches[0].name:`多组匹配（${d.matches.length}组）`,d.matches.length===1?d.matches[0].id:'']);
const detLast=4+detailRows.length;detail.getRange(`A5:M${detLast}`).values=detailRows;const detailTable=detail.tables.add(`A4:M${detLast}`,true,'SearchTermSpendIncreaseDetails');detailTable.style='TableStyleLight1';
for(const [sheet,end,last,keywordCol,widths] of [[summary,'M',sumLast,'B',[18,35,16,16,16,16,14,14,15,15,55,60,20]],[detail,'M',detLast,'B',[18,35,16,34,16,16,16,14,14,15,15,42,18]]]){
  sheet.getRange(`A4:${end}4`).format={fill:white,font:{name:font,size:10,bold:true,color:black},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,borders:{preset:'outside',style:'thin',color:black}};
  sheet.getRange(`A5:${end}${last}`).format={fill:white,font:{name:font,size:9,color:black},verticalAlignment:'center',wrapText:true,borders:{preset:'outside',style:'thin',color:'#808080'}};
  sheet.getRange(`${keywordCol}5:${keywordCol}${last}`).format.font={name:font,size:9,color:red,bold:true};
  for(let i=0;i<widths.length;i++){const col=String.fromCharCode(65+i);sheet.getRange(`${col}:${col}`).format.columnWidth=widths[i];}
  if(sheet===summary){sheet.getRange(`D5:F${last}`).format.numberFormat='$#,##0.00';sheet.getRange(`G5:J${last}`).format.numberFormat='#,##0';}
  else {sheet.getRange(`E5:G${last}`).format.numberFormat='$#,##0.00';sheet.getRange(`H5:K${last}`).format.numberFormat='#,##0';}
  sheet.getRange(`A4:${end}${last}`).format.rowHeight=24;
  sheet.freezePanes.freezeRows(4);
}
summary.getRange(`K5:M${sumLast}`).format.horizontalAlignment='left';
detail.getRange(`L5:M${detLast}`).format.horizontalAlignment='left';
wb.recalculate();
await fs.mkdir(path.dirname(outPath),{recursive:true});
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:100},summary:'increment workbook error scan'});console.log(`ROWS ${summaries.length} keywords; ${details.length} search-term combinations; gross positive increases $${positiveTotal.toFixed(2)}; net +$${(total15-total14).toFixed(2)}`);console.log(errors.ndjson);
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(outPath);console.log(`SAVED ${outPath}`);

