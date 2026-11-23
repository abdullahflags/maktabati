import {FIELDS,clean,normalize} from './core.mjs';
import {ACQUISITIONS,acquisitionGrid} from './acquisitions.mjs';
const NSHEETS=['مكتبتي','الفهرس','اقتناء - كتب','اقتناء - شخصيات','اقتناء - دور نشر','اقتناء - مواضيع','حساب البحث'];
const columns='BCDEFGHIJKLM'.split('');
const esc=v=>String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');
const unesc=v=>v.replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
const attributes=tag=>Object.fromEntries([...tag.matchAll(/([\w:.-]+)\s*=\s*(["'])(.*?)\2/g)].map(m=>[m[1],unesc(m[3])]));
function catalogHeader(wb){for(let row=1;row<=40;row++)if(FIELDS.every(([,label],i)=>clean(wb.Sheets['الفهرس']?.[columns[i]+row]?.v)===label))return row;return 0;}
export function isOriginalLayout(wb){return NSHEETS.every(n=>wb.SheetNames.includes(n))&&catalogHeader(wb)>0;}
async function sheetPath(zip,name){const workbook=await zip.file('xl/workbook.xml').async('string'),rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');const sheet=[...workbook.matchAll(/<sheet\b[^>]*\/?\s*>/g)].map(m=>attributes(m[0])).find(s=>s.name===name);if(!sheet)throw Error('صفحة القالب غير موجودة: '+name);const rel=[...rels.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map(m=>attributes(m[0])).find(r=>r.Id===sheet['r:id']);if(!rel)throw Error('رابط صفحة القالب غير موجود.');return resolvePart('xl/workbook.xml',rel.Target);}
function resolvePart(from,target){const parts=(target.startsWith('/')?target.slice(1):from.slice(0,from.lastIndexOf('/')+1)+target).split('/'),out=[];for(const p of parts){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);}return out.join('/');}
function cellValue(address,value,existing='',defaultStyle='14'){
 let attrs=existing?attributes(existing.match(/^<c\b[^>]*>/)?.[0]||''):{};
 attrs={...attrs,r:address};delete attrs.t;delete attrs.cm;delete attrs.vm;
 if(!('s'in attrs))attrs.s=defaultStyle;
 const opening=()=>'<c '+Object.entries(attrs).map(([k,v])=>k+'="'+esc(v)+'"').join(' ');
 if(value===null||value===undefined||value==='')return opening()+' />';
 if(typeof value==='number'&&Number.isFinite(value))return opening()+' t="n"><v>'+value+'</v></c>';
 if(address.startsWith('B')&&/^\d+$/.test(String(value))&&Number.isSafeInteger(Number(value))&&String(Number(value))===String(value))return opening()+' t="n"><v>'+value+'</v></c>';
 return opening()+' t="inlineStr"><is><t xml:space="preserve">'+esc(value)+'</t></is></c>';
}
function patchRow(xml,row,updates){return xml.replace(new RegExp('<row\\b[^>]*\\br="'+row+'"[^>]*>[\\s\\S]*?<\\/row>'),match=>{let result=match;for(const [col,value]of updates){const address=col+row,re=new RegExp('<c\\b[^>]*\\br="'+address+'"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)');const old=result.match(re)?.[0];if(old)result=result.replace(re,cellValue(address,value,old));else result=result.replace('</row>',cellValue(address,value)+'</row>');}return result;});}
function newRow(row,record,blank=false){const style=row%2?'14':'16',idStyle=row%2?'13':'15';return '<row r="'+row+'" ht="'+(blank?'35':'30')+'" customHeight="1">'+FIELDS.map(([key],i)=>cellValue(columns[i]+row,record?.[key]??null,'',blank?'39':i===0?idStyle:style)).join('')+'</row>';}
async function updateAcquisitions(zip,wb,auxiliary,XLSX){let changed=false;for(const[name]of ACQUISITIONS){const desired=acquisitionGrid(auxiliary.find(s=>s.name===name));if(!desired.items.length||!wb.Sheets[name])continue;const sheet=wb.Sheets[name],range=XLSX.utils.decode_range(sheet['!ref']);range.s={r:0,c:0};const baseline=acquisitionGrid({name,rows:XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,range})});if(!baseline.headers.length||baseline.headerRow!==desired.headerRow||baseline.headers.join('|')!==desired.headers.join('|'))throw Error('قالب قائمة الاقتناء غير متوافق؛ أعد استيراد الملف الكامل قبل التصدير.');const path=await sheetPath(zip,name);let xml=await zip.file(path).async('string'),original=xml;for(const item of desired.items){const match=baseline.items.find(r=>r.rowIndex===item.rowIndex);if(!match)throw Error('سجل الاقتناء غير موجود في القالب. أعد استيراد الملف الكامل.');const updates=[];desired.headers.forEach((_,i)=>{const value=item.values[i]??null;if(clean(value)!==clean(match.values[i]))updates.push([XLSX.utils.encode_col(baseline.followIndex+i),value]);});if(updates.length)xml=patchRow(xml,match.rowIndex+1,updates);}if(xml!==original){zip.file(path,xml,{createFolders:false});changed=true;}}return changed;}

export async function exportOriginalWorkbook(bytes,records,{XLSX,JSZip,auxiliary=[]}){
 const wb=XLSX.read(bytes,{type:'array',cellFormula:true});if(!isOriginalLayout(wb))throw Error('قالب الإكسيل غير مطابق لملف مكتبتك الأصلي؛ لم يُنشأ ملف بديل.');
 const seen=new Set();for(const r of records){const id=normalize(r.id);if(!id||seen.has(id)||!clean(r.title))throw Error('توجد أرقام مفقودة أو مكررة أو عناوين فارغة. راجع الفهرس قبل التصدير.');seen.add(id);}
 const zip=await JSZip.loadAsync(bytes),path=await sheetPath(zip,'الفهرس'),sheet=wb.Sheets['الفهرس'];let xml=await zip.file(path).async('string');const sourceXml=xml;
 const header=catalogHeader(wb),existing=new Map(),maxRow=XLSX.utils.decode_range(sheet['!ref']).e.r+1;let lastDataRow=header;
 for(let row=header+1;row<=maxRow;row++){if(clean(sheet['E'+row]?.v)&&clean(sheet['B'+row]?.v)){const id=normalize(sheet['B'+row].v);if(existing.has(id))throw Error('قالب الإكسيل يتضمن رقمًا مكررًا.');existing.set(id,row);lastDataRow=row;}}
 const wanted=new Map(records.map(r=>[normalize(r.id),r]));
 for(const[id,row]of existing){const record=wanted.get(id),updates=[];FIELDS.forEach(([key],i)=>{const current=sheet[columns[i]+row]?.v??null,value=record?.[key]??null;if(clean(current)!==clean(value))updates.push([columns[i],value]);});if(updates.length)xml=patchRow(xml,row,updates);}
 const added=records.filter(r=>!existing.has(normalize(r.id)));let last=lastDataRow;
 for(const r of added){last++;const old=new RegExp('<row\\b[^>]*\\br="'+last+'"[^>]*>[\\s\\S]*?<\\/row>');if(old.test(xml))xml=xml.replace(old,newRow(last,r));else xml=xml.replace('</sheetData>',newRow(last,r)+'</sheetData>');}
 if(added.length){const blank=last+1;if(!new RegExp('<row\\b[^>]*\\br="'+blank+'"').test(xml))xml=xml.replace('</sheetData>',newRow(blank,null,true)+'</sheetData>');
  const relPath=path.replace(/\/([^/]+)$/,'/_rels/$1.rels'),rels=await zip.file(relPath)?.async('string');if(!rels)throw Error('جدول الفهرس غير موجود في القالب.');const tableRel=[...rels.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map(m=>attributes(m[0])).find(r=>r.Type?.endsWith('/table'));if(!tableRel)throw Error('جدول الفهرس غير موجود في القالب.');const tablePath=resolvePart(path,tableRel.Target),table=await zip.file(tablePath).async('string');const oldRef=attributes(table.match(/<table\b[^>]*>/)[0]).ref;const oldEnd=Number(oldRef.match(/M(\d+)$/)?.[1]);if(!oldEnd)throw Error('نطاق جدول الفهرس غير متوقع.');const end=Math.max(oldEnd,blank),newRef='B'+header+':M'+end;
  if(end>oldEnd){zip.file(tablePath,table.replaceAll('ref="'+oldRef+'"','ref="'+newRef+'"'),{createFolders:false});const helperPath=await sheetPath(zip,'حساب البحث');let helper=await zip.file(helperPath).async('string');for(const col of ['A','B','C'])helper=helper.replace('ref="'+col+'2:'+col+(oldEnd-header+1)+'"','ref="'+col+'2:'+col+(end-header+1)+'"');zip.file(helperPath,helper,{createFolders:false});xml=xml.replace(/(<dimension\b[^>]*ref="[^"]+:)[A-Z]+\d+("[^>]*\/>)/,'$1N'+end+'$2');}
 }
 const auxiliaryChanged=await updateAcquisitions(zip,wb,auxiliary,XLSX);
 if(xml===sourceXml&&added.length===0&&!auxiliaryChanged)return {bytes:new Uint8Array(bytes),changed:false};
 if(records.length!==existing.size){const homePath=await sheetPath(zip,'مكتبتي');let home=await zip.file(homePath).async('string');for(const col of 'BCDEFGHI')home=home.replace(new RegExp('ref="'+col+'(\\d+):'+col+'\\d+"'),(_,start)=>'ref="'+col+start+':'+col+Math.max(Number(start),records.length+Number(start)-1)+'"');zip.file(homePath,home,{createFolders:false});}
 zip.file(path,xml,{createFolders:false});
 // Excel recalculates the original structured-reference formulas on opening.
 if(xml!==sourceXml||added.length){const workbook=await zip.file('xl/workbook.xml').async('string');
 const recalculated=workbook.replace(/<calcPr\b([^>]*?)(\/?>)/,(_,a,end)=>{let attrs=attributes('<calcPr '+a+'>');attrs={...attrs,calcMode:'auto',fullCalcOnLoad:'1',forceFullCalc:'1'};return '<calcPr '+Object.entries(attrs).map(([k,v])=>k+'="'+esc(v)+'"').join(' ')+' />';});
 if(recalculated!==workbook)zip.file('xl/workbook.xml',recalculated,{createFolders:false});}
 return {bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),changed:true};
}
