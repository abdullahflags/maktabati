import {clean} from './core.mjs';
export const ACQUISITIONS=[['اقتناء - كتب','الكتب','الكتاب'],['اقتناء - شخصيات','الشخصيات','الشخصية'],['اقتناء - دور نشر','دور النشر','الدار'],['اقتناء - مواضيع','المواضيع','الموضوع']];
export function acquisitionGrid(sheet){
 if(!sheet?.rows)return {headers:[],items:[],followIndex:-1};
 const headerRow=sheet.rows.findIndex(row=>row.some(v=>clean(v)==='المتابعة'));
 if(headerRow<0)return {headers:[],items:[],followIndex:-1};
 const first=sheet.rows[headerRow].findIndex(v=>clean(v)==='المتابعة');let last=sheet.rows[headerRow].length;while(last>first&&!clean(sheet.rows[headerRow][last-1]))last--;
 const headers=sheet.rows[headerRow].slice(first,last).map(clean),definition=ACQUISITIONS.find(d=>d[0]===sheet.name),keyIndex=Math.max(1,headers.indexOf(definition?.[2]));
 const items=sheet.rows.slice(headerRow+1).map((row,i)=>({rowIndex:headerRow+1+i,values:headers.map((_,j)=>row[first+j]??null)})).filter(item=>clean(item.values[keyIndex]));
 return {headers,items,followIndex:first,headerRow,keyIndex};
}
export const acquisitionBookCount=aux=>acquisitionGrid((aux||[]).find(s=>s.name==='اقتناء - كتب')).items.length;

export function acquisitionChanges(existing,incoming){const changes=[];let added=0,updated=0,deleted=0;for(const[name,label]of ACQUISITIONS){const sheet=(incoming||[]).find(s=>s.name===name);if(!sheet)continue;const old=acquisitionGrid((existing||[]).find(s=>s.name===name)),next=acquisitionGrid(sheet),remaining=[...old.items];for(const item of next.items){const index=remaining.findIndex(r=>clean(r.values[old.keyIndex])===clean(item.values[next.keyIndex]));if(index<0){added++;changes.push({kind:'إضافة',title:clean(item.values[next.keyIndex]),details:'قائمة '+label});}else{const prev=remaining.splice(index,1)[0];if(next.headers.some((_,i)=>clean(item.values[i])!==clean(prev.values[i]))){updated++;changes.push({kind:'تحديث',title:clean(item.values[next.keyIndex]),details:'قائمة '+label});}}}for(const item of remaining){deleted++;changes.push({kind:'حذف',title:clean(item.values[old.keyIndex]),details:'قائمة '+label});}}return {changes,added,updated,deleted};}

export const FOLLOWUP_OPTIONS=[{label:'أضيفت ملاحظة',tone:'blue'},{label:'أُشتري',tone:'green'},{label:'غير متوفر',tone:'neutral'},{label:'لزيارة أخرى',tone:'purple'},{label:'مكلف',tone:'orange'},{label:'بُحث',tone:'teal'},{label:'زِيرَ',tone:'amber'}];

export function parseFollowups(value){return [...new Set(clean(value).split(/\s*(?:\||؛|\n)\s*/).map(clean).filter(Boolean))];}
export function toggleFollowup(value,label){const selected=new Set(parseFollowups(value));if(selected.has(label))selected.delete(label);else selected.add(label);return [...FOLLOWUP_OPTIONS.map(o=>o.label).filter(v=>selected.has(v)),...Array.from(selected).filter(v=>!FOLLOWUP_OPTIONS.some(o=>o.label===v))].join(' | ');}
