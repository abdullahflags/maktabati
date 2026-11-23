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
