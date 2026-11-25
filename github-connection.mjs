const KEY='maktabati.github.connection';
const LEGACY_KEY='maktabati.github.session';
export function rememberConnection(token,storage=localStorage){if(!token?.trim())throw Error('أدخل رمز ربط GitHub.');try{storage.setItem(KEY,token.trim());}catch{throw Error('تعذّر حفظ الربط على هذا الجهاز. راجع السماح بتخزين بيانات الموقع.');}}
export function loadConnection(storage=localStorage,session=sessionStorage){const remembered=storage.getItem(KEY);if(remembered)return remembered;const legacy=session.getItem(LEGACY_KEY);if(!legacy)return null;rememberConnection(legacy,storage);session.removeItem(LEGACY_KEY);return legacy;}
export function forgetConnection(storage=localStorage,session=sessionStorage){storage.removeItem(KEY);session.removeItem(LEGACY_KEY);}
