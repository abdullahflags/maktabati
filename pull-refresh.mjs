// Only the home screen accepts this gesture; forms and scrolling lists keep native touch behavior.
export function installPullRefresh({doc=document,win=window,onRefresh=()=>win.location.reload()}={}){
 const indicator=doc.getElementById('pull-refresh');
 let start=null,distance=0,refreshing=false;
 const blocked=()=>Boolean(doc.querySelector('dialog[open]'))||Boolean(doc.activeElement?.closest('input,textarea,select,[contenteditable="true"]'));
 const reset=()=>{start=null;distance=0;if(!refreshing){indicator.hidden=true;indicator.classList.remove('ready');}};
 doc.addEventListener('touchstart',event=>{
  reset();if(refreshing||event.touches.length!==1||win.scrollY>0||blocked()||event.target.closest('input,textarea,select,button,a,[contenteditable="true"]'))return;
  const touch=event.touches[0];start={x:touch.clientX,y:touch.clientY,id:touch.identifier};
 },{passive:true});
 doc.addEventListener('touchmove',event=>{
  if(!start)return;if(event.touches.length!==1||blocked()){reset();return;}
  const touch=event.touches[0];if(touch.identifier!==start.id){reset();return;}
  const dy=touch.clientY-start.y,dx=Math.abs(touch.clientX-start.x);
  if(dy<0||dx>Math.max(12,dy*.8)){reset();return;}
  if(dy<12)return;
  if(!event.cancelable){reset();return;}
  event.preventDefault();distance=dy;indicator.hidden=false;indicator.classList.toggle('ready',dy>=90);indicator.textContent=dy>=90?'اترك للتحديث':'اسحب للأسفل للتحديث';
 },{passive:false});
 doc.addEventListener('touchend',()=>{
  if(start&&distance>=90&&!blocked()){
   refreshing=true;start=null;indicator.hidden=false;indicator.classList.add('ready');indicator.textContent='جارٍ تحديث المكتبة…';
   win.setTimeout(onRefresh,100);
  }else reset();
 },{passive:true});
 doc.addEventListener('touchcancel',reset,{passive:true});
}
