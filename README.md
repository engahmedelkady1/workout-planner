<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>سجل التمارين</title>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap" rel="stylesheet">
<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-auth-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore-compat.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--bg:#0a0708;--card:#161012;--card2:#1e1518;--wine:#7a1530;--wine2:#4a0c1c;--gold:#d4af37;--gold2:#f0d77a;--tx:#f3ece4;--mu:#a3949a;--ln:#33232a;color-scheme:dark;box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font-family:'Cairo',system-ui,sans-serif;min-height:100vh}
.wrap{max-width:560px;margin:0 auto;padding:18px 14px 90px}
header{text-align:center;padding:22px 10px 8px}
.logo{font-weight:800;font-size:26px;color:var(--gold);letter-spacing:1px}
.sub{color:var(--mu);font-size:14px;margin-top:2px}
.gl{height:2px;background:linear-gradient(90deg,transparent,var(--gold),transparent);margin:14px 0 20px}
h2{font-size:20px;margin:0 0 14px;color:var(--gold2)}
.big{display:block;width:100%;border:1px solid var(--gold);border-radius:18px;padding:26px 18px;margin-bottom:14px;background:linear-gradient(135deg,var(--wine),var(--wine2));color:var(--tx);font:inherit;text-align:right;cursor:pointer}
.big b{display:block;font-size:24px;color:var(--gold2)}
.big span{font-size:13px;color:#e2cfd4}
.big.dark{background:linear-gradient(135deg,#1d1417,#0e0a0b)}
.back{background:none;border:0;color:var(--gold);font:inherit;cursor:pointer;padding:4px 0;margin-bottom:8px}
.ex{background:var(--card);border:1px solid var(--ln);border-radius:16px;margin-bottom:12px;overflow:hidden}
.exh{display:flex;align-items:center;gap:10px;padding:12px 14px;background:linear-gradient(90deg,var(--wine2),var(--card2));border-bottom:1px solid var(--ln)}
.num{min-width:28px;height:28px;border-radius:50%;background:var(--gold);color:#1a0b10;display:grid;place-items:center;font-weight:800;font-size:13px}
.exn{font-weight:600;font-size:15px;direction:ltr;text-align:left;flex:1}
.set{display:grid;grid-template-columns:42px 1fr 1fr;gap:8px;align-items:center;padding:8px 14px;border-top:1px solid #231a1e}
.set:first-of-type{border-top:0}
.sl{color:var(--mu);font-size:13px}
.f{position:relative}
.f input{width:100%;background:#0d0909;border:1px solid var(--ln);color:var(--tx);border-radius:10px;padding:10px 8px 10px 34px;font:inherit;font-size:16px;text-align:center}
.f input:focus{outline:none;border-color:var(--gold)}
.f small{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--mu);font-size:11px;pointer-events:none}
.last{padding:0 14px 10px;color:var(--mu);font-size:12px}
.up{color:#6fd98f}.dn{color:#e57b7b}
.add{background:var(--card);border:1px dashed var(--gold);border-radius:14px;padding:12px;margin-bottom:12px;display:flex;gap:8px;direction:ltr}
.add input{background:#0d0909;border:1px solid var(--ln);color:var(--tx);border-radius:10px;padding:9px;font:inherit;font-size:15px;min-width:0}
.add .n{flex:1}.add .s{width:54px;text-align:center}
.btn{border:0;border-radius:12px;padding:12px 18px;font:inherit;font-weight:800;font-size:16px;cursor:pointer;background:linear-gradient(135deg,var(--gold2),var(--gold));color:#1a0b10;text-decoration:none;display:inline-block;text-align:center}
.btn.o{background:transparent;border:1px solid var(--gold);color:var(--gold)}
.bar{position:fixed;left:0;right:0;bottom:0;padding:12px 14px calc(12px + env(safe-area-inset-bottom,0px));background:linear-gradient(0deg,var(--bg) 70%,transparent)}
.bar .btn{width:100%;max-width:532px;display:block;margin:0 auto;padding:15px}
.box{background:var(--card);border:1px solid var(--ln);border-radius:16px;padding:16px;margin-bottom:14px}
.box label{display:block;color:var(--mu);font-size:13px;margin-bottom:6px}
.box input{width:100%;background:#0d0909;border:1px solid var(--ln);color:var(--tx);border-radius:10px;padding:11px;font:inherit;font-size:16px;direction:ltr;text-align:left}
.row{display:flex;gap:10px;flex-wrap:wrap}.row .btn{flex:1}
.toast{position:fixed;top:calc(14px + env(safe-area-inset-top,0px));left:50%;transform:translateX(-50%);background:var(--wine);border:1px solid var(--gold);padding:10px 18px;border-radius:12px;z-index:9;font-size:14px}
.ck{font-size:54px;text-align:center;color:var(--gold)}
table{width:100%;border-collapse:collapse;font-size:13px}
td,th{padding:6px;border-bottom:1px solid var(--ln);text-align:center}
th{color:var(--gold)}td:nth-child(1){direction:ltr;text-align:left}
.sc{overflow-x:auto}
.pw{position:relative}
.pw input{padding-right:64px}
.eye{position:absolute;right:6px;top:50%;transform:translateY(-50%);background:none;border:0;color:var(--gold);cursor:pointer;font:inherit;font-size:13px;padding:8px}
</style>
</head>
<body>
<div class="wrap">
<header><div class="logo">IRON LOG</div><div class="sub">سجل أوزانك وتقدمك في الحديد</div></header>
<div class="gl"></div>
<div id="app"></div>
</div>
<div id="bar"></div>
<script>
/* ===== الصق إعدادات Firebase بتاعتك هنا ===== */
const CFG={apiKey:"AIzaSyBoQVjRJxTZniLu3jIsr_2Kag0b3XFogn8",authDomain:"workout-planner-8377f.firebaseapp.com",projectId:"workout-planner-8377f",appId:"1:445503159382:web:abc5853abdaa5cd6f54ff7"};
/* ========================================== */
const PLAN={
d1:{t:'اليوم الأول',g:'u',ex:[['Incline Dumbbell Press',2],['T Bar Row Machine',2],['Iliac Lat Pulldown',2],['Chest Press Machine',2],['Lateral Raises',2],['Supported Dumbbell Curls',2],['Overhead Triceps Extension',2],['Abs Rope Crunches',2]]},
d4:{t:'اليوم الرابع',g:'u',ex:[['Wide Grip Lat Pulldown',2],['Incline Press Machine',3],['Wide Grip Seated Row',2],['Reversed Curls',2],['Dumbbell Lateral Raise',2],['Zigzag Bar Tricep Extension',2],['Behind Body Curls',2],['Abs Rope Crunches',2]]},
d2:{t:'اليوم الثاني',g:'l',ex:[['Squat on Smith Machine',3],['Lying or Seated Leg Curl',3],['Abduction Machine',2],['Leg Press Calf Raise',3],['Shoulder Press Machine',2],['Hyper Extension',3]]},
d5:{t:'اليوم الخامس',g:'l',ex:[['Romanian Deadlift',3],['Leg Extensions',3],['Leg Press Calf Raise',2],['Abduction Machine',3]]}
};
const G={u:'الجسم العلوي · UPPER',l:'الجسم السفلي · LOWER'};
const $=id=>document.getElementById(id);
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const ls={get(k,d){try{const v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(e){return d}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}}};
let hist=[],custom={},email=ls.get('il_email',''),user=null,auth=null,fdb=null;
let st={step:'home',g:null,day:null,cur:{},last:null};
const today=()=>{const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')};
const exs=k=>PLAN[k].ex.concat(custom[k]||[]);
async function persist(){
 if(!user||!fdb)return;
 try{await fdb.collection('users').doc(user.uid).set({hist:JSON.stringify(hist),custom:JSON.stringify(custom)})}
 catch(e){toast('تعذر حفظ البيانات على السيرفر')}
}
const ERR={'auth/invalid-email':'الإيميل غير صحيح','auth/missing-password':'اكتب كلمة السر','auth/weak-password':'كلمة السر ضعيفة (6 حروف على الأقل)','auth/email-already-in-use':'الإيميل ده مسجل قبل كده','auth/invalid-credential':'الإيميل أو كلمة السر غلط','auth/user-not-found':'الحساب مش موجود','auth/wrong-password':'كلمة السر غلط','auth/too-many-requests':'محاولات كتيرة، استنى شوية','auth/network-request-failed':'مفيش اتصال بالإنترنت'};
function toast(m){const t=document.createElement('div');t.className='toast';t.textContent=m;document.body.appendChild(t);setTimeout(()=>t.remove(),2600)}
function go(s,p){Object.assign(st,p||{});st.step=s;render();window.scrollTo(0,0)}
function lastOf(k,ex,si){const s=hist.find(h=>h.day===k);if(!s)return null;return s.rows.find(r=>r.ex===ex&&r.set===si+1)||null}
function render(){
 const a=$('app'),b=$('bar');b.innerHTML='';
 if(st.step==='auth'){
  const reg=st.mode==='reg';
  a.innerHTML=`<h2>${reg?'إنشاء حساب جديد':'تسجيل الدخول'}</h2><div class="box"><label>الإيميل</label><input id="ae" type="email" autocomplete="email"><label style="margin-top:12px">كلمة السر</label><div class="pw"><input id="ap" type="password" autocomplete="${reg?'new-password':'current-password'}" placeholder="6 حروف على الأقل"><button type="button" class="eye" id="eye">إظهار</button></div><div id="er" class="dn" style="margin-top:10px;font-size:14px"></div></div><button class="btn" style="width:100%" id="ab2">${reg?'إنشاء الحساب':'دخول'}</button><div class="sub" style="text-align:center;margin-top:16px"><a id="tg" style="color:var(--gold);cursor:pointer">${reg?'عندي حساب؟ سجّل دخول':'مفيش حساب؟ أنشئ حساب جديد'}</a></div>`;
  $('tg').onclick=()=>go('auth',{mode:reg?'login':'reg'});
  $('eye').onclick=()=>{const p=$('ap'),h=p.type==='password';p.type=h?'text':'password';$('eye').textContent=h?'إخفاء':'إظهار'};
  if(!auth)$('er').textContent='لازم تلصق إعدادات Firebase في أول الكود الأول.';
  $('ab2').onclick=async()=>{
   if(!auth)return;const e=$('ae').value.trim(),p=$('ap').value;$('er').textContent='';$('ab2').disabled=true;
   try{reg?await auth.createUserWithEmailAndPassword(e,p):await auth.signInWithEmailAndPassword(e,p)}
   catch(x){$('er').textContent=ERR[x.code]||'حصل خطأ، حاول تاني';$('ab2').disabled=false}
  };
 }
 else if(st.step==='home'){
  a.innerHTML=`<h2>هتتمرن النهاردة إيه؟</h2>
  <button class="big" data-g="u"><b>UPPER BODY</b><span>الجسم العلوي · صدر، ضهر، أكتاف، دراعات</span></button>
  <button class="big dark" data-g="l"><b>LOWER BODY</b><span>الجسم السفلي · رجل، سمانة، أسفل الضهر</span></button>
  <div class="gl"></div><div class="row"><button class="btn o" id="hs">السجل والتصدير (${hist.length})</button></div>
  <div class="sub" style="text-align:center;margin-top:18px;direction:ltr">${esc(user?user.email:'')} · <a id="lo" style="color:var(--gold);cursor:pointer">تسجيل خروج</a></div>`;
  $('lo').onclick=()=>auth.signOut();
  a.querySelectorAll('[data-g]').forEach(x=>x.onclick=()=>go('variant',{g:x.dataset.g}));
  $('hs').onclick=()=>go('hist');
 }
 else if(st.step==='variant'){
  const ks=Object.keys(PLAN).filter(k=>PLAN[k].g===st.g);
  a.innerHTML=`<button class="back" id="bk">← رجوع</button><h2>${G[st.g]}</h2><div class="sub" style="margin-bottom:14px">اختار اليوم الخاص بالجدول</div>`+
  ks.map(k=>`<button class="big ${st.g==='l'?'dark':''}" data-k="${k}"><b>${PLAN[k].t}</b><span>${exs(k).length} تمارين</span></button>`).join('');
  $('bk').onclick=()=>go('home');
  a.querySelectorAll('[data-k]').forEach(x=>x.onclick=()=>go('log',{day:x.dataset.k,cur:{}}));
 }
 else if(st.step==='log'){
  const k=st.day,L=exs(k);
  let h=`<button class="back" id="bk">← رجوع</button><h2>${PLAN[k].t} · ${st.g==='u'?'Upper':'Lower'}</h2><div class="sub" style="margin:-8px 0 14px">${today()}</div>`;
  L.forEach((e,i)=>{
   h+=`<div class="ex"><div class="exh"><div class="num">${i+1}</div><div class="exn">${esc(e[0])}</div></div>`;
   for(let s=0;s<e[1];s++){
    const l=lastOf(k,e[0],s);
    h+=`<div class="set"><div class="sl">سيت ${s+1}</div>
    <div class="f"><input type="number" inputmode="decimal" step="0.5" min="0" data-i="${i}" data-s="${s}" data-f="w" placeholder="${l?l.w:'0'}"><small>كجم</small></div>
    <div class="f"><input type="number" inputmode="numeric" min="0" data-i="${i}" data-s="${s}" data-f="r" placeholder="${l?l.r:'0'}"><small>عدة</small></div></div>`;
   }
   h+=`<div class="last" id="ls${i}"></div></div>`;
  });
  h+=`<div class="add"><input class="n" id="an" placeholder="Add exercise"><input class="s" id="as" type="number" value="3" min="1"><button class="btn" id="ab">+</button></div>`;
  a.innerHTML=h;
  $('bk').onclick=()=>go('variant');
  $('ab').onclick=()=>{const n=$('an').value.trim();if(!n)return;(custom[k]=custom[k]||[]).push([n,Math.max(1,+$('as').value||3)]);persist();go('log',{})};
  a.querySelectorAll('.set input').forEach(inp=>inp.oninput=()=>{
   const key=inp.dataset.i+'_'+inp.dataset.s;(st.cur[key]=st.cur[key]||{})[inp.dataset.f]=inp.value;
   cmp(+inp.dataset.i,L);
  });
  a.querySelectorAll('.set input').forEach(inp=>{const c=st.cur[inp.dataset.i+'_'+inp.dataset.s];if(c&&c[inp.dataset.f]!=null)inp.value=c[inp.dataset.f]});
  L.forEach((_,i)=>cmp(i,L));
  b.innerHTML='<div class="bar"><button class="btn" id="sv">حفظ التمرين وإنشاء الشيت</button></div>';
  $('sv').onclick=save;
 }
 else if(st.step==='done'){
  const s=st.last;
  a.innerHTML=`<div class="ck">✔</div><h2 style="text-align:center">تم حفظ التمرين</h2>
  <div class="box"><div class="sub" style="margin-bottom:8px">${s.date} · ${esc(s.label)} · ${s.rows.length} مجموعة</div>
  <div class="sc"><table><tr><th>التمرين</th><th>سيت</th><th>كجم</th><th>عدة</th></tr>${s.rows.map(r=>`<tr><td>${esc(r.ex)}</td><td>${r.set}</td><td>${r.w}</td><td>${r.r}</td></tr>`).join('')}</table></div></div>
  <div class="box"><label>إيميلك</label><input id="em" type="email" value="${esc(email)}" placeholder="name@email.com">
  <div class="sub" style="margin:10px 0">الموقع مش بيقدر يبعت إيميل لوحده. حمّل الشيت، وبعدها افتح الإيميل (فيه الملخص جاهز) وأرفق الملف.</div>
  <div class="row"><button class="btn" id="dl">تحميل شيت Excel</button><a class="btn o" id="ml" target="_blank" href="#">فتح الإيميل</a></div><div class="row" style="margin-top:10px"><button class="btn o" id="sh">مشاركة الملف (إيميل / واتساب)</button></div></div>
  <button class="btn o" style="width:100%" id="hm">تمرين جديد</button>`;
  const upd=()=>{email=$('em').value.trim();ls.set('il_email',email);const body=s.date+' - '+s.label+'\n\n'+s.rows.map(r=>r.ex+' | سيت '+r.set+' | '+r.w+' كجم x '+r.r).join('\n');$('ml').href='mailto:'+encodeURIComponent(email)+'?subject='+encodeURIComponent('تمرين '+s.date+' - '+s.label)+'&body='+encodeURIComponent(body)};
  $('em').oninput=upd;upd();
  $('dl').onclick=()=>dl(s.date);$('sh').onclick=()=>shareFile(s.date);$('hm').onclick=()=>go('home');
 }
 else if(st.step==='hist'){
  a.innerHTML=`<button class="back" id="bk">← رجوع</button><h2>السجل</h2>`+
  (hist.length?hist.map((h,n)=>`<div class="box hc" data-n="${n}" style="padding:12px;cursor:pointer"><b style="color:var(--gold2)">${h.date}</b> · ${esc(h.label)}<div class="sub">${h.rows.length} مجموعة · اضغط لعرض التفاصيل ‹</div></div>`).join('')+'<button class="btn" style="width:100%" id="dla">تحميل السجل الكامل Excel</button>':'<div class="sub">لسه مفيش تمارين متسجلة.</div>');
  $('bk').onclick=()=>go('home');
  if(hist.length)$('dla').onclick=()=>dl(today());
  a.querySelectorAll('.hc').forEach(x=>x.onclick=()=>go('sess',{n:+x.dataset.n,del:false}));
 }
 else if(st.step==='sess'){
  const s=hist[st.n];if(!s){go('hist');return}
  const g={};s.rows.forEach(r=>(g[r.ex]=g[r.ex]||[]).push(r));
  a.innerHTML=`<button class="back" id="bk">← رجوع للسجل</button><h2>${s.date}</h2><div class="sub" style="margin:-8px 0 14px">${esc(s.label)}</div>`+
  Object.keys(g).map((k,i)=>`<div class="ex"><div class="exh"><div class="num">${i+1}</div><div class="exn">${esc(k)}</div></div>`+
   g[k].map(r=>`<div class="set" style="grid-template-columns:60px 1fr 1fr"><div class="sl">سيت ${r.set}</div><div>${r.w||'-'} <span class="sl">كجم</span></div><div>${r.r||'-'} <span class="sl">عدة</span></div></div>`).join('')+'</div>').join('')+
  `<div class="row"><button class="btn" id="d1">تحميل شيت اليوم ده</button><button class="btn o" id="dx">حذف التمرين</button></div>`;
  $('bk').onclick=()=>go('hist');
  $('d1').onclick=()=>dl(s.date,st.n);
  $('dx').onclick=()=>{if(!st.del){st.del=true;$('dx').textContent='اضغط تاني للتأكيد';return}hist.splice(st.n,1);persist();go('hist')};
 }
}
function cmp(i,L){
 const k=st.day,e=L[i];let best=null;
 for(let s=0;s<e[1];s++){const c=st.cur[i+'_'+s],l=lastOf(k,e[0],s);if(c&&+c.w&&l){const d=(+c.w)-(+l.w);best=best==null?d:Math.max(best,d)}}
 const el=$('ls'+i);if(!el)return;
 el.innerHTML=best==null?'':best>0?`<span class="up">▲ زيادة ${best} كجم عن آخر مرة</span>`:best<0?`<span class="dn">▼ أقل ${-best} كجم عن آخر مرة</span>`:'نفس الوزن بتاع آخر مرة';
}
function save(){
 const k=st.day,L=exs(k),rows=[];
 L.forEach((e,i)=>{for(let s=0;s<e[1];s++){const c=st.cur[i+'_'+s];if(c&&(c.w||c.r))rows.push({ex:e[0],set:s+1,w:c.w||'',r:c.r||''})}});
 if(!rows.length){toast('اكتب وزن أو عدّات الأول');return}
 const sess={date:today(),day:k,label:PLAN[k].t+' - '+(st.g==='u'?'Upper':'Lower'),rows};
 hist.unshift(sess);persist();
 go('done',{last:sess});dl(sess.date);
}
function sheet(rows){
 const h=['التاريخ','اليوم','التمرين','المجموعة','الوزن (كجم)','التكرارات'];
 const ws=XLSX.utils.aoa_to_sheet([h].concat(rows));ws['!cols']=[{wch:12},{wch:22},{wch:30},{wch:10},{wch:12},{wch:12}];return ws;
}
function build(date,idx){
  const cur=(idx!=null&&hist[idx])||hist.find(x=>x.date===date)||hist[0];
  const m=s=>s.rows.map(r=>[s.date,s.label,r.ex,r.set,r.w===''?'':+r.w,r.r===''?'':+r.r]);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,sheet(m(cur)),'جلسة اليوم');
  XLSX.utils.book_append_sheet(wb,sheet([].concat(...hist.slice().reverse().map(m))),'السجل الكامل');
  wb.Workbook={Views:[{RTL:true}]};
  const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  return new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function dl(date,idx){
 try{
  const u=URL.createObjectURL(build(date,idx));
  const a=document.createElement('a');a.href=u;a.download='workout_'+date+'.xlsx';
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(u),3000);
 }catch(e){toast('تعذر إنشاء الملف')}
}
async function shareFile(date){
 try{
  const f=new File([build(date)],'workout_'+date+'.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  if(navigator.canShare&&navigator.canShare({files:[f]}))await navigator.share({files:[f],title:'تمرين '+date});
  else toast('المشاركة غير مدعومة، استخدم التحميل');
 }catch(e){}
}
$('app').innerHTML='<div class="sub" style="text-align:center">جاري التحميل...</div>';
if(CFG.apiKey.indexOf('PASTE')!==0){
 firebase.initializeApp(CFG);auth=firebase.auth();fdb=firebase.firestore();
 auth.onAuthStateChanged(async u=>{
  user=u;
  if(!u){hist=[];custom={};go('auth',{mode:'login'});return}
  try{
   const d=await fdb.collection('users').doc(u.uid).get();
   if(d.exists){hist=JSON.parse(d.data().hist||'[]');custom=JSON.parse(d.data().custom||'{}')}
   else{hist=ls.get('il_hist',[]);custom=ls.get('il_custom',{});await persist();try{localStorage.removeItem('il_hist');localStorage.removeItem('il_custom')}catch(e){}}
  }catch(e){toast('تعذر تحميل بياناتك')}
  if(!email)email=u.email;go('home');
 });
}else go('auth',{mode:'login'});
</script>
</body>
</html>
