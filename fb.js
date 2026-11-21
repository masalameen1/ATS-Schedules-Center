/* Firebase layer: emulates the small db/user API the app uses, on Firestore + Firebase Auth */
(function(){
var cfg={apiKey:"AIzaSyBzia4lKoSi2KCns9xNWAJi5cri1Rd1fEo",authDomain:"ats-schedule-center-e2e1c.firebaseapp.com",projectId:"ats-schedule-center-e2e1c",storageBucket:"ats-schedule-center-e2e1c.firebasestorage.app",messagingSenderId:"278973444209",appId:"1:278973444209:web:333aa846b8ca1bbb266bd0"};
window.__MAILURL=window.__MAILURL||'';var ADMIN=["hatem.hamdeh@actvet.gov.ae", "naser.alashab@actvet.gov.ae", "fahad.almansoori@actvet.gov.ae", "ahmad.debek@actvet.gov.ae", "tamer.ahmed@actvet.gov.ae", "furat.jasim@actvet.gov.ae", "abedalla.banihani@actvet.gov.ae", "waseem.alhausani@actvet.gov.ae", "mohammad.salamin@actvet.gov.ae", "mahmoud.yousef@actvet.gov.ae", "anjali.rajan@actvet.gov.ae", "masalameen@gmail.com"];
var ok=typeof firebase!=='undefined';
if(!ok){window.claude={use:async function(){return null}};window.__fbLogin=async function(){throw {code:'offline'}};window.__fbErr=function(){return 'Cannot reach the sign-in service.'};return}
firebase.initializeApp(cfg);
var auth=firebase.auth(),fs=firebase.firestore();
fs.settings({ignoreUndefinedProperties:true});
var ready=auth.setPersistence(firebase.auth.Auth.Persistence.SESSION).catch(function(){});
var first=new Promise(function(res){var u=auth.onAuthStateChanged(function(x){u();res(x)})});
try{if(!sessionStorage.getItem('tt_role'))first=first.then(function(x){return x?auth.signOut().then(function(){return null}):null})}catch(e){}
window.__fbLogin=async function(email,pw){await ready;var c=await auth.signInWithEmailAndPassword(email,pw);return c.user};
window.__fbErr=function(e){var c=(e&&e.code)||'';
 if(/user-not-found|wrong-password|invalid-credential|invalid-login/.test(c))return 'Incorrect email or password. Ask the coordinator.';
 if(/too-many/.test(c))return 'Too many attempts. Wait a few minutes and try again.';
 if(/network/.test(c))return 'No connection. Check the internet and try again.';
 return 'Could not sign in ('+c+').'};
window.__fbReset=function(email){return auth.sendPasswordResetEmail(email)};
window.__fbChangePw=async function(oldp,newp){var u=auth.currentUser;if(!u)throw {code:'no-user'};await u.reauthenticateWithCredential(firebase.auth.EmailAuthProvider.credential(u.email,oldp));await u.updatePassword(newp)};
window.__pubDoc=function(path){var r=fs.doc(path);return{onSnapshot:function(a,b){return r.onSnapshot(function(d){a(d.exists?JSON.parse(d.data().j||'{}'):undefined)},b||function(){})}}};
/* public create-only write (students submit a broadcast request without signing in; Firestore rules allow create only, small size) */
window.__pubAdd=function(path,o){return fs.doc(path).set({j:JSON.stringify(o),t:Date.now()})};
var pk=function(d){return d.exists?JSON.parse(d.data().j||'{}'):undefined};
function snap(d){return{id:d.id,exists:d.exists,data:function(){return d.exists?JSON.parse(d.data().j||'{}'):undefined}}}
function doc(path){var r=fs.doc(path);return{
 get:async function(){return snap(await r.get())},
 set:function(o){return r.set({j:JSON.stringify(o===undefined?null:o),t:Date.now()})},
 update:async function(p){var s=await r.get(),o=s.exists?JSON.parse(s.data().j||'{}'):{};Object.assign(o,p);return r.set({j:JSON.stringify(o),t:Date.now()})},
 tx:function(fn){return fs.runTransaction(async function(t){var d=await t.get(r),o=d.exists?JSON.parse(d.data().j||'{}'):{};var n=fn(o)||o;t.set(r,{j:JSON.stringify(n),t:Date.now()})})},
 delete:function(){return r.delete()},
 onSnapshot:function(a,b){return r.onSnapshot(function(d){a(snap(d))},b||function(){})}}}
var DB={doc:doc,collection:function(p){var c=fs.collection(p);return{doc:function(i){return doc(p+'/'+i)},
 get:async function(){var q=await c.get();return{docs:q.docs.map(snap)}},
 onSnapshot:function(a,b){return c.onSnapshot(function(q){a({docs:q.docs.map(snap)})},b||function(){})}}}};
var cur=null;
/* owner panel: stores readable role arrays next to the JSON so Firestore rules can honour them */
window.__ownerSave=function(o){var by=function(r){return(o.people||[]).filter(function(p){return p.r===r}).map(function(p){return String(p.e).toLowerCase()})};
 return fs.doc('coord/owner').set({j:JSON.stringify(o),t:Date.now(),tch:by('tch'),wr:by('wr'),au:by('au'),cm:by('cm')})};

/* direct e-mail through the school's Google Apps Script relay (verified with the signed-in Firebase user) */
var MAIL_URL=window.__MAILURL||'';
/* Preferred: send FROM the signed-in teacher's own school mailbox (Microsoft Graph, delegated Mail.Send) */
var MS=window.__MSAL||null;
if(MS&&MS.clientId&&window.msal){var pca=null,msInit=null;
 var msGet=function(){if(!msInit){pca=new msal.PublicClientApplication({auth:{clientId:MS.clientId,authority:'https://login.microsoftonline.com/'+(MS.tenantId||'organizations'),redirectUri:location.origin},cache:{cacheLocation:'sessionStorage'}});msInit=pca.initialize()}return msInit.then(function(){return pca})};
 var msToken=async function(){var c=await msGet();var sc=['https://graph.microsoft.com/Mail.Send'];var acc=c.getAllAccounts()[0];
  if(acc){try{return(await c.acquireTokenSilent({scopes:sc,account:acc})).accessToken}catch(e){}}
  var r=await c.acquireTokenPopup({scopes:sc,prompt:acc?undefined:'select_account'});return r.accessToken};
 window.__sendMail=async function(m){var t=await msToken();
  var r=await fetch('https://graph.microsoft.com/v1.0/me/sendMail',{method:'POST',headers:{'Authorization':'Bearer '+t,'Content-Type':'application/json'},body:JSON.stringify({message:{subject:m.subject,body:{contentType:'Text',content:m.body},toRecipients:m.to.map(function(a){return{emailAddress:{address:a}}})},saveToSentItems:true})});
  if(r.status!==202&&!r.ok){var j=await r.json().catch(function(){return{}});throw new Error((j.error&&j.error.message)||('Mail error '+r.status))}return{ok:true}}}
else if(MAIL_URL){window.__sendMail=async function(m){var u=auth.currentUser;if(!u)throw new Error('Sign in first');var t=await u.getIdToken();
 var r=await fetch(MAIL_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({token:t,to:m.to,subject:m.subject,body:m.body})});
 var j=await r.json().catch(function(){return{ok:false,error:'Bad reply'}});if(!j.ok)throw new Error(j.error||'Failed');return j}}
if(window.__sendMail){var _sm=window.__sendMail;window.__sendMail=function(m){var to=((m&&m.to)||[]).map(String);if(!to.length||to.length>5||to.some(function(a){return !/^[a-z0-9._%+-]+@actvet\.gov\.ae$/i.test(a)}))return Promise.reject(new Error('Recipients must be 1-5 @actvet.gov.ae addresses'));return _sm({to:to,subject:String(m.subject||'').slice(0,200),body:String(m.body||'').slice(0,6000)})}}
var US={id:async function(){return cur?cur.uid:null},
 me:async function(){return cur?{id:cur.uid,email:cur.email,name:cur.displayName||'',canEdit:ADMIN.indexOf((cur.email||'').toLowerCase())>-1,isOwner:false}:null},
 canEdit:async function(){return !!cur&&ADMIN.indexOf((cur.email||'').toLowerCase())>-1},
 isOwner:async function(){return false},can:function(){return null}};
window.claude={use:async function(n){await ready;cur=await first.then(function(x){return x||auth.currentUser});
 if(n==='db')return cur?DB:null;if(n==='user')return cur?US:null;return null}};
auth.onAuthStateChanged(function(u){cur=u;try{window.__fbBanner(u)}catch(e){}});
window.__fbBanner=function(u){var id='fbVerBn',o=document.getElementById(id);if(o)o.remove();if(!u||u.emailVerified||!document.body||ADMIN.indexOf((u.email||'').toLowerCase())<0)return;
 var b=document.createElement('div');b.id=id;b.style.cssText='position:fixed;inset-inline:12px;bottom:12px;z-index:99999;max-width:520px;margin-inline:auto;background:#fff7e6;color:#5a3b00;border:1.5px solid #f0b34a;border-radius:16px;padding:12px 14px;font:600 13px/1.6 system-ui,sans-serif;box-shadow:0 8px 28px rgba(0,0,0,.18);direction:rtl';
 var m=document.createElement('div');m.textContent='⚠ بريدك غير موثّق في Firebase، ولن يُحفظ شيء في الخطة المشتركة حتى يُوثَّق · Your email is not verified, so saving is blocked.';b.appendChild(m);
 var r=document.createElement('div');r.style.cssText='display:flex;gap:6px;flex-wrap:wrap;margin-top:8px';
 var say=function(t){m.textContent=t};
 var mk=function(t,f){var x=document.createElement('button');x.type='button';x.textContent=t;x.style.cssText='border:1px solid #d9a441;background:#fff;border-radius:12px;padding:6px 10px;font:inherit;cursor:pointer';x.onclick=function(){x.disabled=true;Promise.resolve().then(f).catch(function(e){say('تعذّر: '+(e&&e.code||e&&e.message||e))}).then(function(){x.disabled=false})};r.appendChild(x)};
 mk('📧 أعد إرسال رسالة التوثيق',function(){return u.sendEmailVerification().then(function(){say('أُرسلت الرسالة. افحص صندوق الوارد والرسائل غير المرغوبة (Spam/Junk) وقد تتأخر دقائق · Sent — check Spam too.')})});
 mk('🔑 أرسل رابط تعيين كلمة المرور (يوثّق البريد أيضًا)',function(){return auth.sendPasswordResetEmail(u.email).then(function(){say('أُرسل رابط تعيين كلمة المرور. بعد استخدامه يصبح البريد موثّقًا · Reset link sent — using it also verifies the email.')})});
 mk('✔ وثّقتُ بريدي — حدّث الآن',function(){return u.reload().then(function(){return (auth.currentUser||u).getIdToken(true)}).then(function(){var c=auth.currentUser||u;if(c.emailVerified){b.remove();window.__fbVerSent=0;if(window.__smallNote)window.__smallNote(['✔ تم توثيق البريد — يمكنك الحفظ الآن · Verified, you can save now'])}else say('لم يُوثَّق بعد · Still not verified')})});
 mk('✕',function(){b.remove()});
 b.appendChild(r);document.body.appendChild(b)};
window.__fbVerify=function(force){var u=cur||auth.currentUser;if(!u||u.emailVerified)return Promise.resolve(false);if(window.__fbVerSent&&!force)return Promise.resolve('again');window.__fbVerSent=1;return u.sendEmailVerification().then(function(){return true})};
window.__fbWho=function(){var u=cur||auth.currentUser;return u?{e:u.email||'',v:!!u.emailVerified}:null};
})();
