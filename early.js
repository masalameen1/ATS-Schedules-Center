try{if(sessionStorage.getItem('tt_role')&&sessionStorage.getItem('tt_gate')==='1')document.getElementById('boot').style.display='flex'}catch(e){}
try{(function(){var m=(sessionStorage.getItem('tt_mail')||'').toLowerCase();if(sessionStorage.getItem('tt_role')!=='teacher'||!m)return;var c=JSON.parse(localStorage.getItem('tt_dpcache_'+m)||'null'),sk='tt_dpshown_'+m,sn=JSON.parse(sessionStorage.getItem(sk)||'[]');
if(!c||!((c.ds&&c.ds.length)||c.un)||(c.toks&&c.toks.length&&!c.toks.some(function(t){return sn.indexOf(t)<0})))return;
var d=document.createElement('div');d.id='dutyPop';d.setAttribute('role','alertdialog');
var h=document.createElement('h3');h.textContent='🔔 مهمة جديدة · New duty';d.appendChild(h);
if(c.ds&&c.ds.length){var l=document.createElement('div');l.className='dpL';var b=document.createElement('b');b.textContent='✉️ لديك مهام جديدة · You have '+c.ds.length+' new dut'+(c.ds.length>1?'ies':'y')+':';l.appendChild(b);c.ds.slice(0,4).forEach(function(t){var x=document.createElement('div');x.textContent=t;l.appendChild(x)});d.appendChild(l)}
if(c.un){var l2=document.createElement('div');l2.className='dpL';l2.textContent='🔔 '+c.un+' إشعار جديد · new notification'+(c.un>1?'s':'')+' in the bell';d.appendChild(l2)}
function seen(){try{sessionStorage.setItem(sk,JSON.stringify(sn.concat(c.toks||[])))}catch(e){}}
var a=document.createElement('div');a.className='dpA';var g=document.createElement('button');g.className='dpGo';g.textContent=(c.ds&&c.ds.length)?'🔔 عرض · View':'عرض الإشعارات · View';
g.onclick=function(){try{sessionStorage.setItem('tt_gonow',(c.ds&&c.ds.length)?'duty':'bell')}catch(e){}seen();d.remove()};
var n=document.createElement('button');n.className='dpNo';n.textContent='لاحقًا · Later';n.onclick=function(){seen();d.remove()};a.appendChild(g);a.appendChild(n);d.appendChild(a);document.body.appendChild(d)})()}catch(e){}