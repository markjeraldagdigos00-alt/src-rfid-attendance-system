const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const multer = require('multer');
const https = require('https');
const querystring = require('querystring');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');

// ===== DATABASE (MongoDB + Mongoose) =====
const mongoose = require('mongoose');
const { Schema, model } = mongoose;

const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/scholarhub_db?appName=Cluster0';
mongoose.connect(MONGO_URI)
  .then(() => console.log('[DATABASE] Connected to MongoDB Atlas successfully!'))
  .catch(err => console.error('[DATABASE ERROR] Could not connect to MongoDB:', err.message));

const S = { type: String, default: '' };

// students: one document per student (cardStatus: active | blocked | inactive)
const Student = model('Student', new Schema({
  uid: S, studentId: { type: String, required: true }, name: { type: String, required: true },
  yearLevel: { type: String, default: 'Grade 7' }, section: { type: String, default: 'A' },
  photo: S, photoData: Buffer, photoType: S, email: S, phone: S, cardStatus: { type: String, default: 'active' }
}));

// dailyattendance: one document per student per day (AM in/out + PM in/out)
const Daily = model('DailyAttendance', new Schema({
  date: String, uid: S, studentId: S, name: S, yearLevel: S, section: S, photo: S,
  amIn: S, amOut: S, pmIn: S, pmOut: S, late: Boolean, earlyOut: Boolean
}));

// scanlogs: every single tap (TIME-IN / TIME-OUT, AM / PM)
const ScanLog = model('ScanLog', new Schema({
  date: String, time: String, studentId: S, name: S, uid: S, action: S, period: S, status: S,
  at: { type: Date, default: Date.now }
}));

// unauthorized: unknown or blocked RFID cards
const Unauthorized = model('Unauthorized', new Schema({
  uid: String, count: { type: Number, default: 1 }, reason: S, lastSeen: { type: Date, default: Date.now }
}));

const Grade = model('Grade', new Schema({ name: String }));
const Section = model('Section', new Schema({ yearLevel: String, name: String, adviser: S }));
const Teacher = model('Teacher', new Schema({ name: String, email: S, phone: S, subject: S, username: S }));
const Announcement = model('Announcement', new Schema({ title: String, message: S, at: { type: Date, default: Date.now } }));
const Excuse = model('Excuse', new Schema({ studentId: String, name: S, date: String, reason: S }));
const Notification = model('Notification', new Schema({ message: String, level: S, at: { type: Date, default: Date.now } }));
const ActivityLog = model('ActivityLog', new Schema({ user: S, action: String, at: { type: Date, default: Date.now } }));

// config: one document (school info, schedule, admin account, notification settings, ESP status)
const Config = model('Config', new Schema({
  systemName: { type: String, default: 'School RFID Attendance System' }, logoPath: S,
  school: { name: { type: String, default: 'Batac National High School' }, address: S, contact: S },
  sched: {
    opening: { type: String, default: '06:00' }, late: { type: String, default: '07:30' },
    noon: { type: String, default: '12:00' }, dismissal: { type: String, default: '16:00' }
  },
  latestUid: S, lastPing: Date,
  adminUser: { type: String, default: 'admin' },
  adminPass: { type: String, default: require('crypto').createHash('sha256').update('admin123').digest('hex') },
  enableEmail: { type: Boolean, default: true }, enableSms: { type: Boolean, default: false }, semaphoreApiKey: S
}));

async function getConfig() { return (await Config.findOne()) || Config.create({}); }

const D = { mongoose, Student, Daily, ScanLog, Unauthorized, Grade, Section, Teacher, Announcement, Excuse, Notification, ActivityLog, Config, getConfig };

// ===== SERVER =====

const app = express();
const PORT = process.env.PORT || 3000;
const TZ = 'Asia/Manila'; // school time, even when hosted on a UTC server
const ymd = d => d.toLocaleDateString('en-CA', { timeZone: TZ });
const hm = d => d.toLocaleTimeString('en-GB', { timeZone: TZ, hour12: false }).slice(0, 5);
const sha = p => crypto.createHash('sha256').update(String(p)).digest('hex');
const log = a => D.ActivityLog.create({ user: 'admin', action: a }).catch(() => {});
const notify = (message, level = 'info') => D.Notification.create({ message, level }).catch(() => {});


// ===== EMBEDDED PAGES (dashboard, scanner, stylesheet) =====
const CSS = `:root{--ink:#12263a;--bg:#eef1f5;--line:#dde3ea;--green:#1f8a5b;--amber:#e8a317;--red:#d64545;--blue:#2f6fdb;--mut:#6b7a8c}
*{box-sizing:border-box}
body{margin:0;font-family:Figtree,'Segoe UI',sans-serif;background:var(--bg);color:#1c2733;font-size:15px}
.app{display:flex;min-height:100vh}
nav{width:255px;flex-shrink:0;background:var(--ink);color:#cfd8e3;padding:14px 10px;position:sticky;top:0;height:100vh;overflow-y:auto}
nav .brand{display:flex;align-items:center;gap:10px;padding:6px 10px 16px;color:#fff;font-weight:700;line-height:1.2}
nav .brand img{height:42px;border-radius:6px}
nav a{display:block;padding:8px 12px;border-radius:6px;color:inherit;text-decoration:none;font-size:14px}
nav a:hover{background:#1d3a55}nav a.on{background:var(--green);color:#fff;font-weight:600}
main{flex:1;min-width:0;padding:26px 30px}
h2{margin:0 0 4px;font-size:26px;color:var(--ink)}.sub{color:var(--mut);margin:0 0 18px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-bottom:22px}
.stat{background:#fff;border-radius:10px;padding:16px 18px;border-left:5px solid var(--blue)}
.stat b{display:block;font-size:34px;line-height:1.1}.stat span{color:var(--mut);font-size:13px}
.stat.g{border-color:var(--green)}.stat.r{border-color:var(--red)}.stat.a{border-color:var(--amber)}
.card{background:#fff;border-radius:10px;padding:18px;margin-bottom:18px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;align-items:end}
label{font-size:13px;font-weight:600;color:#34495e;display:block}
input,select,textarea{width:100%;padding:9px 10px;margin-top:4px;border:1px solid #c9d2dc;border-radius:6px;font:inherit;background:#fff}
button{background:var(--green);color:#fff;border:0;border-radius:6px;padding:10px 16px;font:inherit;font-weight:600;cursor:pointer}
button.alt{background:var(--blue)}button.danger{background:var(--red)}button.warn{background:var(--amber)}
button.sm{padding:5px 10px;font-size:12px;margin:1px}
.bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px}.bar>*{width:auto;margin:0}
.tw{overflow-x:auto;background:#fff;border-radius:10px}
table{width:100%;border-collapse:collapse}
th{background:#f6f8fa;text-align:left;font-size:13px;color:var(--mut);padding:10px 12px;border-bottom:1px solid var(--line)}
td{padding:9px 12px;border-bottom:1px solid #eef1f4;vertical-align:middle}
.who{display:flex;align-items:center;gap:10px}.who small{display:block;color:var(--mut)}
.av{width:38px;height:38px;border-radius:50%;object-fit:cover;background:#d9e2ec;display:inline-flex;align-items:center;justify-content:center;font-weight:700;color:var(--ink)}
.b{font-style:normal;font-size:12px;font-weight:700;padding:3px 8px;border-radius:5px;color:#fff;margin-right:3px;display:inline-block}
.b.green{background:var(--green)}.b.red{background:var(--red)}.b.amber{background:var(--amber)}.b.blue{background:var(--blue)}
.empty{color:var(--mut);padding:24px;text-align:center}
.bars{display:flex;align-items:flex-end;gap:8px;height:220px}
.col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%;font-size:11px;color:var(--mut)}
.col div{width:100%;background:var(--green);border-radius:4px 4px 0 0;position:relative}.col div i{position:absolute;bottom:0;width:100%;background:var(--red)}
dialog{border:0;border-radius:12px;padding:22px;width:min(560px,94vw)}dialog::backdrop{background:#0008}
.login{display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}
.box{background:#fff;padding:28px;border-radius:12px;width:100%;max-width:380px;display:flex;flex-direction:column;gap:12px;text-align:center}
.box h2{font-size:22px}.box p{margin:0;color:var(--mut)}.err{color:var(--red)!important}.box img{align-self:center}
@media(max-width:800px){.app{flex-direction:column}nav{width:100%;height:auto;position:static;max-height:240px}main{padding:16px}}
`;

const INDEX_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>School Attendance</title>
<link href="https://fonts.googleapis.com/css2?family=Figtree:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/assets/style.css">
</head>
<body>
<div class="app"><nav id="nav"></nav><main id="view"></main></div>
<dialog id="dlg"><form onsubmit="return saveStu(event)">
  <h2 id="dt">Student</h2><input type="hidden" id="s_id">
  <div class="grid">
    <label>RFID card UID<input id="s_uid"></label>
    <label>ID number<input id="s_studentId" required></label>
    <label>Full name<input id="s_name" required></label>
    <label>Grade level<select id="s_yearLevel"></select></label>
    <label>Section<input id="s_section" required></label>
    <label>Parent email<input id="s_email" type="email"></label>
    <label>Parent phone<input id="s_phone"></label>
    <label>Photo (max 1 MB)<input id="s_photo" type="file" accept="image/*"></label>
  </div><p style="margin:14px 0 0"><button>Save student</button> <button type="button" class="alt" onclick="lastUid()">Use last scanned card</button> <button type="button" class="danger" onclick="dlg.close()">Cancel</button></p>
</form></dialog>
<script>
const MENU=[['dashboard','🏠','Dashboard'],['scanner','📡','RFID Scanner','/scanner'],['students','👨‍🎓','Students'],['cards','💳','RFID Cards'],['grades','🏫','Grade Levels'],['sections','📚','Sections'],['teachers','👨‍🏫','Teachers'],['daily','📅','Daily Attendance'],['timeinout','🕐','Time In / Time Out'],['late','⚠️','Late Students'],['absent','❌','Absent Students'],['early','🚪','Early Out'],['reports','📊','Attendance Reports'],['analytics','📈','Attendance Analytics'],['search','🔍','Search Records'],['unauthorized','🚨','Unauthorized RFID'],['esp','📡','ESP8266 Status'],['notifications','🔔','Notifications'],['announcements','📢','Announcements'],['excuses','📋','Excuse/Absence Records'],['export','📤','Export Reports'],['schedule','⚙️','Attendance Schedule'],['school','🏫','School Information'],['admin','👤','Admin Account'],['logs','📝','Activity Logs'],['backup','💾','Backup & Restore'],['logout','🚪','Logout','/logout']];
const $=s=>document.querySelector(s),dlg=$('#dlg');
const api=async(u,o)=>{const r=await fetch(u,o);if(r.status===401)location='/login';const j=await r.json().catch(()=>({}));if(!r.ok)throw new Error(j.error||'Request failed');return j};
const post=(u,b)=>api(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const dayAgo=n=>new Date(Date.now()-n*864e5).toLocaleDateString('en-CA',{timeZone:'Asia/Manila'});
const fmt=v=>/^\\d{4}-\\d\\d-\\d\\dT/.test(v)?new Date(v).toLocaleString():v;
const h=(t,s='')=>\`<h2>\${t}</h2><p class="sub">\${s}</p>\`;
const table=(cols,rows)=>rows.length?\`<div class="tw"><table><thead><tr>\${cols.map(c=>\`<th>\${c[0]}</th>\`).join('')}</tr></thead><tbody>\${rows.map(r=>\`<tr>\${cols.map(c=>\`<td>\${c[1](r)}</td>\`).join('')}</tr>\`).join('')}</tbody></table></div>\`:'<p class="empty">Nothing to show yet.</p>';
let cur='dashboard',GR=[],CACHE={};
const gradeOpts=()=>GR.map(g=>\`<option>\${esc(g.name)}</option>\`).join('');

// attendance table pieces
const av=r=>r.photo?\`<img class="av" src="\${esc(r.photo)}">\`:\`<span class="av">\${esc((r.name||'?')[0])}</span>\`;
const who=r=>\`<div class="who">\${av(r)}<div><b>\${esc(r.name)}</b><small>\${esc(r.studentId)}</small></div></div>\`;
const badge=r=>r.absent?'<i class="b red">ABSENT</i>':(r.late?'<i class="b red">LATE</i>':'')+(r.earlyOut?'<i class="b amber">EARLY OUT</i>':'')||'<i class="b green">ON TIME</i>';
const ATT=[['Student',who],['Grade & Section',r=>\`\${esc(r.yearLevel)} - \${esc(r.section)}\`],['AM In',r=>r.amIn||'–'],['AM Out',r=>r.amOut||'–'],['PM In',r=>r.pmIn||'–'],['PM Out',r=>r.pmOut||'–'],['Status',badge]];
const attTable=async(q,cols=ATT)=>table(cols,await api('/api/attendance?'+new URLSearchParams(q)));
const attView=(t,s,f)=>async()=>h(t,s)+await attTable({filter:f});

const V={};
V.dashboard=async()=>{
  const [s,ann]=await Promise.all([api('/api/stats'),api('/api/announcements')]);
  const c=(n,l,k='')=>\`<div class="stat \${k}"><b>\${n}</b><span>\${l}</span></div>\`;
  return h('Dashboard','Today at school')+\`<div class="stats">\${c(s.total,'Total students')}\${c(s.present,'Present','g')}\${c(s.absent,'Absent','r')}\${c(s.late,'Late','a')}\${c(s.inside,'Inside the school','g')}\${c(s.early,'Early out','a')}</div>\`
   +\`<h3>Latest arrivals</h3>\`+(await attTable({})).split('</tr>').slice(0,9).join('</tr>')+(ann.length?\`</tbody></table></div><div class="card" style="margin-top:18px"><h3>📢 \${esc(ann[0].title)}</h3><p>\${esc(ann[0].message)}</p></div>\`:'</tbody></table></div>');
};
V.daily=attView('Daily Attendance','Everyone who scanned today');
V.late=attView('Late Students','Arrived after the late time');
V.absent=attView('Absent Students','No scan today and no excuse on file','absent');
V.early=attView('Early Out','Left before dismissal','early');
V.timeinout=async()=>h('Time In / Time Out','Every tap recorded today')+table([['Time',r=>r.time],['Student',r=>\`<b>\${esc(r.name)}</b><br><small>\${esc(r.studentId)}</small>\`],['Action',r=>\`<i class="b \${r.action==='TIME-IN'?'blue':'amber'}">\${r.action}</i>\`],['Period',r=>r.period],['Status',r=>esc(r.status)]],await api('/api/scanlogs?date='+dayAgo(0)));

// students + cards
V.students=async()=>{
  CACHE.students=await api('/api/students');
  return h('Students','Add, edit, delete and view student information')+\`<p><button onclick="openStu()">+ Add student</button> <a href="/student-register" target="_blank">Self-registration link</a></p>\`
   +table([['Student',who],['Grade & Section',r=>\`\${esc(r.yearLevel)} - \${esc(r.section)}\`],['Parent contact',r=>esc([r.email,r.phone].filter(Boolean).join(' / '))||'–'],['RFID',r=>r.uid?\`<code>\${esc(r.uid)}</code>\`:'<i class="b amber">No card</i>'],['',r=>\`<button class="sm" onclick="openStu('\${r._id}')">Edit</button><button class="sm danger" onclick="delStu('\${r._id}')">Delete</button>\`]],CACHE.students);
};
V.cards=async()=>h('RFID Cards','Register, assign, replace, deactivate or block cards')+table([['Student',who],['Card UID',r=>r.uid?\`<code>\${esc(r.uid)}</code>\`:'–'],['Card status',r=>r.uid?\`<i class="b \${r.cardStatus==='active'?'green':'red'}">\${r.cardStatus.toUpperCase()}</i>\`:'–'],['',r=>\`<button class="sm alt" onclick="cardAct('\${r._id}','assign')">\${r.uid?'Replace':'Assign'}</button>\`+(r.uid?\`<button class="sm warn" onclick="cardAct('\${r._id}','\${r.cardStatus==='active'?'inactive':'active'}')">\${r.cardStatus==='active'?'Deactivate':'Activate'}</button><button class="sm danger" onclick="cardAct('\${r._id}','blocked')">Block</button>\`:'')]],await api('/api/students'));
function openStu(id){
  const s=(CACHE.students||[]).find(x=>x._id===id)||{};
  $('#dt').innerText=id?'Edit student':'Add student';$('#s_id').value=id||'';$('#s_yearLevel').innerHTML=gradeOpts();
  ['uid','studentId','name','yearLevel','section','email','phone'].forEach(k=>$('#s_'+k).value=s[k]||(k==='yearLevel'?GR[0]?.name:''));
  $('#s_photo').value='';dlg.showModal();
}
async function lastUid(){$('#s_uid').value=(await api('/api/config')).latestUid||''}
async function saveStu(e){
  e.preventDefault();const f=new FormData();f.append('mongoId',$('#s_id').value);
  ['uid','studentId','name','yearLevel','section','email','phone'].forEach(k=>f.append(k,$('#s_'+k).value));
  if($('#s_photo').files[0])f.append('photo',$('#s_photo').files[0]);
  try{await api('/api/register',{method:'POST',body:f});dlg.close();go(cur)}catch(x){alert(x.message)}
}
async function delStu(id){if(confirm('Remove this student?')){await post('/api/delete-student',{id});go(cur)}}
async function cardAct(id,act){
  const b={id,cardStatus:act};
  if(act==='assign'){const u=prompt('Scan the new card, then confirm its UID:',(await api('/api/config')).latestUid);if(!u)return;b.uid=u;b.cardStatus='active'}
  try{await post('/api/card',b)}catch(x){alert(x.message)}go(cur);
}

// simple lists: grades, sections, teachers, announcements, excuses
const CR={
 grades:{t:'Grade Levels',f:[['name','Grade name']],c:['name']},
 sections:{t:'Sections',f:[['yearLevel','Grade level','grade'],['name','Section name'],['adviser','Adviser']],c:['yearLevel','name','adviser']},
 teachers:{t:'Teachers',f:[['name','Full name'],['email','Email'],['phone','Phone'],['subject','Subject'],['username','Login username']],c:['name','subject','email','phone','username']},
 announcements:{t:'Announcements',f:[['title','Title'],['message','Message','area']],c:['title','message','at']},
 excuses:{t:'Excuse / Absence Records',f:[['studentId','Student ID'],['name','Student name'],['date','Date','date'],['reason','Reason']],c:['date','studentId','name','reason']}
};
Object.keys(CR).forEach(k=>V[k]=async()=>{
  const d=CR[k],rows=CACHE[k]=await api('/api/'+k);
  const inp=([n,l,t])=>\`<label>\${l}\${t==='grade'?\`<select id="f_\${n}">\${gradeOpts()}</select>\`:t==='area'?\`<textarea id="f_\${n}"></textarea>\`:\`<input id="f_\${n}" type="\${t||'text'}">\`}</label>\`;
  return h(d.t)+\`<form class="card grid" onsubmit="return saveC('\${k}',event)"><input type="hidden" id="f__id">\${d.f.map(inp).join('')}<button>Save</button></form>\`
   +table([...d.c.map(c=>[c,r=>esc(fmt(r[c]))]),['',r=>\`<button class="sm" onclick="editC('\${k}','\${r._id}')">Edit</button><button class="sm danger" onclick="delC('\${k}','\${r._id}')">Delete</button>\`]],rows);
});
async function saveC(k,e){e.preventDefault();const b={_id:$('#f__id').value};CR[k].f.forEach(([n])=>b[n]=$('#f_'+n).value);await post('/api/'+k,b);if(k==='grades')GR=await api('/api/grades');go(cur);}
function editC(k,id){const r=CACHE[k].find(x=>x._id===id);$('#f__id').value=id;CR[k].f.forEach(([n])=>$('#f_'+n).value=r[n]||'');scrollTo(0,0)}
async function delC(k,id){if(confirm('Delete this record?')){await api('/api/'+k+'/'+id,{method:'DELETE'});if(k==='grades')GR=await api('/api/grades');go(cur)}}

// read-only lists
V.unauthorized=async()=>h('Unauthorized RFID','Unknown, blocked or inactive cards that were tapped')+table([['Card UID',r=>\`<code>\${esc(r.uid)}</code>\`],['Reason',r=>esc(r.reason)],['Taps',r=>r.count],['Last seen',r=>fmt(r.lastSeen)],['',r=>\`<button class="sm danger" onclick="delC('unauthorized','\${r._id}')">Dismiss</button>\`]],await api('/api/unauthorized'));
V.notifications=async()=>h('Notifications','Late students, early outs and card alerts')+table([['When',r=>fmt(r.at)],['Message',r=>\`<i class="b \${r.level==='warning'?'amber':'blue'}">\${r.level||'info'}</i> \${esc(r.message)}\`]],await api('/api/notifications'));
V.logs=async()=>h('Activity Logs','What administrators did')+table([['When',r=>fmt(r.at)],['User',r=>esc(r.user)],['Action',r=>esc(r.action)]],await api('/api/logs'));
V.esp=async()=>{const e=await api('/api/esp');return h('ESP8266 Status','Is the RFID scanner connected?')+\`<div class="card"><h3><i class="b \${e.connected?'green':'red'}">\${e.connected?'CONNECTED':'OFFLINE'}</i></h3><p>Last signal: \${e.lastPing?fmt(e.lastPing):'never'}</p><p>Last card UID: <code>\${esc(e.latestUid||'none')}</code></p><p class="sub">The scanner counts as connected if it sent a scan or <code>POST /api/ping</code> in the last 60 seconds.</p></div>\`};

// reports, search, export (shared filter bar)
const filterBar=()=>\`<div class="bar"><select id="pre" onchange="preset()"><option value="0">Daily</option><option value="6">Weekly</option><option value="29">Monthly</option><option value="179">Semester</option></select><input type="date" id="from" value="\${dayAgo(0)}"><input type="date" id="to" value="\${dayAgo(0)}"><select id="fg"><option value="">All grades</option>\${gradeOpts()}</select><input id="fs" placeholder="Section"><input id="fq" placeholder="Student name or ID"><button onclick="runF()">Search</button><button class="alt" onclick="dl('excel')">Excel</button><button class="alt" onclick="dl('csv')">CSV</button></div><div id="res"></div>\`;
const fq=()=>new URLSearchParams({from:$('#from').value,to:$('#to').value,grade:$('#fg').value,section:$('#fs').value,q:$('#fq').value});
const preset=()=>{$('#from').value=dayAgo(+$('#pre').value);$('#to').value=dayAgo(0)};
const runF=async()=>{$('#res').innerHTML=await attTable(fq(),[['Date',r=>r.date],...ATT])};
const dl=f=>location='/api/export-excel?format='+f+'&'+fq();
V.reports=async()=>h('Attendance Reports','Daily, weekly, monthly and semester reports')+filterBar();
V.search=async()=>h('Search Records','Find attendance by student, section, grade or date')+filterBar();
V.export=async()=>h('Export Reports','Download attendance as Excel or CSV')+filterBar();
V.analytics=async()=>{
  const a=await api('/api/analytics'),avg=Math.round(a.days.reduce((s,d)=>s+d.present,0)/Math.max(a.total*a.days.length,1)*100);
  return h('Attendance Analytics','Last 14 days · green = present, red = late')+\`<div class="stats"><div class="stat g"><b>\${avg}%</b><span>Average attendance</span></div></div><div class="card"><div class="bars">\${a.days.map(d=>{const p=a.total?Math.round(d.present/a.total*100):0;return \`<div class="col">\${p}%<div style="height:\${p*1.7}px"><i style="height:\${d.present?d.late/d.present*100:0}%"></i></div>\${d.date.slice(5)}</div>\`}).join('')}</div></div>\`;
};

// settings
const frm=(id,body,fn)=>\`<form class="card grid" onsubmit="return \${fn}(event)">\${body}<button>Save</button></form>\`;
const cfgPost=async(e,keys)=>{e.preventDefault();const b={};keys.forEach(k=>b[k]=$('#c_'+k.replace('.','_')).value);await post('/api/config',b);alert('Saved');return false};
const cf=(c,k,l,t='text')=>{const v=k.split('.').reduce((o,p)=>o?.[p],c)||'';return \`<label>\${l}<input id="c_\${k.replace('.','_')}" type="\${t}" value="\${esc(v)}"></label>\`};
V.schedule=async()=>{const c=await api('/api/config'),k=['sched.opening','sched.late','sched.noon','sched.dismissal'];window.sk=k;
  return h('Attendance Schedule','Used to decide late, AM/PM and early out')+frm('',cf(c,k[0],'School opening','time')+cf(c,k[1],'Late after','time')+cf(c,k[2],'Morning ends / afternoon starts','time')+cf(c,k[3],'Dismissal (leaving earlier = early out)','time'),'saveSched')};
const saveSched=e=>cfgPost(e,window.sk);
V.school=async()=>{const c=await api('/api/config'),k=['school.name','school.address','school.contact','systemName'];window.sk2=k;
  return h('School Information','Name, logo, address and contact')+frm('',cf(c,k[0],'School name')+cf(c,k[1],'Address')+cf(c,k[2],'Contact')+cf(c,k[3],'System name'),'saveSchool')
   +\`<div class="card"><label>School logo<input type="file" id="logo" accept="image/*"></label><p><button class="alt" onclick="upLogo()">Upload logo</button> \${c.logoPath?\`<button class="danger" onclick="rmLogo()">Remove logo</button> <img src="\${c.logoPath}" height="40">\`:''}</p></div>\`};
const saveSchool=e=>cfgPost(e,window.sk2);
async function upLogo(){const f=new FormData();f.append('logoFile',$('#logo').files[0]);await api('/api/upload-logo',{method:'POST',body:f});init()}
async function rmLogo(){await post('/api/remove-logo',{});init()}
V.admin=async()=>h('Admin Account','Change the administrator username or password')+frm('',\`<label>Username<input id="a_u" value="admin"></label><label>Current password<input id="a_o" type="password" required></label><label>New password<input id="a_n" type="password" placeholder="Leave blank to keep"></label>\`,'saveAdmin');
async function saveAdmin(e){e.preventDefault();try{await post('/api/admin',{username:$('#a_u').value,oldPass:$('#a_o').value,newPass:$('#a_n').value});alert('Account updated')}catch(x){alert(x.message)}return false}
V.backup=async()=>h('Backup & Restore','Save or reload students and attendance data')+\`<div class="card"><p><a href="/api/backup"><button>Download backup</button></a></p><label>Restore from backup file<input type="file" id="rf" accept=".json"></label><p><button class="alt" onclick="restore()">Restore data</button> <button class="danger" onclick="clearLogs()">Clear attendance logs</button></p></div>\`;
async function restore(){const f=$('#rf').files[0];if(!f||!confirm('This replaces current data. Continue?'))return;await post('/api/restore',JSON.parse(await f.text()));alert('Restored')}
async function clearLogs(){if(confirm('Clear ALL attendance logs?')){await post('/api/clear-logs',{});alert('Cleared')}}

// router
const LIVE=['dashboard','daily','late','absent','early','timeinout','unauthorized','esp','notifications'];
async function go(k){
  cur=k||'dashboard';
  $('#nav a.on')?.classList.remove('on');$('#nav a[data-k="'+cur+'"]')?.classList.add('on');
  try{$('#view').innerHTML=await V[cur]();if(cur==='reports'||cur==='search'||cur==='export')runF()}catch(e){$('#view').innerHTML='<p class="empty">'+esc(e.message)+'</p>'}
}
async function init(){
  GR=await api('/api/grades');const c=await api('/api/config');
  $('#nav').innerHTML=\`<div class="brand">\${c.logoPath?\`<img src="\${c.logoPath}">\`:'🎓'}<span>\${esc(c.school.name)}</span></div>\`+MENU.map(m=>\`<a data-k="\${m[0]}" href="\${m[3]||'#'+m[0]}" \${m[0]==='scanner'?'target="_blank"':''}>\${m[1]} \${m[2]}</a>\`).join('');
  go(cur);
}
addEventListener('hashchange',()=>go(location.hash.slice(1)));
setInterval(()=>{if(LIVE.includes(cur)&&!document.hidden)go(cur)},8000);
cur=location.hash.slice(1)||'dashboard';init();
</script>
</body>
</html>
`;

const SCANNER_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>RFID Scanner</title>
<link href="https://fonts.googleapis.com/css2?family=Figtree:wght@400;600;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box}
body{margin:0;height:100vh;display:flex;flex-direction:column;background:#0d1b2a;color:#fff;font-family:Figtree,'Segoe UI',sans-serif;overflow:hidden}
header{display:flex;align-items:center;gap:16px;padding:16px 32px;background:#12263a}
header img{height:56px;border-radius:8px}header h1{margin:0;font-size:26px;flex:1}
#clock{font-size:42px;font-weight:800;text-align:right;line-height:1}#clock small{display:block;font-size:15px;font-weight:400;color:#9fb3c8}
.wrap{flex:1;display:flex;gap:24px;padding:24px 32px;min-height:0}
#stage{flex:1;border-radius:20px;background:#12263a;display:flex;align-items:center;justify-content:center;gap:48px;padding:32px;transition:background .3s}
#stage.ok{background:#14532d}#stage.late{background:#7c4a03}#stage.bad{background:#7f1d1d}
#idle{text-align:center;font-size:40px;font-weight:800}#idle p{font-size:20px;font-weight:400;color:#9fb3c8}
.ring{width:150px;height:150px;border:8px solid #2f6fdb;border-radius:50%;margin:0 auto 24px;animation:pulse 2s infinite}
@keyframes pulse{50%{transform:scale(1.08);opacity:.5}}@media(prefers-reduced-motion:reduce){.ring{animation:none}}
#card{display:none;align-items:center;gap:48px;width:100%}
#photo{width:min(34vw,380px);aspect-ratio:3/4;border-radius:18px;object-fit:cover;background:#1d3a55;font-size:120px;display:flex;align-items:center;justify-content:center}
#info{flex:1;min-width:0}#name{font-size:clamp(34px,5vw,64px);font-weight:800;line-height:1.1}
#grade{font-size:30px;color:#cfe3f5;margin:10px 0}#id{font-size:20px;color:#9fb3c8}
#type{display:inline-block;margin-top:22px;font-size:34px;font-weight:800;padding:10px 26px;border-radius:12px;background:#ffffff26}
#time{font-size:56px;font-weight:800;margin-top:14px}
aside{width:320px;display:flex;flex-direction:column;gap:16px}
.box{background:#12263a;border-radius:16px;padding:16px}.box h3{margin:0 0 10px;font-size:15px;color:#9fb3c8;font-weight:600}
.cnt{display:flex;gap:10px}.cnt div{flex:1;font-size:13px;color:#9fb3c8}.cnt b{display:block;font-size:34px;color:#fff}
#recent div{padding:7px 0;border-bottom:1px solid #1d3a55;font-size:14px;display:flex;justify-content:space-between;gap:8px}
#ann{padding:12px 32px;background:#12263a;font-size:18px;color:#cfe3f5;min-height:48px}
#start{position:fixed;right:16px;bottom:64px;z-index:9;display:flex;align-items:center;gap:10px;background:#12263a;border:2px solid #1f8a5b;padding:10px 16px;border-radius:12px;font-size:16px}
#start button{font:inherit;font-weight:800;padding:8px 18px;border:0;border-radius:8px;background:#1f8a5b;color:#fff;cursor:pointer}
#net{position:fixed;left:16px;bottom:64px;font-size:13px;color:#9fb3c8}
@media(max-width:900px){aside{display:none}#card{flex-direction:column;gap:20px}#photo{width:50vw}}
</style>
</head>
<body>
<div id="start">🔊 Voice is off <button onclick="begin()">Turn on voice</button></div><div id="net">Connecting…</div>
<header><img id="logo" hidden><h1 id="school">School Attendance</h1><div id="clock"></div></header>
<div class="wrap">
  <div id="stage">
    <div id="idle"><div class="ring"></div>Please tap your ID<p>Hold your RFID card near the reader</p></div>
    <div id="card"><div id="photo"></div><div id="info"><div id="name"></div><div id="grade"></div><div id="id"></div><div id="type"></div><div id="time"></div></div></div>
  </div>
  <aside>
    <div class="box"><h3>Today</h3><div class="cnt"><div><b id="cp">0</b>Present</div><div><b id="cl">0</b>Late</div><div><b id="ci">0</b>Inside</div></div></div>
    <div class="box"><h3>Test without the reader</h3><input id="tu" placeholder="Type a card UID" style="width:100%;padding:8px;border-radius:6px;border:0;margin-bottom:8px"><button onclick="testScan()" style="width:100%;padding:8px;border:0;border-radius:6px;background:#2f6fdb;color:#fff;font-weight:600;cursor:pointer">Scan</button></div>
    <div class="box" style="flex:1;overflow:hidden"><h3>Recent scans</h3><div id="recent"></div></div>
  </aside>
</div>
<div id="ann"></div>
<script>
const $=s=>document.querySelector(s);
let seen=null,voice=false,timer,recent=[],anns=[],ai=0;
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function begin(){if(voice)return;voice=true;$('#start')?.remove();speak('Voice is on')}
addEventListener('click',begin);addEventListener('keydown',begin);
function speak(t){if(!voice||!window.speechSynthesis)return;speechSynthesis.cancel();setTimeout(()=>{const u=new SpeechSynthesisUtterance(t);u.rate=.95;speechSynthesis.speak(u)},150)}
if(window.speechSynthesis)speechSynthesis.getVoices();
async function testScan(){const u=$('#tu').value.trim();if(u)await fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({uid:u})})}
setInterval(()=>{const d=new Date();$('#clock').innerHTML=d.toLocaleTimeString('en-US',{timeZone:'Asia/Manila',hour:'2-digit',minute:'2-digit',second:'2-digit'})+'<small>'+d.toLocaleDateString('en-US',{timeZone:'Asia/Manila',weekday:'long',month:'long',day:'numeric'})+'</small>'},1000);

function show(s){
  clearTimeout(timer);
  const stage=$('#stage'),done=s.type==='DONE',late=s.status==='LATE'||s.status==='EARLY OUT';
  stage.className=!s.ok?'bad':late?'late':'ok';
  $('#idle').style.display='none';$('#card').style.display='flex';
  const p=$('#photo');
  if(s.photo){p.style.backgroundImage='url('+s.photo+')';p.style.backgroundSize='cover';p.textContent=''}else{p.style.backgroundImage='';p.textContent=s.ok?(s.name||'?')[0]:'🚫'}
  $('#name').textContent=s.ok?s.name:s.message;
  $('#grade').textContent=s.ok?s.yearLevel+' - '+s.section:'Card: '+s.uid;
  $('#id').textContent=s.ok?'ID '+s.studentId:'Please see the office';
  $('#type').textContent=!s.ok?'ACCESS DENIED':done?'ALREADY COMPLETE':s.type+' · '+s.period+(late?' · '+s.status:'');
  $('#time').textContent=s.time||'';
  const greet=s.period==='PM'?'Good afternoon':'Good morning';
  speak(!s.ok?'Sorry, '+s.message.toLowerCase()+'. Please see the office.':done?s.name+', you already completed your '+s.period+' scans.':greet+', '+s.name+'. '+(s.type==='TIME-IN'?'Time in':'Time out')+' recorded.'+(s.status==='LATE'?' You are late.':s.status==='EARLY OUT'?' Early out recorded.':''));
  recent.unshift(\`<div><span>\${esc(s.ok?s.name:s.message)}</span><span>\${esc(s.ok?(s.type||'')+' '+s.time:'denied')}</span></div>\`);recent=recent.slice(0,8);$('#recent').innerHTML=recent.join('');
  timer=setTimeout(()=>{stage.className='';$('#card').style.display='none';$('#idle').style.display='block'},7000);
  stats();
}
async function poll(){
  try{const s=await (await fetch('/api/last-scan')).json();if(s.status===401||s.error)return location='/login';
    $('#net').textContent='● Connected · last card received: '+(s.uid||'none yet');
    if(seen===null)seen=s.id;else if(s.id!==seen){seen=s.id;show(s)}}catch(e){$('#net').textContent='○ Cannot reach the server'}
}
async function stats(){try{const s=await (await fetch('/api/stats')).json();$('#cp').textContent=s.present;$('#cl').textContent=s.late;$('#ci').textContent=s.inside}catch(e){}}
async function boot(){
  try{const c=await (await fetch('/api/config')).json();$('#school').textContent=c.school.name;if(c.logoPath){$('#logo').src=c.logoPath;$('#logo').hidden=false}anns=await (await fetch('/api/announcements')).json()}catch(e){}
  stats();setInterval(stats,15000);setInterval(poll,800);
  setInterval(()=>{if(anns.length){const a=anns[ai++%anns.length];$('#ann').textContent='📢 '+a.title+': '+a.message}},6000);
}
boot();
</script>
</body>
</html>
`;

// UPLOADS (SCHOOL LOGO + STUDENT PHOTOS)
const uploadsDir = path.join(__dirname, 'uploads');
fs.mkdirSync(uploadsDir, { recursive: true });
const mk = name => multer({
  storage: multer.diskStorage({ destination: uploadsDir, filename: (req, f, cb) => cb(null, name() + path.extname(f.originalname)) }),
  fileFilter: (req, f, cb) => f.mimetype.startsWith('image/') ? cb(null, true) : cb(new Error('Only image files are allowed!'), false)
});
const upload = mk(() => 'school_logo');
const photoUpload = multer({
  storage: multer.memoryStorage(), limits: { fileSize: 1e6 },
  fileFilter: (req, f, cb) => f.mimetype.startsWith('image/') ? cb(null, true) : cb(new Error('Only image files are allowed!'), false)
});

// MIDDLEWARES + LOGIN (cookie session)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));
app.get('/assets/style.css', (req, res) => res.type('css').send(CSS));

const sessions = new Set();
const cookie = req => Object.fromEntries((req.headers.cookie || '').split(';').map(c => c.trim().split('=')));
const open = ['/login', '/api/scan', '/api/ping', '/student-register', '/api/register-student'];
app.use((req, res, next) => {
  if (open.includes(req.path) || req.path.startsWith('/uploads') || req.path.startsWith('/assets') || sessions.has(cookie(req).sid)) return next();
  req.path.startsWith('/api') ? res.status(401).json({ error: 'Login required' }) : res.redirect('/login');
});

app.get('/login', async (req, res) => {
  const c = await getConfig();
  res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Login</title>
  <link rel="stylesheet" href="/assets/style.css"></head><body class="login"><form method="POST" action="/login" class="box">
  ${c.logoPath ? `<img src="${c.logoPath}" height="70">` : ''}<h2>${c.school.name}</h2><p>Attendance system admin</p>
  ${req.query.e ? '<p class="err">Wrong username or password.</p>' : ''}
  <input name="username" placeholder="Username" required><input name="password" type="password" placeholder="Password" required>
  <button>Log in</button></form></body></html>`);
});
app.post('/login', async (req, res) => {
  const c = await getConfig();
  if (req.body.username === c.adminUser && sha(req.body.password) === c.adminPass) {
    const sid = crypto.randomBytes(24).toString('hex');
    sessions.add(sid);
    res.setHeader('Set-Cookie', `sid=${sid}; HttpOnly; Path=/; Max-Age=86400`);
    log('Logged in');
    return res.redirect('/');
  }
  res.redirect('/login?e=1');
});
app.get('/logout', (req, res) => {
  sessions.delete(cookie(req).sid);
  res.setHeader('Set-Cookie', 'sid=; Path=/; Max-Age=0');
  res.redirect('/login');
});

// NOTIFICATIONS (EMAIL & SMS)
const resend = new Resend(process.env.RESEND_API_KEY || 'YOUR_RESEND_API_KEY');

async function sendEmailNotification(config, to, studentName, label, status, timestamp) {
  if (!to) return;
  try {
    await resend.emails.send({
      from: 'School Attendance <onboarding@resend.dev>', to,
      subject: `[${label}] Attendance Alert: ${studentName}`,
      html: `<div style="font-family:Arial;padding:15px;border:1px solid #ddd;border-radius:6px;">
        <h2>${config.school.name}</h2>
        <p><strong>${studentName}</strong> logged <strong>${label}</strong> at ${timestamp}.</p>
        <p>Status: <strong style="color:${status === 'ON TIME' || status === 'COMPLETED' ? '#2ecc71' : '#e74c3c'}">${status}</strong></p></div>`
    });
  } catch (error) { console.error('[EMAIL ERROR]', error.message); }
}

function sendSMSNotification(config, phoneNumber, studentName, label, status, timestamp) {
  if (!config.enableSms || !config.semaphoreApiKey || !phoneNumber) return;
  const postData = querystring.stringify({ apikey: config.semaphoreApiKey, number: phoneNumber, message: `[${config.school.name}] ${studentName}: ${label} at ${timestamp}. Status: ${status}.` });
  const req = https.request({
    hostname: 'api.semaphore.co', port: 443, path: '/api/v4/messages', method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': postData.length }
  }, r => r.on('data', d => console.log('[SMS RESPONSE]', d.toString())));
  req.on('error', e => console.error('[SMS ERROR]', e.message));
  req.write(postData); req.end();
}

// API: ESP8266 SCANNER ENDPOINT (AM in/out + PM in/out)
let lastScan = { id: 0 };
const SLOTS = { am: ['amIn', 'amOut'], pm: ['pmIn', 'pmOut'] };

app.get('/api/scan', (req, res) => res.send('Scan endpoint is online. The ESP8266 must POST uid to this address.'));
app.post('/api/ping', async (req, res) => { await Config.updateOne({}, { lastPing: new Date() }); res.json({ status: 'ok' }); });

app.post('/api/scan', async (req, res) => {
  try {
    const { uid } = req.body;
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID' });
    const cleanUid = uid.trim().toUpperCase();
    console.log('[SCAN] received UID:', cleanUid);
    const config = await getConfig();
    config.latestUid = cleanUid; config.lastPing = new Date();
    await config.save();

    const now = new Date(), date = ymd(now), time = hm(now);
    const show = o => { lastScan = { id: lastScan.id + 1, ts: Date.now(), time, uid: cleanUid, ...o }; };
    if (lastScan.uid === cleanUid && Date.now() - lastScan.ts < 4000) return res.json({ status: 'duplicate', message: 'Already scanned' });

    const compact = cleanUid.replace(/[\s:-]/g, '');
    const student = await Student.findOne({ uid: { $in: [cleanUid, compact] } }).select('-photoData');
    if (!student || student.cardStatus !== 'active') {
      const reason = student ? `Card ${student.cardStatus}` : 'Card not registered';
      await Unauthorized.findOneAndUpdate({ uid: cleanUid }, { $inc: { count: 1 }, reason, lastSeen: now }, { upsert: true });
      notify(`${reason}: ${cleanUid}`, 'warning');
      show({ ok: false, message: reason });
      return res.json({ status: student ? 'blocked' : 'unknown', message: reason });
    }

    const info = { ok: true, studentId: student.studentId, name: student.name, yearLevel: student.yearLevel, section: student.section, photo: student.photo };
    const rec = await Daily.findOne({ date, studentId: student.studentId }) || new Daily({ date, uid: cleanUid, ...info });
    const sc = { opening: '06:00', late: '07:30', noon: '12:00', dismissal: '16:00', ...JSON.parse(JSON.stringify(config.sched || {})) };
    const period = time >= sc.noon ? 'pm' : 'am';
    const slot = SLOTS[period].find(k => !rec[k]);
    if (!slot) {
      show({ ...info, type: 'DONE', period: period.toUpperCase(), status: 'Already completed' });
      return res.json({ status: 'duplicate', message: `${student.name} already completed ${period.toUpperCase()} scans` });
    }

    rec[slot] = time;
    const type = slot.endsWith('In') ? 'TIME-IN' : 'TIME-OUT';
    let status = type === 'TIME-IN' ? 'ON TIME' : 'COMPLETED';
    if (slot === 'amIn' && time > sc.late) { rec.late = true; status = 'LATE'; }
    if (slot === 'pmOut' && time < sc.dismissal) { rec.earlyOut = true; status = 'EARLY OUT'; }
    await rec.save();
    show({ ...info, type, period: period.toUpperCase(), status });
    try {
      await ScanLog.create({ date, time, studentId: student.studentId, name: student.name, uid: cleanUid, action: type, period: period.toUpperCase(), status });
      if (status === 'LATE' || status === 'EARLY OUT') notify(`${student.name} (${student.yearLevel}-${student.section}) ${status} at ${time}`, 'warning');
    } catch (e) { console.error('[SCANLOG ERROR]', e.message); }

    const label = `${type} (${period.toUpperCase()})`;
    if (config.enableEmail) sendEmailNotification(config, student.email, student.name, label, status, time);
    sendSMSNotification(config, student.phone, student.name, label, status, time);
    res.json({ status: 'success', scanType: type, isLate: status === 'LATE', message: `${type} recorded for ${student.name}` });
  } catch (err) {
    console.error('[SCAN ERROR]', err);
    lastScan = { id: lastScan.id + 1, ts: Date.now(), time: hm(new Date()), uid: req.body.uid || '', ok: false, message: 'Server error: ' + err.message };
    res.status(500).json({ status: 'error', message: 'Server Error: ' + err.message });
  }
});

app.get('/api/last-scan', (req, res) => res.json(lastScan));
app.get('/api/esp', async (req, res) => {
  const c = await getConfig();
  res.json({ connected: !!c.lastPing && Date.now() - c.lastPing < 60000, lastPing: c.lastPing, latestUid: c.latestUid });
});

// ATTENDANCE QUERIES (daily, late, absent, early out, reports, search)
async function attendance({ from, to, q, grade, section, filter }) {
  const today = ymd(new Date()), day = from || today;
  const sf = {};
  if (grade) sf.yearLevel = grade;
  if (section) sf.section = section;
  if (q) { const rx = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'); sf.$or = [{ name: rx }, { studentId: rx }]; }
  if (filter === 'absent') {
    const [students, recs, ex] = await Promise.all([Student.find(sf).select('-photoData'), Daily.find({ date: day }), Excuse.find({ date: day })]);
    const skip = new Set([...recs.map(r => r.studentId), ...ex.map(e => e.studentId)]);
    return students.filter(s => !skip.has(s.studentId)).map(s => ({ date: day, name: s.name, studentId: s.studentId, yearLevel: s.yearLevel, section: s.section, photo: s.photo, absent: true }));
  }
  const f = { date: { $gte: day, $lte: to || day }, ...sf };
  if (filter === 'late') f.late = true;
  if (filter === 'early') f.earlyOut = true;
  return Daily.find(f).sort({ date: -1, _id: -1 }).limit(3000);
}
app.get('/api/attendance', async (req, res) => res.json(await attendance(req.query)));

app.get('/api/stats', async (req, res) => {
  const date = ymd(new Date());
  const [total, recs, excused] = await Promise.all([Student.countDocuments(), Daily.find({ date }), Excuse.countDocuments({ date })]);
  const inside = recs.filter(r => { const last = ['pmOut', 'pmIn', 'amOut', 'amIn'].find(k => r[k]); return last && last.endsWith('In'); }).length;
  res.json({ total, present: recs.length, absent: Math.max(total - recs.length - excused, 0), late: recs.filter(r => r.late).length, early: recs.filter(r => r.earlyOut).length, inside });
});

app.get('/api/analytics', async (req, res) => {
  const days = [...Array(14)].map((_, i) => ymd(new Date(Date.now() - i * 864e5))).reverse();
  const [total, recs] = await Promise.all([Student.countDocuments(), Daily.find({ date: { $in: days } })]);
  res.json({ total, days: days.map(d => { const r = recs.filter(x => x.date === d); return { date: d, present: r.length, late: r.filter(x => x.late).length }; }) });
});

// EXPORT TO EXCEL / CSV
app.get('/api/export-excel', async (req, res) => {
  try {
    const c = await getConfig();
    const rows = await attendance(req.query);
    const cols = ['Date', 'ID Number', 'Student Name', 'Grade & Section', 'AM In', 'AM Out', 'PM In', 'PM Out', 'Status'];
    const status = r => r.absent ? 'ABSENT' : [r.late && 'LATE', r.earlyOut && 'EARLY OUT'].filter(Boolean).join(', ') || 'ON TIME';
    const data = rows.map(r => [r.date, r.studentId, r.name, `${r.yearLevel} - ${r.section}`, r.amIn || '', r.amOut || '', r.pmIn || '', r.pmOut || '', status(r)]);
    if (req.query.format === 'csv') {
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename="Attendance_Report.csv"');
      return res.send([cols, ...data].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n'));
    }
    const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Attendance Log');
    ws.addRow([c.school.name]).font = { bold: true, size: 16 };
    ws.addRow(['OFFICIAL ATTENDANCE REPORT']); ws.addRow([]);
    const h = ws.addRow(cols);
    h.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    h.eachCell(x => { x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B365D' } }; });
    data.forEach(r => ws.addRow(r));
    ws.columns.forEach(x => { x.width = 18; });
    [['bnhs_logo.png', 0], ['bnhs_logo.jpg', 0], ['src_logo.png', 8]].forEach(([f, col]) => {
      const p = path.join(__dirname, f);
      if (fs.existsSync(p)) ws.addImage(wb.addImage({ filename: p, extension: f.endsWith('png') ? 'png' : 'jpeg' }), { tl: { col, row: 0 }, ext: { width: 60, height: 60 } });
    });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Attendance_Report.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).send('Error generating Excel file: ' + err.message);
  }
});

// GENERIC LISTS: grades, sections, teachers, announcements, excuses, notifications, logs, unauthorized, scan logs
const lists = { grades: D.Grade, sections: D.Section, teachers: D.Teacher, announcements: D.Announcement, excuses: Excuse, notifications: D.Notification, logs: D.ActivityLog, unauthorized: Unauthorized, scanlogs: ScanLog };
Object.entries(lists).forEach(([k, M]) => {
  app.get('/api/' + k, async (req, res) => res.json(await M.find(req.query.date ? { date: req.query.date } : {}).sort({ _id: -1 }).limit(500)));
  app.post('/api/' + k, async (req, res) => {
    const { _id, ...b } = req.body;
    _id ? await M.findByIdAndUpdate(_id, b) : await M.create(b);
    log(`Saved ${k}`); res.json({ ok: true });
  });
  app.delete('/api/' + k + '/:id', async (req, res) => { await M.findByIdAndDelete(req.params.id); log(`Deleted from ${k}`); res.json({ ok: true }); });
});

// STUDENTS + RFID CARDS
app.get('/api/students', async (req, res) => res.json(await Student.find().select('-photoData').sort({ name: 1 })));
app.get('/photo/:id', async (req, res) => {
  const s = await Student.findById(req.params.id).select('photoData photoType').catch(() => null);
  if (!s || !s.photoData) return res.sendStatus(404);
  res.type(s.photoType).send(s.photoData);
});

app.post('/api/register', photoUpload.single('photo'), async (req, res) => {
  try {
    const { mongoId, uid, name, studentId, yearLevel, section, email, phone } = req.body;
    const cleanUid = (uid || '').trim().toUpperCase();
    if (cleanUid && await Student.findOne({ uid: cleanUid, _id: { $ne: mongoId || undefined } })) return res.status(409).json({ error: 'That RFID card is already assigned' });
    const d = { uid: cleanUid, name, studentId, yearLevel: yearLevel || 'Grade 7', section: section || 'A', email: email || '', phone: phone || '' };
    const st = mongoId ? await Student.findByIdAndUpdate(mongoId, d) : await Student.create(d);
    if (req.file) await Student.findByIdAndUpdate(st?._id || mongoId, { photoData: req.file.buffer, photoType: req.file.mimetype, photo: `/photo/${st?._id || mongoId}?v=${Date.now()}` });
    log(`Saved student ${name}`); res.json({ ok: true });
  } catch (err) { res.status(400).json({ error: err.message }); }
});

app.post('/api/delete-student', async (req, res) => { await Student.findByIdAndDelete(req.body.id); log('Deleted a student'); res.json({ ok: true }); });

app.post('/api/card', async (req, res) => {
  const { id, uid, cardStatus } = req.body, u = {};
  if (uid !== undefined) u.uid = uid.trim().toUpperCase();
  if (cardStatus) u.cardStatus = cardStatus;
  if (u.uid && await Student.findOne({ uid: u.uid, _id: { $ne: id } })) return res.status(409).json({ error: 'That RFID card is already assigned' });
  await Student.findByIdAndUpdate(id, u);
  log('RFID card updated'); res.json({ ok: true });
});

// SETTINGS: schedule, school info, notifications, logo, admin account
app.get('/api/config', async (req, res) => { const c = (await getConfig()).toObject(); delete c.adminPass; res.json(c); });
app.post('/api/config', async (req, res) => {
  const ok = k => /^(sched|school)\./.test(k) || ['systemName', 'enableEmail', 'enableSms', 'semaphoreApiKey'].includes(k);
  await getConfig();
  await Config.updateOne({}, { $set: Object.fromEntries(Object.entries(req.body).filter(([k]) => ok(k))) });
  log('Settings updated'); res.json({ ok: true });
});
app.post('/api/upload-logo', upload.single('logoFile'), async (req, res) => {
  if (req.file) await Config.updateOne({}, { logoPath: `/uploads/${req.file.filename}?v=${Date.now()}` });
  res.json({ ok: true });
});
app.post('/api/remove-logo', async (req, res) => { await Config.updateOne({}, { logoPath: '' }); res.json({ ok: true }); });
app.post('/api/admin', async (req, res) => {
  const c = await getConfig(), { oldPass, username, newPass } = req.body;
  if (sha(oldPass) !== c.adminPass) return res.status(403).json({ error: 'Current password is wrong' });
  c.adminUser = username || c.adminUser;
  if (newPass) c.adminPass = sha(newPass);
  await c.save(); log('Admin account changed'); res.json({ ok: true });
});

// BACKUP & RESTORE
const backupModels = { students: Student, daily: Daily, scanlogs: ScanLog, grades: D.Grade, sections: D.Section, teachers: D.Teacher, announcements: D.Announcement, excuses: Excuse };
app.get('/api/backup', async (req, res) => {
  const out = {};
  for (const [k, M] of Object.entries(backupModels)) out[k] = await M.find().lean();
  res.setHeader('Content-Disposition', `attachment; filename="backup_${ymd(new Date())}.json"`);
  res.json(out);
});
app.post('/api/restore', async (req, res) => {
  for (const [k, M] of Object.entries(backupModels)) {
    if (!Array.isArray(req.body[k])) continue;
    await M.deleteMany({});
    if (req.body[k].length) await M.insertMany(req.body[k]);
  }
  log('Backup restored'); res.json({ ok: true });
});
app.post('/api/clear-logs', async (req, res) => { await Daily.deleteMany({}); await ScanLog.deleteMany({}); log('Cleared attendance logs'); res.json({ ok: true }); });

// PUBLIC STUDENT REGISTRATION FORM
app.get('/student-register', async (req, res) => {
  const grades = (await D.Grade.find()).map(g => `<option>${g.name}</option>`).join('');
  res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Student Registration</title>
  <link rel="stylesheet" href="/assets/style.css"></head><body class="login"><form method="POST" action="/api/register-student" class="box">
  <h2>Student Registration</h2><input name="studentId" placeholder="ID Number (e.g. 2026-1001)" required><input name="name" placeholder="Full Name" required>
  <select name="yearLevel">${grades}</select><input name="section" placeholder="Section (e.g. Diamond)" required>
  <input type="email" name="email" placeholder="Parent / guardian email" required><input type="tel" name="phone" placeholder="Phone (optional)">
  <button>Submit registration</button></form></body></html>`);
});

app.post('/api/register-student', async (req, res) => {
  try {
    const { name, email, studentId, phone, yearLevel, section } = req.body;
    if (await Student.findOne({ email })) return res.send('<h2 style="text-align:center;font-family:Arial;color:#e74c3c">That email is already registered. <a href="/student-register">Back</a></h2>');
    await Student.create({ name, email, studentId, phone: phone || '', yearLevel, section, uid: '' });
    res.send(`<div style="text-align:center;padding:50px;font-family:Arial"><h2 style="color:#2ecc71">Registration Successful!</h2>
      <p>Thank you <strong>${name}</strong>! (${yearLevel} - ${section}). The admin will assign your RFID card.</p><a href="/student-register">Register another student</a></div>`);
  } catch (err) { res.status(500).send('Error: ' + err.message); }
});

// PAGES: admin dashboard ( / ) and separate scanning screen ( /scanner )
app.get('/', (req, res) => res.type('html').send(INDEX_HTML));
app.get('/scanner', (req, res) => res.type('html').send(SCANNER_HTML));

// START SERVER
app.listen(PORT, async () => {
  if (!await D.Grade.countDocuments()) await D.Grade.insertMany([7, 8, 9, 10, 11, 12].map(n => ({ name: 'Grade ' + n })));
  await getConfig();
  console.log(`Server running on port ${PORT}`);
  console.log('[SCANNER] ESP8266 must POST uid to http://<this-server>/api/scan');
});
