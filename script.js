const KEY="KONOHA_MOBILE_CHECKLIST_V1";
const DEFAULT_DAILY=["Check emails / messages","Send morning update","Check Jira / assigned tasks","Update task progress","Send evening update"];
const DEFAULT_WEEKLY=["WSR – Email","WSR – Message","Fill Timesheet"];

// IMPORTANT: paste the SAME deployed Apps Script /exec URL here and in laptop/script.js.
const SHEETS_WEB_APP_URL="https://script.google.com/macros/s/AKfycbxgwDZTs7S-FSSCQg9OF05cJx9yfzyXy4JtApqdXRXfTuDFILA2H8XH_WNx3U1qUAL_/exec";
const SHEETS_ENABLED=SHEETS_WEB_APP_URL.startsWith("https://script.google.com/");
const DEVICE_NAME="Mobile";

function dk(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}
function weekKeyForDate(key){const d=new Date(key+"T12:00:00"),day=d.getDay();d.setDate(d.getDate()+(day===0?-6:1-day));return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}
function wk(){return weekKeyForDate(dk())}
function dailyDefaults(){return DEFAULT_DAILY.map((text,i)=>({id:"d"+i,text,done:false}))}
function weeklyDefaults(){return DEFAULT_WEEKLY.map((text,i)=>({id:"w"+i,text,done:false}))}
function init(){let d;try{d=JSON.parse(localStorage.getItem(KEY)||"null")}catch(e){}if(!d)d={daily:dailyDefaults(),dailyDate:dk(),days:{},weeks:{},notesByDay:{}};d.days??={};d.weeks??={};d.notesByDay??={};if(d.dailyDate!==dk()){d.daily=dailyDefaults();d.dailyDate=dk()}if(!d.days[dk()])d.days[dk()]={tasks:[]};if(!d.weeks[wk()])d.weeks[wk()]={tasks:weeklyDefaults()};return d}
let data=init(),pushTimer=null,pullBusy=false;
function persist(){localStorage.setItem(KEY,JSON.stringify(data))}
function list(t){return t==="daily"?data.daily:t==="weekly"?data.weeks[wk()].tasks:data.days[dk()].tasks}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
function renderList(t,id){document.getElementById(id).innerHTML=list(t).map(x=>`<div class="task ${x.done?"done":""}"><input type="checkbox" ${x.done?"checked":""} onchange="toggle('${t}','${x.id}')"><input type="text" value="${esc(x.text)}" onchange="editTask('${t}','${x.id}',this.value)"><button class="delete" onclick="removeTask('${t}','${x.id}')">×</button></div>`).join("")}
function stat(t){const a=list(t),n=a.filter(x=>x.done).length;return[n,a.length,a.length?Math.round(n/a.length*100):0]}
function render(){const d=new Date();weekday.textContent=d.toLocaleDateString(undefined,{weekday:"long"}).toUpperCase();dateText.textContent=d.toLocaleDateString(undefined,{day:"numeric",month:"long",year:"numeric"});renderList("daily","dailyList");renderList("today","todayList");renderList("weekly","weeklyList");const a=stat("daily"),b=stat("today"),c=stat("weekly"),tot=a[1]+b[1]+c[1],done=a[0]+b[0]+c[0],p=tot?Math.round(done/tot*100):0;dailyStat.textContent=`${a[0]} / ${a[1]}`;todayStat.textContent=`${b[0]} / ${b[1]}`;weeklyStat.textContent=`${c[0]} / ${c[1]}`;overallPct.textContent=p+"%";document.querySelector(".ring").style.setProperty("--deg",p*3.6+"deg")}
function save(sync=true){persist();render();if(sync)queueLivePush()}
function toggle(t,id){const x=list(t).find(a=>a.id===id);if(x){x.done=!x.done;save()}}
function editTask(t,id,v){const x=list(t).find(a=>a.id===id);if(x){x.text=v.trim()||"Untitled task";save()}}
function removeTask(t,id){const a=list(t),i=a.findIndex(x=>x.id===id);if(i>-1){a.splice(i,1);save()}}
function addTask(t){const el=document.getElementById(t+"Input"),v=el.value.trim();if(!v)return;list(t).push({id:t[0]+Date.now()+Math.random().toString(36).slice(2,7),text:v,done:false});el.value="";save()}
function resetToday(){data.daily=dailyDefaults();data.dailyDate=dk();data.days[dk()]={tasks:[]};save()}
function resetWeek(){data.weeks[wk()]={tasks:weeklyDefaults()};save()}
function setSyncStatus(text,cls=""){syncStatus.textContent=text;syncStatus.className=cls}
function livePayload(){return {action:"state",date:dk(),week:wk(),device:DEVICE_NAME,notes:data.notesByDay?.[dk()]||"",tasks:{daily:data.daily,today:data.days[dk()]?.tasks||[],weekly:data.weeks[wk()]?.tasks||[]}}}
function queueLivePush(){clearTimeout(pushTimer);pushTimer=setTimeout(pushLiveState,450)}
function postNoCors(payload){if(!SHEETS_ENABLED)return Promise.reject(new Error("URL missing"));return fetch(SHEETS_WEB_APP_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(payload)})}
function pushLiveState(){if(!SHEETS_ENABLED){setSyncStatus("☁ Add Apps Script URL","err");return}setSyncStatus("☁ Syncing…","busy");postNoCors(livePayload()).then(()=>setSyncStatus("☁ Changes queued","ok")).catch(()=>setSyncStatus("☁ Sync failed","err"))}
function pullLiveState(){if(!SHEETS_ENABLED||pullBusy)return;pullBusy=true;setSyncStatus("☁ Loading shared state…","busy");const cb="__mobileState_"+Date.now();let called=false;const s=document.createElement("script");window[cb]=resp=>{called=true;try{if(resp&&resp.success)applyRemoteState(resp);else throw new Error(resp?.error||"No state")}catch(e){setSyncStatus("☁ Read failed","err")}finally{cleanup()}};function cleanup(){pullBusy=false;delete window[cb];s.remove()}s.src=`${SHEETS_WEB_APP_URL}?action=state&date=${encodeURIComponent(dk())}&week=${encodeURIComponent(wk())}&callback=${encodeURIComponent(cb)}&_=${Date.now()}`;s.onerror=()=>{setSyncStatus("☁ Access blocked","err");cleanup()};s.onload=()=>setTimeout(()=>{if(!called){setSyncStatus("☁ Sign-in/access blocked","err");cleanup()}},1200);document.head.appendChild(s)}
function applyRemoteState(resp){const t=resp.tasks||{};if(Array.isArray(t.daily)&&t.daily.length)data.daily=t.daily;if(!data.days[dk()])data.days[dk()]={tasks:[]};data.days[dk()].tasks=Array.isArray(t.today)?t.today:[];if(!data.weeks[wk()])data.weeks[wk()]={tasks:[]};if(Array.isArray(t.weekly)&&t.weekly.length)data.weeks[wk()].tasks=t.weekly;data.notesByDay[dk()]=resp.notes||"";persist();render();if(document.activeElement!==notes)notes.value=data.notesByDay[dk()]||"";setSyncStatus("☁ Synced with Sheets","ok")}
function calc(a){const done=a.filter(x=>x.done).length;return{done,total:a.length,pct:a.length?Math.round(done/a.length*100):0}}
function buildDailyReport(dateKey){const daily=calc(data.daily),today=calc(data.days[dateKey]?.tasks||[]),weekly=calc(data.weeks[weekKeyForDate(dateKey)]?.tasks||[]),total=daily.total+today.total+weekly.total,done=daily.done+today.done+weekly.done;return{date:dateKey,day:new Date(dateKey+"T12:00:00").toLocaleDateString(undefined,{weekday:"long"}),dailyPct:daily.pct,todayPct:today.pct,weeklyPct:weekly.pct,overallPct:total?Math.round(done/total*100):0,completed:done,missed:Math.max(0,total-done),notes:data.notesByDay?.[dateKey]||"",tasks:[...data.daily.map(x=>({category:"Daily",text:x.text,done:x.done})),...(data.days[dateKey]?.tasks||[]).map(x=>({category:"Today",text:x.text,done:x.done})),...(data.weeks[weekKeyForDate(dateKey)]?.tasks||[]).map(x=>({category:"Weekly",text:x.text,done:x.done}))]}}
function saveReportToSheets(){data.notesByDay[dk()]=notes.value;persist();pushLiveState();setSyncStatus("☁ Saving history…","busy");postNoCors({action:"snapshot",...buildDailyReport(dk())}).then(()=>setSyncStatus("☁ Live + history saved","ok")).catch(()=>setSyncStatus("☁ Save failed","err"))}
let noteTimer;notes.addEventListener("input",()=>{data.notesByDay[dk()]=notes.value;persist();clearTimeout(noteTimer);noteTimer=setTimeout(pushLiveState,700)});["daily","today","weekly"].forEach(t=>document.getElementById(t+"Input").addEventListener("keydown",e=>{if(e.key==="Enter")addTask(t)}));
notes.value=data.notesByDay[dk()]||"";render();pullLiveState();setInterval(pullLiveState,20000);if('serviceWorker' in navigator)window.addEventListener('load',()=>navigator.serviceWorker.register('./service-worker.js').catch(()=>{}));
