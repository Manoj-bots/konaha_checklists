const KEY="MOBILE_CHECKLIST_V1";
const DAILY_DEFAULT=["Check emails / messages","Send morning update","Check Jira / assigned tasks","Update task progress","Send evening update"];
const WEEKLY_DEFAULT=["WSR – Email","WSR – Message","Fill Timesheet"];

function localDayKey(){
 const d=new Date();
 return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function localWeekKey(){
 const d=new Date(), day=d.getDay();
 d.setDate(d.getDate()+(day===0?-6:1-day));
 return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
let data=JSON.parse(localStorage.getItem(KEY)||"null");
if(!data)data={daily:DAILY_DEFAULT.map((text,i)=>({id:"d"+i,text,done:false})),days:{},weeks:{}};
if(!data.days[localDayKey()])data.days[localDayKey()]=[];
if(!data.weeks[localWeekKey()])data.weeks[localWeekKey()]=WEEKLY_DEFAULT.map((text,i)=>({id:"w"+i,text,done:false}));

function save(){localStorage.setItem(KEY,JSON.stringify(data));render();}
function tasks(type){return type==="daily"?data.daily:type==="weekly"?data.weeks[localWeekKey()]:data.days[localDayKey()];}
function esc(s){return s.replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));}
function renderList(type,id){
 document.getElementById(id).innerHTML=tasks(type).map(x=>`
 <div class="task ${x.done?"done":""}">
  <input type="checkbox" ${x.done?"checked":""} onchange="toggleTask('${type}','${x.id}')">
  <input type="text" value="${esc(x.text)}" onchange="editTask('${type}','${x.id}',this.value)">
  <button class="del" onclick="deleteTask('${type}','${x.id}')">×</button>
 </div>`).join("");
}
function stat(type){
 const a=tasks(type),done=a.filter(x=>x.done).length;
 return [done,a.length,a.length?Math.round(done/a.length*100):0];
}
function render(){
 const d=new Date();
 weekday.textContent=d.toLocaleDateString(undefined,{weekday:"long"}).toUpperCase();
 day.textContent=d.getDate();
 month.textContent=d.toLocaleDateString(undefined,{month:"long",year:"numeric"}).toUpperCase();
 clock.textContent=d.toLocaleTimeString(undefined,{hour:"2-digit",minute:"2-digit"});
 renderList("daily","dailyList");renderList("today","todayList");renderList("weekly","weeklyList");
 const a=stat("daily"),b=stat("today"),c=stat("weekly");
 [["daily",a],["today",b],["weekly",c]].forEach(([n,s])=>{
  document.getElementById(n+"Pct").textContent=s[2]+"%";
  document.getElementById(n+"Bar").style.width=s[2]+"%";
 });
 const total=a[1]+b[1]+c[1],done=a[0]+b[0]+c[0],p=total?Math.round(done/total*100):0;
 overallPct.textContent=p+"%";overallBar.style.width=p+"%";
 dailySummary.textContent=`${a[0]}/${a[1]}`;
 todaySummary.textContent=`${b[0]}/${b[1]}`;
 weeklySummary.textContent=`${c[0]}/${c[1]}`;
}
function toggleTask(t,id){const x=tasks(t).find(x=>x.id===id);if(x){x.done=!x.done;save();}}
function editTask(t,id,value){const x=tasks(t).find(x=>x.id===id);if(x){x.text=value.trim()||"Untitled task";save();}}
function deleteTask(t,id){const a=tasks(t),i=a.findIndex(x=>x.id===id);if(i>=0){a.splice(i,1);save();}}
function addTask(t){
 const input=document.getElementById(t+"Input"),value=input.value.trim();
 if(!value)return;
 tasks(t).push({id:t[0]+Date.now(),text:value,done:false});
 input.value="";save();
}
function resetToday(){data.days[localDayKey()]=[];data.daily.forEach(x=>x.done=false);save();}
function resetWeek(){data.weeks[localWeekKey()].forEach(x=>x.done=false);save();}
notes.value=localStorage.getItem(KEY+"_notes")||"";
notes.oninput=()=>localStorage.setItem(KEY+"_notes",notes.value);
render();
setInterval(render,30000);