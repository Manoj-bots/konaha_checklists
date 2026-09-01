const LIVE_SHEET = 'Live_State';
const META_SHEET = 'Live_Meta';
const SUMMARY_SHEET = 'Daily_Summary';
const TASK_SHEET = 'Task_Log';

const DEFAULT_DAILY = [
  ['d0','Check emails / messages'],
  ['d1','Send morning update'],
  ['d2','Check Jira / assigned tasks'],
  ['d3','Update task progress'],
  ['d4','Send evening update']
];
const DEFAULT_WEEKLY = [
  ['w0','WSR – Email'],
  ['w1','WSR – Message'],
  ['w2','Fill Timesheet']
];

function doGet(e) {
  try {
    setupSheets();
    const action = String((e && e.parameter && e.parameter.action) || 'ping');
    let out;
    if (action === 'state') {
      const dateKey = String(e.parameter.date || todayKey());
      const weekKey = String(e.parameter.week || weekKeyForDate(dateKey));
      ensureDefaults(dateKey, weekKey);
      out = getState(dateKey, weekKey);
    } else {
      out = {success:true, message:'Konoha Checklist API is running'};
    }
    const cb = e && e.parameter ? e.parameter.callback : '';
    if (cb) {
      return ContentService.createTextOutput(`${cb}(${JSON.stringify(out)});`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonResponse(out);
  } catch (err) {
    return jsonResponse({success:false,error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    setupSheets();
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.action === 'state') {
      saveLiveState(data);
      return jsonResponse({success:true, action:'state', date:data.date});
    }
    if (data.action === 'snapshot') {
      saveDailySummary(data);
      saveTaskLog(data);
      return jsonResponse({success:true, action:'snapshot', date:data.date});
    }
    throw new Error('Unknown action');
  } catch (err) {
    return jsonResponse({success:false,error:String(err && err.message || err)});
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = [
    [LIVE_SHEET,['Scope','Scope Key','Category','Task ID','Task','Done','Updated At','Updated By']],
    [META_SHEET,['Date','Notes','Updated At','Updated By']],
    [SUMMARY_SHEET,['Date','Day','Daily %',"Today's %",'Weekly %','Overall %','Completed','Missed','Notes','Last Updated']],
    [TASK_SHEET,['Date','Day','Category','Task','Status','Completed','Last Updated']]
  ];
  defs.forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  });
}

function ensureDefaults(dateKey, weekKey) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIVE_SHEET);
  const values = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,8).getValues() : [];
  const hasDaily = values.some(r => r[0] === 'date' && String(r[1]) === dateKey && r[2] === 'Daily');
  const hasWeekly = values.some(r => r[0] === 'week' && String(r[1]) === weekKey && r[2] === 'Weekly');
  const now = new Date();
  if (!hasDaily) appendRows(sh, DEFAULT_DAILY.map(([id,text]) => ['date',dateKey,'Daily',id,text,false,now,'system']));
  if (!hasWeekly) appendRows(sh, DEFAULT_WEEKLY.map(([id,text]) => ['week',weekKey,'Weekly',id,text,false,now,'system']));
}

function getState(dateKey, weekKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LIVE_SHEET);
  const vals = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,8).getValues() : [];
  const tasks = {daily:[],today:[],weekly:[]};
  vals.forEach(r => {
    const scope = String(r[0]);
    const key = String(r[1]);
    const cat = String(r[2]);
    if (scope === 'date' && key === dateKey && cat === 'Daily') tasks.daily.push(rowTask(r));
    if (scope === 'date' && key === dateKey && cat === 'Today') tasks.today.push(rowTask(r));
    if (scope === 'week' && key === weekKey && cat === 'Weekly') tasks.weekly.push(rowTask(r));
  });
  const meta = ss.getSheetByName(META_SHEET);
  let notes = '';
  if (meta.getLastRow() >= 2) {
    const mv = meta.getRange(2,1,meta.getLastRow()-1,4).getValues();
    for (let i=mv.length-1;i>=0;i--) if (String(mv[i][0]) === dateKey) { notes = String(mv[i][1] || ''); break; }
  }
  return {success:true,date:dateKey,week:weekKey,tasks,notes};
}

function rowTask(r) { return {id:String(r[3]), text:String(r[4]), done:r[5] === true || String(r[5]).toLowerCase()==='true'}; }

function saveLiveState(data) {
  const dateKey = String(data.date || todayKey());
  const weekKey = String(data.week || weekKeyForDate(dateKey));
  const device = String(data.device || 'unknown');
  const tasks = data.tasks || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(LIVE_SHEET);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    removeScopedRows(sh, 'date', dateKey, ['Daily','Today']);
    removeScopedRows(sh, 'week', weekKey, ['Weekly']);
    const now = new Date();
    const rows = [];
    (tasks.daily || []).forEach(t => rows.push(['date',dateKey,'Daily',String(t.id),String(t.text||''),!!t.done,now,device]));
    (tasks.today || []).forEach(t => rows.push(['date',dateKey,'Today',String(t.id),String(t.text||''),!!t.done,now,device]));
    (tasks.weekly || []).forEach(t => rows.push(['week',weekKey,'Weekly',String(t.id),String(t.text||''),!!t.done,now,device]));
    appendRows(sh, rows);
    upsertMeta(dateKey, String(data.notes || ''), device);
  } finally { lock.releaseLock(); }
}

function removeScopedRows(sh, scope, key, categories) {
  if (sh.getLastRow() < 2) return;
  const vals = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  for (let i=vals.length-1;i>=0;i--) {
    if (String(vals[i][0])===scope && String(vals[i][1])===key && categories.includes(String(vals[i][2]))) sh.deleteRow(i+2);
  }
}

function upsertMeta(dateKey, notes, device) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(META_SHEET);
  let row = -1;
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (let i=0;i<vals.length;i++) if (String(vals[i][0])===dateKey) { row=i+2; break; }
  }
  const values = [[dateKey, notes, new Date(), device]];
  if (row > 0) sh.getRange(row,1,1,4).setValues(values); else sh.appendRow(values[0]);
}

function saveDailySummary(data) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUMMARY_SHEET);
  let row = -1;
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();
    for (let i=0;i<vals.length;i++) if (vals[i][0]===String(data.date)) {row=i+2;break;}
  }
  const values = [[String(data.date||''),String(data.day||''),Number(data.dailyPct||0)/100,Number(data.todayPct||0)/100,Number(data.weeklyPct||0)/100,Number(data.overallPct||0)/100,Number(data.completed||0),Number(data.missed||0),String(data.notes||''),new Date()]];
  if (row>0) sh.getRange(row,1,1,10).setValues(values); else {sh.appendRow(values[0]);row=sh.getLastRow();}
  sh.getRange(row,3,1,4).setNumberFormat('0%');
}

function saveTaskLog(data) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET);
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();
    for (let i=vals.length-1;i>=0;i--) if (vals[i][0]===String(data.date)) sh.deleteRow(i+2);
  }
  const now = new Date();
  const rows = (data.tasks || []).map(t => [String(data.date||''),String(data.day||''),String(t.category||''),String(t.text||''),t.done?'Done':'Pending',t.done?1:0,now]);
  appendRows(sh, rows);
}

function appendRows(sh, rows) { if (rows && rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
function todayKey() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function weekKeyForDate(key) { const d=new Date(key+'T12:00:00'); const day=d.getDay(); d.setDate(d.getDate()+(day===0?-6:1-day)); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
