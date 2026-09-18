// ── AGKB Teacher Leave Tracker — Google Apps Script backend ──────────────────
const CONFIG_SHEET = '_Config';
const TEAM_SHEET   = 'Team';
const LOGS_SHEET   = 'Logs';
const TYPE_LEAVE        = 'Leave Request';
const TYPE_LEAVE_CREDIT = 'Leave Allowance';
const SUBTYPE_ANNUAL = 'Annual Leave';
const SUBTYPE_SICK   = 'Sick Day';
const ANNUAL_LEAVE_DAYS = 12;
const SICK_LEAVE_DAYS   = 7;

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (params.action==='approve'||params.action==='reject') return handleTokenAction(params);
  return HtmlService.createHtmlOutput('<p>AGKB Leave Tracker API</p>');
}
function doPost(e) {
  try {
    const params=e&&e.parameter?e.parameter:{};
    const action=params.action, sheetId=params.sheetId, name=params.name||'', pin=params.pin||'';
    if (!sheetId) return json({success:false,message:'Missing sheetId.'});
    const ss=SpreadsheetApp.openById(sheetId), config=loadConfig(ss);
    const handlers={
      getConfig:       ()=>publicConfig(config),
      getTeam:         ()=>getTeam(ss,config),
      getMyLogs:       ()=>getMyLogs(ss,name),
      getTeamDashboard:()=>getTeamDashboard(ss,config,name,pin),
      getTeamLogs:     ()=>getTeamLogs(ss,config,name,pin),
      submitEntry:     ()=>submitEntry(ss,config,safeJson(params.entry)),
      importEntry:     ()=>importEntry(ss,config,pin,safeJson(params.entry)),
      reviewEntry:     ()=>reviewEntry(ss,config,pin,params.id,params.reviewAction,name),
      submitHandover:  ()=>submitHandover(ss,config,name,params.id,safeJson(params.handover)),
    };
    if (!handlers[action]) return json({success:false,message:'Unknown action: '+action});
    return json(handlers[action]());
  } catch(err) { return json({success:false,message:err.toString()}); }
}
function json(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function safeJson(s){try{return s?JSON.parse(s):null;}catch(_){return null;}}
function generateId(){const n=new Date();return 'LT-'+n.getFullYear()+String(n.getMonth()+1).padStart(2,'0')+String(n.getDate()).padStart(2,'0')+'-'+String(Math.floor(Math.random()*9000)+1000);}
function generateToken(){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,generateId()+Math.random())).replace(/[^a-zA-Z0-9]/g,'').substring(0,32);}

function handleTokenAction(params) {
  try {
    const token=params.token,action=params.action,sheetId=params.sheetId;
    if (!token||!sheetId||(action!=='approve'&&action!=='reject')) return HtmlService.createHtmlOutput(resultPage('Error','Invalid request.','#e74c3c'));
    const ss=SpreadsheetApp.openById(sheetId),config=loadConfig(ss);
    const logsSheet=ss.getSheetByName(LOGS_SHEET),teamSheet=ss.getSheetByName(TEAM_SHEET);
    if (!logsSheet) return HtmlService.createHtmlOutput(resultPage('Error','Logs sheet not found.','#e74c3c'));
    const data=logsSheet.getDataRange().getValues();
    for (let i=1;i<data.length;i++) {
      if (String(data[i][12]||'').trim()!==token) continue;
      if (data[i][7]!=='PENDING') return HtmlService.createHtmlOutput(resultPage('Already Processed','This request was already '+data[i][7].toLowerCase()+'.','#e67e22'));
      const newStatus=action==='approve'?'APPROVED':'REJECTED';
      const entrySubtype=String(data[i][14]||'').trim(),empName=String(data[i][2]).trim(),reviewedBy='Head of School (email link)';
      logsSheet.getRange(i+1,8).setValue(newStatus);logsSheet.getRange(i+1,9).setValue(reviewedBy);logsSheet.getRange(i+1,13).setValue('');
      if (newStatus==='APPROVED') logsSheet.getRange(i+1,14).setValue('PENDING_HANDOVER');
      const teamRows=teamSheet?teamSheet.getDataRange().getValues().slice(1):[];
      const member=teamRows.find(r=>String(r[0]).trim().toLowerCase()===empName.toLowerCase());
      if (member&&member[1]) sendStatusEmail(member[1],empName,data[i],newStatus,reviewedBy,sheetId);
      if (entrySubtype===SUBTYPE_SICK) {
        const fEmails=getFoundationEmails(config,teamRows);
        if (fEmails.length){let dA=null;try{dA=data[i][9]?JSON.parse(data[i][9]):null;}catch(_){}
          const dv=data[i][11]!==''&&data[i][11]!=null?Number(data[i][11]):1;
          fEmails.forEach(e=>sendFoundationEmail(e,empName,dA,dv,!!data[i][10],String(data[i][6]||''),String(data[i][0]),newStatus,reviewedBy));}
      }
      if (newStatus==='APPROVED'&&config['CAL_ID']){try{const cal=CalendarApp.getCalendarById(config['CAL_ID']);if(cal)createCalEvents(cal,empName,data[i][9]?JSON.parse(data[i][9]):null,data[i]);}catch(e){}}
      return HtmlService.createHtmlOutput(resultPage(newStatus==='APPROVED'?'Approved ✓':'Rejected','Leave request for '+empName+' has been '+newStatus.toLowerCase()+'.',newStatus==='APPROVED'?'#27ae60':'#e74c3c'));
    }
    return HtmlService.createHtmlOutput(resultPage('Not Found','Token not found or already used.','#e67e22'));
  } catch(err){return HtmlService.createHtmlOutput(resultPage('Error',err.toString(),'#e74c3c'));}
}
function resultPage(t,m,c){return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f0f4f8;}.card{background:#fff;border-radius:12px;padding:40px;max-width:420px;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,0.1);}h2{margin:0 0 12px;color:'+c+';}p{color:#4a5568;font-size:15px;line-height:1.6;margin:0;}</style></head><body><div class="card"><h2>'+t+'</h2><p>'+m+'</p></div></body></html>';}

function loadConfig(ss){const sheet=ss.getSheetByName(CONFIG_SHEET);if(!sheet)return{};const c={};sheet.getDataRange().getValues().forEach(r=>{if(r[0])c[String(r[0]).trim()]=String(r[1]||'').trim();});return c;}
function publicConfig(config){const s={};['APP_NAME','POLICY_YEAR','HOLIDAYS'].forEach(k=>{if(config[k])s[k]=config[k];});s.HAS_CALENDAR=config['CAL_ID']?'true':'false';s.HAS_PIN=config['PIN']?'true':'false';return{success:true,config:s};}

function getTeam(ss,config){
  const tS=ss.getSheetByName(TEAM_SHEET),lS=ss.getSheetByName(LOGS_SHEET);
  if(!tS)return{success:false,message:'Team sheet not found.'};
  const py=parseInt(config['POLICY_YEAR'])||new Date().getFullYear();
  const allLogs=lS?lS.getDataRange().getValues().slice(1).map(rowToLog):[];
  const team=tS.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>{
    const name=String(r[0]).trim(),batch=String(r[5]||'').trim(),noA=String(r[4]||'').trim().toUpperCase()==='TRUE';
    return{name,email:String(r[1]||'').trim(),role:String(r[2]||'Teacher').trim(),batch,balance:calcBalance(name,allLogs,py,batch,noA)};
  });
  return{success:true,members:team};
}

function getMyLogs(ss,name){
  const sheet=ss.getSheetByName(LOGS_SHEET);if(!sheet)return{success:false,logs:[]};
  const logs=sheet.getDataRange().getValues().slice(1).filter(r=>String(r[2]||'').trim().toLowerCase()===(name||'').toLowerCase()).map(rowToLog).reverse();
  return{success:true,logs};
}

function getTeamDashboard(ss,config,name,pin){
  const lS=ss.getSheetByName(LOGS_SHEET),tS=ss.getSheetByName(TEAM_SHEET);
  if(!lS||!tS)return{success:false,message:'Sheet not found.'};
  const teamRows=tS.getDataRange().getValues().slice(1).filter(r=>r[0]);
  const py=parseInt(config['POLICY_YEAR'])||new Date().getFullYear();
  const allLogs=lS.getDataRange().getValues().slice(1).map(rowToLog);
  const members=teamRows.map(r=>{
    const mn=String(r[0]).trim(),batch=String(r[5]||'').trim(),noA=String(r[4]||'').trim().toUpperCase()==='TRUE';
    return{name:mn,role:String(r[2]||'Teacher').trim(),batch,balance:calcBalance(mn,allLogs,py,batch,noA)};
  });
  const pending=allLogs.filter(l=>l.status==='PENDING');
  const pendingHandovers=allLogs.filter(l=>l.status==='APPROVED'&&l.handover==='PENDING_HANDOVER');
  const leaveLogs=allLogs.filter(l=>l.type===TYPE_LEAVE||l.type===TYPE_LEAVE_CREDIT);
  return{success:true,members,pending,pendingHandovers,leaveLogs,hasCalendar:!!config['CAL_ID']};
}

function getTeamLogs(ss,config,name,pin){
  const lS=ss.getSheetByName(LOGS_SHEET),tS=ss.getSheetByName(TEAM_SHEET);
  if(!lS||!tS)return{success:false,message:'Sheet not found.'};
  const teamRows=tS.getDataRange().getValues().slice(1).filter(r=>r[0]);
  const req=teamRows.find(r=>String(r[0]).trim().toLowerCase()===String(name||'').trim().toLowerCase());
  if(!req||String(req[2]||'').trim()!=='Head of School')return{success:false,message:'Access denied.'};
  if(config['PIN']&&(!pin||String(pin).trim()!==String(config['PIN']).trim()))return{success:false,message:'Incorrect PIN.'};
  const allLogs=lS.getDataRange().getValues().slice(1).map(rowToLog);
  return{success:true,leaveLogs:allLogs.filter(l=>l.type===TYPE_LEAVE||l.type===TYPE_LEAVE_CREDIT)};
}

function submitEntry(ss,config,entry){
  if(!entry)return{success:false,message:'Missing data.'};
  try{
    const lS=ss.getSheetByName(LOGS_SHEET),tS=ss.getSheetByName(TEAM_SHEET);
    if(!lS)return{success:false,message:'Logs sheet not found.'};
    const holidays=parseHolidays(config['HOLIDAYS']),entrySubtype=entry.entrySubtype||SUBTYPE_ANNUAL;
    const isSick=entrySubtype===SUBTYPE_SICK,halfDay=!!entry.halfDay,driveLink=String(entry.driveLink||'').trim();
    if(!entry.dates||!Array.isArray(entry.dates)||!entry.dates.length)return{success:false,message:'No dates provided.'};
    const datesArr=entry.dates.slice().sort();
    if(isSick&&datesArr.length>=2){
      if(!driveLink)return{success:false,message:'A doctor letter or MCU report link is required for sick leave of 2 or more consecutive days.'};
      if(!/^https:\/\/(drive|docs)\.google\.com\//.test(driveLink))return{success:false,message:'Please provide a valid Google Drive link.'};
    }
    const dayValue=countWorkdays(datesArr,holidays,halfDay),startDate=datesArr[0],endDate=datesArr.length>1?datesArr[datesArr.length-1]:'';
    const id=generateId(),token=generateToken();
    lS.appendRow([id,new Date(),entry.name,startDate?new Date(startDate):'',endDate?new Date(endDate):'',TYPE_LEAVE,entry.notes||'','PENDING','',JSON.stringify(datesArr),halfDay,dayValue,token,'',entrySubtype,driveLink]);
    if(config['HEAD_EMAIL']){
      const py=parseInt(config['POLICY_YEAR'])||new Date().getFullYear();
      const allLogs=lS.getDataRange().getValues().slice(1).map(rowToLog);
      const teamRows=tS?tS.getDataRange().getValues().slice(1):[];
      const mRow=teamRows.find(r=>String(r[0]).trim().toLowerCase()===entry.name.toLowerCase());
      const batch=mRow?String(mRow[5]||'').trim():'',noA=mRow?String(mRow[4]||'').trim().toUpperCase()==='TRUE':false;
      const balance=calcBalance(entry.name,allLogs,py,batch,noA);
      const afterBal=isSick?(balance.sickBalance-dayValue):(balance.annualBalance-dayValue);
      sendApprovalEmail(config['HEAD_EMAIL'],entry,id,balance,dayValue,afterBal,ss.getId(),token,datesArr,halfDay,entrySubtype,driveLink);
    }
    return{success:true,id,status:'PENDING',dayValue};
  }catch(err){return{success:false,message:err.toString()};}
}

function reviewEntry(ss,config,pin,id,action,reviewerName){
  if(!id||!action)return{success:false,message:'Missing data.'};
  try{
    if(!config['PIN']||!pin||String(pin).trim()!==String(config['PIN']).trim())return{success:false,message:'Incorrect PIN.'};
    const lS=ss.getSheetByName(LOGS_SHEET),tS=ss.getSheetByName(TEAM_SHEET);
    if(!lS)return{success:false,message:'Logs sheet not found.'};
    const data=lS.getDataRange().getValues();
    for(let i=1;i<data.length;i++){
      if(String(data[i][0]).trim()!==id)continue;
      const newStatus=action==='approve'?'APPROVED':'REJECTED',entrySubtype=String(data[i][14]||'').trim(),empName=String(data[i][2]).trim();
      lS.getRange(i+1,8).setValue(newStatus);lS.getRange(i+1,9).setValue(reviewerName);lS.getRange(i+1,13).setValue('');
      if(newStatus==='APPROVED')lS.getRange(i+1,14).setValue('PENDING_HANDOVER');
      const teamRows=tS?tS.getDataRange().getValues().slice(1):[];
      const member=teamRows.find(r=>String(r[0]).trim().toLowerCase()===empName.toLowerCase());
      if(member&&member[1])sendStatusEmail(member[1],empName,data[i],newStatus,reviewerName,ss.getId());
      if(entrySubtype===SUBTYPE_SICK){
        const fEmails=getFoundationEmails(config,teamRows);
        if(fEmails.length){let dA=null;try{dA=data[i][9]?JSON.parse(data[i][9]):null;}catch(_){}
          const dv=data[i][11]!==''&&data[i][11]!=null?Number(data[i][11]):1;
          fEmails.forEach(e=>sendFoundationEmail(e,empName,dA,dv,!!data[i][10],String(data[i][6]||''),String(data[i][0]),newStatus,reviewerName));}
      }
      if(newStatus==='APPROVED'&&config['CAL_ID']){try{const cal=CalendarApp.getCalendarById(config['CAL_ID']);if(cal)createCalEvents(cal,empName,data[i][9]?JSON.parse(data[i][9]):null,data[i]);}catch(e){}}
      return{success:true,newStatus};
    }
    return{success:false,message:'Entry not found.'};
  }catch(err){return{success:false,message:err.toString()};}
}

function importEntry(ss,config,pin,entry){
  if(!entry)return{success:false,message:'Missing entry data.'};
  if(!config['PIN']||String(pin||'').trim()!==String(config['PIN']).trim())return{success:false,message:'Incorrect PIN.'};
  try{
    const lS=ss.getSheetByName(LOGS_SHEET);if(!lS)return{success:false,message:'Logs sheet not found.'};
    const id=generateId(),reviewedBy='Historical import by '+(entry.importedBy||'Head of School');
    if(entry.type===TYPE_LEAVE_CREDIT){
      const year=parseInt(entry.year,10),days=parseFloat(entry.days);
      if(!year||isNaN(days)||days<=0)return{success:false,message:'Provide a valid year and days.'};
      lS.appendRow([id,new Date(),entry.name,new Date(year+'-01-01'),'',TYPE_LEAVE_CREDIT,entry.notes||'','APPROVED',reviewedBy,JSON.stringify([year+'-01-01']),false,days,'','','','']);
      return{success:true,id,dayValue:days};
    }
    const holidays=parseHolidays(config['HOLIDAYS']),halfDay=!!entry.halfDay,entrySubtype=entry.entrySubtype||SUBTYPE_ANNUAL,driveLink=String(entry.driveLink||'').trim();
    let datesArr;
    if(entry.dates&&Array.isArray(entry.dates)&&entry.dates.length>0){datesArr=entry.dates.slice().sort();}
    else if(entry.startDate){datesArr=buildDateRange(entry.startDate,entry.endDate||entry.startDate);}
    else return{success:false,message:'No dates provided.'};
    const dayValue=countWorkdays(datesArr,holidays,halfDay),startDate=datesArr[0],endDate=datesArr.length>1?datesArr[datesArr.length-1]:'';
    lS.appendRow([id,new Date(),entry.name,startDate?new Date(startDate):'',endDate?new Date(endDate):'',TYPE_LEAVE,entry.notes||'','APPROVED',reviewedBy,JSON.stringify(datesArr),halfDay,dayValue,'','',entrySubtype,driveLink]);
    return{success:true,id,dayValue};
  }catch(err){return{success:false,message:err.toString()};}
}

function calcBalance(name,allLogs,policyYear,batch,noAllowance){
  const nl=name.trim().toLowerCase(),isBatch1=String(batch||'').trim()==='Batch 1',isBatch2=String(batch||'').trim()==='Batch 2';
  const myY=allLogs.filter(l=>String(l.name||'').trim().toLowerCase()===nl&&l.startDate&&new Date(l.startDate).getFullYear()===policyYear);
  const annualUsed=myY.filter(l=>l.type===TYPE_LEAVE&&l.status==='APPROVED'&&l.entrySubtype===SUBTYPE_ANNUAL).reduce((s,l)=>s+(l.dayValue||1),0);
  const sickUsed=myY.filter(l=>l.type===TYPE_LEAVE&&l.status==='APPROVED'&&l.entrySubtype===SUBTYPE_SICK).reduce((s,l)=>s+(l.dayValue||1),0);
  const annualPending=myY.filter(l=>l.type===TYPE_LEAVE&&l.status==='PENDING'&&l.entrySubtype===SUBTYPE_ANNUAL).reduce((s,l)=>s+(l.dayValue||1),0);
  const sickPending=myY.filter(l=>l.type===TYPE_LEAVE&&l.status==='PENDING'&&l.entrySubtype===SUBTYPE_SICK).reduce((s,l)=>s+(l.dayValue||1),0);
  const annualAlloc=noAllowance?0:(isBatch1?ANNUAL_LEAVE_DAYS:0);
  const sickAlloc=noAllowance?0:(isBatch2?SICK_LEAVE_DAYS:0);
  const prior=allLogs.filter(l=>String(l.name||'').trim().toLowerCase()===nl&&l.startDate&&new Date(l.startDate).getFullYear()<policyYear);
  const pby={};
  prior.forEach(l=>{const yr=new Date(l.startDate).getFullYear();if(!pby[yr])pby[yr]={ac:0,au:0,sc:0,su:0};
    if(l.type===TYPE_LEAVE_CREDIT)pby[yr].ac+=(l.dayValue||1);
    if(l.type===TYPE_LEAVE&&l.status==='APPROVED'&&l.entrySubtype===SUBTYPE_ANNUAL)pby[yr].au+=(l.dayValue||1);
    if(l.type===TYPE_LEAVE&&l.status==='APPROVED'&&l.entrySubtype===SUBTYPE_SICK)pby[yr].su+=(l.dayValue||1);
  });
  let ar=0,sr=0;const rb=[];
  Object.keys(pby).sort().forEach(yr=>{const v=pby[yr];const an=Math.max(0,v.ac-v.au),sk=Math.max(0,v.sc-v.su);
    if(an>0){ar+=an;rb.push({year:parseInt(yr),bucket:'annual',days:an});}
    if(sk>0){sr+=sk;rb.push({year:parseInt(yr),bucket:'sick',days:sk});}
  });
  const annualBalance=annualAlloc+ar-annualUsed,sickBalance=sickAlloc+sr-sickUsed;
  return{batch,annualAlloc,annualRollover:ar,annualUsedDays:annualUsed,annualBalance,annualPending,sickAlloc,sickRollover:sr,sickUsedDays:sickUsed,sickBalance,sickPending,balance:isBatch1?annualBalance:sickBalance,rolloverBreakdown:rb};
}

function submitHandover(ss,config,name,id,handover){
  if(!id||!handover)return{success:false,message:'Missing data.'};
  try{
    const lS=ss.getSheetByName(LOGS_SHEET),tS=ss.getSheetByName(TEAM_SHEET);if(!lS)return{success:false,message:'Logs sheet not found.'};
    const data=lS.getDataRange().getValues();
    for(let i=1;i<data.length;i++){if(String(data[i][0]).trim()!==id)continue;
      lS.getRange(i+1,14).setValue(JSON.stringify(handover));
      const teamRows=tS?tS.getDataRange().getValues().slice(1):[];
      const empName=String(data[i][2]).trim();let dA=null;try{dA=data[i][9]?JSON.parse(data[i][9]):null;}catch(_){}
      const dd=dA&&dA.length>0?(dA.length===1?dA[0]:dA[0]+' to '+dA[dA.length-1]):'';
      teamRows.forEach(r=>{if(r[1]&&String(r[1]).includes('@'))sendHandoverEmail(String(r[1]).trim(),empName,dd,handover);});
      return{success:true};
    }
    return{success:false,message:'Entry not found.'};
  }catch(err){return{success:false,message:err.toString()};}
}

function getFoundationEmails(config,teamRows){
  const e=new Set();
  (config['FOUNDATION_EMAIL']||'').split(',').map(s=>s.trim()).filter(s=>s.includes('@')).forEach(s=>e.add(s.toLowerCase()));
  if(teamRows)teamRows.forEach(r=>{if(String(r[2]||'').trim()==='Foundation Officer'&&String(r[1]||'').includes('@'))e.add(String(r[1]).trim().toLowerCase());});
  return Array.from(e);
}
function fmt(d){try{return Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd MMM yyyy');}catch(_){return String(d);}}
function dStr(dA){return dA&&dA.length>0?(dA.length===1?fmt(dA[0]):fmt(dA[0])+' – '+fmt(dA[dA.length-1])):'—';}

function sendApprovalEmail(headEmail,entry,id,balance,daysReq,afterBal,sheetId,token,datesArr,halfDay,entrySubtype,driveLink){
  const isSick=entrySubtype===SUBTYPE_SICK,sl=isSick?'Sick Day':'Annual Leave';
  const cb=isSick?balance.sickBalance:balance.annualBalance,alloc=isSick?balance.sickAlloc:balance.annualAlloc;
  const rl=isSick?balance.sickRollover:balance.annualRollover,used=isSick?balance.sickUsedDays:balance.annualUsedDays;
  const pend=isSick?balance.sickPending:balance.annualPending,bc=afterBal<0?'#e74c3c':afterBal<=1?'#e67e22':'#27ae60';
  const appUrl=ScriptApp.getService().getUrl();
  const aUrl=appUrl+'?action=approve&token='+token+'&sheetId='+sheetId,rUrl=appUrl+'?action=reject&token='+token+'&sheetId='+sheetId;
  const sub='[Leave Request] '+entry.name+' — '+dStr(datesArr);
  const body='<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;color:#1a2332;">'+
    '<div style="background:#1a2332;padding:20px 24px;border-radius:12px 12px 0 0;"><div style="color:white;font-size:18px;font-weight:700;">Leave Request — '+sl+'</div>'+
    '<div style="color:rgba(255,255,255,0.5);font-size:13px;margin-top:2px;">Requires your review as Head of School</div></div>'+
    '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:24px;border-radius:0 0 12px 12px;">'+
    '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;">'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;width:140px;">Teacher</td><td style="padding:6px 0;font-weight:600;font-size:14px;">'+entry.name+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Type</td><td style="padding:6px 0;font-weight:600;font-size:14px;">'+sl+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Date(s)</td><td style="padding:6px 0;font-weight:600;font-size:14px;">'+dStr(datesArr)+(halfDay?' (half-day)':'')+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Days</td><td style="padding:6px 0;font-weight:600;font-size:14px;">'+daysReq+'</td></tr>'+
    (entry.notes?'<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;vertical-align:top;">Notes</td><td style="padding:6px 0;font-size:14px;">'+entry.notes+'</td></tr>':'')+
    (driveLink?'<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Doctor Letter</td><td style="padding:6px 0;font-size:14px;"><a href="'+driveLink+'" style="color:#0e17c2;">View Document</a></td></tr>':'')+
    '</table><div style="background:#f0f4f8;border-radius:8px;padding:16px;margin-bottom:24px;">'+
    '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:#8a96a3;margin-bottom:12px;">'+sl+' Balance — '+entry.name+'</div>'+
    '<table style="width:100%;border-collapse:collapse;">'+
    '<tr><td style="padding:5px 0;font-size:13px;color:#4a5568;">Annual Allocation</td><td style="padding:5px 0;font-size:13px;font-weight:600;text-align:right;">+'+alloc+' days</td></tr>'+
    '<tr><td style="padding:5px 0;font-size:13px;color:#4a5568;">Rollover</td><td style="padding:5px 0;font-size:13px;font-weight:600;text-align:right;">+'+rl+' days</td></tr>'+
    '<tr><td style="padding:5px 0;font-size:13px;color:#4a5568;">Used</td><td style="padding:5px 0;font-size:13px;font-weight:600;text-align:right;">-'+used+' days</td></tr>'+
    (pend>0?'<tr><td style="padding:5px 0;font-size:13px;color:#4a5568;">Other Pending</td><td style="padding:5px 0;font-size:13px;font-weight:600;text-align:right;">-'+pend+' days</td></tr>':'')+
    '<tr style="border-top:2px solid #e2e8f0;"><td style="padding:8px 0 5px;font-size:14px;font-weight:700;">Current Balance</td><td style="padding:8px 0 5px;font-size:14px;font-weight:700;text-align:right;">'+cb+' days</td></tr>'+
    '<tr style="border-top:1.5px dashed #e2e8f0;"><td style="padding:8px 0 0;font-size:13px;color:#4a5568;">After approval</td><td style="padding:8px 0 0;font-size:15px;font-weight:800;text-align:right;color:'+bc+';">'+afterBal+' days remaining</td></tr>'+
    '</table></div><table style="width:100%;border-collapse:collapse;margin-bottom:16px;"><tr>'+
    '<td style="padding-right:8px;"><a href="'+aUrl+'" style="display:block;text-align:center;background:#27ae60;color:#fff;padding:14px 0;border-radius:8px;font-size:15px;font-weight:700;text-decoration:none;">Approve</a></td>'+
    '<td style="padding-left:8px;"><a href="'+rUrl+'" style="display:block;text-align:center;background:#e74c3c;color:#fff;padding:14px 0;border-radius:8px;font-size:15px;font-weight:700;text-decoration:none;">Reject</a></td>'+
    '</tr></table><div style="font-size:11px;color:#8a96a3;text-align:center;padding-top:12px;border-top:1px solid #e2e8f0;">Entry ID: '+id+' · Links work once and expire after use.</div></div></div>';
  GmailApp.sendEmail(headEmail,sub,'',{htmlBody:body,name:'AGKB Leave Tracker'});
}

function sendStatusEmail(empEmail,empName,row,status,reviewerName,sheetId){
  let dA=null;try{dA=row[9]?JSON.parse(row[9]):null;}catch(_){}
  const sl=String(row[14]||'').trim()===SUBTYPE_SICK?'Sick Day':'Annual Leave',isApp=status==='APPROVED',color=isApp?'#27ae60':'#e74c3c';
  const body='<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;"><div style="background:'+color+';padding:20px 24px;border-radius:12px 12px 0 0;"><div style="color:white;font-size:18px;font-weight:700;">'+(isApp?'Request Approved ✓':'Request Rejected')+'</div></div>'+
    '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:24px;border-radius:0 0 12px 12px;">'+
    '<p style="font-size:14px;color:#4a5568;line-height:1.6;">Your <strong>'+sl+'</strong> request for <strong>'+dStr(dA)+'</strong> has been <strong>'+status.toLowerCase()+'</strong> by '+reviewerName+'.</p>'+
    '<p style="font-size:13px;color:#8a96a3;">Entry ID: '+String(row[0])+'</p></div></div>';
  GmailApp.sendEmail(empEmail,'[AGKB Leave] Your request has been '+status.toLowerCase(),'',{htmlBody:body,name:'AGKB Leave Tracker'});
}

function sendFoundationEmail(email,empName,datesArr,dayValue,halfDay,notes,id,status,reviewedBy){
  const body='<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;"><div style="background:#1a2332;padding:20px 24px;border-radius:12px 12px 0 0;"><div style="color:white;font-size:17px;font-weight:700;">Sick Day Request Update</div><div style="color:rgba(255,255,255,0.5);font-size:13px;">For Foundation Officer records</div></div>'+
    '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;padding:24px;border-radius:0 0 12px 12px;"><table style="width:100%;border-collapse:collapse;">'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;width:130px;">Teacher</td><td style="padding:6px 0;font-size:14px;font-weight:600;">'+empName+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Date(s)</td><td style="padding:6px 0;font-size:14px;">'+dStr(datesArr)+(halfDay?' (half-day)':'')+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Days</td><td style="padding:6px 0;font-size:14px;">'+dayValue+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Status</td><td style="padding:6px 0;font-size:14px;font-weight:700;color:'+(status==='APPROVED'?'#27ae60':'#e74c3c')+';">'+status+'</td></tr>'+
    '<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Decided by</td><td style="padding:6px 0;font-size:14px;">'+reviewedBy+'</td></tr>'+
    (notes?'<tr><td style="padding:6px 0;color:#8a96a3;font-size:13px;">Notes</td><td style="padding:6px 0;font-size:14px;">'+notes+'</td></tr>':'')+
    '</table><p style="font-size:11px;color:#8a96a3;margin-top:16px;">Entry ID: '+id+'</p></div></div>';
  GmailApp.sendEmail(email,'[AGKB Leave] Sick Day '+status+' — '+empName,'',{htmlBody:body,name:'AGKB Leave Tracker'});
}

function sendHandoverEmail(recipientEmail,empName,datesDisplay,handover){
  function row(l,v){if(!v)return '';return '<tr><td style="padding:8px 12px;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.4px;color:#8a96a3;vertical-align:top;width:160px;">'+l+'</td><td style="padding:8px 12px;font-size:14px;color:#1a2332;line-height:1.6;">'+String(v).replace(/\n/g,'<br>')+'</td></tr>';}
  const btn=handover.driveLink?'<div style="margin-top:16px;"><a href="'+handover.driveLink+'" style="display:inline-block;background:#0e17c2;color:#fff;padding:10px 20px;border-radius:8px;font-size:14px;font-weight:600;text-decoration:none;">Open Drive Link</a></div>':'';
  const body='<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;color:#1a2332;"><div style="background:#1a2332;padding:20px 24px;border-radius:12px 12px 0 0;"><div style="color:white;font-size:18px;font-weight:700;">Handover Notice</div><div style="color:rgba(255,255,255,0.45);font-size:13px;margin-top:2px;">'+empName+' · '+datesDisplay+'</div></div>'+
    '<div style="background:#fff;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 12px 12px;overflow:hidden;"><table style="width:100%;border-collapse:collapse;">'+
    row('Tasks in Progress',handover.tasks)+row('Point of Contact',handover.contact)+row('Urgent Items',handover.urgent)+row('Additional Notes',handover.notes)+
    '</table>'+btn+'<div style="padding:14px 16px;background:#f0f4f8;margin-top:16px;font-size:11px;color:#8a96a3;">AGKB Leave Tracker</div></div></div>';
  GmailApp.sendEmail(recipientEmail,'[AGKB Leave] Handover Notice from '+empName,'',{htmlBody:body,name:'AGKB Leave Tracker'});
}

function createCalEvents(cal,name,datesArr,row){
  if(!datesArr||!datesArr.length)return;
  const title=name+' — '+(String(row[14]||'').trim()===SUBTYPE_SICK?'Sick Day':'Annual Leave');
  datesArr.forEach(d=>{try{cal.createAllDayEvent(title,new Date(d+'T00:00:00'),{description:'AGKB Leave Tracker\nEntry: '+String(row[0])});}catch(e){}});
}

function parseHolidays(v){if(!v||String(v).trim()==='')return[];return String(v).split(',').map(s=>s.trim()).filter(s=>s.length===10);}
function buildDateRange(s,e){const dates=[],start=new Date(s+'T00:00:00'),end=new Date(e+'T00:00:00');for(let d=new Date(start);d<=end;d.setDate(d.getDate()+1))dates.push(Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'yyyy-MM-dd'));return dates;}
function countWorkdays(dA,holidays,halfDay){const h=new Set(holidays||[]);let c=0;(dA||[]).forEach(d=>{const day=new Date(d).getDay();if(day!==0&&day!==6&&!h.has(d))c++;});return halfDay?c*0.5:c;}
function rowToLog(r){let dA=null;try{dA=r[9]?JSON.parse(r[9]):null;}catch(_){}
  return{id:String(r[0]||'').trim(),timestamp:r[1]?new Date(r[1]).toISOString():'',name:String(r[2]||'').trim(),
    startDate:r[3]?Utilities.formatDate(new Date(r[3]),Session.getScriptTimeZone(),'yyyy-MM-dd'):'',
    endDate:r[4]?Utilities.formatDate(new Date(r[4]),Session.getScriptTimeZone(),'yyyy-MM-dd'):'',
    type:String(r[5]||'').trim(),notes:String(r[6]||'').trim(),status:String(r[7]||'').trim(),reviewedBy:String(r[8]||'').trim(),
    dates:dA,halfDay:!!r[10],dayValue:r[11]!==''&&r[11]!==undefined&&r[11]!==null?Number(r[11]):null,
    handover:r[13]?String(r[13]).trim():'',entrySubtype:String(r[14]||'').trim(),driveLink:String(r[15]||'').trim()};}
function forceAuth(){GmailApp.getAliases();CalendarApp.getDefaultCalendar();DriveApp.getRootFolder();}
