const APP = {
  NAME: 'INSOMNIA',
  SHEETS: { REG:'Registrations', TICK:'Tickets', LOG:'EntryLogs', USERS:'Users', CONFIG:'Config' },
  NIGHTS: ['NIGHT 1','NIGHT 2','NIGHT 3'],
  BATCHES: ['2021','2022','2023','2024','2025','2026','OTHERS'],
  CR_NAMES: {
    '2021':['Kashish Mahajan',''],
    '2022':['Rhythm Gupta',''],
    '2023':['Gurman Singh Bhatia',''],
    '2024':['Nishant Mittal',''],
    '2025':['Ishan',''],
    '2026':['',''],
    'OTHERS':['','']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('INSOMNIA · Entry Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function setupSystem(){
  const ss = SpreadsheetApp.getActive();
  Object.entries(APP.SHEETS).forEach(([key,name])=>{
    let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
    if(sh.getLastRow()===0){
      const headers={
        REG:['createdAt','registrationId','timestamp','email','name','rollNumber','mobile','batch','paidTo','utr','screenshot','amount','crName','paymentStatus','ticketId','notes'],
        TICK:['createdAt','ticketId','registrationId','name','rollNumber','mobile','batch','qrToken','active','night1','night2','night3','issuedAt'],
        LOG:['timestamp','ticketId','registrationId','night','gate','scanner','result','operator','note'],
        USERS:['username','name','role','batch','gate','pinHash','active','createdAt'],
        CONFIG:['key','value']
      }[key]; sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1);
    }
  });
  seedConfig_(); seedUsers_();
  return 'System ready. Sheets created and initial CR records loaded.';
}

function seedConfig_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.CONFIG);
  const existing=sh.getDataRange().getValues();
  const keys=new Set(existing.slice(1).map(r=>String(r[0])));
  const rows=[];
  [['ADMIN_PASSWORD','CHANGE-ME-1234'],['CURRENT_NIGHT','NIGHT 1'],['ENTRY_OPEN','19:00']].forEach(x=>{if(!keys.has(x[0]))rows.push(x)});
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
}

function seedUsers_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.USERS);
  const existing=sh.getDataRange().getValues();
  const usernames=new Set(existing.slice(1).map(r=>String(r[0])));
  const rows=[];
  APP.BATCHES.forEach(b=>APP.CR_NAMES[b].forEach((name,i)=>{
    const u=`CR-${b}-${i+1}`;
    if(name && !usernames.has(u)) rows.push([u,name,'CR',b,'','',true,new Date()]);
  }));
  ['BOYS-01','BOYS-02','GIRLS-01','GIRLS-02'].forEach(u=>{if(!usernames.has(u)){const gate=u.split('-')[0];rows.push([u,u,'SCANNER','',gate,'',true,new Date()])}});
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,8).setValues(rows);
}

function onFormSubmit(e){
  const row=e.range.getRow(), sh=e.range.getSheet();
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const vals=sh.getRange(row,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const get=(names)=>{for(const n of names){const i=headers.findIndex(h=>h.trim().toLowerCase()===n.toLowerCase()); if(i>=0)return vals[i]} return ''};
  const batch=get(['Batch','Year']);
  const regId='REG-'+Utilities.getUuid().split('-')[0].toUpperCase();
  const cr=getAssignedCr_(batch);
  const reg=[new Date(),regId,get(['Timestamp','Timestamp ']),get(['Email Address','Email']),get(['Name of attendee','Name']),get(['Roll number','Roll No','Roll number']),get(['Phone number','Phone','Mobile']),batch,get(['Money paid to','Paid to']),get(['UTR number','UTR']),get(['Payment screenshot','Payment Screenshot']),get(['Amount','Amount paid']),cr,'PENDING','',''];
  const out=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG);
  out.appendRow(reg);
  return regId;
}

function getAssignedCr_(batch){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.USERS); if(!sh)return '';
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][2])==='CR'&&String(rows[i][3])===String(batch)&&String(rows[i][6])!=='false') return rows[i][1]||rows[i][0];
  return '';
}

function hash_(s){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}
function newToken_(){ return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); }
function getConfig_(key){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.CONFIG); if(!sh)return '';
  const rows=sh.getDataRange().getValues(); for(let i=1;i<rows.length;i++) if(String(rows[i][0])===key)return String(rows[i][1]); return '';
}
function setConfig_(key,value){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.CONFIG); const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++) if(String(rows[i][0])===key){sh.getRange(i+1,2).setValue(value);return;}
  sh.appendRow([key,value]);
}

function login(username,pin){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.USERS); if(!sh)return {ok:false,message:'System is not set up.'};
  const rows=sh.getDataRange().getValues();
  const u=String(username||'').trim();
  const p=String(pin||'');
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===u && String(rows[i][6])!=='false'){
      if(!rows[i][5]){sh.getRange(i+1,6).setValue(hash_(p));}
      if(String(rows[i][5]||hash_(p))!==hash_(p)) return {ok:false,message:'Invalid access code.'};
      const token=newToken_(); CacheService.getScriptCache().put('S:'+token,JSON.stringify({username:u,name:rows[i][1],role:rows[i][2],batch:rows[i][3],gate:rows[i][4]}),21600);
      return {ok:true,token,role:rows[i][2],name:rows[i][1],batch:rows[i][3],gate:rows[i][4]};
    }
  }
  return {ok:false,message:'Invalid username or access code.'};
}

function adminLogin(password){
  if(hash_(password)!==hash_(getConfig_('ADMIN_PASSWORD'))) return {ok:false,message:'Invalid admin password.'};
  const token=newToken_(); CacheService.getScriptCache().put('S:'+token,JSON.stringify({username:'ADMIN',name:'Administrator',role:'ADMIN'}),21600);
  return {ok:true,token,role:'ADMIN',name:'Administrator'};
}
function session_(token,role){const s=CacheService.getScriptCache().get('S:'+token); if(!s)throw new Error('Session expired. Sign in again.'); const x=JSON.parse(s); if(role&&x.role!==role)throw new Error('Not authorized.'); return x;}

function crList(token){const me=session_(token,'CR'); const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG); const rows=sh.getDataRange().getDisplayValues(); const out=[]; for(let i=1;i<rows.length;i++) if(String(rows[i][7])===String(me.batch)) out.push(rowObj_(rows[i],i+1)); return out.reverse().slice(0,200);}
function rowObj_(r,row){return {row,name:r[4],rollNumber:r[5],mobile:r[6],batch:r[7],paidTo:r[8],utr:r[9],screenshot:r[10],amount:r[11],crName:r[12],paymentStatus:r[13],ticketId:r[14],registrationId:r[1],timestamp:r[2],sheetRow:row};}

function verifyPayment(token,registrationId,decision,note){
  const me=session_(token,'CR'); const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG); const rows=sh.getDataRange().getValues();
    for(let i=1;i<rows.length;i++) if(String(rows[i][1])===String(registrationId)){
      if(String(rows[i][7])!==String(me.batch)) throw new Error('This registration is outside your batch.');
      if(String(rows[i][13])!=='PENDING') return {ok:false,message:'Already processed.'};
      const verified=decision==='VERIFY'; sh.getRange(i+1,14).setValue(verified?'VERIFIED':'REJECTED'); sh.getRange(i+1,16).setValue(note||'');
      if(verified){
        const ticketId='INS-'+Utilities.getUuid().split('-')[0].toUpperCase(); const tokenQR=newToken_();
        const t=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.TICK); t.appendRow([new Date(),ticketId,rows[i][1],rows[i][4],rows[i][5],rows[i][6],rows[i][7],tokenQR,true,'UNUSED','UNUSED','UNUSED',new Date()]);
        sh.getRange(i+1,15).setValue(ticketId);
        return {ok:true,message:'Payment verified. Ticket generated.',ticketId};
      }
      return {ok:true,message:'Registration rejected.'};
    }
    throw new Error('Registration not found.');
  } finally {lock.releaseLock();}
}

function studentTicket(mobile,utr){
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG); const rows=sh.getDataRange().getDisplayValues();
  let reg=null; for(let i=1;i<rows.length;i++) if(norm_(rows[i][6])===norm_(mobile)&&norm_(rows[i][9])===norm_(utr)) reg=rows[i];
  if(!reg) return {ok:false,message:'No matching registration found.'};
  if(reg[13]!=='VERIFIED'||!reg[14]) return {ok:false,message:'Payment is not verified yet. Your pass will appear here after CR verification.'};
  const t=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.TICK); const tr=t.getDataRange().getDisplayValues();
  for(let i=1;i<tr.length;i++) if(tr[i][1]===reg[14]) return {ok:true,ticket:{ticketId:tr[i][1],name:tr[i][3],rollNumber:tr[i][4],batch:tr[i][6],qrToken:tr[i][7],nights:tr[i].slice(9,12)}};
  return {ok:false,message:'Ticket record not found. Contact the admin.'};
}
function norm_(s){return String(s||'').replace(/\s+/g,'').trim().toLowerCase();}

function validateQr(token,qrToken,night,gate){
  const me=session_(token,'SCANNER'); if(me.gate!==gate) throw new Error('This scanner is not assigned to this gate.');
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.TICK); const rows=sh.getDataRange().getDisplayValues();
  const idx=APP.NIGHTS.indexOf(night); if(idx<0) throw new Error('Invalid night.');
  for(let i=1;i<rows.length;i++) if(rows[i][7]===qrToken){
    const active=String(rows[i][8])==='true'; const status=rows[i][9+idx];
    if(!active) return {ok:false,result:'INVALID',message:'Ticket is inactive.'};
    if(status==='USED') return {ok:false,result:'ALREADY_USED',message:`Already used for ${night}.`,ticketId:rows[i][1],name:rows[i][3]};
    return {ok:true,result:'VALID',message:'Valid ticket. Check college ID, then confirm entry.',ticketId:rows[i][1],name:rows[i][3],batch:rows[i][6],night};
  }
  return {ok:false,result:'INVALID',message:'QR not recognized.'};
}
function confirmEntry(token,qrToken,night,gate,note){
  const me=session_(token,'SCANNER'); if(me.gate!==gate) throw new Error('Gate mismatch.'); const idx=APP.NIGHTS.indexOf(night); if(idx<0)throw new Error('Invalid night.');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.TICK); const rows=sh.getDataRange().getValues();
    for(let i=1;i<rows.length;i++) if(String(rows[i][7])===String(qrToken)){
      const col=10+idx; if(String(rows[i][col])==='USED') return {ok:false,result:'ALREADY_USED',message:'This pass has already been used tonight.'};
      sh.getRange(i+1,col).setValue('USED');
      SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.LOG).appendRow([new Date(),rows[i][1],rows[i][2],night,gate,me.username,'ENTRY_CONFIRMED',me.name,note||'College ID checked']);
      return {ok:true,message:'ENTRY CONFIRMED',ticketId:rows[i][1],name:rows[i][3]};
    }
    return {ok:false,result:'INVALID',message:'QR not recognized.'};
  } finally {lock.releaseLock();}
}

function dashboard(token){const me=session_(token,'ADMIN'); const reg=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG).getDataRange().getDisplayValues(); const tick=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.TICK).getDataRange().getDisplayValues(); const logs=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.LOG).getDataRange().getDisplayValues(); let p=0,v=0,r=0; for(let i=1;i<reg.length;i++){if(reg[i][13]==='PENDING')p++;else if(reg[i][13]==='VERIFIED')v++;else if(reg[i][13]==='REJECTED')r++;} return {stats:{registrations:Math.max(0,reg.length-1),pending:p,verified:v,rejected:r,tickets:Math.max(0,tick.length-1),entries:Math.max(0,logs.length-1)},rows:reg.slice(1).reverse().slice(0,150).map((x,i)=>rowObj_(x,reg.length-i-1))};}
function listUsers(token){session_(token,'ADMIN'); const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.USERS); return sh.getDataRange().getDisplayValues().slice(1).map(r=>({username:r[0],name:r[1],role:r[2],batch:r[3],gate:r[4],active:r[6]}));}
function createUser(token,data){session_(token,'ADMIN'); const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.USERS); if(!data.username||!data.pin)throw new Error('Username and access code are required.'); const rows=sh.getDataRange().getValues(); if(rows.slice(1).some(r=>String(r[0])===String(data.username)))throw new Error('Username already exists.'); sh.appendRow([data.username,data.name||data.username,data.role,data.batch||'',data.gate||'',hash_(data.pin),true,new Date()]); return {ok:true};}
function setNight(token,night){session_(token,'ADMIN'); if(APP.NIGHTS.indexOf(night)<0)throw new Error('Invalid night.'); setConfig_('CURRENT_NIGHT',night); return {ok:true,night};}
function changeAdminPassword(token,newPassword){session_(token,'ADMIN'); if(!newPassword||newPassword.length<8)throw new Error('Use at least 8 characters.'); setConfig_('ADMIN_PASSWORD',newPassword); return {ok:true};}
function exportCsv(token){session_(token,'ADMIN'); const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.REG); return sh.getDataRange().getDisplayValues();}
