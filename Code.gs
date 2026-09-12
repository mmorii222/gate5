/** PC貸出管理。管理者が setup_ をエディタから実行してください。 */
const HEADERS = {
  devices: ['id','asset_no','model_name','device_type','status','purchased_at'],
  lendings: ['id','device_id','user_id','lent_at','due_date','returned_at','purpose'],
  employees: ['id','name','department','employment_status'],
  employee_accounts: ['email','user_id'],
  requests: ['id','user_id','kind','payload','state','created_at','expires_at','result']
};
const TZ = 'Asia/Tokyo';
const LIMIT = 3;
const TTL = 10 * 60 * 1000;
// 草案の要決定事項。以下は前版の暫定動作を維持しており、確定仕様ではない。
const PROVISIONAL = {purposeMax:100, trimPurpose:true, allowToday:true,
  rejectDayChange:true};
const ERROR_IDS = {REQUIRED:'E-01',LENGTH:'E-02',PAST:'E-04',DAY_CHANGED:'E-05',
  UNAVAILABLE:'E-06',INACTIVE:'E-07',LIMIT:'E-08',BUSY:'E-09',MISMATCH:'E-11',
  RETURNED:'E-12',OVERDUE:'E-13',UNCERTAIN:'E-15',NO_REQUEST:'E-17',
  NO_RESULT:'E-18',FORBIDDEN:'E-19'};

function doGet() {
  // GET は画面を表示するだけ。URL に何を指定しても登録しない。
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('PC貸出管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setup_() {
  const ss = book_();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = HEADERS[name];
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,h.length).setValues([h]);
    if (sh.getRange(1,1,1,h.length).getValues()[0].join('|') !== h.join('|'))
      throw new Error(name + ' の見出しが仕様と異なります。');
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#dbeafe');
    // BIGINT の丸めと入力値の自動変換を避ける。日時は ISO 文字列で保存。
    sh.getRange(2,1,Math.max(1,sh.getMaxRows()-1),h.length).setNumberFormat('@');
  });
}

function getHome() {
  return api_(() => locked_(() => {
    const db = read_(), me = identity_(db), today = day_();
    return {me:{id:me.id,name:me.name,department:me.department},today,
      input_rules:{purposeMax:PROVISIONAL.purposeMax,trimPurpose:PROVISIONAL.trimPurpose},
      devices: db.devices.rows.filter(d => d.device_type === 'LAPTOP' && d.status === 'AVAILABLE')
        .map(d => ({id:d.id,asset_no:d.asset_no,model_name:d.model_name})),
      lendings: active_(db,me.id).map(l => ({id:l.id,device_id:l.device_id,
        asset_no: one_(db.devices,l.device_id).asset_no,due_date:l.due_date,
        overdue:l.due_date < today}))};
  }));
}

/** 確認前に本人・入力・貸出条件を検査し、変更できない申請を保存する。 */
function prepareRequest(input) {
  return api_(() => locked_(() => {
    const db = read_(), me = identity_(db);
    require_(input && typeof input === 'object','INPUT','未入力の項目があります。');
    const id = requestId_(input.request_id);
    const p = normalize_(input,me);
    const old = db.requests.rows.find(r => r.id === id);
    if (old) {
      owned_(old,me);
      require_(old.kind === input.kind && old.payload === JSON.stringify(p),
        'CHANGED','同じ受付番号の申請内容は変更できません。');
      return view_(db,old,me);
    }
    healthy_(db);
    check_(db,me,input.kind,p,id);
    const now = new Date();
    const r = {id,user_id:me.id,kind:input.kind,payload:JSON.stringify(p),state:'PENDING',
      created_at:now.toISOString(),expires_at:new Date(now.getTime()+TTL).toISOString(),result:''};
    batch_([append_(db.requests,r)]);
    return view_(db,r,me);
  }));
}

function getRequest(id) {
  return api_(() => locked_(() => {
    const db = read_(), me = identity_(db), r = request_(db,id,me);
    return view_(db,r,me);
  }));
}

/** 確定時はクライアントの借用者や端末を信用せず保存済み申請を再検査する。 */
function confirmRequest(id) {
  return api_(() => locked_(() => {
    const db = read_(), me = identity_(db), r = request_(db,id,me);
    if (r.state === 'DONE') return completed_(db,r,me);
    healthy_(db);
    pending_(r);
    const p = JSON.parse(r.payload);
    const checked = check_(db,me,r.kind,p,r.id);
    const now = new Date().toISOString();
    let result, updates;
    if (r.kind === 'LEND') {
      const l = {id:nextId_(db.lendings.rows),device_id:p.device_id,user_id:me.id,
        lent_at:now,due_date:p.due_date,returned_at:'',purpose:p.purpose};
      result = {state:'DONE',kind:r.kind,request_id:r.id,lending_id:l.id,
        message:'貸出が完了しました。',warning:'',warning_id:''};
      updates = [append_(db.lendings,l),cell_(db.devices,checked.device,'status','LENT')];
    } else {
      result = {state:'DONE',kind:r.kind,request_id:r.id,lending_id:p.lending_id,
        message:'返却が完了しました。',warning:checked.warning,
        warning_id:checked.warning ? 'W-02' : ''};
      updates = [cell_(db.lendings,checked.lending,'returned_at',now),
        cell_(db.devices,checked.device,'status','AVAILABLE')];
    }
    // 結果不明のAPI実行を無条件に再送しないための永続ガード。
    // DONE は履歴・端末の変更と同じ atomic batch に含める。
    batch_([cell_(db.requests,r,'state','COMMITTING')]);
    updates.push(cell_(db.requests,r,'state','DONE'),cell_(db.requests,r,'result',JSON.stringify(result)));
    try { batch_(updates); }
    catch (e) {
      console.error(e);
      fail_('UNCERTAIN','処理結果を確認できません。同じ受付番号で再確認してください。確認中が続く場合は管理者に受付番号を伝えてください。');
    }
    return result;
  }));
}

function cancelRequest(id) {
  return api_(() => locked_(() => {
    const db = read_(), me = identity_(db), r = request_(db,id,me);
    healthy_(db);
    require_(r.state === 'PENDING','STATE','この申請は取り消せません。結果を再確認してください。');
    batch_([cell_(db.requests,r,'state','CANCELLED')]);
    return {state:'CANCELLED'};
  }));
}

function normalize_(i,me) {
  require_(i.user_id === me.id,'FORBIDDEN','この貸出情報を操作する権限がありません。');
  requiredField_(i.device_id,'device');
  const device = id_(i.device_id);
  if (i.kind === 'RETURN') return {device_id:device,lending_id:id_(i.lending_id)};
  require_(i.kind === 'LEND','INPUT','操作が不正です。');
  requiredField_(i.due_date,'due');
  const purpose = purpose_(i.purpose);
  return {device_id:device,due_date:date_(i.due_date),purpose,application_date:date_(i.application_date)};
}

function check_(db,me,kind,p,requestId) {
  // 保存済み申請も確定時に必須・形式を再検査する（R-07、R-08、R-10）。
  requiredField_(p.device_id,'device');
  id_(p.device_id);
  const device = db.devices.rows.find(d => d.id === p.device_id), today = day_();
  const current = active_(db,me.id);
  if (kind === 'LEND') {
    purpose_(p.purpose);
    requiredField_(p.due_date,'due');
    require_(device,'UNAVAILABLE','この端末は現在貸出不可能です。別の端末を選択してください。');
    require_(me.employment_status === 'ACTIVE','INACTIVE','この端末を借りる権限がありません。');
    // V-07の修正案は未確定。前版の「日またぎ自体を拒否」を維持。
    require_(!PROVISIONAL.rejectDayChange || p.application_date === today,'DAY_CHANGED',
      '日付が変わったため、返却予定日を確認してください。');
    const due = date_(p.due_date);
    fieldRequire_(PROVISIONAL.allowToday ? due >= today : due > today,
      'PAST','返却予定日に過去の日付は指定できません。','due');
    require_(device.device_type === 'LAPTOP' && device.status === 'AVAILABLE',
      'UNAVAILABLE','この端末は現在貸出不可能です。別の端末を選択してください。');
    require_(!db.lendings.rows.some(l => l.device_id === device.id && !l.returned_at),
      'INCONSISTENT','端末情報と貸出履歴が一致しません。管理者に連絡してください。');
    require_(current.length < LIMIT,'LIMIT','貸出できる上限台数を超えているため申請を受け付けられません。');
    require_(!current.some(l => date_(l.due_date) < today),'OVERDUE',
      '返却予定日を過ぎている端末があります。返却してから新規貸出を行ってください。');
    const reserved = db.requests.rows.some(r => r.id !== requestId && r.kind === 'LEND'
      && r.state === 'PENDING' && r.expires_at > new Date().toISOString()
      && JSON.parse(r.payload).device_id === p.device_id);
    require_(!reserved,'BUSY','現在他の利用者もこの端末を選択中です。他の端末を選ぶか、時間をおいてやり直してください。');
    return {device,warning:current.length ? '現在、貸出中のPCがあります。追加で借りる場合は『確定』を押してください。' : ''};
  }
  require_(kind === 'RETURN','INPUT','操作が不正です。');
  const lending = db.lendings.rows.find(l => l.id === p.lending_id);
  require_(lending,'MISMATCH','貸出中の端末の情報と一致しません。他の端末でお試しください。');
  require_(lending.user_id === me.id,'FORBIDDEN','この貸出情報を操作する権限がありません。');
  require_(device && lending.device_id === device.id,'MISMATCH','貸出中の端末の情報と一致しません。他の端末でお試しください。');
  require_(!lending.returned_at,'RETURNED','この端末は返却済みです。');
  require_(device.status === 'LENT' && db.lendings.rows.filter(l => l.device_id === device.id && !l.returned_at).length === 1,
    'INCONSISTENT','端末情報と貸出履歴が一致しません。管理者に連絡してください。');
  return {device,lending,warning:date_(lending.due_date) < today ?
    '返却が完了しました。返却予定日を過ぎています。次回は返却予定日までに返却してください。' : ''};
}

function view_(db,r,me) {
  if (r.state === 'DONE') return completed_(db,r,me);
  require_(r.state !== 'COMMITTING','UNCERTAIN','処理結果を確認中です。再確認しても変わらない場合は管理者に受付番号を伝えてください。');
  pending_(r);
  const p = JSON.parse(r.payload), checked = check_(db,me,r.kind,p,r.id);
  return {state:'PENDING',request_id:r.id,kind:r.kind,payload:p,
    asset_no:checked.device.asset_no,warning:r.kind === 'LEND' ? checked.warning : '',
    warning_id:r.kind === 'LEND' && checked.warning ? 'W-01' : '',expires_at:r.expires_at};
}

// R-20: requestsのDONEだけでなく、実際の貸出と本人を照合して表示する。
function completed_(db,r,me) {
  owned_(r,me);
  let result,p;
  try {result=JSON.parse(r.result);p=JSON.parse(r.payload);} catch(e) {
    fail_('NO_RESULT','貸出情報を確認できません。貸出状況を確認してください。');
  }
  const l = result && db.lendings.rows.find(l => l.id === result.lending_id);
  require_(l && p && l.user_id === me.id && l.device_id === p.device_id
    && result.request_id === r.id && result.kind === r.kind && result.state === 'DONE'
    && (r.kind === 'LEND' || (r.kind === 'RETURN' && l.id === p.lending_id && !!l.returned_at)),
    'NO_RESULT','貸出情報を確認できません。貸出状況を確認してください。');
  return {...result,duplicate:true};
}

function fieldRequire_(ok,code,message,field) {
  if (!ok) {const e=new Error(message);e.appCode=code;e.field=field;throw e;}
}
function requiredField_(value,field) {
  fieldRequire_(typeof value === 'string' && value.length > 0,'REQUIRED','未入力の項目があります。',field);
}
function purpose_(value) {
  requiredField_(value,'purpose');
  const p = PROVISIONAL.trimPurpose ? value.trim() : value;
  fieldRequire_(p.length > 0,'REQUIRED','未入力の項目があります。','purpose');
  fieldRequire_(Array.from(p).length <= PROVISIONAL.purposeMax,'LENGTH',
    '値が長すぎます。もう一度入力しなおしてください。','purpose');
  return p;
}

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  require_(id,'CONFIG','管理者がSPREADSHEET_IDを設定してください。');
  return SpreadsheetApp.openById(id);
}

function read_() {
  const ss = book_(), db = {};
  Object.keys(HEADERS).forEach(name => {
    const sh = ss.getSheetByName(name), headers = HEADERS[name];
    require_(sh,'CONFIG',name+' シートがありません。');
    const values = sh.getDataRange().getValues();
    require_(headers.every((h,i) => values[0][i] === h),'CONFIG',name+' の見出しが違います。');
    const rows = values.slice(1).map((v,i) => {
      const o = {_row:i+2};
      headers.forEach((h,j) => {
        let value = v[j];
        if (value instanceof Date) value = /^(due_date|purchased_at)$/.test(h)
          ? Utilities.formatDate(value,TZ,'yyyy-MM-dd') : value.toISOString();
        if (/^(id|user_id|device_id)$/.test(h) && typeof value === 'number')
          require_(Number.isSafeInteger(value),'DATA','IDは文字列で登録してください。数値精度が不足しています。');
        o[h] = value === null || value === undefined ? '' : String(value);
      });
      return o;
    }).filter(o => headers.some(h => o[h] !== ''));
    const key = name === 'employee_accounts' ? 'email' : 'id';
    const seen = new Set();
    rows.forEach(o => {
      const k = key === 'email' ? o[key].trim().toLowerCase() : o[key];
      require_(k && !seen.has(k),'DATA',name+' のキーが未入力または重複しています。');
      seen.add(k);
      if (['devices','employees','lendings'].includes(name)) id_(o.id);
      if (name === 'lendings') { id_(o.device_id); id_(o.user_id); date_(o.due_date); }
    });
    db[name] = {sheetId:sh.getSheetId(),headers,rows};
  });
  return db;
}

function identity_(db) {
  const email = Session.getActiveUser().getEmail().trim().toLowerCase();
  require_(email,'AUTH','ログインユーザーを確認できません。社内Googleアカウントで開き、管理者に公開設定を確認してください。');
  const account = db.employee_accounts.rows.find(a => a.email.trim().toLowerCase() === email);
  const me = account && db.employees.rows.find(e => e.id === account.user_id);
  require_(me,'AUTH','このGoogleアカウントに対応する社員が登録されていません。');
  return me;
}
function active_(db,id) { return db.lendings.rows.filter(l => l.user_id === id && !l.returned_at); }
function one_(table,id) {
  const row = table.rows.find(r => r.id === id);
  require_(row,'NOT_FOUND','対象データがありません。'); return row;
}
function owned_(r,me) { require_(r.user_id === me.id,'FORBIDDEN','この貸出情報を操作する権限がありません。'); }
function request_(db,id,me) {
  const r = db.requests.rows.find(r => r.id === requestId_(id));
  require_(r,'NO_REQUEST','申請内容がありません。貸出申請画面から入力してください。');
  owned_(r,me); return r;
}
function pending_(r) {
  require_(r.state === 'PENDING','STATE','この申請は無効です。申請画面から入力してください。');
  require_(r.expires_at > new Date().toISOString(),'EXPIRED','確認の有効期限が切れました。申請を取り消し、入力しなおしてください。');
}
function healthy_(db) {
  require_(!db.requests.rows.some(r => r.state === 'COMMITTING'),'UNCERTAIN',
    '結果確認中の処理があります。時間をおいて再確認してください。解消しない場合は管理者に連絡してください。');
}
function id_(s) {
  require_(typeof s === 'string' && /^[1-9][0-9]{0,18}$/.test(s)
    && (s.length < 19 || s <= '9223372036854775807'),'INPUT','IDが未入力または不正です。'); return s;
}
function requestId_(s) {
  require_(typeof s === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(s),
    'INPUT','受付番号が不正です。'); return s;
}
function date_(s) {
  require_(typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s),'DATE','日付を入力してください。');
  const d = new Date(s+'T00:00:00Z');
  require_(!isNaN(d.getTime()) && d.toISOString().slice(0,10) === s && s >= '1900-01-01','DATE','日付が不正です。'); return s;
}
function day_() { return Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'); }
function nextId_(rows) {
  let max = '0';
  rows.forEach(r => { id_(r.id); if (r.id.length > max.length || (r.id.length === max.length && r.id > max)) max = r.id; });
  const a = max.split(''); let carry = 1;
  for (let i=a.length-1;i>=0 && carry;i--) { const n=Number(a[i])+carry; a[i]=String(n%10); carry=n>9?1:0; }
  if (carry) a.unshift('1'); return id_(a.join(''));
}
function cell_(table,row,column,value) {
  return {updateCells:{start:{sheetId:table.sheetId,rowIndex:row._row-1,columnIndex:table.headers.indexOf(column)},
    rows:[{values:[value_(value)]}],fields:'userEnteredValue'}};
}
function value_(v) { return {userEnteredValue:{stringValue:String(v)}}; }
function append_(table,row) {
  return {appendCells:{sheetId:table.sheetId,rows:[{values:table.headers.map(h => value_(row[h] || ''))}],fields:'userEnteredValue'}};
}
function batch_(requests) { Sheets.Spreadsheets.batchUpdate({requests},book_().getId()); }
function locked_(fn) {
  const lock = LockService.getScriptLock();
  require_(lock.tryLock(5000),'BUSY','他の処理を実行中です。時間をおいて再度操作してください。');
  try { return fn(); } finally { lock.releaseLock(); }
}
function require_(ok,code,message) { if (!ok) fail_(code,message); }
function fail_(code,message) { const e = new Error(message); e.appCode=code; throw e; }
function api_(fn) {
  try { return {ok:true,data:fn()}; }
  catch(e) { console.error(e); return {ok:false,code:e.appCode || 'SYSTEM',
    error_id:ERROR_IDS[e.appCode] || '',field:e.field || '',
    redirect:['NO_REQUEST','NO_RESULT'].includes(e.appCode) ? 'home' : '',
    message:e.appCode ? e.message : '処理結果を確認できません。同じ受付番号で再確認してください。改善しない場合は管理者に連絡してください。'}; }
}
