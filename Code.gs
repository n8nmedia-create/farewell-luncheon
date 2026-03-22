// ================================================================
// FAREWELL LUNCHEON — Google Apps Script (Code.gs)
// ================================================================
// After editing: Deploy → Manage Deployments → edit → New version → Deploy
// ================================================================

// ── ⚙️ CONFIG — edit these values ────────────────────────────
const SHEET_ID        = '11GDotTC3Pu8oZdTcvLg0Oe4sKPhnK69tZvrrT7N3G9s';
const ADMIN_EMAIL     = 'n8n.media@gmail.com';
const ORGANIZER_EMAIL = 'hyonok.mattis.naf@army.mil';
const DASHBOARD_URL   = 'https://n8nmedia-create.github.io/farewell-luncheon/dashboard.html';

const ORDERS_TAB      = 'ORDERS';
const ORGANIZERS_TAB  = 'ORGANIZERS';

// ── 📅 Event details ──────────────────────────────────────────
const EVENT = {
  honoree : 'Dominique',
  date    : 'Tuesday, 31 March',
  time    : '11:00 AM',
  location: 'Bakery Café 292LU',
  waze    : 'https://waze.com/ul/hwyd66fk6f',
  contact : 'Kimmy – hyonok.mattis.naf@army.mil',
  farewell: 'We look forward to gathering with you to say a warm and memorable farewell to Dominique as she begins her next chapter.'
};

// ── Response helper ───────────────────────────────────────────
function makeResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── GET router ────────────────────────────────────────────────
function doGet(e) {
  const action=e.parameter.action, token=e.parameter.token;
  // Serve dashboard HTML with data pre-injected (bypasses CORS entirely)
  if (action==='dashboard' || !action) {
    if (!verifyToken(token)) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;color:#c0392b">Invalid or missing token. Please use the link sent to you.</p>');
    }
    const tokenInfo = getTokenInfo(token);
    const ordersData = getOrders();
    const html = buildDashboardHtml(token, tokenInfo.email||'', ordersData.orders||[]);
    return HtmlService.createHtmlOutput(html)
      .setTitle('Farewell Luncheon — Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action==='get_orders')   { if(!verifyToken(token)) return makeResponse({error:'Unauthorized'}); return makeResponse(getOrders()); }
  if (action==='verify_token') return makeResponse(getTokenInfo(token));
  return makeResponse({error:'Unknown action'});
}

function buildDashboardHtml(token, email, orders) {
  // Read dashboard.html from script properties or use inline version
  // We inject the data directly so no fetch needed
  const ordersJson = JSON.stringify(orders);
  const emailJson = JSON.stringify(email);
  const tokenJson = JSON.stringify(token);
  const asuJson = JSON.stringify(ScriptApp.getService().getUrl());
  
  // Return a bootstrap page that loads the real dashboard from GitHub
  // but with data pre-injected via postMessage after load
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Farewell Luncheon — Dashboard</title></head><body style="margin:0;padding:0;background:#1a1a2e;min-height:100vh;display:flex;align-items:center;justify-content:center;font-family:sans-serif;">' +
    '<script>' +
    'var INJECTED={token:'+tokenJson+',email:'+emailJson+',orders:'+ordersJson+',asu:'+asuJson+'};' +
    // Redirect to GitHub Pages dashboard with token - but inject data via sessionStorage
    'try{sessionStorage.setItem("luncheon_data",JSON.stringify(INJECTED));}catch(e){}' +
    'window.location.href="https://n8nmedia-create.github.io/farewell-luncheon/dashboard.html?token="+INJECTED.token+"&from=gas";' +
    '</script>' +
    '<p style="color:#b8873a">Redirecting to dashboard…</p>' +
    '</body></html>';
}

// ── POST router ───────────────────────────────────────────────
function doPost(e) {
  try {
    const data=JSON.parse(e.postData.contents);
    if (data.action==='submit_order')        return makeResponse(submitOrder(data));
    if (data.action==='toggle_paid')         return makeResponse(togglePaid(data));
    if (data.action==='send_reminders')      return makeResponse(sendReminders(data));
    if (data.action==='email_kitchen_sheet') return makeResponse(emailKitchenSheet(data));
    if (data.action==='get_orders')          { if(!verifyToken(data.token)) return makeResponse({error:'Unauthorized'}); return makeResponse(getOrders()); }
    if (data.action==='verify_token')        return makeResponse(getTokenInfo(data.token));
    return makeResponse({error:'Unknown action'});
  } catch(err) { return makeResponse({error:err.toString()}); }
}

// ── Submit order ──────────────────────────────────────────────
function submitOrder(data) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName(ORDERS_TAB);
  if (!sheet) sheet=ss.insertSheet(ORDERS_TAB);
  if (sheet.getLastRow()===0) {
    const h=['Timestamp','Name','Email','Phone','Items (JSON)','Total (KRW)','Payment Status','MYP Details'];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date().toLocaleString('ko-KR'),data.name||'',data.email||'',data.phone||'',JSON.stringify(data.items||[]),data.total||0,'Unpaid',data.mypDetails||'']);
  sheet.autoResizeColumns(1,8);
  try { sendGuestConfirmation(data); }    catch(e){ Logger.log('Guest email err: '+e); }
  try { sendOrganizerNotification(data); } catch(e){ Logger.log('Organizer email err: '+e); }
  return {success:true};
}

// ── Guest confirmation email ──────────────────────────────────
function sendGuestConfirmation(data) {
  const items=data.items||[], total=data.total||0;
  const itemRows=items.map(i=>`<tr><td style="padding:7px 14px;border-bottom:1px solid #f0e8dc;font-size:13px">${i.name}${i.qty>1?' &times;'+i.qty:''}</td><td style="padding:7px 14px;border-bottom:1px solid #f0e8dc;text-align:right;font-weight:600;font-size:13px">&#8361;${Number(i.subtotal||i.price*i.qty).toLocaleString('ko-KR')}</td></tr>`).join('');
  const mypNote=data.mypDetails?`<div style="margin-top:10px;padding:10px 14px;background:#fff8ec;border-left:3px solid #b8873a;border-radius:4px;font-size:12px;color:#7a6f63">🍽️ ${data.mypDetails}</div>`:'';
  const html=`<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0ede8;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:560px;margin:32px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,.09)">
  <div style="background:#1c1917;padding:32px 28px;text-align:center;position:relative">
    <div style="font-size:10px;letter-spacing:.2em;text-transform:uppercase;color:#b8873a;margin-bottom:8px">FAREWELL LUNCHEON</div>
    <h1 style="font-family:Georgia,serif;font-size:28px;font-weight:300;color:#faf7f2;margin:0">Order <em style="color:#b8873a">Confirmed</em></h1>
  </div>
  <div style="padding:28px">
    <div style="font-size:13px;color:#7a6f63;line-height:1.75;font-style:italic;background:#f5f0e8;padding:14px 18px;border-left:3px solid #b8873a;border-radius:4px;margin-bottom:24px">${EVENT.farewell}</div>
    <p style="font-size:14px;color:#1c1917;margin-bottom:6px">Hi <strong>${data.name}</strong>,</p>
    <p style="font-size:13px;color:#7a6f63;margin-bottom:22px;line-height:1.6">Your pre-order has been received! Please bring the <strong>exact cash amount</strong> on the day — all funds will be collected for a single group payment.</p>
    <h3 style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#7a6f63;margin-bottom:10px">Your Order</h3>
    <table style="width:100%;border-collapse:collapse;background:#faf7f2;border-radius:6px;overflow:hidden;margin-bottom:10px">${itemRows}<tr style="background:#1c1917"><td style="padding:11px 14px;color:#faf7f2;font-weight:600;font-size:13px">Total Due</td><td style="padding:11px 14px;color:#b8873a;font-weight:700;text-align:right;font-size:18px">&#8361;${Number(total).toLocaleString('ko-KR')}</td></tr></table>
    ${mypNote}
    <h3 style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#7a6f63;margin:22px 0 12px">Event Details</h3>
    <table style="width:100%;border-collapse:collapse;margin-bottom:24px">
      <tr><td style="padding:5px 0;width:26px;vertical-align:top">📆</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>Date:</strong> ${EVENT.date}</td></tr>
      <tr><td style="padding:5px 0;vertical-align:top">🕐</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>Time:</strong> ${EVENT.time}</td></tr>
      <tr><td style="padding:5px 0;vertical-align:top">📍</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>Location:</strong> ${EVENT.location}</td></tr>
      <tr><td style="padding:5px 0;vertical-align:top">🗺️</td><td style="padding:5px 0;font-size:13px"><a href="${EVENT.waze}" style="color:#b8873a;font-weight:500">Open in Waze →</a></td></tr>
      <tr><td style="padding:5px 0;vertical-align:top">📞</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>Contact:</strong> ${EVENT.contact}</td></tr>
    </table>
    <div style="text-align:center;margin-bottom:10px">
      <a href="https://www.google.com/calendar/render?action=TEMPLATE&text=Farewell+Luncheon+for+${encodeURIComponent(EVENT.honoree)}&dates=20260331T020000Z/20260331T060000Z&location=${encodeURIComponent(EVENT.location)}&details=${encodeURIComponent(EVENT.farewell)}" style="background:#1c1917;color:#faf7f2;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:13px;font-weight:500;display:inline-block">📅 Add to Calendar</a>
    </div>
    <p style="font-size:11px;color:#a09589;text-align:center;margin-top:18px">Questions? Contact ${EVENT.contact}</p>
  </div>
</div></body></html>`;
  MailApp.sendEmail({to:data.email,subject:`Order Confirmed — Farewell Luncheon for ${EVENT.honoree}`,htmlBody:html});
}

// ── Organizer/Admin notification email ────────────────────────
function sendOrganizerNotification(data) {
  const items=data.items||[], total=data.total||0;
  const itemRows=items.map(i=>`<tr><td style="padding:7px 14px;border-bottom:1px solid #f0e8dc;font-size:13px">${i.name}${i.qty>1?' &times;'+i.qty:''}</td><td style="padding:7px 14px;border-bottom:1px solid #f0e8dc;text-align:right;font-size:13px">&#8361;${Number(i.subtotal||i.price*i.qty).toLocaleString('ko-KR')}</td></tr>`).join('');
  const mypLine=data.mypDetails?`<div style="margin-top:10px;padding:10px 14px;background:#fff8ec;border-left:3px solid #b8873a;border-radius:4px;font-size:12px;color:#7a6f63">🍽️ ${data.mypDetails}</div>`:'';
  const html=`<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0ede8;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:560px;margin:32px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,.09)">
  <div style="background:#1c1917;padding:18px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
    <div><div style="font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:#b8873a;margin-bottom:2px">NEW ORDER RECEIVED</div><h2 style="font-family:Georgia,serif;font-size:20px;font-weight:400;color:#faf7f2;margin:0">Farewell Luncheon</h2></div>
    <a href="${DASHBOARD_URL}" style="background:#b8873a;color:#1c1917;text-decoration:none;padding:9px 18px;border-radius:5px;font-size:12px;font-weight:700;white-space:nowrap">View Dashboard →</a>
  </div>
  <div style="padding:24px">
    <table style="width:100%;border-collapse:collapse;margin-bottom:22px">
      <tr><td style="padding:5px 0;font-size:12px;color:#7a6f63;width:70px;font-weight:600">Name</td><td style="padding:5px 0;font-size:13px;color:#1c1917;font-weight:700">${data.name}</td></tr>
      <tr><td style="padding:5px 0;font-size:12px;color:#7a6f63;font-weight:600">Email</td><td style="padding:5px 0;font-size:13px"><a href="mailto:${data.email}" style="color:#b8873a">${data.email}</a></td></tr>
      <tr><td style="padding:5px 0;font-size:12px;color:#7a6f63;font-weight:600">Phone</td><td style="padding:5px 0;font-size:13px"><a href="tel:${data.phone}" style="color:#1c1917">${data.phone}</a></td></tr>
    </table>
    <h3 style="font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#7a6f63;margin-bottom:10px">Order</h3>
    <table style="width:100%;border-collapse:collapse;background:#faf7f2;border-radius:6px;overflow:hidden;margin-bottom:8px">${itemRows}<tr style="background:#1c1917"><td style="padding:11px 14px;color:#faf7f2;font-weight:600;font-size:13px">Total</td><td style="padding:11px 14px;color:#b8873a;font-weight:700;text-align:right;font-size:18px">&#8361;${Number(total).toLocaleString('ko-KR')}</td></tr></table>
    ${mypLine}
    <div style="text-align:center;margin-top:22px"><a href="${DASHBOARD_URL}" style="background:#1c1917;color:#faf7f2;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:13px;font-weight:500;display:inline-block">Open Organizer Dashboard →</a></div>
    <p style="font-size:11px;color:#a09589;text-align:center;margin-top:16px">Payment status: <strong>Unpaid</strong> — update via dashboard after collection.</p>
  </div>
</div></body></html>`;
  const recipients=[ADMIN_EMAIL,ORGANIZER_EMAIL].filter(Boolean).join(',');
  MailApp.sendEmail({to:recipients,subject:`New Order: ${data.name} — \u20A9${Number(total).toLocaleString('ko-KR')}`,htmlBody:html});
}

// ── Manual reminder (from dashboard button) ───────────────────
function sendReminders(data) {
  if (!verifyToken(data.token)) return {error:'Unauthorized'};
  const orders=data.orders||[];
  let sent=0;
  orders.forEach(o=>{
    try { sendReminderEmail(o); sent++; } catch(e){ Logger.log('Reminder err for '+o.email+': '+e); }
  });
  return {success:true, sent};
}

// ── Automated event-day reminder (Time-based trigger) ─────────
// Call setupEventDayReminder() ONCE manually from the editor to schedule it
function setupEventDayReminder() {
  // Clear all existing reminder triggers
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(['sendEventDayReminders','sendDayBeforeReminders'].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });
  // March 31, 2026 at 10:00 AM KST (01:00 UTC)
  ScriptApp.newTrigger('sendEventDayReminders')
    .timeBased().at(new Date('2026-03-31T01:00:00Z')).create();
  // March 30, 2026 at 5:00 PM KST (08:00 UTC) — day before reminder
  ScriptApp.newTrigger('sendDayBeforeReminders')
    .timeBased().at(new Date('2026-03-30T08:00:00Z')).create();
  Logger.log('✅ Event-day reminder: March 31 at 10:00 AM KST');
  Logger.log('✅ Day-before reminder: March 30 at 5:00 PM KST');
}

function sendEventDayReminders() {
  const orders=getOrders().orders||[];
  let sent=0;
  orders.forEach(o=>{
    try { sendReminderEmail(o,'today'); sent++; } catch(e){ Logger.log('Day-of err: '+e); }
  });
  Logger.log('Sent '+sent+' event-day reminders.');
}

function sendDayBeforeReminders() {
  const orders=getOrders().orders||[];
  let sent=0;
  orders.forEach(o=>{
    try { sendReminderEmail(o,'tomorrow'); sent++; } catch(e){ Logger.log('Day-before err: '+e); }
  });
  Logger.log('Sent '+sent+' day-before reminders.');
}

function sendReminderEmail(o, when) { when=when||'today';
  const total=o.total||0;
  const items=Array.isArray(o.items)?o.items:[];
  const itemLines=items.map(i=>`${i.name}${i.qty>1?' ×'+i.qty:''} — ₩${Number(i.subtotal||i.price*i.qty).toLocaleString('ko-KR')}`).join('<br>');
  const isPaid=o.status==='Paid';
  const paymentNote=isPaid
    ? '<div style="background:#eafaf1;border-left:3px solid #27ae60;padding:10px 14px;border-radius:4px;font-size:13px;color:#155724;margin-bottom:18px">✅ <strong>Your payment is confirmed.</strong> See you today!</div>'
    : '<div style="background:#fef9e7;border-left:3px solid #e67e22;padding:10px 14px;border-radius:4px;font-size:13px;color:#856404;margin-bottom:18px">💵 <strong>Please bring ₩'+Number(total).toLocaleString('ko-KR')+'  in exact cash</strong> — payment collected on arrival.</div>';
  const html=`<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0ede8;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:540px;margin:32px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,.09)">
  <div style="background:#1c1917;padding:28px;text-align:center">
    <div style="font-size:22px;margin-bottom:6px">🎉</div>
    <div style="font-size:11px;letter-spacing:.15em;text-transform:uppercase;color:#b8873a;margin-bottom:6px">TODAY'S THE DAY</div>
    <h1 style="font-family:Georgia,serif;font-size:26px;font-weight:300;color:#faf7f2;margin:0">Farewell <em style="color:#b8873a">Luncheon</em></h1>
  </div>
  <div style="padding:26px">
    <p style="font-size:14px;color:#1c1917;margin-bottom:16px">Hi <strong>${o.name}</strong>, the luncheon is today! Here's a reminder of your order and event details.</p>
    ${paymentNote}
    <div style="background:#faf7f2;padding:14px 16px;border-radius:6px;font-size:13px;color:#1c1917;margin-bottom:18px;line-height:1.9">
      <strong style="font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:#7a6f63;display:block;margin-bottom:8px">Your Order</strong>
      ${itemLines}<br>
      <div style="border-top:1px solid #e2d9cc;margin-top:10px;padding-top:10px;font-weight:700">Total: ₩${Number(total).toLocaleString('ko-KR')}</div>
    </div>
    <table style="width:100%;border-collapse:collapse">
      <tr><td style="padding:5px 0;width:26px">📆</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>${EVENT.date}</strong></td></tr>
      <tr><td style="padding:5px 0">🕐</td><td style="padding:5px 0;font-size:13px;color:#1c1917"><strong>${EVENT.time}</strong></td></tr>
      <tr><td style="padding:5px 0">📍</td><td style="padding:5px 0;font-size:13px;color:#1c1917">${EVENT.location}</td></tr>
      <tr><td style="padding:5px 0">🗺️</td><td style="padding:5px 0;font-size:13px"><a href="${EVENT.waze}" style="color:#b8873a;font-weight:500">Open in Waze →</a></td></tr>
    </table>
    <p style="font-size:11px;color:#a09589;text-align:center;margin-top:20px">See you there! Questions? ${EVENT.contact}</p>
  </div>
</div></body></html>`;
  MailApp.sendEmail({to:o.email,subject:`Today! Farewell Luncheon for ${EVENT.honoree} — ${EVENT.time} at ${EVENT.location}`,htmlBody:html});
}

// ── Email kitchen sheet to organizers ────────────────────────
function emailKitchenSheet(data) {
  if (!verifyToken(data.token)) return {error:'Unauthorized'};
  const now = new Date().toLocaleString('en-US',{dateStyle:'full',timeStyle:'short'});
  const recipients = [ADMIN_EMAIL, ORGANIZER_EMAIL].filter(Boolean).join(',');
  const html = `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f0ede8;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:700px;margin:24px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,.09)">
  <div style="background:#1c1917;padding:20px 28px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
    <div>
      <div style="font-size:10px;letter-spacing:.15em;text-transform:uppercase;color:#b8873a;margin-bottom:2px">KITCHEN ORDER SHEET</div>
      <h2 style="font-family:Georgia,serif;font-size:20px;font-weight:400;color:#faf7f2;margin:0">Farewell Luncheon</h2>
    </div>
    <div style="font-size:11px;color:rgba(250,247,242,.5)">Generated: ${now}</div>
  </div>
  <div style="padding:24px">
    ${data.htmlContent}
  </div>
</div>
</body></html>`;

  MailApp.sendEmail({
    to      : recipients,
    subject : `Kitchen Order Sheet — Farewell Luncheon (${now})`,
    htmlBody: html
  });
  return {success:true};
}

// ── Get all orders ────────────────────────────────────────────
function getOrders() {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName(ORDERS_TAB);
  if(!sheet||sheet.getLastRow()<2) return {orders:[]};
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();
  return {orders:rows.filter(r=>r[1]!=='').map((r,i)=>({rowIndex:i+2,timestamp:r[0].toString(),name:r[1],email:r[2],phone:r[3],items:safeJSON(r[4]),total:Number(r[5])||0,status:r[6]||'Unpaid',myp:r[7]||''}))};
}

// ── Toggle paid ───────────────────────────────────────────────
function togglePaid(data) {
  if(!verifyToken(data.token)) return {error:'Unauthorized'};
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName(ORDERS_TAB);
  const cell=sheet.getRange(data.rowIndex,7);
  const next=cell.getValue()==='Paid'?'Unpaid':'Paid';
  cell.setValue(next);
  cell.setBackground(next==='Paid'?'#d4edda':'#fff3cd');
  cell.setFontColor(next==='Paid'?'#155724':'#856404');
  return {success:true,newStatus:next};
}

// ── Token verification ────────────────────────────────────────
function verifyToken(token) {
  if(!token)return false;
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName(ORGANIZERS_TAB);
  if(!sheet||sheet.getLastRow()<2)return false;
  return sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues().some(r=>r[1]===token&&r[2]==='Active');
}
function getTokenInfo(token) {
  if(!token)return {valid:false};
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName(ORGANIZERS_TAB);
  if(!sheet||sheet.getLastRow()<2)return {valid:false};
  const match=sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues().find(r=>r[1]===token&&r[2]==='Active');
  return match?{valid:true,email:match[0]}:{valid:false};
}

// ── Add organizer (run manually) ──────────────────────────────
function addOrganizerManual() {
  const EMAILS=['n8n.media@gmail.com','hyonok.mattis.naf@army.mil'];
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName(ORGANIZERS_TAB);
  if(!sheet){sheet=ss.insertSheet(ORGANIZERS_TAB);sheet.appendRow(['Email','Token','Status','Added']);sheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#FFFFFF');}
  const existing=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat():[];
  EMAILS.forEach(function(EMAIL){
    if(existing.includes(EMAIL)){Logger.log('Already exists: '+EMAIL);return;}
    const token=generateToken();
    sheet.appendRow([EMAIL,token,'Active',new Date().toLocaleDateString()]);
    Logger.log('Access granted: '+EMAIL);
    Logger.log('Dashboard link: '+DASHBOARD_URL+'?token='+token);
  });
  sheet.autoResizeColumns(1,4);
}

function revokeOrganizerManual() {
  const EMAIL='organizer@anyemail.com'; // ← change to email you want to revoke, then Run
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName(ORGANIZERS_TAB);
  if(!sheet||sheet.getLastRow()<2){Logger.log('No organizers.');return;}
  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues();
  for(let i=0;i<rows.length;i++){if(rows[i][0]===EMAIL){sheet.getRange(i+2,3).setValue('Revoked');Logger.log('🚫 Revoked: '+EMAIL);return;}}
  Logger.log('Not found: '+EMAIL);
}

function generateToken(){const c='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';let t='';for(let i=0;i<52;i++)t+=c[Math.floor(Math.random()*c.length)];return t;}
function safeJSON(s){try{return JSON.parse(s);}catch(e){return[];}}
