/**
 * ══════════════════════════════════════════════════════════
 *  FLENBO HUB — Google Apps Script Backend  v3.0
 *  Flenbo FoodWorks Pvt. Ltd.
 * ══════════════════════════════════════════════════════════
 *
 *  SHEETS (auto-created by setupSpreadsheet()):
 *  ▸ Staff          — user accounts & passwords
 *  ▸ Locations      — location list
 *  ▸ Requisitions   — ration orders (line items)
 *  ▸ Leave Requests — leave applications
 *  ▸ Attendance     — daily check in/out
 *  ▸ GRN            — goods received notes
 *  ▸ Wastage        — wastage log
 *  ▸ Error Log
 *
 *  ALL endpoints via GET (?action=X&data=JSON or params)
 *  to avoid CORS preflight issues.
 *
 *  SETUP:
 *  1. script.google.com → New project → paste this code
 *  2. Run setupSpreadsheet() once from the editor
 *  3. Deploy → New deployment → Web app
 *     Execute as: Me | Who has access: Anyone
 *  4. Copy Web App URL → paste as BACKEND_URL in Flenbo_Hub.html
 * ══════════════════════════════════════════════════════════
 */

const ADMIN_EMAIL  = 'nivedita.a.sahay@gmail.com';
const SH_STAFF     = 'Staff';
const SH_LOCATIONS = 'Locations';
const SH_ORDERS    = 'Requisitions';
const SH_LEAVES    = 'Leave Requests';
const SH_ATTEND    = 'Attendance';
const SH_GRN       = 'GRN';
const SH_WASTAGE   = 'Wastage';
const SH_ITEMS     = 'Custom Items';
const SH_LOG       = 'Error Log';

// ── RATION CATEGORY SHEETS ────────────────────────────────
const SH_HP        = 'HYPERPURE';
const SH_BEV       = 'BEVERAGES';
const SH_FRZ       = 'FROZEN';
const SH_VEG       = 'VEGETABLES';
const SH_DAIRY2    = 'DAIRY';
const SH_CLN       = 'CLEANING';
const SH_DISP      = 'DISPOSABLES';
const SH_SR        = 'STAFF RATION';
const RATION_CATS  = ['HYPERPURE','BEVERAGES','FROZEN','VEGETABLES','DAIRY','CLEANING','DISPOSABLES','BIKANERVALA','STAFF RATION'];
const HDR_RATION   = ['Item #','Name','Unit','Added On','Added By'];

const HDR_STAFF     = ['User ID','Name','Profile','Username','Password','Role','Location','Phone','Join Date','Annual Leave','Casual Leave','Casual Used Month','Status'];
const HDR_LOCATIONS = ['Location ID','Name','Type','Address','Added On','Status'];
const HDR_ORDERS    = ['Order ID','Submitted At','Order Type','Event Name','Event Date','PAX','Event Client','Location','Staff Name','User ID','Delivery Date','Urgency','Supplier','Item','Quantity','Unit','Item Remarks','Order Remarks','Status','Manager Note'];
const HDR_LEAVES    = ['Leave ID','User ID','Staff Name','Location','Leave Type','From Date','To Date','Reason','Submitted At','Status','Manager Note','Decided At'];
const HDR_ATTEND    = ['Record ID','User ID','Staff Name','Location','Date','In Time','Out Time','Hours'];
const HDR_GRN       = ['GRN ID','Order ID','Location','Received By','Received At','Remarks','Item','Ordered Qty','Received Qty','Unit','Discrepancy','Discrepancy %'];
const HDR_WASTAGE   = ['Wastage ID','Location','Logged By','Logged At','Remarks','Item','Quantity','Unit','Reason','Supplier'];
const HDR_ITEMS     = ['Item ID','Supplier','Category','Name','Unit','Added On','Added By'];

// ── SPREADSHEET HELPER ───────────────────────────────────
// Works for both standalone and container-bound scripts.
// On first run (setupSheets), creates a new Sheet and stores its ID.
function getSS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // Hardcoded fallback — standalone Web App never has an active spreadsheet
  return SpreadsheetApp.openById('176PL952UBZ2wY6uzrHxepT1KDo_uVLXTqG1fcZmACeM');
}


// ── ENTRY ────────────────────────────────────────────────
function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || '';
    const data   = p.data   ? JSON.parse(decodeURIComponent(p.data)) : {};
    switch(action) {
      case 'login':           return handleLogin(p);
      case 'submit':          return handleSubmit(data);
      case 'submitGRN':       return handleGRN(data);
      case 'logWastage':      return handleWastage(data);
      case 'submitLeave':     return handleSubmitLeave(data);
      case 'updateLeave':     return handleUpdateLeave(data);
      case 'logAttendance':   return handleAttendance(data);
      case 'updateOrderStatus': return handleUpdateOrderStatus(data);
      case 'markOrdered':     return handleMarkOrdered(data);
      case 'getOrders':       return handleGetOrders(p);
      case 'getLeaves':       return handleGetLeaves(p);
      case 'getAttendance':   return handleGetAttendance(p);
      case 'getGRN':          return handleGetGRN(p);
      case 'getWastage':      return handleGetWastage(p);
      case 'getStaff':        return handleGetStaff(p);
      case 'addStaff':        return handleAddStaff(data);
      case 'updateStaff':     return handleUpdateStaff(data);
      case 'getLocations':    return handleGetLocations(p);
      case 'addLocation':     return handleAddLocation(data);
      case 'getCustomItems':  return handleGetCustomItems();
      case 'addCustomItem':   return handleAddCustomItem(data);
      case 'deleteCustomItem':return handleDeleteCustomItem(data);
      case 'getBaseItems':    return handleGetBaseItems();
      case 'addBaseItem':     return handleAddBaseItem(data);
      case 'ping':            return ok({status:'online',time:now()});
      default:                return ok({status:'Flenbo Hub API v3.0'});
    }
  } catch(err) {
    logErr('doGet', err.toString());
    return ok({success:false, error:err.toString()});
  }
}

// ── LOGIN ─────────────────────────────────────────────────
function handleLogin(p) {
  const username = (p.username||'').toLowerCase().trim();
  const password = p.password||'';
  if (!username || !password) return ok({success:false, error:'Missing credentials'});
  const ss = getSS();
  const sheet = ss.getSheetByName(SH_STAFF);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:false, error:'No staff configured'});
  const rows = getRows(sheet);
  const user = rows.find(r =>
    String(r['Username']||'').toLowerCase() === username &&
    String(r['Password']||'') === password &&
    (r['Status']||'Active') === 'Active'
  );
  if (!user) return ok({success:false, error:'Invalid username or password'});
  return ok({success:true, user:{
    userId:   user['User ID'],
    name:     user['Name'],
    username: user['Username'],
    role:     user['Role'],
    location: user['Location'],
    phone:    user['Phone']||'',
  }});
}

// ── SUBMIT ORDER ──────────────────────────────────────────
function handleSubmit(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_ORDERS, HDR_ORDERS, '#1A2642');
  const orderId = d.orderId || genId(d.location);
  const subAt   = now();
  (d.items||[]).forEach(function(item) {
    sheet.appendRow([
      orderId, subAt,
      d.orderType||'Restaurant', d.eventName||'', d.eventDate||'', d.pax||'', d.eventClient||'',
      d.location||'', d.staffName||'', d.userId||'',
      d.deliveryDate||'', d.urgency||'Normal',
      item.supplier||item.sup||'', item.name||'', item.qty||0, item.unit||'',
      item.remarks||'', d.remarks||'', 'Pending', ''
    ]);
  });
  try { sendOrderEmail(d, orderId, subAt); } catch(e) { logErr('email',e.toString()); }
  return ok({success:true, orderId});
}

// ── UPDATE ORDER STATUS (Manager Approve/Reject) ──────────
function handleUpdateOrderStatus(d) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_ORDERS);
  if (!sheet) return ok({success:false, error:'Sheet not found'});
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ok({success:true, updated:0});
  const data    = sheet.getRange(2, 1, lastRow-1, HDR_ORDERS.length).getValues();
  const idCol   = HDR_ORDERS.indexOf('Order ID');
  const qtyCol  = HDR_ORDERS.indexOf('Quantity');
  const nameCol = HDR_ORDERS.indexOf('Item');
  const statCol = HDR_ORDERS.indexOf('Status');
  const noteCol = HDR_ORDERS.indexOf('Manager Note');
  const edits   = d.managerEdits  || {};
  const rejects = d.managerRejects|| {};
  const note    = d.managerNote   || '';
  var count = 0;
  data.forEach(function(row, i) {
    if (String(row[idCol]) !== String(d.orderId)) return;
    const itemKey = d.orderId + '|' + row[nameCol];
    if (rejects[itemKey]) {
      sheet.getRange(i+2, statCol+1).setValue('Item Rejected');
    } else {
      if (edits[itemKey] != null) sheet.getRange(i+2, qtyCol+1).setValue(edits[itemKey]);
      sheet.getRange(i+2, statCol+1).setValue(d.status||'Manager Approved');
    }
    sheet.getRange(i+2, noteCol+1).setValue(note);
    count++;
  });
  try { sendApprovalEmail(d); } catch(e) {}
  return ok({success:true, updated:count});
}

// ── MARK ORDERED ──────────────────────────────────────────
function handleMarkOrdered(d) {
  // d.items is array of "orderId+itemName" strings
  // Update status to 'Ordered' for those rows
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true});
  const data    = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ORDERS.length).getValues();
  const idCol   = HDR_ORDERS.indexOf('Order ID');
  const nameCol = HDR_ORDERS.indexOf('Item');
  const statCol = HDR_ORDERS.indexOf('Status');
  var count = 0;
  (d.items||[]).forEach(function(key) {
    data.forEach(function(row, i) {
      if ((String(row[idCol])+String(row[nameCol])) === key) {
        sheet.getRange(i+2, statCol+1).setValue('Ordered');
        count++;
      }
    });
  });
  return ok({success:true, updated:count});
}

// ── GRN ───────────────────────────────────────────────────
function handleGRN(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_GRN, HDR_GRN, '#0d47a1');
  const grnId = d.grnId || ('GRN-'+now().replace(/[^0-9]/g,'').slice(0,12));
  (d.items||[]).forEach(function(item) {
    const disc = parseFloat(item.receivedQty||0) - parseFloat(item.orderedQty||0);
    const pct  = item.orderedQty>0 ? Math.round(disc/item.orderedQty*100)+'%' : '—';
    sheet.appendRow([grnId, d.orderId||'', d.location||'', d.receivedBy||'', now(), d.remarks||'',
      item.name||'', item.orderedQty||0, item.receivedQty||0, item.unit||'', disc.toFixed(2), pct]);
  });
  updateOrderStatusById(d.orderId, 'GRN Filed');
  const shorts = (d.items||[]).filter(function(i){ return parseFloat(i.receivedQty||0) < parseFloat(i.orderedQty||0)-0.01; });
  if (shorts.length) { try { sendGRNAlert(d, grnId, shorts); } catch(e) {} }
  return ok({success:true, grnId});
}

// ── WASTAGE ───────────────────────────────────────────────
function handleWastage(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_WASTAGE, HDR_WASTAGE, '#c62828');
  const wId   = d.wastageId || ('WST-'+now().replace(/[^0-9]/g,'').slice(0,12));
  (d.items||[]).forEach(function(item) {
    sheet.appendRow([wId, d.location||'', d.loggedBy||'', now(), d.remarks||'',
      item.name||'', item.qty||0, item.unit||'', item.reason||'', item.supplier||'']);
  });
  return ok({success:true, wastageId:wId});
}

// ── LEAVE ─────────────────────────────────────────────────
function handleSubmitLeave(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_LEAVES, HDR_LEAVES, '#7b1fa2');
  const lId   = d.leaveId || ('LV-'+Date.now());
  sheet.appendRow([lId, d.userId||'', d.staffName||'', d.location||'', d.leaveType||'',
    d.fromDate||'', d.toDate||'', d.reason||'', now(), 'Pending', '', '']);
  try { sendLeaveEmail(d, lId); } catch(e) {}
  return ok({success:true, leaveId:lId});
}

function handleUpdateLeave(d) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_LEAVES);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:false, error:'Sheet not found'});
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_LEAVES.length).getValues();
  const idCol   = HDR_LEAVES.indexOf('Leave ID');
  const statCol = HDR_LEAVES.indexOf('Status');
  const noteCol = HDR_LEAVES.indexOf('Manager Note');
  const datCol  = HDR_LEAVES.indexOf('Decided At');
  var count = 0;
  rows.forEach(function(row, i) {
    if (String(row[idCol]) === String(d.leaveId)) {
      sheet.getRange(i+2, statCol+1).setValue(d.status||'Pending');
      sheet.getRange(i+2, noteCol+1).setValue(d.managerNote||'');
      sheet.getRange(i+2, datCol+1).setValue(now());
      count++;
    }
  });
  return ok({success:true, updated:count});
}

// ── ATTENDANCE ────────────────────────────────────────────
function handleAttendance(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_ATTEND, HDR_ATTEND, '#1a4731');
  // Check if row exists for this user+date
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ATTEND.length).getValues();
    const uidC = HDR_ATTEND.indexOf('User ID');
    const dateC= HDR_ATTEND.indexOf('Date');
    const inC  = HDR_ATTEND.indexOf('In Time');
    const outC = HDR_ATTEND.indexOf('Out Time');
    for (var i=0; i<rows.length; i++) {
      if (String(rows[i][uidC])===String(d.userId) && String(rows[i][dateC])===String(d.date)) {
        if (d.type==='out') {
          sheet.getRange(i+2, outC+1).setValue(d.outTime||'');
          // Calculate hours
          try{
            var inT=rows[i][inC];var outT=d.outTime;
            if(inT&&outT){var hrs=calcHours(inT,outT);sheet.getRange(i+2,HDR_ATTEND.indexOf('Hours')+1).setValue(hrs);}
          }catch(e){}
        }
        return ok({success:true, updated:true});
      }
    }
  }
  // New check-in
  const recId = 'ATT-'+d.userId+'-'+d.date.replace(/-/g,'');
  sheet.appendRow([recId, d.userId||'', d.staffName||'', d.location||'', d.date||'', d.inTime||'', d.outTime||'', '']);
  return ok({success:true, created:true});
}

function calcHours(inStr, outStr) {
  function toMins(s){ const p=s.match(/(\d+):(\d+)\s*(AM|PM)?/i);if(!p)return 0;var h=parseInt(p[1]),m=parseInt(p[2]);if(p[3]&&p[3].toUpperCase()==='PM'&&h!==12)h+=12;if(p[3]&&p[3].toUpperCase()==='AM'&&h===12)h=0;return h*60+m;}
  var diff=(toMins(outStr)-toMins(inStr))/60;
  return Math.max(0,diff).toFixed(1);
}

// ── STAFF MANAGEMENT ──────────────────────────────────────
function handleGetStaff() {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_STAFF);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, staff:[]});
  const staff = getRows(sheet).map(function(r) {
    return {userId:r['User ID'],name:r['Name'],username:r['Username'],password:r['Password'],role:r['Role'],location:r['Location'],phone:r['Phone']||'',status:r['Status']||'Active'};
  });
  return ok({success:true, staff});
}

function handleAddStaff(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_STAFF, HDR_STAFF, '#1A2642');
  // Check username unique
  if (sheet.getLastRow() > 1) {
    const rows = getRows(sheet);
    if (rows.find(r=>String(r['Username']).toLowerCase()===String(d.username).toLowerCase()))
      return ok({success:false, error:'Username already exists'});
  }
  const userId = 'U'+Date.now();
  sheet.appendRow([userId, d.name||'', d.profile||'', (d.username||'').toLowerCase(), d.password||'', d.role||'staff', d.location||'', d.phone||'', now().split(' ')[0], d.annualLeave||0, d.casualLeave||0, 0, d.status||'Active']);
  const staff = handleGetStaff().getContent();
  return ok({success:true, userId, staff: JSON.parse(staff).staff});
}

function handleUpdateStaff(d) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_STAFF);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:false});
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,HDR_STAFF.length).getValues();
  const idCol = HDR_STAFF.indexOf('User ID');
  rows.forEach(function(row,i) {
    if (String(row[idCol]) === String(d.userId)) {
      if (d.name)     sheet.getRange(i+2, HDR_STAFF.indexOf('Name')+1).setValue(d.name);
      if (d.password) sheet.getRange(i+2, HDR_STAFF.indexOf('Password')+1).setValue(d.password);
      if (d.role)     sheet.getRange(i+2, HDR_STAFF.indexOf('Role')+1).setValue(d.role);
      if (d.location) sheet.getRange(i+2, HDR_STAFF.indexOf('Location')+1).setValue(d.location);
      if (d.phone)    sheet.getRange(i+2, HDR_STAFF.indexOf('Phone')+1).setValue(d.phone);
      if (d.status)   sheet.getRange(i+2, HDR_STAFF.indexOf('Status')+1).setValue(d.status);
    }
  });
  const staff = handleGetStaff().getContent();
  return ok({success:true, staff: JSON.parse(staff).staff});
}

// ── LOCATIONS ─────────────────────────────────────────────
function handleGetLocations() {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_LOCATIONS);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, locations:[]});
  const locs = getRows(sheet).filter(r=>r['Status']!=='Inactive').map(r=>r['Name']);
  return ok({success:true, locations:locs});
}

function handleAddLocation(d) {
  const ss    = getSS();
  const sheet = getOrCreate(ss, SH_LOCATIONS, HDR_LOCATIONS, '#243254');
  const locId = 'LOC'+Date.now();
  sheet.appendRow([locId, d.name||'', d.type||'Restaurant', d.address||'', now().split(' ')[0], 'Active']);
  const locs = handleGetLocations().getContent();
  return ok({success:true, locId, locations: JSON.parse(locs).locations});
}

// ── CUSTOM ITEMS ──────────────────────────────────────────
function handleGetCustomItems() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SH_ITEMS);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, items:[]});
  var rows = getRows(sheet);
  var items = rows.map(function(r){
    return {sup:r['Supplier']||'CUSTOM', cat:r['Category']||'', name:r['Name']||'', unit:r['Unit']||'KG', custom:true};
  }).filter(function(i){return i.name;});
  return ok({success:true, items:items});
}

function handleAddCustomItem(d) {
  if (!d||!d.name) return ok({success:false, error:'Item name required'});
  var ss = getSS();
  var sheet = getOrCreate(ss, SH_ITEMS, HDR_ITEMS, '#5B21B6');
  // Check for duplicate
  if (sheet.getLastRow() > 1) {
    var existing = getRows(sheet);
    for (var i=0;i<existing.length;i++) {
      if ((existing[i]['Name']||'').toLowerCase() === (d.name||'').toLowerCase()) {
        return ok({success:false, error:'Item already exists'});
      }
    }
  }
  var itemId = 'ITEM'+Date.now();
  sheet.appendRow([itemId, d.sup||'CUSTOM', d.cat||'', d.name, d.unit||'KG', now().split(' ')[0], '']);
  return handleGetCustomItems();
}

function handleDeleteCustomItem(d) {
  if (!d||!d.name) return ok({success:false, error:'Item name required'});
  var ss = getSS();
  var sheet = ss.getSheetByName(SH_ITEMS);
  if (!sheet) return ok({success:true, items:[]});
  var rows = getRows(sheet);
  for (var i=rows.length-1;i>=0;i--) {
    if ((rows[i]['Name']||'').toLowerCase()===(d.name||'').toLowerCase()) {
      sheet.deleteRow(i+2); // +2 for header row and 0-index
      break;
    }
  }
  return handleGetCustomItems();
}

// ── BASE RATION ITEMS (from category sheets) ─────────────
function handleGetBaseItems() {
  var ss = getSS();
  var result = {};
  RATION_CATS.forEach(function(cat) {
    var sheet = ss.getSheetByName(cat);
    if (!sheet || sheet.getLastRow() < 2) { result[cat] = []; return; }
    var rows = sheet.getRange(2, 1, sheet.getLastRow()-1, HDR_RATION.length).getValues();
    result[cat] = rows
      .filter(function(r){ return r[1] && String(r[1]).trim(); })
      .map(function(r){ return [String(r[1]).trim(), String(r[2]||'KG').trim()]; });
  });
  return ok({success:true, items:result});
}

function handleAddBaseItem(d) {
  if (!d||!d.name||!d.cat) return ok({success:false, error:'Name and category required'});
  var sheetName = d.cat; // e.g. 'HYPERPURE', 'BEVERAGES', 'STAFF RATION'
  if (RATION_CATS.indexOf(sheetName) === -1) {
    // Unrecognised category — fall back to Custom Items
    return handleAddCustomItem(d);
  }
  var ss = getSS();
  var sheet = getOrCreate(ss, sheetName, HDR_RATION, '#1A2642');
  // Prevent duplicates
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2,2,sheet.getLastRow()-1,1).getValues().flat();
    if (existing.some(function(n){ return String(n).toLowerCase()===d.name.toLowerCase(); })) {
      return ok({success:false, error:'Item already exists in this category'});
    }
  }
  var num = sheet.getLastRow(); // Item # = row number
  sheet.appendRow([num, d.name, d.unit||'KG', now().split(' ')[0], d.addedBy||'']);
  return handleGetBaseItems();
}

// ── GET QUERIES ───────────────────────────────────────────
function handleGetOrders(p) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, orders:[]});
  var rows = getRows(sheet);
  if (p.status) rows = rows.filter(function(r){ return r['Status']===p.status; });
  if (p.location) rows = rows.filter(function(r){ return r['Location']===p.location; });
  if (p.userId) rows = rows.filter(function(r){ return r['User ID']===p.userId; });
  if (p.dateFrom) rows = rows.filter(function(r){ return fmtRowDate(r['Delivery Date'])>=p.dateFrom; });
  if (p.dateTo)   rows = rows.filter(function(r){ return fmtRowDate(r['Delivery Date'])<=p.dateTo; });
  return ok({success:true, orders:rows});
}

function handleGetLeaves(p) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_LEAVES);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, leaves:[]});
  var rows = getRows(sheet);
  if (p.userId)   rows = rows.filter(function(r){ return r['User ID']===p.userId; });
  if (p.status)   rows = rows.filter(function(r){ return r['Status']===p.status; });
  if (p.location) rows = rows.filter(function(r){ return r['Location']===p.location; });
  const mapped = rows.map(function(r){ return {leaveId:r['Leave ID'],userId:r['User ID'],staffName:r['Staff Name'],location:r['Location'],leaveType:r['Leave Type'],fromDate:r['From Date'],toDate:r['To Date'],reason:r['Reason'],status:r['Status'],managerNote:r['Manager Note']||'',submittedAt:r['Submitted At']}; });
  return ok({success:true, leaves:mapped});
}

function handleGetAttendance(p) {
  const ss    = getSS();
  const sheet = ss.getSheetByName(SH_ATTEND);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, records:[]});
  var rows = getRows(sheet);
  if (p.userId) rows = rows.filter(function(r){ return r['User ID']===p.userId; });
  if (p.date)   rows = rows.filter(function(r){ return fmtRowDate(r['Date'])===p.date; });
  if (p.location) rows = rows.filter(function(r){ return r['Location']===p.location; });
  const mapped = rows.map(function(r){ return {userId:r['User ID'],staffName:r['Staff Name'],location:r['Location'],date:r['Date'],inTime:r['In Time']||'',outTime:r['Out Time']||'',hours:r['Hours']||''}; });
  if (p.date && p.userId && mapped.length===1) return ok({success:true, record:mapped[0]});
  return ok({success:true, records:mapped});
}

function handleGetGRN(p) {
  const ss = getSS();
  const sheet = ss.getSheetByName(SH_GRN);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, grn:[]});
  return ok({success:true, grn:getRows(sheet)});
}

function handleGetWastage(p) {
  const ss = getSS();
  const sheet = ss.getSheetByName(SH_WASTAGE);
  if (!sheet || sheet.getLastRow() < 2) return ok({success:true, wastage:[]});
  return ok({success:true, wastage:getRows(sheet)});
}

// ── EMAIL NOTIFICATIONS ───────────────────────────────────
function sendOrderEmail(d, orderId, subAt) {
  const urgent = d.urgency === 'Urgent';
  const sub = (urgent?'🚨 URGENT — ':'')+'[Flenbo Hub] New Requisition — '+d.orderType+' | '+(d.location||'')+' | '+orderId;
  var body = 'ORDER ID   : '+orderId+'\nType       : '+(d.orderType||'')+'\nLocation   : '+(d.location||'')+'\nStaff      : '+(d.staffName||'')+'\nDelivery   : '+(d.deliveryDate||'')+'\nPriority   : '+(d.urgency||'Normal')+'\nSubmitted  : '+subAt;
  if (d.orderType==='Catering') body += '\nEvent      : '+(d.eventName||'')+'\nEvent Date : '+(d.eventDate||'')+'\nPAX        : '+(d.pax||'');
  body += '\n\nITEMS:\n';
  (d.items||[]).forEach(function(it){ body += '  • '+(it.name||it.supplier||'')+(it.name?' ('+it.supplier+')':'')+': '+(it.qty||0)+' '+(it.unit||'')+'\n'; });
  MailApp.sendEmail({to:ADMIN_EMAIL, subject:sub, body:body});
}

function sendApprovalEmail(d) {
  const sub = '[Flenbo Hub] Manager '+(d.status||'Updated')+' — Order '+d.orderId;
  const body = 'Order '+d.orderId+' has been '+(d.status||'updated')+' by Operations Manager.\n\nNote: '+(d.managerNote||'None');
  MailApp.sendEmail({to:ADMIN_EMAIL, subject:sub, body:body});
}

function sendLeaveEmail(d, leaveId) {
  const sub = '[Flenbo Hub] Leave Request — '+(d.staffName||'')+' | '+(d.leaveType||'')+' | '+leaveId;
  const body = 'Staff     : '+(d.staffName||'')+'\nLocation  : '+(d.location||'')+'\nType      : '+(d.leaveType||'')+'\nDates     : '+(d.fromDate||'')+' to '+(d.toDate||'')+'\nReason    : '+(d.reason||'')+'\n\nPlease review in Flenbo Hub.';
  MailApp.sendEmail({to:ADMIN_EMAIL, subject:sub, body:body});
}

function sendGRNAlert(d, grnId, shorts) {
  const sub = '⚠️ [Flenbo Hub] GRN Discrepancy — '+(d.location||'')+' | '+grnId;
  var body = 'Delivery shortage detected:\n\n';
  shorts.forEach(function(i){ body += '  ▼ '+(i.name||'')+': Ordered '+(i.orderedQty||0)+', Received '+(i.receivedQty||0)+'\n'; });
  MailApp.sendEmail({to:ADMIN_EMAIL, subject:sub, body:body});
}

// ── UTILITIES ─────────────────────────────────────────────
function ok(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function now(){ return Utilities.formatDate(new Date(),'Asia/Kolkata','dd-MMM-yyyy HH:mm'); }
function getRows(sheet) {
  const data=sheet.getDataRange().getValues();
  const hdr=data[0];
  return data.slice(1).map(function(row){
    var o={};hdr.forEach(function(h,i){var v=row[i];if(v instanceof Date)v=Utilities.formatDate(v,'Asia/Kolkata','yyyy-MM-dd');o[h]=v;});return o;
  }).filter(function(r){return r[hdr[0]];});
}
function fmtRowDate(v){if(!v)return'';if(v instanceof Date)return Utilities.formatDate(v,'Asia/Kolkata','yyyy-MM-dd');return String(v).substring(0,10);}
function getOrCreate(ss,name,headers,color){
  var sheet=ss.getSheetByName(name);
  if(!sheet){sheet=ss.insertSheet(name);sheet.appendRow(headers);var r=sheet.getRange(1,1,1,headers.length);r.setFontWeight('bold').setFontColor('#fff').setBackground(color||'#1A2642').setFontSize(10);sheet.setFrozenRows(1);}
  return sheet;
}
function updateOrderStatusById(orderId,status){
  var ss=getSS();var sheet=ss.getSheetByName(SH_ORDERS);if(!sheet||sheet.getLastRow()<2)return;
  var rows=sheet.getRange(2,1,sheet.getLastRow()-1,HDR_ORDERS.length).getValues();
  var idCol=HDR_ORDERS.indexOf('Order ID');var statCol=HDR_ORDERS.indexOf('Status');
  rows.forEach(function(row,i){if(String(row[idCol])===String(orderId))sheet.getRange(i+2,statCol+1).setValue(status);});
}
function logErr(ctx,msg){try{var ss=getSS();var sh=ss.getSheetByName(SH_LOG)||ss.insertSheet(SH_LOG);sh.appendRow([new Date().toISOString(),ctx,msg]);}catch(e){}}

// ── MONTHLY LEAVE CREDIT (run on last day of each month via Time-based trigger) ──
// Leave Policy:
//   Annual Leave: max 15, credited 1.25/month. Once annual hits 15, credit goes to Casual Leave.
//   Casual Leave: salary deducted; expires 31 March every year; max 2 per month.
function monthlyLeaveCredit() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SH_STAFF);
  if (!sheet || sheet.getLastRow() < 2) return;
  var MAX_ANNUAL = 15;
  var CREDIT = 1.25;
  var rows = getRows(sheet);
  var alCol = HDR_STAFF.indexOf('Annual Leave') + 1;
  var clCol = HDR_STAFF.indexOf('Casual Leave') + 1;
  var cuCol = HDR_STAFF.indexOf('Casual Used Month') + 1;
  for (var i = 0; i < rows.length; i++) {
    var al = parseFloat(rows[i]['Annual Leave']) || 0;
    var cl = parseFloat(rows[i]['Casual Leave']) || 0;
    var newAl = al, newCl = cl;
    if (al < MAX_ANNUAL) {
      newAl = Math.min(MAX_ANNUAL, al + CREDIT);
    } else {
      // Annual is full — credit goes to casual
      newCl = cl + CREDIT;
    }
    sheet.getRange(i + 2, alCol).setValue(newAl);
    sheet.getRange(i + 2, clCol).setValue(newCl);
    sheet.getRange(i + 2, cuCol).setValue(0); // reset casual-used-this-month counter
  }
  logErr('monthlyLeaveCredit', 'Credited ' + CREDIT + ' leaves to ' + rows.length + ' staff on ' + new Date().toDateString());
}

// Reset casual leave balance on 31 March each year
function resetCasualLeaveAnnually() {
  var today = new Date();
  if (today.getMonth() !== 2 || today.getDate() !== 31) return; // only run on March 31) return;
  var ss = getSS();
  var sheet = ss.getSheetByName(SH_STAFF);
  if (!sheet || sheet.getLastRow() < 2) return;
  var clCol = HDR_STAFF.indexOf('Casual Leave') + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var range = sheet.getRange(2, clCol, lastRow - 1, 1);
    var vals = range.getValues().map(function(){ return [0]; });
    range.setValues(vals);
  }
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monthlyLeaveCredit').timeBased().onMonthDay(28).atHour(23).create();
  ScriptApp.newTrigger('resetCasualLeaveAnnually').timeBased().onMonthDay(31).atHour(22).create();
}

// ── ONE-TIME SETUP ─────────────────────────────────────────
// Step 1: run setupSheets()  — creates all sheet tabs
// Step 2: run setupStaff()   — seeds staff and locations
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create('Flenbo Hub — Operations Platform');
    PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
    Logger.log('Created new spreadsheet: ' + ss.getUrl());
  } else {
    ss.rename('Flenbo Hub — Operations Platform');
  }
  getOrCreate(ss, SH_STAFF,     HDR_STAFF,     '#1A2642');
  getOrCreate(ss, SH_LOCATIONS, HDR_LOCATIONS, '#243254');
  getOrCreate(ss, SH_ORDERS,    HDR_ORDERS,    '#1A2642');
  getOrCreate(ss, SH_LEAVES,    HDR_LEAVES,    '#7b1fa2');
  getOrCreate(ss, SH_ATTEND,    HDR_ATTEND,    '#166534');
  getOrCreate(ss, SH_GRN,       HDR_GRN,       '#0d47a1');
  getOrCreate(ss, SH_WASTAGE,   HDR_WASTAGE,   '#c62828');
  getOrCreate(ss, SH_ITEMS,     HDR_ITEMS,     '#5B21B6');
  // Create ration category sheets
  RATION_CATS.forEach(function(cat) { getOrCreate(ss, cat, HDR_RATION, '#0f4c2a'); });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Step 1 done! All sheets created.\n\nNow select setupStaff and click Run.');
}

function setupStaff() {
  var ss = getSS();
  var staffSheet = ss.getSheetByName(SH_STAFF);
  if (staffSheet && staffSheet.getLastRow() < 2) {
    var staffData = [
      ['U001','Ravindra Singh Pundeer','Sous Chef',          'ravindrapundeer','982587','staff',  'Flenbo Nawada',  '','',15,   0,0,'Active'],
      ['U002','Amrendra Kumar',        'Utility',            'amrendrakumar',  '587436','staff',  'Flenbo Nawada',  '','',7.5,  0,0,'Active'],
      ['U003','Yogesh Rawat',          'Commi 1 - Tandoor', 'yogeshrawat',    '482695','staff',  'Emaar DigiHomes','','',11.25,0,0,'Active'],
      ['U004','Harish Singh Bisht',    'SR CDP - Indian',   'harishbisht',    '251423','staff',  'Emaar DigiHomes','','',0,    0,0,'Active'],
      ['U005','Umed Singh',            'Captain',            'umedsingh',      '893475','staff',  'Emaar DigiHomes','','',15,   0,0,'Active'],
      ['U006','Ramesh',                'Steward',            'ramesh',         '684256','staff',  'Emaar DigiHomes','','',1.25, 0,0,'Active'],
      ['U007','Roshan',                'Utility',            'roshan',         '374562','staff',  'Emaar DigiHomes','','',5,    0,0,'Active'],
      ['U008','Nandan Ram',            'Commi 1 - Indian',  'nandanram',      '758962','staff',  'Emaar DigiHomes','','',0,    0,0,'Active'],
      ['U009','Bharat Chopra',         'Operations Manager', 'bharatchopra',   '485397','manager','Corporate',      '','',0,    0,0,'Active'],
      ['U010','Nivedita Sahay',        'Director Operations','niveditasahay',  '452478','admin',  'Corporate',      '','',10,   0,0,'Active'],
      ['U011','Ashish Sahay',          'Director Comms',     'ashishsahay',    '597425','admin',  'Corporate',      '','',10,   0,0,'Active']
    ];
    staffSheet.getRange(2, 1, staffData.length, staffData[0].length).setValues(staffData);
  }
  var locSheet = ss.getSheetByName(SH_LOCATIONS);
  if (locSheet && locSheet.getLastRow() < 2) {
    var locData = [
      ['LOC001','Flenbo Nawada',   'Central Kitchen',   'Nawada, New Delhi',  '','Active'],
      ['LOC002','Emaar DigiHomes', 'Society Restaurant','Sector 62, Gurgaon', '','Active'],
      ['LOC003','Saan Verdante',   'Society Restaurant','Sector 95, Gurgaon', '','Active'],
      ['LOC004','Godrej Icon',     'Society Restaurant','Sector 88A, Gurgaon','','Active'],
      ['LOC005','Corporate',       'Head Office',       'Gurgaon',            '','Active']
    ];
    locSheet.getRange(2, 1, locData.length, locData[0].length).setValues(locData);
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Setup Complete!\n11 staff + 5 locations seeded.\n\nNow: Deploy > New deployment > Web App\nCopy the URL and paste into Flenbo_Hub.html');
}

// Alias — keeps any old references working
function setupSpreadsheet() { setupSheets(); }

// ── SEED BASE ITEMS ────────────────────────────────────────
// Run once from GAS editor to populate category sheets from Excel data.
// Safe to re-run: skips sheets that already have data.
function setupBaseItems() {
  var ss = getSS();
  var SEED = {
    'HYPERPURE': [['AASHIRVAAD AATA','KG'],['AASHIRVAAD MAIDA','KG'],['AJINA MOTO- GOLDEN CROWN','KG'],['AJWAIN','KG'],['ALMOND','KG'],['AMUL BUTTER 500 GMS','Pcs'],['AMUL CHEESE BLOCK 1 KG','KG'],['AMUL CHEESE DICED BLEND KG','KG'],['AMUL CREAM 1 LTR','Litres'],['ARHAR DAAL- GROWTH','KG'],['AROMATIC POWDER- KNORR','KG'],['BADI ELAICHI','KG'],['BEDMI POORI AATA- GM FOODS','KG'],['BESAN- RAJDHANI','KG'],['BREAD BROWN- BON','Pcs'],['BREAD CRUMS- PANCO','KG'],['BREAD WHITE- BON','Pcs'],['CASHEW','KG'],['CHAAT MASALA- EASTMADE','KG'],['CHANA DAAL- GROWTH','KG'],['CHICKEN MASALA- MDH','KG'],['CHILLY FLAKES SACHET- CHEFS ART','Packets'],['CHOLE MASALA- MDH','KG'],['CHOTI ELAICHI','KG'],["COCO POWDER- HERSHEY'S",'KG'],['COOKIES CASHEW BADAM- UNIBIC (450 GM)','Pcs'],['CORNFLOUR- JEECON','KG'],['DALCHINI','KG'],['DAWAT BIRYANI RICE- LONGEST GRAINS','KG'],['DEGGI MIRCH- MDH','KG'],['DHANIYA POWDER- CATCH/MDH','KG'],['DHULI MOONG DAAL','KG'],['EGGS','Dozen'],['FRAPPE POWDER- FRAPPIK','KG'],['GARAM MASALA','KG'],['GHEE- AMUL','KG'],['GREEN CHILLY SAUCE- TOPS/CHINGZ','Litres'],['GUD/JAGGERY','KG'],['HALDI POWDER- CATCH/MDH','KG'],['HING- MDH','KG'],['HONEY- NATURES NECTAR','KG'],['IMLLI SUKHI','KG'],['JAIFAL','KG'],['KALA NAMAK- CATCH/MDH','KG'],['KALI MIRCH SABOOT','KG'],['KALI URAD','KG'],['KASOORI METHI- MDH','KG'],['KESAR- BABY/MDH','KG'],['KITCHEN KING MASALA- MDH','KG'],['KITKAT','Pcs'],['LAL MIRCH POWDER- CATCH/MDH','KG'],['LAL MIRCH SABOOT WITHOUT STEM','KG'],['LAUNG','KG'],['LIGHT SOYA SAUCE- LEE KUM KEE','Litres'],['MAGAJ','KG'],['MIX DAAL','KG'],['MOUTH FRESHNER SWEET','KG'],['MUSTARD OIL- FORTUNE KACCHI GHANI','Litres'],['NACHOS- CORNITOS','Pcs'],['NOODLES- KIMS','Pcs'],['NUTELLA','Pcs'],['OLIVES','KG'],['OREGANO SACHET- CHEFS ART','Packets'],['OREO BISCUIT','Pcs'],['OYSTER SAUCE VEG- LEE KUM KEE','Litres'],['PAPAD','KG'],['PASTA PENNE- REGGIA','Pcs'],['PERI PERI MASALA- SNAPIN','KG'],['PILI SARSO','KG'],['POPPY SEEDS','KG'],['RAJMA MASALA- MDH/CATCH','KG'],['RAJMA- TRIPTI','KG'],['SABOOT JEERA- EASTMADE','KG'],['SHEZWAN SAUCE- VEEBA','Litres'],['SOYA SAUCE DARK- LEE KUM KEE','Litres'],['SOYA SAUCE LIGHT- LEE KUM KEE','Litres'],['SUGAR','KG'],['SUGAR FREE- GOLD','KG'],['SUGAR SACHET- DHAMPUR','Packets'],['SUNFLOWER OIL- FORTUNE','Litres'],['TATA SALT','KG'],['TEA- TATA AGNI CARDAMOM','KG'],['TEJ PATTA','KG'],['TOMATO KETCHUP- KISAN','Litres'],['TOMATO PUREE- GOLDEN CROWN','Litres'],['TOMATO SAUCE- VEEBA','Litres'],['VINEGAR- TOPS 1 LTR','Litres'],['WHITE PEPPER','KG']],
    'BEVERAGES': [['BISLERI 250 ML','Pcs'],['BLUE CURACAO SYRUP- 1 LTR','Litres'],['COFFEE BEANS 1 KG','KG'],['COKE- 2 LTRS','Pcs'],['COKE- 180 ML CAN','Pcs'],['COKE- 300 ML CAN','Pcs'],['COSMOPOLITAN- 250 ML CAN','Pcs'],['FRAPPE POWDER','KG'],['LEMON SYRUP- HABIT','Litres'],['ICE CUBES','KG'],['MINT SYRUP- MONIN','Litres'],['NESCAFE COFFEE POWDER','KG'],['REAL CANBERRY JUICE- 1 LTR','Litres'],['REAL MIX FRUIT JUICE- 1 LTR','Litres'],['REAL ORANGE JUICE- 1 LTR','Litres'],['REAL PINEAPPLE JUICE- 1 LTR','Litres'],['RED SUGAR SYRUP','Litres'],['SPRITE- 180 ML CAN','Pcs'],['SPRITE- 2 LTRS','Pcs'],['SPRITE- 300 ML CAN','Pcs'],['THUMS UP- 180 ML CAN','Pcs'],['THUMS UP- 2 LTRS','Pcs'],['THUMS UP- 300 ML CAN','Pcs']],
    'FROZEN': [['BASA FISH','KG'],['BROWNIE','KG'],['CHEESE BALLS','KG'],['CHICKEN MOMO','KG'],['CHICKEN SEEKH KABAB','KG'],['CHICKEN TANDOORI','KG'],['CHICKEN THIGH','KG'],['CHICKEN WINGS','KG'],['FRENCH FRIES','KG'],['FROZEN CORN','KG'],['FROZEN PEAS','KG'],['GULAB JAMUN 40GMS','KG'],['MUTTON','KG'],['PIZZA','KG'],['POTATO WEDGES','KG'],['PRAWNS','KG'],['SOYA CHAAP','KG'],['SPRING ROLL SHEET','KG'],['SPRING ROLLS','KG'],['VEG MOMO','KG']],
    'VEGETABLES': [['AALOO','KG'],['BAIGAN','KG'],['BASIL','KG'],['BEANS','KG'],['BEETROOT','KG'],['BHINDI','KG'],['BROCOLLI','KG'],['CABBAGE','KG'],['CAPSICUM','KG'],['CARROT','KG'],['CAULIFLOWER','KG'],['GARLIC','KG'],['GINGER','KG'],['GREEN CHILLY','KG'],['HARA DHANIYA','KG'],['LAUKI','KG'],['LEMON','KG'],['METHI','KG'],['MINT','KG'],['MUSHROOM','KG'],['PALAK','KG'],['PYAAZ','KG'],['RED CPSICUM','KG'],['TAMATAR','KG'],['YELLOW CAPSICUM','KG']],
    'DAIRY': [['DAHI','KG'],['KHOYA','KG'],['MILK','Litres'],['PANEER','KG']],
    'CLEANING': [['DETERGENT','KG'],['DETTOL','Bottles'],['DISH WASH LIQUID','Bottles'],['GARBAGE BAGS','Pcs'],['JHADU','Pcs'],['MICRO FIBER','Pcs'],['PHENYLE','Litres'],['POCHA','Pcs'],['RED CADDY','Pcs'],['SCOTCH BRITE','Pcs'],['SPONGE','Pcs'],['SPRAY BOTTLES','Pcs'],['STEEL SCRUB','Pcs'],['VIM BAR','Pcs'],['WIPER','Pcs']],
    'DISPOSABLES': [['3 CP CORNSTARCH CONTAINER','Pcs'],['300 ML GLASS BOTTLES WITH CAP','Pcs'],['300 ML PAPER CONTAINER WITH CAP','Pcs'],['300 ML PLASTIC SIPPER WITH CAP','Pcs'],['5 CP CORNSTARCH CONTAINER','Pcs'],['500 ML PAPER CONTAINER WITH CAP','Pcs'],['6 INCH PAPER PLATES','Pcs'],['650 ML PAPER RECTANGLE CONTAINER WITH PLASTIC COVER','Pcs'],['750 ML PAPER BOWL CONTAINER WITH CAP','Pcs'],['750 ML PAPER CONTAINER WITH CAP','Pcs'],['750 ML PLASTIC BOWL CONTAINER WITH CAP','Pcs'],['8 CP CORNSTARCH CONTAINER','Pcs'],['ALUMUNIUM FOIL','Rolls'],['BAGASS BOWLS','Pcs'],['BROWN PAPER BAG','Pcs'],['CLING WRAP','Rolls'],['DOUBLE PLY TISSUE','Rolls'],['FRUIT FORK','Pcs'],['HAIR NET','Pcs'],['PAPER CARRY BAGS','Pcs'],['PAPER STRAW','Pcs'],['PRINTER PAPER ROLL','Rolls'],['SHOE POLISH','Pcs'],['SINGLE PLY TISSUE','Rolls'],['TABLE SPRAY BOTTLE','Pcs'],['TOOTHPICK','Pcs'],['WHITE GLOVES','Pcs'],['WOODEN FORK','Pcs'],['WOODEN SPOONS','Pcs'],['WOODEN SPORK','Pcs'],['WOODEN STIRRER','Pcs'],['ZOMATO RED TAPES','Pcs']],
    'BIKANERVALA': [['GAJAR HALVA','KG'],['JALEBI','KG'],['KAJU KATLI','KG'],['MOONG DAL HALVA','KG'],['PHIRNI','KG'],['RASGULLA','KG'],['RASMALAI','KG'],['SPONGE RASGULLA','KG']],
    'STAFF RATION': [['AALOO','KG'],['BAIGAN','KG'],['BEANS','KG'],['BESAN','KG'],['BHINDI','KG'],['CABBAGE','KG'],['CARROT','KG'],['CAULIFLOWER','KG'],['CHAAT MASALA','KG'],['DAAL- ARHAR','KG'],['DAAL- DHULI MOONG','KG'],['DAAL- MIX','KG'],['GREEN CHILLY','KG'],['KALA CHANA','KG'],['LAUKI','KG'],['MUSTARD OIL','Litres'],['PALAK','KG'],['PYAAZ','KG'],['RAJMA','KG'],['REFINED OIL','Litres'],['STAFF AATA','KG'],['STAFF RICE','KG'],['SUGAR','KG'],['TAMATAR','KG'],['TATA SALT','KG'],['TEA- TATA AGNI CARDAMOM','KG']]
  };
  var today = now().split(' ')[0];
  RATION_CATS.forEach(function(cat) {
    var sheet = getOrCreate(ss, cat, HDR_RATION, '#0f4c2a');
    if (sheet.getLastRow() > 1) return; // already has data, skip
    var items = SEED[cat] || [];
    items.forEach(function(item, i) {
      sheet.appendRow([i+1, item[0], item[1], today, 'System']);
    });
  });
  SpreadsheetApp.flush();
  Logger.log('Base items seeded into all category sheets.');
}
